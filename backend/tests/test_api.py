"""Pruebas automáticas de la API.

Ejecutar desde la carpeta backend:
    pytest -v

Por defecto usan una base SQLite temporal (no tocan tu PostgreSQL).
Para probar contra PostgreSQL:
    TEST_DATABASE_URL=postgresql+psycopg2://vacantes:vacantes@localhost:5432/vacantes_test pytest -v
"""
import os
import tempfile

_tmp = os.path.join(tempfile.gettempdir(), "vacantes_test.db")
os.environ["DATABASE_URL"] = os.getenv("TEST_DATABASE_URL", f"sqlite:///{_tmp}")

import pytest  # noqa: E402
from fastapi.testclient import TestClient  # noqa: E402

from app.database import Base, engine  # noqa: E402
from app.main import app  # noqa: E402


@pytest.fixture()
def client():
    Base.metadata.drop_all(bind=engine)
    with TestClient(app) as c:
        yield c
    Base.metadata.drop_all(bind=engine)


def pos(num, estado="OCUPADA", complejo="Norte", semaforo="", dias=0):
    return {
        "posicion": str(num), "descripcion": f"Puesto {num}", "complejo": complejo,
        "estado": estado, "semaforo": semaforo, "dias": dias, "clasificacion": "PLANTA",
    }


def reporte(nombre, posiciones):
    return {"nombre_archivo": nombre, "posiciones": posiciones}


def test_health(client):
    assert client.get("/api/health").json() == {"status": "ok", "database": "ok"}


def test_guardar_y_resumen(client):
    data = reporte("REPORTE DE POSICIONES P28-2026.xlsx", [
        pos(1), pos(2), pos(3, "VACANTE", semaforo="Rojo", dias=70),
        pos(4, "VACANTE", "Sur", "Verde", 10),
    ])
    r = client.post("/api/reportes", json=data)
    assert r.status_code == 201
    body = r.json()
    assert body["etiqueta"] == "P28-2026"
    assert body["periodo"] == 2026028
    assert body["total_vacantes"] == 2

    res = client.get(f"/api/reportes/{body['id']}/resumen").json()
    assert res["total_posiciones"] == 4
    assert res["cobertura_pct"] == 50.0
    assert res["vacantes_por_semaforo"] == {"Rojo": 1, "Verde": 1}
    assert res["promedio_dias_sin_ocupar"] == 40.0

    norte = client.get(f"/api/reportes/{body['id']}/resumen", params={"complejo": "Norte"}).json()
    assert norte["vacantes"] == 1


def test_resubir_reemplaza(client):
    nombre = "REPORTE P29-2026.xlsx"
    client.post("/api/reportes", json=reporte(nombre, [pos(1)]))
    client.post("/api/reportes", json=reporte(nombre, [pos(1), pos(2, "VACANTE")]))
    lista = client.get("/api/reportes").json()
    assert len(lista) == 1
    assert lista[0]["total_posiciones"] == 2


def test_historial_ordenado_por_periodo(client):
    client.post("/api/reportes", json=reporte("REPORTE P30-2026.xlsx", [pos(1)]))
    client.post("/api/reportes", json=reporte("REPORTE P28-2026.xlsx", [pos(1)]))
    etiquetas = [r["etiqueta"] for r in client.get("/api/reportes").json()]
    assert etiquetas == ["P28-2026", "P30-2026"]


def test_filtrar_posiciones_y_eliminar(client):
    rid = client.post("/api/reportes", json=reporte("P31-2026.xlsx", [
        pos(1), pos(2, "VACANTE", semaforo="Rojo")])).json()["id"]
    vac = client.get(f"/api/reportes/{rid}/posiciones", params={"estado": "vacante"}).json()
    assert [p["posicion"] for p in vac] == ["2"]
    assert client.delete(f"/api/reportes/{rid}").status_code == 204
    assert client.get(f"/api/reportes/{rid}").status_code == 404


def test_reporte_vacio_rechazado(client):
    assert client.post("/api/reportes", json=reporte("x.xlsx", [])).status_code == 422


def test_acciones_upsert(client):
    url = "/api/acciones/semaforo/30000141"
    client.put(url, json={"plan": "Publicar vacante", "responsable": "Sindaly"})
    client.put(url, json={"plan": "Entrevistas el lunes", "responsable": "Sindaly",
                          "fecha_compromiso": "2026-11-03"})
    lista = client.get("/api/acciones/semaforo").json()
    assert len(lista) == 1
    assert lista[0]["plan"] == "Entrevistas el lunes"
    assert client.get("/api/acciones/proyecto").json() == []
    assert client.put("/api/acciones/otro/1", json={}).status_code == 422
