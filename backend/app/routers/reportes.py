"""Endpoints de reportes semanales y sus posiciones."""
import re

from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy import func, select
from sqlalchemy.orm import Session, selectinload

from .. import models, schemas
from ..database import get_db

router = APIRouter(prefix="/api/reportes", tags=["Reportes"])


def detectar_periodo(nombre: str) -> tuple[str, int | None]:
    """'REPORTE DE POSICIONES P28-2026.xlsx' -> ('P28-2026', 2026028)."""
    base = re.sub(r"\.[^.]+$", "", nombre or "").strip()
    m_p = re.search(r"P\s*-?\s*(\d{1,3})", base, re.I)
    m_y = re.search(r"(\d{4})", base)
    if not m_p:
        return base, None
    num = int(m_p.group(1))
    year = int(m_y.group(1)) if m_y else 0
    etiqueta = f"P{num:02d}-{year}" if year else f"P{num:02d}"
    return etiqueta, year * 1000 + num


def es_vacante(estado: str) -> bool:
    return (estado or "").strip().upper() == "VACANTE"


@router.post("", response_model=schemas.ReporteOut, status_code=status.HTTP_201_CREATED)
def guardar_reporte(payload: schemas.ReporteIn, db: Session = Depends(get_db)):
    """Guarda un reporte con todas sus posiciones.

    Si ya existe un reporte con el mismo nombre de archivo se reemplaza
    (volver a subir el mismo Excel no duplica datos).
    """
    if not payload.posiciones:
        raise HTTPException(status_code=422, detail="El reporte no trae posiciones.")

    existente = db.scalar(
        select(models.Reporte).where(models.Reporte.nombre_archivo == payload.nombre_archivo)
    )
    if existente:
        db.delete(existente)
        db.flush()

    etiqueta, periodo = detectar_periodo(payload.nombre_archivo)
    vacantes = sum(1 for p in payload.posiciones if es_vacante(p.estado))
    total = payload.total_posiciones or len(payload.posiciones)
    ocupadas = payload.total_ocupadas if payload.total_ocupadas is not None else total - vacantes

    reporte = models.Reporte(
        nombre_archivo=payload.nombre_archivo,
        etiqueta=payload.etiqueta or etiqueta,
        periodo=periodo,
        total_posiciones=total,
        total_ocupadas=ocupadas,
        total_vacantes=vacantes,
    )
    reporte.posiciones = [models.Posicion(**p.model_dump()) for p in payload.posiciones]
    db.add(reporte)
    db.commit()
    db.refresh(reporte)
    return reporte


@router.get("", response_model=list[schemas.ReporteOut])
def listar_reportes(db: Session = Depends(get_db)):
    """Historial de reportes, del periodo más antiguo al más reciente."""
    stmt = select(models.Reporte).order_by(
        models.Reporte.periodo.asc().nulls_last(), models.Reporte.fecha_carga.asc()
    )
    return db.scalars(stmt).all()


def _obtener(db: Session, reporte_id: int) -> models.Reporte:
    reporte = db.get(models.Reporte, reporte_id)
    if not reporte:
        raise HTTPException(status_code=404, detail="Reporte no encontrado.")
    return reporte


@router.get("/{reporte_id}", response_model=schemas.ReporteDetalleOut)
def obtener_reporte(reporte_id: int, db: Session = Depends(get_db)):
    stmt = (
        select(models.Reporte)
        .options(selectinload(models.Reporte.posiciones))
        .where(models.Reporte.id == reporte_id)
    )
    reporte = db.scalar(stmt)
    if not reporte:
        raise HTTPException(status_code=404, detail="Reporte no encontrado.")
    return reporte


@router.get("/{reporte_id}/posiciones", response_model=list[schemas.PosicionOut])
def listar_posiciones(
    reporte_id: int,
    estado: str | None = Query(None, description="VACANTE u OCUPADA"),
    complejo: str | None = None,
    semaforo: str | None = Query(None, description="Rojo, Amarillo o Verde"),
    db: Session = Depends(get_db),
):
    _obtener(db, reporte_id)
    stmt = select(models.Posicion).where(models.Posicion.reporte_id == reporte_id)
    if estado:
        stmt = stmt.where(func.upper(models.Posicion.estado) == estado.upper())
    if complejo:
        stmt = stmt.where(models.Posicion.complejo == complejo)
    if semaforo:
        stmt = stmt.where(models.Posicion.semaforo == semaforo)
    return db.scalars(stmt.order_by(models.Posicion.posicion)).all()


@router.get("/{reporte_id}/resumen", response_model=schemas.ResumenOut)
def resumen(reporte_id: int, complejo: str | None = None, db: Session = Depends(get_db)):
    """KPIs calculados en la base de datos (totales, cobertura, semáforo)."""
    _obtener(db, reporte_id)
    P = models.Posicion
    base = [P.reporte_id == reporte_id]
    if complejo:
        base.append(P.complejo == complejo)
    vac = func.upper(P.estado) == "VACANTE"

    total = db.scalar(select(func.count()).select_from(P).where(*base)) or 0
    vacantes = db.scalar(select(func.count()).select_from(P).where(*base, vac)) or 0
    ocupadas = total - vacantes
    prom = db.scalar(select(func.avg(P.dias)).where(*base, vac)) or 0

    por_sem = dict(db.execute(select(P.semaforo, func.count()).where(*base, vac).group_by(P.semaforo)).all())
    por_comp = dict(db.execute(select(P.complejo, func.count()).where(*base, vac).group_by(P.complejo)).all())

    return schemas.ResumenOut(
        reporte_id=reporte_id,
        complejo=complejo,
        total_posiciones=total,
        ocupadas=ocupadas,
        vacantes=vacantes,
        cobertura_pct=round(ocupadas / total * 100, 1) if total else 0.0,
        vacantes_por_semaforo={k or "Sin dato": v for k, v in por_sem.items()},
        vacantes_por_complejo={k or "Sin complejo": v for k, v in por_comp.items()},
        promedio_dias_sin_ocupar=round(float(prom), 1),
    )


@router.delete("/{reporte_id}", status_code=status.HTTP_204_NO_CONTENT)
def eliminar_reporte(reporte_id: int, db: Session = Depends(get_db)):
    db.delete(_obtener(db, reporte_id))
    db.commit()
