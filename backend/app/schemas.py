"""Esquemas Pydantic: definen y validan el JSON que entra y sale de la API."""
from datetime import datetime
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


# ── Posiciones ──
class PosicionIn(BaseModel):
    posicion: str
    descripcion: str = ""
    unidad: str = ""
    nivel: str = ""
    subnivel: str = ""
    grupo: str = ""
    area: str = ""
    area_cod: str = ""
    localidad: str = ""
    complejo: str = ""
    gerente: str = ""
    clasificacion: str = ""
    estado: str = ""
    semaforo: str = ""
    estatus: str = ""
    dias: int = 0
    dias_vencidos: int = 0
    vigencia_dias: int = 9999
    tipo_just: str = ""
    justificacion: str = ""


class PosicionOut(PosicionIn):
    model_config = ConfigDict(from_attributes=True)
    id: int


# ── Reportes ──
class ReporteIn(BaseModel):
    nombre_archivo: str
    etiqueta: str | None = None
    total_posiciones: int | None = None
    total_ocupadas: int | None = None
    posiciones: list[PosicionIn] = Field(default_factory=list)


class ReporteOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    nombre_archivo: str
    etiqueta: str
    periodo: int | None
    fecha_carga: datetime
    total_posiciones: int
    total_ocupadas: int
    total_vacantes: int


class ReporteDetalleOut(ReporteOut):
    posiciones: list[PosicionOut]


class ResumenOut(BaseModel):
    reporte_id: int
    complejo: str | None
    total_posiciones: int
    ocupadas: int
    vacantes: int
    cobertura_pct: float
    vacantes_por_semaforo: dict[str, int]
    vacantes_por_complejo: dict[str, int]
    promedio_dias_sin_ocupar: float


# ── Acciones (plan de acción) ──
Contexto = Literal["semaforo", "proyecto"]


class AccionIn(BaseModel):
    plan: str = ""
    responsable: str = ""
    fecha_compromiso: str = ""


class AccionOut(AccionIn):
    model_config = ConfigDict(from_attributes=True)
    id: int
    contexto: str
    posicion: str
    actualizado: datetime | None = None
