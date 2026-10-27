"""Modelos (tablas) de la base de datos.

- reportes:   un registro por cada Excel semanal cargado (P28-2026, P29-2026...).
- posiciones: todas las posiciones de ese reporte (ocupadas y vacantes).
- acciones:   planes de acción capturados en el dashboard. Se guardan por
              número de posición, así se conservan de una semana a otra.
"""
from datetime import datetime

from sqlalchemy import (
    DateTime,
    ForeignKey,
    Integer,
    String,
    Text,
    UniqueConstraint,
    func,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .database import Base


class Reporte(Base):
    __tablename__ = "reportes"

    id: Mapped[int] = mapped_column(primary_key=True)
    nombre_archivo: Mapped[str] = mapped_column(String(255), unique=True, index=True)
    etiqueta: Mapped[str] = mapped_column(String(120))  # p. ej. "P28-2026"
    periodo: Mapped[int | None] = mapped_column(Integer, index=True)  # 2026028 para ordenar
    fecha_carga: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    total_posiciones: Mapped[int] = mapped_column(Integer, default=0)
    total_ocupadas: Mapped[int] = mapped_column(Integer, default=0)
    total_vacantes: Mapped[int] = mapped_column(Integer, default=0)

    posiciones: Mapped[list["Posicion"]] = relationship(
        back_populates="reporte", cascade="all, delete-orphan"
    )


class Posicion(Base):
    __tablename__ = "posiciones"

    id: Mapped[int] = mapped_column(primary_key=True)
    reporte_id: Mapped[int] = mapped_column(
        ForeignKey("reportes.id", ondelete="CASCADE"), index=True
    )
    posicion: Mapped[str] = mapped_column(String(40), index=True)
    descripcion: Mapped[str] = mapped_column(String(255), default="")
    unidad: Mapped[str] = mapped_column(String(255), default="")
    nivel: Mapped[str] = mapped_column(String(40), default="")
    subnivel: Mapped[str] = mapped_column(String(40), default="")
    grupo: Mapped[str] = mapped_column(String(255), default="")
    area: Mapped[str] = mapped_column(String(255), default="")
    area_cod: Mapped[str] = mapped_column(String(20), default="", index=True)
    localidad: Mapped[str] = mapped_column(String(255), default="")
    complejo: Mapped[str] = mapped_column(String(60), default="", index=True)
    gerente: Mapped[str] = mapped_column(String(255), default="")
    clasificacion: Mapped[str] = mapped_column(String(40), default="")
    estado: Mapped[str] = mapped_column(String(20), default="", index=True)  # VACANTE / OCUPADA
    semaforo: Mapped[str] = mapped_column(String(20), default="")  # Rojo / Amarillo / Verde
    estatus: Mapped[str] = mapped_column(String(40), default="")  # VIGENTE / VENCIDA
    dias: Mapped[int] = mapped_column(Integer, default=0)
    dias_vencidos: Mapped[int] = mapped_column(Integer, default=0)
    vigencia_dias: Mapped[int] = mapped_column(Integer, default=9999)
    tipo_just: Mapped[str] = mapped_column(String(255), default="")
    justificacion: Mapped[str] = mapped_column(Text, default="")

    reporte: Mapped[Reporte] = relationship(back_populates="posiciones")


class Accion(Base):
    __tablename__ = "acciones"
    __table_args__ = (UniqueConstraint("contexto", "posicion", name="uq_accion_contexto_posicion"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    contexto: Mapped[str] = mapped_column(String(20), index=True)  # "semaforo" | "proyecto"
    posicion: Mapped[str] = mapped_column(String(40), index=True)
    plan: Mapped[str] = mapped_column(Text, default="")
    responsable: Mapped[str] = mapped_column(String(255), default="")
    fecha_compromiso: Mapped[str] = mapped_column(String(20), default="")  # AAAA-MM-DD
    actualizado: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )
