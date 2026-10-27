"""Endpoints para los planes de acción (Semáforo y Análisis de Proyectos).

Se guardan por número de posición, así el plan capturado una semana
sigue apareciendo cuando se carga el reporte de la semana siguiente.
"""
from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy import select
from sqlalchemy.orm import Session

from .. import models, schemas
from ..database import get_db

router = APIRouter(prefix="/api/acciones", tags=["Planes de acción"])


@router.get("/{contexto}", response_model=list[schemas.AccionOut])
def listar_acciones(contexto: schemas.Contexto, db: Session = Depends(get_db)):
    stmt = select(models.Accion).where(models.Accion.contexto == contexto)
    return db.scalars(stmt).all()


@router.put("/{contexto}/{posicion}", response_model=schemas.AccionOut)
def guardar_accion(
    contexto: schemas.Contexto,
    posicion: str,
    payload: schemas.AccionIn,
    db: Session = Depends(get_db),
):
    """Crea o actualiza (upsert) el plan de acción de una posición."""
    accion = db.scalar(
        select(models.Accion).where(
            models.Accion.contexto == contexto, models.Accion.posicion == posicion
        )
    )
    if accion is None:
        accion = models.Accion(contexto=contexto, posicion=posicion)
        db.add(accion)
    accion.plan = payload.plan
    accion.responsable = payload.responsable
    accion.fecha_compromiso = payload.fecha_compromiso
    db.commit()
    db.refresh(accion)
    return accion


@router.delete("/{contexto}/{posicion}", status_code=status.HTTP_204_NO_CONTENT)
def eliminar_accion(contexto: schemas.Contexto, posicion: str, db: Session = Depends(get_db)):
    accion = db.scalar(
        select(models.Accion).where(
            models.Accion.contexto == contexto, models.Accion.posicion == posicion
        )
    )
    if not accion:
        raise HTTPException(status_code=404, detail="Plan de acción no encontrado.")
    db.delete(accion)
    db.commit()
