"""Punto de entrada de la API REST — Control de Vacantes Pilgrim's.

Ejecutar (desde la carpeta backend):
    uvicorn app.main:app --reload

- API:            http://localhost:8000/api/...
- Documentación:  http://localhost:8000/docs
- Dashboard:      http://localhost:8000/
"""
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import Depends, FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from sqlalchemy import text
from sqlalchemy.orm import Session

from .database import CORS_ORIGINS, Base, engine, get_db
from .routers import acciones, reportes

FRONTEND_DIR = Path(__file__).resolve().parents[2] / "frontend"


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Crea las tablas si no existen (para este proyecto no se usan migraciones).
    Base.metadata.create_all(bind=engine)
    yield


app = FastAPI(
    title="API Control de Vacantes",
    description="Backend del dashboard de Atracción de Talento: guarda reportes semanales, "
    "posiciones y planes de acción en PostgreSQL.",
    version="1.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(reportes.router)
app.include_router(acciones.router)


@app.get("/api/health", tags=["Sistema"])
def health(db: Session = Depends(get_db)):
    """Verifica que la API y la base de datos respondan."""
    db.execute(text("SELECT 1"))
    return {"status": "ok", "database": "ok"}


# El frontend se sirve desde la misma app (montado al final para no tapar /api).
if FRONTEND_DIR.exists():
    app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")
