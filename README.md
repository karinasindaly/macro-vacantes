# Control de Vacantes — Atracción de Talento

Dashboard para el seguimiento semanal de vacantes de Pilgrim's (Semáforo de Vacantes, Detalle de Posiciones y Análisis de Proyectos), migrado de un único archivo HTML a una aplicación **frontend + backend + base de datos**.

| Capa | Tecnología | Qué hace |
|---|---|---|
| Frontend | HTML, CSS, JavaScript, SheetJS, Chart.js | Lee el Excel semanal, calcula el semáforo y muestra los tableros |
| Backend | Python 3.11+, FastAPI, SQLAlchemy | API REST que guarda y consulta reportes y planes de acción |
| Base de datos | PostgreSQL 16 | Historial de reportes, posiciones y planes de acción |

## Participantes

| Nombre | Rol en el proyecto | GitHub |
|---|---|---|
| Sindaly | _(completar)_ | @_(usuario)_ |
| _(nombre)_ | _(completar)_ | @_(usuario)_ |

---

## Arquitectura

```
 Navegador (frontend/)                    Servidor (backend/)              PostgreSQL
 ┌────────────────────────┐  JSON / HTTP  ┌──────────────────────┐  SQL   ┌──────────────┐
 │ 1. Sube el Excel        │ ───────────▶ │ POST /api/reportes    │ ────▶ │ reportes     │
 │ 2. SheetJS lo procesa   │              │ GET  /api/reportes    │       │ posiciones   │
 │ 3. Muestra semáforo     │ ◀─────────── │ PUT  /api/acciones/…  │ ◀──── │ acciones     │
 │ 4. Captura planes       │              │ GET  …/resumen (KPIs) │       └──────────────┘
 └────────────────────────┘              └──────────────────────┘
```

**Flujo:** el navegador procesa el Excel (lógica original, ya probada) y envía las posiciones resultantes a la API, que las guarda en PostgreSQL. Los planes de acción se guardan **por número de posición**, por lo que se conservan al cargar el reporte de la semana siguiente. El *Comparativo Semanal* puede armarse con reportes guardados, sin volver a subir cada Excel.

Si el backend no está encendido, el dashboard funciona igual que la versión original (indicador **"Modo local"** en el encabezado).

## Estructura del repositorio

```
macro-vacantes/
├── original/Macro_de_Analisis.html   ← versión inicial (un solo archivo)
├── frontend/
│   ├── index.html                    ← estructura de la página
│   ├── css/styles.css                ← estilos (diseño original)
│   └── js/
│       ├── app.js                    ← lógica del dashboard
│       └── api.js                    ← cliente de la API REST
├── backend/
│   ├── app/
│   │   ├── main.py                   ← arranque de FastAPI, CORS, sirve el frontend
│   │   ├── database.py               ← conexión a PostgreSQL
│   │   ├── models.py                 ← tablas (SQLAlchemy)
│   │   ├── schemas.py                ← validación de datos (Pydantic)
│   │   └── routers/
│   │       ├── reportes.py           ← endpoints de reportes y posiciones
│   │       └── acciones.py           ← endpoints de planes de acción
│   ├── tests/test_api.py             ← pruebas automáticas (pytest)
│   ├── requirements.txt
│   └── .env.example
├── docs/DEMO.md                      ← guion de la demostración
└── docker-compose.yml                ← PostgreSQL listo para usar
```

---

## Instalación y ejecución

### Requisitos
- Python 3.11 o superior
- PostgreSQL 16 **o** Docker Desktop
- Git

### 1. Clonar el repositorio
```bash
git clone https://github.com/<usuario>/macro-vacantes.git
cd macro-vacantes
```

### 2. Levantar PostgreSQL
**Opción A — Docker (recomendada):**
```bash
docker compose up -d
```
**Opción B — PostgreSQL instalado:** en pgAdmin o `psql` ejecuta:
```sql
CREATE USER vacantes WITH PASSWORD 'vacantes';
CREATE DATABASE vacantes_db OWNER vacantes;
```

### 3. Instalar y arrancar el backend
```bash
cd backend
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
source .venv/bin/activate

pip install -r requirements.txt
copy .env.example .env        # Windows   (en macOS/Linux: cp .env.example .env)
uvicorn app.main:app --reload
```
Las tablas se crean solas la primera vez.

### 4. Abrir el sistema
| Qué | Dirección |
|---|---|
| Dashboard | http://localhost:8000 |
| Documentación interactiva de la API (Swagger) | http://localhost:8000/docs |
| Verificación de salud | http://localhost:8000/api/health |

El encabezado debe mostrar **"Base de datos conectada"**.

---

## API REST

| Método | Ruta | Descripción |
|---|---|---|
| GET | `/api/health` | Estado de la API y la base de datos |
| POST | `/api/reportes` | Guarda un reporte con sus posiciones (si el archivo ya existe, lo reemplaza) |
| GET | `/api/reportes` | Historial de reportes ordenado por periodo (P28, P29…) |
| GET | `/api/reportes/{id}` | Reporte con todas sus posiciones |
| GET | `/api/reportes/{id}/posiciones?estado=&complejo=&semaforo=` | Posiciones filtradas |
| GET | `/api/reportes/{id}/resumen?complejo=` | KPIs calculados en SQL: cobertura, vacantes por semáforo y complejo, promedio de días |
| DELETE | `/api/reportes/{id}` | Elimina un reporte |
| GET | `/api/acciones/{semaforo\|proyecto}` | Planes de acción guardados |
| PUT | `/api/acciones/{contexto}/{posicion}` | Crea o actualiza un plan de acción |
| DELETE | `/api/acciones/{contexto}/{posicion}` | Elimina un plan de acción |

## Base de datos

- **reportes**: `id`, `nombre_archivo` (único), `etiqueta` (P28-2026), `periodo` (2026028, para ordenar), `fecha_carga`, totales.
- **posiciones**: una fila por posición del reporte (`posicion`, `complejo`, `gerente`, `estado`, `semaforo`, `dias`, `clasificacion`, `area_cod`, `justificacion`…). Se borra en cascada con su reporte.
- **acciones**: `contexto` (semaforo / proyecto) + `posicion` (únicos juntos), `plan`, `responsable`, `fecha_compromiso`.

## Pruebas

```bash
cd backend
pytest -v
```
Usan una base SQLite temporal, así que no tocan tus datos. Para probar contra PostgreSQL:
```bash
TEST_DATABASE_URL=postgresql+psycopg2://vacantes:vacantes@localhost:5432/vacantes_test pytest -v
```

Validado con el reporte real P28-2026: 12,823 posiciones guardadas; los KPIs que calcula la base de datos (313 vacantes, 15 rojo / 9 amarillo / 289 verde, 97.6 % de cobertura) coinciden con los del dashboard.

## Confidencialidad

Los reportes de Excel contienen datos de personal. El `.gitignore` bloquea `*.xlsx` / `*.xls` para que **nunca** se suban al repositorio, y el archivo `.env` tampoco se sube.

## Historial de desarrollo

Cada etapa quedó en su propio commit (`git log --oneline`):
1. Versión inicial (un solo HTML)
2. Separación del frontend
3. Backend FastAPI
4. Base de datos PostgreSQL
5. Pruebas automáticas
6. Integración frontend ↔ API
7. Documentación
