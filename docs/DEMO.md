# Guion de la demostración (≈ 10 minutos)

## Antes de empezar
- [ ] PostgreSQL encendido (`docker compose up -d`)
- [ ] Backend encendido (`uvicorn app.main:app --reload` dentro de `backend/`)
- [ ] Tener a la mano **dos** reportes semanales (p. ej. P28 y P29) — no están en el repo
- [ ] Pestañas abiertas: http://localhost:8000 y http://localhost:8000/docs
- [ ] (Opcional) Vaciar la base para empezar limpio: borrar reportes desde `/docs` con `DELETE /api/reportes/{id}`

## 1. El problema (1 min)
La herramienta original era un solo HTML: funcionaba, pero **no guardaba nada**. Al cerrar el navegador se perdían los planes de acción y para comparar semanas había que volver a subir cada Excel.

Mostrar `original/Macro_de_Analisis.html` y el primer commit en GitHub.

## 2. Arquitectura (1.5 min)
Explicar el diagrama del README: frontend → API REST (JSON) → PostgreSQL.
Puntos clave:
- El cálculo del semáforo se mantiene en el navegador (lógica ya validada).
- El backend **persiste** y **calcula KPIs con SQL**.
- Si el servidor no está, el dashboard sigue funcionando (modo local).

## 3. Frontend separado (1 min)
Abrir `frontend/` y mostrar `index.html`, `css/styles.css`, `js/app.js`, `js/api.js`.
"El diseño no cambió: la separación se hizo extrayendo los bloques tal cual."

## 4. Demo en vivo (4 min)
1. Abrir http://localhost:8000 → señalar **"Base de datos conectada"**.
2. Subir el reporte P28 → aparece el aviso *"Reporte P28-2026 guardado… (12,823 posiciones)"*.
3. Abrir **Rojo — Críticas**, escribir un plan de acción, responsable y fecha.
4. **Recargar la página** (F5) y volver a subir el Excel → el plan sigue ahí. *Esto antes no era posible.*
5. Subir el reporte P29 → abrir **Comparativo Semanal** → en *Historial guardado* hacer clic en P28 y P29 → se generan las gráficas de tendencia sin volver a subir archivos.
6. Pestaña **Análisis de Proyectos**: escribir una acción, recargar, mostrar que persiste.

## 5. La API (1.5 min)
En http://localhost:8000/docs:
- `GET /api/reportes` → historial.
- `GET /api/reportes/{id}/resumen` → mostrar que los totales coinciden con las tarjetas del dashboard.
- `GET /api/reportes/{id}/resumen?complejo=Norte` → filtro por complejo calculado en SQL.
- `GET /api/acciones/semaforo` → el plan capturado en el paso 4.

(Opcional) En pgAdmin: `SELECT etiqueta, total_vacantes FROM reportes;`

## 6. Calidad y control de versiones (1 min)
- Ejecutar `pytest -v` → 7 pruebas en verde.
- En GitHub, mostrar el historial de commits (una etapa por commit) y el `.gitignore` que impide subir Excel con datos de personal.

## Posibles preguntas
- **¿Por qué el Excel se procesa en el navegador y no en el servidor?** Se reutilizó la lógica ya probada en producción; mover el parseo a Python (pandas/openpyxl) es la siguiente mejora natural.
- **¿Qué pasa si subo el mismo archivo dos veces?** Se reemplaza; no se duplican datos.
- **¿Por qué los planes se guardan por número de posición?** Para que el seguimiento continúe semana a semana mientras la vacante siga abierta.
- **¿Seguridad?** Para un uso real faltaría autenticación (login) y desplegar con HTTPS.
