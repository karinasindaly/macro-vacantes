// =============================================================
// ── CLIENTE DE LA API REST ──
// Conecta el dashboard con el backend FastAPI + PostgreSQL.
// Si el servidor no está disponible, el dashboard sigue funcionando
// igual que la versión original (modo local, sin guardar nada).
// =============================================================
const API = (() => {
  // Si la página la sirve el propio backend (puerto 8000) se usan rutas
  // relativas; si se abre como archivo o desde otro servidor (Live Server),
  // se apunta a http://localhost:8000. Se puede forzar con window.API_BASE.
  const BASE = window.API_BASE ||
    (location.protocol.startsWith('http') && location.port === '8000' ? '' : 'http://localhost:8000');

  let online = false;

  async function req(path, options = {}) {
    const resp = await fetch(BASE + path, {
      headers: { 'Content-Type': 'application/json' },
      ...options,
    });
    if (!resp.ok) {
      let detalle = resp.statusText;
      try { detalle = (await resp.json()).detail || detalle; } catch (e) {}
      throw new Error(typeof detalle === 'string' ? detalle : JSON.stringify(detalle));
    }
    return resp.status === 204 ? null : resp.json();
  }

  // ── Indicador de conexión en el encabezado ──
  function pintarEstado() {
    const el = document.getElementById('api-status');
    if (!el) return;
    el.className = 'api-status ' + (online ? 'api-on' : 'api-off');
    el.innerHTML = online
      ? '<span class="api-dot"></span>Base de datos conectada'
      : '<span class="api-dot"></span>Modo local (sin servidor)';
    el.title = online
      ? 'Los reportes y planes de acción se guardan en PostgreSQL'
      : 'Inicia el backend (uvicorn app.main:app) para guardar la información';
  }

  function aviso(msg, tipo = 'ok') {
    let t = document.getElementById('api-toast');
    if (!t) {
      t = document.createElement('div');
      t.id = 'api-toast';
      document.body.appendChild(t);
    }
    t.className = 'api-toast api-toast-' + tipo + ' show';
    t.textContent = msg;
    clearTimeout(t._timer);
    t._timer = setTimeout(() => t.classList.remove('show'), 3500);
  }

  async function ping() {
    try { await req('/api/health'); online = true; } catch (e) { online = false; }
    pintarEstado();
    return online;
  }

  const cut = (v, n) => String(v ?? '').slice(0, n);

  // ── Reportes ──
  // Recibe las filas ya procesadas por la pestaña Detalle (detAllRows)
  // y las envía al backend para guardarlas como un reporte semanal.
  async function guardarReporte(nombreArchivo, filas) {
    if (!online && !(await ping())) return null;
    const posiciones = filas.map(r => ({
      posicion: cut(r.posicion, 40),
      descripcion: cut(r.descripcion, 255),
      unidad: cut(r.unidad, 255),
      nivel: cut(r.nivel, 40),
      subnivel: cut(r.subnivel, 40),
      grupo: cut(r.grupo, 255),
      area: cut(r.area, 255),
      area_cod: cut(r.areaCod, 20),
      localidad: cut(r.localidad, 255),
      complejo: cut(r.complejo, 60),
      gerente: cut(r.gerente, 255),
      clasificacion: cut(r.clasif, 40),
      estado: cut(r.estado, 20),
      semaforo: cut(r.semaforo, 20),
      estatus: cut(r.estatus, 40),
      dias: Number(r.dias) || 0,
      dias_vencidos: Number(r.diasVenc) || 0,
      vigencia_dias: Number(r.vigDias) || 9999,
      tipo_just: cut(r.tipoJust, 255),
      justificacion: String(r.justificacion ?? ''),
    }));
    try {
      const rep = await req('/api/reportes', {
        method: 'POST',
        body: JSON.stringify({ nombre_archivo: nombreArchivo, posiciones }),
      });
      aviso(`Reporte ${rep.etiqueta} guardado en la base de datos (${rep.total_posiciones} posiciones)`);
      return rep;
    } catch (e) {
      aviso('No se pudo guardar el reporte: ' + e.message, 'error');
      return null;
    }
  }

  const listarReportes = () => req('/api/reportes');
  const obtenerReporte = id => req('/api/reportes/' + id);

  // ── Planes de acción ──
  // Devuelve un objeto { posicion: {plan, responsable, fecha_compromiso} }
  async function listarAcciones(contexto) {
    if (!online) return {};
    try {
      const lista = await req('/api/acciones/' + contexto);
      return Object.fromEntries(lista.map(a => [a.posicion, a]));
    } catch (e) { return {}; }
  }

  async function guardarAccion(contexto, posicion, datos) {
    if (!online || !posicion) return null;
    try {
      return await req(`/api/acciones/${contexto}/${encodeURIComponent(posicion)}`, {
        method: 'PUT',
        body: JSON.stringify({
          plan: datos.plan || '',
          responsable: datos.responsable || '',
          fecha_compromiso: datos.fecha || datos.fecha_compromiso || '',
        }),
      });
    } catch (e) {
      aviso('No se pudo guardar el plan de acción: ' + e.message, 'error');
      return null;
    }
  }

  document.addEventListener('DOMContentLoaded', ping);

  return {
    ping, aviso, guardarReporte, listarReportes, obtenerReporte,
    listarAcciones, guardarAccion,
    get online() { return online; },
  };
})();
