/* =====================================================================
 * PAGOS Y FLUJO · Administración — Frontend
 * Pantallas: Programación semanal · Registrar pago · Proveedores · Flujo por proyecto
 * Backend: Google Apps Script (Code.gs) + Google Drive
 * ===================================================================== */
'use strict';

/* ============ CONFIGURACIÓN ============ */
const API_URL_DEFAULT = 'https://script.google.com/macros/s/AKfycbz-LNkv5JRcpKo6pLlzd5gBpOH6u3zOYbF3RN0DNjMgVz64hEmQEcx-TYL6TWTgFZA/exec';
const MAX_MB = 30;
const XLSX_CDN = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';

const DOCS = [
  { tipo: 'OC',          label: 'Orden de compra',       accept: '.pdf,.jpg,.jpeg,.png,.xlsx,.docx', col: 'OC_URL' },
  { tipo: 'COTIZACION',  label: 'Cotización',            accept: '.pdf,.jpg,.jpeg,.png,.xlsx,.docx', col: 'COTIZACION_URL' },
  { tipo: 'PAGO',        label: 'Comprobante de pago',   accept: '.pdf,.jpg,.jpeg,.png',             col: 'PAGO_URL' },
  { tipo: 'FACTURA_PDF', label: 'Factura PDF',           accept: '.pdf',                             col: 'FACTURA_PDF_URL' },
  { tipo: 'FACTURA_XML', label: 'Factura XML',           accept: '.xml',                             col: 'FACTURA_XML_URL' },
  { tipo: 'COMPLEMENTO', label: 'Complemento de pago',   accept: '.xml,.pdf',                        col: 'COMPLEMENTO_URL' },
  { tipo: 'OTRO',        label: 'Otros (tickets, fotos…)', accept: '',                               col: 'OTROS_URL' }
];
const EXP_COLS = ['OC_URL', 'COTIZACION_URL', 'PAGO_URL', 'FACTURA_PDF_URL', 'FACTURA_XML_URL', 'COMPLEMENTO_URL'];

const FORMAS_PAGO = ['TRANSFERENCIA 8018', 'EFECTIVALE', 'DEPÓSITO CUENTA PERSONAL', 'AMEX', 'EFECTIVO'];
const CUENTAS_ORIGEN = ['CUENTA EMPRESA 8018', 'EFECTIVALE', 'AMEX', 'EFECTIVO'];

// Clave de banco = 3 primeros dígitos de la CLABE
const BANCOS = { '002': 'BANAMEX', '012': 'BBVA', '014': 'SANTANDER', '021': 'HSBC', '030': 'BAJIO',
  '036': 'INBURSA', '044': 'SCOTIABANK', '058': 'BANREGIO', '062': 'AFIRME', '072': 'BANORTE',
  '127': 'AZTECA', '137': 'BANCOPPEL', '646': 'STP' };

const ESTATUS = {
  'PROGRAMADO': ['Programado', 'c-info'],
  'PAGO PARCIAL': ['Pago parcial', 'c-warn'],
  'PAGADO - FALTAN DOCUMENTOS': ['Pagado · faltan docs', 'c-warn'],
  'COMPLETO': ['Completo', 'c-ok']
};

/* ============ ESTADO ============ */
const cfg = {
  get url() { return localStorage.getItem('pa_url') || API_URL_DEFAULT; },
  get key() { return localStorage.getItem('pa_key') || ''; },
  save(url, key) { localStorage.setItem('pa_url', url.trim()); localStorage.setItem('pa_key', key.trim()); }
};

const state = {
  data: null,          // datos normalizados
  view: 'programacion',
  mes: '',             // 'YYYY-MM'
  proyecto: 'TODOS',
  semana: 'TODAS',
  buscar: '',
  provSel: '',
  flujoProy: '',
  flujoVista: 'mes',
  pago: { proveedor: '', globalTocado: false },
  expediente: null     // partida abierta en el diálogo
};

/* ============ UTILIDADES ============ */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const num = v => { if (typeof v === 'number') return v; const n = parseFloat(String(v ?? '').replace(/[^0-9.\-]/g, '')); return isNaN(n) ? 0 : n; };
const r2 = n => Math.round((n + Number.EPSILON) * 100) / 100;
const up = s => String(s ?? '').toUpperCase().trim();
const sum = (arr, f) => r2(arr.reduce((a, x) => a + f(x), 0));
const uniq = arr => [...new Set(arr.filter(Boolean))];
const money = n => { n = num(n); return (n < 0 ? '-$' : '$') + Math.abs(n).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); };
const pad2 = n => String(n).padStart(2, '0');
const isoDate = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const hoy = () => isoDate(new Date());
const parseISO = s => { const [y, m, d] = String(s).slice(0, 10).split('-').map(Number); return new Date(y, (m || 1) - 1, d || 1); };
const fechaMx = s => { if (!s) return ''; const [y, m, d] = String(s).slice(0, 10).split('-'); return `${d}/${m}/${y}`; };
const mesLabel = ym => { if (!ym) return ''; const [y, m] = ym.split('-').map(Number); return new Date(y, m - 1, 1).toLocaleDateString('es-MX', { month: 'long', year: 'numeric' }); };
const mesCorto = ym => { const [y, m] = ym.split('-').map(Number); return new Date(y, m - 1, 1).toLocaleDateString('es-MX', { month: 'short', year: 'numeric' }); };
const lunesISO = s => { if (!s) return ''; const d = parseISO(s); d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); return isoDate(d); };
const semanaLabel = lun => { const a = parseISO(lun), b = parseISO(lun); b.setDate(b.getDate() + 6); const f = d => d.toLocaleDateString('es-MX', { day: 'numeric', month: 'short' }); return `${f(a)} – ${f(b)}`; };
const pl = (n, uno, varios) => `${n} ${n === 1 ? uno : varios}`;
const pct = (a, b) => (b > 0 ? Math.round(a / b * 100) : 0);

function clabeValida(c) {
  if (!/^\d{18}$/.test(c)) return false;
  const pesos = [3, 7, 1];
  let s = 0;
  for (let i = 0; i < 17; i++) s += (parseInt(c[i], 10) * pesos[i % 3]) % 10;
  return (10 - (s % 10)) % 10 === parseInt(c[17], 10);
}

function toast(msg, tipo = 'ok', ms = 4500) {
  const t = $('#toast');
  const map = { ok: ['var(--ok-bg)', 'var(--ok)'], warn: ['var(--warn-bg)', 'var(--warn)'], err: ['var(--err-bg)', 'var(--err)'], info: ['var(--info-bg)', 'var(--info)'] };
  t.style.background = map[tipo][0]; t.style.color = map[tipo][1];
  t.textContent = msg; t.style.display = 'block';
  clearTimeout(t._h); t._h = setTimeout(() => (t.style.display = 'none'), ms);
}

const chipEstatus = e => { const [l, c] = ESTATUS[e] || [e || 'Programado', 'c-info']; return `<span class="chip ${c}">${esc(l)}</span>`; };
const ocCell = oc => oc ? esc(oc) : '<span class="chip c-err">Falta OC</span>';
const progBar = (p, label) => `<div class="prog"><div class="track" role="img" aria-label="${esc(label || p + ' por ciento')}"><div class="bar ${p > 100 ? 'over' : ''}" style="width:${Math.min(p, 100)}%"></div></div><span>${p}%</span></div>`;
const opts = (arr, sel) => arr.map(v => `<option value="${esc(v)}" ${v === sel ? 'selected' : ''}>${esc(v)}</option>`).join('');
const dlist = (id, arr) => `<datalist id="${id}">${arr.map(v => `<option value="${esc(v)}">`).join('')}</datalist>`;

/* ============ API ============ */
async function api(payload) {
  if (!cfg.url) throw new Error('Falta configurar la URL de la Web App');
  if (!cfg.key) throw new Error('Falta la clave de acceso (botón Configurar)');
  let r;
  try {
    // Sin encabezados: el navegador lo manda como text/plain y Apps Script no exige "preflight" CORS
    r = await fetch(cfg.url, { method: 'POST', body: JSON.stringify({ ...payload, key: cfg.key }) });
  } catch (e) {
    throw new Error('No hay conexión con Apps Script. Revisa internet y que la implementación tenga acceso "Cualquier persona".');
  }
  let data;
  try { data = await r.json(); } catch (e) { throw new Error('Apps Script respondió algo inesperado. Revisa la URL /exec y que esté implementado como Aplicación web.'); }
  if (!data.ok) throw new Error(data.error || 'Error desconocido');
  return data;
}

function leerArchivo(file) {
  return new Promise((resolve, reject) => {
    if (file.size > MAX_MB * 1024 * 1024) return reject(new Error(`"${file.name}" pesa más de ${MAX_MB} MB`));
    const fr = new FileReader();
    fr.onload = () => resolve(String(fr.result).split(',')[1] || '');
    fr.onerror = () => reject(new Error('No se pudo leer ' + file.name));
    fr.readAsDataURL(file);
  });
}

async function subir(partida, tipo, file, extra = {}) {
  const base64 = await leerArchivo(file);
  return api({ action: 'subirArchivo', partida, tipo, nombre: file.name, mimeType: file.type || 'application/octet-stream', base64, ...extra });
}

/* ============ DATOS ============ */
function normalizar(d) {
  const ocs = (d.ocs || []).map(o => ({ ...o, OC: up(o.OC), PROVEEDOR: up(o.PROVEEDOR), PROYECTO: up(o.PROYECTO), RUBRO: up(o.RUBRO), total: num(o.TOTAL) }));
  const ocMap = {};
  ocs.forEach(o => { if (o.OC) ocMap[o.OC] = o; });
  const partidas = (d.partidas || []).map(p => {
    const prog = num(p.MONTO_PROGRAMADO), pag = num(p.MONTO_PAGADO), oc = up(p.OC);
    return {
      ...p, OC: oc, PROYECTO: up(p.PROYECTO), PROVEEDOR: up(p.PROVEEDOR),
      prog, pag, pend: r2(prog - pag),
      rubro: up(p.RUBRO) || (ocMap[oc] ? ocMap[oc].RUBRO : ''),
      ndocs: EXP_COLS.filter(c => p[c]).length
    };
  });
  const presupuesto = (d.presupuesto || []).map(r => ({ ...r, PROYECTO: up(r.PROYECTO), RUBRO: up(r.RUBRO), proyectado: num(r.PROYECTADO) }));
  const pagos = (d.pagos || []).map(p => ({ ...p, PROVEEDOR: up(p.PROVEEDOR), monto: num(p.MONTO_GLOBAL) }));
  return { partidas, ocs, ocMap, presupuesto, pagos, servidor: d.servidor || '', siguiente: d.siguientePartida || '' };
}

async function cargarTodo(silencioso = false) {
  try {
    const d = await api({ action: 'todo' });
    try { localStorage.setItem('pa_cache', JSON.stringify(d)); } catch (e) { /* sin espacio: no importa */ }
    state.data = normalizar(d);
    setConn(true, d.servidor);
    render();
    return true;
  } catch (e) {
    setConn(false);
    if (!state.data) $('#view').innerHTML = `<div class="card empty">${esc(e.message)}<br><br><button class="btn btn-primary" data-action="config">Configurar conexión</button></div>`;
    if (!silencioso) toast(e.message, 'err', 9000);
    return false;
  }
}

function setConn(ok, cuando) {
  $('#conn').innerHTML = ok
    ? `<b><span class="dot on"></span>Conectado a Drive</b>Actualizado ${esc(cuando ? cuando.slice(11) : '')} · cuenta 8018`
    : `<b><span class="dot off"></span>Sin conexión</b>Mostrando la última información guardada`;
}

const proyectosTodos = () => { const D = state.data; return uniq([...D.partidas.map(p => p.PROYECTO), ...D.ocs.map(o => o.PROYECTO), ...D.presupuesto.map(r => r.PROYECTO)]).sort(); };
const proveedoresTodos = () => { const D = state.data; return uniq([...D.partidas.map(p => p.PROVEEDOR), ...D.ocs.map(o => o.PROVEEDOR)]).sort(); };
const rubrosTodos = proy => { const D = state.data; return uniq([...D.presupuesto.filter(r => !proy || r.PROYECTO === proy).map(r => r.RUBRO), ...D.ocs.filter(o => !proy || o.PROYECTO === proy).map(o => o.RUBRO)]).sort(); };
function datosBanco(prov) {
  const ps = state.data.partidas.filter(p => p.PROVEEDOR === prov && (p.BANCO || p.CUENTA_DEPOSITO || p.CLABE_REFERENCIA))
    .sort((a, b) => num(b.PARTIDA) - num(a.PARTIDA));
  const p = ps[0];
  return p ? { banco: p.BANCO, cuenta: p.CUENTA_DEPOSITO, clabe: p.CLABE_REFERENCIA, forma: p.FORMA_PAGO, proyecto: p.PROYECTO } : null;
}

/* ============ RUTAS ============ */
function route() {
  const v = (location.hash.replace(/^#\/?/, '') || 'programacion').split('?')[0];
  state.view = ['programacion', 'pago', 'proveedores', 'flujo'].includes(v) ? v : 'programacion';
  $$('[data-nav]').forEach(a => { a.classList.toggle('active', a.dataset.nav === state.view); a.toggleAttribute('aria-current', a.dataset.nav === state.view); });
  render();
}

function render() {
  if (!state.data) return;
  const vistas = { programacion: vistaProgramacion, pago: vistaPago, proveedores: vistaProveedores, flujo: vistaFlujo };
  $('#view').innerHTML = vistas[state.view]();
  if (state.view === 'pago') recalcPago();
  document.title = { programacion: 'Programación', pago: 'Registrar pago', proveedores: 'Proveedores', flujo: 'Flujo por proyecto' }[state.view] + ' · Pagos y Flujo';
}

/* =====================================================================
 * 1. PROGRAMACIÓN SEMANAL
 * ===================================================================== */
function filasProgramacion() {
  const D = state.data;
  const delMes = D.partidas.filter(p => (p.FECHA_PROGRAMADA || '').slice(0, 7) === state.mes);
  const semanas = uniq(delMes.map(p => lunesISO(p.FECHA_PROGRAMADA))).sort();
  if (state.semana !== 'TODAS' && !semanas.includes(state.semana)) state.semana = 'TODAS';
  const proyectos = uniq(delMes.map(p => p.PROYECTO)).sort();
  if (state.proyecto !== 'TODOS' && !proyectos.includes(state.proyecto)) state.proyecto = 'TODOS';
  const q = state.buscar.toLowerCase();
  const filas = delMes
    .filter(p => state.semana === 'TODAS' || lunesISO(p.FECHA_PROGRAMADA) === state.semana)
    .filter(p => state.proyecto === 'TODOS' || p.PROYECTO === state.proyecto)
    .filter(p => !q || [p.PROVEEDOR, p.OC, p.CONCEPTO, p.PARTIDA, p.PROYECTO].join(' ').toLowerCase().includes(q))
    .sort((a, b) => (a.FECHA_PROGRAMADA || '').localeCompare(b.FECHA_PROGRAMADA || '') || num(a.PARTIDA) - num(b.PARTIDA));
  return { delMes, semanas, proyectos, filas };
}

function vistaProgramacion() {
  const { semanas, proyectos, filas } = filasProgramacion();
  const prog = sum(filas, p => p.prog), pag = sum(filas, p => p.pag), pend = sum(filas, p => Math.max(p.pend, 0));
  const nPend = filas.filter(p => p.pend > 0.009).length;
  const incompletos = filas.filter(p => p.ESTATUS === 'PAGADO - FALTAN DOCUMENTOS').length;
  const periodo = state.semana === 'TODAS' ? mesLabel(state.mes) : 'semana del ' + semanaLabel(state.semana);

  return `
  <header class="page-h">
    <div class="t">
      <span class="eyebrow">Cuentas por pagar · ${esc(periodo)}</span>
      <h1>Programación de pagos</h1>
    </div>
    <div class="actions">
      <label class="f" style="flex-direction:row;align-items:center;gap:8px">Mes
        <input type="month" value="${esc(state.mes)}" data-change="mes" style="width:auto">
      </label>
      <button class="btn" type="button" data-action="exportar-cxp">Exportar para CxP (Excel)</button>
      <button class="btn btn-primary" type="button" data-action="nueva-solicitud">+ Nueva solicitud</button>
    </div>
  </header>

  <section class="kpis" aria-label="Resumen">
    <div class="kpi"><span class="l">Programado</span><span class="v">${money(prog)}</span><span class="s">${pl(filas.length, 'solicitud', 'solicitudes')}</span></div>
    <div class="kpi"><span class="l">Pagado (con comprobante)</span><span class="v v-ok">${money(pag)}</span><span class="s">Se aplica al recibir el comprobante</span></div>
    <div class="kpi"><span class="l">Por pagar</span><span class="v v-warn">${money(pend)}</span><span class="s">${pl(nPend, 'partida pendiente', 'partidas pendientes')}</span></div>
    <div class="kpi"><span class="l">Expedientes incompletos</span><span class="v v-err">${incompletos}</span><span class="s">Pagadas sin factura, XML u OC</span></div>
  </section>

  <div class="info">
    <svg aria-hidden="true" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" style="flex:none"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/></svg>
    <span>Cada solicitud se captura una sola vez y crea su carpeta en Drive. De aquí se alimentan el histórico, los saldos por proveedor y el flujo de cada proyecto.</span>
  </div>

  <section class="card" aria-label="Solicitudes de pago">
    <div class="card-h">
      <div class="fchips" role="group" aria-label="Filtrar por semana">
        <button type="button" class="fchip ${state.semana === 'TODAS' ? 'on' : ''}" data-action="f-semana" data-v="TODAS">Todo el mes</button>
        ${semanas.map(s => `<button type="button" class="fchip ${state.semana === s ? 'on' : ''}" data-action="f-semana" data-v="${s}">${esc(semanaLabel(s))}</button>`).join('')}
      </div>
      <label class="search" style="flex:0 1 300px"><span class="sr">Buscar</span>
        <svg aria-hidden="true" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#5B6474" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>
        <input type="search" placeholder="Proveedor, OC, concepto o partida" value="${esc(state.buscar)}" data-input="buscar">
      </label>
    </div>
    <div class="card-h" style="padding-top:0">
      <div class="fchips" role="group" aria-label="Filtrar por proyecto">
        <button type="button" class="fchip ${state.proyecto === 'TODOS' ? 'on' : ''}" data-action="f-proyecto" data-v="TODOS">Todos los proyectos</button>
        ${proyectos.map(p => `<button type="button" class="fchip ${state.proyecto === p ? 'on' : ''}" data-action="f-proyecto" data-v="${esc(p)}">${esc(p)}</button>`).join('')}
      </div>
    </div>
    ${filas.length ? `
    <div class="tbl-wrap">
      <table class="tbl" style="min-width:980px">
        <thead><tr>
          <th scope="col">N°</th><th scope="col">Fecha req.</th><th scope="col">Proyecto</th><th scope="col">Proveedor / forma de pago</th>
          <th scope="col">OC / ref.</th><th scope="col">Concepto</th>
          <th scope="col" class="num">Programado</th><th scope="col" class="num">Pagado</th><th scope="col">Estatus</th><th scope="col">Expediente</th>
        </tr></thead>
        <tbody>
          ${filas.map(p => `<tr>
            <td class="strong">${esc(p.PARTIDA)}</td>
            <td class="nowrap">${fechaMx(p.FECHA_PROGRAMADA)}</td>
            <td>${esc(p.PROYECTO)}</td>
            <td><div class="strong">${esc(p.PROVEEDOR)}</div><div class="small">${esc(p.FORMA_PAGO)}</div></td>
            <td class="nowrap">${ocCell(p.OC)}</td>
            <td class="wrap">${esc(p.CONCEPTO)}</td>
            <td class="num strong">${money(p.prog)}</td>
            <td class="num ${p.pag ? 'v-ok' : 'muted'}">${p.pag ? money(p.pag) : '—'}</td>
            <td>${chipEstatus(p.ESTATUS)}</td>
            <td><button type="button" class="btn btn-sm" data-action="expediente" data-partida="${esc(p.PARTIDA)}" aria-label="Abrir expediente de la partida ${esc(p.PARTIDA)}, ${p.ndocs} de 6 documentos">${p.ndocs}/6</button></td>
          </tr>`).join('')}
        </tbody>
        <tfoot><tr>
          <td colspan="6" class="num">Total ${esc(state.proyecto === 'TODOS' ? '' : state.proyecto)}</td>
          <td class="num">${money(prog)}</td><td class="num">${money(pag)}</td><td colspan="2"></td>
        </tr></tfoot>
      </table>
    </div>` : `<div class="empty">No hay solicitudes en ${esc(mesLabel(state.mes))}${state.proyecto !== 'TODOS' ? ' para ' + esc(state.proyecto) : ''}. Usa <b>+ Nueva solicitud</b>.</div>`}
  </section>`;
}

/* ---------- Nueva solicitud ---------- */
function abrirNuevaSolicitud() {
  const provs = proveedoresTodos(), proys = proyectosTodos();
  abrirDlg('Nueva solicitud de pago', `
    <form id="formSol" style="display:flex;flex-direction:column;gap:16px">
      <div class="grid">
        <label class="f">Fecha programada <input type="date" name="fechaProgramada" required value="${hoy()}"></label>
        <label class="f">Proveedor / beneficiario
          <input name="proveedor" list="dlProv" required autocomplete="off" placeholder="Ej. DRAKKAR SYSTEMS & ELECTRONICS" data-change="sol-proveedor">
          ${dlist('dlProv', provs)}
        </label>
        <label class="f">Proyecto <input name="proyecto" list="dlProy" required placeholder="Ej. VERACRUZ" data-change="sol-proyecto">${dlist('dlProy', proys)}</label>
        <label class="f">Forma de pago <select name="formaPago">${opts(FORMAS_PAGO)}</select></label>
        <label class="f">OC / referencia
          <input name="oc" list="dlOC" placeholder="Ej. OC-VER-045 o SIN OC" data-change="sol-oc">
          <datalist id="dlOC"></datalist>
          <span class="hint" id="hintOC"></span>
        </label>
        <label class="f">Rubro del presupuesto (flujo) <input name="rubro" list="dlRubro" placeholder="Ej. TRANSMISOR EAS SAME">${dlist('dlRubro', rubrosTodos())}</label>
        <label class="f">Monto programado <input name="montoProgramado" inputmode="decimal" required placeholder="0.00" class="money"></label>
      </div>
      <label class="f">Concepto <input name="concepto" placeholder="Descripción del pago"></label>
      <div>
        <h3>Datos bancarios</h3>
        <div class="small" id="hintProv" style="margin:4px 0 12px">Si el proveedor ya tiene pagos anteriores, se llenan solos.</div>
        <div class="grid">
          <label class="f">Banco <input name="banco" list="dlBancos" placeholder="Ej. BBVA">${dlist('dlBancos', uniq(Object.values(BANCOS)).concat(['EFECTIVALE', 'AMEX']))}</label>
          <label class="f">Cuenta de depósito <input name="cuenta" inputmode="numeric" placeholder="Número de cuenta"></label>
          <label class="f">CLABE interbancaria / línea de captura
            <input name="clabe" inputmode="numeric" placeholder="18 dígitos o referencia" data-input="clabe">
            <span class="hint" id="hintClabe"></span>
          </label>
        </div>
      </div>
      <div class="grid">
        <label class="f">Orden de compra (opcional) <input type="file" name="fileOC" accept=".pdf,.jpg,.jpeg,.png,.xlsx,.docx"></label>
        <label class="f">Cotización (opcional) <input type="file" name="fileCOT" accept=".pdf,.jpg,.jpeg,.png,.xlsx,.docx"></label>
      </div>
      <div class="small">Siguiente número de partida: <b>${esc(state.data.siguiente || '—')}</b>. Se creará su carpeta en Drive dentro del mes de la fecha programada.</div>
    </form>`,
    `<button type="button" class="btn" data-action="cerrar">Cancelar</button>
     <button type="submit" form="formSol" class="btn btn-primary" id="btnSol">Crear partida y carpeta</button>`);
}

function solProveedor(input) {
  const f = $('#formSol');
  const prov = up(input.value);
  const ocs = state.data.ocs.filter(o => o.PROVEEDOR === prov && o.OC);
  $('#dlOC').innerHTML = ocs.map(o => `<option value="${esc(o.OC)}">${esc(o.CONCEPTO)}</option>`).join('');
  const b = datosBanco(prov);
  if (!b) { $('#hintProv').textContent = prov ? 'Proveedor nuevo: captura sus datos bancarios una vez.' : ''; return; }
  if (!f.banco.value) f.banco.value = b.banco;
  if (!f.cuenta.value) f.cuenta.value = b.cuenta;
  if (!f.clabe.value) { f.clabe.value = b.clabe; validarClabe(f.clabe); }
  if (!f.proyecto.value && b.proyecto) f.proyecto.value = b.proyecto;
  if (b.forma && FORMAS_PAGO.includes(b.forma)) f.formaPago.value = b.forma;
  $('#hintProv').textContent = 'Datos tomados del último pago a este proveedor. Revísalos.';
}

function solOC(input) {
  const f = $('#formSol');
  const o = state.data.ocMap[up(input.value)];
  const h = $('#hintOC');
  if (!o) { h.textContent = ''; return; }
  const pagado = sum(state.data.partidas.filter(p => p.OC === o.OC), p => p.pag);
  const comprometido = sum(state.data.partidas.filter(p => p.OC === o.OC), p => Math.max(p.pend, 0));
  h.style.color = 'var(--info)';
  h.textContent = `Total OC ${money(o.total)} · pagado ${money(pagado)} · ya programado ${money(comprometido)} · saldo libre ${money(o.total - pagado - comprometido)}`;
  if (!f.proyecto.value && o.PROYECTO) f.proyecto.value = o.PROYECTO;
  if (!f.rubro.value && o.RUBRO) f.rubro.value = o.RUBRO;
  if (!f.concepto.value && o.CONCEPTO) f.concepto.value = o.CONCEPTO;
}

function validarClabe(input) {
  const v = input.value.replace(/\s+/g, '');
  const h = $('#hintClabe'), f = input.form;
  if (/^\d{18}$/.test(v)) {
    if (clabeValida(v)) {
      const banco = BANCOS[v.slice(0, 3)];
      h.style.color = 'var(--ok)'; h.textContent = 'CLABE válida' + (banco ? ' · ' + banco : '');
      if (banco && f && !f.banco.value) f.banco.value = banco;
    } else { h.style.color = 'var(--err)'; h.textContent = 'Dígito verificador incorrecto. Revisa la CLABE.'; }
  } else if (/^\d{15,17}$/.test(v)) { h.style.color = 'var(--warn)'; h.textContent = `Llevas ${v.length} dígitos; una CLABE tiene 18.`; }
  else h.textContent = '';
}

async function guardarSolicitud(form) {
  const btn = $('#btnSol');
  const d = Object.fromEntries(new FormData(form));
  delete d.fileOC; delete d.fileCOT;
  if (!num(d.montoProgramado)) return toast('Captura el monto programado', 'warn');
  btn.disabled = true; btn.textContent = 'Creando carpeta…';
  try {
    const { partida } = await api({ action: 'crearPartida', ...d });
    const oc = form.fileOC.files[0], cot = form.fileCOT.files[0];
    if (oc) { btn.textContent = 'Subiendo OC…'; await subir(partida.PARTIDA, 'OC', oc); }
    if (cot) { btn.textContent = 'Subiendo cotización…'; await subir(partida.PARTIDA, 'COTIZACION', cot); }
    cerrarDlg();
    state.mes = partida.MES_CARPETA; state.semana = 'TODAS'; state.proyecto = 'TODOS';
    toast(`Partida ${partida.PARTIDA} creada con su carpeta en Drive`);
    await cargarTodo();
  } catch (e) {
    toast(e.message, 'err', 9000);
    btn.disabled = false; btn.textContent = 'Crear partida y carpeta';
  }
}

/* ---------- Exportar a Excel ---------- */
function cargarXLSX() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  return new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = XLSX_CDN; s.onload = () => res(window.XLSX); s.onerror = () => rej(new Error('No se pudo cargar el generador de Excel (revisa internet)'));
    document.head.appendChild(s);
  });
}

async function descargarExcel(nombre, hojas) {
  const XLSX = await cargarXLSX();
  const wb = XLSX.utils.book_new();
  hojas.forEach(h => {
    const ws = XLSX.utils.aoa_to_sheet(h.filas);
    if (h.anchos) ws['!cols'] = h.anchos.map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, h.nombre.slice(0, 31));
  });
  XLSX.writeFile(wb, nombre);
}

async function exportarCxP() {
  const { filas } = filasProgramacion();
  if (!filas.length) return toast('No hay partidas para exportar', 'warn');
  const titulo = `PROGRAMACIÓN DE PAGOS ${state.proyecto === 'TODOS' ? '' : 'PROYECTO ' + state.proyecto + ' '}${fechaMx(hoy())}`.replace(/\s+/g, ' ');
  const aoa = [[titulo], [],
    ['N° PARTIDA', 'FECHA DE PAGO REQUERIDA', 'PROYECTO', 'PROVEEDOR', 'FORMA DE PAGO', 'BANCO', 'CTA. DEPÓSITO', 'CLABE INTERBANCARIA / LÍNEA DE CAPTURA', 'PROGRAMADO', 'PAGADO', 'POR PAGAR', 'CONCEPTO', 'OC', 'ESTATUS']];
  filas.forEach(p => aoa.push([num(p.PARTIDA), fechaMx(p.FECHA_PROGRAMADA), p.PROYECTO, p.PROVEEDOR, p.FORMA_PAGO, p.BANCO, p.CUENTA_DEPOSITO, p.CLABE_REFERENCIA, p.prog, p.pag, r2(p.prog - p.pag), p.CONCEPTO, p.OC || 'SIN OC', (ESTATUS[p.ESTATUS] || [p.ESTATUS])[0]]));
  aoa.push([], ['', '', '', '', '', '', '', 'TOTAL', sum(filas, p => p.prog), sum(filas, p => p.pag), sum(filas, p => p.pend)]);
  try {
    await descargarExcel(`Programacion_pagos_${state.mes}${state.semana !== 'TODAS' ? '_sem_' + state.semana : ''}.xlsx`,
      [{ nombre: 'Programación', filas: aoa, anchos: [10, 14, 14, 34, 22, 14, 16, 26, 14, 14, 14, 50, 16, 20] }]);
  } catch (e) { toast(e.message, 'err'); }
}

/* ---------- Expediente (diálogo) ---------- */
function abrirExpediente(partida) {
  state.expediente = String(partida);
  const p = state.data.partidas.find(x => String(x.PARTIDA) === state.expediente);
  if (!p) return toast('No se encontró la partida', 'err');
  const slots = DOCS.map(d => {
    const url = p[d.col];
    const done = !!url;
    const links = done ? url.split('\n').map((u, i, a) => `<a href="${esc(u)}" target="_blank" rel="noopener">Ver archivo${a.length > 1 ? ' ' + (i + 1) : ''}</a>`).join(' · ') : '';
    let extra = '';
    if (d.tipo === 'FACTURA_XML' && p.FACTURA_UUID) extra = `<small>UUID ${esc(p.FACTURA_UUID)}<br>RFC ${esc(p.FACTURA_RFC)} · Total ${money(p.FACTURA_TOTAL)}</small>`;
    if (d.tipo === 'COMPLEMENTO' && p.COMPLEMENTO_UUID) extra = `<small>UUID ${esc(p.COMPLEMENTO_UUID)}</small>`;
    if (d.tipo === 'PAGO' && p.pag) extra = `<small>Pagado ${money(p.pag)} el ${fechaMx(p.FECHA_PAGO)}${p.ID_PAGO ? ' · ' + esc(p.ID_PAGO) : ''}</small>`;
    const camposPago = d.tipo === 'PAGO' ? `
      <div class="two">
        <label class="sr" for="mp">Monto pagado</label><input id="mp" data-monto inputmode="decimal" class="money" placeholder="Monto" value="${esc(p.pag || p.prog)}">
        <label class="sr" for="fp">Fecha de pago</label><input id="fp" data-fecha type="date" value="${esc(p.FECHA_PAGO || hoy())}">
      </div>
      <small>Pago único de esta partida. Para un pago que cubre varias partidas o parcialidades usa <a href="#/pago" data-action="cerrar">Registrar pago</a>.</small>` : '';
    return `<div class="slot ${done ? 'done' : ''}" data-tipo="${d.tipo}">
      <div class="slot-h"><span>${d.label}</span>${done ? '<span class="chip c-ok">Cargado</span>' : '<span class="chip c-warn">Pendiente</span>'}</div>
      ${done ? `<small>${links}</small>` : ''}${extra}${camposPago}
      <label class="sr" for="f-${d.tipo}">Archivo ${d.label}</label>
      <input id="f-${d.tipo}" type="file" ${d.accept ? `accept="${d.accept}"` : ''} ${d.tipo === 'OTRO' ? 'multiple' : ''}>
      <button class="btn btn-sm ${done ? '' : 'btn-primary'}" type="button" data-action="subir-doc">${done ? (d.tipo === 'OTRO' ? 'Agregar' : 'Reemplazar') : 'Subir'}</button>
    </div>`;
  }).join('');

  abrirDlg(`Expediente · partida ${esc(p.PARTIDA)}`, `
    <div style="display:flex;flex-wrap:wrap;justify-content:space-between;gap:12px;align-items:flex-start">
      <div style="display:flex;flex-direction:column;gap:6px;min-width:0">
        <div style="font-size:17px;font-weight:700">${esc(p.PARTIDA)}. ${esc(p.PROVEEDOR)} ${money(p.pag || p.prog)}</div>
        <div class="meta">
          <span>Proyecto: <b>${esc(p.PROYECTO)}</b></span><span>OC: <b>${esc(p.OC || '—')}</b></span><span>Rubro: <b>${esc(p.rubro || '—')}</b></span>
          <span>Programado: <b>${money(p.prog)}</b> para ${fechaMx(p.FECHA_PROGRAMADA)}</span><span>Pendiente: <b>${money(p.pend)}</b></span>
          <span>Forma: ${esc(p.FORMA_PAGO)}</span><span>Banco: ${esc(p.BANCO || '—')}</span><span>Cuenta: ${esc(p.CUENTA_DEPOSITO || '—')}</span><span>CLABE / ref.: ${esc(p.CLABE_REFERENCIA || '—')}</span>
        </div>
        ${p.CONCEPTO ? `<div class="sub">${esc(p.CONCEPTO)}</div>` : ''}
      </div>
      <div class="actions">${chipEstatus(p.ESTATUS)}<a class="btn btn-sm" href="${esc(p.CARPETA_URL)}" target="_blank" rel="noopener">Abrir carpeta en Drive</a></div>
    </div>
    <div class="slots">${slots}</div>`,
    `<button type="button" class="btn" data-action="cerrar">Cerrar</button>`);
}

async function subirDoc(btn) {
  const slot = btn.closest('.slot'), tipo = slot.dataset.tipo, partida = state.expediente;
  const files = [...slot.querySelector('input[type=file]').files];
  if (!files.length) return toast('Selecciona un archivo primero', 'warn');
  const extra = {};
  if (tipo === 'PAGO') {
    extra.montoPagado = slot.querySelector('[data-monto]').value;
    extra.fechaPago = slot.querySelector('[data-fecha]').value;
    if (!num(extra.montoPagado) || !extra.fechaPago) return toast('Captura monto y fecha del pago', 'warn');
  }
  btn.disabled = true; btn.textContent = 'Subiendo…';
  try {
    let res;
    for (const f of files) res = await subir(partida, tipo, f, extra);
    await cargarTodo(true);
    abrirExpediente(partida);
    if (res.aviso) toast('Guardado. Atención: ' + res.aviso, 'warn', 9000);
    else toast(tipo === 'PAGO' ? 'Pago registrado; la carpeta se renombró con el monto y la fecha reales.' : 'Archivo guardado en Drive');
  } catch (e) {
    toast(e.message, 'err', 9000);
    btn.disabled = false; btn.textContent = 'Subir';
  }
}

/* =====================================================================
 * 2. REGISTRAR PAGO (monto global → varias partidas)
 * ===================================================================== */
function vistaPago() {
  const D = state.data, P = state.pago;
  const pendientes = D.partidas.filter(p => p.pend > 0.009);
  const provs = uniq(pendientes.map(p => p.PROVEEDOR)).sort();
  if (P.proveedor && !provs.includes(P.proveedor)) P.proveedor = '';
  const lista = pendientes.filter(p => p.PROVEEDOR === P.proveedor)
    .sort((a, b) => (a.FECHA_PROGRAMADA || '').localeCompare(b.FECHA_PROGRAMADA || '') || num(a.PARTIDA) - num(b.PARTIDA));
  const sel = P.seleccion || {};

  return `
  <header class="page-h">
    <div class="t">
      <span class="eyebrow">Cuentas por pagar · Registrar comprobante</span>
      <h1>Registrar pago</h1>
      <span class="sub">Un solo comprobante puede cubrir varias partidas u OC, incluso de proyectos distintos, y también parcialidades.</span>
    </div>
  </header>

  <div class="split">
    <div class="main">
      <section class="card card-p" aria-label="Datos del comprobante" style="display:flex;flex-direction:column;gap:16px">
        <h2>1. Datos del comprobante</h2>
        <div class="grid">
          <label class="f">Proveedor
            <select id="pgProv" data-change="pg-proveedor">
              <option value="">Selecciona un proveedor…</option>${opts(provs, P.proveedor)}
            </select>
          </label>
          <label class="f">Cuenta de origen <select id="pgCuenta">${opts(CUENTAS_ORIGEN, P.cuenta)}</select></label>
          <label class="f">Fecha de pago <input id="pgFecha" type="date" value="${esc(P.fecha || hoy())}"></label>
          <label class="f">Clave de rastreo / folio <input id="pgRef" value="${esc(P.ref || '')}" placeholder="Del comprobante SPEI"></label>
        </div>
        <label class="f" style="max-width:380px">Monto global pagado (según comprobante)
          <input id="pgGlobal" class="money big" inputmode="decimal" data-input="pg-global" value="${esc(P.global || '')}" placeholder="0.00">
        </label>
        ${provs.length ? '' : '<div class="small">No hay partidas pendientes de pago. Crea solicitudes en Programación semanal.</div>'}
      </section>

      <section class="card" aria-label="Distribución del pago">
        <div class="card-h"><div class="t">
          <h2>2. ¿A qué partidas se aplica este pago?</h2>
          <span class="sub">Marca las partidas y escribe cuánto del comprobante corresponde a cada una.</span>
        </div></div>
        ${P.proveedor ? (lista.length ? `
        <div class="tbl-wrap">
          <table class="tbl" style="min-width:760px">
            <thead><tr>
              <th scope="col"><label class="sr" for="pgAll">Seleccionar todas</label><input type="checkbox" id="pgAll" data-input="pg-all"></th>
              <th scope="col">N°</th><th scope="col">OC / proyecto</th><th scope="col">Concepto</th>
              <th scope="col" class="num">Pendiente</th>
              <th scope="col" class="num">Monto aplicado</th><th scope="col" class="num">Saldo después</th>
            </tr></thead>
            <tbody>
              ${lista.map(p => { const s = sel[p.PARTIDA]; return `<tr>
                <td><label class="sr" for="chk-${esc(p.PARTIDA)}">Aplicar a partida ${esc(p.PARTIDA)}</label>
                  <input type="checkbox" id="chk-${esc(p.PARTIDA)}" data-pg-check data-partida="${esc(p.PARTIDA)}" data-input="pg-check" ${s ? 'checked' : ''}></td>
                <td class="strong">${esc(p.PARTIDA)}</td>
                <td class="nowrap"><div>${ocCell(p.OC)}</div><div class="small">${esc(p.PROYECTO)} · ${fechaMx(p.FECHA_PROGRAMADA)}</div></td>
                <td class="wrap">${esc(p.CONCEPTO)}</td>
                <td class="num"><div class="strong">${money(p.pend)}</div>${p.pag ? `<div class="small">de ${money(p.prog)}</div>` : ''}</td>
                <td class="num" style="width:160px">
                  <label class="sr" for="mon-${esc(p.PARTIDA)}">Monto aplicado a partida ${esc(p.PARTIDA)}</label>
                  <input id="mon-${esc(p.PARTIDA)}" class="money" inputmode="decimal" data-pg-monto data-partida="${esc(p.PARTIDA)}" data-pend="${p.pend}" data-input="pg-monto" value="${esc(s || '')}" ${s ? '' : 'disabled'}>
                </td>
                <td class="num strong" data-pg-saldo="${esc(p.PARTIDA)}">${money(p.pend)}</td>
              </tr>`; }).join('')}
            </tbody>
            <tfoot><tr><td colspan="5" class="num">Total aplicado</td><td class="num" id="pgSuma">$0.00</td><td></td></tr></tfoot>
          </table>
        </div>` : '<div class="empty">Este proveedor no tiene partidas pendientes.</div>')
        : '<div class="empty">Selecciona un proveedor para ver sus partidas pendientes.</div>'}
        <div style="padding:16px 20px 20px"><div id="pgBal" class="bal bal-info" role="status">Selecciona partidas y captura el monto del comprobante.</div></div>
      </section>
    </div>

    <aside class="aside" aria-label="Comprobante y guardado">
      <section class="card card-p" style="display:flex;flex-direction:column;gap:12px">
        <h2>3. Comprobante</h2>
        <label class="f">Archivo del comprobante (PDF o imagen)
          <input type="file" id="pgFile" accept=".pdf,.jpg,.jpeg,.png">
        </label>
        <div class="small">Se guarda una copia en la carpeta de <b>cada</b> partida, para que cada expediente quede completo. Las facturas se suben después desde el expediente de cada partida.</div>
      </section>
      <section class="card card-p" style="display:flex;flex-direction:column;gap:10px">
        <h2>Al guardar se actualiza</h2>
        <ul style="margin:0;padding-left:18px;font-size:13px;color:var(--ink2);line-height:1.8">
          <li>Monto pagado y estatus de cada partida</li>
          <li>Nombre de cada carpeta en Drive (monto y fecha reales)</li>
          <li>Saldos del proveedor por OC</li>
          <li>Flujo de cada proyecto involucrado</li>
          <li>Registro en la pestaña PAGOS de la hoja</li>
        </ul>
        <div class="actions" style="margin-top:6px">
          <button type="button" class="btn" style="flex:1" data-action="pg-limpiar">Limpiar</button>
          <button type="button" class="btn btn-primary" style="flex:2" id="pgSave" data-action="pg-guardar" disabled>Guardar pago</button>
        </div>
      </section>
    </aside>
  </div>`;
}

function guardarCamposPago() {
  const P = state.pago;
  if (!$('#pgFecha')) return;
  P.fecha = $('#pgFecha').value; P.cuenta = $('#pgCuenta').value; P.ref = $('#pgRef').value; P.global = $('#pgGlobal').value;
  P.seleccion = {};
  $$('[data-pg-check]').forEach(c => { if (c.checked) P.seleccion[c.dataset.partida] = $(`#mon-${CSS.escape(c.dataset.partida)}`).value || ' '; });
}

function recalcPago() {
  const bal = $('#pgBal'), save = $('#pgSave');
  if (!bal) return;
  let suma = 0, excede = false, marcadas = 0;
  $$('[data-pg-monto]').forEach(inp => {
    const id = inp.dataset.partida;
    const chk = $(`#chk-${CSS.escape(id)}`);
    const pend = num(inp.dataset.pend);
    inp.disabled = !chk.checked;
    const v = chk.checked ? num(inp.value) : 0;
    if (chk.checked) marcadas++;
    suma += v;
    if (v - pend > 0.009) excede = true;
    const saldo = r2(pend - v);
    const cell = $(`[data-pg-saldo="${CSS.escape(id)}"]`);
    cell.textContent = money(saldo);
    cell.classList.toggle('neg', saldo < -0.009);
  });
  suma = r2(suma);
  if ($('#pgSuma')) $('#pgSuma').textContent = money(suma);
  const g = $('#pgGlobal');
  if (!state.pago.globalTocado) g.value = suma ? suma.toFixed(2) : '';
  const global = num(g.value), diff = r2(global - suma);
  let cls = 'bal-info', msg = 'Selecciona partidas y captura el monto del comprobante.', ok = false;
  if (marcadas && global > 0) {
    if (Math.abs(diff) < 0.01) {
      ok = true;
      cls = excede ? 'bal-warn' : 'bal-ok';
      msg = excede
        ? `Cuadra con el comprobante (${money(global)}), pero una partida recibe más de lo pendiente: quedará como pago de más.`
        : `Cuadra: ${money(global)} distribuidos entre ${marcadas} partida${marcadas > 1 ? 's' : ''}.`;
    } else if (diff > 0) { cls = 'bal-warn'; msg = `Faltan ${money(diff)} por aplicar a alguna partida.`; }
    else { cls = 'bal-err'; msg = `Aplicaste ${money(-diff)} más de lo que dice el comprobante.`; }
  }
  bal.className = 'bal ' + cls; bal.textContent = msg;
  save.disabled = !ok;
  const all = $('#pgAll');
  if (all) { const cs = $$('[data-pg-check]'); all.checked = cs.length > 0 && cs.every(c => c.checked); }
}

async function guardarPago() {
  const P = state.pago;
  const file = $('#pgFile').files[0];
  if (!file) return toast('Adjunta el archivo del comprobante', 'warn');
  const aplicaciones = $$('[data-pg-check]').filter(c => c.checked)
    .map(c => ({ partida: c.dataset.partida, monto: num($(`#mon-${CSS.escape(c.dataset.partida)}`).value) }))
    .filter(a => a.monto > 0);
  const btn = $('#pgSave');
  btn.disabled = true; btn.textContent = 'Guardando…';
  try {
    const base64 = await leerArchivo(file);
    const res = await api({
      action: 'registrarPagoGlobal', proveedor: P.proveedor, fechaPago: $('#pgFecha').value, cuentaOrigen: $('#pgCuenta').value,
      referencia: $('#pgRef').value, montoGlobal: $('#pgGlobal').value, aplicaciones,
      archivo: { nombre: file.name, mimeType: file.type || 'application/octet-stream', base64 }
    });
    state.pago = { proveedor: '', globalTocado: false };
    toast(`Pago ${res.idPago} registrado en ${aplicaciones.length} partida(s).` + (res.aviso ? ' ' + res.aviso : ''), res.aviso ? 'warn' : 'ok', 8000);
    await cargarTodo();
  } catch (e) {
    toast(e.message, 'err', 9000);
    btn.disabled = false; btn.textContent = 'Guardar pago';
  }
}

/* =====================================================================
 * 3. PROVEEDORES
 * ===================================================================== */
function ocsConSaldo(prov) {
  const D = state.data;
  return D.ocs.filter(o => !prov || o.PROVEEDOR === prov).map(o => {
    const ps = o.OC ? D.partidas.filter(p => p.OC === o.OC) : [];
    const pag = sum(ps, p => p.pag);
    return { ...o, pag, porPagar: r2(o.total - pag), pct: pct(pag, o.total) };
  });
}

function vistaProveedores() {
  const D = state.data;
  const nombres = proveedoresTodos();
  if (!state.provSel || !nombres.includes(state.provSel)) state.provSel = nombres[0] || '';
  const prov = state.provSel;
  if (!prov) return `<header class="page-h"><div class="t"><h1>Proveedores</h1></div>
    <button class="btn btn-primary" data-action="oc-nueva">+ Agregar cotización / OC</button></header>
    <div class="card empty">Aún no hay proveedores. Se agregan al crear solicitudes o cotizaciones.</div>`;

  const ocs = ocsConSaldo(prov);
  const parts = D.partidas.filter(p => p.PROVEEDOR === prov);
  const pend = parts.filter(p => p.pend > 0.009);
  const b = datosBanco(prov) || {};
  const comprometido = sum(ocs, o => o.total), pagado = sum(ocs, o => o.pag);
  const sinOCreg = parts.filter(p => p.pag > 0 && (!p.OC || !D.ocMap[p.OC]));
  const faltantes = ocs.filter(o => !o.COTIZACION).length + ocs.filter(o => !o.OC).length + parts.filter(p => !p.OC).length;

  const historial = [
    ...D.pagos.filter(p => p.PROVEEDOR === prov).map(p => ({ fecha: p.FECHA_PAGO, id: p.ID_PAGO, monto: p.monto, detalle: p.DETALLE, url: p.COMPROBANTE_URL, cuenta: p.CUENTA_ORIGEN })),
    ...parts.filter(p => p.pag > 0 && !p.ID_PAGO).map(p => ({ fecha: p.FECHA_PAGO, id: 'Partida ' + p.PARTIDA, monto: p.pag, detalle: `${p.PARTIDA} (${p.OC || 'SIN OC'})`, url: (p.PAGO_URL || '').split('\n')[0], cuenta: p.FORMA_PAGO }))
  ].sort((a, c) => (c.fecha || '').localeCompare(a.fecha || ''));

  const optsProv = nombres.map(n => {
    const k = D.partidas.filter(p => p.PROVEEDOR === n && p.pend > 0.009).length;
    return `<option value="${esc(n)}" ${n === prov ? 'selected' : ''}>${esc(n)}${k ? ` (${k} por pagar)` : ''}</option>`;
  }).join('');

  return `
  <div style="display:flex;flex-direction:column;gap:24px">
      <div class="card card-p" style="display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end">
        <label class="f" style="flex:1 1 320px">Proveedor
          <select data-change="prov-sel">${optsProv}</select>
        </label>
        <div class="small" style="flex:1 1 260px;padding-bottom:12px">${pl(nombres.length, 'proveedor registrado', 'proveedores registrados')}. Escribe las primeras letras sobre la lista para buscar.</div>
      </div>

      <header class="page-h">
        <div class="t">
          <span class="eyebrow">Proveedor · cotizaciones y saldos</span>
          <h1>${esc(prov)}</h1>
          <span class="sub">Banco: <b>${esc(b.banco || '—')}</b> · Cuenta: ${esc(b.cuenta || '—')} · CLABE / ref.: ${esc(b.clabe || '—')}</span>
        </div>
        <div class="actions">
          <button class="btn" type="button" data-action="exportar-prov">Estado de cuenta (Excel)</button>
          <button class="btn btn-primary" type="button" data-action="oc-nueva">+ Agregar cotización / OC</button>
        </div>
      </header>

      <section class="kpis" aria-label="Resumen del proveedor">
        <div class="kpi"><span class="l">Total comprometido (OC)</span><span class="v">${money(comprometido)}</span><span class="s">${pl(ocs.length, 'cotización / OC', 'cotizaciones / OC')}</span></div>
        <div class="kpi"><span class="l">Pagado</span><span class="v v-ok">${money(pagado)}</span><span class="s">Calculado con los pagos registrados</span></div>
        <div class="kpi"><span class="l">Por pagar</span><span class="v v-warn">${money(comprometido - pagado)}</span><span class="s">${pl(pend.length, 'partida pendiente', 'partidas pendientes')}</span></div>
        <div class="kpi"><span class="l">Datos faltantes</span><span class="v v-err">${faltantes}</span><span class="s">Cotización u OC sin registrar</span></div>
      </section>

      <section class="card" aria-label="Cotizaciones y saldos">
        <div class="card-h"><div class="t"><h2>Cotizaciones y saldos</h2>
          <span class="sub">El pagado se suma solo con las partidas que tienen esa OC. Ya no se copia del CxP.</span></div></div>
        ${ocs.length ? `<div class="tbl-wrap"><table class="tbl" style="min-width:880px">
          <thead><tr><th scope="col">Cotización / OC</th><th scope="col">Concepto</th><th scope="col">Proyecto / rubro</th>
            <th scope="col" class="num">Total</th><th scope="col" class="num">Pagado</th><th scope="col" class="num">Por pagar</th><th scope="col">Avance</th><th scope="col"><span class="sr">Acciones</span></th></tr></thead>
          <tbody>${ocs.map(o => `<tr>
            <td class="nowrap"><div>${o.COTIZACION ? esc(o.COTIZACION) : '<span class="chip c-err">Falta cotización</span>'}</div>
              <div style="margin-top:4px">${o.OC ? `<span class="small">${esc(o.OC)}</span>` : '<span class="chip c-err">Falta OC</span>'}</div></td>
            <td class="wrap">${esc(o.CONCEPTO)}</td>
            <td><div>${esc(o.PROYECTO)}</div><div class="small">${esc(o.RUBRO || 'Sin rubro')}</div></td>
            <td class="num">${money(o.total)}</td><td class="num">${money(o.pag)}</td>
            <td class="num strong ${o.porPagar < -0.009 ? 'neg' : ''}">${money(o.porPagar)}</td>
            <td>${progBar(o.pct, 'Pagado ' + o.pct + ' por ciento')}</td>
            <td><button type="button" class="btn-link" data-action="oc-editar" data-id="${esc(o.ID)}">Editar</button></td>
          </tr>`).join('')}</tbody>
          <tfoot><tr><td colspan="3" class="num">Totales</td><td class="num">${money(comprometido)}</td><td class="num">${money(pagado)}</td><td class="num">${money(comprometido - pagado)}</td><td colspan="2"></td></tr></tfoot>
        </table></div>` : '<div class="empty">Sin cotizaciones / OC registradas para este proveedor. Usa <b>+ Agregar cotización / OC</b>.</div>'}
        ${sinOCreg.length ? `<div style="padding:0 20px 18px" class="small">Hay ${sinOCreg.length} pago(s) a este proveedor con OC que no está registrada aquí (${money(sum(sinOCreg, p => p.pag))}). Regístrala para que cuente en los saldos.</div>` : ''}
      </section>

      <section class="card" aria-label="Partidas pendientes">
        <div class="card-h"><div class="t"><h2>Partidas programadas pendientes de pago</h2></div>
          ${pend.length ? `<button class="btn btn-primary btn-sm" type="button" data-action="prov-pagar" data-v="${esc(prov)}">Registrar pago de este proveedor</button>` : ''}</div>
        ${pend.length ? `<div class="tbl-wrap"><table class="tbl">
          <thead><tr><th scope="col">N°</th><th scope="col">Fecha prog.</th><th scope="col">OC</th><th scope="col">Proyecto</th><th scope="col">Concepto</th><th scope="col" class="num">Pendiente</th><th scope="col">Estatus</th></tr></thead>
          <tbody>${pend.map(p => `<tr><td class="strong">${esc(p.PARTIDA)}</td><td>${fechaMx(p.FECHA_PROGRAMADA)}</td><td>${ocCell(p.OC)}</td><td>${esc(p.PROYECTO)}</td><td class="wrap">${esc(p.CONCEPTO)}</td><td class="num strong">${money(p.pend)}</td><td>${chipEstatus(p.ESTATUS)}</td></tr>`).join('')}</tbody>
        </table></div>` : '<div class="empty">Sin partidas pendientes.</div>'}
      </section>

      <section class="card" aria-label="Historial de pagos">
        <div class="card-h"><div class="t"><h2>Historial de pagos</h2></div></div>
        ${historial.length ? `<div class="tbl-wrap"><table class="tbl">
          <thead><tr><th scope="col">Fecha de pago</th><th scope="col">Pago</th><th scope="col">Cuenta / forma</th><th scope="col" class="num">Monto</th><th scope="col">Aplicado a</th><th scope="col">Comprobante</th></tr></thead>
          <tbody>${historial.map(h => `<tr><td>${fechaMx(h.fecha)}</td><td class="strong">${esc(h.id)}</td><td>${esc(h.cuenta || '')}</td><td class="num strong">${money(h.monto)}</td><td class="wrap">${esc(h.detalle)}</td>
            <td>${h.url ? `<a href="${esc(h.url)}" target="_blank" rel="noopener">Ver</a>` : '—'}</td></tr>`).join('')}</tbody>
        </table></div>` : '<div class="empty">Sin pagos registrados.</div>'}
      </section>
  </div>`;
}

function abrirOC(id) {
  const o = id ? state.data.ocs.find(x => x.ID === id) : null;
  const provs = proveedoresTodos(), proys = proyectosTodos();
  abrirDlg(o ? 'Editar cotización / OC' : 'Agregar cotización / OC', `
    <form id="formOC" style="display:flex;flex-direction:column;gap:14px">
      <input type="hidden" name="id" value="${esc(o ? o.ID : '')}">
      <div class="grid">
        <label class="f">Proveedor <input name="proveedor" list="dlProvOC" required value="${esc(o ? o.PROVEEDOR : state.provSel)}">${dlist('dlProvOC', provs)}</label>
        <label class="f">N° de cotización <input name="cotizacion" value="${esc(o ? o.COTIZACION : '')}" placeholder="Ej. CO_MDRK233-DP-0826_V1"></label>
        <label class="f">Orden de compra <input name="oc" value="${esc(o ? o.OC : '')}" placeholder="Ej. OC-VER-045"></label>
        <label class="f">Proyecto <input name="proyecto" list="dlProyOC" value="${esc(o ? o.PROYECTO : '')}">${dlist('dlProyOC', proys)}</label>
        <label class="f">Rubro del presupuesto <input name="rubro" list="dlRubOC" value="${esc(o ? o.RUBRO : '')}">${dlist('dlRubOC', rubrosTodos())}</label>
        <label class="f">Total de la cotización / OC <input name="total" class="money" inputmode="decimal" required value="${esc(o ? o.total : '')}" placeholder="0.00"></label>
      </div>
      <label class="f">Concepto <input name="concepto" value="${esc(o ? o.CONCEPTO : '')}"></label>
      <label class="f">Observaciones <textarea name="observaciones">${esc(o ? o.OBSERVACIONES : '')}</textarea></label>
    </form>`,
    `<button type="button" class="btn" data-action="cerrar">Cancelar</button><button type="submit" form="formOC" class="btn btn-primary">Guardar</button>`);
}

async function exportarProveedor() {
  const prov = state.provSel, D = state.data;
  const ocs = ocsConSaldo(prov);
  const aoa = [[`ESTADO DE CUENTA ${prov} · ${fechaMx(hoy())}`], [], ['COTIZACIÓN', 'CONCEPTO', 'OC', 'PROYECTO', 'RUBRO', 'TOTAL', 'PAGADO', 'POR PAGAR', '% AVANCE']];
  ocs.forEach(o => aoa.push([o.COTIZACION || 'FALTA', o.CONCEPTO, o.OC || 'FALTA', o.PROYECTO, o.RUBRO, o.total, o.pag, o.porPagar, o.pct / 100]));
  aoa.push([], ['', '', '', '', 'TOTAL', sum(ocs, o => o.total), sum(ocs, o => o.pag), sum(ocs, o => o.porPagar)]);
  const pagos = [['N° PARTIDA', 'FECHA PAGO', 'OC', 'PROYECTO', 'CONCEPTO', 'PROGRAMADO', 'PAGADO', 'ID PAGO']];
  D.partidas.filter(p => p.PROVEEDOR === prov).forEach(p => pagos.push([num(p.PARTIDA), fechaMx(p.FECHA_PAGO), p.OC, p.PROYECTO, p.CONCEPTO, p.prog, p.pag, p.ID_PAGO]));
  try {
    await descargarExcel(`Estado_cuenta_${prov.replace(/[^A-Z0-9]+/g, '_')}.xlsx`, [
      { nombre: 'Cotizaciones y saldos', filas: aoa, anchos: [26, 50, 16, 14, 22, 14, 14, 14, 10] },
      { nombre: 'Partidas', filas: pagos, anchos: [10, 12, 16, 14, 50, 14, 14, 16] }]);
  } catch (e) { toast(e.message, 'err'); }
}

/* =====================================================================
 * 4. FLUJO POR PROYECTO
 * ===================================================================== */
function datosFlujo(P) {
  const D = state.data;
  const pres = D.presupuesto.filter(r => r.PROYECTO === P);
  const ocs = D.ocs.filter(o => o.PROYECTO === P);
  const parts = D.partidas.filter(p => p.PROYECTO === P);
  const R = x => x || 'SIN RUBRO';
  const rubros = uniq([...pres.map(r => R(r.RUBRO)), ...ocs.map(o => R(o.RUBRO)), ...parts.map(p => R(p.rubro))]);
  const filas = rubros.map(rub => {
    const pr = pres.filter(x => R(x.RUBRO) === rub);
    const oo = ocs.filter(o => R(o.RUBRO) === rub);
    const pp = parts.filter(p => R(p.rubro) === rub);
    const proyectado = sum(pr, x => x.proyectado);
    const estimado = !oo.length && proyectado > 0;
    const cotizado = estimado ? proyectado : sum(oo, o => o.total);
    const pagado = sum(pp, p => p.pag);
    const proveedor = pr.map(x => x.PROVEEDOR).filter(Boolean)[0] || uniq([...oo.map(o => o.PROVEEDOR), ...pp.map(p => p.PROVEEDOR)]).slice(0, 2).join(', ');
    return { rubro: rub, proveedor, proyectado, cotizado, estimado, pagado, porPagar: r2(cotizado - pagado), varianza: r2(proyectado - cotizado),
      pct: pct(pagado, cotizado), sinPres: proyectado <= 0, presId: pr[0] ? pr[0].ID : '', partidas: pp };
  });
  return { filas, parts };
}

function calendario(P, vista) {
  const { filas, parts } = datosFlujo(P);
  let cols, clave, etiqueta;
  if (vista === 'mes') {
    cols = uniq(parts.map(p => (p.FECHA_PROGRAMADA || '').slice(0, 7))).sort();
    clave = p => (p.FECHA_PROGRAMADA || '').slice(0, 7); etiqueta = mesCorto;
  } else {
    cols = uniq(parts.filter(p => (p.FECHA_PROGRAMADA || '').slice(0, 7) === state.mes).map(p => lunesISO(p.FECHA_PROGRAMADA))).sort();
    clave = p => lunesISO(p.FECHA_PROGRAMADA); etiqueta = semanaLabel;
  }
  const filasCal = filas.map(f => {
    const vals = cols.map(c => sum(f.partidas.filter(p => clave(p) === c), p => p.prog));
    return { rubro: f.rubro, vals, total: sum(vals, v => v) };
  }).filter(f => f.total > 0);
  const totCols = cols.map((c, i) => sum(filasCal, f => f.vals[i]));
  return { cols, etiquetas: cols.map(etiqueta), filasCal, totCols, grand: sum(totCols, v => v) };
}

function vistaFlujo() {
  const proys = proyectosTodos();
  if (!state.flujoProy || !proys.includes(state.flujoProy)) state.flujoProy = proys[0] || '';
  const P = state.flujoProy;
  if (!P) return `<header class="page-h"><div class="t"><h1>Flujo por proyecto</h1></div>
    <button class="btn btn-primary" data-action="rubro-nuevo">+ Agregar rubro al presupuesto</button></header>
    <div class="card empty">Aún no hay proyectos. Agrega rubros de presupuesto o crea solicitudes.</div>`;

  const { filas } = datosFlujo(P);
  const T = { proy: sum(filas, f => f.proyectado), cot: sum(filas, f => f.cotizado), pag: sum(filas, f => f.pagado) };
  T.pp = r2(T.cot - T.pag); T.vari = r2(T.proy - T.cot);
  const cal = calendario(P, state.flujoVista);

  return `
  <header class="page-h">
    <div class="t">
      <span class="eyebrow">Flujo de dirección · montos con IVA</span>
      <h1>Flujo del proyecto ${esc(P)}</h1>
      <span class="sub">Se actualiza solo con cada solicitud, OC y pago registrado.</span>
    </div>
    <div class="actions">
      <button class="btn" type="button" data-action="exportar-flujo">Exportar flujo (Excel)</button>
      <button class="btn btn-primary" type="button" data-action="rubro-nuevo">+ Agregar rubro</button>
    </div>
  </header>

  <div class="fchips" role="group" aria-label="Proyecto">
    ${proys.map(p => `<button type="button" class="fchip ${p === P ? 'on' : ''}" data-action="flujo-proy" data-v="${esc(p)}">${esc(p)}</button>`).join('')}
  </div>

  <section class="kpis" aria-label="Resumen del proyecto">
    <div class="kpi"><span class="l">Proyectado</span><span class="v">${money(T.proy)}</span></div>
    <div class="kpi"><span class="l">Cotizado</span><span class="v">${money(T.cot)}</span></div>
    <div class="kpi"><span class="l">Pagado</span><span class="v v-ok">${money(T.pag)}</span></div>
    <div class="kpi"><span class="l">Por pagar</span><span class="v v-warn">${money(T.pp)}</span></div>
    <div class="kpi"><span class="l">Avance de pago</span><span class="v">${pct(T.pag, T.cot)}%</span><span class="s">Pagado entre cotizado</span></div>
  </section>

  <section class="card" aria-label="Presupuesto por rubro">
    <div class="card-h"><div class="t"><h2>Presupuesto vs. real por rubro</h2>
      <span class="sub">Varianza = proyectado − cotizado. Si un rubro no tiene OC, se toma el proyectado como cotizado (estimado).</span></div></div>
    ${filas.length ? `<div class="tbl-wrap"><table class="tbl" style="min-width:1100px">
      <thead><tr><th scope="col">Rubro</th><th scope="col">Proveedor</th><th scope="col" class="num">Proyectado</th><th scope="col" class="num">Cotizado</th>
        <th scope="col" class="num">Pagado</th><th scope="col" class="num">Por pagar</th><th scope="col" class="num">Varianza</th><th scope="col">Avance</th><th scope="col"><span class="sr">Acciones</span></th></tr></thead>
      <tbody>${filas.map(f => `<tr>
        <td class="strong">${esc(f.rubro)}</td>
        <td class="wrap">${esc(f.proveedor || '—')}</td>
        <td class="num">${money(f.proyectado)}</td>
        <td class="num">${money(f.cotizado)}${f.estimado ? ' <span class="small">(est.)</span>' : ''}</td>
        <td class="num">${money(f.pagado)}</td>
        <td class="num ${f.porPagar < -0.009 ? 'neg' : ''}">${f.porPagar < -0.009 ? '(' + money(-f.porPagar) + ')' : money(f.porPagar)}</td>
        <td class="num">${money(f.varianza)}</td>
        <td>${f.sinPres && !f.cotizado ? '<span class="chip c-err">Sin presupuesto</span>' : `<div class="actions" style="gap:8px">${progBar(f.pct, 'Pagado ' + f.pct + ' por ciento de lo cotizado')}${f.pct > 100 ? '<span class="chip c-warn">Excedido</span>' : ''}${f.sinPres ? '<span class="chip c-err">Sin presupuesto</span>' : ''}</div>`}</td>
        <td><button type="button" class="btn-link" data-action="rubro-editar" data-id="${esc(f.presId)}" data-rubro="${esc(f.rubro)}">${f.presId ? 'Editar' : 'Presupuestar'}</button></td>
      </tr>`).join('')}</tbody>
      <tfoot><tr><td colspan="2">Totales</td><td class="num">${money(T.proy)}</td><td class="num">${money(T.cot)}</td><td class="num">${money(T.pag)}</td><td class="num">${money(T.pp)}</td><td class="num">${money(T.vari)}</td><td colspan="2"></td></tr></tfoot>
    </table></div>` : '<div class="empty">Sin rubros. Usa <b>+ Agregar rubro</b>.</div>'}
  </section>

  <section class="card" aria-label="Calendario de dispersión">
    <div class="card-h">
      <div class="t"><h2>Calendario de dispersión</h2>
        <span class="sub">${state.flujoVista === 'mes' ? 'Monto programado por rubro en cada mes' : 'Monto programado por semana en ' + esc(mesLabel(state.mes))}</span></div>
      <div class="actions">
        ${state.flujoVista === 'sem' ? `<label class="sr" for="mesFlujo">Mes</label><input id="mesFlujo" type="month" value="${esc(state.mes)}" data-change="mes" style="width:auto">` : ''}
        <div class="seg" role="group" aria-label="Vista del calendario">
          <button type="button" class="${state.flujoVista === 'mes' ? 'on' : ''}" data-action="flujo-vista" data-v="mes">Mensual</button>
          <button type="button" class="${state.flujoVista === 'sem' ? 'on' : ''}" data-action="flujo-vista" data-v="sem">Semanal</button>
        </div>
      </div>
    </div>
    ${cal.filasCal.length ? `<div class="tbl-wrap"><table class="tbl cal" style="min-width:${300 + cal.cols.length * 130}px">
      <thead><tr><th scope="col">Rubro</th>${cal.etiquetas.map(e => `<th scope="col" class="num">${esc(e)}</th>`).join('')}<th scope="col" class="num">Total</th></tr></thead>
      <tbody>${cal.filasCal.map(f => `<tr><td class="strong">${esc(f.rubro)}</td>${f.vals.map(v => `<td class="num ${v ? 'has' : 'muted'}">${v ? money(v) : '—'}</td>`).join('')}<td class="num strong">${money(f.total)}</td></tr>`).join('')}</tbody>
      <tfoot><tr><td>Total a dispersar</td>${cal.totCols.map(v => `<td class="num">${money(v)}</td>`).join('')}<td class="num">${money(cal.grand)}</td></tr></tfoot>
    </table></div>` : '<div class="empty">Sin pagos programados en este periodo.</div>'}
  </section>`;
}

function abrirRubro(id, rubro) {
  const r = id ? state.data.presupuesto.find(x => x.ID === id) : null;
  abrirDlg(r ? 'Editar rubro del presupuesto' : 'Agregar rubro al presupuesto', `
    <form id="formRubro" style="display:flex;flex-direction:column;gap:14px">
      <input type="hidden" name="id" value="${esc(r ? r.ID : '')}">
      <div class="grid">
        <label class="f">Proyecto <input name="proyecto" list="dlProyR" required value="${esc(r ? r.PROYECTO : state.flujoProy)}">${dlist('dlProyR', proyectosTodos())}</label>
        <label class="f">Rubro / concepto <input name="rubro" list="dlRubR" required value="${esc(r ? r.RUBRO : (rubro && rubro !== 'SIN RUBRO' ? rubro : ''))}" placeholder="Ej. TRANSMISOR EAS SAME">${dlist('dlRubR', rubrosTodos(state.flujoProy))}</label>
        <label class="f">Proveedor (estimado) <input name="proveedor" value="${esc(r ? r.PROVEEDOR : '')}"></label>
        <label class="f">Monto proyectado <input name="proyectado" class="money" inputmode="decimal" required value="${esc(r ? r.proyectado : '')}" placeholder="0.00"></label>
      </div>
      <div class="small">Para que los pagos sumen a este rubro, captura el mismo rubro en las solicitudes o en las OC.</div>
    </form>`,
    `<button type="button" class="btn" data-action="cerrar">Cancelar</button><button type="submit" form="formRubro" class="btn btn-primary">Guardar</button>`);
}

async function exportarFlujo() {
  const P = state.flujoProy;
  const { filas } = datosFlujo(P);
  const a = [[`FLUJO PROYECTO ${P} · ${fechaMx(hoy())} · MONTOS CON IVA`], [], ['RUBRO', 'PROVEEDOR', 'PROYECTADO', 'COTIZADO', 'PAGADO', 'POR PAGAR', 'VARIANZA', '% AVANCE']];
  filas.forEach(f => a.push([f.rubro, f.proveedor, f.proyectado, f.cotizado, f.pagado, f.porPagar, f.varianza, f.pct / 100]));
  a.push([], ['TOTALES', '', sum(filas, f => f.proyectado), sum(filas, f => f.cotizado), sum(filas, f => f.pagado), sum(filas, f => f.porPagar), sum(filas, f => f.varianza)]);
  const hojas = [{ nombre: 'Flujo', filas: a, anchos: [34, 34, 16, 16, 16, 16, 16, 10] }];
  ['mes', 'sem'].forEach(v => {
    const c = calendario(P, v);
    if (!c.filasCal.length) return;
    const t = [['RUBRO', ...c.etiquetas, 'TOTAL']];
    c.filasCal.forEach(f => t.push([f.rubro, ...f.vals, f.total]));
    t.push(['TOTAL', ...c.totCols, c.grand]);
    hojas.push({ nombre: v === 'mes' ? 'Calendario mensual' : 'Semanal ' + state.mes, filas: t, anchos: [34, ...c.cols.map(() => 16), 16] });
  });
  try { await descargarExcel(`Flujo_${P.replace(/[^A-Z0-9]+/g, '_')}.xlsx`, hojas); } catch (e) { toast(e.message, 'err'); }
}

/* =====================================================================
 * DIÁLOGOS, CONFIGURACIÓN E INSTALACIÓN
 * ===================================================================== */
function abrirDlg(titulo, body, foot) {
  $('#dlgTitulo').innerHTML = titulo;
  $('#dlgBody').innerHTML = body;
  $('#dlgFoot').innerHTML = foot || '';
  const d = $('#dlg');
  if (!d.open) d.showModal();
}
function cerrarDlg() { $$('dialog[open]').forEach(d => d.close()); state.expediente = null; }

function abrirConfig() {
  const f = $('#formConfig');
  f.url.value = cfg.url; f.key.value = cfg.key;
  $('#dlgConfig').showModal();
}

let promptInstalar = null;
const esApp = () => window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); promptInstalar = e; });
window.addEventListener('appinstalled', () => { promptInstalar = null; $('#btnInstalar').hidden = true; toast('Aplicación instalada en tu escritorio'); });

async function instalar() {
  if (promptInstalar) {
    promptInstalar.prompt();
    const r = await promptInstalar.userChoice;
    promptInstalar = null;
    if (r.outcome === 'accepted') $('#btnInstalar').hidden = true;
  } else {
    $('#dlgInstalar').showModal();
  }
}

/* ============ EVENTOS ============ */
const acciones = {
  'cerrar': () => cerrarDlg(),
  'config': () => abrirConfig(),
  'recargar': async () => { toast('Actualizando…', 'info', 1500); if (await cargarTodo()) toast('Datos actualizados'); },
  'instalar': () => instalar(),
  // Programación
  'f-semana': el => { state.semana = el.dataset.v; render(); },
  'f-proyecto': el => { state.proyecto = el.dataset.v; render(); },
  'nueva-solicitud': () => abrirNuevaSolicitud(),
  'exportar-cxp': () => exportarCxP(),
  'expediente': el => abrirExpediente(el.dataset.partida),
  'subir-doc': el => subirDoc(el),
  // Registrar pago
  'pg-guardar': () => guardarPago(),
  'pg-limpiar': () => { state.pago = { proveedor: '', globalTocado: false }; render(); },
  // Proveedores
  'prov-pagar': el => { state.pago = { proveedor: el.dataset.v, globalTocado: false }; location.hash = '#/pago'; },
  'oc-nueva': () => abrirOC(),
  'oc-editar': el => abrirOC(el.dataset.id),
  'exportar-prov': () => exportarProveedor(),
  // Flujo
  'flujo-proy': el => { state.flujoProy = el.dataset.v; render(); },
  'flujo-vista': el => { state.flujoVista = el.dataset.v; render(); },
  'rubro-nuevo': () => abrirRubro(),
  'rubro-editar': el => abrirRubro(el.dataset.id, el.dataset.rubro),
  'exportar-flujo': () => exportarFlujo()
};

document.addEventListener('click', e => {
  const el = e.target.closest('[data-action]');
  if (!el) return;
  const fn = acciones[el.dataset.action];
  if (!fn) return;
  if (el.tagName !== 'A') e.preventDefault();
  fn(el, e);
});

document.addEventListener('change', e => {
  const t = e.target.dataset.change;
  if (!t) return;
  if (t === 'mes' && e.target.value) { state.mes = e.target.value; state.semana = 'TODAS'; render(); }
  if (t === 'pg-proveedor') { guardarCamposPago(); state.pago.proveedor = e.target.value; state.pago.seleccion = {}; state.pago.globalTocado = false; state.pago.global = ''; render(); }
  if (t === 'prov-sel') { state.provSel = e.target.value; render(); }
  if (t === 'sol-proveedor') solProveedor(e.target);
  if (t === 'sol-oc') solOC(e.target);
});

document.addEventListener('input', e => {
  const t = e.target.dataset.input;
  if (!t) return;
  if (t === 'buscar') {
    state.buscar = e.target.value;
    const pos = e.target.selectionStart;
    render();
    const inp = $('[data-input="buscar"]'); inp.focus(); inp.setSelectionRange(pos, pos);
  }
  if (t === 'clabe') validarClabe(e.target);
  if (t === 'pg-global') { state.pago.globalTocado = e.target.value.trim() !== ''; recalcPago(); }
  if (t === 'pg-monto') recalcPago();
  if (t === 'pg-check') {
    const inp = $(`#mon-${CSS.escape(e.target.dataset.partida)}`);
    if (e.target.checked && !num(inp.value)) inp.value = num(inp.dataset.pend).toFixed(2);
    recalcPago();
    if (e.target.checked) inp.focus();
  }
  if (t === 'pg-all') {
    $$('[data-pg-check]').forEach(c => {
      c.checked = e.target.checked;
      const inp = $(`#mon-${CSS.escape(c.dataset.partida)}`);
      if (c.checked && !num(inp.value)) inp.value = num(inp.dataset.pend).toFixed(2);
    });
    recalcPago();
  }
});

document.addEventListener('submit', async e => {
  const f = e.target;
  const fid = f.getAttribute('id'); // (f.id no sirve: el campo oculto name="id" lo tapa)
  if (!['formSol', 'formOC', 'formRubro', 'formConfig'].includes(fid)) return;
  e.preventDefault();
  if (fid === 'formSol') return guardarSolicitud(f);
  if (fid === 'formConfig') {
    cfg.save(f.url.value, f.key.value);
    $('#dlgConfig').close();
    toast('Probando conexión…', 'info', 2000);
    if (await cargarTodo()) toast('Conectado a Google Drive');
    return;
  }
  const btn = $('#dlgFoot .btn-primary');
  btn.disabled = true;
  try {
    const d = Object.fromEntries(new FormData(f));
    if (fid === 'formOC') { await api({ action: 'guardarOC', ...d }); toast('Cotización / OC guardada'); state.provSel = up(d.proveedor); }
    if (fid === 'formRubro') { await api({ action: 'guardarRubro', ...d }); toast('Rubro guardado'); state.flujoProy = up(d.proyecto); }
    cerrarDlg();
    await cargarTodo();
  } catch (err) {
    toast(err.message, 'err', 9000);
    btn.disabled = false;
  }
});

$('#dlg').addEventListener('close', () => { state.expediente = null; });
window.addEventListener('hashchange', route);

/* ============ INICIO ============ */
(function iniciar() {
  state.mes = hoy().slice(0, 7);
  if (esApp()) $('#btnInstalar').hidden = true;
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  try {
    const cache = JSON.parse(localStorage.getItem('pa_cache') || 'null');
    if (cache) { state.data = normalizar(cache); setConn(false); }
  } catch (e) { /* cache dañado */ }
  route();
  if (!cfg.key) {
    if (!state.data) $('#view').innerHTML = '<div class="card empty">Captura la clave de acceso para conectar con Google Drive.<br><br><button class="btn btn-primary" data-action="config">Configurar conexión</button></div>';
    abrirConfig();
    return;
  }
  cargarTodo(!!state.data);
})();
