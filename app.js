/* =====================================================================
 * PAGOS Y FLUJO · Administración — Frontend v3
 * Pantallas: Programación (día / semana / mes) · Registrar pago · Proveedores · Flujo por proyecto
 * Las descargas de Excel están en excel.js
 * ===================================================================== */
'use strict';

/* ============ CONFIGURACIÓN ============ */
const API_URL_DEFAULT = 'https://script.google.com/macros/s/AKfycbz-LNkv5JRcpKo6pLlzd5gBpOH6u3zOYbF3RN0DNjMgVz64hEmQEcx-TYL6TWTgFZA/exec';
const MAX_MB = 30;

const DOCS = [
  { tipo: 'OC',          label: 'Órdenes de compra',     accept: '.pdf,.jpg,.jpeg,.png,.xlsx,.docx', col: 'OC_URL' },
  { tipo: 'COTIZACION',  label: 'Cotizaciones',          accept: '.pdf,.jpg,.jpeg,.png,.xlsx,.docx', col: 'COTIZACION_URL' },
  { tipo: 'PAGO',        label: 'Comprobantes de pago',  accept: '.pdf,.jpg,.jpeg,.png',             col: 'PAGO_URL' },
  { tipo: 'FACTURA_PDF', label: 'Facturas PDF',          accept: '.pdf',                             col: 'FACTURA_PDF_URL' },
  { tipo: 'FACTURA_XML', label: 'Facturas XML',          accept: '.xml',                             col: 'FACTURA_XML_URL' },
  { tipo: 'COMPLEMENTO', label: 'Complementos de pago',  accept: '.xml,.pdf',                        col: 'COMPLEMENTO_URL' },
  { tipo: 'OTRO',        label: 'Otros (tickets, fotos…)', accept: '',                               col: 'OTROS_URL' }
];
const EXP_COLS = ['OC_URL', 'COTIZACION_URL', 'PAGO_URL', 'FACTURA_PDF_URL', 'FACTURA_XML_URL', 'COMPLEMENTO_URL'];

const FORMAS_PAGO = ['TRANSFERENCIA 8018', 'EFECTIVALE', 'DEPÓSITO CUENTA PERSONAL', 'AMEX', 'EFECTIVO'];
const CUENTAS_ORIGEN = ['CUENTA EMPRESA 8018', 'EFECTIVALE', 'AMEX', 'EFECTIVO'];

// Clave de banco = 3 primeros dígitos de la CLABE
const BANCOS = { '002': 'BANAMEX', '012': 'BBVA', '014': 'SANTANDER', '021': 'HSBC', '030': 'BAJIO',
  '036': 'INBURSA', '044': 'SCOTIABANK', '058': 'BANREGIO', '062': 'AFIRME', '072': 'BANORTE',
  '112': 'MONEX', '127': 'AZTECA', '137': 'BANCOPPEL', '646': 'STP' };

const ESTATUS = {
  'PROGRAMADO': ['Programado', 'c-info'],
  'PAGO PARCIAL': ['Pago parcial', 'c-warn'],
  'PAGADO - FALTAN DOCUMENTOS': ['Pagado · faltan docs', 'c-warn'],
  'COMPLETO': ['Completo', 'c-ok'],
  'CANCELADA': ['Cancelada', 'c-gray']
};

/* ============ ESTADO ============ */
const cfg = {
  get url() { return localStorage.getItem('pa_url') || API_URL_DEFAULT; },
  get key() { return localStorage.getItem('pa_key') || ''; },
  save(url, key) { localStorage.setItem('pa_url', url.trim()); localStorage.setItem('pa_key', key.trim()); }
};

const state = {
  data: null,
  view: 'programacion',
  periodo: 'semana',   // dia | semana | mes
  mes: '', semana: '', dia: '',
  proyecto: 'TODOS',
  buscar: '',
  verCanceladas: false,
  provSel: '',
  flujoProy: '',
  flujoVista: 'mes',
  pago: { proveedor: '', globalTocado: false },
  expediente: null,
  ocForm: null
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
const fechaMx = s => { if (!s) return ''; const [y, m, d] = String(s).slice(0, 10).split('-'); return d ? `${d}/${m}/${y}` : ''; };
const mesLabel = ym => { if (!ym) return ''; const [y, m] = ym.split('-').map(Number); return new Date(y, m - 1, 1).toLocaleDateString('es-MX', { month: 'long', year: 'numeric' }); };
const mesCorto = ym => { const [y, m] = ym.split('-').map(Number); return new Date(y, m - 1, 1).toLocaleDateString('es-MX', { month: 'short', year: 'numeric' }); };
const lunesISO = s => { if (!s) return ''; const d = parseISO(s); d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); return isoDate(d); };
const masDias = (iso, n) => { const d = parseISO(iso); d.setDate(d.getDate() + n); return isoDate(d); };
const semanaLabel = lun => { const a = parseISO(lun), b = parseISO(masDias(lun, 6)); const f = d => d.toLocaleDateString('es-MX', { day: 'numeric', month: 'short' }); return `${f(a)} – ${f(b)}`; };
const diaLabel = iso => parseISO(iso).toLocaleDateString('es-MX', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
const pct = (a, b) => (b > 0 ? Math.round(a / b * 100) : 0);
const pl = (n, uno, varios) => `${n} ${n === 1 ? uno : varios}`;
const lineas = t => String(t || '').split('\n');
const esViaticos = p => /VI[AÁ]TIC/.test(up(p.rubro) + ' ' + up(p.CONCEPTO)) || ['EFECTIVALE', 'DEPÓSITO CUENTA PERSONAL'].includes(up(p.FORMA_PAGO));

/** Semanas (lunes) que tocan un mes 'YYYY-MM'. */
function semanasDelMes(ym) {
  const [y, m] = ym.split('-').map(Number);
  const fin = isoDate(new Date(y, m, 0));
  const out = [];
  for (let l = lunesISO(`${ym}-01`); l <= fin; l = masDias(l, 7)) out.push(l);
  return out;
}

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
const icoOk = '<svg aria-hidden="true" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12l5 5L20 7"/></svg>';

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
  const provCat = (d.proveedores || []).map(p => ({ ...p, NOMBRE: up(p.NOMBRE) }));
  const provMap = {};
  provCat.forEach(p => { provMap[p.NOMBRE] = p; });

  const ocs = (d.ocs || []).map(o => ({ ...o, OC: up(o.OC), PROVEEDOR: up(o.PROVEEDOR), PROYECTO: up(o.PROYECTO), RUBRO: up(o.RUBRO), total: num(o.TOTAL) }));
  const ocById = {}, ocByOC = {};
  ocs.forEach(o => { ocById[o.ID] = o; if (o.OC && !ocByOC[o.OC]) ocByOC[o.OC] = o; });

  const partidas = (d.partidas || []).map(p => {
    const prog = num(p.MONTO_PROGRAMADO), pag = num(p.MONTO_PAGADO), oc = up(p.OC);
    const cot = p.ID_COTIZACION === '-' ? null : (ocById[p.ID_COTIZACION] || (!p.ID_COTIZACION && oc && ocByOC[oc]) || null);
    return {
      ...p, OC: oc, PROYECTO: up(p.PROYECTO), PROVEEDOR: up(p.PROVEEDOR),
      prog, pag, pend: r2(prog - pag), cancelada: p.ESTATUS === 'CANCELADA',
      cot, cotId: cot ? cot.ID : '', rubro: up(p.RUBRO) || (cot ? cot.RUBRO : ''),
      ndocs: EXP_COLS.filter(c => p[c]).length
    };
  });
  const porNum = {};
  partidas.forEach(p => { porNum[String(p.PARTIDA)] = p; });
  const activas = partidas.filter(p => !p.cancelada);

  const apps = (d.aplicaciones || []).filter(a => a.ESTATUS !== 'CANCELADO')
    .map(a => ({ ...a, monto: num(a.MONTO), p: porNum[String(a.PARTIDA)] })).filter(a => a.p && !a.p.cancelada);

  const pars = (d.parcialidades || []).map(x => ({ ...x, monto: num(x.MONTO), n: num(x.NUM), cot: ocById[x.ID_COTIZACION] })).filter(x => x.cot);
  pars.forEach(x => {
    const p = x.PARTIDA ? porNum[String(x.PARTIDA)] : null;
    x.partida = p && !p.cancelada ? p : null;
    x.estado = !x.partida ? 'TENTATIVA' : (x.partida.pend <= 0.009 ? 'PAGADA' : (x.partida.pag > 0 ? 'PAGO PARCIAL' : 'PROGRAMADA'));
  });

  ocs.forEach(o => {
    o.partidas = activas.filter(p => p.cotId === o.ID);
    o.pag = sum(o.partidas, p => p.pag);
    o.porPagar = r2(o.total - o.pag);
    o.pct = pct(o.pag, o.total);
    o.pars = pars.filter(x => x.ID_COTIZACION === o.ID).sort((a, b) => a.n - b.n);
  });

  return {
    partidas, activas, porNum, ocs, ocById, ocByOC, pars, apps, provCat, provMap,
    presupuesto: (d.presupuesto || []).map(r => ({ ...r, PROYECTO: up(r.PROYECTO), RUBRO: up(r.RUBRO), proyectado: num(r.PROYECTADO) })),
    pagos: (d.pagos || []).filter(p => p.ESTATUS !== 'CANCELADO').map(p => ({ ...p, PROVEEDOR: up(p.PROVEEDOR), monto: num(p.MONTO_GLOBAL) })),
    cobranza: (d.cobranza || []).map(c => ({ ...c, PROYECTO: up(c.PROYECTO), monto: num(c.MONTO) })),
    servidor: d.servidor || '', siguiente: d.siguientePartida || ''
  };
}

async function cargarTodo(silencioso = false) {
  try {
    const d = await api({ action: 'todo' });
    try { localStorage.setItem('pa_cache', JSON.stringify(d)); } catch (e) { /* sin espacio */ }
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

const D = () => state.data;
const proyectosTodos = () => uniq([...D().activas.map(p => p.PROYECTO), ...D().ocs.map(o => o.PROYECTO), ...D().presupuesto.map(r => r.PROYECTO), ...D().cobranza.map(c => c.PROYECTO)]).sort();
const proveedoresTodos = () => uniq([...D().provCat.map(p => p.NOMBRE), ...D().partidas.map(p => p.PROVEEDOR), ...D().ocs.map(o => o.PROVEEDOR)]).sort();
const rubrosTodos = proy => uniq([...D().presupuesto.filter(r => !proy || r.PROYECTO === proy).map(r => r.RUBRO), ...D().ocs.filter(o => !proy || o.PROYECTO === proy).map(o => o.RUBRO), ...D().activas.filter(p => !proy || p.PROYECTO === proy).map(p => p.rubro)]).sort();

function datosBanco(prov) {
  const c = D().provMap[prov];
  const ps = D().partidas.filter(p => p.PROVEEDOR === prov && (p.BANCO || p.CUENTA_DEPOSITO || p.CLABE_REFERENCIA)).sort((a, b) => num(b.PARTIDA) - num(a.PARTIDA));
  const p = ps[0] || {};
  const last = D().partidas.filter(x => x.PROVEEDOR === prov).sort((a, b) => num(b.PARTIDA) - num(a.PARTIDA))[0] || {};
  if (!c && !ps.length && !last.PARTIDA) return null;
  return {
    banco: (c && c.BANCO) || p.BANCO || '', cuenta: (c && c.CUENTA) || p.CUENTA_DEPOSITO || '', clabe: (c && c.CLABE) || p.CLABE_REFERENCIA || '',
    rfc: c ? c.RFC : '', contacto: c ? [c.CONTACTO, c.TELEFONO, c.CORREO].filter(Boolean).join(' · ') : '',
    forma: last.FORMA_PAGO || '', proyecto: last.PROYECTO || ''
  };
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
 * 1. PROGRAMACIÓN (día / semana / mes)
 * ===================================================================== */
function periodoTexto() {
  if (state.periodo === 'dia') return diaLabel(state.dia);
  if (state.periodo === 'semana') return 'semana del ' + semanaLabel(state.semana);
  return mesLabel(state.mes);
}

function enPeriodo(fecha) {
  const f = String(fecha || '').slice(0, 10);
  if (state.periodo === 'dia') return f === state.dia;
  if (state.periodo === 'semana') return lunesISO(f) === state.semana;
  return f.slice(0, 7) === state.mes;
}

function filasProgramacion() {
  const base = (state.verCanceladas ? D().partidas : D().activas).filter(p => enPeriodo(p.FECHA_PROGRAMADA));
  const proyectos = uniq(base.map(p => p.PROYECTO)).sort();
  if (state.proyecto !== 'TODOS' && !proyectos.includes(state.proyecto)) state.proyecto = 'TODOS';
  const q = state.buscar.toLowerCase();
  const filas = base
    .filter(p => state.proyecto === 'TODOS' || p.PROYECTO === state.proyecto)
    .filter(p => !q || [p.PROVEEDOR, p.OC, p.CONCEPTO, p.PARTIDA, p.PROYECTO].join(' ').toLowerCase().includes(q))
    .sort((a, b) => (a.FECHA_PROGRAMADA || '').localeCompare(b.FECHA_PROGRAMADA || '') || num(a.PARTIDA) - num(b.PARTIDA));
  return { proyectos, filas };
}

function vistaProgramacion() {
  const semanas = semanasDelMes(state.mes);
  if (!semanas.includes(state.semana)) state.semana = semanas.includes(lunesISO(hoy())) ? lunesISO(hoy()) : semanas[0];
  const { proyectos, filas } = filasProgramacion();
  const vivas = filas.filter(p => !p.cancelada);
  const prog = sum(vivas, p => p.prog), pag = sum(vivas, p => p.pag), pend = sum(vivas, p => Math.max(p.pend, 0));
  const nPend = vivas.filter(p => p.pend > 0.009).length;
  const incompletos = vivas.filter(p => p.ESTATUS === 'PAGADO - FALTAN DOCUMENTOS').length;

  const selector = state.periodo === 'dia'
    ? `<label class="f inline">Día <input type="date" value="${esc(state.dia)}" data-change="dia"></label>
       <button type="button" class="btn btn-sm" data-action="dia-mover" data-v="-1" aria-label="Día anterior">‹</button>
       <button type="button" class="btn btn-sm" data-action="dia-mover" data-v="1" aria-label="Día siguiente">›</button>
       <button type="button" class="btn btn-sm" data-action="dia-hoy">Hoy</button>`
    : `<label class="f inline">Mes <input type="month" value="${esc(state.mes)}" data-change="mes"></label>`;

  return `
  <header class="page-h">
    <div class="t">
      <span class="eyebrow">Cuentas por pagar · ${esc(periodoTexto())}</span>
      <h1>Programación de pagos</h1>
    </div>
    <div class="actions">
      <button class="btn" type="button" data-action="exportar-cxp-ind" title="Formato de programación para enviar a Cuentas por Pagar">Descargar CxP individual</button>
      <button class="btn" type="button" data-action="exportar-cxp-gen" title="Histórico con pestañas VIÁTICOS, CXP PROYECTOS 8018 y AMEX">Descargar CxP general</button>
      <button class="btn btn-primary" type="button" data-action="nueva-solicitud">+ Nueva solicitud</button>
    </div>
  </header>

  <section class="card card-p periodo" aria-label="Periodo">
    <div class="seg" role="group" aria-label="Consultar por">
      ${[['dia', 'Por día'], ['semana', 'Por semana'], ['mes', 'Por mes']].map(([v, l]) => `<button type="button" class="${state.periodo === v ? 'on' : ''}" data-action="periodo" data-v="${v}">${l}</button>`).join('')}
    </div>
    <div class="actions">${selector}</div>
    ${state.periodo === 'semana' ? `<div class="fchips" role="group" aria-label="Semana">${semanas.map(s => `<button type="button" class="fchip ${state.semana === s ? 'on' : ''}" data-action="f-semana" data-v="${s}">${esc(semanaLabel(s))}</button>`).join('')}</div>` : ''}
  </section>

  <section class="kpis" aria-label="Resumen">
    <div class="kpi"><span class="l">Programado</span><span class="v">${money(prog)}</span><span class="s">${pl(vivas.length, 'solicitud', 'solicitudes')}</span></div>
    <div class="kpi"><span class="l">Pagado (con comprobante)</span><span class="v v-ok">${money(pag)}</span><span class="s">Se aplica al recibir el comprobante</span></div>
    <div class="kpi"><span class="l">Por pagar</span><span class="v v-warn">${money(pend)}</span><span class="s">${pl(nPend, 'partida pendiente', 'partidas pendientes')}</span></div>
    <div class="kpi"><span class="l">Expedientes incompletos</span><span class="v v-err">${incompletos}</span><span class="s">Pagadas sin factura, XML u OC</span></div>
  </section>

  <section class="card" aria-label="Solicitudes de pago">
    <div class="card-h">
      <div class="fchips" role="group" aria-label="Filtrar por proyecto">
        <button type="button" class="fchip ${state.proyecto === 'TODOS' ? 'on' : ''}" data-action="f-proyecto" data-v="TODOS">Todos los proyectos</button>
        ${proyectos.map(p => `<button type="button" class="fchip ${state.proyecto === p ? 'on' : ''}" data-action="f-proyecto" data-v="${esc(p)}">${esc(p)}</button>`).join('')}
      </div>
      <div class="actions">
        <label class="check"><input type="checkbox" data-change="ver-canceladas" ${state.verCanceladas ? 'checked' : ''}> Ver canceladas</label>
        <label class="search" style="flex:0 1 280px"><span class="sr">Buscar</span>
          <svg aria-hidden="true" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#5B6474" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>
          <input type="search" placeholder="Proveedor, OC, concepto o N°" value="${esc(state.buscar)}" data-input="buscar">
        </label>
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
          ${filas.map(p => `<tr class="${p.cancelada ? 'row-off' : ''}">
            <td class="strong">${esc(p.PARTIDA)}</td>
            <td class="nowrap">${fechaMx(p.FECHA_PROGRAMADA)}</td>
            <td>${esc(p.PROYECTO)}</td>
            <td><div class="strong">${esc(p.PROVEEDOR)}</div><div class="small">${esc(p.FORMA_PAGO)}</div></td>
            <td class="nowrap">${ocCell(p.OC)}${p.cot && p.cot.COTIZACION ? `<div class="small">Cot. ${esc(p.cot.COTIZACION)}</div>` : ''}</td>
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
    </div>` : `<div class="empty">No hay solicitudes para ${esc(periodoTexto())}${state.proyecto !== 'TODOS' ? ' en ' + esc(state.proyecto) : ''}. Usa <b>+ Nueva solicitud</b>.</div>`}
  </section>`;
}

/* ---------- Nueva solicitud / editar partida ---------- */
function abrirSolicitud(opc = {}) {
  const p = opc.partida ? D().porNum[String(opc.partida)] : null;
  const provs = proveedoresTodos(), proys = proyectosTodos();
  const v = (campo, def = '') => esc(p ? (p[campo] ?? def) : def);
  const prov = p ? p.PROVEEDOR : (opc.proveedor || '');
  abrirDlg(p ? `Editar partida ${esc(p.PARTIDA)}` : 'Nueva solicitud de pago', `
    <form id="formSol" data-partida="${p ? esc(p.PARTIDA) : ''}" style="display:flex;flex-direction:column;gap:16px">
      <div class="paso"><span class="paso-n">1</span><div><b>¿A quién se le paga?</b><div class="small">Si el proveedor ya existe, sus datos bancarios se llenan solos. Si es nuevo, se da de alta con lo que captures.</div></div></div>
      <div class="grid">
        <label class="f">Proveedor / beneficiario
          <input name="proveedor" list="dlProv" required autocomplete="off" value="${esc(prov)}" placeholder="Escribe o elige de la lista" data-change="sol-proveedor">
          ${dlist('dlProv', provs)}
        </label>
        <label class="f">Forma de pago <select name="formaPago">${opts(FORMAS_PAGO, p ? p.FORMA_PAGO : '')}</select></label>
        <label class="f">Banco <input name="banco" list="dlBancos" value="${v('BANCO')}" placeholder="Ej. BBVA">${dlist('dlBancos', uniq(Object.values(BANCOS)).concat(['EFECTIVALE', 'AMEX']))}</label>
        <label class="f">Cuenta de depósito <input name="cuenta" inputmode="numeric" value="${v('CUENTA_DEPOSITO')}" placeholder="Número de cuenta"></label>
        <label class="f">CLABE interbancaria / línea de captura
          <input name="clabe" inputmode="numeric" value="${v('CLABE_REFERENCIA')}" placeholder="18 dígitos o referencia" data-input="clabe">
          <span class="hint" id="hintClabe"></span>
        </label>
      </div>
      <div class="small" id="hintProv"></div>

      <div class="paso"><span class="paso-n">2</span><div><b>¿De qué cotización y parcialidad es?</b><div class="small">Opcional. Al elegirla se llenan monto, OC, proyecto y rubro, y se descuenta del saldo de la cotización.</div></div></div>
      <div class="grid">
        <label class="f">Cotización / OC <select name="idCotizacion" data-change="sol-cot"></select></label>
        <label class="f">Parcialidad <select name="idParcialidad" data-change="sol-par"></select></label>
      </div>
      <div class="small" id="hintCot"></div>

      <div class="paso"><span class="paso-n">3</span><div><b>Datos del pago</b></div></div>
      <div class="grid">
        <label class="f">Fecha programada <input type="date" name="fechaProgramada" required value="${v('FECHA_PROGRAMADA', hoy())}"></label>
        <label class="f">Monto programado <input name="montoProgramado" inputmode="decimal" required placeholder="0.00" class="money" value="${p ? esc(p.prog) : ''}"></label>
        <label class="f">Proyecto <input name="proyecto" list="dlProy" required value="${v('PROYECTO')}" placeholder="Ej. VERACRUZ">${dlist('dlProy', proys)}</label>
        <label class="f">OC / referencia <input name="oc" value="${v('OC')}" placeholder="Ej. OC-VER-045 o SIN OC"></label>
        <label class="f">Rubro del flujo <input name="rubro" list="dlRubro" value="${p ? esc(p.RUBRO || p.rubro) : ''}" placeholder="Ej. TRANSMISOR EAS SAME">${dlist('dlRubro', rubrosTodos())}</label>
      </div>
      <label class="f">Concepto <input name="concepto" value="${v('CONCEPTO')}" placeholder="Descripción del pago"></label>
      <label class="f">Observaciones <textarea name="observaciones" placeholder="Notas para CxP o Administración">${v('OBSERVACIONES')}</textarea></label>
      ${p ? '' : `<div class="grid">
        <label class="f">Orden de compra (opcional) <input type="file" name="fileOC" accept=".pdf,.jpg,.jpeg,.png,.xlsx,.docx"></label>
        <label class="f">Cotización (opcional) <input type="file" name="fileCOT" accept=".pdf,.jpg,.jpeg,.png,.xlsx,.docx"></label>
      </div>
      <div class="small">Siguiente número de partida: <b>${esc(D().siguiente || '—')}</b>. Se creará su carpeta en Drive dentro del mes de la fecha programada.</div>`}
    </form>`,
    `<button type="button" class="btn" data-action="cerrar">Cancelar</button>
     <button type="submit" form="formSol" class="btn btn-primary" id="btnSol">${p ? 'Guardar cambios' : 'Crear partida y carpeta'}</button>`);
  const f = $('#formSol');
  llenarCotizaciones(f, prov, p ? (p.ID_COTIZACION === '-' ? '' : p.cotId) : (opc.cotId || ''));
  llenarParcialidades(f, p ? p.ID_PARCIALIDAD : (opc.parId || ''), p);
  if (!p && opc.parId) aplicarParcialidad(f, true);
  else if (!p && opc.cotId) aplicarCotizacion(f);
  if (!p && prov) solProveedor(f.proveedor, true);
  if (f.clabe.value) validarClabe(f.clabe);
}

function llenarCotizaciones(f, prov, sel) {
  const cots = D().ocs.filter(o => o.PROVEEDOR === up(prov));
  f.idCotizacion.innerHTML = `<option value="">— Sin cotización —</option>` + cots.map(o =>
    `<option value="${esc(o.ID)}" ${o.ID === sel ? 'selected' : ''}>${esc([o.COTIZACION || 'Sin número', o.OC, o.CONCEPTO].filter(Boolean).join(' · '))} · saldo ${money(o.porPagar)}</option>`).join('');
}

function llenarParcialidades(f, sel, partidaActual) {
  const o = D().ocById[f.idCotizacion.value];
  const pars = o ? o.pars.filter(x => !x.partida || x.ID === sel || (partidaActual && String(x.PARTIDA) === String(partidaActual.PARTIDA))) : [];
  f.idParcialidad.innerHTML = `<option value="">— Ninguna —</option>` + pars.map(x =>
    `<option value="${esc(x.ID)}" ${x.ID === sel ? 'selected' : ''}>${x.n}. ${esc(x.CONCEPTO)} · ${money(x.monto)}${x.FECHA_TENTATIVA ? ' · ' + fechaMx(x.FECHA_TENTATIVA) : ''}</option>`).join('');
  f.idParcialidad.disabled = !o;
  const h = $('#hintCot');
  if (h) h.textContent = o ? `Total ${money(o.total)} · pagado ${money(o.pag)} · por pagar ${money(o.porPagar)}${o.pars.length ? ' · ' + pl(o.pars.filter(x => !x.partida).length, 'parcialidad sin programar', 'parcialidades sin programar') : ''}` : '';
}

function aplicarCotizacion(f) {
  const o = D().ocById[f.idCotizacion.value];
  if (!o) return;
  if (o.OC) f.oc.value = o.OC;
  if (o.PROYECTO) f.proyecto.value = o.PROYECTO;
  if (o.RUBRO) f.rubro.value = o.RUBRO;
  if (!f.concepto.value && o.CONCEPTO) f.concepto.value = o.CONCEPTO;
}

function aplicarParcialidad(f, forzar) {
  const o = D().ocById[f.idCotizacion.value];
  const x = o && o.pars.find(y => y.ID === f.idParcialidad.value);
  aplicarCotizacion(f);
  if (!x) return;
  if (forzar || !num(f.montoProgramado.value)) f.montoProgramado.value = x.monto.toFixed(2);
  if (x.FECHA_TENTATIVA && (forzar || !f.fechaProgramada.value)) f.fechaProgramada.value = x.FECHA_TENTATIVA;
  f.concepto.value = [o.CONCEPTO, x.CONCEPTO].filter(Boolean).join(' · ');
}

function solProveedor(input, silencioso) {
  const f = input.form;
  const prov = up(input.value);
  llenarCotizaciones(f, prov, f.idCotizacion.value);
  llenarParcialidades(f, f.idParcialidad.value);
  const b = datosBanco(prov);
  const h = $('#hintProv');
  if (!b) { h.textContent = prov ? 'Proveedor nuevo: se dará de alta con estos datos bancarios.' : ''; return; }
  if (!f.banco.value) f.banco.value = b.banco;
  if (!f.cuenta.value) f.cuenta.value = b.cuenta;
  if (!f.clabe.value) { f.clabe.value = b.clabe; validarClabe(f.clabe); }
  if (!f.proyecto.value && b.proyecto) f.proyecto.value = b.proyecto;
  if (!silencioso && b.forma && FORMAS_PAGO.includes(b.forma)) f.formaPago.value = b.forma;
  h.textContent = 'Datos bancarios tomados del catálogo de proveedores. Revísalos.';
}

function validarClabe(input) {
  const v = input.value.replace(/\s+/g, '');
  const h = input.closest('label').querySelector('.hint'), f = input.form;
  if (!h) return;
  if (/^\d{18}$/.test(v)) {
    if (clabeValida(v)) {
      const banco = BANCOS[v.slice(0, 3)];
      h.style.color = 'var(--ok)'; h.textContent = 'CLABE válida' + (banco ? ' · ' + banco : '');
      if (banco && f && f.banco && !f.banco.value) f.banco.value = banco;
    } else { h.style.color = 'var(--err)'; h.textContent = 'Dígito verificador incorrecto. Revisa la CLABE.'; }
  } else if (/^\d{15,17}$/.test(v)) { h.style.color = 'var(--warn)'; h.textContent = `Llevas ${v.length} dígitos; una CLABE tiene 18.`; }
  else h.textContent = '';
}

async function guardarSolicitud(form) {
  const btn = $('#btnSol');
  const d = Object.fromEntries(new FormData(form));
  delete d.fileOC; delete d.fileCOT;
  if (!num(d.montoProgramado)) return toast('Captura el monto programado', 'warn');
  const editar = form.dataset.partida;
  btn.disabled = true; btn.textContent = editar ? 'Guardando…' : 'Creando carpeta…';
  try {
    if (editar) {
      await api({ action: 'actualizarPartida', partida: editar, ...d });
      cerrarDlg();
      toast(`Partida ${editar} actualizada`);
      await cargarTodo();
      abrirExpediente(editar);
      return;
    }
    const { partida } = await api({ action: 'crearPartida', ...d });
    const oc = form.fileOC.files[0], cot = form.fileCOT.files[0];
    if (oc) { btn.textContent = 'Subiendo OC…'; await subir(partida.PARTIDA, 'OC', oc); }
    if (cot) { btn.textContent = 'Subiendo cotización…'; await subir(partida.PARTIDA, 'COTIZACION', cot); }
    cerrarDlg();
    state.mes = partida.MES_CARPETA; state.semana = lunesISO(partida.FECHA_PROGRAMADA); state.dia = partida.FECHA_PROGRAMADA; state.proyecto = 'TODOS';
    if (location.hash !== '#/programacion' && location.hash !== '') location.hash = '#/programacion';
    toast(`Partida ${partida.PARTIDA} creada con su carpeta en Drive`);
    await cargarTodo();
  } catch (e) {
    toast(e.message, 'err', 9000);
    btn.disabled = false; btn.textContent = editar ? 'Guardar cambios' : 'Crear partida y carpeta';
  }
}

/* ---------- Expediente (diálogo) ---------- */
function itemArchivo(tipo, url, i, extra) {
  return `<li><a href="${esc(url)}" target="_blank" rel="noopener">Archivo ${i + 1}</a>${extra ? ` <span class="small">${extra}</span>` : ''}
    <button type="button" class="btn-link danger" data-action="quitar-doc" data-tipo="${tipo}" data-url="${esc(url)}">Quitar</button></li>`;
}

function abrirExpediente(partida) {
  state.expediente = String(partida);
  const p = D().porNum[state.expediente];
  if (!p) return toast('No se encontró la partida', 'err');
  const apps = D().apps.filter(a => String(a.PARTIDA) === state.expediente).sort((a, b) => (a.FECHA_PAGO || '').localeCompare(b.FECHA_PAGO || ''));
  const par = p.ID_PARCIALIDAD && p.cot ? p.cot.pars.find(x => x.ID === p.ID_PARCIALIDAD) : null;

  const slots = DOCS.map(d => {
    let items = '', n = 0;
    if (d.tipo === 'PAGO') {
      n = apps.length;
      items = apps.map((a, i) => `<li><a href="${esc(a.COMPROBANTE_URL)}" target="_blank" rel="noopener">Pago ${i + 1}</a> <span class="small">${money(a.monto)} · ${fechaMx(a.FECHA_PAGO)} · ${esc(a.ID_PAGO)}</span>
        <button type="button" class="btn-link danger" data-action="quitar-doc" data-tipo="PAGO" data-url="${esc(a.COMPROBANTE_URL)}">Quitar</button></li>`).join('');
    } else {
      const urls = lineas(p[d.col]).filter(Boolean);
      n = urls.length;
      const uu = lineas(p.FACTURA_UUID), tt = lineas(p.FACTURA_TOTAL), cu = lineas(p.COMPLEMENTO_UUID);
      items = urls.map((u, i) => {
        let ex = '';
        if (d.tipo === 'FACTURA_XML') ex = [uu[i] ? 'UUID ' + esc(uu[i]) : '', tt[i] ? money(tt[i]) : ''].filter(Boolean).join(' · ');
        if (d.tipo === 'COMPLEMENTO' && cu[i]) ex = 'UUID ' + esc(cu[i]);
        return itemArchivo(d.tipo, u, i, ex);
      }).join('');
    }
    const sumaFact = d.tipo === 'FACTURA_XML' && n ? `<small>Total facturado: <b>${money(sum(lineas(p.FACTURA_TOTAL), num))}</b></small>` : '';
    const sumaPag = d.tipo === 'PAGO' && n ? `<small>Pagado: <b>${money(p.pag)}</b> de ${money(p.prog)}</small>` : '';
    const campos = d.tipo === 'PAGO' ? `
      <div class="two">
        <label class="sr" for="mp">Monto de este pago</label><input id="mp" data-monto inputmode="decimal" class="money" placeholder="Monto" value="${p.pend > 0.009 ? esc(p.pend.toFixed(2)) : ''}">
        <label class="sr" for="fp">Fecha de pago</label><input id="fp" data-fecha type="date" value="${hoy()}">
      </div>` : '';
    const boton = { PAGO: 'Agregar pago', FACTURA_PDF: 'Agregar factura', FACTURA_XML: 'Agregar XML', COMPLEMENTO: 'Agregar complemento' }[d.tipo] || 'Agregar';
    return `<div class="slot ${n ? 'done' : ''}" data-tipo="${d.tipo}">
      <div class="slot-h"><span>${d.label}</span>${n ? `<span class="chip c-ok">${n}</span>` : '<span class="chip c-warn">Pendiente</span>'}</div>
      ${n ? `<ul class="files">${items}</ul>` : ''}${sumaFact}${sumaPag}${campos}
      <label class="sr" for="f-${d.tipo}">Archivo ${d.label}</label>
      <input id="f-${d.tipo}" type="file" ${d.accept ? `accept="${d.accept}"` : ''} ${d.tipo !== 'PAGO' ? 'multiple' : ''} ${p.cancelada ? 'disabled' : ''}>
      <button class="btn btn-sm ${n ? '' : 'btn-primary'}" type="button" data-action="subir-doc" ${p.cancelada ? 'disabled' : ''}>${boton}</button>
    </div>`;
  }).join('');

  abrirDlg(`Expediente · partida ${esc(p.PARTIDA)}`, `
    <div style="display:flex;flex-wrap:wrap;justify-content:space-between;gap:12px;align-items:flex-start">
      <div style="display:flex;flex-direction:column;gap:6px;min-width:0;flex:1 1 420px">
        <div style="font-size:17px;font-weight:700">${esc(p.PARTIDA)}. ${esc(p.PROVEEDOR)} ${money(p.pag || p.prog)}</div>
        <div class="meta">
          <span>Proyecto: <b>${esc(p.PROYECTO)}</b></span><span>OC: <b>${esc(p.OC || '—')}</b></span><span>Rubro: <b>${esc(p.rubro || '—')}</b></span>
          <span>Programado: <b>${money(p.prog)}</b> para ${fechaMx(p.FECHA_PROGRAMADA)}</span><span>Pendiente: <b>${money(p.pend)}</b></span>
          <span>Forma: ${esc(p.FORMA_PAGO)}</span><span>Banco: ${esc(p.BANCO || '—')}</span><span>Cuenta: ${esc(p.CUENTA_DEPOSITO || '—')}</span><span>CLABE / ref.: ${esc(p.CLABE_REFERENCIA || '—')}</span>
          ${p.cot ? `<span>Cotización: <b>${esc(p.cot.COTIZACION || p.cot.OC || '—')}</b>${par ? ` · parcialidad ${par.n} de ${p.cot.pars.length} (${esc(par.CONCEPTO)})` : ''}</span>` : ''}
        </div>
        ${p.CONCEPTO ? `<div class="sub">${esc(p.CONCEPTO)}</div>` : ''}
      </div>
      <div class="actions">${chipEstatus(p.ESTATUS)}
        <a class="btn btn-sm" href="${esc(p.CARPETA_URL)}" target="_blank" rel="noopener">Abrir carpeta en Drive</a>
        <button type="button" class="btn btn-sm" data-action="editar-partida">Editar datos</button>
        ${p.cancelada ? '<button type="button" class="btn btn-sm" data-action="reactivar-partida">Reactivar</button>'
          : (p.pag ? '' : '<button type="button" class="btn btn-sm danger" data-action="cancelar-partida">Cancelar partida</button>')}
      </div>
    </div>
    <div class="slots">${slots}</div>
    <label class="f">Observaciones
      <textarea id="obsExp" placeholder="Notas: comprobación entregada, pagos directos, diferencias…">${esc(p.OBSERVACIONES)}</textarea>
    </label>
    <div><button type="button" class="btn btn-sm" data-action="guardar-obs">Guardar observaciones</button></div>`,
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
    if (!num(extra.montoPagado) || !extra.fechaPago) return toast('Captura monto y fecha de este pago', 'warn');
  }
  btn.disabled = true; btn.textContent = 'Subiendo…';
  try {
    const avisos = [];
    for (const f of files) { const r = await subir(partida, tipo, f, extra); if (r.aviso) avisos.push(r.aviso); }
    await cargarTodo(true);
    abrirExpediente(partida);
    if (avisos.length) toast('Guardado. Atención: ' + avisos.join(' '), 'warn', 10000);
    else toast(tipo === 'PAGO' ? 'Pago agregado; la carpeta se renombró con el monto y la fecha.' : `${pl(files.length, 'archivo guardado', 'archivos guardados')} en Drive`);
  } catch (e) {
    toast(e.message, 'err', 9000);
    btn.disabled = false; btn.textContent = 'Reintentar';
  }
}

async function quitarDoc(el) {
  const tipo = el.dataset.tipo;
  const msg = tipo === 'PAGO' ? '¿Quitar este comprobante de pago? Se restará su monto de lo pagado y el archivo se mandará a la papelera de Drive.'
    : '¿Quitar este archivo? Se mandará a la papelera de Drive (se puede recuperar 30 días).';
  if (!window.confirm(msg)) return;
  try {
    await api({ action: 'quitarArchivo', partida: state.expediente, tipo, url: el.dataset.url });
    const n = state.expediente;
    await cargarTodo(true);
    abrirExpediente(n);
    toast('Archivo quitado');
  } catch (e) { toast(e.message, 'err', 9000); }
}

async function accionPartida(tipo) {
  const n = state.expediente;
  try {
    if (tipo === 'obs') await api({ action: 'actualizarPartida', partida: n, observaciones: $('#obsExp').value });
    if (tipo === 'cancelar') {
      if (!window.confirm('¿Cancelar esta partida? Ya no aparecerá en la programación ni en el flujo. La carpeta se queda en Drive marcada como CANCELADA.')) return;
      await api({ action: 'cancelarPartida', partida: n });
    }
    if (tipo === 'reactivar') await api({ action: 'cancelarPartida', partida: n, reactivar: true });
    await cargarTodo(true);
    abrirExpediente(n);
    toast({ obs: 'Observaciones guardadas', cancelar: 'Partida cancelada', reactivar: 'Partida reactivada' }[tipo]);
  } catch (e) { toast(e.message, 'err', 9000); }
}

/* =====================================================================
 * 2. REGISTRAR PAGO (monto global → varias partidas)
 * ===================================================================== */
function vistaPago() {
  const P = state.pago;
  const pendientes = D().activas.filter(p => p.pend > 0.009);
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
        ${provs.length ? '' : '<div class="small">No hay partidas pendientes de pago. Crea solicitudes en Programación.</div>'}
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
          <li>Saldos de cotizaciones y parcialidades</li>
          <li>Flujo de cada proyecto involucrado</li>
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
        : `Cuadra: ${money(global)} distribuidos entre ${pl(marcadas, 'partida', 'partidas')}.`;
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
    toast(`Pago ${res.idPago} registrado en ${pl(aplicaciones.length, 'partida', 'partidas')}.` + (res.aviso ? ' ' + res.aviso : ''), res.aviso ? 'warn' : 'ok', 8000);
    await cargarTodo();
  } catch (e) {
    toast(e.message, 'err', 9000);
    btn.disabled = false; btn.textContent = 'Guardar pago';
  }
}

/* =====================================================================
 * 3. PROVEEDORES
 * ===================================================================== */
const ESTADO_PAR = { TENTATIVA: ['Por programar', 'c-gray'], PROGRAMADA: ['Programada', 'c-info'], 'PAGO PARCIAL': ['Pago parcial', 'c-warn'], PAGADA: ['Pagada', 'c-ok'] };

function vistaProveedores() {
  const nombres = proveedoresTodos();
  if (!state.provSel || !nombres.includes(state.provSel)) state.provSel = nombres[0] || '';
  const prov = state.provSel;
  const botones = `<div class="actions">
      <button class="btn" type="button" data-action="prov-nuevo">+ Nuevo proveedor</button>
      ${prov ? '<button class="btn" type="button" data-action="prov-editar">Editar datos del proveedor</button>' : ''}
    </div>`;
  if (!prov) return `<header class="page-h"><div class="t"><h1>Proveedores</h1></div>${botones}</header>
    <div class="card empty">Aún no hay proveedores. Da de alta el primero con <b>+ Nuevo proveedor</b>.</div>`;

  const ocs = D().ocs.filter(o => o.PROVEEDOR === prov);
  const parts = D().activas.filter(p => p.PROVEEDOR === prov);
  const pend = parts.filter(p => p.pend > 0.009);
  const b = datosBanco(prov) || {};
  const comprometido = sum(ocs, o => o.total), pagado = sum(ocs, o => o.pag);
  const sinCot = parts.filter(p => p.pag > 0 && !p.cotId);
  const faltantes = ocs.filter(o => !o.COTIZACION).length + ocs.filter(o => !o.OC).length + parts.filter(p => !p.OC).length;
  const tentativas = ocs.reduce((a, o) => a + o.pars.filter(x => x.estado === 'TENTATIVA').length, 0);

  const pagosProv = D().pagos.filter(p => p.PROVEEDOR === prov);
  const idsEnPagos = new Set(pagosProv.map(p => p.ID_PAGO));
  const historial = [
    ...pagosProv.map(p => ({ fecha: p.FECHA_PAGO, id: p.ID_PAGO, monto: p.monto, detalle: p.DETALLE, url: p.COMPROBANTE_URL, cuenta: p.CUENTA_ORIGEN })),
    ...D().apps.filter(a => a.p.PROVEEDOR === prov && !idsEnPagos.has(a.ID_PAGO)).map(a => ({ fecha: a.FECHA_PAGO, id: a.ID_PAGO || 'Partida ' + a.PARTIDA, monto: a.monto, detalle: `${a.PARTIDA} (${a.p.OC || 'SIN OC'})`, url: a.COMPROBANTE_URL, cuenta: a.p.FORMA_PAGO }))
  ].sort((a, c) => (c.fecha || '').localeCompare(a.fecha || ''));

  const optsProv = nombres.map(n => {
    const k = D().activas.filter(p => p.PROVEEDOR === n && p.pend > 0.009).length;
    return `<option value="${esc(n)}" ${n === prov ? 'selected' : ''}>${esc(n)}${k ? ` (${k} por pagar)` : ''}</option>`;
  }).join('');

  return `
  <div style="display:flex;flex-direction:column;gap:24px">
    <div class="card card-p" style="display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end;justify-content:space-between">
      <label class="f" style="flex:1 1 320px">Proveedor
        <select data-change="prov-sel">${optsProv}</select>
      </label>
      ${botones}
    </div>

    <header class="page-h">
      <div class="t">
        <span class="eyebrow">Proveedor · cotizaciones, parcialidades y saldos</span>
        <h1>${esc(prov)}</h1>
        <span class="sub">${b.rfc ? `RFC: <b>${esc(b.rfc)}</b> · ` : ''}Banco: <b>${esc(b.banco || '—')}</b> · Cuenta: ${esc(b.cuenta || '—')} · CLABE / ref.: ${esc(b.clabe || '—')}${b.contacto ? ' · ' + esc(b.contacto) : ''}</span>
      </div>
      <div class="actions">
        <button class="btn" type="button" data-action="exportar-prov">Estado de cuenta (Excel)</button>
        <button class="btn btn-primary" type="button" data-action="oc-nueva">+ Agregar cotización / OC</button>
      </div>
    </header>

    <section class="kpis" aria-label="Resumen del proveedor">
      <div class="kpi"><span class="l">Total comprometido</span><span class="v">${money(comprometido)}</span><span class="s">${pl(ocs.length, 'cotización / OC', 'cotizaciones / OC')}</span></div>
      <div class="kpi"><span class="l">Pagado</span><span class="v v-ok">${money(pagado)}</span><span class="s">Calculado con los pagos registrados</span></div>
      <div class="kpi"><span class="l">Por pagar</span><span class="v v-warn">${money(comprometido - pagado)}</span><span class="s">${pl(tentativas, 'parcialidad por programar', 'parcialidades por programar')}</span></div>
      <div class="kpi"><span class="l">Datos faltantes</span><span class="v v-err">${faltantes}</span><span class="s">Cotización u OC sin registrar</span></div>
    </section>

    <section class="card" aria-label="Cotizaciones">
      <div class="card-h"><div class="t"><h2>Cotizaciones / OC y sus parcialidades</h2>
        <span class="sub">El pagado se suma solo con las partidas ligadas a cada cotización. Desde cada parcialidad puedes programar el pago.</span></div></div>
      ${ocs.length ? ocs.map(o => `
        <div class="cot">
          <div class="cot-h">
            <div class="cot-t">
              <div class="strong">${o.COTIZACION ? esc(o.COTIZACION) : '<span class="chip c-err">Falta n° de cotización</span>'} · ${o.OC ? esc(o.OC) : '<span class="chip c-err">Falta OC</span>'}</div>
              <div class="sub">${esc(o.CONCEPTO || 'Sin concepto')}</div>
              <div class="small">${esc(o.PROYECTO || 'Sin proyecto')} · rubro ${esc(o.RUBRO || '—')} · ${pl(o.partidas.length, 'partida ligada', 'partidas ligadas')}</div>
            </div>
            <div class="cot-n">
              <div><span class="small">Total</span><b>${money(o.total)}</b></div>
              <div><span class="small">Pagado</span><b class="v-ok">${money(o.pag)}</b></div>
              <div><span class="small">Por pagar</span><b class="${o.porPagar < -0.009 ? 'neg' : 'v-warn'}">${money(o.porPagar)}</b></div>
              ${progBar(o.pct, 'Pagado ' + o.pct + ' por ciento')}
              <button type="button" class="btn btn-sm" data-action="oc-editar" data-id="${esc(o.ID)}">Editar</button>
            </div>
          </div>
          ${o.pars.length ? `<ol class="pars">${o.pars.map(x => { const [l, c] = ESTADO_PAR[x.estado]; return `<li>
              <span class="par-n">${x.n}</span>
              <span class="par-c">${esc(x.CONCEPTO)}</span>
              <span class="par-m">${money(x.monto)}</span>
              <span class="par-f">${x.FECHA_TENTATIVA ? fechaMx(x.FECHA_TENTATIVA) : 'Sin fecha'}</span>
              <span class="chip ${c}">${l}${x.partida ? ' · partida ' + esc(x.partida.PARTIDA) : ''}</span>
              ${x.partida ? `<button type="button" class="btn-link" data-action="expediente" data-partida="${esc(x.partida.PARTIDA)}">Ver</button>`
                : `<button type="button" class="btn btn-sm btn-primary" data-action="par-programar" data-cot="${esc(o.ID)}" data-par="${esc(x.ID)}">Programar pago</button>`}
            </li>`; }).join('')}</ol>`
            : `<div class="small" style="padding:0 20px 14px">Sin parcialidades. <button type="button" class="btn-link" data-action="oc-editar" data-id="${esc(o.ID)}">Definir en cuántos pagos se liquida</button></div>`}
        </div>`).join('')
        : '<div class="empty">Sin cotizaciones / OC para este proveedor. Usa <b>+ Agregar cotización / OC</b>.</div>'}
      ${sinCot.length ? `<div style="padding:12px 20px 18px" class="small">Hay ${pl(sinCot.length, 'pago', 'pagos')} a este proveedor sin cotización ligada (${money(sum(sinCot, p => p.pag))}). Edita la cotización correspondiente y márcalos en "Partidas que se descuentan".</div>` : ''}
    </section>

    <section class="card" aria-label="Partidas pendientes">
      <div class="card-h"><div class="t"><h2>Partidas programadas pendientes de pago</h2></div>
        ${pend.length ? `<button class="btn btn-primary btn-sm" type="button" data-action="prov-pagar" data-v="${esc(prov)}">Registrar pago de este proveedor</button>` : ''}</div>
      ${pend.length ? `<div class="tbl-wrap"><table class="tbl" style="min-width:760px">
        <thead><tr><th scope="col">N°</th><th scope="col">Fecha prog.</th><th scope="col">OC</th><th scope="col">Proyecto</th><th scope="col">Concepto</th><th scope="col" class="num">Pendiente</th><th scope="col">Estatus</th></tr></thead>
        <tbody>${pend.map(p => `<tr><td class="strong">${esc(p.PARTIDA)}</td><td>${fechaMx(p.FECHA_PROGRAMADA)}</td><td>${ocCell(p.OC)}</td><td>${esc(p.PROYECTO)}</td><td class="wrap">${esc(p.CONCEPTO)}</td><td class="num strong">${money(p.pend)}</td><td>${chipEstatus(p.ESTATUS)}</td></tr>`).join('')}</tbody>
      </table></div>` : '<div class="empty">Sin partidas pendientes.</div>'}
    </section>

    <section class="card" aria-label="Historial de pagos">
      <div class="card-h"><div class="t"><h2>Historial de pagos</h2></div></div>
      ${historial.length ? `<div class="tbl-wrap"><table class="tbl" style="min-width:760px">
        <thead><tr><th scope="col">Fecha de pago</th><th scope="col">Pago</th><th scope="col">Cuenta / forma</th><th scope="col" class="num">Monto</th><th scope="col">Aplicado a</th><th scope="col">Comprobante</th></tr></thead>
        <tbody>${historial.map(h => `<tr><td>${fechaMx(h.fecha)}</td><td class="strong">${esc(h.id)}</td><td>${esc(h.cuenta || '')}</td><td class="num strong">${money(h.monto)}</td><td class="wrap">${esc(h.detalle)}</td>
          <td>${h.url ? `<a href="${esc(h.url)}" target="_blank" rel="noopener">Ver</a>` : '—'}</td></tr>`).join('')}</tbody>
      </table></div>` : '<div class="empty">Sin pagos registrados.</div>'}
    </section>
  </div>`;
}

function abrirProveedor(nombre) {
  const c = nombre ? D().provMap[nombre] : null;
  const b = nombre ? (datosBanco(nombre) || {}) : {};
  abrirDlg(nombre ? 'Datos del proveedor' : 'Nuevo proveedor', `
    <form id="formProv" style="display:flex;flex-direction:column;gap:14px">
      <input type="hidden" name="id" value="${esc(c ? c.ID : '')}">
      <input type="hidden" name="actualizar" value="${nombre ? '1' : ''}">
      <div class="grid">
        <label class="f">Nombre o razón social <input name="nombre" required value="${esc(nombre || '')}" ${nombre ? 'readonly' : ''} placeholder="Ej. ASAJI AUDIO"></label>
        <label class="f">RFC <input name="rfc" value="${esc(c ? c.RFC : '')}" maxlength="13"></label>
        <label class="f">Banco <input name="banco" list="dlBancosP" value="${esc(b.banco || '')}">${dlist('dlBancosP', uniq(Object.values(BANCOS)))}</label>
        <label class="f">Cuenta <input name="cuenta" inputmode="numeric" value="${esc(b.cuenta || '')}"></label>
        <label class="f">CLABE interbancaria <input name="clabe" inputmode="numeric" value="${esc(b.clabe || '')}" data-input="clabe"><span class="hint"></span></label>
        <label class="f">Contacto <input name="contacto" value="${esc(c ? c.CONTACTO : '')}"></label>
        <label class="f">Teléfono <input name="telefono" type="tel" value="${esc(c ? c.TELEFONO : '')}"></label>
        <label class="f">Correo <input name="correo" type="email" value="${esc(c ? c.CORREO : '')}"></label>
      </div>
      <label class="f">Notas <textarea name="notas">${esc(c ? c.NOTAS : '')}</textarea></label>
      ${nombre ? '<div class="small">El nombre no se puede cambiar porque ya está ligado a partidas y cotizaciones.</div>' : ''}
    </form>`,
    `<button type="button" class="btn" data-action="cerrar">Cancelar</button><button type="submit" form="formProv" class="btn btn-primary">Guardar proveedor</button>`);
  const cl = $('#formProv').clabe; if (cl.value) validarClabe(cl);
}

/* ---------- Cotización / OC con parcialidades ---------- */
function abrirOC(id) {
  const o = id ? D().ocById[id] : null;
  const provs = proveedoresTodos(), proys = proyectosTodos();
  state.ocForm = {
    id: o ? o.ID : '',
    pars: o ? o.pars.map(x => ({ id: x.ID, concepto: x.CONCEPTO, monto: x.monto, fecha: x.FECHA_TENTATIVA, partida: x.partida ? x.partida.PARTIDA : '' })) : []
  };
  abrirDlg(o ? 'Editar cotización / OC' : 'Agregar cotización / OC', `
    <form id="formOC" style="display:flex;flex-direction:column;gap:16px">
      <input type="hidden" name="id" value="${esc(o ? o.ID : '')}">
      <div class="paso"><span class="paso-n">1</span><div><b>Datos de la cotización</b><div class="small">Captura los datos tal como vienen en la cotización u orden de compra.</div></div></div>
      <div class="grid">
        <label class="f">Proveedor <input name="proveedor" list="dlProvOC" required value="${esc(o ? o.PROVEEDOR : state.provSel)}" data-change="oc-proveedor">${dlist('dlProvOC', provs)}</label>
        <label class="f">N° de cotización <input name="cotizacion" value="${esc(o ? o.COTIZACION : '')}" placeholder="Ej. CO_MDRK233-DP-0826_V1"></label>
        <label class="f">OC relacionada <input name="oc" value="${esc(o ? o.OC : '')}" placeholder="Ej. OC-VER-045"></label>
        <label class="f">Proyecto <input name="proyecto" list="dlProyOC" value="${esc(o ? o.PROYECTO : '')}">${dlist('dlProyOC', proys)}</label>
        <label class="f">Rubro del flujo <input name="rubro" list="dlRubOC" value="${esc(o ? o.RUBRO : '')}" placeholder="Ej. TROMPETAS, AMPLIFICADORES">${dlist('dlRubOC', rubrosTodos())}</label>
        <label class="f">Monto total de la cotización <input name="total" class="money" inputmode="decimal" required value="${esc(o ? o.total : '')}" placeholder="0.00" data-input="oc-total"></label>
      </div>
      <label class="f">Concepto del proyecto / cotización <input name="concepto" value="${esc(o ? o.CONCEPTO : '')}" placeholder="Ej. Diseño y manufactura de prototipo monitor de nivel"></label>

      <div class="paso"><span class="paso-n">2</span><div><b>¿En cuántos pagos se liquida?</b><div class="small">Captura las parcialidades que dice la cotización con su monto y fecha tentativa. Con eso se llena el calendario del flujo.</div></div></div>
      <div class="actions">
        <label class="f inline">Número de pagos <input type="number" min="0" max="24" id="ocN" value="${state.ocForm.pars.length}" data-change="oc-n" style="width:90px"></label>
        <span class="small">Rápido:</span>
        <button type="button" class="btn btn-sm" data-action="oc-preset" data-v="100">1 pago</button>
        <button type="button" class="btn btn-sm" data-action="oc-preset" data-v="50,50">50 / 50</button>
        <button type="button" class="btn btn-sm" data-action="oc-preset" data-v="60,40">60 / 40</button>
        <button type="button" class="btn btn-sm" data-action="oc-preset" data-v="30,70">30 / 70</button>
        <button type="button" class="btn btn-sm" data-action="oc-iguales">Partes iguales</button>
      </div>
      <div id="ocPars"></div>

      <div class="paso"><span class="paso-n">3</span><div><b>Partidas de la programación que se descuentan de esta cotización</b><div class="small">Marca los pagos ya programados o pagados que corresponden a esta cotización.</div></div></div>
      <div id="ocPartidas"></div>
      <label class="f">Observaciones <textarea name="observaciones">${esc(o ? o.OBSERVACIONES : '')}</textarea></label>
    </form>`,
    `${o ? '<button type="button" class="btn danger" data-action="oc-borrar" style="margin-right:auto">Eliminar cotización</button>' : ''}
     <button type="button" class="btn" data-action="cerrar">Cancelar</button><button type="submit" form="formOC" class="btn btn-primary">Guardar cotización</button>`);
  renderParsOC();
  renderPartidasOC();
}

function leerParsOC() {
  $$('#ocPars [data-par-row]').forEach((row, i) => {
    const x = state.ocForm.pars[i];
    if (!x) return;
    x.concepto = row.querySelector('[data-k="concepto"]').value;
    x.monto = num(row.querySelector('[data-k="monto"]').value);
    x.fecha = row.querySelector('[data-k="fecha"]').value;
  });
}

function renderParsOC() {
  const total = num($('#formOC').total.value);
  const pars = state.ocForm.pars;
  if (!pars.length) { $('#ocPars').innerHTML = '<div class="small">Sin parcialidades: el saldo se controla solo con las partidas ligadas.</div>'; return; }
  $('#ocPars').innerHTML = `<div class="tbl-wrap"><table class="tbl" style="min-width:720px">
    <thead><tr><th scope="col">N°</th><th scope="col">Concepto (ej. Anticipo 50%)</th><th scope="col" class="num">%</th><th scope="col" class="num">Monto</th><th scope="col">Fecha tentativa</th><th scope="col">Estado</th></tr></thead>
    <tbody>${pars.map((x, i) => `<tr data-par-row>
      <td class="strong">${i + 1}</td>
      <td><label class="sr" for="pc-${i}">Concepto ${i + 1}</label><input id="pc-${i}" data-k="concepto" value="${esc(x.concepto || '')}" placeholder="Pago ${i + 1}"></td>
      <td class="num" style="width:90px"><label class="sr" for="pp-${i}">Porcentaje ${i + 1}</label><input id="pp-${i}" data-k="pct" class="money" inputmode="decimal" value="${total ? r2(x.monto / total * 100) : ''}" data-input="oc-pct"></td>
      <td class="num" style="width:150px"><label class="sr" for="pm-${i}">Monto ${i + 1}</label><input id="pm-${i}" data-k="monto" class="money" inputmode="decimal" value="${x.monto ? x.monto.toFixed(2) : ''}" data-input="oc-monto"></td>
      <td style="width:160px"><label class="sr" for="pf-${i}">Fecha ${i + 1}</label><input id="pf-${i}" data-k="fecha" type="date" value="${esc(x.fecha || '')}"></td>
      <td class="small">${x.partida ? 'Programada · partida ' + esc(x.partida) : 'Por programar'}</td>
    </tr>`).join('')}</tbody></table></div>
    <div id="ocSuma" class="bal" style="margin-top:10px"></div>`;
  sumaParsOC();
}

function sumaParsOC() {
  const total = num($('#formOC').total.value);
  const s = r2($$('#ocPars [data-k="monto"]').reduce((a, i) => a + num(i.value), 0));
  const el = $('#ocSuma');
  if (!el) return;
  const ok = total > 0 && Math.abs(s - total) < 0.01;
  el.className = 'bal ' + (ok ? 'bal-ok' : 'bal-warn');
  el.textContent = ok ? `Las parcialidades suman ${money(s)}: cuadra con el total.` : `Las parcialidades suman ${money(s)} de ${money(total)}. ${s < total ? 'Faltan ' + money(total - s) : 'Sobran ' + money(s - total)}.`;
}

function ajustarNumPars(n) {
  leerParsOC();
  const pars = state.ocForm.pars;
  n = Math.max(0, Math.min(24, n | 0));
  while (pars.length < n) pars.push({ concepto: '', monto: 0, fecha: '' });
  pars.length = n;
  $('#ocN').value = n;
  renderParsOC();
}

function presetPars(pcts) {
  const total = num($('#formOC').total.value);
  if (!total) return toast('Captura primero el monto total', 'warn');
  ajustarNumPars(pcts.length);
  const pars = state.ocForm.pars;
  let acum = 0;
  pcts.forEach((p, i) => {
    const m = i === pcts.length - 1 ? r2(total - acum) : r2(total * p / 100);
    acum = r2(acum + m);
    pars[i].monto = m;
    if (!pars[i].concepto) pars[i].concepto = pcts.length === 1 ? 'Pago único' : (i === 0 ? `Anticipo ${p}%` : (i === pcts.length - 1 ? `Finiquito ${p}%` : `Pago ${i + 1} (${p}%)`));
  });
  renderParsOC();
}

function renderPartidasOC() {
  const f = $('#formOC');
  const prov = up(f.proveedor.value);
  const id = state.ocForm.id;
  const lista = D().activas.filter(p => p.PROVEEDOR === prov && (!p.cotId || p.cotId === id))
    .sort((a, b) => (a.FECHA_PROGRAMADA || '').localeCompare(b.FECHA_PROGRAMADA || ''));
  $('#ocPartidas').innerHTML = !prov ? '<div class="small">Elige primero el proveedor.</div>'
    : (!lista.length ? '<div class="small">Este proveedor no tiene partidas libres para ligar.</div>'
    : `<div class="tbl-wrap"><table class="tbl" style="min-width:640px">
      <thead><tr><th scope="col"><span class="sr">Ligar</span></th><th scope="col">N°</th><th scope="col">Fecha</th><th scope="col">OC</th><th scope="col">Concepto</th><th scope="col" class="num">Programado</th><th scope="col" class="num">Pagado</th></tr></thead>
      <tbody>${lista.map(p => `<tr>
        <td><label class="sr" for="lk-${esc(p.PARTIDA)}">Ligar partida ${esc(p.PARTIDA)}</label><input type="checkbox" id="lk-${esc(p.PARTIDA)}" data-liga="${esc(p.PARTIDA)}" ${id && p.cotId === id ? 'checked' : ''}></td>
        <td class="strong">${esc(p.PARTIDA)}</td><td>${fechaMx(p.FECHA_PROGRAMADA)}</td><td>${esc(p.OC || '—')}</td><td class="wrap">${esc(p.CONCEPTO)}</td>
        <td class="num">${money(p.prog)}</td><td class="num">${p.pag ? money(p.pag) : '—'}</td></tr>`).join('')}</tbody></table></div>`);
}

async function guardarOC(f, btn) {
  leerParsOC();
  const d = Object.fromEntries(new FormData(f));
  const total = num(d.total);
  const pars = state.ocForm.pars.filter(x => num(x.monto) > 0);
  if (pars.length && Math.abs(sum(pars, x => num(x.monto)) - total) > 0.01) return toast('Las parcialidades deben sumar el total de la cotización', 'warn');
  d.parcialidades = pars.map(x => ({ id: x.id, concepto: x.concepto, monto: x.monto, fecha: x.fecha }));
  d.partidas = $$('#ocPartidas [data-liga]').filter(c => c.checked).map(c => c.dataset.liga);
  btn.disabled = true;
  try {
    await api({ action: 'guardarOC', ...d });
    toast('Cotización guardada');
    state.provSel = up(d.proveedor);
    cerrarDlg();
    await cargarTodo();
  } catch (e) { toast(e.message, 'err', 9000); btn.disabled = false; }
}

/* =====================================================================
 * 4. FLUJO POR PROYECTO
 * ===================================================================== */
const SIN_RUBRO = 'SIN RUBRO';

function datosFlujo(P) {
  const pres = D().presupuesto.filter(r => r.PROYECTO === P);
  const ocs = D().ocs.filter(o => o.PROYECTO === P);
  const parts = D().activas.filter(p => p.PROYECTO === P);
  const R = x => x || SIN_RUBRO;
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
    return { rubro: rub, proveedor, proyectado, cotizado, estimado, pagado, porPagar: r2(cotizado - pagado), varianza: r2(proyectado - pagado),
      pct: pct(pagado, cotizado), sinPres: proyectado <= 0, presId: pr[0] ? pr[0].ID : '' };
  });
  return { filas };
}

/** Movimientos del proyecto para el calendario: pagado, programado pendiente y parcialidades tentativas. */
function movimientosFlujo(P) {
  const R = x => x || SIN_RUBRO;
  const m = [];
  D().apps.filter(a => a.p.PROYECTO === P).forEach(a => m.push({ fecha: a.FECHA_PAGO, monto: a.monto, tipo: 'pagado', rubro: R(a.p.rubro) }));
  D().activas.filter(p => p.PROYECTO === P && p.pend > 0.009).forEach(p => m.push({ fecha: p.FECHA_PROGRAMADA, monto: p.pend, tipo: 'programado', rubro: R(p.rubro) }));
  D().pars.filter(x => x.cot.PROYECTO === P && !x.partida).forEach(x => m.push({ fecha: x.FECHA_TENTATIVA, monto: x.monto, tipo: 'tentativo', rubro: R(x.cot.RUBRO) }));
  const ingresos = D().cobranza.filter(c => c.PROYECTO === P).map(c => ({ fecha: c.FECHA, monto: c.monto, tipo: c.ESTATUS === 'COBRADO' ? 'cobrado' : 'porcobrar' }));
  return { m, ingresos };
}

/**
 * Columnas del calendario. vista 'mes': meses del primer al último movimiento (máx. maxCols);
 * 'sem': semanas del mes elegido (o de todo el rango si rangoCompleto). Lo que no cabe va a "POSTERIOR".
 */
function calendario(P, vista, opc = {}) {
  const { filas } = datosFlujo(P);
  const { m, ingresos } = movimientosFlujo(P);
  const fechas = [...m, ...ingresos].map(x => x.fecha).filter(Boolean).sort();
  let cols = [], clave, etiqueta;
  const maxCols = opc.maxCols || (vista === 'mes' ? 18 : 26);
  if (vista === 'mes') {
    clave = f => String(f || '').slice(0, 7); etiqueta = mesCorto;
    if (fechas.length) { let ym = fechas[0].slice(0, 7); const fin = fechas[fechas.length - 1].slice(0, 7); while (ym <= fin && cols.length < maxCols) { cols.push(ym); const [y, mm] = ym.split('-').map(Number); ym = `${mm === 12 ? y + 1 : y}-${pad2(mm === 12 ? 1 : mm + 1)}`; } }
  } else {
    clave = f => lunesISO(f); etiqueta = semanaLabel;
    if (opc.rangoCompleto) { if (fechas.length) { for (let l = lunesISO(fechas[0]); l <= lunesISO(fechas[fechas.length - 1]) && cols.length < maxCols; l = masDias(l, 7)) cols.push(l); } }
    else cols = semanasDelMes(state.mes);
  }
  const ultimo = cols[cols.length - 1];
  const idx = f => { if (!f) return -1; const k = clave(f); const i = cols.indexOf(k); if (i >= 0) return i; return (opc.rangoCompleto || vista === 'mes') && ultimo && k > ultimo ? cols.length : -2; };
  const conPost = opc.rangoCompleto || vista === 'mes';
  const n = cols.length + (conPost ? 1 : 0);
  const vacio = () => Array.from({ length: n }, () => ({ total: 0, pagado: 0, programado: 0, tentativo: 0 }));
  const filasCal = filas.map(f => {
    const cel = vacio();
    m.filter(x => x.rubro === f.rubro).forEach(x => { let i = idx(x.fecha); if (i === -1 && conPost) i = cols.length; if (i < 0) return; cel[i].total = r2(cel[i].total + x.monto); cel[i][x.tipo] = r2(cel[i][x.tipo] + x.monto); });
    return { rubro: f.rubro, proveedor: f.proveedor, cel, total: sum(cel, c => c.total) };
  });
  const ing = Array.from({ length: n }, () => 0);
  ingresos.forEach(x => { let i = idx(x.fecha); if (i === -1 && conPost) i = cols.length; if (i >= 0) ing[i] = r2(ing[i] + x.monto); });
  const totCols = Array.from({ length: n }, (_, i) => sum(filasCal, f => f.cel[i].total));
  return { cols, etiquetas: cols.map(etiqueta).concat(conPost ? ['Posterior'] : []), conPost, filasCal, totCols, ing, grand: sum(totCols, v => v) };
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
  T.pp = r2(T.cot - T.pag); T.vari = r2(T.proy - T.pag);
  const cal = calendario(P, state.flujoVista);
  const cobros = D().cobranza.filter(c => c.PROYECTO === P).sort((a, b) => (a.FECHA || '').localeCompare(b.FECHA || ''));
  const cobrado = sum(cobros.filter(c => c.ESTATUS === 'COBRADO'), c => c.monto), porCobrar = sum(cobros.filter(c => c.ESTATUS !== 'COBRADO'), c => c.monto);
  const hayPost = cal.conPost && (cal.totCols[cal.cols.length] || cal.ing[cal.cols.length]);
  const nCols = cal.cols.length + (hayPost ? 1 : 0);
  const celda = c => {
    if (!c.total) return '<td class="num muted">—</td>';
    const cls = c.tentativo && !c.pagado && !c.programado ? 'cal-tent' : (c.pagado && !c.programado && !c.tentativo ? 'cal-pag' : 'cal-prog');
    const det = [c.pagado ? 'Pagado ' + money(c.pagado) : '', c.programado ? 'Programado ' + money(c.programado) : '', c.tentativo ? 'Tentativo ' + money(c.tentativo) : ''].filter(Boolean).join(' · ');
    return `<td class="num ${cls}" title="${esc(det)}">${money(c.total)}</td>`;
  };

  return `
  <header class="page-h">
    <div class="t">
      <span class="eyebrow">Flujo de dirección · montos con IVA</span>
      <h1>Flujo del proyecto ${esc(P)}</h1>
      <span class="sub">Se actualiza solo con cada solicitud, cotización, parcialidad y pago registrado.</span>
    </div>
    <div class="actions">
      <button class="btn" type="button" data-action="exportar-flujo">Descargar flujo (Excel)</button>
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

  <section class="card" aria-label="Calendario de dispersión">
    <div class="card-h">
      <div class="t"><h2>Calendario de dispersión</h2>
        <span class="sub">${state.flujoVista === 'mes' ? 'Cuánto se dispersa por rubro cada mes' : 'Cuánto se dispersa por semana en ' + esc(mesLabel(state.mes))}</span></div>
      <div class="actions">
        ${state.flujoVista === 'sem' ? `<label class="sr" for="mesFlujo">Mes</label><input id="mesFlujo" type="month" value="${esc(state.mes)}" data-change="mes" style="width:auto">` : ''}
        <div class="seg" role="group" aria-label="Vista del calendario">
          <button type="button" class="${state.flujoVista === 'mes' ? 'on' : ''}" data-action="flujo-vista" data-v="mes">Mensual</button>
          <button type="button" class="${state.flujoVista === 'sem' ? 'on' : ''}" data-action="flujo-vista" data-v="sem">Semanal</button>
        </div>
      </div>
    </div>
    <div class="leyenda"><span><i class="cal-pag"></i>Pagado</span><span><i class="cal-prog"></i>Programado (pendiente)</span><span><i class="cal-tent"></i>Parcialidad tentativa (sin programar)</span></div>
    ${nCols && (cal.filasCal.some(f => f.total) || cal.ing.some(Boolean)) ? `<div class="tbl-wrap"><table class="tbl cal" style="min-width:${260 + nCols * 125}px">
      <thead><tr><th scope="col">Rubro</th>${cal.etiquetas.slice(0, nCols).map(e => `<th scope="col" class="num">${esc(e)}</th>`).join('')}<th scope="col" class="num">Total</th></tr></thead>
      <tbody>
        ${cal.ing.some(Boolean) ? `<tr class="row-ing"><td class="strong">Cuentas por cobrar (ingresos)</td>${cal.ing.slice(0, nCols).map(v => `<td class="num">${v ? money(v) : '—'}</td>`).join('')}<td class="num strong">${money(sum(cal.ing, v => v))}</td></tr>` : ''}
        ${cal.filasCal.filter(f => f.total).map(f => `<tr><td class="strong">${esc(f.rubro)}</td>${f.cel.slice(0, nCols).map(celda).join('')}<td class="num strong">${money(f.total)}</td></tr>`).join('')}
      </tbody>
      <tfoot>
        <tr><td>Total a dispersar</td>${cal.totCols.slice(0, nCols).map(v => `<td class="num">${money(v)}</td>`).join('')}<td class="num">${money(cal.grand)}</td></tr>
        ${cal.ing.some(Boolean) ? `<tr class="row-neto"><td>Neto (cobros − pagos)</td>${cal.ing.slice(0, nCols).map((v, i) => { const n = r2(v - cal.totCols[i]); return `<td class="num ${n < 0 ? 'neg' : ''}">${money(n)}</td>`; }).join('')}<td class="num">${money(sum(cal.ing, v => v) - cal.grand)}</td></tr>` : ''}
      </tfoot>
    </table></div>` : '<div class="empty">Sin movimientos en este periodo. Agrega parcialidades con fecha tentativa en las cotizaciones del proveedor para proyectar pagos.</div>'}
  </section>

  <section class="card" aria-label="Presupuesto por rubro">
    <div class="card-h"><div class="t"><h2>Presupuesto vs. real por rubro</h2>
      <span class="sub">Varianza = proyectado − pagado. Si un rubro no tiene cotización, se toma el proyectado como cotizado (estimado).</span></div></div>
    ${filas.length ? `<div class="tbl-wrap"><table class="tbl" style="min-width:1050px">
      <thead><tr><th scope="col">Rubro</th><th scope="col">Proveedor</th><th scope="col" class="num">Proyectado</th><th scope="col" class="num">Cotizado</th>
        <th scope="col" class="num">Pagado</th><th scope="col" class="num">Por pagar</th><th scope="col" class="num">Varianza</th><th scope="col">Avance</th><th scope="col"><span class="sr">Acciones</span></th></tr></thead>
      <tbody>${filas.map(f => `<tr>
        <td class="strong">${esc(f.rubro)}</td>
        <td class="wrap">${esc(f.proveedor || '—')}</td>
        <td class="num">${money(f.proyectado)}</td>
        <td class="num">${money(f.cotizado)}${f.estimado ? ' <span class="small">(est.)</span>' : ''}</td>
        <td class="num">${money(f.pagado)}</td>
        <td class="num ${f.porPagar < -0.009 ? 'neg' : ''}">${f.porPagar < -0.009 ? '(' + money(-f.porPagar) + ')' : money(f.porPagar)}</td>
        <td class="num ${f.varianza < -0.009 ? 'neg' : ''}">${money(f.varianza)}</td>
        <td><div class="actions" style="gap:8px">${f.cotizado ? progBar(f.pct, 'Pagado ' + f.pct + ' por ciento de lo cotizado') : ''}${f.pct > 100 ? '<span class="chip c-warn">Excedido</span>' : ''}${f.sinPres ? '<span class="chip c-err">Sin presupuesto</span>' : ''}</div></td>
        <td><button type="button" class="btn-link" data-action="rubro-editar" data-id="${esc(f.presId)}" data-rubro="${esc(f.rubro)}">${f.presId ? 'Editar' : 'Presupuestar'}</button></td>
      </tr>`).join('')}</tbody>
      <tfoot><tr><td colspan="2">Totales</td><td class="num">${money(T.proy)}</td><td class="num">${money(T.cot)}</td><td class="num">${money(T.pag)}</td><td class="num">${money(T.pp)}</td><td class="num">${money(T.vari)}</td><td colspan="2"></td></tr></tfoot>
    </table></div>` : '<div class="empty">Sin rubros. Usa <b>+ Agregar rubro</b>.</div>'}
  </section>

  <section class="card" aria-label="Cuentas por cobrar">
    <div class="card-h"><div class="t"><h2>Cuentas por cobrar (ingresos del proyecto)</h2>
      <span class="sub">Cobrado ${money(cobrado)} · por cobrar ${money(porCobrar)}. Aparecen en el calendario y en el Excel del flujo.</span></div>
      <button type="button" class="btn btn-sm" data-action="cobro-nuevo">+ Agregar cobro</button></div>
    ${cobros.length ? `<div class="tbl-wrap"><table class="tbl" style="min-width:600px">
      <thead><tr><th scope="col">Fecha</th><th scope="col">Concepto</th><th scope="col" class="num">Monto</th><th scope="col">Estatus</th><th scope="col"><span class="sr">Acciones</span></th></tr></thead>
      <tbody>${cobros.map(c => `<tr><td>${fechaMx(c.FECHA)}</td><td>${esc(c.CONCEPTO)}</td><td class="num strong">${money(c.monto)}</td>
        <td>${c.ESTATUS === 'COBRADO' ? '<span class="chip c-ok">Cobrado</span>' : '<span class="chip c-info">Por cobrar</span>'}</td>
        <td><button type="button" class="btn-link" data-action="cobro-editar" data-id="${esc(c.ID)}">Editar</button></td></tr>`).join('')}</tbody>
    </table></div>` : '<div class="empty">Sin cobros registrados para este proyecto.</div>'}
  </section>`;
}

function abrirRubro(id, rubro) {
  const r = id ? D().presupuesto.find(x => x.ID === id) : null;
  abrirDlg(r ? 'Editar rubro del presupuesto' : 'Agregar rubro al presupuesto', `
    <form id="formRubro" style="display:flex;flex-direction:column;gap:14px">
      <input type="hidden" name="id" value="${esc(r ? r.ID : '')}">
      <div class="grid">
        <label class="f">Proyecto <input name="proyecto" list="dlProyR" required value="${esc(r ? r.PROYECTO : state.flujoProy)}">${dlist('dlProyR', proyectosTodos())}</label>
        <label class="f">Rubro / concepto <input name="rubro" list="dlRubR" required value="${esc(r ? r.RUBRO : (rubro && rubro !== SIN_RUBRO ? rubro : ''))}" placeholder="Ej. TRANSMISOR EAS SAME">${dlist('dlRubR', rubrosTodos(state.flujoProy))}</label>
        <label class="f">Proveedor (estimado) <input name="proveedor" value="${esc(r ? r.PROVEEDOR : '')}"></label>
        <label class="f">Monto proyectado <input name="proyectado" class="money" inputmode="decimal" required value="${esc(r ? r.proyectado : '')}" placeholder="0.00"></label>
      </div>
      <div class="small">Para que los pagos sumen a este rubro, usa el mismo rubro en las cotizaciones o en las solicitudes.</div>
    </form>`,
    `${r ? '<button type="button" class="btn danger" data-action="borrar" data-tabla="PRESUPUESTO" data-id="' + esc(r.ID) + '" style="margin-right:auto">Eliminar</button>' : ''}
     <button type="button" class="btn" data-action="cerrar">Cancelar</button><button type="submit" form="formRubro" class="btn btn-primary">Guardar</button>`);
}

function abrirCobro(id) {
  const c = id ? D().cobranza.find(x => x.ID === id) : null;
  abrirDlg(c ? 'Editar cobro' : 'Agregar cuenta por cobrar', `
    <form id="formCobro" style="display:flex;flex-direction:column;gap:14px">
      <input type="hidden" name="id" value="${esc(c ? c.ID : '')}">
      <div class="grid">
        <label class="f">Proyecto <input name="proyecto" list="dlProyC" required value="${esc(c ? c.PROYECTO : state.flujoProy)}">${dlist('dlProyC', proyectosTodos())}</label>
        <label class="f">Concepto <input name="concepto" value="${esc(c ? c.CONCEPTO : '')}" placeholder="Ej. Estimación 2"></label>
        <label class="f">Fecha (tentativa o real) <input name="fecha" type="date" required value="${esc(c ? c.FECHA : '')}"></label>
        <label class="f">Monto con IVA <input name="monto" class="money" inputmode="decimal" required value="${esc(c ? c.monto : '')}"></label>
        <label class="f">Estatus <select name="estatus">${opts(['POR COBRAR', 'COBRADO'], c ? c.ESTATUS : 'POR COBRAR')}</select></label>
      </div>
    </form>`,
    `${c ? '<button type="button" class="btn danger" data-action="borrar" data-tabla="COBRANZA" data-id="' + esc(c.ID) + '" style="margin-right:auto">Eliminar</button>' : ''}
     <button type="button" class="btn" data-action="cerrar">Cancelar</button><button type="submit" form="formCobro" class="btn btn-primary">Guardar</button>`);
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
  $('#dlgBody').scrollTop = 0;
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

async function descargar(fn) {
  toast('Preparando Excel…', 'info', 2500);
  try { await fn(); } catch (e) { toast(e.message, 'err', 9000); }
}

/* ============ EVENTOS ============ */
const acciones = {
  'cerrar': () => cerrarDlg(),
  'config': () => abrirConfig(),
  'recargar': async () => { toast('Actualizando…', 'info', 1500); if (await cargarTodo()) toast('Datos actualizados'); },
  'instalar': () => instalar(),
  // Programación
  'periodo': el => { state.periodo = el.dataset.v; if (el.dataset.v === 'dia' && !state.dia.startsWith(state.mes)) state.dia = state.mes === hoy().slice(0, 7) ? hoy() : state.mes + '-01'; render(); },
  'dia-mover': el => { state.dia = masDias(state.dia, num(el.dataset.v)); state.mes = state.dia.slice(0, 7); render(); },
  'dia-hoy': () => { state.dia = hoy(); state.mes = hoy().slice(0, 7); render(); },
  'f-semana': el => { state.semana = el.dataset.v; render(); },
  'f-proyecto': el => { state.proyecto = el.dataset.v; render(); },
  'nueva-solicitud': () => abrirSolicitud(),
  'exportar-cxp-ind': () => descargar(exportarCxPIndividual),
  'exportar-cxp-gen': () => descargar(exportarCxPGeneral),
  'expediente': el => abrirExpediente(el.dataset.partida),
  'subir-doc': el => subirDoc(el),
  'quitar-doc': el => quitarDoc(el),
  'editar-partida': () => abrirSolicitud({ partida: state.expediente }),
  'cancelar-partida': () => accionPartida('cancelar'),
  'reactivar-partida': () => accionPartida('reactivar'),
  'guardar-obs': () => accionPartida('obs'),
  // Registrar pago
  'pg-guardar': () => guardarPago(),
  'pg-limpiar': () => { state.pago = { proveedor: '', globalTocado: false }; render(); },
  // Proveedores
  'prov-nuevo': () => abrirProveedor(),
  'prov-editar': () => abrirProveedor(state.provSel),
  'prov-pagar': el => { state.pago = { proveedor: el.dataset.v, globalTocado: false }; location.hash = '#/pago'; },
  'oc-nueva': () => abrirOC(),
  'oc-editar': el => abrirOC(el.dataset.id),
  'oc-preset': el => presetPars(el.dataset.v.split(',').map(Number)),
  'oc-iguales': () => { const n = Math.max(1, num($('#ocN').value) || 1); presetPars(Array.from({ length: n }, () => 100 / n)); },
  'oc-borrar': async () => {
    if (!window.confirm('¿Eliminar esta cotización y sus parcialidades? Las partidas ligadas no se borran, solo se desligan.')) return;
    try { await api({ action: 'eliminar', tabla: 'OCS', id: state.ocForm.id }); cerrarDlg(); toast('Cotización eliminada'); await cargarTodo(); } catch (e) { toast(e.message, 'err'); }
  },
  'par-programar': el => abrirSolicitud({ proveedor: D().ocById[el.dataset.cot].PROVEEDOR, cotId: el.dataset.cot, parId: el.dataset.par }),
  'exportar-prov': () => descargar(exportarProveedor),
  // Flujo
  'flujo-proy': el => { state.flujoProy = el.dataset.v; render(); },
  'flujo-vista': el => { state.flujoVista = el.dataset.v; render(); },
  'rubro-nuevo': () => abrirRubro(),
  'rubro-editar': el => abrirRubro(el.dataset.id, el.dataset.rubro),
  'cobro-nuevo': () => abrirCobro(),
  'cobro-editar': el => abrirCobro(el.dataset.id),
  'borrar': async el => {
    if (!window.confirm('¿Eliminar este registro?')) return;
    try { await api({ action: 'eliminar', tabla: el.dataset.tabla, id: el.dataset.id }); cerrarDlg(); toast('Eliminado'); await cargarTodo(); } catch (e) { toast(e.message, 'err'); }
  },
  'exportar-flujo': () => descargar(exportarFlujo)
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
  if (t === 'mes' && e.target.value) { state.mes = e.target.value; state.semana = ''; render(); }
  if (t === 'dia' && e.target.value) { state.dia = e.target.value; state.mes = state.dia.slice(0, 7); render(); }
  if (t === 'ver-canceladas') { state.verCanceladas = e.target.checked; render(); }
  if (t === 'prov-sel') { state.provSel = e.target.value; render(); }
  if (t === 'pg-proveedor') { guardarCamposPago(); state.pago.proveedor = e.target.value; state.pago.seleccion = {}; state.pago.globalTocado = false; state.pago.global = ''; render(); }
  if (t === 'sol-proveedor') solProveedor(e.target);
  if (t === 'sol-cot') { llenarParcialidades(e.target.form, ''); aplicarCotizacion(e.target.form); }
  if (t === 'sol-par') aplicarParcialidad(e.target.form, true);
  if (t === 'oc-n') ajustarNumPars(num(e.target.value));
  if (t === 'oc-proveedor') renderPartidasOC();
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
  if (t === 'oc-pct' || t === 'oc-monto') {
    const total = num($('#formOC').total.value);
    const row = e.target.closest('tr');
    if (t === 'oc-pct' && total) row.querySelector('[data-k="monto"]').value = r2(total * num(e.target.value) / 100).toFixed(2);
    if (t === 'oc-monto' && total) row.querySelector('[data-k="pct"]').value = r2(num(e.target.value) / total * 100);
    sumaParsOC();
  }
  if (t === 'oc-total') {
    const total = num(e.target.value);
    $$('#ocPars [data-par-row]').forEach(row => { const m = num(row.querySelector('[data-k="monto"]').value); row.querySelector('[data-k="pct"]').value = total ? r2(m / total * 100) : ''; });
    sumaParsOC();
  }
});

document.addEventListener('submit', async e => {
  const f = e.target;
  const fid = f.getAttribute('id'); // (f.id no sirve: el campo oculto name="id" lo tapa)
  if (!['formSol', 'formOC', 'formRubro', 'formCobro', 'formProv', 'formConfig'].includes(fid)) return;
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
  if (fid === 'formOC') return guardarOC(f, btn);
  btn.disabled = true;
  try {
    const d = Object.fromEntries(new FormData(f));
    if (fid === 'formRubro') { await api({ action: 'guardarRubro', ...d }); toast('Rubro guardado'); state.flujoProy = up(d.proyecto); }
    if (fid === 'formCobro') { await api({ action: 'guardarCobro', ...d }); toast('Cobro guardado'); state.flujoProy = up(d.proyecto); }
    if (fid === 'formProv') { await api({ action: 'guardarProveedor', ...d }); toast('Proveedor guardado'); state.provSel = up(d.nombre); }
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
  state.dia = hoy();
  state.semana = lunesISO(hoy());
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
