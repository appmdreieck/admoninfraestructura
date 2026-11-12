/* =====================================================================
 * DESCARGAS DE EXCEL — con el mismo diseño de los archivos que usa Administración
 *   · CxP individual   → formato "PROGRAMACIÓN DE PAGOS PROYECTO …" (una hoja por proyecto)
 *   · CxP general      → pestañas VIÁTICOS · CXP PROYECTOS 8018 · AMEX
 *   · Flujo            → FLUJO MENSUAL · FLUJO SEMANAL · PROYECCIÓN · CXP <PROYECTO>
 *   · Estado de cuenta → Proy. <proveedor> · Pagos <proveedor> · Parcialidades
 * Usa ExcelJS (se descarga solo la primera vez que se exporta).
 * ===================================================================== */
'use strict';

const EXCELJS_CDN = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';

function cargarExcelJS() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  return new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = EXCELJS_CDN;
    s.onload = () => res(window.ExcelJS);
    s.onerror = () => rej(new Error('No se pudo cargar el generador de Excel (revisa tu conexión a internet).'));
    document.head.appendChild(s);
  });
}

/* ---------- Estilos y utilidades ---------- */
const XL = {
  acct: '_-"$"* #,##0.00_-;\\-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-',
  acctV: '_-\\$* #,##0.00_-;"-$"* #,##0.00_-;_-\\$* \\-??_-;_-@_-',
  mon: '\\$#,##0.00;"($"#,##0.00\\);\\-',
  monCero: '\\$#,##0.00;"($"#,##0.00\\);;',
  fecha: 'dd/mm/yyyy',
  hoyLargo: 'dddd", "d" DE "mmmm" DE "yyyy',
  borde: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
};
const xFill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const xDate = iso => { if (!iso) return null; const [y, m, d] = String(iso).slice(0, 10).split('-').map(Number); return d ? new Date(Date.UTC(y, m - 1, d)) : null; };
const xCol = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const xHoja = s => String(s).replace(/[\\/?*[\]:]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 31);
const xRef = hoja => `'${hoja.replace(/'/g, "''")}'`;
const xClabe = c => (/^\d{18}$/.test(c || '') ? c.replace(/(\d{3})(?=\d)/g, '$1 ') : (c || ''));
const xF = (formula, result) => ({ formula, result });
const xLink = (text, url) => (url ? { text: String(text), hyperlink: url } : String(text || ''));

/** Aplica estilo a una celda. s = {font, fill, border, align, fmt} */
function xs(cell, s = {}) {
  if (s.font) cell.font = s.font;
  if (s.fill) cell.fill = xFill(s.fill);
  if (s.border) cell.border = XL.borde;
  if (s.align) cell.alignment = s.align;
  if (s.fmt) cell.numFmt = s.fmt;
  return cell;
}
function xSet(ws, ref, value, s) { const c = ws.getCell(ref); if (value !== undefined) c.value = value; return xs(c, s); }
function xAnchos(ws, anchos, desde = 1) { anchos.forEach((w, i) => { ws.getColumn(desde + i).width = w; }); }
const CEN = { horizontal: 'center', vertical: 'middle', wrapText: true };
const IZQ = { horizontal: 'left', vertical: 'middle', wrapText: true };
const DER = { horizontal: 'right', vertical: 'middle', wrapText: true };
const font = (name, size, bold, argb, extra) => Object.assign({ name, size, bold: !!bold }, argb ? { color: { argb } } : {}, extra || {});

async function guardarLibro(wb, nombre) {
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = nombre;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  toast('Excel descargado: ' + nombre);
}

async function libroNuevo() {
  const ExcelJS = await cargarExcelJS();
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Pagos y Flujo · Administración';
  wb.created = new Date();
  return wb;
}

const altoTexto = (txt, porLinea, min, linea) => Math.max(min, Math.ceil(String(txt || '').length / porLinea) * linea);
const fechaArchivo = () => fechaMx(hoy()).replace(/\//g, '_');

/* =====================================================================
 * 1. CxP INDIVIDUAL  (diseño: hoja "05_11" de CXP_PROYECTOS_MARIEL)
 * ===================================================================== */
async function exportarCxPIndividual() {
  const { filas } = filasProgramacion();
  const rows = filas.filter(p => !p.cancelada);
  if (!rows.length) throw new Error('No hay partidas en el periodo seleccionado.');
  const wb = await libroNuevo();
  let titFecha, sufijo;
  if (state.periodo === 'dia') { titFecha = fechaMx(state.dia); sufijo = fechaMx(state.dia).slice(0, 5).replace('/', '_'); }
  else if (state.periodo === 'semana') { titFecha = `DEL ${fechaMx(state.semana).slice(0, 5)} AL ${fechaMx(masDias(state.semana, 6))}`; sufijo = 'SEM ' + fechaMx(state.semana).slice(0, 5).replace('/', '_'); }
  else { titFecha = mesLabel(state.mes).toUpperCase(); sufijo = mesCorto(state.mes).toUpperCase(); }

  const proyectos = uniq(rows.map(p => p.PROYECTO || 'SIN PROYECTO'));
  proyectos.forEach(pr => hojaCxPIndividual(wb, pr, rows.filter(p => (p.PROYECTO || 'SIN PROYECTO') === pr), titFecha, sufijo));
  const nombre = `CXP_INDIVIDUAL_${proyectos.length === 1 ? proyectos[0].replace(/[^A-Z0-9]+/g, '_') : 'VARIOS'}_${sufijo.replace(/[^A-Z0-9_]+/gi, '_')}.xlsx`;
  await guardarLibro(wb, nombre);
}

function hojaCxPIndividual(wb, proyecto, rows, titFecha, sufijo) {
  const ws = wb.addWorksheet(xHoja(`${proyecto} ${sufijo}`), {
    views: [{ showGridLines: false, zoomScale: 50 }],
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 } }
  });
  xAnchos(ws, [34.4, 26.2, 62.8, 29.5, 34.2, 51.6, 37.6, 79.1, 24]);
  const AZUL = 'FF080B96';
  ws.mergeCells('E1:I1'); ws.mergeCells('E2:I2');
  xSet(ws, 'E1', `PROGRAMACIÓN DE PAGOS PROYECTO ${proyecto} ${titFecha}`, { font: font('Arial', 22, true, AZUL), align: { horizontal: 'center', vertical: 'middle' } });
  xSet(ws, 'E2', '', { font: font('Arial', 16, true, AZUL), align: { horizontal: 'center', vertical: 'middle' } });
  ws.getRow(1).height = 28; ws.getRow(2).height = 20; ws.getRow(3).height = 15; ws.getRow(4).height = 51;

  const heads = ['FECHA DE PAGO REQUERIDA', 'PROYECTO', 'PROVEEDOR', 'BANCO', 'CTA. DEPOSITO', 'CLABE INTERBANCARIA/LINEA DE CAPTURA', 'PROGRAMADO', 'CONCEPTO', 'OC'];
  heads.forEach((h, i) => xSet(ws, xCol(i + 1) + '4', h, { font: font('Arial', i === 0 ? 20 : 16, true), fill: i === 0 ? 'FF00B0F0' : 'FFEEECE1', border: 1, align: CEN }));

  let r = 5;
  rows.forEach(p => {
    const vals = [xDate(p.FECHA_PROGRAMADA), p.PROYECTO, p.PROVEEDOR, p.BANCO, p.CUENTA_DEPOSITO, xClabe(p.CLABE_REFERENCIA), p.prog, p.CONCEPTO, p.OC || 'SIN OC'];
    vals.forEach((v, i) => {
      const c = ws.getCell(xCol(i + 1) + r);
      c.value = v === '' ? null : v;
      if (i === 0) xs(c, { font: font('Arial', 20, true, 'FF4BACC6'), border: 1, align: CEN, fmt: XL.fecha });
      else if (i === 6) xs(c, { font: font('Arial', 18), border: 1, align: DER, fmt: XL.acct });
      else xs(c, { font: font('Arial', 18), border: 1, align: [1, 3, 4, 5, 8].includes(i) ? CEN : IZQ });
    });
    ws.getRow(r).height = Math.max(altoTexto(p.CONCEPTO, 40, 60, 24), altoTexto(p.PROVEEDOR, 30, 60, 24));
    r++;
  });
  const t = r + 1;
  xSet(ws, 'F' + t, 'Total', { font: font('Calibri', 22, true), align: { horizontal: 'right', vertical: 'middle' } });
  xSet(ws, 'G' + t, xF(`SUM(G5:G${r - 1})`, sum(rows, p => p.prog)), { font: font('Calibri', 22, true), fill: 'FFFFFF00', border: 1, fmt: XL.acct });
  ws.getRow(t).height = 30;
}

/* =====================================================================
 * 2. CxP GENERAL  (diseño: CXP_PARA_RUBI — VIÁTICOS · CXP PROYECTOS 8018 · AMEX)
 * ===================================================================== */
async function exportarCxPGeneral() {
  const todas = D().activas.slice().sort((a, b) => (a.FECHA_PROGRAMADA || '').localeCompare(b.FECHA_PROGRAMADA || '') || num(a.PARTIDA) - num(b.PARTIDA));
  if (!todas.length) throw new Error('Aún no hay partidas registradas.');
  const wb = await libroNuevo();
  hojaViaticos(wb, todas.filter(esViaticos));
  hojaCxP8018(wb, todas.filter(p => !['EFECTIVALE', 'AMEX'].includes(up(p.FORMA_PAGO))));
  hojaAmex(wb, todas.filter(p => up(p.FORMA_PAGO) === 'AMEX'));
  await guardarLibro(wb, `CXP_GENERAL_${fechaArchivo()}.xlsx`);
}

function hojaViaticos(wb, rows) {
  const ws = wb.addWorksheet('VIÁTICOS', { views: [{ state: 'frozen', ySplit: 1, zoomScale: 80 }] });
  xAnchos(ws, [7.6, 13.1, 13, 21, 15.5, 22.5, 12.8, 11.3, 13.3, 11.3, 11.9, 39.5, 38, 20, 60, 3, 18]);
  const heads = ['NO ZIP', 'FECHA DE SOLICITUD', 'FECHA DEPÓSITO', 'QUIEN SOLICITA', 'PROYECTO', 'CONCEPTO', 'MONTO SOLICITADO', 'COMPROBACIÓN ENTREGADA', 'FORMA DE PAGO', 'FECHA DE COMPROBACIÓN', 'VÍA', 'OBSERVACIONES', 'COMPROBACIÓN(ES)', 'PERIODO COMPROBADO', 'OBSERVACIONES DE LA REVISIÓN'];
  heads.forEach((h, i) => xSet(ws, xCol(i + 1) + '1', h, { font: font('Calibri', i === 0 ? 11 : 8, true, 'FFFFFFFF'), fill: i === 0 ? 'FF31859B' : (i === 6 ? 'FF7030A0' : 'FF205867'), border: 1, align: CEN }));
  ws.getRow(1).height = 30;
  rows.forEach((p, k) => {
    const r = k + 2;
    const docs = lineas(p.FACTURA_PDF_URL).concat(lineas(p.FACTURA_XML_URL), lineas(p.OTROS_URL)).filter(Boolean);
    const comprobado = docs.length > 0;
    const fondo = comprobado ? 'FFC2D69B' : (p.pag > 0 ? 'FFFF0000' : null);
    const vals = [num(p.PARTIDA), xDate(String(p.FECHA_REGISTRO || '').slice(0, 10)), xDate(p.FECHA_PAGO), p.PROVEEDOR, p.PROYECTO, p.rubro || 'VIÁTICOS', p.prog,
      comprobado ? 'Sí' : 'No', p.FORMA_PAGO, null, comprobado ? 'DRIVE' : '', p.OBSERVACIONES || p.CONCEPTO,
      xLink(comprobado ? `Carpeta ${p.PARTIDA} (${docs.length} archivo${docs.length > 1 ? 's' : ''})` : `Carpeta ${p.PARTIDA}`, p.CARPETA_URL), '', ''];
    vals.forEach((v, i) => {
      const c = ws.getCell(xCol(i + 1) + r);
      c.value = v === '' ? null : v;
      xs(c, { font: font('Calibri', i === 0 ? 11 : 8, i === 0, i === 12 ? 'FF0563C1' : null, i === 12 ? { underline: true } : null), border: 1, align: CEN,
        fill: fondo || undefined, fmt: i === 0 ? '0' : ([1, 2].includes(i) ? XL.fecha : (i === 6 ? XL.acctV : undefined)) });
    });
    ws.getRow(r).height = altoTexto(p.OBSERVACIONES || p.CONCEPTO, 45, 24, 12);
  });
  if (rows.length) ws.autoFilter = `A1:O${rows.length + 1}`;
  xSet(ws, 'Q1', 'Colores', { font: font('Calibri', 11, true), align: CEN });
  [['Comprobado', 'FFC2D69B'], ['Sin comprobación', 'FFFF0000'], ['Por depositar', null]].forEach(([t, f], i) =>
    xSet(ws, 'Q' + (i + 2), t, { font: font('Calibri', 8), fill: f || undefined, border: 1, align: { vertical: 'middle', wrapText: true } }));
}

function hojaCxP8018(wb, rows) {
  const ws = wb.addWorksheet('CXP PROYECTOS 8018', { views: [{ state: 'frozen', ySplit: 1, zoomScale: 80, showGridLines: false }] });
  xAnchos(ws, [1.5, 13.1, 13, 18.3, 10.3, 13.5, 47.5, 16.1, 12.1, 10.8, 12.7, 20.5, 14, 11.5, 11.5, 14.5, 11.5]);
  const heads = ['FECHA DE PROGRAMACIÓN', 'PROYECTO', 'PROVEEDOR', 'FORMA DE PAGO', 'MONTO PROGRAMADO', 'CONCEPTO', 'OC', 'MONTO PAGADO', 'POR PAGAR', 'ESTATUS DE PAGO', 'OBSERVACIONES', 'NO. CARPETA ZIP FACTURA', 'FECHA DE PAGO', 'MES', 'FACTURA', 'COMPLEMENTO'];
  heads.forEach((h, i) => {
    const col = xCol(i + 2);
    const esZip = col === 'M';
    xSet(ws, col + '1', h, { font: esZip ? font('Calibri', 12, true, 'FF632423') : font('Calibri', 9, false, 'FFFFFFFF'),
      fill: esZip ? 'FFFFFF00' : ({ F: 'FF31859B', I: 'FFE36C09' }[col] || 'FF002060'), border: 1, align: CEN });
  });
  ws.getRow(1).height = 59.25;
  rows.forEach((p, k) => {
    const r = k + 2;
    const fact = lineas(p.FACTURA_PDF_URL).filter(Boolean)[0] || lineas(p.FACTURA_XML_URL).filter(Boolean)[0];
    const comp = lineas(p.COMPLEMENTO_URL).filter(Boolean)[0];
    const mes = p.FECHA_PAGO ? parseISO(p.FECHA_PAGO).toLocaleDateString('es-MX', { month: 'long' }) : '';
    const C = (col, v, s) => xSet(ws, col + r, v, Object.assign({ font: font('Calibri', 9), border: 1, align: CEN }, s));
    C('B', xDate(p.FECHA_PROGRAMADA), { fmt: XL.fecha });
    C('C', p.PROYECTO);
    C('D', p.PROVEEDOR, { align: IZQ });
    C('E', p.FORMA_PAGO);
    C('F', p.prog, { fmt: XL.acctV, align: DER });
    C('G', p.CONCEPTO, { align: IZQ });
    C('H', p.OC || 'SIN OC');
    C('I', p.pag || null, { fmt: XL.acctV, align: DER });
    C('J', xF(`F${r}-I${r}`, p.pend), { fmt: XL.acctV, align: DER });
    C('K', xF(`IF(ROUND(J${r},2)<=0,"PAGADO","X PAGAR")`, p.pend <= 0.009 ? 'PAGADO' : 'X PAGAR'));
    C('L', p.OBSERVACIONES || null, { align: IZQ });
    C('M', xLink(p.PARTIDA, p.CARPETA_URL), { font: font('Calibri', 12, true, 'FF632423', { underline: true }) });
    C('N', xDate(p.FECHA_PAGO), { fmt: XL.fecha });
    C('O', xF(`IF(N${r}="","",TEXT(N${r},"[$-080A]mmmm"))`, mes));
    C('P', fact ? xLink('Ver factura', fact) : null, { font: font('Calibri', 9, false, 'FF0563C1', { underline: true }) });
    C('Q', comp ? xLink('Ver complemento', comp) : null, { font: font('Calibri', 9, false, 'FF0563C1', { underline: true }) });
    ws.getRow(r).height = altoTexto(p.CONCEPTO, 55, 30, 12);
  });
  const last = Math.max(rows.length + 1, 2);
  ws.autoFilter = `B1:Q${last}`;
  ws.addConditionalFormatting({ ref: `B2:Q${last}`, rules: [
    { type: 'expression', priority: 1, formulae: ['$K2="PAGADO"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFF5C345' } } } },
    { type: 'expression', priority: 2, formulae: ['$K2="X PAGAR"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF76923C' } } } }
  ] });
}

function hojaAmex(wb, rows) {
  const ws = wb.addWorksheet('AMEX', { views: [{ showGridLines: false }] });
  xAnchos(ws, [18, 18, 66.8, 14]);
  ['FECHA DEL CARGO', 'PROVEEDOR', 'CONCEPTO', 'MONTO CARGADO'].forEach((h, i) =>
    xSet(ws, xCol(i + 1) + '1', h, { font: font('Calibri', 11, true, 'FFFFFFFF'), fill: 'FF205867', border: 1, align: CEN }));
  rows.forEach((p, k) => {
    const r = k + 2;
    [[xDate(p.FECHA_PAGO || p.FECHA_PROGRAMADA), XL.fecha], [p.PROVEEDOR], [p.CONCEPTO], [p.pag || p.prog, XL.acct]].forEach(([v, f], i) =>
      xSet(ws, xCol(i + 1) + r, v, { font: font('Calibri', 11), border: 1, align: CEN, fmt: f }));
    ws.getRow(r).height = altoTexto(p.CONCEPTO, 70, 30, 15);
  });
}

/* =====================================================================
 * 3. FLUJO  (diseño: FLUJO - PROVEEDORES VERACRUZ)
 * ===================================================================== */
const AZ = { titulo: 'FF1F3864', cobrar: 'FF548DD4', concepto: 'FF215868', pagar: 'FF365F91', fila: 'FFD9E2F3', ingreso: 'FFEDEDDE', proy: 'FF833C00', semana: 'FF1F5C5C', semanaLbl: 'FF31859B', verde: 'FF4F6228', verdeClaro: 'FFEBF1DE', cxp: 'FF1F497D' };
const BLANCO = 'FFFFFFFF';

async function exportarFlujo() {
  const P = state.flujoProy;
  if (!P) throw new Error('Elige un proyecto.');
  const wb = await libroNuevo();
  const { filas } = datosFlujo(P);
  const hojaCxp = xHoja('CXP ' + P);
  const ref = xRef(hojaCxp);
  const anio = hoy().slice(0, 4);
  const act = parseISO(hoy()).toLocaleDateString('es-MX', { day: '2-digit', month: 'long', year: 'numeric' }).toUpperCase();
  const pagadoRubro = {};
  D().apps.filter(a => a.p.PROYECTO === P).forEach(a => { const k = a.p.rubro || SIN_RUBRO; pagadoRubro[k] = r2((pagadoRubro[k] || 0) + a.monto); });
  const cobros = D().cobranza.filter(c => c.PROYECTO === P);
  const ingTotal = sum(cobros, c => c.monto), ingCobrado = sum(cobros.filter(c => c.ESTATUS === 'COBRADO'), c => c.monto);

  const T = hojaFlujoMensual(wb, P, filas, ref, pagadoRubro, { anio, act, ingTotal, ingCobrado });
  hojaFlujoSemanal(wb, P, filas, T, { anio, ingTotal, ingCobrado, pagadoRubro });
  hojaProyeccion(wb, P, filas, pagadoRubro, anio);
  hojaCxpProyecto(wb, P, hojaCxp);
  await guardarLibro(wb, `FLUJO_${P.replace(/[^A-Z0-9]+/g, '_')}_${fechaArchivo()}.xlsx`);
}

/** Celda de encabezado del flujo. */
const hF = (ws, ref, v, fill, extra) => xSet(ws, ref, v, Object.assign({ font: font('Arial', 9, true, BLANCO), fill, border: 1, align: CEN }, extra || {}));

function hojaFlujoMensual(wb, P, filas, ref, pagadoRubro, o) {
  const ws = wb.addWorksheet('FLUJO MENSUAL', { views: [{ state: 'frozen', xSplit: 2, ySplit: 10, zoomScale: 80, showGridLines: false }],
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
  const cal = calendario(P, 'mes', { maxCols: 12 });
  if (!cal.cols.length) { cal.cols = [hoy().slice(0, 7)]; cal.filasCal.forEach(f => f.cel.unshift({ total: 0 })); cal.ing.unshift(0); cal.totCols.unshift(0); }
  const nm = cal.cols.length;
  const c0 = 10, cPost = c0 + nm, cTot = cPost + 1;
  const L = xCol, lPost = L(cPost), lTot = L(cTot), lIni = L(c0), lFin = L(cPost);
  xAnchos(ws, [24.3, 19.4, 15.1, 15.5, 16, 15.5, 13.8, 11.7, 2.3]);
  for (let c = c0; c < cPost; c++) ws.getColumn(c).width = 15;
  ws.getColumn(cPost).width = 16; ws.getColumn(cTot).width = 16;

  ws.mergeCells('A1:H1');
  xSet(ws, 'A1', `FLUJO MENSUAL ${o.anio} — PROYECTO ${P} - ÚLTIMA ACTUALIZACIÓN ${o.act}`, { font: font('Arial', 14, true, AZ.titulo), align: { horizontal: 'left', vertical: 'middle' } });
  ws.mergeCells(`J1:${L(Math.max(13, c0 + 3))}1`);
  xSet(ws, 'J1', xF('TODAY()'), { font: font('Arial', 14, true), align: { horizontal: 'center' }, fmt: XL.hoyLargo });
  ws.getRow(1).height = 19.5; ws.getRow(2).height = 7.5; ws.getRow(8).height = 5.25;
  ws.mergeCells('B4:H4'); hF(ws, 'B4', 'MONTOS CON IVA', AZ.titulo);

  // Cuentas por cobrar
  ws.mergeCells(`B5:${lFin}5`); hF(ws, 'B5', 'CUENTAS POR COBRAR', AZ.cobrar);
  ws.mergeCells(`${lTot}5:${lTot}6`); hF(ws, lTot + '5', 'TOTALES', AZ.cobrar);
  ws.mergeCells('B6:C6');
  [['B6', 'CONCEPTO'], ['D6', 'PROYECTADO'], ['E6', 'COBRADO'], ['F6', 'POR COBRAR'], ['G6', ''], ['H6', '% AVANCE']].forEach(([r, v]) => hF(ws, r, v, AZ.concepto));
  cal.cols.forEach((ym, i) => hF(ws, L(c0 + i) + '6', xDate(ym + '-01'), AZ.concepto, { fmt: 'mmm-yy' }));
  hF(ws, lPost + '6', 'POSTERIOR', AZ.proy);
  ws.getRow(6).height = 22.5;
  const sI = { font: font('Arial', 10, true), fill: AZ.ingreso, border: 1, align: CEN, fmt: XL.mon };
  xSet(ws, 'A7', 'INGRESOS', sI); ws.mergeCells('B7:C7'); xSet(ws, 'B7', 'CUENTAS POR COBRAR', sI);
  xSet(ws, 'D7', o.ingTotal, sI); xSet(ws, 'E7', o.ingCobrado, sI); xSet(ws, 'F7', xF('D7-E7', r2(o.ingTotal - o.ingCobrado)), sI);
  xSet(ws, 'G7', null, sI); xSet(ws, 'H7', xF('IFERROR(E7/D7,0)', o.ingTotal ? o.ingCobrado / o.ingTotal : 0), Object.assign({}, sI, { fmt: '0%' }));
  for (let i = 0; i <= nm; i++) xSet(ws, L(c0 + i) + '7', cal.ing[i] || null, Object.assign({}, sI, { fmt: XL.monCero }));
  xSet(ws, lTot + '7', xF(`SUM(${lIni}7:${lFin}7)`, sum(cal.ing, v => v)), Object.assign({}, sI, { font: font('Arial', 10, true, 'FF0000FF') }));
  ws.getRow(7).height = 18;

  // Cuentas por pagar
  ws.mergeCells(`B9:${lFin}9`); hF(ws, 'B9', 'CUENTAS POR PAGAR', AZ.pagar);
  ws.mergeCells(`${lTot}9:${lTot}10`); hF(ws, lTot + '9', 'TOTALES', AZ.pagar);
  ['CONCEPTO', 'PROVEEDOR', 'PROYECTADO', 'COTIZADO', 'PAGADO', 'POR PAGAR', 'VARIANZA', '% AVANCE PAGO'].forEach((h, i) => hF(ws, L(i + 1) + '10', h, AZ.titulo));
  cal.cols.forEach((ym, i) => hF(ws, L(c0 + i) + '10', xF(`${L(c0 + i)}6`, xDate(ym + '-01')), AZ.titulo, { fmt: 'mmm-yy' }));
  hF(ws, lPost + '10', 'POSTERIOR', AZ.proy);
  ws.getRow(10).height = 30;

  const r0 = 11;
  filas.forEach((f, i) => {
    const r = r0 + i;
    const cel = cal.filasCal[i].cel;
    const pag = pagadoRubro[f.rubro] || 0;
    const s = (extra) => Object.assign({ font: font('Arial', 9), fill: AZ.fila, border: 1, align: DER, fmt: XL.mon }, extra || {});
    xSet(ws, 'A' + r, f.rubro, f.sinPres ? s({ font: font('Arial', 9, true, BLANCO), fill: 'FFFF0000', align: IZQ }) : s({ font: font('Arial', 9, true), align: IZQ }));
    xSet(ws, 'B' + r, f.proveedor || null, s({ font: font('Arial', 9, true), align: IZQ }));
    xSet(ws, 'C' + r, f.proyectado, s());
    xSet(ws, 'D' + r, f.cotizado, s());
    xSet(ws, 'E' + r, xF(`SUMIFS(${ref}!$F:$F,${ref}!$A:$A,$A${r})`, pag), s());
    xSet(ws, 'F' + r, xF(`D${r}-E${r}`, r2(f.cotizado - pag)), s());
    xSet(ws, 'G' + r, xF(`C${r}-E${r}`, r2(f.proyectado - pag)), s());
    xSet(ws, 'H' + r, xF(`IFERROR(E${r}/D${r},0)`, f.cotizado ? pag / f.cotizado : 0), s({ align: CEN, fmt: '0%' }));
    for (let k = 0; k <= nm; k++) xSet(ws, L(c0 + k) + r, cel[k] && cel[k].total ? cel[k].total : null, s({ align: CEN, fmt: XL.monCero }));
    xSet(ws, lTot + r, xF(`SUM(${lIni}${r}:${lFin}${r})`, sum(cel, c => c.total || 0)), s({ font: font('Arial', 9, true), fmt: XL.monCero }));
    ws.getRow(r).height = altoTexto(f.rubro, 24, 24, 12);
  });
  const T = r0 + Math.max(filas.length, 1);
  const rf = T - 1;
  const sT = { font: font('Arial', 9, true, BLANCO), fill: AZ.titulo, border: 1, align: DER, fmt: XL.mon };
  xSet(ws, 'A' + T, null, sT); xSet(ws, 'B' + T, 'TOTALES', Object.assign({}, sT, { align: CEN }));
  const tot = { C: sum(filas, f => f.proyectado), D: sum(filas, f => f.cotizado), E: sum(Object.values(pagadoRubro), v => v) };
  ['C', 'D', 'E'].forEach(c => xSet(ws, c + T, xF(`SUM(${c}${r0}:${c}${rf})`, tot[c]), sT));
  xSet(ws, 'F' + T, xF(`SUM(F${r0}:F${rf})`, r2(tot.D - tot.E)), sT);
  xSet(ws, 'G' + T, xF(`SUM(G${r0}:G${rf})`, r2(tot.C - tot.E)), sT);
  xSet(ws, 'H' + T, xF(`IFERROR(E${T}/C${T},0)`, tot.C ? tot.E / tot.C : 0), Object.assign({}, sT, { align: CEN, fmt: '0%' }));
  for (let k = 0; k <= nm + 1; k++) { const l = L(c0 + k); xSet(ws, l + T, xF(`SUM(${l}${r0}:${l}${rf})`, k <= nm ? (cal.totCols[k] || 0) : cal.grand), sT); }
  ws.getRow(T).height = 20;

  xSet(ws, 'B' + (T + 2), 'CUADRE (meses vs. cotizado):', { font: font('Arial', 9, true) });
  xSet(ws, 'C' + (T + 2), xF(`SUM(${lIni}${T}:${lFin}${T})-D${T}`, r2(cal.grand - tot.D)), { font: font('Arial', 9), fmt: XL.mon });
  xSet(ws, 'B' + (T + 3), 'Meses = pagado (fecha de pago) + programado pendiente (fecha programada) + parcialidades tentativas (fecha tentativa).', { font: font('Arial', 8, false, 'FF6A6688', { italic: true }) });
  xSet(ws, 'C' + (T + 5), 'Sin IVA', { font: font('Arial', 9, true) });
  xSet(ws, 'D' + (T + 5), xF(`D${T}/1.16`, r2(tot.D / 1.16)), { font: font('Arial', 9), fmt: XL.mon });
  return T;
}

function hojaFlujoSemanal(wb, P, filas, Tm, o) {
  const ws = wb.addWorksheet('FLUJO SEMANAL', { views: [{ state: 'frozen', xSplit: 2, ySplit: 11, zoomScale: 80, showGridLines: false }],
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
  const cal = calendario(P, 'sem', { rangoCompleto: true, maxCols: 26 });
  if (!cal.cols.length) { cal.cols = [lunesISO(hoy())]; cal.filasCal.forEach(f => f.cel.unshift({ total: 0 })); cal.ing.unshift(0); cal.totCols.unshift(0); }
  const ns = cal.cols.length;
  const c0 = 9, cPost = c0 + ns, cTot = cPost + 1;
  const L = xCol, lIni = L(c0), lFin = L(cPost), lTot = L(cTot), lPost = L(cPost);
  const FM = xRef('FLUJO MENSUAL');
  xAnchos(ws, [24.3, 30.7, 18, 16, 16, 16, 11, 2.3]);
  for (let c = c0; c < cPost; c++) ws.getColumn(c).width = 14;
  ws.getColumn(cPost).width = 16; ws.getColumn(cTot).width = 16;

  ws.mergeCells('A1:G1');
  xSet(ws, 'A1', `FLUJO SEMANAL ${o.anio} — PROYECTO ${P}`, { font: font('Arial', 14, true, AZ.titulo), align: { horizontal: 'left', vertical: 'middle' } });
  ws.getRow(2).height = 7.5; ws.getRow(8).height = 5.25;
  ws.mergeCells('B4:G4'); hF(ws, 'B4', 'MONTOS CON IVA', AZ.titulo);
  ws.mergeCells(`B5:${lFin}5`); hF(ws, 'B5', 'CUENTAS POR COBRAR', AZ.cobrar);
  hF(ws, lTot + '5', 'TOTAL', AZ.titulo);
  ws.mergeCells('B6:C6');
  [['B6', 'CONCEPTO'], ['D6', 'PROYECTADO'], ['E6', 'COBRADO'], ['F6', 'POR COBRAR'], ['G6', '% AVANCE']].forEach(([r, v]) => hF(ws, r, v, AZ.concepto));
  cal.cols.forEach((l, i) => hF(ws, L(c0 + i) + '6', xDate(l), AZ.semana, { fmt: 'dd/mm' }));
  hF(ws, lPost + '6', 'POSTERIOR', AZ.proy); hF(ws, lTot + '6', '', AZ.titulo);
  const sI = { font: font('Arial', 10, true), fill: AZ.ingreso, border: 1, align: CEN, fmt: XL.mon };
  xSet(ws, 'A7', 'INGRESOS', sI); ws.mergeCells('B7:C7'); xSet(ws, 'B7', 'CUENTAS POR COBRAR', sI);
  xSet(ws, 'D7', xF(`${FM}!D7`, o.ingTotal), sI); xSet(ws, 'E7', xF(`${FM}!E7`, o.ingCobrado), sI);
  xSet(ws, 'F7', xF('D7-E7', r2(o.ingTotal - o.ingCobrado)), sI);
  xSet(ws, 'G7', xF('IFERROR(E7/D7,0)', o.ingTotal ? o.ingCobrado / o.ingTotal : 0), Object.assign({}, sI, { fmt: '0%' }));
  for (let i = 0; i <= ns; i++) xSet(ws, L(c0 + i) + '7', cal.ing[i] || null, Object.assign({}, sI, { fmt: XL.monCero }));
  xSet(ws, lTot + '7', xF(`SUM(${lIni}7:${lFin}7)`, sum(cal.ing, v => v)), sI);

  ws.mergeCells(`B9:${lFin}9`); hF(ws, 'B9', 'CUENTAS POR PAGAR', AZ.pagar);
  cal.cols.forEach((l, i) => hF(ws, L(c0 + i) + '10', 'SEMANA', AZ.semanaLbl));
  ['CONCEPTO', 'PROVEEDOR', 'PROYECTADO', 'COTIZADO', 'PAGADO', 'POR PAGAR', '% AVANCE'].forEach((h, i) => hF(ws, L(i + 1) + '11', h, AZ.titulo));
  cal.cols.forEach((l, i) => hF(ws, L(c0 + i) + '11', xF(`${L(c0 + i)}6`, xDate(l)), AZ.titulo, { fmt: 'dd/mm' }));
  hF(ws, lPost + '11', 'POSTERIOR', AZ.proy); hF(ws, lTot + '11', 'TOTAL', AZ.titulo);
  ws.getRow(11).height = 30;

  const r0 = 12;
  filas.forEach((f, i) => {
    const r = r0 + i, rm = 11 + i; // renglón equivalente en FLUJO MENSUAL
    const pag = o.pagadoRubro[f.rubro] || 0;
    const cel = cal.filasCal[i].cel;
    const s = extra => Object.assign({ font: font('Arial', 9), fill: AZ.fila, border: 1, align: DER, fmt: XL.mon }, extra || {});
    xSet(ws, 'A' + r, xF(`${FM}!A${rm}`, f.rubro), s({ font: font('Arial', 9, true), align: IZQ }));
    xSet(ws, 'B' + r, xF(`${FM}!B${rm}`, f.proveedor || ''), s({ font: font('Arial', 9, true), align: IZQ }));
    xSet(ws, 'C' + r, xF(`${FM}!C${rm}`, f.proyectado), s());
    xSet(ws, 'D' + r, xF(`${FM}!D${rm}`, f.cotizado), s());
    xSet(ws, 'E' + r, xF(`${FM}!E${rm}`, pag), s());
    xSet(ws, 'F' + r, xF(`${FM}!F${rm}`, r2(f.cotizado - pag)), s());
    xSet(ws, 'G' + r, xF(`${FM}!H${rm}`, f.cotizado ? pag / f.cotizado : 0), s({ align: CEN, fmt: '0%' }));
    for (let k = 0; k <= ns; k++) xSet(ws, L(c0 + k) + r, cel[k] && cel[k].total ? cel[k].total : null, s({ align: CEN, fmt: XL.monCero }));
    xSet(ws, lTot + r, xF(`SUM(${lIni}${r}:${lFin}${r})`, sum(cel, c => c.total || 0)), s({ font: font('Arial', 9, true), fmt: XL.monCero }));
    ws.getRow(r).height = altoTexto(f.rubro, 24, 24, 12);
  });
  const T = r0 + Math.max(filas.length, 1), rf = T - 1;
  const sT = { font: font('Arial', 9, true, BLANCO), fill: AZ.titulo, border: 1, align: DER, fmt: XL.mon };
  xSet(ws, 'A' + T, 'TOTALES', Object.assign({}, sT, { align: CEN })); xSet(ws, 'B' + T, null, sT);
  const tP = sum(Object.values(o.pagadoRubro), v => v), tD = sum(filas, f => f.cotizado);
  [['C', sum(filas, f => f.proyectado)], ['D', tD], ['E', tP], ['F', r2(tD - tP)]].forEach(([c, v]) => xSet(ws, c + T, xF(`SUM(${c}${r0}:${c}${rf})`, v), sT));
  xSet(ws, 'G' + T, xF(`${FM}!H${Tm}`, sum(filas, f => f.proyectado) ? tP / sum(filas, f => f.proyectado) : 0), Object.assign({}, sT, { align: CEN, fmt: '0%' }));
  for (let k = 0; k <= ns + 1; k++) { const l = L(c0 + k); xSet(ws, l + T, xF(`SUM(${l}${r0}:${l}${rf})`, k <= ns ? (cal.totCols[k] || 0) : cal.grand), sT); }
  xSet(ws, 'A' + (T + 2), 'CUADRE (semanas vs. cotizado):', { font: font('Arial', 9, true) });
  xSet(ws, 'C' + (T + 2), xF(`SUM(${lIni}${T}:${lFin}${T})-D${T}`, r2(cal.grand - tD)), { font: font('Arial', 9), fmt: XL.mon });
}

function hojaProyeccion(wb, P, filas, pagadoRubro, anio) {
  const ws = wb.addWorksheet('PROYECCIÓN', { views: [{ state: 'frozen', ySplit: 4, zoomScale: 80, showGridLines: false }] });
  xAnchos(ws, [24, 28, 13, 19.2, 16.4, 18, 16, 10]);
  xSet(ws, 'A1', `PROYECCIÓN ${anio} — PROYECTO ${P}`, { font: font('Arial', 14, true, AZ.titulo), align: { horizontal: 'left' } });
  ws.mergeCells('A2:D2');
  xSet(ws, 'A2', xF('TODAY()'), { font: font('Arial', 14, true), align: { horizontal: 'left' }, fmt: XL.hoyLargo });
  ws.mergeCells('A3:H3');
  xSet(ws, 'A3', 'MONTOS CON IVA', { font: font('Arial', 11, true, BLANCO), fill: AZ.concepto, align: CEN });
  ws.getRow(3).height = 24; ws.getRow(4).height = 30;
  ['CONCEPTO', 'PROVEEDOR', 'CANTIDAD', 'PROYECTADO TOTAL', 'COTIZADO', 'PAGADO', 'POR PAGAR', '% AVANCE'].forEach((h, i) =>
    hF(ws, xCol(i + 1) + '4', h, i === 3 ? AZ.verde : (i === 6 ? AZ.pagar : AZ.titulo)));
  const r0 = 5;
  filas.forEach((f, i) => {
    const r = r0 + i;
    const pag = pagadoRubro[f.rubro] || 0;
    const s = extra => Object.assign({ font: font('Arial', 9), border: 1, align: DER, fmt: XL.mon }, extra || {});
    xSet(ws, 'A' + r, f.rubro, s({ align: IZQ }));
    xSet(ws, 'B' + r, f.proveedor || null, s({ align: IZQ }));
    xSet(ws, 'C' + r, null, s({ align: CEN, fmt: 'General' }));
    xSet(ws, 'D' + r, f.proyectado, s({ fill: AZ.verdeClaro }));
    xSet(ws, 'E' + r, f.cotizado, s());
    xSet(ws, 'F' + r, xF(`${xRef('FLUJO MENSUAL')}!E${11 + i}`, pag), s());
    xSet(ws, 'G' + r, xF(`D${r}-F${r}`, r2(f.proyectado - pag)), s());
    xSet(ws, 'H' + r, xF(`IFERROR(F${r}/D${r},0)`, f.proyectado ? pag / f.proyectado : 0), s({ align: CEN, fmt: '0%' }));
    ws.getRow(r).height = altoTexto(f.rubro, 26, 22, 12);
  });
  const T = r0 + Math.max(filas.length, 1), rf = T - 1;
  const sT = { font: font('Arial', 9, true, BLANCO), fill: AZ.titulo, border: 1, align: DER, fmt: XL.acctV };
  const tD = sum(filas, f => f.proyectado), tE = sum(filas, f => f.cotizado), tF = sum(Object.values(pagadoRubro), v => v);
  xSet(ws, 'A' + T, 'TOTALES', Object.assign({}, sT, { align: CEN }));
  ['B', 'C'].forEach(c => xSet(ws, c + T, null, sT));
  [['D', tD], ['E', tE], ['F', tF], ['G', r2(tD - tF)]].forEach(([c, v]) => xSet(ws, c + T, xF(`SUM(${c}${r0}:${c}${rf})`, v), sT));
  xSet(ws, 'H' + T, xF(`IFERROR(F${T}/D${T},0)`, tD ? tF / tD : 0), Object.assign({}, sT, { align: CEN, fmt: '0%' }));
  xSet(ws, 'A' + (T + 1), 'SIN IVA', { font: font('Arial', 9, true) });
  xSet(ws, 'D' + (T + 1), xF(`D${T}/1.16`, r2(tD / 1.16)), { font: font('Arial', 9), fmt: XL.acctV });
}

function hojaCxpProyecto(wb, P, nombreHoja) {
  const ws = wb.addWorksheet(nombreHoja, { views: [{ state: 'frozen', ySplit: 4, zoomScale: 80, showGridLines: false }] });
  xAnchos(ws, [18, 35.1, 17.7, 56.2, 17.1, 16.9, 16.6, 50]);
  ws.mergeCells('A1:H1');
  xSet(ws, 'A1', `CUENTAS POR PAGAR DEL PROYECTO ${P}`, { font: font('Arial', 14, true, BLANCO), fill: AZ.pagar, align: { vertical: 'middle', indent: 1 } });
  xSet(ws, 'A2', `Pagos registrados en la plataforma al ${fechaMx(hoy())}. TIPO = rubro del flujo.`, { font: font('Arial', 8, false, 'FF6A6688', { italic: true }) });
  ws.getRow(1).height = 25.5; ws.getRow(4).height = 31.5;
  ['TIPO', 'PROVEEDOR', 'FORMA DE PAGO', 'CONCEPTO', 'ORDEN DE COMPRA', 'MONTO PAGADO', 'FECHA DE PAGO', 'OBSERVACIONES'].forEach((h, i) =>
    xSet(ws, xCol(i + 1) + '4', h, { font: font('Arial', 9, true, BLANCO), fill: AZ.cxp, border: 1, align: CEN }));
  const apps = D().apps.filter(a => a.p.PROYECTO === P).sort((a, b) => (a.FECHA_PAGO || '').localeCompare(b.FECHA_PAGO || ''));
  apps.forEach((a, k) => {
    const r = 5 + k, p = a.p;
    const s = extra => Object.assign({ font: font('Arial', 9), border: 1, align: CEN }, extra || {});
    xSet(ws, 'A' + r, p.rubro || SIN_RUBRO, s());
    xSet(ws, 'B' + r, p.PROVEEDOR, s({ align: IZQ }));
    xSet(ws, 'C' + r, p.FORMA_PAGO, s());
    xSet(ws, 'D' + r, p.CONCEPTO, s({ align: IZQ }));
    xSet(ws, 'E' + r, p.OC || 'SIN OC', s());
    xSet(ws, 'F' + r, a.monto, s({ align: DER, fmt: '\\$#,##0.00' }));
    xSet(ws, 'G' + r, xDate(a.FECHA_PAGO), s({ fmt: XL.fecha }));
    xSet(ws, 'H' + r, [a.ID_PAGO ? 'Pago ' + a.ID_PAGO : '', 'Partida ' + p.PARTIDA, p.OBSERVACIONES].filter(Boolean).join(' · '), s({ align: IZQ }));
    ws.getRow(r).height = altoTexto(p.CONCEPTO, 60, 26.5, 12);
  });
  const t = 5 + apps.length;
  xSet(ws, 'E' + t, 'TOTAL', { font: font('Arial', 9, true), align: DER });
  xSet(ws, 'F' + t, xF(`SUM(F5:F${Math.max(t - 1, 5)})`, sum(apps, a => a.monto)), { font: font('Arial', 9, true), border: 1, fmt: '\\$#,##0.00' });
}

/* =====================================================================
 * 4. ESTADO DE CUENTA DEL PROVEEDOR  (diseño: pestañas "Proy. Drakkar" / "Pagos Drakkar")
 * ===================================================================== */
async function exportarProveedor() {
  const prov = state.provSel;
  if (!prov) throw new Error('Elige un proveedor.');
  const wb = await libroNuevo();
  const corto = prov.split(/\s+/).slice(0, 2).join(' ').slice(0, 20);
  const hProy = xHoja('Proy. ' + corto), hPagos = xHoja('Pagos ' + corto);
  const NAR = 'FFE26B0A';
  const head = (ws, heads, anchos) => {
    xAnchos(ws, anchos);
    heads.forEach((h, i) => xSet(ws, xCol(i + 1) + '1', h, { font: font('Calibri', 12, true, BLANCO), fill: NAR, border: 1, align: CEN }));
    ws.getRow(1).height = 45;
  };
  const sD = extra => Object.assign({ font: font('Calibri', 11), border: 1, align: CEN }, extra || {});
  const llave = o => o.OC || ('SIN OC · ' + (o.COTIZACION || o.ID));
  const ocs = D().ocs.filter(o => o.PROVEEDOR === prov);
  const apps = D().apps.filter(a => a.p.PROVEEDOR === prov).sort((a, b) => (a.FECHA_PAGO || '').localeCompare(b.FECHA_PAGO || ''));
  const llaveApp = a => (a.p.cot ? llave(a.p.cot) : (a.p.OC || 'SIN OC'));

  // Proy.
  const ws = wb.addWorksheet(hProy, { views: [{ state: 'frozen', ySplit: 1, zoomScale: 90 }] });
  head(ws, ['#', 'COTIZACION', 'CONCEPTO DEL PROYECTO', `OC RELACIONADA ${corto}`, 'TOTAL DEL PROYECTO', 'PAGADO', 'POR PAGAR', 'OBSERVACIONES'], [6, 28, 52, 24, 20, 20, 20, 44]);
  const filas = ocs.map((o, i) => ({ n: i + 1, cot: o.COTIZACION || 'FALTA # DE COTIZACIÓN', concepto: o.CONCEPTO, llave: llave(o), total: o.total, obs: [o.OBSERVACIONES, o.pars.length ? `${o.pars.length} parcialidades` : ''].filter(Boolean).join(' · ') }));
  uniq(apps.filter(a => !a.p.cot).map(llaveApp)).forEach(k => filas.push({ n: '', cot: '—', concepto: 'Pagos sin cotización registrada', llave: k, total: null, obs: 'Registra la cotización en la plataforma' }));
  const ref = xRef(hPagos);
  filas.forEach((f, i) => {
    const r = i + 2;
    const pag = sum(apps.filter(a => llaveApp(a) === f.llave), a => a.monto);
    xSet(ws, 'A' + r, f.n || null, sD());
    xSet(ws, 'B' + r, f.cot, sD({ font: font('Calibri', 11, false, f.cot.startsWith('FALTA') ? 'FFFF0000' : null) }));
    xSet(ws, 'C' + r, f.concepto || null, sD({ align: IZQ }));
    xSet(ws, 'D' + r, f.llave, sD({ font: font('Calibri', 11, false, f.llave.startsWith('SIN OC') ? 'FFFF0000' : null) }));
    xSet(ws, 'E' + r, f.total, sD({ fmt: XL.acct, align: DER }));
    xSet(ws, 'F' + r, xF(`SUMIF(${ref}!$C:$C,$D${r},${ref}!$E:$E)`, pag), sD({ fmt: XL.acct, align: DER }));
    xSet(ws, 'G' + r, xF(`E${r}-F${r}`, r2((f.total || 0) - pag)), sD({ fmt: XL.acct, align: DER }));
    xSet(ws, 'H' + r, f.obs || null, sD({ align: IZQ }));
    ws.getRow(r).height = altoTexto(f.concepto, 50, 22, 15);
  });
  const t = filas.length + 2;
  ['E', 'F', 'G'].forEach(c => xSet(ws, c + t, xF(`SUM(${c}2:${c}${t - 1})`, c === 'E' ? sum(filas, f => f.total || 0) : (c === 'F' ? sum(apps, a => a.monto) : r2(sum(filas, f => f.total || 0) - sum(apps, a => a.monto)))),
    { font: font('Calibri', 12, true), border: 1, fmt: XL.acct }));
  ws.autoFilter = `A1:H${Math.max(t - 1, 2)}`;

  // Pagos
  const wp = wb.addWorksheet(hPagos, { views: [{ state: 'frozen', ySplit: 1, zoomScale: 90 }] });
  head(wp, ['#', 'CONCEPTO DEL PROYECTO', 'ORDEN DE COMPRA RELACIONADA', 'FOLIO DE PAGO', 'TOTAL PAGADO', 'FECHA DE PAGO', 'OBSERVACIONES'], [8, 52, 26, 18, 20, 18, 50]);
  apps.forEach((a, i) => {
    const r = i + 2, p = a.p;
    xSet(wp, 'A' + r, num(p.PARTIDA), sD());
    xSet(wp, 'B' + r, p.CONCEPTO, sD({ align: IZQ }));
    xSet(wp, 'C' + r, llaveApp(a), sD());
    xSet(wp, 'D' + r, xLink(a.ID_PAGO || '—', a.COMPROBANTE_URL), sD({ font: font('Calibri', 11, false, 'FF0563C1', { underline: true }) }));
    xSet(wp, 'E' + r, a.monto, sD({ fmt: XL.acct, align: DER }));
    xSet(wp, 'F' + r, xDate(a.FECHA_PAGO), sD({ fmt: XL.fecha }));
    xSet(wp, 'G' + r, p.OBSERVACIONES || null, sD({ align: IZQ }));
    wp.getRow(r).height = altoTexto(p.CONCEPTO, 50, 22, 15);
  });
  const tp = apps.length + 3;
  xSet(wp, 'E' + tp, xF(`SUM(E2:E${tp - 1})`, sum(apps, a => a.monto)), { font: font('Calibri', 12, true), border: 1, fmt: XL.acct });
  wp.autoFilter = `A1:G${Math.max(apps.length + 1, 2)}`;

  // Parcialidades
  const wq = wb.addWorksheet('Parcialidades', { views: [{ state: 'frozen', ySplit: 1, zoomScale: 90 }] });
  head(wq, ['COTIZACIÓN', 'OC', 'N°', 'CONCEPTO', 'MONTO', 'FECHA TENTATIVA', 'ESTADO', 'PARTIDA'], [26, 18, 6, 36, 18, 16, 16, 10]);
  let r = 2;
  ocs.forEach(o => o.pars.forEach(x => {
    [o.COTIZACION || '—', o.OC || 'SIN OC', x.n, x.CONCEPTO, x.monto, xDate(x.FECHA_TENTATIVA), ESTADO_PAR[x.estado][0], x.partida ? num(x.partida.PARTIDA) : null].forEach((v, i) =>
      xSet(wq, xCol(i + 1) + r, v, sD({ align: i === 3 ? IZQ : CEN, fmt: i === 4 ? XL.acct : (i === 5 ? XL.fecha : undefined) })));
    r++;
  }));
  if (r === 2) xSet(wq, 'A2', 'Sin parcialidades registradas', { font: font('Calibri', 11, false, 'FF6A6688', { italic: true }) });

  await guardarLibro(wb, `ESTADO_CUENTA_${prov.replace(/[^A-Z0-9]+/g, '_')}_${fechaArchivo()}.xlsx`);
}
