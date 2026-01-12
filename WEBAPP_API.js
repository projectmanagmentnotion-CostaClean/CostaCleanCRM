/*************************************************
 * WEBAPP_API.gs — Backend para la App móvil (Costa Clean CRM)
 * - APIs para Dashboard, Listas, Detalles, Create/Update y Acciones
 * - Setup de hojas faltantes (GASTOS / CIERRES_TRIMESTRE)
 *************************************************/

function apiPing() {
  return { ok: true, ts: new Date().toISOString() };
}

/** ========= CONFIG DE HOJAS ========= **/
const CC_SHEETS = {
  CLIENTES: 'CLIENTES',
  LEADS: 'LEADS',
  FACTURA: 'FACTURA',              // tu hoja real de facturas
  LINEAS: 'LINEAS',                // líneas de factura
  PRESUPUESTOS: 'PRESUPUESTOS',    // proformas
  PRES_LINEAS: 'PRES_LINEAS',      // líneas de proforma
  GASTOS: 'GASTOS',
  CIERRES: 'CIERRES_TRIMESTRE',
  CONFIG: 'CONFIG',
};

/** Headers recomendados (solo para crear hojas faltantes)
 *  Si ya tienes la hoja, NO se sobreescribe.
 */
const HEADERS_GASTOS = [
  'Gasto_ID','Fecha','Proveedor','Concepto','Categoria',
  'Base','IVA','Total','Metodo_pago','Referencia','Deducible','Notas','PDF_link'
];

const HEADERS_CIERRES = [
  'Cierre_ID','Periodo','Desde','Hasta','Fecha_cierre',
  'Ventas_base','IVA_repercutido','Gastos_base','IVA_soportado','IVA_neto',
  'IRPF_estimado','Facturas_emitidas','Facturas_pagadas','Facturas_pendientes',
  'Notas','Snapshot_json'
];

/** ========= HELPERS BASE ========= **/
function _ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function _sh_(name) {
  const sh = _ss_().getSheetByName(name);
  if (!sh) throw new Error('No existe la hoja: ' + name);
  return sh;
}

function _ensureSheet_(name, headers) {
  const ss = _ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  } else {
    // Si existe pero está vacía, ponemos headers recomendados
    if (headers && headers.length && sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function setupSheetsIfMissing_() {
  _ensureSheet_(CC_SHEETS.GASTOS, HEADERS_GASTOS);
  _ensureSheet_(CC_SHEETS.CIERRES, HEADERS_CIERRES);
  return true;
}

function _getHeaders_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
}

function _rowToObj_(headers, row) {
  const o = {};
  headers.forEach((h, i) => o[h] = row[i]);
  return o;
}

function _getAll_(sheetName) {
  const sh = _sh_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).map(r => _rowToObj_(headers, r));
}

function _findById_(sheetName, idCol, id) {
  const sh = _sh_(sheetName);
  const headers = _getHeaders_(sh);
  const idIndex = headers.indexOf(idCol);
  if (idIndex === -1) throw new Error(`No existe columna ${idCol} en ${sheetName}`);
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIndex]) === String(id)) {
      return { rowNumber: r + 1, headers, row: values[r], obj: _rowToObj_(headers, values[r]) };
    }
  }
  return null;
}

/** ========= DASHBOARD / KPIs ========= **/
function apiDashboard(period) {
  setupSheetsIfMissing_();

  // period: { year: 2025, month: 12 } o { year: 2025, quarter: 4 }
  const now = new Date();
  const year = period?.year || now.getFullYear();
  const month = period?.month || (now.getMonth() + 1);
  const quarter = period?.quarter || (Math.floor((month - 1) / 3) + 1);

  const rangeMonth = _monthRange_(year, month);
  const rangeQuarter = _quarterRange_(year, quarter);

  const facturas = _getAll_(CC_SHEETS.FACTURA);
  const gastos = _getAll_(CC_SHEETS.GASTOS);

  const ventasMes = _sumByDateRange_(facturas, 'Fecha', 'Total', rangeMonth.from, rangeMonth.to, (f) => true);
  const pendientes = facturas.filter(f => _isInRange_(f.Fecha, rangeMonth.from, rangeMonth.to) && _isPendingInvoice_(f)).length;

  const ivaRepTri = _sumByDateRange_(facturas, 'Fecha', 'IVA_total', rangeQuarter.from, rangeQuarter.to, (f) => true);
  const gastoDedTri = _sumByDateRange_(gastos, 'Fecha', 'Base', rangeQuarter.from, rangeQuarter.to, (g) => _toBool_(g.Deducible, true));
  const ivaSopTri = _sumByDateRange_(gastos, 'Fecha', 'IVA', rangeQuarter.from, rangeQuarter.to, (g) => _toBool_(g.Deducible, true));
  const ivaNeto = (Number(ivaRepTri) || 0) - (Number(ivaSopTri) || 0);

  return {
    periodo: {
      year, month, quarter,
      monthLabel: _monthLabel_(month),
      quarterLabel: `Q${quarter} ${year}`,
      fromQuarter: rangeQuarter.from,
      toQuarter: rangeQuarter.to
    },
    kpis: {
      ventasMes: Number(ventasMes) || 0,
      facturasPendientes: Number(pendientes) || 0,
      ivaRepercutidoTrimestre: Number(ivaRepTri) || 0,
      gastoDeducibleTrimestre: Number(gastoDedTri) || 0,
      ivaSoportadoTrimestre: Number(ivaSopTri) || 0,
      ivaNetoTrimestre: Number(ivaNeto) || 0
    }
  };
}

/** ========= LISTAS / DETALLES ========= **/
function apiList(entity, params) {
  setupSheetsIfMissing_();

  const q = (params?.q || '').toString().trim().toLowerCase();
  const limit = Number(params?.limit || 40);

  const map = _entityMap_();
  const cfg = map[entity];
  if (!cfg) throw new Error('Entidad no soportada: ' + entity);

  const rows = _getAll_(cfg.sheet);
  const filtered = q
    ? rows.filter(r => JSON.stringify(r).toLowerCase().includes(q))
    : rows;

  return filtered.slice(0, limit);
}

function apiGet(entity, id) {
  const map = _entityMap_();
  const cfg = map[entity];
  if (!cfg) throw new Error('Entidad no soportada: ' + entity);

  const found = _findById_(cfg.sheet, cfg.idCol, id);
  if (!found) throw new Error('No encontrado: ' + entity + ' ' + id);
  return found.obj;
}

/** ========= CREATE / UPDATE ========= **/
function apiCreate(entity, payload) {
  setupSheetsIfMissing_();
  const map = _entityMap_();
  const cfg = map[entity];
  if (!cfg) throw new Error('Entidad no soportada: ' + entity);

  if (entity === 'clientes') return _createCliente_(payload);
  if (entity === 'leads') return _createLead_(payload);
  if (entity === 'facturas') return _createFactura_(payload);
  if (entity === 'proformas') return _createProforma_(payload);
  if (entity === 'gastos') return _createGasto_(payload);

  throw new Error('Create no implementado para: ' + entity);
}

function apiCrearPresupuestoLead(leadId) {
  const id = String(leadId || '').trim();
  if (!id) throw new Error('Lead_ID requerido');
  if (typeof crearPresupuestoParaLead_ !== 'function') {
    throw new Error('No existe crearPresupuestoParaLead_');
  }
  const res = crearPresupuestoParaLead_(id);
  return { ok: true, leadId: id, presId: res && res.presId ? res.presId : '' };
}

function apiUpdate(entity, id, payload) {
  const map = _entityMap_();
  const cfg = map[entity];
  if (!cfg) throw new Error('Entidad no soportada: ' + entity);

  const found = _findById_(cfg.sheet, cfg.idCol, id);
  if (!found) throw new Error('No encontrado: ' + entity + ' ' + id);

  const sh = _sh_(cfg.sheet);
  const headers = found.headers;

  // Actualiza solo claves existentes en headers
  headers.forEach((h, i) => {
    if (payload && Object.prototype.hasOwnProperty.call(payload, h)) {
      sh.getRange(found.rowNumber, i + 1).setValue(payload[h]);
    }
  });

  return { ok: true, id };
}

/** ========= ACTIONS (PDF / EMAIL / PAGADA / CONVERTIR) ========= **/
function apiAction(entity, id, action, payload) {
  if (entity === 'facturas') {
    if (action === 'markPaid') return _markFacturaPagada_(id, payload);
    if (action === 'pdf') return _pdfFactura_(id);          // placeholder
    if (action === 'email') return _emailFactura_(id);      // placeholder
  }
  if (entity === 'proformas') {
    if (action === 'convertToFactura') return _convertProformaToFactura_(id);
    if (action === 'pdf') return _pdfProforma_(id);
    if (action === 'email') return _emailProforma_(id);
  }
  if (entity === 'cierres') {
    if (action === 'closeQuarter') return apiCloseQuarter(payload);
  }
  throw new Error(`Acción no soportada: ${entity}.${action}`);
}

/** ========= CIERRE TRIMESTRAL ========= **/
function apiCloseQuarter(payload) {
  setupSheetsIfMissing_();

  const now = new Date();
  const year = payload?.year || now.getFullYear();
  const quarter = payload?.quarter || (Math.floor((now.getMonth()) / 3) + 1);
  const range = _quarterRange_(year, quarter);

  const dash = apiDashboard({ year, quarter });
  const k = dash.kpis;

  // IRPF estimado (simple): si quieres lo refinamos luego según tu modelo fiscal real.
  // Por defecto: 0 (porque depende de tu situación, retenciones, etc.)
  const irpf = Number(payload?.irpf_estimado || 0);

  const cierreId = `CIERRE-${year}-Q${quarter}`;
  const sh = _sh_(CC_SHEETS.CIERRES);
  const headers = _getHeaders_(sh);

  const row = {};
  headers.forEach(h => row[h] = '');

  row['Cierre_ID'] = cierreId;
  row['Periodo'] = `Q${quarter} ${year}`;
  row['Desde'] = range.from;
  row['Hasta'] = range.to;
  row['Fecha_cierre'] = new Date();

  row['Ventas_base'] = 0; // si tienes Base en facturas lo calculamos luego
  row['IVA_repercutido'] = k.ivaRepercutidoTrimestre;
  row['Gastos_base'] = k.gastoDeducibleTrimestre;
  row['IVA_soportado'] = k.ivaSoportadoTrimestre;
  row['IVA_neto'] = k.ivaNetoTrimestre;

  row['IRPF_estimado'] = irpf;

  // Contadores útiles
  const facturas = _getAll_(CC_SHEETS.FACTURA);
  row['Facturas_emitidas'] = facturas.filter(f => _isInRange_(f.Fecha, range.from, range.to)).length;
  row['Facturas_pagadas'] = facturas.filter(f => _isInRange_(f.Fecha, range.from, range.to) && _isPaidInvoice_(f)).length;
  row['Facturas_pendientes'] = facturas.filter(f => _isInRange_(f.Fecha, range.from, range.to) && _isPendingInvoice_(f)).length;

  row['Notas'] = payload?.notas || '';
  row['Snapshot_json'] = JSON.stringify(dash);

  // Upsert por Cierre_ID
  const existing = _findById_(CC_SHEETS.CIERRES, 'Cierre_ID', cierreId);
  if (existing) {
    apiUpdate('cierres', cierreId, row);
  } else {
    const out = headers.map(h => row[h]);
    sh.appendRow(out);
  }

  return { ok: true, cierreId, dash };
}

/** ========= IMPLEMENTACIONES CREATE ========= **/
function _createCliente_(p) {
  const sh = _sh_(CC_SHEETS.CLIENTES);
  const headers = _getHeaders_(sh);

  // Tu generador real:
  const clienteId = generarSiguienteClienteId_();

  const row = {};
  headers.forEach(h => row[h] = '');

  // Ajusta a tus columnas reales típicas:
  row['Cliente_ID'] = clienteId;
  row['Nombre'] = p?.Nombre || p?.nombre || '';
  row['NIF'] = p?.NIF || p?.nif || '';
  row['Direccion'] = p?.Direccion || p?.direccion || '';
  row['CP'] = p?.CP || p?.cp || '';
  row['Ciudad'] = p?.Ciudad || p?.ciudad || '';
  row['Email'] = p?.Email || p?.email || '';
  row['Estado'] = p?.Estado || p?.estado || 'Activo';
  row['Fecha_alta'] = p?.Fecha_alta || new Date();

  sh.appendRow(headers.map(h => row[h]));
  return { ok: true, id: clienteId };
}

function _createLead_(p) {
  const sh = _sh_(CC_SHEETS.LEADS);
  const headers = _getHeaders_(sh);

  const leadId = generarLeadId_(sh); // tu función existente
  const row = {};
  headers.forEach(h => row[h] = '');

  row['Lead_ID'] = leadId;
  row['Fecha'] = p?.Fecha || new Date();
  row['Nombre'] = p?.Nombre || p?.nombre || '';
  row['Telefono'] = p?.Telefono || p?.telefono || '';
  row['Email'] = p?.Email || p?.email || '';
  row['Servicio'] = p?.Servicio || p?.servicio || '';
  row['Estado'] = p?.Estado || 'Nuevo';
  row['Origen'] = p?.Origen || 'App móvil';

  sh.appendRow(headers.map(h => row[h]));
  return { ok: true, id: leadId };
}

function _createFactura_(p) {
  const sh = _sh_(CC_SHEETS.FACTURA);
  const headers = _getHeaders_(sh);

  // Si ya tienes generador de Factura_ID en tu CRM, lo conectamos en el PASO 2.
  // Por ahora: generador simple basado en última fila.
  const facturaId = _nextIdFromSheet_(sh, 'Factura_ID', 'F');

  const row = {};
  headers.forEach(h => row[h] = '');

  row['Factura_ID'] = facturaId;
  row['Fecha'] = p?.Fecha || new Date();
  row['Estado'] = p?.Estado || 'Borrador';

  row['Cliente_ID'] = p?.Cliente_ID || '';
  row['Cliente'] = p?.Cliente || '';
  row['Email_cliente'] = p?.Email_cliente || '';
  row['NIF'] = p?.NIF || '';
  row['Direccion'] = p?.Direccion || '';
  row['CP'] = p?.CP || '';
  row['Ciudad'] = p?.Ciudad || '';

  row['Base'] = Number(p?.Base || 0);
  row['IVA_total'] = Number(p?.IVA_total || 0);
  row['Total'] = Number(p?.Total || 0);
  row['Notas'] = p?.Notas || '';

  sh.appendRow(headers.map(h => row[h]));
  return { ok: true, id: facturaId };
}

function _createProforma_(p) {
  // Reutiliza tu estructura pro de PRESUPUESTOS (headers)
  const sh = _sh_(CC_SHEETS.PRESUPUESTOS);
  const headers = _getHeaders_(sh);

  const presId = _nextIdFromSheet_(sh, 'Pres_ID', 'PRO');
  const row = {};
  headers.forEach(h => row[h] = '');

  row['Pres_ID'] = presId;
  row['Fecha'] = p?.Fecha || new Date();
  row['Validez_dias'] = p?.Validez_dias || 7;
  row['Estado'] = p?.Estado || 'Borrador';

  row['Cliente_ID'] = p?.Cliente_ID || '';
  row['Cliente'] = p?.Cliente || '';
  row['Email_cliente'] = p?.Email_cliente || '';
  row['NIF'] = p?.NIF || '';
  row['Direccion'] = p?.Direccion || '';
  row['CP'] = p?.CP || '';
  row['Ciudad'] = p?.Ciudad || '';

  row['Base'] = Number(p?.Base || 0);
  row['IVA_total'] = Number(p?.IVA_total || 0);
  row['Total'] = Number(p?.Total || 0);
  row['Notas'] = p?.Notas || '';

  sh.appendRow(headers.map(h => row[h]));
  return { ok: true, id: presId };
}

function _createGasto_(p) {
  const sh = _sh_(CC_SHEETS.GASTOS);
  const headers = _getHeaders_(sh);

  const gastoId = _nextIdFromSheet_(sh, 'Gasto_ID', 'G');
  const row = {};
  headers.forEach(h => row[h] = '');

  row['Gasto_ID'] = gastoId;
  row['Fecha'] = p?.Fecha || new Date();
  row['Proveedor'] = p?.Proveedor || '';
  row['Concepto'] = p?.Concepto || '';
  row['Categoria'] = p?.Categoria || 'General';

  row['Base'] = Number(p?.Base || 0);
  row['IVA'] = Number(p?.IVA || 0);
  row['Total'] = Number(p?.Total || (Number(p?.Base||0) + Number(p?.IVA||0)));

  row['Metodo_pago'] = p?.Metodo_pago || '';
  row['Referencia'] = p?.Referencia || '';
  row['Deducible'] = (p?.Deducible === undefined) ? true : p?.Deducible;
  row['Notas'] = p?.Notas || '';
  row['PDF_link'] = p?.PDF_link || '';

  sh.appendRow(headers.map(h => row[h]));
  return { ok: true, id: gastoId };
}

/** ========= ACTIONS (placeholders + hooks) ========= **/
function _markFacturaPagada_(facturaId, payload) {
  // Intenta marcar por columna Estado o Pagada/Fecha_pago si existe
  const found = _findById_(CC_SHEETS.FACTURA, 'Factura_ID', facturaId);
  if (!found) throw new Error('Factura no encontrada: ' + facturaId);

  const sh = _sh_(CC_SHEETS.FACTURA);
  const headers = found.headers;

  const idxEstado = headers.indexOf('Estado');
  if (idxEstado !== -1) sh.getRange(found.rowNumber, idxEstado + 1).setValue('Pagada');

  const idxFechaPago = headers.indexOf('Fecha_pago');
  if (idxFechaPago !== -1) sh.getRange(found.rowNumber, idxFechaPago + 1).setValue(new Date());

  return { ok: true, id: facturaId };
}

function _pdfFactura_(facturaId) {
  // En PASO 2 conectamos con tu generador real de PDF de facturas (si ya lo tienes)
  return { ok: true, id: facturaId, pdfLink: '' };
}

function _emailFactura_(facturaId) {
  // En PASO 2 conectamos con tu envío real (GmailApp / MailApp)
  return { ok: true, id: facturaId, sent: true };
}

function _convertProformaToFactura_(presId) {
  // Aquí conectaremos con tu función real en PRESUPUESTOS.gs (tú ya lo tienes pro)
  // Por ahora: devolvemos ok para que la UI tenga endpoint.
  return { ok: true, presId, facturaId: '' };
}

function _pdfProforma_(presId) {
  return { ok: true, presId, pdfLink: '' };
}

function _emailProforma_(presId) {
  return { ok: true, presId, sent: true };
}

/** ========= UTILIDADES ========= **/
function _entityMap_() {
  return {
    clientes: { sheet: CC_SHEETS.CLIENTES, idCol: 'Cliente_ID' },
    leads: { sheet: CC_SHEETS.LEADS, idCol: 'Lead_ID' },
    facturas: { sheet: CC_SHEETS.FACTURA, idCol: 'Factura_ID' },
    proformas: { sheet: CC_SHEETS.PRESUPUESTOS, idCol: 'Pres_ID' },
    gastos: { sheet: CC_SHEETS.GASTOS, idCol: 'Gasto_ID' },
    cierres: { sheet: CC_SHEETS.CIERRES, idCol: 'Cierre_ID' },
  };
}

function _toBool_(v, defaultValue) {
  if (v === true || v === false) return v;
  if (v === '' || v === null || v === undefined) return defaultValue;
  const s = String(v).toLowerCase().trim();
  if (['true','sí','si','1','yes'].includes(s)) return true;
  if (['false','no','0'].includes(s)) return false;
  return defaultValue;
}

function _parseDate_(v) {
  if (v instanceof Date) return v;
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function _isInRange_(vDate, from, to) {
  const d = _parseDate_(vDate);
  if (!d) return false;
  return d >= from && d <= to;
}

function _sumByDateRange_(rows, dateKey, valueKey, from, to, predicateFn) {
  let sum = 0;
  rows.forEach(r => {
    if (!_isInRange_(r[dateKey], from, to)) return;
    if (predicateFn && !predicateFn(r)) return;
    sum += Number(r[valueKey] || 0);
  });
  return sum;
}

function _isPaidInvoice_(f) {
  const s = String(f.Estado || '').toLowerCase();
  return s.includes('pag') || s === 'cobrada';
}

function _isPendingInvoice_(f) {
  const s = String(f.Estado || '').toLowerCase();
  return ['enviada','pendiente','vencida','impagada'].some(x => s.includes(x));
}

function _monthRange_(year, month) {
  const from = new Date(year, month - 1, 1);
  const to = new Date(year, month, 0, 23, 59, 59, 999);
  return { from, to };
}

function _quarterRange_(year, quarter) {
  const m1 = (quarter - 1) * 3;
  const from = new Date(year, m1, 1);
  const to = new Date(year, m1 + 3, 0, 23, 59, 59, 999);
  return { from, to };
}

function _monthLabel_(m) {
  return ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'][m - 1] || '';
}

function _nextIdFromSheet_(sh, idCol, prefix) {
  const headers = _getHeaders_(sh);
  const idx = headers.indexOf(idCol);
  if (idx === -1) {
    // Si no hay columna, creamos fallback en A
    return `${prefix}-${new Date().getFullYear()}-0001`;
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return `${prefix}-${new Date().getFullYear()}-0001`;

  const lastVal = String(sh.getRange(lastRow, idx + 1).getValue() || '').trim();
  const n = Number(lastVal.replace(/[^\d]/g, ''));
  const next = isNaN(n) ? 1 : (n + 1);
  const y = new Date().getFullYear();
  return `${prefix}-${y}-${String(next).padStart(4, '0')}`;
}
