/*************************************************
 * WEBAPP_API.gs — Backend para la App móvil (Costa Clean CRM)
 * - APIs para Dashboard, Listas, Detalles, Create/Update y Acciones
 * - Setup de hojas faltantes (GASTOS / CIERRES_TRIMESTRE)
 *************************************************/

function apiPing() {
  return { ok: true, ts: new Date().toISOString() };
}

function diagSheets_() {
  const ss = _ss_();
  const names = ss.getSheets().map(s => s.getName());
  const counts = {};
  names.forEach(n => {
    try {
      const sh = ss.getSheetByName(n);
      counts[n] = sh ? sh.getLastRow() : 0;
    } catch(e){
      counts[n] = -1;
    }
  });
  return {
    ok: true,
    ssId: (typeof SS_ID !== 'undefined' ? SS_ID : null),
    wants: { views: CC_VIEWS, sheets: CC_SHEETS },
    names,
    lastRow: counts
  };
}

/** ========= CONFIG DE HOJAS ========= **/
const CC_SHEETS = {
  CLIENTES: 'CLIENTES',
  LEADS: 'LEADS',
  FACTURA: 'FACTURA',              // tu hoja real de facturas
  LINEAS: 'LINEAS',                // líneas de factura
  PRESUPUESTOS: 'PRESUPUESTOS',    // proformas
  PRES_HIST: 'HISTORIAL_PRESUPUESTOS',
  PRES_LINEAS: 'PRES_LINEAS',      // líneas de proforma
  PRES_LINEAS_HIST: 'LINEAS_PRES_HIST',
  GASTOS: 'GASTOS',
  CIERRES: 'CIERRES_TRIMESTRE',
  CONFIG: 'CONFIG',
};

const CC_VIEWS = {
  CLIENTES: 'VW_CLIENTES',
  LEADS: 'VW_LEADS',
  PRESUPUESTOS: 'VW_PRESUPUESTOS',
  FACTURAS: 'VW_FACTURAS',
  GASTOS: 'VW_GASTOS'
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
function _ss_() {
  if (typeof SS_ID === 'undefined' || !SS_ID) {
    throw new Error('SS_ID no está definido. Revisa API.js (const SS_ID=...)');
  }
  try {
    return SpreadsheetApp.openById(SS_ID);
  } catch (err) {
    throw new Error('No pude abrir el Spreadsheet por SS_ID. ID=' + SS_ID + ' | ' + (err && err.message ? err.message : String(err)));
  }
}

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

function _getAllWithHeaders_(sheetName) {
  const sh = _sh_(sheetName);
  const values = sh.getDataRange().getValues();
  if (!values.length) return { headers: [], rows: [] };
  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null && c !== undefined))
    .map(r => _rowToObj_(headers, r));
  return { headers, rows };
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

  // FACTURAS: en tu sistema real suelen estar en HISTORIAL (y/o VW_FACTURAS).
  // Fallback robusto para no depender de ""FACTURA"".
  const shFact =
    _getSheetIfExists_('HISTORIAL') ||
    _getSheetIfExists_('VW_FACTURAS') ||
    _getSheetIfExists_('FACTURAS') ||
    _getSheetIfExists_('FACTURA');
  const facturas = shFact ? (_getAllWithHeadersFromSheet_(shFact).rows || []) : [];
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
  _ensureViews_();
  const map = _entityMap_();
  const cfg = map[entity];
  if (!cfg) throw new Error('Entidad no soportada: ' + entity);

  const viewName = cfg.view || cfg.sheet;
  const out = _listFromView_(viewName, params || {}, Number(params?.limit || 40));
  return (out === undefined || out === null) ? [] : out;
}

function apiGet(entity, id) {
  _ensureViews_();
  const map = _entityMap_();
  const cfg = map[entity];
  if (!cfg) throw new Error('Entidad no soportada: ' + entity);

  const viewName = cfg.view || cfg.sheet;
  const found = _findByIdInView_(viewName, cfg.idCol, id);
  if (!found) throw new Error('No encontrado: ' + entity + ' ' + id);
  return found.obj;
}

/** ========= API UI: PRESUPUESTOS ========= **/
function apiListPresupuestos(params){
  params = params || {};
  if (params.includeHistorial === undefined) params.includeHistorial = true;
  const ss = _ss_();
  const limit = Number(params?.limit || 100);

  try {
    _ensureViews_();
    const result = _listFromView_(CC_VIEWS.PRESUPUESTOS, params || {}, limit);
    const mapItem = (r) => ({
      id: r.Pres_ID || r.Presupuesto_ID || r.ID,
      cliente: r.Cliente || r.Cliente_ID || r.Lead_ID || '',
      clienteId: r.Cliente_ID || r.Lead_ID || '',
      estado: r.estado_normalizado || r.Estado || '',
      base: _firstNumber_(r.total_base, r.Base, r.Importe, r.Subtotal),
      total: _firstNumber_(r.total, r.Total, r.Importe_total, r.Total_con_IVA),
      fecha: r.Fecha || r.created_at || '',
      fechaRaw: r.Fecha || '',
      email: r.Email_cliente || r.Lead_Email || r.Email || '',
      nif: r.NIF || r.DNI || r.CIF || '',
      direccion: r.Direccion || r.Direccion_cliente || r.Direccion_servicio || r.Direccion_facturacion || '',
      cp: r.CP || r.Codigo_postal || r.Codigo_Postal || '',
      ciudad: r.Ciudad || r.Municipio || '',
      pdfUrl: r.pdf_url || r.PDF_link || r.Pdf_link || '',
      notas: r.Notas || r.Nota || r.Observaciones || '',
      sourceSheet: r.SourceSheet || ''
    });

    const out = _mapListResult_(result, mapItem);

    // Blindaje: si por alguna razón queda undefined, no devolvemos null al frontend
    if (out === undefined || out === null) {
      return { items: [], page: 1, pageSize: limit, total: 0 };
    }
    return out;

  } catch (err) {
    logEvent_(ss, 'WEBAPP', 'listPresupuestos', 'presupuestos', '', 'ERROR', err.message, { stack: err.stack });
    throw err;
  }
}
function _findByIdInView_(sheetName, idCol, id) {
  const data = _getViewData_(sheetName) || { headers: [], rows: [] };
  const needle = String(id || '').trim();
  if (!needle) return null;

  const getCI = (obj, key) => {
    if (!obj) return '';
    if (obj[key] !== undefined && obj[key] !== null) return obj[key];
    const k = String(key || '').toLowerCase();
    const foundKey = Object.keys(obj).find(h => String(h).toLowerCase() === k);
    if (foundKey && obj[foundKey] !== undefined && obj[foundKey] !== null) return obj[foundKey];
    return '';
  };

  const row = (data.rows || []).find(r => String(getCI(r, idCol) || '').trim() === needle);
  return row ? { headers: data.headers, obj: row } : null;
}

function _ensureViews_() {
  if (typeof ccEnsureViews_ === 'function') ccEnsureViews_(false);
}

function _getViewData_(viewName) {
  _ensureViews_();
  try {
    return _getAllWithHeaders_(viewName);
  } catch (e) {
    return { headers: [], rows: [] };
  }
}
function testListPresupuestos() {
  const ss = _ss_();
  try {
    const items = apiListPresupuestos({ includeHistorial: false });
    const sampleIds = items.slice(0, 5).map(p => p.id).filter(Boolean);
    logEvent_(ss, 'TEST', 'listPresupuestos', 'presupuestos', '', 'OK', `items=${items.length}`, { sampleIds });
    return { ok: true, total: items.length, sampleIds };
  } catch (err) {
    logEvent_(ss, 'TEST', 'listPresupuestos', 'presupuestos', '', 'ERROR', err.message || String(err), { stack: err.stack });
    throw err;
  }
}
function apiGetPresupuesto(id) {
  const ss = _ss_();
  const presId = String(id || '').trim();
  if (!presId) throw new Error('Pres_ID requerido');

  try {
    _ensureViews_();
    const found = _findByIdInView_(CC_VIEWS.PRESUPUESTOS, 'Pres_ID', presId);
    if (!found || !found.obj) throw new Error('No se encontro el presupuesto ' + presId);
    const row = found.obj;

    return {
      id: presId,
      raw: row,
      cliente: row.Cliente || row.Cliente_ID || row.Lead_ID || '',
      clienteId: row.Cliente_ID || row.Lead_ID || '',
      estado: row.estado_normalizado || row.Estado || '',
      base: _firstNumber_(row.total_base, row.Base),
      total: _firstNumber_(row.total, row.Total),
      fecha: row.Fecha || row.created_at || '',
      email: row.Email_cliente || row.Lead_Email || row.Email || '',
      nif: row.NIF || row.DNI || row.CIF || '',
      direccion: row.Direccion || row.Direccion_cliente || row.Direccion_servicio || row.Direccion_facturacion || '',
      cp: row.CP || row.Codigo_postal || row.Codigo_Postal || '',
      ciudad: row.Ciudad || row.Municipio || '',
      pdfUrl: row.pdf_url || row.PDF_link || row.Pdf_link || '',
      notas: row.Notas || row.Nota || row.Observaciones || '',
      lineas: _getPresupuestoLineas_(presId)
    };
  } catch (err) {
    logEvent_(ss, 'WEBAPP', 'getPresupuesto', 'presupuestos', presId, 'ERROR', err.message, { stack: err.stack });
    throw err;
  }
}
function apiGeneratePresupuestoPdf(presId) {
  const ss = _ss_();
  const id = String(presId || '').trim();
  if (!id) throw new Error('Pres_ID requerido');

  try {
    const res = generatePresupuestoPdfById(id);
    logEvent_(ss, 'WEBAPP', 'pdfPresupuesto', 'presupuestos', id, 'OK', '', { url: res });
    return { ok: true, pdfUrl: res };
  } catch (err) {
    logEvent_(ss, 'WEBAPP', 'pdfPresupuesto', 'presupuestos', id, 'ERROR', err.message, { stack: err.stack });
    throw err;
  }
}

function apiGenerateFacturaPdf(factId) {
  const ss = _ss_();
  const id = String(factId || '').trim();
  if (!id) throw new Error('Factura_ID requerido');

  try {
    const res = generateFacturaPdfById(id);
    logEvent_(ss, 'WEBAPP', 'pdfFactura', 'facturas', id, 'OK', '', { url: res });
    return { ok: true, pdfUrl: res };
  } catch (err) {
    logEvent_(ss, 'WEBAPP', 'pdfFactura', 'facturas', id, 'ERROR', err.message, { stack: err.stack });
    throw err;
  }
}

function apiCrearFacturaDesdePresupuesto(presId, options) {
  const ss = _ss_();
  const id = String(presId || '').trim();
  if (!id) throw new Error('Pres_ID requerido');

  try {
    const res = createFacturaDesdePresupuesto_(id, options || {});
    logEvent_(ss, 'WEBAPP', 'crearFacturaDesdePresupuesto', 'presupuestos', id, 'OK', '', res || null);
    return { ok: true, facturaId: res.facturaId || '', pdfUrl: res.pdfUrl || '', alreadyExisted: !!res.alreadyExisted };
  } catch (err) {
    logEvent_(ss, 'WEBAPP', 'crearFacturaDesdePresupuesto', 'presupuestos', id, 'ERROR', err.message, { stack: err.stack });
    throw err;
  }
}

function apiListClientes(params) {
  const ss = _ss_();
  try {
    _ensureViews_();
    const result = _listFromView_(CC_VIEWS.CLIENTES, params || {}, 200);
    const mapItem = (r) => ({
      id: r.Cliente_ID || r.ID,
      nombre: r.Nombre || r.Cliente || '',
      email: r.Email || r.Email_cliente || '',
      telefono: r.Telefono || r.Telefono || r.Phone || '',
      estado: r.estado_normalizado || r.Estado || ''
    });
    return _mapListResult_(result, mapItem);
  } catch (err) {
    logEvent_(ss, 'WEBAPP', 'listClientes', 'clientes', '', 'ERROR', err.message, { stack: err.stack });
    throw err;
  }
}

function apiListLeads(params) {
  const ss = _ss_();
  try {
    _ensureViews_();
    const result = _listFromView_(CC_VIEWS.LEADS, params || {}, 200);
    const mapItem = (r) => ({
      id: r.Lead_ID || r.ID,
      nombre: r.Nombre || r.Lead_Nombre || '',
      email: r.Email || r.Lead_Email || '',
      telefono: r.Telefono || r.Lead_Telefono || '',
      estado: r.estado_normalizado || r.Estado || ''
    });
    return _mapListResult_(result, mapItem);
  } catch (err) {
    logEvent_(ss, 'WEBAPP', 'listLeads', 'leads', '', 'ERROR', err.message, { stack: err.stack });
    throw err;
  }
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

function apiLeadMarcarGanado(leadId){
  const ss = _ss_();
  const id = String(leadId || '').trim();
  if (!id) throw new Error('Lead_ID requerido');

  // Encontrar lead en hoja LEADS por Lead_ID
  const found = _findById_(CC_SHEETS.LEADS, 'Lead_ID', id);
  if (!found) throw new Error('No encontrado: lead ' + id);

  const sh = _sh_(CC_SHEETS.LEADS);

  // Set Estado = Ganado (por header si existe; fallback: col V=22)
  const headers = found.headers || [];
  const estadoIx = headers.indexOf('Estado') + 1; // 1-based
  if (estadoIx > 0) {
    sh.getRange(found.rowNumber, estadoIx).setValue('Ganado');
  } else {
    sh.getRange(found.rowNumber, 22).setValue('Ganado'); // V
  }

  // Convertir a cliente usando tu lógica existente
  if (typeof convertirLeadEnCliente_ !== 'function') {
    throw new Error('No existe convertirLeadEnCliente_ (LEADS_A_CLIENTES.js)');
  }
  convertirLeadEnCliente_(ss, found.rowNumber);

  // Leer Cliente_ID (W=23) luego de convertir
  const clienteId = String(sh.getRange(found.rowNumber, 23).getValue() || '').trim();
  return { ok: true, leadId: id, clienteId };
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
  // FACTURAS: en tu sistema real suelen estar en HISTORIAL (y/o VW_FACTURAS).
  // Fallback robusto para no depender de ""FACTURA"".
  const shFact =
    _getSheetIfExists_('HISTORIAL') ||
    _getSheetIfExists_('VW_FACTURAS') ||
    _getSheetIfExists_('FACTURAS') ||
    _getSheetIfExists_('FACTURA');
  const facturas = shFact ? (_getAllWithHeadersFromSheet_(shFact).rows || []) : [];
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
  const ss = _ss_();
  const sh = ss.getSheetByName('HISTORIAL') || ss.getSheetByName('FACTURAS') || ss.getSheetByName('FACTURA');
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
  const url = generateFacturaPdfById(facturaId);
  return { ok: true, id: facturaId, pdfLink: url };
}

function _emailFactura_(facturaId) {
  // En PASO 2 conectamos con tu envío real (GmailApp / MailApp)
  return { ok: true, id: facturaId, sent: true };
}

function _convertProformaToFactura_(presId) {
  const res = createFacturaDesdePresupuesto_(presId);
  return { ok: true, presId, facturaId: res.facturaId || '', pdfUrl: res.pdfUrl || '', alreadyExisted: !!res.alreadyExisted };
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
    clientes: { sheet: CC_SHEETS.CLIENTES, view: CC_VIEWS.CLIENTES, idCol: 'Cliente_ID' },
    leads: { sheet: CC_SHEETS.LEADS, view: CC_VIEWS.LEADS, idCol: 'Lead_ID' },
    facturas: { sheet: CC_SHEETS.FACTURA, view: CC_VIEWS.FACTURAS, idCol: 'Factura_ID' },
    proformas: { sheet: CC_SHEETS.PRESUPUESTOS, view: CC_VIEWS.PRESUPUESTOS, idCol: 'Pres_ID' },
    gastos: { sheet: CC_SHEETS.GASTOS, view: CC_VIEWS.GASTOS, idCol: 'Gasto_ID' },
    cierres: { sheet: CC_SHEETS.CIERRES, idCol: 'Cierre_ID' }
  };
}

function _presupuestoSheet_() {
  const ss = _ss_();
  const hist = ss.getSheetByName(CC_SHEETS.PRES_HIST);
  const pres = ss.getSheetByName(CC_SHEETS.PRESUPUESTOS);

  if (pres) return pres;
  if (hist) return hist;

  return null;
}
function _findHeader_(headers, candidates) {
  const lower = headers.map(h => _normalizeKey_(h));
  for (let i = 0; i < candidates.length; i++) {
    const idx = lower.indexOf(_normalizeKey_(candidates[i]));
    if (idx !== -1) return headers[idx];
  }
  return '';
}
function _pickValue_(obj, keys) {
  for (let i = 0; i < keys.length; i++) {
    const v = obj[keys[i]];
    if (v !== undefined && v !== null && String(v).trim() !== '') return v;
  }
  return '';
}

function _safeNumber_(v) {
  if (v === undefined || v === null) return null;
  if (v instanceof Date) return null;

  let s = String(v).trim();
  if (!s) return null;

  // quitar € y espacios
  s = s.replace(/€/g, '').replace(/\s+/g, '');

  // EU: 1.234,56 -> 1234.56
  if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (/^\d+,\d+$/.test(s) && !s.includes('.')) {
    // 1234,56 -> 1234.56
    s = s.replace(',', '.');
  } else if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) {
    // US: 1,234.56 -> 1234.56
    s = s.replace(/,/g, '');
  }

  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

// Primer número válido (evita que un Date "truthy" gane por ||)
function _firstNumber_() {
  for (let i = 0; i < arguments.length; i++) {
    const n = _safeNumber_(arguments[i]);
    if (n !== null) return n;
  }
  return null;
}

function _formatDateIso_(v) {
  const d = _parseDate_(v);
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _getPresupuestoLineas_(presId) {
  const ss = _ss_();
  const sheetNames = [CC_SHEETS.PRES_LINEAS_HIST || 'LINEAS_PRES_HIST', CC_SHEETS.PRES_LINEAS];

  for (let i = 0; i < sheetNames.length; i++) {
    const name = sheetNames[i];
    if (!name) continue;
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const { headers, rows } = _getAllWithHeaders_(name);
    const idHeader = _findHeader_(headers, ['Pres_ID', 'Presupuesto_ID', 'ID']);
    if (!idHeader) continue;

    const filtered = rows.filter(r => String(r[idHeader]).trim() === presId);
    if (!filtered.length) continue;

    const lineHeader = _findHeader_(headers, ['Linea_n', 'Linea', 'Linea_num']);

    return filtered.map(r => ({
      linea: lineHeader ? r[lineHeader] : '',
      concepto: _pickValue_(r, ['Concepto', 'Descripcion', 'Descripción']),
      cantidad: _safeNumber_(_pickValue_(r, ['Cantidad', 'Cant'])),
      precio: _safeNumber_(_pickValue_(r, ['Precio', 'Precio_unitario'])),
      dto: _safeNumber_(_pickValue_(r, ['Dto_%', 'Dto'])),
      iva: _safeNumber_(_pickValue_(r, ['IVA_%', 'IVA'])),
      subtotal: _safeNumber_(_pickValue_(r, ['Subtotal', 'Importe', 'Base'])),
    })).sort((a, b) => (Number(a.linea) || 0) - (Number(b.linea) || 0));
  }

  return [];
}


function apiPresupuestosDebug() {
  const ss = _ss_();
  const pres = _getSheetIfExists_(CC_SHEETS.PRESUPUESTOS);
  const hist = _getSheetIfExists_(CC_SHEETS.PRES_HIST);

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    sheets: {
      PRESUPUESTOS: _sheetDebug_(pres),
      HISTORIAL_PRESUPUESTOS: _sheetDebug_(hist)
    }
  };
}
function _compareValues_(a, b) {
  const da = _parseDate_(a);
  const db = _parseDate_(b);
  if (da && db) return da.getTime() - db.getTime();
  const na = Number(a);
  const nb = Number(b);
  if (!isNaN(na) && !isNaN(nb)) return na - nb;
  return String(a || '').localeCompare(String(b || ''), 'es', { sensitivity: 'base' });
}

function _applyListFilters_(rows, params) {
  let out = rows || [];
  const q = (params?.q || '').toString().trim().toLowerCase();
  const estado = (params?.estado || params?.status || '').toString().trim().toUpperCase();
  const clienteId = (params?.clienteId || params?.cliente_id || params?.cliente || '').toString().trim().toLowerCase();
  const desde = params?.desde || params?.from || '';
  const hasta = params?.hasta || params?.to || '';

  if (q) {
    out = out.filter(r => {
      const search = String(r.search_text || '').toLowerCase();
      if (search) return search.includes(q);
      return JSON.stringify(r).toLowerCase().includes(q);
    });
  }

  if (estado) {
    out = out.filter(r => {
      const v = String(r.estado_normalizado || r.Estado || r.estado || '').toUpperCase();
      return v === estado;
    });
  }

  if (clienteId) {
    out = out.filter(r => {
      const v = String(r.Cliente_ID || r.clienteId || r.cliente_id || r.Lead_ID || '').toLowerCase();
      return v === clienteId;
    });
  }

  if (desde || hasta) {
    const from = _parseDate_(desde);
    const to = _parseDate_(hasta);
    out = out.filter(r => {
      const raw = r.updated_at || r.created_at || r.Fecha || r.Fecha_envio || r.Fecha_aceptacion || r.Fecha_pago;
      const d = _parseDate_(raw);
      if (!d) return false;
      if (from && d < from) return false;
      if (to && d > to) return false;
      return true;
    });
  }

  return out;
}

function _sortRows_(rows, params) {
  const sortBy = (params?.sortBy || params?.orderBy || '').toString().trim();
  const dir = (params?.sortDir || params?.order || 'desc').toString().toLowerCase();
  const key = sortBy || (rows[0] && (rows[0].updated_at ? 'updated_at' : (rows[0].Fecha ? 'Fecha' : (rows[0].created_at ? 'created_at' : ''))));
  if (!key) return rows;
  const factor = dir === 'asc' ? 1 : -1;
  return rows.slice().sort((a, b) => _compareValues_(a[key], b[key]) * factor);
}

function _paginateRows_(rows, params, defaultLimit) {
  const pageSize = Number(params?.pageSize || params?.limit || defaultLimit || 40);
  const page = Number(params?.page || 1);
  if (params && (params.page || params.pageSize)) {
    const start = (page - 1) * pageSize;
    return { items: rows.slice(start, start + pageSize), page, pageSize, total: rows.length };
  }
  return rows.slice(0, pageSize);
}

function _listFromView_(viewName, params, defaultLimit) {
  const data = _getViewData_(viewName);
  let rows = _applyListFilters_(data.rows || [], params);
  rows = _sortRows_(rows, params || {});
  return _paginateRows_(rows, params || {}, defaultLimit || 40);
}

function _mapListResult_(result, mapper) {
  if (result && Array.isArray(result.items)) {
    return {
      items: result.items.map(mapper).filter(r => r && r.id),
      page: result.page,
      pageSize: result.pageSize,
      total: result.total
    };
  }
  return (result || []).map(mapper).filter(r => r && r.id);
}

function _getSheetIfExists_(name) {
  return _ss_().getSheetByName(name) || null;
}

function _getAllWithHeadersFromSheet_(sh) {
  const values = sh.getDataRange().getValues();
  if (!values.length) return { headers: [], rows: [] };
  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null && c !== undefined))
    .map(r => _rowToObj_(headers, r));
  return { headers, rows };
}

function _normalizeKey_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '');
}

function _buildHeaderMap_(headers) {
  const map = {};
  headers.forEach(h => {
    const k = _normalizeKey_(h);
    if (!k) return;
    if (!map[k]) map[k] = h;
  });
  return map;
}

function _pickValueByMap_(row, headerMap, keys) {
  if (!row) return '';
  for (let i = 0; i < keys.length; i++) {
    const direct = row[keys[i]];
    if (direct !== undefined && direct !== null && String(direct).trim() !== '') return direct;

    const normalized = _normalizeKey_(keys[i]);
    const header = headerMap ? headerMap[normalized] : '';
    if (header && row[header] !== undefined && row[header] !== null && String(row[header]).trim() !== '') {
      return row[header];
    }
  }
  return '';
}

function _findPresupuestoRow_(rows, headerMap, presId) {
  if (!rows || !rows.length) return null;
  const idNeedle = String(presId || '').trim();
  if (!idNeedle) return null;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const idVal = _pickValueByMap_(row, headerMap, ['Pres_ID', 'Presupuesto_ID', 'ID']);
    if (String(idVal).trim() === idNeedle) return row;
  }
  return null;
}

function _mergePresupuestoItems_(presData, histData) {
  const presMap = _buildHeaderMap_(presData.headers || []);
  const histMap = _buildHeaderMap_(histData.headers || []);

  const presItems = (presData.rows || []).map(r => ({
    row: r,
    headerMap: presMap,
    sourceSheet: CC_SHEETS.PRESUPUESTOS
  }));
  const histItems = (histData.rows || []).map(r => ({
    row: r,
    headerMap: histMap,
    sourceSheet: CC_SHEETS.PRES_HIST
  }));

  if (!presItems.length) return histItems;

  const presIds = new Set();
  presItems.forEach(it => {
    const idVal = _pickValueByMap_(it.row, it.headerMap, ['Pres_ID', 'Presupuesto_ID', 'ID']);
    const id = String(idVal || '').trim();
    if (id) presIds.add(id);
  });

  const merged = presItems.slice();
  histItems.forEach(it => {
    const idVal = _pickValueByMap_(it.row, it.headerMap, ['Pres_ID', 'Presupuesto_ID', 'ID']);
    const id = String(idVal || '').trim();
    if (!id || !presIds.has(id)) merged.push(it);
  });

  return merged;
}

function _sheetDebug_(sh) {
  if (!sh) return { exists: false, name: '', lastRow: 0, lastColumn: 0, dataRows: 0 };
  const lastRow = sh.getLastRow();
  const lastColumn = sh.getLastColumn();
  let dataRows = 0;
  if (lastRow > 1 && lastColumn > 0) {
    const values = sh.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    dataRows = values.filter(r => r.some(c => c !== '' && c !== null && c !== undefined)).length;
  }
  return { exists: true, name: sh.getName(), lastRow, lastColumn, dataRows };
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






function testWebappListPresupuestosDirect(){
  const items = apiListPresupuestos({ includeHistorial: true, limit: 20 });
  console.log('[testWebappListPresupuestosDirect] items=', items && items.length);
  if (items && items.length) console.log('[sample]', JSON.stringify(items[0]));
  return items;
}



function testDiagPresFact(){
  const out = {};
  out.ss = (typeof apiPresupuestosDebug==='function') ? apiPresupuestosDebug() : null;
  out.pres = (typeof apiListPresupuestos==='function') ? apiListPresupuestos({includeHistorial:true,limit:300}) : null;
  out.fact = (typeof apiListFacturas==='function') ? apiListFacturas({limit:300}) : null;
  console.log(JSON.stringify({TEST:'testDiagPresFact', sample:{pres: out.pres && out.pres[0], fact: out.fact && out.fact[0]}, counts:{pres: out.pres && out.pres.length, fact: out.fact && out.fact.length}}, null, 2));
  return out;
}





function apiDbInfo(){
  const ss = _ss_();
  const names = ss.getSheets().map(s => s.getName());

  function info_(name){
    const sh = ss.getSheetByName(name);
    if (!sh) return { exists:false, lastRow:0, lastCol:0 };
    return { exists:true, lastRow:sh.getLastRow(), lastCol:sh.getLastColumn() };
  }

  const out = {
    ok: true,
    spreadsheetId: ss.getId(),
    sheets: names,
    targets: {
      CLIENTES: info_(CC_SHEETS.CLIENTES),
      LEADS: info_(CC_SHEETS.LEADS),
      FACTURA: info_(CC_SHEETS.FACTURA),
      LINEAS: info_(CC_SHEETS.LINEAS),
      PRESUPUESTOS: info_(CC_SHEETS.PRESUPUESTOS),
      PRES_HIST: info_(CC_SHEETS.PRES_HIST),
      PRES_LINEAS: info_(CC_SHEETS.PRES_LINEAS),
      GASTOS: info_(CC_SHEETS.GASTOS),
      CONFIG: info_(CC_SHEETS.CONFIG),
    }
  };

  return out;
}





