const CC_VIEW_NAMES = {
  CLIENTES: 'VW_CLIENTES',
  LEADS: 'VW_LEADS',
  PRESUPUESTOS: 'VW_PRESUPUESTOS',
  FACTURAS: 'VW_FACTURAS',
  GASTOS: 'VW_GASTOS'
};

const CC_AUDIT_KEYS = ['Pres_ID', 'Cliente_ID', 'Lead_ID', 'Factura_ID', 'Gasto_ID'];
const CC_PRES_ESTADOS = ['BORRADOR', 'ENVIADO', 'ACEPTADO', 'PERDIDO'];
const CC_FACT_ESTADOS = ['BORRADOR', 'EMITIDA', 'ENVIADA', 'PENDIENTE', 'VENCIDA', 'IMPAGADA', 'PAGADA', 'ANULADA'];
const CC_LEAD_ESTADOS = ['NUEVO', 'GANADO', 'PERDIDO'];

const CC_INDEX_KEY = 'cc_index_v1';
const CC_VIEWS_DIRTY_KEY = 'cc_views_dirty';
const CC_VIEWS_LAST_BUILD_KEY = 'cc_views_last_build';

function ccSetupWebAppLayer_() {
  ccEnsureViews_(true);
  ccBuildIndex_();
  if (typeof setupValidationsPresupuestos === 'function') setupValidationsPresupuestos();
  if (typeof factApplyValidations_ === 'function') factApplyValidations_();
  return true;
}

function ccSetupAndAudit() {
  if (typeof setupAll === 'function') setupAll();
  return runSpreadsheetAudit();
}

function ccEnsureViews_(force) {
  const props = PropertiesService.getScriptProperties();
  const dirty = props.getProperty(CC_VIEWS_DIRTY_KEY) === '1';
  const lastBuild = props.getProperty(CC_VIEWS_LAST_BUILD_KEY);

// AUTO_REBUILD_IF_EMPTY_VIEWS
// Si las vistas existen pero están vacías (solo headers), forzamos rebuild.
// Esto evita depender de clasp run / permisos y hace el sistema auto-reparable.
try {
  const ss = _ss_ ? _ss_() : SpreadsheetApp.getActiveSpreadsheet();
  const mustHave = [CC_VIEW_NAMES.CLIENTES, CC_VIEW_NAMES.LEADS, CC_VIEW_NAMES.PRESUPUESTOS, CC_VIEW_NAMES.FACTURAS, CC_VIEW_NAMES.GASTOS];
  let anyEmpty = false;
  mustHave.forEach((name) => {
    const sh = ccGetSheet_(name, false);
    if (sh && sh.getLastRow() <= 1) anyEmpty = true;
  });
  if (anyEmpty) {
    force = true;
    try { props.setProperty(CC_VIEWS_DIRTY_KEY, '1'); } catch(e){}
  }
} catch(e) {}

  if (!force && !dirty && lastBuild) {
    const anyMissing = !ccGetSheet_(CC_VIEW_NAMES.CLIENTES, false)
      || !ccGetSheet_(CC_VIEW_NAMES.LEADS, false)
      || !ccGetSheet_(CC_VIEW_NAMES.PRESUPUESTOS, false)
      || !ccGetSheet_(CC_VIEW_NAMES.FACTURAS, false);
    if (!anyMissing) return false;
  }

  ccBuildViews_();
  props.setProperty(CC_VIEWS_DIRTY_KEY, '0');
  props.setProperty(CC_VIEWS_LAST_BUILD_KEY, new Date().toISOString());
  return true;
}

function ccMarkViewsDirty_() {
  PropertiesService.getScriptProperties().setProperty(CC_VIEWS_DIRTY_KEY, '1');
}

function ccInvalidateIndex_() {
  CacheService.getScriptCache().remove(CC_INDEX_KEY);
}

function ccBuildIndex_() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return null;
  try {
    ccEnsureViews_(false);

    const index = {
      ts: new Date().toISOString(),
      views: {}
    };

    Object.keys(CC_VIEW_NAMES).forEach((key) => {
      const name = CC_VIEW_NAMES[key];
      const sh = ccGetSheet_(name, false);
      if (!sh) return;

      const data = ccGetSheetData_(sh);
      const idHeader = ccFindHeader_(data.headers, ['Cliente_ID', 'Lead_ID', 'Pres_ID', 'Factura_ID', 'Gasto_ID', 'ID']);
      const estadoHeader = ccFindHeader_(data.headers, ['estado_normalizado', 'Estado']);

      const byId = {};
      const byEstado = {};

      data.rows.forEach((row, i) => {
        const id = idHeader ? String(row[idHeader] || '').trim() : '';
        if (id) byId[id] = i + 2;
        if (estadoHeader) {
          const estado = String(row[estadoHeader] || '').trim().toUpperCase();
          if (estado) {
            if (!byEstado[estado]) byEstado[estado] = [];
            byEstado[estado].push(id || String(i + 2));
          }
        }
      });

      index.views[name] = { byId, byEstado };
    });

    CacheService.getScriptCache().put(CC_INDEX_KEY, JSON.stringify(index), 21600);
    return index;
  } finally {
    lock.releaseLock();
  }
}

function ccGetIndex_() {
  const cache = CacheService.getScriptCache();
  const raw = cache.get(CC_INDEX_KEY);
  if (raw) {
    try { return JSON.parse(raw); } catch (_) {}
  }
  return ccBuildIndex_();
}

function ccNormalizeEstadoOnEdit_(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

  const headers = ccGetHeaders_(sh);
  const idx = headers.indexOf('Estado');
  if (idx === -1 || col !== idx + 1) return;

  const cache = CacheService.getScriptCache();
  const guardKey = 'cc_state_norm_guard';
  if (cache && cache.get(guardKey)) return;
  if (cache) cache.put(guardKey, '1', 5);

  try {
    const value = String(e.range.getDisplayValue() || '').trim();
    const upper = value.toUpperCase();
    if (upper && upper !== value) {
      sh.getRange(row, idx + 1).setValue(upper);
    }
  } finally {
    if (cache) cache.remove(guardKey);
  }
}

function ccBuildViews_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ccBuildClientesView_(ss);
  ccBuildLeadsView_(ss);
  ccBuildPresupuestosView_(ss);
  ccBuildFacturasView_(ss);
  ccBuildGastosView_(ss);

  return true;
}

function ccBuildClientesView_(ss) {
  const source = ss.getSheetByName('CLIENTES');
  const view = ccGetSheet_(CC_VIEW_NAMES.CLIENTES, true);
  const headers = [
    'Cliente_ID','Nombre','Email','Telefono','NIF','Direccion','CP','Ciudad','Tipo_cliente','Origen','Estado',
    'Fecha_alta','created_at','updated_at','estado_normalizado','search_text'
  ];
  if (!source) return ccWriteView_(view, headers, []);

  const data = ccGetSheetData_(source);
  const rows = data.rows.map((row) => {
    const id = ccPick_(row, data.headers, ['Cliente_ID','ID']);
    const nombre = ccPick_(row, data.headers, ['Nombre','Cliente']);
    const email = ccPick_(row, data.headers, ['Email','Email_cliente']);
    const telefono = ccPick_(row, data.headers, ['Telefono','TelÃ©fono','Phone']);
    const nif = ccPick_(row, data.headers, ['NIF','DNI','CIF']);
    const direccion = ccPick_(row, data.headers, ['Direccion','DirecciÃ³n']);
    const cp = ccPick_(row, data.headers, ['CP','Codigo_postal','Codigo_Postal']);
    const ciudad = ccPick_(row, data.headers, ['Ciudad','Municipio','Poblacion']);
    const tipo = ccPick_(row, data.headers, ['Tipo_cliente','Tipo']);
    const origen = ccPick_(row, data.headers, ['Origen']);
    const estado = ccPick_(row, data.headers, ['Estado']);
    const fechaAlta = ccPick_(row, data.headers, ['Fecha_alta','Fecha']);
    const createdAt = ccFormatDateIso_(fechaAlta);
    const updatedAt = createdAt;
    const estadoNorm = ccNormalizeEstado_(estado);
    const search = ccBuildSearchText_([id, nombre, email, telefono, nif, ciudad, direccion]);

    return [
      id, nombre, email, telefono, nif, direccion, cp, ciudad, tipo, origen, estado,
      ccFormatDateIso_(fechaAlta), createdAt, updatedAt, estadoNorm, search
    ];
  });

  ccWriteView_(view, headers, rows);
}

function ccBuildLeadsView_(ss) {
  const source = ss.getSheetByName('LEADS');
  const view = ccGetSheet_(CC_VIEW_NAMES.LEADS, true);
  const headers = [
    'Lead_ID','Nombre','Email','Telefono','NIF','Direccion','CP','Ciudad','Estado','Cliente_ID','Origen',
    'Fecha_entrada','Ultimo_contacto','created_at','updated_at','estado_normalizado','search_text'
  ];
  if (!source) return ccWriteView_(view, headers, []);

  const data = ccGetSheetData_(source);
  const rows = data.rows.map((row) => {
    const id = ccPick_(row, data.headers, ['Lead_ID','ID']);
    const nombre = ccPick_(row, data.headers, ['Nombre','Lead_Nombre']);
    const email = ccPick_(row, data.headers, ['Email','Lead_Email']);
    const telefono = ccPick_(row, data.headers, ['Telefono','Lead_Telefono']);
    const nif = ccPick_(row, data.headers, ['NIF/CIF','NIF','CIF']);
    const direccion = ccPick_(row, data.headers, ['Direccion','DirecciÃ³n']);
    const cp = ccPick_(row, data.headers, ['CP','Codigo_postal','Codigo_Postal']);
    const ciudad = ccPick_(row, data.headers, ['Poblacion','Ciudad','Municipio']);
    const estado = ccPick_(row, data.headers, ['Estado']);
    const clienteId = ccPick_(row, data.headers, ['Cliente_ID']);
    const origen = ccPick_(row, data.headers, ['Origen']);
    const fechaEntrada = ccPick_(row, data.headers, ['Fecha_entrada','Fecha']);
    const ultimoContacto = ccPick_(row, data.headers, ['Ultimo_contacto']);
    const createdAt = ccFormatDateIso_(fechaEntrada);
    const updatedAt = ccFormatDateIso_(ultimoContacto || fechaEntrada);
    const estadoNorm = ccNormalizeEstado_(estado);
    const search = ccBuildSearchText_([id, nombre, email, telefono, nif, ciudad, direccion, clienteId]);

    return [
      id, nombre, email, telefono, nif, direccion, cp, ciudad, estado, clienteId, origen,
      ccFormatDateIso_(fechaEntrada), ccFormatDateIso_(ultimoContacto),
      createdAt, updatedAt, estadoNorm, search
    ];
  });

  ccWriteView_(view, headers, rows);
}

function ccBuildPresupuestosView_(ss) {
  const shPres = ss.getSheetByName('PRESUPUESTOS');
  const shHist = ss.getSheetByName('HISTORIAL_PRESUPUESTOS');
  const view = ccGetSheet_(CC_VIEW_NAMES.PRESUPUESTOS, true);

  // Headers normalizados para UI/API (VW_PRESUPUESTOS)
  const headers = [
    'Pres_ID','Cliente_ID','Cliente','Email_cliente','NIF','Direccion','CP','Ciudad','Estado',
    'Base','IVA_total','Total','Fecha','Fecha_envio','Fecha_aceptacion','PDF_link','Notas',
    'Tipo','IsDemo','SourceSheet','Updated_at'
  ];

  view.clearContents();
  view.getRange(1,1,1,headers.length).setValues([headers]);

  // Merge seguro: PRESUPUESTOS + HISTORIAL_PRESUPUESTOS (si existe)
  const merged = ccMergePresupuestoSources_(shPres, shHist);
  if (!merged || !merged.rows || !merged.rows.length) return;

  const rows = merged.rows.map((row) => {
    const id = ccPick_(row, merged.headers, ['Pres_ID','Presupuesto_ID','ID']);
    const cliId = ccPick_(row, merged.headers, ['Cliente_ID','CLI_ID','Lead_ID']);
    const cli = ccPick_(row, merged.headers, ['Cliente','Nombre','Razon_social']);
    const email = ccPick_(row, merged.headers, ['Email_cliente','Email','Lead_Email']);
    const nif = ccPick_(row, merged.headers, ['NIF','DNI','CIF']);
    const dir = ccPick_(row, merged.headers, ['Direccion','Direccion_cliente','Direccion_servicio','Direccion_facturacion']);
    const cp = ccPick_(row, merged.headers, ['CP','Codigo_postal','Codigo_Postal']);
    const ciudad = ccPick_(row, merged.headers, ['Ciudad','Municipio']);
    const estado = ccPick_(row, merged.headers, ['Estado','estado_normalizado','Status']);
    const base = ccPick_(row, merged.headers, ['Base','Subtotal','total_base']);
    const iva = ccPick_(row, merged.headers, ['IVA_total','IVA','iva_total']);
    const total = ccPick_(row, merged.headers, ['Total','Importe_total','total','Total_con_IVA']);
    const fecha = ccPick_(row, merged.headers, ['Fecha','created_at','Fecha_presupuesto']);
    const fEnv = ccPick_(row, merged.headers, ['Fecha_envio']);
    const fAcep = ccPick_(row, merged.headers, ['Fecha_aceptacion']);
    const pdf = ccPick_(row, merged.headers, ['PDF_link','Pdf_link','pdf_url','pdfUrl']);
    const notas = ccPick_(row, merged.headers, ['Notas','Nota','Observaciones']);

    // Tipo/IsDemo (sin tocar datos source)
    const idStr = String(id || '');
    const tipo = (/^(PRO|PF|PROFORMA)[-_]/i.test(idStr)) ? 'PROFORMA' : 'PRESUPUESTO';
    const isDemo = (/demo|mock|fictic/i.test(String(notas||''))) ? 'TRUE' : 'FALSE';

    const updated = new Date().toISOString();
    const source = ccPick_(row, merged.headers, ['SourceSheet']) || '';

    return [
      id, cliId, cli, email, nif, dir, cp, ciudad, estado,
      base, iva, total, fecha, fEnv, fAcep, pdf, notas,
      tipo, isDemo, source, updated
    ];
  });

  if (rows.length){
    view.getRange(2,1,rows.length,headers.length).setValues(rows);
  }
}

function ccBuildFacturasView_(ss) {
  // SOURCE OF TRUTH (reales): HISTORIAL (segun tu captura)
  // Fallbacks: FACTURAS / FACTURA (por compatibilidad)
  const shFact =
    ss.getSheetByName('HISTORIAL') ||
    ss.getSheetByName('FACTURAS') ||
    ss.getSheetByName('FACTURA');

  const view = ccGetSheet_(CC_VIEW_NAMES.FACTURAS, true);

  // Headers normalizados para UI/API (VW_FACTURAS)
  const headers = [
    'Factura_ID','Pres_ID','Cliente_ID','Cliente','Email','NIF','Direccion','CP','Ciudad','Estado',
    'Fecha','Fecha_envio','Fecha_pago','Base','IVA_total','Total','PDF_link',
    'created_at','updated_at','estado_normalizado','total_base','total_iva','total','pdf_url','search_text',
    'SourceSheet'
  ];

  view.clearContents();
  view.getRange(1,1,1,headers.length).setValues([headers]);

  if (!shFact) return;

  const data = ccGetSheetData_(shFact);
  if (!data || !data.rows || !data.rows.length) return;

  const rows = data.rows.map((row) => {
    // ID: en tu hoja real parece ser Numero_factura
    const id = ccPick_(row, data.headers, ['Factura_ID','Numero_factura','ID']);
    const presId = ccPick_(row, data.headers, ['Pres_ID']);

    const clienteId = ccPick_(row, data.headers, ['Cliente_ID','Lead_ID']);
    const cliente   = ccPick_(row, data.headers, ['Cliente','Nombre']);
    const email     = ccPick_(row, data.headers, ['Email','Email_cliente']);
    const nif       = ccPick_(row, data.headers, ['NIF','DNI','CIF']);
    const dir       = ccPick_(row, data.headers, ['Direccion','Direccion_cliente','Direccion_facturacion']);
    const cp        = ccPick_(row, data.headers, ['CP','Codigo_postal','Codigo_Postal']);
    const ciudad    = ccPick_(row, data.headers, ['Ciudad','Municipio']);
    const estado    = ccPick_(row, data.headers, ['Estado','estado_normalizado','Status']);

    const fecha     = ccPick_(row, data.headers, ['Fecha','created_at']);
    const fEnv      = ccPick_(row, data.headers, ['Fecha_envio']);
    const fPago     = ccPick_(row, data.headers, ['Fecha_pago']);

    // Totales: tu hoja real usa Base / IVA / Total
    const base      = ccPick_(row, data.headers, ['Base','Subtotal','total_base']);
    const iva       = ccPick_(row, data.headers, ['IVA_total','IVA','total_iva']);
    const total     = ccPick_(row, data.headers, ['Total','total','Importe_total']);

    const pdf       = ccPick_(row, data.headers, ['PDF_link','PDF (link)','pdf_url','Pdf_link','pdfUrl']);

    // Campos extra para UI/busqueda (no rompen si vacios)
    const createdAt = ccPick_(row, data.headers, ['created_at']) || '';
    const updatedAt = ccPick_(row, data.headers, ['updated_at']) || '';
    const estadoNorm= ccPick_(row, data.headers, ['estado_normalizado']) || estado || '';
    const totalBase = ccPick_(row, data.headers, ['total_base']) || base || '';
    const totalIva  = ccPick_(row, data.headers, ['total_iva']) || iva || '';
    const totalAny  = ccPick_(row, data.headers, ['total']) || total || '';
    const pdfUrl    = ccPick_(row, data.headers, ['pdf_url']) || pdf || '';

    const st = [
      id, presId, clienteId, cliente, email, nif, dir, cp, ciudad, estado,
      fecha, fEnv, fPago, base, iva, total, pdf,
      createdAt, updatedAt, estadoNorm, totalBase, totalIva, totalAny, pdfUrl,
      '', // search_text (lo puedes mejorar luego)
      (shFact.getName() || '')
    ];

    return st;
  });

  if (rows.length){
    view.getRange(2,1,rows.length,headers.length).setValues(rows);
  }
}

function ccBuildGastosView_(ss) {
  const sh = ss.getSheetByName('GASTOS');
  const view = ccGetSheet_(CC_VIEW_NAMES.GASTOS, true);

  const headers = [
    'Gasto_ID','Cliente_ID','Cliente','Categoria','Concepto','Importe','Fecha','Estado','Notas',
    'SourceSheet','Updated_at'
  ];

  view.clearContents();
  view.getRange(1,1,1,headers.length).setValues([headers]);

  if (!sh) return;
  const data = ccGetSheetData_(sh);
  if (!data || !data.rows || !data.rows.length) return;

  const rows = data.rows.map((row) => {
    const id = ccPick_(row, data.headers, ['Gasto_ID','ID']);
    const cliId = ccPick_(row, data.headers, ['Cliente_ID','CLI_ID','Lead_ID']);
    const cli = ccPick_(row, data.headers, ['Cliente','Nombre','Razon_social']);
    const cat = ccPick_(row, data.headers, ['Categoria','Tipo']);
    const con = ccPick_(row, data.headers, ['Concepto','Descripcion']);
    const imp = ccPick_(row, data.headers, ['Importe','Total','Monto']);
    const fecha = ccPick_(row, data.headers, ['Fecha','created_at']);
    const estado = ccPick_(row, data.headers, ['Estado','Status']);
    const notas = ccPick_(row, data.headers, ['Notas','Observaciones']);
    const updated = new Date().toISOString();
    return [id, cliId, cli, cat, con, imp, fecha, estado, notas, (sh.getName()||''), updated];
  });

  if (rows.length){
    view.getRange(2,1,rows.length,headers.length).setValues(rows);
  }
}

function ccMergePresupuestoSources_(shPres, shHist) {
  const presData = shPres ? ccGetSheetData_(shPres) : { headers: [], rows: [] };
  const histData = shHist ? ccGetSheetData_(shHist) : { headers: [], rows: [] };
  if (!presData.rows.length) return histData;
  if (!histData.rows.length) return presData;

  const presIdHeader = ccFindHeader_(presData.headers, ['Pres_ID','Presupuesto_ID','ID']);
  const histIdHeader = ccFindHeader_(histData.headers, ['Pres_ID','Presupuesto_ID','ID']);
  const presIds = new Set();
  presData.rows.forEach((row) => {
    const id = presIdHeader ? String(row[presIdHeader] || '').trim() : '';
    if (id) presIds.add(id);
  });

  const mergedRows = presData.rows.slice();
  histData.rows.forEach((row) => {
    const id = histIdHeader ? String(row[histIdHeader] || '').trim() : '';
    if (!id || !presIds.has(id)) mergedRows.push(row);
  });

  return { headers: presData.headers, rows: mergedRows };
}

function ccWriteView_(sh, headers, rows) {
  if (!sh) return;
  sh.clearContents();
  if (headers && headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  if (rows && rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function runSpreadsheetAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = new Date();
  const rows = [];
  const add = (section, sheet, issue, details) => {
    rows.push([timestamp, section, sheet || '', issue || '', details || '']);
  };

  const sheets = ss.getSheets();
  sheets.forEach((sh) => {
    const data = ccGetSheetData_(sh);
    const headerRow = data.headerRow;
    const headers = data.headers;
    const emptyCols = ccFindEmptyColumns_(data);
    const types = ccDetectColumnTypes_(data);

    add('SHEET', sh.getName(), 'SUMMARY', JSON.stringify({
      headerRow,
      headers,
      rows: data.rows.length,
      cols: headers.length
    }));

    if (emptyCols.length) {
      add('SHEET', sh.getName(), 'EMPTY_COLUMNS', emptyCols.join(', '));
    }

    add('SHEET', sh.getName(), 'COLUMN_TYPES', JSON.stringify(types));

    CC_AUDIT_KEYS.forEach((key) => {
      const dup = ccDetectDuplicates_(data, key);
      if (dup.count > 0) {
        add('DUPLICATES', sh.getName(), key, `count=${dup.count} values=${dup.values.join(', ')}`);
      }
    });
  });

  ccAuditRelationships_(ss, add);
  ccAuditEstados_(ss, add);
  ccAuditPdfLinks_(ss, add);

  ccWriteAudit_(ss, rows);
  return { ok: true, rows: rows.length };
}

function ccAuditRelationships_(ss, add) {
  const presIds = ccCollectIds_(ss, ['PRESUPUESTOS', 'HISTORIAL_PRESUPUESTOS'], 'Pres_ID');
  const factIds = ccCollectIds_(ss, ['HISTORIAL','FACTURAS','FACTURA'], 'Factura_ID');

  const presLineas = ccCollectIds_(ss, ['PRES_LINEAS', 'LINEAS_PRES_HIST'], 'Pres_ID', true);
  const missingPres = presLineas.filter(id => id && !presIds.has(id));
  if (missingPres.length) {
    add('REL', 'PRES_LINEAS', 'ORPHAN_PRES_ID', `count=${missingPres.length}`);
  }

  const factLineas = ccCollectIds_(ss, ['FACT_LINEAS'], 'Factura_ID', true);
  const lineasLegacy = ccCollectIds_(ss, ['LINEAS'], 'Numero_factura', true);
  const facturaRefs = factLineas.concat(lineasLegacy);
  const missingFact = facturaRefs.filter(id => id && !factIds.has(id));
  if (missingFact.length) {
    add('REL', 'FACT_LINEAS', 'ORPHAN_FACTURA_ID', `count=${missingFact.length}`);
  }
}

function ccAuditEstados_(ss, add) {
  ccAuditEstadoSheet_(ss, 'PRESUPUESTOS', CC_PRES_ESTADOS, ['Fecha_envio','Fecha_aceptacion'], add);
  ccAuditEstadoSheet_(ss, 'HISTORIAL_PRESUPUESTOS', CC_PRES_ESTADOS, ['Fecha_envio','Fecha_aceptacion'], add);
  ccAuditEstadoSheet_(ss, 'FACTURAS', CC_FACT_ESTADOS, ['Fecha_envio','Fecha_pago'], add);
  ccAuditEstadoSheet_(ss, 'FACTURA', CC_FACT_ESTADOS, ['Fecha_envio','Fecha_pago'], add);
  ccAuditEstadoSheet_(ss, 'LEADS', CC_LEAD_ESTADOS, [], add);
}

function ccAuditEstadoSheet_(ss, sheetName, allowed, dateFields, add) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const data = ccGetSheetData_(sh);
  const estadoHeader = ccFindHeader_(data.headers, ['Estado']);
  if (!estadoHeader) return;

  const invalid = [];
  const missingDates = {};
  dateFields.forEach((f) => { missingDates[f] = 0; });

  data.rows.forEach((row) => {
    const estado = ccNormalizeEstado_(row[estadoHeader]);
    if (estado && allowed.indexOf(estado) === -1) invalid.push(estado);

    if (estado === 'ENVIADO' && dateFields.indexOf('Fecha_envio') !== -1) {
      if (!row['Fecha_envio']) missingDates['Fecha_envio'] += 1;
    }
    if (estado === 'ACEPTADO' && dateFields.indexOf('Fecha_aceptacion') !== -1) {
      if (!row['Fecha_aceptacion']) missingDates['Fecha_aceptacion'] += 1;
    }
    if (estado === 'ENVIADA' && dateFields.indexOf('Fecha_envio') !== -1) {
      if (!row['Fecha_envio']) missingDates['Fecha_envio'] += 1;
    }
    if (estado === 'PAGADA' && dateFields.indexOf('Fecha_pago') !== -1) {
      if (!row['Fecha_pago']) missingDates['Fecha_pago'] += 1;
    }
  });

  if (invalid.length) {
    add('ESTADO', sheetName, 'INVALID', `count=${invalid.length} sample=${invalid.slice(0, 10).join(', ')}`);
  }

  Object.keys(missingDates).forEach((k) => {
    if (missingDates[k] > 0) {
      add('FECHAS', sheetName, `MISSING_${k}`, `count=${missingDates[k]}`);
    }
  });
}

function ccAuditPdfLinks_(ss, add) {
  const check = (sheetName, estados) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const data = ccGetSheetData_(sh);
    const estadoHeader = ccFindHeader_(data.headers, ['Estado']);
    const pdfHeader = ccFindHeader_(data.headers, ['PDF_link','Pdf_link']);
    if (!estadoHeader || !pdfHeader) return;

    let missing = 0;
    data.rows.forEach((row) => {
      const estado = ccNormalizeEstado_(row[estadoHeader]);
      if (estados.indexOf(estado) === -1) return;
      const pdf = String(row[pdfHeader] || '').trim();
      if (!pdf) missing += 1;
    });
    if (missing > 0) add('PDF', sheetName, 'MISSING_PDF', `count=${missing}`);
  };

  check('PRESUPUESTOS', ['ENVIADO','ACEPTADO']);
  check('HISTORIAL_PRESUPUESTOS', ['ENVIADO','ACEPTADO']);
  check('FACTURAS', ['EMITIDA','ENVIADA','PAGADA']);
  check('FACTURA', ['EMITIDA','ENVIADA','PAGADA']);
}

function ccWriteAudit_(ss, rows) {
  const sh = ss.getSheetByName('AI_AUDIT') || ss.getSheetByName('AI_LOG') || ss.insertSheet('AI_AUDIT');
  const headers = ['Timestamp','Section','Sheet','Issue','Details'];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }
}

function ccGetSheet_(name, createIfMissing) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh && createIfMissing) sh = ss.insertSheet(name);
  return sh;
}

function ccGetHeaders_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
}

function ccGetSheetData_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { headerRow: 1, headers: [], rows: [] };

  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  let headerRow = 1;
  for (let i = 0; i < values.length; i++) {
    const nonEmpty = values[i].filter(v => String(v || '').trim() !== '');
    if (nonEmpty.length >= 2) {
      headerRow = i + 1;
      break;
    }
  }
  const headers = values[headerRow - 1].map(h => String(h || '').trim());
  const rows = values.slice(headerRow)
    .filter(r => r.some(c => c !== '' && c !== null && c !== undefined))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });

  return { headerRow, headers, rows };
}

function ccFindEmptyColumns_(data) {
  const empty = [];
  data.headers.forEach((h) => {
    const has = data.rows.some(r => String(r[h] || '').trim() !== '');
    if (!has) empty.push(h || '(sin nombre)');
  });
  return empty;
}

function ccDetectColumnTypes_(data) {
  const out = {};
  data.headers.forEach((h) => {
    let date = 0;
    let num = 0;
    let text = 0;
    data.rows.forEach((r) => {
      const v = r[h];
      if (v === '' || v === null || v === undefined) return;
      if (v instanceof Date) { date += 1; return; }
      if (typeof v === 'number' && !isNaN(v)) { num += 1; return; }
      text += 1;
    });
    const type = (date && !num && !text) ? 'date'
      : (num && !date && !text) ? 'number'
      : (text && !date && !num) ? 'text'
      : (date || num || text) ? 'mixed' : 'empty';
    out[h || '(sin nombre)'] = { type, date, number: num, text };
  });
  return out;
}

function ccDetectDuplicates_(data, header) {
  const key = ccFindHeader_(data.headers, [header]);
  if (!key) return { count: 0, values: [] };

  const seen = {};
  const dups = {};
  data.rows.forEach((r) => {
    const v = String(r[key] || '').trim();
    if (!v) return;
    if (seen[v]) dups[v] = true;
    seen[v] = true;
  });
  const values = Object.keys(dups);
  return { count: values.length, values };
}

function ccCollectIds_(ss, sheetNames, header, returnArray) {
  const set = new Set();
  const arr = [];
  sheetNames.forEach((name) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const data = ccGetSheetData_(sh);
    const key = ccFindHeader_(data.headers, [header]);
    if (!key) return;
    data.rows.forEach((r) => {
      const id = String(r[key] || '').trim();
      if (!id) return;
      set.add(id);
      if (returnArray) arr.push(id);
    });
  });
  return returnArray ? arr : set;
}

function ccFindHeader_(headers, candidates) {
  const map = {};
  headers.forEach((h) => {
    const k = ccNormalizeKey_(h);
    if (!map[k]) map[k] = h;
  });
  for (let i = 0; i < candidates.length; i++) {
    const key = ccNormalizeKey_(candidates[i]);
    if (map[key]) return map[key];
  }
  return '';
}

function ccPick_(row, headers, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const direct = row[candidates[i]];
    if (direct !== undefined && direct !== null && String(direct).trim() !== '') return direct;
    const found = ccFindHeader_(headers, [candidates[i]]);
    if (found && row[found] !== undefined && row[found] !== null && String(row[found]).trim() !== '') {
      return row[found];
    }
  }
  return '';
}

function ccNormalizeKey_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '');
}

function ccNormalizeEstado_(value) {
  return String(value || '').trim().toUpperCase();
}

function ccFormatDateIso_(value) {
  const d = ccParseDate_(value);
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function ccParseDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (!value) return null;
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function ccLatestDate_(values) {
  const dates = (values || [])
    .map(v => ccParseDate_(v))
    .filter(d => d);
  if (!dates.length) return null;
  dates.sort((a, b) => b.getTime() - a.getTime());
  return dates[0];
}

function ccToNumber_(value) {
  const n = Number(value);
  return isNaN(n) ? null : n;
}

function ccBuildSearchText_(parts) {
  return (parts || [])
    .map(p => String(p || '').trim())
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
}





function ccRebuildViews() {
  try {
    PropertiesService.getScriptProperties().setProperty(CC_VIEWS_DIRTY_KEY, '1');
  } catch(e){}
  // fuerza rebuild completo
  return ccEnsureViews_(true);
}


