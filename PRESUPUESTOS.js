/*************************************************
 * PRESUPUESTOS / PROFORMAS - Costa Clean (PRO)
 * Archivo: PRESUPUESTOS.gs
 *
 * ✅ PDF automático (sin pedir nombre)
 * ✅ Carpeta fija en Drive (override)
 * ✅ Archiva líneas: PRES_LINEAS -> LINEAS_PRES_HIST
 * ✅ Archiva presupuesto: PRESUPUESTOS -> HISTORIAL_PRESUPUESTOS
 * ✅ Convertir a factura desde HISTORIAL (lee líneas desde LINEAS_PRES_HIST)
 * ✅ Email + WhatsApp (modal + link)
 * ✅ Página AI + integración OpenAI (Responses API /v1/responses)
 *
 * NOTA: Este archivo NO rompe tu CRM_NUEVO.
 * Usa tus funciones si existen:
 * - consumirSiguienteNumero_()
 * - workflowGenerarFactura()
 *************************************************/

/** ====== AJUSTES ====== */

// Carpeta fija pedida (override)
const PRES_FOLDER_ID_OVERRIDE = CC_DEFAULT_IDS.PRESUPUESTOS_FOLDER_ID;

// Cuántas líneas reserva al crear presupuesto
const PRES_LINEAS_PRECREADAS = 5;

// Hojas
const SH_PRES            = 'PRESUPUESTOS';
const SH_PRES_LINEAS     = 'PRES_LINEAS';
const SH_PRES_LINEAS_HIS = 'LINEAS_PRES_HIST';
const SH_PRES_HIST       = 'HISTORIAL_PRESUPUESTOS';

// AI
const SH_AI_CONFIG = 'AI_CONFIG';
const SH_AI_LOG    = 'AI_LOG';
const PROP_OPENAI_KEY = 'OPENAI_API_KEY';
const PROP_PRES_PDF_FOLDER_ID = 'PRES_Pdf_Folder_Id';
const PROP_PRES_TEMPLATE_ID   = 'PRES_Template_DocId';
const DEFAULT_PRES_PDF_FOLDER_NAME = 'Costa Clean - Presupuestos PDF';
const DEFAULT_PRES_TEMPLATE_NAME   = 'Plantilla Presupuesto Costa Clean';

// Estados dropdown (hoja PRESUPUESTOS)
const PRES_ESTADOS = ['BORRADOR', 'ENVIADO', 'ACEPTADO', 'PERDIDO'];

// Cabeceras esperadas (según tu estructura A..T y añadimos U=Archivado_el para profesionalizar)
const PRES_HEADERS = [
  'Pres_ID','Fecha','Validez_dias','Vence_el','Estado','Cliente_ID',
  'Cliente','Email_cliente','NIF','Direccion','CP','Ciudad',
  'Base','IVA_total','Total','Notas','PDF_link','Factura_ID','Fecha_envio','Fecha_aceptacion','Archivado_el'
];

// Cabecera líneas (A..H) y archivado extra
const PRES_LINEAS_HEADERS = ['Pres_ID','Linea_n','Concepto','Cantidad','Precio','Dto_%','IVA_%','Subtotal'];
const PRES_LINEAS_HIST_HEADERS = ['Pres_ID','Linea_n','Concepto','Cantidad','Precio','Dto_%','IVA_%','Subtotal','Archivado_el'];

/** =========================
 * MENÚ PRESUPUESTOS
 * ========================= */
function menuPresupuestos_() {
  SpreadsheetApp.getUi()
    .createMenu('Presupuestos')
    .addItem('➕ Crear presupuesto (nuevo)', 'crearPresupuesto')
    .addSeparator()
    .addItem('📄 Generar PDF (fila seleccionada) + Archivar', 'uiGenerarPdfPresupuesto')
    .addItem('✉️ Enviar email (fila seleccionada)', 'uiEnviarPresupuestoEmail')
    .addItem('💬 WhatsApp (fila seleccionada)', 'uiWhatsAppPresupuesto')
    .addSeparator()
    .addItem('✅ Marcar como Enviado (fila seleccionada)', 'uiMarcarPresupuestoEnviado')
    .addItem('✅ Marcar como Aceptado (fila seleccionada)', 'uiMarcarPresupuestoAceptado')
    .addItem('⛔ Marcar como Perdido (fila seleccionada)', 'uiMarcarPresupuestoPerdido')
    .addSeparator()
    .addItem('🧾 Convertir a factura (desde HISTORIAL, fila seleccionada)', 'uiConvertirPresupuestoAFactura')
    .addSeparator()
    .addItem('🧠 AI: Crear página AI', 'aiCrearPagina_')
    .addItem('🧠 AI: Configurar API Key', 'aiConfigurarApiKey_')
    .addItem('🧠 AI: Generar texto email/WhatsApp (fila seleccionada)', 'aiGenerarTextoComercial_')
    .addToUi();
}

/** =========================
 * UTIL: asegurar hojas / cabeceras / validaciones
 * ========================= */
function presAsegurarEstructura_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // PRESUPUESTOS
  let sh = ss.getSheetByName(SH_PRES);
  if (!sh) sh = ss.insertSheet(SH_PRES);
  presAsegurarHeaders_(sh, PRES_HEADERS);
  presAsegurarLeadColumns_(sh);

  // PRES_LINEAS
  let shL = ss.getSheetByName(SH_PRES_LINEAS);
  if (!shL) shL = ss.insertSheet(SH_PRES_LINEAS);
  presAsegurarHeaders_(shL, PRES_LINEAS_HEADERS);

  // LINEAS_PRES_HIST
  let shLH = ss.getSheetByName(SH_PRES_LINEAS_HIS);
  if (!shLH) shLH = ss.insertSheet(SH_PRES_LINEAS_HIS);
  presAsegurarHeaders_(shLH, PRES_LINEAS_HIST_HEADERS);

  // HISTORIAL_PRESUPUESTOS
  let shH = ss.getSheetByName(SH_PRES_HIST);
  if (!shH) shH = ss.insertSheet(SH_PRES_HIST);
  presAsegurarHeaders_(shH, PRES_HEADERS);

  // Validaciones dinámicas
  presApplyValidations_();
}

function presAsegurarHeaders_(sh, headers) {
  const first = sh.getRange(1,1,1,headers.length).getValues()[0];
  const empty = first.every(v => String(v || '').trim() === '');
  if (empty) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    // Si hay headers pero con menos columnas, no rompas: solo asegura ancho mínimo
    if (sh.getLastColumn() < headers.length) sh.insertColumnsAfter(sh.getLastColumn(), headers.length - sh.getLastColumn());
  }
}

function presAsegurarLeadColumns_(sh) {
  const leadHeaders = ['Tipo_destinatario','Lead_ID','Lead_RowKey','Lead_Nombre','Lead_Email','Lead_NIF','Lead_Telefono','Lead_Direccion'];
  const headers = presGetHeaders_(sh);
  const missing = leadHeaders.filter((h) => headers.indexOf(h) === -1);
  if (!missing.length) return;

  const lastCol = sh.getLastColumn();
  sh.insertColumnsAfter(lastCol, missing.length);
  sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
}

function presAsegurarValidacionEstado_(sh, col) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PRES_ESTADOS, true)
    .setAllowInvalid(true)
    .build();

  // aplica de fila 2 a 1000 (puedes subir)
  sh.getRange(2, col, Math.max(1, sh.getMaxRows()-1), 1).setDataValidation(rule);
}

function presApplyValidations_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPres = ss.getSheetByName(SH_PRES);
  const shHist = ss.getSheetByName(SH_PRES_HIST);
  const shCli = ss.getSheetByName('CLIENTES');
  const shLeads = ss.getSheetByName('LEADS');

  [shPres, shHist].forEach((sheet) => {
    if (!sheet) return;
    const { map } = presGetHeaderMap_(sheet);
    const maxRows = Math.max(1, sheet.getMaxRows() - 1);

    if (map['Estado']) {
      const ruleEstado = SpreadsheetApp.newDataValidation()
        .requireValueInList(PRES_ESTADOS, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, map['Estado'], maxRows, 1).setDataValidation(ruleEstado);
    }

    const cliRule = presBuildValidationRuleFromSheet_(shCli, 'Cliente_ID');
    if (map['Cliente_ID'] && cliRule) {
      sheet.getRange(2, map['Cliente_ID'], maxRows, 1).setDataValidation(cliRule);
    } else if (map['Cliente_ID']) {
      sheet.getRange(2, map['Cliente_ID'], maxRows, 1).clearDataValidations();
    }

    const leadRule = presBuildValidationRuleFromSheet_(shLeads, 'Lead_ID');
    if (map['Lead_ID'] && leadRule) {
      sheet.getRange(2, map['Lead_ID'], maxRows, 1).setDataValidation(leadRule);
    } else if (map['Lead_ID']) {
      sheet.getRange(2, map['Lead_ID'], maxRows, 1).clearDataValidations();
    }
  });
}

function presBuildValidationRuleFromSheet_(sheet, headerName) {
  if (!sheet) return null;
  const { map } = presGetHeaderMap_(sheet);
  const col = map[headerName] || 0;
  const lastRow = sheet.getLastRow();
  if (!col || lastRow < 2) return null;

  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange(2, col, lastRow - 1, 1), true)
    .setAllowInvalid(true)
    .build();
}

function setupValidationsPresupuestos() {
  presAsegurarEstructura_();
  presApplyValidations_();
}

function presGetHeaders_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
}

function presGetHeaderMap_(sh) {
  const headers = presGetHeaders_(sh);
  const map = {};
  headers.forEach((h, i) => {
    map[h] = i + 1;
  });
  return { headers, map };
}

function presBuildSheetCache_(sh, keyHeader, extraKeys) {
  if (!sh) return { headers: [], headerMap: {}, byId: {}, extra: {} };

  const { headers, map } = presGetHeaderMap_(sh);
  const keyCol = map[keyHeader] || 0;
  if (!keyCol) return { headers, headerMap: map, byId: {}, extra: {} };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { headers, headerMap: map, byId: {}, extra: {} };

  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const byId = {};
  const extra = {};
  (extraKeys || []).forEach((k) => { extra[k] = {}; });

  data.forEach((row, idx) => {
    const id = String(row[keyCol - 1] || '').trim();
    if (!id) return;
    const entry = { row: idx + 2, values: row };
    byId[id] = entry;

    (extraKeys || []).forEach((k) => {
      const col = map[k] || 0;
      if (!col) return;
      const val = String(row[col - 1] || '').trim();
      if (val) extra[k][val] = entry;
    });
  });

  return { headers, headerMap: map, byId, extra };
}

function presBuildClientesCache_(ss) {
  const shCli = ss.getSheetByName('CLIENTES');
  return presBuildSheetCache_(shCli, 'Cliente_ID', []);
}

function presBuildLeadsCache_(ss) {
  const shLeads = ss.getSheetByName('LEADS');
  const cache = presBuildSheetCache_(shLeads, 'Lead_ID', ['RowKey']);
  return {
    headers: cache.headers,
    headerMap: cache.headerMap,
    byId: cache.byId,
    byRowKey: cache.extra.RowKey || {}
  };
}

function presExtractClienteFromCache_(entry, headerMap) {
  if (!entry || !headerMap) return null;
  const row = entry.values;
  return {
    clienteId: String(row[(headerMap['Cliente_ID'] || 1) - 1] || '').trim(),
    nombre: String(row[(headerMap['Nombre'] || 2) - 1] || '').trim(),
    email: String(row[(headerMap['Email'] || 8) - 1] || '').trim(),
    nif: String(row[(headerMap['NIF'] || 3) - 1] || '').trim(),
    direccion: String(row[(headerMap['Direccion'] || 4) - 1] || '').trim(),
    cp: String(row[(headerMap['CP'] || 5) - 1] || '').trim(),
    ciudad: String(row[(headerMap['Ciudad'] || 6) - 1] || '').trim(),
    telefono: String(row[(headerMap['Telefono'] || 7) - 1] || '').trim()
  };
}

function presExtractLeadFromCache_(entry, headerMap) {
  if (!entry || !headerMap) return null;
  const row = entry.values;
  const get = (h, defIdx) => String(row[(headerMap[h] || defIdx) - 1] || '').trim();

  return {
    leadId: get('Lead_ID', 1),
    rowKey: get('RowKey', 26),
    nombre: get('Nombre', 3),
    email: get('Email', 4),
    telefono: get('Telefono', 5),
    nif: get('NIF/CIF', 6),
    direccion: get('Direccion', 7),
    cp: get('CP', 8),
    ciudad: get('Poblacion', 9),
    clienteId: get('Cliente_ID', 23)
  };
}

function presSetValuesByHeaders_(sh, row, headerMap, values) {
  if (!sh || !headerMap || !values) return;
  Object.keys(values).forEach((header) => {
    const col = headerMap[header];
    if (!col) return;
    sh.getRange(row, col).setValue(values[header]);
  });
}

function presClearByHeaders_(sh, row, headerMap, headers) {
  if (!sh || !headerMap || !headers || !headers.length) return;
  headers.forEach((h) => {
    const col = headerMap[h];
    if (!col) return;
    sh.getRange(row, col).clearContent();
  });
}

function presBuildRow_(headers, values) {
  return headers.map((h) => (Object.prototype.hasOwnProperty.call(values, h) ? values[h] : ''));
}

function presPickValue_(obj, candidates) {
  if (!obj) return '';
  const keys = Object.keys(obj);
  for (let i = 0; i < candidates.length; i++) {
    const target = String(candidates[i] || '').toLowerCase();
    const match = keys.find((k) => String(k || '').toLowerCase() === target);
    if (match) return obj[match];
  }
  return '';
}

function presToDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function presGetLeadSelection_(ss) {
  const sh = ss.getActiveSheet();
  if (sh.getName() !== 'LEADS') return null;
  const row = ss.getActiveRange().getRow();
  if (row < 2) return null;

  const data = sh.getRange(row, 1, 1, 26).getValues()[0];
  const leadId = String(data[0] || '').trim();
  const rowKey = String(data[25] || '').trim();
  if (!leadId) return null;

  return {
    row,
    leadId,
    rowKey,
    nombre: String(data[2] || '').trim(),
    email: String(data[3] || '').trim(),
    telefono: String(data[4] || '').trim(),
    nif: String(data[5] || '').trim(),
    direccion: String(data[6] || '').trim(),
    cp: String(data[7] || '').trim(),
    ciudad: String(data[8] || '').trim()
  };
}

function presFindLeadRowByRowKey_(shLeads, rowKey) {
  const lastRow = shLeads.getLastRow();
  if (lastRow < 2) return 0;

  const data = shLeads.getRange(2, 26, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === String(rowKey || '').trim()) {
      return i + 2;
    }
  }
  return 0;
}

function presFindLeadRowById_(shLeads, leadId) {
  const lastRow = shLeads.getLastRow();
  if (lastRow < 2) return 0;

  const data = shLeads.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === String(leadId || '').trim()) {
      return i + 2;
    }
  }
  return 0;
}

function presGetLeadData_(shLeads, leadRow) {
  const data = shLeads.getRange(leadRow, 1, 1, 26).getValues()[0];
  return {
    leadId: String(data[0] || '').trim(),
    rowKey: String(data[25] || '').trim(),
    nombre: String(data[2] || '').trim(),
    email: String(data[3] || '').trim(),
    telefono: String(data[4] || '').trim(),
    nif: String(data[5] || '').trim(),
    direccion: String(data[6] || '').trim(),
    cp: String(data[7] || '').trim(),
    ciudad: String(data[8] || '').trim()
  };
}

/** =========================
 * CONFIG (lee tu CONFIG por headers como ya tenías)
 * ========================= */
function getCfg_(header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('CONFIG');
  if (!cfg) throw new Error('No existe hoja CONFIG');

  const headers = cfg.getRange(1,1,1,cfg.getLastColumn()).getDisplayValues()[0];
  const col = headers.indexOf(header) + 1;
  if (col < 1) throw new Error('CONFIG: no existe ' + header);

  return String(cfg.getRange(2, col).getDisplayValue()).trim();
}

function getCfgNum_(header, fallback) {
  const v = Number(getCfg_(header).replace(',', '.'));
  return isNaN(v) ? fallback : v;
}

function getCfgOptional_(header) {
  try {
    return getCfg_(header);
  } catch (_) {
    return '';
  }
}

function getCfgAny_(headers) {
  for (let i = 0; i < headers.length; i++) {
    const v = getCfgOptional_(headers[i]);
    if (String(v || '').trim()) return v;
  }
  return '';
}

function getCfgFromSheet_(header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('CONFIG');
  if (!cfg) return '';

  const headers = cfg.getRange(1, 1, 1, Math.max(cfg.getLastColumn(), 1)).getDisplayValues()[0];
  const col = headers.indexOf(header) + 1;
  if (col < 1) return '';
  return String(cfg.getRange(2, col).getDisplayValue()).trim();
}

function setCfgValueIfSheet_(header, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('CONFIG');
  if (!cfg) return false;

  const headers = cfg.getRange(1, 1, 1, Math.max(cfg.getLastColumn(), 1)).getDisplayValues()[0];
  let col = headers.indexOf(header) + 1;
  if (col < 1) {
    const lastCol = Math.max(cfg.getLastColumn(), 1);
    cfg.insertColumnsAfter(lastCol, 1);
    col = lastCol + 1;
    cfg.getRange(1, col).setValue(header);
  }

  cfg.getRange(2, col).setValue(value || '');
  return true;
}

function presFindHeaderIndex_(headers, candidates) {
  const lower = (headers || []).map((h) => String(h || '').toLowerCase());
  for (let i = 0; i < candidates.length; i++) {
    const idx = lower.indexOf(String(candidates[i] || '').toLowerCase());
    if (idx !== -1) return idx + 1;
  }
  return 0;
}

function presPickHeaderName_(headers, candidates) {
  const lower = (headers || []).map((h) => String(h || '').toLowerCase());
  for (let i = 0; i < candidates.length; i++) {
    const idx = lower.indexOf(String(candidates[i] || '').toLowerCase());
    if (idx !== -1) return headers[idx];
  }
  return '';
}

function getPdfConfig_() {
  const props = PropertiesService.getScriptProperties();
  const folderId = getCfgAny_(['PRES_Pdf_Folder_Id']) || props.getProperty(PROP_PRES_PDF_FOLDER_ID) || '';
  const templateId = getCfgAny_(['PRES_Template_DocId']) || props.getProperty(PROP_PRES_TEMPLATE_ID) || CC_DEFAULT_IDS.PRESUPUESTO_TEMPLATE_ID || '';

  const resolvedFolderId = folderId || props.getProperty('PRES_FOLDER_ID_OVERRIDE') || PRES_FOLDER_ID_OVERRIDE || '';
  return { folderId: resolvedFolderId, templateId: templateId || '' };
}

function savePdfConfigIds_(folderId, templateId) {
  const props = PropertiesService.getScriptProperties();
  if (folderId) props.setProperty(PROP_PRES_PDF_FOLDER_ID, folderId);
  if (templateId) props.setProperty(PROP_PRES_TEMPLATE_ID, templateId);

  setCfgValueIfSheet_('PRES_Pdf_Folder_Id', folderId || '');
  setCfgValueIfSheet_('PRES_Template_DocId', templateId || '');
}

function createDefaultPresTemplate_(folder) {
  const doc = DocumentApp.create(DEFAULT_PRES_TEMPLATE_NAME);
  const body = doc.getBody();

  body.appendParagraph('Presupuesto {{Pres_ID}}').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('');
  body.appendParagraph('Datos del cliente').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendTable([
    ['Fecha', '{{Fecha}}'],
    ['Cliente', '{{Cliente}}'],
    ['Email', '{{Email_cliente}}'],
    ['Direccion', '{{Direccion}}'],
    ['Ciudad', '{{Ciudad}}'],
    ['Base imponible', '{{Base}}'],
    ['Total', '{{Total}}']
  ]);
  body.appendParagraph('Notas').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{Notas}}');
  body.appendParagraph('');
  body.appendParagraph('Lineas del presupuesto').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{LINEAS_TABLA}}');

  doc.saveAndClose();

  if (folder) {
    const file = DriveApp.getFileById(doc.getId());
    try { file.moveTo(folder); } catch (_) { folder.addFile(file); }
  }

  return doc.getId();
}

function setupPdfSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = (resultado, mensaje, data) => {
    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PRESUPUESTOS', 'SETUP_PDF', 'SYSTEM', '', resultado, mensaje || '', data || null);
    }
  };

  try {
    presAsegurarEstructura_();
    const cfg = getPdfConfig_();

    let folder = null;
    if (cfg.folderId) {
      try { folder = DriveApp.getFolderById(cfg.folderId); } catch (_) {}
    }
    let createdFolder = false;
    if (!folder) {
      const it = DriveApp.getFoldersByName(DEFAULT_PRES_PDF_FOLDER_NAME);
      folder = it.hasNext() ? it.next() : DriveApp.createFolder(DEFAULT_PRES_PDF_FOLDER_NAME);
      createdFolder = true;
    }

    let templateId = cfg.templateId;
    let createdTemplate = false;
    if (templateId) {
      try { DriveApp.getFileById(templateId); } catch (_) { templateId = ''; }
    }
    if (!templateId) {
      templateId = createDefaultPresTemplate_(folder);
      createdTemplate = true;
    }

    savePdfConfigIds_(folder.getId(), templateId);
    log('OK', '', { folderId: folder.getId(), templateId, createdFolder, createdTemplate });
    return { ok: true, folderId: folder.getId(), templateId, createdFolder, createdTemplate };
  } catch (err) {
    log('ERROR', err.message || String(err), { stack: err.stack });
    throw err;
  }
}

function getConfigPres_() {
  const validezRaw = getCfgAny_(['PRES_Validez_default']);
  const validezDefault = Number(String(validezRaw || '').replace(',', '.')) || 15;
  return {
    anio: getCfgAny_(['PRES_Año', 'PRES_A?o', 'PRES_Ano']) || String(new Date().getFullYear()),
    validezDefault: validezDefault,
    carpetaId: getPdfConfig_().folderId || PRES_FOLDER_ID_OVERRIDE,
    templateId: getPdfConfig_().templateId || ''
  };
}

function consumirSiguientePresId_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('CONFIG');

  const headers = cfg.getRange(1,1,1,Math.max(cfg.getLastColumn(),1)).getDisplayValues()[0];
  const colUlt = presFindHeaderIndex_(headers, ['PRES_Ultimo_numero']);
  const colAnio = presFindHeaderIndex_(headers, ['PRES_Año','PRES_A?o','PRES_Ano']);
  if (!colUlt || !colAnio) throw new Error('CONFIG: faltan columnas PRES_Ultimo_numero / PRES_Año');

  const anio = String(cfg.getRange(2, colAnio).getDisplayValue()).trim();
  const ultimo = Number(cfg.getRange(2, colUlt).getValue()) || 0;

  const next = ultimo + 1;
  cfg.getRange(2, colUlt).setValue(next);

  return `PRO-${anio}-${String(next).padStart(4,'0')}`;
}

/** =========================
 * CREAR PRESUPUESTO
 * ========================= */
function crearPresupuestoParaLead_(leadRowOrId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shLeads = ss.getSheetByName('LEADS');
  if (!shLeads) throw new Error('No existe hoja LEADS');

  let leadRow = 0;
  let leadId = '';

  try {
    if (typeof leadRowOrId === 'number') {
      leadRow = leadRowOrId;
    } else {
      leadId = String(leadRowOrId || '').trim();
      if (leadId) leadRow = presFindLeadRowById_(shLeads, leadId);
    }

    if (!leadRow) throw new Error('Lead no encontrado');

    presAsegurarEstructura_();
    const lead = presGetLeadData_(shLeads, leadRow);
    if (!lead.leadId) throw new Error('Lead_ID vacio');

    const shPres = ss.getSheetByName(SH_PRES);
    const presId = consumirSiguientePresId_();
    const cfg = getConfigPres_();
    const hoy = new Date();
    const vence = new Date(hoy.getTime() + cfg.validezDefault * 86400000);

    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PRESUPUESTOS', 'CREATE', 'LEAD', lead.leadId, 'START', '', { presId: presId });
    }

    const headerInfo = presGetHeaderMap_(shPres);
    const rowValues = presBuildRow_(headerInfo.headers, {
      Pres_ID: presId,
      Fecha: hoy,
      Validez_dias: cfg.validezDefault,
      Vence_el: vence,
      Estado: 'BORRADOR',
      Cliente_ID: '',
      Cliente: lead.nombre,
      Email_cliente: lead.email,
      NIF: lead.nif,
      Direccion: lead.direccion,
      CP: lead.cp,
      Ciudad: lead.ciudad,
      Tipo_destinatario: 'LEAD',
      Lead_ID: lead.leadId,
      Lead_RowKey: lead.rowKey,
      Lead_Nombre: lead.nombre,
      Lead_Email: lead.email,
      Lead_NIF: lead.nif,
      Lead_Telefono: lead.telefono,
      Lead_Direccion: lead.direccion
    });

    shPres.appendRow(rowValues);

    const r = shPres.getLastRow();
    const colEstado = headerInfo.map['Estado'];
    if (colEstado) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(PRES_ESTADOS, true)
        .setAllowInvalid(true)
        .build();
      shPres.getRange(r, colEstado).setDataValidation(rule);
    }

    reservarLineasPres_(presId, PRES_LINEAS_PRECREADAS);

    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PRESUPUESTOS', 'CREATE', 'PRESUPUESTO', presId, 'OK', '', { leadId: lead.leadId });
    }

    return { presId: presId, row: r, leadId: lead.leadId };
  } catch (err) {
    const id = leadId || String(leadRowOrId || '');
    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PRESUPUESTOS', 'CREATE', 'LEAD', id, 'ERROR', err.message || String(err), null);
    }
    throw err;
  }
}

function crearPresupuesto() {
  presAsegurarEstructura_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_PRES);
  const leadInfo = presGetLeadSelection_(ss);

  if (leadInfo && leadInfo.row) {
    const res = crearPresupuestoParaLead_(leadInfo.row);
    SpreadsheetApp.getUi().alert(
      '? Presupuesto creado:\n' + res.presId +
      '\n\n?? Rellena las l¡neas en ' + SH_PRES_LINEAS +
      '\nLuego genera PDF desde la fila del presupuesto.'
    );
    return;
  }

  const presId = consumirSiguientePresId_();
  const { validezDefault } = getConfigPres_();

  const hoy = new Date();
  const vence = new Date(hoy.getTime() + validezDefault * 86400000);

  const headerInfo = presGetHeaderMap_(sh);
  const rowValues = presBuildRow_(headerInfo.headers, {
    Pres_ID: presId,
    Fecha: hoy,
    Validez_dias: validezDefault,
    Vence_el: vence,
    Estado: 'BORRADOR',
    Cliente_ID: '',
    Cliente: leadInfo ? leadInfo.nombre : '',
    Email_cliente: leadInfo ? leadInfo.email : '',
    NIF: leadInfo ? leadInfo.nif : '',
    Direccion: leadInfo ? leadInfo.direccion : '',
    CP: leadInfo ? leadInfo.cp : '',
    Ciudad: leadInfo ? leadInfo.ciudad : '',
    Tipo_destinatario: leadInfo ? 'LEAD' : '',
    Lead_RowKey: leadInfo ? leadInfo.rowKey : '',
    Lead_Nombre: leadInfo ? leadInfo.nombre : '',
    Lead_Email: leadInfo ? leadInfo.email : '',
    Lead_NIF: leadInfo ? leadInfo.nif : '',
    Lead_Telefono: leadInfo ? leadInfo.telefono : '',
    Lead_Direccion: leadInfo ? leadInfo.direccion : ''
  });

  sh.appendRow(rowValues);

  const r = sh.getLastRow();
  const colEstado = headerInfo.map['Estado'];
  if (colEstado) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(PRES_ESTADOS, true)
      .setAllowInvalid(true)
      .build();
    sh.getRange(r, colEstado).setDataValidation(rule);
  }

  // Reservar líneas
  reservarLineasPres_(presId, PRES_LINEAS_PRECREADAS);

  SpreadsheetApp.getUi().alert(
    '✅ Presupuesto creado:\n' + presId +
    '\n\n👉 Rellena las líneas en ' + SH_PRES_LINEAS +
    '\nLuego genera PDF desde la fila del presupuesto.'
  );
}

/** =========================
 * RESERVAR LÍNEAS (sin insertar filas)
 * ========================= */
function reservarLineasPres_(presId, n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_PRES_LINEAS);
  if (!sh) throw new Error('No existe ' + SH_PRES_LINEAS);

  const lastRow = Math.max(sh.getLastRow(), 2);
  let startRow = 2;

  if (lastRow >= 2) {
    const data = sh.getRange(2,1,lastRow-1,8).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i].some(v => String(v).trim() !== '')) {
        startRow = i + 3;
        break;
      }
    }
  }

  for (let i = 0; i < n; i++) {
    const r = startRow + i;

    sh.getRange(r,1).setValue(presId); // Pres_ID
    sh.getRange(r,2).setValue(i + 1);  // Linea_n
    sh.getRange(r,3).setValue('');     // Concepto
    sh.getRange(r,4).setValue('');     // Cantidad
    sh.getRange(r,5).setValue('');     // Precio
    sh.getRange(r,6).setValue(0);      // Dto %
    sh.getRange(r,7).setValue(21);     // IVA %

    // Subtotal
    sh.getRange(r,8).setFormula(
      `=IF(OR(C${r}="";D${r}="";E${r}="");"";ROUND(D${r}*E${r}*(1-(F${r}/100));2))`
    );
  }
}

/** =========================
 * AUTORELLENO CLIENTE (cuando editas Cliente_ID en col F)
 * ========================= */
function onEditPresupuestos_(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== SH_PRES) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

  const cache = CacheService.getScriptCache();
  const guardKey = 'pres_onedit_guard';
  if (cache && cache.get(guardKey)) return;

  const headerInfo = presGetHeaderMap_(sh);
  const colClienteId = headerInfo.map['Cliente_ID'] || 0;
  const colLeadId = headerInfo.map['Lead_ID'] || 0;
  const colEstado = headerInfo.map['Estado'] || 0;
  const relevantCols = [colClienteId, colLeadId, colEstado].filter(Boolean);
  if (relevantCols.indexOf(col) === -1) return;

  const lock = LockService.getUserLock();
  if (!lock.tryLock(5000)) return;
  if (cache) cache.put(guardKey, '1', 5);

  try {
    const ss = sh.getParent() || SpreadsheetApp.getActiveSpreadsheet();
    const caches = { clientes: null, leads: null };
    const ensureClientes = () => (caches.clientes || (caches.clientes = presBuildClientesCache_(ss)));
    const ensureLeads = () => (caches.leads || (caches.leads = presBuildLeadsCache_(ss)));

    if (col === colClienteId) {
      const cliId = String(e.range.getDisplayValue() || '').trim();
      presHandleClienteSelection_(sh, row, cliId, ensureClientes(), headerInfo.map);
    }

    if (col === colLeadId) {
      const leadId = String(e.range.getDisplayValue() || '').trim();
      presHandleLeadSelection_(sh, row, leadId, ensureLeads(), ensureClientes(), headerInfo.map);
    }

    if (col === colEstado) {
      const estado = String(e.range.getDisplayValue() || '').trim();
      presHandleEstadoChange_(sh, row, estado, headerInfo.map, ensureLeads(), ensureClientes(), e);
    }
  } finally {
    lock.releaseLock();
    if (cache) cache.remove(guardKey);
  }
}

function presHandleClienteSelection_(shPres, row, cliId, clientesCache, headerMapPres) {
  const targetHeaders = ['Cliente', 'Email_cliente', 'NIF', 'Direccion', 'CP', 'Ciudad'];
  const clearHeaders = targetHeaders.concat(['Tipo_destinatario']);
  const normalizedId = String(cliId || '').trim();

  if (!normalizedId) {
    presClearByHeaders_(shPres, row, headerMapPres, clearHeaders);
    return;
  }

  const entry = clientesCache && clientesCache.byId[normalizedId];
  if (!entry) {
    presClearByHeaders_(shPres, row, headerMapPres, clearHeaders);
    return;
  }

  const cliente = presExtractClienteFromCache_(entry, clientesCache.headerMap);
  if (!cliente) {
    presClearByHeaders_(shPres, row, headerMapPres, clearHeaders);
    return;
  }

  presSetValuesByHeaders_(shPres, row, headerMapPres, {
    Cliente_ID: normalizedId,
    Cliente: cliente.nombre,
    Email_cliente: cliente.email,
    NIF: cliente.nif,
    Direccion: cliente.direccion,
    CP: cliente.cp,
    Ciudad: cliente.ciudad
  });

  if (headerMapPres['Tipo_destinatario']) {
    presSetValuesByHeaders_(shPres, row, headerMapPres, { Tipo_destinatario: 'CLIENTE' });
  }
}

function presHandleLeadSelection_(shPres, row, leadId, leadCache, clientesCache, headerMapPres) {
  const leadFields = ['Lead_RowKey', 'Lead_Nombre', 'Lead_Email', 'Lead_NIF', 'Lead_Telefono', 'Lead_Direccion'];
  const normalizedId = String(leadId || '').trim();

  if (!normalizedId) {
    presClearByHeaders_(shPres, row, headerMapPres, leadFields.concat(['Tipo_destinatario']));
    return;
  }

  const entry = leadCache && leadCache.byId[normalizedId];
  if (!entry) {
    presClearByHeaders_(shPres, row, headerMapPres, leadFields);
    return;
  }

  const lead = presExtractLeadFromCache_(entry, leadCache.headerMap) || {};
  presSetValuesByHeaders_(shPres, row, headerMapPres, {
    Lead_ID: lead.leadId,
    Lead_RowKey: lead.rowKey,
    Lead_Nombre: lead.nombre,
    Lead_Email: lead.email,
    Lead_NIF: lead.nif,
    Lead_Telefono: lead.telefono,
    Lead_Direccion: lead.direccion
  });

  if (headerMapPres['Tipo_destinatario']) {
    presSetValuesByHeaders_(shPres, row, headerMapPres, { Tipo_destinatario: 'LEAD' });
  }

  presSetValuesByHeaders_(shPres, row, headerMapPres, {
    Cliente: lead.nombre,
    Email_cliente: lead.email,
    NIF: lead.nif,
    Direccion: lead.direccion,
    CP: lead.cp,
    Ciudad: lead.ciudad
  });

  if (lead.clienteId) {
    presSetValuesByHeaders_(shPres, row, headerMapPres, { Cliente_ID: lead.clienteId });
    const cliEntry = clientesCache && clientesCache.byId[lead.clienteId];
    if (cliEntry) {
      const cli = presExtractClienteFromCache_(cliEntry, clientesCache.headerMap);
      if (cli) {
        presSetValuesByHeaders_(shPres, row, headerMapPres, {
          Cliente: cli.nombre || lead.nombre,
          Email_cliente: cli.email || lead.email,
          NIF: cli.nif || lead.nif,
          Direccion: cli.direccion || lead.direccion,
          CP: cli.cp || lead.cp,
          Ciudad: cli.ciudad || lead.ciudad
        });
      }
    }
  }
}

function presHandleEstadoChange_(shPres, row, estadoRaw, headerMapPres, leadCache, clientesCache, e) {
  const estadoInput = String(estadoRaw || '').trim();
  const estadoUpper = estadoInput.toUpperCase();

  if (estadoUpper && headerMapPres['Estado'] && estadoUpper !== estadoInput) {
    shPres.getRange(row, headerMapPres['Estado']).setValue(estadoUpper);
  }

  if (estadoUpper.toLowerCase() === 'aceptado') {
    if (headerMapPres['Fecha_aceptacion']) {
      shPres.getRange(row, headerMapPres['Fecha_aceptacion']).setValue(new Date());
    }
    presAutoConvertLeadOnAccept_(shPres, row, leadCache, clientesCache);
    const oldEstado = e ? String(e.oldValue || '').trim().toUpperCase() : '';
    if (e && estadoUpper !== oldEstado) {
      try {
        const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
        ss.toast('Presupuesto ACEPTADO. Accion recomendada: Crear factura.', 'PRESUPUESTOS', 6);
      } catch (_) {}
    }
  }
}

function presAutoConvertLeadOnAccept_(shPres, rowPres, leadCache, clientesCache) {
  const ss = shPres.getParent ? shPres.getParent() : SpreadsheetApp.getActiveSpreadsheet();
  const headerInfo = presGetHeaderMap_(shPres);
  const colTipo = headerInfo.map['Tipo_destinatario'];
  const colLeadId = headerInfo.map['Lead_ID'];
  const colLeadKey = headerInfo.map['Lead_RowKey'];
  const colClienteId = headerInfo.map['Cliente_ID'] || 0;

  const leadId = colLeadId ? String(shPres.getRange(rowPres, colLeadId).getValue()).trim() : '';
  const leadRowKey = colLeadKey ? String(shPres.getRange(rowPres, colLeadKey).getValue()).trim() : '';
  const leadRef = leadId || leadRowKey || String(rowPres);
  const clienteIdActual = colClienteId ? String(shPres.getRange(rowPres, colClienteId).getValue()).trim() : '';

  if (!leadId && !leadRowKey) return;
  if (clienteIdActual) return;

  const tipo = colTipo ? String(shPres.getRange(rowPres, colTipo).getValue()).trim().toUpperCase() : '';
  if (tipo && tipo !== 'LEAD') return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) {
    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PRESUPUESTOS', 'AUTO_CONVERT', 'LEAD', leadRef, 'SKIP', 'lock busy', null);
    }
    return;
  }

  try {
    const currentClienteId = colClienteId ? String(shPres.getRange(rowPres, colClienteId).getValue()).trim() : '';
    if (currentClienteId) return;

    const leadsCache = leadCache || presBuildLeadsCache_(ss);
    const leadEntry = leadId ? leadsCache.byId[leadId] : (leadRowKey ? leadsCache.byRowKey[leadRowKey] : null);
    if (!leadEntry) throw new Error('Lead no encontrado en LEADS');

    const leadData = presExtractLeadFromCache_(leadEntry, leadsCache.headerMap) || {};
    if (leadData.clienteId) {
      if (colClienteId) shPres.getRange(rowPres, colClienteId).setValue(leadData.clienteId);
      rellenarClienteEnPresupuesto_(shPres, rowPres, leadData.clienteId, clientesCache);
      if (typeof logEvent_ === 'function') {
        logEvent_(ss, 'PRESUPUESTOS', 'AUTO_CONVERT', 'CLIENTE', leadData.clienteId, 'SKIP', 'cliente ya existente para lead', { leadId: leadData.leadId });
      }
      return;
    }

    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PRESUPUESTOS', 'AUTO_CONVERT', 'LEAD', leadRef, 'START', '', { leadRow: leadEntry.row });
    }

    if (typeof convertirLeadEnCliente_ !== 'function') {
      throw new Error('No existe convertirLeadEnCliente_');
    }

    convertirLeadEnCliente_(ss, leadEntry.row);

    const shLeads = ss.getSheetByName('LEADS');
    if (!shLeads) throw new Error('No existe hoja LEADS');

    const leadRowValues = shLeads.getRange(leadEntry.row, 1, 1, shLeads.getLastColumn()).getValues()[0];
    const colCliLead = leadsCache.headerMap['Cliente_ID'] || 23;
    const newClienteId = String(leadRowValues[colCliLead - 1] || '').trim();
    if (!newClienteId) throw new Error('Cliente_ID no generado en LEADS');

    if (colClienteId) shPres.getRange(rowPres, colClienteId).setValue(newClienteId);
    rellenarClienteEnPresupuesto_(shPres, rowPres, newClienteId, null);

    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PRESUPUESTOS', 'AUTO_CONVERT', 'CLIENTE', newClienteId, 'OK', '', { leadId: leadData.leadId, leadRowKey: leadRowKey || leadData.rowKey });
    }
  } catch (err) {
    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PRESUPUESTOS', 'AUTO_CONVERT', 'LEAD', leadRef, 'ERROR', err.message || String(err), null);
    }
  } finally {
    lock.releaseLock();
  }
}

function presGetClienteDataById_(ss, clienteId, clientesCache) {
  const cache = clientesCache || presBuildClientesCache_(ss);
  const entry = cache.byId[String(clienteId || '').trim()];
  if (!entry) return null;
  return presExtractClienteFromCache_(entry, cache.headerMap);
}

function presVincularPresupuestosPorLead_(ss, leadId, clienteId, clientesCache) {
  if (!leadId || !clienteId) return 0;
  try {
    const shPres = ss.getSheetByName(SH_PRES);
    if (!shPres) return 0;

    const headerInfo = presGetHeaderMap_(shPres);
    const colLeadId = headerInfo.map['Lead_ID'];
    const colClienteId = headerInfo.map['Cliente_ID'] || 0;
    if (!colLeadId || !colClienteId) return 0;

    const lastRow = shPres.getLastRow();
    if (lastRow < 2) return 0;

    const clienteData = presGetClienteDataById_(ss, clienteId, clientesCache);
    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PRESUPUESTOS', 'LINK', 'LEAD', leadId, 'START', '', { clienteId: clienteId });
    }

    let updated = 0;
    for (let row = 2; row <= lastRow; row++) {
      const rowLeadId = String(shPres.getRange(row, colLeadId).getValue()).trim();
      if (rowLeadId !== String(leadId || '').trim()) continue;

      const currentClienteId = String(shPres.getRange(row, colClienteId).getValue()).trim();
      if (!currentClienteId) {
        shPres.getRange(row, colClienteId).setValue(clienteId);
        updated++;
      }

      if (clienteData) {
        const colCliente = headerInfo.map['Cliente'];
        const colEmail = headerInfo.map['Email_cliente'];
        const colNif = headerInfo.map['NIF'];
        const colDireccion = headerInfo.map['Direccion'];
        const colCp = headerInfo.map['CP'];
        const colCiudad = headerInfo.map['Ciudad'];

        if (colCliente && !String(shPres.getRange(row, colCliente).getValue()).trim()) {
          shPres.getRange(row, colCliente).setValue(clienteData.nombre);
        }
        if (colEmail && !String(shPres.getRange(row, colEmail).getValue()).trim()) {
          shPres.getRange(row, colEmail).setValue(clienteData.email);
        }
        if (colNif && !String(shPres.getRange(row, colNif).getValue()).trim()) {
          shPres.getRange(row, colNif).setValue(clienteData.nif);
        }
        if (colDireccion && !String(shPres.getRange(row, colDireccion).getValue()).trim()) {
          shPres.getRange(row, colDireccion).setValue(clienteData.direccion);
        }
        if (colCp && !String(shPres.getRange(row, colCp).getValue()).trim()) {
          shPres.getRange(row, colCp).setValue(clienteData.cp);
        }
        if (colCiudad && !String(shPres.getRange(row, colCiudad).getValue()).trim()) {
          shPres.getRange(row, colCiudad).setValue(clienteData.ciudad);
        }
      }
    }

    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PRESUPUESTOS', 'LINK', 'PRESUPUESTO', clienteId, 'OK', '', { leadId: leadId, updated: updated });
    }

    return updated;
  } catch (err) {
    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PRESUPUESTOS', 'LINK', 'LEAD', leadId, 'ERROR', err.message || String(err), null);
    }
    return 0;
  }
}

function rellenarClienteEnPresupuesto_(shPres, rowPres, cliId, clientesCache) {
  const ss = shPres.getParent ? shPres.getParent() : SpreadsheetApp.getActiveSpreadsheet();
  const cache = clientesCache || presBuildClientesCache_(ss);
  const headerInfo = presGetHeaderMap_(shPres);
  const targetHeaders = ['Cliente', 'Email_cliente', 'NIF', 'Direccion', 'CP', 'Ciudad'];

  const entry = cache.byId[String(cliId || '').trim()];
  if (!entry) {
    presClearByHeaders_(shPres, rowPres, headerInfo.map, targetHeaders);
    return;
  }

  const cliente = presExtractClienteFromCache_(entry, cache.headerMap);
  if (!cliente) {
    presClearByHeaders_(shPres, rowPres, headerInfo.map, targetHeaders);
    return;
  }

  presSetValuesByHeaders_(shPres, rowPres, headerInfo.map, {
    Cliente: cliente.nombre,
    Email_cliente: cliente.email,
    NIF: cliente.nif,
    Direccion: cliente.direccion,
    CP: cliente.cp,
    Ciudad: cliente.ciudad
  });
}

/** =========================
 * UI: obtener fila seleccionada (PRESUPUESTOS o HISTORIAL)
 * ========================= */
function presGetSelectedRow_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (sh.getName() !== sheetName) throw new Error('Ve a la hoja ' + sheetName + ' y selecciona una celda de su fila.');
  const row = ss.getActiveRange().getRow();
  if (row < 2) throw new Error('Selecciona una fila (desde la 2 en adelante).');
  return { ss, sh, row };
}

function presGetRowData_(sh, row) {
  // lee A..U (21)
  const vals = sh.getRange(row, 1, 1, 21).getValues()[0];
  const obj = {};
  PRES_HEADERS.forEach((h, i) => obj[h] = vals[i]);
  return obj;
}

function presFindPresRow_(shPres, presId) {
  const headerInfo = presGetHeaderMap_(shPres);
  const idHeader = presPickHeaderName_(headerInfo.headers, ['Pres_ID']);
  if (!idHeader) throw new Error('No existe columna Pres_ID en ' + shPres.getName());

  const idCol = headerInfo.map[idHeader] || 1;
  const last = shPres.getLastRow();
  if (last < 2) throw new Error('No hay presupuestos');

  const data = shPres.getRange(2, 1, last - 1, shPres.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idCol - 1]).trim() === String(presId).trim()) {
      const rowObj = {};
      headerInfo.headers.forEach((h, idx) => { rowObj[h] = data[i][idx]; });
      return { rowNumber: i + 2, rowObj, headerInfo };
    }
  }
  throw new Error('No existe Pres_ID: ' + presId);
}

function presReadLineasPorPresId_(shLin, presId) {
  const { headers, map } = presGetHeaderMap_(shLin);
  const idHeader = presPickHeaderName_(headers, ['Pres_ID']);
  if (!idHeader) throw new Error('No existe columna Pres_ID en ' + shLin.getName());

  const idCol = map[idHeader] || 1;
  const last = shLin.getLastRow();
  if (last < 2) return { headers, lineas: [] };

  const data = shLin.getRange(2, 1, last - 1, shLin.getLastColumn()).getValues();
  const lineas = [];
  data.forEach((row) => {
    if (String(row[idCol - 1]).trim() === String(presId).trim()) {
      const obj = {};
      headers.forEach((h, idx) => { obj[h] = row[idx]; });
      lineas.push(obj);
    }
  });

  return { headers, lineas };
}

function presInsertLineasTable_(doc, lineas) {
  const tokens = ['LINEA_CONCEPTO', 'LINEA_CANTIDAD', 'LINEA_PRECIO', 'LINEA_SUBTOTAL'];
  const body = doc.getBody();
  if (!body) throw new Error('Documento sin body.');

  const tables = body.getTables();
  let table = null;
  let rowIndex = -1;

  for (let t = 0; t < tables.length && !table; t++) {
    const rows = tables[t].getNumRows();
    for (let r = 0; r < rows; r++) {
      const row = tables[t].getRow(r);
      const text = row.getText();
      if (tokens.some((k) => new RegExp(presBuildTokenPattern_(k), 'i').test(text))) {
        table = tables[t];
        rowIndex = r;
        break;
      }
    }
  }

  if (!table || rowIndex < 0) {
    throw new Error('No se encontro la fila plantilla con placeholders de lineas.');
  }

  const templateRow = table.getRow(rowIndex).copy();
  table.removeRow(rowIndex);

  (lineas || []).forEach((linea, idx) => {
    const row = templateRow.copy();
    const replacements = {
      LINEA_CONCEPTO: linea.concepto || '',
      LINEA_CANTIDAD: presFormatCantidad_(linea.cantidad),
      LINEA_PRECIO: presMoney2_(linea.precio),
      LINEA_SUBTOTAL: presMoney2_(linea.subtotal)
    };

    presReplaceTokensInRow_(row, replacements);
    table.insertTableRow(rowIndex + idx, row);
  });
}

function presReplaceTokensInRow_(row, map) {
  const cells = row.getNumCells();
  for (let c = 0; c < cells; c++) {
    const cell = row.getCell(c);
    const text = cell.editAsText();
    Object.keys(map).forEach((key) => {
      const pattern = presBuildTokenPattern_(key);
      text.replaceText(pattern, String(map[key] || ''));
    });
  }
}

function presMoney2_(n) {
  const num = Number(String(n).replace(',', '.'));
  if (isNaN(num)) return '';
  return Utilities.formatString('%.2f', num).replace('.', ',');
}

function presFormatCantidad_(value) {
  if (value === null || value === undefined || value === '') return '';
  const raw = String(value).replace(',', '.');
  const num = Number(raw);
  if (isNaN(num)) return String(value);
  if (Math.floor(num) === num) return String(num);
  return Utilities.formatString('%.2f', num).replace('.', ',');
}

/** =========================
 * UI: generar PDF + archivar
 * ========================= */
function uiGenerarPdfPresupuesto() {
  presAsegurarEstructura_();

  const { sh, row } = presGetSelectedRow_(SH_PRES);
  const presId = String(sh.getRange(row, 1).getDisplayValue()).trim();
  if (!presId) throw new Error('Esta fila no tiene Pres_ID en la columna A.');

  const url = generarPDFPresupuesto(presId, { archivar: true });
  SpreadsheetApp.getUi().alert('✅ PDF generado y archivado:\n' + url);
}

/** =========================
 * NOMBRE SEGURO
 * ========================= */
function sanitizeFileName_(name) {
  return String(name || '')
    .replace(/[\\\/:*?"<>|#]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120) || 'Sin_cliente';
}

/** =========================
 * EMISOR desde hoja FACTURA (mismo origen de tu sistema de facturas)
 * ========================= */
function getEmisorDesdeFactura_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('FACTURA');
  if (!sh) return { nombre:'', nif:'', direccion:'', cp:'', ciudad:'' };

  return {
    nombre: String(sh.getRange('B15').getValue() || ''),
    nif: String(sh.getRange('B16').getValue() || ''),
    direccion: String(sh.getRange('B17').getValue() || ''),
    cp: String(sh.getRange('B18').getDisplayValue() || ''),
    ciudad: String(sh.getRange('B19').getValue() || '')
  };
}

/** =========================
 * REPLACE TOKENS (body + header + footer si existen)
 * ========================= */
function replaceTokensEverywhere_(doc, map) {
  const containers = [];
  const body = doc.getBody();
  if (body) containers.push(body);

  // header/footer pueden ser null (y pueden lanzar en algunos docs)
  try { const h = doc.getHeader(); if (h) containers.push(h); } catch(e){}
  try { const f = doc.getFooter(); if (f) containers.push(f); } catch(e){}

  const keys = Object.keys(map || {});
  containers.forEach(c => {
    keys.forEach(key => {
      const value = (map[key] === null || map[key] === undefined) ? '' : String(map[key]);
      // Soporta {{TOKEN}} con espacios: {{ TOKEN }}
      const pattern = presBuildTokenPattern_(key);
      c.replaceText(pattern, value);
    });
  });
}

function presBuildTokenPattern_(key) {
  return `\\{\\{\\s*${presCaseInsensitiveKeyPattern_(key)}\\s*\\}\\}`;
}

function presCaseInsensitiveKeyPattern_(key) {
  const raw = String(key || '');
  let out = '';
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i];
    if (/[a-zA-Z]/.test(ch)) {
      const lower = ch.toLowerCase();
      const upper = ch.toUpperCase();
      out += lower === upper ? escapeRegex_(ch) : `[${lower}${upper}]`;
    } else {
      out += escapeRegex_(ch);
    }
  }
  return out;
}

function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** =========================
 * ARCHIVAR LÍNEAS: PRES_LINEAS -> LINEAS_PRES_HIST + limpiar
 * ========================= */
function archivarYLimpiarLineasPres_(presId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shLin = ss.getSheetByName(SH_PRES_LINEAS);
  const hist = ss.getSheetByName(SH_PRES_LINEAS_HIS);
  if (!shLin) throw new Error('No existe hoja ' + SH_PRES_LINEAS);
  if (!hist) throw new Error('No existe hoja ' + SH_PRES_LINEAS_HIS);

  const lastRow = shLin.getLastRow();
  if (lastRow < 2) return;

  const data = shLin.getRange(2, 1, lastRow - 1, 8).getValues(); // A..H
  const rowsToArchive = [];
  const rowsToClear = [];

  data.forEach((r, idx) => {
    if (String(r[0]).trim() === String(presId).trim()) {
      const concepto = String(r[2] || '').trim();
      const cant     = String(r[3] || '').trim();
      const precio   = String(r[4] || '').trim();
      const tieneAlgo = concepto || cant || precio;

      if (tieneAlgo) {
        rowsToArchive.push([r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], new Date()]);
      }
      rowsToClear.push(idx + 2);
    }
  });

  if (rowsToArchive.length) {
    const start = hist.getLastRow() + 1;
    hist.getRange(start, 1, rowsToArchive.length, 9).setValues(rowsToArchive);
  }

  // limpia A..H en PRES_LINEAS para ese Pres_ID
  rowsToClear.forEach(r => shLin.getRange(r, 1, 1, 8).clearContent());
}

/** =========================
 * ARCHIVAR PRESUPUESTO: PRESUPUESTOS -> HISTORIAL_PRESUPUESTOS
 * (no borra tu fila si no quieres; la marcamos como Archivado y set Archivado_el)
 * ========================= */
function archivarPresupuestoEnHistorial_(presId, pdfUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_PRES);
  const hist = ss.getSheetByName(SH_PRES_HIST);
  if (!sh || !hist) throw new Error('Faltan hojas para archivar');

  const last = sh.getLastRow();
  if (last < 2) throw new Error('No hay presupuestos');

  const data = sh.getRange(2, 1, last - 1, 21).getValues(); // A..U
  const idx = data.findIndex(r => String(r[0]).trim() === String(presId).trim());
  if (idx === -1) throw new Error('No existe Pres_ID en PRESUPUESTOS: ' + presId);

  const row = idx + 2;
  const vals = data[idx];

  // Actualiza PDF_link y Archivado_el en el objeto a archivar
  if (pdfUrl) vals[16] = pdfUrl;    // Q PDF_link
  vals[4] = 'Archivado';            // E Estado
  vals[20] = new Date();            // U Archivado_el

  // Si ya existe en historial, actualiza. Si no, agrega.
  const lastH = hist.getLastRow();
  let histRow = -1;
  if (lastH >= 2) {
    const ids = hist.getRange(2, 1, lastH - 1, 1).getValues().flat();
    const j = ids.findIndex(v => String(v).trim() === String(presId).trim());
    if (j !== -1) histRow = j + 2;
  }

  if (histRow === -1) {
    hist.appendRow(vals);
  } else {
    hist.getRange(histRow, 1, 1, 21).setValues([vals]);
  }

  // ✅ BORRAR la fila de PRESUPUESTOS (para “limpiar” y que quede solo historial)
  sh.deleteRow(row);
}

/** =========================
 * GENERAR PDF (principal)
 * options: { archivar: true|false }
 * ========================= */
function generarPDFPresupuesto(presId, options) {
  presAsegurarEstructura_();
  const opts = options || {};
  const archivar = opts.archivar !== false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPres = ss.getSheetByName(SH_PRES);
  const shLin  = ss.getSheetByName(SH_PRES_LINEAS);
  if (!shPres || !shLin) throw new Error('Faltan hojas ' + SH_PRES + ' o ' + SH_PRES_LINEAS);

  const log = (resultado, mensaje, data) => {
    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PRESUPUESTOS', 'PDF', 'PRESUPUESTO', presId, resultado, mensaje || '', data || null);
    }
  };

  try {
    const presInfo = presFindPresRow_(shPres, presId);
    const rowObj = presInfo.rowObj;

    const fechaVal = presPickValue_(rowObj, ['Fecha']);
    const venceVal = presPickValue_(rowObj, ['Vence_el', 'Vence']);
    const cliente = String(presPickValue_(rowObj, ['Cliente', 'Lead_Nombre']) || 'Sin_cliente').trim() || 'Sin_cliente';
    const email = String(presPickValue_(rowObj, ['Email_cliente', 'Lead_Email', 'Email']) || '').trim();
    const nif = String(presPickValue_(rowObj, ['NIF', 'DNI', 'CIF']) || '').trim();
    const dir = String(presPickValue_(rowObj, ['Direccion', 'Dirección']) || '').trim();
    const cp = String(presPickValue_(rowObj, ['CP']) || '').trim();
    const ciudad = String(presPickValue_(rowObj, ['Ciudad']) || '').trim();
    const notas = String(presPickValue_(rowObj, ['Notas']) || '').trim();

    const lineasData = presReadLineasPorPresId_(shLin, presId);
    const conceptHeader = presPickHeaderName_(lineasData.headers, ['Concepto','Descripcion','Descripción']) || 'Concepto';
    const cantidadHeader = presPickHeaderName_(lineasData.headers, ['Cantidad','Cant']) || 'Cantidad';
    const precioHeader = presPickHeaderName_(lineasData.headers, ['Precio','Precio_unitario']) || 'Precio';
    const dtoHeader = presPickHeaderName_(lineasData.headers, ['Dto_%','Dto']) || 'Dto_%';
    const ivaHeader = presPickHeaderName_(lineasData.headers, ['IVA_%','IVA']) || 'IVA_%';
    const subtotalHeader = presPickHeaderName_(lineasData.headers, ['Subtotal','Importe','Base']) || 'Subtotal';

    const lineas = (lineasData.lineas || []).map((l) => {
      const concepto = String(presPickValue_(l, [conceptHeader]) || '').trim();
      const cantidad = Number(presPickValue_(l, [cantidadHeader]) || 0);
      const precio = Number(presPickValue_(l, [precioHeader]) || 0);
      const dto = Number(presPickValue_(l, [dtoHeader]) || 0);
      const iva = Number(presPickValue_(l, [ivaHeader]) || 21);
      let subtotal = Number(presPickValue_(l, [subtotalHeader]) || 0);
      if (!subtotal && cantidad && precio) subtotal = Number((cantidad * precio * (1 - (dto/100))).toFixed(2));
      const enriched = Object.assign({}, l);
      enriched[subtotalHeader] = subtotal;
      return { concepto, cantidad, precio, dto, iva, subtotal, enriched };
    }).filter((l) => l.concepto && (l.cantidad || l.precio || l.subtotal));

    if (!lineas.length) throw new Error('No hay lineas validas en ' + SH_PRES_LINEAS + ' para ' + presId);

    let base = 0, ivaTotal = 0;
    lineas.forEach((l) => {
      const sub = l.subtotal || (l.cantidad * l.precio * (1 - (l.dto/100)));
      base += sub;
      ivaTotal += sub * (l.iva/100);
    });
    base = Number(base.toFixed(2));
    ivaTotal = Number(ivaTotal.toFixed(2));
    const total = Number((base + ivaTotal).toFixed(2));

    presSetValuesByHeaders_(shPres, presInfo.rowNumber, presInfo.headerInfo.map, {
      Base: base,
      IVA_total: ivaTotal,
      Total: total
    });

    const cfg = getConfigPres_();
    if (!cfg.carpetaId) throw new Error('Config incompleta: falta carpeta destino para PDFs');
    if (!cfg.templateId) throw new Error('Config incompleta: falta template de Docs para PDFs');
    const folder = DriveApp.getFolderById(cfg.carpetaId);

    const safeCliente = sanitizeFileName_(cliente);
    const baseName = sanitizeFileName_(`Presupuesto_${presId}_${safeCliente}`);

    const copy = DriveApp.getFileById(cfg.templateId).makeCopy(baseName, folder);
    const doc = DocumentApp.openById(copy.getId());

    const em = getEmisorDesdeFactura_();
    const tz = Session.getScriptTimeZone();
    const fechaDate = presToDate_(fechaVal);
    const venceDate = presToDate_(venceVal);
    const FECHA_TXT = fechaDate ? Utilities.formatDate(fechaDate, tz, 'dd/MM/yyyy') : '';
    const VENCE_TXT = venceDate ? Utilities.formatDate(venceDate, tz, 'dd/MM/yyyy') : '';

    const uniqueIvas = [...new Set(lineas.map(l => l.iva))];
    const ivaPorc = (uniqueIvas.length === 1) ? String(uniqueIvas[0]) : String(uniqueIvas[0] || 21);

    const map = {
      'PRES_ID': presId,
      'Pres_ID': presId,
      'FECHA': FECHA_TXT,
      'Fecha': FECHA_TXT,
      'VENCE': VENCE_TXT,
      'Vence_el': VENCE_TXT,
      'Emisor_nombre': em.nombre || '',
      'Emisor_NIF': em.nif || '',
      'Emisor_direccion': em.direccion || '',
      'Emisor_CP': em.cp || '',
      'Emisor_ciudad': em.ciudad || '',
      'CLIENTE': cliente,
      'Cliente': cliente,
      'NIF': nif,
      'DIRECCION': dir,
      'Direccion': dir,
      'CP': cp,
      'CIUDAD': ciudad,
      'Ciudad': ciudad,
      'EMAIL': email,
      'Email_cliente': email,
      'BASE': base.toFixed(2),
      'Base': base.toFixed(2),
      'IVA_PORC': ivaPorc,
      'IVA_TOTAL': ivaTotal.toFixed(2),
      'IVA_total': ivaTotal.toFixed(2),
      'TOTAL': total.toFixed(2),
      'Total': total.toFixed(2),
      'NOTAS': notas,
      'Notas': notas
    };

    replaceTokensEverywhere_(doc, map);
    presInsertLineasTable_(doc, lineas);
    doc.saveAndClose();

    const pdfBlob = DriveApp.getFileById(copy.getId()).getAs(MimeType.PDF).setName(baseName + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);
    const url = pdfFile.getUrl();

    const updates = { PDF_link: url };
    if (presInfo.headerInfo.map['Fecha_envio']) {
      const current = shPres.getRange(presInfo.rowNumber, presInfo.headerInfo.map['Fecha_envio']).getValue();
      if (!current) updates.Fecha_envio = new Date();
    }
    presSetValuesByHeaders_(shPres, presInfo.rowNumber, presInfo.headerInfo.map, updates);

    if (archivar) {
      archivarYLimpiarLineasPres_(presId);
      archivarPresupuestoEnHistorial_(presId, url);
    }

    log('OK', '', { url, archivar });
    return url;
  } catch (err) {
    log('ERROR', err.message || String(err), { stack: err.stack });
    throw err;
  }
}

/** =========================
 * UI: estados rápidos
 * ========================= */
function presGetSelectedRowAny_(allowedSheets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const name = sh.getName();

  if (!allowedSheets.includes(name)) {
    throw new Error('Ve a la hoja ' + allowedSheets.join(' o ') + ' y selecciona una celda de su fila.');
  }

  const row = ss.getActiveRange().getRow();
  if (row < 2) throw new Error('Selecciona una fila (desde la 2 en adelante).');

  return { ss, sh, row, name };
}

function uiMarcarPresupuestoEnviado() {
  presAsegurarEstructura_();

  const { sh, row } = presGetSelectedRowAny_([SH_PRES, SH_PRES_HIST]);
  const headerInfo = presGetHeaderMap_(sh);
  const colEstado = headerInfo.map['Estado'];
  const colFechaEnvio = headerInfo.map['Fecha_envio'];

  if (colEstado) sh.getRange(row, colEstado).setValue('ENVIADO');
  if (colFechaEnvio) sh.getRange(row, colFechaEnvio).setValue(new Date());
  SpreadsheetApp.getUi().alert('✅ Marcado como ENVIADO.');
}

function uiMarcarPresupuestoAceptado() {
  presAsegurarEstructura_();

  const { sh, row } = presGetSelectedRowAny_([SH_PRES, SH_PRES_HIST]);
  const headerInfo = presGetHeaderMap_(sh);
  const colEstado = headerInfo.map['Estado'];
  const colFechaAcept = headerInfo.map['Fecha_aceptacion'];

  if (colEstado) sh.getRange(row, colEstado).setValue('ACEPTADO');
  if (colFechaAcept) sh.getRange(row, colFechaAcept).setValue(new Date());
  SpreadsheetApp.getUi().alert('✅ Marcado como ACEPTADO.');
}

function uiMarcarPresupuestoRechazado() {
  return uiMarcarPresupuestoPerdido();
}

function uiMarcarPresupuestoPerdido() {
  presAsegurarEstructura_();

  const { sh, row } = presGetSelectedRowAny_([SH_PRES, SH_PRES_HIST]);
  const headerInfo = presGetHeaderMap_(sh);
  const colEstado = headerInfo.map['Estado'];
  if (colEstado) sh.getRange(row, colEstado).setValue('PERDIDO');
  SpreadsheetApp.getUi().alert('✅ Marcado como PERDIDO.');
}

/** =========================
 * EMAIL (desde fila seleccionada en PRESUPUESTOS o HISTORIAL)
 * ========================= */
function uiEnviarPresupuestoEmail() {
  presAsegurarEstructura_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const name = sh.getName();
  if (name !== SH_PRES && name !== SH_PRES_HIST) {
    SpreadsheetApp.getUi().alert('Ve a PRESUPUESTOS o HISTORIAL_PRESUPUESTOS y selecciona una fila.');
    return;
  }
  const row = ss.getActiveRange().getRow();
  if (row < 2) return;

  const o = presGetRowData_(sh, row);
  const email = String(o['Email_cliente'] || '').trim();
  const pdf = String(o['PDF_link'] || '').trim();
  const presId = String(o['Pres_ID'] || '').trim();
  const cliente = String(o['Cliente'] || '').trim() || 'cliente';

  if (!email) {
    SpreadsheetApp.getUi().alert('Esta fila no tiene Email_cliente.');
    return;
  }
  if (!pdf) {
    SpreadsheetApp.getUi().alert('Esta fila no tiene PDF_link. Genera el PDF primero.');
    return;
  }

  const subject = `Presupuesto ${presId} - Costa Clean`;
  const body =
`Hola ${cliente},

Te adjunto el enlace del presupuesto ${presId}:
${pdf}

Si te va bien, respóndeme a este email y lo dejamos confirmado.

Un saludo,
Costa Clean`;

  MailApp.sendEmail({
    to: email,
    subject,
    body
  });

  // registrar fecha_envio (S=19) si existe
  sh.getRange(row, 19).setValue(new Date());

  SpreadsheetApp.getUi().alert('✅ Email enviado a ' + email);
}

/** =========================
 * WHATSAPP (modal + enlace wa.me)
 * ========================= */
function uiWhatsAppPresupuesto() {
  presAsegurarEstructura_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const name = sh.getName();
  if (name !== SH_PRES && name !== SH_PRES_HIST) {
    SpreadsheetApp.getUi().alert('Ve a PRESUPUESTOS o HISTORIAL_PRESUPUESTOS y selecciona una fila.');
    return;
  }
  const row = ss.getActiveRange().getRow();
  if (row < 2) return;

  const o = presGetRowData_(sh, row);
  const presId = String(o['Pres_ID'] || '').trim();
  const cliente = String(o['Cliente'] || '').trim() || 'cliente';
  const pdf = String(o['PDF_link'] || '').trim();

  if (!pdf) {
    SpreadsheetApp.getUi().alert('Esta fila no tiene PDF_link. Genera el PDF primero.');
    return;
  }

  const msg = `Hola ${cliente} 👋\nTe paso el presupuesto ${presId}:\n${pdf}\n\nCualquier cosa me dices y lo dejamos confirmado.`;
  const waLink = 'https://wa.me/?text=' + encodeURIComponent(msg);

  const html = HtmlService.createHtmlOutput(
`<div style="font-family:Arial;padding:12px">
  <h3>WhatsApp listo</h3>
  <p>Copia si quieres, o abre WhatsApp con el mensaje ya preparado:</p>
  <textarea style="width:100%;height:120px">${escapeHtml_(msg)}</textarea>
  <p style="margin-top:10px">
    <a href="${waLink}" target="_blank" style="display:inline-block;padding:10px 12px;background:#25D366;color:white;text-decoration:none;border-radius:8px">
      Abrir WhatsApp
    </a>
  </p>
</div>`
  ).setWidth(420).setHeight(320);

  SpreadsheetApp.getUi().showModalDialog(html, 'Enviar por WhatsApp');

  // registrar fecha_envio (S=19) si existe
  sh.getRange(row, 19).setValue(new Date());
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/** =========================
 * CONVERTIR A FACTURA (desde HISTORIAL_PRESUPUESTOS)
 * SOLUCIONA tu error: ya NO busca en PRES_LINEAS,
 * usa LINEAS_PRES_HIST.
 * ========================= */
function uiConvertirPresupuestoAFactura() {
  presAsegurarEstructura_();

  const { ss, sh, row } = presGetSelectedRow_(SH_PRES_HIST);
  const o = presGetRowData_(sh, row);

  const estado = String(o['Estado'] || '').trim().toLowerCase();
  if (estado !== 'aceptado') {
    SpreadsheetApp.getUi().alert('Este presupuesto no está en estado "ACEPTADO".\nCámbialo a ACEPTADO en el historial y vuelve a intentar.');
    return;
  }

  const presId = String(o['Pres_ID'] || '').trim();
  if (!presId) throw new Error('Fila sin Pres_ID.');

  const facturaIdExistente = String(o['Factura_ID'] || '').trim();
  if (facturaIdExistente) {
    SpreadsheetApp.getUi().alert('Este presupuesto ya fue convertido a factura: ' + facturaIdExistente);
    return;
  }

  const facturaId = convertirPresupuestoAFactura_(o);
  sh.getRange(row, 18).setValue(facturaId); // R Factura_ID

  SpreadsheetApp.getUi().alert('✅ Convertido a factura: ' + facturaId);
}

function convertirPresupuestoAFactura_(presObj) {
  // Requiere tus hojas FACTURA y LINEAS + tu workflowGenerarFactura funcionando
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shFactura = ss.getSheetByName('FACTURA');
  const shLineas  = ss.getSheetByName('LINEAS');
  const shLinHist = ss.getSheetByName(SH_PRES_LINEAS_HIS);

  if (!shFactura || !shLineas) throw new Error('Falta hoja FACTURA o LINEAS (sistema facturas).');
  if (!shLinHist) throw new Error('Falta hoja ' + SH_PRES_LINEAS_HIS);

  const presId   = String(presObj['Pres_ID'] || '').trim();
  const clienteId= String(presObj['Cliente_ID'] || '').trim();
  const cliente  = String(presObj['Cliente'] || '').trim();
  const nif      = String(presObj['NIF'] || '').trim();
  const dir      = String(presObj['Direccion'] || '').trim();
  const cp       = String(presObj['CP'] || '').trim();
  const ciudad   = String(presObj['Ciudad'] || '').trim();

  // 1) Leer líneas desde LINEAS_PRES_HIST (NO PRES_LINEAS)
  const last = shLinHist.getLastRow();
  if (last < 2) throw new Error('No hay líneas en ' + SH_PRES_LINEAS_HIS);

  const data = shLinHist.getRange(2, 1, last - 1, 9).getValues(); // A..I
  const lineas = data
    .filter(r => String(r[0]).trim() === presId)
    .map(r => ({
      concepto: String(r[2] || '').trim(),
      cantidad: Number(r[3]) || 0,
      precio:   Number(r[4]) || 0,
      iva:      Number(r[6]) || 0
    }))
    .filter(l => l.concepto && l.cantidad > 0 && l.precio > 0);

  if (!lineas.length) throw new Error('No hay líneas para este presupuesto en ' + SH_PRES_LINEAS_HIS);

  // 2) Obtener número de factura (usa tu función del CRM_NUEVO)
  if (typeof consumirSiguienteNumero_ !== 'function') {
    throw new Error('No existe consumirSiguienteNumero_() en el proyecto (CRM_NUEVO).');
  }
  const numeroFactura = consumirSiguienteNumero_();

  // 3) Rellenar hoja FACTURA (campos que tu workflow usa)
  // B1 = Cliente_ID (lo usas en HISTORIAL facturas)
  shFactura.getRange('B1').setValue(clienteId || '');
  shFactura.getRange('B2').setValue(cliente || '');
  shFactura.getRange('B3').setValue(nif || '');
  shFactura.getRange('B4').setValue(dir || '');
  shFactura.getRange('B5').setValue(cp || '');
  shFactura.getRange('B6').setValue(ciudad || '');

  shFactura.getRange('B8').setValue(numeroFactura);
  shFactura.getRange('B9').setValue(new Date());

  // IVA%: si varias tasas, tomamos la primera
  const uniqueIvas = [...new Set(lineas.map(l => l.iva))];
  const ivaPorc = Number(uniqueIvas[0] || 21);
  shFactura.getRange('B11').setValue(ivaPorc);

  // 4) Crear líneas en hoja LINEAS para ese número (A=Numero, B Concepto, C Cantidad, D Precio)
  const nextRow = findNextFreeRow_(shLineas, 2, 4); // mira A:D
  lineas.forEach((l, i) => {
    const r = nextRow + i;
    shLineas.getRange(r, 1).setValue(numeroFactura);
    shLineas.getRange(r, 2).setValue(l.concepto);
    shLineas.getRange(r, 3).setValue(l.cantidad);
    shLineas.getRange(r, 4).setValue(l.precio);
  });

  // 5) Generar PDF factura con tu workflow existente
  if (typeof workflowGenerarFactura !== 'function') {
    throw new Error('No existe workflowGenerarFactura() en el proyecto.');
  }
  workflowGenerarFactura();

  return numeroFactura;
}

function findNextFreeRow_(sh, startRow, cols) {
  const lastRow = Math.max(sh.getLastRow(), startRow);
  let nextRow = startRow;

  if (lastRow > startRow) {
    const data = sh.getRange(startRow, 1, lastRow - startRow + 1, cols).getValues(); // A:D
    for (let i = data.length - 1; i >= 0; i--) {
      const hasData = data[i].some(v => String(v).trim() !== '');
      if (hasData) { nextRow = startRow + i + 1; break; }
    }
  }
  return nextRow;
}

/** =========================
 * ===== AI: Página + API =====
 * ========================= */

function aiCrearPagina_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let shCfg = ss.getSheetByName(SH_AI_CONFIG);
  if (!shCfg) shCfg = ss.insertSheet(SH_AI_CONFIG);

  let shLog = ss.getSheetByName(SH_AI_LOG);
  if (!shLog) shLog = ss.insertSheet(SH_AI_LOG);

  shCfg.clear();
  shCfg.getRange(1,1,1,2).setValues([['Clave','Valor']]);
  shCfg.getRange(2,1,6,2).setValues([
    ['MODELO', 'gpt-5'],
    ['TEMPERATURA', '0.7'],
    ['MAX_OUTPUT_TOKENS', '350'],
    ['IDIOMA', 'es'],
    ['INSTRUCCION_BASE', 'Eres un asistente comercial de Costa Clean. Genera mensajes claros, profesionales y persuasivos sin ser agresivo.'],
    ['NOTA', 'La API Key se guarda en Propiedades del Script (no se muestra aquí). Usa el menú para configurarla.']
  ]);
  shCfg.setFrozenRows(1);

  shLog.clear();
  shLog.getRange(1,1,1,6).setValues([['Fecha','Accion','Pres_ID','Resultado','Error','Payload_resumen']]);
  shLog.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('✅ Página AI creada: AI_CONFIG y AI_LOG');
}

function aiConfigurarApiKey_() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Configurar OpenAI API Key', 'Pega tu OPENAI_API_KEY (se guarda de forma segura en el Script)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const key = String(r.getResponseText() || '').trim();
  if (!key) {
    ui.alert('API Key vacía.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty(PROP_OPENAI_KEY, key);
  ui.alert('✅ API Key guardada.');
}

/**
 * Genera texto comercial (email + whatsapp) con AI para la fila seleccionada
 * en PRESUPUESTOS o HISTORIAL_PRESUPUESTOS.
 */
function aiGenerarTextoComercial_() {
  presAsegurarEstructura_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const name = sh.getName();
  if (name !== SH_PRES && name !== SH_PRES_HIST) {
    SpreadsheetApp.getUi().alert('Ve a PRESUPUESTOS o HISTORIAL_PRESUPUESTOS y selecciona una fila.');
    return;
  }
  const row = ss.getActiveRange().getRow();
  if (row < 2) return;

  const o = presGetRowData_(sh, row);
  const presId = String(o['Pres_ID'] || '').trim();
  const cliente = String(o['Cliente'] || '').trim();
  const total = String(o['Total'] || '').trim();
  const pdf = String(o['PDF_link'] || '').trim();

  const prompt =
`Genera 2 textos:
1) Email corto (asunto + cuerpo) para enviar el presupuesto.
2) WhatsApp corto y directo.
Contexto:
- Empresa: Costa Clean (limpieza)
- Cliente: ${cliente}
- Presupuesto: ${presId}
- Total: ${total}
- Enlace PDF: ${pdf}
Estilo: castellano Barcelona, profesional, persuasivo sin agresividad, claro, con llamada a la acción suave.`;

  const out = openaiResponsesText_(prompt);

  // Modal con el resultado
  const html = HtmlService.createHtmlOutput(
`<div style="font-family:Arial;padding:12px">
  <h3>Texto AI generado</h3>
  <textarea style="width:100%;height:260px;white-space:pre-wrap">${escapeHtml_(out)}</textarea>
</div>`
  ).setWidth(520).setHeight(380);

  SpreadsheetApp.getUi().showModalDialog(html, 'AI - Costa Clean');

  aiLog_('GENERAR_TEXTO', presId, 'OK', '', out.slice(0, 300));
}

/**
 * Llamada a OpenAI Responses API (/v1/responses)
 * Docs: https://api.openai.com/v1/responses  (según OpenAI) :contentReference[oaicite:1]{index=1}
 */
function openaiResponsesText_(inputText) {
  const key = PropertiesService.getScriptProperties().getProperty(PROP_OPENAI_KEY);
  if (!key) throw new Error('No hay API Key. Menú Presupuestos → 🧠 AI: Configurar API Key');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCfg = ss.getSheetByName(SH_AI_CONFIG);

  const model = (shCfg ? String(shCfg.getRange('B2').getValue() || 'gpt-4.1-mini') : 'gpt-4.1-mini');
  const temp  = (shCfg ? Number(shCfg.getRange('B3').getValue() || 0.7) : 0.7);
  const maxOut= (shCfg ? Number(shCfg.getRange('B4').getValue() || 350) : 350);
  const idioma= (shCfg ? String(shCfg.getRange('B5').getValue() || 'es') : 'es');
  const baseI = (shCfg ? String(shCfg.getRange('B6').getValue() || '') : '');

  // Construye payload base
  const basePayload = {
    model: model,
    input: [
      {
        role: "system",
        content: [{ type: "input_text", text: `${baseI}\nIdioma: ${idioma}` }]
      },
      {
        role: "user",
        content: [{ type: "input_text", text: String(inputText || '') }]
      }
    ],
    max_output_tokens: maxOut
  };

  // 1) intento con temperature
  let payload = { ...basePayload, temperature: temp };

  const call = (p) => UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify(p),
    muteHttpExceptions: true
  });

  let res = call(payload);
  let code = res.getResponseCode();
  let raw  = res.getContentText();

  // Si el modelo no soporta temperature → reintenta sin temperature
  if (code === 400 && raw && raw.includes('Unsupported parameter') && raw.includes('"temperature"')) {
    payload = { ...basePayload }; // sin temperature
    res = call(payload);
    code = res.getResponseCode();
    raw  = res.getContentText();
  }

  if (code < 200 || code >= 300) {
    aiLog_('OPENAI_CALL', '', 'ERROR', raw.slice(0, 2000), JSON.stringify(payload).slice(0, 300));
    throw new Error('OpenAI API error (' + code + '): ' + raw.slice(0, 800));
  }

  const json = JSON.parse(raw);

// Extraer texto correctamente del Responses API
let text = '';

if (json.output_text) {
  text = String(json.output_text);
} else if (Array.isArray(json.output)) {
  for (const block of json.output) {
    if (Array.isArray(block.content)) {
      for (const part of block.content) {
        if (part.type === 'output_text' && part.text) {
          text += part.text;
        }
      }
    }
  }
}

text = String(text || '').trim();
return text || '(Sin texto devuelto por el modelo)';

}


function aiLog_(accion, presId, resultado, error, payloadResumen) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SH_AI_LOG);
  if (!sh) {
    sh = ss.insertSheet(SH_AI_LOG);
    sh.getRange(1,1,1,6).setValues([['Fecha','Accion','Pres_ID','Resultado','Error','Payload_resumen']]);
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), accion, presId || '', resultado || '', error || '', payloadResumen || '']);
}

/** =========================
 * Trigger router (tu TRIGGERS.gs lo llama, pero por si acaso)
 * ========================= */
function onEdit(e) {
  try { if (typeof onEditPresupuestos_ === 'function') onEditPresupuestos_(e); } catch(err) {}
}

/***************
 * FORM -> PRESUPUESTO AUTO
 ***************/

// 1) Pon aquí el nombre exacto de tu hoja de respuestas del Form
const SH_FORM_RESPUESTAS = 'Respuestas de formulario';



// 2) Mapea aquí las columnas del Form (por nombre de pregunta)
//    OJO: deben coincidir con los títulos de columna de la hoja de respuestas
const FORM_MAP = {
  cliente: 'Nombre / Empresa',
  email: 'Email',
  nif: 'NIF',
  direccion: 'Dirección',
  cp: 'CP',
  ciudad: 'Ciudad',
  tipoServicio: 'Tipo de servicio',
  habitaciones: 'Habitaciones',
  banos: 'Baños',
  metros: 'm²',
  cristales: 'Incluye cristales',
  notas: 'Notas'
};

/**
 * Trigger instalado: From spreadsheet -> On form submit
 */
function onFormSubmitPresupuesto(e)
 {
  try {
    presAsegurarEstructura_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shResp = ss.getSheetByName(SH_FORM_RESPUESTAS);
    if (!shResp) throw new Error('No existe la hoja de respuestas: ' + SH_FORM_RESPUESTAS);

    // Convierte la respuesta a objeto por header
    const obj = formEventToObject_(e);

    // Lee campos (con fallback)
    const cliente = (obj[FORM_MAP.cliente] || '').toString().trim();
    const email   = (obj[FORM_MAP.email] || '').toString().trim();
    const nif     = (obj[FORM_MAP.nif] || '').toString().trim();
    const dir     = (obj[FORM_MAP.direccion] || '').toString().trim();
    const cp      = (obj[FORM_MAP.cp] || '').toString().trim();
    const ciudad  = (obj[FORM_MAP.ciudad] || '').toString().trim();
    const notas   = (obj[FORM_MAP.notas] || '').toString().trim();

    // Validación mínima
    if (!cliente) throw new Error('El formulario llegó sin Cliente/Empresa.');

    // Crear presupuesto
    const shPres = ss.getSheetByName(SH_PRES);
    const presId = consumirSiguientePresId_();
    const cfg = getConfigPres_();
    const hoy = new Date();
    const vence = new Date(hoy.getTime() + (cfg.validezDefault || 15) * 86400000);

    // Fila PRESUPUESTOS (A..U = 21 columnas)
    const row = [
      presId, hoy, cfg.validezDefault || 15, vence, 'BORRADOR',
      '', // Cliente_ID (si luego quieres enlazar a CLIENTES)
      cliente, email, nif, dir, cp, ciudad,
      '', '', '', notas, '', '', '', '', '' // Base, IVA_total, Total, PDF_link, Factura_ID, Fechas..., Archivado_el
    ];

    shPres.appendRow(row);
    const presRow = shPres.getLastRow();

    // Reservar líneas (base)
    reservarLineasPres_(presId, PRES_LINEAS_PRECREADAS);

    // Rellenar líneas automáticas según respuestas
    crearLineasDesdeForm_(presId, obj);

    // Opcional: marcar como ENVIADO automáticamente si quieres
    // shPres.getRange(presRow, 5).setValue('ENVIADO');
    // shPres.getRange(presRow, 19).setValue(new Date());

    // Opcional: generar PDF automático y archivar
    // const url = generarPDFPresupuesto(presId, { archivar: true });

  } catch (err) {
    // Si quieres log, lo mandas a AI_LOG o a Logger
    Logger.log(err);
  }
}

/**
 * Convierte el event e a objeto {Header: Value}
 * Funciona tanto con e.namedValues como con e.values.
 */
function formEventToObject_(e) {
  if (e && e.namedValues) {
    const out = {};
    Object.keys(e.namedValues).forEach(k => out[k] = e.namedValues[k][0]);
    return out;
  }

  // fallback si no viene namedValues
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_FORM_RESPUESTAS);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const vals = e.values || [];
  const obj = {};
  headers.forEach((h,i) => obj[h] = vals[i]);
  return obj;
}

/**
 * Crea líneas automáticas en PRES_LINEAS para ese Pres_ID
 * Ajusta esta lógica a tus servicios reales.
 */
function crearLineasDesdeForm_(presId, obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shLin = ss.getSheetByName(SH_PRES_LINEAS);
  if (!shLin) throw new Error('No existe ' + SH_PRES_LINEAS);

  const tipo = (obj[FORM_MAP.tipoServicio] || '').toString().toLowerCase();
  const hab  = Number(obj[FORM_MAP.habitaciones] || 0) || 0;
  const ban  = Number(obj[FORM_MAP.banos] || 0) || 0;
  const m2   = Number(obj[FORM_MAP.metros] || 0) || 0;
  const cris = (obj[FORM_MAP.cristales] || '').toString().toLowerCase();

  // Busca las filas ya reservadas para ese Pres_ID y rellena empezando por la primera libre
  const last = shLin.getLastRow();
  if (last < 2) throw new Error('No hay filas en PRES_LINEAS.');

  const data = shLin.getRange(2,1,last-1,8).getValues(); // A..H
  const idxs = [];
  data.forEach((r, i) => {
    if (String(r[0]).trim() === String(presId).trim()) idxs.push(i + 2);
  });
  if (!idxs.length) throw new Error('No encontré filas reservadas para ' + presId);

  // Construye lista de líneas según el formulario
  const lineas = [];

  // Ejemplo: limpieza base
  lineas.push({ concepto: `Limpieza ${tipo || 'general'} (${m2 ? m2 + ' m²' : 'servicio'})`, cantidad: 1, precio: '' });

  // Ejemplo: extras por habitaciones/baños (si tu negocio lo usa)
  if (hab) lineas.push({ concepto: `Habitaciones (${hab})`, cantidad: hab, precio: '' });
  if (ban) lineas.push({ concepto: `Baños (${ban})`, cantidad: ban, precio: '' });

  // Cristales
  if (cris.includes('sí') || cris.includes('si')) {
    lineas.push({ concepto: 'Limpieza de cristales', cantidad: 1, precio: '' });
  }

  // Rellena en filas reservadas
  for (let i = 0; i < Math.min(lineas.length, idxs.length); i++) {
    const r = idxs[i];
    shLin.getRange(r, 3).setValue(lineas[i].concepto); // Concepto
    shLin.getRange(r, 4).setValue(lineas[i].cantidad); // Cantidad
    // Precio lo dejas vacío para que tú lo pongas, o puedes calcularlo si tienes tarifa
    // shLin.getRange(r, 5).setValue(lineas[i].precio);
  }
}


