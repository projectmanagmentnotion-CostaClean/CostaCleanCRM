var __CC_SETUP = (function(){
var setupAll_ = function() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = (accion, resultado, mensaje, data) => {
    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'SETUP', accion, 'SYSTEM', '', resultado, mensaje || '', data || null);
    }
  };

  log('RUN', 'START', '', null);
  try {
    const headersConfig = [
      'Ano','Ultimo_numero','Leads_Sheet_ID','Leads_Tab_Name','Leads_Destino',
      'CLI_Ano','CLI_Ultimo_numero','PRES_A�o','PRES_Ultimo_numero','PRES_Validez_default',
      'PRES_Pdf_Folder_Id','PRES_Template_DocId','FACT_Pdf_Folder_Id','FACT_Template_DocId'
    ];

    const headersClientes = [
      'Cliente_ID','Nombre','NIF','Direccion','CP','Ciudad','Telefono','Email','Tipo_cliente','Origen','Fecha_alta'
    ];

    const headersLeads = [
      'Lead_ID','Fecha_entrada','Nombre','Email','Telefono','NIF/CIF','Direccion','CP','Poblacion',
      'Tipo_servicio','Tipo_propiedad','m2','Habitaciones','Banos','Terraza','Mascotas',
      'Fecha_servicio','Hora_preferida','Frecuencia','Canal_preferido','Mensaje/Notas',
      'Estado','Cliente_ID','Ultimo_contacto','Origen','RowKey'
    ];

    const headersPresupuestos = [
      'Pres_ID','Fecha','Validez_dias','Vence_el','Estado','Cliente_ID',
      'Cliente','Email_cliente','NIF','Direccion','CP','Ciudad',
      'Base','IVA_total','Total','Notas','PDF_link','Factura_ID','Fecha_envio','Fecha_aceptacion','Archivado_el'
    ];

    const headersPresLineas = ['Pres_ID','Linea_n','Concepto','Cantidad','Precio','Dto_%','IVA_%','Subtotal'];
    const headersLineas = ['Numero_factura','Concepto','Cantidad','Precio','Subtotal'];

    const headersFactura = [
      'Factura_ID','Fecha','Estado','Cliente_ID','Cliente','Email_cliente','NIF','Direccion','CP','Ciudad',
      'Base','IVA_total','Total','Notas','PDF_link','Fecha_pago'
    ];
    const headersFacturas = [
      'Factura_ID','Fecha','Estado','Pres_ID','Cliente_ID','Cliente','Email','NIF','Direccion','CP','Ciudad',
      'Base','IVA_total','Total','Notas','PDF_link','Fecha_envio','Fecha_pago'
    ];
    const headersFactLineas = ['Factura_ID','Linea_n','Concepto','Cantidad','Precio','Dto_%','IVA_%','Subtotal'];

    const headersLog = ['Fecha','Modulo','Accion','Entidad','ID','Resultado','Mensaje','DataJSON'];

    const shConfig = ensureSheet_(ss, 'CONFIG');
    ensureHeaders_(shConfig, headersConfig);
    ensureConfigDefaults_(shConfig);
    ensureHeaders_(ensureSheet_(ss, 'CLIENTES'), headersClientes);
    ensureHeaders_(ensureSheet_(ss, 'LEADS'), headersLeads);

    const shPres = ensureSheet_(ss, 'PRESUPUESTOS');
    ensureHeaders_(shPres, headersPresupuestos);
    ensurePresupuestoLeadColumns_(shPres, [
      'Tipo_destinatario','Lead_ID','Lead_RowKey','Lead_Nombre','Lead_Email','Lead_NIF','Lead_Telefono','Lead_Direccion'
    ]);

    ensureHeaders_(ensureSheet_(ss, 'PRES_LINEAS'), headersPresLineas);
    ensureHeaders_(ensureSheet_(ss, 'LINEAS'), headersLineas);
    ensureHeaders_(ensureSheet_(ss, 'FACTURA'), headersFactura);
    ensureHeaders_(ensureSheet_(ss, 'FACTURAS'), headersFacturas);
    ensureHeaders_(ensureSheet_(ss, 'FACT_LINEAS'), headersFactLineas);
    ensureHeaders_(ensureSheet_(ss, 'LOG'), headersLog);

    applyListValidation_(ss.getSheetByName('LEADS'), 22, ['Nuevo','Ganado','Perdido']);
    setupValidationsPresupuestos();
    if (typeof factApplyValidations_ === 'function') factApplyValidations_();
    if (typeof ccSetupWebAppLayer_ === 'function') ccSetupWebAppLayer_();

    installTriggers_(ss);

    log('RUN', 'OK', '', null);
  } catch (err) {
    log('RUN', 'ERROR', err.message || String(err), null);
    throw err;
  }
}
var setupAll__legacy_ = function() {
  return setupAll_();
}
var ensureSheet_ = function(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}
var ensureHeaders_ = function(sh, headers) {
  const lastRow = sh.getLastRow();
  if (lastRow === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return;
  }

  const row = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const empty = row.every((v) => String(v || '').trim() === '');
  if (empty) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
}
var ensureConfigDefaults_ = function(sh) {
  if (!sh) return;
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const values = sh.getRange(2, 1, 1, lastCol).getValues()[0];

  const ensureHeader = (header) => {
    let idx = headers.indexOf(header);
    if (idx !== -1) return idx;
    const newCol = sh.getLastColumn() + 1;
    sh.insertColumnsAfter(sh.getLastColumn(), 1);
    sh.getRange(1, newCol).setValue(header);
    headers.push(header);
    values.push('');
    return headers.length - 1;
  };

  const setIfEmpty = (header, value) => {
    const idx = ensureHeader(header);
    const current = values[idx];
    if (String(current || '').trim()) return;
    sh.getRange(2, idx + 1).setValue(value || '');
  };

  setIfEmpty('PRES_Pdf_Folder_Id', CC_DEFAULT_IDS.PRESUPUESTOS_FOLDER_ID);
  setIfEmpty('FACT_Pdf_Folder_Id', CC_DEFAULT_IDS.FACTURAS_FOLDER_ID);
  setIfEmpty('PRES_Template_DocId', CC_DEFAULT_IDS.PRESUPUESTO_TEMPLATE_ID);
  setIfEmpty('FACT_Template_DocId', CC_DEFAULT_IDS.FACTURA_TEMPLATE_ID);
}
var ensurePresupuestoLeadColumns_ = function(sh, headers) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = headers.filter((h) => existing.indexOf(h) === -1);
  if (!missing.length) return;

  const startCol = lastCol + 1;
  sh.insertColumnsAfter(lastCol, missing.length);
  sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
}
var applyListValidation_ = function(sh, col, values) {
  if (!sh) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)
    .build();
  const rows = Math.max(1, sh.getMaxRows() - 1);
  sh.getRange(2, col, rows, 1).setDataValidation(rule);
}
var setupTriggers_ = function(ss) {
  const targets = {
    onEdit: { handler: 'onEdit', type: ScriptApp.EventType.ON_EDIT },
    onOpen: { handler: 'onOpen', type: ScriptApp.EventType.ON_OPEN },
    onFormSubmit: { handler: 'onFormSubmit', type: ScriptApp.EventType.ON_FORM_SUBMIT }
  };

  ScriptApp.getProjectTriggers().forEach((t) => {
    const handler = t.getHandlerFunction();
    const type = t.getEventType();
    if (
      (handler === targets.onEdit.handler && type === targets.onEdit.type) ||
      (handler === targets.onOpen.handler && type === targets.onOpen.type) ||
      (handler === targets.onFormSubmit.handler && type === targets.onFormSubmit.type)
    ) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger(targets.onEdit.handler).forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger(targets.onOpen.handler).forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger(targets.onFormSubmit.handler).forSpreadsheet(ss).onFormSubmit().create();
}
var installTriggers_ = function(ss) {
  setupTriggers_(ss);
}











  return {
    setupAll: setupAll_
  };
})();

// Entry-point visible (solo este debe quedar como 'function' top-level)
function setupAll(){ return __CC_SETUP.setupAll(); }




/** Fix PRO: limpia props viejas y fija CC_DB_SPREADSHEET_ID con SS_ID */
function ccFixDbIdProps(){
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  Object.keys(all).forEach(function(k){
    if(/^1m62QB04/.test(k)) props.deleteProperty(k);
  });
  props.setProperty("CC_DB_SPREADSHEET_ID", String(SS_ID));
  return props.getProperties();
}


function run_ccFixDbIdProps(){ return ccFixDbIdProps(); }





/** PRO DIAG: obtiene el DB ID real desde Script Properties (fallback a SS_ID si existe) */
function ccGetDbId_(){
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CC_DB_SPREADSHEET_ID');
  if (id) return String(id).trim();
  try {
    if (typeof SS_ID !== 'undefined' && SS_ID) return String(SS_ID).trim();
  } catch(e){}
  return '';
}

/** PRO DIAG: valida estructura del spreadsheet y headers */
function ccDiagDbStructure(){
  var expected = ['CLIENTES','LEADS','HISTORIAL','HISTORIAL_PRESUPUESTOS','GASTOS'];

  var id = ccGetDbId_();
  if (!id) throw new Error('No hay CC_DB_SPREADSHEET_ID (ni SS_ID). Ejecuta run_ccFixDbIdProps primero.');

  var ss = SpreadsheetApp.openById(id);
  var allSheets = ss.getSheets().map(function(s){ return s.getName(); });

  var out = {
    ok: true,
    spreadsheetId: id,
    spreadsheetName: ss.getName(),
    allSheets: allSheets,
    expected: expected,
    missing: [],
    details: {}
  };

  expected.forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (!sh){
      out.missing.push(name);
      return;
    }
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();

    var headers = [];
    if (lastCol > 0){
      headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(function(h){
        return String(h || '').trim();
      });
    }

    out.details[name] = {
      exists: true,
      lastRow: lastRow,
      lastCol: lastCol,
      headers: headers
    };
  });

  if (out.missing.length) out.ok = false;

  var msg = 'ccDiagDbStructure => ' + JSON.stringify(out, null, 2);
  try { Logger.log(msg); } catch(e){}
  try { console.log(msg); } catch(e){}
  return out;
}

/** Runner visible para dropdown */
function run_ccDiagDbStructure(){ return ccDiagDbStructure(); }
/** DIAG PRO: inspecciona SpreadsheetApp + prueba openById con SS_ID y con ScriptProperty */
function diagSpreadsheetApp(){
  var out = {};
  out.typeof_SpreadsheetApp = typeof SpreadsheetApp;

  // SS_ID (global)
  try { out.SS_ID = (typeof SS_ID !== "undefined") ? String(SS_ID) : null; }
  catch(e){ out.SS_ID_err = String(e); }

  // Script Property
  try {
    out.PROP_CC_DB_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty("CC_DB_SPREADSHEET_ID");
  } catch(e){ out.PROP_err = String(e); }

  // ¿Existe openById?
  try { out.has_openById = !!SpreadsheetApp.openById; }
  catch(e){ out.has_openById_err = String(e); }

  // Prueba openById con SS_ID
  try {
    var ss1 = SpreadsheetApp.openById(out.SS_ID);
    out.openById_SS_ID_ok = ss1.getId();
    out.openById_SS_ID_name = ss1.getName();
  } catch(e){ out.openById_SS_ID_err = String(e); }

  // Prueba openById con PROP
  try {
    var ss2 = SpreadsheetApp.openById(out.PROP_CC_DB_SPREADSHEET_ID);
    out.openById_PROP_ok = ss2.getId();
    out.openById_PROP_name = ss2.getName();
  } catch(e){ out.openById_PROP_err = String(e); }

  // Active spreadsheet (si aplica)
  try {
    var a = SpreadsheetApp.getActiveSpreadsheet();
    out.activeSpreadsheetId = a ? a.getId() : null;
  } catch(e){ out.activeSpreadsheet_err = String(e); }
  var msg = 'diagSpreadsheetApp => ' + JSON.stringify(out, null, 2);
  try { Logger.log(msg); } catch(e){}
  try { console.log(msg); } catch(e){}


  return out;
}
function run_diagSpreadsheetApp(){ return diagSpreadsheetApp(); }



