function setupAll() {
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
      'CLI_Ano','CLI_Ultimo_numero','PRES_Año','PRES_Ultimo_numero','PRES_Validez_default','PRES_Template_DocId'
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

    const headersLog = ['Fecha','Modulo','Accion','Entidad','ID','Resultado','Mensaje','DataJSON'];

    ensureHeaders_(ensureSheet_(ss, 'CONFIG'), headersConfig);
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
    ensureHeaders_(ensureSheet_(ss, 'LOG'), headersLog);

    applyListValidation_(ss.getSheetByName('LEADS'), 22, ['Nuevo','Ganado','Perdido']);
    applyListValidation_(shPres, 5, ['Borrador','Enviado','Aceptado','Rechazado','Archivado']);

    installTriggers_(ss);

    log('RUN', 'OK', '', null);
  } catch (err) {
    log('RUN', 'ERROR', err.message || String(err), null);
    throw err;
  }
}

function setupAll_() {
  return setupAll();
}

function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeaders_(sh, headers) {
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

function ensurePresupuestoLeadColumns_(sh, headers) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = headers.filter((h) => existing.indexOf(h) === -1);
  if (!missing.length) return;

  const startCol = lastCol + 1;
  sh.insertColumnsAfter(lastCol, missing.length);
  sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
}

function applyListValidation_(sh, col, values) {
  if (!sh) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)
    .build();
  const rows = Math.max(1, sh.getMaxRows() - 1);
  sh.getRange(2, col, rows, 1).setDataValidation(rule);
}

function setupTriggers_(ss) {
  const targets = {
    onEdit: { handler: 'onEdit', type: ScriptApp.EventType.ON_EDIT },
    onOpen: { handler: 'onOpen', type: ScriptApp.EventType.ON_OPEN }
  };

  ScriptApp.getProjectTriggers().forEach((t) => {
    const handler = t.getHandlerFunction();
    const type = t.getEventType();
    if (
      (handler === targets.onEdit.handler && type === targets.onEdit.type) ||
      (handler === targets.onOpen.handler && type === targets.onOpen.type)
    ) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger(targets.onEdit.handler).forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger(targets.onOpen.handler).forSpreadsheet(ss).onOpen().create();
}
function installTriggers_(ss) {
  setupTriggers_(ss);
}


