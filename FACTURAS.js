const SH_FACTURAS = 'FACTURAS';
const SH_FACT_LINEAS = 'FACT_LINEAS';

const FACT_HEADERS = [
  'Factura_ID','Fecha','Estado','Pres_ID','Cliente_ID','Cliente','Email','NIF','Direccion','CP','Ciudad',
  'Base','IVA_total','Total','Notas','PDF_link','Fecha_envio','Fecha_pago'
];

const FACT_ESTADOS = ['BORRADOR','EMITIDA','ENVIADA','PENDIENTE','VENCIDA','IMPAGADA','PAGADA','ANULADA'];

const FACT_LINEAS_HEADERS = ['Factura_ID','Linea_n','Concepto','Cantidad','Precio','Dto_%','IVA_%','Subtotal'];

const FACT_FOLDER_ID_DEFAULT = '111q4NpNNT_W_jZ8Mgw0AVv2VN_PFsGp';
const FACT_TEMPLATE_ID_DEFAULT = '1OU_1CxZBEc46OP5W1d98al1ylIO0UV7Tln7qShH6as';
const PRES_FOLDER_ID_DEFAULT = '1b4R5P30DULl-Fp_PY8dmuVLg6UfuJjo9';
const PRES_TEMPLATE_ID_DEFAULT = '1M2tpK-Iq6_WuVmHxahkHbtJrmOtLmPu502-TjnqkQ8';

function factAsegurarEstructura_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SH_FACTURAS);
  if (!sh) sh = ss.insertSheet(SH_FACTURAS);
  factEnsureHeaders_(sh, FACT_HEADERS);

  let shL = ss.getSheetByName(SH_FACT_LINEAS);
  if (!shL) shL = ss.insertSheet(SH_FACT_LINEAS);
  factEnsureHeaders_(shL, FACT_LINEAS_HEADERS);

  factEnsurePdfConfig_();
}

function factEnsureHeaders_(sh, headers) {
  const lastRow = sh.getLastRow();
  if (lastRow === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return;
  }
  const first = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const empty = first.every(v => String(v || '').trim() === '');
  if (empty) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < headers.length) {
    sh.insertColumnsAfter(sh.getLastColumn(), headers.length - sh.getLastColumn());
  }
}

function factEnsurePdfConfig_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('FACT_Pdf_Folder_Id')) {
    props.setProperty('FACT_Pdf_Folder_Id', FACT_FOLDER_ID_DEFAULT);
  }
  if (!props.getProperty('FACT_Template_DocId')) {
    props.setProperty('FACT_Template_DocId', FACT_TEMPLATE_ID_DEFAULT);
  }
  if (!props.getProperty('PRES_Pdf_Folder_Id')) {
    props.setProperty('PRES_Pdf_Folder_Id', PRES_FOLDER_ID_DEFAULT);
  }
  if (!props.getProperty('PRES_Template_DocId')) {
    props.setProperty('PRES_Template_DocId', PRES_TEMPLATE_ID_DEFAULT);
  }

  if (typeof setCfgValueIfSheet_ === 'function') {
    setCfgValueIfSheet_('FACT_Pdf_Folder_Id', props.getProperty('FACT_Pdf_Folder_Id'));
    setCfgValueIfSheet_('FACT_Template_DocId', props.getProperty('FACT_Template_DocId'));
    setCfgValueIfSheet_('PRES_Pdf_Folder_Id', props.getProperty('PRES_Pdf_Folder_Id'));
    setCfgValueIfSheet_('PRES_Template_DocId', props.getProperty('PRES_Template_DocId'));
  }
}

function factApplyValidations_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shFact = ss.getSheetByName(SH_FACTURAS);
  const shFactSingle = ss.getSheetByName('FACTURA');
  const shCli = ss.getSheetByName('CLIENTES');
  const shLeads = ss.getSheetByName('LEADS');

  [shFact, shFactSingle].forEach((sheet) => {
    if (!sheet) return;
    const headerInfo = presGetHeaderMap_(sheet);
    const maxRows = Math.max(1, sheet.getMaxRows() - 1);

    if (headerInfo.map['Estado']) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(FACT_ESTADOS, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, headerInfo.map['Estado'], maxRows, 1).setDataValidation(rule);
    }

    const cliRule = (typeof presBuildValidationRuleFromSheet_ === 'function')
      ? presBuildValidationRuleFromSheet_(shCli, 'Cliente_ID')
      : null;
    if (headerInfo.map['Cliente_ID'] && cliRule) {
      sheet.getRange(2, headerInfo.map['Cliente_ID'], maxRows, 1).setDataValidation(cliRule);
    }

    const leadRule = (typeof presBuildValidationRuleFromSheet_ === 'function')
      ? presBuildValidationRuleFromSheet_(shLeads, 'Lead_ID')
      : null;
    if (headerInfo.map['Lead_ID'] && leadRule) {
      sheet.getRange(2, headerInfo.map['Lead_ID'], maxRows, 1).setDataValidation(leadRule);
    }
  });
}

function createFacturaDesdePresupuesto_(presId, options) {
  const opts = options || {};
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('No se pudo obtener el bloqueo para crear la factura.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = String(presId || '').trim();
  if (!id) throw new Error('Pres_ID requerido');

  const log = (resultado, mensaje, data) => {
    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'FACTURAS', 'CREATE_FROM_PRES', 'PRESUPUESTOS', id, resultado, mensaje || '', data || null);
    }
  };

  try {
    presAsegurarEstructura_();
    factAsegurarEstructura_();

    const shPres = ss.getSheetByName(SH_PRES);
    if (!shPres) throw new Error('No existe hoja PRESUPUESTOS');

    const presRow = presFindPresRow_(shPres, id);
    const presObj = presRow.rowObj || {};
    const estado = String(presObj.Estado || '').trim().toUpperCase();
    if (estado !== 'ACEPTADO') {
      throw new Error('El presupuesto debe estar en estado ACEPTADO.');
    }

    const shFact = ss.getSheetByName(SH_FACTURAS);
    const existingFacturaId = String(presObj.Factura_ID || '').trim();
    const existingByPres = factFindFacturaByPresId_(shFact, id);

    const facturaId = existingFacturaId || (existingByPres ? existingByPres.facturaId : '');
    if (facturaId) {
      if (opts.regenerarPdf) {
        const pdfUrl = generateFacturaPdfById(facturaId);
        factUpdatePdfLink_(shFact, facturaId, pdfUrl);
        presUpdateFacturaLink_(shPres, presRow, facturaId, pdfUrl);
        log('OK', 'PDF regenerado', { facturaId, pdfUrl });
        return { facturaId, pdfUrl, alreadyExisted: true };
      }
      if (!existingFacturaId) {
        presUpdateFacturaLink_(shPres, presRow, facturaId, existingByPres ? existingByPres.pdfUrl : '');
      }
      log('SKIP', 'Factura ya existe', { facturaId });
      return { facturaId, pdfUrl: existingByPres ? existingByPres.pdfUrl : '', alreadyExisted: true };
    }

    const nuevaFacturaId = (typeof consumirSiguienteNumero_ === 'function')
      ? consumirSiguienteNumero_()
      : factFallbackNextId_(shFact);

    const lineasInfo = factGetLineasFromPresupuesto_(ss, id);
    if (!lineasInfo.lineas.length) throw new Error('No hay lineas para este presupuesto');

    const totals = factComputeTotals_(lineasInfo.lineas);
    const base = Number(presObj.Base || 0) || totals.base;
    const ivaTotal = Number(presObj.IVA_total || 0) || totals.ivaTotal;
    const total = Number(presObj.Total || 0) || totals.total;

    const headerInfo = presGetHeaderMap_(shFact);
    const rowValues = presBuildRow_(headerInfo.headers, {
      Factura_ID: nuevaFacturaId,
      Fecha: new Date(),
      Estado: 'EMITIDA',
      Pres_ID: id,
      Cliente_ID: presObj.Cliente_ID || '',
      Cliente: presObj.Cliente || '',
      Email: presObj.Email_cliente || '',
      NIF: presObj.NIF || '',
      Direccion: presObj.Direccion || '',
      CP: presObj.CP || '',
      Ciudad: presObj.Ciudad || '',
      Base: base,
      IVA_total: ivaTotal,
      Total: total,
      Notas: presObj.Notas || '',
      PDF_link: ''
    });

    shFact.appendRow(rowValues);
    factInsertLineas_(ss, nuevaFacturaId, lineasInfo.lineas);

    const pdfUrl = generateFacturaPdfById(nuevaFacturaId);
    factUpdatePdfLink_(shFact, nuevaFacturaId, pdfUrl);
    presUpdateFacturaLink_(shPres, presRow, nuevaFacturaId, pdfUrl);

    log('OK', '', { facturaId: nuevaFacturaId, pdfUrl });
    return { facturaId: nuevaFacturaId, pdfUrl, alreadyExisted: false };
  } catch (err) {
    log('ERROR', err.message || String(err), { stack: err.stack });
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function aceptarYFacturarPresupuesto_(presId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = String(presId || '').trim();
  if (!id) throw new Error('Pres_ID requerido');

  const shPres = ss.getSheetByName(SH_PRES);
  if (!shPres) throw new Error('No existe hoja PRESUPUESTOS');

  const presRow = presFindPresRow_(shPres, id);
  const { headerInfo } = presRow;
  const estadoCol = headerInfo.map['Estado'];
  const fechaAceptCol = headerInfo.map['Fecha_aceptacion'];
  if (estadoCol) shPres.getRange(presRow.rowNumber, estadoCol).setValue('ACEPTADO');
  if (fechaAceptCol) shPres.getRange(presRow.rowNumber, fechaAceptCol).setValue(new Date());

  return createFacturaDesdePresupuesto_(id);
}

function uiCrearFacturaDesdePresupuestoActivo_() {
  presAsegurarEstructura_();
  factAsegurarEstructura_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (sh.getName() !== SH_PRES) {
    SpreadsheetApp.getUi().alert('Ve a la hoja PRESUPUESTOS y selecciona una fila.');
    return;
  }
  const row = ss.getActiveRange().getRow();
  if (row < 2) return;

  const data = presGetRowData_(sh, row);
  const presId = String(data.Pres_ID || '').trim();
  if (!presId) throw new Error('Fila sin Pres_ID');

  const res = createFacturaDesdePresupuesto_(presId);
  const msg = res.alreadyExisted
    ? 'La factura ya existe: ' + res.facturaId
    : 'Factura creada: ' + res.facturaId;
  SpreadsheetApp.getUi().alert(msg);
}

function uiGenerarPdfPresupuestoActivo_() {
  presAsegurarEstructura_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (sh.getName() !== SH_PRES) {
    SpreadsheetApp.getUi().alert('Ve a la hoja PRESUPUESTOS y selecciona una fila.');
    return;
  }
  const row = ss.getActiveRange().getRow();
  if (row < 2) return;

  const data = presGetRowData_(sh, row);
  const presId = String(data.Pres_ID || '').trim();
  if (!presId) throw new Error('Fila sin Pres_ID');

  const url = (typeof generatePresupuestoPdfById === 'function')
    ? generatePresupuestoPdfById(presId)
    : generarPDFPresupuesto(presId, { archivar: true });
  SpreadsheetApp.getUi().alert('PDF generado:\n' + url);
}

function uiGenerarPdfFacturaActiva_() {
  factAsegurarEstructura_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (sh.getName() !== SH_FACTURAS) {
    SpreadsheetApp.getUi().alert('Ve a la hoja FACTURAS y selecciona una fila.');
    return;
  }
  const row = ss.getActiveRange().getRow();
  if (row < 2) return;

  const headerInfo = presGetHeaderMap_(sh);
  const colId = headerInfo.map['Factura_ID'];
  const facturaId = colId ? String(sh.getRange(row, colId).getDisplayValue()).trim() : '';
  if (!facturaId) throw new Error('Fila sin Factura_ID');

  const url = generateFacturaPdfById(facturaId);
  SpreadsheetApp.getUi().alert('PDF generado:\n' + url);
}

function factFindFacturaByPresId_(shFact, presId) {
  const headerInfo = presGetHeaderMap_(shFact);
  const colPres = headerInfo.map['Pres_ID'];
  const colFact = headerInfo.map['Factura_ID'];
  const colPdf = headerInfo.map['PDF_link'];
  if (!colPres || !colFact) return null;

  const lastRow = shFact.getLastRow();
  if (lastRow < 2) return null;
  const data = shFact.getRange(2, 1, lastRow - 1, shFact.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][colPres - 1]).trim() === String(presId).trim()) {
      return {
        row: i + 2,
        facturaId: String(data[i][colFact - 1] || '').trim(),
        pdfUrl: colPdf ? String(data[i][colPdf - 1] || '').trim() : ''
      };
    }
  }
  return null;
}

function factGetLineasFromPresupuesto_(ss, presId) {
  const shLin = ss.getSheetByName(SH_PRES_LINEAS);
  if (shLin) {
    const data = presReadLineasPorPresId_(shLin, presId);
    if (data.lineas.length) return data;
  }
  const shHist = ss.getSheetByName(SH_PRES_LINEAS_HIS);
  if (shHist) return presReadLineasPorPresId_(shHist, presId);
  return { headers: [], lineas: [] };
}

function factComputeTotals_(lineas) {
  let base = 0;
  let ivaTotal = 0;
  lineas.forEach((l) => {
    const cantidad = factParseNumber_(l.Cantidad || l.cantidad || 0);
    const precio = factParseNumber_(l.Precio || l.precio || 0);
    const dto = factParseNumber_(l['Dto_%'] || l.dto || 0);
    const iva = factParseNumber_(l['IVA_%'] || l.iva || 0);
    const subtotal = factParseNumber_(l.Subtotal || l.subtotal || 0) || (cantidad * precio * (1 - (dto / 100)));
    base += subtotal;
    if (iva) ivaTotal += subtotal * (iva / 100);
  });
  const baseRound = Math.round(base * 100) / 100;
  const ivaRound = Math.round(ivaTotal * 100) / 100;
  return { base: baseRound, ivaTotal: ivaRound, total: Math.round((baseRound + ivaRound) * 100) / 100 };
}

function factParseNumber_(v) {
  if (v === null || v === undefined || v === '') return 0;
  const raw = String(v).replace(',', '.');
  const num = Number(raw);
  return isNaN(num) ? 0 : num;
}

function factInsertLineas_(ss, facturaId, lineas) {
  const sh = ss.getSheetByName(SH_FACT_LINEAS);
  if (!sh) throw new Error('No existe hoja ' + SH_FACT_LINEAS);

  const headerInfo = presGetHeaderMap_(sh);
  const baseRow = sh.getLastRow() + 1;
  const rows = (lineas || []).map((l, idx) => presBuildRow_(headerInfo.headers, {
    Factura_ID: facturaId,
    Linea_n: l.Linea_n || l.linea || (idx + 1),
    Concepto: l.Concepto || l.concepto || '',
    Cantidad: l.Cantidad || l.cantidad || '',
    Precio: l.Precio || l.precio || '',
    'Dto_%': l['Dto_%'] || l.dto || '',
    'IVA_%': l['IVA_%'] || l.iva || '',
    Subtotal: l.Subtotal || l.subtotal || ''
  }));

  if (rows.length) {
    sh.getRange(baseRow, 1, rows.length, headerInfo.headers.length).setValues(rows);
  }
}

function factUpdatePdfLink_(shFact, facturaId, pdfUrl) {
  const headerInfo = presGetHeaderMap_(shFact);
  const colId = headerInfo.map['Factura_ID'];
  const colPdf = headerInfo.map['PDF_link'];
  if (!colId || !colPdf) return;

  const lastRow = shFact.getLastRow();
  if (lastRow < 2) return;
  const data = shFact.getRange(2, colId, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(facturaId).trim()) {
      shFact.getRange(i + 2, colPdf).setValue(pdfUrl || '');
      return;
    }
  }
}

function presUpdateFacturaLink_(shPres, presRow, facturaId, pdfUrl) {
  const map = presRow.headerInfo.map;
  if (map['Factura_ID']) shPres.getRange(presRow.rowNumber, map['Factura_ID']).setValue(facturaId);
  if (map['PDF_link_factura']) shPres.getRange(presRow.rowNumber, map['PDF_link_factura']).setValue(pdfUrl || '');
}

function factFallbackNextId_(shFact) {
  const year = new Date().getFullYear();
  const lastRow = shFact.getLastRow();
  if (lastRow < 2) return year + '-001';
  const headerInfo = presGetHeaderMap_(shFact);
  const colId = headerInfo.map['Factura_ID'];
  if (!colId) return year + '-001';
  const lastVal = String(shFact.getRange(lastRow, colId).getValue() || '').trim();
  const n = Number(lastVal.replace(/[^\d]/g, ''));
  const next = isNaN(n) ? 1 : n + 1;
  return year + '-' + String(next).padStart(3, '0');
}

function seedSampleData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  presAsegurarEstructura_();
  factAsegurarEstructura_();

  const shClientes = ss.getSheetByName('CLIENTES');
  const shLeads = ss.getSheetByName('LEADS');
  const shPres = ss.getSheetByName(SH_PRES);
  const shLineas = ss.getSheetByName(SH_PRES_LINEAS);

  if (!shClientes || !shLeads || !shPres || !shLineas) {
    throw new Error('Faltan hojas CLIENTES, LEADS, PRESUPUESTOS o PRES_LINEAS');
  }

  const clientes = [
    {
      Cliente_ID: 'CLI-SEED-0001',
      Nombre: 'Comunidad Pineda',
      NIF: 'B66778899',
      Direccion: 'C/ Pineda 123',
      CP: '08397',
      Ciudad: 'Pineda de Mar',
      Telefono: '+34 611 222 333',
      Email: 'admin@pineda.com',
      Tipo_cliente: 'Comunidad',
      Origen: 'Seed',
      Fecha_alta: new Date()
    },
    {
      Cliente_ID: 'CLI-SEED-0002',
      Nombre: 'Marta Ramos',
      NIF: 'X1234567A',
      Direccion: 'C/ Mallorca 55',
      CP: '08013',
      Ciudad: 'Barcelona',
      Telefono: '+34 600 123 123',
      Email: 'marta@email.com',
      Tipo_cliente: 'Particular',
      Origen: 'Seed',
      Fecha_alta: new Date()
    }
  ];

  seedUpsertById_(shClientes, 'Cliente_ID', clientes);

  const leads = [
    {
      Lead_ID: 'LSEED-0001',
      Fecha_entrada: new Date(),
      Nombre: 'Comunidad Pineda (Lead)',
      Email: 'lead@pineda.com',
      Telefono: '+34 611 222 333',
      'NIF/CIF': 'B66778899',
      Direccion: 'C/ Pineda 123',
      CP: '08397',
      Poblacion: 'Pineda de Mar',
      Tipo_servicio: 'Limpieza comunidades',
      Tipo_propiedad: 'Comunidad',
      m2: 500,
      Habitaciones: '',
      Banos: '',
      Terraza: 'Si',
      Mascotas: 'No',
      Fecha_servicio: '',
      Hora_preferida: '',
      Frecuencia: 'Semanal',
      Canal_preferido: 'Email',
      'Mensaje/Notas': 'Necesitamos presupuesto mensual.',
      Estado: 'Nuevo',
      Cliente_ID: '',
      Ultimo_contacto: '',
      Origen: 'Seed',
      RowKey: 'seed|pineda|1'
    },
    {
      Lead_ID: 'LSEED-0002',
      Fecha_entrada: new Date(),
      Nombre: 'Laura Costa (Lead)',
      Email: 'laura@costa.com',
      Telefono: '+34 699 111 222',
      'NIF/CIF': 'Y7654321B',
      Direccion: 'C/ Girona 44',
      CP: '08009',
      Poblacion: 'Barcelona',
      Tipo_servicio: 'Limpieza hogar',
      Tipo_propiedad: 'Piso',
      m2: 90,
      Habitaciones: 3,
      Banos: 2,
      Terraza: 'No',
      Mascotas: 'No',
      Fecha_servicio: '',
      Hora_preferida: '',
      Frecuencia: 'Quincenal',
      Canal_preferido: 'WhatsApp',
      'Mensaje/Notas': 'Interesa limpieza profunda.',
      Estado: 'Nuevo',
      Cliente_ID: '',
      Ultimo_contacto: '',
      Origen: 'Seed',
      RowKey: 'seed|laura|1'
    }
  ];

  seedUpsertById_(shLeads, 'Lead_ID', leads);

  const presupuestos = [
    { id: 'PRO-SEED-0001', estado: 'BORRADOR', clienteId: 'CLI-SEED-0001' },
    { id: 'PRO-SEED-0002', estado: 'BORRADOR', clienteId: 'CLI-SEED-0002' },
    { id: 'PRO-SEED-0003', estado: 'ENVIADO', clienteId: 'CLI-SEED-0001' },
    { id: 'PRO-SEED-0004', estado: 'ENVIADO', clienteId: 'CLI-SEED-0002' },
    { id: 'PRO-SEED-0005', estado: 'ACEPTADO', clienteId: 'CLI-SEED-0001' }
  ];

  presupuestos.forEach((p, idx) => {
    const presId = p.id;
    if (seedRowExists_(shPres, 'Pres_ID', presId)) return;

    const cli = clientes.find(c => c.Cliente_ID === p.clienteId);
    const fecha = new Date();
    const vence = new Date(fecha.getTime() + 15 * 86400000);

    const headerInfo = presGetHeaderMap_(shPres);
    const rowValues = presBuildRow_(headerInfo.headers, {
      Pres_ID: presId,
      Fecha: fecha,
      Validez_dias: 15,
      Vence_el: vence,
      Estado: p.estado,
      Cliente_ID: cli ? cli.Cliente_ID : '',
      Cliente: cli ? cli.Nombre : '',
      Email_cliente: cli ? cli.Email : '',
      NIF: cli ? cli.NIF : '',
      Direccion: cli ? cli.Direccion : '',
      CP: cli ? cli.CP : '',
      Ciudad: cli ? cli.Ciudad : '',
      Base: '',
      IVA_total: '',
      Total: '',
      Notas: 'Presupuesto seed ' + (idx + 1),
      PDF_link: '',
      Factura_ID: '',
      Fecha_envio: (p.estado === 'ENVIADO') ? new Date() : '',
      Fecha_aceptacion: (p.estado === 'ACEPTADO') ? new Date() : '',
      Archivado_el: ''
    });

    shPres.appendRow(rowValues);

    const lineas = [
      { Linea_n: 1, Concepto: 'Servicio limpieza base', Cantidad: 1, Precio: 120, 'Dto_%': 0, 'IVA_%': 21 },
      { Linea_n: 2, Concepto: 'Cristales', Cantidad: 1, Precio: 35, 'Dto_%': 0, 'IVA_%': 21 }
    ];

    seedInsertPresLineas_(shLineas, presId, lineas);
  });

  SpreadsheetApp.getUi().alert('Datos de prueba cargados (seed).');
}

function seedUpsertById_(sh, idHeader, rows) {
  const headerInfo = presGetHeaderMap_(sh);
  const headers = headerInfo.headers;
  const idCol = headerInfo.map[idHeader];
  if (!idCol) throw new Error('No existe columna ' + idHeader + ' en ' + sh.getName());

  const existing = new Set();
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const ids = sh.getRange(2, idCol, lastRow - 1, 1).getValues();
    ids.forEach(r => { if (r[0]) existing.add(String(r[0]).trim()); });
  }

  const newRows = rows
    .filter(r => !existing.has(String(r[idHeader] || '').trim()))
    .map(r => presBuildRow_(headers, r));

  if (newRows.length) {
    sh.getRange(lastRow + 1, 1, newRows.length, headers.length).setValues(newRows);
  }
}

function seedRowExists_(sh, idHeader, idValue) {
  const headerInfo = presGetHeaderMap_(sh);
  const col = headerInfo.map[idHeader];
  if (!col) return false;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  const ids = sh.getRange(2, col, lastRow - 1, 1).getValues();
  return ids.some(r => String(r[0] || '').trim() === String(idValue).trim());
}

function seedInsertPresLineas_(shLineas, presId, lineas) {
  const headerInfo = presGetHeaderMap_(shLineas);
  const existing = seedRowExists_(shLineas, 'Pres_ID', presId);
  if (existing) return;

  const rows = (lineas || []).map(l => presBuildRow_(headerInfo.headers, {
    Pres_ID: presId,
    Linea_n: l.Linea_n || '',
    Concepto: l.Concepto || '',
    Cantidad: l.Cantidad || '',
    Precio: l.Precio || '',
    'Dto_%': l['Dto_%'] || 0,
    'IVA_%': l['IVA_%'] || 21,
    Subtotal: l.Subtotal || (Number(l.Cantidad || 0) * Number(l.Precio || 0))
  }));

  if (rows.length) {
    const start = shLineas.getLastRow() + 1;
    shLineas.getRange(start, 1, rows.length, headerInfo.headers.length).setValues(rows);
  }
}
