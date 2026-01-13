const PRESUPUESTO_TEMPLATE_ID = CC_DEFAULT_IDS.PRESUPUESTO_TEMPLATE_ID;
const FACTURA_TEMPLATE_ID = CC_DEFAULT_IDS.FACTURA_TEMPLATE_ID;
const PRESUPUESTO_FOLDER_ID = CC_DEFAULT_IDS.PRESUPUESTOS_FOLDER_ID;
const FACTURA_FOLDER_ID = CC_DEFAULT_IDS.FACTURAS_FOLDER_ID;

const PDF_PROP_PRES_FOLDER_ID = 'PRES_Pdf_Folder_Id';
const PDF_PROP_PRES_TEMPLATE_ID = 'PRES_Template_DocId';
const PDF_PROP_FACT_FOLDER_ID = 'FACT_Pdf_Folder_Id';
const PDF_PROP_FACT_TEMPLATE_ID = 'FACT_Template_DocId';

const PDF_PRES_FOLDER_NAME = 'Costa Clean - Presupuestos PDF';
const PDF_FACT_FOLDER_NAME = 'Costa Clean - Facturas PDF';

function generatePresupuestoPdfById(presId) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('No se pudo obtener el bloqueo para generar el PDF.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = String(presId || '').trim();
  if (!id) throw new Error('Pres_ID requerido');

  const log = (resultado, mensaje, data) => {
    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PDF', 'GENERATE_PRESUPUESTO', 'PRESUPUESTOS', id, resultado, mensaje || '', data || null);
    }
  };

  try {
    const shPres = pdfGetSheetByNames_(ss, ['PRESUPUESTOS', 'HISTORIAL_PRESUPUESTOS']);
    if (!shPres) throw new Error('No se encontro hoja PRESUPUESTOS.');

    const presRow = pdfFindRowById_(shPres, ['Pres_ID', 'Presupuesto_ID', 'ID'], id);
    if (!presRow) throw new Error('Presupuesto no encontrado: ' + id);

    const lineas = pdfGetPresupuestoLineas_(ss, id);
    if (!lineas.length) throw new Error('No hay lineas para el presupuesto ' + id);

    const cfg = pdfEnsurePdfConfig_(
      PDF_PRES_FOLDER_NAME,
      PDF_PROP_PRES_FOLDER_ID,
      PDF_PROP_PRES_TEMPLATE_ID,
      PRESUPUESTO_TEMPLATE_ID,
      'PRES_Pdf_Folder_Id',
      'PRES_Template_DocId',
      PRESUPUESTO_FOLDER_ID
    );
    const folder = DriveApp.getFolderById(cfg.folderId);

    const emisor = (typeof getEmisorDesdeFactura_ === 'function')
      ? getEmisorDesdeFactura_()
      : { nombre: '', nif: '', direccion: '', cp: '', ciudad: '' };

    const fechaTxt = pdfFormatDate_(presRow.obj[pdfFindHeader_(presRow.headers, ['Fecha'])]);
    const venceTxt = pdfFormatDate_(presRow.obj[pdfFindHeader_(presRow.headers, ['Vence_el', 'Vence'])]);

    const cliente = pdfPickValue_(presRow.obj, ['Cliente', 'Lead_Nombre', 'Cliente_nombre']);
    const email = pdfPickValue_(presRow.obj, ['Email_cliente', 'Lead_Email', 'Email']);
    const nif = pdfPickValue_(presRow.obj, ['NIF', 'DNI', 'CIF']);
    const direccion = pdfPickValue_(presRow.obj, ['Direccion']);
    const cp = pdfPickValue_(presRow.obj, ['CP']);
    const ciudad = pdfPickValue_(presRow.obj, ['Ciudad']);

    const totals = pdfComputeLineTotals_(lineas);
    const ivaPorc = totals.ivaPorc;

    const notasHeader = pdfFindHeader_(presRow.headers, ['Notas', 'NOTAS']);
    let notas = String(notasHeader ? presRow.obj[notasHeader] || '' : '').trim();

    const copyName = pdfSafeFileName_('Presupuesto_' + id + '_' + (cliente || 'Cliente'));
    const copy = DriveApp.getFileById(cfg.templateId).makeCopy(copyName, folder);
    const doc = DocumentApp.openById(copy.getId());

    if (pdfShouldAppendNoIvaNote_(notas, doc, notasHeader)) {
      notas = pdfAppendNote_(notas, 'Precios no incluyen IVA.');
      if (notasHeader) {
        shPres.getRange(presRow.rowNumber, presRow.headerMap[notasHeader]).setValue(notas);
      }
    }

    const placeholders = {
      PRES_ID: id,
      FECHA: fechaTxt,
      VENCE: venceTxt,
      CLIENTE: cliente,
      NIF: nif,
      DIRECCION: direccion,
      CP: cp,
      CIUDAD: ciudad,
      EMAIL: email,
      BASE: pdfMoney2_(totals.base),
      IVA_PORC: pdfNumber_(ivaPorc),
      IVA_TOTAL: pdfMoney2_(totals.ivaTotal),
      TOTAL: pdfMoney2_(totals.base),
      NOTAS: notas,
      Emisor_nombre: emisor.nombre || '',
      Emisor_NIF: emisor.nif || '',
      Emisor_direccion: emisor.direccion || '',
      Emisor_CP: emisor.cp || '',
      Emisor_ciudad: emisor.ciudad || ''
    };

    pdfReplaceTokens_(doc, placeholders);
    pdfFillLineTable_(doc, lineas, {
      LINEA_CONCEPTO: 'concepto',
      LINEA_CANTIDAD: 'cantidad',
      LINEA_PRECIO: 'precio',
      LINEA_SUBTOTAL: 'subtotal'
    });
    doc.saveAndClose();

    const pdfBlob = DriveApp.getFileById(copy.getId()).getAs(MimeType.PDF).setName(copyName + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);
    const pdfUrl = pdfFile.getUrl();

    pdfUpdateRowValues_(shPres, presRow, {
      Base: totals.base,
      IVA_total: totals.ivaTotal,
      Total: totals.base,
      PDF_link: pdfUrl
    });

    log('OK', '', { url: pdfUrl });
    return pdfUrl;
  } catch (err) {
    log('ERROR', err.message || String(err), { stack: err.stack });
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function regenerarPdfPresupuesto_(presId) {
  const id = String(presId || '').trim();
  if (!id) throw new Error('Pres_ID requerido');
  return generatePresupuestoPdfById(id);
}

function generateFacturaPdfById(factId) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('No se pudo obtener el bloqueo para generar el PDF.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = String(factId || '').trim();
  if (!id) throw new Error('Factura_ID requerido');

  const log = (resultado, mensaje, data) => {
    if (typeof logEvent_ === 'function') {
      logEvent_(ss, 'PDF', 'GENERATE_FACTURA', 'FACTURAS', id, resultado, mensaje || '', data || null);
    }
  };

  try {
    const shFact = pdfGetSheetByNames_(ss, ['FACTURAS', 'FACTURA']);
    if (!shFact) throw new Error('No se encontro hoja FACTURAS/FACTURA.');

    const factRow = pdfFindRowById_(shFact, ['Numero_factura', 'Factura_ID', 'ID'], id);
    if (!factRow) throw new Error('Factura no encontrada: ' + id);

    const lineas = pdfGetFacturaLineas_(ss, factRow.obj, id);
    if (!lineas.length) throw new Error('No hay lineas para la factura ' + id);

    const cfg = pdfEnsurePdfConfig_(
      PDF_FACT_FOLDER_NAME,
      PDF_PROP_FACT_FOLDER_ID,
      PDF_PROP_FACT_TEMPLATE_ID,
      FACTURA_TEMPLATE_ID,
      'FACT_Pdf_Folder_Id',
      'FACT_Template_DocId',
      FACTURA_FOLDER_ID
    );
    const folder = DriveApp.getFolderById(cfg.folderId);

    const emisor = (typeof getEmisorDesdeFactura_ === 'function')
      ? getEmisorDesdeFactura_()
      : { nombre: '', nif: '', direccion: '', cp: '', ciudad: '' };

    const fechaTxt = pdfFormatDate_(factRow.obj[pdfFindHeader_(factRow.headers, ['Fecha'])]);

    const clienteNombre = pdfPickValue_(factRow.obj, ['Cliente', 'Cliente_nombre', 'Nombre']);
    const clienteNif = pdfPickValue_(factRow.obj, ['NIF', 'Cliente_NIF']);
    const clienteDireccion = pdfPickValue_(factRow.obj, ['Direccion', 'Cliente_direccion']);
    const clienteCp = pdfPickValue_(factRow.obj, ['CP', 'Cliente_CP']);
    const clienteCiudad = pdfPickValue_(factRow.obj, ['Ciudad', 'Cliente_ciudad']);

    const totals = pdfComputeLineTotals_(lineas);
    const ivaTotal = pdfResolveIvaTotal_(factRow.obj, totals.base);
    const ivaPorc = pdfResolveIvaPorc_(factRow.obj, totals.base, ivaTotal);
    const total = pdfRound2_(totals.base + ivaTotal);

    const numeroFactura = pdfPickValue_(factRow.obj, ['Numero_factura', 'Factura_ID', 'ID']) || id;

    const copyName = pdfSafeFileName_('Factura_' + numeroFactura + '_' + (clienteNombre || 'Cliente'));
    const copy = DriveApp.getFileById(cfg.templateId).makeCopy(copyName, folder);
    const doc = DocumentApp.openById(copy.getId());

    const placeholders = {
      Numero_factura: numeroFactura,
      Fecha: fechaTxt,
      Cliente_nombre: clienteNombre,
      Cliente_NIF: clienteNif,
      Cliente_direccion: clienteDireccion,
      Cliente_CP: clienteCp,
      Cliente_ciudad: clienteCiudad,
      Base_imponible: pdfMoney2_(totals.base),
      IVA_PORC: pdfNumber_(ivaPorc),
      IVA_EUR: pdfMoney2_(ivaTotal),
      TOTAL: pdfMoney2_(total),
      Emisor_nombre: emisor.nombre || '',
      Emisor_NIF: emisor.nif || '',
      Emisor_direccion: emisor.direccion || '',
      Emisor_CP: emisor.cp || '',
      Emisor_ciudad: emisor.ciudad || ''
    };

    pdfReplaceTokens_(doc, placeholders);
    pdfFillLineTable_(doc, lineas, {
      LINEA_CONCEPTO: 'concepto',
      LINEA_CANTIDAD: 'cantidad',
      LINEA_PRECIO: 'precio',
      LINEA_SUBTOTAL: 'subtotal'
    });
    doc.saveAndClose();

    const pdfBlob = DriveApp.getFileById(copy.getId()).getAs(MimeType.PDF).setName(copyName + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);
    const pdfUrl = pdfFile.getUrl();

    pdfUpdateRowValues_(shFact, factRow, {
      Base: totals.base,
      IVA_total: ivaTotal,
      Total: total,
      PDF_link: pdfUrl
    });

    log('OK', '', { url: pdfUrl });
    return pdfUrl;
  } catch (err) {
    log('ERROR', err.message || String(err), { stack: err.stack });
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function pdfGetSheetByNames_(ss, names) {
  for (let i = 0; i < names.length; i++) {
    const sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  return null;
}

function pdfGetHeaderMap_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const map = {};
  headers.forEach((h, i) => { map[h] = i + 1; });
  return { headers, map };
}

function pdfFindHeader_(headers, candidates) {
  const lower = (headers || []).map((h) => String(h || '').toLowerCase());
  for (let i = 0; i < candidates.length; i++) {
    const idx = lower.indexOf(String(candidates[i] || '').toLowerCase());
    if (idx !== -1) return headers[idx];
  }
  return '';
}

function pdfFindRowById_(sh, idCandidates, id) {
  const headerInfo = pdfGetHeaderMap_(sh);
  const idHeader = pdfFindHeader_(headerInfo.headers, idCandidates);
  if (!idHeader) throw new Error('No existe columna ID en ' + sh.getName());

  const idCol = headerInfo.map[idHeader];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idCol - 1]).trim() === String(id).trim()) {
      const obj = {};
      headerInfo.headers.forEach((h, idx) => { obj[h] = data[i][idx]; });
      return { rowNumber: i + 2, headers: headerInfo.headers, headerMap: headerInfo.map, obj };
    }
  }
  return null;
}

function pdfPickValue_(obj, candidates) {
  if (!obj) return '';
  const keys = Object.keys(obj);
  for (let i = 0; i < candidates.length; i++) {
    const target = String(candidates[i] || '').toLowerCase();
    const match = keys.find((k) => String(k || '').toLowerCase() === target);
    if (match) return obj[match];
  }
  return '';
}

function pdfGetPresupuestoLineas_(ss, presId) {
  const sheetNames = ['PRES_LINEAS', 'LINEAS_PRES_HIST'];
  for (let i = 0; i < sheetNames.length; i++) {
    const sh = ss.getSheetByName(sheetNames[i]);
    if (!sh) continue;
    const lines = pdfGetLineasById_(sh, ['Pres_ID', 'Presupuesto_ID', 'ID'], presId);
    if (lines.length) return lines;
  }
  return [];
}

function pdfGetFacturaLineas_(ss, factRowObj, factId) {
  const sheetNames = ['FACT_LINEAS', 'LINEAS'];
  const numeroFactura = pdfPickValue_(factRowObj, ['Numero_factura', 'Factura_ID', 'ID']) || factId;
  for (let i = 0; i < sheetNames.length; i++) {
    const sh = ss.getSheetByName(sheetNames[i]);
    if (!sh) continue;
    const lines = pdfGetLineasById_(sh, ['Numero_factura', 'Factura_ID', 'ID'], numeroFactura);
    if (lines.length) return lines;
  }
  return [];
}

function pdfGetLineasById_(sh, idCandidates, idValue) {
  const headerInfo = pdfGetHeaderMap_(sh);
  const idHeader = pdfFindHeader_(headerInfo.headers, idCandidates);
  if (!idHeader) return [];
  const idCol = headerInfo.map[idHeader];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const lines = [];
  data.forEach((row) => {
    if (String(row[idCol - 1]).trim() !== String(idValue).trim()) return;
    const obj = {};
    headerInfo.headers.forEach((h, idx) => { obj[h] = row[idx]; });
    const concepto = pdfPickValue_(obj, ['Concepto', 'Descripcion', 'DescripciÃ³n']);
    const cantidad = pdfNumber_(pdfPickValue_(obj, ['Cantidad', 'Cant'])) || 0;
    const precio = pdfNumber_(pdfPickValue_(obj, ['Precio', 'Precio_unitario'])) || 0;
    const subtotalRaw = pdfPickValue_(obj, ['Subtotal', 'Importe', 'Base']);
    const subtotal = pdfRound2_(pdfNumber_(subtotalRaw) || (cantidad * precio));
    if (!concepto && !cantidad && !precio && !subtotal) return;
    lines.push({
      concepto: String(concepto || '').trim(),
      cantidad: pdfNumber_(cantidad),
      precio: pdfNumber_(precio),
      subtotal: pdfNumber_(subtotal),
      raw: obj
    });
  });
  return lines;
}

function pdfComputeLineTotals_(lineas) {
  let base = 0;
  let ivaTotal = 0;
  let ivaPorc = 0;
  let ivaFound = false;

  lineas.forEach((l) => {
    const subtotal = pdfNumber_(l.subtotal) || 0;
    base += subtotal;

    const iva = pdfNumber_(pdfPickValue_(l.raw, ['IVA_%', 'IVA_PORC', 'IVA'])) || 0;
    if (iva && !ivaFound) {
      ivaPorc = iva;
      ivaFound = true;
    }
    if (iva) ivaTotal += subtotal * (iva / 100);
  });

  return {
    base: pdfRound2_(base),
    ivaTotal: pdfRound2_(ivaTotal),
    ivaPorc: ivaPorc || 0
  };
}

function pdfResolveIvaTotal_(factRowObj, base) {
  const ivaTotal = pdfNumber_(pdfPickValue_(factRowObj, ['IVA_total', 'IVA_EUR', 'IVA'])) || 0;
  if (ivaTotal) return pdfRound2_(ivaTotal);
  const ivaPorc = pdfResolveIvaPorc_(factRowObj, base, 0);
  if (!base || !ivaPorc) return 0;
  return pdfRound2_(base * (ivaPorc / 100));
}

function pdfResolveIvaPorc_(factRowObj, base, ivaTotal) {
  const ivaPorc = pdfNumber_(pdfPickValue_(factRowObj, ['IVA_PORC', 'IVA_%'])) || 0;
  if (ivaPorc) return ivaPorc;
  if (!base || !ivaTotal) return 0;
  return pdfRound2_((ivaTotal / base) * 100);
}

function pdfReplaceTokens_(doc, map) {
  if (typeof replaceTokensEverywhere_ === 'function') {
    replaceTokensEverywhere_(doc, map);
    return;
  }
  pdfReplaceTokensEverywhere_(doc, map);
}

function pdfReplaceTokensEverywhere_(doc, map) {
  const containers = [];
  const body = doc.getBody();
  if (body) containers.push(body);
  try { const h = doc.getHeader(); if (h) containers.push(h); } catch (e) {}
  try { const f = doc.getFooter(); if (f) containers.push(f); } catch (e) {}

  const keys = Object.keys(map || {});
  containers.forEach((c) => {
    keys.forEach((key) => {
      const value = (map[key] === null || map[key] === undefined) ? '' : String(map[key]);
      const pattern = pdfBuildTokenPattern_(key);
      c.replaceText(pattern, value);
    });
  });
}

function pdfFillLineTable_(doc, lineas, fieldMap) {
  const tokens = Object.keys(fieldMap || {});
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
      if (tokens.some((k) => new RegExp(pdfBuildTokenPattern_(k), 'i').test(text))) {
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

  lineas.forEach((linea, idx) => {
    const row = templateRow.copy();
    const replacements = {
      LINEA_CONCEPTO: linea.concepto || '',
      LINEA_CANTIDAD: pdfFormatCantidad_(linea.cantidad),
      LINEA_PRECIO: pdfMoney2_(linea.precio),
      LINEA_SUBTOTAL: pdfMoney2_(linea.subtotal)
    };

    pdfReplaceTokensInRow_(row, replacements);
    table.insertTableRow(rowIndex + idx, row);
  });
}

function pdfReplaceTokensInRow_(row, map) {
  const cells = row.getNumCells();
  for (let c = 0; c < cells; c++) {
    const cell = row.getCell(c);
    const text = cell.editAsText();
    Object.keys(map).forEach((key) => {
      const pattern = pdfBuildTokenPattern_(key);
      text.replaceText(pattern, String(map[key] || ''));
    });
  }
}

function pdfEnsurePdfConfig_(folderName, folderProp, templateProp, defaultTemplateId, cfgFolderHeader, cfgTemplateHeader, defaultFolderId) {
  const props = PropertiesService.getScriptProperties();
  const configFolderId = pdfGetConfigValue_([cfgFolderHeader]);
  const configTemplateId = pdfGetConfigValue_([cfgTemplateHeader]);

  let folderId = configFolderId || props.getProperty(folderProp) || defaultFolderId || '';
  if (folderId) {
    try { DriveApp.getFolderById(folderId); } catch (e) { folderId = ''; }
  }
  if (!folderId) {
    const it = DriveApp.getFoldersByName(folderName);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
    folderId = folder.getId();
  }

  let templateId = configTemplateId || props.getProperty(templateProp) || defaultTemplateId;
  try { DriveApp.getFileById(templateId); } catch (e) {
    templateId = defaultTemplateId;
    DriveApp.getFileById(templateId);
  }

  props.setProperty(folderProp, folderId);
  props.setProperty(templateProp, templateId);
  pdfSetConfigValue_(cfgFolderHeader, folderId);
  pdfSetConfigValue_(cfgTemplateHeader, templateId);

  return { folderId, templateId };
}

function pdfGetConfigValue_(headers) {
  if (typeof getCfgAny_ === 'function') return getCfgAny_(headers);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('CONFIG');
  if (!cfg) return '';
  const head = cfg.getRange(1, 1, 1, Math.max(cfg.getLastColumn(), 1)).getDisplayValues()[0];
  for (let i = 0; i < headers.length; i++) {
    const idx = head.indexOf(headers[i]);
    if (idx !== -1) return String(cfg.getRange(2, idx + 1).getDisplayValue()).trim();
  }
  return '';
}

function pdfSetConfigValue_(header, value) {
  if (typeof setCfgValueIfSheet_ === 'function') {
    setCfgValueIfSheet_(header, value);
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('CONFIG');
  if (!cfg) return;
  const head = cfg.getRange(1, 1, 1, Math.max(cfg.getLastColumn(), 1)).getDisplayValues()[0];
  let col = head.indexOf(header) + 1;
  if (!col) {
    const lastCol = Math.max(cfg.getLastColumn(), 1);
    cfg.insertColumnsAfter(lastCol, 1);
    col = lastCol + 1;
    cfg.getRange(1, col).setValue(header);
  }
  cfg.getRange(2, col).setValue(value || '');
}

function pdfUpdateRowValues_(sh, rowInfo, values) {
  const headerMap = rowInfo.headerMap;
  Object.keys(values || {}).forEach((key) => {
    const col = headerMap[key];
    if (!col) return;
    sh.getRange(rowInfo.rowNumber, col).setValue(values[key]);
  });
}

function pdfShouldAppendNoIvaNote_(notas, doc, notasHeader) {
  const lower = String(notas || '').toLowerCase();
  if (lower.includes('precios no incluyen iva')) return false;
  if (notasHeader) return true;
  return pdfDocHasPlaceholder_(doc, 'NOTAS');
}

function pdfDocHasPlaceholder_(doc, key) {
  const body = doc.getBody();
  if (!body) return false;
  const pattern = pdfBuildTokenPattern_(key);
  return !!body.findText(pattern);
}

function pdfAppendNote_(notas, note) {
  if (!notas) return note;
  const trimmed = String(notas).trim();
  if (!trimmed) return note;
  return trimmed + ' ' + note;
}

function pdfFormatDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  if (!v) return '';
  const d = new Date(v);
  if (isNaN(d.getTime())) return String(v);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function pdfMoney2_(n) {
  const num = Number(n);
  if (isNaN(num)) return '';
  return Utilities.formatString('%.2f', num).replace('.', ',');
}

function pdfFormatCantidad_(value) {
  if (value === null || value === undefined || value === '') return '';
  const raw = String(value).replace(',', '.');
  const num = Number(raw);
  if (isNaN(num)) return String(value);
  if (Math.floor(num) === num) return String(num);
  return Utilities.formatString('%.2f', num).replace('.', ',');
}

function pdfNumber_(n) {
  if (n === null || n === undefined || n === '') return 0;
  const raw = String(n).replace(',', '.');
  const num = Number(raw);
  return isNaN(num) ? 0 : num;
}

function pdfRound2_(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

function pdfSafeFileName_(name) {
  return String(name || '')
    .replace(/[\\\/:*?"<>|#]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120) || 'Documento';
}

function pdfBuildTokenPattern_(key) {
  return '\\{\\{\\s*' + pdfCaseInsensitiveKeyPattern_(key) + '\\s*\\}\\}';
}

function pdfCaseInsensitiveKeyPattern_(key) {
  const raw = String(key || '');
  let out = '';
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i];
    if (/[a-zA-Z]/.test(ch)) {
      const lower = ch.toLowerCase();
      const upper = ch.toUpperCase();
      out += lower === upper ? pdfEscapeRegex_(ch) : `[${lower}${upper}]`;
    } else {
      out += pdfEscapeRegex_(ch);
    }
  }
  return out;
}

function pdfEscapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
