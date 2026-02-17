var __CC_CRM_NUEVO = (function(){
/***************
 * COSTA CLEAN - SISTEMA FACTURAS (DEFINITIVO) + MENÚ CRM
 * HOJAS:
 *  - FACTURA
 *  - LINEAS (A: Numero_factura, B: Concepto, C: Cantidad, D: Precio, E: Subtotal)
 *  - CONFIG (B1=Año, B2=Último número emitido, B4 SheetID leads, B5 Tab respuestas, B6 Hoja destino LEADS)
 *  - HISTORIAL (A:Numero, B:Fecha, C:Cliente, D:Base, E:IVA, F:Total, G:PDF link)
 *  - LINEAS_HIST (se crea sola si no existe)
 ***************/

const ID_PLANTILLA_DOC = '1OU_1CxZBEc46OP5W1d98aI1LyllO0UV7Tln7qShH6as';
const ID_CARPETA_PDF  = '1l11q4NpNNT_W_jZ8Mgw0AVv2VN_PFsGp';

const LINEAS_PRECREADAS = 5; // cuántas filas crea el botón "AÑADIR LÍNEA(S)"

// === Formato dinero (España) a 2 decimales ===
var ccMoney2_ = function(n) {
  const num = Number(n);
  if (isNaN(num)) return '';
  // 1234.5 -> "1234,50"
  return Utilities.formatString('%.2f', num).replace('.', ',');
}


/***************
 * MENU (ÚNICO)
 ***************/
var onOpenMain_ = function() {
  SpreadsheetApp.getUi()
    .createMenu('Costa Clean')
    .addItem('➕ Añadir líneas (siguiente factura)', 'workflowAnadirLineasSiguiente')
    .addItem('🧾 Generar factura PDF', 'workflowGenerarFactura')
    .addSeparator()
    .addItem('Crear factura del presupuesto activo', 'uiCrearFacturaDesdePresupuestoActivo_')
    .addItem('Generar PDF presupuesto activo', 'uiGenerarPdfPresupuestoActivo_')
    .addItem('Generar PDF factura activa', 'uiGenerarPdfFacturaActiva_')
    .addSeparator()
    .addItem('Cargar datos de prueba (seed)', 'seedSampleData_')
    .addToUi();

  // Menú CRM (definido en CRM_LEADS.gs)
  if (typeof menuCRM_ === 'function') menuCRM_();
  if (typeof menuPresupuestos_ === 'function') menuPresupuestos_();

}

/***************
 * CONFIG
 ***************/
var getConfig_ = function() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('CONFIG');
  if (!cfg) throw new Error("No existe la hoja 'CONFIG'.");

  const anio = String(cfg.getRange('B1').getValue()).trim();
  const ultimo = Number(cfg.getRange('B2').getValue());

  if (!anio) throw new Error("CONFIG!B1 (Año) está vacío.");
  if (isNaN(ultimo)) throw new Error("CONFIG!B2 (Último número) no es un número.");

  return { cfg, anio, ultimo };
}

// Ver siguiente sin consumir
var peekSiguienteNumero_ = function() {
  const { anio, ultimo } = getConfig_();
  const siguiente = ultimo + 1;
  return anio + '-' + String(siguiente).padStart(3, '0');
}

// Consumir siguiente (incrementa CONFIG)
var consumirSiguienteNumero_ = function() {
  const { cfg, anio, ultimo } = getConfig_();
  const siguiente = ultimo + 1;
  cfg.getRange('B2').setValue(siguiente);
  return anio + '-' + String(siguiente).padStart(3, '0');
}

/***************
 * BOTÓN 1: AÑADIR LÍNEAS
 ***************/
var workflowAnadirLineasSiguiente = function() {
  asegurarFormulaSubtotal_();
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const factura = ss.getSheetByName('FACTURA');
  const lineas  = ss.getSheetByName('LINEAS');
  if (!factura || !lineas) throw new Error("Falta hoja FACTURA o LINEAS.");

  const numero = peekSiguienteNumero_();

  // Mostrar número “a usar” en FACTURA
  factura.getRange('B8').setValue(numero);
  factura.getRange('B9').setValue(new Date());

  // Buscar la siguiente fila libre real (mirando A:D)
  const lastRow = Math.max(lineas.getLastRow(), 2);
  let nextRow = 2;

  if (lastRow > 2) {
    const data = lineas.getRange(2, 1, lastRow - 1, 4).getValues(); // A:D
    for (let i = data.length - 1; i >= 0; i--) {
      const rowHasData = data[i].some(v => String(v).trim() !== '');
      if (rowHasData) { nextRow = i + 3; break; }
    }
  }

  // Crear filas
  for (let i = 0; i < LINEAS_PRECREADAS; i++) {
    const r = nextRow + i;
    lineas.getRange(r, 1).setValue(numero); // A Numero
    lineas.getRange(r, 2).clearContent();   // B Concepto
    lineas.getRange(r, 3).clearContent();   // C Cantidad
    lineas.getRange(r, 4).clearContent();   // D Precio
    // E se calcula por ARRAYFORMULA en E2 (no tocar)
  }

  lineas.setActiveSelection(lineas.getRange(nextRow, 2));

  ui.alert(
    '✅ Líneas creadas para ' + numero +
    '\n\nRellena en LINEAS:' +
    '\nB: Concepto' +
    '\nC: Cantidad' +
    '\nD: Precio'
  );
}

/***************
 * LEER LINEAS DE UNA FACTURA
 ***************/
var obtenerLineasFactura_ = function(sheetLineas, numeroFactura) {
  const lastRow = sheetLineas.getLastRow();
  if (lastRow < 2) return [];

  const data = sheetLineas.getRange(2, 1, lastRow - 1, 5).getValues(); // A:E
  return data
    .map((r, idx) => ({ row: idx + 2, vals: r }))
    .filter(o => String(o.vals[0]).trim() === String(numeroFactura).trim());
}
var lineasReales_ = function(objs) {
  return objs.filter(o => {
    const r = o.vals;
    return String(r[1]).trim() || String(r[2]).trim() || String(r[3]).trim();
  });
}

/***************
 * ARCHIVAR Y BORRAR LINEAS
 ***************/
var archivarYLimpiarLineas_ = function(numeroFactura) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lineas = ss.getSheetByName('LINEAS');
  if (!lineas) throw new Error("No existe la hoja LINEAS.");

  let hist = ss.getSheetByName('LINEAS_HIST');
  if (!hist) {
    hist = ss.insertSheet('LINEAS_HIST');
    hist.getRange(1,1,1,6).setValues([['Numero_factura','Concepto','Cantidad','Precio','Subtotal','Archivado_el']]);
  }

  const lastRow = lineas.getLastRow();
  if (lastRow < 2) return;

  const data = lineas.getRange(2,1,lastRow-1,5).getValues(); // A:E
  const rowsToArchive = [];
  const rowsToClear = [];

  data.forEach((r, idx) => {
    if (String(r[0]).trim() === String(numeroFactura).trim()) {
      const tieneAlgo = String(r[1]).trim() || String(r[2]).trim() || String(r[3]).trim();
      if (tieneAlgo) rowsToArchive.push([r[0], r[1], r[2], r[3], r[4], new Date()]);
      rowsToClear.push(idx + 2);
    }
  });

  if (rowsToArchive.length) {
    const start = hist.getLastRow() + 1;
    hist.getRange(start,1,rowsToArchive.length,6).setValues(rowsToArchive);
  }

  // Limpia SOLO A:D (no tocamos E porque E2 tiene ArrayFormula)
  rowsToClear.forEach(r => {
    lineas.getRange(r, 1, 1, 4).clearContent();
  });
}

/***************
 * BOTÓN 2: GENERAR FACTURA PDF
 ***************/
var workflowGenerarFactura = function() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const factura = ss.getSheetByName('FACTURA');
  const lineas = ss.getSheetByName('LINEAS');
  const hist = ss.getSheetByName('HISTORIAL');

  if (!factura || !lineas || !hist) throw new Error("Falta hoja FACTURA, LINEAS o HISTORIAL.");

  const numeroEnFactura = String(factura.getRange('B8').getDisplayValue()).trim();
  const peek = peekSiguienteNumero_();

  let numeroFactura = numeroEnFactura;

  if (!numeroFactura) {
    numeroFactura = consumirSiguienteNumero_();
    factura.getRange('B8').setValue(numeroFactura);
  } else {
    if (numeroFactura === peek) consumirSiguienteNumero_();
  }

  factura.getRange('B9').setValue(new Date());

  const objs = obtenerLineasFactura_(lineas, numeroFactura);
  const reales = lineasReales_(objs);

  if (reales.length === 0) {
    ui.alert(
      '⚠️ No hay líneas en LINEAS para ' + numeroFactura +
      '\n\nPulsa primero "➕ Añadir líneas" y rellena Concepto/Cantidad/Precio.'
    );
    return;
  }

  // Base usando la ArrayFormula de E (C*D). Si E está en blanco, calcula igual.
  let base = 0;
  reales.forEach(o => {
    const concepto = String(o.vals[1] || '').trim();
    const cant = Number(o.vals[2]) || 0;
    const precio = Number(o.vals[3]) || 0;
    const subtotal = (concepto || cant || precio) ? (cant * precio) : 0;
    base += subtotal;
  });

  const ivaPorc = Number(factura.getRange('B11').getValue()) || 0;
  const iva = base * (ivaPorc / 100);
  const total = base + iva;

  factura.getRange('B10').setValue(base);
  factura.getRange('B12').setValue(iva);
  factura.getRange('B13').setValue(total);

  const datos = {
    Numero_factura: factura.getRange('B8').getDisplayValue(),
    Fecha: factura.getRange('B9').getDisplayValue(),

    Emisor_nombre: factura.getRange('B15').getValue(),
    Emisor_NIF: factura.getRange('B16').getValue(),
    Emisor_direccion: factura.getRange('B17').getValue(),
    Emisor_CP: factura.getRange('B18').getDisplayValue(),
    Emisor_ciudad: factura.getRange('B19').getValue(),

    Cliente_nombre: factura.getRange('B2').getValue(),
    Cliente_NIF: factura.getRange('B3').getValue(),
    Cliente_direccion: factura.getRange('B4').getValue(),
    Cliente_CP: factura.getRange('B5').getDisplayValue(),
    Cliente_ciudad: factura.getRange('B6').getValue(),

    Base_imponible: ccMoney2_(base),
    IVA_PORC: factura.getRange('B11').getDisplayValue(),
    IVA_EUR: ccMoney2_(iva),
    TOTAL: ccMoney2_(total)

  };

  const copia = DriveApp.getFileById(ID_PLANTILLA_DOC).makeCopy('Factura ' + datos.Numero_factura);
  const doc = DocumentApp.openById(copia.getId());
  const body = doc.getBody();

  for (const k in datos) body.replaceText('{{' + k + '}}', String(datos[k]));

  let tabla = null;
  const tables = body.getTables();
  for (const t of tables) {
    if (t.getText().includes('{{LINEA_CONCEPTO}}')) { tabla = t; break; }
  }
  if (!tabla) throw new Error("No encuentro en el Doc la fila molde con {{LINEA_CONCEPTO}}.");

  const filaMolde = tabla.getRow(1).copy();
  tabla.removeRow(1);

  reales.forEach((o, idx) => {
    const r = o.vals;
    const cant = r[2];
    const precio = r[3];
    const subtotal = (Number(cant) || 0) * (Number(precio) || 0);

    const nueva = filaMolde.copy();
    nueva.getCell(0).setText(String(r[1]));        // Concepto
    nueva.getCell(1).setText(String(cant || ''));  // Cantidad
    nueva.getCell(2).setText(ccMoney2_(precio));    // Precio (2 decimales)
    nueva.getCell(3).setText(ccMoney2_(subtotal));  // Subtotal (2 decimales)

    tabla.insertTableRow(1 + idx, nueva);
  });

  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(copia.getId()).getAs('application/pdf');
  const pdfFile = DriveApp.getFolderById(ID_CARPETA_PDF).createFile(pdfBlob);
  const pdfUrl = pdfFile.getUrl();
  const clienteId = String(factura.getRange('B1').getDisplayValue()).trim();

hist.appendRow([
  datos.Numero_factura,  // A Numero_factura
  datos.Fecha,           // B Fecha
  clienteId,             // C Cliente_ID
  datos.Cliente_nombre,  // D Cliente
  datos.Cliente_NIF,     // E NIF
  datos.Base_imponible,  // F Base
  datos.IVA_EUR,         // G IVA
  datos.TOTAL,           // H Total
  pdfUrl                 // I PDF_link
]);


  archivarYLimpiarLineas_(numeroFactura);

  ui.alert('✅ Factura emitida: ' + datos.Numero_factura + '\n📄 PDF: ' + pdfUrl);
}
var asegurarFormulaSubtotal_ = function() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('LINEAS');
  if (!sh) throw new Error("No existe la hoja LINEAS.");

  const celda = sh.getRange('E2');
  const formulaActual = celda.getFormula();

  if (!formulaActual) {
    celda.setFormula('=ARRAYFORMULA(SI((C2:C<>"")*(D2:D<>"");C2:C*D2:D;""))');
  }
}



  // Exports (solo lo que se usa fuera)
  return {
    onOpenMain_: onOpenMain_,
    consumirSiguienteNumero_: consumirSiguienteNumero_,
    workflowAnadirLineasSiguiente: workflowAnadirLineasSiguiente,
    workflowGenerarFactura: workflowGenerarFactura
  };
})();


// Entry-points visibles (solo estos deben quedar como 'function' top-level)
function onOpenMain_(e){ return __CC_CRM_NUEVO.onOpenMain_(e); }
function consumirSiguienteNumero_(){ return __CC_CRM_NUEVO.consumirSiguienteNumero_(); }
function workflowAnadirLineasSiguiente(){ return __CC_CRM_NUEVO.workflowAnadirLineasSiguiente(); }
function workflowGenerarFactura(){ return __CC_CRM_NUEVO.workflowGenerarFactura(); }

