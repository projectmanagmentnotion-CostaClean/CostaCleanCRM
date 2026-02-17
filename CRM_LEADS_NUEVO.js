var __CC_CRM_LEADS = (function(){
/***************
 * CRM - LEADS (COSTA CLEAN) - DEFINITIVO
 * Importa nuevas respuestas del Google Form hacia la hoja LEADS del CRM
 * - Evita duplicados por Email o Teléfono o RowKey
 * - Genera Lead_ID incremental L0001...
 * - Setea Estado="Nuevo" y Origen="Formulario Web"
 *
 * Requiere CONFIG:
 *  B4 = Leads_Sheet_ID (ID del spreadsheet de respuestas)
 *  B5 = Leads_Tab_Name (nombre pestaña en ese spreadsheet)
 *  B6 = Nombre hoja destino en CRM (ej: "LEADS")
 ***************/

var menuCRM_ = function() {
  SpreadsheetApp.getUi()
    .createMenu('CRM')
    .addItem('📥 Importar nuevos leads (Form → LEADS)', 'importarNuevosLeads_')
    .addToUi();
}

var importarNuevosLeads_ = function() {
  const ui = SpreadsheetApp.getUi();
  const ssCRM = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ssCRM.getSheetByName('CONFIG');
  if (!cfg) throw new Error("No existe la hoja CONFIG.");

  const leadsSheetId = String(cfg.getRange('B4').getValue()).trim();
  const leadsTabName = String(cfg.getRange('B5').getValue()).trim();
  const destinoName  = String(cfg.getRange('B6').getValue()).trim();

  if (!leadsSheetId || !leadsTabName || !destinoName) {
    throw new Error("CONFIG incompleta. Revisa B4 (Sheet ID), B5 (Tab Name), B6 (Destino).");
  }

  const shLeads = ssCRM.getSheetByName(destinoName);
  if (!shLeads) throw new Error(`No existe la hoja destino "${destinoName}" en el CRM.`);

  asegurarCabeceraLeads_(shLeads);

  // --- 1) Leer respuestas del Form (origen) ---
  const ssSrc = SpreadsheetApp.openById(leadsSheetId);
  const shSrc = ssSrc.getSheetByName(leadsTabName);
  if (!shSrc) throw new Error(`No existe la pestaña "${leadsTabName}" en el spreadsheet de respuestas.`);

  const lastRowSrc = shSrc.getLastRow();
  const lastColSrc = shSrc.getLastColumn();

  if (lastRowSrc < 2) {
    ui.alert("No hay respuestas nuevas (el origen está vacío).");
    return;
  }

  const headersSrc = shSrc.getRange(1, 1, 1, lastColSrc).getValues()[0].map(h => String(h).trim());
  const rowsSrc = shSrc.getRange(2, 1, lastRowSrc - 1, lastColSrc).getValues();

  // Mapa: header -> índice
  const idx = {};
  headersSrc.forEach((h, i) => idx[h] = i);

  // --- 2) Leer existentes en LEADS para deduplicar ---
  const lastRowLeads = shLeads.getLastRow();
  const existingKeys = new Set();
  const existingEmail = new Set();
  const existingTel = new Set();

  if (lastRowLeads >= 2) {
    // A:Z => 26
    const dataLeads = shLeads.getRange(2, 1, lastRowLeads - 1, 26).getValues();
    dataLeads.forEach(r => {
      const email = String(r[3] || '').trim().toLowerCase(); // D
      const tel   = normalizarTelefono_(r[4]);              // E
      const rowKey = String(r[25] || '').trim();            // Z

      if (rowKey) existingKeys.add(rowKey);
      if (email) existingEmail.add(email);
      if (tel) existingTel.add(tel);
    });
  }
  // --- 2.5) Preparar contador incremental de Lead_ID (PRO) ---
  let nextLeadNum = 1;
  const lastRow = shLeads.getLastRow();
  if (lastRow >= 2) {
    const lastId = String(shLeads.getRange(lastRow, 1).getValue()).trim(); // Col A
    const n = Number(lastId.replace(/[^\d]/g, ''));
    if (!isNaN(n) && n > 0) nextLeadNum = n + 1;
  }

  // --- 3) Construir nuevos leads ---
  const nuevos = [];
  let saltados = 0;

  rowsSrc.forEach((r) => {
    const ts = pick_(r, idx, ['Marca temporal', 'Timestamp', 'Fecha', 'Fecha y hora'], 0);
    const nombre = pick_(r, idx, ['Nombre completo', 'Nombre', 'Nombre y apellidos'], '');
    const email  = String(pick_(r, idx, ['Correo electrónico', 'Dirección de correo electrónico', 'Email', 'Correo'], '')).trim().toLowerCase();
    const tel    = normalizarTelefono_(pick_(r, idx, ['Teléfono de contacto (con prefijo)', 'Teléfono', 'Telefono', 'Móvil', 'Movil'], ''));

    const rowKey = construirRowKey_(ts, email, tel);

    if (existingKeys.has(rowKey) || (email && existingEmail.has(email)) || (tel && existingTel.has(tel))) {
      saltados++;
      return;
    }

    const leadId = 'L' + String(nextLeadNum++).padStart(4, '0');
    const fechaEntrada = (ts instanceof Date) ? ts : String(ts || '');

    const tipoServicio   = pick_(r, idx, ['¿Qué tipo de servicio necesitas?', 'Tipo de servicio'], '');
    const tipoPropiedad  = pick_(r, idx, ['¿Qué tipo de propiedad es?', 'Tipo de propiedad'], '');
    const m2             = pick_(r, idx, ['¿Cuántos metros cuadrados tiene aproximadamente?', 'Metros cuadrados', 'm2'], '');
    const hab            = pick_(r, idx, ['¿Cuántas habitaciones tiene?', 'Habitaciones'], '');
    const banos          = pick_(r, idx, ['¿Cuántos baños tiene?', 'Baños', 'Banos'], '');
    const terraza        = pick_(r, idx, ['¿Tiene terraza, balcón o zonas exteriores?', 'Terraza'], '');
    const mascotas       = pick_(r, idx, ['¿Hay mascotas en el lugar?', 'Mascotas'], '');
    const fechaServicio  = pick_(r, idx, ['¿Para qué fecha necesitas el servicio?', 'Fecha servicio'], '');
    const horaPreferida  = pick_(r, idx, ['¿A qué hora puede realizarse la limpieza?', 'Hora preferida'], '');
    const frecuencia     = pick_(r, idx, ['¿Con qué frecuencia necesitas el servicio?', 'Frecuencia'], '');
    const canalPreferido = pick_(r, idx, ['¿Cómo prefieres recibir tu presupuesto?', 'Canal preferido'], '');
    const mensaje        = pick_(r, idx, ['¿Podrías contarnos brevemente qué necesitas?', 'Mensaje', 'Notas', 'Observaciones'], '');

    const nif  = pick_(r, idx, ['NIF/CIF', 'NIF', 'CIF'], '');
    const dir  = pick_(r, idx, ['Dirección', 'Direccion'], '');
    const cp   = pick_(r, idx, ['Código Postal', 'Codigo Postal', 'CP'], '');
    const pob  = pick_(r, idx, ['Población', 'Poblacion', 'Ciudad'], '');

    nuevos.push([
      leadId,           // A Lead_ID
      fechaEntrada,     // B Fecha_entrada
      nombre,           // C Nombre
      email,            // D Email
      tel,              // E Teléfono
      nif,              // F NIF/CIF
      dir,              // G Dirección
      cp,               // H CP
      pob,              // I Población
      tipoServicio,     // J Tipo_servicio
      tipoPropiedad,    // K Tipo_propiedad
      m2,               // L m2
      hab,              // M Habitaciones
      banos,            // N Baños
      terraza,          // O Terraza
      mascotas,         // P Mascotas
      fechaServicio,    // Q Fecha_servicio
      horaPreferida,    // R Hora_preferida
      frecuencia,       // S Frecuencia
      canalPreferido,   // T Canal_preferido
      mensaje,          // U Mensaje/Notas
      "Nuevo",          // V Estado
      "",               // W Cliente_ID
      "",               // X Último_contacto
      "Formulario Web", // Y Origen
      rowKey            // Z RowKey
    ]);

    existingKeys.add(rowKey);
    if (email) existingEmail.add(email);
    if (tel) existingTel.add(tel);
  });

  if (nuevos.length === 0) {
    ui.alert(`No hay leads nuevos. Duplicados saltados: ${saltados}`);
    return;
  }

  const startRow = shLeads.getLastRow() + 1;
  shLeads.getRange(startRow, 1, nuevos.length, 26).setValues(nuevos);

  ui.alert(`✅ Importación lista.\nNuevos leads: ${nuevos.length}\nDuplicados saltados: ${saltados}`);
}

/***************
 * Helpers
 ***************/
var asegurarCabeceraLeads_ = function(shLeads) {
  const headers = [
    'Lead_ID','Fecha_entrada','Nombre','Email','Teléfono','NIF/CIF','Dirección','CP','Población',
    'Tipo_servicio','Tipo_propiedad','m2','Habitaciones','Baños','Terraza','Mascotas',
    'Fecha_servicio','Hora_preferida','Frecuencia','Canal_preferido','Mensaje/Notas',
    'Estado','Cliente_ID','Último_contacto','Origen','RowKey'
  ];
  const firstRow = shLeads.getRange(1,1,1,26).getValues()[0];
  const empty = firstRow.every(v => String(v || '').trim() === '');
  if (empty) shLeads.getRange(1,1,1,26).setValues([headers]);
}

var pick_ = function(row, idx, names, fallback) {
  for (const n of names) {
    if (Object.prototype.hasOwnProperty.call(idx, n)) {
      return row[idx[n]];
    }
  }
  // fallback por índice numérico si te pasé 0 arriba
  if (typeof fallback === 'number') return row[fallback];
  return fallback;
}

var normalizarTelefono_ = function(tel) {
  const t = String(tel || '').trim();
  if (!t) return '';
  return t.replace(/[^\d+]/g, '');
}

var construirRowKey_ = function(ts, email, tel) {
  const tsStr = (ts instanceof Date)
    ? Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss")
    : String(ts || '');
  return [tsStr, email || '', tel || ''].join('|').trim();
}

var generarLeadId_ = function(shLeads) {
  const lastRow = shLeads.getLastRow();
  if (lastRow < 2) return 'L0001';

  const lastId = String(shLeads.getRange(lastRow, 1).getValue()).trim(); // A
  const n = Number(lastId.replace(/[^\d]/g, ''));
  const next = (isNaN(n) ? 1 : n + 1);
  return 'L' + String(next).padStart(4, '0');
}

  // Exports (solo lo que se usa fuera)
  return {
    menuCRM_: menuCRM_,
    importarNuevosLeads_: importarNuevosLeads_,
    generarLeadId_: generarLeadId_
  };
})();


// Entry-points visibles (solo estos deben quedar como 'function' top-level)
function menuCRM_(){ return __CC_CRM_LEADS.menuCRM_(); }
function importarNuevosLeads_(){ return __CC_CRM_LEADS.importarNuevosLeads_(); }
function generarLeadId_(shLeads){ return __CC_CRM_LEADS.generarLeadId_(shLeads); }


