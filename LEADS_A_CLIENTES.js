/*************************************************
 * LEADS -> CLIENTES automático
 * Regla: si LEADS!Estado (col V) pasa a "Ganado"
 *   - crea cliente en CLIENTES si no existe por NIF o Email
 *   - escribe el Cliente_ID en LEADS col W
 *   - registra Fecha_alta y Origen
 *************************************************/

function onEdit_leads(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (String(sh.getName()).trim() !== 'LEADS') return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  const COL_ESTADO = 22; // V
  if (row < 2 || col !== COL_ESTADO) return;

  const estado = String(e.range.getDisplayValue() || '').trim().toLowerCase();
  if (estado !== 'ganado') return;

  console.log('LEADS: pasando a GANADO en fila', row);
  logEvent_(e.source, 'LEADS', 'ON_EDIT', 'LEAD', row, 'START', '', null);
  try {
    convertirLeadEnCliente_(e.source, row);
  } catch (err) {
    logEvent_(e.source, 'LEADS', 'CONVERT', 'CLIENTE', row, 'ERROR', err.message || String(err), null);
    throw err;
  }
}

function convertirLeadEnCliente_(ss, row) {
  // ss viene del trigger (e.source) => SIEMPRE es el spreadsheet correcto
  const shLeads = ss.getSheetByName('LEADS');
  const shCli = ss.getSheetByName('CLIENTES');
  if (!shLeads || !shCli) throw new Error('Falta hoja LEADS o CLIENTES');
  // LEADS: A..Z (26 cols)
  const lead = shLeads.getRange(row, 1, 1, 26).getValues()[0];

  const leadId  = String(lead[0] || '').trim();   // A
  const nombre  = String(lead[2] || '').trim();   // C
  const email   = String(lead[3] || '').trim().toLowerCase(); // D
  const tel     = String(lead[4] || '').trim();   // E
  const nif     = String(lead[5] || '').trim().toUpperCase(); // F
  const dir     = String(lead[6] || '').trim();   // G
  const cp      = String(lead[7] || '').trim();   // H
  const ciudad  = String(lead[8] || '').trim();   // I

  const COL_CLIENTE_ID = 23; // W
  const yaTieneCliente = String(lead[COL_CLIENTE_ID - 1] || '').trim();
  if (yaTieneCliente) {
    _safeToast_(`ℹ️ Lead ${leadId} ya tiene Cliente_ID: ${yaTieneCliente}`, 'LEADS → CLIENTE', 4);
    return;
  }

  // Validación mínima
  if (!nombre) {
    _safeToast_('⚠️ No puedo convertir: falta Nombre en el lead.', 'LEADS → CLIENTE', 6);
    return;
  }
  if (!nif && !email) {
    _safeToast_('⚠️ No puedo convertir: falta NIF o Email para evitar duplicados.', 'LEADS → CLIENTE', 6);
    return;
  }

  // Buscar duplicado en CLIENTES por NIF o Email
  const existente = buscarClienteExistente_(shCli, nif, email);
  let clienteId = existente;

  if (!clienteId) {
    // Crear nuevo Cliente_ID (usa tu CONFIG D2/E2)
    clienteId = generarSiguienteClienteId_(); // <- viene del script CLIENTES_ID.gs

    // Tipo_cliente (si no tienes, lo dejamos “Empresa” por defecto)
    const tipoCliente = 'Empresa';
    const origen = 'Lead (Ganado)';
    const fechaAlta = new Date();

    const targetRow = nextEmptyRow_(shCli, 1, 2); // col A, desde fila 2
shCli.getRange(targetRow, 1, 1, 11).setValues([[
  clienteId,   // A
  nombre,      // B
  nif,         // C
  dir,         // D
  cp,          // E
  ciudad,      // F
  tel,         // G
  email,       // H
  tipoCliente, // I
  origen,      // J
  fechaAlta    // K
]]);
  }
  if (existente) {
    logEvent_(ss, 'LEADS', 'CONVERT', 'CLIENTE', clienteId, 'SKIP', 'cliente ya existe', null);
  }

  // Escribir Cliente_ID en el lead (W) + fecha de conversión en Último_contacto (X)
  shLeads.getRange(row, 23).setValue(clienteId);     // W
  shLeads.getRange(row, 24).setValue(new Date());    // X (ultimo_contacto)
  logEvent_(ss, 'LEADS', 'CONVERT', 'CLIENTE', clienteId, 'OK', '', null);
  try {
    if (typeof presVincularPresupuestosPorLead_ === 'function') {
      presVincularPresupuestosPorLead_(ss, leadId, clienteId);
    }
  } catch (err) {
    logEvent_(ss, 'LEADS', 'LINK_PRES', 'PRESUPUESTO', leadId, 'ERROR', err.message || String(err), null);
  }
  _safeToast_(`✅ Lead convertido: ${leadId} → ${clienteId}`, 'LEADS → CLIENTE', 6);
}

function buscarClienteExistente_(shCli, nif, email) {
  const lastRow = shCli.getLastRow();
  if (lastRow < 2) return '';

  // CLIENTES: A..H mínimo (ID..Email) = 8 cols
  const data = shCli.getRange(2, 1, lastRow - 1, 8).getValues();
  const nifNorm = String(nif || '').trim().toUpperCase();
  const emailNorm = String(email || '').trim().toLowerCase();

  for (const r of data) {
    const id = String(r[0] || '').trim();               // A
    const nifC = String(r[2] || '').trim().toUpperCase(); // C
    const emailC = String(r[7] || '').trim().toLowerCase(); // H

    if (nifNorm && nifC && nifNorm === nifC) return id;
    if (emailNorm && emailC && emailNorm === emailC) return id;
  }
  return '';
}

function nextEmptyRow_(sh, keyCol = 1, startRow = 2) {
  const last = Math.max(sh.getLastRow(), startRow);
  const values = sh.getRange(startRow, keyCol, last - startRow + 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const v = String(values[i][0] ?? '').trim();
    if (v === '') return startRow + i; // primera fila vacía real en la col keyCol
  }
  return last + 1; // si no encontró huecos, continúa al final lógico
}






function _safeToast_(msg, title, secs){
  try { SpreadsheetApp.getActive().toast(msg, title || 'INFO', secs || 4); } catch(e) { console.log(String(title||'INFO')+': '+String(msg)); }
}


