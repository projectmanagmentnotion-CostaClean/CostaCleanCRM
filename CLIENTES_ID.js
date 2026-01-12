/*************************************************
 * CLIENTES: ID AUTOGENERADO + INICIALIZADOR
 * Formato: CLI-AAAA-0001
 * CONFIG:
 *  D2 = CLI_Año
 *  E2 = CLI_Ultimo_numero
 *************************************************/

function onEdit__CLIENTES(e) {
  // 🔒 Evita el error si lo ejecutas desde el editor
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== 'CLIENTES') return;

  const row = e.range.getRow();
  if (row < 2) return;

  // Si ya tiene ID, no tocar
  const idCell = sh.getRange(row, 1); // Col A = Cliente_ID
  const currentId = String(idCell.getValue()).trim();
  if (currentId) return;

  // Condición mínima: Nombre + NIF o Email (para evitar IDs en filas vacías)
  const nombre = String(sh.getRange(row, 2).getValue()).trim(); // B
  const nif    = String(sh.getRange(row, 3).getValue()).trim(); // C
  const email  = String(sh.getRange(row, 8).getValue()).trim(); // H

  if (!nombre || (!nif && !email)) return;

  const newId = generarSiguienteClienteId_();
  idCell.setValue(newId);

  // Fecha_alta si está vacía (col K = 11 según tu captura)
  const fechaAltaCell = sh.getRange(row, 11);
  if (!fechaAltaCell.getValue()) fechaAltaCell.setValue(new Date());

  SpreadsheetApp.getActive().toast(`✅ Cliente_ID asignado: ${newId}`, 'CLIENTES', 3);
}

/**
 * EJECUTA ESTO 1 VEZ para enumerar clientes ya existentes
 * - Respeta IDs existentes (CL001, CLI-2025-0001, etc.)
 * - Asigna IDs solo a filas sin ID pero con datos
 * - Actualiza CONFIG (CLI_Ultimo_numero)
 */
function inicializarClientesIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('CLIENTES');
  const cfg = ss.getSheetByName('CONFIG');
  if (!sh || !cfg) throw new Error('Falta hoja CLIENTES o CONFIG.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const cliYear = String(cfg.getRange('D2').getValue()).trim() || String(new Date().getFullYear());
  let ultimo = Number(cfg.getRange('E2').getValue()) || 0;

  // Lee columnas A..K (hasta Fecha_alta)
  const data = sh.getRange(2, 1, lastRow - 1, 11).getValues();

  // 1) Detectar el máximo ya usado para no pisar nada
  data.forEach(r => {
    const id = String(r[0] || '').trim(); // A
    const m = id.match(/^CLI-(\d{4})-(\d{4})$/);
    if (m && m[1] === cliYear) {
      const n = Number(m[2]);
      if (!isNaN(n) && n > ultimo) ultimo = n;
    }
  });

  // 2) Asignar IDs faltantes
  let asignados = 0;

  for (let i = 0; i < data.length; i++) {
    const rowIndex = i + 2;
    const id = String(data[i][0] || '').trim();

    const nombre = String(data[i][1] || '').trim(); // B
    const nif    = String(data[i][2] || '').trim(); // C
    const email  = String(data[i][7] || '').trim(); // H

    // Si ya hay ID, saltar
    if (id) continue;

    // Si fila vacía, saltar
    if (!nombre || (!nif && !email)) continue;

    // Generar siguiente
    ultimo++;
    const newId = `CLI-${cliYear}-${String(ultimo).padStart(4, '0')}`;
    sh.getRange(rowIndex, 1).setValue(newId);

    // Fecha_alta si vacía
    const fechaAlta = sh.getRange(rowIndex, 11).getValue();
    if (!fechaAlta) sh.getRange(rowIndex, 11).setValue(new Date());

    asignados++;
  }

  // 3) Guardar contador
  cfg.getRange('D2').setValue(cliYear);
  cfg.getRange('E2').setValue(ultimo);

  SpreadsheetApp.getActive().toast(
    `✅ Inicialización lista. IDs asignados: ${asignados}. Último: ${ultimo}`,
    'CLIENTES',
    6
  );
}

/**
 * Genera el siguiente Cliente_ID y actualiza CONFIG
 */
function generarSiguienteClienteId_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('CONFIG');
  if (!cfg) throw new Error("No existe CONFIG.");

  const cliYear = String(cfg.getRange('D2').getValue()).trim() || String(new Date().getFullYear());
  let ultimo = Number(cfg.getRange('E2').getValue()) || 0;

  ultimo++;
  cfg.getRange('D2').setValue(cliYear);
  cfg.getRange('E2').setValue(ultimo);

  return `CLI-${cliYear}-${String(ultimo).padStart(4, '0')}`;
}