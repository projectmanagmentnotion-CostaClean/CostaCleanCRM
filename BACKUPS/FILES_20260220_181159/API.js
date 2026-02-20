const SS_ID = "1m62QB04_aDrxeXjSiK6QHrdAztGPqTAhx5zA9cKa8kk";
var __CC_API = (function(){
/***************
 * API — CostaClean CRM (Sheets)
 * Basado en tus pestañas reales:
 *  - CLIENTES
 *  - LEADS
 *  - HISTORIAL (facturas emitidas)
 *  - HISTORIAL_PRESUPUESTOS (proformas/presupuestos emitidos)
 *  - (opcional) GASTOS
 *
 * IMPORTANTE:
 * 1) Pega el ID REAL del Spreadsheet en SS_ID.
 * 2) Si algún nombre de hoja no coincide EXACTO, cámbialo en SHEETS.
 ***************/

const SHEETS = {
  clientes: 'CLIENTES',
  leads: 'LEADS',
  facturas: 'HISTORIAL',
  proformas: 'HISTORIAL_PRESUPUESTOS',
  gastos: 'GASTOS' // si no existe, déjalo pero no lo uses o cámbialo
};

var _ss = function() {
  return SpreadsheetApp.openById(SS_ID);
}

var _sh = function(name) {
  const sh = _ss().getSheetByName(name);
  if (!sh) throw new Error('No existe la hoja: ' + name);
  return sh;
}

/**
 * Lee toda la hoja como array de objetos:
 * - Primera fila = headers
 * - Resto = rows
 * - Convierte fechas a string ISO simple si hace falta (opcional)
 */
var _getDataWithHeaders = function(sheetName) {
  const sh = _sh(sheetName);
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const headers = values.shift().map(h => String(h).trim());

  return values
    .filter(r => r.some(c => c !== '' && c !== null)) // ignora filas vacías
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        let v = r[i];

        // Normaliza fechas (para que el frontend no reciba objetos raros)
        if (v instanceof Date && !isNaN(v.getTime())) {
          // YYYY-MM-DD
          v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }

        obj[h] = v;
      });
      return obj;
    });
}

/** Helpers de búsqueda */
var _includesQuery = function(rowObj, qLower) {
  // búsqueda simple full-text
  return JSON.stringify(rowObj).toLowerCase().includes(qLower);
}

/** ====== 1) Dashboard KPIs (por ahora placeholder REALISTA, luego lo calculamos) ====== */
var apiDashboardLegacy_ = function() {
  const tz = Session.getScriptTimeZone();

  const shFact = _sh(SHEETS.facturas); // HISTORIAL
  const fact = _getDataWithHeaders(SHEETS.facturas);

  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth(); // 0-11

  // Mes actual: sumar Total
  const ventasMes = fact
    .filter(r => r.Fecha)
    .filter(r => {
      const d = new Date(r.Fecha);
      return d.getFullYear() === y && d.getMonth() === m;
    })
    .reduce((acc, r) => acc + (Number(r.Total) || 0), 0);

  // Trimestre actual: IVA repercutido (sum IVA en facturas)
  const q = Math.floor(m / 3) + 1; // 1..4
  const qStartMonth = (q - 1) * 3;

  const ivaRepercutido = fact
    .filter(r => r.Fecha)
    .filter(r => {
      const d = new Date(r.Fecha);
      return d.getFullYear() === y && d.getMonth() >= qStartMonth && d.getMonth() <= qStartMonth + 2;
    })
    .reduce((acc, r) => acc + (Number(r.IVA) || 0), 0);

  // (Opcional) IVA soportado desde GASTOS si existe
  let ivaSoportado = 0;
  try {
    const gastos = _getDataWithHeaders(SHEETS.gastos);
    ivaSoportado = gastos
      .filter(r => r.Fecha)
      .filter(r => {
        const d = new Date(r.Fecha);
        return d.getFullYear() === y && d.getMonth() >= qStartMonth && d.getMonth() <= qStartMonth + 2;
      })
      .reduce((acc, r) => acc + (Number(r.IVA) || 0), 0);
  } catch(e){ /* si no existe hoja, queda 0 */ }

  const ivaNetoTrimestre = ivaRepercutido - ivaSoportado;

  // Pendientes: si no tienes columna Estado/Pagado aún, lo dejamos 0 por ahora
  const facturasPendientes = 0;

  return {
    ventasMes: round2(ventasMes),
    facturasPendientes,
    ivaNetoTrimestre: round2(ivaNetoTrimestre),
    gastoDeducibleTrimestre: 0,
    periodo: { mes: Utilities.formatDate(now, tz, 'MMM'), trimestre: 'Q' + q + ' ' + y }
  };
}

var round2 = function(n){ return Math.round((Number(n)||0)*100)/100; }


/** ====== 2) Listas ====== */
var legacy_apiList = function(entity, params) {
  const sheetName = SHEETS[entity];
  if (!sheetName) throw new Error('Entidad no soportada: ' + entity);

  const data = _getDataWithHeaders(sheetName);

  const q = (params && params.q ? String(params.q).toLowerCase().trim() : '');
  const filtered = q ? data.filter(row => _includesQuery(row, q)) : data;

  const limit = params && params.limit ? Math.max(1, Number(params.limit)) : 30;
  return filtered.slice(0, limit);
}

/** ====== 3) Detalle por ID ======
 * Mapea la columna ID real según tus hojas.
 * Ajusta estas keys si cambian los headers:
 * - CLIENTES: Cliente_ID
 * - LEADS: Lead_ID
 * - HISTORIAL (facturas): Numero_factura (según tu captura)
 * - HISTORIAL_PRESUPUESTOS (proformas): Pres_ID
 * - GASTOS: Gasto_ID (si lo tienes)
 */
var legacy_apiGet = function(entity, id) {
  const sheetName = SHEETS[entity];
  if (!sheetName) throw new Error('Entidad no soportada: ' + entity);

  const data = _getDataWithHeaders(sheetName);

  const ID_KEYS = {
    clientes: 'Cliente_ID',
    leads: 'Lead_ID',
    facturas: 'Numero_factura',
    proformas: 'Pres_ID',
    gastos: 'Gasto_ID'
  };

  const key = ID_KEYS[entity];
  if (!key) throw new Error('No hay ID key definida para: ' + entity);

  const found = data.find(r => String(r[key]).trim() === String(id).trim());
  if (!found) throw new Error('No encontrado: ' + entity + ' ' + id);
  return found;
}

/** ====== 4) Healthcheck / debug rápido (opcional) ======
 * Útil para confirmar que el Spreadsheet ID y hojas existen.
 */
var apiHealth = function() {
  const ss = _ss();
  const sheets = ss.getSheets().map(s => s.getName());
  return {
    ok: true,
    spreadsheetId: SS_ID,
    availableSheets: sheets,
    mapped: SHEETS
  };
}


  // Exports (solo lo que se usa fuera)
  return {
    apiHealth: apiHealth
  };
})();


// Entry-point visible (único top-level permitido)
function apiHealth(){ return __CC_API.apiHealth(); }




//
// ===== DIAGNÓSTICO TEMPORAL (no borrar hasta resolver) =====
function diagOpenById(){
  var out = { ssId: SS_ID };

  // 1) Drive visibility: si esto falla -> no tienes acceso real al archivo por ID (permiso/propietario/cuenta)
  try{
    var f = DriveApp.getFileById(SS_ID);
    out.drive_ok = true;
    out.drive_name = f.getName();
    out.drive_owner = (f.getOwner && f.getOwner()) ? f.getOwner().getEmail() : null;
  }catch(e){
    out.drive_ok = false;
    out.drive_error = String(e && e.message ? e.message : e);
  }

  // 2) SpreadsheetApp: si Drive_OK pero esto falla -> problema específico SpreadsheetApp/openById o scopes
  try{
    var ss = SpreadsheetApp.openById(SS_ID);
    out.sheets_ok = true;
    out.ss_name = ss.getName();
    out.sheet_names = ss.getSheets().map(function(s){ return s.getName(); });
  }catch(e2){
    out.sheets_ok = false;
    out.sheets_error = String(e2 && e2.message ? e2.message : e2);
  }

  // 3) Tipos (para detectar runtime raro)
  out.typeof_SpreadsheetApp = typeof SpreadsheetApp;
  out.typeof_DriveApp = typeof DriveApp;

  return out;
}




