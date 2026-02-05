function __logJson_(label, obj){
  try{
    Logger.log(label + ': ' + JSON.stringify(obj, null, 2));
  }catch(e){
    Logger.log(label + ': [no-json] ' + String(obj));
  }
}

function __test_apiPing(){
  const r = apiPing();
  __logJson_('apiPing', r);
}

function __test_apiDbInfo(){
  const r = apiDbInfo();
  __logJson_('apiDbInfo', r);
}

function __test_apiListClientes(){
  const r = apiListClientes({ q: '' });
  __logJson_('apiListClientes', r);
}

function __test_apiListLeads(){
  const r = apiListLeads({ q: '' });
  __logJson_('apiListLeads', r);
}

function __test_apiListFacturas_5(){
  const r = apiList('facturas', { q: '', limit: 5 });
  __logJson_('apiList facturas', r);
}

function __test_apiListPresupuestos(){
  const r = apiListPresupuestos({ q: '', includeHistorial: true });
  __logJson_('apiListPresupuestos', r);
}

function __test_apiDashboard(){
  // usa la firma que te interese:
  // - si la buena es WEBAPP_API.js -> apiDashboard(period)
  // - si la buena es API.js        -> apiDashboard()
  const r = apiDashboard();
  __logJson_('apiDashboard', r);
}

function __test_diagSheets(){
  const r = diagSheets_();
  __logJson_('diagSheets_', r);
  return r;
}


function __test_forceRebuildViews(){
  // Fuerza rebuild de vistas + index, usando tu capa cc*
  const did = (typeof ccEnsureViews_ === 'function') ? ccEnsureViews_(true) : null;
  const idx = (typeof ccBuildIndex_ === 'function') ? ccBuildIndex_() : null;
  const diag = (typeof diagSheets_ === 'function') ? diagSheets_() : { ok:false, error:'diagSheets_ missing' };
  __logJson_('forceRebuildViews.did', did);
  __logJson_('forceRebuildViews.index', idx);
  __logJson_('forceRebuildViews.diag', diag);
  return diag;
}

function __test_ccSetupWebAppLayer(){
  // Wrapper visible para ejecutar ccSetupWebAppLayer_ desde el desplegable
  const r = (typeof ccSetupWebAppLayer_ === 'function') ? ccSetupWebAppLayer_() : { ok:false, error:'ccSetupWebAppLayer_ missing' };
  __logJson_('ccSetupWebAppLayer_', r);
  const diag = (typeof diagSheets_ === 'function') ? diagSheets_() : { ok:false, error:'diagSheets_ missing' };
  __logJson_('diagSheets_', diag);
  return diag;
}

// =========================
// DIAG: origen (sheets reales)  headers + columna ID + conteos
// =========================

function __normKey__(s){
  try{
    if (typeof ccNormalizeKey_ === 'function') return ccNormalizeKey_(s);
  }catch(_){}
  return String(s||'')
    .trim()
    .toLowerCase()
    .replace(/\s+/g,'_')
    .replace(/[^\w]+/g,'_')
    .replace(/^_+|_+$/g,'');
}

function __diagOneSheet__(sheetName, idCandidates){
  const out = { sheet: sheetName, ok:false };

  try{
    const ss = (typeof _ss_ === 'function') ? _ss_() : SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(sheetName);
    if (!sh){
      out.error = 'Sheet no existe';
      return out;
    }

    // Si existe tu helper oficial, lo usamos (exactamente el mismo algoritmo)
    let data = null;
    if (typeof ccGetSheetData_ === 'function'){
      data = ccGetSheetData_(sh);
      out.used = 'ccGetSheetData_';
    }else{
      // Fallback simple si no existe (no debería pasar)
      const lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      if (lastRow < 1 || lastCol < 1) {
        out.error = 'Sheet vacía';
        return out;
      }
      const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
      let headerRow = 1;
      for (let i=0;i<Math.min(values.length,50);i++){
        const nonEmpty = values[i].filter(v => String(v||'').trim() !== '');
        if (nonEmpty.length >= 2){ headerRow = i+1; break; }
      }
      const headers = values[headerRow-1].map(h => String(h||'').trim());
      const rows = values.slice(headerRow);
      data = { headerRow, headers, rows };
      out.used = 'fallback';
    }

    out.headerRow = data.headerRow;
    out.headersCount = (data.headers||[]).length;
    out.headers = (data.headers||[]).slice(0, 60);

    // Detectar columna ID por candidatos
    const headersNorm = (data.headers||[]).map(h => __normKey__(h));
    const candNorm = (idCandidates||[]).map(x => __normKey__(x));

    let idCol = -1;
    let idHeader = null;
    for (let i=0;i<candNorm.length;i++){
      const idx = headersNorm.indexOf(candNorm[i]);
      if (idx !== -1){ idCol = idx; idHeader = data.headers[idx]; break; }
    }

    // Si no encontró exacto, intento por contains (por si es "Cliente ID", etc.)
    if (idCol === -1){
      for (let i=0;i<candNorm.length;i++){
        const needle = candNorm[i];
        const idx = headersNorm.findIndex(h => h === needle || h.includes(needle));
        if (idx !== -1){ idCol = idx; idHeader = data.headers[idx]; break; }
      }
    }

    out.idCol = idCol;
    out.idHeader = idHeader;

    // Conteos
    const rows = data.rows || [];
    out.rowsCount = rows.length;

    if (idCol === -1){
      out.ok = true;
      out.warning = 'No pude detectar columna ID con los candidatos';
      // devuelve ejemplos de 5 primeras filas no vacías (para inspección)
      const ex = [];
      for (let i=0;i<rows.length && ex.length<5;i++){
        const r = rows[i] || [];
        if (r.some(c => String(c||'').trim() !== '')){
          ex.push({ row: data.headerRow + 1 + i, firstCells: r.slice(0, 8) });
        }
      }
      out.examples = ex;
      return out;
    }

    let nonEmpty = 0;
    const examples = [];

    // Soportar rows como:
    // - Array de arrays (vía getValues)
    // - Array de objetos (vía ccGetSheetData_)
    const getIdValue = (row) => {
      if (!row) return '';
      // Caso A: array
      if (Array.isArray(row)) {
        const v = (idCol >= 0 && row[idCol] != null) ? String(row[idCol]).trim() : '';
        return v;
      }
      // Caso B: objeto
      if (typeof row === 'object') {
        const key = String(idHeader || '').trim();
        if (key && row[key] != null) return String(row[key]).trim();

        // Fallback: match case-insensitive por si cambia el case
        const low = key.toLowerCase();
        const k2 = Object.keys(row).find(k => String(k).toLowerCase() === low);
        if (k2 && row[k2] != null) return String(row[k2]).trim();
      }
      return '';
    };

    for (let i=0;i<rows.length;i++){
      const v = getIdValue(rows[i]);
      if (v){
        nonEmpty++;
        if (examples.length < 10) examples.push({ row: data.headerRow + 1 + i, id: v });
      }
    }

    out.ok = true;
    out.nonEmptyId = nonEmpty;
    out.examples = examples;
    return out;

  }catch(e){
    out.error = (e && e.message) ? e.message : String(e);
    return out;
  }
}

function __test_diagSourceSheets(){
  const result = {
    ts: new Date().toISOString(),
    sheets: {}
  };

  // Origen real (las que importan)
  result.sheets.CLIENTES    = __diagOneSheet__('CLIENTES',    ['Cliente_ID','ID','ClienteID','Cliente Id']);
  result.sheets.LEADS       = __diagOneSheet__('LEADS',       ['Lead_ID','Cliente_ID','ID','LeadID','Lead Id']);
  result.sheets.FACTURA     = __diagOneSheet__('FACTURA',     ['Factura_ID','ID','FacturaID','Factura Id','Cliente_ID']);
  result.sheets.PRESUPUESTOS= __diagOneSheet__('PRESUPUESTOS',['Pres_ID','Presupuesto_ID','ID','PresID','Cliente_ID']);
  result.sheets.GASTOS      = __diagOneSheet__('GASTOS',      ['Gasto_ID','ID','GastoID','Cliente_ID']);

  result.sheets.HISTORIAL   = __diagOneSheet__('HISTORIAL',   ['Numero_factura','Factura_ID','ID','Cliente_ID']);
  result.sheets.HIST_PRES   = __diagOneSheet__('HISTORIAL_PRESUPUESTOS', ['Pres_ID','Presupuesto_ID','ID','Cliente_ID']);
  __logJson_('diagSourceSheets_', result);
  return result;
}

// =========================
// PEEK: ver tail de una sheet (para confirmar si es tabla o plantilla)
// =========================
function __test_peekSheetTail(){
  const sheetName = 'FACTURA'; // <-- cambia aquí cuando lo vuelvas a ejecutar
  const tail = 12;            // últimas N filas
  const cols = 12;            // primeras N columnas

  const ss = (typeof _ss_ === 'function') ? _ss_() : SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    const out = { ok:false, sheetName, error:'Sheet no existe' };
    __logJson_('peekSheetTail', out);
    return out;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const startRow = Math.max(1, lastRow - tail + 1);
  const width = Math.min(cols, Math.max(1, lastCol));
  const height = Math.max(1, lastRow - startRow + 1);

  const values = sh.getRange(startRow, 1, height, width).getDisplayValues();
  const out = { ok:true, sheetName, lastRow, lastCol, startRow, height, width, values };
  __logJson_('peekSheetTail', out);
  return out;
}

// =========================
// PEEK: ver head de una sheet (para confirmar headers reales)
// =========================
function __test_peekSheetHead(){
  const sheetName = 'HISTORIAL'; // <-- cambia a HISTORIAL_PRESUPUESTOS y re-ejecutas
  const head = 12;              // primeras N filas
  const cols = 12;              // primeras N columnas

  const ss = (typeof _ss_ === 'function') ? _ss_() : SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    const out = { ok:false, sheetName, error:'Sheet no existe' };
    __logJson_('peekSheetHead', out);
    return out;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const width = Math.min(cols, Math.max(1, lastCol));
  const height = Math.min(head, Math.max(1, lastRow));

  const values = sh.getRange(1, 1, height, width).getDisplayValues();
  const out = { ok:true, sheetName, lastRow, lastCol, height, width, values };
  __logJson_('peekSheetHead', out);
  return out;
}





function __test_apiGetPresupuesto_PRO_2025_0019(){
  const r = apiGetPresupuesto('PRO-2025-0019');
  __logJson_('apiGetPresupuesto PRO-2025-0019', r);
  return r;
}

