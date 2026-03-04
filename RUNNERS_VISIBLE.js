// CC_RUNNERS_VERSION=20260304-1
// RUNNERS_VISIBLE.js
// Runners 100% visibles en el desplegable (function top-level)
// Objetivo: verificar qué funciones están realmente disponibles y setear/ver DB Spreadsheet ID.

function run_verifyRunnableFunctions_(){
  // Lista funciones globales (mejor esfuerzo en Apps Script V8)
  var names = [];
  try {
    var g = (typeof globalThis !== 'undefined') ? globalThis : this;
    for (var k in g){
      try{
        if (typeof g[k] === 'function') names.push(k);
      }catch(_){}
    }
  } catch (e){
    names = [];
  }

  names.sort();

  var must = [
    'run_verifyRunnableFunctions_',
    'run_setDbSpreadsheetId_',
    'run_getDbSpreadsheetId_',
    'ccSetDbSpreadsheetId',
    'ccGetDbSpreadsheetId_',
    'ccGetDbSpreadsheet_',
    'ccDiagInitToken_',
    'ccDiagGetToken_',
    '__debug_get_db_property__'
  ];

  var present = {};
  must.forEach(function(n){ present[n] = (names.indexOf(n) >= 0); });

  console.log('[RUNNABLE VERIFY] present=', JSON.stringify(present, null, 2));
  console.log('[RUNNABLE VERIFY] sample globals (first 120)=', names.slice(0,120).join(', '));

  return { ok:true, present:present, count:names.length, sample:names.slice(0,120) };
}

function run_setDbSpreadsheetId_(){
  // NO depende de que ccSetDbSpreadsheetId sea visible en UI.
  // Solo depende de que exista en runtime (lo verificamos con run_verifyRunnableFunctions_).
  var id = '1m62QB04_aDrxeXjSiK6QHRdAztGPqTAhx5zA9cKa8kk';
  if (typeof ccSetDbSpreadsheetId !== 'function') {
    throw new Error('ccSetDbSpreadsheetId NO existe en runtime. Revisa CONFIG_IDS.js y que esté top-level como function ccSetDbSpreadsheetId(...)');
  }
  var out = ccSetDbSpreadsheetId(id);
  console.log('[DB SET] out=', JSON.stringify(out));
  return out;
}

function run_getDbSpreadsheetId_(){
  // Lectura directa de ScriptProperties (sin depender de __debug_get_db_property__)
  var props = PropertiesService.getScriptProperties();
  var v = props.getProperty('CC_DB_SPREADSHEET_ID') || '';
  console.log('[DB GET] CC_DB_SPREADSHEET_ID=', v);
  return { ok:true, CC_DB_SPREADSHEET_ID: v || null };
}


