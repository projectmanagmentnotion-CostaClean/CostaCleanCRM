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

