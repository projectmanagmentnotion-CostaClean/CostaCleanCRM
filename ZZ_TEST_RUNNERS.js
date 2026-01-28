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
