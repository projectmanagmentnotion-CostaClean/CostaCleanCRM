/**
 * RUNNERS PRO (logs)  CostaClean CRM
 * Objetivo: ejecutar APIs reales y ver output en Cloud Logging
 * No afecta la arquitectura, solo runners para debug.
 */

function _log_(label, obj){
  var msg = label + " => " + JSON.stringify(obj, null, 2);
  try { Logger.log(msg); } catch(e){}
  try { console.log(msg); } catch(e){}
  return obj;
}

function run_apiPing(){
  return _log_("apiPing", apiPing());
}

function run_apiDbInfo(){
  return _log_("apiDbInfo", apiDbInfo());
}

function run_apiDashboard(){
  return _log_("apiDashboard", apiDashboard());
}

function run_apiListClientes(){
  // si soporta params, esto ayuda; si no, igual funcionará o te dirá el error
  return _log_("apiListClientes", apiListClientes({ limit: 10 }));
}

function run_apiListPresupuestos(){
  return _log_("apiListPresupuestos", apiListPresupuestos({ limit: 10 }));
}

function run_apiListLeads(){
  return _log_("apiListLeads", apiListLeads({ limit: 10 }));
}
