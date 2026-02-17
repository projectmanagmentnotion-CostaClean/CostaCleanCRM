function doGet(e) {
  // === DIAG RUNNER (JSON) ===
  // Llama: ?diag=1&fn=apiPing&token=XXXX&args=%5B...%5D
  try{
    const p0 = e && e.parameter ? e.parameter : {};
    if (String(p0.diag || '').trim() === '1') {
      return __ccDiagHandleGet_(e);
    }
  }catch(_){}
  const t = HtmlService.createTemplateFromFile('app');

  // Debug server-driven: ?debug=1 fuerza ON, ?debug=0 fuerza OFF, undefined = no fuerza
  let dbg;
  try{
    const p = e && e.parameter ? e.parameter : {};
    if (String(p.debug || '').trim() === '1') dbg = true;
    else if (String(p.debug || '').trim() === '0') dbg = false;
  }catch(_){}

  t.CC_DEBUG = (typeof dbg !== 'undefined') ? dbg : null;

  return t.evaluate()
    .setTitle('Costa Clean CRM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
 
// =========================
// DIAGNOSTIC RUNNER (WebApp)
// =========================
var __ccDiagJson_ = function(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
};

var __ccDiagHandleGet_ = function(e){
  const p = (e && e.parameter) ? e.parameter : {};
  const token = String(p.token || '').trim();

  const props = PropertiesService.getScriptProperties();
  const expected = String(props.getProperty('CC_DIAG_TOKEN') || '').trim();

  if (!expected) {
    return __ccDiagJson_({
      ok:false,
      code:500,
      error:"CC_DIAG_TOKEN no está configurado. Ejecuta ccDiagInitToken_() una vez desde Apps Script."
    });
  }
  if (!token || token !== expected) {
    return __ccDiagJson_({ ok:false, code:403, error:"Forbidden (token inválido)" });
  }

  const fnName = String(p.fn || '').trim();

  // Whitelist estricta (solo entrypoints seguros)
  const ALLOW = {
    apiPing: apiPing,
    apiDbInfo: apiDbInfo,
    apiDashboard: apiDashboard,
    apiListClientes: apiListClientes,
    apiListLeads: apiListLeads,
    apiListPresupuestos: apiListPresupuestos,
    apiGetPresupuesto: apiGetPresupuesto,
    apiCreate: apiCreate,
    apiUpdate: apiUpdate,
    apiAction: apiAction,
    apiGeneratePresupuestoPdf: apiGeneratePresupuestoPdf,
    apiGenerateFacturaPdf: apiGenerateFacturaPdf,
    apiCrearFacturaDesdePresupuesto: apiCrearFacturaDesdePresupuesto,
    apiCrearPresupuestoLead: apiCrearPresupuestoLead,
    diagSheets_: diagSheets_,
    _ss_: _ss_
  };

  if (!ALLOW[fnName]) {
    return __ccDiagJson_({
      ok:false,
      code:400,
      error:"fn no permitida",
      allowed:Object.keys(ALLOW).sort()
    });
  }

  // args: JSON array URL-encoded (opcional)
  let args = [];
  try{
    if (p.args) {
      args = JSON.parse(String(p.args));
      if (!Array.isArray(args)) args = [args];
    }
  }catch(err){
    return __ccDiagJson_({
      ok:false,
      code:400,
      error:"args inválidos (JSON)",
      detail:String(err && err.message ? err.message : err)
    });
  }

  try{
    const data = ALLOW[fnName].apply(null, args);
    return __ccDiagJson_({ ok:true, fn: fnName, args: args, data: data });
  }catch(err){
    return __ccDiagJson_({
      ok:false,
      code:500,
      fn: fnName,
      error: String(err && err.message ? err.message : err),
      stack: err && err.stack ? String(err.stack) : null
    });
  }
};

// Ejecuta 1 vez en el editor de Apps Script (Run) para generar token
function ccDiagInitToken_(){
  const props = PropertiesService.getScriptProperties();
  let t = String(props.getProperty('CC_DIAG_TOKEN') || '').trim();
  if (!t) {
    t = Utilities.getUuid().replace(/-/g,'');
    props.setProperty('CC_DIAG_TOKEN', t);
  }
  console.log("[ccDiagInitToken_] CC_DIAG_TOKEN=", t);
  return { ok:true, token: t };
}
// ✅ necesario para incluir styles.html dentro de app.html
var include = function(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}





// STUB para evitar errores por triggers antiguos
var onOpenRouter_ = function(){
  console.log('onOpenRouter_ STUB called');
}






