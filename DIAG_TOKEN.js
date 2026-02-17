/**
 * DIAG_TOKEN.js  funciones globales para inicializar token de diagnóstico
 * (Aparecen en el desplegable del editor Apps Script)
 */
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

function ccDiagGetToken_(){
  const props = PropertiesService.getScriptProperties();
  const t = String(props.getProperty('CC_DIAG_TOKEN') || '').trim();
  console.log("[ccDiagGetToken_] CC_DIAG_TOKEN=", t);
  return { ok:true, token: t };
}
