function doGet(e) {
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
// ✅ necesario para incluir styles.html dentro de app.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}





// STUB para evitar errores por triggers antiguos
function onOpenRouter_(){
  console.log('onOpenRouter_ STUB called');
}


