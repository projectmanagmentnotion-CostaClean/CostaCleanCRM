function doGet() {
  return HtmlService.createTemplateFromFile('app')
    .evaluate()
    .setTitle('Costa Clean CRM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ✅ necesario para incluir styles.html dentro de app.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
