function logEvent_(ss, modulo, accion, entidad, id, resultado, mensaje, data) {
  try {
    if (!ss) return;
    const sheetName = 'LOG';
    let sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
    }

    const headers = ['Fecha', 'Modulo', 'Accion', 'Entidad', 'ID', 'Resultado', 'Mensaje', 'DataJSON'];
    const lastRow = sh.getLastRow();
    if (lastRow === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else if (lastRow === 1) {
      const existing = sh.getRange(1, 1, 1, headers.length).getValues()[0];
      const isEmpty = existing.every((v) => String(v || '').trim() === '');
      if (isEmpty) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
    }

    const dataJson = data == null ? '' : JSON.stringify(data);
    sh.appendRow([
      new Date(),
      modulo || '',
      accion || '',
      entidad || '',
      id || '',
      resultado || '',
      mensaje || '',
      dataJson
    ]);
  } catch (err) {
    console.error('logEvent_ error', err);
  }
}
