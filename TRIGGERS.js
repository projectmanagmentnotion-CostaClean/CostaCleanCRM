function onEdit(e) {
  try {
    if (typeof onEdit_clientes === 'function') onEdit_clientes(e);
  } catch (err) {
    logTriggerError_('onEdit_clientes', err, e);
  }

  try {
    if (typeof onEdit_leads === 'function') onEdit_leads(e);
  } catch (err) {
    logTriggerError_('onEdit_leads', err, e);
  }

  try {
    if (typeof onEditPresupuestos_ === 'function') onEditPresupuestos_(e);
  } catch (err) {
    logTriggerError_('onEditPresupuestos_', err, e);
  }
}

function onOpen(e) {
  try {
    if (typeof onOpenMain_ === 'function') onOpenMain_(e);
  } catch (err) {
    logTriggerError_('onOpenMain_', err, e);
  }
}

function logTriggerError_(where, err, e) {
  console.error(where, err);

  // toast visible cuando editas tú mismo en el sheet
  try {
    const ss = (e && e.source) ? e.source : SpreadsheetApp.getActive();
    ss.toast(`ERROR ${where}: ${err.message}`, 'TRIGGER', 8);
  } catch (_) {}
}
