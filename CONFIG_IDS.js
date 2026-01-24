// Centralized default IDs for Drive folders and Docs templates.
const CC_DEFAULT_IDS = {
  PRESUPUESTOS_FOLDER_ID: '1b4R5P3ODUL1-Fp_PY8dmuVLg6UfuJjo9',
  FACTURAS_FOLDER_ID: '1111q4NpNNT_w_jZ8Mgw0AVv2VN_PFsGp',
  PRESUPUESTO_TEMPLATE_ID: '1M2tpK-Iq6_WuVmHxahkHbtJrmOtLmPu502-TjnqkQ8',
  FACTURA_TEMPLATE_ID: '10U_1CxZBEc46OP5W1d98a11y1I0OUV7T1n7qShH6as'
};

// Build stamp (solo debug)
const BUILD_STAMP = new Date().toISOString();


/* =========================
   DB Spreadsheet binding
   - Guarda Spreadsheet ID de la "DB" en ScriptProperties
   - WebApp NO debe depender de getActiveSpreadsheet()
========================= */

const CC_DB_SPREADSHEET_ID_KEY = "CC_DB_SPREADSHEET_ID";

/**
 * Setea el Spreadsheet ID de la DB.
 * Uso recomendado:
 *  - ccSetDbSpreadsheetId("TU_SPREADSHEET_ID")
 * Opcional:
 *  - ccSetDbSpreadsheetId() si estás ejecutando desde un editor ligado a la DB (no webapp)
 */
function ccSetDbSpreadsheetId(spreadsheetIdOpt){
  const props = PropertiesService.getScriptProperties();
  let ssid = spreadsheetIdOpt || "";

  if (!ssid){
    try{
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      ssid = ss ? ss.getId() : "";
    }catch(e){
      ssid = "";
    }
  }

  if (!ssid){
    throw new Error(
      'CC_DB_SPREADSHEET_ID no está seteado. Ejecuta ccSetDbSpreadsheetId("SPREADSHEET_ID") (recomendado).'
    );
  }

  props.setProperty(CC_DB_SPREADSHEET_ID_KEY, ssid);
  return { ok:true, spreadsheetId:ssid };
}

function ccGetDbSpreadsheetId_(){
  return PropertiesService.getScriptProperties().getProperty(CC_DB_SPREADSHEET_ID_KEY) || "";
}

function ccGetDbSpreadsheet_(){
  const ssid = ccGetDbSpreadsheetId_();
  if (!ssid){
    throw new Error('CC_DB_SPREADSHEET_ID is not set. Run ccSetDbSpreadsheetId("SPREADSHEET_ID") once.');
  }
  return SpreadsheetApp.openById(ssid);
}

