/*************************************************************************
 * Branding.gs — applies your logo to the Form header and the Dashboard.
 * Run this any time from the AM Tools menu (safe to run repeatedly —
 * it won't create duplicate logos).
 *
 * SETUP: paste your logo's Google Drive file ID into CONFIG.LOGO_FILE_ID
 * in Config.gs, save, then run "🖼️ Apply/Refresh Logo" from AM Tools.
 *************************************************************************/

function applyLogo() {
  var ui = SpreadsheetApp.getUi();
  if (!CONFIG.LOGO_FILE_ID) {
    ui.alert('No logo file ID yet. Paste your logo\'s Google Drive file ID into CONFIG.LOGO_FILE_ID in Config.gs, save, then run this again.');
    return;
  }
  var formId = PropertiesService.getDocumentProperties().getProperty('FORM_ID');
  if (!formId) {
    ui.alert('Run "Setup System" first — no form exists yet.');
    return;
  }

  var blob;
  try {
    blob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
  } catch (err) {
    ui.alert('Could not open that file. Double-check the file ID in Config.gs is correct and that the file still exists in your Drive.');
    return;
  }

  FormApp.openById(formId).setImage(blob);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(SHEETS.DASH);
  dash.getImages().forEach(function (img) { img.remove(); }); // clear any previous logo first
  var img = dash.insertImage(blob, 1, 1);
  img.setWidth(50).setHeight(50);

  ui.alert('Done — logo applied to the Form header and the Dashboard.');
}

// Called at the end of every dashboard refresh so the logo survives a
// full rebuild without ever being duplicated.
function ensureLogoOnDashboard_(dash) {
  if (!CONFIG.LOGO_FILE_ID) return;
  if (dash.getImages().length > 0) return;
  try {
    var blob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
    var img = dash.insertImage(blob, 1, 1);
    img.setWidth(50).setHeight(50);
  } catch (err) {
    // Logo not set up yet, or file not accessible — dashboard still works fine without it.
  }
}
