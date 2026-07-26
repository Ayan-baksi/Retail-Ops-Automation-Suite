// Branding.gs
// Applies the configured logo to the Form header and the Dashboard sheet.
// Run from the AM Tools menu. Idempotent — safe to re-run without duplicating the logo.
//
// Setup: set CONFIG.LOGO_FILE_ID in Config.gs to your logo's Drive file ID,
// then run "Apply/Refresh Logo" from the AM Tools menu.

function applyLogo() {
  var ui = SpreadsheetApp.getUi();
  if (!CONFIG.LOGO_FILE_ID) {
    ui.alert('No logo file ID set. Add one to CONFIG.LOGO_FILE_ID in Config.gs, then run this again.');
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
    ui.alert('Could not open that file. Verify the file ID in Config.gs and that the file still exists in Drive.');
    return;
  }

  FormApp.openById(formId).setImage(blob);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(SHEETS.DASH);
  dash.getImages().forEach(function (img) { img.remove(); });
  var img = dash.insertImage(blob, 1, 1);
  img.setWidth(50).setHeight(50);

  ui.alert('Done — logo applied to the Form header and Dashboard.');
}

// Re-applies the logo after a dashboard rebuild, without duplicating it.
function ensureLogoOnDashboard_(dash) {
  if (!CONFIG.LOGO_FILE_ID) return;
  if (dash.getImages().length > 0) return;
  try {
    var blob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
    var img = dash.insertImage(blob, 1, 1);
    img.setWidth(50).setHeight(50);
  } catch (err) {
    // Logo not configured, or file inaccessible — dashboard still works without it.
  }
}
