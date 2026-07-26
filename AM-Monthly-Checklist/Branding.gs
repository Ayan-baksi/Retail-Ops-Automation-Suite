/**
 * Branding.gs
 * Handles logo management for the Google Form and Dashboard.
 *
 * Usage:
 * 1. Set CONFIG.LOGO_FILE_ID in Config.gs.
 * 2. Run "Apply/Refresh Logo" from the AM Tools menu.
 *
 * Safe to run multiple times. Existing dashboard logos are replaced
 * to prevent duplicates.
 */

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

  // Apply logo to the Google Form.
  FormApp.openById(formId).setImage(blob);

  // Replace any existing logo on the dashboard.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(SHEETS.DASH);

  dash.getImages().forEach(function (img) {
    img.remove();
  });

  var img = dash.insertImage(blob, 1, 1);
  img.setWidth(50).setHeight(50);

  ui.alert('Done — logo applied to the Form header and the Dashboard.');
}

/**
 * Ensures the dashboard logo exists after a dashboard refresh.
 * Prevents duplicate images from being inserted.
 */
function ensureLogoOnDashboard_(dash) {
  if (!CONFIG.LOGO_FILE_ID) return;
  if (dash.getImages().length > 0) return;

  try {
    var blob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
    var img = dash.insertImage(blob, 1, 1);
    img.setWidth(50).setHeight(50);
  } catch (err) {
    // Ignore if the logo is unavailable or inaccessible.
  }
}
