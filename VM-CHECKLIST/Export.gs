// Export.gs
// Exports the Dashboard tab to PDF and saves it to Drive.

function exportDashboardToPdf() {
  var ui = SpreadsheetApp.getUi();
  var log = [];

  try {
    log.push('Step 1: Locating spreadsheet and Dashboard sheet...');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dash = ss.getSheetByName(SHEETS.DASH);
    if (!dash) throw new Error('Could not find a sheet named "' + SHEETS.DASH + '".');
    var ssId = ss.getId();
    var gid = dash.getSheetId();

    log.push('Step 2: Building export URL...');
    var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export' +
      '?format=pdf&gid=' + gid +
      '&size=A4&portrait=false&fitw=true' +
      '&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3' +
      '&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=false&horizontal_alignment=CENTER';

    log.push('Step 3: Fetching OAuth token...');
    var token = ScriptApp.getOAuthToken();
    if (!token) throw new Error('Got an empty OAuth token — authorization likely was not completed.');

    log.push('Step 4: Requesting the PDF from Google...');
    var response;
    try {
      response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    } catch (fetchErr) {
      throw new Error('UrlFetchApp.fetch threw an error — almost always the missing "script.external_request" scope. Raw: ' + fetchErr.message);
    }

    var code = response.getResponseCode();
    if (code !== 200) throw new Error('Google returned HTTP ' + code + '. Response snippet: ' + response.getContentText().substring(0, 300));

    log.push('Step 5: Building the PDF file blob...');
    var fileName = CONFIG.ORG_NAME + ' VM Dashboard - ' + Utilities.formatDate(new Date(), TZ, 'dd-MMM-yyyy HHmm') + '.pdf';
    var blob = response.getBlob().setName(fileName);
    if (blob.getBytes().length < 1000) throw new Error('The PDF came back suspiciously small — likely an error page, not a real PDF.');

    log.push('Step 6: Finding or creating the Drive export folder...');
    var folder = getOrCreateExportFolder_();

    log.push('Step 7: Saving the PDF...');
    var file = folder.createFile(blob);

    Logger.log(log.join('\n'));
    ui.alert('Exported ✅', 'Saved to your Drive folder "DeeHub VM Dashboard Exports" as:\n' + file.getName(), ui.ButtonSet.OK);

  } catch (err) {
    Logger.log('EXPORT FAILED\n' + log.join('\n') + '\n\nError: ' + err.message);
    ui.alert('Export failed', 'Something went wrong:\n' + err.message + '\n\nFull log in Executions.', ui.ButtonSet.OK);
  }
}

function getOrCreateExportFolder_() {
  var name = 'DeeHub VM Dashboard Exports';
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function authorizeAllPermissions() {
  SpreadsheetApp.getActiveSpreadsheet().getName();
  DriveApp.getRootFolder().getName();
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  Logger.log('If this ran with no error, all required permissions are granted.');
}
