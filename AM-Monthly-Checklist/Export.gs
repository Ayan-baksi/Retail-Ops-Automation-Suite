// Export.gs
// Exports the Dashboard sheet to PDF and saves it to Drive. Triggered
// by the "EXPORT TO PDF" checkbox on the Dashboard (see onEdit() in
// Dashboard.gs), or can be run manually.
//
// Troubleshooting a permissions error on UrlFetchApp
// (script.external_request scope):
//   1. Apps Script editor → Project Settings → enable "Show
//      appsscript.json manifest file in editor".
//   2. Open appsscript.json and remove any oauthScopes array that
//      restricts scopes (or use the version in this repo) → Save.
//   3. Go to myaccount.google.com/permissions, find this project, and
//      remove access — this forces a full consent screen on next run.
//   4. Run exportDashboardToPdf from the function dropdown and approve
//      all requested permissions (Sheets, Drive, external requests).

function exportDashboardToPdf() {
  var ui = SpreadsheetApp.getUi();
  var log = [];

  try {
    log.push('Step 1: Locating spreadsheet and Dashboard sheet...');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dash = ss.getSheetByName(SHEETS.DASH);
    if (!dash) throw new Error('Could not find a sheet named "' + SHEETS.DASH + '". Has it been renamed?');
    var ssId = ss.getId();
    var gid = dash.getSheetId();
    log.push('  Spreadsheet ID: ' + ssId + ' | Dashboard GID: ' + gid);

    log.push('Step 2: Building export URL...');
    var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export' +
      '?format=pdf&gid=' + gid +
      '&size=A4&portrait=false&fitw=true' +
      '&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3' +
      '&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=false&horizontal_alignment=CENTER';
    log.push('  URL: ' + url);

    log.push('Step 3: Fetching OAuth token...');
    var token = ScriptApp.getOAuthToken();
    if (!token) throw new Error('Got an empty OAuth token — authorization likely was not completed.');

    log.push('Step 4: Requesting the PDF from Google (UrlFetchApp)...');
    var response;
    try {
      response = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
    } catch (fetchErr) {
      throw new Error('UrlFetchApp.fetch threw an error — this is almost always the missing '
        + '"script.external_request" scope. See the troubleshooting note at the top of this file. Raw error: ' + fetchErr.message);
    }

    var code = response.getResponseCode();
    log.push('  HTTP response code: ' + code);
    if (code !== 200) {
      var bodySnippet = response.getContentText().substring(0, 300);
      throw new Error('Google returned HTTP ' + code + ' instead of 200. This usually means the GID is wrong, '
        + 'or the spreadsheet couldn\'t be read with this token. Response snippet: ' + bodySnippet);
    }

    log.push('Step 5: Building the PDF file blob...');
    var fileName = CONFIG.ORG_NAME + ' Dashboard - ' + Utilities.formatDate(new Date(), TZ, 'dd-MMM-yyyy HHmm') + '.pdf';
    var blob = response.getBlob().setName(fileName);
    var blobSize = blob.getBytes().length;
    log.push('  Blob size: ' + blobSize + ' bytes');
    if (blobSize < 1000) {
      throw new Error('The PDF came back suspiciously small (' + blobSize + ' bytes) — likely an error page rather than a real export.');
    }

    log.push('Step 6: Finding or creating the Drive export folder...');
    var folder = getOrCreateExportFolder_();
    log.push('  Folder: "' + folder.getName() + '" (ID: ' + folder.getId() + ')');

    log.push('Step 7: Saving the PDF into that folder...');
    var file = folder.createFile(blob);
    log.push('  Saved as: ' + file.getName() + ' (File ID: ' + file.getId() + ')');

    Logger.log(log.join('\n'));
    ui.alert('Exported ✅', 'Saved to your Drive folder "DeeHub Dashboard Exports" as:\n' + file.getName(), ui.ButtonSet.OK);

  } catch (err) {
    Logger.log('EXPORT FAILED\n' + log.join('\n') + '\n\nError: ' + err.message);
    ui.alert('Export failed',
      'Something went wrong:\n' + err.message +
      '\n\nA full step-by-step log has been written to the Apps Script Execution log '
      + '(Executions in the left sidebar) so you can see exactly which step failed.',
      ui.ButtonSet.OK);
  }
}

function getOrCreateExportFolder_() {
  var name = 'DeeHub Dashboard Exports';
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

// Run manually once, after fixing permissions per the note at the top
// of this file, to trigger the consent screen for every required scope
// in one shot instead of one at a time.
function authorizeAllPermissions() {
  SpreadsheetApp.getActiveSpreadsheet().getName();      // Sheets scope
  DriveApp.getRootFolder().getName();                    // Drive scope
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); // external_request scope
  Logger.log('If this ran with no error, all required permissions are granted.');
}
