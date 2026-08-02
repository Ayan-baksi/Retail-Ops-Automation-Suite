// Setup.gs
// Entry point: setup() wires every file together in order.

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('FORM_ID')) {
    Logger.log('A form already exists. Run resetEverything() first if you really want to rebuild from scratch.');
    return;
  }

  var form = buildForm_();
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  buildAllMasterData_(ss);
  buildMasterLogSheet_(ss);
  buildIssueLogSheet_(ss);
  buildItemScoreLogSheet_(ss);
  hideFormResponsesSheet_(ss);
  buildDashboardShell_(ss);
  installTriggers_(form);

  props.setProperty('FORM_ID', form.getId());
  refreshDashboard();

  Logger.log('DONE ✅\nShare this checklist link with VMs:\n' + form.getPublishedUrl());
  Logger.log('Edit the form here:\n' + form.getEditUrl());
}

function resetEverything() {
  PropertiesService.getDocumentProperties().deleteProperty('FORM_ID');
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  Logger.log('Reset done. Run setup() again. Existing sheets and form are not deleted — remove them manually first for a fully clean rebuild.');
}
