// Setup.gs
// Entry point: setup() wires every file together in the correct order.
// After this runs once, everything else is driven by triggers and the
// AM Tools menu.
//
// First-time use:
//   1. Ensure Config.gs, Utils.gs, MasterData.gs, FormBuilder.gs,
//      SubmitHandler.gs, Dashboard.gs, Menu.gs, Triggers.gs, and this
//      file are all in the same Apps Script project (each as its own
//      file — they share one global scope, so file order doesn't matter).
//   2. Select "setup" from the function dropdown, run it, and approve
//      permissions (first time: Advanced -> Go to project (unsafe) -> Allow).
//   3. Check View > Logs for the form link to share with Area Managers.
//   4. Reopen or refresh the Sheet to see the "AM Tools" menu.

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
  hideFormResponsesSheet_(ss);
  buildDashboardShell_(ss);
  installTriggers_(form);

  props.setProperty('FORM_ID', form.getId());
  refreshDashboard();

  Logger.log('DONE ✅\nShare this checklist link with Area Managers:\n' + form.getPublishedUrl());
  Logger.log('Edit the form here:\n' + form.getEditUrl());
}

function resetEverything() {
  PropertiesService.getDocumentProperties().deleteProperty('FORM_ID');
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  Logger.log('Reset done. Run setup() again. Existing sheets and form are not deleted — remove them manually first for a fully clean rebuild.');
}
