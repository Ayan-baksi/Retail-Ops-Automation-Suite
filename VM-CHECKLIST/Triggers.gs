// Triggers.gs
// Installs the Form-submit trigger and a daily dashboard refresh.

function installTriggers_(form) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'onFormSubmitHandler' || fn === 'dailyRefresh') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmitHandler').forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger('dailyRefresh').timeBased().everyDays(1).atHour(9).create();
}

function dailyRefresh() { refreshDashboard(); }
