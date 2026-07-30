// Triggers.gs
// Installs the two triggers this system relies on: one fires on every
// Form submission, the other refreshes the dashboard once a day as a
// safety net (e.g. so "days since last visit" stays current even with
// zero submissions that day).

function installTriggers_(form) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'onFormSubmitHandler' || fn === 'dailyRefresh') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmitHandler').forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger('dailyRefresh').timeBased().everyDays(1).atHour(9).create();
}

function dailyRefresh() { refreshDashboard(); }
