function createMonthlyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'duplicateAndRenameSheet') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('duplicateAndRenameSheet')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();
}
