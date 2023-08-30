function createMonthlyTrigger() {
  // Delete existing triggers to avoid any dupes
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  // Create a new trigger to run on the 1st day of each month
  ScriptApp.newTrigger('duplicateAndRenameSheet')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();
}
