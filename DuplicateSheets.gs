function duplicateAndRenameSheet() {
  const ss = SpreadsheetApp.getActive();
  const currentDate = new Date();
  const previousMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1);
  const previousMonthName = Utilities.formatDate(previousMonth, Session.getScriptTimeZone(), "MM-yyyy");
  const currentMonth = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MM-yyyy");

  const previousSheet = ss.getSheetByName(previousMonthName);
  if (!previousSheet) {
    // If the previous month's sheet doesn't exist, display an error message and exit
    SpreadsheetApp.getUi().alert("The previous month's sheet doesn't exist.");
    return;
  }

  const newSheet = previousSheet.copyTo(ss);
  newSheet.setName(currentMonth);

  // Set the value of the new sheet to the current month and year
  newSheet.getRange("A1").setValue(`Month: ${currentMonth}`);
  newSheet.getRange("A2").setValue(`Year: ${currentDate.getFullYear()}`);
}
