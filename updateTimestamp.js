function updateTimestamp() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var now = new Date();
  sheet.getRange('D1').setValue(now);
}