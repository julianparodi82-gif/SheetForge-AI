function logEvent_(type, prompt, plan, status, message, affectedRanges, durationMs) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('LOG');
  if (!sheet) return;
  var user = Session.getActiveUser().getEmail() || 'desconocido';
  var row = [
    new Date(),
    user,
    prompt || '',
    plan ? JSON.stringify(plan) : '',
    status || '',
    message || '',
    affectedRanges || '',
    durationMs || 0
  ];
  sheet.appendRow(row);
}

function getLastLogs(n) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('LOG');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var count = Math.min(n || 10, sheet.getLastRow() - 1);
  var values = sheet.getRange(sheet.getLastRow() - count + 1, 1, count, 8).getValues();
  return values.map(function (row) {
    return row[0] + ' | ' + row[4] + ' | ' + row[5];
  }).reverse();
}
