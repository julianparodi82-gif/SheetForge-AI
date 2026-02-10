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
    var timestamp = row[0] instanceof Date
      ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
      : row[0];
    var status = normalizeLogStatus_(row[4]);
    var message = row[5] || 'sin detalles';
    var summary = buildLogSummary_(row[3]);
    var summaryText = summary ? '\nResumen: ' + summary : '';
    return 'Fecha: ' + timestamp + '\nEstado: ' + status + '\nMensaje: ' + message + summaryText;
  }).reverse();
}

function clearLogs() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('LOG');
  if (!sheet || sheet.getLastRow() < 2) return { ok: true };
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  return { ok: true };
}

function normalizeLogStatus_(status) {
  if (!status) return 'sin estado';
  var value = status.toString().toLowerCase();
  if (value === 'ok') return 'Aplicado';
  if (value === 'error') return 'Error';
  return status;
}

function buildLogSummary_(planJson) {
  if (!planJson) return '';
  try {
    var plan = JSON.parse(planJson);
    if (!plan || !plan.actions) return '';
    return buildSummary_(plan);
  } catch (e) {
    return '';
  }
}
