function exportSheetPdf_(sheetName, filename) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  var url = buildExportUrl_(ss.getId(), sheet.getSheetId());
  var blob = fetchPdfBlob_(url).setName(filename);
  var file = DriveApp.createFile(blob);
  return file.getUrl();
}

function exportSpreadsheetPdf_(filename) {
  var ss = SpreadsheetApp.getActive();
  var url = buildExportUrl_(ss.getId(), null);
  var blob = fetchPdfBlob_(url).setName(filename || (ss.getName() + '.pdf'));
  var file = DriveApp.createFile(blob);
  return file.getUrl();
}

function buildExportUrl_(spreadsheetId, sheetId) {
  var base = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=pdf';
  if (sheetId) {
    base += '&gid=' + sheetId;
  }
  return base;
}

function fetchPdfBlob_(url) {
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token }
  });
  return response.getBlob();
}

function insertImageFromDrive_(action) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(action.sheetName);
  var cell = sheet.getRange(action.cell);
  var file = DriveApp.getFileById(action.fileId);
  var blob = file.getBlob();
  var image = sheet.insertImage(blob, cell.getColumn(), cell.getRow());
  if (action.width) image.setWidth(action.width);
  if (action.height) image.setHeight(action.height);
}
