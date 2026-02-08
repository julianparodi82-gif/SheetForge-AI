function onOpen() {
  SpreadsheetApp.getUi().createMenu('SheetForge AI')
    .addItem('AI Operator', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  ensureSpecialSheets_();
  var html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('AI Operator');
  SpreadsheetApp.getUi().showSidebar(html);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ensureSpecialSheets_() {
  getOrCreateSheet_('LOG');
  getOrCreateSheet_('CONFIG');
  getOrCreateSheet_('CONTEXTO_IA');
  ensureLogHeaders_();
  ensureConfigDefaults_();
  ensureContextHeader_();
}

function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function ensureLogHeaders_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('LOG');
  var headers = ['timestamp', 'user', 'prompt', 'plan_json', 'status', 'message', 'affected_ranges', 'duration_ms'];
  var range = sheet.getRange(1, 1, 1, headers.length);
  var values = range.getValues()[0];
  var needs = false;
  for (var i = 0; i < headers.length; i++) {
    if (values[i] !== headers[i]) {
      needs = true;
      break;
    }
  }
  if (needs) {
    range.setValues([headers]);
  }
}

function ensureConfigDefaults_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('CONFIG');
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 6).setValues([[
      'label', 'sheetName', 'rangeA1', 'includeValues', 'includeFormulas', 'includeFormats'
    ]]);
  }
}

function ensureContextHeader_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('CONTEXTO_IA');
  if (sheet.getLastRow() === 0) {
    writeContextLines_([
      'CONTEXTO_IA — Estado actual del Spreadsheet',
      'ultima_actualizacion_iso',
      'spreadsheet_id',
      'spreadsheet_nombre',
      'timezone',
      'modo_contexto (quick|full)',
      'indice_hojas (name(sheetId),…)',
      '—'
    ]);
  }
}

function writeContextLines_(lines) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('CONTEXTO_IA');
  sheet.clearContents();
  var rows = lines.map(function (line) { return [line]; });
  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
}
