function onOpen() {
  SpreadsheetApp.getUi().createMenu('SheetForge AI')
    .addItem('Operador IA', 'showSidebar')
    .addItem('Configuración', 'showConfig')
    .addToUi();
}

function showSidebar() {
  ensureSpecialSheets_();
  var html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('Operador IA');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showConfig() {
  ensureSpecialSheets_();
  var html = HtmlService.createTemplateFromFile('config')
    .evaluate()
    .setWidth(420)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configuración');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ensureSpecialSheets_() {
  hideSheet_(getOrCreateSheet_('LOG'));
  hideSheet_(getOrCreateSheet_('CONFIG'));
  hideSheet_(getOrCreateSheet_('CONTEXTO_IA'));
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

function hideSheet_(sheet) {
  if (sheet && !sheet.isSheetHidden()) {
    sheet.hideSheet();
  }
}

function getSettings() {
  var props = PropertiesService.getDocumentProperties();
  var userProps = PropertiesService.getUserProperties();
  return {
    language: props.getProperty('APP_LANGUAGE') || Session.getActiveUserLocale(),
    aiEndpoint: props.getProperty('AI_ENDPOINT') || '',
    aiApiKey: userProps.getProperty('AI_API_KEY') || props.getProperty('AI_API_KEY') || ''
  };
}

function saveSettings(settings) {
  if (!settings) {
    throw new Error('Datos de configuración vacíos.');
  }
  var props = PropertiesService.getDocumentProperties();
  var userProps = PropertiesService.getUserProperties();
  var language = (settings.language || '').toString().trim();
  var endpoint = settings.aiEndpoint !== undefined ? settings.aiEndpoint.toString().trim() : '';
  var apiKey = (settings.aiApiKey || '').toString().trim();
  if (!apiKey) {
    throw new Error('La API key no puede estar vacía.');
  }
  if (language) props.setProperty('APP_LANGUAGE', language);
  props.setProperty('AI_ENDPOINT', endpoint);
  userProps.setProperty('AI_API_KEY', apiKey);
  props.setProperty('AI_API_KEY', apiKey);
  return {
    ok: true,
    message: 'Clave guardada correctamente.',
    settings: {
      language: language || Session.getActiveUserLocale(),
      aiEndpoint: endpoint,
      aiApiKey: apiKey
    }
  };
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
