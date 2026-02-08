var ALLOWED_OPS = [
  'setValue', 'setFormula', 'setFormulas', 'setBackground', 'setBackgrounds', 'setNumberFormat', 'clearContent',
  'copyPasteValues', 'moveRange', 'insertRows', 'deleteRows', 'renameSheet', 'createSheet',
  'deleteSheet', 'setColumnWidth', 'setRowHeight', 'renameFile', 'exportPdf',
  'exportSpreadsheetPdf', 'insertImageFromDrive', 'setImageFormula'
];

function validatePlan_(plan, flags) {
  if (!plan || !plan.meta || !Array.isArray(plan.actions)) {
    return 'Plan inválido.';
  }
  for (var i = 0; i < plan.actions.length; i++) {
    var action = plan.actions[i];
    if (!action.op || ALLOWED_OPS.indexOf(action.op) === -1) {
      return 'Operación no permitida: ' + action.op + ' (no está en whitelist)';
    }
    var v = validateAction_(action, flags);
    if (v) return v;
  }
  return null;
}

function validateAction_(action, flags) {
  if (flags && flags.safeMode) {
    var blockedOps = getSafeBlockedOps_();
    if (blockedOps.indexOf(action.op) !== -1) {
      return 'Operación no permitida: ' + action.op + ' (bloqueada por Safe Mode)';
    }
    if (action.op === 'deleteRows' && action.numRows && action.numRows > 50 && blockedOps.indexOf('deleteRows') !== -1) {
      return 'Operación no permitida: deleteRows (bloqueada por Safe Mode)';
    }
    if (action.op === 'clearContent' && action.rangeA1 && rangeIsLarge_(action.rangeA1) && blockedOps.indexOf('clearContent') !== -1) {
      return 'Operación no permitida: clearContent (bloqueada por Safe Mode)';
    }
  }
  if (requiresSheet_(action.op)) {
    if (!action.sheetName) {
      return 'Operación no permitida: falta sheetName';
    }
    var sheet = SpreadsheetApp.getActive().getSheetByName(action.sheetName);
    if (!sheet) {
      return 'Hoja no encontrada: ' + action.sheetName;
    }
  }
  if (requiresRange_(action.op) && !action.rangeA1) {
    return 'Rango inválido: falta rangeA1';
  }
  if (action.op === 'setBackgrounds' && !Array.isArray(action.colors)) {
    return 'Color inválido: falta colors';
  }
  if (action.op === 'copyPasteValues') {
    if (!action.sourceA1 || !action.targetA1) {
      return 'Rango inválido: falta sourceA1/targetA1';
    }
  }
  if (action.op === 'moveRange') {
    if (!action.sourceA1 || !action.targetA1) {
      return 'Rango inválido: falta sourceA1/targetA1';
    }
  }
  if (action.op === 'insertImageFromDrive' && !action.cell) {
    return 'Rango inválido: falta cell';
  }
  if (action.rangeA1 && !isValidA1_(action.rangeA1)) {
    return 'Rango inválido: ' + action.rangeA1;
  }
  if (action.color && !/^#[0-9a-fA-F]{6}$/.test(action.color)) {
    return 'Color inválido: ' + action.color;
  }
  return null;
}

function requiresSheet_(op) {
  return [
    'setValue', 'setFormula', 'setFormulas', 'setBackground', 'setBackgrounds', 'setNumberFormat', 'clearContent',
    'copyPasteValues', 'moveRange', 'insertRows', 'deleteRows', 'renameSheet', 'createSheet',
    'deleteSheet', 'setColumnWidth', 'setRowHeight', 'exportPdf', 'insertImageFromDrive',
    'setImageFormula'
  ].indexOf(op) !== -1;
}

function requiresRange_(op) {
  return [
    'setValue', 'setFormula', 'setFormulas', 'setBackground', 'setBackgrounds', 'setNumberFormat', 'clearContent',
    'setImageFormula'
  ].indexOf(op) !== -1;
}

function isValidA1_(rangeA1) {
  try {
    var sheet = SpreadsheetApp.getActive().getActiveSheet();
    sheet.getRange(rangeA1);
    return true;
  } catch (e) {
    return false;
  }
}

function rangeIsLarge_(rangeA1) {
  try {
    var sheet = SpreadsheetApp.getActive().getActiveSheet();
    var range = sheet.getRange(rangeA1);
    return range.getNumRows() * range.getNumColumns() > 500;
  } catch (e) {
    return true;
  }
}
