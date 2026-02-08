var ALLOWED_OPS = [
  'setValue', 'setFormula', 'setFormulas', 'setBackground', 'setNumberFormat', 'clearContent',
  'copyPasteValues', 'moveRange', 'insertRows', 'deleteRows', 'renameSheet', 'createSheet',
  'deleteSheet', 'setColumnWidth', 'setRowHeight', 'renameFile', 'exportPdf',
  'exportSpreadsheetPdf', 'insertImageFromDrive', 'setImageFormula'
];

function validatePlan_(plan, flags) {
  if (!plan || !plan.meta || !Array.isArray(plan.actions)) {
    return 'Plan inválido.';
  }
  if (ALLOWED_OPS.indexOf(plan.actions[0] && plan.actions[0].op) === -1 && plan.actions.length > 0) {
    return 'Operación no permitida.';
  }
  for (var i = 0; i < plan.actions.length; i++) {
    var action = plan.actions[i];
    if (!action.op || ALLOWED_OPS.indexOf(action.op) === -1) {
      return 'Operación no permitida: ' + action.op;
    }
    var v = validateAction_(action, flags);
    if (v) return v;
  }
  return null;
}

function validateAction_(action, flags) {
  if (flags && flags.safeMode) {
    if (action.op === 'deleteSheet' || action.op === 'renameFile') {
      return 'Safe Mode bloquea ' + action.op;
    }
    if (action.op === 'deleteRows' && action.numRows && action.numRows > 50) {
      return 'Safe Mode bloquea deleteRows masivo';
    }
    if (action.op === 'clearContent' && action.rangeA1 && rangeIsLarge_(action.rangeA1)) {
      return 'Safe Mode bloquea clearContent masivo';
    }
  }
  if (action.rangeA1 && !isValidA1_(action.rangeA1)) {
    return 'Rango inválido: ' + action.rangeA1;
  }
  if (action.color && !/^#[0-9a-fA-F]{6}$/.test(action.color)) {
    return 'Color inválido: ' + action.color;
  }
  return null;
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
