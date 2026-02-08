var ALLOWED_OPS = [
  'setValue', 'setValues', 'setFormula', 'setFormulas', 'setBackground', 'setBackgrounds',
  'setFontColor', 'setFontWeight', 'setFontSize', 'setHorizontalAlignment', 'setVerticalAlignment',
  'setWrap', 'setBorder', 'setNumberFormat', 'clearContent', 'copyPasteValues', 'moveRange',
  'insertRows', 'deleteRows', 'insertColumns', 'deleteColumns', 'hideRows', 'showRows',
  'hideColumns', 'showColumns', 'renameSheet', 'createSheet', 'deleteSheet', 'hideSheet',
  'showSheet', 'setColumnWidth', 'setRowHeight', 'renameFile', 'exportPdf',
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
      var activeSheet = SpreadsheetApp.getActive().getActiveSheet();
      action.sheetName = activeSheet ? activeSheet.getName() : null;
      if (!action.sheetName) {
        return 'Operación no permitida: falta sheetName';
      }
    }
    var sheet = SpreadsheetApp.getActive().getSheetByName(action.sheetName);
    if (!sheet) {
      return 'Hoja no encontrada: ' + action.sheetName;
    }
  }
  if (requiresRange_(action.op) && !action.rangeA1) {
    var activeRange = SpreadsheetApp.getActive().getActiveRange();
    action.rangeA1 = activeRange ? activeRange.getA1Notation() : null;
    if (!action.rangeA1 && action.op === 'setBackground') {
      action.rangeA1 = 'A:A';
    }
    if (!action.rangeA1) {
      return 'Rango inválido: falta rangeA1';
    }
  }
  if (action.op === 'setValues' && !Array.isArray(action.values)) {
    return 'Rango inválido: falta values';
  }
  if (action.op === 'setValues' && action.rangeA1 && Array.isArray(action.values)) {
    var valuesRows = action.values.length;
    var valuesCols = valuesRows ? action.values[0].length : 0;
    if (valuesRows && valuesCols) {
      var sheetForValues = SpreadsheetApp.getActive().getSheetByName(action.sheetName);
      var startRange = sheetForValues.getRange(action.rangeA1);
      var targetRange = sheetForValues.getRange(
        startRange.getRow(),
        startRange.getColumn(),
        valuesRows,
        valuesCols
      );
      action.rangeA1 = targetRange.getA1Notation();
    }
  }
  if (action.op === 'setBackground' && !action.color) {
    action.color = '#ffeb3b';
  }
  if (action.op === 'setBackgrounds' && !Array.isArray(action.colors)) {
    return 'Color inválido: falta colors';
  }
  if (action.op === 'setBorder' && action.border === undefined) {
    return 'Borde inválido: falta border';
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
  if (action.op === 'insertRows' || action.op === 'deleteRows') {
    if (!action.rowIndex || !action.numRows) {
      return 'Rango inválido: falta rowIndex/numRows';
    }
  }
  if (action.op === 'insertColumns' || action.op === 'deleteColumns') {
    if (!action.columnIndex || !action.numColumns) {
      return 'Rango inválido: falta columnIndex/numColumns';
    }
  }
  if (action.op === 'hideRows' || action.op === 'showRows') {
    if (!action.rowIndex || !action.numRows) {
      return 'Rango inválido: falta rowIndex/numRows';
    }
  }
  if (action.op === 'hideColumns' || action.op === 'showColumns') {
    if (!action.columnIndex || !action.numColumns) {
      return 'Rango inválido: falta columnIndex/numColumns';
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
    'setValue', 'setValues', 'setFormula', 'setFormulas', 'setBackground', 'setBackgrounds',
    'setFontColor', 'setFontWeight', 'setFontSize', 'setHorizontalAlignment', 'setVerticalAlignment',
    'setWrap', 'setBorder', 'setNumberFormat', 'clearContent', 'copyPasteValues', 'moveRange',
    'insertRows', 'deleteRows', 'insertColumns', 'deleteColumns', 'hideRows', 'showRows',
    'hideColumns', 'showColumns', 'renameSheet', 'createSheet', 'deleteSheet', 'hideSheet',
    'showSheet', 'setColumnWidth', 'setRowHeight', 'exportPdf', 'insertImageFromDrive',
    'setImageFormula'
  ].indexOf(op) !== -1;
}

function requiresRange_(op) {
  return [
    'setValue', 'setValues', 'setFormula', 'setFormulas', 'setBackground', 'setBackgrounds',
    'setFontColor', 'setFontWeight', 'setFontSize', 'setHorizontalAlignment', 'setVerticalAlignment',
    'setWrap', 'setBorder', 'setNumberFormat', 'clearContent', 'setImageFormula'
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
