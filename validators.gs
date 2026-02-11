var ALLOWED_OPS = [
  'setValue', 'setValues', 'setFormula', 'setFormulas', 'setBackground', 'setBackgrounds',
  'setBackgroundColor',
  'paintCell', 'paintRow', 'paintColumn', 'paintSheet',
  'moveCells', 'moveRows', 'moveColumns',
  'copyCells', 'copyRows', 'copyColumns',
  'createFormula', 'dragFormula', 'fillFormula',
  'setFontColor', 'setFontWeight', 'setFontSize', 'setFontFamily', 'setFontStyle', 'setFontLine',
  'setHorizontalAlignment', 'setVerticalAlignment', 'setWrap', 'setBorder', 'setBorders',
  'setTextRotation',
  'setNumberFormat', 'clearContent', 'copyPasteValues', 'copyPasteFormats', 'copyPasteAll',
  'copyPasteFormulas', 'copyRange', 'moveRange', 'createPivotTable',
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
    normalizeAction_(action);
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
    normalizeRangeObject_(action, sheet);
    expandPaintTargets_(action, sheet);
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
  if (action.op === 'setBackgrounds' && Array.isArray(action.colors) && action.colors.length) {
    normalizeBackgroundRangeFromColors_(action, sheet || SpreadsheetApp.getActive().getActiveSheet());
  }
  if (action.op === 'setFormula' && !action.formula) {
    action.formula = '';
  }
  if (action.op === 'setFormulas' && !Array.isArray(action.formulas)) {
    var sheetForFormulas = sheet || SpreadsheetApp.getActive().getActiveSheet();
    var rangeForFormulas = sheetForFormulas.getRange(action.rangeA1);
    action.formulas = buildFillMatrix_(rangeForFormulas.getNumRows(), rangeForFormulas.getNumColumns(), '');
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
  if (action.op === 'setBackgroundColor' && !action.color) {
    action.color = '#ffeb3b';
  }
  if (action.op === 'setBackgrounds' && !Array.isArray(action.colors)) {
    var sheetForColors = sheet || SpreadsheetApp.getActive().getActiveSheet();
    var rangeForColors = sheetForColors.getRange(action.rangeA1);
    var fillColor = action.color || '#ffeb3b';
    action.colors = buildFillMatrix_(rangeForColors.getNumRows(), rangeForColors.getNumColumns(), fillColor);
  }
  if (action.op === 'setBackgrounds' && Array.isArray(action.colors) && action.rangeA1) {
    var sheetForBgMatrix = sheet || SpreadsheetApp.getActive().getActiveSheet();
    var rangeForBgMatrix = sheetForBgMatrix.getRange(action.rangeA1);
    action.colors = normalizeColorMatrix_(action.colors, rangeForBgMatrix.getNumRows(), rangeForBgMatrix.getNumColumns(), action.color || '#ffeb3b');
  }
  if (action.op === 'setTextRotation' && action.rotation === undefined) {
    action.rotation = 0;
  }
  if (action.op === 'createPivotTable') {
    if (!action.sourceA1 || !action.targetA1) {
      return 'Rango inválido: falta sourceA1/targetA1';
    }
  }
  if (action.op === 'setBorder' || action.op === 'setBorders') {
    if (action.border === undefined) {
      action.border = true;
    }
  }
  if (action.op === 'setFontFamily' && !action.family) {
    action.family = 'Arial';
  }
  if (action.op === 'setFontStyle' && !action.style) {
    action.style = 'normal';
  }
  if (action.op === 'setFontLine' && !action.line) {
    action.line = 'none';
  }
  if (action.op === 'setFontColor' && !action.color) {
    action.color = '#000000';
  }
  if (action.op === 'setFontWeight' && !action.weight) {
    action.weight = 'normal';
  }
  if (action.op === 'setFontSize' && !action.size) {
    action.size = 10;
  }
  if (action.op === 'setHorizontalAlignment' && !action.alignment) {
    action.alignment = 'left';
  }
  if (action.op === 'setVerticalAlignment' && !action.alignment) {
    action.alignment = 'top';
  }
  if (action.op === 'setWrap' && action.wrap === undefined) {
    action.wrap = false;
  }
  if (action.op === 'setNumberFormat' && !action.format) {
    action.format = '@';
  }
  if (action.op === 'copyPasteValues') {
    if (!action.sourceA1 || !action.targetA1) {
      return 'Rango inválido: falta sourceA1/targetA1';
    }
  }
  if (action.op === 'copyRange') {
    if (!action.sourceA1 || !action.targetA1) {
      return 'Rango inválido: falta sourceA1/targetA1';
    }
  }
  if (action.op === 'copyPasteFormats' || action.op === 'copyPasteAll' || action.op === 'copyPasteFormulas') {
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

function normalizeAction_(action) {
  if (!action || !action.op) return;
  if (action.op === 'setBackgroundColor') {
    action.op = 'setBackground';
    return;
  }
  if (action.op === 'moveCells' || action.op === 'moveRows' || action.op === 'moveColumns') {
    action.op = 'moveRange';
    return;
  }
  if (action.op === 'copyCells' || action.op === 'copyRows' || action.op === 'copyColumns') {
    action.op = 'copyRange';
    return;
  }
  if (action.op === 'createFormula') {
    action.op = 'setFormula';
    return;
  }
  if (action.op === 'dragFormula' || action.op === 'fillFormula') {
    action.op = Array.isArray(action.formulas) ? 'setFormulas' : 'setFormula';
    return;
  }
  if (action.op === 'paintCell') {
    action.op = 'setBackground';
    if (action.cell && !action.rangeA1) {
      action.rangeA1 = action.cell;
    }
    return;
  }
  if (action.op === 'paintRow' || action.op === 'paintColumn' || action.op === 'paintSheet') {
    if (action.op === 'paintRow') {
      action.paintTarget = 'row';
    }
    if (action.op === 'paintColumn') {
      action.paintTarget = 'column';
    }
    if (action.op === 'paintSheet') {
      action.paintTarget = 'sheet';
    }
    action.op = 'setBackground';
  }
}

function expandPaintTargets_(action, sheet) {
  if (!action || !sheet || action.op !== 'setBackground') return;
  if (action.rangeA1) return;
  if (action.rowIndex) {
    action.rangeA1 = sheet.getRange(action.rowIndex, 1, 1, sheet.getMaxColumns()).getA1Notation();
    return;
  }
  if (action.columnIndex) {
    action.rangeA1 = sheet.getRange(1, action.columnIndex, sheet.getMaxRows(), 1).getA1Notation();
    return;
  }
  if (action.paintTarget === 'sheet' || action.paintSheet) {
    action.rangeA1 = sheet.getDataRange().getA1Notation();
  }
}

function normalizeRangeObject_(action, sheet) {
  if (!action || !sheet || action.rangeA1 || !action.range) return;
  var rangeObj = action.range;
  if (rangeObj && rangeObj.startRow !== undefined && rangeObj.endRow !== undefined &&
      rangeObj.startColumn !== undefined && rangeObj.endColumn !== undefined) {
    var startRow = rangeObj.startRow + 1;
    var startColumn = rangeObj.startColumn + 1;
    var numRows = rangeObj.endRow - rangeObj.startRow;
    var numColumns = rangeObj.endColumn - rangeObj.startColumn;
    if (numRows > 0 && numColumns > 0) {
      action.rangeA1 = sheet.getRange(startRow, startColumn, numRows, numColumns).getA1Notation();
    }
  }
}

function requiresSheet_(op) {
  return [
    'setValue', 'setValues', 'setFormula', 'setFormulas', 'setBackground', 'setBackgrounds',
    'setBackgroundColor', 'paintCell', 'paintRow', 'paintColumn', 'paintSheet',
    'moveCells', 'moveRows', 'moveColumns', 'copyCells', 'copyRows', 'copyColumns',
    'createFormula', 'dragFormula', 'fillFormula',
    'setFontColor', 'setFontWeight', 'setFontSize', 'setFontFamily', 'setFontStyle', 'setFontLine',
    'setHorizontalAlignment', 'setVerticalAlignment', 'setWrap', 'setBorder', 'setBorders',
    'setTextRotation',
    'setNumberFormat', 'clearContent', 'copyPasteValues', 'copyPasteFormats', 'copyPasteAll',
    'copyPasteFormulas', 'copyRange', 'moveRange', 'createPivotTable',
    'insertRows', 'deleteRows', 'insertColumns', 'deleteColumns', 'hideRows', 'showRows',
    'hideColumns', 'showColumns', 'renameSheet', 'createSheet', 'deleteSheet', 'hideSheet',
    'showSheet', 'setColumnWidth', 'setRowHeight', 'exportPdf', 'insertImageFromDrive',
    'setImageFormula'
  ].indexOf(op) !== -1;
}

function requiresRange_(op) {
  return [
    'setValue', 'setValues', 'setFormula', 'setFormulas', 'setBackground', 'setBackgrounds',
    'setBackgroundColor', 'paintCell', 'createFormula', 'dragFormula', 'fillFormula',
    'setFontColor', 'setFontWeight', 'setFontSize', 'setFontFamily', 'setFontStyle', 'setFontLine',
    'setHorizontalAlignment', 'setVerticalAlignment', 'setWrap', 'setBorder', 'setBorders',
    'setTextRotation',
    'setNumberFormat', 'clearContent', 'setImageFormula'
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

function buildFillMatrix_(rows, cols, value) {
  var matrix = [];
  for (var r = 0; r < rows; r++) {
    var row = [];
    for (var c = 0; c < cols; c++) {
      row.push(value);
    }
    matrix.push(row);
  }
  return matrix;
}

function normalizeColorMatrix_(inputColors, rows, cols, fallbackColor) {
  var safeFallback = fallbackColor || '#ffeb3b';
  if (!Array.isArray(inputColors) || rows <= 0 || cols <= 0) {
    return buildFillMatrix_(rows, cols, safeFallback);
  }

  var is2D = Array.isArray(inputColors[0]);
  if (is2D) {
    var matrix = [];
    for (var r = 0; r < rows; r++) {
      var srcRow = inputColors[Math.min(r, inputColors.length - 1)] || [];
      var newRow = [];
      for (var c = 0; c < cols; c++) {
        var srcCell = srcRow[Math.min(c, srcRow.length - 1)];
        newRow.push((typeof srcCell === 'string' && /^#[0-9a-fA-F]{6}$/.test(srcCell)) ? srcCell : safeFallback);
      }
      matrix.push(newRow);
    }
    return matrix;
  }

  var palette = inputColors.filter(function(color) {
    return typeof color === 'string' && /^#[0-9a-fA-F]{6}$/.test(color);
  });
  if (palette.length === 0) {
    return buildFillMatrix_(rows, cols, safeFallback);
  }

  var paletteMatrix = [];
  for (var pr = 0; pr < rows; pr++) {
    var paletteRow = [];
    for (var pc = 0; pc < cols; pc++) {
      var index = (pr * cols + pc) % palette.length;
      paletteRow.push(palette[index]);
    }
    paletteMatrix.push(paletteRow);
  }
  return paletteMatrix;
}

function normalizeBackgroundRangeFromColors_(action, sheet) {
  if (!action || !sheet || !Array.isArray(action.colors) || action.colors.length === 0) return;
  var rows = action.colors.length;
  var cols = Array.isArray(action.colors[0]) ? action.colors[0].length : action.colors.length;
  if (!rows || !cols) return;

  if (!action.rangeA1) {
    var activeRange = SpreadsheetApp.getActive().getActiveRange();
    if (!activeRange) return;
    action.rangeA1 = sheet.getRange(activeRange.getRow(), activeRange.getColumn(), rows, cols).getA1Notation();
    return;
  }

  try {
    var current = sheet.getRange(action.rangeA1);
    if (current.getNumRows() === 1 && current.getNumColumns() === 1 && (rows > 1 || cols > 1)) {
      action.rangeA1 = sheet.getRange(current.getRow(), current.getColumn(), rows, cols).getA1Notation();
    }
  } catch (e) {
    // no-op: validation will catch invalid ranges later.
  }
}
