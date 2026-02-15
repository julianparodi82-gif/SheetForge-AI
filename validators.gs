var ALLOWED_OPS = [
  'getValue',
  'getValues',
  'getDisplayValue',
  'getDisplayValues',
  'getFormula',
  'getFormulas',
  'getNumberFormat',
  'getNumberFormats',
  'getBackground',
  'getBackgrounds',
  'getFontColor',
  'getFontColors',
  'getFontFamily',
  'getFontFamilies',
  'getFontSize',
  'getFontSizes',
  'getFontStyle',
  'getFontStyles',
  'getFontWeight',
  'getFontWeights',
  'getHorizontalAlignment',
  'getHorizontalAlignments',
  'getVerticalAlignment',
  'getVerticalAlignments',
  'getWrap',
  'getWraps',
  'getTextRotation',
  'getTextRotations',
  'getBorder',
  'getBorders',
  'getNote',
  'getNotes',
  'getDataValidation',
  'getDataValidations',
  'getA1Notation',
  'getRow',
  'getColumn',
  'getNumRows',
  'getNumColumns',
  'setValue',
  'setValues',
  'setDisplayValue',
  'setDisplayValues',
  'clearContent',
  'clearContents',
  'clear',
  'clearAll',
  'setFormula',
  'setFormulas',
  'setFormulaR1C1',
  'setFormulasR1C1',
  'createFormula',
  'dragFormula',
  'fillFormula',
  'setBackground',
  'setBackgrounds',
  'setBackgroundColor',
  'setFontColor',
  'setFontColors',
  'setFontWeight',
  'setFontWeights',
  'setFontSize',
  'setFontSizes',
  'setFontFamily',
  'setFontFamilies',
  'setFontStyle',
  'setFontStyles',
  'setFontLine',
  'setNumberFormat',
  'setNumberFormats',
  'setTextRotation',
  'setTextRotations',
  'setWrap',
  'setWraps',
  'setHorizontalAlignment',
  'setHorizontalAlignments',
  'setVerticalAlignment',
  'setVerticalAlignments',
  'setHorizontalAlign',
  'setVerticalAlign',
  'alignLeft',
  'alignCenter',
  'alignRight',
  'alignJustify',
  'alignTop',
  'alignMiddle',
  'alignBottom',
  'setAlignment',
  'setBorder',
  'setBorders',
  'paintCell',
  'paintRow',
  'paintColumn',
  'paintSheet',
  'setNote',
  'setNotes',
  'clearNote',
  'clearNotes',
  'getRange',
  'copyRange',
  'moveRange',
  'copyCells',
  'moveCells',
  'copyRows',
  'moveRows',
  'copyColumns',
  'moveColumns',
  'insertRows',
  'deleteRows',
  'insertColumns',
  'deleteColumns',
  'hideRows',
  'showRows',
  'hideColumns',
  'showColumns',
  'setRowHeight',
  'setRowHeights',
  'setColumnWidth',
  'setColumnWidths',
  'autoResizeRows',
  'autoResizeColumns',
  'mergeCells',
  'mergeAcross',
  'mergeVertically',
  'unmergeCells',
  'breakApart',
  'setFrozenRows',
  'setFrozenColumns',
  'sortRange',
  'removeDuplicates',
  'createFilter',
  'removeFilter',
  'setFilterCriteria',
  'clearFilterCriteria',
  'setDataValidation',
  'setDataValidations',
  'clearDataValidation',
  'clearDataValidations',
  'getConditionalFormatRules',
  'setConditionalFormatRules',
  'clearConditionalFormatRules',
  'createNamedRange',
  'updateNamedRange',
  'deleteNamedRange',
  'getNamedRanges',
  'createPivotTable',
  'updatePivotTable',
  'deletePivotTable',
  'createChart',
  'updateChart',
  'deleteChart',
  'getCharts',
  'renameSheet',
  'createSheet',
  'deleteSheet',
  'hideSheet',
  'showSheet',
  'duplicateSheet',
  'moveSheet',
  'setActiveSheet',
  'renameFile',
  'exportPdf',
  'exportSpreadsheetPdf',
  'exportXlsx',
  'exportCsv',
  'insertImageFromDrive',
  'insertImageFromUrl',
  'setImageFormula',
  'removeImage',
  'copyPasteValues',
  'copyPasteFormats',
  'copyPasteAll',
  'copyPasteFormulas',
  'driveFindFiles',
  'driveListFolderFiles',
  'driveCreateFolder',
  'driveMoveFile',
  'driveCopyFile',
  'driveTrashFile',
  'driveRestoreFile',
  'driveDeleteFile',
  'driveShareFile',
  'driveUnshareFile',
  'driveSetPermissions',
  'driveGetFileMetadata',
  'driveExportFile',
  'createTimeTrigger',
  'deleteTrigger',
  'listTriggers',
  'lock',
  'unlock',
  'getProperties',
  'setProperties',
  'deleteProperties',
  'cacheGet',
  'cachePut',
  'cacheRemove',
  'urlFetch'
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

function validatePlanReadOnly_(plan, flags) {
  var cloned;
  try {
    cloned = JSON.parse(JSON.stringify(plan));
  } catch (e) {
    return 'Plan inválido: no se pudo validar estructura.';
  }
  return validatePlan_(cloned, flags);
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
  expandColorFormula_(action, sheet || SpreadsheetApp.getActive().getActiveSheet());
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
  if ((action.op === 'setBackground' || action.op === 'setBackgroundColor') && !action.color) {
    return 'Color inválido: falta color';
  }
  if (action.op === 'setBackground' || action.op === 'setBackgroundColor') {
    action.color = normalizeColorValue_(action.color);
    if (!action.color) return 'Color inválido: ' + action.color;
  }
  if (action.op === 'setBackgrounds' && !Array.isArray(action.colors)) {
    return 'Color inválido: falta colors';
  }
  if (action.op === 'setBackgrounds' && Array.isArray(action.colors) && action.rangeA1) {
    var sheetForBgMatrix = sheet || SpreadsheetApp.getActive().getActiveSheet();
    var rangeForBgMatrix = sheetForBgMatrix.getRange(action.rangeA1);
    var expectedRows = rangeForBgMatrix.getNumRows();
    var expectedCols = rangeForBgMatrix.getNumColumns();
    action.colors = normalizeColorMatrixValues_(action.colors, expectedRows, expectedCols);
    if (!isStrictHexColorMatrix_(action.colors, expectedRows, expectedCols)) {
      return 'Color inválido: colors debe ser matriz ' + expectedRows + 'x' + expectedCols + ' con valores #RRGGBB';
    }
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
  normalizeCommonActionFields_(action);
  if (action.op === 'setBackgroundColor') {
    action.op = 'setBackground';
  }
  if (action.op === 'moveCells' || action.op === 'moveRows' || action.op === 'moveColumns') {
    action.op = 'moveRange';
  }
  if (action.op === 'copyCells' || action.op === 'copyRows' || action.op === 'copyColumns') {
    action.op = 'copyRange';
  }
  if (action.op === 'createFormula') {
    action.op = 'setFormula';
  }
  if (action.op === 'dragFormula' || action.op === 'fillFormula') {
    action.op = Array.isArray(action.formulas) ? 'setFormulas' : 'setFormula';
  }
  if (action.op === 'setHorizontalAlign') {
    action.op = 'setHorizontalAlignment';
  }
  if (action.op === 'setVerticalAlign') {
    action.op = 'setVerticalAlignment';
  }
  if (action.op === 'alignLeft') {
    action.op = 'setHorizontalAlignment';
    action.alignment = 'left';
  }
  if (action.op === 'alignCenter') {
    action.op = 'setHorizontalAlignment';
    action.alignment = 'center';
  }
  if (action.op === 'alignRight') {
    action.op = 'setHorizontalAlignment';
    action.alignment = 'right';
  }
  if (action.op === 'alignJustify') {
    action.op = 'setHorizontalAlignment';
    action.alignment = 'justify';
  }
  if (action.op === 'alignTop') {
    action.op = 'setVerticalAlignment';
    action.alignment = 'top';
  }
  if (action.op === 'alignMiddle') {
    action.op = 'setVerticalAlignment';
    action.alignment = 'middle';
  }
  if (action.op === 'alignBottom') {
    action.op = 'setVerticalAlignment';
    action.alignment = 'bottom';
  }
  if (action.op === 'setAlignment') {
    if (action.horizontal && !action.alignment) {
      action.op = 'setHorizontalAlignment';
      action.alignment = action.horizontal;
    } else if (action.vertical && !action.alignment) {
      action.op = 'setVerticalAlignment';
      action.alignment = action.vertical;
    }
  }
  if (action.op === 'paintCell') {
    action.op = 'setBackground';
    if (action.cell && !action.rangeA1) {
      action.rangeA1 = action.cell;
    }
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

function normalizeCommonActionFields_(action) {
  if (!action) return;
  if (!action.rangeA1 && typeof action.range === 'string') {
    action.rangeA1 = action.range;
  }
  if (!action.color && typeof action.backgroundColor === 'string') {
    action.color = action.backgroundColor;
  }
  if (!action.color && typeof action.bgColor === 'string') {
    action.color = action.bgColor;
  }
  if (action.op === 'setBackgrounds' && Array.isArray(action.backgrounds) && !Array.isArray(action.colors)) {
    action.colors = action.backgrounds;
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
  if (rangeObj) {
    var startRowRaw = rangeObj.startRow !== undefined ? rangeObj.startRow : rangeObj.startRowIndex;
    var endRowRaw = rangeObj.endRow !== undefined ? rangeObj.endRow : rangeObj.endRowIndex;
    var startColumnRaw = rangeObj.startColumn !== undefined ? rangeObj.startColumn : rangeObj.startColumnIndex;
    var endColumnRaw = rangeObj.endColumn !== undefined ? rangeObj.endColumn : rangeObj.endColumnIndex;
    if (startRowRaw === undefined || endRowRaw === undefined || startColumnRaw === undefined || endColumnRaw === undefined) return;
    var startRow = Number(startRowRaw) + 1;
    var startColumn = Number(startColumnRaw) + 1;
    var numRows = Number(endRowRaw) - Number(startRowRaw);
    var numColumns = Number(endColumnRaw) - Number(startColumnRaw);
    if (numRows > 0 && numColumns > 0) {
      action.rangeA1 = sheet.getRange(startRow, startColumn, numRows, numColumns).getA1Notation();
    }
  }
}

function expandColorFormula_(action, sheet) {
  if (!action || !sheet || action.op !== 'setBackground' || !action.colorFormula || !action.rangeA1) return;
  var match = action.colorFormula.match(/RGB\s*\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)/i);
  if (!match) return;
  var baseR = Math.max(0, Math.min(255, Number(match[1])));
  var baseG = Math.max(0, Math.min(255, Number(match[2])));
  var baseB = Math.max(0, Math.min(255, Number(match[3])));
  var range = sheet.getRange(action.rangeA1);
  var rows = range.getNumRows();
  var cols = range.getNumColumns();
  if (rows <= 1 && cols <= 1) return;
  var colors = [];
  for (var r = 0; r < rows; r++) {
    var row = [];
    var rValue = Math.max(0, Math.min(255, baseR + r));
    var hex = '#' + toHex_(rValue) + toHex_(baseG) + toHex_(baseB);
    for (var c = 0; c < cols; c++) row.push(hex);
    colors.push(row);
  }
  action.op = 'setBackgrounds';
  action.colors = colors;
  action.color = null;
}

function toHex_(n) {
  var hex = Number(n).toString(16);
  return hex.length === 1 ? '0' + hex : hex;
}

function requiresSheet_(op) {
  return [
    'setValue', 'setValues', 'setFormula', 'setFormulas', 'setBackground', 'setBackgrounds',
    'setBackgroundColor', 'paintCell', 'paintRow', 'paintColumn', 'paintSheet',
    'moveCells', 'moveRows', 'moveColumns', 'copyCells', 'copyRows', 'copyColumns',
    'createFormula', 'dragFormula', 'fillFormula',
    'setFontColor', 'setFontWeight', 'setFontSize', 'setFontFamily', 'setFontStyle', 'setFontLine',
    'setHorizontalAlignment', 'setVerticalAlignment', 'setHorizontalAlign', 'setVerticalAlign',
  'alignLeft', 'alignCenter', 'alignRight', 'alignJustify', 'alignTop', 'alignMiddle', 'alignBottom',
  'setAlignment', 'setWrap', 'setBorder', 'setBorders',
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
    'setHorizontalAlignment', 'setVerticalAlignment', 'setHorizontalAlign', 'setVerticalAlign',
  'alignLeft', 'alignCenter', 'alignRight', 'alignJustify', 'alignTop', 'alignMiddle', 'alignBottom',
  'setAlignment', 'setWrap', 'setBorder', 'setBorders',
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
  var safeFallback = fallbackColor || '#000000';
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

function isStrictHexColorMatrix_(colors, rows, cols) {
  if (!Array.isArray(colors) || colors.length !== rows) return false;
  for (var r = 0; r < rows; r++) {
    var row = colors[r];
    if (!Array.isArray(row) || row.length !== cols) return false;
    for (var c = 0; c < cols; c++) {
      if (typeof row[c] !== 'string' || !/^#[0-9a-fA-F]{6}$/.test(row[c])) return false;
    }
  }
  return true;
}

function normalizeColorMatrixValues_(colors, rows, cols) {
  if (!Array.isArray(colors)) return colors;
  var targetRows = Math.max(1, Number(rows) || 0);
  var targetCols = Math.max(1, Number(cols) || 0);
  var normalized = colors;

  if (!Array.isArray(colors[0])) {
    normalized = coerceFlatColorsToMatrix_(colors, targetRows || colors.length, targetCols || 1);
  } else if (targetRows && targetCols) {
    var sourceRows = colors.length;
    var sourceCols = Array.isArray(colors[0]) ? colors[0].length : 0;
    if (sourceRows !== targetRows || sourceCols !== targetCols) {
      var flatColors = flattenColorMatrix_(colors);
      normalized = coerceFlatColorsToMatrix_(flatColors, targetRows, targetCols);
    }
  }

  return normalized.map(function(row) {
    return Array.isArray(row)
      ? row.map(function(cell) { return normalizeColorValue_(cell); })
      : row;
  });
}

function flattenColorMatrix_(colors) {
  var flat = [];
  for (var r = 0; r < colors.length; r++) {
    var row = colors[r];
    if (Array.isArray(row)) {
      for (var c = 0; c < row.length; c++) flat.push(row[c]);
    } else {
      flat.push(row);
    }
  }
  return flat;
}

function coerceFlatColorsToMatrix_(colors, rows, cols) {
  var targetRows = Math.max(1, Number(rows) || 1);
  var targetCols = Math.max(1, Number(cols) || 1);
  var total = targetRows * targetCols;
  var flat = colors.slice();
  if (flat.length === 1 && total > 1) {
    var fill = [];
    for (var i = 0; i < total; i++) fill.push(flat[0]);
    flat = fill;
  }
  if (targetCols === 1 && flat.length === targetRows) {
    return flat.map(function(cell) { return [cell]; });
  }
  if (targetRows === 1 && flat.length === targetCols) {
    return [flat.slice(0, targetCols)];
  }
  var matrix = [];
  for (var r = 0; r < targetRows; r++) {
    var row = [];
    for (var c = 0; c < targetCols; c++) {
      var idx = r * targetCols + c;
      row.push(flat[Math.min(idx, flat.length - 1)]);
    }
    matrix.push(row);
  }
  return matrix;
}

function normalizeColorValue_(value) {
  if (typeof value === 'string') {
    var v = value.trim();
    if (/^#[0-9a-fA-F]{6}$/.test(v)) return v.toLowerCase();
    var shortHex = v.match(/^#([0-9a-fA-F]{3})$/);
    if (shortHex) {
      var sh = shortHex[1].toLowerCase();
      return '#' + sh[0] + sh[0] + sh[1] + sh[1] + sh[2] + sh[2];
    }
    var rgbMatch = v.match(/^rgba?\s*\(([^)]+)\)$/i);
    if (rgbMatch) {
      var parts = rgbMatch[1].split(',').map(function(part) { return part.trim(); });
      if (parts.length >= 3) {
        return '#' + toHexColorChannel_(parts[0]) + toHexColorChannel_(parts[1]) + toHexColorChannel_(parts[2]);
      }
    }
    var hslMatch = v.match(/^hsla?\s*\(([^)]+)\)$/i);
    if (hslMatch) {
      var hslParts = hslMatch[1].split(',').map(function(part) { return part.trim(); });
      if (hslParts.length >= 3) {
        var rgb = hslToRgb_(Number(hslParts[0]), parsePercentage_(hslParts[1]), parsePercentage_(hslParts[2]));
        return '#' + toHexColorChannel_(rgb.r) + toHexColorChannel_(rgb.g) + toHexColorChannel_(rgb.b);
      }
    }
    var named = {
      blanco: '#ffffff', white: '#ffffff',
      negro: '#000000', black: '#000000',
      gris: '#808080', gray: '#808080', grey: '#808080',
      rojo: '#ff0000', red: '#ff0000',
      verde: '#00ff00', green: '#00ff00',
      azul: '#0000ff', blue: '#0000ff',
      amarillo: '#ffff00', yellow: '#ffff00',
      naranja: '#ffa500', orange: '#ffa500',
      morado: '#800080', purple: '#800080'
    };
    var mapped = named[v.toLowerCase()];
    return mapped || null;
  }
  if (value && typeof value === 'object' && value.red !== undefined && value.green !== undefined && value.blue !== undefined) {
    return '#' + toHexColorChannel_(value.red) + toHexColorChannel_(value.green) + toHexColorChannel_(value.blue);
  }
  return null;
}

function toHexColorChannel_(n) {
  var raw = n;
  if (typeof raw === 'string') {
    var trimmed = raw.trim();
    if (/^-?\d+(\.\d+)?%$/.test(trimmed)) {
      raw = (parseFloat(trimmed) / 100) * 255;
    } else {
      raw = Number(trimmed);
    }
  }
  var num = Number(raw);
  if (!isFinite(num)) num = 0;
  if (num >= 0 && num <= 1) {
    num = num * 255;
  }
  num = Math.max(0, Math.min(255, num));
  var hex = Math.round(num).toString(16);
  return hex.length === 1 ? '0' + hex : hex;
}

function parsePercentage_(value) {
  var str = String(value || '').trim();
  if (/^-?\d+(\.\d+)?%$/.test(str)) {
    return Math.max(0, Math.min(1, parseFloat(str) / 100));
  }
  var num = Number(str);
  if (!isFinite(num)) return 0;
  if (num > 1) return Math.max(0, Math.min(1, num / 100));
  return Math.max(0, Math.min(1, num));
}

function hslToRgb_(h, s, l) {
  var hue = ((Number(h) % 360) + 360) % 360;
  var sat = Math.max(0, Math.min(1, Number(s)));
  var lig = Math.max(0, Math.min(1, Number(l)));
  var c = (1 - Math.abs(2 * lig - 1)) * sat;
  var x = c * (1 - Math.abs((hue / 60) % 2 - 1));
  var m = lig - c / 2;
  var r1 = 0;
  var g1 = 0;
  var b1 = 0;
  if (hue < 60) {
    r1 = c; g1 = x;
  } else if (hue < 120) {
    r1 = x; g1 = c;
  } else if (hue < 180) {
    g1 = c; b1 = x;
  } else if (hue < 240) {
    g1 = x; b1 = c;
  } else if (hue < 300) {
    r1 = x; b1 = c;
  } else {
    r1 = c; b1 = x;
  }
  return {
    r: (r1 + m) * 255,
    g: (g1 + m) * 255,
    b: (b1 + m) * 255
  };
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
