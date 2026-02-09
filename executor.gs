function applyPlan(plan, flags) {
  ensureSpecialSheets_();
  var start = Date.now();
  var validationError = validatePlan_(plan, flags);
  if (validationError) {
    return { error: validationError };
  }
  if (flags && flags.dryRun) {
    return { error: 'Dry Run activo. Desactiva para aplicar.' };
  }
  var affected = [];
  try {
    for (var i = 0; i < plan.actions.length; i++) {
      var action = plan.actions[i];
      executeAction_(action);
      affected.push(action.sheetName || 'spreadsheet');
    }
    var duration = Date.now() - start;
    var summary = buildActionSummary_(plan.actions);
    logEvent_('apply', '', plan, 'ok', summary, affected.join(', '), duration);
    updateContextAfterCommit_(plan.actions);
    return { message: 'Aplicado correctamente.' };
  } catch (e) {
    var durationErr = Date.now() - start;
    logEvent_('apply', '', plan, 'error', 'Error al aplicar: ' + e.message, affected.join(', '), durationErr);
    return { error: e.message };
  }
}

function buildActionSummary_(actions) {
  if (!Array.isArray(actions) || actions.length === 0) {
    return 'Aplicado: sin acciones detalladas.';
  }
  var parts = actions.slice(0, 3).map(function(action) {
    var op = action.op || 'acción';
    var sheet = action.sheetName ? ' en ' + action.sheetName : '';
    var range = action.rangeA1 || action.targetA1 || action.sourceA1 || action.cell;
    var location = range ? ' (' + range + ')' : '';
    return op + sheet + location;
  });
  var extra = actions.length > 3 ? ' y ' + (actions.length - 3) + ' más' : '';
  return 'Aplicado: ' + parts.join('; ') + extra + '.';
}

function executeAction_(action) {
  var ss = SpreadsheetApp.getActive();
  if (action.op === 'setValue') {
    var sheet = ss.getSheetByName(action.sheetName);
    sheet.getRange(action.rangeA1).setValue(action.value);
    return;
  }
  if (action.op === 'setValues') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setValues(action.values);
    return;
  }
  if (action.op === 'setFormula') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setFormula(action.formula);
    return;
  }
  if (action.op === 'setFormulas') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setFormulas(action.formulas);
    return;
  }
  if (action.op === 'setBackground') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setBackground(action.color);
    return;
  }
  if (action.op === 'setBackgroundColor') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setBackground(action.color);
    return;
  }
  if (action.op === 'setBackgrounds') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setBackgrounds(action.colors);
    return;
  }
  if (action.op === 'setFontColor') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setFontColor(action.color);
    return;
  }
  if (action.op === 'setFontWeight') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setFontWeight(action.weight);
    return;
  }
  if (action.op === 'setFontSize') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setFontSize(action.size);
    return;
  }
  if (action.op === 'setFontFamily') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setFontFamily(action.family);
    return;
  }
  if (action.op === 'setFontStyle') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setFontStyle(action.style);
    return;
  }
  if (action.op === 'setFontLine') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setFontLine(action.line);
    return;
  }
  if (action.op === 'setHorizontalAlignment') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setHorizontalAlignment(action.alignment);
    return;
  }
  if (action.op === 'setVerticalAlignment') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setVerticalAlignment(action.alignment);
    return;
  }
  if (action.op === 'setWrap') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setWrap(!!action.wrap);
    return;
  }
  if (action.op === 'setTextRotation') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setTextRotation(action.rotation);
    return;
  }
  if (action.op === 'setBorder' || action.op === 'setBorders') {
    var range = ss.getSheetByName(action.sheetName).getRange(action.rangeA1);
    var border = action.border;
    if (typeof border === 'object') {
      range.setBorder(
        !!border.top,
        !!border.left,
        !!border.bottom,
        !!border.right,
        !!border.vertical,
        !!border.horizontal
      );
    } else {
      range.setBorder(
        !!border,
        !!border,
        !!border,
        !!border,
        !!border,
        !!border
      );
    }
    return;
  }
  if (action.op === 'setNumberFormat') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setNumberFormat(action.format);
    return;
  }
  if (action.op === 'clearContent') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).clearContent();
    return;
  }
  if (action.op === 'copyPasteValues') {
    var src = ss.getSheetByName(action.sheetName).getRange(action.sourceA1);
    var dest = ss.getSheetByName(action.sheetName).getRange(action.targetA1);
    src.copyTo(dest, { contentsOnly: true });
    return;
  }
  if (action.op === 'copyPasteFormats') {
    var srcFormats = ss.getSheetByName(action.sheetName).getRange(action.sourceA1);
    var destFormats = ss.getSheetByName(action.sheetName).getRange(action.targetA1);
    srcFormats.copyTo(destFormats, { formatOnly: true });
    return;
  }
  if (action.op === 'copyPasteFormulas') {
    var srcFormulas = ss.getSheetByName(action.sheetName).getRange(action.sourceA1);
    var destFormulas = ss.getSheetByName(action.sheetName).getRange(action.targetA1);
    srcFormulas.copyTo(destFormulas, { contentsOnly: false });
    return;
  }
  if (action.op === 'copyPasteAll') {
    var srcAll = ss.getSheetByName(action.sheetName).getRange(action.sourceA1);
    var destAll = ss.getSheetByName(action.sheetName).getRange(action.targetA1);
    srcAll.copyTo(destAll);
    return;
  }
  if (action.op === 'copyRange') {
    var srcRange = ss.getSheetByName(action.sheetName).getRange(action.sourceA1);
    var destRange = ss.getSheetByName(action.sheetName).getRange(action.targetA1);
    srcRange.copyTo(destRange);
    return;
  }
  if (action.op === 'createPivotTable') {
    var pivotSheet = ss.getSheetByName(action.sheetName);
    var sourceRange = pivotSheet.getRange(action.sourceA1);
    var targetRange = pivotSheet.getRange(action.targetA1);
    var pivot = sourceRange.createPivotTable(targetRange);
    if (Array.isArray(action.rowGroups)) {
      action.rowGroups.forEach(function(group) {
        if (group && group.sourceColumn) {
          pivot.addRowGroup(group.sourceColumn);
        }
      });
    }
    if (Array.isArray(action.columnGroups)) {
      action.columnGroups.forEach(function(group) {
        if (group && group.sourceColumn) {
          pivot.addColumnGroup(group.sourceColumn);
        }
      });
    }
    if (Array.isArray(action.values)) {
      action.values.forEach(function(value) {
        if (value && value.sourceColumn && value.summarizeFunction) {
          pivot.addPivotValue(value.sourceColumn, value.summarizeFunction);
        }
      });
    }
    return;
  }
  if (action.op === 'moveRange') {
    var sheetMove = ss.getSheetByName(action.sheetName);
    var moveRange = sheetMove.getRange(action.sourceA1);
    moveRange.moveTo(sheetMove.getRange(action.targetA1));
    return;
  }
  if (action.op === 'insertRows') {
    ss.getSheetByName(action.sheetName).insertRows(action.rowIndex, action.numRows);
    return;
  }
  if (action.op === 'deleteRows') {
    ss.getSheetByName(action.sheetName).deleteRows(action.rowIndex, action.numRows);
    return;
  }
  if (action.op === 'insertColumns') {
    ss.getSheetByName(action.sheetName).insertColumns(action.columnIndex, action.numColumns);
    return;
  }
  if (action.op === 'deleteColumns') {
    ss.getSheetByName(action.sheetName).deleteColumns(action.columnIndex, action.numColumns);
    return;
  }
  if (action.op === 'hideRows') {
    ss.getSheetByName(action.sheetName).hideRows(action.rowIndex, action.numRows);
    return;
  }
  if (action.op === 'showRows') {
    ss.getSheetByName(action.sheetName).showRows(action.rowIndex, action.numRows);
    return;
  }
  if (action.op === 'hideColumns') {
    ss.getSheetByName(action.sheetName).hideColumns(action.columnIndex, action.numColumns);
    return;
  }
  if (action.op === 'showColumns') {
    ss.getSheetByName(action.sheetName).showColumns(action.columnIndex, action.numColumns);
    return;
  }
  if (action.op === 'renameSheet') {
    ss.getSheetByName(action.sheetName).setName(action.newName);
    return;
  }
  if (action.op === 'createSheet') {
    ss.insertSheet(action.sheetName);
    return;
  }
  if (action.op === 'deleteSheet') {
    var del = ss.getSheetByName(action.sheetName);
    ss.deleteSheet(del);
    return;
  }
  if (action.op === 'hideSheet') {
    ss.getSheetByName(action.sheetName).hideSheet();
    return;
  }
  if (action.op === 'showSheet') {
    ss.getSheetByName(action.sheetName).showSheet();
    return;
  }
  if (action.op === 'setColumnWidth') {
    ss.getSheetByName(action.sheetName).setColumnWidth(action.columnIndex, action.width);
    return;
  }
  if (action.op === 'setRowHeight') {
    ss.getSheetByName(action.sheetName).setRowHeight(action.rowIndex, action.height);
    return;
  }
  if (action.op === 'renameFile') {
    DriveApp.getFileById(ss.getId()).setName(action.name);
    return;
  }
  if (action.op === 'exportPdf') {
    exportSheetPdf_(action.sheetName, action.filename);
    return;
  }
  if (action.op === 'exportSpreadsheetPdf') {
    exportSpreadsheetPdf_(action.filename);
    return;
  }
  if (action.op === 'insertImageFromDrive') {
    insertImageFromDrive_(action);
    return;
  }
  if (action.op === 'setImageFormula') {
    ss.getSheetByName(action.sheetName).getRange(action.rangeA1).setFormula(action.formula);
  }
}
