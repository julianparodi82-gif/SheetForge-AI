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
    logEvent_('apply', '', plan, 'ok', 'Aplicado', affected.join(', '), duration);
    updateContextAfterCommit_(plan.actions);
    return { message: 'Aplicado correctamente.' };
  } catch (e) {
    var durationErr = Date.now() - start;
    logEvent_('apply', '', plan, 'error', e.message, affected.join(', '), durationErr);
    return { error: e.message };
  }
}

function executeAction_(action) {
  var ss = SpreadsheetApp.getActive();
  if (action.op === 'setValue') {
    var sheet = ss.getSheetByName(action.sheetName);
    sheet.getRange(action.rangeA1).setValue(action.value);
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
