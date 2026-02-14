function scanProjectAndUpdateContext(options) {
  ensureSpecialSheets_();
  var mode = options && options.mode ? options.mode : 'quick';
  var ss = SpreadsheetApp.getActive();
  var context = buildContext_(ss, mode);
  writeContextLines_(context.lines);
  return { message: 'Contexto actualizado (' + mode + ').' };
}

function updateContextAfterCommit_(actions) {
  var sheetIds = extractAffectedSheetIds_(actions);
  var ss = SpreadsheetApp.getActive();
  var mode = 'quick';
  var existing = readContextLines_();
  var rebuilt = buildContext_(ss, mode, sheetIds, existing);
  writeContextLines_(rebuilt.lines);
}

function getContext_() {
  var ss = SpreadsheetApp.getActive();
  var activeSheet = ss.getActiveSheet();
  var selection = activeSheet.getSelection();
  var range = selection.getActiveRange();
  var limit = 5;
  var values = range.getValues().slice(0, limit).map(function (row) { return row.slice(0, limit); });
  var formulas = range.getFormulas().slice(0, limit).map(function (row) { return row.slice(0, limit); });
  var backgrounds = range.getBackgrounds().slice(0, limit).map(function (row) { return row.slice(0, limit); });
  return {
    meta: { spreadsheetId: ss.getId(), sheetName: activeSheet.getName() },
    selection: { values: values, formulas: formulas, backgrounds: backgrounds },
    contextFromSheet: readContextLines_().join('\n'),
    configBlocks: readConfigBlocks_()
  };
}

function buildContext_(ss, mode, onlySheetIds, existingLines) {
  var lines = [];
  var timezone = ss.getSpreadsheetTimeZone();
  var sheets = ss.getSheets();
  var includeValues = mode === 'full';
  lines.push('CONTEXTO_IA — Estado actual del Spreadsheet');
  lines.push(new Date().toISOString());
  lines.push(ss.getId());
  lines.push(ss.getName());
  lines.push(timezone);
  lines.push('modo_contexto (' + mode + ')');
  lines.push('indice_hojas (' + sheets.map(function (s) { return s.getName() + '(' + s.getSheetId() + ')'; }).join(',') + ')');
  lines.push('—');
  var configRanges = readConfigBlocks_();
  sheets.forEach(function (sheet) {
    if (onlySheetIds && onlySheetIds.length && onlySheetIds.indexOf(sheet.getSheetId()) === -1) {
      if (existingLines) {
        lines = lines.concat(extractSheetBlock_(existingLines, sheet.getSheetId()));
      }
      return;
    }
    lines = lines.concat(buildSheetBlock_(sheet, configRanges, includeValues));
  });
  return { lines: lines };
}

function buildSheetBlock_(sheet, configRanges, includeValues) {
  var lines = [];
  var sheetId = sheet.getSheetId();
  var dataRange = sheet.getDataRange();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var maxRows = sheet.getMaxRows();
  var maxColumns = sheet.getMaxColumns();
  lines.push('### BEGIN_SHEET_BLOCK::' + sheetId);
  lines.push('## HOJA: ' + sheet.getName());
  lines.push('- sheetId ' + sheetId);
  lines.push('- filas | columnas (max) ' + maxRows + ' | ' + maxColumns);
  lines.push('- filas | columnas (usadas) ' + lastRow + ' | ' + lastColumn);
  lines.push('- rango_usado ' + dataRange.getA1Notation());
  lines.push('- hoja_oculta ' + (sheet.isSheetHidden() ? 'si' : 'no'));
  lines.push('- filas_congeladas | columnas_congeladas ' + sheet.getFrozenRows() + ' | ' + sheet.getFrozenColumns());
  var ranges = configRanges.filter(function (r) { return r.sheetName === sheet.getName(); });
  lines.push('- rangos_clave ' + (ranges.map(function (r) { return r.rangeA1; }).join(', ') || 'desconocido'));
  lines.push('- filtros_activos ' + (sheet.getFilter() ? 'si' : 'no'));
  lines.push('- columnas_importantes ' + summarizeHeader_(sheet, lastColumn));
  lines.push('- celdas_con_formula ' + countFormulaCells_(sheet));
  lines.push('- celdas_con_validacion ' + countValidationCells_(sheet));
  lines.push('- celdas_con_nota ' + countNoteCells_(sheet));
  lines.push('- celdas_combinadas ' + countMergedCells_(sheet));
  lines.push('- reglas_formato_condicional ' + countConditionalRules_(sheet));
  var formats = sampleFormats_(sheet);
  lines.push('- formatos_relevantes ' + formats.join(', '));
  var formulas = sampleFormulas_(sheet);
  lines.push('- formulas_representativas ' + formulas.join(', '));
  lines.push('- dependencias posibles entre hojas/formulas (muestreo)');
  if (includeValues) {
    lines.push('- valores_muestra ' + JSON.stringify(sampleValues_(sheet)));
  }
  lines.push('- notas_operativas ' + (includeValues ? 'valores muestreados' : 'muestreo quick'));
  lines.push('### END_SHEET_BLOCK::' + sheetId);
  return lines;
}

function summarizeHeader_(sheet, lastColumn) {
  if (!lastColumn) return 'desconocido';
  var header = sheet.getRange(1, 1, 1, Math.min(lastColumn, 15)).getDisplayValues()[0] || [];
  var cleaned = header.filter(function (v) { return v; });
  return cleaned.length ? cleaned.join(' | ') : 'sin encabezados detectados';
}

function countFormulaCells_(sheet) {
  var formulas = sheet.getDataRange().getFormulas();
  var count = 0;
  formulas.forEach(function (row) {
    row.forEach(function (cell) { if (cell) count += 1; });
  });
  return count;
}

function countValidationCells_(sheet) {
  var vals = sheet.getDataRange().getDataValidations();
  var count = 0;
  vals.forEach(function (row) {
    row.forEach(function (cell) { if (cell) count += 1; });
  });
  return count;
}

function countNoteCells_(sheet) {
  var notes = sheet.getDataRange().getNotes();
  var count = 0;
  notes.forEach(function (row) {
    row.forEach(function (cell) { if (cell) count += 1; });
  });
  return count;
}

function countMergedCells_(sheet) {
  var ranges = sheet.getDataRange().getMergedRanges();
  return ranges ? ranges.length : 0;
}

function countConditionalRules_(sheet) {
  return sheet.getConditionalFormatRules().length;
}

function sampleValues_(sheet) {
  var range = sheet.getDataRange();
  var values = range.getDisplayValues();
  return values.slice(0, 5).map(function (row) { return row.slice(0, 5); });
}

function sampleFormulas_(sheet) {
  var range = sheet.getDataRange();
  var formulas = range.getFormulas();
  var collected = [];
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      if (formulas[i][j]) {
        collected.push(formulas[i][j]);
      }
      if (collected.length >= 10) return collected;
    }
  }
  return collected.length ? collected : ['desconocido'];
}

function sampleFormats_(sheet) {
  var range = sheet.getDataRange();
  var backgrounds = range.getBackgrounds();
  var format = range.getNumberFormats();
  var colors = {};
  var formats = {};
  for (var i = 0; i < backgrounds.length; i++) {
    for (var j = 0; j < backgrounds[i].length; j++) {
      colors[backgrounds[i][j]] = true;
      formats[format[i][j]] = true;
    }
  }
  var topColors = Object.keys(colors).slice(0, 5);
  var topFormats = Object.keys(formats).slice(0, 5);
  return topColors.concat(topFormats);
}

function extractSheetBlock_(lines, sheetId) {
  var start = '### BEGIN_SHEET_BLOCK::' + sheetId;
  var end = '### END_SHEET_BLOCK::' + sheetId;
  var inside = false;
  var block = [];
  lines.forEach(function (line) {
    if (line === start) {
      inside = true;
    }
    if (inside) {
      block.push(line);
    }
    if (line === end) {
      inside = false;
    }
  });
  return block.length ? block : buildSheetBlock_(SpreadsheetApp.getActive().getSheetById(sheetId), [], false);
}

function readContextLines_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('CONTEXTO_IA');
  if (!sheet) return [];
  var values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  return values.map(function (row) { return row[0]; });
}

function readConfigBlocks_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('CONFIG');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  return values.map(function (row) {
    return {
      label: row[0],
      sheetName: row[1],
      rangeA1: row[2],
      includeValues: row[3],
      includeFormulas: row[4],
      includeFormats: row[5]
    };
  }).filter(function (row) { return row.sheetName && row.rangeA1; });
}

function extractAffectedSheetIds_(actions) {
  var ss = SpreadsheetApp.getActive();
  var ids = [];
  actions.forEach(function (action) {
    if (!action.sheetName) return;
    var sheet = ss.getSheetByName(action.sheetName);
    if (sheet) ids.push(sheet.getSheetId());
  });
  return ids.filter(function (v, i, a) { return a.indexOf(v) === i; });
}
