function buildPlan(prompt, flags) {
  ensureSpecialSheets_();
  var trimmed = (prompt || '').trim();
  if (!trimmed) return { error: 'Prompt vacío.' };
  if (trimmed.indexOf('/') === 0) {
    var plan = commandToPlan_(trimmed, flags);
    if (plan.error) return plan;
    var summary = buildSummary_(plan);
    return { plan: plan, summary: summary };
  }
  var aiResult = buildPlanWithAi_(trimmed, flags, null);
  if (aiResult.error) return aiResult;
  return aiResult;
}
function editPlan(editPrompt, currentPlan, flags) {
  ensureSpecialSheets_();
  var trimmed = (editPrompt || '').trim();
  if (!currentPlan) return { error: 'Sin plan actual.' };
  if (trimmed.indexOf('/') === 0) {
    var updated = editPlanWithCommand_(trimmed, currentPlan);
    if (updated.error) return updated;
    var summary = buildSummary_(updated);
    return { plan: updated, summary: summary };
  }
  var aiResult = buildPlanWithAi_(trimmed, flags, currentPlan);
  if (aiResult.error) return aiResult;
  return aiResult;
}
function buildPlanWithAi_(prompt, flags, currentPlan) {
  var props = PropertiesService.getDocumentProperties();
  var userProps = PropertiesService.getUserProperties();
  var endpoint = props.getProperty('AI_ENDPOINT') || '';
  var apiKey = userProps.getProperty('AI_API_KEY') || props.getProperty('AI_API_KEY') || '';
  if (!apiKey) {
    return { error: 'IA no configurada. Define tu API key en Configuración.' };
  }
  var payload = buildAiPayload_(prompt, flags, currentPlan);
  var response = callAiEndpoint_(endpoint, apiKey, payload);
  if (response.error) return response;
  var parsed = parseAiPlanResponse_(response);
  if (parsed.error) return parsed;
  var validationError = validatePlanStructureOnly_(parsed.plan, flags);
  if (validationError) return { error: validationError };
  if (!parsed.summary) return { error: 'Plan IA inválido: falta summary.' };
  var summary = parsed.summary;
  return {
    plan: parsed.plan,
    summary: summary,
    aiDebug: {
      timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      payload: JSON.stringify(payload, null, 2),
      rawResponse: response.text
    }
  };
}
function commandToPlan_(prompt, flags) {
  var meta = {
    version: '1.0',
    dryRun: !!(flags && flags.dryRun),
    safeMode: !!(flags && flags.safeMode),
    notes: ''
  };
  var action = parseCommand_(prompt);
  if (action.error) return action;
  return { meta: meta, actions: [action] };
}
function editPlanWithCommand_(prompt, plan) {
  var parts = prompt.split(' ');
  var cmd = parts[0];
  if (cmd === '/replaceAction') {
    var index = extractIndex_(prompt);
    var actionJson = extractJson_(prompt);
    if (index < 0 || !actionJson) return { error: 'Uso: /replaceAction index=2 {..}' };
    plan.actions[index] = actionJson;
    return plan;
  }
  if (cmd === '/removeAction') {
    var idx = extractIndex_(prompt);
    if (idx < 0) return { error: 'Uso: /removeAction index=1' };
    plan.actions.splice(idx, 1);
    return plan;
  }
  if (cmd === '/addAction') {
    var addJson = extractJson_(prompt);
    if (!addJson) return { error: 'Uso: /addAction {..}' };
    plan.actions.push(addJson);
    return plan;
  }
  return { error: 'Comando de edición no soportado.' };
}
function buildAiPayload_(prompt, flags, currentPlan) {
  var system = [
    'Eres un generador de planes JSON para Google Sheets.',
    'Tu salida se usa para ejecutar acciones reales; evita ambigüedad.',
    'Responde SOLO JSON válido y sin texto extra.',
    'Formato de salida obligatorio:',
    '{"plan":{"meta":{"version":"1.0","dryRun":false,"safeMode":false,"notes":""},"actions":[{"op":"...","sheetName":"...","rangeA1":"..."}]},"summary":"..."}',
    'Reglas de generación (guía estandarizada):',
    '1) Usa únicamente operaciones incluidas en allowedOpsWhitelist.',
    '2) En cada acción usa campos canónicos: op, sheetName, rangeA1 y parámetros del op.',
    '3) Para color único usa color en formato #RRGGBB.',
    '4) Para setBackgrounds usa colors como matriz 2D exacta del rango (mismas filas/columnas).',
    '5) En colors cada celda debe ser string #RRGGBB (sin objetos, sin nombres, sin rgb()).',
    '6) No uses campos alternativos: range, backgrounds, bgColor, backgroundColor, operation, action, type, method.',
    '7) No inventes operaciones fuera de whitelist aunque parezcan útiles.',
    '8) Si falta un dato crítico, completa con el valor más probable y explícitalo en plan.meta.notes.',
    '9) Para gradientes calcula explícitamente todos los colores celda por celda y respeta orientación del rango.',
    '10) summary debe describir objetivo, acciones y ubicación en lenguaje profesional breve.',
    '11) Mantén plan.meta.version="1.0", dryRun=false y safeMode=false.',
    '12) No devuelvas markdown ni bloques de código.'
  ].join('\n');
  var userPayload = {
    userMessage: String(prompt || ''),
    allowedOpsWhitelist: (typeof ALLOWED_OPS !== 'undefined' && Array.isArray(ALLOWED_OPS)) ? ALLOWED_OPS : []
  };
  var user = JSON.stringify(userPayload);
  return {
    model: 'gpt-4o-mini',
    temperature: 0.1,
    messages: [
      { role: 'system', content: system },
      { role: 'user', content: user }
    ]
  };
}

function callAiEndpoint_(endpoint, apiKey, payload) {
  var url = endpoint || 'https://api.openai.com/v1/chat/completions';
  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    var text = response.getContentText();
    if (code < 200 || code >= 300) {
      return { error: 'Error IA: ' + code + ' ' + text };
    }
    return { ok: true, text: text };
  } catch (e) {
    return { error: 'Error IA: ' + e.message };
  }
}
function parseAiPlanResponse_(response) {
  var raw = response.text;
  var data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    return { error: 'La respuesta del servicio IA no llegó en JSON válido. Intenta nuevamente.' };
  }
  var content = '';
  if (data && data.choices && data.choices.length && data.choices[0].message) {
    content = data.choices[0].message.content;
  } else if (data && data.plan) {
    return { plan: data.plan, summary: data.summary };
  }
  if (!content) return { error: 'Respuesta IA vacía.' };
  var parsed;
  try {
    parsed = JSON.parse(content);
  } catch (e) {
    return { error: 'La IA devolvió contenido no-JSON. Debe devolver solo JSON válido.' };
  }
  if (!parsed.plan || !parsed.plan.actions) return { error: 'Plan IA inválido.' };
  return { plan: parsed.plan, summary: parsed.summary };
}
function parseCommand_(prompt) {
  var cmd = prompt.split(' ')[0];
  if (cmd === '/color') {
    var match = prompt.match(/\/color\s+"([^"]+)"\s+([^\s]+)\s+(#[0-9a-fA-F]{6})/);
    if (!match) return { error: 'Uso: /color "HOJA" A1:B2 #RRGGBB' };
    return { op: 'setBackground', sheetName: match[1], rangeA1: match[2], color: match[3] };
  }
  if (cmd === '/formula') {
    var fmatch = prompt.match(/\/formula\s+"([^"]+)"\s+([^\s]+)\s+(.+)/);
    if (!fmatch) return { error: 'Uso: /formula "HOJA" A1 =FORMULA' };
    return { op: 'setFormula', sheetName: fmatch[1], rangeA1: fmatch[2], formula: fmatch[3] };
  }
  if (cmd === '/move') {
    var mmatch = prompt.match(/\/move\s+"([^"]+)"\s+([^\s]+)\s+([^\s]+)/);
    if (!mmatch) return { error: 'Uso: /move "HOJA" A1:B2 C1:D2' };
    return { op: 'moveRange', sheetName: mmatch[1], sourceA1: mmatch[2], targetA1: mmatch[3] };
  }
  if (cmd === '/renamefile') {
    var nmatch = prompt.match(/\/renamefile\s+(.+)/);
    if (!nmatch) return { error: 'Uso: /renamefile NuevoNombre' };
    return { op: 'renameFile', name: nmatch[1].trim() };
  }
  if (cmd === '/exportpdf') {
    var ematch = prompt.match(/\/exportpdf\s+"([^"]+)"\s+filename=(.+)/);
    if (!ematch) return { error: 'Uso: /exportpdf "HOJA" filename=archivo.pdf' };
    return { op: 'exportPdf', sheetName: ematch[1], filename: ematch[2].trim() };
  }
  if (cmd === '/insertimage') {
    var imatch = prompt.match(/\/insertimage\s+sheet="([^"]+)"\s+cell=([^\s]+)\s+fileId=([^\s]+)\s+width=(\d+)\s+height=(\d+)/);
    if (!imatch) return { error: 'Uso: /insertimage sheet="HOJA" cell=A1 fileId=ID width=200 height=80' };
    return { op: 'insertImageFromDrive', sheetName: imatch[1], cell: imatch[2], fileId: imatch[3], width: Number(imatch[4]), height: Number(imatch[5]) };
  }
  if (cmd === '/createSheet') {
    var cmatch = prompt.match(/\/createSheet\s+"([^"]+)"/);
    if (!cmatch) return { error: 'Uso: /createSheet "Nombre"' };
    return { op: 'createSheet', sheetName: cmatch[1] };
  }
  if (cmd === '/deleteSheet') {
    var dmatch = prompt.match(/\/deleteSheet\s+"([^"]+)"/);
    if (!dmatch) return { error: 'Uso: /deleteSheet "Nombre"' };
    return { op: 'deleteSheet', sheetName: dmatch[1] };
  }
  return { error: 'Comando no reconocido.' };
}
function extractIndex_(prompt) {
  var match = prompt.match(/index=(\d+)/);
  if (!match) return -1;
  return parseInt(match[1], 10);
}
function extractJson_(prompt) {
  var match = prompt.match(/\{[\s\S]+\}/);
  if (!match) return null;
  try {
    return JSON.parse(match[0]);
  } catch (e) {
    return null;
  }
}
function buildSummary_(plan) {
  var lines = [];
  lines.push('Objetivo: aplicar los cambios solicitados en la hoja respetando prioridad de instrucciones del usuario.');
  lines.push('Acciones:');
  lines.push('- Bloque de ejecución: ' + plan.actions.length + ' acción(es) coordinadas para aplicar formato/datos según el plan.');
  lines.push('Ubicación exacta:');
  plan.actions.forEach(function (action) {
    if (action.sheetName) {
      var range = action.rangeA1 || action.sourceA1 || action.targetA1 || action.cell || 'desconocido';
      lines.push('- ' + action.sheetName + ' ' + range);
    }
  });
  lines.push('Impacto:');
  lines.push(plan.actions.map(function (a) { return describeImpact_(a); }).join(' | '));
  return lines.join('\n');
}
function describeAction_(action) {
  var op = friendlyOp_(action.op || 'acción');
  var target = action.rangeA1 || action.targetA1 || action.sourceA1 || action.cell || 'sin rango';
  var color = action.color ? ' con color ' + formatColorLabel_(action.color) : '';
  if (action.op === 'setBackgrounds' && Array.isArray(action.colors)) {
    color = ' con colores distintos por celda';
  }
  return op + ' en ' + (action.sheetName || 'hoja activa') + ' (' + target + ')' + color + '.';
}
function describeImpact_(action) {
  if (action.op === 'setBackground' || action.op === 'setBackgrounds') {
    return 'Se verá un cambio visual de color en ' + (action.rangeA1 || action.targetA1 || action.cell || 'el rango elegido');
  }
  if (action.op === 'setFormula' || action.op === 'setFormulas') {
    return 'Las celdas calcularán resultados automáticos en ' + (action.rangeA1 || 'el rango elegido');
  }
  if (action.op === 'copyRange' || action.op === 'moveRange') {
    return 'Se trasladará contenido desde ' + (action.sourceA1 || 'origen') + ' hacia ' + (action.targetA1 || 'destino');
  }
  return 'Se aplicará un cambio en ' + (action.sheetName || 'la hoja activa');
}
function friendlyOp_(op) {
  var map = {
    setBackground: 'Pintar celdas',
    setBackgrounds: 'Pintar celdas',
    setValue: 'Escribir valor',
    setValues: 'Escribir valores',
    setFormula: 'Aplicar fórmula',
    setFormulas: 'Aplicar fórmulas',
    copyRange: 'Copiar rango',
    moveRange: 'Mover rango',
    setBorder: 'Aplicar bordes',
    setBorders: 'Aplicar bordes'
  };
  return map[op] || ('Aplicar ' + op);
}
function formatColorLabel_(hex) {
  var normalized = String(hex || '').toLowerCase();
  var names = {
    '#ff0000': 'Rojo',
    '#00ff00': 'Verde',
    '#0000ff': 'Azul',
    '#ffff00': 'Amarillo',
    '#ffa500': 'Naranja',
    '#800080': 'Morado',
    '#000000': 'Negro',
    '#ffffff': 'Blanco',
    '#ffeb3b': 'Amarillo claro'
  };
  var name = names[normalized] || 'Color personalizado';
  return name + ' (' + hex + ')';
}
