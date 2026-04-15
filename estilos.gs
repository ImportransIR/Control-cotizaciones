/** ========= ESTILOS & VALIDACIONES ========= */

function initValidationsAndFormats() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Datos');
  if (!sh) throw new Error('No existe la hoja "Datos".');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const rngEstado  = sh.getRange(2, 8,  lastRow - 1, 1); // H
  const rngProceso = sh.getRange(2, 14, lastRow - 1, 1); // N

  const dvEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(['En curso','Descartada','Aprobada','Rechazada'], true)
    .setAllowInvalid(false).build();

  const dvProceso = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Sin facturar','Facturado'], true)
    .setAllowInvalid(false).build();

  rngEstado.setDataValidation(dvEstado);
  rngProceso.setDataValidation(dvProceso);

  const rules = sh.getConditionalFormatRules().filter(r =>
    !r.getRanges().some(g =>
      (g.getColumn() === 8  && g.getLastColumn() === 8 ) ||
      (g.getColumn() === 14 && g.getLastColumn() === 14)
    )
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LOWER($H2)="en curso"')
      .setBackground('#ffeb3b').setRanges([sh.getRange(2,8, sh.getMaxRows()-1,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LOWER($H2)="descartada"')
      .setBackground('#2196f3').setRanges([sh.getRange(2,8, sh.getMaxRows()-1,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LOWER($H2)="aprobada"')
      .setBackground('#4caf50').setRanges([sh.getRange(2,8, sh.getMaxRows()-1,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LOWER($H2)="rechazada"')
      .setBackground('#f44336').setRanges([sh.getRange(2,8, sh.getMaxRows()-1,1)]).build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LOWER($N2)="sin facturar"')
      .setBackground('#ffeb3b').setRanges([sh.getRange(2,14, sh.getMaxRows()-1,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LOWER($N2)="facturado"')
      .setBackground('#4caf50').setRanges([sh.getRange(2,14, sh.getMaxRows()-1,1)]).build()
  );

  sh.setConditionalFormatRules(rules);
}

function ensureRowValidation_(sh, row) {
  if (row < 2) return;
  const dvEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(['En curso','Descartada','Aprobada','Rechazada'], true)
    .setAllowInvalid(false).build();
  const dvProceso = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Sin facturar','Facturado'], true)
    .setAllowInvalid(false).build();

  sh.getRange(row, 8).setDataValidation(dvEstado);
  sh.getRange(row, 14).setDataValidation(dvProceso);
}

function applyValidationAndFormatsToSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    sh.getRange(2, 8,  lastRow-1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['En curso','Descartada','Aprobada','Rechazada'], true)
        .setAllowInvalid(false).build()
    );
    sh.getRange(2, 14, lastRow-1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Sin facturar','Facturado'], true)
        .setAllowInvalid(false).build()
    );
  }

  const rules = sh.getConditionalFormatRules().filter(r =>
    !r.getRanges().some(g =>
      (g.getColumn() === 8  && g.getLastColumn() === 8 ) ||
      (g.getColumn() === 14 && g.getLastColumn() === 14)
    )
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LOWER($H2)="en curso"')
      .setBackground('#ffeb3b').setRanges([sh.getRange(2,8, sh.getMaxRows()-1,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LOWER($H2)="descartada"')
      .setBackground('#2196f3').setRanges([sh.getRange(2,8, sh.getMaxRows()-1,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LOWER($H2)="aprobada"')
      .setBackground('#4caf50').setRanges([sh.getRange(2,8, sh.getMaxRows()-1,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LOWER($H2)="rechazada"')
      .setBackground('#f44336').setRanges([sh.getRange(2,8, sh.getMaxRows()-1,1)]).build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LOWER($N2)="sin facturar"')
      .setBackground('#ffeb3b').setRanges([sh.getRange(2,14, sh.getMaxRows()-1,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LOWER($N2)="facturado"')
      .setBackground('#4caf50').setRanges([sh.getRange(2,14, sh.getMaxRows()-1,1)]).build()
  );

  sh.setConditionalFormatRules(rules);
}

function refreshFormatsAllUserSheets() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets()
    .filter(s => s.getName() !== 'Datos' && /@/.test(s.getName()))
    .forEach(applyValidationAndFormatsToSheet_);
}

/* ===== Validación de factura multi-línea (1–8 alfanuméricos) ===== 
const INVOICE_MAX_LEN_STYLES = 8; // nombre cambiado para evitar conflicto
function validateInvoices_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return { ok: false, normalized: '', invalid: ['(vacío)'] };

  const lines = raw.split(/\r?\n/);
  const normalized = [];
  const invalid = [];

  lines.forEach((ln, i) => {
    const code = (ln || '').replace(/[^A-Za-z0-9]/g, '').toUpperCase();
    if (code.length >= 1 && code.length <= INVOICE_MAX_LEN_STYLES) {
      normalized.push(code);
    } else {
      invalid.push(`${i + 1}: "${ln}"`);
    }
  });

  return {
    ok: invalid.length === 0,
    normalized: normalized.join('\n'),
    invalid
  };
}*/

function validateInvoices_(value) {
  return {
    ok: true,
    normalized: String(value == null ? '' : value).trim(),
    invalid: []
  };
}

/* ====== Utilidades ====== */
function setValueSafeWithValidation_(range, value) {
  const rule = range.getDataValidation();
  try {
    range.setValue(value);
  } catch (err) {
    range.clearDataValidations();
    range.setValue(value);
    if (rule) range.setDataValidation(rule);
  }
}

function ensureWrap_(range) {
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

function _daysDiff_(fromDate, toDate) {
  if (!fromDate || !toDate) return NaN;
  const tz = 'America/Bogota';
  const d1 = new Date(Utilities.formatDate(new Date(fromDate), tz, 'yyyy-MM-dd'));
  const d2 = new Date(Utilities.formatDate(new Date(toDate),   tz, 'yyyy-MM-dd'));
  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  return Math.floor((d2 - d1) / MS_PER_DAY);
}

function _findRowById_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
  const idx = ids.findIndex(v => v === id);
  return (idx >= 0) ? (idx + 2) : -1;
}

function _setDescartadaSafe_(sheet, row) {
  setValueSafeWithValidation_(sheet.getRange(row, 8), 'Descartada'); // H
  ensureRowValidation_(sheet, row);
}

/* ===== Triggers — Descarte automático 90 días ===== */
function setupAutoDiscardTrigger() {
  const handler = 'autoDescartar90Dias';
  deleteTriggersByHandler_(handler);
  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert(
    'Descarte automático activado',
    'Se programó autoDescartar90Dias para ejecutarse todos los días a las 08:00.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function deleteTriggersByHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Control Comercial')
    .addItem('▶ Ejecutar descarte 90 días ahora', 'autoDescartar90Dias')
    .addItem('⏰ Activar disparador diario (08:00)', 'setupAutoDiscardTrigger')
    .addItem('🗑 Desactivar disparador diario', 'disableAutoDiscardTrigger')
    .addToUi();

  // Aplica estilos si la función existe
  if (typeof aplicarEstilosATodas === 'function') {
    aplicarEstilosATodas();
  }
}

function disableAutoDiscardTrigger() {
  deleteTriggersByHandler_('autoDescartar90Dias');
  SpreadsheetApp.getUi().alert(
    'Descarte automático desactivado',
    'Se eliminaron los disparadores de autoDescartar90Dias.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
