/** Respuesta estandar exitosa. */
function ok(data) {
  return { ok: true, data: data == null ? null : data };
}

/** Respuesta estandar de error. */
function err(message, extra) {
  const out = { ok: false, err: String(message || 'Error no especificado') };
  if (extra != null) out.extra = extra;
  return out;
}

/** Normaliza encabezados para mapeo robusto. */
function normalizeHeader_(text) {
  return String(text == null ? '' : text)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/** Busca hoja por nombre sin lanzar excepcion. */
function findSheetByNameSafe_(name) {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name) || null;
  } catch (error) {
    console.error('findSheetByNameSafe_ error:', error);
    return null;
  }
}

/** Construye mapa encabezado-normalizado -> indice de columna (1-based). */
function getHeaderMap_(sheet) {
  if (!sheet) return {};
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};

  const headers = sheet
    .getRange(CC_CONFIG.HEADER_ROW, 1, 1, lastCol)
    .getValues()[0];

  const map = {};
  headers.forEach(function (h, idx) {
    const key = normalizeHeader_(h);
    if (key) map[key] = idx + 1;
  });
  return map;
}

/** Nombre de hoja seguro basado en correo. */
function safeSheetNameFromEmail_(email) {
  if (!email) return 'usuario_sin_correo';
  return String(email).replace(/[^A-Za-z0-9@._-]/g, '_').slice(0, 99);
}

/** UI segura para ejecuciones con contexto de interfaz. */
function safeUi_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (error) {
    return null;
  }
}

/** Notificacion no intrusiva (funciona en onEdit simple). */
function toast_(message, title) {
  try {
    SpreadsheetApp.getActive().toast(String(message || ''), String(title || 'Control Comercial'));
  } catch (error) {
    console.error('toast_ error:', error);
  }
}

/** Diferencia de dias normalizada por fecha. */
function daysDiff_(fromDate, toDate) {
  if (!fromDate || !toDate) return NaN;
  const d1 = new Date(Utilities.formatDate(new Date(fromDate), CC_CONFIG.TIMEZONE, 'yyyy-MM-dd'));
  const d2 = new Date(Utilities.formatDate(new Date(toDate), CC_CONFIG.TIMEZONE, 'yyyy-MM-dd'));
  return Math.floor((d2 - d1) / (24 * 60 * 60 * 1000));
}

/** Escritura tolerante a validaciones de celda. */
function setValueSafeWithValidation_(range, value) {
  const rule = range.getDataValidation();
  try {
    range.setValue(value);
  } catch (error) {
    range.clearDataValidations();
    range.setValue(value);
    if (rule) range.setDataValidation(rule);
  }
}

/** Mantiene saltos de linea visibles en celda. */
function ensureWrap_(range) {
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

/** Encuentra fila por ID (columna A). */
function findRowById_(sheet, id) {
  const last = sheet.getLastRow();
  if (!sheet || last < 2) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
  const idx = ids.findIndex(function (v) { return v === id; });
  return idx >= 0 ? idx + 2 : -1;
}

/** Detecta si una hoja es de usuario (correo). */
function isUserSheet_(sheetName) {
  return !!sheetName && sheetName !== CC_CONFIG.MAIN_SHEET_NAME && /@/.test(sheetName);
}

/** Registra trazabilidad en hoja Logs. */
function appendLog_(rowData) {
  try {
    const sh = ensureLogsSheet_();
    sh.appendRow([
      new Date(),
      rowData.usuario || '',
      rowData.accion || '',
      rowData.hoja || '',
      rowData.fila || '',
      rowData.columna || '',
      rowData.valorAnterior || '',
      rowData.valorNuevo || '',
      rowData.resultado || '',
      rowData.detalle || ''
    ]);
  } catch (error) {
    console.error('appendLog_ error:', error);
  }
}

/** Compatibilidad con nombres historicos en el proyecto. */
function _daysDiff_(fromDate, toDate) {
  return daysDiff_(fromDate, toDate);
}

/** Compatibilidad con nombres historicos en el proyecto. */
function _findRowById_(sheet, id) {
  return findRowById_(sheet, id);
}

/** Compatibilidad con nombres historicos en el proyecto. */
function _setDescartadaSafe_(sheet, row) {
  const headerMap = getHeaderMap_(sheet);
  const estadoCol = headerMap[normalizeHeader_('Estado')] || 8;
  setValueSafeWithValidation_(sheet.getRange(row, estadoCol), 'Descartada');
}

/** Alerta centrada (cuando UI esta disponible). */
function _uiAlert_(title, message) {
  const ui = safeUi_();
  if (ui) {
    ui.alert(String(title || 'Control Comercial'), String(message || ''), ui.ButtonSet.OK);
  } else {
    toast_(message, title);
  }
}
