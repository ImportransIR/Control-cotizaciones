/**
 * Sincronizacion incremental entre hojas de usuario y hoja principal.
 * Sin onEdit global para evitar colision con Security.gs.
 */

/** Normaliza fechas comunes antes de escribir en la hoja principal. */
function _normalizeRowForDatos_CC_(rowData, headerMap) {
  const out = rowData.slice();
  const dateHeaders = ['Fecha de solicitud', 'Fecha de inicio', 'Fecha final', 'Fecha de factura'];

  dateHeaders.forEach(function (header) {
    const col = headerMap[normalizeHeader_(header)];
    if (!col) return;

    const idx = col - 1;
    const value = out[idx];
    if (value === '' || value == null) {
      out[idx] = '';
      return;
    }

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return;

    const parsed = new Date(value);
    out[idx] = parsed && !isNaN(parsed) ? parsed : '';
  });

  return out;
}

/** Escribe una fila respetando validaciones de celdas. */
function setRowValuesSafe_CC_(sheet, row, values) {
  for (let c = 1; c <= values.length; c++) {
    setValueSafeWithValidation_(sheet.getRange(row, c), values[c - 1]);
  }
}

/** Sincroniza todas las hojas de usuario (@correo) con hoja principal. */
function syncAllUserSheetsToDatos() {
  const ss = SpreadsheetApp.getActive();
  const hojaDatos = findSheetByNameSafe_(CC_CONFIG.MAIN_SHEET_NAME);
  if (!hojaDatos) {
    _uiAlert_('Sincronizacion', 'No existe la hoja principal.');
    return;
  }

  const colCount = hojaDatos.getLastColumn();
  const headerMapDatos = getHeaderMap_(hojaDatos);

  const userSheets = ss.getSheets().filter(function (s) {
    return isUserSheet_(s.getName());
  });

  const lastDatos = hojaDatos.getLastRow();
  const idsDatos = lastDatos > 1
    ? hojaDatos.getRange(2, 1, lastDatos - 1, 1).getValues().flat()
    : [];

  let totalUpdated = 0;
  let totalInserted = 0;

  userSheets.forEach(function (sh) {
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const data = sh.getRange(2, 1, lastRow - 1, colCount).getValues();

    data.forEach(function (rowData) {
      const id = rowData[0];
      if (!id) return;

      const normalizedRow = _normalizeRowForDatos_CC_(rowData, headerMapDatos);
      const idx = idsDatos.findIndex(function (v) { return v === id; });

      if (idx >= 0) {
        const targetRow = idx + 2;
        try {
          hojaDatos.getRange(targetRow, 1, 1, colCount).setValues([normalizedRow]);
        } catch (error) {
          setRowValuesSafe_CC_(hojaDatos, targetRow, normalizedRow);
        }
        ensureRowValidation_(hojaDatos, targetRow);
        totalUpdated++;
      } else {
        const newRow = hojaDatos.getLastRow() + 1;
        try {
          hojaDatos.getRange(newRow, 1, 1, colCount).setValues([normalizedRow]);
        } catch (error) {
          setRowValuesSafe_CC_(hojaDatos, newRow, normalizedRow);
        }
        ensureRowValidation_(hojaDatos, newRow);
        idsDatos.push(id);
        totalInserted++;
      }
    });
  });

  appendLog_({
    usuario: getCurrentUserEmail_(),
    accion: 'syncAllUserSheetsToDatos',
    resultado: 'ok',
    detalle: 'Actualizados: ' + totalUpdated + ', nuevos: ' + totalInserted
  });

  _uiAlert_('Sincronizacion completada', 'Actualizados: ' + totalUpdated + '\nNuevos: ' + totalInserted);
}
