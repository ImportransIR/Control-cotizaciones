/** Reglas de negocio sobre edicion autorizada. */
function enforceBusinessRules_(e, role, headerMap) {
  const range = e.range;
  const sh = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();

  const estadoCol = headerMap[normalizeHeader_('Estado')];
  const estadoProcesoCol = headerMap[normalizeHeader_('Estado del proceso')];
  const numeroFacturaCol = headerMap[normalizeHeader_('Numero factura')];

  // Normaliza formato de fechas en columnas conocidas.
  const dateHeaders = ['Fecha de solicitud', 'Fecha de inicio', 'Fecha final', 'Fecha de factura'];
  dateHeaders.forEach(function (h) {
    const idx = headerMap[normalizeHeader_(h)];
    if (idx && idx === col) {
      range.setNumberFormat('dd/MM/yyyy');
    }
  });

  // Si cambian Numero factura y no esta vacio, forzar flujo facturado/aprobada.
  if (numeroFacturaCol && col === numeroFacturaCol) {
    const factura = String(range.getValue() || '').trim();
    if (factura) {
      if (estadoProcesoCol) setValueSafeWithValidation_(sh.getRange(row, estadoProcesoCol), 'Facturado');
      if (estadoCol) setValueSafeWithValidation_(sh.getRange(row, estadoCol), 'Aprobada');
    }
  }

  // Si marcan Facturado, exigir Numero factura.
  if (estadoProcesoCol && col === estadoProcesoCol) {
    const estadoProceso = normalizeHeader_(range.getValue());
    if (estadoProceso === 'facturado' && numeroFacturaCol) {
      const facturaNow = String(sh.getRange(row, numeroFacturaCol).getValue() || '').trim();
      if (!facturaNow) {
        setValueSafeWithValidation_(range, e.oldValue || 'Sin facturar');
        appendLog_({
          usuario: getCurrentUserEmail_(),
          accion: 'validateFacturado',
          hoja: sh.getName(),
          fila: row,
          columna: col,
          valorAnterior: typeof e.oldValue === 'undefined' ? '' : e.oldValue,
          valorNuevo: typeof e.value === 'undefined' ? '' : e.value,
          resultado: 'blocked',
          detalle: 'Se requiere Numero factura para estado Facturado'
        });
        toast_('Debes registrar Numero factura antes de marcar Facturado', 'Control Comercial');
        return;
      }
      if (estadoCol) setValueSafeWithValidation_(sh.getRange(row, estadoCol), 'Aprobada');
    }
  }

  // Regla opcional de descarte automatico en 90 dias (si aplica por fecha/estado).
  applyDiscardRuleForRow_(sh, row, headerMap);
}

/** Descarta automaticamente cotizacion despues de 90 dias si no fue aprobada/facturada. */
function applyDiscardRuleForRow_(sheet, row, headerMap) {
  const fechaCol = headerMap[normalizeHeader_('Fecha de solicitud')];
  const estadoCol = headerMap[normalizeHeader_('Estado')];
  const estadoProcesoCol = headerMap[normalizeHeader_('Estado del proceso')];

  if (!fechaCol || !estadoCol || !estadoProcesoCol) return;

  const fecha = sheet.getRange(row, fechaCol).getValue();
  if (!fecha) return;

  const estado = normalizeHeader_(sheet.getRange(row, estadoCol).getValue());
  const proceso = normalizeHeader_(sheet.getRange(row, estadoProcesoCol).getValue());

  const dias = daysDiff_(fecha, new Date());
  if (isNaN(dias) || dias < 90) return;
  if (estado === 'aprobada' || proceso === 'facturado') return;

  setValueSafeWithValidation_(sheet.getRange(row, estadoCol), 'Descartada');
}

/** Sincroniza fila editada hacia hoja principal manteniendo fuente central. */
function syncEditedRowToMain_(e, headerMap) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();

  if (row <= 1) return;

  const main = findSheetByNameSafe_(CC_CONFIG.MAIN_SHEET_NAME);
  if (!main) return;

  const cols = main.getLastColumn();
  const currentRow = sheet.getRange(row, 1, 1, Math.min(cols, sheet.getLastColumn())).getValues()[0];
  const id = currentRow[0];

  if (!id) return;

  // Si la edicion viene de hoja principal, opcionalmente reflejar en hoja de usuario.
  if (sheetName === CC_CONFIG.MAIN_SHEET_NAME) {
    syncMainRowToUserSheet_(main, row);
    return;
  }

  // Si viene de hoja de usuario, actualizar/inserta en principal por ID.
  const targetRow = findRowById_(main, id);
  if (targetRow > 0) {
    main.getRange(targetRow, 1, 1, currentRow.length).setValues([currentRow]);
  } else {
    main.appendRow(currentRow);
  }
}

/** Refleja fila de hoja principal a hoja de usuario segun correo (si existe). */
function syncMainRowToUserSheet_(mainSheet, row) {
  const headerMap = getHeaderMap_(mainSheet);
  const correoCol = headerMap[normalizeHeader_('Correo electronico')];
  if (!correoCol) return;

  const email = String(mainSheet.getRange(row, correoCol).getValue() || '').trim().toLowerCase();
  if (!email) return;

  const userSheetName = safeSheetNameFromEmail_(email);
  const userSheet = findSheetByNameSafe_(userSheetName);
  if (!userSheet) return;

  const data = mainSheet.getRange(row, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  const id = data[0];
  if (!id) return;

  const userRow = findRowById_(userSheet, id);
  if (userRow > 0) {
    userSheet.getRange(userRow, 1, 1, data.length).setValues([data]);
  } else {
    userSheet.appendRow(data);
  }
}
