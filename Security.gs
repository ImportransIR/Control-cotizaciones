/** Trigger principal onOpen: menu y verificacion ligera. */
function onOpen(e) {
  try {
    buildMenu_();

    const role = getUserRole_(getCurrentUserEmail_());
    if (role === CC_CONFIG.ROLES.ADMIN) {
      ensureUsersSheet_();
      ensureConfigSheet_();
      ensureLogsSheet_();
    }
  } catch (error) {
    console.error('onOpen error:', error);
  }
}

/** Trigger principal onEdit: valida permisos por rol y sincroniza. */
function onEdit(e) {
  handleEditByRole_(e);
}

/** Trigger onChange (instalable): trazabilidad de cambios estructurales. */
function onChange(e) {
  try {
    if (!e || !e.changeType) return;
    if (e.changeType === 'REMOVE_GRID') {
      appendLog_({
        usuario: getCurrentUserEmail_(),
        accion: 'onChange_REMOVE_GRID',
        resultado: 'warn',
        detalle: 'Se detecto eliminacion de hoja. Revisar estructura y re-aplicar protecciones.'
      });
    }
  } catch (error) {
    console.error('onChange error:', error);
  }
}

/** Aplica protecciones de encabezados y columnas sensibles en hojas operativas. */
function applySheetProtections_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheets = ss.getSheets().filter(function (sh) {
      const name = sh.getName();
      return name === CC_CONFIG.MAIN_SHEET_NAME || isUserSheet_(name);
    });

    sheets.forEach(function (sh) {
      clearManagedProtections_(sh);
      protectHeaders_(sh);
      protectSensitiveColumns_(sh);
    });

    appendLog_({
      usuario: getCurrentUserEmail_(),
      accion: 'applySheetProtections',
      resultado: 'ok',
      detalle: 'Protecciones aplicadas en ' + sheets.length + ' hoja(s)'
    });

    toast_('Protecciones aplicadas correctamente', 'Control Comercial');
    return ok({ sheets: sheets.map(function (s) { return s.getName(); }) });
  } catch (error) {
    console.error('applySheetProtections_ error:', error);
    return err('No fue posible aplicar protecciones', String(error && error.message || error));
  }
}

/** Reaplica protecciones (wrapper administrativo). */
function reapplyProtections_() {
  return applySheetProtections_();
}

/** Protege fila de encabezado para evitar alteraciones/borrados. */
function protectHeaders_(sheet) {
  if (!sheet) return;
  const range = sheet.getRange(CC_CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn());
  const p = range.protect().setDescription(CC_CONFIG.PROTECTION_TAG + ':headers:' + sheet.getName());
  applyAdminsOnly_(p);
}

/** Protege columnas sensibles por encabezado en cada hoja operativa. */
function protectSensitiveColumns_(sheet) {
  if (!sheet) return;
  const headerMap = getHeaderMap_(sheet);
  const maxRows = Math.max(1, sheet.getMaxRows() - 1);

  CC_CONFIG.SENSITIVE_HEADERS.forEach(function (h) {
    const idx = headerMap[normalizeHeader_(h)];
    if (!idx) return;

    const range = sheet.getRange(2, idx, maxRows, 1);
    const p = range.protect().setDescription(CC_CONFIG.PROTECTION_TAG + ':sensitive:' + sheet.getName() + ':' + idx);
    applyAdminsOnly_(p);
  });
}

/** Maneja permisos por rol y revierte ediciones no autorizadas. */
function handleEditByRole_(e) {
  try {
    if (!e || !e.range) return;

    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    const row = range.getRow();
    const col = range.getColumn();

    if (row <= CC_CONFIG.HEADER_ROW) {
      revertUnauthorizedEdit_(e, 'No se permite editar encabezados.');
      return;
    }

    if (sheetName === CC_CONFIG.USERS_SHEET_NAME || sheetName === CC_CONFIG.SYSTEM_SHEET_NAME || sheetName === CC_CONFIG.LOGS_SHEET_NAME) {
      const emailCfg = getCurrentUserEmail_();
      if (!isAdmin_(emailCfg)) {
        revertUnauthorizedEdit_(e, 'Solo admin puede editar hojas de configuracion.');
        return;
      }
    }

    if (!(sheetName === CC_CONFIG.MAIN_SHEET_NAME || isUserSheet_(sheetName))) {
      return;
    }

    const email = getCurrentUserEmail_();
    const role = getUserRole_(email);
    const headerMap = getHeaderMap_(sheet);

    if (!canEditCell_(role, col, headerMap)) {
      revertUnauthorizedEdit_(e, 'No autorizado para editar esta columna.');
      return;
    }

    enforceBusinessRules_(e, role, headerMap);
    syncEditedRowToMain_(e, headerMap);
  } catch (error) {
    console.error('handleEditByRole_ error:', error);
    appendLog_({
      usuario: getCurrentUserEmail_(),
      accion: 'handleEditByRole_error',
      resultado: 'error',
      detalle: String(error && error.message || error)
    });
  }
}

/** Revierte cambios no autorizados y registra evento. */
function revertUnauthorizedEdit_(e, reason) {
  const range = e.range;
  const sheet = range.getSheet();

  // Si fue una celda simple, restaurar oldValue; en rangos, limpiar edicion.
  if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
    if (typeof e.oldValue !== 'undefined') {
      setValueSafeWithValidation_(range, e.oldValue);
    } else {
      range.clearContent();
    }
  } else {
    range.clearContent();
  }

  appendLog_({
    usuario: getCurrentUserEmail_(),
    accion: 'unauthorizedEdit',
    hoja: sheet.getName(),
    fila: range.getRow(),
    columna: range.getColumn(),
    valorAnterior: typeof e.oldValue === 'undefined' ? '' : e.oldValue,
    valorNuevo: typeof e.value === 'undefined' ? '' : e.value,
    resultado: 'blocked',
    detalle: reason || 'Edicion no permitida'
  });

  toast_(reason || 'Edicion no autorizada', 'Control Comercial');
}

/** Crea trigger instalable onOpen/onChange si se requiere ejecucion con permisos extendidos. */
function installableOnOpen_() {
  const handlers = ['onOpen', 'onChange'];

  handlers.forEach(function (handler) {
    const exists = ScriptApp.getProjectTriggers().some(function (t) {
      return t.getHandlerFunction() === handler;
    });

    if (!exists) {
      if (handler === 'onOpen') {
        ScriptApp.newTrigger(handler).forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
      }
      if (handler === 'onChange') {
        ScriptApp.newTrigger(handler).forSpreadsheet(SpreadsheetApp.getActive()).onChange().create();
      }
    }
  });

  return ok({ installed: handlers });
}

/** Elimina protecciones administradas por este sistema en la hoja. */
function clearManagedProtections_(sheet) {
  const all = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  all.forEach(function (p) {
    if ((p.getDescription() || '').indexOf(CC_CONFIG.PROTECTION_TAG + ':') === 0) {
      p.remove();
    }
  });
}

/** Restringe editorado de proteccion solo a admins declarados. */
function applyAdminsOnly_(protection) {
  const usersMap = getUsersMap_();
  const admins = Object.keys(usersMap).filter(function (email) {
    const u = usersMap[email];
    return u && u.rol === CC_CONFIG.ROLES.ADMIN && u.activo;
  });

  protection.setWarningOnly(false);

  const currentEditors = protection.getEditors();
  if (currentEditors && currentEditors.length) {
    protection.removeEditors(currentEditors);
  }

  if (admins.length) {
    protection.addEditors(admins);
  }

  try {
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  } catch (error) {
    // Ignorar cuando dominio no aplica.
  }
}
