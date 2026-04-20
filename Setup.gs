/** Inicializa sistema de seguridad/permisos sin alterar estructura central. */
function initSystem() {
  try {
    const usersResult = ensureUsersSheet_();
    const configResult = ensureConfigSheet_();
    const logsResult = ensureLogsSheet_();
    const validateResult = validateMainSheetStructure_();
    const clientesResult = ensureClientesSupportSheet_();
    const catalogsResult = ensureCatalogs_();

    applySheetProtections_();

    return ok({
      usersSheet: usersResult.data,
      configSheet: configResult.data,
      logsSheet: logsResult.data,
      mainValidation: validateResult,
      clientesSheet: clientesResult.data,
      catalogsSheet: catalogsResult.data
    });
  } catch (error) {
    console.error('initSystem error:', error);
    return err('No fue posible inicializar el sistema', String(error && error.message || error));
  }
}

/** Crea hoja de usuarios/permisos si no existe. */
function ensureUsersSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = findSheetByNameSafe_(CC_CONFIG.USERS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CC_CONFIG.USERS_SHEET_NAME);
    sh.getRange(1, 1, 1, CC_CONFIG.USERS_HEADERS.length).setValues([CC_CONFIG.USERS_HEADERS]);
    sh.setFrozenRows(1);
  }

  const currentHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(normalizeHeader_);
  const expected = CC_CONFIG.USERS_HEADERS.map(normalizeHeader_);
  const needsReset = expected.some(function (h, i) { return currentHeaders[i] !== h; });
  if (needsReset) {
    sh.getRange(1, 1, 1, CC_CONFIG.USERS_HEADERS.length).setValues([CC_CONFIG.USERS_HEADERS]);
  }

  return ok({ name: sh.getName(), rows: sh.getLastRow() });
}

/** Crea hoja minima de configuracion del sistema. */
function ensureConfigSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = findSheetByNameSafe_(CC_CONFIG.SYSTEM_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CC_CONFIG.SYSTEM_SHEET_NAME);
    sh.getRange(1, 1, 1, 2).setValues([['clave', 'valor']]);
    sh.getRange(2, 1, 1, 2).setValues([['mainSheetName', CC_CONFIG.MAIN_SHEET_NAME]]);
    sh.setFrozenRows(1);
  }
  return ok({ name: sh.getName() });
}

/** Crea hoja de logs si no existe. */
function ensureLogsSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = findSheetByNameSafe_(CC_CONFIG.LOGS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CC_CONFIG.LOGS_SHEET_NAME);
    sh.getRange(1, 1, 1, CC_CONFIG.LOG_HEADERS.length).setValues([CC_CONFIG.LOG_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Valida hoja principal y encabezados clave sin cambiar orden actual. */
function validateMainSheetStructure_() {
  const sh = findSheetByNameSafe_(CC_CONFIG.MAIN_SHEET_NAME);
  if (!sh) return err('No existe la hoja principal: ' + CC_CONFIG.MAIN_SHEET_NAME);

  const headerMap = getHeaderMap_(sh);
  const missing = CC_CONFIG.EXPECTED_MAIN_HEADERS.filter(function (h) {
    return !headerMap[normalizeHeader_(h)];
  });

  if (missing.length) {
    return err('Faltan encabezados esperados en hoja principal', missing);
  }

  return ok({
    sheet: sh.getName(),
    totalHeaders: Object.keys(headerMap).length,
    validatedHeaders: CC_CONFIG.EXPECTED_MAIN_HEADERS.length
  });
}

/** Hoja de soporte futura para formulario de clientes. */
function ensureClientesSupportSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = findSheetByNameSafe_(CC_CONFIG.CLIENTES_SUPPORT_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CC_CONFIG.CLIENTES_SUPPORT_SHEET_NAME);
    sh.getRange(1, 1, 1, 6).setValues([[
      'rut_nit', 'razon_social', 'nombre_comercial', 'carpeta_link', 'activo', 'actualizado_en'
    ]]);
    sh.setFrozenRows(1);
  }
  return ok({ name: sh.getName() });
}

/** Catalogos minimos para futuras listas desplegables. */
function ensureCatalogs_() {
  const ss = SpreadsheetApp.getActive();
  let sh = findSheetByNameSafe_(CC_CONFIG.CATALOGS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CC_CONFIG.CATALOGS_SHEET_NAME);
    sh.getRange(1, 1, 1, 2).setValues([['catalogo', 'valor']]);
    sh.getRange(2, 1, 6, 2).setValues([
      ['estado', 'En curso'],
      ['estado', 'Descartada'],
      ['estado', 'Aprobada'],
      ['estado', 'Rechazada'],
      ['estadoProceso', 'Sin facturar'],
      ['estadoProceso', 'Facturado']
    ]);
    sh.setFrozenRows(1);
  }
  return ok({ name: sh.getName() });
}

/** Construye menu segun rol del usuario. */
function buildMenu_() {
  const ui = safeUi_();
  if (!ui) return;

  const email = getCurrentUserEmail_();
  const role = getUserRole_(email);

  const menu = ui.createMenu(CC_CONFIG.USER_MENU_NAME)
    .addItem('Estado del sistema', 'showSystemStatus_');

  if (role === CC_CONFIG.ROLES.ADMIN) {
    ui.createMenu(CC_CONFIG.ADMIN_MENU_NAME)
      .addItem('Inicializar sistema', 'initSystem')
      .addItem('Aplicar protecciones', 'applySheetProtections_')
      .addItem('Reaplicar protecciones', 'reapplyProtections_')
      .addItem('Sincronizar permisos', 'syncPermissions_')
      .addItem('Validar estructura', 'validateMainSheetStructure_')
      .addItem('Sincronizar hojas usuario -> principal', 'syncAllUserSheetsToDatos')
      .addSeparator()
      .addItem('Aplicar validaciones/formatos (principal)', 'initValidationsAndFormats')
      .addItem('Aplicar validaciones/formatos (usuarios)', 'refreshFormatsAllUserSheets')
      .addSeparator()
      .addItem('Descartar >=90 dias (ahora)', 'autoDescartar90Dias')
      .addItem('Instalar job diario descarte', 'installDailyAutoDiscard_')
      .addToUi();
  }

  menu.addToUi();
}

/** Muestra estado general para diagnostico rapido. */
function showSystemStatus_() {
  const checks = {
    main: !!findSheetByNameSafe_(CC_CONFIG.MAIN_SHEET_NAME),
    users: !!findSheetByNameSafe_(CC_CONFIG.USERS_SHEET_NAME),
    logs: !!findSheetByNameSafe_(CC_CONFIG.LOGS_SHEET_NAME),
    config: !!findSheetByNameSafe_(CC_CONFIG.SYSTEM_SHEET_NAME)
  };

  const msg = [
    'Hoja principal: ' + (checks.main ? 'OK' : 'FALTA'),
    'Hoja usuarios: ' + (checks.users ? 'OK' : 'FALTA'),
    'Hoja logs: ' + (checks.logs ? 'OK' : 'FALTA'),
    'Hoja config: ' + (checks.config ? 'OK' : 'FALTA')
  ].join('\n');

  const ui = safeUi_();
  if (ui) {
    ui.alert('Estado del sistema', msg, ui.ButtonSet.OK);
  } else {
    toast_(msg, 'Estado del sistema');
  }

  return ok(checks);
}

/** Refresca cache de usuarios/permisos desde hoja Config_Usuarios. */
function syncPermissions_() {
  const cache = CacheService.getScriptCache();
  cache.remove('CC_USERS_MAP');
  appendLog_({
    usuario: getCurrentUserEmail_(),
    accion: 'syncPermissions',
    resultado: 'ok',
    detalle: 'Cache de permisos limpiada'
  });
  return ok({ refreshed: true });
}
