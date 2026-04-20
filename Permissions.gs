/** Obtiene correo del usuario actual con fallback seguro. */
function getCurrentUserEmail_() {
  const active = (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (active) return active;
  const effective = (Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  return effective;
}

/** Lee hoja de usuarios y crea mapa correo -> metadata. */
function getUsersMap_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('CC_USERS_MAP');
  if (cached) {
    try { return JSON.parse(cached); } catch (error) {}
  }

  const sh = findSheetByNameSafe_(CC_CONFIG.USERS_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return {};

  const headerMap = getHeaderMap_(sh);
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const out = {};

  values.forEach(function (row) {
    const email = String(row[(headerMap.correo || 1) - 1] || '').trim().toLowerCase();
    if (!email) return;

    const role = String(row[(headerMap.rol || 3) - 1] || CC_CONFIG.ROLES.USUARIO).trim().toLowerCase();
    const activoRaw = String(row[(headerMap.activo || 4) - 1] || 'si').trim().toLowerCase();
    const puedeEditarRaw = String(row[(headerMap.puedeeditar || 5) - 1] || 'si').trim().toLowerCase();

    out[email] = {
      correo: email,
      nombre: String(row[(headerMap.nombre || 2) - 1] || '').trim(),
      rol: role,
      activo: ['si', 'true', '1', 'x'].includes(activoRaw),
      puedeEditar: ['si', 'true', '1', 'x'].includes(puedeEditarRaw),
      observaciones: String(row[(headerMap.observaciones || 6) - 1] || '').trim()
    };
  });

  cache.put('CC_USERS_MAP', JSON.stringify(out), 300);
  return out;
}

/** Determina rol por correo consultando hoja Config_Usuarios. */
function getUserRole_(email) {
  const mail = String(email || getCurrentUserEmail_() || '').trim().toLowerCase();
  if (!mail) return CC_CONFIG.ROLES.AUDITOR;

  const usersMap = getUsersMap_();
  const user = usersMap[mail];

  if (!user || !user.activo) return CC_CONFIG.ROLES.AUDITOR;
  if (!user.puedeEditar && user.rol !== CC_CONFIG.ROLES.ADMIN) return CC_CONFIG.ROLES.AUDITOR;

  if (!Object.values(CC_CONFIG.ROLES).includes(user.rol)) {
    return CC_CONFIG.ROLES.USUARIO;
  }
  return user.rol;
}

/** Compatibilidad con funciones existentes. */
function getUserRole(email) {
  return getUserRole_(email);
}

/** Indica si un correo pertenece a admin. */
function isAdmin_(email) {
  return getUserRole_(email) === CC_CONFIG.ROLES.ADMIN;
}

/** Obtiene columnas permitidas por rol en base a encabezados reales. */
function getAllowedColumnsByRole_(role, headerMap) {
  if (role === CC_CONFIG.ROLES.ADMIN) {
    return Object.keys(headerMap).map(function (k) { return headerMap[k]; });
  }

  const allowHeaders = CC_CONFIG.ROLE_ALLOWED_HEADERS[role] || [];
  return allowHeaders
    .map(function (h) { return headerMap[normalizeHeader_(h)]; })
    .filter(function (idx) { return typeof idx === 'number' && idx > 0; });
}

/** Evalua si rol puede editar una columna concreta. */
function canEditCell_(role, colIndex, headerMap) {
  if (role === CC_CONFIG.ROLES.ADMIN) return true;
  if (colIndex <= CC_CONFIG.HEADER_ROW) return false;

  const allowed = getAllowedColumnsByRole_(role, headerMap);
  return allowed.includes(colIndex);
}
