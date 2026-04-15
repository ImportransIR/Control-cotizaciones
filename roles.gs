/** ===========================
 *  ROLES Y UTILIDADES GENERALES
 *  =========================== */

function getUserRole(email) {
  const e = String(email || '').toLowerCase().trim();

  const admins = [
    "tics@importransradiactivos.com",
    "analista.tics@importransradiactivos.com"
  ];

  const financieras = [
    "facturacion@importransradiactivos.com",
    "financiera@importransradiactivos.com",
    "asistenteadministrativo@importransradiactivos.com"
  ];

  const usuarios = [
    "info@importransradiactivos.com",
    "industria1@importransradiactivos.com",
    "comercial@importransradiactivos.com",
    "proteccionradiologica@importransradiactivos.com",
    "asesorias@importransradiactivos.com",
    "radioproteccion@importransradiactivos.com",
    "asistentehseq@importransradiactivos.com"
  ];

  const auditores = [
    "asistentegerencia@importransradiactivos.com",
    "saulacero@importransradiactivos.com",
    "contabilidad@importransradiactivos.com",
    "jmanosalva@importransradiactivos.com",
  ];

  if (admins.includes(e)) return "admin";
  if (financieras.includes(e)) return "financiera";
  if (auditores.includes(e)) return "auditor";
  if (usuarios.includes(e)) return "usuario";
  return "usuario";
}

function safeUi_() {
  try { return SpreadsheetApp.getUi(); } catch (e) { return null; }
}

// Nombre seguro de hoja a partir del correo
function safeSheetNameFromEmail_(email) {
  if (!email) return 'usuario_sin_correo';
  return String(email).replace(/[^A-Za-z0-9@._-]/g, '_').slice(0, 99);
}

// Regresa true si es una hoja de usuario (correo@...)
function _isUserSheet_(name) {
  return name !== 'Datos' && /@/.test(name);
}

// Diferencia de días normalizando a medianoche local
function _daysDiff_(fromDate, toDate) {
  if (!fromDate || !toDate) return NaN;
  const tz = 'America/Bogota';
  const d1 = new Date(Utilities.formatDate(new Date(fromDate), tz, 'yyyy-MM-dd'));
  const d2 = new Date(Utilities.formatDate(new Date(toDate),   tz, 'yyyy-MM-dd'));
  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  return Math.floor((d2 - d1) / MS_PER_DAY);
}

// Alerta centrada
function _uiAlert_(title, message) {
  const ui = safeUi_();
  if (ui) ui.alert(title, message, ui.ButtonSet.OK);
}

// Set value sin chocar validaciones (restaura luego)
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

// Wrap para ver saltos de línea
function ensureWrap_(range) {
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

// Validación/normalización de números de factura: N líneas, cada una 5 alfanuméricos
function validateInvoices_(value) {
  const s = String(value == null ? '' : value);
  const lines = s.split(/\r?\n/).map(x => x.trim()).filter(x => x.length > 0);

  if (lines.length === 0) {
    return { ok: false, normalized: '', invalid: ['(vacío)'] };
  }

  const normalized = [];
  const invalid = [];

  lines.forEach(line => {
    const cleaned = line.replace(/[^A-Za-z0-9]/g, '').toUpperCase();
    if (cleaned.length === 5) {
      normalized.push(cleaned);
    } else {
      invalid.push(line);
    }
  });

  return {
    ok: invalid.length === 0,
    normalized: normalized.join('\n'),
    invalid
  };
}

// Marca "Descartada" con respeto por validaciones
function _setDescartadaSafe_(sheet, row) {
  const r = sheet.getRange(row, 8); // H
  setValueSafeWithValidation_(r, 'Descartada');
}

// Busca fila por ID (col A)
function _findRowById_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
  const idx = ids.findIndex(v => v === id);
  return (idx >= 0) ? (idx + 2) : -1;
}
