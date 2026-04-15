/** Sincroniza ediciones de hojas @usuario → "Datos" y aplica reglas. */
const CC_TZ = 'America/Bogota';

/* ====== UI helper ====== */
function _uiAlert_(title, message) {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ====== Helpers base ====== */
function _ss() { return SpreadsheetApp.getActive(); }
function _hoja(n) { return _ss().getSheetByName(n); }

function _nextId_(hojaDatos) {
  const last = hojaDatos.getLastRow();
  if (last < 2) return 1;
  const colA = hojaDatos.getRange(2, 1, last - 1, 1).getValues().flat();
  const nums = colA.filter(v => typeof v === 'number');
  return nums.length ? Math.max.apply(null, nums) + 1 : 1;
}

function _isUserSheet_(name) {
  return name !== 'Datos' && /@/.test(name);
}

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

/* ====== Fechas / Descarte 90d ====== */
function _daysDiff_(fromDate, toDate) {
  if (!fromDate || !toDate) return NaN;
  const d1 = new Date(Utilities.formatDate(new Date(fromDate), CC_TZ, 'yyyy-MM-dd'));
  const d2 = new Date(Utilities.formatDate(new Date(toDate),   CC_TZ, 'yyyy-MM-dd'));
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

/* ====== Validaciones y formatos ====== */
function initValidationsAndFormats() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Datos');
  if (!sh) throw new Error('No existe la hoja "Datos".');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return; // sin datos

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

  // Formato condicional (cubre desde fila 2)
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

  // P con wrap
  sh.getRange(2, 16, Math.max(lastRow-1, 1), 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

function ensureRowValidation_(sh, row) {
  if (row < 2) return;
  const dvEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(['En curso','Descartada','Aprobada','Rechazada'], true)
    .setAllowInvalid(false).build();
  const dvProceso = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Sin facturar','Facturado'], true)
    .setAllowInvalid(false).build();

  sh.getRange(row, 8).setDataValidation(dvEstado);   // H
  sh.getRange(row, 14).setDataValidation(dvProceso); // N
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

  // P con wrap
  sh.getRange(2, 16, Math.max(lastRow-1, 1), 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

function refreshFormatsAllUserSheets() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets()
    .filter(s => s.getName() !== 'Datos' && /@/.test(s.getName()))
    .forEach(applyValidationAndFormatsToSheet_);
}

/* ====== Validación de factura multi-línea (1..8 chars por línea) ====== 
const INVOICE_MAX_LEN = 8;
function validateInvoices_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return { ok: false, normalized: '', invalid: ['(vacío)'] };

  const lines = raw.split(/\r?\n/);
  const normalized = [];
  const invalid = [];

  lines.forEach((ln, i) => {
    const code = (ln || '').replace(/[^A-Za-z0-9]/g, '').toUpperCase();
    if (code.length >= 1 && code.length <= INVOICE_MAX_LEN) {
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

/* ====== Validación de factura multi-línea (SIN VALIDACIÓN, TODO SE ACEPTA) ====== */
//const INVOICE_MAX_LEN = 8; // ya no se usa, pero lo dejamos por compatibilidad

function validateInvoicesRelaxed_(value) {
  const raw = String(value == null ? '' : value).trim();
  return {
    ok: true,
    normalized: raw,
    invalid: []
  };
}

/* ====== Chequeo puntual de 90 días para una fila editada ====== 
function checkAndAutoDiscardByDate_(sheet, row) {
  try {
    if (row < 2) return;

    const ss        = SpreadsheetApp.getActive();
    const hojaDatos = ss.getSheetByName('Datos');
    if (!hojaDatos) return;

    const fecha   = sheet.getRange(row, 2).getValue();      // B
    const estado  = String(sheet.getRange(row, 8).getValue()  || ''); // H
    const proceso = String(sheet.getRange(row, 14).getValue() || ''); // N
    if (!fecha) return;

    // Solo descartar si NO está Aprobada y NO está Facturado
    const estadoOk  = estado.trim().toLowerCase() !== 'aprobada';
    const procesoOk = proceso.trim().toLowerCase() !== 'facturado';

    const hoy  = new Date();
    const dias = _daysDiff_(fecha, hoy);
    if (isNaN(dias) || dias < 90 || !estadoOk || !procesoOk) return;

    // Marcar "Descartada" en la hoja actual
    _setDescartadaSafe_(sheet, row);

    // Reflejar en la otra hoja por ID
    const id = sheet.getRange(row, 1).getValue(); // A
    if (!id) return;

    if (sheet.getName() === 'Datos') {
      const correo = sheet.getRange(row, 18).getValue(); // R
      const userSheet = correo ? ss.getSheetByName(safeSheetNameFromEmail_(correo)) : null;
      if (userSheet) {
        const rUser = _findRowById_(userSheet, id);
        if (rUser > 0) _setDescartadaSafe_(userSheet, rUser);
      }
    } else {
      const rDatos = _findRowById_(hojaDatos, id);
      if (rDatos > 0) _setDescartadaSafe_(hojaDatos, rDatos);
    }
  } catch (err) {
    console.error('checkAndAutoDiscardByDate_ →', err);
  }
}*/

/* ====== onEdit: validaciones, sincronía y reglas ====== */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sh        = e.range.getSheet();
    const sheetName = sh.getName();
    const row       = e.range.getRow();
    const col       = e.range.getColumn();
    if (row === 1) return; // encabezados

    const ss        = SpreadsheetApp.getActive();
    const hojaDatos = ss.getSheetByName('Datos');
    if (!hojaDatos) {
      _uiAlert_(
        '⚠️ Configuración incompleta',
        'No existe la hoja "Datos". Crea una hoja llamada exactamente "Datos" con los encabezados definidos.'
      );
      return;
    }

    // --- Formato automático de fechas en B, L, M, O ---
    if ([2, 12, 13, 15].includes(col)) {
      e.range.setNumberFormat('dd/MM/yyyy');
    }

    // ===================== Permisos por rol =====================
    const email = Session.getActiveUser().getEmail();
    const role  = (typeof getUserRole === 'function') ? getUserRole(email) : 'admin';
    const ui    = (typeof safeUi_ === 'function') ? safeUi_() : SpreadsheetApp.getUi();

    const colsUsuario  = [8, 9, 10, 11, 12, 13]; // H–M
    const colsFinanzas = [14, 15, 16];           // N–P

    // Rol sin permisos
    if (role === 'auditor' || !role) {
      if (ui) ui.alert('Acceso denegado', 'No tienes permisos para editar esta hoja.', ui.ButtonSet.OK);
      e.range.setValue(e.oldValue || '');
      return;
    }

    // Solo admin puede editar "Datos"
    if (sheetName === 'Datos' && role !== 'admin') {
      if (ui) ui.alert('Edición bloqueada', 'No puedes editar directamente la hoja "Datos".', ui.ButtonSet.OK);
      e.range.setValue(e.oldValue || '');
      return;
    }

    // No admin → solo puede editar hojas de usuario (@correo)
    if (role !== 'admin' && !_isUserSheet_(sheetName)) {
      if (ui) ui.alert('Edición bloqueada', 'Solo puedes editar hojas de usuario.', ui.ButtonSet.OK);
      e.range.setValue(e.oldValue || '');
      return;
    }

    // Restricciones para usuario
    if (role === 'usuario') {
      const mySheet = safeSheetNameFromEmail_(email).toLowerCase();
      if (sheetName.toLowerCase() !== mySheet) {
        if (ui) ui.alert('Edición bloqueada', 'Solo puedes editar tu propia hoja.', ui.ButtonSet.OK);
        e.range.setValue(e.oldValue || '');
        return;
      }
      if (!colsUsuario.includes(col)) {
        if (ui) ui.alert('Edición bloqueada', 'Solo puedes editar columnas H–M.', ui.ButtonSet.OK);
        e.range.setValue(e.oldValue || '');
        return;
      }
    }

    // Restricciones para financiera
    if (role === 'financiera' && !colsFinanzas.includes(col)) {
      if (ui) ui.alert('Edición bloqueada', 'Solo puedes editar columnas N–P.', ui.ButtonSet.OK);
      e.range.setValue(e.oldValue || '');
      return;
    }

    // ===================== Lógica de negocio =====================

    // A) Número(s) de factura (P=16)
    //    Validación (ahora relajada) y sincronía con "Datos"
    if (_isUserSheet_(sheetName) && col === 16) {
      // 🔹 Usamos la versión relajada: siempre ok, sin popups de error
      const check = validateInvoicesRelaxed_(e.range.getValue());

      if (!check.ok) {
        // Este bloque ya no se ejecutará nunca, pero se deja por compatibilidad
        setValueSafeWithValidation_(e.range, '');
        ensureWrap_(e.range);
        _uiAlert_(
          '⚠️ Validación de Número(s) de Factura',
          'Cada línea debe tener entre 1 y 8 caracteres alfanuméricos (A–Z, 0–9).\n\n' +
          (check.invalid && check.invalid.length
            ? 'Líneas inválidas:\n• ' + check.invalid.join('\n• ')
            : 'Ejemplos válidos:\n• ABC12345\n• 12A3B\n• 12345678')
        );
        return;
      }

      // Válido → normaliza y fuerza estados
      setValueSafeWithValidation_(e.range, check.normalized);
      ensureWrap_(e.range);
      setValueSafeWithValidation_(sh.getRange(row, 14), 'Facturado'); // N
      setValueSafeWithValidation_(sh.getRange(row, 8),  'Aprobada');  // H
      ensureRowValidation_(sh, row);

      // Reflejar en "Datos"
      const id = sh.getRange(row, 1).getValue();
      if (id) {
        const last = hojaDatos.getLastRow();
        const ids  = last > 1
          ? hojaDatos.getRange(2, 1, last - 1, 1).getValues().flat()
          : [];
        const idx  = ids.findIndex(v => v === id);
        if (idx >= 0) {
          const rowDatos = idx + 2;
          setValueSafeWithValidation_(hojaDatos.getRange(rowDatos, 16), check.normalized); // P
          ensureWrap_(hojaDatos.getRange(rowDatos, 16));
          setValueSafeWithValidation_(hojaDatos.getRange(rowDatos, 14), 'Facturado');       // N
          setValueSafeWithValidation_(hojaDatos.getRange(rowDatos, 8),  'Aprobada');        // H
          ensureRowValidation_(hojaDatos, rowDatos);
        }
      }

      checkAndAutoDiscardByDate_(sh, row);
      return; // manejado
    }

    // B) Estado del proceso (N=14) → si "Facturado", exigir P válido; si ok, H = "Aprobada"
    if (_isUserSheet_(sheetName) && col === 14) {
      const val = String(e.range.getValue() || '').trim().toLowerCase();

      if (val === 'facturado') {
        const pVal = String(sh.getRange(row, 16).getValue() || '');
        // 🔹 También usamos la versión relajada aquí
        const chk  = validateInvoicesRelaxed_(pVal);

        if (!chk.ok) {
          // Nunca se ejecutará con la validación relajada,
          // pero mantenemos la lógica original por compatibilidad
          setValueSafeWithValidation_(sh.getRange(row, 14), 'Sin facturar');
          ensureRowValidation_(sh, row);
          _uiAlert_(
            '⚠️ No puedes marcar "Facturado"',
            'Antes debes ingresar el/los número(s) de factura en la columna P (uno por línea), ' +
            'cada uno con entre 1 y 8 caracteres alfanuméricos.'
          );
          checkAndAutoDiscardByDate_(sh, row);
          return;
        }

        setValueSafeWithValidation_(sh.getRange(row, 8), 'Aprobada');
        ensureRowValidation_(sh, row);

        const id = sh.getRange(row, 1).getValue();
        if (id) {
          const last = hojaDatos.getLastRow();
          const ids  = last > 1
            ? hojaDatos.getRange(2, 1, last - 1, 1).getValues().flat()
            : [];
          const idx  = ids.findIndex(v => v === id);
          if (idx >= 0) {
            const rowDatos = idx + 2;
            setValueSafeWithValidation_(hojaDatos.getRange(rowDatos, 14), 'Facturado'); // N
            setValueSafeWithValidation_(hojaDatos.getRange(rowDatos, 8),  'Aprobada');  // H
            ensureRowValidation_(hojaDatos, rowDatos);
          }
        }
      }
      // si no es "Facturado", no forzamos H
    }

    // C) Sincronía general: @usuario → Datos (cualquier edición)
    if (_isUserSheet_(sheetName)) {
      const cols     = hojaDatos.getLastColumn();
      const startRow = e.range.getRow();
      const numRows  = e.range.getNumRows();
      const userRows = sh.getRange(startRow, 1, numRows, cols).getValues();

      const datosLast = hojaDatos.getLastRow();
      const datosIds  = datosLast > 1
        ? hojaDatos.getRange(2, 1, datosLast - 1, 1).getValues().flat()
        : [];

      for (let i = 0; i < userRows.length; i++) {
        const filaUsuario = userRows[i];
        const id = filaUsuario[0];
        if (!id) continue;

        const idx = datosIds.findIndex(v => v === id);
        if (idx >= 0) {
          hojaDatos.getRange(idx + 2, 1, 1, cols).setValues([filaUsuario]);
          ensureRowValidation_(hojaDatos, idx + 2);
        } else {
          hojaDatos.appendRow(filaUsuario);
          const newRow = hojaDatos.getLastRow();
          ensureRowValidation_(hojaDatos, newRow);
        }
      }
    }

    // D) Descarte por 90 días para la fila editada
    //checkAndAutoDiscardByDate_(sh, row);

    // Estilos (si tienes esta función en otro archivo)
    if (typeof aplicarEstilosATodas === 'function') {
      aplicarEstilosATodas();
    }

  } catch (err) {
    console.error('❌ Error en onEdit sync →', err);
    _uiAlert_(
      '❌ Error en Control Comercial',
      (err && err.message) ? String(err.message) : 'Ocurrió un error no especificado durante la sincronización.'
    );
  }
}


/**
 * Sincroniza todas las hojas de usuario (@correo) con la hoja principal "Datos".
 * Copia o actualiza filas según el ID (columna A).
 */
/** Normaliza posibles fechas (B=2, L=12, M=13, O=15) a Date o '' */
function _normalizeRowForDatos_CC_(fila) {
  const DATE_COLS = [2, 12, 13, 15]; // 1-based
  const out = fila.slice(); // copia superficial

  DATE_COLS.forEach(c => {
    const idx = c - 1;
    const v = out[idx];
    if (v === '' || v == null) { out[idx] = ''; return; }
    // Si ya es Date válido, conservar
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return;
    // Intentar parseo prudente
    const d = new Date(v);
    out[idx] = (d && !isNaN(d)) ? d : '';
  });

  return out;
}

/** Escribe una fila completa, celda por celda, respetando validaciones */
function setRowValuesSafe_CC_(sheet, row, values) {
  for (let c = 1; c <= values.length; c++) {
    const cell = sheet.getRange(row, c);
    setValueSafeWithValidation_(cell, values[c - 1]);
  }
}

/**
 * Sincroniza todas las hojas de usuario (@correo) con la hoja principal "Datos".
 * - Actualiza por ID (col A); inserta si no existe
 * - Convierte fechas (B, L, M, O) a Date
 * - Evita errores de validación con escritura celda-a-celda cuando es necesario
 */
function syncAllUserSheetsToDatos() {
  const ss = SpreadsheetApp.getActive();
  const hojaDatos = ss.getSheetByName('Datos');
  if (!hojaDatos) {
    SpreadsheetApp.getUi().alert('No existe la hoja "Datos".');
    return;
  }

  const colCount = hojaDatos.getLastColumn();
  const userSheets = ss.getSheets().filter(s => /@/.test(s.getName()));

  // IDs ya presentes en Datos
  const lastDatos = hojaDatos.getLastRow();
  const idsDatos = lastDatos > 1
    ? hojaDatos.getRange(2, 1, lastDatos - 1, 1).getValues().flat()
    : [];

  let totalActualizados = 0;
  let totalInsertados = 0;

  userSheets.forEach(sh => {
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return; // sin datos

    const data = sh.getRange(2, 1, lastRow - 1, colCount).getValues();

    data.forEach(fila => {
      const id = fila[0];
      if (!id) return;

      // Normalizar antes de escribir (fechas)
      const filaNorm = _normalizeRowForDatos_CC_(fila);

      const idx = idsDatos.findIndex(v => v === id);
      if (idx >= 0) {
        // Actualizar fila existente
        const targetRow = idx + 2;
        try {
          hojaDatos.getRange(targetRow, 1, 1, colCount).setValues([filaNorm]);
        } catch (e) {
          // Si hay validaciones que bloquean, escribir celda a celda
          setRowValuesSafe_CC_(hojaDatos, targetRow, filaNorm);
        }
        ensureRowValidation_(hojaDatos, targetRow);
        if (colCount >= 16) hojaDatos.getRange(targetRow, 16).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        totalActualizados++;
      } else {
        // Insertar nueva fila
        const newRow = hojaDatos.getLastRow() + 1;
        try {
          hojaDatos.getRange(newRow, 1, 1, colCount).setValues([filaNorm]);
        } catch (e) {
          setRowValuesSafe_CC_(hojaDatos, newRow, filaNorm);
        }
        ensureRowValidation_(hojaDatos, newRow);
        if (colCount >= 16) hojaDatos.getRange(newRow, 16).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        totalInsertados++;
        idsDatos.push(id);
      }
    });
  });

  // Opcional: re-aplicar estilos/condicionales si tienes estas funciones
  if (typeof aplicarEstilosEnHoja === 'function') {
    try { aplicarEstilosEnHoja(hojaDatos); } catch (_) {}
  }

  SpreadsheetApp.getUi().alert(
    '✅ Sincronización completada',
    `Actualizados: ${totalActualizados}\nNuevos: ${totalInsertados}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

