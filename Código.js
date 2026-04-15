function doGet() {
  return HtmlService.createHtmlOutputFromFile('formulario')
    .setTitle('Control Comercial Prueba')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function safeSheetNameFromEmail_(email) {
  if (!email) return 'usuario_sin_correo';
  return String(email).replace(/[^A-Za-z0-9@._-]/g, '_').slice(0, 99);
}

function ensureUserSheet_(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDatos = ss.getSheetByName('Datos');
  if (!hojaDatos) throw new Error('No existe la hoja "Datos".');

  const userSheetName = safeSheetNameFromEmail_(email);
  let sh = ss.getSheetByName(userSheetName);
  if (!sh) {
    sh = ss.insertSheet(userSheetName);
    const headers = hojaDatos.getRange(1,1,1,hojaDatos.getLastColumn()).getValues();
    sh.getRange(1,1,1,headers[0].length).setValues(headers);
    sh.setFrozenRows(1);
  }

  // Aplica validaciones/formatos a la hoja del usuario
  applyValidationAndFormatsToSheet_(sh);
  return sh;
}

function guardarDatos(form) {
  if (!form) throw new Error('No llegaron datos del formulario');

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Datos');
  if (!hoja) throw new Error('No existe una hoja llamada "Datos".');

  // CORREO del usuario (la Web App debe correr como "Tú" y acceso del dominio)
  const correo = (Session.getActiveUser().getEmail() || '').trim();
  if (!correo) {
    throw new Error('El correo llegó vacío. Revisa los permisos de la Web App.');
  }

  // Unir múltiples descripciones (vengan como string, array u objeto)
  let descripcion = '';
  const raw = form['descripcion[]'] ?? form.descripcion ?? form.detalle;

  if (Array.isArray(raw)) {
    descripcion = raw.map(v => String(v || '').trim()).filter(Boolean).join('\n');
  } else if (raw && typeof raw === 'object') {
    // por si llega como {0:"a",1:"b"} (no es Array real)
    const arr = Object.keys(raw).sort().map(k => raw[k]);
    descripcion = arr.map(v => String(v || '').trim()).filter(Boolean).join('\n');
  } else if (raw != null) {
    descripcion = String(raw).trim();
  }

  console.log('FORM keys:', Object.keys(form));
  console.log('RAW descripcion[]:', form['descripcion[]']);



  // Calcular ID y fila nueva
  const ultimaFila = hoja.getLastRow() + 1;
  let nuevoID = 1;
  if (hoja.getLastRow() > 1) {
    const colA = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues().flat();
    const nums = colA.filter(v => typeof v === 'number');
    nuevoID = nums.length ? Math.max(...nums) + 1 : 1;
  }

  const ahora = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss');

  // Fila (A..R)
  const fila = [
    nuevoID,                          // A ID
    form.fecha_solicitud || '',       // B Fecha de solicitud
    form.centro_costos || '',         // C Centro de costos
    form.cliente || '',               // D Cliente
    descripcion || '',                // E Descripción
    form.numero_cotizacion || '',     // F Número Cotización
    form.link_cotizacion || '',       // G Link
    'En curso',                       // H Estado (inicial)
    '',                               // I Observaciones
    '',                               // J PO/OC
    '',                               // K HE/ACTA
    '',                               // L Fecha de inicio
    '',                               // M Fecha final
    'Sin facturar',                   // N Estado del proceso
    '',                               // O Fecha de factura
    '',                               // P Número Factura (multi-línea, 5 chars c/u)
    ahora,                            // Q Fecha y Hora de Registro
    correo                            // R Correo electrónico
  ];

  // Guardar en "Datos"
  hoja.getRange(ultimaFila, 1, 1, fila.length).setValues([fila]);
  ensureRowValidation_(hoja, ultimaFila); // desplegables solo a esa fila
  hoja.getRange(ultimaFila, 16).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // P con wrap

  // Guardar en hoja del usuario
  const hojaUsuario = ensureUserSheet_(correo);
  hojaUsuario.appendRow(fila);
  const rowUser = hojaUsuario.getLastRow();
  ensureRowValidation_(hojaUsuario, rowUser);
  hojaUsuario.getRange(rowUser, 16).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // P con wrap

  aplicarEstilosEnHoja(hoja);        
  aplicarEstilosEnHoja(hojaUsuario); 


  return { ok: true, id: nuevoID, row: ultimaFila, sheet: hojaUsuario.getName() };
}

function autoDescartar90Dias() {
  const ss = SpreadsheetApp.getActive();
  const hojaDatos = ss.getSheetByName('Datos');
  if (!hojaDatos) {
    _uiAlert_('⚠️ Configuración incompleta', 'No existe la hoja "Datos".');
    return;
  }

  const last = hojaDatos.getLastRow();
  if (last < 2) return;

  const values = hojaDatos.getRange(2, 1, last - 1, hojaDatos.getLastColumn()).getValues(); // A..R
  const hoy = new Date();

  for (let i = 0; i < values.length; i++) {
    const rowIndex = i + 2;
    const id        = values[i][0];  // A
    const fecha     = values[i][1];  // B
    const estado    = String(values[i][7]  || ''); // H
    const proceso   = String(values[i][13] || ''); // N
    const correo    = values[i][17]; // R

    if (!fecha) continue;

    // Solo descartar si NO está Aprobada y NO está Facturado
    const estadoOk  = estado.trim().toLowerCase() !== 'aprobada';
    const procesoOk = proceso.trim().toLowerCase() !== 'facturado';

    const dias = _daysDiff_(fecha, hoy);
    if (!isNaN(dias) && dias >= 90 && estadoOk && procesoOk) {
      if (estado.trim() !== 'Descartada') {
        _setDescartadaSafe_(hojaDatos, rowIndex);

        // Reflejar en hoja de usuario por ID
        if (id && correo) {
          const userSheet = ss.getSheetByName(safeSheetNameFromEmail_(correo));
          if (userSheet) {
            const rUser = _findRowById_(userSheet, id);
            if (rUser > 0) _setDescartadaSafe_(userSheet, rUser);
          }
        }
      }
    }
  }

  aplicarEstilosATodas();

}

/* ===== Menú y trigger diario ===== */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Control Comercial')
    .addItem('Aplicar validaciones/formatos (Datos)', 'initValidationsAndFormats')
    .addItem('Aplicar a todas las hojas de usuario', 'refreshFormatsAllUserSheets')
    .addSeparator()
    .addItem('Descartar ≥90 días (ahora)', 'autoDescartar90Dias')
    .addItem('Instalar job diario (07:00 Bogotá)', 'installDailyAutoDiscard_')
    .addToUi();
}

function installDailyAutoDiscard_() {
  // Elimina triggers previos de la misma función
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'autoDescartar90Dias')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('autoDescartar90Dias')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .inTimezone('America/Bogota')
    .create();

  _uiAlert_('✅ Trigger instalado',
    'Se ejecutará "autoDescartar90Dias" todos los días a las 07:00 (America/Bogota).');
}
