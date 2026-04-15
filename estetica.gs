function aplicarEstilosEnHoja(hoja) {
  if (!hoja) return;

  const lastRow = hoja.getLastRow();
  const lastCol = hoja.getLastColumn();
  if (lastCol < 1) return;

  /* 0) Limpiar bandings previos (evita el error de "ya tiene colores alternos") */
  try {
    hoja.getBandings().forEach(b => b.remove());
  } catch (err) {
    // Ignoramos cualquier detalle aquí para que nunca bloquee
    console.error('aplicarEstilosEnHoja.removeBandings →', err);
  }

  /* 1) Encabezado */
  const hdr = hoja.getRange(1, 1, 1, lastCol);
  hoja.setFrozenRows(1);
  hdr
    .setFontWeight('bold')
    .setBackground('#023047')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  /* 2) Anchos sugeridos por columna (A..R) */
  const widths = [60,110,120,160,260,120,120,120,200,120,120,110,110,140,120,120,170,240];
  const limit = Math.min(widths.length, lastCol);
  for (let c = 1; c <= limit; c++) {
    try { hoja.setColumnWidth(c, widths[c-1]); } catch (_) {}
  }

  /* 3) Si no hay filas de datos, terminamos */
  const dataRows = Math.max(0, lastRow - 1);
  if (dataRows === 0) return;

  const datos = hoja.getRange(2, 1, dataRows, lastCol);

  // Alineación / ajuste general
  datos
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  /* 4) Cebra por bloques (NO tocamos H=8 ni N=14 para que sus colores condicionales se vean) */
  const applyCebra = (colStart, colCount) => {
    if (colStart > lastCol || colCount <= 0) return;
    const count = Math.min(colCount, lastCol - colStart + 1);
    try {
      hoja
        .getRange(2, colStart, dataRows, count)
        .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    } catch (err) {
      console.error('aplicarEstilosEnHoja.applyCebra →', err);
    }
  };

  // Bloques: A:G, I:M, O:last
  applyCebra(1, 7);                         // A..G
  if (lastCol >= 9)  applyCebra(9, 5);      // I..M
  if (lastCol >= 15) applyCebra(15, lastCol - 14); // O..last

  // Bordes externos
  try {
    datos.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  } catch (_) {}

  /* 5) Formatos de fecha: B, L, M, O */
  try {
    if (lastCol >= 2)  hoja.getRange(2,  2, dataRows, 1).setNumberFormat('dd/MM/yyyy'); // B
    if (lastCol >= 12) hoja.getRange(2, 12, dataRows, 1).setNumberFormat('dd/MM/yyyy'); // L
    if (lastCol >= 13) hoja.getRange(2, 13, dataRows, 1).setNumberFormat('dd/MM/yyyy'); // M
    if (lastCol >= 15) hoja.getRange(2, 15, dataRows, 1).setNumberFormat('dd/MM/yyyy'); // O
  } catch (_) {}

  /* 6) Altura de filas suave */
  try {
    for (let r = 2; r <= lastRow; r++) hoja.setRowHeight(r, 28);
  } catch (_) {}

  /* 7) Validaciones (sin tocar reglas condicionales existentes) */
  try {
    if (lastCol >= 8) {
      const dvEstado = SpreadsheetApp.newDataValidation()
        .requireValueInList(['En curso','Descartada','Aprobada','Rechazada'], true)
        .setAllowInvalid(false)
        .build();
      hoja.getRange(2, 8, dataRows, 1).setDataValidation(dvEstado);
    }
    if (lastCol >= 14) {
      const dvProceso = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Sin facturar','Facturado'], true)
        .setAllowInvalid(false)
        .build();
      hoja.getRange(2, 14, dataRows, 1).setDataValidation(dvProceso);
    }
  } catch (e) {
    console.error('aplicarEstilosEnHoja.setValidations →', e);
  }

  /* 8) Columna P (número(s) de factura) con WRAP para saltos de línea */
  if (lastCol >= 16) {
    try { hoja.getRange(2, 16, dataRows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); } catch (_) {}
  }
}

/** Aplica estilos a TODAS las hojas (Datos y todas las que contengan "@") */
function aplicarEstilosATodas() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (name === 'Datos' || /@/.test(name)) {
      aplicarEstilosEnHoja(sh);
    }
  });
}
