/** ========= ESTILOS & VALIDACIONES ========= */

function initValidationsAndFormats() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CC_CONFIG.MAIN_SHEET_NAME);
  if (!sh) throw new Error('No existe la hoja principal.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const headerMap = getHeaderMap_(sh);
  const estadoCol = headerMap[normalizeHeader_('Estado')];
  const procesoCol = headerMap[normalizeHeader_('Estado del proceso')];

  if (estadoCol) {
    const rngEstado = sh.getRange(2, estadoCol, lastRow - 1, 1);
    const dvEstado = SpreadsheetApp.newDataValidation()
      .requireValueInList(['En curso', 'Descartada', 'Aprobada', 'Rechazada'], true)
      .setAllowInvalid(false)
      .build();
    rngEstado.setDataValidation(dvEstado);
  }

  if (procesoCol) {
    const rngProceso = sh.getRange(2, procesoCol, lastRow - 1, 1);
    const dvProceso = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Sin facturar', 'Facturado'], true)
      .setAllowInvalid(false)
      .build();
    rngProceso.setDataValidation(dvProceso);
  }
}

function ensureRowValidation_(sh, row) {
  if (!sh || row < 2) return;

  const headerMap = getHeaderMap_(sh);
  const estadoCol = headerMap[normalizeHeader_('Estado')];
  const procesoCol = headerMap[normalizeHeader_('Estado del proceso')];

  if (estadoCol) {
    sh.getRange(row, estadoCol).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['En curso', 'Descartada', 'Aprobada', 'Rechazada'], true)
        .setAllowInvalid(false)
        .build()
    );
  }

  if (procesoCol) {
    sh.getRange(row, procesoCol).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Sin facturar', 'Facturado'], true)
        .setAllowInvalid(false)
        .build()
    );
  }
}

function applyValidationAndFormatsToSheet_(sh) {
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  for (let r = 2; r <= lastRow; r++) {
    ensureRowValidation_(sh, r);
  }

  aplicarEstilosEnHoja(sh);
}

function refreshFormatsAllUserSheets() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets()
    .filter(function (s) { return isUserSheet_(s.getName()); })
    .forEach(applyValidationAndFormatsToSheet_);
}

function aplicarEstilosEnHoja(hoja) {
  if (!hoja) return;

  const lastRow = hoja.getLastRow();
  const lastCol = hoja.getLastColumn();
  if (lastCol < 1) return;

  try {
    hoja.getBandings().forEach(function (b) { b.remove(); });
  } catch (error) {
    console.error('aplicarEstilosEnHoja.removeBandings:', error);
  }

  const hdr = hoja.getRange(1, 1, 1, lastCol);
  hoja.setFrozenRows(1);
  hdr
    .setFontWeight('bold')
    .setBackground('#023047')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const widths = [60, 110, 120, 160, 260, 120, 120, 120, 200, 120, 120, 110, 110, 140, 120, 120, 170, 240];
  const limit = Math.min(widths.length, lastCol);
  for (let c = 1; c <= limit; c++) {
    try { hoja.setColumnWidth(c, widths[c - 1]); } catch (error) {}
  }

  const dataRows = Math.max(0, lastRow - 1);
  if (dataRows === 0) return;

  const datos = hoja.getRange(2, 1, dataRows, lastCol);
  datos
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  const applyCebra = function (colStart, colCount) {
    if (colStart > lastCol || colCount <= 0) return;
    const count = Math.min(colCount, lastCol - colStart + 1);
    try {
      hoja
        .getRange(2, colStart, dataRows, count)
        .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    } catch (error) {
      console.error('aplicarEstilosEnHoja.applyCebra:', error);
    }
  };

  applyCebra(1, 7);
  if (lastCol >= 9) applyCebra(9, 5);
  if (lastCol >= 15) applyCebra(15, lastCol - 14);

  try {
    datos.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  } catch (error) {}

  const headerMap = getHeaderMap_(hoja);
  ['Fecha de solicitud', 'Fecha de inicio', 'Fecha final', 'Fecha de factura'].forEach(function (h) {
    const idx = headerMap[normalizeHeader_(h)];
    if (idx) {
      try { hoja.getRange(2, idx, dataRows, 1).setNumberFormat('dd/MM/yyyy'); } catch (error) {}
    }
  });

  try {
    for (let r = 2; r <= lastRow; r++) hoja.setRowHeight(r, 28);
  } catch (error) {}

  const facturaCol = headerMap[normalizeHeader_('Numero factura')];
  if (facturaCol) {
    try { hoja.getRange(2, facturaCol, dataRows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); } catch (error) {}
  }
}

function aplicarEstilosATodas() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(function (sh) {
    const name = sh.getName();
    if (name === CC_CONFIG.MAIN_SHEET_NAME || isUserSheet_(name)) {
      aplicarEstilosEnHoja(sh);
    }
  });
}
