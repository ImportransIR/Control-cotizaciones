/**
 * Base futura para formulario de clientes.
 * Esta capa no altera la operacion actual; solo prepara validaciones y utilidades.
 */

/** Valida payload de cliente para futura integracion de formulario web. */
function validateClientePayload_(payload) {
  const data = payload || {};
  const rutNit = String(data.rut_nit || data.rut || data.nit || '').trim();
  if (!rutNit) return err('RUT/NIT es obligatorio');

  return ok({
    rut_nit: rutNit,
    razon_social: String(data.razon_social || data.razonSocial || '').trim(),
    nombre_comercial: String(data.nombre_comercial || data.nombreComercial || '').trim(),
    carpeta_link: String(data.carpeta_link || '').trim(),
    activo: true,
    actualizado_en: new Date()
  });
}

/** Upsert de cliente por RUT/NIT en hoja Clientes (soporte futuro). */
function upsertCliente_(payload) {
  const validated = validateClientePayload_(payload);
  if (!validated.ok) return validated;

  const shResult = ensureClientesSupportSheet_();
  if (!shResult.ok) return shResult;

  const sh = findSheetByNameSafe_(CC_CONFIG.CLIENTES_SUPPORT_SHEET_NAME);
  const headerMap = getHeaderMap_(sh);
  const data = validated.data;

  const rutCol = headerMap[normalizeHeader_('rut_nit')] || 1;
  const last = sh.getLastRow();
  const values = last >= 2 ? sh.getRange(2, rutCol, last - 1, 1).getValues().flat() : [];
  const idx = values.findIndex(function (v) {
    return String(v || '').trim().toLowerCase() === data.rut_nit.toLowerCase();
  });

  const rowData = [
    data.rut_nit,
    data.razon_social,
    data.nombre_comercial,
    data.carpeta_link,
    data.activo,
    data.actualizado_en
  ];

  if (idx >= 0) {
    sh.getRange(idx + 2, 1, 1, rowData.length).setValues([rowData]);
    return ok({ action: 'updated', row: idx + 2 });
  }

  sh.appendRow(rowData);
  return ok({ action: 'inserted', row: sh.getLastRow() });
}
