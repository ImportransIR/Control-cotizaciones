/**
 * Base futura para formulario de cotizaciones.
 * Mantiene la hoja principal como fuente de operacion.
 */

/** Valida payload minimo de cotizacion para uso futuro de formulario web. */
function validateCotizacionPayload_(payload) {
  const data = payload || {};

  const required = {
    fecha_solicitud: String(data.fecha_solicitud || '').trim(),
    centro_costos: String(data.centro_costos || '').trim(),
    cliente: String(data.cliente || '').trim(),
    descripcion: String(data.descripcion || '').trim(),
    numero_cotizacion: String(data.numero_cotizacion || '').trim()
  };

  const missing = Object.keys(required).filter(function (k) { return !required[k]; });
  if (missing.length) return err('Faltan campos obligatorios', missing);

  return ok({
    fecha_solicitud: required.fecha_solicitud,
    centro_costos: required.centro_costos,
    cliente: required.cliente,
    descripcion: required.descripcion,
    numero_cotizacion: required.numero_cotizacion,
    link: String(data.link || data.link_cotizacion || '').trim()
  });
}

/**
 * Placeholder de insercion futura.
 * Se mantiene separado para que formulario cliente/cotizacion se integren sin romper operacion actual.
 */
function registrarCotizacionDesdeFormulario_(payload) {
  const validated = validateCotizacionPayload_(payload);
  if (!validated.ok) return validated;

  // En esta etapa no se migra flujo; se reutiliza la funcion operativa existente.
  if (typeof guardarDatos !== 'function') {
    return err('No existe guardarDatos en el proyecto actual');
  }

  try {
    const res = guardarDatos({
      fecha_solicitud: validated.data.fecha_solicitud,
      centro_costos: validated.data.centro_costos,
      cliente: validated.data.cliente,
      descripcion: validated.data.descripcion,
      numero_cotizacion: validated.data.numero_cotizacion,
      link_cotizacion: validated.data.link
    });

    return ok(res);
  } catch (error) {
    return err('Error registrando cotizacion', String(error && error.message || error));
  }
}
