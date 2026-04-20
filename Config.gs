/**
 * Configuracion central del sistema.
 * Se mantiene enfoque evolutivo: hoja principal existente + hojas auxiliares minimas.
 */
const CC_CONFIG = {
  MAIN_SHEET_NAME: 'Datos',
  USERS_SHEET_NAME: 'Config_Usuarios',
  SYSTEM_SHEET_NAME: 'Config_Sistema',
  LOGS_SHEET_NAME: 'Logs',
  CLIENTES_SUPPORT_SHEET_NAME: 'Clientes',
  CATALOGS_SHEET_NAME: 'Catalogos',
  HEADER_ROW: 1,
  TIMEZONE: 'America/Bogota',

  ROLES: {
    ADMIN: 'admin',
    USUARIO: 'usuario',
    FINANCIERA: 'financiera',
    AUDITOR: 'auditor'
  },

  USERS_HEADERS: ['correo', 'nombre', 'rol', 'activo', 'puedeEditar', 'observaciones'],
  LOG_HEADERS: [
    'fechaHora', 'usuario', 'accion', 'hoja', 'fila', 'columna',
    'valorAnterior', 'valorNuevo', 'resultado', 'detalle'
  ],

  EXPECTED_MAIN_HEADERS: [
    'id',
    'fecha de solicitud',
    'centro de costos',
    'cliente',
    'descripcion',
    'numero cotizacion',
    'link',
    'estado',
    'observaciones',
    'po/oc',
    'he/acta',
    'fecha de inicio',
    'fecha final',
    'estado del proceso',
    'fecha de factura',
    'numero factura'
  ],

  ROLE_ALLOWED_HEADERS: {
    usuario: [
      'estado',
      'observaciones',
      'po/oc',
      'he/acta',
      'fecha de inicio',
      'fecha final'
    ],
    financiera: [
      'estado del proceso',
      'fecha de factura',
      'numero factura'
    ],
    auditor: [],
    admin: ['*']
  },

  // Columnas estructurales que no deben ser editadas por no-admin.
  SENSITIVE_HEADERS: [
    'id',
    'fecha de solicitud',
    'centro de costos',
    'cliente',
    'descripcion',
    'numero cotizacion',
    'link',
    'fecha y hora de registro',
    'correo electronico'
  ],

  ADMIN_MENU_NAME: 'Control Comercial Admin',
  USER_MENU_NAME: 'Control Comercial',
  PROTECTION_TAG: 'CC_PROTECT'
};
