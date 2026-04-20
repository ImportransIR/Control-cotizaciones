# Control Comercial / Control de Cotizaciones

Sistema evolutivo en Google Sheets + Google Apps Script para gestionar cotizaciones sin rehacer la operacion actual.

## 1. Objetivo del sistema

Este proyecto permite:

- Registrar cotizaciones desde formulario web.
- Mantener una hoja maestra principal como fuente operativa.
- Sincronizar informacion con hojas de usuario cuando aplica.
- Aplicar formato y validaciones de negocio.
- Controlar permisos por rol.
- Proteger encabezados y columnas sensibles.
- Registrar trazabilidad de intentos no autorizados.
- Dejar base preparada para formularios futuros (Clientes y Cotizaciones) sin migracion agresiva.

## 2. Principios de implementacion

- No se rehace el sistema desde cero.
- No se altera de forma agresiva la estructura de la hoja principal.
- La hoja principal sigue siendo la fuente de verdad operativa.
- La mejora es incremental y compatible con lo que ya funciona.

## 3. Estructura de archivos

### Nucleo operativo existente

- `Código.gs`: web app (`doGet`), registro de datos (`guardarDatos`), descarte automatico 90 dias, trigger diario.
- `formulario.html`: formulario de captura para cotizaciones.
- `estilos.gs`: estilos y validaciones visuales/operativas.
- `sincronizacion.gs`: sincronizacion masiva de hojas de usuario hacia la hoja principal.

### Capa de seguridad y evolucion

- `Config.gs`: constantes centrales (nombres de hojas, roles, encabezados, permisos por rol).
- `Utils.gs`: utilidades comunes, normalizacion, logs, helpers de seguridad.
- `Setup.gs`: inicializacion del sistema, creacion de hojas auxiliares, validaciones estructurales, menu.
- `Permissions.gs`: obtencion de usuario, lectura de matriz de permisos, rol por correo, evaluacion de columnas editables.
- `Security.gs`: `onOpen`, `onEdit`, `onChange`, protecciones por rango, reversion de edicion no autorizada.
- `Validation.gs`: reglas de negocio sobre ediciones autorizadas y sincronizacion de filas.

### Base futura (sin romper operacion actual)

- `Clientes.gs`: validacion y upsert de clientes por RUT/NIT (base para formulario de cliente).
- `Cotizaciones.gs`: validacion de payload y puente a `guardarDatos` (base para formulario de cotizacion).

### Compatibilidad legacy

- `roles.gs`: wrapper de compatibilidad.
- `estetica.gs`: wrapper legacy sin duplicar funciones globales.

## 4. Hoja principal

Nombre esperado: **Datos**

Encabezados operativos actuales (A-R):

1. ID
2. Fecha de solicitud
3. Centro de costos
4. Cliente
5. Descripcion
6. Numero cotizacion
7. Link
8. Estado
9. Observaciones
10. PO/OC
11. HE/ACTA
12. Fecha de inicio
13. Fecha final
14. Estado del proceso
15. Fecha de factura
16. Numero factura
17. Fecha y hora de registro
18. Correo electronico

Importante:

- No cambiar orden ni nombre de encabezados sin revisar `Config.gs`.
- La capa de permisos detecta columnas por nombre normalizado del encabezado (no por indice fijo siempre que sea posible).

## 5. Flujo funcional actual

### 5.1 Captura web

1. `doGet()` publica `formulario.html`.
2. El formulario envia datos a `guardarDatos(form)`.
3. `guardarDatos`:
	 - valida payload.
	 - calcula ID consecutivo.
	 - arma fila completa A-R.
	 - inserta en hoja `Datos`.
	 - crea/actualiza hoja de usuario por correo si aplica.
	 - aplica validaciones y estilo.

### 5.2 Automatizacion de descarte

- `autoDescartar90Dias()` marca como `Descartada` registros con >= 90 dias, excepto aprobados/facturados.
- `installDailyAutoDiscard_()` crea trigger diario programado.

### 5.3 Seguridad en edicion

- `onEdit(e)` central en `Security.gs` llama `handleEditByRole_(e)`.
- Si una edicion no esta autorizada:
	- se revierte con `revertUnauthorizedEdit_`.
	- se registra en hoja `Logs`.
	- se notifica por toast.
- Si la edicion es autorizada:
	- se aplican reglas de negocio (`Validation.gs`).
	- se sincroniza la fila hacia hoja principal cuando corresponde.

## 6. Roles y permisos

Roles soportados:

- `admin`
- `usuario`
- `financiera`
- `auditor`

Definicion de edicion por rol (sobre encabezados):

- admin: todas las columnas.
- usuario:
	- Estado
	- Observaciones
	- PO/OC
	- HE/ACTA
	- Fecha de inicio
	- Fecha final
- financiera:
	- Estado del proceso
	- Fecha de factura
	- Numero factura
- auditor: sin edicion (solo lectura).

Columnas sensibles protegidas para no-admin:

- ID
- Fecha de solicitud
- Centro de costos
- Cliente
- Descripcion
- Numero cotizacion
- Link
- Fecha y hora de registro
- Correo electronico

## 7. Hojas auxiliares

Se crean con `initSystem()` cuando no existen:

- `Config_Usuarios`: matriz de permisos por correo.
- `Config_Sistema`: parametros base del sistema.
- `Logs`: trazabilidad de acciones y bloqueos.
- `Clientes`: base futura de clientes.
- `Catalogos`: catalogos base para listas.

### 7.1 Estructura de Config_Usuarios

Encabezados:

- correo
- nombre
- rol
- activo
- puedeEditar
- observaciones

Valores sugeridos:

- `activo`: si/true/1/x para habilitar.
- `puedeEditar`: si/true/1/x para habilitar edicion.

## 8. Menu del sistema

Se construye en `buildMenu_()` (onOpen):

- Menu general:
	- Estado del sistema.
- Menu admin (solo rol admin):
	- Inicializar sistema
	- Aplicar protecciones
	- Reaplicar protecciones
	- Sincronizar permisos
	- Validar estructura
	- Sincronizar hojas usuario -> principal
	- Aplicar validaciones/formatos
	- Ejecutar descarte 90 dias
	- Instalar job diario

## 9. Instalacion y puesta en marcha

Orden recomendado:

1. Abrir el proyecto Apps Script vinculado a la hoja operativa.
2. Ejecutar `initSystem()` una vez.
3. Llenar `Config_Usuarios` con correos reales y rol.
4. Ejecutar `applySheetProtections_()`.
5. Ejecutar `installableOnOpen_()` para triggers instalables (`onOpen` y `onChange`).
6. Ejecutar `installDailyAutoDiscard_()` si se desea descarte automatico diario.
7. Validar con cuentas de prueba por cada rol.

## 10. Trazabilidad y auditoria

La hoja `Logs` registra eventos como:

- intentos de edicion no autorizada.
- errores en flujo de validacion.
- acciones administrativas (sincronizacion de permisos, aplicacion de protecciones).
- eventos de cambio estructural detectados (`onChange`).

Campos registrados:

- fechaHora
- usuario
- accion
- hoja
- fila
- columna
- valorAnterior
- valorNuevo
- resultado
- detalle

## 11. Consideraciones de seguridad importantes

1. Apps Script ayuda a bloquear/revertir cambios, pero la seguridad total tambien depende de permisos de comparticion del archivo.
2. Para maxima seguridad operativa:
	 - limitar editores del archivo a personal autorizado.
	 - usar viewers para perfiles de solo consulta.
	 - mantener actualizado `Config_Usuarios`.
3. `onChange` permite traza de cambios estructurales, pero no evita todo evento destructivo en tiempo real en todos los escenarios.

## 12. Extensibilidad futura

El sistema queda listo para crecer a dos formularios web:

- Formulario de cliente:
	- validacion de RUT/NIT.
	- upsert de cliente evitando duplicados.
- Formulario de cotizacion:
	- validacion de payload comercial.
	- uso del flujo actual `guardarDatos` para no romper operacion.

## 13. Mantenimiento recomendado

- Revisar `Logs` semanalmente.
- Reaplicar `applySheetProtections_()` cuando haya cambios de estructura.
- Ejecutar `syncPermissions_()` despues de modificar `Config_Usuarios`.
- Mantener encabezados de la hoja `Datos` estables para evitar desalineaciones.

## 14. Estado actual del proyecto

- Arquitectura modular activa.
- `onOpen` y `onEdit` centralizados en una sola version.
- Sin duplicados globales de funciones.
- Sin errores de sintaxis detectados en la ultima validacion del proyecto.