# Panel de consulta SEATRAZ

Interfaz adicional en `TRAZ/panel.html`. El acceso original `TRAZ/TRAZ.html`
continúa disponible y no se modifica.

## Conexión y permisos

- Usa el mismo `auth.js`, Client ID de Google y endpoint que SEATRAZ actual.
- Requiere usuario activo con `SEATRAZ = TRUE` en `USUARIOS_AUTORIZADOS`.
- El backend continúa comprobando permisos en cada consulta.
- Solo solicita `trazResumen` y `trazDetalle`; la comprobación de acceso del
  módulo compartido usa `verificarAcceso`. Aunque el transporte sea POST,
  estas acciones no escriben datos.
- No agrega base de datos, hojas, columnas ni automatizaciones. Solo guarda
  preferencias de consulta en el navegador, separadas por cuenta Google.
- Publicar bajo el mismo origen Google autorizado que el frontend actual.
  No necesita un nuevo despliegue de Apps Script.

## Experiencia de consulta

Una sola tabla muestra orden, cliente/sucursal, servicio, estado TRAZ, todos los
folios de informe relacionados y acceso al expediente/detalle. Una fila por OT:
varios informes no duplican la orden. Sin resumen general ni vistas separadas
de informes o alertas.

La búsqueda encuentra OT, clientes, sucursales, personal y folios de informe.
Los filtros principales son serie (Todas, OT, OTB, OTC), estado y servicio.
Personal está en Más filtros y se señala cuando está activo. La serie se toma
del tipo explícito OT/OTB/OTC; si falta, se reconoce un prefijo de folio válido.
No se asume OT cuando la serie es desconocida.

El detalle abre un diálogo lateral con ruta, datos, documentos y bitácora real.
La tabla mantiene búsqueda, filtros y posición de desplazamiento; cerrar o
pulsar Escape devuelve el foco al botón de origen. El foco de teclado queda
dentro del diálogo mientras está abierto.

Consultar carpeta obtiene solo el detalle de esa orden y enfoca sus documentos,
donde el enlace existente abre Drive. El resumen no expone URLs de carpetas:
no se inventan enlaces ni se consultan masivamente las órdenes para obtenerlos.

Los datos faltantes aparecen junto al dato afectado. Una OTC sin informes muestra
No aplica; el detalle omite únicamente la advertencia genérica de informe ausente
y marca ese hito como no aplicable. Si una OTC sí tiene informes, se muestran.
Los informes con OT inexistente solo aparecen cuando hay registros, en un
bloque desplegable independiente de los filtros.

## Retomar consultas y encontrar documentos

Después de una consulta autenticada correcta, recupera búsqueda y filtros de esa
cuenta desde localStorage. Guarda seis folios consultados con éxito y muestra
sus nombres actuales usando exclusivamente el resumen autorizado de esa sesión.
No copia datos del expediente, bitácora ni tokens. Si el navegador bloquea el
almacenamiento o las preferencias están dañadas, la consulta sigue funcionando.
Limpiar filtros y Limpiar recientes son acciones locales. Las preferencias no
se sincronizan entre equipos ni cuentas.

Los folios de informe son accesos a la ficha correspondiente dentro de la OT;
Ver documentos enfoca los enlaces de carpeta. El backend entrega carpetas,
no URLs directas de los PDF. El detalle presenta primero documentos y último
cambio con fecha de AUDITORIA, seguido de la ruta y la historia desplegable.
Ver cambios enfoca ese resumen; Ver bitácora completa abre la tabla original.

El último cambio se calcula únicamente entre eventos auditados con fecha
reconocible, sin usar fechas de programación como evidencia de actividad.
Se avisa si hay eventos sin fecha reconocible. No se consultan todas las
bitácoras para inventar una columna global de últimos movimientos.

## Interpretación

Estado TRAZ y estatus SEA se conservan separados. Los enlaces no comprueban el
contenido de Drive. Los hitos de ejecución dependen de fechas registradas;
revisión y emisión solo aparecen si fueron auditadas.

Actualizar es manual; ante error se conserva la última consulta y se muestra su
antigüedad. Solo se consulta detalle al abrir una orden; no hay carga masiva.

## Verificación

```sh
node --check TRAZ/panel.js
node --check TRAZ/panel-detalle.js
node --test TRAZ/tests/panel-consulta.test.cjs
```

La prueba usa datos sintéticos sin conexión a SEA. La validación con datos reales
requiere iniciar sesión con una cuenta autorizada en el origen de producción.

## Reversión

Los únicos archivos añadidos son `panel.html`, `panel.css`, `panel.js`,
`panel-detalle.js`, esta documentación y la prueba. Retirarlos elimina el panel
sin afectar el frontend original, backend, autenticación o información de SEA.
