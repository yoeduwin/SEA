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
- No agrega base de datos, hojas, columnas, automatizaciones ni almacenamiento
  de datos operativos en el navegador.
- Publicar bajo el mismo origen Google autorizado que el frontend actual.
  No necesita un nuevo despliegue de Apps Script.

## Vistas

1. **Resumen:** indicadores y distribución por estado dentro de los filtros.
2. **Órdenes:** búsqueda por OT, cliente, sucursal, informe o personal, con
   filtros por estado TRAZ, servicio y personal de OT.
3. **Informes:** índice de los folios relacionados expuestos por el resumen.
   Estados y expedientes individuales se consultan en el detalle de la OT.
4. **Alertas:** ausencia de carpeta de OT o de informe relacionado. Los informes
   con OT inexistente se muestran como consulta global independiente del filtro.
5. **Detalle:** ruta, historia, datos, informes y bitácora real, con enlaces a
   las carpetas existentes. Conserva separados estado TRAZ y estatus SEA.

## Interpretación

Los indicadores usan el estado y las relaciones que devuelve el backend actual.
Las alertas de resumen no comprueban el contenido de Drive ni equivalen a una
evaluación de cumplimiento. Una OT entregada también puede presentar datos
faltantes. El detalle expone las advertencias completas del backend al abrirla.
Los hitos de ejecución dependen de fechas registradas, no de una verificación
independiente de campo. Revisión y emisión solo aparecen si fueron auditadas.

El panel realiza una consulta de resumen y consulta el detalle únicamente al
abrir una OT. Actualizar es una acción manual; ante error conserva la última
consulta y muestra su antigüedad. No carga masivamente detalles ni bitácoras.

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
