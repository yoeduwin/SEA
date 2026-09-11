# Propuesta — PAIC autónomo con modelo cliente/sucursal persistente

**Estado:** Diseño corregido para implementación.
**Fecha:** 2026-09-10
**Alcance:** `PAIC.html` y ramas aditivas de `BACKEND_FIXES.gs`. **`SEAPD.html` no se toca. `SEAOT.html` no se toca.**

## 1. Regla de negocio correcta

PAIC y SEAPD no crean servicios ni órdenes de trabajo. Ambos sirven para dar de alta o actualizar un **cliente/sucursal** que puede reutilizarse indefinidamente mientras sus datos sigan vigentes.

La identidad operativa del registro maestro es:

`RFC + sucursal`

Ejemplo:

`AAC151218QKA + Planta Puebla`

Ese mismo registro puede utilizarse hoy, mañana o dentro de diez años. Si cambian datos del cliente o de la sucursal, PAIC/SEAPD actualizan el mismo registro y reutilizan la misma estructura padre en Drive.

La orden de trabajo es otra entidad y nace exclusivamente en SEAOT:

`cliente/sucursal existente + 1 servicio = 1 OT`

Por tanto:

- **1 registro PAIC/SEAPD = 1 cliente/sucursal.**
- **1 OT = 1 servicio para ese cliente/sucursal.**
- Un mismo cliente/sucursal puede acumular cualquier número de OT a lo largo del tiempo.
- PAIC no genera, reserva, predice ni administra folios de OT.
- PAIC no crea una entidad `SOLICITUDES` ni usa `sol_folio`, `sol_folios` u `OT_RESERVADA`.
- SEAOT conserva su responsabilidad vigente de consultar `getSiguienteFolioOT`, generar el folio y registrar la OT.
- SEAINF conserva su responsabilidad de crear el expediente después de que existe una OT.

## 2. Responsabilidad de cada módulo

### PAIC / SEAPD — maestro de cliente y sucursal

El registro crea o reutiliza:

```text
{RFC} - {EMPRESA}/
└── {Sucursal}/
    └── 01_Cliente/
```

Y actualiza `CLIENTES_MAESTRO` para la pareja RFC+sucursal.

La carpeta de empresa y la carpeta de sucursal son estructuras persistentes. No representan un servicio concreto.

### SEAOT — orden de trabajo

SEAOT busca un cliente/sucursal existente, obtiene el enlace de su carpeta padre, genera el siguiente folio mediante su mecanismo vigente y registra una OT para **un servicio**.

PAIC no participa en el consecutivo de OT.

### SEAINF — expediente

SEAINF trabaja a partir de una OT ya existente y crea el expediente correspondiente. No necesita una solicitud PAIC intermedia.

## 3. Qué cambia en PAIC

PAIC sigue siendo el portal específico para asesores, intermediarios y consultorías, pero deja de presentarse como formulario de solicitud de un estudio.

### 3.1 Copy y significado del formulario

Cambiar:

`1 registro = 1 estudio`

por:

`1 registro = 1 cliente / sucursal`

El texto de bienvenida y confirmación debe explicar que el formulario da de alta o actualiza los datos maestros de una ubicación del cliente y que las órdenes de trabajo se generan posteriormente por Ejecutiva Ambiental.

### 3.2 Quitar del registro maestro lo que pertenece a un servicio concreto

Eliminar de PAIC como parte del registro de cliente/sucursal:

- selector `estudio_laboratorio`;
- selector `estudio_higiene` y `otra_nom_higiene`;
- bloques de hojas de campo, fotografías y croquis por estudio;
- textos “1 registro = 1 estudio”;
- `fechas_preferidas` para evaluación;
- cualquier lógica que intente convertir esos campos en solicitud, folio o vínculo con una OT.

Los documentos y fechas propios de un servicio deben asociarse después al proceso de la OT/expediente correspondiente, no al maestro del cliente.

### 3.3 NOM-020 y PIPC

PAIC no debe convertir NOM-020 en una “solicitud de servicio”. En esta implementación se elimina del formulario PAIC el bloque específico de servicio NOM-020 y su documentación operativa (testigos, calibración, etc.).

Como esos campos dejan de llegar desde PAIC, el backend PAIC debe aplicar esta regla al actualizar una fila existente:

- si `aplica_nom020` está ausente, **conservar el valor vigente** de `CLIENTES_MAESTRO`;
- si `requiere_pipc` está ausente, **conservar el valor vigente** de `CLIENTES_MAESTRO`.

Para un cliente/sucursal nuevo registrado por PAIC, esos dos campos se guardan vacíos en lugar de afirmar `NO` sobre algo que PAIC no preguntó.

SEAPD conserva exactamente su comportamiento actual para ambos campos.

### 3.4 Documentación que sí permanece en PAIC

PAIC puede seguir recibiendo documentos generales de cliente/sucursal, por ejemplo:

- layout o plano general;
- requisitos generales de ingreso a planta;
- documentación general ya existente que no dependa de una OT concreta.

Esos documentos permanecen en `01_Cliente` porque describen al cliente/sucursal y pueden reutilizarse en futuros servicios.

No se crean subcarpetas por servicio desde PAIC.

## 4. Correo del intermediario y correo del cliente

PAIC necesita distinguir dos destinos:

- `correo_informe`: correo del **cliente**, que permanece en `CLIENTES_MAESTRO` y sigue siendo el destino de informes/PORTAL;
- `correo_acuse`: correo del **asesor/intermediario** que está llenando PAIC y recibe la confirmación del registro.

Agregar en PAIC un campo obligatorio:

**Correo para confirmación y seguimiento del registro** (`correo_acuse`).

Regla dura de backend:

```javascript
if (data.portal_origen === 'PAIC') {
  if (esCorreoValido_(data.correo_acuse)) {
    enviarConfirmacionPAIC_(data, carpetaCliente, files);
  }
  // Si falta o es inválido, se omite el acuse.
  // Nunca se usa correo_informe como fallback.
} else {
  enviarConfirmacionCliente(data, carpetaCliente, files);
}
```

El camino de SEAPD queda intacto.

El correo interno debe indicar claramente:

- origen: PAIC;
- asesor/consultor;
- correo de acuse del asesor;
- correo del cliente para informe.

En PAIC se omiten referencias a “servicios solicitados”, PIPC y fechas preferidas, porque el evento recibido es **alta/actualización de cliente/sucursal**, no una OT.

## 5. Actualización de CLIENTES_MAESTRO

La fila sigue teniendo el contrato vigente de 22 columnas. No se agregan columnas.

Para `portal_origen === 'PAIC'`:

- RFC+sucursal siguen identificando la fila a actualizar;
- `REQUIERE_PIPC` se conserva si PAIC no envía el campo;
- `APLICA_NOM020` se conserva si PAIC no envía el campo;
- `ASESOR_CONSULTOR` se actualiza cuando PAIC envía un valor y se conserva cuando llega vacío;
- `correo_informe` sigue siendo el correo del cliente;
- `correo_acuse` no se escribe en `CLIENTES_MAESTRO`.

No se crea una hoja `SOLICITUDES` porque no existe una entidad de negocio “solicitud PAIC” entre cliente y OT.

## 6. Reintentos

Un reintento de PAIC debe entenderse únicamente como un nuevo intento de **actualizar el mismo cliente/sucursal**, no como una solicitud de servicio.

No se introduce un folio de negocio para el formulario.

Se acepta el comportamiento normal de upsert por RFC+sucursal. Si una respuesta se pierde y el formulario se reenvía, el backend vuelve a asegurar los datos del mismo cliente/sucursal.

No se diseña en este PR una transacción distribuida para Drive/Sheets/Gmail. Los efectos externos se mantienen con el comportamiento existente salvo las protecciones explícitas del correo PAIC.

## 7. Cambios técnicos previstos

### `PAIC.html`

1. Cambiar copy de “1 registro = 1 estudio” a “1 registro = 1 cliente/sucursal”.
2. Agregar `correo_acuse` obligatorio junto a los datos del asesor.
3. Aclarar que `correo_informe` pertenece al cliente.
4. Eliminar Laboratorio, Higiene y sus archivos por estudio.
5. Eliminar la sección de coordinación/fechas preferidas.
6. Eliminar el bloque de servicio NOM-020 y su documentación específica.
7. Simplificar el modal de confirmación a datos del cliente/sucursal + archivos generales.
8. Limpiar JavaScript asociado a selectores y bloques eliminados.
9. Mantener `portal_origen: 'PAIC'` en el payload.

### `BACKEND_FIXES.gs`

1. Mantener el camino vigente de SEAPD sin cambios funcionales.
2. En la rama PAIC, preservar `APLICA_NOM020` y `REQUIERE_PIPC` cuando el payload no los contiene; en alta nueva dejarlos vacíos.
3. Enviar el acuse PAIC exclusivamente a `correo_acuse`; nunca hacer fallback a `correo_informe`.
4. Ajustar correo interno para que PAIC se describa como alta/actualización de cliente/sucursal, no como servicio solicitado.
5. No agregar `SOLICITUDES`, reserva de OT ni cambios en SEAOT/SEAINF.

## 8. Gate de regresión

Antes de publicar cualquier cambio de `BACKEND_FIXES.gs`, verificar con un payload vigente de SEAPD que:

- crea/reutiliza la misma estructura de Drive;
- mantiene el contrato de 22 columnas;
- conserva su manejo actual de NOM-020 y PIPC;
- conserva el destinatario y contenido funcional de su acuse;
- no requiere `portal_origen` ni `correo_acuse`.

Si cambia el comportamiento de SEAPD, el backend no se publica.

## 9. Fuera de alcance

Este trabajo **no** debe:

- tocar `SEAPD.html`;
- modificar `SEAOT.html`;
- modificar el generador `getSiguienteFolioOT`;
- reservar folios de OT desde PAIC;
- crear `SOLICITUDES`;
- crear `sol_folio`, `sol_folios` u `OT_RESERVADA`;
- agrupar servicios u órdenes;
- crear carpetas de Drive por servicio desde PAIC;
- modificar SEAINF para buscar archivos de una supuesta solicitud PAIC.

## 10. Criterio de aceptación

El cambio está correcto cuando se puede demostrar este flujo:

```text
PAIC
  ↓
Alta/actualización de RFC + sucursal
  ↓
CLIENTES_MAESTRO + carpeta padre persistente
  ↓
(otro momento, incluso años después)
  ↓
SEAOT busca ese cliente/sucursal
  ↓
SEAOT genera 1 OT para 1 servicio
  ↓
SEAINF crea el expediente de esa OT
```

Y cuando un registro PAIC nunca manda su acuse al correo del cliente.