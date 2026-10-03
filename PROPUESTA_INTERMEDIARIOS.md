# Propuesta — Identidad del intermediario en PAIC (código `A4-R577`)

**Estado:** Diseño acordado, pendiente de implementación.
**Fecha:** 2026-10-03
**Alcance:** `PAIC.html` y ramas aditivas de `BACKEND_FIXES.gs`. **`SEAOT.html`, la orden de trabajo y `SEAPD.html` no se tocan. `SEADB.html` no se modifica.**

## 1. Problema

El intermediario que registra clientes en PAIC no tiene identidad: el campo "Asesor / Consultor (empresa)" es texto libre, opcional, y se guarda tal cual en la columna 22 de `CLIENTES_MAESTRO`.

Con sólo 8 registros ya existen todos estos casos:

- la misma consultoría escrita de dos formas (con y sin acentos y comas);
- espacios invisibles al final del nombre;
- nombre de persona en lugar de despacho;
- una palabra genérica (`CONSULTOR`) en lugar del nombre;
- la misma intermediaria registrada en dos columnas distintas con dos nombres distintos.

Consecuencias:

- El filtro de SEADB (`toLowerCase().includes()`) no ignora acentos ni puntuación: buscar `protección` y `proteccion` devuelve registros distintos de la misma consultoría.
- No se puede responder con confianza "¿cuántos clientes trae cada intermediario?".
- No se guarda el portal de origen, así que un registro PAIC sin intermediario no se distingue de uno de SEAPD.
- Si el cliente actualiza sus datos en SEAPD, la columna 22 se sobrescribe vacía y se pierde quién era su intermediario.

Normalizar el texto no lo resuelve: ninguna regla sobre el texto puede saber que `ADRIANA FLORES` y `CONSULTOR` son la misma intermediaria. El intermediario necesita una identidad asignada por Ejecutiva Ambiental, igual que el cliente la tiene en su RFC.

## 2. Regla de negocio

- **Intermediario = despacho = asesor = consultor.** Es una sola entidad.
- **Ejecutiva Ambiental asigna el código.** No hay autorregistro: así nadie acumula varios códigos por la prisa de avanzar.
- **Sin código válido, PAIC no abre el formulario.**
- Con código válido, los datos del intermediario se **muestran precargados y no editables**. El intermediario sólo captura los datos del cliente/sucursal y sube documentos.
- El intermediario no avisa cuando termina; por eso **cada registro debe llegar identificado por sí mismo**.
- Un cliente/sucursal puede cambiar de intermediario. **Se guarda el historial.**

## 3. Formato del código

```text
A4-R577
└┬┘ └┬─┘
 │   └─ Clave: 1 letra + 3 números (aleatoria)
 └───── Número de intermediario: "A" + consecutivo, fijo para siempre
```

| Parte | Función | Reglas |
|---|---|---|
| `A4` | **Identidad.** Va en `CLIENTES_MAESTRO`, en el historial y en reportes. | Consecutivo `A1`…`A100`, sin ceros a la izquierda. Nunca cambia ni se reutiliza. |
| `R577` | **Llave.** Sólo sirve para abrir PAIC. | Letra sin `I`, `L`, `O` + 3 dígitos. 23 000 combinaciones por número. Se puede cambiar sin cambiar el número. |

Fundamento (prácticas de ERP y portales de proveedores):

- Los números de cliente/proveedor recomendados son **sin significado y estables**: el historial cuelga de ellos y cambiarlos rompe referencias.
- Los portales de proveedores (Oracle, Workday) funcionan por **invitación de la empresa**, no por autorregistro libre, y revisan duplicados antes de dar de alta.
- Un consecutivo es adivinable; si por sí solo abriera el portal, cualquiera podría probar números. Por eso el número (identidad) y la clave (llave) se separan, como número de cliente y NIP en un banco.

### 3.1 Captura tolerante

El sistema normaliza lo que se teclea antes de validar:

1. Mayúsculas; se eliminan espacios, guiones y guiones bajos.
2. Después de la `A` inicial: `O` → `0`; `I` y `L` → `1`.
3. Debe cumplir `^A(\d{1,3})([A-HJKMNP-Z])(\d{3})$` y el número estar entre 1 y 100.

Así `a4r577`, `A4 R577`, `A4-R577` y `A04-R577` son el mismo código.

### 3.2 Mensaje de error único

Cualquier fallo (formato, número inexistente, clave incorrecta, código `DISPONIBLE` o `BAJA`) responde lo mismo:

> Código no válido. Solicítelo a Atención a Clientes: aclientes@ejecutivambiental.com · WhatsApp 56 5282 1561.

No se indica si falló el número o la clave, para no dar pistas.

### 3.3 Freno contra intentos repetidos

- Validar el código exige token reCAPTCHA, igual que `registrarCliente`.
- **Bloqueo por número:** 5 claves incorrectas para el mismo número en 1 hora bloquean ese número 1 hora (`CacheService`). Los demás intermediarios no se ven afectados.
- Al activarse un bloqueo se envía un aviso interno a `CONFIG.EMAIL_TO`.

Riesgo residual aceptado: con este freno, adivinar la clave de un número concreto toma del orden de meses de intentos sostenidos, cada uno con reCAPTCHA, y cada registro resultante llega notificado a Ejecutiva Ambiental. Si una clave se filtra o se sospecha abuso, se cambia la clave (sección 4.2).

## 4. Hoja `INTERMEDIARIOS` (catálogo)

Hoja nueva en el mismo spreadsheet. Nace con **100 filas pregeneradas** (`A1`…`A100`), cada una con su clave aleatoria y estatus `DISPONIBLE`.

| Col. | Campo | Notas |
|---|---|---|
| A | Número | `A1`…`A100`. Fijo. |
| B | Clave | `R577`. Generada al crear la hoja. |
| C | Código | `A1-R577`, para copiar y enviar. |
| D | Estatus | `DISPONIBLE` · `ACTIVO` · `BAJA` |
| E | Nombre del intermediario | Nombre oficial; es el que se escribe en `CLIENTES_MAESTRO`. |
| F | Contacto (nombre y puesto) | |
| G | Correo | Destino del acuse. Obligatorio para `ACTIVO`. |
| H | Teléfono | |
| I | RFC | Opcional. Sirve para revisar duplicados. |
| J | Fecha de alta | |
| K | Notas | |

Sólo un código `ACTIVO` con nombre y correo abre PAIC.

### 4.1 Alta de un intermediario (manual, sin pantalla)

1. Buscar en la hoja que el intermediario no exista ya (por nombre o RFC).
2. Tomar la siguiente fila `DISPONIBLE`.
3. Llenar nombre, contacto, correo, teléfono y RFC si se tiene.
4. Cambiar estatus a `ACTIVO`, poner la fecha de alta y enviarle su código.

La pantalla de administración queda fuera de esta etapa.

### 4.2 Cambio de clave y baja

- **Clave filtrada:** se reemplaza sólo la clave de esa fila. El número y todo su historial se conservan. Se le envía el nuevo código.
- **Baja:** estatus `BAJA`. El código deja de abrir PAIC. El número nunca se reasigna.

### 4.3 Generación de las claves

- Función `crearHojaIntermediarios()` en `BACKEND_FIXES.gs`, ejecutada **una sola vez desde el editor** de Apps Script (mismo patrón que `crearHojaUsuarios()`).
- Aleatoriedad derivada de `Utilities.getUuid()`. Las claves no se repiten dentro de la hoja.
- Si la hoja ya existe, la función no hace nada (no sobrescribe claves vigentes).
- **Los códigos nunca se guardan en el repositorio:** el repositorio se publica en GitHub Pages. Viven sólo en la hoja.

### 4.4 Asignación inicial

Por fecha del primer registro existente:

| Número | Intermediario |
|---|---|
| A1 | FRATUM |
| A2 | Adriana Flores Flores |
| A3 | Carolina Castillo Abriz |
| A4 | Protección Ambiental, Consultoría, Ingeniería y Análisis |
| A5 | CRECYDE |
| A6 | Servicios de Apoyo Industrial |

Ejecutiva Ambiental completa contacto, correo y teléfono de cada uno antes de enviarles su código.

## 5. Cambios en `CLIENTES_MAESTRO`

Se agrega **una columna al final** (no se mueve ninguna existente):

| Col. | Campo |
|---|---|
| 23 (índice 22) | ID Intermediario (`A4`) |

No se agrega columna de portal de origen: con código obligatorio, toda fila con ID de intermediario proviene de PAIC, y cada envío queda además en el historial (sección 6).

Reglas al registrar desde **PAIC**:

| Columna | Valor |
|---|---|
| 8 · Contacto (Solicitante) | Contacto del catálogo. Se mantiene el comportamiento actual porque SEAOT lee esta columna como solicitante de la OT. |
| 22 · Asesor / Intermediario | Nombre oficial del catálogo. |
| 23 · ID Intermediario | Número del catálogo (`A4`). |

Reglas al registrar desde **SEAPD** (corrección necesaria):

- Al **actualizar** una fila existente, la columna 22 **conserva** su valor vigente en lugar de sobrescribirse vacía. Es la misma protección que PAIC ya aplica a `REQUIERE_PIPC`.
- La columna 23 queda intacta de forma natural, porque SEAPD sólo escribe las 22 columnas del contrato.

El preflight `getMaxColumns() < 22` se conserva; la columna 23 es aditiva.

### 5.1 Migración de las filas existentes

Función `migrarIntermediariosExistentes()` ejecutada una vez desde el editor. Recorre `CLIENTES_MAESTRO` y, según el valor actual de la columna 22 (comparado sin espacios extremos ni mayúsculas/acentos), escribe nombre oficial y número:

| Valor actual en columna 22 | Número |
|---|---|
| `FRATUM` | A1 |
| `ADRIANA FLORES`, `CONSULTOR` | A2 |
| `CAROLINA CASTILLO ABRIZ` | A3 |
| `PROTECCIÓN AMBIENTAL, CONSULTORIA, INGENIERÍA Y ANÁLISIS`, `PROTECCION AMBIENTAL CONSULTORIA INGENIERIA Y ANALISIS` | A4 |
| `CRECYDE` | A5 |
| `Servicios de Apoyo Industrial` | A6 |

La función registra en el log cada fila modificada y no toca filas cuyo valor no esté en la tabla. Afecta 8 filas.

## 6. Hoja `HISTORIAL_INTERMEDIARIOS`

Hoja nueva, **sólo se agregan filas, nunca se modifican**. Una fila por cada envío exitoso de PAIC:

| Fecha | ID Intermediario | Intermediario (nombre al momento) | RFC | Razón social | Sucursal | Tipo (`ALTA` / `ACTUALIZACIÓN`) | Link Drive |
|---|---|---|---|---|---|---|---|

La migración (5.1) agrega también una fila por cada uno de los 8 registros existentes, con su fecha original y tipo `MIGRACIÓN`.

### 6.1 Limitación aceptada

SEADB obtiene el intermediario de cada OT consultando `CLIENTES_MAESTRO` en el momento (`fase4_GetTablero`). Si un cliente cambia de intermediario, el tablero mostrará el nuevo también en sus OT anteriores. **La fuente de verdad de quién registró qué y cuándo es `HISTORIAL_INTERMEDIARIOS`.** No se modifica la OT para evitarlo.

## 7. Cambios técnicos previstos

### `BACKEND_FIXES.gs`

1. `CONFIG`: agregar `SHEET_INTERMEDIARIOS`, `SHEET_HIST_INTERMEDIARIOS` y `CLIENTES.ID_INTERMEDIARIO: 22`.
2. Nueva acción pública `validarCodigoIntermediario` con `AUTH_MODE: 'RECAPTCHA'`:
   - normaliza (3.1), aplica el freno (3.3) y valida contra el catálogo;
   - si es válido devuelve sólo `{ valid, numero, nombre, contacto, correo, telefono }`;
   - si no, devuelve `{ valid:false }` con el mensaje único (3.2).
3. En `registrarCliente`, rama PAIC:
   - **revalidar el código en el servidor** y tomar nombre, contacto y correo **del catálogo**, ignorando lo que venga en el payload (los campos son de sólo lectura en pantalla, pero el payload se puede alterar);
   - sin código válido: rechazar **antes** de crear carpetas o archivos;
   - escribir columnas 8, 22 y 23 según la sección 5;
   - agregar la fila a `HISTORIAL_INTERMEDIARIOS`;
   - enviar el acuse al correo del catálogo; el correo interno incluye el número de intermediario.
4. En `registrarCliente`, rama SEAPD: conservar la columna 22 al actualizar. Ningún otro cambio.
5. Funciones de ejecución única: `crearHojaIntermediarios()` y `migrarIntermediariosExistentes()`.

### `PAIC.html`

1. Pantalla inicial con un solo campo "Código de intermediario" y botón "Continuar". El formulario permanece oculto.
2. Con código válido se muestra el formulario con la sección del intermediario **precargada y en sólo lectura**: número, nombre, contacto, correo y teléfono.
3. Se eliminan los campos editables `nombre_solicitante`, `asesor_consultor` y `correo_acuse`. El payload envía `codigo_intermediario`.
4. La precarga por RFC del cliente (`applyClientData`) deja de escribir `nombre_solicitante`, para no pisar los datos del intermediario.
5. "Limpiar datos" (`clearPrefill`) limpia sólo los datos del cliente y conserva la sesión del intermediario.
6. Opción "Cambiar código" para volver a la pantalla inicial.
7. El código validado se guarda en `sessionStorage` para no pedirlo de nuevo si la página se recarga en la misma pestaña; nunca en `localStorage`.

## 8. Pruebas (`TESTS_BACKEND.gs`)

- Normalización: `a4r577`, `A4 R577`, `A04-R577` y `AO4-R577` equivalen a `A4-R577`; `Al-R577` equivale a `A1-R577`; `A4-RS77` y `A101-R577` son inválidos.
- Validación: `ACTIVO` válido; clave incorrecta; `DISPONIBLE`; `BAJA`; `ACTIVO` sin correo; número fuera de rango. Todas las fallas devuelven el mismo mensaje.
- Freno: la sexta clave incorrecta en una hora bloquea ese número y no otro.
- PAIC con payload alterado (otro nombre/correo de intermediario): se escriben los datos del catálogo.
- PAIC sin código o con código inválido: no se crea carpeta ni fila.
- PAIC alta y actualización: columnas 8, 22 y 23 correctas; una fila nueva en el historial.
- SEAPD actualizando una fila con intermediario: columnas 22 y 23 se conservan.
- Migración: las 8 filas quedan con nombre oficial y número; filas sin intermediario no cambian.

## 9. Gate de regresión

Antes de publicar `BACKEND_FIXES.gs`, verificar con un payload vigente de SEAPD que:

- crea/reutiliza la misma estructura de Drive;
- mantiene el contrato de 22 columnas;
- conserva NOM-020, PIPC y su acuse al cliente;
- no requiere `codigo_intermediario`.

Si cambia el comportamiento de SEAPD, el backend no se publica.

## 10. Fuera de alcance

- Tocar `SEAOT.html`, la orden de trabajo o `ORDENES_TRABAJO`.
- Tocar `SEAPD.html`.
- Modificar `SEADB.html` (el filtro funciona porque la columna 22 queda con nombres oficiales).
- Pantalla de administración de intermediarios.
- Autorregistro de intermediarios.
- Guardar códigos en el repositorio.

## 11. Orden de puesta en marcha

En Apps Script, ejecutar una función desde el editor usa el código guardado, mientras que la aplicación web sigue sirviendo la versión desplegada. Eso permite preparar los datos sin cambiar lo que ven los usuarios:

1. Guardar el nuevo `BACKEND_FIXES.gs` en el editor **sin crear una nueva implementación**.
2. Ejecutar `crearHojaIntermediarios()`.
3. Completar en la hoja los datos de A1–A6 y marcarlos `ACTIVO`.
4. Ejecutar `migrarIntermediariosExistentes()` y revisar el log.
5. Pasar el gate de regresión (sección 9).
6. Crear la nueva implementación del backend y publicar el nuevo `PAIC.html` al mismo tiempo.
7. Enviar a cada intermediario su código.

## 12. Criterio de aceptación

```text
Intermediario escribe A4-R577
  ↓
PAIC muestra sus datos (sólo lectura) y abre el formulario
  ↓
Captura cliente/sucursal y documentos
  ↓
CLIENTES_MAESTRO: col. 22 = nombre oficial, col. 23 = A4
HISTORIAL_INTERMEDIARIOS: +1 fila
Acuse → correo del catálogo
  ↓
SEADB: filtrar por el nombre oficial devuelve todos sus clientes
```

Y cuando ningún texto tecleado por el intermediario llega a las columnas 8, 22 o 23.
