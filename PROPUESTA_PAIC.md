# Propuesta — Alineación de PAIC con SEAPD

**Estado:** Análisis y plan para revisión (NO implementado).
**Flujo previsto:** Eduwin revisa → decide las 4 preguntas de §4 → se abre PR por fases.
**Fecha:** 2026-09-10
**Alcance:** `SEAPD.html`, `PAIC.html`, `BACKEND_FIXES.gs` (ruta de registro), `TESTS_BACKEND.gs`, `MANUAL.md`.

> **Resumen en una línea:** PAIC y SEAPD son **el mismo formulario bifurcado**: CSS idéntico, JavaScript
> idéntico salvo una función, y el **mismo endpoint de backend**. Pero la bifurcación nunca se cerró:
> PAIC **dejó fuera** partes de SEAPD que sí necesita, **agregó** partes que el backend nunca aprendió a
> recibir, y **heredó** los bugs compartidos sin heredar las correcciones.

---

## 0. El punto de partida: dos ramas del mismo árbol

| | SEAPD | PAIC |
|---|---|---|
| Archivo | `SEAPD.html` · 1,628 líneas | `PAIC.html` · 1,351 líneas |
| Acceso | Público · reCAPTCHA Enterprise | Público · reCAPTCHA Enterprise |
| Acción del backend | `registrarCliente` | `registrarCliente` — **la misma** |
| Etiqueta de versión | `v3.0 (Multi-Sucursal)` | `PAIC v1.0` |
| Clases CSS propias | — | 2 (`.paic-badge`, `.study-reminder`) |
| Funciones JS propias | 1 (`resetPIPCConditionals`) | **0** |
| Quién llena el formulario | El cliente | El asesor / intermediario |

Las funciones de JavaScript de PAIC son **las mismas 18**, con el mismo cuerpo casi carácter por carácter
(`clearFileUI`, `showOnRecordIndicators`, `showFileInput`, `clearPrefill`, `searchClientByRFC`,
`applyClientData`, `confirmAndSend`, `fileToBase64`…). Lo que cambió es el marcado del formulario.

### Secciones del formulario

| SEAPD | PAIC |
|---|---|
| Información General | Información General **+ Asesor / Consultor** |
| Sobre el Informe de Resultados | Sobre el Informe de Resultados |
| Coordinación de Evaluación | Coordinación de Evaluación |
| Descripción del Proceso | Descripción del Proceso |
| Archivos Adjuntos (4 documentos) | Archivos Adjuntos (**2** documentos, numerados 1 y 4) |
| NOM-020-STPS | INSPECCIÓN — NOM-020-STPS |
| **Programa Interno de Protección Civil (PIPC)** | — |
| — | **LABORATORIO** (estudio + hojas de campo, fotos, croquis) |
| — | **HIGIENE** (NOM STPS + hojas de campo, fotos, croquis) |

### Qué se le modificó a SEAPD desde la bifurcación

| Commit | Cambio | ¿Tocó PAIC? |
|---|---|---|
| `3e4ac3b` | **Flujo NOM-020 completo**: campo `jefe_mantenimiento`, testigos y jefe en la hoja de perfil y en el correo, renombrado de INE con el nombre del testigo | No |
| `2d6262a` + 4 fixes | **Rediseño completo de los correos de registro** (+474 líneas): correo interno y acuse al cliente, chips de servicios, botón de WhatsApp, normalización de teléfonos mexicanos | Backend compartido — llega a PAIC **a medias** (§2, grupo B) |
| `4789d1c` | reCAPTCHA Classic → Enterprise | **Sí** — única modificación que PAIC recibió |
| `637fbeb` | Site Key real en SEAPD | No aplicaba |

---

## 1. Cómo opera SEAPD (revisión)

### 1.1 El recorrido del usuario

```
  Modal de bienvenida (cookie 30 días)
        │
        ▼
  ¿Cliente recurrente?  ──►  buscarClienteRFC  ──►  elige sucursal existente
        │  no                                   └─►  o "Registrar Nueva Sucursal"
        │                                              │
        │                                              ▼
        │                                    prellenado + banner +
        │                                    "Ya tenemos este documento"
        ▼
  Formulario (7 secciones, condicionales anidados)
        │
        ▼
  Modal de confirmación  ──►  confirmAndSend  ──►  POST único con archivos en base64
                                                          │
                                                          ▼
                                                 fase1_RegistrarCliente
```

### 1.2 Los condicionales — donde vive casi toda la lógica

**NOM-020-STPS**: un radio `aplica_nom020` abre la sección legal y marca `required` a todo lo que lleve
`data-conditional-required="nom020"`: INE de quien atiende, dos testigos con nombre + INE, poder
notarial, INE del representante, constancia fiscal, licencia, DC-3, dictamen de calibración y —desde
`3e4ac3b`— el **jefe de mantenimiento**. Un checkbox `sin_calibracion` libera el dictamen y cambia el
texto de ayuda. Apagar el radio limpia archivos y valores.

**PIPC**: un radio `requiere_pipc` abre **12 documentos obligatorios** y cuatro sub-compuertas, cada una
con su propio radio y su propio archivo condicional:

| Sub-compuerta | Documento que habilita |
|---|---|
| `pipc_tiene_medidas` | Medidas preventivas |
| `pipc_tiene_gas` | Gas natural |
| `pipc_tiene_quimicos` | Sustancias químicas |
| `pipc_tiene_montacargas` | DC-3 de operadores |

`resetPIPCConditionals()` —la única función que PAIC no tiene— reinicia las cuatro.

### 1.3 El envío

`confirmAndSend()` recorre el `FormData`, convierte cada archivo a base64 (`data:...`), acumula el
tamaño total, corta en **25 MB**, y manda un solo `POST` con `AbortController` a 120 segundos. El límite
por archivo en el navegador es 10 MB; el backend valida hasta 50 MB. Los errores están tipificados:
`ARCHIVOS_MUY_GRANDES`, `HTTP_ERROR_nnn`, `RESPUESTA_INVALIDA`, `AbortError`, fallo de red.

### 1.4 Qué hace el backend con eso — `fase1_RegistrarCliente`

```
1.  Valida razón social y formato de RFC
2.  Preflight: CLIENTES_MAESTRO debe tener ≥ 22 columnas, si no aborta sin tocar Drive
3.  Carpeta padre  {RFC} - {EMPRESA}          ← el RFC es la identidad estable
4.  Carpeta hija   /{Sucursal}
5.  Perfil         /{Sucursal}/01_Cliente
6.  guardarArchivos()    → whitelist fija de 29 campos, renombra por etiqueta
7.  generarPerfilSheet() → hoja "PERFIL DE DATOS" con formato fijo por celda
8.  Upsert de la fila de 22 columnas, buscando hacia atrás por RFC + Sucursal
9.  enviarNotificacionRobusta() → correo al equipo + acuse al cliente
    3 reintentos, y fallback a correo simple si los dos fallan
10. Log de la ejecución en Drive
```

### 1.5 Los dos correos (el rediseño de `2d6262a`)

**Al equipo** (`enviarNotificacionEquipo`): botón de WhatsApp con el teléfono utilizable, mensaje de
seguimiento listo para copiar, **chips de servicios**, contacto, empresa, bloque NOM-020 si aplica,
fechas preferidas y lista de documentos con enlace directo.

**Al cliente** (`enviarConfirmacionCliente`): acuse a `correo_informe` con los datos registrados, chips
de servicios, los tres pasos siguientes y la línea de Atención a Clientes.

Los chips de servicios son **exactamente dos**: `NOM-020-STPS` y `PIPC`. Recuérdalo — es la raíz del
grupo B.

---

## 2. Qué le falta a PAIC

Cuatro grupos. No es una sola lista porque no son el mismo tipo de problema: unos son funcionalidad que
PAIC **dejó fuera**, otros son funcionalidad que PAIC **agregó y nadie recibe**, otro es un **choque de
modelo de negocio**, y el último es **deuda compartida** que PAIC heredó.

---

### GRUPO A — Lo que PAIC dejó fuera de SEAPD

#### 🔴 A1 · La sección PIPC completa, y con ella el campo `requiere_pipc`

PAIC no tiene ni la compuerta ni los 12 documentos ni las 4 sub-compuertas. Consecuencias:

1. **Un asesor no puede registrar un servicio de Protección Civil.** Es una línea de negocio entera
   cerrada para el canal.
2. El backend evalúa `data.requiere_pipc === 'si' ? 'SÍ' : 'NO'` — como PAIC nunca manda el campo, la
   **columna 16 de `CLIENTES_MAESTRO` queda en `'NO'`**, que es una afirmación, no una ausencia.
3. Peor: como el registro es un **upsert que sobrescribe la fila completa** (`BACKEND_FIXES.gs:783`),
   si el cliente ya se había registrado por SEAPD declarando `SÍ`, **el registro por PAIC lo pisa a `NO`**.
4. El correo interno imprime el chip `PIPC · NO REQUIERE` (`BACKEND_FIXES.gs:2337`) — Operaciones lee
   una negación explícita de algo que nunca se preguntó.

#### 🟠 A2 · `jefe_mantenimiento` — el flujo NOM-020 queda a medias

`3e4ac3b` completó NOM-020 en SEAPD (`SEAPD.html:438`) y adaptó el backend para persistir el dato en la
hoja de perfil y mostrarlo en el correo. **PAIC tiene toda la sección NOM-020 menos ese campo.** Para un
registro de PAIC con NOM-020 = sí, la fila 45 de la hoja de perfil sale vacía y el correo interno
imprime “Jefe de mantenimiento: —”. La mitad de la corrección llegó (el renombrado de INE con el nombre
del testigo sí funciona, porque PAIC sí manda `testigo1` y `testigo2`); la otra mitad no.

#### 🟡 A3 · Adjuntos 2 y 3 ausentes, numeración rota, llaves muertas

SEAPD pide cuatro documentos generales. PAIC eliminó *2) Programa de mantenimiento* y
*3) Proceso de producción / Hojas de Seguridad*, pero **dejó la numeración**: el asesor ve “1)” y luego
“4)” (`PAIC.html:337` y `:350`). Y `showOnRecordIndicators()` (`PAIC.html:1227`) sigue recorriendo las
cuatro llaves, incluidas las dos que ya no existen en el DOM.

---

### GRUPO B — Lo que PAIC agregó y el backend nunca aprendió a recibir

#### 🔴 B1 · Los seis archivos de laboratorio e higiene se descartan en silencio

`guardarArchivos()` (`BACKEND_FIXES.gs:1978`) tiene una whitelist fija. Estos seis campos de PAIC **no
están en ella**:

```
hojas_campo_laboratorio   fotografias_laboratorio   croquis_laboratorio
hojas_campo_higiene       fotografias_higiene       croquis_higiene
```

El asesor los selecciona, se codifican en base64, viajan por la red, cuentan contra el tope de 25 MB…
y el backend los ignora. El formulario responde “✓ Información enviada”.

**El síntoma es visible y medible:** el acuse al cliente imprime `Documentos recibidos: N archivos`
contando **solo los que el backend guardó**. Un asesor que adjuntó planos + hojas de campo + fotos +
croquis recibe un acuse que dice **“1 archivo”**. Ese contador es la prueba de campo del bug.

#### 🔴 B2 · El estudio solicitado no existe para el sistema

`estudio_laboratorio`, `estudio_higiene` y `otra_nom_higiene` — el corazón de la regla “1 registro =
1 estudio” — tienen **cero referencias en el backend**. No van a `CLIENTES_MAESTRO`, no van a la hoja de
perfil, no van a ningún correo.

Como el correo de servicios solo conoce dos chips:

- **Al equipo**: un registro cuyo propósito era NOM-025 de iluminación llega anunciando
  `NOM-020-STPS · NO APLICA` y `PIPC · NO REQUIERE`. El correo **afirma que no se pidió nada**.
- **Al cliente**: `chipsServicios` queda vacío (`BACKEND_FIXES.gs:2474`) y la sección
  “Servicios solicitados” **desaparece del acuse**. El cliente nunca ve confirmado qué se le va a hacer.

La única constancia del estudio es el modal de confirmación que el asesor vio en pantalla antes de
enviar — y ese no se guarda en ninguna parte.

#### 🟠 B3 · `asesor_consultor` se guarda, pero nadie lo lee y cualquiera lo borra

Es el único campo verdaderamente propio de PAIC (`PAIC.html:206`) y llega a la columna 22 de
`CLIENTES_MAESTRO`. Pero:

- **No se prellena al actualizar.** `applyClientData()` no lo incluye en su `fieldMap`, aunque
  `fase2_BuscarClienteRFC` sí lo devuelve. Reregistrar por PAIC lo vacía.
- **SEAPD tampoco lo repuebla** — no tiene el campo. Un cliente que entró por el canal y luego se
  actualiza por SEAPD pierde a su asesor.
- **No aparece en ningún correo.** Ni el interno ni el acuse lo mencionan. El equipo que recibe el
  registro no sabe quién lo refirió sin abrir la hoja.

#### 🟠 B4 · `portal_origen: 'PAIC'` se envía y se ignora

`PAIC.html:966` lo manda en cada payload. Cero referencias en el backend; `CLIENTES_MAESTRO` no tiene
columna de origen; ningún correo lo imprime. **Operaciones no puede distinguir un registro de PAIC de
uno de SEAPD**, ni medir cuánto negocio produce el canal.

---

### GRUPO C — El modelo de correo de SEAPD no encaja con el de PAIC

#### 🟠 C1 · El acuse va al cliente, pero está dirigido al asesor

En SEAPD, quien llena el formulario y quien recibe el informe son la misma persona, así que el acuse
funciona. En PAIC son **dos personas distintas**: `nombre_solicitante` es “Nombre y puesto del asesor /
intermediario”, y `correo_informe` es el correo del cliente final.

`enviarConfirmacionCliente` construye el saludo así (`BACKEND_FIXES.gs:2453`):

```javascript
const saludo = nombre ? `Estimado(a) <strong>${data.nombre_solicitante || data.responsable}</strong>,` : …
GmailApp.sendEmail(data.correo_informe, 'Recibimos su información · …', …)
```

Es decir: **el cliente final recibe un correo que lo saluda con el nombre del intermediario.** Además de
verse mal, revela la identidad del asesor al cliente — una decisión de relación comercial que nadie tomó
explícitamente.

Y al revés: **el asesor no recibe ningún acuse**, porque PAIC no captura su correo. El único
comprobante de que su registro llegó es el mensaje verde en pantalla.

---

### GRUPO D — Deuda compartida que PAIC heredó

#### 🔴 D1 · El dictamen de calibración de NOM-020 se pierde SIEMPRE, en los dos portales

Hay **dos inputs con el mismo `name="calibracion_valvula"`**:

| | SEAPD | PAIC |
|---|---|---|
| Input visible (condicional NOM-020) | línea 546 | línea 648 |
| Input dentro de `#seccion_archivo_calibracion` | línea 562 | línea 664 |

Ese segundo contenedor está en `display:none` y **ninguna función de JavaScript lo muestra jamás** — es
marcado muerto en ambos portales. Pero `display:none` **no excluye un control del `FormData`**. En el
bucle de `confirmAndSend` (`PAIC.html:969`, `SEAPD.html:1291`):

```javascript
for (let [key, value] of formData.entries()) {
  if (value instanceof File && value.size > 0) { data[key] = await fileToBase64(value); … }
  else { data[key] = value; }          // ← el input vacío entra por aquí
}
```

El primer input deja el base64 en `data.calibracion_valvula`; el segundo, vacío, entra por el `else` y
**sobrescribe ese base64 con un `File` vacío**. `guardarArchivos()` exige `typeof fileData === 'string'
&& fileData.startsWith('data:')`, así que descarta el archivo.

**Resultado: el dictamen de calibración nunca se guarda, en ningún registro, desde ningún portal.** Es
un agujero dentro del mismo flujo NOM-020 que `3e4ac3b` dio por completo.

#### 🟠 D2 · El upsert sobrescribe la fila completa

`fase1_RegistrarCliente` no hace merge: al encontrar la fila del RFC + Sucursal la reescribe con los 22
valores del payload (`BACKEND_FIXES.gs:783`). Cualquier campo que el portal no mande se borra. Es el
mecanismo detrás de A1 y B3, y afecta a los dos portales.

#### 🟡 D3 · `sucursal` es texto libre y tiene tres normalizaciones distintas

| Consumidor | Normalización |
|---|---|
| `CLIENTES_MAESTRO` | **ninguna** — `data.sucursal` crudo |
| Carpeta de Drive | `sanitizeFileName()` + `toLowerCase()` |
| Búsqueda de cliente | `.trim()`, con `'Matriz'` por defecto |

`sanitizeFileName()` reemplaza lo no alfanumérico por `_`, **no hace trim** y corta a 50 caracteres. Un
espacio al inicio basta para que la carpeta no coincida y para que SEAOT no encuentre al cliente.

#### 🟡 D4 · El prellenado promete dos campos que el backend nunca devuelve

`applyClientData()` mapea `actividad_principal` y `descripcion_proceso`; `fase2_BuscarClienteRFC` no los
devuelve porque no existen esas columnas. Siempre llegan `undefined`. **Idéntico en los dos archivos** —
es un bug copiado, no una divergencia.

#### 🟡 D5 · Los teléfonos de los portales ya no coinciden con los del backend

Ambos portales enlazan `wa.me/522791113533`. El backend usa `SUPPORT_WHATSAPP: '56 5282 1561'`, la línea
móvil de Atención a Clientes que fijó `ebbb70e`. El cliente ve un número en el formulario y otro
distinto en el correo de confirmación.

#### ⚪ D6 · Escapado, `alert()` y etiquetas de versión

Ninguno de los dos portales escapa `data.razon_social` al pintar el resultado de la búsqueda por RFC
(`PAIC.html:1121`), pese a que el backend ya tiene `escHtml_()` para sus correos. Ambos usan `alert()`
para validar la selección de sucursal. Y las etiquetas de versión (`v3.0` / `PAIC v1.0`) no se han
tocado desde la bifurcación.

#### ⚪ D7 · Pruebas y documentación

El caso E01 de `TESTS_BACKEND.gs` prueba `registrarCliente` etiquetado como “SEAPD”; **no hay ningún
caso PAIC**, ni con archivos de laboratorio, ni con reregistro. El manual (§6.1) no documenta
`asesor_consultor`, ni los campos de archivo, ni `portal_origen`. El bloque de correos del backend está
rotulado “CORREOS DE REGISTRO SEADB” cuando pertenece a la ruta SEAPD/PAIC.

---

### Resumen: modificación de SEAPD → estado en PAIC

| Modificación de SEAPD | Estado en PAIC | Brecha |
|---|---|---|
| Flujo NOM-020 completo (`3e4ac3b`) | **Parcial** — testigos sí, jefe de mantenimiento no | A2 |
| Sección PIPC | **Ausente** | A1 |
| 4 documentos generales | **2 de 4**, numeración rota | A3 |
| Correo interno rediseñado | **Llega, pero miente** — anuncia NO APLICA / NO REQUIERE | B2, B4 |
| Acuse al cliente rediseñado | **Llega incompleto** — sin servicios, con el conteo de archivos equivocado, dirigido al asesor | B1, B2, C1 |
| Renombrado de archivos por etiqueta | Funciona | — |
| reCAPTCHA Enterprise (`4789d1c`) | **Adoptado** | — |
| Guardado del dictamen de calibración | **Roto en ambos** | D1 |

---

## 3. Plan por fases

Principio rector, heredado de `PROPUESTA_TRAZABILIDAD.md`: **no romper contratos.**
`CLIENTES_MAESTRO` = 22 columnas (guard en `BACKEND_FIXES.gs:694`). Todo lo que sigue lo respeta.

---

### Fase 0 — Decidir el modelo de convergencia

Antes de escribir código hay que decidir **cómo dejan de divergir**. Tres caminos:

| | A · Un solo formulario con modo | B · Núcleo compartido | C · Seguir sincronizando a mano |
|---|---|---|---|
| Forma | Un HTML, `?modo=asesor` activa la capa de PAIC | `portal-core.js` compartido + un HTML por portal | Lo de hoy |
| Elimina la bifurcación | Sí, de raíz | En el JavaScript (que es el 95% idéntico) | No |
| Costo inicial | Alto | Medio | Cero |
| Costo por cada cambio futuro | Nulo | Bajo | **Se paga otra vez, siempre** |
| Riesgo | Un cambio afecta a los dos portales a la vez | Acotado | Que la próxima corrección vuelva a llegar a uno solo |

**Recomendación: B.** El repo ya tiene el patrón funcionando — `auth.js` y `recaptcha.js` son módulos
compartidos por varios HTML. Extraer `portal-core.js` (manejo de archivos, condicionales genéricos,
búsqueda por RFC, prellenado, envío y errores) deja en cada HTML solo su propio marcado y su propia
configuración. A es el destino ideal, pero PAIC y SEAPD difieren en secciones completas y conviene
llegar ahí después de la Fase 2, no antes.

---

### Fase 1 — Reparar lo que se pierde

*Riesgo: bajo · Solo backend y marcado muerto · Beneficia a los dos portales.*

1. **Quitar el input duplicado de `calibracion_valvula`** y el contenedor muerto
   `#seccion_archivo_calibracion` de ambos portales. Es la corrección de mayor valor por línea escrita
   de todo el plan. *(D1)*
2. **Agregar los 6 campos de laboratorio e higiene** a la whitelist de `guardarArchivos()`, con
   etiquetas propias. `validarArchivo_` ya cubre tipo y tamaño. *(B1)*
3. **Merge en vez de sobrescritura** en `fase1_RegistrarCliente`: si el payload no trae un campo y la
   fila existente sí tiene valor, conservarlo. Aplicar al menos a `ASESOR_CONSULTOR` y `REQUIERE_PIPC`.
   Sutileza: `requiere_pipc` colapsa hoy a `SÍ`/`NO` y pierde la diferencia entre “dijo que no” y “no se
   preguntó” — el merge debe distinguir **campo ausente** de **campo en `'no'`**. *(D2)*
4. **Pruebas**: caso “el dictamen de calibración se guarda”, caso E01 con variante PAIC y archivos de
   laboratorio, caso “reregistro por PAIC no borra `ASESOR_CONSULTOR` ni `REQUIERE_PIPC`”. *(D7)*

---

### Fase 2 — Paridad de captura: completar PAIC contra SEAPD

*Riesgo: bajo · Solo marcado en `PAIC.html`.*

1. **`jefe_mantenimiento`** con el mismo `data-conditional-required="nom020"` de SEAPD. Es un `<input>`;
   el backend ya lo persiste y ya lo imprime. *(A2)*
2. **Compuerta `requiere_pipc`.** Mínimo el radio, para que la columna 16 deje de mentir. Si además el
   asesor puede tramitar PIPC, portar la sección completa con sus 12 documentos, sus 4 sub-compuertas y
   `resetPIPCConditionals()`. **Esto es la pregunta 2 de §4.** *(A1)*
3. **Adjuntos 2 y 3**: portarlos o renumerar a “1) 2)”. Y limpiar las dos llaves muertas de
   `showOnRecordIndicators()`. *(A3)*
4. **Prellenar `asesor_consultor`** en el `fieldMap`, y quitar los dos campos que el backend nunca
   devuelve. *(B3, D4)*

---

### Fase 3 — Que el sistema entienda a PAIC

*Riesgo: medio · Backend y correos · Aquí hay una decisión de arquitectura.*

1. **Persistir el estudio solicitado.** *(B2)*

   | | Hoja `SOLICITUDES` nueva (append-only) | Columna 23 en `CLIENTES_MAESTRO` |
   |---|---|---|
   | Contrato de 22 columnas | intacto | **roto** (guard + índices `CL.*` en 5 funciones) |
   | Modelo | “1 registro = 1 estudio” es un **evento** | lo trata como atributo del cliente |
   | Historial | conserva cada solicitud | la última pisa a la anterior |
   | Beneficio extra | embudo medible **solicitud → OT**, y SEAOT puede prellenar la NOM | ninguno |

   **Recomendación: hoja nueva**, con timestamp, folio, `portal_origen`, asesor, RFC, sucursal, estudio
   solicitado, fechas preferidas, link de carpeta y estatus de atención.

2. **Chips de servicios que incluyan el estudio de PAIC.** Hoy son dos constantes; deben construirse a
   partir de lo que el registro realmente pidió. El correo interno debe decir “NOM-025-STPS ·
   Iluminación”, no “NOM-020 NO APLICA”. *(B2)*
3. **Imprimir `asesor_consultor` y `portal_origen`** en el correo al equipo — dos filas en el bloque de
   contacto. Es el cambio más barato del plan y el que más contexto le da a Operaciones. *(B3, B4)*
4. **Corregir el destinatario del acuse.** Si el registro trae asesor, el saludo al cliente no debe usar
   el nombre del intermediario, y el asesor debería recibir su propio acuse — lo que implica capturar su
   correo en PAIC. **Esto es la pregunta 3 de §4.** *(C1)*

---

### Fase 4 — Ejecutar la convergencia elegida en Fase 0

*Riesgo: medio · Es refactor, no funcionalidad.*

Extraer `portal-core.js` con lo que hoy está duplicado carácter por carácter, dejar en cada HTML solo su
marcado y su configuración (`FORM_ID`, `PORTAL_ORIGEN`, `VERSION`), y actualizar las etiquetas de
versión. A partir de aquí, **una corrección se escribe una vez**. *(D6)*

---

### Fase 5 — Normalizar `sucursal`

*Riesgo: medio · Toca coincidencias existentes.*

1. **Diagnóstico primero, solo lectura**: listar por RFC todas las variantes de sucursal en
   `CLIENTES_MAESTRO`, `ORDENES_TRABAJO` y Drive. Sin ese inventario no se decide nada — dirá si el
   problema son 3 filas o 300.
2. **Una sola función** `normalizarSucursal_()`: trim → colapsar espacios → Title Case, aplicada al
   escribir y al comparar.
3. **En los portales**: selector con las sucursales que el backend ya devuelve, más opción “nueva” con
   normalización en vivo y previsualización del nombre de carpeta.
4. **Sin reescribir históricos en el mismo PR.**

---

### Fase 6 — Cerrar la deuda

- Sincronizar los teléfonos de los portales con `CONFIG.SUPPORT_PHONE` / `SUPPORT_WHATSAPP`. *(D5)*
- `escHtml()` en el render del modal de RFC de ambos portales. *(D6)*
- Manual §3.5 (PAIC real), §6.1 (payload completo con `asesor_consultor`, archivos y `portal_origen`).
  Corregir el rótulo “CORREOS DE REGISTRO SEADB”. *(D7)*

---

## 4. Decisiones que necesito de ti

Las Fases 1 y 2.1 **no dependen de estas respuestas** — se pueden arrancar ya.

1. **§3 Fase 0 — ¿modelo de convergencia?** A (un formulario con modo), **B (núcleo compartido,
   recomendado)** o C (seguir sincronizando a mano). Define si esta conversación se repite en seis meses.
2. **§3 Fase 2.2 — ¿el asesor puede tramitar PIPC?** Si sí, se porta la sección completa (12 documentos
   + 4 sub-compuertas). Si no, basta el radio para que la columna deje de mentir.
3. **§3 Fase 3.4 — ¿el acuse al cliente puede revelar al asesor?** Y en paralelo: ¿PAIC debe capturar el
   correo del asesor para mandarle su propio acuse?
4. **§3 Fase 3.1 — ¿hoja `SOLICITUDES` nueva (recomendado) o columna 23?** En
   `PROPUESTA_TRAZABILIDAD.md` decidiste “sin crear hojas nuevas”; aquí la propongo igual porque no es un
   control operativo interno sino el registro del canal de negocio. Tu llamada.

---

## 5. Orden sugerido (de menor a mayor riesgo)

| # | Trabajo | Riesgo | Desbloquea |
|---|---|---|---|
| 1 | D1 — quitar el input duplicado de calibración | bajo | El dictamen NOM-020 deja de perderse **en los dos portales** |
| 2 | Fase 1 completa (B1, D2) + sus pruebas | bajo | Deja de perderse información hoy |
| 3 | Fase 2 — paridad de captura | bajo | PAIC llena lo mismo que SEAPD |
| 4 | Fase 3.2–3.3 — correos que digan la verdad | bajo | Operaciones ve el estudio y el asesor |
| 5 | Fase 3.1 — hoja `SOLICITUDES` | medio | Embudo del canal + prellenado de NOM en SEAOT |
| 6 | Fase 3.4 — destinatarios del acuse | medio | Relación asesor ↔ cliente bien planteada |
| 7 | Fase 4 — convergencia | medio | Que la próxima corrección llegue a los dos |
| 8 | Fase 5 — normalizar sucursal | medio | Carpetas y búsquedas confiables |

---

## 6. Qué NO hacer

- **No** ampliar `CLIENTES_MAESTRO` más allá de 22 columnas sin migrar a la vez el guard y los índices
  `CL.*`; hay lecturas por índice en cinco funciones distintas.
- **No** corregir un bug solo en un portal. Todo lo del grupo D vive en los dos archivos: si se arregla
  en uno, la bifurcación se ensancha.
- **No** portar la sección PIPC a PAIC antes de responder la pregunta 2: son ~340 líneas de marcado que
  no se deben escribir dos veces ni por gusto.
- **No** reescribir los datos históricos de sucursal en el mismo PR que introduce la normalización.
- **No** tocar el orden ni el formato de las celdas de `generarPerfilSheet`: la hoja de perfil se llena
  por coordenada fija (`B3`, `D5`, `B45`…) y cualquier inserción desplaza todo lo de abajo.
