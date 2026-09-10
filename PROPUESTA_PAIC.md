# Propuesta — PAIC autónomo, con SEAPD congelado

**Estado:** Plan para revisión (NO implementado).
**Fecha:** 2026-09-10
**Alcance:** `PAIC.html` y ramas aditivas en `BACKEND_FIXES.gs`. **`SEAPD.html` no se toca.**

> **Decisiones de Eduwin que esta versión incorpora:**
> 1. **Ninguna modificación a SEAPD.** Ni al archivo ni a su comportamiento.
> 2. PAIC debe **operar como SEAPD** pero conservando sus distinciones: **sin PIPC**, **un solo
>    servicio por registro** (hoy no se cumple) y **la notificación al correo del intermediario,
>    nunca al cliente**.
> 3. Pregunta abierta que este documento responde en §3: **¿compagina con los registros derivados
>    en Sheets, Drive y los módulos de abajo?**

---

## 1. El principio que hace posible todo lo demás: `portal_origen` es la costura

PAIC y SEAPD comparten el mismo endpoint `registrarCliente` y las mismas funciones de backend. Si
SEAPD no se puede tocar, **toda diferencia de PAIC tiene que colgar de una condición explícita**:

```javascript
if (data.portal_origen === 'PAIC') { … }   // rama nueva
else { …lo que ya hace hoy, intacto… }     // camino de SEAPD
```

PAIC ya manda `portal_origen: 'PAIC'` en cada payload (`PAIC.html:966`) y hoy el backend lo ignora.
Convertirlo en la costura no cuesta nada y da una **regla de revisión verificable**:

> **Si borras todas las ramas `portal_origen === 'PAIC'`, el backend debe quedar exactamente como
> está hoy.** Cualquier PR que no cumpla eso está tocando SEAPD.

SEAPD nunca envía ese campo, así que su camino queda idéntico byte por byte.

### 1.1 Lo que la restricción cuesta — dicho de frente

Congelar SEAPD tiene tres consecuencias que conviene aceptar a ojos abiertos, no descubrirlas después:

| Costo | Detalle |
|---|---|
| 🔴 **El dictamen de calibración se seguirá perdiendo en SEAPD** | Hay dos `<input name="calibracion_valvula">` en cada portal: el visible del condicional NOM-020 y otro dentro de `#seccion_archivo_calibracion`, un contenedor `display:none` que **ningún JavaScript muestra jamás**. `display:none` **no** excluye un control del `FormData`, así que el segundo (vacío) sobrescribe el base64 del primero y `guardarArchivos()` lo descarta. Se corrige en PAIC (`PAIC.html:648` y `:664`); en SEAPD (`:546` y `:562`) **queda vivo**. Cuando lo autorices, son 12 líneas borradas. |
| 🟠 **No habrá núcleo compartido** | Las 18 funciones de JavaScript de PAIC son las mismas de SEAPD, carácter por carácter. Sin un `portal-core.js`, cada corrección futura hay que decidir a mano si se replica. Es el precio de la autonomía. |
| 🟠 **La columna 22 seguirá siendo frágil** | `fase1_RegistrarCliente` reescribe la fila completa (`BACKEND_FIXES.gs:783`). Una rama por `portal_origen` protege a los envíos de PAIC, pero **un envío de SEAPD sobre ese mismo cliente seguirá vaciando `ASESOR_CONSULTOR`**, y evitarlo exige tocar el camino de SEAPD. → **Por eso el asesor no puede vivir solo en la columna 22.** Ver §3.5: esto convierte la hoja `SOLICITUDES` de recomendable en **necesaria**. |

---

## 2. Los tres requisitos, diseñados

### 2.1 · Sin PIPC — qué falta además de no portar la sección

La sección ya no existe en PAIC, así que el requisito está cumplido en el formulario. Lo que **no**
está resuelto es lo que ese hueco provoca aguas abajo:

- El backend evalúa `data.requiere_pipc === 'si' ? 'SÍ' : 'NO'`. Como PAIC nunca manda el campo, la
  **columna 16 de `CLIENTES_MAESTRO` se escribe como `'NO'`** — una afirmación, no una ausencia.
- Y como el registro es un upsert que reescribe la fila completa, si el cliente ya se había
  registrado por SEAPD declarando `SÍ`, **el registro por PAIC lo pisa a `NO`**.
- El correo interno imprime el chip `PIPC · NO REQUIERE`: Operaciones lee una negación explícita de
  algo que en PAIC nunca se preguntó.

**Diseño.** Dos ramas por `portal_origen`, ninguna toca a SEAPD:

1. En el upsert: si la fila ya existe, **conservar el valor previo** de `REQUIERE_PIPC` en vez de
   escribir `'NO'`.
2. En el correo interno: **omitir el chip de PIPC** cuando el origen es PAIC. No se niega lo que no
   se preguntó.

### 2.2 · Un solo servicio por registro — hoy **no** se cumple

La regla “1 registro = 1 estudio” está en el copy del modal de bienvenida y en la insignia del
encabezado, pero **nada la hace cumplir**. Hoy PAIC tiene tres controles independientes:

| Control | Tipo | Opciones | ¿Obligatorio? |
|---|---|---|---|
| `estudio_laboratorio` | `<select>` | 6 NOM | No |
| `estudio_higiene` | `<select>` | 13 NOM + “otra” | No |
| `aplica_nom020` | `<radio>` | sí / no | **Sí** |

Un asesor puede enviar **Laboratorio + Higiene + NOM-020 en un mismo registro**. Y como el radio de
NOM-020 es obligatorio, hoy hay que contestarlo aunque se venga a registrar un estudio de iluminación.

**Diseño: un selector maestro.**

```
¿Qué servicio va a registrar?  (obligatorio, una sola opción)
├── Laboratorio        → 6 NOM
├── Higiene            → 13 NOM + "Otra NOM STPS"
└── Inspección         → NOM-020-STPS · Recipientes, calderas y generadores de vapor
```

- Al elegir se abre **un solo** bloque de adjuntos; el resto ni se muestra ni queda `required`.
- **`aplica_nom020` se deriva del selector** y se sigue enviando con el mismo nombre y los mismos
  valores (`'si'` / `'no'`). Así la columna 15, la hoja de perfil y los chips del correo **siguen
  funcionando exactamente igual** — el contrato con el backend no se toca.
- El modal de confirmación pasa de listar bloques a mostrar **un solo servicio**.

**Hallazgo que hay que resolver al unificar las listas:** cuatro NOM están en ambos selectores con
valores distintos.

| NOM | Valor en Laboratorio | Valor en Higiene |
|---|---|---|
| Ruido | `NOM-011-STPS` | `NOM-011-STPS-2001` |
| Condiciones térmicas | `NOM-015-STPS` | `NOM-015-STPS-2001` |
| Vibraciones | `NOM-024-STPS` | `NOM-024-STPS-2001` |
| Iluminación | `NOM-025-STPS` | `NOM-025-STPS-2008` |

Al fundirlas en una sola lista hay que **fijar un valor canónico por NOM**, porque ese texto es el que
el operador vuelve a teclear en SEAOT y termina en `ORDENES_TRABAJO` e `INFORMES` — donde SEADB lo usa
para calcular renovaciones. **Es la pregunta 4 de §5.**

### 2.3 · La notificación al intermediario, nunca al cliente

Hoy pasa exactamente lo contrario de lo que pides. `enviarConfirmacionCliente` manda el acuse a
`data.correo_informe` —el correo del **cliente**— y lo saluda con el nombre del **asesor**:

```javascript
const saludo = `Estimado(a) <strong>${data.nombre_solicitante || data.responsable}</strong>,`;
GmailApp.sendEmail(data.correo_informe, 'Recibimos su información · …', …);   // línea 2554
```

En SEAPD eso funciona porque quien llena el formulario y quien recibe el informe son la misma persona.
En PAIC son dos, y el resultado es que **el cliente final recibe un correo dirigido al intermediario**.

#### Lo que muestran los registros reales

Dos filas reales de `CLIENTES_MAESTRO` que entraron por PAIC (se reconocen porque traen asesor en la
columna 22, que SEAPD no captura):

| | Fila A · CRECYDE | Fila B · Protección Ambiental |
|---|---|---|
| Col. 8 · Solicitante | `YOLANDA QUINTO SANCHEZ` | `PRISCILA HERNANDEZ PULIDO` |
| **Col. 9 · Correo** | `yolandaquinto3@yahoo.com` | `proteccionambiental08@gmail.com` |
| Col. 17 · Responsable en sitio | `YARIBET ORDUÑO CRUZ` | `LIZETH CINTHIA AGUILAR TORRES` |
| Col. 22 · Asesor | `CRECYDE` | `PROTECCIÓN AMBIENTAL, CONSULTORÍA…` |

En las dos, **el correo de la columna 9 es el del intermediario, no el del cliente**: en la fila B
coincide con el despacho asesor, y en la A con la persona que llenó el formulario, que no es la
responsable en sitio.

> **Conclusión: “no ha habido problemas” porque los intermediarios ya lo resuelven solos**, poniendo su
> propio correo en el campo del cliente. El sistema no resolvió nada — lo evadieron. El acuse les llega
> por accidente, no por diseño.

Y deja dos efectos colaterales: el **cliente final no puede entrar a PORTAL** (su correo no está
registrado), y si mañana se registra otra sucursal del mismo RFC con el correo real del cliente,
**el código de acceso se manda a los dos a la vez** (§3.3).

#### El caso difícil: el intermediario que también es el cliente

A veces el intermediario está dentro de un corporativo que además es cliente. Ahí “asesor” y “cliente”
no son dos organizaciones distintas, así que **cualquier regla que intente adivinar quién es quién se
va a equivocar**. La salida es no adivinar: que lo diga quien llena el formulario.

#### Diseño — un campo y un condicional

**En PAIC**, un campo nuevo obligatorio, junto a los datos del asesor:

> **Correo para la confirmación y el seguimiento** — *aquí te llega el acuse de este registro*
> ☐ Es el mismo correo del informe

`correo_informe` se queda como está (correo del cliente, destino del informe y llave de acceso a
PORTAL), solo con la etiqueta aclarada para que quede claro que **ahí no llega el acuse**.

**En el backend**, tres líneas detrás de la costura:

```javascript
const destinoAcuse = (data.portal_origen === 'PAIC' && data.correo_acuse)
  ? data.correo_acuse
  : data.correo_informe;
```

| Caso | `correo_acuse` | `correo_informe` | Resultado |
|---|---|---|---|
| Intermediario externo | del asesor | del cliente | Acuse al asesor; informe y portal al cliente |
| Intermediario corporativo, que también es el cliente | el suyo | el mismo | Todo al mismo, coherente |
| Cliente que se registra solo por PAIC | el suyo | el suyo | Igual que hoy |

Tres razones por las que esta es la ruta barata:

1. **SEAPD ni se entera**: nunca manda `portal_origen` ni `correo_acuse`, así que el ternario siempre
   cae del lado de hoy.
2. **El saludo se corrige solo.** El acuse ya dice “Estimado(a) {nombre_solicitante}”, que es el
   intermediario; al cambiar el destinatario, el saludo pasa a ser correcto **sin tocar el correo**.
3. **No hay que migrar nada.** Lo ya registrado sigue funcionando igual.

Lo demás del despacho ya es seguro: el correo interno va a `CONFIG.EMAIL_TO`, y
`enviarEmailSimpleFallback` también — la ruta de error no filtra nada al cliente.

**Pero hay una fuga que el correo no cubre.** El correo interno trae un botón *“Contactar por
WhatsApp”* cuyo destino es `contactoWhatsAppCliente_(data)`, que toma **`telefono_responsable`
primero** — el contacto del cliente en sitio. Es decir: la acción de un clic que Operaciones
efectivamente ejecuta va **directo al cliente, saltándose al asesor**. Si la regla es “nunca al
cliente”, esto pesa más que el correo. **Es la pregunta 1 de §5.**

### 2.4 · Otros hallazgos de la validación de campo

Del mismo par de registros reales, tres cosas que hoy no rompen nada pero conviene atender:

| Hallazgo | Qué pasa | Gravedad |
|---|---|---|
| `SUCURSAL = "N/A"` | La carpeta se creó como `N_A` y `folderMatchesClientBranch_` sí la encuentra, así que SEAOT funciona. Pero “N/A” no es una sucursal: el día que se registre una planta real de ese RFC quedarán `N/A` y `Planta X` como hermanas, con el historial colgando de la que no significa nada. Debería ser `Matriz`. | 🟡 |
| Teléfono de 11 dígitos (`22212246015`) | `normalizarTelefonoWhatsApp_` acepta 10, 12 con `52` o 13 con `521`; con 11 devuelve `null`, así que el botón de WhatsApp **cae silenciosamente al teléfono de la empresa**. Aquí no importó porque era casi el mismo número, pero **PAIC no valida formato de teléfono en ningún campo**. | 🟡 |
| `… S.A DE C.V` sin todos los puntos | `LEGAL_SUFFIX_REGEX_` busca `S.A. DE C.V.`, ` SA DE CV`, ` S.A.` o ` S.C.`; este texto no coincide con ninguno, así que el sufijo no se limpió y la carpeta quedó `AAC151218QKA - AGUACATES ACUITZIO DEL CANJE S_A DE C_V`. No fragmenta nada —la reutilización por RFC lo cubre—, solo se ve mal. | ⚪ |

Y confirmado en las dos filas: **`REQUIERE_PIPC = NO`**, tal como describe §2.1. Los RFC de ambas
(`CCO8605231N4`, `AAC151218QKA`) pasan `validarRFC_` sin problema.

### 2.5 · Límite estructural: la columna 22 guarda un asesor por cliente, no por servicio

Al precisar la regla de preservación (§4, Fase 1.2) aparece un problema que **ninguna versión de esa
regla resuelve**, porque no está en la regla sino en la forma del dato.

`fase4_GetTablero` no lee el asesor de la OT: lo resuelve en vivo con un mapa de
**una sola entrada por cliente y sucursal**, y se lo asigna a *todas* las OT de esa pareja:

```javascript
asesorMap[rfc + '|' + suc] = asesor;                    // una entrada por RFC|Sucursal
…
asesor_consultor: asesorMap[rfc + '|' + suc] || '',     // aplicada a TODAS las OT de esa pareja
```

Entonces, si el asesor A trajo a un cliente y meses después el asesor B registra otro servicio para la
misma sucursal, ninguna de las dos opciones es correcta:

| Regla | Qué rompe |
|---|---|
| Conservar siempre a A | El servicio nuevo de B queda atribuido a A |
| Guardar siempre a B (lo de hoy) | **Todas las OT históricas de A pasan a mostrar a B**, retroactivamente |

La columna 22 simplemente no tiene la forma del dato: el asesor es un atributo **del servicio**, no
**del cliente**. Por eso:

- **`SOLICITUDES` es la fuente autoritativa** del asesor por registro — ahí sí hay una fila por
  servicio, con su `ot_folio`.
- **La columna 22 se degrada a conveniencia**: “último asesor conocido de esta sucursal”. La regla de
  §4 (guardar el que viene, conservar solo si viene vacío) es la correcta para ese propósito.
- **Límite conocido y aceptado por ahora**: mientras SEADB siga resolviendo el asesor con `asesorMap`,
  su columna “Asesor/Consultor” seguirá siendo *el último*, no *el de cada OT*. Corregirlo es trabajo
  de SEADB —resolver el asesor por OT vía `SOLICITUDES.ot_folio`— y queda **fuera de este alcance**,
  anotado aquí para no descubrirlo después.

---

## 3. ¿Compagina con lo que se deriva? — el rastreo completo

Esta es la pregunta que faltaba. Seguí un registro de PAIC desde el `POST` hasta los cuatro módulos
que lo consumen.

### 3.1 Recorrido

```
PAIC ──POST registrarCliente──► fase1_RegistrarCliente
                                      │
        ┌─────────────────────────────┼─────────────────────────────┐
        ▼                             ▼                             ▼
  CLIENTES_MAESTRO            Drive: {RFC} - {EMPRESA}/        2 correos
  fila de 22 columnas           └── {Sucursal}/
        │                            ├── 01_Cliente/  ← perfil + archivos
        │                            └── 02_Expediente_…/  ← lo crea SEAINF después
        │
        ├──► SEAOT    busca por RFC, exige carpeta exacta, crea la OT
        ├──► SEAINF   crea el expediente dentro de la carpeta de la sucursal
        ├──► SEADB    resuelve el asesor en vivo con asesorMap[RFC|Sucursal]
        └──► PORTAL   autentica al cliente por el correo de la columna 9
```

### 3.2 Columna por columna

| Col | Campo | Qué escribe PAIC | ¿Compagina? |
|---|---|---|---|
| 1–8, 10–14 | Datos generales | Igual que SEAPD | ✅ |
| 9 | `CORREO` | `correo_informe` = correo del **cliente** | ✅ **y debe seguir así** — ver §3.3 |
| 15 | `APLICA_NOM020` | Derivado del selector maestro | ✅ contrato intacto |
| 16 | `REQUIERE_PIPC` | Siempre `'NO'` | ⚠️ pisa el `SÍ` de un registro previo → §2.1 |
| 17–20 | Responsable y destinatarios | Igual que SEAPD | ✅ |
| 21 | `LINK_DRIVE` | Carpeta de la sucursal | ✅ |
| 22 | `ASESOR_CONSULTOR` | El asesor | ⚠️ se borra al re-registrar → §1.1 y §3.5 |

### 3.3 🔴 El hallazgo grave: PORTAL manda el código de acceso a **todos** los correos del RFC

`PORTAL/` autentica a los clientes con un código de un solo uso enviado al correo registrado en
`CLIENTES_MAESTRO` **columna 9**. Y `portal_resolverCliente_` no toma un correo: **reúne todos los
correos distintos de todas las filas de ese RFC** y los manda juntos:

```javascript
correosMap[c.toLowerCase()] = c;          // recorre TODAS las filas del RFC
GmailApp.sendEmail(cliente.correos.join(','), asunto, …);   // envía a TODOS
```

Consecuencia directa para este proyecto: **si el correo del asesor llegara alguna vez a la columna 9
—en cualquier sucursal de ese RFC— el código de acceso del cliente le llegaría también al asesor**, y
al revés. Sería la violación más grave posible de “la notificación nunca al cliente”, pero en espejo.

> **Regla que sale de aquí: el correo del asesor NO puede vivir en `CLIENTES_MAESTRO`.**
> Ni en la columna 9 ni en ninguna otra que PORTAL indexe. Tiene que estar en una estructura que
> PORTAL no lea — y eso apunta a la hoja `SOLICITUDES` de §3.5.

### 3.4 🟠 Los archivos del estudio caen donde nadie los va a buscar

`fase1_RegistrarCliente` deposita todo lo que suba el portal en `{Sucursal}/01_Cliente`. Más tarde
SEAINF crea un **hermano**, `02_Expediente_{consecutivo}_{OT}_{NOM}`, con sus seis subcarpetas:

```
{Sucursal}/
├── 01_Cliente/                       ← aquí caen las hojas de campo de PAIC,
│                                        mezcladas con INEs y actas constitutivas
└── 02_Expediente_0001_OT2603-001_NOM-025/
    ├── 1. ORDEN_TRABAJO/
    ├── 2. HDC/                       ← aquí es donde deberían estar
    ├── 3. CROQUIS/                   ← y aquí
    ├── 4. FOTOS/                     ← y aquí
    ├── 5. INFORMES Y MEMORIAS/
    └── 6. INFORME PRELIMINAR/
```

Las hojas de campo, croquis y fotos que sube el asesor son **exactamente** el material de las
subcarpetas 2, 3 y 4 del expediente. Pero el expediente **todavía no existe** cuando se registra: nace
con la OT, después. Así que no se pueden depositar ahí directamente — hay que decidir dónde esperan.
**Es la pregunta 3 de §5.**

> Hoy ni siquiera llegan a `01_Cliente`: los seis campos de archivo de PAIC no están en la whitelist
> de `guardarArchivos()` y **se descartan en silencio**. El síntoma comprobable es que el acuse
> imprime “Documentos recibidos: N archivos” contando solo los guardados — un asesor que subió cuatro
> archivos recibe un acuse que dice **“1 archivo”**.

### 3.5 🟠 El servicio solicitado se pierde y hay que volver a teclearlo

`estudio_laboratorio`, `estudio_higiene` y `otra_nom_higiene` tienen **cero referencias en el
backend**. No van a `CLIENTES_MAESTRO`, no van a la hoja de perfil, no van a ningún correo. La única
constancia es el modal que el asesor vio en pantalla.

Eso rompe la cadena en dos puntos:

- **SEAOT**: la NOM de la OT la vuelve a teclear un operador. No hay forma de auditar “lo que pidió el
  asesor” contra “lo que se registró”, y ese texto es el que después alimenta las renovaciones de SEADB.
- **SEADB**: sin vínculo solicitud → OT, no se puede medir cuánto negocio produce el canal.

Súmale que (a) el correo del asesor no puede vivir en `CLIENTES_MAESTRO` (§3.3) y (b) la columna 22
seguirá siendo frágil mientras SEAPD esté congelado (§1.1), y las tres necesidades convergen en la
misma solución:

> **Una hoja `SOLICITUDES`, solo append.** Deja de ser “recomendable”: es la única estructura donde
> caben el servicio solicitado, el asesor y su correo, sin tocar el contrato de 22 columnas, sin que
> PORTAL los vea y sin depender de que SEAPD respete un campo que no conoce.

Columnas propuestas: `timestamp · folio · portal_origen · asesor · correo_acuse · telefono_asesor ·
RFC · sucursal · servicio_canonico · fechas_preferidas · link_carpeta · estatus · ot_folio`
(`RECIBIDA` → `OT_GENERADA` → `DESCARTADA`).

#### Por qué `ot_folio` no es opcional

Con la regla de un servicio por registro, **el mismo RFC y sucursal va a tener varias solicitudes
abiertas a la vez** — es la consecuencia esperada, no un caso raro. Sin un identificador que ate cada
solicitud a su OT:

- SEAOT no puede saber **cuál** de las tres solicitudes pendientes está atendiendo, así que el
  prellenado de la NOM no tiene de dónde elegir.
- El embudo solicitud → OT no se puede auditar: `estatus = OT_GENERADA` diría *que* se generó una, pero
  no *cuál*.

El enlace tiene que vivir del lado de `SOLICITUDES`, no de `ORDENES_TRABAJO`. Se podría agregar una
columna R a la hoja de OT —el guard es `getMaxColumns() < 17`, un mínimo, así que no reventaría—, pero
`fase2_RegistrarOT` escribe exactamente 17 valores con `appendRow`, así que habría que tocarlo y
arrastrar a SEAOT a la conversación. Poner `ot_folio` en la hoja nueva cuesta cero y no toca ningún
contrato vigente.

### 3.6 🟠 La hoja de perfil no tiene dónde poner lo de PAIC

`generarPerfilSheet` escribe **por coordenada fija** (`B3`, `D5`, `B45`…). No hay celda para el
servicio solicitado ni para el asesor, y el bloque NOM-020 de las filas 45–50 —que `3e4ac3b` agregó
para SEAPD— sale con el jefe de mantenimiento vacío, porque PAIC no captura ese campo.

**Cuidado al tocarla:** cualquier inserción de filas desplaza todo lo de abajo. Si se le agrega un
bloque a PAIC, tiene que ir **al final** y detrás de la rama `portal_origen`.

### 3.7 Lo que sí compagina sin tocar nada

| Módulo | Estado |
|---|---|
| **SEAOT** — búsqueda por RFC y resolución de carpeta | ✅ funciona, siempre que `sucursal` coincida exactamente |
| **SEAINF** — expediente dentro de la sucursal | ✅ funciona, no depende del portal de origen |
| **TRAZ** — trazabilidad OT ↔ INFORMES | ✅ no toca el registro |
| **Renombrado de archivos por etiqueta** | ✅ PAIC sí manda `testigo1` y `testigo2` |
| **reCAPTCHA Enterprise** | ✅ ya adoptado |

⚠️ Con una salvedad transversal: **`sucursal` es texto libre** y el sistema la normaliza de tres
maneras distintas — cruda en la hoja, `sanitizeFileName()` + minúsculas en Drive, `.trim()` en las
búsquedas. `sanitizeFileName()` **no hace trim**, así que un espacio al inicio basta para que SEAOT no
encuentre la carpeta y falle con `CARPETA_NO_ENCONTRADA`. Es deuda compartida, pero PAIC la puede
mitigar por su cuenta (§4, Fase 4).

---

## 4. Plan por fases

Todas las fases respetan la regla de §1: **borrar las ramas de PAIC devuelve el backend a su estado
actual.**

### Fase 1 — Abrir la costura `portal_origen` (backend, aditivo)

*Riesgo: bajo · Nada de esto cambia el camino de SEAPD.*

1. **Ampliar la whitelist de `guardarArchivos()`** con los seis campos de PAIC:
   `hojas_campo_laboratorio`, `fotografias_laboratorio`, `croquis_laboratorio`,
   `hojas_campo_higiene`, `fotografias_higiene`, `croquis_higiene`. SEAPD no envía esas llaves, así
   que para él es literalmente un no-op. *(§3.4)*
2. **Rama de preservación en el upsert**, con `portal_origen === 'PAIC'` y fila existente. La regla es
   **conservar solo cuando el payload no trae el dato**, nunca de forma incondicional:
   - `REQUIERE_PIPC`: PAIC nunca lo manda, así que siempre se conserva el valor previo en vez de
     escribir `'NO'`. Sutileza: hay que distinguir **campo ausente** de **campo en `'no'`**, porque hoy
     los dos colapsan al mismo `'NO'`.
   - `ASESOR_CONSULTOR`: se conserva **solo si el payload viene vacío**. Si el registro trae asesor, se
     guarda el que trae — de lo contrario un asesor B que registra un servicio nuevo quedaría atribuido
     al asesor A del registro anterior. *(§2.1 y §2.5)*
3. **Rama de correos**: el acuse se manda a `correo_acuse` en vez de a `correo_informe`. Tres líneas,
   sin tocar el cuerpo del correo. *(§2.3)*
4. **Omitir el chip de PIPC** en el correo interno cuando el origen es PAIC. *(§2.1)*

### Fase 2 — PAIC: un solo servicio y los datos del asesor

*Riesgo: bajo · Solo `PAIC.html`.*

1. **Selector maestro** con los tres grupos, excluyente, obligatorio. *(§2.2)*
2. **Derivar `aplica_nom020`** del selector, manteniendo nombre y valores.
3. **Fijar el valor canónico** de las cuatro NOM duplicadas. *(pregunta 4)*
4. **Campo `correo_acuse`** obligatorio, con la casilla “Es el mismo correo del informe” para el caso
   del intermediario corporativo. Y recalibrar la etiqueta de `correo_informe` para que quede claro
   que ahí **no** llega el acuse: es el destino del informe final y la llave de acceso del cliente al
   portal. *(§2.3, §3.3)*
5. **Quitar el input duplicado de `calibracion_valvula`** y su contenedor muerto — **solo en PAIC**.
   *(§1.1)*
6. **Modal de confirmación** con un servicio, no una lista de bloques.

### Fase 3 — Que el registro compagine hacia abajo

*Riesgo: medio · Aquí están las decisiones de arquitectura.*

1. **Hoja `SOLICITUDES`** con las columnas de §3.5. Es lo que resuelve a la vez el servicio
   solicitado, el asesor y su correo fuera del alcance de PORTAL.
2. **Subcarpeta propia para los archivos del estudio** dentro de la carpeta de la sucursal, en espera
   de que nazca el expediente. *(pregunta 3)*
3. **Chips del correo interno construidos desde el servicio real**, no desde dos constantes: que diga
   “NOM-025-STPS · Iluminación” en vez de “NOM-020 NO APLICA”.
4. **Imprimir el asesor y el origen** en el correo interno — dos filas en el bloque de contacto. Es el
   cambio más barato del plan y el que más contexto le da a Operaciones.
5. **SEAOT elige la solicitud y la cierra.** Al crear una OT para un RFC + sucursal, SEAOT lista las
   solicitudes en estatus `RECIBIDA` de esa pareja; el operador escoge una y de ahí sale la NOM
   prellenada. Al registrarse la OT, SEAOT escribe su folio en `ot_folio` de esa fila y la pasa a
   `OT_GENERADA`. Sin ese paso de selección el prellenado es ambiguo en cuanto hay dos solicitudes
   abiertas, que es el caso normal bajo la regla de un servicio por registro. *(§3.5)*

### Fase 4 — Deuda propia de PAIC

*Riesgo: bajo · Solo `PAIC.html`.*

- Adjuntos 2 y 3 ausentes con la numeración “1) … 4)” a la vista, y las dos llaves muertas que
  `showOnRecordIndicators()` sigue recorriendo (`PAIC.html:1227`).
- `escHtml()` en el render del modal de búsqueda por RFC (`PAIC.html:1121`), y `alert()` → aviso en
  línea.
- Teléfonos desincronizados: PAIC enlaza `wa.me/522791113533`; el backend usa
  `SUPPORT_WHATSAPP: '56 5282 1561'` desde `ebbb70e`.
- Mitigar `sucursal`: selector con las sucursales que `buscarClienteRFC` ya devuelve, más opción
  “nueva” con normalización en vivo y previsualización del nombre de carpeta.
- Etiqueta de versión viva.
- Quitar del prellenado `actividad_principal` y `descripcion_proceso`, que el backend nunca devuelve.

### Fase 5 — Pruebas y documentación

- Casos E2E de PAIC: registro con archivos de estudio; reregistro que **no** borra `REQUIERE_PIPC` ni
  `ASESOR_CONSULTOR`; acuse que llega **solo** al asesor; dictamen de calibración que sí se guarda.
- Prueba de regresión de SEAPD: **el mismo payload de hoy debe producir el mismo resultado de hoy.**
  Es la red de seguridad de la regla de §1.
- Manual §3.5 y §6.1 con el payload real de PAIC.

---

## 5. Decisiones que necesito de ti

Las Fases 1 y 2 (salvo el punto 3) **no dependen de estas respuestas**.

1. **¿El botón de WhatsApp del correo interno debe apuntar al asesor en los registros de PAIC?**
   Hoy apunta al teléfono del cliente en sitio. Si la regla es “nunca al cliente”, el clic que
   Operaciones sí ejecuta pesa más que el correo. *(§2.3)*
2. ~~¿El `correo_informe` del cliente sigue siendo obligatorio?~~ **Resuelta** (§2.3): sí, se queda
   como está —es la llave de PORTAL y el destino del informe— y el acuse se va por el campo nuevo
   `correo_acuse`. Queda una tarea manual asociada: en los registros ya cargados la columna 9 tiene
   el correo del intermediario; si quieres que esos clientes puedan entrar a PORTAL algún día, hay
   que corregir esas celdas a mano.
3. **¿Dónde esperan los archivos del estudio hasta que exista el expediente?** Dentro de
   `01_Cliente/` en una subcarpeta por servicio, o en un `00_Solicitudes/` aparte que SEAINF pueda
   vaciar hacia el expediente cuando lo cree. *(§3.4)*
4. **¿Qué valor canónico llevan las cuatro NOM duplicadas?** `NOM-025-STPS` o `NOM-025-STPS-2008`, y
   lo mismo para 011, 015 y 024. Ese texto termina en `ORDENES_TRABAJO` e `INFORMES`. *(§2.2)*

---

## 6. Orden sugerido

| # | Trabajo | Riesgo | Qué desbloquea |
|---|---|---|---|
| 1 | Fase 1.1 — whitelist de archivos | bajo | Los estudios dejan de perderse |
| 2 | Fase 1.2–1.4 — costura `portal_origen` | bajo | PAIC deja de pisar datos y de escribirle al cliente |
| 3 | Fase 2 — un solo servicio + datos del asesor | bajo | El requisito de negocio queda cumplido |
| 4 | Fase 5 — prueba de regresión de SEAPD | bajo | Garantiza el congelamiento |
| 5 | Fase 3.1 — hoja `SOLICITUDES` | medio | Servicio y asesor sobreviven, fuera del alcance de PORTAL |
| 6 | Fase 3.2 — destino de los archivos | medio | El expediente nace con su material |
| 7 | Fase 3.3–3.5 — correos y prellenado de SEAOT | medio | Embudo solicitud → OT medible |
| 8 | Fase 4 — deuda propia | bajo | Calidad del portal |

---

## 7. Qué NO hacer

- **No** tocar `SEAPD.html`. Ni una línea, ni siquiera para corregir el bug del dictamen de
  calibración: ese arreglo va en un PR aparte, cuando lo autorices.
- **No** escribir cambios en el backend fuera de una rama `portal_origen === 'PAIC'`. Si al borrar
  esas ramas el archivo no queda como hoy, el PR está tocando SEAPD.
- **No** poner el correo del asesor en `CLIENTES_MAESTRO`. PORTAL manda el código de acceso a **todos**
  los correos del RFC. *(§3.3)*
- **No** ampliar `CLIENTES_MAESTRO` más allá de 22 columnas: el upsert escribe exactamente 22 y hay
  lecturas por índice `CL.*` en cinco funciones.
- **No** insertar filas en medio de `generarPerfilSheet`: se llena por coordenada fija y todo lo de
  abajo se desplaza.
- **No** reescribir los datos históricos de sucursal en el mismo PR que introduzca la normalización.
