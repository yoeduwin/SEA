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

**En el backend**, una guarda detrás de la costura. Ojo con la forma: un ternario con *fallback* a
`correo_informe` sería una trampa — si por lo que sea el registro llega sin `correo_acuse`, el acuse
se iría **al cliente**, justo lo que se quiere evitar. La regla tiene que ser dura:

```javascript
if (data.portal_origen === 'PAIC') {
  // Invariante: o va al asesor, o no va. Nunca al cliente.
  if (esCorreoValido_(data.correo_acuse)) enviarConfirmacionAsesor_(data, data.correo_acuse, …);
  // si no hay correo de acuse, NO se manda nada y se avisa en el correo interno
} else {
  enviarConfirmacionCliente(data, …);   // camino de SEAPD, intacto
}
```

Marcar el campo como `required` en el formulario **no** es suficiente: `registrarCliente` es un
endpoint público protegido con reCAPTCHA, no con validación de esquema, así que una página vieja en
caché o una petición malformada pueden llegar sin el campo. Cuando falte, el registro se guarda
igual —no se tira el trabajo del asesor ni sus archivos— pero el acuse se **omite** y el correo
interno lleva una línea que lo dice, para que Atención a Clientes lo contacte a mano.

> **Y hay que desplegar en el orden correcto: el backend va primero, no el campo.**
>
> En una versión anterior de este plan yo tenía lo contrario, con este argumento: *“el backend ignora
> las llaves que no conoce, así que publicar `correo_acuse` en el formulario es un no-op”*. El
> argumento es cierto sobre la **llave** y falso sobre la **página**, y esa distinción es justo la que
> importa. La misma publicación que agrega `correo_acuse` **redefine `correo_informe`**: hoy su
> etiqueta dice *“aquí recibirá su informe”* en segunda persona, dirigida a quien llena el formulario
> —el asesor—, y por eso las dos filas reales de la validación de campo traen el correo del
> intermediario en la columna 9. En cuanto la página nueva contrasta los dos campos, el asesor empieza
> a escribir el correo **real del cliente** en `correo_informe`. Y `enviarConfirmacionCliente` sigue
> leyendo exactamente ese campo:
>
> ```javascript
> const emailCliente = data.correo_informe;      // BACKEND_FIXES.gs:2448
> …
> GmailApp.sendEmail(emailCliente, …);            // BACKEND_FIXES.gs:2554
> ```
>
> Es decir: **cada registro de PAIC de esa ventana mandaría el acuse al cliente**, que es precisamente
> lo que este requisito prohíbe. Publicar el campo no es un no-op; cambia el significado de un campo
> que el backend vigente todavía lee.
>
> **Y no hay opción atómica.** `PAIC.html` es una página estática que hace `fetch` a una `SCRIPT_URL`
> fija (`PAIC.html:711`); el backend se publica aparte, como despliegue del proyecto de Apps Script.
> Son dos superficies y dos publicaciones: no existe el “las dos a la vez”. Hay que elegir orden.
>
> **El orden seguro es el backend primero**, y su costo es asimétrico:
>
> | Orden | Qué pasa en la ventana |
> |---|---|
> | Campo primero | El asesor pone el correo del cliente en `correo_informe` y **el acuse le llega al cliente**. Se rompe la invariante |
> | **Backend primero** | La página vieja no manda `correo_acuse`, así que el acuse **se omite**. Nadie recibe de más; el asesor recibe de menos |
>
> Se paga con un acuse ausente, no con uno mal entregado. Y ni siquiera se pierde el registro: el
> correo interno sale igual y lleva la línea para que Atención a Clientes contacte al asesor a mano.
> La ventana dura lo que tarde la segunda publicación.
>
> Rechacé la tentación de acortarla haciendo que el guard caiga a `correo_informe` cuando falte
> `correo_acuse` “solo durante la transición”: eso es exactamente el agujero cerrado más arriba, y una
> petición fabricada es indistinguible de una página en caché. Ver §6.

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
(`RECIBIDA` → `OT_RESERVADA` → `OT_GENERADA`; `DESCARTADA` como salida terminal alternativa).

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

   **Pero no puede ir sola: necesita destino propio en el mismo paso.** `guardarArchivos()` nombra
   cada archivo con la **etiqueta fija del campo** y lo escribe con `carpetaCliente.createFile(blob)`,
   sin pasar por `storeBlobSafely_()` ni `versionedFileName_()` — ese versionado existe, pero solo lo
   usa la ruta del expediente. Como Drive **sí permite nombres repetidos en una misma carpeta**, dos
   estudios del mismo RFC y sucursal dejarían dos archivos llamados idénticamente
   `L1) Hojas de campo - Laboratorio.pdf` en `01_Cliente`, **sin forma de saber cuál vino de cuál
   solicitud** — y eso ya no se reconstruye después. Por eso, en el mismo paso:
   - **Se genera el folio de la solicitud al registrar.** No con precisión de minuto: dos registros
     del mismo minuto compartirían folio, y como el folio es a la vez la llave de la fila y el nombre
     de la carpeta, eso volvería ambigua la búsqueda y —mismo RFC, sucursal y servicio— **fundiría
     cargas de dos solicitudes distintas en una sola carpeta**. El formato es
     **`SOL-aammdd-hhmmss-XXXX`**, con cuatro caracteres aleatorios, verificando que la carpeta no
     exista antes de usarlo.

     > No es un problema de concurrencia: `doPost` ya toma un `LockService.getScriptLock()` con
     > `waitLock(15000)` para toda la ruta POST, así que dos registros nunca se ejecutan entrelazados.
     > Es un problema de **precisión**. Por eso mismo, si prefieres folios legibles y consecutivos, un
     > contador secuencial también es seguro bajo ese lock; `Utilities.getUuid()` es la alternativa
     > sin pensarlo.
   - **Los archivos del estudio van a `01_Cliente/{folio}_{servicio}/`**, no a la raíz de
     `01_Cliente`. Cada registro queda con su carpeta y su identidad desde el primer día.
   - **La hoja `SOLICITUDES` y su fila van en esta MISMA entrega.** Es tentador dejarla para la Fase 3
     —así lo tenía yo— pero eso abre una ventana en la que se generan folios y carpetas **sin fila que
     los describa**, y esos registros ya no se recuperan: `correo_acuse` no se guarda en ningún lado,
     el `servicio_canonico` tampoco, y del asesor solo sobrevive el **último** en la columna 22 (§2.5).
     Las `fechas_preferidas` quedarían únicamente dentro de la hoja de perfil de ese cliente, que no es
     consultable. Sin fila, SEAOT no puede seleccionar la solicitud y no hay backfill posible.
     **El folio, la carpeta y la fila son una sola unidad de entrega: se publican juntos o no se
     publican.**
   - **Y hace falta una llave de idempotencia, porque “unidad de entrega” no es “unidad
     transaccional”.** Que las tres cosas se *publiquen* juntas no las vuelve atómicas *en tiempo de
     ejecución*: Drive y Sheets son dos servicios y no hay commit conjunto. Si la ejecución muere
     después de crear la carpeta y antes de escribir la fila, o si la respuesta se pierde después de
     escribirla, el reintento genera un folio nuevo y deja **una carpeta huérfana o una solicitud
     duplicada**.

     Y el reintento no es un escenario de laboratorio en este portal — es el camino que la propia
     página invita a tomar:

     | Evidencia en `PAIC.html` | Consecuencia |
     |---|---|
     | `setTimeout(() => controller.abort(), 120000)` sobre un endpoint que sube hasta 25 MB | A los 120 s el navegador aborta, pero **Apps Script no cancela**: el registro se completa en el servidor mientras el asesor ve un error |
     | `showError('… Por favor intente nuevamente.')`, en dos ramas del `catch` | La interfaz **le pide explícitamente** que reintente |
     | `finally { submitBtn.disabled = false; }` | El botón se rearma; reintentar es un clic |

     La solución, en la misma entrega:

     - **`submission_id`**, generado por la página **una vez por llenado** (no por envío) y mandado en
       cada intento. Es la llave estable que hoy no existe.
     - **El folio queda asociado a esa llave, no al intento.** Antes de crear nada, el backend busca
       `submission_id` en `SOLICITUDES`. Si ya existe, **reutiliza la misma fila y el mismo folio, pero
       no devuelve todavía**; si no existe, genera el folio y agrega la fila. Encontrar la fila solo
       prueba que la solicitud empezó, no que terminó.
     - **Orden de escritura: fila → carpeta → archivos.** La fila es el punto de commit porque es la
       escritura más barata y la única imprescindible; si algo muere después, el reintento la
       encuentra y continúa con el mismo folio.
     - **Cada paso se asegura antes de responder éxito**: localizar o crear la carpeta con
       `getFoldersByName(folio…)` y, después, localizar o crear cada archivo esperado con
       `getFilesByName(nombre)`. Si el primer intento murió después de la fila, el segundo crea la
       carpeta; si murió a media carga, completa únicamente los archivos faltantes.
     - **La salida exitosa ocurre al final del “ensure”.** Solo después de verificar fila + carpeta +
       todos los archivos que venían en ese payload se devuelve el folio al formulario. Así el
       navegador nunca interpreta una fila parcial como registro terminado ni borra sus archivos
       locales antes de que el servidor los haya asegurado.

     > Con esto el endpoint deja de ser “crea todo” y pasa a ser **“asegura que exista”**. La rama
     > `submission_id encontrado` no es una salida temprana: es la entrada al camino de reanudación.

     Y si el `submission_id` no llega —una copia en caché de la página anterior—, **se degrada**: folio
     aleatorio y registro normal, aceptando que un reintento desde esa copia podría duplicar. Es
     coherente con el criterio del documento: aquí el peor caso es una fila de más, revisable a mano,
     no un correo mal entregado ni un archivo perdido. Rechazar costaría el trabajo del asesor.
2. **Validar en el servidor que el registro resuelve a exactamente un servicio.** El selector maestro
   de la Fase 2 impone la regla en pantalla, pero **no es una invariante**: una copia en caché de la
   página actual —o cualquier petición directa al endpoint público— puede mandar
   `estudio_laboratorio` + `estudio_higiene` + `aplica_nom020` a la vez, o ninguno, dejando el
   `servicio_canonico` ambiguo mientras el backend acepta el registro tan campante. Con
   `portal_origen === 'PAIC'`, el backend resuelve el servicio y **rechaza** si no sale exactamente
   uno, con un mensaje que invite a recargar la página.

   > **Dónde va el rechazo, y por qué aquí sí se rechaza.** `fase1_RegistrarCliente` tiene un
   > preflight limpio —razón social, RFC y contrato de la hoja, los tres con `return` **antes** del
   > primer `createFolder`—. Ahí cabe esta validación sin dejar basura. Es la diferencia con el correo
   > de acuse (§2.3), donde elegí degradar en vez de rechazar: eso solo se sabe cuando la carpeta ya
   > existe. **Lo que se puede validar en el preflight se rechaza; lo que no, se degrada.**

3. **Rama de preservación en el upsert**, con `portal_origen === 'PAIC'` y fila existente. La regla es
   **conservar solo cuando el payload no trae el dato**, nunca de forma incondicional:
   - `REQUIERE_PIPC`: PAIC nunca lo manda, así que siempre se conserva el valor previo en vez de
     escribir `'NO'`. Sutileza: hay que distinguir **campo ausente** de **campo en `'no'`**, porque hoy
     los dos colapsan al mismo `'NO'`.
   - `ASESOR_CONSULTOR`: se conserva **solo si el payload viene vacío**. Si el registro trae asesor, se
     guarda el que trae — de lo contrario un asesor B que registra un servicio nuevo quedaría atribuido
     al asesor A del registro anterior. *(§2.1 y §2.5)*
4. **Rama de correos**: guarda dura, el acuse va a `correo_acuse` o no va. *(§2.3)*
5. **Omitir el chip de PIPC** en el correo interno cuando el origen es PAIC. *(§2.1)*

### Fase 2 — PAIC: un solo servicio y los datos del asesor

*Riesgo: bajo · Solo `PAIC.html`.* ⛔ **Depende de que la Fase 1.3–1.5 ya esté publicada**: esta fase
redefine `correo_informe`, y el backend vigente todavía manda el acuse a ese campo (§2.3).

1. **Selector maestro** con los tres grupos, excluyente, obligatorio. *(§2.2)*
2. **Derivar `aplica_nom020`** del selector, manteniendo nombre y valores.
3. **Fijar el valor canónico** de las cuatro NOM duplicadas. *(pregunta 4)*
4. **Campo `correo_acuse`** obligatorio, con la casilla “Es el mismo correo del informe” para el caso
   del intermediario corporativo. Y recalibrar la etiqueta de `correo_informe` para que quede claro
   que ahí **no** llega el acuse: es el destino del informe final y la llave de acceso del cliente al
   portal. *(§2.3, §3.3)*
5. **Campo oculto `submission_id`**, generado **una vez por llenado del formulario** —no por envío— y
   mandado en cada intento. Es la llave de idempotencia de la que cuelga todo el paso 3: sin ella, un
   reintento crea una solicitud duplicada, y la página *pide* reintentar en dos ramas de su `catch`.
   *(Fase 1.1)*
6. **Quitar el input duplicado de `calibracion_valvula`** y su contenedor muerto — **solo en PAIC**.
   *(§1.1)*
7. **Modal de confirmación** con un servicio, no una lista de bloques.

### Fase 3 — Que el registro compagine hacia abajo

*Riesgo: medio · Aquí están las decisiones de arquitectura.*

1. **Hoja `SOLICITUDES`** con las columnas de §3.5. Es lo que resuelve a la vez el servicio
   solicitado, el asesor y su correo fuera del alcance de PORTAL.
2. **Traspaso de los archivos en espera al expediente.** La subcarpeta por solicitud ya se creó en la
   Fase 1, pero **crearla no basta**: `fase3_CrearExpediente` puebla el expediente *únicamente* con
   los archivos que le manda SEAINF (`validateDriveFiles_(payload.files)` →
   `uploadValidatedFiles_()`), así que sin un paso explícito el material del asesor **se queda en la
   carpeta de espera para siempre** y el expediente no nace con nada.

   Hay que agregar un paso **explícito e idempotente** al crear el expediente: resolver la fila de
   `SOLICITUDES` por `ot_folio`, y mover sus archivos a las subcarpetas que les tocan —
   hojas de campo → `2. HDC`, croquis → `3. CROQUIS`, fotos → `4. FOTOS`. Se **mueven**, no se copian:
   el expediente es el artefacto operativo y `01_Cliente` es el perfil del cliente; dejar copias en
   los dos lados crea dos fuentes de verdad. La fila de `SOLICITUDES` queda apuntando al expediente.

   ⚠️ **Esto mete a SEAINF en el alcance por primera vez.** Hasta aquí el plan solo tocaba PAIC y la
   ruta de registro; `fase3_CrearExpediente` es código de SEAINF. No está congelado —solo SEAPD lo
   está— pero conviene saberlo antes de empezar. *(pregunta 3)*
3. **Chips del correo interno construidos desde el servicio real**, no desde dos constantes: que diga
   “NOM-025-STPS · Iluminación” en vez de “NOM-020 NO APLICA”.
4. **Imprimir el asesor y el origen** en el correo interno — dos filas en el bloque de contacto. Es el
   cambio más barato del plan y el que más contexto le da a Operaciones.
5. **SEAOT elige la solicitud, reserva su folio y después la cierra.** Al crear una OT para un RFC +
   sucursal, SEAOT lista las solicitudes en `RECIBIDA`; el operador escoge una y de ahí sale la NOM
   prellenada. **Antes de devolver un número de OT a la interfaz**, el backend reserva el folio para
   ese `sol_folio`. La reserva es parte de la operación de servidor, no un cálculo read-only en el
   navegador.

   Esto importa porque el `getSiguienteFolioOT_()` vigente **solo lee** `ORDENES_TRABAJO` y devuelve
   el siguiente valor: no lo reserva. Dos operadores que abran SEAOT antes de que alguno guarde pueden
   recibir el mismo folio aunque las llamadas estén serializadas; entre una llamada y la siguiente no
   cambió ninguna hoja.

   La reserva propuesta mantiene intacto el contrato A–Q:

   - `reservarFolioOTParaSolicitud_(sol_folio, serie)` corre bajo el mismo script lock. Si esa
     solicitud ya tiene `ot_folio`, devuelve **ese mismo** folio; recargar SEAOT no genera otro.
   - Si aún no tiene, calcula el siguiente candidato usando **dos conjuntos ocupados**: los folios ya
     presentes en `ORDENES_TRABAJO` y los `ot_folio` no vacíos de `SOLICITUDES` (incluidas reservas
     todavía no materializadas como OT). Después escribe el folio en la solicitud y cambia su estado a
     `OT_RESERVADA` **antes** de devolverlo a la interfaz.
   - Una solicitud distinta nunca puede recibir una reserva ya ocupada. Si por datos históricos se
     detecta el mismo `ot_folio` en dos solicitudes, se bloquea con conflicto explícito; no se
     reconcilia automáticamente.
   - `fase2_RegistrarOT` recibe `sol_folio` y toma como autoridad el `ot_folio` reservado en
     `SOLICITUDES`; un `data.ot_folio` distinto se rechaza en vez de usarse para decidir identidad.

   El guardado final también es **idempotente y recuperable**. Para la reserva de esa solicitud:

   | Estado encontrado | Acción |
   |---|---|
   | `OT_RESERVADA` y el folio no existe en `ORDENES_TRABAJO` | Inserta una sola fila A–Q con el folio reservado |
   | `OT_RESERVADA` y el folio ya existe | Verifica RFC + sucursal + servicio/NOM esperado; si coincide, no duplica |
   | La OT existe pero la solicitud sigue `OT_RESERVADA` | Completa únicamente el cambio a `OT_GENERADA` |
   | Solicitud ya `OT_GENERADA` | Devuelve la OT existente; no escribe nada |
   | El folio existente contradice la solicitud o está reservado por otra fila | Error de conflicto; **nunca** enlaza la segunda solicitud con la primera OT |

   Así se cubren los dos fallos parciales: si muere después de reservar y antes del `appendRow`, el
   reintento usa la misma reserva y crea la OT; si muere después del `appendRow` y antes de cerrar la
   solicitud, encuentra la OT correcta y solo marca `OT_GENERADA`.

   > El cambio clave es de identidad: **`sol_folio` identifica el intento lógico; `ot_folio` es un
   > recurso reservado para ese intento.** Ya no se deduplica “porque el texto del folio coincide”.
   > Y el generador vigente no puede seguir siendo la fuente autoritativa mientras ignore las reservas.

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

### Fase 5 — Pruebas, gates y documentación

- **Gate de regresión de SEAPD antes de publicar backend.** Primero se captura el comportamiento base
  con el payload vigente de SEAPD. Después, **cada entrega que modifique `BACKEND_FIXES.gs` debe pasar
  esa misma regresión sobre el candidato antes de publicarse**, como mínimo las entregas 1 y 3 del
  orden de §6. Una prueba posterior al despliegue no sustituye este gate: si falla, el backend no se
  publica.
- Casos E2E de PAIC: registro con archivos de estudio; reregistro que **no** borra `REQUIERE_PIPC` ni
  `ASESOR_CONSULTOR`; acuse que llega **solo** al asesor; dictamen de calibración que sí se guarda.
- Casos de recuperación: reintento después de crear solo la fila de `SOLICITUDES`; reintento después
  de subir solo parte de los archivos; ambos deben terminar la misma solicitud y conservar el folio.
- Caso de concurrencia SEAOT: dos solicitudes distintas de la misma serie reservadas antes de guardar
  deben recibir folios distintos; recargar la misma solicitud debe recuperar su misma reserva.
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

El orden **no es libre**: los pasos marcados con ⛔ abren, si se hacen fuera de lugar, una ventana en
la que el sistema hace justo lo que queremos evitar. Los pasos 1 y 2 son un **par ordenado** —backend
y luego página, nunca al revés— y el 3 es un **paquete** que no se puede partir.

La regresión de SEAPD ya no es un paso posterior del cronograma: es un **gate de publicación**. Se
captura su baseline antes del primer cambio y se vuelve a ejecutar sobre cada candidato de backend
antes de desplegarlo. En particular, **no se publica el paso 1 ni el paso 3 si ese gate no pasa**.

| # | Trabajo | Riesgo | Qué desbloquea |
|---|---|---|---|
| 1 | ⛔ **Fase 1.3–1.5 — costura `portal_origen`** (preservación, guard duro de correos, chip de PIPC) | bajo | **Gate SEAPD previo obligatorio.** Después va primero: el guard tiene que estar vivo antes de que la página redefina `correo_informe`; si no, el acuse se le va **al cliente** (§2.3) |
| 2 | ⛔ **Fase 2 — campos de PAIC** (`correo_acuse`, selector de un servicio, datos del asesor) | bajo | Cierra la ventana del paso 1: el asesor recupera su acuse, ahora por el campo correcto |
| 3 | ⛔ **Fase 1.1 + 1.2 + 3.1, en una sola entrega**: archivos, folio ligado a `submission_id`, reanudación idempotente, carpeta por solicitud, hoja `SOLICITUDES` y validación de un solo servicio | medio | **Gate SEAPD previo obligatorio.** Los estudios dejan de perderse y quedan atribuibles, descritos y reintentables sin duplicar |
| 4 | Fase 3.5 — selección de solicitud en SEAOT, **reserva de folio por `sol_folio`** y cierre idempotente de la OT | medio | Embudo solicitud → OT auditable, recuperable y sin cruces entre operadores |
| 5 | Fase 3.2 — traspaso al expediente (toca SEAINF) | medio | Ahora sí: el expediente nace con su material |
| 6 | Fase 3.3–3.4 — correos que dicen la verdad | bajo | Operaciones ve el servicio y el asesor |
| 7 | Fase 4 — deuda propia | bajo | Calidad del portal |
| 8 | Fase 5 — E2E de PAIC, recuperación/concurrencia y documentación | bajo | Evidencia final; **no reemplaza los gates previos** de las entregas de backend |

> **Nota sobre cualquier entrega posterior de backend.** Aunque Codex señaló específicamente los
> pasos 1 y 3, el criterio queda generalizado: si los pasos 4, 5 o 6 terminan modificando el backend
> compartido, también pasan la regresión SEAPD **antes** de publicarse.

> **Nota honesta sobre el paso 3.** En la versión anterior de este plan, “que los estudios dejen de
> perderse” era la victoria barata del principio. Ya no lo es: para que esos archivos sirvan de algo
> tienen que llegar con folio, carpeta y fila, y eso arrastra la hoja `SOLICITUDES` al mismo paquete.
> Sube de *bajo* a *medio* y crece en tamaño. Es el precio de que el arreglo sea real y no cosmético.

> **Nota honesta sobre los pasos 1 y 2.** Este par estuvo invertido en la versión anterior, con el
> argumento de que publicar el campo era un no-op. No lo es: la misma publicación redefine
> `correo_informe`, que el backend vigente todavía lee para mandar el acuse (§2.3). Como `PAIC.html`
> y el backend viven en superficies distintas, **no existe el despliegue simultáneo**; hay que elegir,
> y la elección correcta es la que falla omitiendo el acuse en vez de la que falla mandándoselo al
> cliente.

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
