# Propuesta — Mejoras de Trazabilidad SEAOT

**Estado:** Propuesta para revisión (NO implementada).
**Flujo previsto:** Eduwin revisa esta primera propuesta → ajustes → se abre PR con Codex para implementar.
**Fecha:** 2026-08-26

> **Decisiones tomadas por Eduwin (2026-08-26) — esta versión ya las incorpora:**
> 1. **Sin extender el esquema** de `ORDENES_TRABAJO` (contrato A–Q intacto). Solo **validar**, no persistir fechas nuevas.
> 2. Las validaciones cronológicas y de agenda son **advertencia** (avisan y dejan continuar), **no bloqueo duro**.
> 3. **Sin crear hojas nuevas.** El control de "salida abierta" de equipo se hace **dentro del propio formato** (sin backend), no con una hoja `MOVIMIENTOS_EQUIPO`.
> 4. **Prioridad inmediata:** que al generar los formatos se **reescriba lo menos posible** (máximo prellenado desde la OT). Ver **§0-bis**.

---

## 0. Contexto real del sistema (lo que hoy existe)

Antes de proponer, esto es lo que hay hoy en el código, para que las propuestas encajen sin romper contratos:

| Pieza | Dónde vive | Notas relevantes |
|---|---|---|
| Registro de OT (backend) | `BACKEND_FIXES.gs → fase2_RegistrarOT(data)` (línea ~900) | Escribe en `ORDENES_TRABAJO` con **contrato inmutable A–Q** (17 columnas). Rechaza si la hoja tiene < 17 columnas. |
| Formulario de OT (frontend) | `SEAOT.html` | Campos: `wo_issueDate` (Fecha de Emisión), `wo_contractReviewDate` (Fecha de revisión de contrato), y tabla "Desarrollo del Trabajo" con una **FECHA por servicio**. `fecha_visita = dates[0]` (la primera fecha de la tabla). |
| Mapa de columnas OT | `CONFIG.COLUMNS.ORDENES` (`CO`) | `FECHA=A`, `OT=B`, `PERSONAL=H(7)`, `FECHA_VISITA=I(8)`, `ESTATUS_EXTERNO=L(11)`. |
| Estatus válidos OT | `ESTATUS_EXTERNO_VALIDOS_` / `..._TERMINALES_` | Terminales = `FINALIZADO`, `CANCELADO`. Activas = el resto. |
| Formato de salida de equipos | `registro-equipos.html` | **Es un formato de impresión/PDF, sin backend.** "Entregó/Recibió" son solo líneas de firma a mano. `fechaSalida` = fecha de salida. Prellenado desde SEAOT vía `sessionStorage` (`applyOtHandoff`). |
| Auditoría | `registrarAuditoria_(usuario, accion, ot, campo, antes, despues)` | Bitácora en hoja `AUDITORIA`. |
| Trazabilidad de lectura | `TRAZ/` | Solo lectura, no escribe nada. |

**Dos hechos que condicionan todo:**

1. **`ORDENES_TRABAJO` tiene contrato "inmutable A–Q".** Hoy **NO** guarda ni "Fecha de Emisión" ni "Fecha de Revisión de Contrato" como columnas (solo `FECHA` timestamp A, `FECHA_VISITA` I y `FECHA_ENTREGA` J). Para trazabilidad cronológica persistida hay que decidir entre **validar sin persistir** o **extender el contrato de forma controlada**.
2. **No existe ninguna hoja de equipos.** El control de "salida abierta" por número de inventario requiere crear una estructura de datos nueva (hoy no hay dónde consultarlo).

---

## 0-bis. PRIORIDAD INMEDIATA — Reescribir lo menos posible al generar los formatos

El handoff SEAOT→formato ya viaja con este payload (mismo-origen, sin PII en URL):
`ot, serie, cliente, sucursal, rfc, direccion, contacto, telefono, correo, emision, servicios, personal, fecha`.

Lo que **hoy** se prellena vs. lo que **aún se reescribe a mano**:

| Formato | Ya prellenado (automático) | Se reescribe a mano (evitable) |
|---|---|---|
| `registro-equipos.html` | `#destino`(OT), `#motivoSalida`, `#solicitante`, `#fechaSalida`, `#area`, equipos por NOM | **Entregó / Recibió** (salida y entrada) — hoy son líneas de firma en blanco |
| `supervision-gabinete.html` | `#sg_ot`, `#sg_cliente` | **`#sg_fecha`** (hoy vacía, se teclea), `#sg_folio` (folio propio del revisor) |

**Cambios mínimos propuestos (aditivos, sin tocar backend ni esquema):**

**A) `supervision-gabinete.html` — prellenar la fecha (1 línea en su `applyOtHandoff`):**
```js
// dentro del applyOtHandoff ya existente, junto a sg_ot / sg_cliente:
if (data.fecha)   { var f = document.getElementById('sg_fecha');   if (f && !f.value) f.value = fechaADMY_(data.fecha); }
// (si se prefiere la emisión de la OT en vez de la fecha de servicio: usar data.emision)
```
**Nota (hallazgo P2 de Codex):** `#sg_fecha` es un campo de **texto** que se imprime tal cual, y `data.fecha` viene en ISO (`YYYY-MM-DD`). Hay que convertirla con **`fechaADMY_(data.fecha)`** para que la supervisión salga en `dd/mm/yyyy` y no en ISO. `#sg_folio` se deja manual (consecutivo propio, no viene de la OT).

> **⚠️ Dependencia obligatoria (hallazgo P2 de Codex):** `fechaADMY_` (definida en §3) **NO existe hoy** en `supervision-gabinete.html`. Como §0-bis-A la invoca, **su definición debe incluirse en este mismo cambio** (pegar el helper de §3 en el `<script>` del formato). No es opcional ni posterior: sin ella, todo handoff con `data.fecha` lanzaría `ReferenceError`. Por eso en §7 el helper de §3 pasa a ser **prerrequisito**, no un paso final opcional.

> **⚠️ Captura manual de `#sg_fecha` (hallazgo P2 de Codex):** `#sg_fecha` es `type="text"` y **sigue editable** tras el prellenado. El handoff solo normaliza el valor recibido; una captura manual como `08/27/2026` o `31/02/2026` se imprimiría sin validar. Dos opciones (elige una en el PR):
> - **Simple:** cambiar `#sg_fecha` a `<input type="date">` y mostrar/imprimir con `fechaADMY_` — así el navegador impide formatos inválidos.
> - **Mínimo cambio de marcado:** dejarlo como texto pero validar al salir del campo:
> ```js
> document.getElementById('sg_fecha').addEventListener('blur', function(){
>   const v = this.value.trim(); if (!v) return;
>   const iso = normalizarFechaISO_(v);              // acepta dd/mm/yyyy o ISO; nunca mm/dd
>   if (!iso) { alert('⚠️ Fecha inválida en Supervisión. Usa día/mes/año.'); return; }
>   this.value = fechaADMY_(iso);                     // re-formatea a dd/mm/yyyy
> });
> ```
> Esto también requiere que `normalizarFechaISO_` (§3) esté definida en el formato.

**B) `registro-equipos.html` — autollenar SOLO el nombre del técnico en sus dos campos.**

El bloque de firmas tiene 4 espacios. La regla, ya confirmada por Eduwin:

| Momento | Campo | Quién | ¿Autollenar? |
|---|---|---|---|
| SALIDA | 1. Entregó | Almacén (presta el equipo) | **No** — se queda como línea de firma tal cual (sin selector) |
| SALIDA | 2. **Recibió** | **Técnico** = `#solicitante` | **Sí** ← handoff |
| ENTRADA (regreso) | 3. **Entregó** | **Técnico** = `#solicitante` (el mismo) | **Sí** ← handoff |
| ENTRADA (regreso) | 4. Recibió | Almacén (recibe de vuelta) | **No** — línea de firma tal cual |

Es decir: **los dos campos del técnico se autollenan con el mismo nombre** (el `solicitante` de la OT: recibe a la salida, entrega al regreso). Los dos de almacén **no cambian** — siguen como el texto de firma actual, sin `<select>`.

Solo esos dos espacios pasan a tener un pequeño `<span>` con el nombre (no selector), antepuesto a la firma manuscrita:

```html
<!-- SALIDA: 2. Recibió -->
<div class="signature-line"><strong>2. Recibió:</strong>
  <span id="salidaRecibioNombre" class="nombre-auto"></span> — Firma y Fecha</div>
<!-- ENTRADA: 3. Entregó -->
<div class="signature-line"><strong>3. Entregó:</strong>
  <span id="entradaEntregoNombre" class="nombre-auto"></span> — Firma y Fecha</div>
<!-- Los espacios 1 (Entregó salida) y 4 (Recibió entrada) NO se tocan. -->
```
La sincronización se hace con **una sola función** que se llama desde el handoff, desde el `change` de `#solicitante` **y** desde los reinicios del formulario. **Es obligatoria** (hallazgo P2 de Codex): si no, cambiar `#solicitante` a mano dejaría los `<span>` con el nombre anterior y el impreso identificaría a una persona como solicitante y a otra en las firmas.

```js
// Fuente única de verdad: refleja #solicitante en los dos span del técnico.
function sincronizarNombreTecnico() {
  var tecnico = (document.getElementById('solicitante') || {}).value || '';
  var r = document.getElementById('salidaRecibioNombre');  if (r) r.textContent = tecnico;
  var e = document.getElementById('entradaEntregoNombre'); if (e) e.textContent = tecnico;
}

// 1) tras setSolicitante(data.personal) en applyOtHandoff():
sincronizarNombreTecnico();

// 2) cambio manual del solicitante (OBLIGATORIO, no opcional):
document.getElementById('solicitante').addEventListener('change', sincronizarNombreTecnico);

// 3) reinicios programáticos del formulario (limpiarFormulario y cualquier reset):
//    llamar sincronizarNombreTecnico() al final para que los span queden vacíos también.
```
El punto 3 cubre el reseteo: `limpiarFormulario()` ya limpia los `select`, pero debe invocar `sincronizarNombreTecnico()` para vaciar también los `<span>`.

Resultado: al abrir el formato desde la OT, el nombre del técnico ya aparece impreso en "Recibió" (salida) y "Entregó" (entrada); si se cambia el solicitante, ambos se actualizan solos; y un reinicio los vacía. Los dos campos de almacén quedan idénticos a hoy.

> Este bloque **0-bis es lo que puede entrar primero al PR** porque es el de menor riesgo (solo frontend, aditivo) y resuelve directo tu pedido de "reescribir lo menos posible".

---

## 1. Trazabilidad cronológica: Revisión de Contrato ≤ Emisión de OT ≤ Fecha de Visita

**Regla de negocio:**
```
Fecha Revisión de Contrato  ≤  Fecha de Emisión de OT  ≤  Fecha de Visita (mínima de la tabla de servicios)
```
Y además: **toda** fecha de la tabla de servicios ≥ Fecha de Emisión (no solo la primera).

### 1.a Validación en frontend (SEAOT.html) — barrera inmediata, sin tocar esquema

Antes de armar el `payload` en el submit (`SEAOT.html`, junto al armado de `services/personnel/dates`):

```js
// Devuelve 'YYYY-MM-DD' o '' — los <input type="date"> ya entregan ISO.
function _d(id){ return (document.getElementById(id)?.value || '').trim(); }

// Acumula TODAS las infracciones (hallazgo P2 de Codex): no corta en la primera,
// y comprueba cada relación solo cuando sus dos operandos están disponibles.
function validarCronologiaOT(revision, emision, fechasVisita){
  // fechasVisita = array de 'YYYY-MM-DD' de la tabla de servicios (todas las filas)
  const errs = [];
  const visitas = (fechasVisita || []).filter(Boolean).sort();  // ISO ordena lexicográfico = cronológico
  const minVisita = visitas[0] || '';

  if (!emision)  errs.push('Falta la Fecha de Emisión de la OT.');
  if (!revision) errs.push('Falta la Fecha de revisión de contrato.');
  if (!visitas.length) errs.push('Falta al menos una fecha de visita en la tabla de servicios.');

  // cada relación se evalúa aunque falten otras fechas
  if (revision && emision && revision > emision)
    errs.push(`La revisión de contrato (${revision}) no puede ser posterior a la emisión (${emision}).`);
  if (emision && minVisita && emision > minVisita)
    errs.push(`La emisión (${emision}) no puede ser posterior a la visita más temprana (${minVisita}).`);

  return errs.join('\n');  // '' si todo OK; varias líneas si hay varias infracciones
}
```
El aviso (`confirmarCronologiaOT`, más abajo) muestra el `join('\n')` completo, de modo que si coexisten "revisión posterior a emisión" **y** "emisión posterior a visita", el usuario ve ambas y no solo la primera.

**Enganche en TODAS las salidas de la OT (hallazgo P2 de Codex), no solo en el registro.** En `SEAOT.html` hay botones independientes que producen salida sin pasar por `registerToSheets`: `printWorkOrder()` (línea 765), `sendWorkOrder()` (787) y `openFormatosModal()` (994). Para no dejar huecos, se **centraliza** la validación en una función que recolecta las fechas del DOM y se invoca al inicio de las cuatro acciones:

```js
// recolecta fechas de la tabla de servicios y CUENTA las filas con servicio pero sin fecha
// (hallazgo P2 de Codex): una fila fechada en blanco no debe descartarse en silencio.
function _visitasDOM(){
  const dates = [];
  let sinFecha = 0;
  document.querySelectorAll('#workTableBody tr').forEach(row => {
    const f = row.querySelectorAll('input, select, textarea');
    if (f.length >= 3 && f[0].value){            // hay servicio en la fila
      const d = (f[2].value || '').trim();
      if (d) dates.push(d); else sinFecha++;     // servicio sin fecha = infracción a reportar
    }
  });
  return { dates, sinFecha };
}
// aviso centralizado (modo advertencia, decisión #2). Devuelve false si el usuario cancela.
function confirmarCronologiaOT(){
  const { dates, sinFecha } = _visitasDOM();
  let err = validarCronologiaOT(_d('wo_contractReviewDate'), _d('wo_issueDate'), dates);
  if (sinFecha > 0) err = (err ? err + '\n' : '') +
    `${sinFecha} servicio(s) sin fecha de visita: no se puede comprobar su cronología.`;
  return !err || confirm('⚠️ ' + err + '\n\n¿Continuar de todos modos?');
}

// invocar al inicio de CADA salida:
async function registerToSheets(e){ if (!confirmarCronologiaOT()) return; /* … */ }
function printWorkOrder(){        if (!confirmarCronologiaOT()) return; /* … */ }
function sendWorkOrder(){         if (!confirmarCronologiaOT()) return; /* … */ }
function openFormatosModal(){     if (!confirmarCronologiaOT()) return; /* … */ }
```
Así, registrar, imprimir/PDF, enviar por correo y abrir formatos asociados pasan todos por el mismo aviso; una revisión posterior a la emisión (o emisión posterior a la visita) avisa en cualquier ruta.

### 1.a-bis Propagar la fecha MÍNIMA, no `dates[0]` (hallazgo P2 de Codex)

La regla define "Fecha de Visita" como la **mínima** de la tabla. Pero hoy el código envía la **primera fila**:
- `registerToSheets`: `fecha_visita: dates[0]`
- `collectOtContext` (handoff a formatos): `fecha: fechas[0]`

Si las filas no están en orden cronológico, la hoja y los formatos (incluida la validación de equipos §2.a) recibirían una visita distinta a la que valida §1. **Corrección:** usar la mínima calculada en ambos destinos:
```js
const minVisita = dates.filter(Boolean).sort()[0] || '';
// payload de registrarOT:
fecha_visita: minVisita,
// collectOtContext:
fecha: (fechas.filter(Boolean).sort()[0] || ''),
```
(Alternativa: cambiar explícitamente el contrato de la regla a "primera fila". Recomiendo la mínima, que es lo que valida §1.)

### 1.b Sin persistir fechas nuevas (decisión #1)

La validación es **solo en frontend** (§1.a). **No** se extiende el contrato A–Q ni se agregan columnas: las fechas de Emisión y Revisión de Contrato siguen viviendo únicamente en el formato/PDF de la OT, como hoy. La trazabilidad cronológica se **garantiza en el momento de capturar la OT** (el aviso obliga a confirmar si el orden está mal), pero no se registra como dato consultable en la hoja.

> Nota: si más adelante quieres que TRAZ muestre estas fechas, la vía es extender el contrato al final (A–S) — queda documentada como mejora futura, fuera de este PR.

---

## 2. Equipos: fecha de solicitud ≤ fecha de visita + nombres en "Recibió" / "Entregó"

### 2.a Fecha de solicitud ≤ Fecha de visita
En `registro-equipos.html`, `fechaSalida` es la fecha de solicitud/salida y la fecha de visita llega en el handoff (`data.fecha`). Propuesta: guardar la fecha de visita recibida y validarla al imprimir/guardar:

```js
// en applyOtHandoff(): conservar la fecha de visita de la OT
if (data.fecha){ window.__fechaVisitaOT = data.fecha; }

// validación al imprimir/exportar — modo advertencia (decisión #2)
function validarFechaSolicitud(){
  const fs = (document.getElementById('fechaSalida').value || '').trim();
  const fv = window.__fechaVisitaOT;
  const avisos = [];
  // Fecha de salida ausente = infracción confirmable (hallazgo P2 de Codex):
  // sin ella no se puede demostrar solicitud ≤ visita.
  if (!fs) avisos.push('Falta la fecha de solicitud/salida.');
  if (fs && fv && fs > fv) avisos.push(`La fecha de solicitud (${fs}) es posterior a la fecha de visita (${fv}).`);
  return avisos.length === 0 ||
    confirm('⚠️ ' + avisos.join('\n') + '\n\n¿Continuar de todos modos?');
}
```

**Importante (hallazgo P2 de Codex):** definir la función no basta — hay que **invocarla al inicio de las DOS acciones de salida** (`imprimir()` y `exportarPDF()`) y **abortar** si el usuario declina:
```js
function imprimir(){
  if (!validarFechaSolicitud()) return;   // ← nuevo, primera línea
  /* … resto igual … */
}
function exportarPDF(){
  if (!validarFechaSolicitud()) return;   // ← nuevo, primera línea
  /* … resto igual … */
}
```
Sin estas dos guardas, cambiar `fechaSalida` a una fecha posterior a la visita seguiría generando el PDF/impresión sin aviso.

**Limpiar el estado al reiniciar (hallazgo P2 de Codex):** `window.__fechaVisitaOT` sobrevive a `limpiarFormulario()`. Si se reutiliza la misma pestaña para capturar **otra salida manual**, la comparación se haría contra la visita de la OT anterior → aviso falso. Por eso `limpiarFormulario()` debe **borrar la fecha de visita interna**:
```js
function limpiarFormulario(){
  /* … reset de campos … */
  window.__fechaVisitaOT = '';          // ← nuevo: olvida la visita de la OT previa
  sincronizarNombreTecnico();           // (del §0-bis-B: vacía también los span de nombre)
}
```

### 2.b Nombres en "Recibió" (salida) y "Entregó" (entrada)
Ver **§0-bis, bloque B** (con la tabla de los 4 campos ya confirmada). Resumen: se autollena **solo el nombre del técnico** (`#solicitante`) en sus dos espacios —Recibió a la salida y Entregó al regreso, es la misma persona— con un `<span>` impreso, no un selector. Los dos campos de almacén se quedan como están.

---

## 3. Validación de formato de fecha (día/mes/año) para no romper backend

**Diagnóstico:** los campos `<input type="date">` **ya entregan `YYYY-MM-DD` (ISO)**, que es no ambiguo y seguro para backend. El riesgo real aparece si:
- se muestra/captura como texto `dd/mm/yyyy` y Google Sheets lo **auto-interpreta como mm/dd** (US), o
- se agregan campos de texto libre para fechas.

**Propuesta — helpers de frontend (sin backend) + display consistente.** `normalizarFechaISO_` solo se usa como salvaguarda si aparece algún campo de texto libre de fecha; `fechaADMY_` es para mostrar en documentos:

```js
/**
 * Normaliza cualquier entrada de fecha a ISO 'YYYY-MM-DD' o '' si es inválida.
 * Acepta: 'YYYY-MM-DD', 'DD/MM/YYYY', 'DD-MM-YYYY'. NUNCA interpreta MM/DD.
 */
function normalizarFechaISO_(v){
  if (!v) return '';
  var s = String(v).trim();
  var m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/))) return _validaISO_(m[1], m[2], m[3]);
  if ((m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/))) {           // DD/MM/YYYY
    return _validaISO_(m[3], ('0'+m[2]).slice(-2), ('0'+m[1]).slice(-2));
  }
  return '';
}
function _validaISO_(y, mo, d){
  var dt = new Date(+y, +mo - 1, +d);
  if (dt.getFullYear() != +y || dt.getMonth() != +mo - 1 || dt.getDate() != +d) return ''; // fecha inexistente (p.ej. 31/02)
  return y + '-' + mo + '-' + d;
}
/** Para mostrar en documentos/PDF en formato humano dd/mm/yyyy. */
function fechaADMY_(iso){
  var m = String(iso||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? (m[3] + '/' + m[2] + '/' + m[1]) : String(iso||'');
}
```

**Regla operativa recomendada (sin tocar backend, decisión #1):**
- **La mayoría de los campos de fecha ya son `type="date"`** (SEAOT y `fechaSalida` en equipos) → el navegador entrega ISO `YYYY-MM-DD`, no ambiguo. Ahí el pedido de "día/mes/año sin romper backend" ya está cubierto: la fecha nunca sale como `mm/dd`.
- **Excepción confirmada: `#sg_fecha` en `supervision-gabinete.html` es `type="text"`** (hallazgo P2 de Codex). No basta con formatear el valor del handoff: hay que **validar la captura manual** con `normalizarFechaISO_` (o cambiar el campo a `type="date"`), según §0-bis-A. Sin eso, un `08/27/2026` o `31/02/2026` tecleado se imprime sin control.
- **Solo se convierte a `dd/mm/yyyy` para mostrar** (PDF, formatos) con `fechaADMY_`.
- Regla general: **todo campo de fecha de texto libre** pasa por `normalizarFechaISO_` en cliente; si devuelve `''`, se avisa. No se agrega validación en backend (decisión #1).

---

## 4. Antichoque de agenda: mismo "Personal Asignado" ya tiene OT activa en la misma "Fecha de Visita" — **FUERA DE ALCANCE (decisión: cero backend)**

> **Decisión de Eduwin: cero backend.** Este punto se **descarta del alcance actual**. A continuación queda el porqué, para no volver a proponerlo sin persistencia.

Para detectar el choque hay que **leer todas las OT activas** (personal + fecha de visita) desde el servidor. El único endpoint de lectura existente que devuelve eso es `getOrdenes`, y **no sirve** (confirmado en el código, hallazgo P1 de Codex):

1. `getOrdenes` está autorizado bajo el módulo **`SEAINF`** (`ACTION_MODULE.getOrdenes = 'SEAINF'`). Un operador con acceso a **SEAOT pero no a SEAINF** recibe una respuesta sin `data` → no detectaría ningún choque.
2. `getOrdenesSafe_` **elimina toda OT que ya tenga registro en `INFORMES`** aunque su estatus siga activo → una OT con expediente creado pero visita pendiente **quedaría fuera** del chequeo.

Una verificación confiable **requeriría** un endpoint de lectura propio de SEAOT (una función `getAgendaActiva_` + su fila en `ACTION_MODULE`) — es decir, **backend**. Como la decisión es **cero backend**, esta validación **no se implementa ahora**.

**Alternativa 100% frontend (limitada, opcional):** dentro de la misma sesión de SEAOT no hay forma de conocer OT de otras sesiones/días sin leer el servidor, así que **no hay** un sustituto real en cliente. Se deja como **mejora futura**, condicionada a que en algún momento se autorice esa pequeña lectura de backend.

---

## 5. Control de "salida abierta" de equipo por número de inventario — **sin hoja nueva** (decisión #3)

**Realidad:** `registro-equipos.html` no persiste nada y **no se creará** una hoja. Por lo tanto, **no es posible** verificar de forma confiable si `EA-SO11-01` "ya salió en otro documento/otro día", porque no hay dónde consultar el historial. Sería deshonesto prometerlo sin persistencia.

**Lo que SÍ se puede hacer hoy, dentro del propio formato (frontend, sin backend):**

### 5.a Impedir/avisar equipo duplicado dentro del mismo formato
El caso más común de error humano —seleccionar dos veces el mismo inventario en la misma salida— sí se detecta en la tabla actual. Ya existe la base `existeEquipo(clave)` (recorre los `.equipo-select`). Se refuerza al cambiar un equipo:

La función real es `alCambiarEquipo(selectElement)` (parámetro `selectElement`, verificado en `registro-equipos.html:675`). La comprobación va **al inicio**, justo tras leer `selectElement.value`:

```js
function inventarioYaEnTabla(clave, selfSelect){
  return Array.from(document.querySelectorAll('#tablaEquipos tbody .equipo-select'))
    .some(s => s !== selfSelect && s.value === clave);
}

function alCambiarEquipo(selectElement) {
  const clave = selectElement.value;
  if (clave && inventarioYaEnTabla(clave, selectElement)){        // ← nuevo, antes de todo
    alert(`⚠️ El equipo ${clave} ya está en esta salida. Revisa que no lo estés registrando dos veces.`);
    selectElement.value = '';              // rechaza: limpia la selección duplicada
    actualizarInventario(selectElement);   // limpia también el nº de inventario de esa fila
    return;                                // corta: no continúa con autollenado ni grupos
  }
  actualizarInventario(selectElement);     // (resto original sin cambios)
  /* … grupos … */
}
```
**Notas (hallazgos de Codex):**
- **P1:** el parámetro real es `selectElement`, **no** `select` — usar `select` lanzaría `ReferenceError` y rompería también el autollenado. Corregido arriba.
- **P2:** el aviso por sí solo no impide el duplicado; hay que **limpiar la selección y cortar** (`return`) para que la antiduplicación sea real.

### 5.b "Función base" documentada para el futuro (gancho, no se implementa ahora)
Se deja escrita la firma de la función que haría la verificación real *si algún día se autoriza persistencia*, para que quede clara la estructura. **No se incluye en este PR** (requeriría una hoja):

```js
/**
 * [FUTURO — requiere persistencia, hoy no disponible por decisión #3]
 * ¿El equipo `inventario` tiene una salida abierta que se traslape con `fechaISO`?
 * Contrato de datos sugerido (cuando exista): {inventario, fecha_salida, fecha_retorno_real}.
 * Regla: hay salida abierta si existe un registro del inventario sin fecha_retorno_real
 * cuyo rango de salida cubre `fechaISO`.
 */
function tieneSalidaAbierta_(inventario, fechaISO){ /* pendiente: sin fuente de datos aún */ }
```

**Resumen honesto:** con las reglas actuales (sin hoja), el control de inventario se limita a **evitar duplicados dentro del mismo formato**. El bloqueo por "salida abierta entre documentos/fechas" queda **fuera de alcance** hasta que se autorice una fuente de datos persistente.

---

## 6. Resumen de decisiones (ya aplicadas en esta versión)

| # | Tema | Decisión de Eduwin |
|---|---|---|
| 1 | Emisión + Revisión de Contrato | **Solo validar en frontend, sin extender esquema** |
| 2 | Cronología (y agenda) | **Advertencia** (avisa y deja continuar) |
| 3 | Control de equipos por inventario | **Sin crear hoja** → solo antiduplicado dentro del formato |
| 4 | Prioridad inmediata | **Reescribir lo menos posible** al generar formatos (§0-bis) |
| 5 | **Cero backend** | **§4 (choque de agenda) queda FUERA de alcance** — requeriría un endpoint SEAOT nuevo. Nada de lo demás toca backend. |

---

## 7. Orden sugerido de implementación para el PR (menor a mayor riesgo)

0. **§3 — Helper de fecha `fechaADMY_` (PRERREQUISITO):** debe existir **antes** que §0-bis, porque §0-bis-A lo usa para `#sg_fecha`. Pegar el helper en el `<script>` de cada formato que lo invoque (`supervision-gabinete.html`; y `registro-equipos.html` si se muestra alguna fecha en `dd/mm/yyyy`). **No es opcional.**
1. **§0-bis — Prellenado de formatos** (el pedido inmediato): fecha (con `fechaADMY_`) en supervisión + nombre del técnico en Recibió/Entregó de equipos. Solo frontend, aditivo.
2. **§1.a — Validación cronológica** Revisión ≤ Emisión ≤ Visita, como aviso en SEAOT (incluye §1.a-bis: propagar la fecha mínima).
3. **§2.a — fecha solicitud ≤ visita** como aviso en registro-equipos (incluye limpiar `window.__fechaVisitaOT` en el reinicio).
4. **§5.a — Antiduplicado de inventario** dentro del formato.

> **§4 (choque de agenda) queda fuera** por la decisión de cero backend (mejora futura).

Todo es **frontend y aditivo**: no reordena el contrato A–Q, **no toca backend**, no crea hojas. Cada punto puede ir como commit independiente dentro del mismo PR.
