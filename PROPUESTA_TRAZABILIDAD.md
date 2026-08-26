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
if (data.fecha)   { var f = document.getElementById('sg_fecha');   if (f && !f.value) f.value = data.fecha; }
// (si se prefiere la emisión de la OT en vez de la fecha de servicio: usar data.emision)
```
`#sg_folio` se deja manual (es el consecutivo de supervisión, no viene de la OT).

**B) `registro-equipos.html` — convertir "Entregó/Recibió" en campos prellenables.**
Hoy en el bloque de firmas son texto fijo (`Nombre, Firma y Fecha`). Se propone dejar la **firma manuscrita** igual, pero anteponer un `<select>`/`<input>` de **nombre** que se autollena desde el handoff, reutilizando el catálogo de personal que ya usa `#solicitante`:

```html
<!-- SALIDA DE EQUIPOS -->
<div class="signature-line"><strong>1. Entregó:</strong>
  <select id="salidaEntrego">…mismas opciones que #solicitante…</select> — Firma y Fecha</div>
<div class="signature-line"><strong>2. Recibió:</strong>
  <select id="salidaRecibio">…</select> — Firma y Fecha</div>
```
```js
// en applyOtHandoff() de registro-equipos.html, tras setSolicitante(data.personal):
var rec = document.getElementById('salidaRecibio');
if (rec && !rec.value) rec.value = document.getElementById('solicitante').value; // Recibió = quien sale a campo
// 'Entregó' (almacén) se deja para elegir, o se fija un responsable por defecto si lo defines.
```
Resultado: al abrir el formato desde la OT, **Recibió** ya trae el nombre del técnico asignado y solo queda firmar. "Entregó" (almacén) es el único nombre a elegir.

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

function validarCronologiaOT(revision, emision, fechasVisita){
  // fechasVisita = array de 'YYYY-MM-DD' de la tabla de servicios (todas las filas)
  if (!emision)  return 'Falta la Fecha de Emisión de la OT.';
  if (!revision) return 'Falta la Fecha de revisión de contrato.';
  const visitas = fechasVisita.filter(Boolean).sort();      // ISO ordena lexicográfico = cronológico
  if (!visitas.length) return 'Falta al menos una fecha de visita en la tabla de servicios.';
  const minVisita = visitas[0];

  if (revision > emision)   return `La revisión de contrato (${revision}) no puede ser posterior a la emisión (${emision}).`;
  if (emision  > minVisita) return `La emisión (${emision}) no puede ser posterior a la primera visita (${minVisita}).`;
  return ''; // OK
}
```

Uso en el submit — **modo advertencia** (avisa y deja continuar, decisión #2):
```js
const err = validarCronologiaOT(_d('wo_contractReviewDate'), _d('wo_issueDate'), dates);
if (err && !confirm('⚠️ ' + err + '\n\n¿Registrar la OT de todos modos?')) return;
```

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
  const fs = document.getElementById('fechaSalida').value;
  const fv = window.__fechaVisitaOT;
  if (fs && fv && fs > fv){
    return confirm(`⚠️ La fecha de solicitud (${fs}) es posterior a la fecha de visita (${fv}).\n\n¿Continuar de todos modos?`);
  }
  return true;
}
```

### 2.b Nombres estructurados en "Entregó" (ingresó equipo) y "Recibió" (recibir equipo)
Ver **§0-bis, bloque B** — es la misma mejora y es la de mayor prioridad (reduce reescritura). Resumen: dejar la firma manuscrita igual, anteponer un `<select>` de nombre con el catálogo de `#solicitante`, y autollenar `salidaRecibio = solicitante` desde el handoff.

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
- **Mantener `type="date"` en todos los campos de fecha** (SEAOT y formatos ya lo usan) → el navegador entrega ISO `YYYY-MM-DD`, no ambiguo. Esto **ya cubre** el pedido de "formato día/mes/año sin romper backend": la fecha nunca sale como `mm/dd` ambiguo.
- **Solo se convierte a `dd/mm/yyyy` para mostrar** (PDF, formatos) con `fechaADMY_` (helper de frontend).
- Si en algún formato hubiera un campo de **texto libre** de fecha, se le pasa `normalizarFechaISO_` en cliente antes de usarla y, si devuelve `''`, se avisa. No se agrega validación en el backend (queda como salvaguarda futura si algún día se persisten fechas).

---

## 4. Antichoque de agenda: mismo "Personal Asignado" ya tiene OT activa en la misma "Fecha de Visita"

**Enfoque sin tocar backend (decisión #1 y #2):** la verificación se hace en **frontend**, antes de enviar el registro, usando la acción de **lectura que ya existe** (`getOrdenes`, que devuelve `personal`, `fecha_visita` y `estatus_externo` por OT). Es **advertencia**: avisa y deja continuar.

```js
// En SEAOT.html, antes de armar el payload de registrarOT:
async function hayChoqueAgenda(personalAsignado, fechaVisitaISO, otExcluir){
  if (!fechaVisitaISO) return null;
  const resp = await SEAAuth.wrapFetch(SHEETS_SCRIPT_URL + '?action=getOrdenes', { method:'GET' });
  const { data: ordenes = [] } = await resp.json();
  const TERMINALES = ['FINALIZADO','CANCELADO'];
  const norm = s => String(s||'').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/\s+/g,' ').trim();
  const solicitados = personalAsignado.split(',').map(norm).filter(Boolean);

  for (const o of ordenes){
    if (TERMINALES.includes(String(o.estatus_externo||'').toUpperCase())) continue;
    if (otExcluir && String(o.ot).trim() === String(otExcluir).trim()) continue;
    if (String(o.fecha_visita||'').trim() !== String(fechaVisitaISO).trim()) continue; // misma fecha exacta
    const yaAsignados = String(o.personal||'').split(',').map(norm);
    const choca = solicitados.find(p => yaAsignados.includes(p));
    if (choca) return `${choca} ya tiene la OT ${o.ot} programada para la misma fecha (${o.fecha_visita}).`;
  }
  return null;
}

// Uso (modo advertencia):
const aviso = await hayChoqueAgenda(payload.personal_asignado, payload.fecha_visita, payload.ot_folio);
if (aviso && !confirm('⚠️ Choque de agenda: ' + aviso + '\n\n¿Registrar la OT de todos modos?')) return;
```

**Notas:**
- `getOrdenes` compara `fecha_visita` tal cual está en la hoja. Conviene confirmar que ese valor se guarda en ISO (viene de `dates[0]`, que es `type="date"` → ISO). Si en la hoja hubiera fechas en otro formato, la comparación exacta fallaría; por eso mantener `type="date"` (§3) es lo que hace confiable esta verificación.
- Los nombres en la OT usan formato corto ("Martín Luna"); la normalización `norm` (sin acentos, minúsculas) los homogeniza. Conviene revisar que el catálogo de personal sea consistente entre SEAOT y equipos.
- Si más adelante se quiere que la verificación sea **infalible** (no saltable desde el cliente), se movería a `fase2_RegistrarOT` — queda como mejora futura, no en este PR.

---

## 5. Control de "salida abierta" de equipo por número de inventario — **sin hoja nueva** (decisión #3)

**Realidad:** `registro-equipos.html` no persiste nada y **no se creará** una hoja. Por lo tanto, **no es posible** verificar de forma confiable si `EA-SO11-01` "ya salió en otro documento/otro día", porque no hay dónde consultar el historial. Sería deshonesto prometerlo sin persistencia.

**Lo que SÍ se puede hacer hoy, dentro del propio formato (frontend, sin backend):**

### 5.a Impedir/avisar equipo duplicado dentro del mismo formato
El caso más común de error humano —seleccionar dos veces el mismo inventario en la misma salida— sí se detecta en la tabla actual. Ya existe la base `existeEquipo(clave)` (recorre los `.equipo-select`). Se refuerza al cambiar un equipo:

```js
// en alCambiarEquipo(select), tras leer la clave elegida:
function inventarioYaEnTabla(clave, selfSelect){
  return Array.from(document.querySelectorAll('#tablaEquipos tbody .equipo-select'))
    .some(s => s !== selfSelect && s.value === clave);
}
// dentro de alCambiarEquipo:
if (clave && inventarioYaEnTabla(clave, select)){
  alert(`⚠️ El equipo ${clave} ya está en esta salida. Revisa que no lo estés registrando dos veces.`);
}
```

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
| 2 | Choque de agenda y cronología | **Advertencia** (avisa y deja continuar) |
| 3 | Control de equipos por inventario | **Sin crear hoja** → solo antiduplicado dentro del formato |
| 4 | Prioridad inmediata | **Reescribir lo menos posible** al generar formatos (§0-bis) |

---

## 7. Orden sugerido de implementación para el PR (menor a mayor riesgo)

1. **§0-bis — Prellenado de formatos** (el pedido inmediato): fecha en supervisión + Entregó/Recibió prellenados en equipos. Solo frontend, aditivo. *(Entra primero.)*
2. **§1.a — Validación cronológica** Revisión ≤ Emisión ≤ Visita, como aviso en SEAOT.
3. **§2.a — fecha solicitud ≤ visita** como aviso en registro-equipos.
4. **§4 — Choque de agenda** vía `getOrdenes` (lectura existente) + aviso.
5. **§5.a — Antiduplicado de inventario** dentro del formato.
6. **§3 — Helper de fecha** `fechaADMY_` solo si se necesita mostrar `dd/mm/yyyy` en algún formato.

Todo es **frontend y aditivo**: no reordena el contrato A–Q, no cambia el backend, no crea hojas. Cada punto puede ir como commit independiente dentro del mismo PR.
