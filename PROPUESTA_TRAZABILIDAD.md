# Propuesta — Mejoras de Trazabilidad SEAOT

**Estado:** Propuesta para revisión (NO implementada).
**Flujo previsto:** Eduwin revisa esta primera propuesta → ajustes → se abre PR con Codex para implementar.
**Fecha:** 2026-08-26

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

Uso en el submit:
```js
const err = validarCronologiaOT(_d('wo_contractReviewDate'), _d('wo_issueDate'), dates);
if (err){ showStatus('workorderStatus', '❌ ' + err, 'error'); return; }
```

### 1.b Revalidación en backend (defensa en profundidad)

El frontend se puede saltar. Se añade la misma comprobación en `fase2_RegistrarOT` usando los campos que **ya se pueden enviar** en el payload (hoy no van; hay que agregarlos al `payload` de SEAOT: `emision`, `fecha_revision_contrato`). No cambia el esquema de la hoja, solo valida:

```js
// dentro de fase2_RegistrarOT(data), antes del appendRow
var errCron = validarCronologiaOT_(data.fecha_revision_contrato, data.emision, data.fecha_visita);
if (errCron) return { success:false, error: errCron };
```
```js
function validarCronologiaOT_(revision, emision, visita){
  var r = normalizarFechaISO_(revision), e = normalizarFechaISO_(emision), v = normalizarFechaISO_(visita);
  if (!e || !v) return 'Faltan fechas obligatorias (emisión o visita) o tienen formato inválido.';
  if (r && r > e) return 'La revisión de contrato es posterior a la emisión de la OT.';
  if (e > v)      return 'La emisión de la OT es posterior a la fecha de visita.';
  return '';
}
```
> `normalizarFechaISO_` se define en el punto 3.

### 1.c ¿Persistir Emisión y Revisión? — decisión a tomar

Hoy no se guardan. Dos caminos (elige uno; **recomiendo B**):

- **Opción A — Solo validar (cero cambios de esquema).** Rápido y sin riesgo. Desventaja: la trazabilidad cronológica no queda *registrada* en la hoja, solo se garantiza en el momento del registro. TRAZ no podría mostrar esas fechas.
- **Opción B — Extender el contrato a A–S (recomendada).** Agregar dos columnas al final: `R = Fecha Emisión OT`, `S = Fecha Revisión Contrato`. Se hace de forma segura porque se **agrega al final** (no reordena A–Q): `fase2_RegistrarOT` valida `getMaxColumns() >= 19`, el `appendRow` añade dos valores más, y `setupSheets()` incluye los encabezados nuevos. Así la trazabilidad queda persistida y TRAZ puede leerla. Requiere `MIGRACION.gs` para rellenar filas viejas (o dejarlas vacías).

---

## 2. Equipos: fecha de solicitud ≤ fecha de visita + nombres en "Recibió" / "Entregó"

### 2.a Fecha de solicitud ≤ Fecha de visita
En `registro-equipos.html`, `fechaSalida` es la fecha de solicitud/salida y la fecha de visita llega en el handoff (`data.fecha`). Propuesta: guardar la fecha de visita recibida y validarla al imprimir/guardar:

```js
// en applyOtHandoff(): conservar la fecha de visita de la OT
if (data.fecha){ window.__fechaVisitaOT = data.fecha; }

// validación al imprimir/exportar (o al enviar a backend, ver 5)
function validarFechaSolicitud(){
  const fs = document.getElementById('fechaSalida').value;
  const fv = window.__fechaVisitaOT;
  if (fs && fv && fs > fv){
    alert(`La fecha de solicitud (${fs}) no puede ser posterior a la fecha de visita (${fv}).`);
    return false;
  }
  return true;
}
```

### 2.b Nombres estructurados en "Entregó" (ingresó equipo) y "Recibió" (recibir equipo)
Hoy son solo líneas de firma a mano. Propuesta: convertirlas en campos con el **mismo catálogo de personal** que ya usa `solicitante` (reutilizar la lista `EDUWIN IVÁN…`, `EDUARDO CAMPOS…`, `MARTÍN LUNA…`, `JIMMY AYALA`), dejando la firma manuscrita como está:

```html
<!-- SALIDA -->
<label>Entregó (nombre):</label>   <select id="salidaEntrego">  … mismas opciones que #solicitante …</select>
<label>Recibió (nombre):</label>   <select id="salidaRecibio">  … </select>
<!-- ENTRADA / retorno -->
<label>Entregó (nombre):</label>   <select id="entradaEntrego"> … </select>
<label>Recibió (nombre):</label>   <select id="entradaRecibio"> … </select>
```
Sugerencia de prellenado coherente con lo que ya hace `applyOtHandoff`: `salidaRecibio = solicitante` (quien recibe el equipo para ir a campo es normalmente el solicitante) y `salidaEntrego` = responsable de almacén/equipos. Ajustable.

---

## 3. Validación de formato de fecha (día/mes/año) para no romper backend

**Diagnóstico:** los campos `<input type="date">` **ya entregan `YYYY-MM-DD` (ISO)**, que es no ambiguo y seguro para backend. El riesgo real aparece si:
- se muestra/captura como texto `dd/mm/yyyy` y Google Sheets lo **auto-interpreta como mm/dd** (US), o
- se agregan campos de texto libre para fechas.

**Propuesta — utilería única compartida (backend) + display consistente:**

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

**Regla operativa recomendada:**
- **Internamente todo se maneja en ISO `YYYY-MM-DD`** (captura con `type="date"`, almacenamiento en hoja, comparaciones).
- **Solo se convierte a `dd/mm/yyyy` para mostrar** (PDF, formatos, TRAZ) con `fechaADMY_`.
- Cualquier fecha que entre al backend pasa por `normalizarFechaISO_` y, si devuelve `''`, se rechaza el registro con mensaje claro. Esto elimina el riesgo de que Sheets guarde una fecha en formato equivocado.

---

## 4. Antichoque de agenda: mismo "Personal Asignado" ya tiene OT activa en la misma "Fecha de Visita"

Validación nueva dentro de `fase2_RegistrarOT`, **antes** del `appendRow`. Lee `ORDENES_TRABAJO`, y para cada persona del `Personal Asignado` busca otra OT **activa** (estatus no terminal) con **exactamente la misma** `FECHA_VISITA`.

```js
/**
 * Devuelve {conflicto:true, detalle:'...'} si alguna persona ya tiene OT activa esa fecha.
 * personalAsignado: string separado por comas (como lo arma SEAOT).
 * fechaVisitaISO:   'YYYY-MM-DD' ya normalizada.
 * otExcluir:        folio de la OT actual (para permitir re-registros/edición).
 */
function verificarChoqueAgenda_(personalAsignado, fechaVisitaISO, otExcluir){
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  var values = sheet.getDataRange().getValues();
  var fechaObjetivo = normalizarFechaISO_(fechaVisitaISO);
  if (!fechaObjetivo) return { conflicto:false };

  var solicitados = _split_(personalAsignado);   // ['martin luna', ...] normalizado
  for (var i = 1; i < values.length; i++){
    var row = values[i];
    var estatus = String(row[CO.ESTATUS_EXTERNO] || '').trim().toUpperCase();
    if (ESTATUS_EXTERNO_TERMINALES_.indexOf(estatus) !== -1) continue;        // ignora FINALIZADO/CANCELADO
    if (otExcluir && String(row[CO.OT]).trim() === String(otExcluir).trim()) continue;
    if (normalizarFechaISO_(row[CO.FECHA_VISITA]) !== fechaObjetivo) continue; // misma fecha exacta

    var yaAsignados = _split_(row[CO.PERSONAL]);
    for (var p = 0; p < solicitados.length; p++){
      if (yaAsignados.indexOf(solicitados[p]) !== -1){
        return { conflicto:true, detalle: 'La persona "' + solicitados[p] +
                 '" ya tiene la OT ' + row[CO.OT] + ' programada para ' + fechaADMY_(fechaObjetivo) + '.' };
      }
    }
  }
  return { conflicto:false };
}
function _split_(s){
  return String(s||'').split(',').map(function(x){
    return x.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/\s+/g,' ').trim();
  }).filter(Boolean);
}
```

Enganche en `fase2_RegistrarOT`:
```js
var choque = verificarChoqueAgenda_(data.personal_asignado, data.fecha_visita, data.ot_folio);
if (choque.conflicto) return { success:false, error: '⚠️ Choque de agenda: ' + choque.detalle };
```

**Decisiones a confirmar:**
- ¿Bloqueo duro (no deja registrar) o advertencia con opción "registrar de todos modos"? (recomiendo bloqueo duro con mensaje claro).
- Los nombres en la OT usan formato corto ("Martín Luna") — la normalización `_split_` ya lo homogeniza; conviene revisar que el catálogo de personal sea consistente entre SEAOT y equipos.

---

## 5. Estructura base para control de "salida abierta" de equipo por número de inventario

**Problema:** hoy `registro-equipos.html` no persiste nada, así que no hay forma de saber si `EA-SO11-01` ya está fuera. Se propone una **hoja nueva** y una **función base** de verificación. Es la pieza más grande porque introduce persistencia donde no había.

### 5.a Nueva hoja `MOVIMIENTOS_EQUIPO` (contrato A–J)

| Col | Campo | Ejemplo |
|---|---|---|
| A | Timestamp registro | `2026-08-26 10:15` |
| B | Inventario | `EA-SO11-01` |
| C | OT / Proyecto | `OT2608-1` |
| D | Solicitante | `MARTÍN LUNA CÓRDOVA` |
| E | Entregó (salida) | `EDUWIN IVÁN…` |
| F | Recibió (salida) | `MARTÍN LUNA…` |
| G | Fecha salida (ISO) | `2026-08-27` |
| H | Fecha retorno esperada (ISO) | `2026-08-28` |
| I | Fecha retorno real (ISO) | `` (vacío = **salida abierta**) |
| J | Estatus | `ABIERTA` / `CERRADA` |

La "salida abierta" = fila con `J=ABIERTA` (equivalente: `I` vacío).

### 5.b Función base de verificación (backend)

```js
CONFIG.SHEET_MOV_EQUIPO = 'MOVIMIENTOS_EQUIPO';
CONFIG.COLUMNS.MOV = { TS:0, INV:1, OT:2, SOLIC:3, ENTREGO:4, RECIBIO:5, F_SALIDA:6, F_RET_ESP:7, F_RET_REAL:8, ESTATUS:9 };

/**
 * ¿El equipo `inventario` tiene una salida abierta que se traslape con `fechaISO`?
 * Regla: existe fila ABIERTA cuya [F_SALIDA .. F_RET_ESP] cubre la fecha,
 * o cualquier fila ABIERTA sin retorno real (equipo aún fuera).
 * Devuelve {abierta:true, detalle} o {abierta:false}.
 */
function tieneSalidaAbierta_(inventario, fechaISO){
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_MOV_EQUIPO);
  if (!sheet) return { abierta:false };                 // sin hoja aún → no bloquea
  var MV = CONFIG.COLUMNS.MOV;
  var inv = String(inventario||'').toUpperCase().trim();
  var f = normalizarFechaISO_(fechaISO);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++){
    var row = values[i];
    if (String(row[MV.INV]||'').toUpperCase().trim() !== inv) continue;
    if (String(row[MV.ESTATUS]||'').toUpperCase().trim() !== 'ABIERTA') continue;
    var ini = normalizarFechaISO_(row[MV.F_SALIDA]);
    var fin = normalizarFechaISO_(row[MV.F_RET_ESP]) || '9999-12-31'; // sin retorno esperado = sigue fuera
    if (!f || (f >= ini && f <= fin)){
      return { abierta:true, detalle: 'El equipo ' + inv + ' tiene una salida abierta (OT ' +
               row[MV.OT] + ', desde ' + fechaADMY_(ini) + ').' };
    }
  }
  return { abierta:false };
}

/** Registra una salida (llamada desde registro-equipos.html vía nueva acción 'registrarSalidaEquipo'). */
function registrarSalidaEquipo_(data){
  // 1) validar formato de fechas con normalizarFechaISO_
  // 2) por cada inventario: if (tieneSalidaAbierta_(inv, data.fecha_salida).abierta) → rechazar
  // 3) appendRow por equipo con J='ABIERTA'
  // (esqueleto — detalle en el PR)
}
```

### 5.c Cierre de salida
Al retornar el equipo, una acción `cerrarSalidaEquipo_(inventario)` pone `I = fecha retorno real` y `J = CERRADA`. Esto es lo que "libera" el inventario para futuras validaciones.

**Alcance mínimo viable (para el PR con Codex):**
1. Crear hoja + `setupSheets()` la incluye.
2. `tieneSalidaAbierta_` + `normalizarFechaISO_` (utilería compartida con puntos 1, 3, 4).
3. Acción `registrarSalidaEquipo_` conectada al botón de `registro-equipos.html` (además del PDF actual).
4. Auditoría con `registrarAuditoria_` en salida y cierre.

---

## 6. Resumen de decisiones que necesito de ti antes del PR

| # | Decisión | Recomendación |
|---|---|---|
| 1 | ¿Persistir Emisión+Revisión (extender a A–S) o solo validar? | **Extender A–S** (trazabilidad real) |
| 2 | Choque de agenda: ¿bloqueo duro o advertencia? | **Bloqueo duro** con mensaje claro |
| 3 | Control de equipos: ¿implementamos la hoja `MOVIMIENTOS_EQUIPO` ahora o solo dejamos la función base? | **Función base + hoja en el mismo PR** (si no, no hay dónde consultar) |
| 4 | ¿Prellenado de "Recibió" = solicitante? | Sí, ajustable |
| 5 | Catálogo de personal: unificar nombres entre SEAOT (corto) y equipos (completo) | Sí, tabla de mapeo única |

---

## 7. Orden sugerido de implementación (transversal → específico)

1. **Utilería de fechas** (`normalizarFechaISO_`, `fechaADMY_`) — la usan 1, 3, 4 y 5.
2. **Validación cronológica** (frontend + backend) — bajo riesgo, alto valor.
3. **Choque de agenda** — solo backend, aditivo.
4. **Campos Entregó/Recibió + fecha solicitud ≤ visita** en `registro-equipos.html`.
5. **Hoja `MOVIMIENTOS_EQUIPO` + salida abierta** — el cambio más grande, al final.

Cada bloque es **aditivo** y no reordena el contrato A–Q existente (el punto 1 opción B solo agrega columnas al final).
