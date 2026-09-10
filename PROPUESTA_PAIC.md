# Propuesta — Alineación de PAIC con las modificaciones de SEADB

**Estado:** Análisis y plan para revisión (NO implementado).
**Flujo previsto:** Eduwin revisa → decide las 4 preguntas de §4 → se abre PR por fases.
**Fecha:** 2026-09-10
**Alcance:** `SEADB.html`, `PAIC.html`, `BACKEND_FIXES.gs`, `TESTS_BACKEND.gs`, `MANUAL.md`.

> **Resumen en una línea:** SEADB evolucionó a un tablero que depende de tres llaves de datos —
> `RFC|Sucursal`, `EstatusExterno` (con `EN PAUSA`) y `NOM` — y PAIC, que es la puerta de entrada de esas
> tres llaves, sigue en el contrato de hace un año: **manda datos que el backend descarta en silencio,
> pisa columnas que SEADB lee en vivo, y no tiene ningún camino de regreso hacia el asesor.**

---

## 0. Qué se modificó en SEADB (el "objetivo móvil")

Historial real de `SEADB.html` (commits que cambiaron comportamiento, del más antiguo al más reciente):

| Commit | Cambio | Alcance real |
|---|---|---|
| `1c96727` | **Separar `ORDENES_TRABAJO` de `INFORMES`** — estatus independientes | Backend + SEADB + SEAINF |
| `3019f53` | Corrección de cálculo de fechas y formato **DD-MM-YYYY** | SEADB + SEAINF |
| `9250373` | Renombrar **"Cliente Final" → "Sucursal"**, etiquetas SLA `+20d` / `+25d` | SEADB |
| `78f3027` → `f2aca9b` | **Pestaña Renovaciones** (primero backend, luego recalculada en el cliente desde el CSV de Informes) | Backend + SEADB |
| `508abc1` | Cliente Inicial en el modal de OTs Vencidas | SEADB |
| `659434e` | **Mecanismo `EN PAUSA` / Reanudar** (congelar y recalcular SLA) | Backend + SEADB |
| `dc353cb` | `safeDriveUrl()` — el botón Drive abría el propio dashboard | SEADB |

**PAIC recibió en el mismo periodo un solo commit:** `4789d1c` (migración reCAPTCHA Classic → Enterprise),
que fue una obligación de Google, no una evolución del producto. **Cero adopciones funcionales.**

---

## 1. Cómo opera SEADB hoy (revisión)

### 1.1 Identidad del módulo

| Aspecto | Valor |
|---|---|
| Archivo | `SEADB.html` (1,375 líneas, todo el JS embebido en `<script>` líneas 452–1358) |
| Acceso | Google OAuth. `SEAAuth.init(..., { gasUrl: API_URL })` (`SEADB.html:1369`) verifica la whitelist **antes** de pintar la UI |
| Autorización | `AUTH_MODE` = `GOOGLE` + `ACTION_MODULE` = `SEADB` (`BACKEND_FIXES.gs:150-205`) |
| Pestañas | Órdenes de Trabajo · Informes · Calendario · Renovaciones |
| Librerías | FullCalendar 6.1.10 y Chart.js 4.4.0 por CDN |

### 1.2 Las DOS fuentes de datos (el hecho más importante del módulo)

```
                    ┌──────────────────────────────────────────┐
   Pestaña          │  A) API autenticada  (Google OAuth)      │
   Órdenes  ───────►│  GET  ?action=getTablero                 │
   Calendario       │  POST updateEstatus / pausarOT /         │
                    │       reanudarOT                          │
                    │  Lee: ORDENES_TRABAJO (A–Q)              │
                    │     + CLIENTES_MAESTRO (col 22, asesor)  │
                    └──────────────────────────────────────────┘

                    ┌──────────────────────────────────────────┐
   Pestaña          │  B) CSV publicado a la web (SIN AUTH)    │
   Informes  ──────►│  docs.google.com/.../pub?output=csv      │
   Renovaciones     │  gid=808750033        (SEADB.html:454)   │
                    │  Layout: fecha, empresa, nom, persona,   │
                    │  correo, link, remitente, rfc, …,        │
                    │  sucursal(col L)                          │
                    └──────────────────────────────────────────┘
```

Consecuencias que hay que tener presentes antes de tocar nada:

1. **El CSV NO es la hoja `INFORMES` del sistema.** El contrato `CONFIG.COLUMNS.INFORMES` (A–Q, `SUCURSAL`
   en Q/índice 16) no coincide con el layout que lee `loadInformes()` (`SEADB.html:1019`, `sucursal` en
   índice 11). Es **otra hoja, fuera del contrato del backend**.
2. **El endpoint `getRenovaciones()` del backend quedó huérfano.** Existe (`BACKEND_FIXES.gs:2773`, ~119
   líneas, con auth `GOOGLE` y módulo `SEADB`) pero **SEADB ya no lo llama**: tras `f2aca9b` calcula las
   renovaciones en el navegador desde el CSV. Hoy hay dos motores de renovación con reglas distintas
   (el backend exige `ESTATUS = FINALIZADO` + `FECHA_REAL`; el cliente usa la fecha del informe).
3. **Ese CSV expone razón social, RFC, correo de contacto y link de Drive de cada informe sin
   autenticación**, y la URL viaja dentro del HTML del dashboard. Es una decisión de arquitectura que
   PAIC **no debe heredar** (ver §3, Fase 5).

### 1.3 Pipeline de la pestaña Órdenes (el corazón)

`loadData()` (`SEADB.html:484`) hace un solo `GET ?action=getTablero` y normaliza cada fila:

```
fase4_GetTablero()  (BACKEND_FIXES.gs:1865)
  ├── getDisplayValues() de ORDENES_TRABAJO      → texto tal como se ve en la hoja
  ├── Deduplica por OT quedándose con la ÚLTIMA fila (recorre de abajo hacia arriba)
  ├── asesorMap[RFC + '|' + Sucursal] ← CLIENTES_MAESTRO col 22   ◄── join en vivo
  └── Devuelve: ot, nom, cliente, sucursal, rfc, tipo_orden, responsable,
                fecha_visita, fechaEntrega, fechaRealEntrega, estatus, link_drive,
                asesor_consultor, fechaPausa, motivoPausa, fechaInfoCompleta
```

En el cliente, cada registro recibe un **estatus calculado** (`SEADB.html:509-530`):

| Condición | Estatus mostrado | Categoría (`getStatusCategory`) |
|---|---|---|
| `ENTREGADO` o `FINALIZADO` | ✅ Entregado | `ENTREGADO` |
| `EN PAUSA` | ⏸ En Pausa (desde dd/mm/aaaa) | `PAUSA` |
| `diffDias < 0` | 🔴 Alerta [N días] | `ALERTA` |
| `0 ≤ diffDias ≤ 7` | 🟠 En límite [N días] | `LIMITE` |
| `diffDias > 7` | 🟢 En tiempo [N días] | `TIEMPO` |
| Sin fecha de visita ni de entrega | 🟡 Pendiente | `PENDIENTE` |

**Reglas SLA vigentes** (`SEADB.html:476-477`): `SLA_DIGITAL_DIAS = 20`, `SLA_FISICO_EXTRA = 5`.
La fecha límite es `FECHA_ENTREGA` de la hoja si existe; si no, `fecha_visita + 20`. El físico
siempre es digital `+ 5`. Formato de salida **DD-MM-YYYY** (`formatearFechaSLA`, `SEADB.html:479`).
`parseFecha()` (`SEADB.html:675`) acepta ISO, `YYYY-MM-DD`, `DD-MM-YYYY` y `DD/MM/YYYY` —
tolerancia deliberada porque las fuentes escriben en formatos distintos.

Al terminar la carga, `checkOtsVencidas()` abre un **modal bloqueante** con las OTs en `ALERTA`
(excluye explícitamente las `EN PAUSA`, `SEADB.html:569-575`).

### 1.4 El mecanismo `EN PAUSA` (la modificación más profunda)

Es lo único que agregó columnas nuevas al contrato de `ORDENES_TRABAJO` (ahora A–Q, 17 columnas):

| Col | Campo | Escrito por |
|---|---|---|
| O (14) | `FECHA_PAUSA` | `pausarOT_` |
| P (15) | `MOTIVO_PAUSA` | `pausarOT_` |
| Q (16) | `FECHA_INFO_COMPLETA` | `reanudarOT_` |

```
   NO INICIADO ─┐
                ├──► pausarOT_    ──► EN PAUSA ──► reanudarOT_ ──► EN PROCESO
   EN PROCESO ──┘    (motivo ≥10)     (SLA           (fechaInfoCompleta      (FECHA_ENTREGA
                                       congelado)      + motivo ≥10)          recalculada)
```

- `pausarOT_` (`BACKEND_FIXES.gs:1715`) solo acepta origen `NO INICIADO` / `EN PROCESO`
  (`ESTATUS_PAUSABLES_`), exige motivo ≥ 10 caracteres y escribe 3 renglones de auditoría.
- `reanudarOT_` (`BACKEND_FIXES.gs:1753`) recalcula `FECHA_ENTREGA = fechaInfoCompleta + SLA`,
  con `SLA = 25` si `TIPO = OTB`, si no `20`. **`FECHA_VISITA` nunca se toca** — decisión de diseño
  para no falsear el historial de campo. Limpia O y P, y **acumula** el motivo en `OBSERVACIONES`
  con sello de tiempo.
- `EN PAUSA` es **reversible**: no está en `ESTATUS_EXTERNO_TERMINALES_` (`BACKEND_FIXES.gs:141`).

**La razón de negocio de todo el mecanismo es literal:** *"el cliente no entregó información completa"*.
Guárdala, porque es exactamente el hueco de PAIC (§2, G6).

### 1.5 Pestañas Informes y Renovaciones

- **Informes**: carga perezosa del CSV al primer clic; vistas tarjeta/tabla, filtros en cascada
  (empresa → NOM), y acciones de compartir por WhatsApp / correo / copiar enlace.
- **Renovaciones**: `calcularRenovacionesDesdeInformes()` (`SEADB.html:1229`) agrupa por
  `RFC|Sucursal|Servicio`, se queda con el informe más reciente y proyecta el vencimiento:

| Servicio | Ciclo | | Urgencia | Umbral |
|---|---|---|---|---|
| NOM-022-STPS, NOM-081-SEMARNAT, PIPC / Programa Interno | 1 año | | Vencido | `< 0 días` |
| NOM-015, NOM-024, NOM-025-STPS | 2 años | | Urgente | `< 30 días` |
| | | | Próximo | `30–90 días` |
| | | | Al corriente | `> 90 días` |

### 1.6 Escrituras y auditoría

Solo tres acciones escriben desde SEADB: `updateEstatus`, `pausarOT`, `reanudarOT`. Las tres pasan por
`SEAAuth.wrapFetch` y las tres dejan rastro en la hoja `AUDITORIA` vía `registrarAuditoria_`
(usuario, acción, OT, campo, antes, después). **PAIC no genera ni un solo renglón de auditoría.**

### 1.7 Convenciones de robustez que SEADB ya adoptó

`escHtml()` en todo render, `escapeQuotes()` en handlers inline, `safeDriveUrl()` (solo acepta
`https://`, si no deshabilita el botón), toasts no bloqueantes en vez de `alert()`, y carga perezosa
por pestaña. Son el estándar de facto del repo. PAIC **no las tiene**.

---

## 2. Qué le falta a PAIC — análisis de brechas

**Qué es PAIC hoy:** un formulario público de una sola dirección. Un `POST` con
`action: 'registrarCliente'` + `portal_origen: 'PAIC'` (`PAIC.html:966`), protegido con reCAPTCHA
Enterprise, más una búsqueda por RFC para prellenar. Comparte backend con SEAPD: **el mismo
`fase1_RegistrarCliente`**, sin ninguna ramificación por portal.

Las brechas, ordenadas por severidad. Cada una con evidencia verificable en el código.

---

### 🔴 G1 — Los archivos de estudio de PAIC se descartan en silencio

`guardarArchivos()` (`BACKEND_FIXES.gs:1978`) tiene una **whitelist fija** de campos de archivo.
Estos 6 campos que PAIC sí sube **no están en ella**:

```
hojas_campo_laboratorio   fotografias_laboratorio   croquis_laboratorio
hojas_campo_higiene       fotografias_higiene       croquis_higiene
```

El asesor los selecciona, se codifican en base64, viajan por la red, **cuentan contra el tope de
25 MB** (`PAIC.html:979`) … y el backend los ignora. El usuario ve *"✓ Información enviada"*.

**Impacto en SEADB:** son precisamente las hojas de campo del estudio que se registra. Sin ellas el
expediente nace vacío, SEAINF no las encuentra, y la OT arranca sin insumos → **es la causa típica de
un `EN PAUSA`**. La modificación estrella de SEADB está compensando un bug de PAIC.

> Nota: `mantenimiento` y `proceso_produccion` sí están en la whitelist pero **no existen en PAIC**
> (son de SEAPD). La lista se escribió para SEAPD y nunca se revisó contra PAIC.

---

### 🔴 G2 — Reregistrar desde PAIC **borra** Asesor/Consultor y Requiere PIPC

`fase1_RegistrarCliente` no hace merge: cuando encuentra la fila del `RFC + Sucursal`, la
**sobrescribe completa** (`BACKEND_FIXES.gs:783`) con los 22 valores del payload.

Y PAIC no repuebla dos de ellos:

| Columna | Qué pasa | Por qué |
|---|---|---|
| 22 · `ASESOR_CONSULTOR` | Queda **vacía** | `applyClientData()` (`PAIC.html:1172`) no incluye `asesor_consultor` en su `fieldMap`, aunque `fase2_BuscarClienteRFC` **sí lo devuelve** (`BACKEND_FIXES.gs:846`) |
| 16 · `REQUIERE_PIPC` | Queda en **`'NO'`** | PAIC no tiene el campo `requiere_pipc`; el backend evalúa `data.requiere_pipc === 'si' ? 'SÍ' : 'NO'` |

**Impacto en SEADB — directo y retroactivo:** la columna "Asesor/Consultor" del tablero **no está
almacenada en la OT**; se resuelve en cada carga con `asesorMap[RFC + '|' + Sucursal]`
(`BACKEND_FIXES.gs:1869-1877, 1903`). Al vaciarse la celda de `CLIENTES_MAESTRO`, **todas las OTs de
ese cliente pierden el asesor, incluidas las históricas**, y el filtro "Asesor / Consultor" deja de
encontrarlas. El canal por el que entró el negocio se borra solo.

---

### 🔴 G3 — `sucursal` es texto libre y es la llave primaria de todo el sistema

`<input type="text" name="sucursal" required>` (`PAIC.html:219`). Sin normalización, sin lista de
sucursales conocidas, sin previsualización. Y sobre ese texto se apoyan **cuatro normalizaciones
distintas**:

| Consumidor | Normalización | Referencia |
|---|---|---|
| Almacenamiento en `CLIENTES_MAESTRO` | **ninguna** — `data.sucursal` crudo | `BACKEND_FIXES.gs:750` |
| Carpeta de Drive | `sanitizeFileName()` + `toLowerCase()` | `BACKEND_FIXES.gs:1041-1043` |
| Join del tablero (`asesorMap`) | `.trim()`, **sensible a mayúsculas** | `BACKEND_FIXES.gs:1873-1876` |
| Renovaciones (SEADB) | `.trim()`, sensible a mayúsculas | `SEADB.html:1236` |

`sanitizeFileName()` (`BACKEND_FIXES.gs:2630`) reemplaza todo lo no alfanumérico por `_`, **no hace
trim** y corta a 50 caracteres. Un espacio al inicio ya rompe la coincidencia de carpeta.

**Tres fallas de SEADB que nacen aquí:**
1. La OT se bloquea con `CARPETA_NO_ENCONTRADA` (`BACKEND_FIXES.gs:927`) — y el mensaje dice
   *"Registra o corrige el cliente en **SEAPD**"*, aunque el registro haya entrado por PAIC.
2. Columna Asesor/Consultor en blanco (falla el join exacto).
3. **Renovaciones duplicadas**: "Planta Norte" y "planta norte" generan dos series de vencimiento
   para la misma sucursal, cada una con su propia fecha.

---

### 🟠 G4 — El estudio solicitado no se persiste en ningún lado

PAIC captura `estudio_laboratorio`, `estudio_higiene` y `otra_nom_higiene` — **el corazón de la regla
"1 registro = 1 estudio"**. En el backend tienen **cero referencias**: no van a `CLIENTES_MAESTRO`, no
van a la hoja de perfil (`generarPerfilSheet`, `BACKEND_FIXES.gs:2026-2073`), no van a ninguna hoja.

**Impacto en SEADB:** la NOM que llega al tablero es la que **alguien vuelve a teclear** en SEAOT. No
hay forma de auditar "lo que pidió el asesor" contra "lo que se registró", y los KPIs por NOM y toda
la pestaña de Renovaciones heredan ese error de captura sin posibilidad de detectarlo.

---

### 🟠 G5 — `portal_origen: 'PAIC'` se envía y se ignora

`PAIC.html:966` lo manda en cada payload. **Cero referencias en el backend.** `CLIENTES_MAESTRO` no
tiene columna de origen. SEADB no puede distinguir negocio directo de negocio de canal, ni medir
cuántas OTs produce el portal de asesores. El único proxy es la columna 22 … que G2 borra.

---

### 🟠 G6 — No hay camino de regreso: el ciclo `EN PAUSA` está roto por diseño

Este es **el hueco funcional real** entre SEADB y PAIC:

```
  Asesor ──PAIC──► registro ──► OT ──► SEADB detecta info incompleta
                                            │
                                            ▼
                                     pausarOT_ escribe MOTIVO_PAUSA (col P)
                                            │
                                            ✗  nadie se lo dice al asesor
                                            ✗  no hay dónde subir lo faltante
                                            ✗  fechaInfoCompleta se teclea a mano
```

SEADB ya sabe **por qué** está detenida cada OT (el motivo es obligatorio, ≥10 caracteres) y **desde
cuándo**. Esa información nunca sale del tablero interno. El asesor —que es quien puede resolverlo—
no se entera, y el reloj del SLA queda congelado esperando una llamada telefónica.

---

### 🟡 G7 — `fechas_preferidas` es texto libre

`<textarea name="fechas_preferidas" required>` con ejemplo *"primera semana de febrero"*. SEAOT
necesita una `fecha_visita` real: es la que alimenta el semáforo, el fallback `+20 días` y el
calendario de SEADB. Hoy la conversión es manual y sin trazabilidad.

---

### 🟡 G8 — El prellenado está desalineado con lo que devuelve el backend

`applyClientData()` (`PAIC.html:1172`) mapea `actividad_principal` y `descripcion_proceso`, que
`fase2_BuscarClienteRFC` **nunca devuelve** (no existen esas columnas en `CLIENTES_MAESTRO`) →
siempre `undefined`. Y **no** mapea tres campos que el backend **sí** devuelve: `asesor_consultor`
(→ G2), `aplica_nom020` y `requiere_pipc`.

---

### 🟡 G9 — PAIC no adoptó las convenciones de robustez de SEADB

| Convención en SEADB | Estado en PAIC |
|---|---|
| `escHtml()` en todo render | ❌ `innerHTML` con `${data.razon_social}` sin escapar (`PAIC.html:1121`) — el dato viene de la propia base |
| `safeDriveUrl()` | ❌ no existe |
| Toasts no bloqueantes | ❌ `alert()` (`PAIC.html:1147`) |
| Fechas DD-MM-YYYY consistentes | n/a (no maneja fechas) |
| Etiqueta de versión viva | ❌ congelada en `PAIC v1.0` (`PAIC.html:143`) |

---

### 🟡 G10 — Si PAIC va a mostrar informes, no debe heredar el atajo del CSV

Ver §1.2. Antes de que PAIC muestre nada de informes o renovaciones hay que decidir **cuál es la
fuente de verdad**: la hoja `INFORMES` del contrato (A–Q) o la hoja del CSV publicado. Hoy conviven
dos motores de renovación con reglas distintas y uno de ellos es código muerto.

---

### ⚪ G11 — Deuda de pruebas y documentación

- `TESTS_BACKEND.gs` E01 prueba `registrarCliente` **etiquetado como "SEAPD"** (`TESTS_BACKEND.gs:216`).
  No hay ningún caso PAIC.
- **No existe ninguna prueba de `pausarOT_` / `reanudarOT_`** — cero coincidencias en el archivo de
  tests, pese a ser el mecanismo que toca el SLA.
- `MANUAL.md §6.1` no documenta `asesor_consultor`, ni los campos de archivo, ni `portal_origen`.
- `MANUAL.md §3.4` describe pestañas de SEADB que ya no existen (dice "Resumen / Órdenes / Reportes /
  Calendario"; las reales son Órdenes / Informes / Calendario / Renovaciones) y no menciona `EN PAUSA`
  en la tabla de estatus externos.

---

### Matriz de trazabilidad: modificación de SEADB → qué necesita PAIC

| Modificación de SEADB | ¿PAIC la puede adoptar hoy? | Brecha que lo impide |
|---|---|---|
| Separación `ORDENES_TRABAJO` / `INFORMES` | Parcial | G10 (fuente de verdad sin definir) |
| Renombrar a "Sucursal" + SLA `+20d/+25d` | ❌ | G3 (la llave no está normalizada) |
| Formato DD-MM-YYYY | ❌ | G7 (PAIC no captura fechas reales) |
| Pestaña Renovaciones | ❌ | G3 + G10 (duplica series, fuente ambigua) |
| Columna Asesor/Consultor | ❌ **la rompe activamente** | G2 |
| Mecanismo `EN PAUSA` / Reanudar | ❌ | G6 (sin canal de retorno) + G1 (sin archivos) |
| `safeDriveUrl` / `escHtml` | ❌ | G9 |

---

## 3. Plan por fases

Principio rector, heredado de `PROPUESTA_TRAZABILIDAD.md`: **no romper contratos.**
`CLIENTES_MAESTRO` = 22 columnas (guard en `BACKEND_FIXES.gs:694`), `ORDENES_TRABAJO` = A–Q
(guard en `BACKEND_FIXES.gs:935`). Todo lo que sigue respeta ambos.

---

### Fase 1 — Detener la pérdida de datos (sin tocar el esquema)

*Riesgo: bajo · Sin cambio de contrato · Es la fase que más dolor quita por línea escrita.*

**1.1 — Registrar los 6 campos de archivo de PAIC** (G1)
Agregar a `fileFields` en `guardarArchivos()`:
```
hojas_campo_laboratorio  → 'L1) Hojas de campo - Laboratorio'
fotografias_laboratorio  → 'L2) Fotografías - Laboratorio'
croquis_laboratorio      → 'L3) Croquis - Laboratorio'
hojas_campo_higiene      → 'H1) Hojas de campo - Higiene'
fotografias_higiene      → 'H2) Fotografías - Higiene'
croquis_higiene          → 'H3) Croquis - Higiene'
```
Se guardan en `01_Cliente` como el resto. `validarArchivo_` ya cubre tipo y tamaño.

**1.2 — Merge en vez de sobrescritura ciega** (G2) — *la corrección de raíz*
En `fase1_RegistrarCliente`, cuando `rowIndex > -1`: si el payload **no trae** un campo y la fila
existente **sí tiene valor**, conservar el valor previo. Aplicar al menos a `ASESOR_CONSULTOR` y
`REQUIERE_PIPC`. Protege también a SEAPD y a cualquier portal futuro.
> Sutileza: `requiere_pipc` hoy se colapsa a `'SÍ'`/`'NO'` y pierde la diferencia entre "dijo que no"
> y "no preguntó". El merge debe distinguir **campo ausente** de **campo en `'no'`**.

**1.3 — Prellenado honesto en PAIC** (G8)
Agregar `asesor_consultor` al `fieldMap`; quitar `actividad_principal` y `descripcion_proceso`
(no existen en el origen); prellenar `aplica_nom020` y `requiere_pipc` si se agrega el campo.

**1.4 — Escapado en PAIC** (G9)
`escHtml()` (portado tal cual de `SEADB.html:1190`) en el render del modal de cliente encontrado.

**Pruebas:** extender E01 con variante PAIC (con archivos de laboratorio) + nuevo caso
"reregistro por PAIC no borra `ASESOR_CONSULTOR` ni `REQUIERE_PIPC`".

---

### Fase 2 — Normalizar `sucursal`, la llave del sistema

*Riesgo: medio (toca joins existentes) · Requiere diagnóstico previo.*

**2.1 — Diagnóstico primero, solo lectura**
Script que liste, por RFC, todas las variantes de `sucursal` presentes en `CLIENTES_MAESTRO`,
`ORDENES_TRABAJO` y la hoja de informes, señalando cuáles rompen hoy el `asesorMap` y cuáles duplican
series de renovación. **Sin este inventario no se decide nada** — dirá si el problema son 3 filas o 300.

**2.2 — Una sola función de normalización**
`normalizarSucursal_(s)` en el backend: `trim` → colapsar espacios múltiples → Title Case.
Se aplica **al escribir** (`fase1`) y **al comparar** (`asesorMap`, `folderMatchesClientBranch_`).

**2.3 — PAIC deja de ser texto libre**
El portal ya conoce las sucursales del RFC (`fase2_BuscarClienteRFC` las devuelve). Convertir el input
en selector + opción "nueva sucursal" con normalización en vivo y **previsualización del nombre de
carpeta** que se creará. El asesor ve exactamente lo que el sistema va a guardar.

**2.4 — Sin reescribir históricos en el mismo PR**
La corrección de datos existentes va en un paso aparte, revisado a mano contra el inventario de 2.1.

---

### Fase 3 — Persistir la intención del asesor

*Riesgo: bajo-medio · Aquí sí hay una decisión de arquitectura.*

Problema a resolver: G4 (estudio solicitado) + G5 (origen) + G7 (fechas preferidas).

| | Opción A — hoja `SOLICITUDES` (append-only) | Opción B — ampliar `CLIENTES_MAESTRO` |
|---|---|---|
| Contrato de 22 columnas | intacto | **roto** (guard e índices `CL.*`) |
| Modelo | "1 registro = 1 estudio" es un **evento** | lo trata como atributo del cliente |
| Historial | conserva cada solicitud | la última pisa a la anterior |
| Beneficio extra | embudo medible **solicitud → OT**; SEAOT puede prellenar la NOM | ninguno |

**Recomendación: Opción A.** Hoja nueva, solo append, con: timestamp · folio de solicitud ·
`portal_origen` · `asesor_consultor` · RFC · sucursal normalizada · estudio/NOM solicitada ·
fechas preferidas · link de carpeta · estatus de atención (`RECIBIDA` → `OT_GENERADA` → `DESCARTADA`).
No toca ningún contrato existente y le da a SEADB, por primera vez, un embudo real del canal de asesores.

**Sobre G7 (fechas):** cambiar el textarea por dos `<input type="date">` (rango preferido) **más** el
textarea como nota. La fecha estructurada viaja a `SOLICITUDES`; SEAOT la propone como `fecha_visita`.
Formato de captura ISO, render DD-MM-YYYY — misma regla que ya aplica `SEADB.html:479`.

---

### Fase 4 — Cerrar el ciclo `EN PAUSA` ↔ PAIC

*Riesgo: alto (toca el camino crítico del SLA) · Detrás de bandera · Exige las pruebas de G11 primero.*

**4.1 — `pausarOT_` notifica al asesor**
Al pausar, enviar correo a `asesor_consultor` + `correo_informe` con el motivo (que ya es obligatorio).
Reusa la infraestructura de correo recién rediseñada (`enviarNotificacionRobusta`, `botonEmail_`,
`contactoWhatsAppCliente_`) — no hay que construir nada nuevo.

**4.2 — PAIC: "Completar información pendiente"**
Pantalla con folio de OT + token del correo, que sube únicamente los archivos faltantes a la carpeta
del expediente. Escritura acotada: **no** modifica `CLIENTES_MAESTRO` ni `ORDENES_TRABAJO`.

**4.3 — `reanudarOT_` toma la fecha real de la subida**
`fechaInfoCompleta` deja de teclearse: es la fecha en que el asesor efectivamente entregó. El operador
confirma en SEADB en vez de capturar. El SLA se vuelve auditable de punta a punta.

**Orden obligatorio:** 4.1 → 4.2 → 4.3. Cada paso es útil por sí solo y 4.1 ya resuelve el 80% del dolor.

---

### Fase 5 — Espejo de seguimiento para el asesor (PAIC v2)

*Riesgo: medio · Módulo nuevo, aditivo y aislado.*

El repo ya tiene el patrón resuelto **dos veces**: `PORTAL/` y `TRAZ/` son proyectos de Apps Script
independientes, de solo lectura, que no tocan `BACKEND_FIXES.gs`. `PORTAL/` incluso resuelve la
autenticación de externos sin cuenta de Google: **OTP por correo + token de sesión firmado con HMAC**,
con aislamiento por RFC.

**PAIC-Seguimiento = mismo patrón, alcance por `asesor_consultor` en vez de por RFC.** Muestra:
- el mismo semáforo de SEADB (mismas reglas de §1.3, no una segunda implementación);
- `EN PAUSA` **con su motivo** — el cierre natural de la Fase 4;
- fechas digital y físico en DD-MM-YYYY;
- renovaciones de su cartera.

**Prerrequisito duro: resolver G10** (definir la fuente de verdad de informes). Si no, se construye
un segundo consumidor sobre un CSV público y el problema se duplica en vez de resolverse.

---

### Fase 6 — Cerrar la deuda

- Pruebas E2E de `pausarOT_` / `reanudarOT_` (transiciones válidas, rechazo desde `ENTREGADO`,
  motivo corto, recálculo `20` vs `25` días, limpieza de O y P).
- Caso E2E de PAIC con archivos de laboratorio y con reregistro.
- `MANUAL.md`: §3.4 (pestañas reales + `EN PAUSA`), §3.5 (PAIC real), §6.1 (payload completo),
  §4.2 (columnas O, P, Q).
- Etiqueta de versión de PAIC viva.

---

## 4. Decisiones que necesito de ti

Las Fases 1 y 2 **no dependen de estas respuestas** — se pueden arrancar ya.

1. **§3 Fase 3:** ¿hoja `SOLICITUDES` nueva (recomendado) o se queda todo en `CLIENTES_MAESTRO`?
   → En `PROPUESTA_TRAZABILIDAD.md` decidiste "sin crear hojas nuevas". Aquí la propongo igual porque
   el caso es distinto: no es un control operativo interno, es el registro del canal de negocio. Tu llamada.
2. **§2 G10:** ¿la fuente de verdad de informes es la hoja `INFORMES` (contrato A–Q) o la hoja del CSV
   publicado? De esto depende si `getRenovaciones()` se arregla o se borra.
3. **§3 Fase 4:** ¿el asesor puede subir información faltante directamente, o solo recibe el aviso y
   responde por WhatsApp/correo? (La segunda opción es Fase 4.1 sola y cuesta una fracción.)
4. **§3 Fase 5:** ¿PAIC v2 es prioridad este trimestre, o basta con cerrar Fases 1–4?

---

## 5. Orden sugerido (de menor a mayor riesgo)

| # | Trabajo | Riesgo | Desbloquea |
|---|---|---|---|
| 1 | Fase 1 completa (G1, G2, G8, G9) | bajo | Deja de perderse información **hoy** |
| 2 | Fase 2.1 — diagnóstico de sucursales | nulo (solo lectura) | Dimensiona G3 con datos reales |
| 3 | Fase 6 — pruebas de pausa/reanudación | bajo | Prerrequisito de la Fase 4 |
| 4 | Fase 2.2–2.3 — normalización | medio | Asesor/Consultor y Renovaciones confiables |
| 5 | Fase 3 — hoja `SOLICITUDES` | medio | Embudo del canal + prellenado de NOM en SEAOT |
| 6 | Fase 4.1 — aviso de pausa al asesor | medio | 80% del ciclo cerrado |
| 7 | Fase 4.2–4.3 — subida y reanudación | alto | SLA auditable punta a punta |
| 8 | Fase 5 — PAIC v2 | medio | Autoservicio del canal |

---

## 6. Qué NO hacer

- **No** ampliar `CLIENTES_MAESTRO` más allá de 22 columnas sin migrar a la vez el guard y los
  índices `CL.*`; hay lecturas por índice en 5 funciones distintas.
- **No** reordenar ni insertar columnas en `ORDENES_TRABAJO`: el contrato A–Q está validado en
  `fase2_RegistrarOT` y asumido por `fase4_GetTablero`, `pausarOT_` y `reanudarOT_`.
- **No** dar a PAIC un segundo CSV publicado. Si el asesor necesita leer datos, va por endpoint
  autenticado o por el patrón OTP de `PORTAL/`.
- **No** corregir los datos históricos de sucursal en el mismo PR que introduce la normalización.
- **No** tocar `FECHA_VISITA` al reanudar: es una decisión de diseño explícita del mecanismo de pausa.
