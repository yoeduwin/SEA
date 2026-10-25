# Arquitectura de datos y ciclo operativo de SEA

**Sistema:** SEA — Ejecutiva Ambiental  
**Documento:** fuente de referencia para la relación entre clientes, órdenes de trabajo, expedientes/informes y envío de informes.  
**Fecha de definición:** 25/10/2026

> Este documento distingue entre **ESTADO ACTUAL** (lo que existe hoy) y **CAMBIO ACORDADO** (lo que se decidió implementar posteriormente). No asumir que un cambio acordado ya está desarrollado sólo por aparecer aquí.

---

## 1. Principio general

El ciclo natural del servicio debe ser:

```text
CLIENTE / SUCURSAL
        ↓
ORDEN DE TRABAJO
        ↓
EXPEDIENTE / INFORME
        ↓
ENVÍO DEL INFORME
        ↓
CIERRE DEL SERVICIO
```

Regla operativa acordada:

```text
1 OT = 1 servicio = 1 expediente/informe
```

La **OT** es la llave transversal del ciclo operativo.

---

# 2. Sistemas involucrados

## 2.1 SEAPD

**Función:** alta y actualización de clientes/sucursales.

**Base principal:** `CLIENTES_MAESTRO`.

SEAPD registra la identidad y datos maestros del cliente. No debe convertirse en una base de seguimiento operativo del servicio.

---

## 2.2 SEAOT

**Función:** registrar la Orden de Trabajo de un servicio.

**Base principal:** `ORDENES_TRABAJO`.

Consulta `CLIENTES_MAESTRO` para recuperar los datos del cliente/sucursal y genera la OT.

---

## 2.3 SEAINF

**Función:** crear y administrar el expediente asociado a una OT y, cuando aplica, generar el número de informe.

**Base principal:** `INFORMES` del archivo principal de SEA.

El backend actual aplica la regla:

```text
1 OT = 1 expediente
```

Una OT que ya tiene expediente deja de mostrarse como disponible para crear otro expediente.

---

## 2.4 SEADB

**Función:** tablero de consulta y gestión del estado de las OTs y de los informes enviados.

SEADB no constituye una base independiente del ciclo. Principalmente consulta:

- `ORDENES_TRABAJO`
- `INFORMES` de SEAINF
- hoja externa `INFORMES` de informes enviados

Por ello, cuando `ORDENES_TRABAJO` cambia a `FINALIZADO`, SEADB refleja el cierre al volver a consultar los datos.

---

## 2.5 Sistema local Python — envío de informes

**Función actual:**

1. Seleccionar Excel/perfil del cliente.
2. Leer datos del cliente.
3. Pedir norma/servicio.
4. Seleccionar PDF final.
5. Subir o sustituir el PDF en Google Drive.
6. Generar enlace.
7. Registrar el envío en Google Sheets.
8. Generar QR.
9. Enviar correo mediante Gmail/SMTP.

**Base principal:** hoja externa `INFORMES`, en el archivo denominado actualmente `PANEL DE EJECUTIVA`.

Este sistema permanece local por el momento. Su futura migración a web queda fuera del cambio inmediato.

---

# 3. Base `CLIENTES_MAESTRO`

## 3.1 Función

`CLIENTES_MAESTRO` es la **base maestra de clientes/sucursales de SEA**.

Cada fila representa una combinación de:

```text
RFC + Sucursal
```

Por ello, un mismo RFC puede aparecer en varias filas cuando existen distintas sucursales o plantas.

## 3.2 Contrato actual — 22 columnas

| Col. | Campo |
|---|---|
| A | Fecha Registro |
| B | Razón Social |
| C | Sucursal |
| D | RFC |
| E | Representante Legal |
| F | Dirección Evaluación |
| G | Teléfono Empresa |
| H | Nombre Solicitante |
| I | Correo Informe |
| J | Giro |
| K | Registro Patronal |
| L | Capacidad Instalada |
| M | Capacidad Operación |
| N | Días / Turnos / Horarios |
| O | Aplica NOM-020 — **en desuso** |
| P | Requiere PIPC — **en desuso** |
| Q | Responsable |
| R | Teléfono Responsable |
| S | Nombre dirigido / nombre para informe |
| T | Puesto dirigido / puesto para informe |
| U | Link Drive |
| V | Asesor / Consultor |

Las columnas O y P se conservan por compatibilidad del contrato de 22 columnas, pero actualmente no se escriben como atributos maestros del cliente.

## 3.3 Uso acordado

`CLIENTES_MAESTRO` **no se modifica para el proyecto de cierre automático**.

Debe usarse para:

- consultar datos del cliente/sucursal;
- recuperar datos al registrar una OT;
- actualizar datos maestros cuando corresponda;
- localizar la carpeta Drive del cliente/sucursal.

No debe recibir información operativa de informes, envíos o cierres.

---

# 4. Base `ORDENES_TRABAJO`

## 4.1 Función

Es la **base maestra de las órdenes de trabajo**.

Una fila describe el servicio que debe ejecutarse y su estado general.

La llave principal es:

```text
OT
```

## 4.2 Contrato actual A–Q

| Col. | Campo |
|---|---|
| A | Fecha |
| B | OT |
| C | Tipo |
| D | NOM / Servicio |
| E | Cliente |
| F | Sucursal |
| G | RFC |
| H | Personal asignado |
| I | Fecha Visita |
| J | Fecha Entrega |
| K | Fecha Real |
| L | Estatus Externo |
| M | Link Drive |
| N | Observaciones |
| O | Fecha Pausa |
| P | Motivo Pausa |
| Q | Fecha Información Completa |

## 4.3 Estados actuales

Valores válidos actuales de `Estatus Externo`:

```text
NO INICIADO
EN PROCESO
EN PAUSA
ENTREGADO
FINALIZADO
CANCELADO
```

Estados terminales:

```text
FINALIZADO
CANCELADO
```

`EN PAUSA` es reversible.

## 4.4 Regla de uso

`ORDENES_TRABAJO` sí puede ser actualizada por otros procesos, pero únicamente con información propia de la OT.

Ejemplos correctos:

- estatus;
- fecha real;
- responsable;
- pausa/reanudación;
- observaciones de la orden.

No agregar en esta hoja datos propios de otras etapas, como bitácoras de envío, destinatarios de correos o archivos enviados.

Regla:

```text
Si el dato describe la OT → ORDENES_TRABAJO
Si describe otra etapa/proceso → base correspondiente
```

---

# 5. Base `INFORMES` de SEAINF

## 5.1 Función

Es la **base de expedientes/informes generados desde las OTs**.

SEAINF selecciona una OT existente y crea el expediente correspondiente.

La relación directa es:

```text
ORDENES_TRABAJO.B (OT)
        ↓
INFORMES.D (OT)
```

## 5.2 Contrato actual A–Q

| Col. | Campo |
|---|---|
| A | Timestamp |
| B | Número de Informe |
| C | Tipo de Orden |
| D | OT |
| E | NOM / Servicio |
| F | Cliente |
| G | Solicitante |
| H | RFC |
| I | Teléfono |
| J | Dirección |
| K | Fecha Servicio |
| L | Fecha Entrega |
| M | Es Capacitación |
| N | Estatus |
| O | Link Drive del expediente |
| P | Responsable |
| Q | Sucursal |

## 5.3 Estados actuales del informe

```text
NO INICIADO
EN PROCESO
PARA REVISION
PARA IMPRESION
FINALIZADO
CANCELADO
```

Estados terminales:

```text
FINALIZADO
CANCELADO
```

## 5.4 Regla actual de expediente

El backend aplica:

```text
1 OT = 1 expediente
```

Si la OT ya tiene registro en `INFORMES`, SEAINF reutiliza el expediente existente y no crea un segundo expediente para la misma OT.

## 5.5 Número de informe

- OT / OTB: normalmente llevan número de informe.
- OTC: no lleva número de informe.

Por ello, el **número de informe no puede ser la única llave universal del sistema**.

La OT sí puede cumplir esa función.

---

# 6. Drive — relación con las bases

Los enlaces de Drive de cada base representan cosas distintas.

## 6.1 `CLIENTES_MAESTRO.U`

Apunta a la carpeta correspondiente al cliente/sucursal.

Estructura general:

```text
RFC - RAZÓN SOCIAL/
└── SUCURSAL/
    └── 01_Cliente/
```

## 6.2 `ORDENES_TRABAJO.M`

Mantiene el enlace de referencia a la carpeta de la sucursal del cliente.

## 6.3 `SEAINF.INFORMES.O`

Apunta a la carpeta específica del expediente.

Estructura actual:

```text
RFC - RAZÓN SOCIAL/
└── SUCURSAL/
    ├── 01_Cliente/
    └── 02_Expediente_.../
        ├── 1. ORDEN_TRABAJO/
        ├── 2. HDC/
        ├── 3. CROQUIS/
        ├── 4. FOTOS/
        ├── 5. INFORMES Y MEMORIAS/
        └── 6. INFORME PRELIMINAR/
```

## 6.4 Sistema Python — link del PDF

La hoja externa de informes enviados registra el enlace directo al PDF final almacenado por el sistema Python.

Por tanto:

```text
CLIENTES_MAESTRO.U  = carpeta cliente/sucursal
ORDENES_TRABAJO.M   = carpeta cliente/sucursal usada por la OT
SEAINF.INFORMES.O   = carpeta del expediente
PYTHON.INFORMES.F   = archivo PDF final enviado al cliente
```

No son campos duplicados.

## 6.5 Regla de diseño Drive

**No agregar niveles de carpetas sólo para representar relaciones lógicas.**

Drive debe seguir siendo almacenamiento documental. Las relaciones del negocio deben resolverse mediante los datos y las llaves de las hojas.

---

# 7. Hoja externa `INFORMES` — sistema Python

## 7.1 ESTADO ACTUAL

El script actual escribe en:

```text
INFORMES!A:K
```

Columnas actuales:

| Col. | Campo actual |
|---|---|
| A | Fecha y hora |
| B | Empresa |
| C | Norma / servicio |
| D | Destinatario |
| E | Correo destinatario |
| F | Link PDF |
| G | Remitente |
| H | RFC |
| I | File ID |
| J | Nombre del archivo |
| K | Estatus (`Enviado`) |

El sistema actual solicita manualmente la norma y obtiene cliente/RFC/destinatario/correo desde el Excel seleccionado.

## 7.2 Comportamiento actual del archivo PDF

El nombre se construye aproximadamente como:

```text
MMAA_CLIENTE_NORMA.pdf
```

Si dentro de la carpeta mensual ya existe un archivo con el mismo nombre, el script actualiza ese archivo y conserva el `file_id`.

## 7.3 Integración actual con SEA

Aunque Python es un programa local independiente, SEA ya conoce esta hoja.

`BACKEND_FIXES.gs` contiene configuración para el archivo externo de informes enviados y `SEADB` consulta esos registros mediante:

```text
getInformesEnviados_()
```

El backend actualmente intenta leer también:

```text
L = Sucursal
```

pero el script Python actual sólo escribe `A:K`.

Esta diferencia debe corregirse en la siguiente modificación del sistema Python.

---

# 8. Problema actual de trazabilidad

`ORDENES_TRABAJO` y `SEAINF.INFORMES` están relacionadas directamente por `OT`.

Sin embargo, la hoja del sistema Python actualmente **no registra**:

```text
OT
Número de informe
Sucursal
```

Por ello, actualmente sólo podría inferirse una relación usando combinaciones como:

```text
RFC + Cliente + Norma
```

Esto no es una llave suficientemente segura porque un mismo cliente puede contratar el mismo servicio varias veces.

## 8.1 Llave acordada

La llave principal para conectar el ciclo será:

```text
OT
```

El número de informe se conservará también cuando exista, pero será una referencia adicional.

---

# 9. CAMBIO ACORDADO — selección segura desde Python

## 9.1 No capturar manualmente OT ni número de informe

Para evitar errores tipográficos, el usuario **no escribirá manualmente** la OT ni el número de informe en Python.

Python deberá consultar los registros de `SEAINF.INFORMES` y presentar una selección similar a:

```text
EA-2610-025-0042
Empresa ABC
NOM-025-STPS
OT2610-123
```

Al seleccionar un registro, Python debe recibir automáticamente:

```text
OT
Número de informe
Cliente
RFC
Sucursal
Servicio / NOM
```

Esto evita depender de captura manual para las llaves del proceso.

## 9.2 Uso mínimo de SEAINF

No es requisito que el personal haya mantenido actualizado el estatus de SEAINF durante el desarrollo del trabajo.

Puede ocurrir:

```text
SEAINF se usó para:
- seleccionar la OT;
- crear expediente;
- obtener número de informe.
```

Aunque el informe siga con un estado anterior, el envío exitoso en Python podrá cerrar automáticamente tanto el informe como la OT.

---

# 10. CAMBIO ACORDADO — columnas futuras del registro de enviados

La hoja externa de informes enviados deberá conservar A:K y agregar:

| Col. | Campo futuro |
|---|---|
| L | Sucursal |
| M | OT |
| N | Número de informe |

Estructura propuesta completa:

| Col. | Campo |
|---|---|
| A | Fecha/hora de envío |
| B | Empresa |
| C | Servicio / NOM |
| D | Destinatario |
| E | Correo destinatario |
| F | Link PDF |
| G | Remitente |
| H | RFC |
| I | File ID |
| J | Nombre del archivo |
| K | Estado del envío |
| L | Sucursal |
| M | OT |
| N | Número de informe |

El registro de envíos será una **bitácora**. Por ello puede existir más de una fila de envío para el mismo informe en caso de reenvíos posteriores.

Los reenvíos/correcciones posteriores continuarán gestionándose manualmente y no requieren una automatización adicional en esta etapa.

---

# 11. CAMBIO ACORDADO — cierre automático por envío

## 11.1 Evento de cierre

El cierre natural del servicio será el **primer envío exitoso del informe**.

Flujo acordado:

```text
PYTHON
Envío exitoso
        ↓
Registro de envío = ENVIADO
        ↓
SEAINF.INFORMES.N = FINALIZADO
        ↓
ORDENES_TRABAJO.L = FINALIZADO
ORDENES_TRABAJO.K = fecha real de envío
```

## 11.2 Consecuencia en SEADB

SEADB consulta `ORDENES_TRABAJO`.

Por ello:

```text
Python cierra OT
        ↓
ORDENES_TRABAJO = FINALIZADO
        ↓
SEADB refleja automáticamente el servicio como finalizado
```

No es necesario crear una segunda lógica de cierre dentro de SEADB.

## 11.3 Control manual se conserva

El cierre automático por envío **no elimina** el control manual existente.

Deben seguir disponibles las operaciones manuales actuales, entre ellas:

```text
EN PAUSA
EN PROCESO / reanudación
FINALIZADO
CANCELADO
```

El envío exitoso únicamente agrega otro camino válido hacia `FINALIZADO`.

---

# 12. Importante — cuándo registrar `ENVIADO`

## 12.1 Problema del script actual

Actualmente el script escribe la fila en Google Sheets con:

```text
K = Enviado
```

**antes** de ejecutar el envío SMTP.

Esto permite un caso inconsistente:

```text
Sheets = Enviado
Correo = Falló
```

También puede ocurrir el caso inverso si falla el registro en Sheets y el script continúa:

```text
Correo = Enviado
Sheets = Sin registro
```

## 12.2 Regla futura

Para utilizar el envío como evento de cierre, `ENVIADO` debe significar:

```text
El correo fue confirmado como enviado correctamente.
```

Por tanto, la siguiente modificación de Python deberá ordenar/transaccionar el proceso para que el cierre sólo ocurra después del envío exitoso.

---

# 13. CAMBIO ACORDADO — backend SEA

No se modificarán las estructuras A:Q de `ORDENES_TRABAJO` ni `SEAINF.INFORMES`.

Se agregará al backend una operación de cierre por envío que reciba al menos la OT validada y realice:

```text
1. Buscar OT en ORDENES_TRABAJO.
2. Buscar expediente por OT en INFORMES.
3. Marcar INFORMES.N = FINALIZADO.
4. Marcar ORDENES_TRABAJO.L = FINALIZADO.
5. Registrar ORDENES_TRABAJO.K = fecha real de envío.
6. Registrar cambios en AUDITORIA.
```

Debe ser idempotente o, como mínimo, tolerar que una OT ya esté `FINALIZADO` sin generar un estado incorrecto.

---

# 14. Qué NO se modifica en esta etapa

```text
CLIENTES_MAESTRO          → no cambiar estructura
ORDENES_TRABAJO A:Q       → no agregar/reordenar columnas
SEAINF.INFORMES A:Q       → no agregar/reordenar columnas
Estructura Drive          → no agregar niveles de carpetas
SEADB                     → no requiere nueva base ni cierre propio
```

El trabajo inmediato se concentra en:

```text
1. Python local
2. Hoja externa de informes enviados
3. Nueva operación de cierre en BACKEND_FIXES.gs
```

---

# 15. Modelo de relaciones

## 15.1 Relación principal

```text
CLIENTES_MAESTRO
   RFC + Sucursal
        │
        │ consulta
        ▼
ORDENES_TRABAJO
        OT
        │
        │ 1 : 1
        ▼
SEAINF.INFORMES
        OT + N.º Informe (cuando aplica)
        │
        │ 1 : N posibles eventos de envío
        ▼
INFORMES ENVIADOS (Python)
```

## 15.2 Interpretación de cada base

```text
CLIENTES_MAESTRO
= quién es el cliente / sucursal

ORDENES_TRABAJO
= qué servicio se debe ejecutar

SEAINF.INFORMES
= qué expediente/informe produjo la orden

INFORMES ENVIADOS
= cuándo, cómo y a quién se entregó el documento
```

---

# 16. Estados: no mezclar conceptos

Los estados de las distintas hojas describen etapas diferentes.

```text
ORDENES_TRABAJO.L
= estado global del servicio

SEAINF.INFORMES.N
= estado de elaboración/control del informe

INFORMES ENVIADOS.K
= resultado del evento de envío
```

No deben sustituirse entre sí.

El evento `ENVIADO` es el disparador que sincroniza los cierres:

```text
ENVIADO
   ↓
Informe FINALIZADO
   ↓
OT FINALIZADO
```

---

# 17. Fecha real y renovaciones

El backend actual usa `ORDENES_TRABAJO.K` como `Fecha Real`.

El módulo de renovaciones utiliza OTs `FINALIZADO` con `Fecha Real` para determinar el último servicio realizado.

Por ello, al usar la fecha efectiva del envío como `Fecha Real`, el ciclo queda conectado también con las renovaciones:

```text
Envío
 ↓
Fecha Real
 ↓
OT FINALIZADO
 ↓
Cálculo de renovación futura
```

---

# 18. Reenvíos y correcciones

Decisión operativa:

- el **primer envío exitoso** cierra automáticamente informe y OT;
- los reenvíos, sustituciones o correcciones posteriores se gestionan manualmente;
- no se desarrollará por ahora una lógica especial de reapertura/cierre por reenvío.

La bitácora de enviados puede conservar nuevos eventos de envío sin alterar la identidad original de la OT ni del informe.

---

# 19. Migración futura del sistema Python

El sistema Python puede migrarse posteriormente a una aplicación web conservando su lógica funcional:

- lectura del perfil Excel;
- carga/selección del PDF;
- Drive;
- sustitución del PDF conservando `file_id`;
- Sheets;
- QR;
- correo HTML;
- CC;
- bitácora de envío;
- cierre automático de SEA.

La migración **no forma parte del cambio inmediato**.

Decisión actual:

```text
Mantener Python local
→ integrar correctamente con SEA ahora
→ migrar a web después
```

---

# 20. Resumen de decisiones vigentes

1. `CLIENTES_MAESTRO` permanece como base maestra de clientes/sucursales.
2. `ORDENES_TRABAJO` permanece como base maestra de OTs.
3. `SEAINF.INFORMES` permanece como base de expedientes/informes.
4. La hoja externa del Python permanece como bitácora de informes enviados.
5. No se fusionarán las tres bases operativas.
6. La **OT** será la llave principal de trazabilidad.
7. Una OT representa un servicio y un expediente/informe.
8. OT y número de informe no se capturarán manualmente en Python; se seleccionarán desde SEAINF.
9. La hoja externa agregará `Sucursal`, `OT` y `Número de informe` en L:N.
10. `ENVIADO` sólo debe registrarse después de confirmar el envío correcto del correo.
11. El primer envío exitoso marcará el informe como `FINALIZADO`.
12. El mismo evento marcará la OT como `FINALIZADO` y colocará la fecha real.
13. SEADB reflejará el cierre automáticamente al consultar `ORDENES_TRABAJO`.
14. La pausa, cancelación y finalización manual continúan disponibles.
15. Reenvíos/correcciones se gestionan manualmente.
16. No se agregan carpetas Drive para representar relaciones lógicas.
17. No se modifican los contratos A:Q de `ORDENES_TRABAJO` ni `SEAINF.INFORMES`.
18. La migración de Python a web se posterga; la integración actual debe diseñarse para poder reutilizarse posteriormente.

---

# 21. Próximo trabajo técnico previsto

Cuando se retome la implementación:

```text
A. Python
   - consultar SEAINF.INFORMES;
   - mostrar selector de informe/OT;
   - tomar OT, N.º informe, RFC, sucursal y servicio sin captura manual;
   - corregir secuencia de confirmación del envío;
   - escribir columnas L:N;
   - solicitar cierre de SEA después de envío exitoso.

B. BACKEND_FIXES.gs
   - implementar operación autenticada de cierre por envío;
   - finalizar INFORMES por OT;
   - finalizar ORDENES_TRABAJO por OT;
   - registrar Fecha Real;
   - registrar AUDITORIA.

C. Validación
   - envío exitoso;
   - error SMTP;
   - fallo de Sheets;
   - OT ya finalizada;
   - informe ya finalizado;
   - OTC sin número de informe;
   - reenvío posterior.
```

---

## Fuente de verdad

Para cambios futuros sobre este flujo, actualizar primero este documento cuando cambie cualquiera de los siguientes contratos:

- columnas de las hojas;
- definición de llaves;
- reglas de relación;
- estados terminales;
- evento de cierre;
- comportamiento del sistema de envíos.

Esto evita depender de conversaciones históricas para reconstruir la arquitectura del sistema.
