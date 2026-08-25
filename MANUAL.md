# MANUAL DEL SISTEMA SEA
## Sistema Ejecutivo Ambiental — Ejecutiva Ambiental
### Versión 3.0 — Multi-Sucursal

---

## ÍNDICE

1. [Descripción General](#1-descripción-general)
2. [Arquitectura del Sistema](#2-arquitectura-del-sistema)
3. [Módulos del Sistema](#3-módulos-del-sistema)
   - 3.1 [SEAPD — Registro de Clientes](#31-seapd--registro-de-clientes)
   - 3.2 [SEAOT — Órdenes de Trabajo](#32-seaot--órdenes-de-trabajo)
   - 3.3 [SEAINF — Gestión de Expedientes](#33-seainf--gestión-de-expedientes)
   - 3.4 [SEADB — Dashboard y Control](#34-seadb--dashboard-y-control)
   - 3.5 [PAIC — Portal de Asesores y Consultores](#35-paic--portal-de-asesores-y-consultores)
4. [Base de Datos (Google Sheets)](#4-base-de-datos-google-sheets)
5. [Estructura de Carpetas en Drive](#5-estructura-de-carpetas-en-drive)
6. [API Backend — Referencia de Funciones](#6-api-backend--referencia-de-funciones)
7. [Autenticación y Seguridad](#7-autenticación-y-seguridad)
8. [Configuración e Instalación](#8-configuración-e-instalación)
9. [Sistema de Respaldos](#9-sistema-de-respaldos)
10. [Gestión de Usuarios](#10-gestión-de-usuarios)
11. [Pruebas Automatizadas](#11-pruebas-automatizadas)
12. [Solución de Problemas](#12-solución-de-problemas)
13. [Glosario](#13-glosario)

---

## 1. Descripción General

El **Sistema SEA** (Sistema Ejecutivo Ambiental) es una plataforma web integral para la gestión de clientes, órdenes de trabajo y expedientes ambientales de **Ejecutiva Ambiental**. Permite administrar el ciclo completo de un servicio ambiental: desde el registro del cliente hasta la entrega del informe final.

### Características principales

- Registro centralizado de clientes con carpetas en Google Drive creadas automáticamente
- Generación y seguimiento de Órdenes de Trabajo (OT y OTB)
- Creación de expedientes con estructura de carpetas estandarizada
- Dashboard en tiempo real con semáforos de estatus y alertas de vencimiento
- Portal público para asesores, intermediarios y consultores (PAIC)
- Soporte multi-sucursal: un mismo cliente puede tener varias sucursales o plantas
- Autenticación Google OAuth para usuarios internos y reCAPTCHA para portales públicos
- Respaldos automáticos semanales con retención de 12 semanas

### Tecnologías utilizadas

| Componente | Tecnología |
|---|---|
| Backend | Google Apps Script (GAS) |
| Frontend | HTML5 + JavaScript puro |
| Base de datos | Google Sheets |
| Almacenamiento | Google Drive |
| Autenticación interna | Google OAuth 2.0 |
| Autenticación pública | reCAPTCHA v3 |
| Zona horaria | GMT-6 (México) |

---

## 2. Arquitectura del Sistema

```
┌─────────────────────────────────────────────────────────┐
│                     USUARIOS                            │
│  Internos (Google Auth)    │   Externos (reCAPTCHA)     │
│  SEAPD · SEAOT · SEAINF   │   PAIC                     │
│         SEADB              │                             │
└────────────┬───────────────┴────────────┬────────────────┘
             │                            │
             ▼                            ▼
┌────────────────────────────────────────────────────────┐
│              BACKEND (Google Apps Script)              │
│                    BACKEND_FIXES.gs                    │
│                                                        │
│  ┌──────────────────┐   ┌────────────────────────┐    │
│  │ Módulo de        │   │ Módulo de Seguridad     │    │
│  │ Negocio          │   │ OAuth + reCAPTCHA       │    │
│  │                  │   │                         │    │
│  │ fase1_Registrar  │   │ verificarIdToken_       │    │
│  │ fase2_Buscar     │   │ verificarRecaptcha_     │    │
│  │ fase2_RegistrarOT│   │ verificarUsuario_       │    │
│  │ fase3_Expediente │   └────────────────────────┘    │
│  │ getOrdenes       │                                  │
│  │ updateEstatus    │                                  │
│  └──────────────────┘                                  │
└──────────────┬─────────────────────────────────────────┘
               │
       ┌───────┴────────┐
       ▼                ▼
┌─────────────┐  ┌─────────────────┐
│ Google      │  │ Google Drive    │
│ Sheets      │  │                 │
│             │  │ /SEA_ROOT/      │
│ CLIENTES_   │  │  /RFC_EMPRESA/  │
│ MAESTRO     │  │   /Sucursal/    │
│             │  │    /Expediente/ │
│ ORDENES_    │  │     /OT/        │
│ TRABAJO     │  │     /HDC/       │
│             │  │     /CROQUIS/   │
│ USUARIOS_   │  │     /FOTOS/     │
│ AUTORIZADOS │  └─────────────────┘
│             │
│ AUDITORIA   │
└─────────────┘
```

### Flujo de datos principal

```
Cliente →[SEAPD]→ Registro BD + Carpeta Drive
                        ↓
Operador →[SEAOT]→ OT en ORDENES_TRABAJO
                        ↓
Informe →[SEAINF]→ Expediente en Drive + Actualización OT
                        ↓
Gestión →[SEADB]→ Dashboard + Semáforos + Alertas
```

---

## 3. Módulos del Sistema

### 3.1 SEAPD — Registro de Clientes

**Archivo:** `SEAPD.html`
**Acceso:** Portal público con reCAPTCHA v3
**Función:** Registro de nuevos clientes o actualización de datos existentes

#### Campos del formulario

| Campo | Requerido | Descripción |
|---|---|---|
| Razón Social | Sí | Nombre legal de la empresa |
| RFC | Sí | Registro Federal de Contribuyentes (13 caracteres) |
| Sucursal / Planta | Sí | Nombre de la sucursal o "Matriz" |
| Representante Legal | Sí | Nombre del representante legal |
| Dirección de evaluación | Sí | Domicilio donde se realizará el servicio |
| Teléfono | Sí | Teléfono de contacto de la empresa |
| Correo | Sí | Correo para envío de informes |
| Nombre del solicitante | Sí | Persona que llena el formulario |
| Giro | Sí | Giro o actividad económica |
| Registro patronal | No | Número de registro ante el IMSS |
| Capacidad instalada | No | Capacidad de producción instalada |
| Capacidad de operación | No | Capacidad de producción real |
| Días/turnos/horarios | No | Jornada laboral |
| Aplica NOM-020 | Sí | Sí / No |
| Requiere PIPC | Sí | Sí / No |
| Responsable del área | No | Encargado del área de seguridad |
| Teléfono del responsable | No | Teléfono del responsable |
| Nombre para informe | No | Nombre al que va dirigido el informe |
| Puesto para informe | No | Cargo de quien recibe el informe |
| Asesor/Consultor | No | Intermediario que refiere al cliente |

#### Flujo de SEAPD

```
1. Usuario llena el formulario
2. Frontend valida campos requeridos
3. Se genera token reCAPTCHA v3
4. Se envía payload a fase1_RegistrarCliente()
5. Backend crea carpeta en Drive: /ROOT/RFC_EMPRESA/Sucursal/
6. Backend agrega fila en CLIENTES_MAESTRO
7. Backend envía correo de notificación al equipo interno
8. Frontend muestra confirmación con link a Drive
```

#### Búsqueda de cliente existente

El formulario incluye un modal de búsqueda por RFC que pre-llena los campos si el cliente ya está registrado. Esto permite actualizar datos sin crear duplicados.

---

### 3.2 SEAOT — Órdenes de Trabajo

**Archivo:** `SEAOT.html`
**Acceso:** Usuarios internos con Google Auth
**Función:** Creación y gestión de Órdenes de Trabajo (OT)

#### Tipos de Orden de Trabajo

| Tipo | Descripción |
|---|---|
| OTA | Orden de Trabajo Ambiental (estándar) |
| OTB | Orden de Trabajo de Brigada (campo) |

#### Campos de la Orden de Trabajo

| Campo | Requerido | Descripción |
|---|---|---|
| Folio OT | Sí | Identificador único (ej. EA-2026-001) |
| Tipo de orden | Sí | OTA / OTB |
| Servicio NOM | Sí | Norma a aplicar (NOM-035, NOM-036, etc.) |
| Cliente (razón social) | Sí | Nombre del cliente |
| Sucursal | Sí | Planta o sucursal |
| RFC | Sí | RFC del cliente |
| Personal asignado | Sí | Técnico o consultor responsable |
| Fecha de visita | Sí | Fecha de la visita de campo |
| Fecha límite de entrega | Sí | Fecha máxima de entrega del informe |
| Link Drive cliente | Auto | Se obtiene de la búsqueda por RFC |
| Observaciones | No | Notas adicionales para el equipo |

#### Flujo de SEAOT

```
1. Operador busca al cliente por RFC o nombre
2. Sistema pre-llena datos del cliente desde CLIENTES_MAESTRO
3. Operador completa los campos de la OT
4. Se envía a fase2_RegistrarOT()
5. Backend crea fila en ORDENES_TRABAJO con estatus inicial "NO INICIADO"
6. El link_drive_cliente se guarda en la OT para referencia posterior
```

#### Carga desde Excel

SEAOT permite cargar datos del cliente desde un archivo Excel (.xlsx). El sistema mapea las columnas del archivo a los campos del formulario.

#### Documentación y Formatos Asociados

Desde la barra **Gestión de Orden** (botón **📄 Formatos asociados**), SEAOT abre un modal que vincula los formatos del proceso a la OT activa y los abre ya contextualizados —con el folio y los datos del cliente precargados— para eliminar la recaptura manual al volver de campo.

**Formatos integrados** (viven en el mismo repo/origen que SEAOT: `registro-equipos.html` y `supervision-gabinete.html`)

| Formato | Estado | Comportamiento al abrir desde la OT |
|---|---|---|
| Solicitud de Equipos (`registro-equipos.html`) | Activo | Precarga N° de OT, **Solicitante** (= personal asignado de la OT), **Fecha de Salida** (= fecha del servicio de la OT), **Área/Departamento** (según la norma: `-STPS` → Ambiente Laboral; `SEMARNAT`/NOM-081 → Fuentes Fijas), fija el motivo en «Trabajo de Campo» y **preselecciona los equipos según las NOM del servicio** (acompañantes por la lógica de grupos). Cada fila nace con Estado «Bueno» y sus checkboxes de componentes marcados como conformes. Exporta a PDF con el botón existente. |
| Supervisión de Gabinete (`supervision-gabinete.html`) | Activo | Hereda el N° de OT y la razón social; Fecha y Folio los completa el revisor. Descarga en PDF con «Imprimir». |
| Encuesta de Satisfacción | Pendiente | Diferida a una segunda iteración cuando exista el formato. |

**Transferencia de datos (handoff mismo-origen, sin PII en la URL)**

Los datos del cliente (RFC, dirección, teléfono, correo) son PII y **no viajan en la URL** (evitando historial, encabezado `Referer` y logs). En su lugar SEAOT y los formatos —al estar bajo el mismo origen (GitHub Pages `/SEA/`)— usan un *handoff* vía `sessionStorage`, cuyo contenido **nunca se escribe en disco ni sobrevive a la pestaña**:

1. SEAOT guarda el contexto de la OT en `sessionStorage` bajo la clave `sea_ot_handoff_<token>`, con un token opaco de un solo uso.
2. Abre el formato **sin `noopener`**: el navegador **clona** el `sessionStorage` a la pestaña nueva. En la URL viaja **solo** el token: `registro-equipos.html?h=<token>`.
3. La pestaña del formato lee su clon, **lo borra** (uso único) y prellena; **SEAOT elimina su propia copia de inmediato**.

Así el borrado **no depende de temporizadores ni de que SEAOT siga abierto**: la copia de SEAOT se elimina al instante y el clon del formato desaparece al cerrarse esa pestaña (nunca persiste). Si `sessionStorage` no está disponible, el formato se abre sin precarga: nunca se degrada a exponer PII en la URL.

Payload disponible en el handoff: `ot`, `serie`, `cliente`, `sucursal`, `rfc`, `direccion`, `contacto`, `telefono`, `correo`, `emision`, `servicios`, `personal`, `fecha`.

El destino de los formatos se controla con la constante `FORMATOS_BASE` en `SEAOT.html` (por defecto vacía = mismo directorio y origen que SEAOT).

---

### 3.3 SEAINF — Gestión de Expedientes

**Archivo:** `SEAINF.html`
**Acceso:** Usuarios internos con Google Auth
**Función:** Creación de expedientes y carga de documentos

#### Flujo de SEAINF (pasos secuenciales bloqueados)

```
Paso 1: Seleccionar OT
        ↓ (desbloquea Paso 2)
Paso 2: Asignar número de informe (consecutivo automático EA-AAMM-NOM-0000)
        ↓ (desbloquea Paso 3)
Paso 3: Crear estructura de carpetas en Drive
        ↓ (desbloquea Paso 4)
Paso 4: Subir documentos a las carpetas correspondientes
        ↓
Paso 5: Confirmar expediente completo
```

#### Formato del número de informe

```
EA-AAMM-NOMXXX-0000
│   │    │      └── Consecutivo 4 dígitos (por tipo de OT)
│   │    └───────── Código del servicio NOM
│   └────────────── Año (2 dígitos) + Mes (2 dígitos)
└────────────────── Prefijo Ejecutiva Ambiental
```

**Ejemplo:** `EA-2603-NOM035-0042`

#### Estructura de carpetas del expediente

```
/Carpeta_Cliente_(RFC)/
  /Sucursal/
    /02_Expediente_0042_OT-001_NOM035/
      /1. ORDEN_TRABAJO/       ← OT, contratos
      /2. HDC/                  ← Hojas de Campo
      /3. CROQUIS/              ← Planos y croquis
      /4. FOTOS/                ← Fotografías de campo
      /5. INFORMES Y MEMORIAS/  ← Informes y memorias técnicas
      /6. INFORME PRELIMINAR/   ← Versiones preliminares
```

#### Cadena de búsqueda de carpeta del cliente

El backend exige una coincidencia exacta de **RFC + sucursal**:

1. **CLIENTES_MAESTRO** — enlace exacto del RFC y la sucursal.
2. **ORDENES_TRABAJO, columna M** — se acepta únicamente si el nombre de la carpeta y su padre coinciden con la sucursal y el RFC.
3. **Ruta exacta en Drive** — búsqueda por padre con prefijo RFC y nombre exacto de sucursal.

Si ninguna ruta coincide, la operación se detiene sin crear carpetas. El sistema nunca elige otra sucursal ni utiliza la carpeta raíz como fallback. La comparación de nombres aplica la misma normalización segura usada al crear las carpetas, para tolerar caracteres que Drive haya sanitizado.

Para clientes con el centinela `SIN_RFC`, la carpeta padre solo se reutiliza cuando también coincide exactamente la razón social; `SIN_RFC` nunca se considera una identidad compartida entre empresas. Cuando una fila histórica no tiene enlace, SEAOT puede resolver la ruta exacta RFC + sucursal en modo de solo lectura, sin escribir ni completar celdas en el Spreadsheet.

---

### 3.4 SEADB — Dashboard y Control

**Archivo:** `SEADB.html`
**Acceso:** Usuarios internos con Google Auth
**Función:** Visualización del estado de todas las órdenes, reportes y estadísticas

#### Pestañas del dashboard

| Pestaña | Contenido |
|---|---|
| Resumen | KPIs principales: total OTs, en proceso, entregadas, vencidas |
| Órdenes | Tabla filtrable con todas las OTs, semáforos de estatus |
| Reportes | Gráficas y estadísticas de cumplimiento |
| Calendario | Vista mensual de fechas de visita y entrega |

#### Sistema de semáforos

| Color | Significado |
|---|---|
| Verde | OT entregada a tiempo o con margen suficiente |
| Amarillo | OT próxima a vencer (< 3 días) |
| Rojo | OT vencida (fecha límite superada sin entrega) |

#### Cálculo de fechas de entrega

- **Fecha Entrega Digital**: fecha límite registrada en ORDENES_TRABAJO
- **Fecha Entrega Físico**: Fecha Digital + 5 días hábiles
- El semáforo se calcula respecto a la fecha actual vs. la fecha límite

#### Estatus externos (visibles al cliente)

| Estatus | Descripción |
|---|---|
| NO INICIADO | La OT está creada pero no se ha comenzado |
| EN PROCESO | Se está trabajando en el informe |
| ENTREGADO | Informe digital entregado |
| FINALIZADO | Informe físico y digital entregados |
| CANCELADO | OT cancelada |

---

### 3.5 PAIC — Portal de Asesores y Consultores

**Archivo:** `PAIC.html`
**Acceso:** Portal público con reCAPTCHA v3
**Función:** Portal externo para intermediarios y consultores que refieren clientes

#### Características

- Formulario de registro de estudios (1 estudio = 1 registro)
- Búsqueda de cliente existente por RFC para pre-llenar datos
- Subida de documentos con validación de tipo y tamaño
- Detección de cliente con expediente previo
- Integración con `recaptcha.js` para protección contra bots

---

## 4. Base de Datos (Google Sheets)

**ID del Spreadsheet:** `1MoScea4CYg0NCjvPjHqZwV0cKhrd2nxfW8LYhz_4pDo`

### 4.1 Hoja CLIENTES_MAESTRO

| # Col | Índice | Campo | Descripción |
|---|---|---|---|
| 1 | 0 | Fecha Registro | Fecha de alta del cliente |
| 2 | 1 | Razón Social | Nombre legal de la empresa |
| 3 | 2 | Sucursal | Nombre de la sucursal o planta |
| 4 | 3 | RFC | RFC de 13 caracteres |
| 5 | 4 | Representante Legal | Nombre del representante |
| 6 | 5 | Dirección Evaluación | Domicilio del servicio |
| 7 | 6 | Teléfono Empresa | Teléfono de contacto |
| 8 | 7 | Nombre Solicitante | Persona que registró |
| 9 | 8 | Correo Informe | Correo para envío de informes |
| 10 | 9 | Giro | Actividad económica |
| 11 | 10 | Registro Patronal | Número IMSS |
| 12 | 11 | Cap. Instalada | Capacidad instalada |
| 13 | 12 | Cap. Operación | Capacidad de operación |
| 14 | 13 | Días/Turnos/Horarios | Jornada laboral |
| 15 | 14 | Aplica NOM-020 | SÍ / NO |
| 16 | 15 | Requiere PIPC | SÍ / NO |
| 17 | 16 | Responsable | Responsable del área |
| 18 | 17 | Teléfono Responsable | Teléfono del responsable |
| 19 | 18 | Nombre Dirigido | Destinatario del informe |
| 20 | 19 | Puesto Dirigido | Cargo del destinatario |
| 21 | **20** | **Link Drive** | **URL de la carpeta en Drive** |
| 22 | 21 | Asesor/Consultor | Intermediario que refirió |

> **Importante:** El link de Drive está en el índice **20** (columna 21). Este índice es crítico para el funcionamiento de toda la cadena de búsqueda de carpetas.

### 4.2 Hoja ORDENES_TRABAJO

> **Contrato inmutable:** el sistema usa exactamente las columnas A–Q existentes. El código no agrega, elimina, renombra ni reordena columnas y no migra información histórica. Las validaciones de código solamente comprueban que esa estructura ya exista.

| # Col | Índice | Campo | Descripción |
|---|---|---|---|
| 1 | 0 | Fecha | Fecha de creación de la OT |
| 2 | 1 | OT (Folio) | Identificador único |
| 3 | 2 | Tipo | OTA / OTB |
| 4 | 3 | Servicio NOM | Norma aplicada |
| 5 | 4 | Cliente | Razón social |
| 6 | 5 | Sucursal | Planta o sucursal |
| 7 | 6 | RFC | RFC del cliente |
| 8 | 7 | Personal Asignado | Técnico responsable |
| 9 | 8 | Fecha Visita | Fecha de visita de campo |
| 10 | 9 | Fecha Entrega Límite | Fecha máxima de entrega |
| 11 | 10 | Fecha Real Entrega | Fecha real de entrega |
| 12 | 11 | Estatus Externo | Estado visible en SEADB |
| 13 | **12** | **Link Drive** | **URL de la carpeta de la sucursal** |
| 14 | 13 | Observaciones | Notas adicionales |
| 15 | 14 | Fecha Pausa | Fecha en que se pausó la OT |
| 16 | 15 | Motivo Pausa | Motivo registrado para la pausa |
| 17 | 16 | Fecha Info Completa | Fecha de recepción de la información completa |

Los datos propios del expediente —número de informe, estatus interno y enlace del expediente— se almacenan en la hoja **INFORMES**, sin modificar la estructura A–Q de ORDENES_TRABAJO.

### 4.3 Hoja USUARIOS_AUTORIZADOS

| Campo | Tipo | Descripción |
|---|---|---|
| Email | Texto | Correo de Google del usuario |
| Nombre | Texto | Nombre completo |
| Rol | Texto | admin / operador / auxiliar_operador |
| Activo | Boolean | TRUE = activo, FALSE = bloqueado |
| Fecha Alta | Fecha | Fecha de registro |
| SEADB | Boolean | TRUE = acceso al dashboard |
| SEAOT | Boolean | TRUE = acceso a órdenes de trabajo |
| SEAINF | Boolean | TRUE = acceso a expedientes |
| Notas | Texto | Observaciones internas |

### 4.4 Hoja AUDITORIA

Registro automático de todos los cambios de estatus. Se crea automáticamente al primer cambio. Campos: Timestamp, Usuario, Acción, OT, Campo, Valor Anterior, Valor Nuevo.

---

## 5. Estructura de Carpetas en Drive

**ID de Carpeta Raíz:** `1nHd-70uUeciClDm_3_pgbmqGF7II1lfQ`

```
/SEA_ROOT/
  /XTES000000TST — EMPRESA EJEMPLO SA DE CV/   ← una por RFC único
    /Matriz/                                      ← una por sucursal
      /01_Cliente/                                ← datos del cliente
      /02_Expediente_0001_EA-001_NOM035/          ← por OT+informe
        /1. ORDEN_TRABAJO/
        /2. HDC/
        /3. CROQUIS/
        /4. FOTOS/
        /5. INFORMES Y MEMORIAS/
        /6. INFORME PRELIMINAR/
    /Planta Norte/
      /02_Expediente_0002_.../
```

### Convenciones de nombres

| Elemento | Formato |
|---|---|
| Carpeta raíz del cliente | `{RFC} — {RAZON_SOCIAL_LIMPIA}` |
| Carpeta de sucursal | `{NOMBRE_SUCURSAL}` |
| Carpeta de expediente | `02_Expediente_{CONSECUTIVO}_{FOLIO_OT}_{NOM}` |

La función `cleanCompanyName()` elimina sufijos legales (SA DE CV, S.A., S.C., etc.) del nombre para mantener los nombres de carpeta cortos y legibles. La función `sanitizeFileName()` elimina caracteres especiales y trunca a 50 caracteres.

---

## 6. API Backend — Referencia de Funciones

Todas las funciones del backend se invocan vía `google.script.run` desde el frontend, o directamente desde el editor de GAS para pruebas.

### 6.1 Fase 1 — Registro de Clientes

#### `fase1_RegistrarCliente(payload)`

Registra un nuevo cliente o actualiza datos existentes.

**Payload:**
```javascript
{
  action: 'registrarCliente',
  razon_social: 'EMPRESA SA DE CV',
  sucursal: 'Matriz',
  rfc: 'EMP010101AAA',
  representante_legal: 'Juan Pérez',
  direccion_evaluacion: 'Av. Principal 123, Puebla',
  telefono_empresa: '2221234567',
  nombre_solicitante: 'María García',
  correo_informe: 'informes@empresa.com',
  giro: 'Manufactura',
  aplica_nom020: 'si',    // 'si' | 'no'
  requiere_pipc: 'no',   // 'si' | 'no'
  // Opcionales:
  registro_patronal: 'IMSS-01234',
  capacidad_instalada: '500 ton',
  capacidad_operacion: '450 ton',
  dias_turnos_horarios: 'L-V 08:00-18:00',
  responsable: 'Ing. Seguridad',
  telefono_responsable: '2229876543',
  nombre_dirigido: 'Lic. Director',
  puesto_dirigido: 'Director General',
  _skipEmail: false   // true para pruebas (omite envío de correo)
}
```

**Respuesta exitosa:**
```javascript
{ success: true, linkDrive: 'https://drive.google.com/drive/folders/...' }
```

**Respuesta de error:**
```javascript
{ success: false, error: 'Descripción del error' }
```

---

### 6.2 Fase 2 — Órdenes de Trabajo

#### `fase2_BuscarClienteRFC(rfc)`

Busca un cliente por RFC y devuelve todas sus sucursales registradas.

**Respuesta exitosa:**
```javascript
{
  found: true,
  sucursales: [{
    razon_social: 'EMPRESA SA DE CV',
    sucursal: 'Matriz',
    rfc: 'EMP010101AAA',
    nombre_solicitante: 'María García',
    correo_informe: 'informes@empresa.com',
    telefono_empresa: '2221234567',
    representante_legal: 'Juan Pérez',
    link_drive_cliente: 'https://drive.google.com/drive/folders/...'
  }]
}
```

**No encontrado:**
```javascript
{ found: false }
```

---

#### `fase2_BuscarClienteNombre(nombreBuscado)`

Busca clientes por coincidencia parcial en razón social (mínimo 3 caracteres).

**Respuesta exitosa:**
```javascript
{
  found: true,
  resultados: [{
    razon_social: '...',
    sucursal: '...',
    rfc: '...',
    link_drive_cliente: '...'
    // ... mismos campos que buscarClienteRFC
  }]
}
```

---

#### `fase2_RegistrarOT(data)`

Crea una nueva Orden de Trabajo en ORDENES_TRABAJO.

**Payload:**
```javascript
{
  action: 'registrarOT',
  ot_folio: 'EA-2026-001',
  tipo_orden: 'OTA',           // 'OTA' | 'OTB'
  nom_servicio: 'NOM-035-STPS',
  cliente_razon_social: 'EMPRESA SA DE CV',
  sucursal: 'Matriz',
  rfc: 'EMP010101AAA',
  personal_asignado: 'Ing. Técnico',
  fecha_visita: '2026-03-20',
  fecha_entrega_limite: '2026-03-27',
  link_drive_cliente: 'https://drive.google.com/drive/folders/...',
  observaciones: 'Sin observaciones'
}
```

**Respuesta:**
```javascript
{ success: true, message: 'OT Registrada correctamente' }
```

---

### 6.3 Fase 3 — Expedientes

#### `fase3_CrearExpediente(payload)`

Crea de forma idempotente la carpeta del expediente en Drive y registra el expediente en **INFORMES**. No altera la estructura de ORDENES_TRABAJO.

**Payload:**
```javascript
{
  action: 'createExpediente',
  data: {
    ot: 'EA-2026-001',
    nom: 'NOM035',
    numInforme: 'EA-2603-NOM035-0001',
    cliente: 'EMPRESA SA DE CV',
    sucursal: 'Matriz',
    rfc: 'EMP010101AAA',
    linkDrive: 'https://drive.google.com/drive/folders/...',
    fecha: '20/03/2026',
    entrega: '27/03/2026',
    tipoOrden: 'OTA',
    solicitante: 'María García',
    telefono: '2221234567',
    direccion: 'Av. Principal 123',
    responsable: 'Ing. Técnico',
    estatus: 'NO INICIADO'
  },
  files: [
    {
      content: '<base64>',   // contenido en Base64
      type: 'application/pdf',
      name: 'orden_trabajo.pdf',
      category: 'ORDEN_TRABAJO'  // ORDEN_TRABAJO|HOJAS_CAMPO|CROQUIS|FOTOS|INFORMES_MEMORIA|INF_PRELIMINAR|PERFIL_DATOS
    }
  ]
}
```

**Respuesta:**
```javascript
{
  success: true,
  url: 'https://drive.google.com/drive/folders/...',
  alreadyExists: false,
  uploadedFiles: 3,
  rejectedFiles: [] // puede contener archivos rechazados si los demás sí se guardaron
}
```

La operación es idempotente por OT. Si el expediente ya existe, el backend reutiliza la carpeta registrada y carga únicamente los archivos recibidos que aún falten. Los archivos válidos no se pierden porque otro archivo sea rechazado; la interfaz muestra la advertencia y conserva el formulario para reemplazar los rechazados y reintentar sobre el mismo expediente.

---

#### `fase3_AddFilesToExpediente(payload)`

Agrega archivos a un expediente existente. La carpeta autorizada siempre se obtiene del registro de la OT en **INFORMES**. Si el cliente envía `expedienteUrl`, su ID debe coincidir con ese registro; una URL de otro expediente se rechaza antes de escribir en Drive.

**Payload:**
```javascript
{
  ot: 'EA-2026-001',
  expedienteUrl: 'https://drive.google.com/drive/folders/...',
  files: [ /* igual que en createExpediente */ ]
}
```

---

### 6.4 Funciones Internas (Safe)

Estas funciones son invocadas desde `doPost()` después de validar autenticación:

| Función | Descripción |
|---|---|
| `getOrdenesSafe_()` | Devuelve todas las OTs activas (excluye FINALIZADO/CANCELADO) |
| `getConsecutivoSafe_(params)` | Genera el siguiente número de informe consecutivo |
| `updateEstatusSafe_(data, usuario)` | Actualiza el estatus externo de una OT |
| `updateEstatusInformeSafe_(data, usuario)` | Actualiza el estatus interno del informe |

---

## 7. Autenticación y Seguridad

### 7.1 Modos de autenticación

| Modo | Descripción | Usado por |
|---|---|---|
| `GOOGLE` | Requiere `id_token` válido + usuario en whitelist | Módulos internos |
| `RECAPTCHA` | Requiere `recaptcha_token` válido (score ≥ 0.5) | Portales públicos |
| `EITHER` | Acepta Google Auth o reCAPTCHA | `buscarClienteRFC/Nombre` |

### 7.2 Tabla de modos por acción

| Acción | Modo | Módulo |
|---|---|---|
| `registrarCliente` | RECAPTCHA | SEAPD / PAIC |
| `buscarClienteRFC` | EITHER | SEAOT / PAIC |
| `buscarClienteNombre` | EITHER | SEAOT / PAIC |
| `registrarOT` | GOOGLE | SEAOT |
| `resolverCarpetaCliente` | GOOGLE | SEAOT |
| `getOrdenes` | GOOGLE | SEAINF |
| `getConsecutivo` | GOOGLE | SEAINF |
| `createExpediente` | GOOGLE | SEAINF |
| `addFilesToExpediente` | GOOGLE | SEAINF |
| `updateEstatusInforme` | GOOGLE | SEAINF |
| `getTablero` | GOOGLE | SEADB |
| `updateEstatus` | GOOGLE | SEADB |
| `updateResponsable` | GOOGLE | SEADB |
| `verificarAcceso` | GOOGLE | Todos los internos |

### 7.3 Archivos de autenticación frontend

**`auth.js`** — Para módulos internos (SEAPD, SEAOT, SEAINF, SEADB, index.html)
- Gestiona el flujo OAuth 2.0 de Google
- Almacena el `id_token` en `sessionStorage`
- Refresca el token automáticamente 60 segundos antes de expirar
- Muestra overlay de login animado si el usuario no está autenticado
- Verifica que el usuario esté en la whitelist llamando a `verificarAcceso`

**`recaptcha.js`** — Para portales públicos (PAIC, SEAPD)
- Genera token invisible reCAPTCHA v3 por cada solicitud
- Inyecta el token en los headers de cada petición fetch
- `siteKey` en el frontend; `secretKey` en GAS Script Properties (nunca en el código)

### 7.4 Validación de archivos

El backend valida todos los archivos subidos antes de guardarlos en Drive:

- Tipos permitidos: PDF, JPG/JPEG, PNG, HEIC/HEIF, XLS/XLSX, DOC/DOCX, ZIP/RAR, DWG/DXF y MP4/MOV.
- Tamaño máximo: 50 MB por archivo.
- Los nombres se sanitizan conservando la extensión y los guiones, con un máximo de 180 caracteres.
- Si un nombre ya existe con contenido distinto, se crea una versión `(v2)`, `(v3)`, etc.; los reintentos idénticos no duplican el archivo.
- Cuando un envío mezcla archivos válidos y rechazados, se guardan los válidos y la respuesta enumera los rechazados. SEAINF muestra la advertencia y conserva el formulario para completar el mismo expediente.

---

## 8. Configuración e Instalación

### 8.1 Prerrequisitos

- Cuenta de Google Workspace (dominio `ejecutivambiental.com`)
- Acceso al Spreadsheet principal
- Acceso a la carpeta raíz en Drive
- Editor de Google Apps Script (script.google.com)

### 8.2 Configuración inicial (una sola vez)

**Paso 1: Configurar Script Properties en GAS**

En el editor de GAS → Proyecto → Propiedades del proyecto → Propiedades de script:

```
GOOGLE_CLIENT_ID     = 407541868250-5pbtl3me85quu1nl38b1c57ebi3nn9a6.apps.googleusercontent.com
RECAPTCHA_SECRET_KEY = <clave secreta de reCAPTCHA v3>
```

**Paso 2: Crear hojas del Spreadsheet**

```javascript
// En el editor GAS, ejecutar una sola vez:
setupSheets()
```

Esto crea las hojas `CLIENTES_MAESTRO`, `ORDENES_TRABAJO`, `USUARIOS_AUTORIZADOS` y `AUDITORIA` si no existen.

**Paso 3: Crear usuarios autorizados**

```javascript
// En el editor GAS, ejecutar una sola vez:
crearHojaUsuarios()
```

Esto crea la hoja `USUARIOS_AUTORIZADOS` con los 4 usuarios iniciales. Después puedes agregar/editar usuarios directamente en la hoja.

**Paso 4: Configurar respaldo automático**

```javascript
// En el editor GAS, ejecutar una sola vez:
configurarTriggerRespaldo()
```

Configura el trigger para ejecutar `crearRespaldoSemanal()` cada lunes a las 3:00 AM GMT-6.

### 8.3 Actualizar IDs en CONFIG

Si se cambia el Spreadsheet o la carpeta raíz en Drive, actualizar en `BACKEND_FIXES.gs`:

```javascript
const CONFIG = {
  SPREADSHEET_ID: 'nuevo_id_del_spreadsheet',
  FOLDER_ID:      'nuevo_id_de_la_carpeta_raiz',
  // ...
};
```

### 8.4 Desplegar como Web App

En el editor GAS → Implementar → Nueva implementación → Aplicación web:
- Ejecutar como: `Yo (ejecutiva@gmail.com)`
- Quién tiene acceso: `Cualquier usuario`

Copiar la URL de la implementación y usarla como endpoint en el frontend.

---

## 9. Sistema de Respaldos

**Archivo:** `RESPALDO.gs`

### 9.1 Configuración

| Parámetro | Valor |
|---|---|
| Frecuencia | Semanal (lunes 3:00 AM GMT-6) |
| Retención | 12 copias (≈ 3 meses) |
| Notificaciones | Email en éxito y falla |

### 9.2 Funciones disponibles

| Función | Descripción |
|---|---|
| `configurarTriggerRespaldo()` | Crea el trigger automático (ejecutar una vez) |
| `crearRespaldoSemanal()` | Ejecuta el respaldo manualmente |
| `crearCarpetaRespaldos()` | Crea la estructura de carpetas para respaldos |
| `verificarEstadoRespaldos()` | Muestra el estado actual de los respaldos |

### 9.3 Ejecutar respaldo manual

```javascript
// En el editor GAS:
crearRespaldoSemanal()
```

El respaldo crea una copia del Spreadsheet en la carpeta de respaldos con formato:
`SEA_Respaldo_YYYY-MM-DD_HH-mm`

Cuando se supera el máximo de 12 copias, el respaldo más antiguo se mueve automáticamente a la papelera.

---

## 10. Gestión de Usuarios

### 10.1 Agregar un usuario nuevo

1. Abrir el Spreadsheet principal
2. Ir a la hoja `USUARIOS_AUTORIZADOS`
3. Agregar una fila con:
   - Email: correo de Google del usuario
   - Nombre: nombre completo
   - Rol: `admin`, `operador` o `auxiliar_operador`
   - Activo: `TRUE`
   - Fecha Alta: fecha de hoy
   - SEADB, SEAOT, SEAINF: `TRUE` para dar acceso, `FALSE` para restringir

> El sistema actualiza los permisos en un máximo de 10 minutos (cache de GAS).

### 10.2 Desactivar un usuario

Cambiar la celda `Activo` de `TRUE` a `FALSE`. El sistema bloqueará el acceso en la próxima solicitud y enviará una alerta al correo de dirección general si el usuario intenta acceder.

### 10.3 Roles disponibles

| Rol | Descripción |
|---|---|
| `admin` | Acceso completo, puede gestionar usuarios |
| `operador` | Acceso a módulos asignados (SEADB, SEAOT, SEAINF) |
| `auxiliar_operador` | Acceso limitado a módulos específicos |

### 10.4 Usuarios iniciales del sistema

| Email | Nombre | Rol |
|---|---|---|
| `eduwin.ejecutiva@gmail.com` | Administrador | admin |
| `aclientes.ejecutiva@gmail.com` | Operador A Clientes | operador |
| `operaciones.ejecutivamx@gmail.com` | Operador Operaciones | operador |
| `calidad.ejecutivamx@gmail.com` | Aux. Operador Calidad | auxiliar_operador |

---

## 11. Pruebas Automatizadas

**Archivo:** `TESTS_BACKEND.gs`

### 11.1 Tipos de prueba

| Tipo | Descripción |
|---|---|
| Unitarias | 74 comprobaciones de lógica pura, sin Drive ni Sheets |
| E2E | Pruebas completas contra recursos exclusivos de staging |
| Correo | Prueba manual de envío; no forma parte del runner E2E |

### 11.2 Seguridad obligatoria para E2E

Las E2E se niegan a comenzar si no existen estas Script Properties:

| Propiedad | Valor |
|---|---|
| `SEA_E2E_ENABLED` | `TRUE` |
| `SEA_TEST_SPREADSHEET_ID` | ID de una copia de staging del Spreadsheet |
| `SEA_TEST_FOLDER_ID` | ID de una carpeta raíz exclusiva de staging |

El Spreadsheet de staging debe conservar los contratos `CLIENTES_MAESTRO` A–V, `ORDENES_TRABAJO` A–Q e `INFORMES` A–Q. El runner aborta si falta una hoja, si el ancho es menor o si alguno de los IDs de staging coincide con producción.

> Las E2E nunca deben apuntar al Spreadsheet ni a la carpeta Drive productivos. `runUnitTests()` no requiere ninguna propiedad de staging y no produce efectos secundarios.

### 11.3 Pruebas E2E disponibles

| ID | Módulo | Descripción |
|---|---|---|
| E01 | SEAPD | Registro de cliente, `01_Cliente`, carpeta Drive y fila CLIENTES |
| E02 | SEAOT | Búsqueda por RFC y creación de OT |
| E03 | SEAINF | Creación completa del expediente y sus seis subcarpetas |
| E04 | SEAOT | Búsqueda positiva por nombre |
| E05 | SEAOT | RFC inexistente |
| E06 | SEAOT | Validación de nombre demasiado corto |
| E07 | SEAOT | Registro de OT tipo OTB |
| E08 | SEADB | Cambio de estatus externo y fecha real |
| E09 | SEAINF | Cambio de estatus interno |
| E10 | SEAINF | Reintento idempotente: misma URL, mismo archivo omitido y una sola fila INFORMES |
| E11 | Drive | Mismo nombre con contenido distinto crea `(v2)` conservando el original |
| E12 | Seguridad | Una URL perteneciente a otra OT no recibe archivos |
| E13 | SEAINF | Lote parcial guarda válidos y reporta el rechazado |
| E14 | SEAOT | Cliente histórico sin enlace se resuelve sin escribir la celda |
| E15 | Seguridad | RFC+sucursal inexistentes no crean carpeta ni fila INFORMES |
| E16 | SEAOT | Enlace vacío o inválido bloquea el alta de la OT |
| E17 | Drive | Expediente legado de cuatro subcarpetas se completa a seis sin duplicados |

### 11.4 Cómo ejecutar

**Solo unitarias, sin efectos secundarios:**

1. Seleccionar `runUnitTests`.
2. Ejecutar y revisar el registro: el resultado esperado es `74 PASS | 0 FAIL`.

**E2E completas, únicamente después de configurar staging:**

1. Confirmar las tres Script Properties de staging.
2. Seleccionar `runE2ETests`.
3. Ejecutar y revisar los resultados E01–E17.
4. El runner realiza limpieza previa y final. Elimina filas de prueba en CLIENTES, ORDENES, INFORMES y AUDITORIA, y envía sus carpetas de staging —incluida `01_Cliente`— a la papelera.

Cada `runTest_E01` … `runTest_E17` vuelve a comprobar el guard de staging cuando se ejecuta individualmente.

### 11.5 Datos reservados para pruebas

| Dato | Valor principal |
|---|---|
| RFC | `XTES000000TST` |
| Folios | `TEST-E2E-001`, `TEST-E2E-002` y auxiliares `TEST-E2E-*` |
| Sucursal | `Sucursal Test E2E` |
| Empresa | `EMPRESA TEST E2E SA DE CV` |

### 11.6 Verificaciones manuales previas al despliegue

Estas comprobaciones son de solo lectura salvo la prueba de humo, que debe realizarse exclusivamente en staging:

- Inventariar filas de CLIENTES cuyo Link Drive no contenga `/folders/`.
- Comparar nombres de sucursal en Drive con la normalización de `sanitizeFileName()`.
- Confirmar que ORDENES_TRABAJO e INFORMES llegan como mínimo hasta la columna Q y que INFORMES existe.
- Probar en staging un MP4 real de teléfono cercano al límite permitido.
- Cronometrar en staging el alta de una OT histórica sin enlace para establecer una línea base de rendimiento.

Poblar enlaces históricos o renombrar carpetas sería una migración separada y no forma parte de este PR.

---

## 12. Solución de Problemas

### Error: "No se encontró una carpeta exacta para el RFC y la sucursal"

**Causa:** El cliente/sucursal no está registrado en SEAPD, el enlace de Drive no es accesible o no coincide exactamente con el RFC y la sucursal de la OT.

**Solución:**
1. Verificar en SEAPD que exista el RFC con la sucursal correcta.
2. Volver a seleccionar el cliente en SEAOT para validar su carpeta exacta.
3. No editar, insertar ni reordenar columnas manualmente en el Excel productivo.
4. El backend detiene la operación; no crea expedientes en otra sucursal ni en la raíz.

---

### Error: "OT no encontrada" al crear expediente

**Causa:** El folio de la OT enviado por SEAINF no coincide con ninguna fila en ORDENES_TRABAJO.

**Solución:**
1. Verificar que el folio esté escrito exactamente igual (mayúsculas, sin espacios extra)
2. Refrescar la lista de OTs en SEAINF (salir y volver a entrar)

---

### Error: "Token rechazado: aud=..."

**Causa:** El `id_token` de Google fue emitido para una aplicación diferente.

**Solución:**
1. Verificar que `GOOGLE_CLIENT_ID` en Script Properties coincida con el Client ID en `auth.js`
2. Limpiar el `sessionStorage` del navegador y volver a autenticar

---

### El usuario no puede acceder aunque está en la lista

**Causa:** El cache de GAS (10 minutos) aún tiene los permisos anteriores.

**Solución:** Esperar hasta 10 minutos o limpiar la cache del script:
```javascript
// En el editor GAS:
CacheService.getScriptCache().removeAll([]);
```

---

### Error al enviar correo: "Límite de cuota excedido"

**Causa:** Google Apps Script tiene un límite de 100 correos/día para cuentas gratuitas.

**Solución:**
- Con Google Workspace Business: 1,500 correos/día
- Verificar cuota en: GAS → Ver → Cuotas

---

### El test falla con "Error cleanup carpeta raíz"

**Causa:** La carpeta de prueba en Drive ya fue eliminada en una ejecución anterior.

**Solución:** Es un error no crítico. El test de limpieza intenta borrar lo que puede. Si persiste, buscar y eliminar manualmente carpetas con `XTES000000TST` en el nombre dentro de la carpeta raíz de Drive.

---

## 13. Glosario

| Término | Definición |
|---|---|
| **OT** | Orden de Trabajo |
| **OTA** | Orden de Trabajo Ambiental (estudio en campo) |
| **OTB** | Orden de Trabajo de Brigada |
| **RFC** | Registro Federal de Contribuyentes |
| **HDC** | Hoja de Campo |
| **PIPC** | Programa Interno de Protección Civil |
| **NOM** | Norma Oficial Mexicana |
| **GAS** | Google Apps Script |
| **CLIENTES_MAESTRO** | Hoja de Google Sheets con el registro maestro de clientes |
| **ORDENES_TRABAJO** | Hoja de Google Sheets con todas las OTs |
| **Expediente** | Carpeta en Drive con todos los documentos de una OT |
| **Sucursal** | Planta o ubicación de un cliente (un RFC puede tener varias) |
| **Folio** | Identificador único de una OT o informe |
| **Consecutivo** | Número secuencial del informe (ej. 0042) |
| **Semáforo** | Indicador visual de color en SEADB según urgencia |
| **Whitelist** | Lista de usuarios autorizados en `USUARIOS_AUTORIZADOS` |
| **reCAPTCHA** | Sistema de verificación anti-bots de Google |
| **JWT** | JSON Web Token (formato del `id_token` de Google) |
| **PAIC** | Portal para Asesores, Intermediarios y Consultores |
| **SEADB** | Módulo Dashboard de SEA |
| **SEAOT** | Módulo de Órdenes de Trabajo de SEA |
| **SEAINF** | Módulo de Informes/Expedientes de SEA |
| **SEAPD** | Módulo de Perfil/Datos del cliente de SEA |

---

*Sistema SEA v3.0 — Ejecutiva Ambiental — Última actualización: Marzo 2026*
