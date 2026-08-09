# TRAZ — Módulo de Trazabilidad de Servicios

Trazabilidad **sencilla y de solo lectura** para Ejecutiva Ambiental:

```
ORDEN DE TRABAJO  →  EXPEDIENTE (carpeta Drive)  →  INFORME(S)
```

La trazabilidad se organiza **por OT relacionada con su o sus informe(s)**.

Este módulo es **aditivo y aislado**: vive en la carpeta `TRAZ/` y **no modifica ningún
archivo ni dato del sistema SEA existente**. Solo **lee** las hojas que ya existen.

---

## Principios de diseño (por decisión del negocio)

- **Solo lectura, intocable.** `TRAZ.gs` nunca usa `setValue`, `appendRow`, `insertSheet` ni
  escribe en Drive. Solo llama `getDataRange().getDisplayValues()`.
- **Sin cambios de esquema.** No agrega columnas, hojas ni campos. Se alimenta únicamente de lo
  que ya capturan `SEAOT` y `SEAINF`.
- **`AUDITORIA` es la bitácora real.** Se muestran los eventos **realmente registrados**; no se
  infieren acciones que no estén en el log (revisión/emisión solo aparecen si fueron auditadas).
- **El expediente es `INFORMES.LINK_DRIVE`.** No se crea otro repositorio documental. El botón
  **ABRIR EXPEDIENTE** abre esa carpeta de Drive.
- **Relación OT ↔ Informe por folio de OT** (texto). Un id interno estable queda como mejora futura.
- **No incluye** memoria de cálculo ni equipos (no existen como dato estructurado en el sistema).
- **Sin autenticación (v1).** Se quitó el acceso con Google OAuth. La app web se publica abierta
  (ver aviso de seguridad abajo).

> ⚠️ **Aviso de seguridad.** Al no tener control de acceso, **cualquier persona que conozca la URL
> del despliegue puede ver los datos de trazabilidad** (clientes, RFC, folios, fechas). Como
> mitigación básica: no compartas la URL públicamente y regenera el despliegue si se filtra. Si más
> adelante quieres proteger el acceso sin OAuth, la opción más simple es un código/PIN compartido.

---

## Qué lee (fuente de datos)

Mismo Spreadsheet del SEA (`SPREADSHEET_ID` en `TRAZ.gs`), hojas:

| Hoja | Uso en TRAZ |
|---|---|
| `ORDENES_TRABAJO` | OT: folio, cliente, NOM, personal, fechas, estatus, link Drive |
| `INFORMES` | Informe(s) de la OT: folio, estatus, responsable, fechas, **link del expediente** |
| `AUDITORIA` | Bitácora real de cambios (timestamp, usuario, acción, campo, antes/después) |

---

## Archivos

| Archivo | Rol |
|---|---|
| `TRAZ.gs` | Backend de solo lectura (app web de Apps Script **independiente**). Acciones: `trazResumen`, `trazDetalle`. |
| `TRAZ.html` | Frontend estático (línea de tiempo, informes, ABRIR EXPEDIENTE, bitácora, advertencias). Sin dependencias externas. |

---

## Despliegue (una sola vez)

TRAZ se despliega como un **proyecto de Apps Script separado**, para no tocar el backend SEA
(`BACKEND_FIXES.gs`).

1. **Crear proyecto Apps Script.** En [script.google.com](https://script.google.com) → *Nuevo
   proyecto*. Pega el contenido de `TRAZ.gs`.
2. **Permisos.** La cuenta que despliega debe tener acceso de lectura al Spreadsheet
   `1MoScea4CYg0NCjvPjHqZwV0cKhrd2nxfW8LYhz_4pDo`.
3. **Desplegar como app web.** *Implementar* → *Nueva implementación* → *Aplicación web*:
   - Ejecutar como: **Yo** (dueño con acceso al Spreadsheet).
   - Quién tiene acceso: **Cualquier usuario**.
   - Autoriza los permisos que pida (lectura de Sheets) y copia la URL `.../exec`.
4. **Conectar el frontend.** En `TRAZ.html`, reemplaza:
   ```js
   const API_URL = 'REEMPLAZAR_CON_URL_DEL_DESPLIEGUE_TRAZ';
   ```
   por la URL del paso 3.
5. **Publicar el frontend.** `TRAZ.html` se sirve junto al resto (GitHub Pages). Abre
   `…/TRAZ/TRAZ.html`.

> No requiere configurar OAuth ni `USUARIOS_AUTORIZADOS`: la v1 no tiene control de acceso
> (ver aviso de seguridad arriba).

---

## Alcance de la v1 (y qué NO hace)

- **Sí:** muestra la cadena OT → Expediente → Informe(s) con datos existentes, botón ABRIR
  EXPEDIENTE, bitácora real de AUDITORIA y advertencias viables (OT sin carpeta, OT sin informe,
  informe sin expediente, informe sin fecha de ejecución, entrega faltante pese a estatus
  entregado/finalizado).
- **No:** no escribe nada, no captura memoria/equipos, no sintetiza hitos de revisión/emisión
  (solo aparecen si están en la bitácora), no agrega un id interno.

## Nota de acoplamiento

`TRAZ.gs` copia los índices de columna del esquema SEA (`TRAZ_CONFIG.COL_*`). Si en el futuro se
reordenan columnas en `ORDENES_TRABAJO` / `INFORMES` / `AUDITORIA`, hay que actualizar esos
índices aquí también.
