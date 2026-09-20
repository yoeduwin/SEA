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
- **Solo un control de acceso adicional.** TRAZ no modifica el esquema operativo de OT/INFORMES;
  utiliza la columna `SEATRAZ` en `USUARIOS_AUTORIZADOS` para controlar quién puede consultar el módulo.
- **`AUDITORIA` es la bitácora real.** Se muestran los eventos **realmente registrados**; no se
  infieren acciones que no estén en el log (revisión/emisión solo aparecen si fueron auditadas).
- **El expediente es `INFORMES.LINK_DRIVE`.** No se crea otro repositorio documental. El botón
  **ABRIR EXPEDIENTE** abre esa carpeta de Drive.
- **Relación OT ↔ Informe por folio de OT** (texto). Un id interno estable queda como mejora futura.
- **No incluye** memoria de cálculo ni equipos (no existen como dato estructurado en el sistema).
- **Autenticación Google OAuth.** Usa el mismo `auth.js` de los módulos internos SEA.
- **Autorización por módulo.** El backend valida que el usuario esté activo y tenga `SEATRAZ = TRUE`
  en `USUARIOS_AUTORIZADOS`.
- **Protección real del endpoint.** `trazResumen` y `trazDetalle` requieren un `id_token` válido;
  no es únicamente un bloqueo visual del frontend.

---

## Qué lee (fuente de datos)

Mismo Spreadsheet del SEA (`SPREADSHEET_ID` en `TRAZ.gs`), hojas:

| Hoja | Uso en TRAZ |
|---|---|
| `ORDENES_TRABAJO` | OT: folio, cliente, NOM, personal, fechas, estatus, link Drive |
| `INFORMES` | Informe(s) de la OT: folio, estatus, responsable, fechas, **link del expediente** |
| `AUDITORIA` | Bitácora real de cambios (timestamp, usuario, acción, campo, antes/después) |
| `USUARIOS_AUTORIZADOS` | Autorización de acceso mediante la columna `SEATRAZ` |

---

## Archivos

| Archivo | Rol |
|---|---|
| `TRAZ.gs` | Backend de solo lectura (app web de Apps Script **independiente**). Valida Google OAuth y expone `verificarAcceso`, `trazResumen`, `trazDetalle`. |
| `TRAZ.html` | Frontend estático protegido con el `auth.js` compartido de SEA. |

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
   - La URL puede ser pública a nivel técnico, pero **los datos no se entregan sin un token Google válido y permiso SEATRAZ**.
   - Autoriza los permisos solicitados (Sheets y verificación de token mediante UrlFetch) y copia la URL `.../exec`.
4. **Conectar el frontend.** En `TRAZ.html`, reemplaza:
   ```js
   const API_URL = 'REEMPLAZAR_CON_URL_DEL_DESPLIEGUE_TRAZ';
   ```
   por la URL del paso 3.
5. **Publicar el frontend.** `TRAZ.html` se sirve junto al resto (GitHub Pages). Abre
   `…/TRAZ/TRAZ.html`.

> El OAuth usa el mismo Client ID que `auth.js`. En `USUARIOS_AUTORIZADOS`, agrega o administra
> la columna `SEATRAZ`; solo usuarios activos con valor `TRUE` pueden consultar el módulo.

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
