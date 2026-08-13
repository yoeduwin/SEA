# PORTAL — Portal de Clientes (Ejecutiva Ambiental)

Portal web donde **cada cliente inicia sesión y consulta únicamente sus propios informes**, con su
estatus, y descarga el PDF del informe.

```
Cliente  →  inicia sesión con un código enviado a su correo  →  ve SOLO sus informes  →  descarga el PDF
```

Este módulo es **aditivo y aislado**: vive en la carpeta `PORTAL/` y **no modifica ningún archivo
ni dato del sistema SEA existente**. Es un **proyecto de Apps Script independiente** (igual que
`TRAZ/`) que solo **lee** las hojas que ya existen y **sirve su propia página** con `HtmlService`.
**No usa GitHub para alojarse** — la página la sirve el propio Apps Script.

---

## Principios de diseño

- **No modifica lo actual.** No toca `BACKEND_FIXES.gs` ni el proyecto interno de Apps Script.
  Se despliega como un proyecto **aparte**.
- **Solo lectura de datos SEA.** `PORTAL.gs` nunca usa `setValue`, `appendRow`, `insertSheet` ni
  escribe en Drive. Solo `getDataRange().getDisplayValues()` y lectura de archivos de Drive.
- **Autenticación propia de clientes.** Los clientes no tienen cuenta de Google en la whitelist
  interna. El acceso es por **código de un solo uso (OTP) enviado a su correo registrado** en
  `CLIENTES_MAESTRO`, seguido de un **token de sesión firmado (HMAC)**.
- **Aislamiento por RFC.** El alcance de los datos **siempre** sale del RFC firmado en el token,
  nunca de un parámetro suelto del cliente. Un cliente jamás ve informes de otro RFC.
- **El informe se entrega, Drive queda oculto.** El backend transmite el **PDF** del informe; el
  cliente no accede a la carpeta de Drive ni ve material interno (hojas de campo, croquis, fotos).

---

## Qué lee (fuente de datos)

Mismo Spreadsheet del SEA (`SPREADSHEET_ID` en `PORTAL.gs`), en **solo lectura**:

| Hoja | Uso en PORTAL |
|---|---|
| `CLIENTES_MAESTRO` | Resolver RFC ↔ correo(s) del cliente para enviar el código de acceso. |
| `ORDENES_TRABAJO` | Estatus externo del servicio (visible al cliente) y relación por folio de OT. |
| `INFORMES` | Informes del cliente: número, NOM, fechas, estatus y enlace del expediente (Drive). |

---

## Archivos

| Archivo | Rol |
|---|---|
| `PORTAL.gs` | Backend independiente: `doGet` (sirve la página), autenticación (OTP + token HMAC), lecturas filtradas por RFC y entrega del PDF. Acciones: `portal_solicitarCodigo`, `portal_verificarCodigo`, `portal_misServicios`, `portal_descargarInforme`. |
| `PORTAL.html` | Página servida por `HtmlService`: acceso por código y panel de informes. Usa `google.script.run` (sin CORS). Sin dependencias externas. |

---

## Despliegue (una sola vez)

PORTAL se despliega como un **proyecto de Apps Script separado**, para no tocar el backend SEA.

1. **Crear proyecto Apps Script.** En [script.google.com](https://script.google.com) → *Nuevo
   proyecto*.
   - Crea un archivo de script y pega el contenido de `PORTAL.gs`.
   - Crea un archivo **HTML** llamado exactamente `PORTAL` (menú *+* → *HTML*) y pega el contenido
     de `PORTAL.html`. (El nombre debe ser `PORTAL` porque `doGet` hace
     `HtmlService.createHtmlOutputFromFile('PORTAL')`.)
2. **Propiedad de script (obligatoria).** *Configuración del proyecto* → *Propiedades de script* →
   agregar:
   ```
   PORTAL_HMAC_SECRET = <una cadena aleatoria larga, ej. 40+ caracteres>
   ```
   Es el secreto que firma los tokens de sesión. Guárdalo bien; si lo cambias, se cierran todas
   las sesiones.
3. **Permisos.** La cuenta que despliega debe tener acceso de **lectura** al Spreadsheet
   `1MoScea4CYg0NCjvPjHqZwV0cKhrd2nxfW8LYhz_4pDo` y a las carpetas de expedientes en Drive, y poder
   enviar correo (Gmail) para el código.
4. **Desplegar como app web.** *Implementar* → *Nueva implementación* → *Aplicación web*:
   - Ejecutar como: **Yo** (dueño con acceso al Spreadsheet y Drive).
   - Quién tiene acceso: **Cualquier usuario**.
   - Autoriza los permisos que pida (leer Sheets/Drive, enviar Gmail) y copia la URL `.../exec`.
5. **Compartir con los clientes.** Esa URL `.../exec` **es el portal**. Compártela (puedes
   acortarla o enlazarla desde tu web/redes). No hay GitHub Pages ni merge a `main`.

> Cada vez que edites `PORTAL.gs` o `PORTAL.html`, crea una **nueva versión** de la implementación
> (o usa *Administrar implementaciones* → editar → *Nueva versión*) para que los cambios salgan en
> vivo.

---

## Cómo lo usa el cliente

1. Abre la URL del portal.
2. Escribe su **RFC** (o el correo registrado con EA) y pulsa *Enviar código*.
3. Recibe un **código de 6 dígitos** en su correo (válido 10 min) y lo ingresa.
4. Ve la lista de **sus** servicios/informes con su estatus y descarga el PDF de los ya entregados.

La sesión dura 30 días en ese navegador (token en `localStorage`); *Salir* la cierra.

---

## Entrega del informe (importante)

El PDF que descarga el cliente sale de la **carpeta del expediente** (`INFORMES.LINK_DRIVE`). Para
que el portal entregue **solo el informe** y no el material interno:

> **Convención recomendada:** coloca el **PDF final del informe en la raíz** de la carpeta del
> expediente. El material de trabajo permanece en las subcarpetas internas (`1. ORDEN_TRABAJO`,
> `2. HDC`, `3. CROQUIS`, `4. FOTOS`), que el portal **nunca** expone.

Comportamiento de `portal_descargarInforme`:
- Toma los **PDF de la raíz** del expediente (no de las subcarpetas). Si hay varios, prefiere el
  que contiene el número de informe en el nombre; si no, el más reciente.
- Si no hay PDF en la raíz (o el archivo supera el tope de descarga directa), **cae de vuelta** al
  enlace de la carpeta como respaldo.

Puedes cambiar el modo en `PORTAL_CONFIG`:
- `MODO_ENTREGA: 'PDF'` (por defecto) — el portal transmite el PDF, Drive oculto.
- `MODO_ENTREGA: 'CARPETA'` — el portal abre el enlace de la carpeta de Drive (requiere que
  compartas permisos de esas carpetas con el cliente y expone el material interno).

---

## Seguridad

- Los datos siempre están detrás de sesión; el alcance sale del **RFC del token firmado**,
  verificado en el servidor en cada llamada.
- Código de un solo uso: 6 dígitos, válido 10 min, máx. 5 intentos, 1 código por minuto por RFC.
- Descarga: se valida que el informe pertenezca al RFC del token antes de entregar bytes; solo se
  exponen PDF de la raíz del expediente.
- El secreto HMAC vive en *Propiedades de script*, nunca en el frontend.
- La página es pública, pero **sin sesión válida no devuelve ningún dato**.

> Endurecimiento opcional: activar reCAPTCHA en `portal_solicitarCodigo` (reusando las llaves ya
> documentadas del sistema) si se detecta abuso. No es imprescindible, porque el código solo llega
> al correo ya registrado del cliente.

---

## Prueba rápida (sin desplegar la web)

En el editor de Apps Script, con `PORTAL_HMAC_SECRET` ya configurado, ejecuta la función
`test_portal()` y revisa *Registros* (Ctrl+Enter). Verifica: firma/verificación del token, rechazo
de token manipulado, hash de código estable/sensible, y que `misServicios`/`descargarInforme`
exigen token válido y filtran por RFC (RFC de prueba `XTES000000TST`).

---

## Nota de acoplamiento

`PORTAL.gs` copia los índices de columna del esquema SEA (`PORTAL_CONFIG.COL_*`), igual que
`TRAZ.gs`. Si en el futuro se reordenan columnas en `CLIENTES_MAESTRO` / `ORDENES_TRABAJO` /
`INFORMES`, hay que actualizar esos índices aquí también.
