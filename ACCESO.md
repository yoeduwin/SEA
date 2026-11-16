# Acceso real — EA Suite (SEA)

## Qué protege este candado

- **SEAOT, SEADB, SEAINF, SEATRAZ, registro-equipos y supervision-gabinete:**
  login Google + verificación en el servidor contra la lista de acceso.
  Sin sesión verificada no se ve ni se usa la página (el contenido queda
  oculto detrás del overlay de login) y no se carga ningún dato.
- **SEAPD:** el formulario de alta para clientes **nuevos sigue público**.
  "Ya soy cliente" (consulta de datos por RFC) exige login del equipo. La
  búsqueda sigue devolviendo solo la vista pública enmascarada (razón
  social, sucursales, correo/teléfono enmascarados); nunca la ficha completa.
- **Índice (index.html):** ya no pasa el `id_token` en la URL (quedaba en
  historial y logs); ahora va por POST en el body.

## Cómo funciona

1. Google Identity Services (GIS) emite un `id_token`.
2. `auth.js` lo envía por POST al backend de Apps Script (acción
   `verificarAcceso`), que lo valida contra Google (tokeninfo + audience) y
   comprueba el correo en la hoja **USUARIOS_AUTORIZADOS** (columna Activo y
   columna del módulo).
3. Solo entonces se muestra la página y se cargan datos (siempre con
   `SEAAuth.wrapFetch`, que inyecta el token en el body, nunca en la URL).
4. El backend niega por defecto cualquier acción no registrada (`AUTH_MODE`
   en `BACKEND_FIXES.gs` / `TRAZ/TRAZ.gs`).

## Configuración (un solo valor)

- `CLIENT_ID` en `auth.js`: **un solo cliente OAuth de Google** para todos
  los módulos. El backend valida el `aud` del token contra la Script
  Property `GOOGLE_CLIENT_ID` (en `TRAZ.gs` va fijo en el código).
- La lista de acceso es la hoja **USUARIOS_AUTORIZADOS** del Sheet del
  backend. Para dar o quitar acceso se edita esa hoja; no se toca código.

## Qué NO hace este candado

- El HTML/JS estático de GitHub Pages siempre es descargable por diseño;
  la garantía real de los datos está en el backend (GAS). El candado evita
  el uso casual y la lectura de la estructura sin login.
- La sesión en el navegador dura hasta que expira el token de Google
  (~1 h). Si alguien se da de baja en la lista, su sesión abierta sigue
  válida hasta que expire el token, pero cada llamada a GAS vuelve a
  verificar y se le niega.

## Pendiente (fuera de este PR)

- [ ] Confirmar en Google Cloud Console que el cliente OAuth existe y que
      sus orígenes JavaScript autorizados incluyen
      `https://yoeduwin.github.io` y `https://intranet.ejecutivambiental.com`.
- [ ] Confirmar la Script Property `GOOGLE_CLIENT_ID` en los dos despliegues
      de Apps Script (SEA y SEATRAZ) y que la hoja `USUARIOS_AUTORIZADOS`
      tiene al equipo con las columnas de módulo correctas.
- [ ] Revisar este PR, hacer merge y comprobar en producción que cada
      módulo pide login y que "Ya soy cliente" lo pide también.

## Cómo probarlo (tras el merge)

1. Abrir cada módulo en incógnito: debe aparecer solo el login de Google.
2. Entrar con una cuenta del equipo: la página se muestra y carga datos.
3. Entrar con una cuenta fuera de la lista: mensaje "no tiene acceso" y la
   página no se muestra.
4. En SEAPD sin login: el formulario de alta funciona; "Ya soy cliente"
   pide login primero.
