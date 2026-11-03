# Migración de SEAPD y PAIC a ejecutivambiental.com

| Portal | Dirección actual | Dirección nueva |
|---|---|---|
| SEAPD (Perfil de Datos, clientes) | `yoeduwin.github.io/SEA/SEAPD.html` | **`https://ejecutivambiental.com/clientes`** |
| PAIC (asesores, intermediarios y consultorías) | `yoeduwin.github.io/SEA/PAIC.html` | **`https://ejecutivambiental.com/asesores`** |

Solo cambian de lugar las páginas. El formulario, el backend de Apps Script,
las hojas de Google y Drive quedan igual. Las demás páginas internas (SEAOT,
SEAINF, etc.) siguen en GitHub sin cambios.

Sigue los pasos **en este orden**. Si haces el paso 3 antes del 1 y el 2,
los formularios marcarán error de reCAPTCHA.

---

## Paso 1 — Permitir el dominio nuevo en reCAPTCHA (Google Cloud)

1. Entra a <https://console.cloud.google.com/security/recaptcha> con la cuenta
   dueña del proyecto.
2. Abre la clave que empieza con `6LdIq4cs…`.
3. En **Dominios**, agrega `ejecutivambiental.com`.
   No borres `yoeduwin.github.io` todavía.
4. Guarda.

## Paso 2 — Permitir el dominio en el backend (Apps Script)

1. Abre el proyecto de Apps Script de SEA.
2. Ve a **Configuración del proyecto (engrane) → Propiedades del script**.
3. Busca `RECAPTCHA_HOSTNAMES`:
   - **Si existe:** cambia su valor a
     `ejecutivambiental.com,yoeduwin.github.io`
   - **Si no existe:** no hagas nada. Sin esa propiedad el backend acepta
     cualquier dominio. Es recomendable crearla con el valor de arriba,
     porque así nadie puede copiar el formulario en otro sitio.
4. Guarda. No hace falta volver a implementar el script.

## Paso 3 — Subir los archivos (cPanel → Administrador de archivos)

1. Descarga `sea-portales-hostgator.zip`.
2. En cPanel abre **Administrador de archivos → public_html**. Debe ser la
   carpeta de `ejecutivambiental.com`, no la de `intranet`.
3. Antes de subir, activa **Configuración (arriba a la derecha) → Mostrar
   archivos ocultos (dotfiles)**. Si no lo haces, no verás los `.htaccess`.
4. Clic en **Cargar** y sube el `.zip` a `public_html`.
5. Selecciona el zip → **Extraer**. Se crean dos carpetas:
   ```
   public_html/
   ├── clientes/   (index.html, recaptcha.js, .htaccess)
   └── asesores/   (index.html, recaptcha.js, .htaccess)
   ```
6. Borra el `.zip` de `public_html`.

> Si ya existiera una carpeta `clientes` o `asesores` de otra cosa, avísame
> antes de extraer para cambiar el nombre.

## Paso 4 — Probar

En una ventana de incógnito, revisa en cada dirección:

- [ ] `https://ejecutivambiental.com/clientes` abre el formulario de SEAPD.
- [ ] `https://ejecutivambiental.com/asesores` abre PAIC.
- [ ] Sale el candado (HTTPS). Si entras con `http://`, la página cambia
      sola a `https://`.
- [ ] Abajo a la derecha aparece el ícono de reCAPTCHA.
- [ ] PAIC: buscar un RFC que ya exista. Debe encontrarlo.
- [ ] Registrar un cliente de prueba en cada portal y confirmar que llega
      a la hoja de Google.
- [ ] `https://ejecutivambiental.com/clientes/.htaccess` debe dar
      **Forbidden** o **403**. Eso es correcto.

Si algo falla con *"No autorizado"* o *"AUTH_REQUIRED"*, el problema está en
el paso 1 o el 2. En Apps Script, en **Ejecuciones**, el registro dice
`reCAPTCHA Enterprise: ... hostname=...` y muestra qué dominio llegó.

## Paso 5 — Cambiar los enlaces

Actualiza donde se comparten las direcciones viejas: plantillas de correo,
WhatsApp, firmas, QR impresos, tu página web.

## Paso 6 — Cierre (2 a 4 semanas después)

Cuando ya nadie use las direcciones viejas:

1. Avísame y dejo en GitHub una redirección de las direcciones viejas a las
   nuevas.
2. Quita `yoeduwin.github.io` de reCAPTCHA (paso 1) y de
   `RECAPTCHA_HOSTNAMES` (paso 2).

---

## Cuando cambie el código de SEAPD o PAIC

En GitHub siempre se edita `SEAPD.html` y `PAIC.html` como hasta ahora.
Para actualizar el servidor:

1. Genera el paquete nuevo:
   ```bash
   bash scripts/empaquetar-hostgator.sh
   ```
   El zip queda en `dist/sea-portales-hostgator.zip`. También me lo puedes
   pedir y te lo preparo.
2. Súbelo y extráelo igual que en el paso 3. Cuando pregunte, acepta
   **sobrescribir**.

No edites los archivos directamente en el servidor. Si lo haces, la
siguiente subida borra tus cambios.
