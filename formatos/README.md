# Formatos Asociados a la OT — Contrato de Integración

SEAOT (`SEAOT.html`) abre los formatos del proceso ya **contextualizados a la
Orden de Trabajo activa**. Los formatos viven en el **mismo repo y origen** que
SEAOT (GitHub Pages, bajo `/SEA/`), por lo que la transferencia de datos es
**mismo-origen** y **no expone PII en la URL**.

Formatos integrados (en la raíz del repo):

- `registro-equipos.html` — Solicitud/registro de equipos (código EA-FCEQ-03.05)
- `supervision-gabinete.html` — Supervisión de gabinete (código EA-FSDG-12.01)

> Estado: **implementado en ambos lados** (SEAOT emite el handoff; cada formato
> ya trae su receptor). La Encuesta de Satisfacción queda pendiente.

## Cómo viajan los datos (handoff `localStorage`, sin PII en la URL)

Los datos del cliente (RFC, dirección, teléfono, correo) son PII y no deben
quedar en el historial, el encabezado `Referer` ni en logs. Por eso **no se pasan
como query params**. En su lugar:

1. SEAOT guarda el contexto en `localStorage['sea_ot_handoff_<token>']` con un
   token opaco de un solo uso y `exp` (caducidad, 5 min).
2. Abre el formato con solo el token: `registro-equipos.html?h=<token>`.
3. El formato lee la clave, **la borra** (uso único), valida `exp` y prellena.

Payload: `ot`, `serie`, `cliente`, `sucursal`, `rfc`, `direccion`, `contacto`,
`telefono`, `correo`, `emision`, `servicios`, `personal`, `fecha`.

El emisor está en `SEAOT.html` (`abrirFormato` / `makeHandoffToken` /
`sweepHandoffs`); `FORMATOS_BASE` vacío = mismo directorio/origen que SEAOT.

## Receptor (patrón, ya implementado en cada formato)

```html
<script>
(function () {
  function applyOtHandoff() {
    var data;
    try {
      var token = new URLSearchParams(location.search).get('h');
      if (!token) return;
      var key = 'sea_ot_handoff_' + token;
      var raw = localStorage.getItem(key);
      localStorage.removeItem(key);              // uso único
      if (!raw) return;
      var parsed = JSON.parse(raw);
      if (!parsed || (parsed.exp && parsed.exp < Date.now())) return;
      data = parsed.data || {};
    } catch (e) { return; }
    // Asignar a .value / .textContent (nunca innerHTML sin escapar):
    // p. ej. document.getElementById('destino').value = data.ot;
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', applyOtHandoff);
  else applyOtHandoff();
})();
</script>
```

### Mapeo por formato

- **registro-equipos.html**: `#destino` ← `ot`; `#motivoSalida` = `trabajo-campo`;
  equipos preseleccionados según las NOM de `servicios` (el código de inventario
  codifica la NOM: `EA-SO11`=011, `EA-CV24`=024, `EA-LX25`=025, `EA-MT15`=015,
  `EA-TK22`=022; los acompañantes los agrega `GRUPOS_EQUIPOS`). `#fechaSalida`
  se deja **pendiente** a propósito.
- **supervision-gabinete.html**: `#sg_ot` ← `ot`; `#sg_cliente` ← `cliente`.
  `#sg_fecha` y `#sg_folio` los completa el revisor.

## Pendiente

- 🚫 Encuesta de Satisfacción: diferida (el formato aún no existe).
