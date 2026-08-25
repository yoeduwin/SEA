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

## Cómo viajan los datos (handoff `sessionStorage`, sin PII en la URL)

Los datos del cliente (RFC, dirección, teléfono, correo) son PII y no deben
quedar en el historial, el encabezado `Referer` ni en logs. Por eso **no se pasan
como query params**. Tampoco se usa `localStorage` (persiste en disco). Se usa
`sessionStorage`, que **nunca se escribe en disco ni sobrevive a la pestaña**:

1. SEAOT guarda el contexto en `sessionStorage['sea_ot_handoff_<token>']` con un
   token opaco de un solo uso.
2. Abre el formato **sin `noopener`**: el navegador **clona** el `sessionStorage`
   a la pestaña nueva. En la URL viaja **solo** el token: `...?h=<token>`.
3. La pestaña del formato lee su clon, **lo borra** (uso único) y prellena;
   **SEAOT elimina su propia copia de inmediato**.

No hay temporizadores ni dependencia de que SEAOT siga abierto: la copia de SEAOT
se borra al instante y el clon del formato desaparece al cerrar esa pestaña. El
token de la URL es opaco (no es PII).

Payload: `ot`, `serie`, `cliente`, `sucursal`, `rfc`, `direccion`, `contacto`,
`telefono`, `correo`, `emision`, `servicios`, `personal`, `fecha`.

El emisor está en `SEAOT.html` (`abrirFormato` / `makeHandoffToken`);
`FORMATOS_BASE` vacío = mismo directorio/origen que SEAOT.

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
      var raw = sessionStorage.getItem(key);      // clon propio de esta pestaña
      sessionStorage.removeItem(key);             // uso único
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
