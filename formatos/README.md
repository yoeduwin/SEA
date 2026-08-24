# Formatos Asociados a la OT — Contrato de Integración

SEAOT (`SEAOT.html`) puede abrir los formatos del proceso ya **contextualizados a la
Orden de Trabajo activa**, pasando el folio y los datos del cliente como parámetros
de URL. Así se elimina la recaptura manual al volver de campo.

Este README define el **contrato** que deben cumplir las páginas de formato para
recibir esos datos. Aplica igual si los formatos viven en `yoeduwin.github.io/formatos/`
o dentro de este repo (carpeta `formatos/`).

> Para apuntar SEAOT a los formatos de este repo, cambia en `SEAOT.html`:
> `const FORMATOS_BASE = 'https://yoeduwin.github.io/formatos/';` → `'formatos/'`.

## Parámetros que envía SEAOT

| Parámetro    | Origen en SEAOT                                   |
|--------------|---------------------------------------------------|
| `ot`         | N° de Orden de Trabajo                             |
| `serie`      | Serie OT / OTB                                     |
| `cliente`    | Razón social                                      |
| `sucursal`   | Sucursal                                          |
| `rfc`        | RFC del cliente                                   |
| `direccion`  | Dirección del servicio                            |
| `contacto`   | Responsable de atendernos                         |
| `telefono`   | Teléfono de contacto                              |
| `correo`     | Correo del cliente                                |
| `emision`    | Fecha de emisión de la OT                         |
| `servicios`  | NOMs de la tabla de trabajo (separadas por coma)  |
| `personal`   | Personal asignado (sin duplicados)                |
| `fecha`      | Fecha de visita — **se omite** en Solicitud de Equipos |

`fecha` no se envía a `registro-equipos.html` de forma intencional: la fecha del
equipo se define al confirmar la visita, no al crear la OT.

## Adaptador de prellenado (pegar en cada formato)

Ajusta `FIELD_MAP` a los `id` reales de los campos de cada página y pega el bloque
antes de `</body>`. Rellena solo lo que llega; si un parámetro no viene, deja el
campo intacto.

```html
<script>
(function () {
  var q = new URLSearchParams(location.search);
  // Mapea cada parámetro de la OT al id del campo en ESTE formato:
  var FIELD_MAP = {
    ot:        'ot',          // <input id="ot">
    cliente:   'cliente',
    sucursal:  'sucursal',
    rfc:       'rfc',
    direccion: 'direccion',
    contacto:  'contacto',
    telefono:  'telefono',
    correo:    'correo',
    servicios: 'servicios',
    personal:  'personal',
    emision:   'emision',
    fecha:     'fecha'        // en registro-equipos no llega: se queda vacío
  };
  Object.keys(FIELD_MAP).forEach(function (param) {
    if (!q.has(param)) return;
    var el = document.getElementById(FIELD_MAP[param]);
    if (el && 'value' in el) el.value = q.get(param); // .value es seguro (no innerHTML)
  });
})();
</script>
```

> Seguridad: prefiere asignar a `.value` / `.textContent`. Nunca inyectes los
> parámetros con `innerHTML` sin escaparlos.

## Estado

- ✅ Lado SEAOT: implementado (modal **Documentación y Formatos Asociados**).
- ⏳ Lado formatos: pendiente pegar el adaptador y (opcional) el botón de descarga PDF.
- 🚫 Encuesta de Satisfacción: diferida a una segunda iteración (formato aún no existe).
