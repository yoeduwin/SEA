# Pendientes de mejora — sistemas SEA

Notas de lo que quedó por decidir o hacer después de la revisión de SEAINF, SEADB, PAIC, SEAPD y SEAOT (octubre 2026). Cada punto indica el estado y lo que falta definir.

## 1. Fecha meta de entrega: no aplica igual a todos los servicios

**Situación.** Hoy todo se mide contra un plazo fijo: SEAINF propone visita + 20 días, y SEADB calcula "en tiempo / atrasada" y el tiempo promedio contra 20 días (25 para el físico de OTB).

**Problema.** No todos los servicios tienen una fecha meta real:
- Algunos dependen de **tiempos de la autoridad** y tardan mucho más.
- En otros, después de la visita el **cliente tarda en corregir** lo que le corresponde.

Ese tiempo no es responsabilidad de EA, pero hoy cuenta como atraso y sube el promedio.

**Lo que ya existe y por qué no basta.** El botón **Pausar** de SEADB detiene el reloj y lo reanuda con nueva fecha. Pero depende de que alguien lo use a mano y a tiempo, que es justo lo que se quiere evitar. Debe quedar como una **alternativa**, no como la única opción.

**Ideas a evaluar (no decididas):**
- Indicar en el catálogo (`catalogo-servicios.js`) qué servicios **no tienen fecha meta** o tienen un **plazo propio**. Esos no contarían como atrasados ni entrarían al promedio contra la meta.
- Un estatus del tipo **"En espera del cliente / autoridad"** que se pueda marcar desde la OT o el expediente, y que excluya ese tiempo del cálculo. Sería parecido a la pausa, pero con causa explícita y visible en el tablero.
- Medir por separado el tiempo **de EA** (visita → informe listo) y el tiempo **total** (visita → entrega), para que la meta evalúe solo lo que depende de EA.

## 2. Personal: una sola lista para todos los sistemas (para después)

SEAOT, SEAINF y SEADB escriben distinto a las mismas personas, y SEADB acumula unas 38 variantes, así que el filtro por persona no sirve.

**Personal confirmado:** Jimmy Ayala, Martín Luna, Eduardo Campos e Iván Hernández. "Iván Hernández" es la misma persona que aparece como "Eduwin Iván", "Iván Zárate", etc.; su nombre completo es Eduwin Iván Hernández Zárate.

**Falta definir:**
- La forma exacta en que se escribe cada nombre en el sistema (mayúsculas y acentos).
- Si se quedan "Externo - Adrián H" y "Externo".
- Si se corrige el histórico de la hoja o solo se aplica de aquí en adelante.

## Resueltos

- **Meta de entrega:** es de **25 días**.
- **Backend:** ya publicado en Apps Script, con los folios a 4 dígitos y la validación de folios repetidos.
- **Botones de serie en SEAOT:** no se bloquean.
- **Responsable en SEAINF:** se deja como está.
