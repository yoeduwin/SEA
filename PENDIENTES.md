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

## 2. Meta de entrega: ¿20 o 25 días?

SEADB compara el tiempo promedio contra **20 días** (digital), y el correo al equipo menciona **25**. Hay que definir cuál es la meta oficial y alinear el sistema y la comunicación.

## 3. Personal: una sola lista para todos los sistemas

SEAOT, SEAINF y SEADB escriben distinto a las mismas personas, y SEADB acumula unas 38 variantes, así que el filtro por persona no sirve. Falta definir:
- La forma oficial de cada nombre (mayúsculas y acentos).
- Si "Iván Zárate" y "EDUWIN IVAN" son la misma persona.
- Si se quedan "Externo - Adrián H" y "Externo".
- Si se corrige el histórico de la hoja o solo de aquí en adelante.

## 4. SEAOT: ¿bloquear los botones de serie?

La serie se asigna sola al elegir el servicio. Los botones de serie siguen visibles por si hay que corregir. Falta decidir si se bloquean.

## 5. SEAINF: el responsable elegido no se guarda

Al crear el expediente se elige un responsable, pero el backend guarda el personal asignado en la OT (`fase3_CrearExpediente`, columna RESPONSABLE). Falta decidir cuál debe prevalecer.

## 6. Publicación del backend

Después de cada cambio en `BACKEND_FIXES.gs` hay que copiarlo a Apps Script y publicar una **nueva versión** de la implementación. El último cambio pendiente de publicar es el de los folios a 4 dígitos y la validación de folios repetidos.
