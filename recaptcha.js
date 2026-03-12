/**
 * recaptcha.js — Módulo reCAPTCHA v3 para portales públicos de SEA
 *
 * USO: Páginas públicas/clientes (paic.html, SEADB.html).
 * NO incluir en páginas internas — esas usan auth.js.
 *
 * CONFIGURACIÓN:
 *   1. Reemplaza SITE_KEY con tu clave pública de reCAPTCHA v3.
 *      La Site Key es pública — va en el HTML.
 *   2. La Secret Key va en GAS Script Properties como RECAPTCHA_SECRET_KEY.
 *      Nunca la pongas en el frontend.
 *   3. En cada HTML: <script src="recaptcha.js"></script>
 *      Llama a SEARecaptcha.wrapFetch(url, options) en lugar de fetch().
 *
 * REUTILIZABLE: Copiar a cualquier repo público del proyecto.
 * Solo cambiar SITE_KEY si usas un proyecto reCAPTCHA diferente.
 */

const SEARecaptcha = (() => {
  // ─── CONFIGURACIÓN ─────────────────────────────────────────────────────────
  // Reemplazar con tu Site Key de reCAPTCHA v3 (Google reCAPTCHA Console)
  const SITE_KEY = 'TU_RECAPTCHA_V3_SITE_KEY';

  // Acción usada para identificar el contexto en el dashboard de reCAPTCHA
  const ACTION = 'sea_portal_submit';

  // Puntuación mínima aceptada (0.0 a 1.0 — 0.5 es recomendado por Google)
  // El umbral real se valida en GAS, este valor es solo informativo.
  const MIN_SCORE_INFO = 0.5;

  // ─── Obtener token de reCAPTCHA v3 ───────────────────────────────────────
  /**
   * Solicita un token de reCAPTCHA v3 a Google.
   * No requiere interacción del usuario (invisible).
   * @returns {Promise<string>} token
   */
  async function getToken() {
    if (!window.grecaptcha) {
      throw new Error('reCAPTCHA no cargó. Verifica tu conexión a internet.');
    }
    return new Promise((resolve, reject) => {
      grecaptcha.ready(() => {
        grecaptcha.execute(SITE_KEY, { action: ACTION })
          .then(resolve)
          .catch(reject);
      });
    });
  }

  // ─── wrapFetch: inyecta el token en cada request ─────────────────────────
  /**
   * Reemplaza fetch() para incluir automáticamente el recaptcha_token.
   *
   * GET:  token como query param &recaptcha_token=...
   * POST: token en el campo recaptcha_token del body JSON
   *
   * El token se obtiene fresco en cada llamada (reCAPTCHA v3 tokens son de un solo uso).
   */
  async function wrapFetch(url, options = {}) {
    let rcToken;
    try {
      rcToken = await getToken();
    } catch (e) {
      // Si falla reCAPTCHA, dejar pasar con indicador de error
      // El backend decidirá si rechazar o permitir con advertencia
      console.warn('reCAPTCHA error:', e.message);
      rcToken = '';
    }

    const method = (options.method || 'GET').toUpperCase();

    if (method === 'GET') {
      const sep = url.includes('?') ? '&' : '?';
      url = `${url}${sep}recaptcha_token=${encodeURIComponent(rcToken)}`;
      return fetch(url, options);
    } else {
      let body = {};
      try {
        body = JSON.parse(options.body || '{}');
      } catch (_) {
        return fetch(url, options);
      }
      body.recaptcha_token = rcToken;
      return fetch(url, { ...options, body: JSON.stringify(body) });
    }
  }

  // ─── API pública del módulo ───────────────────────────────────────────────
  return { wrapFetch, getToken };
})();
