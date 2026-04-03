/**
 * recaptcha.js — Módulo reCAPTCHA Enterprise para portales públicos de SEA
 *
 * USO: Páginas públicas/clientes (PAIC.html, SEAPD.html).
 * NO incluir en páginas internas — esas usan auth.js.
 *
 * CONFIGURACIÓN:
 *   1. La Site Key pública está definida en SITE_KEY abajo (pública por diseño).
 *   2. El backend requiere tres Script Properties en GAS:
 *        RECAPTCHA_API_KEY  → API Key de GCP (restringida a reCAPTCHA Enterprise API)
 *        GCP_PROJECT_ID     → ID del proyecto GCP
 *        RECAPTCHA_SITE_KEY → mismo valor que SITE_KEY abajo
 *   3. En cada HTML:
 *        <script src="https://www.google.com/recaptcha/enterprise.js?render=SITE_KEY"></script>
 *        <script src="recaptcha.js"></script>
 *      Llama a SEARecaptcha.wrapFetch(url, options) en lugar de fetch().
 *
 * MIGRACIÓN: Classic reCAPTCHA v3 → Enterprise (Google obligó la migración).
 *   - SDK: recaptcha/api.js → recaptcha/enterprise.js
 *   - JS:  grecaptcha.execute → grecaptcha.enterprise.execute
 *   - Backend: api/siteverify + Secret Key → GCP REST API + API Key + Project ID
 */

const SEARecaptcha = (() => {
  // ─── CONFIGURACIÓN ─────────────────────────────────────────────────────────
  // Site Key pública de reCAPTCHA Enterprise (migrada desde Classic v3)
  const SITE_KEY = '6LdIq4csAAAAAHCggurLxQbsc15M8MnAD8UIOR1E';

  // Acción usada para identificar el contexto en el dashboard de reCAPTCHA
  const ACTION = 'sea_portal_submit';

  // ─── Obtener token de reCAPTCHA Enterprise ───────────────────────────────
  /**
   * Solicita un token de reCAPTCHA Enterprise.
   * No requiere interacción del usuario (invisible).
   * Incluye timeout de 8s para evitar que el formulario quede congelado
   * si el servicio de reCAPTCHA no responde.
   * @returns {Promise<string>} token
   */
  async function getToken() {
    if (!window.grecaptcha || !window.grecaptcha.enterprise) {
      throw new Error('reCAPTCHA Enterprise no cargó. Verifica tu conexión a internet.');
    }
    const tokenPromise = new Promise((resolve, reject) => {
      grecaptcha.enterprise.ready(() => {
        grecaptcha.enterprise.execute(SITE_KEY, { action: ACTION })
          .then(resolve)
          .catch(reject);
      });
    });
    const timeoutPromise = new Promise((_, reject) =>
      setTimeout(() => reject(new Error('reCAPTCHA timeout — recarga la página e intenta de nuevo')), 8000)
    );
    return Promise.race([tokenPromise, timeoutPromise]);
  }

  // ─── wrapFetch: inyecta el token en cada request ─────────────────────────
  /**
   * Reemplaza fetch() para incluir automáticamente el recaptcha_token.
   *
   * GET:  token como query param &recaptcha_token=...
   * POST: token en el campo recaptcha_token del body JSON
   *
   * El token se obtiene fresco en cada llamada (tokens son de un solo uso).
   */
  async function wrapFetch(url, options = {}) {
    let rcToken;
    try {
      rcToken = await getToken();
    } catch (e) {
      // Si falla reCAPTCHA, dejar pasar con token vacío.
      // El backend decidirá si rechazar o permitir.
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
