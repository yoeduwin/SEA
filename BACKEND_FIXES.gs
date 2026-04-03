// =========================================================================
// API CENTRAL - EJECUTIVA AMBIENTAL (SISTEMA UNIFICADO v3.0 - MULTI-SUCURSAL)
// VERSIÓN CORREGIDA — Listo para copiar/pegar en Google Apps Script
// =========================================================================
// CAMBIOS aplicados sobre el backend actual:
//   1. fase1_RegistrarCliente()  → NO envía correo al registrar; sí crea carpeta
//   2. fase2_BuscarClienteRFC/Nombre() → corregido índice link_drive_cliente [22]→[20]
//      (nuevo esquema CLIENTES_MAESTRO tiene 22 columnas; Drive link en col 21 índice 20, Asesor/Consultor en col 22 índice 21)
//   3. fase3_CrearExpediente()   → fallback #3 corregido índice [22]→[20]
//   4. getOrdenesSafe_()         → ahora devuelve rfc, sucursal, personal, nom, fecha_visita
//      (sin rfc, SEAINF nunca podía buscar carpeta por RFC → expediente caía en raíz)
// =========================================================================
const CONFIG = {
  // ── Spreadsheet y Drive ─────────────────────────────────────────────────
  SPREADSHEET_ID: '1MoScea4CYg0NCjvPjHqZwV0cKhrd2nxfW8LYhz_4pDo',
  FOLDER_ID:      '1nHd-70uUeciClDm_3_pgbmqGF7II1lfQ',

  // ── Nombres de hojas ─────────────────────────────────────────────────────
  SHEET_CLIENTES:  'CLIENTES_MAESTRO',
  SHEET_OT:        'ORDENES_TRABAJO',
  SHEET_INFORMES:  'INFORMES',
  SHEET_USUARIOS:  'USUARIOS_AUTORIZADOS',
  SHEET_AUDITORIA: 'AUDITORIA',

  // ── Zona horaria ─────────────────────────────────────────────────────────
  TIMEZONE: 'GMT-6',

  // ── Correos de notificación interna ──────────────────────────────────────
  EMAIL_TO: [
    'direccion.general@ejecutivambiental.com',
    'operaciones@ejecutivambiental.com',
    'aclientes@ejecutivambiental.com'
  ],

  // ── Datos de contacto de soporte (aparecen en correos al cliente) ─────────
  SUPPORT_EMAIL: 'aclientes@ejecutivambiental.com',
  SUPPORT_PHONE: '222 941 7295',

  // ── Identidad de la empresa ───────────────────────────────────────────────
  COMPANY_NAME: 'Ejecutiva Ambiental',

  // ── Reintentos de correo ──────────────────────────────────────────────────
  EMAIL_RETRY_ATTEMPTS: 3,
  EMAIL_RETRY_DELAY_MS: 2000,

  // ── Estructura de subcarpetas en cada expediente Drive ────────────────────
  FOLDER_STRUCTURE: {
    ORDEN_TRABAJO: '1. ORDEN_TRABAJO',
    HOJAS_CAMPO:   '2. HDC',
    CROQUIS:       '3. CROQUIS',
    FOTOS:         '4. FOTOS'
  },

  // ── Índices de columna (0-based) por hoja ────────────────────────────────
  COLUMNS: {
    CLIENTES: {
      FECHA_REGISTRO:   0,
      RAZON_SOCIAL:     1,
      SUCURSAL:         2,
      RFC:              3,
      REPRESENTANTE:    4,
      DIRECCION:        5,
      TELEFONO:         6,
      SOLICITANTE:      7,
      CORREO:           8,
      GIRO:             9,
      REGISTRO_PATRONAL:10,
      CAP_INSTALADA:    11,
      CAP_OPERACION:    12,
      DIAS_TURNOS:      13,
      APLICA_NOM020:    14,
      REQUIERE_PIPC:    15,
      RESPONSABLE:      16,
      TEL_RESPONSABLE:  17,
      NOMBRE_DIRIGIDO:  18,
      PUESTO_DIRIGIDO:  19,
      LINK_DRIVE:       20,
      ASESOR_CONSULTOR: 21
    },
    ORDENES: {
      FECHA:            0,  // A
      OT:               1,  // B
      TIPO:             2,  // C
      NOM:              3,  // D  (NUM_INFORME eliminado)
      CLIENTE:          4,  // E
      SUCURSAL:         5,  // F
      RFC:              6,  // G
      PERSONAL:         7,  // H
      FECHA_VISITA:     8,  // I
      FECHA_ENTREGA:    9,  // J
      FECHA_REAL:       10, // K
      ESTATUS_EXTERNO:  11, // L
      LINK_DRIVE:       12, // M
      OBSERVACIONES:    13  // N  (ESTATUS_INFORME eliminado)
    },
    INFORMES: {
      TIMESTAMP:        0,  // A
      NUM_INFORME:      1,  // B
      TIPO_ORDEN:       2,  // C
      OT:               3,  // D
      NOM:              4,  // E
      CLIENTE:          5,  // F
      SOLICITANTE:      6,  // G
      RFC:              7,  // H
      TELEFONO:         8,  // I
      DIRECCION:        9,  // J
      FECHA_SERVICIO:   10, // K
      FECHA_ENTREGA:    11, // L
      ES_CAPACITACION:  12, // M
      ESTATUS:          13, // N
      LINK_DRIVE:       14, // O
      RESPONSABLE:      15, // P
      SUCURSAL:         16  // Q
    },
    USUARIOS: {
      EMAIL:  0,
      NOMBRE: 1,
      ROL:    2,
      ACTIVO: 3
    }
  }
};
// Shorthands de columnas para legibilidad en funciones
const CL = CONFIG.COLUMNS.CLIENTES;
const CO = CONFIG.COLUMNS.ORDENES;
const CI = CONFIG.COLUMNS.INFORMES;
const CU = CONFIG.COLUMNS.USUARIOS;

// Valores válidos para INFORMES.Estatus (hoja INFORMES col N)
const ESTATUS_INFORME_VALIDOS_ = ['NO INICIADO', 'EN PROCESO', 'PARA REVISION', 'PARA IMPRESION', 'FINALIZADO', 'CANCELADO'];
const ESTATUS_INFORME_TERMINALES_ = ['FINALIZADO', 'CANCELADO'];
// Valores válidos para ORDENES_TRABAJO.EstatusExterno (col L)
const ESTATUS_EXTERNO_TERMINALES_ = ['FINALIZADO', 'CANCELADO'];
// =========================================================================
// MÓDULO DE SEGURIDAD — Autenticación Google OAuth + reCAPTCHA v3
// =========================================================================
// Modos de autenticación por acción:
//   GOOGLE    → requiere id_token válido + usuario en whitelist
//   RECAPTCHA → requiere recaptcha_token válido (portales públicos de clientes)
//   EITHER    → acepta id_token (Google) o recaptcha_token
//
// Páginas internas con Google Auth: SEADB, SEAOT, SEAINF
// Portales públicos con reCAPTCHA:  PAIC.html, SEAPD.html (registro de clientes)
//
// Módulos para control por columna en USUARIOS_AUTORIZADOS:
//   SEADB → Dashboard / control de entregas
//   SEAOT → Órdenes de trabajo
//   SEAINF → Expedientes e informes
// =========================================================================
const AUTH_MODE = {
  // ── Requiere Google Auth (herramientas internas) ─────────────────────────
  // SEADB
  getTablero:             'GOOGLE',
  getTableroInf:          'GOOGLE',
  updateEstatus:          'GOOGLE',
  updateRespInf:          'GOOGLE',
  updateResponsable:      'GOOGLE',
  // SEAOT
  buscarClienteRFC:       'EITHER',   // SEAOT usa Google Auth; PAIC/SEAPD (públicos) usan reCAPTCHA
  buscarClienteNombre:    'EITHER',   // igual que buscarClienteRFC
  registrarOT:            'GOOGLE',
  // SEAINF
  getOrdenes:             'GOOGLE',
  getConsecutivo:         'GOOGLE',
  createExpediente:       'GOOGLE',
  addFilesToExpediente:   'GOOGLE',
  updateEstatusInforme:   'GOOGLE',
  getRenovaciones:        'GOOGLE',
  // ── Requiere reCAPTCHA (portales públicos de registro de clientes) ────────
  // PAIC.html y SEAPD.html son portales donde los CLIENTES se registran
  registrarCliente:       'RECAPTCHA',
  // ── Ping de verificación previo a la carga de la app ────────────────────
  // auth.js llama esto ANTES de mostrar la UI, para bloquear no autorizados
  verificarAcceso:        'GOOGLE'
};

// Mapeo acción → módulo (para verificar acceso por columna en la hoja)
const ACTION_MODULE = {
  getTablero:             'SEADB',
  getTableroInf:          'SEAINF',
  updateEstatus:          'SEADB',
  updateResponsable:      'SEADB',
  updateRespInf:          'SEAINF',
  buscarClienteRFC:       'SEAOT',
  buscarClienteNombre:    'SEAOT',
  registrarOT:            'SEAOT',
  getOrdenes:             'SEAINF',
  getConsecutivo:         'SEAINF',
  createExpediente:       'SEAINF',
  addFilesToExpediente:   'SEAINF',
  updateEstatusInforme:   'SEAINF',
  getRenovaciones:        'SEADB'
};

// ── verificarIdToken_ ──────────────────────────────────────────────────────
/**
 * Verifica el id_token con Google y retorna datos del usuario.
 * Usa CacheService 10 min para no llamar tokeninfo en cada request.
 * @param {string} idToken
 * @returns {{email:string, name:string, sub:string}|null}
 */
function verificarIdToken_(idToken) {
  if (!idToken || typeof idToken !== 'string' || idToken.length < 100) return null;

  const cacheKey = 'idtok_' + idToken.slice(-32);
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (_) { /* continúa */ }
  }

  try {
    const url = 'https://www.googleapis.com/oauth2/v3/tokeninfo?id_token=' + encodeURIComponent(idToken);
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;

    const data = JSON.parse(resp.getContentText());
    if (data.error_description) return null;

    // Verificar audience contra nuestro Client ID (protege contra tokens de otras apps)
    const CLIENT_ID = PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID');
    if (CLIENT_ID && data.aud !== CLIENT_ID) {
      Logger.log('Token rechazado: aud=' + data.aud + ' esperado=' + CLIENT_ID);
      return null;
    }

    const userInfo = { email: data.email || '', name: data.name || '', sub: data.sub || '' };
    const ttl = Math.min(600, Math.max(5, Number(data.exp) - Math.floor(Date.now() / 1000) - 60));
    cache.put(cacheKey, JSON.stringify(userInfo), ttl);
    return userInfo;

  } catch (e) {
    Logger.log('verificarIdToken_ error: ' + e.message);
    return null;
  }
}

// ── verificarUsuarioAutorizado_ ───────────────────────────────────────────
/**
 * Verifica si el email está activo en USUARIOS_AUTORIZADOS y tiene acceso al módulo.
 * Columnas: Email | Nombre | Rol | Activo | FechaAlta | SEADB | SEAOT | SEAINF | Notas
 * Para crear/recrear la hoja con los datos iniciales, ejecutar crearHojaUsuarios() desde
 * el editor de GAS (una sola vez).
 * @param {string} email
 * @param {string} modulo  'SEADB' | 'SEAOT' | 'SEAINF' | '' (sin restricción de módulo)
 * @returns {boolean}
 */
function verificarUsuarioAutorizado_(email, modulo) {
  if (!email) return false;
  const emailLower = email.toLowerCase().trim();

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_USUARIOS);

    if (!sheet) {
      Logger.log('Hoja ' + CONFIG.SHEET_USUARIOS + ' no existe. Ejecuta setupSheets() desde el editor GAS.');
      return false;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return false;

    // Leer encabezados para encontrar columnas de módulos dinámicamente
    const headers = data[0].map(h => String(h).toUpperCase().replace(/[^A-Z0-9]/g, ''));
    const moduloCol = modulo ? headers.indexOf(modulo.toUpperCase()) : -1;

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][CU.EMAIL]).toLowerCase().trim();
      const activo   = data[i][CU.ACTIVO];

      if (rowEmail !== emailLower) continue;
      if (activo !== true) {
        // Alerta de acceso con cuenta desactivada (máximo 1 email/hora por usuario)
        Logger.log('ALERTA: Usuario desactivado intentó acceder: ' + email + ' [' + new Date().toISOString() + ']');
        try {
          const alertKey = 'baja_alerta_' + emailLower.replace(/[^a-z0-9]/g, '').slice(-20);
          const alertCache = CacheService.getScriptCache();
          if (!alertCache.get(alertKey)) {
            GmailApp.sendEmail(
              CONFIG.EMAIL_TO[0],
              'ALERTA: Acceso de usuario desactivado — Sistema SEA',
              'El usuario ' + email + ' intentó acceder al sistema pero su cuenta está marcada como INACTIVA.\n' +
              'Fecha: ' + new Date().toISOString() + '\n\n' +
              'Si esta persona ya no trabaja en la empresa, no se requiere acción adicional.\n' +
              'Si fue un error, activa la cuenta en la hoja USUARIOS_AUTORIZADOS.\n\n' +
              'Sistema SEA — Ejecutiva Ambiental'
            );
            alertCache.put(alertKey, '1', 3600); // silencio de 1 hora para no saturar emails
          }
        } catch (e) { /* no interrumpir el flujo de autenticación */ }
        return false;
      }

      // Si hay columna para este módulo, verificar acceso
      if (moduloCol >= 0) {
        return data[i][moduloCol] === true;
      }
      return true; // sin columna de módulo → permitir
    }
    return false; // email no encontrado

  } catch (e) {
    Logger.log('verificarUsuarioAutorizado_ error: ' + e.message);
    return false;
  }
}

// ── crearHojaUsuarios ─────────────────────────────────────────────────────
/**
 * EJECUTAR UNA VEZ desde el editor de GAS (menú Ejecutar → crearHojaUsuarios).
 * Crea (o recrea) la hoja USUARIOS_AUTORIZADOS con las columnas correctas
 * y los 4 usuarios autorizados pre-cargados.
 *
 * Si la hoja ya existe, la elimina y la recrea para garantizar columnas actualizadas.
 */
function crearHojaUsuarios() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Eliminar hoja existente si hay
  const existing = ss.getSheetByName(CONFIG.SHEET_USUARIOS);
  if (existing) {
    ss.deleteSheet(existing);
    Logger.log('Hoja anterior eliminada.');
  }

  // Crear nueva hoja
  const sheet = ss.insertSheet(CONFIG.SHEET_USUARIOS);

  // Encabezados — SEADB | SEAOT | SEAINF controlan acceso por módulo
  const headers = ['Email', 'Nombre', 'Rol', 'Activo (TRUE/FALSE)', 'Fecha Alta', 'SEADB', 'SEAOT', 'SEAINF', 'Notas'];
  sheet.appendRow(headers);

  // Usuarios autorizados iniciales (todos los módulos activos)
  const hoy = new Date();
  sheet.appendRow(['eduwin.ejecutiva@gmail.com',       'Administrador',        'admin',             true, hoy, true, true, true, '']);
  sheet.appendRow(['aclientes.ejecutiva@gmail.com',    'Operador A Clientes',  'operador',          true, hoy, true, true, true, '']);
  sheet.appendRow(['operaciones.ejecutivamx@gmail.com','Operador Operaciones', 'operador',          true, hoy, true, true, true, '']);
  sheet.appendRow(['calidad.ejecutivamx@gmail.com',    'Aux. Operador Calidad','auxiliar_operador', true, hoy, true, true, true, '']);

  // Formato encabezado
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // Ajustar ancho de columnas
  sheet.setColumnWidth(1, 280); // Email
  sheet.setColumnWidth(2, 200); // Nombre
  sheet.setColumnWidth(5, 120); // Fecha Alta

  Logger.log('✅ Hoja ' + CONFIG.SHEET_USUARIOS + ' creada con 4 usuarios.');
  Logger.log('   Para agregar/desactivar usuarios: edita directamente la hoja.');
  Logger.log('   Columnas SEADB/SEAOT/SEAINF: TRUE = acceso permitido, FALSE = bloqueado.');
  Logger.log('   El GAS refresca permisos en máximo 10 minutos (cache).');
}

// =========================================================================
// SETUP — Inicialización del Spreadsheet
// =========================================================================
/**
 * setupSheets()
 *
 * Crea todas las hojas necesarias para el sistema SEA si no existen.
 * Ejecutar UNA VEZ desde el editor de GAS (menú Ejecutar → setupSheets)
 * al desplegar el sistema en un Spreadsheet nuevo.
 *
 * Hojas que crea / verifica:
 *   1. CLIENTES_MAESTRO   — registro maestro de clientes y sucursales
 *   2. ORDENES_TRABAJO    — órdenes de trabajo (OTs)
 *   3. USUARIOS_AUTORIZADOS — control de acceso por módulo
 *   4. AUDITORIA          — log de cambios (se auto-crea al primer cambio,
 *                           pero aquí se puede precrear con formato)
 *
 * Las hojas ya existentes NO se tocan ni se eliminan.
 */
function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const headerStyle = (range) => range
    .setFontWeight('bold')
    .setBackground(CONFIG.COLUMNS ? '#1e5a3e' : '#1a73e8') // verde corporativo
    .setFontColor('#ffffff')
    .setFrozenRows(1);

  // ── 1. CLIENTES_MAESTRO ────────────────────────────────────────────────
  if (!ss.getSheetByName(CONFIG.SHEET_CLIENTES)) {
    const s = ss.insertSheet(CONFIG.SHEET_CLIENTES);
    const headers = [
      'Fecha Registro', 'Razón Social', 'Sucursal', 'RFC',
      'Representante Legal', 'Dirección Evaluación', 'Teléfono Empresa',
      'Contacto (Solicitante)', 'Email Contacto', 'Giro / Actividad',
      'Registro Patronal', 'Capacidad Instalada', 'Capacidad Operación',
      'Días / Turnos', 'Aplica NOM-020', 'Requiere PIPC',
      'Nombre quien Atiende', 'Teléfono quien Atiende',
      'A quien se dirige informe', 'Puesto a quien se dirige',
      'Link Drive', 'Asesor / Consultor'
    ];
    s.appendRow(headers);
    headerStyle(s.getRange(1, 1, 1, headers.length));
    s.setColumnWidth(2, 220); s.setColumnWidth(3, 160); s.setColumnWidth(4, 140);
    s.setColumnWidth(5, 200); s.setColumnWidth(6, 260); s.setColumnWidth(9, 220);
    Logger.log('✅ Hoja ' + CONFIG.SHEET_CLIENTES + ' creada.');
  } else {
    Logger.log('ℹ️  Hoja ' + CONFIG.SHEET_CLIENTES + ' ya existe — sin cambios.');
  }

  // ── 2. ORDENES_TRABAJO ─────────────────────────────────────────────────
  if (!ss.getSheetByName(CONFIG.SHEET_OT)) {
    const s = ss.insertSheet(CONFIG.SHEET_OT);
    const headers = [
      'Fecha Alta', 'OT Folio', 'Tipo Orden', 'Num Informe', 'NOM / Servicio',
      'Cliente Inicial (Razón Social)', 'Cliente Final (Sucursal)', 'RFC',
      'Personal Asignado', 'Fecha Visita', 'Fecha Entrega Límite',
      'Fecha Real Entrega', 'Estatus (Externo SEADB)', 'Link Drive',
      'Observaciones', 'Estatus Informe (Interno SEAINF)'
    ];
    s.appendRow(headers);
    headerStyle(s.getRange(1, 1, 1, headers.length));
    s.setColumnWidth(2, 130); s.setColumnWidth(6, 220); s.setColumnWidth(7, 180);
    s.setColumnWidth(9, 160); s.setColumnWidth(14, 280);
    Logger.log('✅ Hoja ' + CONFIG.SHEET_OT + ' creada.');
  } else {
    Logger.log('ℹ️  Hoja ' + CONFIG.SHEET_OT + ' ya existe — sin cambios.');
  }

  // ── 3. USUARIOS_AUTORIZADOS ────────────────────────────────────────────
  if (!ss.getSheetByName(CONFIG.SHEET_USUARIOS)) {
    crearHojaUsuarios(); // reutiliza la función existente que ya tiene usuarios iniciales
  } else {
    Logger.log('ℹ️  Hoja ' + CONFIG.SHEET_USUARIOS + ' ya existe — sin cambios.');
  }

  // ── 4. AUDITORIA ───────────────────────────────────────────────────────
  if (!ss.getSheetByName(CONFIG.SHEET_AUDITORIA)) {
    const s = ss.insertSheet(CONFIG.SHEET_AUDITORIA);
    const headers = ['Timestamp', 'Usuario', 'Accion', 'OT', 'Campo', 'Valor_Anterior', 'Valor_Nuevo'];
    s.appendRow(headers);
    const auditHeaderRange = s.getRange(1, 1, 1, headers.length);
    auditHeaderRange.setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    s.setFrozenRows(1);
    s.setColumnWidth(1, 180); s.setColumnWidth(2, 220);
    Logger.log('✅ Hoja ' + CONFIG.SHEET_AUDITORIA + ' creada.');
  } else {
    Logger.log('ℹ️  Hoja ' + CONFIG.SHEET_AUDITORIA + ' ya existe — sin cambios.');
  }

  Logger.log('');
  Logger.log('=== setupSheets() completado ===');
  Logger.log('Hojas verificadas: ' + [CONFIG.SHEET_CLIENTES, CONFIG.SHEET_OT, CONFIG.SHEET_USUARIOS, CONFIG.SHEET_AUDITORIA].join(', '));
}

// ── verificarRecaptcha_ ───────────────────────────────────────────────────
/**
 * Verifica un token de reCAPTCHA Enterprise con la API de GCP.
 * Requiere tres Script Properties:
 *   RECAPTCHA_API_KEY  → API Key de GCP (restringida a reCAPTCHA Enterprise API)
 *   GCP_PROJECT_ID     → ID del proyecto GCP
 *   RECAPTCHA_SITE_KEY → Site Key pública (misma que en el frontend)
 * @param {string} rcToken  token enviado por el frontend
 * @param {number} minScore mínimo aceptable (0.0–1.0). Default 0.5
 * @returns {boolean}
 */
function verificarRecaptcha_(rcToken, minScore) {
  if (!rcToken || typeof rcToken !== 'string' || rcToken.length < 20) return false;

  const props    = PropertiesService.getScriptProperties();
  const API_KEY  = props.getProperty('RECAPTCHA_API_KEY');
  const PROJ_ID  = props.getProperty('GCP_PROJECT_ID');
  const SITE_KEY = props.getProperty('RECAPTCHA_SITE_KEY');

  if (!API_KEY || !PROJ_ID || !SITE_KEY) {
    Logger.log('reCAPTCHA Enterprise: faltan Script Properties (RECAPTCHA_API_KEY, GCP_PROJECT_ID, RECAPTCHA_SITE_KEY).');
    return false;
  }

  try {
    const url  = 'https://recaptchaenterprise.googleapis.com/v1/projects/' + PROJ_ID + '/assessments?key=' + API_KEY;
    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        event: { token: rcToken, siteKey: SITE_KEY, expectedAction: 'sea_portal_submit' }
      }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      Logger.log('reCAPTCHA Enterprise HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText());
      return false;
    }
    const r         = JSON.parse(resp.getContentText());
    const valid     = r.tokenProperties && r.tokenProperties.valid === true;
    const score     = (r.riskAnalysis && r.riskAnalysis.score) || 0;
    const threshold = (typeof minScore === 'number') ? minScore : 0.5;
    Logger.log('reCAPTCHA Enterprise: valid=' + valid + ' score=' + score + ' action=' + (r.tokenProperties && r.tokenProperties.action));
    return valid && score >= threshold;

  } catch (e) {
    Logger.log('verificarRecaptcha_ error: ' + e.message);
    return false;
  }
}

// ── respuestaNoAutorizado_ ────────────────────────────────────────────────
function respuestaNoAutorizado_(detalle) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: false,
      error: 'AUTH_REQUIRED',
      message: detalle || 'Autenticación requerida. Por favor inicia sesión.'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── verificarAcceso_ ──────────────────────────────────────────────────────
/**
 * Punto único de verificación de acceso para doGet y doPost.
 * Retorna null si el acceso es válido, o una respuesta de error si no.
 *
 * @param {string} action  nombre de la acción
 * @param {string|null} idToken  token Google (puede ser null)
 * @param {string|null} rcToken  token reCAPTCHA (puede ser null)
 * @returns {GoogleAppsScript.Content.TextOutput|null}
 */
function verificarAcceso_(action, idToken, rcToken) {
  const mode = AUTH_MODE[action];

  if (!mode) {
    // Acción no registrada → denegar por defecto
    return respuestaNoAutorizado_('Acción no autorizada.');
  }

  if (mode === 'GOOGLE' || mode === 'EITHER') {
    if (idToken) {
      const usuario = verificarIdToken_(idToken);
      if (!usuario) return respuestaNoAutorizado_('Token de sesión inválido o expirado.');
      const modulo = ACTION_MODULE[action] || '';
      if (!verificarUsuarioAutorizado_(usuario.email, modulo)) {
        return respuestaNoAutorizado_('Tu cuenta (' + usuario.email + ') no tiene acceso a este módulo.');
      }
      return null; // ✅ acceso permitido
    }
    if (mode === 'GOOGLE') {
      return respuestaNoAutorizado_('Se requiere iniciar sesión con Google.');
    }
    // Si mode === 'EITHER', intentar con reCAPTCHA
  }

  if (mode === 'RECAPTCHA' || mode === 'EITHER') {
    if (rcToken) {
      if (!verificarRecaptcha_(rcToken)) {
        return respuestaNoAutorizado_('Verificación de seguridad fallida. Intenta de nuevo.');
      }
      return null; // ✅ acceso permitido
    }
    return respuestaNoAutorizado_('Token de seguridad faltante. Recarga la página.');
  }

  return respuestaNoAutorizado_('Método de autenticación no soportado.');
}

// =========================================================================
// ENDPOINTS PRINCIPALES
// =========================================================================
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    // ── Verificar autenticación antes de cualquier operación ──────────────
    const idToken = data.id_token || null;
    const rcToken = data.recaptcha_token || null;
    const authError = verificarAcceso_(action, idToken, rcToken);
    if (authError) return authError;

    // Obtener email del usuario para trazabilidad (cacheado, bajo costo)
    const _usuario = idToken ? ((verificarIdToken_(idToken) || {}).email || 'desconocido') : 'portal_publico';

    switch(action) {
      case 'registrarCliente': return output_(fase1_RegistrarCliente(data));
      case 'registrarOT': return output_(fase2_RegistrarOT(data));
      case 'createExpediente': return output_(fase3_CrearExpediente(data));
      case 'addFilesToExpediente': return output_(fase3_AddFilesToExpediente(data));
      case 'updateEstatus': return output_(updateEstatusSafe_(data, _usuario));
      case 'updateEstatusInforme': return output_(updateEstatusInformeSafe_(data, _usuario));
      case 'updateResponsable': return output_(updateResponsableSafe_(data, _usuario));
      // Acciones GET migradas a POST para mantener token fuera de la URL (B-03)
      case 'verificarAcceso': {
        const u = verificarIdToken_(idToken || '');
        const mod = data.modulo || '';
        const ok = u && verificarUsuarioAutorizado_(u.email, mod);
        return output_({ success: !!ok, email: u ? u.email : '' });
      }
      case 'buscarClienteRFC': return output_(fase2_BuscarClienteRFC(data.rfc));
      case 'buscarClienteNombre': return output_(fase2_BuscarClienteNombre(data.nombre));
      case 'getTablero': return output_(fase4_GetTablero());
      case 'getTableroInf': return output_(fase4_GetTableroInf());
      case 'getOrdenes': return output_(getOrdenesSafe_());
      case 'getInformes': return output_(getInformesSafe_());
      case 'getConsecutivo': return output_(getConsecutivoSafe_(data));
      case 'updateRespInf': return output_(updateResponsableInformeSafe_(data, _usuario));
      case 'getRenovaciones': return output_(getRenovaciones());
      default: return output_({ success: false, error: 'Acción POST no reconocida.' });
    }
  } catch (err) {
    Logger.log('ERROR doPost [' + new Date().toISOString() + ']: ' + err.message + '\n' + (err.stack || ''));
    return output_({ success: false, error: 'Error en el servidor. Código: ERR-' + Date.now() });
  } finally {
    lock.releaseLock();
  }
}
function doGet(e) {
  if (!e || !e.parameter || !e.parameter.action) return ContentService.createTextOutput("API Ejecutiva Ambiental v3.0 - Multi-Sucursal Activa");
  const action = e.parameter.action;

  // ── Verificar autenticación ──────────────────────────────────────────────
  const idToken = e.parameter.id_token || null;
  const rcToken = e.parameter.recaptcha_token || null;
  const authError = verificarAcceso_(action, idToken, rcToken);
  if (authError) return authError;

  try {
    switch(action) {
      case 'verificarAcceso': {
        const u = verificarIdToken_(e.parameter.id_token || '');
        const mod = e.parameter.modulo || '';
        const ok = u && verificarUsuarioAutorizado_(u.email, mod);
        return output_({ success: !!ok, email: u ? u.email : '' });
      }
      case 'buscarClienteRFC': return output_(fase2_BuscarClienteRFC(e.parameter.rfc));
      case 'buscarClienteNombre': return output_(fase2_BuscarClienteNombre(e.parameter.nombre));
      case 'getTablero': return output_(fase4_GetTablero());
      case 'getTableroInf': return output_(fase4_GetTableroInf());
      case 'getOrdenes': return output_(getOrdenesSafe_());
      case 'getInformes': return output_(getInformesSafe_());
      case 'getConsecutivo': return output_(getConsecutivoSafe_(e.parameter));
      case 'getRenovaciones': return output_(getRenovaciones());
      default: return output_({ success: false, error: 'Acción GET no reconocida.' });
    }
  } catch (err) {
    Logger.log('ERROR doGet [' + new Date().toISOString() + ']: ' + err.message);
    return output_({ success: false, error: 'Error en el servidor. Código: ERR-' + Date.now() });
  }
}
// =========================================================================
// FASE 1: REGISTRO CLIENTE (16 Columnas)
// =========================================================================
function fase1_RegistrarCliente(data) {
  let logEntries = [];
  function addLog(message) { logEntries.push(`[${new Date().toISOString()}] ${message}`); Logger.log(message); }
  try {
    addLog('=== INICIO PROCESO MULTI-SUCURSAL ===');
    // Validar campos requeridos
    if (!data.razon_social || !String(data.razon_social).trim()) {
      return { success: false, error: 'El campo Razón Social es obligatorio.' };
    }
    const folderRaiz = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const rfcClean = (data.rfc || 'SIN_RFC').toUpperCase().trim();
    // Validar formato de RFC
    if (!validarRFC_(rfcClean)) {
      return { success: false, error: 'RFC inválido: "' + rfcClean + '". Formato requerido: 12 o 13 caracteres alfanuméricos (ej: XAXX010101000).' };
    }
    const companyClean = cleanCompanyName(data.razon_social || 'Cliente');
    const branchClean = sanitizeFileName(data.sucursal || 'Matriz');
    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd');
    // 1. LÓGICA DE CARPETA PADRE (Empresa)
    const parentFolderName = `${rfcClean} - ${companyClean}`;
    let parentFolder;
    const pIter = folderRaiz.getFoldersByName(parentFolderName);
    if (pIter.hasNext()) {
      parentFolder = pIter.next();
      addLog(`Carpeta Padre encontrada: ${parentFolderName}`);
    } else {
      parentFolder = folderRaiz.createFolder(parentFolderName);
      addLog(`Carpeta Padre creada: ${parentFolderName}`);
    }
    // 2. LÓGICA DE CARPETA HIJO (Sucursal)
    let carpetaCliente;
    const bIter = parentFolder.getFoldersByName(branchClean);
    if (bIter.hasNext()) {
      carpetaCliente = bIter.next();
      addLog(`Carpeta Sucursal encontrada. Archivos en: ${branchClean}`);
    } else {
      carpetaCliente = parentFolder.createFolder(branchClean);
      addLog(`Carpeta Sucursal creada: ${branchClean}`);
    }
    // 3. Guardar archivos y Excel
    const processedFiles = guardarArchivos(data, carpetaCliente, addLog);
    const sheetUrl = generarPerfilSheet(data, carpetaCliente, companyClean, branchClean, timestamp, addLog);
    // 4. REGISTRAR O ACTUALIZAR EN CLIENTES_MAESTRO (22 Columnas)
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_CLIENTES);

    const rowData = [
      Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss'), // 1. Fecha Registro
      data.razon_social || '',                                          // 2. Razón Social
      data.sucursal || '',                                              // 3. Sucursal
      data.rfc || '',                                                   // 4. RFC
      data.representante_legal || '',                                   // 5. Representante Legal
      data.direccion_evaluacion || '',                                  // 6. Dirección Evaluación
      data.telefono_empresa || '',                                      // 7. Teléfono Empresa
      data.nombre_solicitante || '',                                    // 8. Contacto (Solicitante)
      data.correo_informe || '',                                        // 9. Email Contacto
      data.giro || '',                                                  // 10. Giro / Actividad
      data.registro_patronal || '',                                     // 11. Registro Patronal
      data.capacidad_instalada || '',                                   // 12. Capacidad Instalada
      data.capacidad_operacion || '',                                   // 13. Capacidad Operación
      data.dias_turnos_horarios || '',                                  // 14. Días / Turnos
      data.aplica_nom020 === 'si' ? 'SÍ' : 'NO',                        // 15. Aplica NOM-020
      data.requiere_pipc === 'si' ? 'SÍ' : 'NO',                        // 16. Requiere PIPC
      data.responsable || '',                                           // 17. Nombre quien Atiende
      data.telefono_responsable || '',                                  // 18. Teléfono quien Atiende
      data.nombre_dirigido || '',                                       // 19. A quien se dirige informe
      data.puesto_dirigido || '',                                       // 20. Puesto a quien se dirige
      carpetaCliente.getUrl(),                                          // 21. Link Drive (índice 20)
      data.asesor_consultor || ''                                       // 22. Asesor / Consultor (índice 21)
    ];
    const allData = sheet.getDataRange().getValues();
    let rowIndex = -1;
    // Búsqueda backward → actualiza la fila más reciente (igual que el test),
    // evitando desalineación cuando hay filas residuales de ejecuciones previas.
    for (let i = allData.length - 1; i >= 1; i--) {
      if (String(allData[i][3]).toUpperCase().trim() === rfcClean &&
          String(allData[i][2]).trim() === branchClean) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      addLog(`Registro actualizado en la fila ${rowIndex}`);
    } else {
      sheet.appendRow(rowData);
      addLog('Nuevo registro agregado en BD.');
    }
    // 5. Enviar correo de notificación (omitir si _skipEmail está activo, p.ej. en tests)
    if (!data._skipEmail) {
      const emailResult = enviarNotificacionRobusta(data, processedFiles, carpetaCliente, sheetUrl, addLog);
      if (emailResult && !emailResult.success) {
        addLog('WARNING: Email no enviado: ' + (emailResult.error || 'error desconocido'));
      }
    }
    guardarLogEnDrive(carpetaCliente, logEntries, data);
    return { success: true, message: 'Registro/Actualización exitosa', files: processedFiles.length };
  } catch (error) {
    try { enviarEmailEmergencia(error, logEntries); } catch (e4) {}
    return { success: false, error: error.toString() };
  }
}
// =========================================================================
// FASE 2: BÚSQUEDA Y REGISTRO DE OT
// =========================================================================
function fase2_BuscarClienteRFC(rfcBuscado) {
  if(!rfcBuscado) return { found: false, error: 'RFC vacío' };
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_CLIENTES);
  const data = sheet.getDataRange().getValues();
  let sucursalesEncontradas = [];
  let razonSocialFija = "";
  let setSucursalesUnicas = new Set();
  for (let i = data.length - 1; i >= 1; i--) {
    const rfcFila = String(data[i][CL.RFC]).toUpperCase().trim();
    if (rfcFila === rfcBuscado.toUpperCase().trim()) {
      const nombreSucursal = String(data[i][CL.SUCURSAL]).trim() || 'Matriz';
      if (!setSucursalesUnicas.has(nombreSucursal)) {
        setSucursalesUnicas.add(nombreSucursal);
        if (!razonSocialFija) razonSocialFija = data[i][CL.RAZON_SOCIAL];
        sucursalesEncontradas.push({
          razon_social:        data[i][CL.RAZON_SOCIAL],
          sucursal:            nombreSucursal,
          rfc:                 data[i][CL.RFC],
          nombre_solicitante:  data[i][CL.SOLICITANTE],
          correo_informe:      data[i][CL.CORREO],
          telefono_empresa:    data[i][CL.TELEFONO],
          representante_legal: data[i][CL.REPRESENTANTE],
          direccion_evaluacion:data[i][CL.DIRECCION],
          giro:                data[i][CL.GIRO],
          registro_patronal:   data[i][CL.REGISTRO_PATRONAL],
          link_drive_cliente:  data[i][CL.LINK_DRIVE]
        });
      }
    }
  }
  if (sucursalesEncontradas.length > 0) {
    return { found: true, razon_social: razonSocialFija, sucursales: sucursalesEncontradas };
  }
  return { found: false };
}
function fase2_BuscarClienteNombre(nombreBuscado) {
  if(!nombreBuscado || nombreBuscado.length < 3) return { found: false, error: 'Nombre demasiado corto (mínimo 3 caracteres)' };
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_CLIENTES);
  const data = sheet.getDataRange().getValues();
  var resultados = [];
  var setUnicos = new Set();
  var termino = nombreBuscado.toUpperCase().trim();
  for (var i = data.length - 1; i >= 1; i--) {
    var razonSocial = String(data[i][CL.RAZON_SOCIAL]).toUpperCase().trim();
    if (razonSocial.indexOf(termino) !== -1) {
      var nombreSucursal = String(data[i][CL.SUCURSAL]).trim() || 'Matriz';
      var clave = razonSocial + '||' + nombreSucursal;
      if (!setUnicos.has(clave)) {
        setUnicos.add(clave);
        resultados.push({
          razon_social:        data[i][CL.RAZON_SOCIAL],
          sucursal:            nombreSucursal,
          rfc:                 data[i][CL.RFC],
          nombre_solicitante:  data[i][CL.SOLICITANTE],
          correo_informe:      data[i][CL.CORREO],
          telefono_empresa:    data[i][CL.TELEFONO],
          representante_legal: data[i][CL.REPRESENTANTE],
          direccion_evaluacion:data[i][CL.DIRECCION],
          giro:                data[i][CL.GIRO],
          registro_patronal:   data[i][CL.REGISTRO_PATRONAL],
          link_drive_cliente:  data[i][CL.LINK_DRIVE]
        });
      }
    }
  }
  if (resultados.length > 0) {
    return { found: true, resultados: resultados };
  }
  return { found: false };
}
function fase2_RegistrarOT(data) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  // 14 columnas A-N: FECHA, OT, TIPO, NOM, CLIENTE, SUCURSAL, RFC, PERSONAL,
  //                  FECHA_VISITA, FECHA_ENTREGA, FECHA_REAL, ESTATUS_EXTERNO, LINK_DRIVE, OBSERVACIONES
  sheet.appendRow([
    new Date(),
    data.ot_folio || '',
    data.tipo_orden || 'OTA',
    data.nom_servicio || '',
    data.cliente_razon_social || '',
    data.sucursal || '',
    data.rfc || '',
    data.personal_asignado || '',
    data.fecha_visita || '',
    data.fecha_entrega_limite || '',
    '',               // FECHA_REAL (col K) - vacía al crear
    'NO INICIADO',    // ESTATUS_EXTERNO (col L)
    data.link_drive_cliente || '',
    data.observaciones || ''
  ]);
  return { success: true, message: 'OT Registrada correctamente' };
}
// =========================================================================
// FASE 3: SISTEMA DE EXPEDIENTES
// =========================================================================
// Cadena de fallback para siempre encontrar la carpeta del cliente:
//   1) Hoja ORDENES_TRABAJO columna 14 (link_drive_cliente guardado al registrar OT)
//   2) Payload del frontend (linkDrive enviado por SEAINF)
//   3) Buscar en CLIENTES_MAESTRO por RFC + sucursal (índice [15], esquema 16 col)
//   4) Último recurso: carpeta raíz

function normalizeOtForSeainf_(ot) {
  return String(ot == null ? '' : ot).trim().toUpperCase();
}
/**
 * Busca la fila vigente de ORDENES_TRABAJO para una OT, recorriendo de abajo hacia arriba.
 * @param {Array<Array<*>>} sheetValues Valores de getDataRange().getValues()/getDisplayValues().
 * @param {string} ot OT a buscar.
 * @param {{minSheetRow:number}} options Opciones de búsqueda.
 * @returns {{arrayIndex:number, sheetRow:number}|null}
 */
function findOtRowForSeainf_(sheetValues, ot, options) {
  const values = Array.isArray(sheetValues) ? sheetValues : [];
  const normalizedOt = normalizeOtForSeainf_(ot);
  if (!normalizedOt || values.length < 2) return null;

  const opts = options || {};
  const minSheetRow = Math.max(2, Number(opts.minSheetRow) || 2);
  const startIndex = Math.max(1, minSheetRow - 1);

  for (let i = values.length - 1; i >= startIndex; i--) {
    if (normalizeOtForSeainf_(values[i][CO.OT]) === normalizedOt) {
      return { arrayIndex: i, sheetRow: i + 1 };
    }
  }
  return null;
}

function fase3_CrearExpediente(payload) {
  const info = payload.data || {};
  const files = payload.files || [];
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const values = sheet.getDataRange().getValues();
  const otMatch = findOtRowForSeainf_(values, info.ot, { minSheetRow: 2 });
  if (!otMatch) return { success: false, error: 'OT no encontrada.' };

  const filaOT = otMatch.sheetRow;
  const linkCarpetaSucursal = values[otMatch.arrayIndex][CO.LINK_DRIVE];
  // --- Cadena de fallback para encontrar la carpeta correcta ---
  let carpetaSucursal = null;
  // 1) Desde la hoja ORDENES_TRABAJO (columna 14)
  if (linkCarpetaSucursal) {
    var m1 = String(linkCarpetaSucursal).match(/folders\/([a-zA-Z0-9_-]+)/);
    if (m1) { try { carpetaSucursal = DriveApp.getFolderById(m1[1]); } catch(e) {} }
  }
  // 2) Desde el payload del frontend (linkDrive enviado por SEAINF)
  if (!carpetaSucursal && info.linkDrive) {
    var m2 = String(info.linkDrive).match(/folders\/([a-zA-Z0-9_-]+)/);
    if (m2) { try { carpetaSucursal = DriveApp.getFolderById(m2[1]); } catch(e) {} }
  }
  // 3) Buscar en CLIENTES_MAESTRO por RFC + sucursal
  //    NOTA: nuevo esquema 21 columnas → link Drive en índice [20] (columna 21)
  if (!carpetaSucursal && info.rfc) {
    try {
      var sheetCli = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_CLIENTES);
      var cliData = sheetCli.getDataRange().getValues();
      var rfcBusc = String(info.rfc).toUpperCase().trim();
      var sucBusc = String(info.sucursal || '').trim();
      for (var j = cliData.length - 1; j >= 1; j--) {
        if (String(cliData[j][CL.RFC]).toUpperCase().trim() === rfcBusc) {
          if (!sucBusc || String(cliData[j][CL.SUCURSAL]).trim() === sucBusc) {
            var linkCli = cliData[j][CL.LINK_DRIVE];
            if (linkCli) {
              var m3 = String(linkCli).match(/folders\/([a-zA-Z0-9_-]+)/);
              if (m3) { try { carpetaSucursal = DriveApp.getFolderById(m3[1]); } catch(e) {} }
            }
            break;
          }
        }
      }
    } catch(e) {
      Logger.log('Error buscando carpeta por RFC: ' + e.message);
    }
  }
  // 4) Último recurso: carpeta raíz
  if (!carpetaSucursal) {
    carpetaSucursal = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    Logger.log('ADVERTENCIA: Expediente creado en carpeta raíz porque no se encontró carpeta del cliente. OT: ' + info.ot);
  }
  var consecutivoMatch = info.numInforme.match(/-(\d{4})$/);
  var consecutivoPrefix = consecutivoMatch ? consecutivoMatch[1] : '0000';
  var nombreCarpetaOT = '02_Expediente_' + consecutivoPrefix + '_' + info.ot + '_' + info.nom;
  var carpetaOT = carpetaSucursal.createFolder(nombreCarpetaOT);
  const FS = CONFIG.FOLDER_STRUCTURE;
  var folders = {
    ORDEN_TRABAJO: carpetaOT.createFolder(FS.ORDEN_TRABAJO),
    HOJAS_CAMPO:   carpetaOT.createFolder(FS.HOJAS_CAMPO),
    CROQUIS:       carpetaOT.createFolder(FS.CROQUIS),
    FOTOS:         carpetaOT.createFolder(FS.FOTOS)
  };
  files.forEach(function(file) {
    if (!file || !file.content) return;
    try {
      var v = validarArchivo_(file.content, file.type || '', file.name || 'archivo');
      if (!v.valid) { Logger.log('Archivo rechazado: ' + v.error); return; }
      var decoded = Utilities.base64Decode(file.content);
      var blob = Utilities.newBlob(decoded, file.type, sanitizeFileName(file.name || 'archivo'));
      var targetFolder = folders[file.category] || carpetaOT;
      targetFolder.createFile(blob);
    } catch (err) { Logger.log('Error archivo: ' + err.message); }
  });
  // Escribir nueva fila en INFORMES (nunca tocar ORDENES_TRABAJO para datos de informe)
  const sheetInf = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_INFORMES);
  const otRow = values[otMatch.arrayIndex];
  sheetInf.appendRow([
    new Date(),                                    // A: Timestamp
    info.numInforme || '',                         // B: NumInforme
    otRow[CO.TIPO] || '',                          // C: TipoOrden
    info.ot || '',                                 // D: OT
    otRow[CO.NOM] || '',                           // E: NOM
    otRow[CO.CLIENTE] || '',                       // F: Cliente
    info.solicitante || '',                        // G: Solicitante
    otRow[CO.RFC] || '',                           // H: RFC
    info.telefono || '',                           // I: Telefono
    info.direccion || '',                          // J: Direccion
    info.fechaServicio || otRow[CO.FECHA_VISITA] || '', // K: FechaServicio
    otRow[CO.FECHA_ENTREGA] || '',                 // L: FechaEntrega
    info.esCapacitacion || 'NO',                   // M: EsCapacitacion
    'NO INICIADO',                                 // N: Estatus (independiente de ESTATUS_EXTERNO)
    carpetaOT.getUrl(),                            // O: LinkDrive (carpeta del expediente)
    otRow[CO.PERSONAL] || '',                      // P: Responsable
    otRow[CO.SUCURSAL] || ''                       // Q: Sucursal
  ]);
  return { success: true, url: carpetaOT.getUrl() };
}
function fase3_AddFilesToExpediente(payload) {
  const ot = payload.ot;
  const files = payload.files || [];
  if (!ot) return { success: false, error: 'Falta OT' };

  // Buscar el link del expediente en INFORMES (donde se guarda carpetaOT.getUrl())
  const sheetInf = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_INFORMES);
  const infData = sheetInf.getDataRange().getValues();
  const normalizedOt = normalizeOtForSeainf_(ot);
  let driveLink = '';
  for (let i = infData.length - 1; i >= 1; i--) {
    if (normalizeOtForSeainf_(infData[i][CI.OT]) === normalizedOt) {
      driveLink = infData[i][CI.LINK_DRIVE];
      break;
    }
  }
  if (!driveLink) return { success: false, error: 'No se encontró el expediente en INFORMES para OT: ' + ot };
  const folderIdMatch = driveLink.match(/folders\/([a-zA-Z0-9_-]+)/);
  const expedienteFolder = DriveApp.getFolderById(folderIdMatch[1]);
  const subfolderNames = CONFIG.FOLDER_STRUCTURE;
  const folders = {};
  const existingFolders = expedienteFolder.getFolders();
  while (existingFolders.hasNext()) {
    const f = existingFolders.next();
    for (const [key, name] of Object.entries(subfolderNames)) { if (f.getName() === name) folders[key] = f; }
  }
  for (const [key, name] of Object.entries(subfolderNames)) { if (!folders[key]) folders[key] = expedienteFolder.createFolder(name); }
  files.forEach(file => {
    if (!file || !file.content) return;
    try {
      const v = validarArchivo_(file.content, file.type || '', file.name || 'archivo');
      if (!v.valid) { Logger.log('Archivo rechazado en addFiles: ' + v.error); return; }
      const decoded = Utilities.base64Decode(file.content);
      const blob = Utilities.newBlob(decoded, file.type, sanitizeFileName(file.name || 'archivo'));
      const targetFolder = folders[file.category] || expedienteFolder;
      targetFolder.createFile(blob);
    } catch (err) { Logger.log('Error archivo addFiles: ' + err.message); }
  });
  return { success: true };
}
// =========================================================================
// getOrdenesSafe_ — devuelve rfc, sucursal, personal y nom para que SEAINF
// pueda encontrar la carpeta del cliente via fallback #3 (búsqueda por RFC)
// =========================================================================
function getOrdenesSafe_() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const values = sheet.getDataRange().getDisplayValues();
  const ordenes = values.slice(1).map(row => ({
    ot:             row[CO.OT],
    tipo_orden:     row[CO.TIPO],
    nom_servicio:   row[CO.NOM],
    clienteInicial: row[CO.CLIENTE],
    clienteFinal:   row[CO.SUCURSAL],
    cliente:        row[CO.CLIENTE],
    sucursal:       row[CO.SUCURSAL],
    rfc:            row[CO.RFC],
    personal:       row[CO.PERSONAL],
    fecha_visita:   row[CO.FECHA_VISITA],
    fecha_entrega:  row[CO.FECHA_ENTREGA],
    link_drive:     row[CO.LINK_DRIVE],
    estatus_externo: row[CO.ESTATUS_EXTERNO]
  })).filter(orden =>
    orden.ot && orden.ot.trim() !== '' &&
    ESTATUS_EXTERNO_TERMINALES_.indexOf(String(orden.estatus_externo || '').toUpperCase()) === -1
  );
  return { success: true, data: ordenes };
}
function getConsecutivoSafe_(params) {
  // Lee de INFORMES para encontrar el mayor consecutivo registrado
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_INFORMES);
  const dataRange = sheet.getDataRange().getDisplayValues().slice(1);
  const regex = /^EA-\d{4}-.+-(\d{4})$/;
  let maxConsecutivo = 0;
  dataRange.forEach(row => {
    const valNum = row[CI.NUM_INFORME];
    const valTipo = String(row[CI.TIPO_ORDEN] || '').trim().toUpperCase();
    if (valTipo === (params.tipo || 'OTA').toUpperCase()) {
      const match = String(valNum || '').trim().match(regex);
      if (match) {
        const consecutivo = parseInt(match[1], 10);
        if (consecutivo > maxConsecutivo) maxConsecutivo = consecutivo;
      }
    }
  });
  const siguiente = String(maxConsecutivo + 1).padStart(4, '0');
  return { success: true, numeroInforme: `EA-${params.anio}${params.mes}-${params.nom}-${siguiente}` };
}
function updateEstatusSafe_(data, usuario) {
  if (!data || !data.ot || !data.estatus) return { success: false, error: 'Faltan campos requeridos: ot, estatus.' };
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][CO.OT]).trim() === String(data.ot).trim()) {
      const valorAnterior = String(values[i][CO.ESTATUS_EXTERNO]);
      const nuevoEstatus = data.estatus.toUpperCase();
      sheet.getRange(i + 1, CO.ESTATUS_EXTERNO + 1).setValue(nuevoEstatus);
      if(nuevoEstatus === 'ENTREGADO' || nuevoEstatus === 'FINALIZADO') {
         sheet.getRange(i + 1, CO.FECHA_REAL + 1).setValue(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "dd/MM/yyyy"));
      }
      registrarAuditoria_(usuario || 'desconocido', 'UPDATE_ESTATUS_EXTERNO', data.ot, 'estatus_externo', valorAnterior, nuevoEstatus);
      return { success: true, message: 'Actualizado' };
    }
  }
  return { success: false, error: 'OT no encontrada' };
}
// Actualiza el estatus interno del informe en la hoja INFORMES (col N).
// NO toca ORDENES_TRABAJO ni ESTATUS_EXTERNO.
function updateEstatusInformeSafe_(data, usuario) {
  if (!data || !data.ot || !data.estatus) return { success: false, error: 'Faltan campos requeridos: ot, estatus.' };
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_INFORMES);
  const values = sheet.getDataRange().getValues();
  const normalizedOt = normalizeOtForSeainf_(data.ot);

  // Buscar la última fila de INFORMES que coincida con el OT
  for (let i = values.length - 1; i >= 1; i--) {
    if (normalizeOtForSeainf_(values[i][CI.OT]) === normalizedOt) {
      const valorAnterior = String(values[i][CI.ESTATUS]);
      const nuevoEstatus = String(data.estatus).toUpperCase();
      const targetRow = i + 1;
      const targetCol = CI.ESTATUS + 1;
      sheet.getRange(targetRow, targetCol).setValue(nuevoEstatus);
      SpreadsheetApp.flush();
      registrarAuditoria_(usuario || 'desconocido', 'UPDATE_ESTATUS_INFORME', data.ot, 'estatus_informe', valorAnterior, nuevoEstatus);
      return { success: true, message: 'Estatus informe actualizado' };
    }
  }
  return { success: false, error: 'Informe no encontrado en INFORMES para OT: ' + data.ot };
}
function updateResponsableSafe_(data, usuario) {
  if (!data || !data.ot || data.responsable === undefined) return { success: false, error: 'Faltan campos requeridos: ot, responsable.' };
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const values = sheet.getDataRange().getValues();
  const otMatch = findOtRowForSeainf_(values, data.ot, { minSheetRow: 2 });
  if (!otMatch) return { success: false, error: 'OT no encontrada' };

  const valorAnterior = String(values[otMatch.arrayIndex][CO.PERSONAL]);
  sheet.getRange(otMatch.sheetRow, CO.PERSONAL + 1).setValue(data.responsable);
  registrarAuditoria_(usuario || 'desconocido', 'UPDATE_RESPONSABLE', data.ot, 'personal_asignado', valorAnterior, data.responsable);
  return { success: true, _debug: { fila: otMatch.sheetRow } };
}
// Actualiza Responsable en la hoja INFORMES (col P = CI.RESPONSABLE).
// Usada por SEAINF via acción updateRespInf. No toca ORDENES_TRABAJO.
function updateResponsableInformeSafe_(data, usuario) {
  if (!data || !data.ot || data.responsable === undefined)
    return { success: false, error: 'Faltan campos requeridos: ot, responsable.' };
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_INFORMES);
  const values = sheet.getDataRange().getValues();
  const normalizedOt = normalizeOtForSeainf_(data.ot);
  for (let i = values.length - 1; i >= 1; i--) {
    if (normalizeOtForSeainf_(values[i][CI.OT]) === normalizedOt) {
      const valorAnterior = String(values[i][CI.RESPONSABLE]);
      sheet.getRange(i + 1, CI.RESPONSABLE + 1).setValue(data.responsable);
      SpreadsheetApp.flush();
      registrarAuditoria_(usuario || 'desconocido', 'UPDATE_RESPONSABLE_INFORME',
        data.ot, 'responsable', valorAnterior, data.responsable);
      return { success: true };
    }
  }
  return { success: false, error: 'Informe no encontrado en INFORMES para OT: ' + data.ot };
}
// =========================================================================
// FASE 4: TABLERO / DASHBOARD
// =========================================================================
function fase4_GetTablero() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_OT);
  const data = sheet.getDataRange().getDisplayValues();
  // Construir mapa RFC+Sucursal → asesor_consultor desde CLIENTES_MAESTRO
  const clientesSheet = ss.getSheetByName(CONFIG.SHEET_CLIENTES);
  const clientesData = clientesSheet.getDataRange().getDisplayValues();
  const asesorMap = {};
  clientesData.slice(1).forEach(row => {
    const rfc = String(row[CL.RFC] || '').toUpperCase().trim();
    const suc = String(row[CL.SUCURSAL] || '').trim();
    const asesor = String(row[CL.ASESOR_CONSULTOR] || '').trim();
    if (rfc && suc) asesorMap[rfc + '|' + suc] = asesor;
  });
  const latestRowsByOt = {};
  for (let i = data.length - 1; i >= 1; i--) {
    const normalizedOt = normalizeOtForSeainf_(data[i][CO.OT]);
    if (!normalizedOt || latestRowsByOt[normalizedOt]) continue;
    latestRowsByOt[normalizedOt] = data[i];
  }

  const registros = Object.keys(latestRowsByOt).map(normalizedOt => {
    const row = latestRowsByOt[normalizedOt];
    const rfc = String(row[CO.RFC] || '').toUpperCase().trim();
    const suc = String(row[CO.SUCURSAL] || '').trim();
    return {
      ot:               row[CO.OT],
      nom:              row[CO.NOM],
      cliente:          row[CO.CLIENTE],
      sucursal:         row[CO.SUCURSAL],
      rfc:              rfc,
      tipo_orden:       row[CO.TIPO],
      responsable:      row[CO.PERSONAL],
      fecha_visita:     row[CO.FECHA_VISITA],
      fechaEntrega:     row[CO.FECHA_ENTREGA],
      fechaRealEntrega: row[CO.FECHA_REAL],
      estatus:          row[CO.ESTATUS_EXTERNO],
      link_drive:       row[CO.LINK_DRIVE],
      asesor_consultor: asesorMap[rfc + '|' + suc] || ''
    };
  });
  return { success: true, data: registros };
}
// =========================================================================
// INFORMES — Lectura y dashboard interno (SEAINF)
// =========================================================================
function getInformesSafe_() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_INFORMES);
  const values = sheet.getDataRange().getDisplayValues();
  const informes = values.slice(1).map(row => ({
    numInforme:     row[CI.NUM_INFORME],
    tipoOrden:      row[CI.TIPO_ORDEN],
    ot:             row[CI.OT],
    nom:            row[CI.NOM],
    cliente:        row[CI.CLIENTE],
    sucursal:       row[CI.SUCURSAL],
    rfc:            row[CI.RFC],
    solicitante:    row[CI.SOLICITANTE],
    telefono:       row[CI.TELEFONO],
    direccion:      row[CI.DIRECCION],
    fechaServicio:  row[CI.FECHA_SERVICIO],
    fechaEntrega:   row[CI.FECHA_ENTREGA],
    esCapacitacion: row[CI.ES_CAPACITACION],
    estatus:        row[CI.ESTATUS],
    link_drive:     row[CI.LINK_DRIVE],
    responsable:    row[CI.RESPONSABLE]
  })).filter(inf => inf.ot && inf.ot.trim() !== '');
  return { success: true, data: informes };
}
function fase4_GetTableroInf() {
  return getInformesSafe_();
}
// =========================================================================
// TEST DIAGNÓSTICO — ejecutar manualmente desde el editor GAS
// =========================================================================
function testDiagnosticoEstatusInforme() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_OT);
  if (!sheet) { Logger.log('ERROR: hoja ORDENES_TRABAJO no encontrada'); return; }

  const values = sheet.getDataRange().getValues();
  Logger.log('Total filas (con encabezado): ' + values.length);
  Logger.log('Encabezados: ' + JSON.stringify(values[0]));
  Logger.log('col[15] del encabezado: "' + values[0][15] + '"');

  // Buscar primera fila con OT
  const primeraFila = values[1];
  if (!primeraFila) { Logger.log('No hay filas de datos'); return; }
  Logger.log('Primera OT encontrada: "' + primeraFila[1] + '"');
  Logger.log('Valor actual col P (idx15): "' + primeraFila[15] + '"');

  // Intentar escribir TEST en col P de la fila 2
  sheet.getRange(2, 16).setValue('DIAGNOSTICO_TEST');
  SpreadsheetApp.flush();
  const leido = sheet.getRange(2, 16).getValue();
  Logger.log('Escrito "DIAGNOSTICO_TEST", leído de vuelta: "' + leido + '"');
  if (leido === 'DIAGNOSTICO_TEST') {
    Logger.log('✅ WRITE FUNCIONA correctamente en col P');
    // Limpiar el test
    sheet.getRange(2, 16).setValue(primeraFila[15]);
    Logger.log('Valor restaurado a: "' + primeraFila[15] + '"');
  } else {
    Logger.log('❌ WRITE FALLÓ — posible protección de celda o rango');
  }
}
// =========================================================================
// UTILIDADES
// =========================================================================
function guardarArchivos(data, carpetaCliente, addLog) {
  const files = [];
  const fileFields = [
    { key: 'planos', label: '1) Planos generales' }, { key: 'mantenimiento', label: '2) Prog. Mantenimiento' },
    { key: 'proceso_produccion', label: '3) Proceso Producción' }, { key: 'requisitos_ingreso', label: '4) Requisitos Ingreso' },
    { key: 'ine_atencion', label: 'A) INE Atiende' }, { key: 'ine_testigo1', label: 'B) INE Testigo 1' },
    { key: 'ine_testigo2', label: 'C) INE Testigo 2' }, { key: 'poder_notarial', label: 'D) Poder Notarial (NOM-020)' },
    { key: 'ine_representante', label: 'E) INE Representante (NOM-020)' }, { key: 'situacion_fiscal', label: 'F) Sit. Fiscal (NOM-020)' },
    { key: 'licencia', label: 'G) Licencia/Cédula' }, { key: 'dc3', label: 'H) DC-3 Operador' },
    { key: 'calibracion_valvula', label: 'I) Calibración Válvula' }, { key: 'pipc_licencia_funcionamiento', label: 'PIPC - 1) Lic. Funcionamiento' },
    { key: 'pipc_uso_suelo', label: 'PIPC - 2) Uso de Suelo' }, { key: 'pipc_predial', label: 'PIPC - 3) Predial' },
    { key: 'pipc_poliza_seguro', label: 'PIPC - 4) Póliza Seguro' }, { key: 'pipc_mant_extintores', label: 'PIPC - 5) Mant. Extintores' },
    { key: 'pipc_situacion_fiscal', label: 'PIPC - 6) Sit. Fiscal' }, { key: 'pipc_ine_representante', label: 'PIPC - 7) INE Rep. Legal' },
    { key: 'pipc_acta_constitutiva', label: 'PIPC - 8) Acta Constitutiva' }, { key: 'pipc_poder_notarial', label: 'PIPC - 9) Poder Notarial' },
    { key: 'pipc_evidencia_simulacros', label: 'PIPC - 10) Simulacros' }, { key: 'pipc_organigrama_brigadas', label: 'PIPC - 11) Brigadas' },
    { key: 'pipc_detectores_humo', label: 'PIPC - 12) Detectores Humo' }, { key: 'pipc_medidas_preventivas', label: 'PIPC - 13) Medidas Preventivas' },
    { key: 'pipc_gas_natural', label: 'PIPC - 14) Gas Natural' }, { key: 'pipc_sustancias_quimicas', label: 'PIPC - 15) Sust. Químicas' },
    { key: 'pipc_dc3_operadores', label: 'PIPC - 16) DC3 Montacargas' }
  ];
  fileFields.forEach(field => {
    const fileData = data[field.key]; const fileName = data[field.key + '_filename'];
    if (fileData && fileName && typeof fileData === 'string' && fileData.startsWith('data:')) {
      try {
        const parts = fileData.split(',');
        const mimeMatch = parts[0].match(/:(.*?);/);
        if (!mimeMatch) { if (addLog) addLog('⚠️ MIME inválido en: ' + fileName); return; }
        const mimeType = mimeMatch[1];
        // Validar tamaño y tipo antes de decodificar
        const validacion = validarArchivo_(parts[1] || '', mimeType, fileName);
        if (!validacion.valid) { if (addLog) addLog('⚠️ Archivo rechazado: ' + validacion.error); return; }
        const blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), mimeType, sanitizeFileName(fileName));
        const file = carpetaCliente.createFile(blob);
        files.push({ name: fileName, label: field.label, url: file.getUrl(), size: data[field.key + '_size'] || 0 });
        if (addLog) addLog(`  - Archivo guardado: ${fileName}`);
      } catch (error) { if (addLog) addLog('⚠️ Error: ' + error.toString()); }
    }
  });
  return files;
}
function generarPerfilSheet(data, carpetaCliente, cleanedCompany, cleanedBranch, timestamp, addLog) {
  const ss = SpreadsheetApp.create(`${timestamp}_Perfil_${cleanedCompany}_${cleanedBranch}`);
  const sheet = ss.getSheets()[0]; sheet.setName('PERFIL DE DATOS');
  const verde = '#1e5a3e'; const valueBg = '#F0F0F0';
  const setHeader = (range, text) => { sheet.getRange(range).merge().setValue(text).setBackground(verde).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center'); };
  const fillFieldCell = (labelRange, label, valueCellA1, value, opt = {}) => {
    sheet.getRange(labelRange).merge().setValue(label).setFontWeight('bold');
    const r = sheet.getRange(valueCellA1); r.setValue(value || '').setBackground(valueBg).setBorder(true, true, true, true, false, false);
    if (opt.boldValue) r.setFontWeight('bold'); if (opt.wrap) r.setWrap(true).setVerticalAlignment('top');
  };
  sheet.getRange('A1:L1').merge().setValue('PERFIL DE DATOS TÉCNICOS - EJECUTIVA AMBIENTAL').setBackground(verde).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  fillFieldCell('B3:C3', 'Nombre solicitante:', 'D3', data.nombre_solicitante, { wrap: true });
  fillFieldCell('B5:C5', 'Razón Social:', 'D5', data.razon_social, { wrap: true });
  sheet.getRange('B6:C6').merge().setValue('SUCURSAL:').setFontWeight('bold').setFontColor('red');
  sheet.getRange('D6').setValue(data.sucursal || '').setBackground(valueBg).setBorder(true, true, true, true, false, false).setFontWeight('bold').setWrap(true);
  fillFieldCell('B7:C7', 'RFC:', 'D7', data.rfc);
  fillFieldCell('H7:H7', 'Teléfono Empresa:', 'I7', data.telefono_empresa);
  fillFieldCell('B9:C9', 'Representante Legal:', 'D9', data.representante_legal, { wrap: true });
  fillFieldCell('B11:C11', 'Dirección evaluación:', 'D11', data.direccion_evaluacion, { wrap: true });
  fillFieldCell('B13:C13', 'Responsable atiende:', 'D13', data.responsable, { wrap: true });
  fillFieldCell('H13:H13', 'Contacto Atiende:', 'I13', data.telefono_responsable);
  fillFieldCell('B15:C15', 'Giro:', 'D15', data.giro);
  fillFieldCell('B17:C17', 'Actividad:', 'D17', data.actividad_principal, { wrap: true });
  fillFieldCell('B19:C19', 'Registro Patronal:', 'D19', data.registro_patronal);
  fillFieldCell('B21:C21', 'Capacidad operación:', 'D21', data.capacidad_operacion, { wrap: true });
  fillFieldCell('H21:H21', 'Capacidad instalada:', 'I21', data.capacidad_instalada, { wrap: true });
  fillFieldCell('B23:C23', 'Horarios y Turnos:', 'D23', data.dias_turnos_horarios, { wrap: true });
  fillFieldCell('B26:C26', 'A quien se dirige:', 'D26', data.nombre_dirigido, { wrap: true });
  fillFieldCell('B28:C28', 'Puesto de quien dirige:', 'D28', data.puesto_dirigido, { wrap: true });
  fillFieldCell('B30:C30', 'Correo envío informe:', 'D30', data.correo_informe, { wrap: true });
  sheet.getRange('A32:L32').merge().setValue('DESCRIPCIÓN DEL PROCESO').setBackground(verde).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('A33:L36').merge().setValue(data.descripcion_proceso || '').setVerticalAlignment('top').setWrap(true).setBackground(valueBg).setBorder(true, true, true, true, false, false);
  sheet.getRange('A1:L36').setFontFamily('Arial');
  sheet.setColumnWidth(2, 170); sheet.setColumnWidth(4, 280); sheet.setColumnWidth(8, 160); sheet.setColumnWidth(9, 220);
  let currentRow = 38; setHeader(`A${currentRow}:L${currentRow}`, 'FECHAS PREFERIDAS');
  sheet.getRange(`B${currentRow+1}:L${currentRow+3}`).merge().setValue(data.fechas_preferidas || 'No especificadas').setBackground(valueBg).setWrap(true).setBorder(true, true, true, true, false, false);
  if (data.aplica_nom020) { setHeader(`A${currentRow+5}:L${currentRow+5}`, 'NOM-020-STPS'); sheet.getRange(`B${currentRow+6}:L${currentRow+6}`).merge().setValue(data.aplica_nom020 === 'si' ? 'SÍ APLICA' : 'NO APLICA').setHorizontalAlignment('center').setFontWeight('bold'); }
  const archivoSheet = DriveApp.getFileById(ss.getId()); archivoSheet.moveTo(carpetaCliente);
  return archivoSheet.getUrl();
}
function enviarNotificacionRobusta(data, files, carpetaCliente, sheetUrl, addLog) {
  let lastError = null;
  for (let attempt = 1; attempt <= CONFIG.EMAIL_RETRY_ATTEMPTS; attempt++) {
    try {
      enviarNotificacionEquipo(data, files, carpetaCliente, sheetUrl);
      Utilities.sleep(1000);
      enviarConfirmacionCliente(data, carpetaCliente);
      return { success: true };
    } catch (error) {
      lastError = error;
      if (attempt < CONFIG.EMAIL_RETRY_ATTEMPTS) Utilities.sleep(CONFIG.EMAIL_RETRY_DELAY_MS);
    }
  }
  try { enviarEmailSimpleFallback(data, carpetaCliente, sheetUrl); return { success: true, usedFallback: true }; }
  catch (fallbackError) { return { success: false, error: fallbackError.toString() }; }
}
function enviarNotificacionEquipo(data, files, carpetaCliente, sheetUrl) {
  const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
  let filesListHTML = '';
  if (files.length > 0) {
    files.forEach(f => { filesListHTML += `<tr><td style="padding: 8px 0; border-bottom: 1px solid #eee;"><a href="${f.url}" style="color:#1e5a3e; text-decoration:none; font-weight:600;">${f.label}</a></td><td style="padding: 8px 0; border-bottom: 1px solid #eee; text-align:right; color:#777; font-size:12px;">${formatFileSize(f.size)}</td></tr>`; });
  } else { filesListHTML = '<tr><td colspan="2" style="padding:10px; color:#999; font-style:italic;">No se adjuntaron archivos</td></tr>'; }
  const tagNom = data.aplica_nom020 === 'si' ? '<span style="background:#fff3cd; color:#856404; padding:4px 8px; border-radius:4px; font-weight:bold; font-size:11px;">SI APLICA</span>' : '<span style="background:#e8f5e9; color:#2e7d32; padding:4px 8px; border-radius:4px; font-weight:bold; font-size:11px;">NO APLICA</span>';
  const tagPipc = data.requiere_pipc === 'si' ? '<span style="background:#fff3cd; color:#856404; padding:4px 8px; border-radius:4px; font-weight:bold; font-size:11px;">SI REQUIERE</span>' : '<span style="background:#e8f5e9; color:#2e7d32; padding:4px 8px; border-radius:4px; font-weight:bold; font-size:11px;">NO REQUIERE</span>';
  const htmlBody = `<!DOCTYPE html><html><body style="margin:0; padding:0; background-color:#f4f6f8; font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;"><table width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color:#f4f6f8; padding:20px;"><tr><td align="center"><table width="600" border="0" cellpadding="0" cellspacing="0" style="background-color:#ffffff; border-radius:8px; overflow:hidden; box-shadow:0 4px 15px rgba(0,0,0,0.05);"><tr><td style="background-color:#1e5a3e; padding:30px; text-align:center;"><h1 style="color:#ffffff; margin:0; font-size:24px; font-weight:700;">Nuevo Perfil de Datos</h1><p style="color:#a8e6cf; margin:5px 0 0 0; font-size:14px;">${data.razon_social || 'Cliente Nuevo'}</p><p style="color:#a8e6cf; margin:3px 0 0 0; font-size:12px;">Sucursal: ${data.sucursal || 'N/A'}</p></td></tr><tr><td style="padding:30px;"><p style="text-align:right; font-size:11px; color:#999; margin-top:0;">Recibido: ${timestamp}</p><h3 style="color:#1e5a3e; border-bottom:2px solid #1e5a3e; padding-bottom:8px; margin-top:0;">Información de Contacto</h3><table width="100%" border="0" cellspacing="0" cellpadding="5" style="font-size:14px; color:#333; margin-bottom:20px;"><tr><td width="30%" style="font-weight:bold; color:#555;">Solicitante:</td><td>${data.nombre_solicitante || '-'}</td></tr><tr><td style="font-weight:bold; color:#555;">Sucursal:</td><td><strong>${data.sucursal || '-'}</strong></td></tr><tr><td style="font-weight:bold; color:#555;">Teléfono:</td><td><a href="tel:${data.telefono_empresa}" style="text-decoration:none; color:#333;">${data.telefono_empresa || '-'}</a></td></tr><tr><td style="font-weight:bold; color:#555;">Correo:</td><td><a href="mailto:${data.correo_informe}" style="color:#1e5a3e; font-weight:bold;">${data.correo_informe || '-'}</a></td></tr><tr><td style="font-weight:bold; color:#555;">Giro:</td><td>${data.giro || '-'}</td></tr></table><h3 style="color:#1e5a3e; border-bottom:2px solid #1e5a3e; padding-bottom:8px;">Servicios Solicitados</h3><table width="100%" border="0" cellspacing="0" cellpadding="5" style="font-size:14px; color:#333; margin-bottom:20px;"><tr><td width="50%"><strong>NOM-020-STPS:</strong> ${tagNom}</td><td width="50%"><strong>Prot. Civil (PIPC):</strong> ${tagPipc}</td></tr></table><div style="background-color:#fff8e1; border-left:4px solid #ffc107; padding:15px; margin-bottom:25px; border-radius:4px;"><strong style="color:#f57f17; font-size:12px; text-transform:uppercase;">FECHAS PREFERIDAS PARA EVALUACIÓN:</strong><p style="margin:8px 0 0 0; font-size:14px; color:#333; line-height:1.6;">${data.fechas_preferidas || 'No especificadas'}</p></div><h3 style="color:#1e5a3e; border-bottom:2px solid #1e5a3e; padding-bottom:8px;">Documentación Adjunta (${files.length})</h3><table width="100%" border="0" cellspacing="0" cellpadding="0" style="font-size:13px; color:#333;">${filesListHTML}</table><table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:30px;"><tr><td align="center"><a href="${carpetaCliente.getUrl()}" style="background-color:#1e5a3e; color:#ffffff; padding:12px 25px; text-decoration:none; border-radius:5px; font-weight:bold; font-size:14px; margin-right:10px; display:inline-block;">Ver Carpeta Drive</a><a href="${sheetUrl}" style="background-color:#2196f3; color:#ffffff; padding:12px 25px; text-decoration:none; border-radius:5px; font-weight:bold; font-size:14px; display:inline-block;">Ver Perfil Excel</a></td></tr></table></td></tr><tr><td style="background-color:#f8f9fa; padding:15px; text-align:center; border-top:1px solid #eee; font-size:11px; color:#888;">Sistema de Registro Automático v7.1 - ${CONFIG.COMPANY_NAME}<br>Este es un mensaje automático, no responder.</td></tr></table></td></tr></table></body></html>`;
  GmailApp.sendEmail(CONFIG.EMAIL_TO.join(','), `Nuevo Registro - ${data.razon_social} - ${data.sucursal}`, 'Su cliente de correo no soporta HTML.', { htmlBody: htmlBody, name: CONFIG.COMPANY_NAME });
}
function enviarConfirmacionCliente(data, carpetaCliente) {
  const emailCliente = data.correo_informe;
  if (!emailCliente || emailCliente.trim() === '') return;
  const htmlCliente = `<!DOCTYPE html><html><body style="margin:0; padding:0; background-color:#f4f6f8; font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;"><table width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color:#f4f6f8; padding:20px;"><tr><td align="center"><table width="600" border="0" cellpadding="0" cellspacing="0" style="background-color:#ffffff; border-radius:8px; overflow:hidden; box-shadow:0 4px 15px rgba(0,0,0,0.05);"><tr><td style="background-color:#1e5a3e; padding:30px; text-align:center;"><h1 style="color:#ffffff; margin:0; font-size:24px; font-weight:700;">¡Información Recibida!</h1><p style="color:#a8e6cf; margin:8px 0 0 0; font-size:14px;">Ejecutiva Ambiental</p></td></tr><tr><td style="padding:30px;"><p style="font-size:16px; color:#333; line-height:1.6; margin-top:0;">Estimado(a) <strong>${data.nombre_solicitante || 'Cliente'}</strong>,</p><p style="font-size:15px; color:#333; line-height:1.8;">Hemos recibido correctamente su <strong>Perfil de Datos</strong> para:</p><div style="background-color:#e8f5e9; border-left:4px solid #4caf50; padding:15px; margin:20px 0; border-radius:4px;"><p style="margin:0; font-size:14px; color:#2e7d32; line-height:1.6;"><strong>Empresa:</strong> ${data.razon_social}<br><strong>Sucursal:</strong> ${data.sucursal || 'N/A'}<br><strong>RFC:</strong> ${data.rfc || 'N/A'}</p></div><p style="font-size:15px; color:#333; line-height:1.8;">Nuestro equipo de <strong>Atención a Clientes</strong> revisará la información y se comunicará con usted en las próximas <strong>24 horas</strong> para coordinar los detalles del servicio.</p><div style="background-color:#fff3cd; border-left:4px solid #ffc107; padding:15px; margin:25px 0; border-radius:4px;"><p style="margin:0; font-size:13px; color:#856404; line-height:1.6;"><strong>¿Necesita realizar algún cambio?</strong><br>Por favor comuníquese con nosotros:<br><br><strong>${CONFIG.SUPPORT_PHONE}</strong><br><strong>${CONFIG.SUPPORT_EMAIL}</strong></p></div><p style="font-size:14px; color:#666; line-height:1.6;">Gracias por su confianza en <strong>Ejecutiva Ambiental</strong>.</p><p style="font-size:14px; color:#666; margin-bottom:0;">Atentamente,<br><strong style="color:#1e5a3e;">Equipo de Ejecutiva Ambiental</strong></p></td></tr><tr><td style="background-color:#f8f9fa; padding:15px; text-align:center; border-top:1px solid #eee; font-size:11px; color:#888;">Sistema de Registro Automático - ${CONFIG.COMPANY_NAME}<br>Este es un mensaje automático, por favor no responder a este correo.</td></tr></table></td></tr></table></body></html>`;
  GmailApp.sendEmail(emailCliente, '✓ Información Recibida - Ejecutiva Ambiental', 'Su cliente de correo no soporta HTML.', { htmlBody: htmlCliente, name: CONFIG.COMPANY_NAME });
}
function enviarEmailSimpleFallback(data, carpetaCliente, sheetUrl) {
  const subject = `NUEVO REGISTRO: ${data.razon_social} - ${data.sucursal}`;
  const body = `NUEVO PERFIL DE DATOS RECIBIDO\nCliente: ${data.razon_social}\nSucursal: ${data.sucursal}\nRFC: ${data.rfc}\nCarpeta Drive: ${carpetaCliente.getUrl()}`;
  GmailApp.sendEmail(CONFIG.EMAIL_TO.join(','), subject, body);
}
function enviarEmailEmergencia(error, logEntries) {
  GmailApp.sendEmail(CONFIG.EMAIL_TO[0], 'ERROR CRÍTICO - Sistema v2', `Error: ${error.toString()}\n\nLOGS:\n${logEntries.join('\n')}`);
}
function formatFileSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024; const sizes = ['B', 'KB', 'MB', 'GB']; const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}
function cleanCompanyName(name) { return sanitizeFileName((name || 'Cliente').replace(/ S\.A\. DE C\.V\.| SA DE CV| S\.A\.| S\.C\./gi, '').trim()); }
function sanitizeFileName(name) { return String(name || 'Sin_nombre').replace(/[^a-z0-9áéíóúñü ]/gi, '_').substring(0, 50); }
function guardarLogEnDrive(carpetaCliente, logEntries, data) {
  try { const blob = Utilities.newBlob(logEntries.join('\n'), 'text/plain', `LOG_${Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss')}.txt`); carpetaCliente.createFile(blob); } catch (e) {}
}
function output_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

// ── registrarAuditoria_ ──────────────────────────────────────────────────
/**
 * Registra cada mutación de datos en la hoja AUDITORIA del Spreadsheet.
 * Si la hoja no existe, la crea con encabezados automáticamente.
 *
 * @param {string} usuario       Email del operador que realizó el cambio
 * @param {string} accion        Identificador de la operación (ej: 'UPDATE_ESTATUS_EXTERNO')
 * @param {string} ot            Folio de la OT afectada
 * @param {string} campo         Nombre del campo modificado
 * @param {string} valorAnterior Valor antes del cambio
 * @param {string} valorNuevo    Valor después del cambio
 */
function registrarAuditoria_(usuario, accion, ot, campo, valorAnterior, valorNuevo) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_AUDITORIA);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_AUDITORIA);
      const headers = ['Timestamp', 'Usuario', 'Accion', 'OT', 'Campo', 'Valor_Anterior', 'Valor_Nuevo'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 180); sheet.setColumnWidth(2, 220);
      Logger.log('Hoja AUDITORIA creada automáticamente.');
    }
    sheet.appendRow([
      new Date(), usuario || 'desconocido', accion, ot, campo,
      valorAnterior !== undefined ? String(valorAnterior) : '',
      valorNuevo    !== undefined ? String(valorNuevo)    : ''
    ]);
  } catch (e) {
    Logger.log('registrarAuditoria_ error: ' + e.message);
  }
}

// ── validarRFC_ ───────────────────────────────────────────────────────────
/**
 * Valida el formato de un RFC mexicano (personas morales 12 chars, físicas 13).
 * Acepta la cadena especial 'SIN_RFC' para casos donde no aplica.
 */
function validarRFC_(rfc) {
  if (rfc === 'SIN_RFC') return true;
  return /^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$/.test(rfc);
}

// ── validarArchivo_ ───────────────────────────────────────────────────────
/**
 * Valida tamaño (desde longitud base64) y tipo MIME de un archivo.
 * Tamaño máximo: 50 MB. Solo tipos de negocio permitidos.
 */
function validarArchivo_(content, mimeType, fileName) {
  const MAX_BYTES = 50 * 1024 * 1024; // 50 MB
  const ALLOWED_MIMES = [
    'application/pdf',
    'image/jpeg', 'image/png', 'image/gif', 'image/webp',
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  ];
  // Estimación de bytes a partir de longitud base64 (evita decode costoso)
  const estimatedBytes = Math.ceil((content || '').length * 0.75);
  if (estimatedBytes > MAX_BYTES) {
    return { valid: false, error: 'Archivo "' + fileName + '" supera el límite de 50 MB.' };
  }
  if (!ALLOWED_MIMES.includes(mimeType)) {
    return { valid: false, error: 'Tipo de archivo no permitido: "' + mimeType + '". Solo PDF, imágenes y Office.' };
  }
  return { valid: true };
}
function autorizarPermisos() {
  var tempFolder = DriveApp.createFolder("Test_Permisos_EA");
  tempFolder.setTrashed(true);
  SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Permisos Drive + Sheets concedidos");
}
// Ejecutar UNA VEZ desde el editor GAS para autorizar el scope de Gmail.
// Aparecerá la pantalla de permisos; aceptar y redesplegar la webapp.
function autorizarGmail() {
  GmailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    'Test autorización Gmail — Ejecutiva Ambiental',
    'Si recibes este correo, el scope de Gmail está autorizado correctamente.'
  );
  Logger.log('Gmail autorizado. Redesplega la webapp.');
}

// =========================================================================
// RENOVACIONES — Alertas de servicios periódicos
// =========================================================================
function getRenovaciones() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheetOt = ss.getSheetByName(CONFIG.SHEET_OT);
  const sheetCl = ss.getSheetByName(CONFIG.SHEET_CLIENTES);

  // Servicios rastreados: clave normalizada (sin tildes, mayúsculas) → ciclo en años
  const SERVICIOS_RENOVACION = {
    'NOM-022-STPS':    1,
    'NOM-081-SEMARNAT':1,
    'NOM-025-STPS':    2,
    'NOM-024-STPS':    2,
    'NOM-015-STPS':    2,
    'PROGRAMA INTERNO':1,  // PIPC
    'PIPC':            1,
    'PROTECCION CIVIL':1
  };

  function normalizar_(str) {
    return String(str || '').toUpperCase().trim()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  function matchServicio_(nom) {
    const n = normalizar_(nom);
    for (const key in SERVICIOS_RENOVACION) {
      if (n.indexOf(key) !== -1) return { key: key, ciclo: SERVICIOS_RENOVACION[key] };
    }
    return null;
  }

  // Construir mapa RFC|Sucursal → email contacto desde CLIENTES_MAESTRO
  const clData = sheetCl.getDataRange().getDisplayValues().slice(1);
  const emailMap = {};
  clData.forEach(function(row) {
    const rfc = String(row[CL.RFC] || '').toUpperCase().trim();
    const suc = String(row[CL.SUCURSAL] || '').trim();
    if (rfc) emailMap[rfc + '|' + suc] = String(row[CL.CORREO] || '').trim();
  });

  // Leer OTs finalizadas con Fecha Real Entrega
  const otData = sheetOt.getDataRange().getDisplayValues().slice(1);
  const mapaUltimos = {};  // clave: RFC|Sucursal|ServicioKey → mejor entrada

  otData.forEach(function(row) {
    const estatus = String(row[CO.ESTATUS_EXTERNO] || '').toUpperCase().trim();
    if (estatus !== 'FINALIZADO') return;
    const fechaReal = String(row[CO.FECHA_REAL] || '').trim();
    if (!fechaReal) return;

    const nom = String(row[CO.NOM] || '').trim();
    const match = matchServicio_(nom);
    if (!match) return;

    const rfc = String(row[CO.RFC] || '').toUpperCase().trim();
    const suc = String(row[CO.SUCURSAL] || '').trim();
    const clave = rfc + '|' + suc + '|' + match.key;

    // Parsear fecha dd/mm/yyyy (formato es-MX de GAS getDisplayValues)
    const partes = fechaReal.split('/');
    let fechaObj;
    if (partes.length === 3) {
      fechaObj = new Date(+partes[2], +partes[1] - 1, +partes[0]);
    } else {
      fechaObj = new Date(fechaReal);
    }
    if (isNaN(fechaObj.getTime())) return;

    if (!mapaUltimos[clave] || fechaObj > mapaUltimos[clave].fecha) {
      mapaUltimos[clave] = {
        fecha:      fechaObj,
        fechaStr:   fechaReal,
        nom:        nom,
        cliente:    String(row[CO.CLIENTE]  || '').trim(),
        sucursal:   suc,
        rfc:        rfc,
        cicloAnios: match.ciclo,
        ot:         String(row[CO.OT]       || '').trim()
      };
    }
  });

  const hoy = new Date();
  const MS_DIA = 1000 * 60 * 60 * 24;

  function fmt_(d) {
    return String(d.getDate()).padStart(2,'0') + '/' +
           String(d.getMonth() + 1).padStart(2,'0') + '/' +
           d.getFullYear();
  }

  const renovaciones = Object.values(mapaUltimos).map(function(item) {
    const proxima = new Date(item.fecha);
    proxima.setFullYear(proxima.getFullYear() + item.cicloAnios);
    const diasRestantes = Math.ceil((proxima - hoy) / MS_DIA);
    return {
      cliente:           item.cliente,
      sucursal:          item.sucursal,
      rfc:               item.rfc,
      servicio:          item.nom,
      cicloAnios:        item.cicloAnios,
      ultimoServicio:    item.fechaStr,
      proximaRenovacion: fmt_(proxima),
      diasRestantes:     diasRestantes,
      emailContacto:     emailMap[item.rfc + '|' + item.sucursal] || '',
      ultimaOT:          item.ot
    };
  });

  // Ordenar: más urgentes primero (vencidos incluidos)
  renovaciones.sort(function(a, b) { return a.diasRestantes - b.diasRestantes; });

  return { success: true, data: renovaciones };
}

