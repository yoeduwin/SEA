// =========================================================================
// API CENTRAL - EJECUTIVA AMBIENTAL (SISTEMA UNIFICADO v3.0 - MULTI-SUCURSAL)
// VERSIÓN CORREGIDA — Listo para copiar/pegar en Google Apps Script
// =========================================================================
// CAMBIOS aplicados sobre el backend actual:
//   1. fase1_RegistrarCliente()  → NO envía correo al registrar; sí crea carpeta
//   2. fase2_BuscarClienteRFC/Nombre() → corregido índice link_drive_cliente [22]→[15]
//      (nuevo esquema CLIENTES_MAESTRO tiene 16 columnas; Drive link en col 16, índice 15)
//   3. fase3_CrearExpediente()   → fallback #3 corregido índice [22]→[15]
//   4. getOrdenesSafe_()         → ahora devuelve rfc, sucursal, personal, nom, fecha_visita
//      (sin rfc, SEAINF nunca podía buscar carpeta por RFC → expediente caía en raíz)
// =========================================================================
const CONFIG = {
  SPREADSHEET_ID: '1MoScea4CYg0NCjvPjHqZwV0cKhrd2nxfW8LYhz_4pDo',
  SHEET_CLIENTES: 'CLIENTES_MAESTRO',
  SHEET_OT: 'ORDENES_TRABAJO',
  FOLDER_ID: '1nHd-70uUeciClDm_3_pgbmqGF7II1lfQ',
  EMAIL_TO: [
    'direccion.general@ejecutivambiental.com',
    'operaciones@ejecutivambiental.com',
    'aclientes@ejecutivambiental.com'
  ],
  COMPANY_NAME: 'Ejecutiva Ambiental',
  EMAIL_RETRY_ATTEMPTS: 3,
  EMAIL_RETRY_DELAY_MS: 2000
};
// =========================================================================
// MÓDULO DE SEGURIDAD — Autenticación Google OAuth + reCAPTCHA v3
// =========================================================================
// Modos de autenticación por acción:
//   GOOGLE    → requiere id_token válido + usuario en whitelist
//   RECAPTCHA → requiere recaptcha_token válido (portales públicos de clientes)
//   EITHER    → acepta id_token (Google) o recaptcha_token
//
// Páginas internas con Google Auth: SEADB, SEAOT, SEAINF
// Portales públicos con reCAPTCHA:  paic.html, SEAPD.html (registro de clientes)
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
  updateEstatus:          'GOOGLE',
  updateResponsable:      'GOOGLE',
  // SEAOT
  buscarClienteRFC:       'EITHER',   // SEAOT usa Google Auth; paic/SEAPD (públicos) usan reCAPTCHA
  buscarClienteNombre:    'EITHER',   // igual que buscarClienteRFC
  registrarOT:            'GOOGLE',
  // SEAINF
  getOrdenes:             'GOOGLE',
  getConsecutivo:         'GOOGLE',
  createExpediente:       'GOOGLE',
  addFilesToExpediente:   'GOOGLE',
  updateEstatusInforme:   'GOOGLE',
  // ── Requiere reCAPTCHA (portales públicos de registro de clientes) ────────
  // paic.html y SEAPD.html son portales donde los CLIENTES se registran
  registrarCliente:       'RECAPTCHA',
  // ── Ping de verificación previo a la carga de la app ────────────────────
  // auth.js llama esto ANTES de mostrar la UI, para bloquear no autorizados
  verificarAcceso:        'GOOGLE'
};

// Mapeo acción → módulo (para verificar acceso por columna en la hoja)
const ACTION_MODULE = {
  getTablero:             'SEADB',
  updateEstatus:          'SEADB',
  updateResponsable:      'SEADB',
  buscarClienteRFC:       'SEAOT',
  buscarClienteNombre:    'SEAOT',
  registrarOT:            'SEAOT',
  getOrdenes:             'SEAINF',
  getConsecutivo:         'SEAINF',
  createExpediente:       'SEAINF',
  addFilesToExpediente:   'SEAINF',
  updateEstatusInforme:   'SEAINF'
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
    const sheet = ss.getSheetByName('USUARIOS_AUTORIZADOS');

    if (!sheet) {
      Logger.log('Hoja USUARIOS_AUTORIZADOS no existe. Ejecuta crearHojaUsuarios() desde el editor GAS.');
      return false;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return false;

    // Leer encabezados para encontrar columnas de módulos dinámicamente
    const headers = data[0].map(h => String(h).toUpperCase().replace(/[^A-Z0-9]/g, ''));
    const moduloCol = modulo ? headers.indexOf(modulo.toUpperCase()) : -1;

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][0]).toLowerCase().trim();
      const activo   = data[i][3]; // columna D: Activo

      if (rowEmail !== emailLower) continue;
      if (activo !== true) return false; // usuario desactivado

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
  const existing = ss.getSheetByName('USUARIOS_AUTORIZADOS');
  if (existing) {
    ss.deleteSheet(existing);
    Logger.log('Hoja anterior eliminada.');
  }

  // Crear nueva hoja
  const sheet = ss.insertSheet('USUARIOS_AUTORIZADOS');

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

  Logger.log('✅ Hoja USUARIOS_AUTORIZADOS creada con 4 usuarios.');
  Logger.log('   Para agregar/desactivar usuarios: edita directamente la hoja.');
  Logger.log('   Columnas SEADB/SEAOT/SEAINF: TRUE = acceso permitido, FALSE = bloqueado.');
  Logger.log('   El GAS refresca permisos en máximo 10 minutos (cache).');
}

// ── verificarRecaptcha_ ───────────────────────────────────────────────────
/**
 * Verifica un token de reCAPTCHA v3 con Google.
 * La Secret Key se obtiene de Script Properties (nunca en el frontend).
 * @param {string} rcToken  token enviado por el frontend
 * @param {number} minScore mínimo aceptable (0.0–1.0). Default 0.5
 * @returns {boolean}
 */
function verificarRecaptcha_(rcToken, minScore) {
  if (!rcToken || typeof rcToken !== 'string' || rcToken.length < 20) return false;

  const SECRET = PropertiesService.getScriptProperties().getProperty('RECAPTCHA_SECRET_KEY');
  if (!SECRET) {
    Logger.log('RECAPTCHA_SECRET_KEY no configurada en Script Properties.');
    return false;
  }

  try {
    const resp = UrlFetchApp.fetch('https://www.google.com/recaptcha/api/siteverify', {
      method: 'post',
      payload: { secret: SECRET, response: rcToken },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return false;

    const result = JSON.parse(resp.getContentText());
    const threshold = (typeof minScore === 'number') ? minScore : 0.5;
    Logger.log('reCAPTCHA: success=' + result.success + ' score=' + result.score + ' action=' + result.action);
    return result.success === true && (result.score || 0) >= threshold;

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

    switch(action) {
      case 'registrarCliente': return output_(fase1_RegistrarCliente(data));
      case 'registrarOT': return output_(fase2_RegistrarOT(data));
      case 'createExpediente': return output_(fase3_CrearExpediente(data));
      case 'addFilesToExpediente': return output_(fase3_AddFilesToExpediente(data));
      case 'updateEstatus': return output_(updateEstatusSafe_(data));
      case 'updateEstatusInforme': return output_(updateEstatusInformeSafe_(data));
      case 'updateResponsable': return output_(updateResponsableSafe_(data));
      default: return output_({ success: false, error: 'Acción POST no reconocida.' });
    }
  } catch (err) {
    return output_({ success: false, error: 'Error crítico en Servidor: ' + err.message });
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
      case 'verificarAcceso': return output_({ success: true, email: (function(){ const u = verificarIdToken_(e.parameter.id_token||''); return u ? u.email : ''; })() });
      case 'buscarClienteRFC': return output_(fase2_BuscarClienteRFC(e.parameter.rfc));
      case 'buscarClienteNombre': return output_(fase2_BuscarClienteNombre(e.parameter.nombre));
      case 'getTablero': return output_(fase4_GetTablero());
      case 'getOrdenes': return output_(getOrdenesSafe_());
      case 'getConsecutivo': return output_(getConsecutivoSafe_(e.parameter));
      default: return output_({ success: false, error: 'Acción GET no reconocida.' });
    }
  } catch (err) {
    return output_({ success: false, error: 'Error GET: ' + err.message });
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
    const folderRaiz = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const rfcClean = (data.rfc || 'SIN_RFC').toUpperCase().trim();
    const companyClean = cleanCompanyName(data.razon_social || 'Cliente');
    const branchClean = sanitizeFileName(data.sucursal || 'Matriz');
    const timestamp = Utilities.formatDate(new Date(), 'GMT-6', 'yyMMdd');
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
    // 4. REGISTRAR O ACTUALIZAR EN CLIENTES_MAESTRO (16 Columnas exactas)
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_CLIENTES);

    const capacidadInstalada = data.capacidad_instalada || 'N/A';
    const capacidadOperacion = data.capacidad_operacion || 'N/A';
    const capacidadUnida = `${capacidadInstalada} / ${capacidadOperacion}`;
    const rowData = [
      Utilities.formatDate(new Date(), 'GMT-6', 'dd/MM/yyyy HH:mm:ss'), // 1. Fecha Registro
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
      capacidadUnida,                                                   // 12. Capacidad Instalada / Operación
      data.dias_turnos_horarios || '',                                  // 13. Días / Turnos
      data.aplica_nom020 === 'si' ? 'SÍ' : 'NO',                        // 14. Aplica NOM-020
      data.requiere_pipc === 'si' ? 'SÍ' : 'NO',                        // 15. Requiere PIPC
      carpetaCliente.getUrl()                                           // 16. Link Drive (índice 15)
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
    const rfcFila = String(data[i][3]).toUpperCase().trim();
    if (rfcFila === rfcBuscado.toUpperCase().trim()) {
      const nombreSucursal = String(data[i][2]).trim() || 'Matriz';
      if (!setSucursalesUnicas.has(nombreSucursal)) {
        setSucursalesUnicas.add(nombreSucursal);
        if (!razonSocialFija) razonSocialFija = data[i][1];
        sucursalesEncontradas.push({
          razon_social: data[i][1],
          sucursal: nombreSucursal,
          rfc: data[i][3],
          nombre_solicitante: data[i][7],   // col 8
          correo_informe: data[i][8],        // col 9
          telefono_empresa: data[i][6],      // col 7
          representante_legal: data[i][4],   // col 5
          direccion_evaluacion: data[i][5],  // col 6
          giro: data[i][9],                  // col 10
          registro_patronal: data[i][10],    // col 11
          link_drive_cliente: data[i][15]    // col 16 (índice 15) — nuevo esquema 16 col
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
    var razonSocial = String(data[i][1]).toUpperCase().trim();
    if (razonSocial.indexOf(termino) !== -1) {
      var nombreSucursal = String(data[i][2]).trim() || 'Matriz';
      var clave = razonSocial + '||' + nombreSucursal;
      if (!setUnicos.has(clave)) {
        setUnicos.add(clave);
        resultados.push({
          razon_social: data[i][1],
          sucursal: nombreSucursal,
          rfc: data[i][3],
          nombre_solicitante: data[i][7],   // col 8
          correo_informe: data[i][8],        // col 9
          telefono_empresa: data[i][6],      // col 7
          representante_legal: data[i][4],   // col 5
          direccion_evaluacion: data[i][5],  // col 6
          giro: data[i][9],                  // col 10
          registro_patronal: data[i][10],    // col 11
          link_drive_cliente: data[i][15]    // col 16 (índice 15) — nuevo esquema 16 col
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
  sheet.appendRow([
    new Date(), data.ot_folio || '', data.tipo_orden || 'OTA', '', data.nom_servicio || '',
    data.cliente_razon_social || '', data.sucursal || '', data.rfc || '', data.personal_asignado || '',
    data.fecha_visita || '', data.fecha_entrega_limite || '', '', 'NO INICIADO',  // col 13 = estatus externo (SEADB)
    data.link_drive_cliente || '', data.observaciones || '',
    'NO INICIADO'  // col 16 = estatus_informe (interno SEAINF)
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
function fase3_CrearExpediente(payload) {
  const info = payload.data || {};
  const files = payload.files || [];
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const values = sheet.getDataRange().getValues();
  let filaOT = -1;
  let linkCarpetaSucursal = '';
  // Búsqueda backward → actualiza la fila más reciente, igual que el test,
  // evitando desalineación cuando hay filas residuales de TEST_FOLIO.
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][1]).trim() === String(info.ot).trim()) {
      filaOT = i + 1;
      linkCarpetaSucursal = values[i][13];
      break;
    }
  }
  if (filaOT === -1) return { success: false, error: 'OT no encontrada.' };
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
  //    NOTA: nuevo esquema 16 columnas → link Drive en índice [15] (columna 16)
  if (!carpetaSucursal && info.rfc) {
    try {
      var sheetCli = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_CLIENTES);
      var cliData = sheetCli.getDataRange().getValues();
      var rfcBusc = String(info.rfc).toUpperCase().trim();
      var sucBusc = String(info.sucursal || '').trim();
      for (var j = cliData.length - 1; j >= 1; j--) {
        if (String(cliData[j][3]).toUpperCase().trim() === rfcBusc) {
          if (!sucBusc || String(cliData[j][2]).trim() === sucBusc) {
            var linkCli = cliData[j][15]; // índice 15 = columna 16 (nuevo esquema)
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
  var folders = {
    ORDEN_TRABAJO: carpetaOT.createFolder('1. ORDEN_TRABAJO'),
    HOJAS_CAMPO:   carpetaOT.createFolder('2. HDC'),
    CROQUIS:       carpetaOT.createFolder('3. CROQUIS'),
    FOTOS:         carpetaOT.createFolder('4. FOTOS')
  };
  files.forEach(function(file) {
    if (!file || !file.content) return;
    try {
      var decoded = Utilities.base64Decode(file.content);
      var blob = Utilities.newBlob(decoded, file.type, file.name);
      var targetFolder = folders[file.category] || carpetaOT;
      targetFolder.createFile(blob);
    } catch (err) { Logger.log('Error archivo: ' + err.message); }
  });
  sheet.getRange(filaOT, 4).setValue(info.numInforme);
  sheet.getRange(filaOT, 13).setValue('EN PROCESO');
  sheet.getRange(filaOT, 14).setValue(carpetaOT.getUrl());
  return { success: true, url: carpetaOT.getUrl() };
}
function fase3_AddFilesToExpediente(payload) {
  const ot = payload.ot;
  const files = payload.files || [];
  if (!ot) return { success: false, error: 'Falta OT' };
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const data = sheet.getDataRange().getValues();
  let driveLink = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(ot).trim()) { driveLink = data[i][13]; break; }
  }
  if (!driveLink) return { success: false, error: 'No se encontró el expediente' };
  const folderIdMatch = driveLink.match(/folders\/([a-zA-Z0-9_-]+)/);
  const expedienteFolder = DriveApp.getFolderById(folderIdMatch[1]);
  const subfolderNames = { ORDEN_TRABAJO: '1. ORDEN_TRABAJO', HOJAS_CAMPO: '2. HDC', CROQUIS: '3. CROQUIS', FOTOS: '4. FOTOS' };
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
      const decoded = Utilities.base64Decode(file.content);
      const blob = Utilities.newBlob(decoded, file.type, file.name);
      const targetFolder = folders[file.category] || expedienteFolder;
      targetFolder.createFile(blob);
    } catch (err) {}
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
  const ordenes = values.slice(1).filter(row => {
    // Filtra por estatus INTERNO (col 16, índice 15): excluye informes ya finalizados/cancelados internamente
    // col 13 (estatus externo) lo gestiona SEADB de forma independiente
    const estatusInforme = String(row[15] || '').trim().toUpperCase();
    return estatusInforme !== 'FINALIZADO' && estatusInforme !== 'CANCELADO';
  }).map(row => ({
    ot: row[1],
    tipo_orden: row[2],
    nom_servicio: row[4],
    clienteInicial: row[5],
    clienteFinal: row[6],
    cliente: row[5],
    sucursal: row[6],
    rfc: row[7],
    personal: row[8],
    fecha_visita: row[9],
    link_drive: row[13],
    estatus_informe: row[15] || 'NO INICIADO'  // estado interno del dpto. informes
  })).filter(orden => orden.ot && orden.ot.trim() !== '');
  return { success: true, data: ordenes };
}
function getConsecutivoSafe_(params) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const dataRange = sheet.getDataRange().getDisplayValues().slice(1);
  const regex = /^EA-\d{4}-.+-(\d{4})$/;
  let maxConsecutivo = 0;
  dataRange.forEach(row => {
    const valNum = row[3];
    const valTipo = String(row[2] || '').trim().toUpperCase();
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
function updateEstatusSafe_(data) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]).trim() === String(data.ot).trim()) {
      sheet.getRange(i + 1, 13).setValue(data.estatus.toUpperCase());
      if(data.estatus.toUpperCase() === 'ENTREGADO' || data.estatus.toUpperCase() === 'FINALIZADO') {
         sheet.getRange(i + 1, 12).setValue(Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy"));
      }
      return { success: true, message: 'Actualizado' };
    }
  }
  return { success: false, error: 'OT no encontrada' };
}
// Actualiza SOLO el estatus interno del dpto. de informes (col 16).
// NO toca col 13 (estatus externo que lee SEADB) ni la fecha real de entrega.
function updateEstatusInformeSafe_(data) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]).trim() === String(data.ot).trim()) {
      sheet.getRange(i + 1, 16).setValue(data.estatus.toUpperCase()); // col 16 = Estatus Informe (interno)
      return { success: true, message: 'Estatus informe actualizado' };
    }
  }
  return { success: false, error: 'OT no encontrada' };
}
function updateResponsableSafe_(data) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]).trim() === String(data.ot).trim()) {
      sheet.getRange(i + 1, 9).setValue(data.responsable);
      return { success: true };
    }
  }
  return { success: false, error: 'OT no encontrada' };
}
// =========================================================================
// FASE 4: TABLERO / DASHBOARD
// =========================================================================
function fase4_GetTablero() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const data = sheet.getDataRange().getDisplayValues();
  const registros = data.slice(1).map(row => ({
    ot: row[1], numInforme: row[3], nom: row[4], cliente: row[5], sucursal: row[6],
    personal: row[8], fecha_visita: row[9], fechaEntrega: row[10], fechaRealEntrega: row[11],
    estatus: row[12], link_drive: row[13]
  })).reverse();
  return { success: true, data: registros };
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
        const parts = fileData.split(','); const mimeType = parts[0].match(/:(.*?);/)[1];
        const blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), mimeType, fileName);
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
  const timestamp = Utilities.formatDate(new Date(), 'GMT-6', 'dd/MM/yyyy HH:mm');
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
  const htmlCliente = `<!DOCTYPE html><html><body style="margin:0; padding:0; background-color:#f4f6f8; font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;"><table width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color:#f4f6f8; padding:20px;"><tr><td align="center"><table width="600" border="0" cellpadding="0" cellspacing="0" style="background-color:#ffffff; border-radius:8px; overflow:hidden; box-shadow:0 4px 15px rgba(0,0,0,0.05);"><tr><td style="background-color:#1e5a3e; padding:30px; text-align:center;"><h1 style="color:#ffffff; margin:0; font-size:24px; font-weight:700;">¡Información Recibida!</h1><p style="color:#a8e6cf; margin:8px 0 0 0; font-size:14px;">Ejecutiva Ambiental</p></td></tr><tr><td style="padding:30px;"><p style="font-size:16px; color:#333; line-height:1.6; margin-top:0;">Estimado(a) <strong>${data.nombre_solicitante || 'Cliente'}</strong>,</p><p style="font-size:15px; color:#333; line-height:1.8;">Hemos recibido correctamente su <strong>Perfil de Datos</strong> para:</p><div style="background-color:#e8f5e9; border-left:4px solid #4caf50; padding:15px; margin:20px 0; border-radius:4px;"><p style="margin:0; font-size:14px; color:#2e7d32; line-height:1.6;"><strong>Empresa:</strong> ${data.razon_social}<br><strong>Sucursal:</strong> ${data.sucursal || 'N/A'}<br><strong>RFC:</strong> ${data.rfc || 'N/A'}</p></div><p style="font-size:15px; color:#333; line-height:1.8;">Nuestro equipo de <strong>Atención a Clientes</strong> revisará la información y se comunicará con usted en las próximas <strong>24 horas</strong> para coordinar los detalles del servicio.</p><div style="background-color:#fff3cd; border-left:4px solid #ffc107; padding:15px; margin:25px 0; border-radius:4px;"><p style="margin:0; font-size:13px; color:#856404; line-height:1.6;"><strong>¿Necesita realizar algún cambio?</strong><br>Por favor comuníquese con nosotros:<br><br><strong>222 941 7295</strong><br><strong>aclientes@ejecutivambiental.com</strong></p></div><p style="font-size:14px; color:#666; line-height:1.6;">Gracias por su confianza en <strong>Ejecutiva Ambiental</strong>.</p><p style="font-size:14px; color:#666; margin-bottom:0;">Atentamente,<br><strong style="color:#1e5a3e;">Equipo de Ejecutiva Ambiental</strong></p></td></tr><tr><td style="background-color:#f8f9fa; padding:15px; text-align:center; border-top:1px solid #eee; font-size:11px; color:#888;">Sistema de Registro Automático - ${CONFIG.COMPANY_NAME}<br>Este es un mensaje automático, por favor no responder a este correo.</td></tr></table></td></tr></table></body></html>`;
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
  try { const blob = Utilities.newBlob(logEntries.join('\n'), 'text/plain', `LOG_${Utilities.formatDate(new Date(), 'GMT-6', 'yyyyMMdd_HHmmss')}.txt`); carpetaCliente.createFile(blob); } catch (e) {}
}
function output_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
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
