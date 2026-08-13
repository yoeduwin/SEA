// =========================================================================
// PORTAL — Portal de Clientes (Ejecutiva Ambiental)
// BACKEND · SOLO LECTURA de datos SEA · Proyecto Apps Script INDEPENDIENTE
// =========================================================================
//
// Qué hace:
//   Sirve (HtmlService) una página donde cada CLIENTE inicia sesión con un
//   código enviado a su correo y consulta ÚNICAMENTE sus propios informes,
//   con su estatus, y descarga el PDF del informe.
//
//   Lee (nunca escribe) las hojas CLIENTES_MAESTRO, ORDENES_TRABAJO e
//   INFORMES del Spreadsheet del sistema SEA.
//
// Principios (por decisión del negocio):
//   • NO MODIFICA nada del sistema actual. Es un proyecto Apps Script APARTE
//     (no toca BACKEND_FIXES.gs ni el proyecto interno). Igual que TRAZ.
//   • SOLO LECTURA de datos SEA: solo getDisplayValues() y lectura de Drive.
//     Jamás setValue/appendRow/insertSheet ni escritura sobre datos SEA.
//   • Autenticación propia de clientes (los clientes NO tienen cuenta Google
//     en la whitelist interna): código de un solo uso (OTP) por correo +
//     token de sesión firmado (HMAC). El alcance de datos SIEMPRE sale del
//     RFC firmado en el token, nunca de un parámetro suelto del cliente.
//
// Despliegue: proyecto Apps Script nuevo, App web, "ejecutar como yo",
//   acceso "cualquier usuario". Requiere Script Property PORTAL_HMAC_SECRET.
//   Ver PORTAL/README.md.
//
// NOTA GAS: las funciones invocables por google.script.run NO llevan guion
//   bajo final. Las funciones internas terminan en "_" (privadas).
// =========================================================================

// ── Configuración (espejo de solo lectura del esquema SEA) ────────────────
// Si el esquema del SEA cambia (se reordenan columnas), actualizar aquí.
var PORTAL_CONFIG = {
  SPREADSHEET_ID: '1MoScea4CYg0NCjvPjHqZwV0cKhrd2nxfW8LYhz_4pDo',

  SHEET_CLIENTES: 'CLIENTES_MAESTRO',
  SHEET_OT:       'ORDENES_TRABAJO',
  SHEET_INFORMES: 'INFORMES',

  // Índices de columna (0-based), idénticos a CONFIG.COLUMNS del SEA.
  COL_CLI: { RAZON_SOCIAL: 1, SUCURSAL: 2, RFC: 3, CORREO: 8, LINK_DRIVE: 20 },
  COL_OT:  {
    OT: 1, TIPO: 2, NOM: 3, CLIENTE: 4, SUCURSAL: 5, RFC: 6, PERSONAL: 7,
    FECHA_VISITA: 8, FECHA_ENTREGA: 9, FECHA_REAL: 10, ESTATUS_EXTERNO: 11,
    LINK_DRIVE: 12
  },
  COL_INF: {
    NUM_INFORME: 1, TIPO_ORDEN: 2, OT: 3, NOM: 4, CLIENTE: 5, RFC: 7,
    FECHA_SERVICIO: 10, FECHA_ENTREGA: 11, ES_CAPACITACION: 12, ESTATUS: 13,
    LINK_DRIVE: 14, SUCURSAL: 16
  },

  // Subcarpetas internas del expediente (NO se exponen al cliente).
  FOLDER_STRUCTURE: ['1. ORDEN_TRABAJO', '2. HDC', '3. CROQUIS', '4. FOTOS'],

  MODO_ENTREGA: 'PDF',        // 'PDF' = el portal entrega el PDF | 'CARPETA' = enlace de Drive
  SESION_DIAS: 30,            // vigencia del token de sesión
  OTP_TTL_SEG: 600,          // 10 min de validez del código
  OTP_MAX_INTENTOS: 5,       // intentos por código
  OTP_REENVIO_SEG: 60,       // 1 código por minuto por RFC
  MAX_PDF_BYTES: 25 * 1024 * 1024,  // tope para descarga directa (base64)

  ESTATUS_ENTREGADO: ['ENTREGADO', 'FINALIZADO'],
  COMPANY_NAME: 'Ejecutiva Ambiental',
  SOPORTE_TEL: '222 941 7295'
};

// =========================================================================
// PÁGINA (HtmlService)
// =========================================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('PORTAL')
    .setTitle('Portal de Clientes — Ejecutiva Ambiental')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

// =========================================================================
// LECTURA DE DATOS (solo getDisplayValues — nunca escribe)
// =========================================================================
function portal_leerHoja_(nombre) {
  var sheet = SpreadsheetApp.openById(PORTAL_CONFIG.SPREADSHEET_ID).getSheetByName(nombre);
  if (!sheet) return [];
  return sheet.getDataRange().getDisplayValues();
}

function portal_normRfc_(v) { return String(v == null ? '' : v).trim().toUpperCase(); }
function portal_normOt_(v)  { return String(v == null ? '' : v).trim().toUpperCase(); }
function portal_esHttp_(v)  { return /^https?:\/\//i.test(String(v || '').trim()); }

function portal_parseFecha_(s) {
  var m = String(s || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!m) return 0;
  var y = parseInt(m[3], 10); if (y < 100) y += 2000;
  return new Date(y, parseInt(m[2], 10) - 1, parseInt(m[1], 10)).getTime();
}

function portal_maskEmail_(email) {
  var s = String(email || '').trim();
  var at = s.indexOf('@');
  if (at < 1) return '';
  return s.slice(0, 1) + '***' + s.slice(at);
}

// =========================================================================
// SEGURIDAD — secreto, hash de código, token de sesión firmado (HMAC)
// =========================================================================
function portal_secret_() {
  var s = PropertiesService.getScriptProperties().getProperty('PORTAL_HMAC_SECRET');
  if (!s) throw new Error('Falta la propiedad de script PORTAL_HMAC_SECRET.');
  return s;
}

// base64url (web-safe). Se CONSERVA el padding "=" a propósito: el token no
// viaja en una URL (se pasa por google.script.run / localStorage) y así el
// decodificador de Apps Script reconstruye el payload sin ambigüedad.
function portal_b64url_(bytesOrStr) {
  return (typeof bytesOrStr === 'string')
    ? Utilities.base64EncodeWebSafe(Utilities.newBlob(bytesOrStr).getBytes())
    : Utilities.base64EncodeWebSafe(bytesOrStr);
}

// Comparación en tiempo constante (evita fugas por timing).
function portal_eq_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= (a.charCodeAt(i) ^ b.charCodeAt(i));
  return diff === 0;
}

function portal_hashCodigo_(rfc, codigo) {
  var raw = portal_normRfc_(rfc) + ':' + String(codigo) + ':' + portal_secret_();
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

function portal_firmarToken_(rfc) {
  var exp = Date.now() + PORTAL_CONFIG.SESION_DIAS * 24 * 60 * 60 * 1000;
  var payloadB64 = portal_b64url_(JSON.stringify({ rfc: portal_normRfc_(rfc), exp: exp }));
  var sigB64 = portal_b64url_(Utilities.computeHmacSha256Signature(payloadB64, portal_secret_()));
  return payloadB64 + '.' + sigB64;
}

// Devuelve { rfc } si el token es válido y no expiró; si no, null.
function portal_verificarToken_(token) {
  try {
    if (!token) return null;
    var parts = String(token).split('.');
    if (parts.length !== 2) return null;
    var esperado = portal_b64url_(Utilities.computeHmacSha256Signature(parts[0], portal_secret_()));
    if (!portal_eq_(parts[1], esperado)) return null;
    var json = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    var obj = JSON.parse(json);
    if (!obj || !obj.rfc || !obj.exp) return null;
    if (Date.now() > Number(obj.exp)) return null;
    return { rfc: portal_normRfc_(obj.rfc) };
  } catch (err) {
    return null;
  }
}

// =========================================================================
// CLIENTES — resolver RFC + correos a partir de RFC o correo escrito
// =========================================================================
function portal_buscarCliente_(identificador) {
  var id = String(identificador || '').trim();
  if (!id) return null;
  var idU = id.toUpperCase(), idL = id.toLowerCase();
  var CL = PORTAL_CONFIG.COL_CLI;
  var rows = portal_leerHoja_(PORTAL_CONFIG.SHEET_CLIENTES).slice(1);

  // 1) Resolver el RFC (por RFC exacto o por correo exacto).
  var rfc = '';
  for (var i = 0; i < rows.length; i++) {
    var rRfc = portal_normRfc_(rows[i][CL.RFC]);
    var rCorreo = String(rows[i][CL.CORREO] || '').trim().toLowerCase();
    if ((rRfc && rRfc === idU) || (rCorreo && rCorreo === idL)) { rfc = rRfc; break; }
  }
  if (!rfc) return null;

  // 2) Reunir todos los correos distintos y la razón social de ese RFC.
  var correosMap = {}, razon = '';
  rows.forEach(function (r) {
    if (portal_normRfc_(r[CL.RFC]) !== rfc) return;
    var c = String(r[CL.CORREO] || '').trim();
    if (c && c.indexOf('@') > 0) correosMap[c.toLowerCase()] = c;
    if (!razon) razon = String(r[CL.RAZON_SOCIAL] || '').trim();
  });
  var correos = Object.keys(correosMap).map(function (k) { return correosMap[k]; });
  return { rfc: rfc, razon_social: razon, correos: correos };
}

// =========================================================================
// AUTENTICACIÓN — solicitar código (OTP) y verificar
// =========================================================================
function portal_solicitarCodigo(payload) {
  try {
    payload = payload || {};
    var identificador = String(payload.identificador || '').trim();
    if (!identificador) return { success: false, error: 'Escribe tu RFC o tu correo.' };

    var cliente = portal_buscarCliente_(identificador);
    if (!cliente || cliente.correos.length === 0) {
      return { success: false, error: 'No encontramos un cliente con correo registrado para ese dato. Contáctanos al ' + PORTAL_CONFIG.SOPORTE_TEL + '.' };
    }

    var cache = CacheService.getScriptCache();
    var rlKey = 'otprl_' + cliente.rfc;
    if (cache.get(rlKey)) {
      return { success: false, error: 'Ya enviamos un código hace poco. Espera un minuto e inténtalo de nuevo.' };
    }

    var codigo = '' + Math.floor(100000 + Math.random() * 900000); // 6 dígitos
    var registro = {
      hash: portal_hashCodigo_(cliente.rfc, codigo),
      intentos: 0,
      exp: Date.now() + PORTAL_CONFIG.OTP_TTL_SEG * 1000
    };
    cache.put('otp_' + cliente.rfc, JSON.stringify(registro), PORTAL_CONFIG.OTP_TTL_SEG);
    cache.put(rlKey, '1', PORTAL_CONFIG.OTP_REENVIO_SEG);

    portal_enviarCodigo_(cliente, codigo);

    return {
      success: true,
      enviadoA: cliente.correos.map(portal_maskEmail_),
      razon_social: cliente.razon_social
    };
  } catch (err) {
    return { success: false, error: 'No se pudo enviar el código. Intenta más tarde.' };
  }
}

function portal_verificarCodigo(payload) {
  try {
    payload = payload || {};
    var identificador = String(payload.identificador || '').trim();
    var codigo = String(payload.codigo || '').trim();
    if (!identificador || !codigo) return { success: false, error: 'Faltan datos.' };

    var cliente = portal_buscarCliente_(identificador);
    if (!cliente) return { success: false, error: 'Código inválido o expirado.' };

    var cache = CacheService.getScriptCache();
    var key = 'otp_' + cliente.rfc;
    var raw = cache.get(key);
    if (!raw) return { success: false, error: 'El código expiró. Solicita uno nuevo.' };

    var registro = JSON.parse(raw);
    if (Date.now() > Number(registro.exp)) { cache.remove(key); return { success: false, error: 'El código expiró. Solicita uno nuevo.' }; }
    if (registro.intentos >= PORTAL_CONFIG.OTP_MAX_INTENTOS) { cache.remove(key); return { success: false, error: 'Demasiados intentos. Solicita un código nuevo.' }; }

    if (!portal_eq_(portal_hashCodigo_(cliente.rfc, codigo), registro.hash)) {
      registro.intentos += 1;
      var restanteSeg = Math.max(1, Math.round((Number(registro.exp) - Date.now()) / 1000));
      cache.put(key, JSON.stringify(registro), restanteSeg);
      return { success: false, error: 'Código incorrecto. Te quedan ' + (PORTAL_CONFIG.OTP_MAX_INTENTOS - registro.intentos) + ' intentos.' };
    }

    cache.remove(key); // un solo uso
    return {
      success: true,
      token: portal_firmarToken_(cliente.rfc),
      cliente: { rfc: cliente.rfc, razon_social: cliente.razon_social }
    };
  } catch (err) {
    return { success: false, error: 'No se pudo verificar el código.' };
  }
}

function portal_enviarCodigo_(cliente, codigo) {
  var asunto = 'Tu código de acceso · Portal de Clientes ' + PORTAL_CONFIG.COMPANY_NAME;
  var textoPlano = 'Tu código de acceso es: ' + codigo + '\n\nEs válido por 10 minutos. Si no lo solicitaste, ignora este mensaje.';
  var html =
    '<div style="font-family:Segoe UI,Roboto,Arial,sans-serif;max-width:460px;margin:0 auto;border:1px solid #e0e0e0;border-radius:12px;overflow:hidden">' +
      '<div style="background:#1e5a3e;color:#fff;padding:18px 22px;font-weight:700;letter-spacing:.5px">' + PORTAL_CONFIG.COMPANY_NAME + ' · Portal de Clientes</div>' +
      '<div style="padding:24px 22px;color:#2c3e50">' +
        '<p style="margin:0 0 10px">Hola' + (cliente.razon_social ? ' <strong>' + portal_escHtml_(cliente.razon_social) + '</strong>' : '') + ',</p>' +
        '<p style="margin:0 0 16px">Usa este código para entrar al portal:</p>' +
        '<div style="font-size:34px;font-weight:800;letter-spacing:8px;color:#1e5a3e;background:#e8f5e9;border-radius:10px;text-align:center;padding:16px 0">' + codigo + '</div>' +
        '<p style="margin:16px 0 0;font-size:13px;color:#7f8c8d">El código es válido por 10 minutos. Si tú no lo solicitaste, ignora este correo.</p>' +
      '</div>' +
    '</div>';
  GmailApp.sendEmail(cliente.correos.join(','), asunto, textoPlano, {
    name: PORTAL_CONFIG.COMPANY_NAME,
    htmlBody: html
  });
}

// Escapa texto para el correo HTML (el frontend tiene su propio esc()).
function portal_escHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}

// =========================================================================
// DATOS DEL CLIENTE — informes + estatus, filtrados por el RFC del token
// =========================================================================
function portal_misServicios(token) {
  try {
    var sesion = portal_verificarToken_(token);
    if (!sesion) return { success: false, error: 'AUTH_REQUIRED' };
    var rfc = sesion.rfc;
    var CO = PORTAL_CONFIG.COL_OT, CI = PORTAL_CONFIG.COL_INF;

    // OTs del cliente por folio (para estatus externo).
    var ots = portal_leerHoja_(PORTAL_CONFIG.SHEET_OT).slice(1);
    var otPorFolio = {};
    ots.forEach(function (r) {
      if (portal_normRfc_(r[CO.RFC]) !== rfc) return;
      var folio = portal_normOt_(r[CO.OT]);
      if (folio) otPorFolio[folio] = r; // la última fila de la hoja gana
    });

    var servicios = [];
    var foliosConInforme = {};
    var nombre = '';

    // 1) Informes del cliente (cada informe es una tarjeta).
    var infs = portal_leerHoja_(PORTAL_CONFIG.SHEET_INFORMES).slice(1);
    infs.forEach(function (r) {
      if (portal_normRfc_(r[CI.RFC]) !== rfc) return;
      if (!nombre) nombre = String(r[CI.CLIENTE] || '').trim();
      var folioOt = portal_normOt_(r[CI.OT]);
      if (folioOt) foliosConInforme[folioOt] = true;
      var otRow = otPorFolio[folioOt] || null;
      var estatus = otRow ? String(otRow[CO.ESTATUS_EXTERNO] || '').trim() : String(r[CI.ESTATUS] || '').trim();
      if (estatus.toUpperCase() === 'CANCELADO') return;
      var entregado = PORTAL_CONFIG.ESTATUS_ENTREGADO.indexOf(estatus.toUpperCase()) !== -1;
      servicios.push({
        numInforme:   String(r[CI.NUM_INFORME] || '').trim(),
        nom:          String(r[CI.NOM] || '').trim(),
        sucursal:     String(r[CI.SUCURSAL] || '').trim(),
        fechaServicio: String(r[CI.FECHA_SERVICIO] || '').trim(),
        fechaEntrega: String(r[CI.FECHA_ENTREGA] || '').trim(),
        estatus:      estatus,
        entregado:    entregado,
        descargable:  entregado && portal_esHttp_(r[CI.LINK_DRIVE])
      });
    });

    // 2) OTs del cliente sin informe todavía (trabajos en proceso).
    Object.keys(otPorFolio).forEach(function (folio) {
      if (foliosConInforme[folio]) return;
      var r = otPorFolio[folio];
      var estatus = String(r[CO.ESTATUS_EXTERNO] || '').trim();
      if (estatus.toUpperCase() === 'CANCELADO') return;
      if (!nombre) nombre = String(r[CO.CLIENTE] || '').trim();
      servicios.push({
        numInforme:   '',
        nom:          String(r[CO.NOM] || '').trim(),
        sucursal:     String(r[CO.SUCURSAL] || '').trim(),
        fechaServicio: String(r[CO.FECHA_VISITA] || '').trim(),
        fechaEntrega: String(r[CO.FECHA_ENTREGA] || '').trim(),
        estatus:      estatus,
        entregado:    false,
        descargable:  false
      });
    });

    // Orden: más reciente primero (por fecha de entrega, luego de servicio).
    servicios.sort(function (a, b) {
      var fa = portal_parseFecha_(a.fechaEntrega) || portal_parseFecha_(a.fechaServicio);
      var fb = portal_parseFecha_(b.fechaEntrega) || portal_parseFecha_(b.fechaServicio);
      return fb - fa;
    });

    return { success: true, cliente: { rfc: rfc, razon_social: nombre }, servicios: servicios };
  } catch (err) {
    return { success: false, error: 'No se pudieron cargar tus informes.' };
  }
}

// =========================================================================
// DESCARGA — entrega el PDF del informe (base64) o el enlace de carpeta
// =========================================================================
function portal_descargarInforme(token, numInforme) {
  try {
    var sesion = portal_verificarToken_(token);
    if (!sesion) return { success: false, error: 'AUTH_REQUIRED' };
    var rfc = sesion.rfc;
    var num = String(numInforme || '').trim();
    if (!num) return { success: false, error: 'Informe no especificado.' };

    var CI = PORTAL_CONFIG.COL_INF;
    var infs = portal_leerHoja_(PORTAL_CONFIG.SHEET_INFORMES).slice(1);
    var row = null;
    for (var i = 0; i < infs.length; i++) {
      if (String(infs[i][CI.NUM_INFORME] || '').trim() === num && portal_normRfc_(infs[i][CI.RFC]) === rfc) { row = infs[i]; break; }
    }
    if (!row) return { success: false, error: 'Informe no encontrado.' };

    // Defensa en profundidad: solo se entrega si el servicio está ENTREGADO/
    // FINALIZADO en el servidor (no basta con que la UI muestre el botón).
    var estatus = portal_estatusDeInforme_(row, rfc);
    if (PORTAL_CONFIG.ESTATUS_ENTREGADO.indexOf(String(estatus).toUpperCase()) === -1) {
      return { success: false, error: 'Este informe aún no está disponible para descarga.' };
    }

    var link = String(row[CI.LINK_DRIVE] || '').trim();
    if (!portal_esHttp_(link)) return { success: false, error: 'Este informe aún no tiene archivo disponible.' };

    if (PORTAL_CONFIG.MODO_ENTREGA === 'CARPETA') return { success: true, modo: 'carpeta', url: link };

    var folderId = portal_extraerIdDrive_(link);
    if (!folderId) return { success: true, modo: 'carpeta', url: link };

    var archivo = portal_buscarPdfInforme_(folderId, num);
    if (!archivo) return { success: true, modo: 'carpeta', url: link }; // fallback: sin PDF en la raíz

    var bytes = archivo.getBlob().getBytes();
    if (bytes.length > PORTAL_CONFIG.MAX_PDF_BYTES) {
      return { success: true, modo: 'carpeta', url: link, aviso: 'El archivo es muy grande para descarga directa.' };
    }
    return {
      success: true,
      modo: 'archivo',
      nombre: archivo.getName(),
      mimeType: archivo.getMimeType() || 'application/pdf',
      base64: Utilities.base64Encode(bytes)
    };
  } catch (err) {
    return { success: false, error: 'No se pudo obtener el informe.' };
  }
}

// Estatus (externo) del servicio al que pertenece un informe: se toma de la OT
// por folio (misma dueña del RFC) y, si no hay OT, del propio informe.
function portal_estatusDeInforme_(infRow, rfc) {
  var CO = PORTAL_CONFIG.COL_OT, CI = PORTAL_CONFIG.COL_INF;
  var folio = portal_normOt_(infRow[CI.OT]);
  if (folio) {
    var ots = portal_leerHoja_(PORTAL_CONFIG.SHEET_OT).slice(1);
    for (var i = ots.length - 1; i >= 0; i--) {
      if (portal_normOt_(ots[i][CO.OT]) === folio && portal_normRfc_(ots[i][CO.RFC]) === rfc) {
        return String(ots[i][CO.ESTATUS_EXTERNO] || '').trim();
      }
    }
  }
  return String(infRow[CI.ESTATUS] || '').trim();
}

// Extrae el ID de una URL de carpeta/archivo de Drive.
function portal_extraerIdDrive_(url) {
  var s = String(url || '');
  var m = s.match(/\/folders\/([a-zA-Z0-9_-]{10,})/); if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);          if (m) return m[1];
  m = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/);            if (m) return m[1];
  return '';
}

// Busca el PDF del informe SOLO en la raíz del expediente (no en subcarpetas
// internas 1..4, que nunca se exponen al cliente).
function portal_buscarPdfInforme_(folderId, num) {
  var folder;
  try { folder = DriveApp.getFolderById(folderId); } catch (e) { return null; }
  var it = folder.getFilesByType(MimeType.PDF);
  var pdfs = [];
  while (it.hasNext()) pdfs.push(it.next());
  if (pdfs.length === 0) return null;
  if (pdfs.length === 1) return pdfs[0];

  var numNorm = String(num).replace(/\s+/g, '').toUpperCase();
  for (var i = 0; i < pdfs.length; i++) {
    if (pdfs[i].getName().replace(/\s+/g, '').toUpperCase().indexOf(numNorm) !== -1) return pdfs[i];
  }
  pdfs.sort(function (a, b) { return b.getLastUpdated().getTime() - a.getLastUpdated().getTime(); });
  return pdfs[0];
}

// =========================================================================
// PRUEBA (ejecutar en el editor GAS; requiere PORTAL_HMAC_SECRET configurado)
// =========================================================================
function test_portal() {
  var rfc = 'XTES000000TST'; // RFC de prueba usado por TESTS_BACKEND.gs
  var log = [];

  // 1) Token: firma / verifica / rechaza manipulación.
  var t = portal_firmarToken_(rfc);
  var v = portal_verificarToken_(t);
  log.push('[' + (v && v.rfc === rfc ? 'PASS' : 'FAIL') + '] token válido verifica al RFC');
  log.push('[' + (portal_verificarToken_(t + 'x') === null ? 'PASS' : 'FAIL') + '] token manipulado se rechaza');
  log.push('[' + (portal_verificarToken_('') === null ? 'PASS' : 'FAIL') + '] token vacío se rechaza');

  // 2) Hash de código: determinista y sensible al código.
  log.push('[' + (portal_hashCodigo_(rfc, '123456') === portal_hashCodigo_(rfc, '123456') ? 'PASS' : 'FAIL') + '] hash estable');
  log.push('[' + (portal_hashCodigo_(rfc, '123456') !== portal_hashCodigo_(rfc, '654321') ? 'PASS' : 'FAIL') + '] hash cambia con el código');

  // 3) Datos: misServicios solo del RFC del token, sin fugas de otros RFC.
  var res = portal_misServicios(t);
  log.push('[' + (res.success ? 'PASS' : 'FAIL') + '] misServicios responde (servicios=' + (res.servicios ? res.servicios.length : 'n/a') + ')');
  var sinToken = portal_misServicios('token-invalido');
  log.push('[' + (!sinToken.success && sinToken.error === 'AUTH_REQUIRED' ? 'PASS' : 'FAIL') + '] misServicios exige token válido');

  // 4) Descarga: exige token válido.
  var d = portal_descargarInforme('token-invalido', 'EA-0000-NOM000-0000');
  log.push('[' + (!d.success && d.error === 'AUTH_REQUIRED' ? 'PASS' : 'FAIL') + '] descargarInforme exige token válido');

  Logger.log('=== test_portal ===\n' + log.join('\n'));
  return log;
}
