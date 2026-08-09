// =========================================================================
// TRAZ — Módulo de Trazabilidad de Servicios (Ejecutiva Ambiental)
// BACKEND · SOLO LECTURA · v1
// =========================================================================
//
// Qué hace:
//   Lee (nunca escribe) las hojas ORDENES_TRABAJO, INFORMES y AUDITORIA del
//   Spreadsheet del sistema SEA y construye la TRAZABILIDAD de un servicio:
//
//        ORDEN DE TRABAJO  →  EXPEDIENTE (carpeta Drive)  →  INFORME(S)
//
//   La trazabilidad se organiza POR OT relacionada con su o sus informe(s).
//
// Principios (por decisión del negocio):
//   • SOLO LECTURA. Este módulo JAMÁS modifica datos existentes. No usa
//     setValue/appendRow/insertSheet ni escribe en Drive. Es "intocable".
//   • Alimentado ÚNICAMENTE por lo que ya existe. No agrega columnas ni campos.
//   • La hoja AUDITORIA es la bitácora real: se muestran los eventos REALMENTE
//     registrados, sin inferir acciones que no estén en el log.
//   • El expediente es INFORMES.LINK_DRIVE (col O). No se crea otro repositorio.
//   • La relación OT ↔ Informe es por FOLIO de OT (texto). El id interno queda
//     como mejora futura.
//
// Despliegue: es una app web de Apps Script INDEPENDIENTE del backend SEA
//   (no se toca BACKEND_FIXES.gs). Ver TRAZ/README.md.
// =========================================================================

// ── Configuración (espejo de solo lectura del esquema SEA) ────────────────
// Si el esquema del SEA cambia, actualizar estos índices aquí también.
var TRAZ_CONFIG = {
  SPREADSHEET_ID: '1MoScea4CYg0NCjvPjHqZwV0cKhrd2nxfW8LYhz_4pDo',

  SHEET_OT:        'ORDENES_TRABAJO',
  SHEET_INFORMES:  'INFORMES',
  SHEET_USUARIOS:  'USUARIOS_AUTORIZADOS',
  SHEET_AUDITORIA: 'AUDITORIA',

  // Módulo de USUARIOS_AUTORIZADOS que da acceso a la trazabilidad.
  // Se reutiliza la columna 'SEAINF' existente — NO se agrega ninguna columna.
  MODULO_ACCESO: 'SEAINF',

  // Índices de columna (0-based), idénticos a CONFIG.COLUMNS del SEA.
  COL_OT: {
    FECHA: 0, OT: 1, TIPO: 2, NOM: 3, CLIENTE: 4, SUCURSAL: 5, RFC: 6,
    PERSONAL: 7, FECHA_VISITA: 8, FECHA_ENTREGA: 9, FECHA_REAL: 10,
    ESTATUS_EXTERNO: 11, LINK_DRIVE: 12, OBSERVACIONES: 13
  },
  COL_INF: {
    TIMESTAMP: 0, NUM_INFORME: 1, TIPO_ORDEN: 2, OT: 3, NOM: 4, CLIENTE: 5,
    SOLICITANTE: 6, RFC: 7, TELEFONO: 8, DIRECCION: 9, FECHA_SERVICIO: 10,
    FECHA_ENTREGA: 11, ES_CAPACITACION: 12, ESTATUS: 13, LINK_DRIVE: 14,
    RESPONSABLE: 15, SUCURSAL: 16
  },
  COL_USR: { EMAIL: 0, NOMBRE: 1, ROL: 2, ACTIVO: 3 },
  COL_AUD: {
    TIMESTAMP: 0, USUARIO: 1, ACCION: 2, OT: 3, CAMPO: 4,
    VALOR_ANTERIOR: 5, VALOR_NUEVO: 6
  }
};

var TRAZ_ESTATUS_TERMINALES = ['ENTREGADO', 'FINALIZADO', 'CANCELADO'];

// =========================================================================
// ENDPOINTS
// =========================================================================
function doGet(e) {
  if (!e || !e.parameter || !e.parameter.action) {
    return ContentService.createTextOutput('TRAZ — Trazabilidad Ejecutiva Ambiental (solo lectura) activa');
  }
  return trazRouter_(e.parameter, e.parameter.id_token || null);
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (_) { body = {}; }
  return trazRouter_(body, body.id_token || null);
}

/**
 * Router único. Verifica sesión (Google Auth + whitelist) y despacha.
 * Todas las acciones son de SOLO LECTURA.
 */
function trazRouter_(params, idToken) {
  var action = params.action;

  // ── Autenticación (misma política que los módulos internos SEA) ──────────
  if (action === 'verificarAcceso') {
    var u0 = trazVerificarIdToken_(idToken);
    var ok0 = u0 && trazVerificarUsuario_(u0.email);
    return trazOut_({ success: !!ok0, email: u0 ? u0.email : '' });
  }

  var usuario = trazVerificarIdToken_(idToken);
  if (!usuario) {
    return trazOut_({ success: false, error: 'AUTH_REQUIRED', message: 'Inicia sesión con Google.' });
  }
  if (!trazVerificarUsuario_(usuario.email)) {
    return trazOut_({ success: false, error: 'AUTH_REQUIRED', message: 'Tu cuenta (' + usuario.email + ') no tiene acceso a Trazabilidad.' });
  }

  switch (action) {
    case 'trazResumen': return trazOut_(trazResumen_());
    case 'trazDetalle': return trazOut_(trazDetalle_(params.ot));
    default:            return trazOut_({ success: false, error: 'Acción no reconocida.' });
  }
}

function trazOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// =========================================================================
// AUTENTICACIÓN — idéntica en enfoque a SEA (solo verifica, no escribe)
// =========================================================================
function trazVerificarIdToken_(idToken) {
  if (!idToken || typeof idToken !== 'string' || idToken.length < 100) return null;
  var cacheKey = 'traz_idtok_' + idToken.slice(-32);
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch (_) {} }
  try {
    var url = 'https://www.googleapis.com/oauth2/v3/tokeninfo?id_token=' + encodeURIComponent(idToken);
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    var data = JSON.parse(resp.getContentText());
    if (data.error_description) return null;
    var CLIENT_ID = PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID');
    if (CLIENT_ID && data.aud !== CLIENT_ID) return null;
    var userInfo = { email: data.email || '', name: data.name || '', sub: data.sub || '' };
    var ttl = Math.min(600, Math.max(5, Number(data.exp) - Math.floor(Date.now() / 1000) - 60));
    cache.put(cacheKey, JSON.stringify(userInfo), ttl);
    return userInfo;
  } catch (e) { return null; }
}

/**
 * Verifica que el email esté ACTIVO en USUARIOS_AUTORIZADOS y tenga acceso al
 * módulo SEAINF. SOLO LECTURA de la hoja.
 */
function trazVerificarUsuario_(email) {
  if (!email) return false;
  var emailLower = String(email).toLowerCase().trim();
  try {
    var ss = SpreadsheetApp.openById(TRAZ_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TRAZ_CONFIG.SHEET_USUARIOS);
    if (!sheet) return false;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return false;
    var headers = data[0].map(function (h) { return String(h).toUpperCase().replace(/[^A-Z0-9]/g, ''); });
    var moduloCol = headers.indexOf(TRAZ_CONFIG.MODULO_ACCESO.toUpperCase());
    var CU = TRAZ_CONFIG.COL_USR;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][CU.EMAIL]).toLowerCase().trim() !== emailLower) continue;
      if (data[i][CU.ACTIVO] !== true) return false;
      if (moduloCol >= 0) return data[i][moduloCol] === true;
      return true;
    }
    return false;
  } catch (e) { return false; }
}

// =========================================================================
// LECTURA DE DATOS (solo getDisplayValues — nunca escribe)
// =========================================================================
function trazNormOt_(ot) {
  return String(ot == null ? '' : ot).trim().toUpperCase();
}

function trazLeerHoja_(nombre) {
  var sheet = SpreadsheetApp.openById(TRAZ_CONFIG.SPREADSHEET_ID).getSheetByName(nombre);
  if (!sheet) return [];
  return sheet.getDataRange().getDisplayValues();
}

// =========================================================================
// RESUMEN — lista de OTs con su(s) informe(s) y banderas de trazabilidad
// =========================================================================
function trazResumen_() {
  var CO = TRAZ_CONFIG.COL_OT, CI = TRAZ_CONFIG.COL_INF;
  var ots = trazLeerHoja_(TRAZ_CONFIG.SHEET_OT).slice(1);
  var infs = trazLeerHoja_(TRAZ_CONFIG.SHEET_INFORMES).slice(1);

  // Índice de informes por folio de OT
  var infPorOt = {};
  infs.forEach(function (r) {
    var k = trazNormOt_(r[CI.OT]);
    if (!k) return;
    if (!infPorOt[k]) infPorOt[k] = [];
    infPorOt[k].push(r);
  });

  var lista = ots.filter(function (r) { return String(r[CO.OT] || '').trim() !== ''; })
    .map(function (r) {
      var folio = String(r[CO.OT]).trim();
      var informesDeOT = infPorOt[trazNormOt_(folio)] || [];
      var estatus = String(r[CO.ESTATUS_EXTERNO] || '').toUpperCase();
      return {
        ot:            folio,
        tipo:          r[CO.TIPO],
        nom:           r[CO.NOM],
        cliente:       r[CO.CLIENTE],
        sucursal:      r[CO.SUCURSAL],
        personal:      r[CO.PERSONAL],
        fecha_visita:  r[CO.FECHA_VISITA],
        estatus_ot:    r[CO.ESTATUS_EXTERNO],
        tiene_carpeta: String(r[CO.LINK_DRIVE] || '').indexOf('http') === 0,
        num_informes:  informesDeOT.length,
        folios_informe: informesDeOT.map(function (x) { return x[CI.NUM_INFORME]; }).filter(Boolean),
        entregado:     TRAZ_ESTATUS_TERMINALES.indexOf(estatus) !== -1
      };
    });

  // Informes huérfanos: referencian una OT que no existe en ORDENES_TRABAJO
  var foliosOT = {};
  ots.forEach(function (r) { foliosOT[trazNormOt_(r[CO.OT])] = true; });
  var huerfanos = infs
    .filter(function (r) { var k = trazNormOt_(r[CI.OT]); return k && !foliosOT[k]; })
    .map(function (r) { return { num_informe: r[CI.NUM_INFORME], ot: r[CI.OT] }; });

  return { success: true, data: lista, huerfanos: huerfanos };
}

// =========================================================================
// DETALLE — trazabilidad completa de UNA OT y su(s) informe(s)
// =========================================================================
function trazDetalle_(otFolio) {
  var CO = TRAZ_CONFIG.COL_OT, CI = TRAZ_CONFIG.COL_INF, CA = TRAZ_CONFIG.COL_AUD;
  var norm = trazNormOt_(otFolio);
  if (!norm) return { success: false, error: 'Folio de OT requerido.' };

  // ── OT (última coincidencia, de abajo hacia arriba como en SEA) ──────────
  var ots = trazLeerHoja_(TRAZ_CONFIG.SHEET_OT);
  var otRow = null;
  for (var i = ots.length - 1; i >= 1; i--) {
    if (trazNormOt_(ots[i][CO.OT]) === norm) { otRow = ots[i]; break; }
  }
  if (!otRow) return { success: false, error: 'OT no encontrada: ' + otFolio };

  var ot = {
    folio:        String(otRow[CO.OT]).trim(),
    tipo:         otRow[CO.TIPO],
    nom:          otRow[CO.NOM],
    cliente:      otRow[CO.CLIENTE],
    sucursal:     otRow[CO.SUCURSAL],
    rfc:          otRow[CO.RFC],
    personal:     otRow[CO.PERSONAL],
    fecha_alta:   otRow[CO.FECHA],
    fecha_visita: otRow[CO.FECHA_VISITA],
    fecha_entrega_limite: otRow[CO.FECHA_ENTREGA],
    fecha_real_entrega:   otRow[CO.FECHA_REAL],
    estatus:      otRow[CO.ESTATUS_EXTERNO],
    link_drive:   otRow[CO.LINK_DRIVE],
    observaciones: otRow[CO.OBSERVACIONES]
  };

  // ── Informe(s) de esta OT ────────────────────────────────────────────────
  var infs = trazLeerHoja_(TRAZ_CONFIG.SHEET_INFORMES).slice(1);
  var informes = infs
    .filter(function (r) { return trazNormOt_(r[CI.OT]) === norm; })
    .map(function (r) {
      return {
        num_informe:    r[CI.NUM_INFORME],
        tipo_orden:     r[CI.TIPO_ORDEN],
        nom:            r[CI.NOM],
        solicitante:    r[CI.SOLICITANTE],
        fecha_servicio: r[CI.FECHA_SERVICIO],
        fecha_entrega:  r[CI.FECHA_ENTREGA],
        es_capacitacion: r[CI.ES_CAPACITACION],
        estatus:        r[CI.ESTATUS],
        link_drive:     r[CI.LINK_DRIVE],   // ← EXPEDIENTE
        responsable:    r[CI.RESPONSABLE],
        timestamp:      r[CI.TIMESTAMP]
      };
    });

  // ── Bitácora real (AUDITORIA) de esta OT — sin inferir nada ──────────────
  var auds = trazLeerHoja_(TRAZ_CONFIG.SHEET_AUDITORIA).slice(1);
  var bitacora = auds
    .filter(function (r) { return trazNormOt_(r[CA.OT]) === norm; })
    .map(function (r) {
      return {
        timestamp:       r[CA.TIMESTAMP],
        usuario:         r[CA.USUARIO],
        accion:          r[CA.ACCION],
        campo:           r[CA.CAMPO],
        valor_anterior:  r[CA.VALOR_ANTERIOR],
        valor_nuevo:     r[CA.VALOR_NUEVO]
      };
    });

  // Expediente preferente: el del informe; si no, la carpeta de la OT.
  var expediente = '';
  for (var j = 0; j < informes.length; j++) {
    if (String(informes[j].link_drive || '').indexOf('http') === 0) { expediente = informes[j].link_drive; break; }
  }
  if (!expediente && String(ot.link_drive || '').indexOf('http') === 0) expediente = ot.link_drive;

  return {
    success: true,
    ot: ot,
    informes: informes,
    expediente: expediente,
    bitacora: bitacora,
    linea_tiempo: trazLineaTiempo_(ot, informes),
    advertencias: trazAdvertencias_(ot, informes)
  };
}

// =========================================================================
// LÍNEA DE TIEMPO — solo con datos que existen (✓ / ⚠)
// No se sintetizan hitos de "revisión"/"emisión": esos solo se ven en la
// bitácora real (AUDITORIA) si fueron registrados.
// =========================================================================
function trazLineaTiempo_(ot, informes) {
  function nodo(clave, titulo, fecha, ok, extra) {
    return { clave: clave, titulo: titulo, fecha: fecha || '', estado: ok ? 'ok' : 'pendiente', extra: extra || '' };
  }
  var tieneExp = false;
  var foliosInf = [];
  informes.forEach(function (inf) {
    if (String(inf.link_drive || '').indexOf('http') === 0) tieneExp = true;
    if (inf.num_informe) foliosInf.push(inf.num_informe);
  });

  var pasos = [];
  pasos.push(nodo('ot', 'Orden de Trabajo', ot.fecha_alta, !!ot.folio, ot.folio));
  pasos.push(nodo('ejecucion', 'Servicio ejecutado', ot.fecha_visita, !!String(ot.fecha_visita || '').trim()));
  pasos.push(nodo('expediente', 'Expediente', '', tieneExp, tieneExp ? 'Carpeta de trabajo disponible' : 'Sin carpeta relacionada'));
  pasos.push(nodo('informe', 'Informe', '', foliosInf.length > 0, foliosInf.join(', ')));
  pasos.push(nodo('entrega', 'Entrega', ot.fecha_real_entrega, !!String(ot.fecha_real_entrega || '').trim()));
  return pasos;
}

// =========================================================================
// ADVERTENCIAS — solo validaciones viables con los datos existentes
// (No se validan memoria ni equipos: no existen como dato.)
// =========================================================================
function trazAdvertencias_(ot, informes) {
  var w = [];
  var estatus = String(ot.estatus || '').toUpperCase();

  if (String(ot.link_drive || '').indexOf('http') !== 0) {
    w.push('La OT no tiene carpeta de trabajo (expediente) registrada.');
  }
  if (informes.length === 0) {
    w.push('La OT no tiene ningún informe relacionado.');
  }
  informes.forEach(function (inf) {
    var etq = inf.num_informe ? ('Informe ' + inf.num_informe) : 'Un informe';
    if (String(inf.link_drive || '').indexOf('http') !== 0) {
      w.push(etq + ' no tiene expediente (carpeta Drive).');
    }
    if (!String(inf.fecha_servicio || '').trim()) {
      w.push(etq + ' no tiene fecha de ejecución del servicio.');
    }
  });
  if ((estatus === 'ENTREGADO' || estatus === 'FINALIZADO') && !String(ot.fecha_real_entrega || '').trim()) {
    w.push('La OT figura como ' + estatus + ' pero no tiene fecha real de entrega.');
  }
  return w;
}
