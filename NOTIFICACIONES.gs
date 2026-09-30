// =========================================================================
// NOTIFICACIONES AUTOMÁTICAS — Sistema SEA (Ejecutiva Ambiental)
// =========================================================================
// Convierte en correo lo que hoy sólo se ve si alguien abre SEADB.
//
//   1. Digest operativo    → lunes y jueves 07:30 → operaciones, dirección, aclientes
//   2. Resumen de dirección→ viernes 17:00        → dirección general
//   3. Renovaciones        → lunes 08:00          → ventas
//
// DEPENDENCIAS
//   Vive en el mismo proyecto de Apps Script que BACKEND_FIXES.gs y reutiliza
//   de ahí: CONFIG, CO, CI, EMAIL_COLORS_, escHtml_, encabezadoEmail_,
//   envolturaEmail_, pieEmail_, etiquetaEmail_, botonEmail_.
//   No modifica doGet ni doPost: no hace falta publicar una versión nueva del
//   Web App para instalar este módulo.
//
// INSTALACIÓN (una sola vez, desde el editor de Apps Script)
//   1. Pegar este archivo como NOTIFICACIONES.gs y guardar.
//   2. Ejecutar  notif_estado()                  → verifica configuración y zona horaria.
//   3. Ejecutar  notif_previsualizar()           → escribe los 3 correos en el Log, sin enviar.
//   4. Poner NOTIF_DRY_RUN en propiedades del script con tu correo y
//      ejecutar  notif_probarEnvio()             → te llegan los 3 marcados [PRUEBA].
//   5. Borrar NOTIF_DRY_RUN y ejecutar  configurarTriggersNotificaciones().
//
// INTERRUPTORES (Configuración del proyecto → Propiedades del script)
//   NOTIF_ACTIVO   = 'false'            → apaga todos los envíos sin borrar triggers.
//   NOTIF_DRY_RUN  = 'correo@dominio'   → redirige TODO a ese buzón, con asunto [PRUEBA].
//
// CUOTA
//   Cuenta gratuita: 100 destinatarios/día. Este módulo consume 8 por semana
//   (6 del digest + 1 dirección + 1 renovaciones), compartidos con los correos
//   transaccionales que ya envía BACKEND_FIXES.gs.
// =========================================================================

const NOTIF_CONFIG = {

  // ── Destinatarios por tipo de correo ────────────────────────────────────
  DESTINATARIOS: {
    OPERATIVO: [
      'operaciones@ejecutivambiental.com',
      'direccion.general@ejecutivambiental.com',
      'aclientes@ejecutivambiental.com'
    ],
    DIRECCION:    ['direccion.general@ejecutivambiental.com'],
    RENOVACIONES: ['ventas@ejecutivambiental.com']
  },

  // Enlace al tablero que llevan los botones de los correos.
  // Ajustar si cambia el dominio de publicación.
  URL_SEADB: 'https://yoeduwin.github.io/SEA/SEADB.html',

  // ── Umbrales de negocio ─────────────────────────────────────────────────
  DIAS_LIMITE:        3,   // "en límite": entrega en menos de N días
  SLA_DIGITAL_DIAS:   20,  // fallback: visita + N días cuando no hay fecha límite capturada
  DIAS_ANTIGUO:       60,  // informe abierto "atascado" a partir de N días
  VENTANA_RENOVACION: 30,  // renovaciones que entran al correo: vencidas y a menos de N días
  MAX_FILAS_TABLA:    15,  // corte por bloque; el resto se resume en "+N más"

  // ── Ciclo de renovación por norma, en años ──────────────────────────────
  // La llave es el número de norma ya normalizado (ver notif_clavesServicio_).
  CICLOS_RENOVACION: {
    '022':  1,   // NOM-022-STPS
    '081':  1,   // NOM-081-SEMARNAT
    '025':  2,   // NOM-025-STPS
    '024':  2,   // NOM-024-STPS
    '015':  2,   // NOM-015-STPS
    'PIPC': 1    // Programa Interno de Protección Civil
  },

  // Nombre legible de cada servicio rastreado, para los correos.
  NOMBRES_SERVICIO: {
    '022':  'NOM-022-STPS',
    '081':  'NOM-081-SEMARNAT',
    '025':  'NOM-025-STPS',
    '024':  'NOM-024-STPS',
    '015':  'NOM-015-STPS',
    'PIPC': 'Protección Civil / PIPC'
  },

  // ── Infraestructura ─────────────────────────────────────────────────────
  SHEET_LOG:        'NOTIFICACIONES_LOG',
  PROP_ACTIVO:      'NOTIF_ACTIVO',
  PROP_DRY_RUN:     'NOTIF_DRY_RUN',
  REINTENTOS:       3,
  REINTENTO_MS:     2000,
  LOG_FILAS_REVISA: 500   // cuántas filas del log se revisan al buscar duplicados
};

// Estatus de OT que ya no requieren seguimiento operativo.
const NOTIF_ESTATUS_CERRADOS_ = ['ENTREGADO', 'FINALIZADO', 'CANCELADO'];


// =========================================================================
// UTILIDADES BÁSICAS
// =========================================================================

/** Lee una propiedad del script, con valor por omisión. */
function notif_prop_(clave, porOmision) {
  try {
    const v = PropertiesService.getScriptProperties().getProperty(clave);
    return (v === null || v === undefined || v === '') ? porOmision : v;
  } catch (e) {
    return porOmision;
  }
}

/** ¿El módulo está habilitado? Se apaga poniendo NOTIF_ACTIVO = 'false'. */
function notif_activo_() {
  return String(notif_prop_(NOTIF_CONFIG.PROP_ACTIVO, 'true')).toLowerCase() !== 'false';
}

/** Devuelve el buzón de pruebas si NOTIF_DRY_RUN está puesto, o null. */
function notif_dryRun_() {
  const v = String(notif_prop_(NOTIF_CONFIG.PROP_DRY_RUN, '')).trim();
  return v && v.indexOf('@') > 0 ? v : null;
}

/** Fecha de hoy a medianoche, para comparar días sin arrastrar la hora. */
function notif_hoy_() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * Convierte a Date lo que venga de la hoja: objeto Date, número de serie de
 * Sheets, o texto en dd/mm/yyyy o yyyy-mm-dd. Devuelve null si no se puede.
 */
function notif_parseFecha_(valor) {
  if (valor === null || valor === undefined || valor === '') return null;

  if (Object.prototype.toString.call(valor) === '[object Date]') {
    if (isNaN(valor.getTime())) return null;
    const d = new Date(valor.getTime());
    d.setHours(0, 0, 0, 0);
    return d;
  }

  // Número de serie de Google Sheets (época 30/12/1899).
  if (typeof valor === 'number' && isFinite(valor)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(valor) * 86400000);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  const txt = String(valor).trim().split(' ')[0];
  if (!txt) return null;

  let m = txt.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);   // dd/mm/yyyy
  if (m) {
    let anio = parseInt(m[3], 10);
    if (anio < 100) anio += 2000;
    const d = new Date(anio, parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    return isNaN(d.getTime()) ? null : d;
  }

  m = txt.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);            // yyyy-mm-dd
  if (m) {
    const d = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
    return isNaN(d.getTime()) ? null : d;
  }

  // Último recurso: si el número viene como texto, tratarlo como serie.
  if (/^\d+(\.\d+)?$/.test(txt)) return notif_parseFecha_(parseFloat(txt));

  return null;
}

/** dd/mm/yyyy para mostrar en los correos. */
function notif_fmtFecha_(d) {
  if (!d) return '—';
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd/MM/yyyy');
}

/** Días completos entre dos fechas (b - a). */
function notif_dias_(a, b) {
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

/** MAYÚSCULAS sin acentos y con espacios colapsados, para comparar. */
function notif_normalizar_(s) {
  return String(s === null || s === undefined ? '' : s)
    .toUpperCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Recorta un texto largo para que no rompa el maquetado del correo. */
function notif_corta_(s, n) {
  const t = String(s === null || s === undefined ? '' : s).trim();
  return t.length > n ? t.substring(0, n - 1) + '…' : t;
}

/** Clave ISO de semana, p. ej. 2026-W38. Se usa para no duplicar envíos. */
function notif_claveSemana_(d) {
  const t = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  t.setDate(t.getDate() + 4 - (t.getDay() || 7));            // jueves de esa semana
  const inicio = new Date(t.getFullYear(), 0, 1);
  const semana = Math.ceil((((t - inicio) / 86400000) + 1) / 7);
  return t.getFullYear() + '-W' + String(semana).padStart(2, '0');
}

/** Lunes de la semana de la fecha dada. */
function notif_lunesDe_(d) {
  const l = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  l.setDate(l.getDate() - ((l.getDay() + 6) % 7));
  return l;
}

/**
 * Extrae las claves de servicio rastreable que aparecen en la columna NOM.
 *
 * La columna es texto libre y en la práctica trae formatos muy distintos:
 * 'NOM-025-STPS', 'NOM-025-STPS-2008', 'NOM-025', '025', '25', 'PIPC',
 * 'PROTECCIÓN CIVIL' e incluso varias normas en una sola celda
 * ('NOM-015-STPS, NOM-022-STPS'). Devuelve todas las que reconozca.
 */
function notif_clavesServicio_(nomRaw) {
  const n = notif_normalizar_(nomRaw);
  if (!n) return [];

  const claves = {};

  if (/\bPIPC\b/.test(n) ||
      n.indexOf('PROGRAMA INTERNO') >= 0 ||
      n.indexOf('PROTECCION CIVIL') >= 0) {
    claves['PIPC'] = true;
  }

  // Forma explícita: NOM-025, NOM 25, NOM-036-1-STPS…
  // El (?!\d) evita que el año de 'NOM-022-STPS-2015' se lea como norma.
  let m;
  const reNom = /NOM[\s\-]*0*(\d{1,3})(?!\d)/g;
  while ((m = reNom.exec(n)) !== null) {
    claves[String(parseInt(m[1], 10)).padStart(3, '0')] = true;
  }

  // Forma abreviada: la celda trae sólo el número ('025', '81', '22').
  n.split(/[,;\/]+/).forEach(function (token) {
    const t = token.trim();
    if (/^0*\d{1,3}$/.test(t)) {
      claves[String(parseInt(t, 10)).padStart(3, '0')] = true;
    }
  });

  // Sólo interesan las normas con ciclo de renovación conocido.
  return Object.keys(claves).filter(function (k) {
    return Object.prototype.hasOwnProperty.call(NOTIF_CONFIG.CICLOS_RENOVACION, k);
  });
}

/**
 * Acumula nombres de persona bajo una clave normalizada, conservando la
 * escritura más legible. La hoja guarda 'MARTIN LUNA' y 'Martín Luna' como
 * valores distintos; sin esto cada persona aparecería dos veces.
 */
function notif_acumulaPersona_(mapa, nombreRaw) {
  const nombre = String(nombreRaw || '').trim();
  if (!nombre) return;
  const clave = notif_normalizar_(nombre);
  if (!mapa[clave]) {
    mapa[clave] = { nombre: nombre, total: 0 };
  } else if (mapa[clave].nombre === mapa[clave].nombre.toUpperCase() &&
             nombre !== nombre.toUpperCase()) {
    // Preferir 'Martín Luna' sobre 'MARTIN LUNA' para mostrar.
    mapa[clave].nombre = nombre;
  }
  mapa[clave].total++;
}

/** Lee una hoja completa; devuelve [] si no existe. */
function notif_leerHoja_(nombre) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    Logger.log('notif: no existe la hoja ' + nombre);
    return [];
  }
  const valores = hoja.getDataRange().getValues();
  return valores.length > 1 ? valores.slice(1) : [];
}


// =========================================================================
// BITÁCORA DE ENVÍOS — evita que un doble disparo del trigger duplique correo
// =========================================================================

function notif_hojaLog_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let hoja = ss.getSheetByName(NOTIF_CONFIG.SHEET_LOG);
  if (!hoja) {
    hoja = ss.insertSheet(NOTIF_CONFIG.SHEET_LOG);
    const encabezados = ['Timestamp', 'Tipo', 'Clave', 'Destinatarios', 'Asunto', 'Resultado', 'Detalle'];
    hoja.appendRow(encabezados);
    hoja.getRange(1, 1, 1, encabezados.length)
        .setFontWeight('bold').setBackground('#123A28').setFontColor('#FFFFFF');
    hoja.setFrozenRows(1);
    hoja.setColumnWidth(1, 160);
    hoja.setColumnWidth(5, 320);
    Logger.log('notif: hoja ' + NOTIF_CONFIG.SHEET_LOG + ' creada.');
  }
  return hoja;
}

/** Últimas filas del log, para buscar duplicados sin leer años de historia. */
function notif_ultimasFilasLog_() {
  const hoja = notif_hojaLog_();
  const total = hoja.getLastRow();
  if (total < 2) return [];
  const desde = Math.max(2, total - NOTIF_CONFIG.LOG_FILAS_REVISA + 1);
  return hoja.getRange(desde, 1, total - desde + 1, 7).getValues();
}

/** ¿Ya se envió este tipo de correo con esta clave? */
function notif_yaEnviado_(tipo, clave) {
  const filas = notif_ultimasFilasLog_();
  for (let i = filas.length - 1; i >= 0; i--) {
    if (String(filas[i][1]).trim() === tipo &&
        String(filas[i][2]).trim() === clave &&
        String(filas[i][5]).trim().toUpperCase() === 'ENVIADO') {
      return true;
    }
  }
  return false;
}

/** Detalle guardado en el último envío de este tipo, como objeto. */
function notif_detalleAnterior_(tipo) {
  const filas = notif_ultimasFilasLog_();
  for (let i = filas.length - 1; i >= 0; i--) {
    if (String(filas[i][1]).trim() === tipo &&
        String(filas[i][5]).trim().toUpperCase() === 'ENVIADO') {
      try {
        return JSON.parse(String(filas[i][6] || '{}'));
      } catch (e) {
        return {};
      }
    }
  }
  return {};
}

function notif_registrarEnvio_(tipo, clave, destinatarios, asunto, resultado, detalle) {
  try {
    notif_hojaLog_().appendRow([
      new Date(), tipo, clave, destinatarios.join(', '),
      asunto, resultado, detalle ? JSON.stringify(detalle) : ''
    ]);
  } catch (e) {
    Logger.log('notif_registrarEnvio_ error: ' + e.message);
  }
}


// =========================================================================
// ENVÍO
// =========================================================================

/**
 * Envía un correo aplicando todas las salvaguardas: interruptor general,
 * modo prueba, antiduplicado, reintentos y bitácora.
 *
 * @param {Object} op  { tipo, clave, destinatarios[], asunto, html, texto, detalle }
 * @return {Object} { enviado, razon }
 */
function notif_enviar_(op) {
  if (!notif_activo_()) {
    Logger.log('notif: desactivado por NOTIF_ACTIVO; no se envía ' + op.tipo);
    return { enviado: false, razon: 'desactivado' };
  }

  const prueba = notif_dryRun_();
  const destinos = prueba ? [prueba] : op.destinatarios.slice();
  const asunto   = prueba ? '[PRUEBA] ' + op.asunto : op.asunto;

  if (!destinos.length) {
    Logger.log('notif: sin destinatarios para ' + op.tipo);
    return { enviado: false, razon: 'sin_destinatarios' };
  }

  // En modo prueba se permite reenviar cuantas veces haga falta.
  const lock = LockService.getScriptLock();
  let conLock = false;
  try {
    conLock = lock.tryLock(15000);

    if (!prueba && notif_yaEnviado_(op.tipo, op.clave)) {
      Logger.log('notif: ' + op.tipo + ' con clave ' + op.clave + ' ya se había enviado.');
      return { enviado: false, razon: 'duplicado' };
    }

    let ultimoError = null;
    for (let intento = 1; intento <= NOTIF_CONFIG.REINTENTOS; intento++) {
      try {
        GmailApp.sendEmail(destinos.join(','), asunto, op.texto, {
          htmlBody: op.html,
          name: CONFIG.COMPANY_NAME
        });
        notif_registrarEnvio_(
          prueba ? op.tipo + '_PRUEBA' : op.tipo,
          op.clave, destinos, asunto, 'ENVIADO', op.detalle
        );
        Logger.log('notif: enviado ' + op.tipo + ' → ' + destinos.join(', '));
        return { enviado: true };
      } catch (err) {
        ultimoError = err;
        if (intento < NOTIF_CONFIG.REINTENTOS) Utilities.sleep(NOTIF_CONFIG.REINTENTO_MS);
      }
    }

    notif_registrarEnvio_(
      prueba ? op.tipo + '_PRUEBA' : op.tipo,
      op.clave, destinos, asunto, 'ERROR',
      { error: String(ultimoError) }
    );
    Logger.log('notif: FALLÓ ' + op.tipo + ' — ' + ultimoError);
    return { enviado: false, razon: 'error', error: String(ultimoError) };

  } finally {
    if (conLock) {
      try { lock.releaseLock(); } catch (e) { /* nada que hacer */ }
    }
  }
}


// =========================================================================
// LECTURA DE DATOS
// =========================================================================

/**
 * Clasifica las órdenes de trabajo que siguen abiertas.
 * La fecha límite es la capturada; si falta, se calcula como visita + SLA,
 * igual que hace SEADB.
 */
function notif_ordenesAbiertas_() {
  const hoy = notif_hoy_();
  const filas = notif_leerHoja_(CONFIG.SHEET_OT);

  const res = { vencidas: [], limite: [], pausa: [], holgura: [], sinFecha: [], total: 0 };

  filas.forEach(function (fila) {
    const folio = String(fila[CO.OT] || '').trim();
    if (!folio) return;

    const estatus = notif_normalizar_(fila[CO.ESTATUS_EXTERNO]);
    if (NOTIF_ESTATUS_CERRADOS_.indexOf(estatus) >= 0) return;

    res.total++;

    const ot = {
      folio:     folio,
      nom:       String(fila[CO.NOM] || '').trim(),
      cliente:   String(fila[CO.CLIENTE] || '').trim(),
      sucursal:  String(fila[CO.SUCURSAL] || '').trim(),
      personal:  String(fila[CO.PERSONAL] || '').trim(),
      estatus:   estatus
    };

    if (estatus === 'EN PAUSA') {
      const fp = notif_parseFecha_(fila[CO.FECHA_PAUSA]);
      ot.fechaPausa = fp;
      ot.diasPausa  = fp ? notif_dias_(fp, hoy) : null;
      ot.motivo     = String(fila[CO.MOTIVO_PAUSA] || '').trim();
      res.pausa.push(ot);
      return;
    }

    let limite = notif_parseFecha_(fila[CO.FECHA_ENTREGA]);
    if (!limite) {
      const visita = notif_parseFecha_(fila[CO.FECHA_VISITA]);
      if (visita) {
        limite = new Date(visita.getTime());
        limite.setDate(limite.getDate() + NOTIF_CONFIG.SLA_DIGITAL_DIAS);
        ot.limiteEstimado = true;
      }
    }

    if (!limite) { res.sinFecha.push(ot); return; }

    ot.limite = limite;
    ot.dias   = notif_dias_(hoy, limite);      // negativo = vencida

    if (ot.dias < 0)                              res.vencidas.push(ot);
    else if (ot.dias < NOTIF_CONFIG.DIAS_LIMITE)  res.limite.push(ot);
    else                                          res.holgura.push(ot);
  });

  res.vencidas.sort(function (a, b) { return a.dias - b.dias; });          // más atrasada primero
  res.limite.sort(function (a, b) { return a.dias - b.dias; });
  res.pausa.sort(function (a, b) { return (b.diasPausa || 0) - (a.diasPausa || 0); });

  return res;
}

/** Cartera de informes que aún no están finalizados, con su antigüedad. */
function notif_carteraInformes_() {
  const hoy = notif_hoy_();
  const filas = notif_leerHoja_(CONFIG.SHEET_INFORMES);

  const res = {
    abiertos: [], porEtapa: {}, porResponsable: {},
    antiguedad: { d0_30: 0, d31_60: 0, d61_90: 0, d90: 0 },
    atascados: 0, totalRegistros: 0
  };

  filas.forEach(function (fila) {
    const num = String(fila[CI.NUM_INFORME] || '').trim();
    const cliente = String(fila[CI.CLIENTE] || '').trim();
    if (!num && !cliente) return;

    res.totalRegistros++;

    const etapa = notif_normalizar_(fila[CI.ESTATUS]);
    if (etapa === 'FINALIZADO') return;

    const alta = notif_parseFecha_(fila[CI.TIMESTAMP]);
    const dias = alta ? notif_dias_(alta, hoy) : null;

    const inf = {
      num:         num,
      cliente:     cliente,
      sucursal:    String(fila[CI.SUCURSAL] || '').trim(),
      nom:         String(fila[CI.NOM] || '').trim(),
      etapa:       etapa || 'SIN ETAPA',
      responsable: String(fila[CI.RESPONSABLE] || '').trim(),
      alta:        alta,
      dias:        dias
    };

    res.abiertos.push(inf);
    res.porEtapa[inf.etapa] = (res.porEtapa[inf.etapa] || 0) + 1;
    notif_acumulaPersona_(res.porResponsable, inf.responsable);

    if (dias === null)    return;
    if (dias <= 30)       res.antiguedad.d0_30++;
    else if (dias <= 60)  res.antiguedad.d31_60++;
    else if (dias <= 90)  res.antiguedad.d61_90++;
    else                  res.antiguedad.d90++;

    if (dias > NOTIF_CONFIG.DIAS_ANTIGUO) res.atascados++;
  });

  // Más antiguos primero; los que no tienen fecha van al final.
  res.abiertos.sort(function (a, b) {
    if (a.dias === null) return 1;
    if (b.dias === null) return -1;
    return b.dias - a.dias;
  });

  return res;
}

/**
 * Renovaciones calculadas sobre INFORMES: por cada cliente / sucursal /
 * servicio rastreado se toma el informe más reciente y se le suma el ciclo.
 */
function notif_renovaciones_() {
  const hoy = notif_hoy_();
  const filas = notif_leerHoja_(CONFIG.SHEET_INFORMES);
  const ultimos = {};

  filas.forEach(function (fila) {
    const fecha = notif_parseFecha_(fila[CI.TIMESTAMP]);
    if (!fecha) return;

    const claves = notif_clavesServicio_(fila[CI.NOM]);
    if (!claves.length) return;

    const cliente  = String(fila[CI.CLIENTE] || '').trim();
    const rfc      = notif_normalizar_(fila[CI.RFC]);
    const sucursal = String(fila[CI.SUCURSAL] || '').trim();
    const idCliente = rfc || notif_normalizar_(cliente);
    if (!idCliente) return;

    claves.forEach(function (clave) {
      const id = idCliente + '|' + notif_normalizar_(sucursal) + '|' + clave;
      if (!ultimos[id] || fecha.getTime() > ultimos[id].fecha.getTime()) {
        ultimos[id] = {
          fecha: fecha, clave: clave, cliente: cliente,
          sucursal: sucursal, rfc: rfc
        };
      }
    });
  });

  const lista = Object.keys(ultimos).map(function (id) {
    const it = ultimos[id];
    const ciclo = NOTIF_CONFIG.CICLOS_RENOVACION[it.clave];
    const prox = new Date(it.fecha.getTime());
    prox.setFullYear(prox.getFullYear() + ciclo);
    return {
      cliente:  it.cliente,
      sucursal: it.sucursal,
      rfc:      it.rfc,
      servicio: NOTIF_CONFIG.NOMBRES_SERVICIO[it.clave] || it.clave,
      ciclo:    ciclo,
      ultimo:   it.fecha,
      proxima:  prox,
      dias:     notif_dias_(hoy, prox)
    };
  });

  lista.sort(function (a, b) { return a.dias - b.dias; });
  return lista;
}


// =========================================================================
// COMPOSICIÓN HTML — se apoya en los ayudantes de BACKEND_FIXES.gs
// =========================================================================

function notif_seccion_(contenido, paddingTop) {
  const pt = paddingTop === undefined ? 28 : paddingTop;
  return '<tr><td style="padding:' + pt + 'px 30px 0 30px;">' + contenido + '</td></tr>';
}

function notif_separador_() {
  return '<div style="height:1px; line-height:1px; font-size:0; background-color:' +
         EMAIL_COLORS_.borde + '; margin:8px 0 24px 0;">&nbsp;</div>';
}

/** Chip de color según qué tan crítico es el dato. */
function notif_chip_(texto, tono) {
  const paleta = {
    rojo:   { bg: '#F7E2DE', fg: '#A8452F' },
    ambar:  { bg: '#FBF0D8', fg: '#8A5F0C' },
    verde:  { bg: '#E9F2EC', fg: '#2C6B47' },
    neutro: { bg: '#EEF2EC', fg: '#5A665A' }
  }[tono] || { bg: '#EEF2EC', fg: '#5A665A' };
  return '<span style="display:inline-block; background-color:' + paleta.bg +
         '; color:' + paleta.fg + '; font-size:11px; font-weight:700; padding:5px 9px;' +
         ' border-radius:5px; white-space:nowrap;">' + escHtml_(texto) + '</span>';
}

/**
 * Tabla con encabezado. `filas` es un arreglo de arreglos de HTML ya escapado
 * por quien la construye; `alineaDerecha` marca las columnas alineadas a la
 * derecha por índice.
 */
function notif_tabla_(encabezados, filas, alineaDerecha) {
  const der = alineaDerecha || [];
  const th = encabezados.map(function (h) {
    return '<td style="padding:9px 12px; font-size:10px; letter-spacing:.6px;' +
           ' text-transform:uppercase; color:' + EMAIL_COLORS_.etiqueta +
           '; font-weight:700;">' + escHtml_(h) + '</td>';
  }).join('');

  const tr = filas.map(function (fila, i) {
    const borde = i === 0 ? '' : 'border-top:1px solid ' + EMAIL_COLORS_.bordeSuave + ';';
    const tds = fila.map(function (celda, j) {
      const align = der.indexOf(j) >= 0 ? ' align="right"' : '';
      return '<td' + align + ' style="padding:11px 12px; ' + borde +
             ' font-size:13px; color:' + EMAIL_COLORS_.texto + '; vertical-align:top;">' +
             celda + '</td>';
    }).join('');
    return '<tr>' + tds + '</tr>';
  }).join('');

  return '<table width="100%" border="0" cellpadding="0" cellspacing="0" style="border:1px solid ' +
         EMAIL_COLORS_.bordeSuave + '; border-radius:8px;">' +
         '<tr style="background-color:' + EMAIL_COLORS_.panel + ';">' + th + '</tr>' +
         tr + '</table>';
}

/** Nombre del cliente con su folio debajo, para las tablas. */
function notif_celdaCliente_(titulo, subtitulo) {
  return '<strong style="color:' + EMAIL_COLORS_.textoFuerte + '; font-weight:600;">' +
         escHtml_(notif_corta_(titulo, 38)) + '</strong>' +
         (subtitulo ? '<br><span style="color:#8A948A; font-size:11px;">' +
                      escHtml_(notif_corta_(subtitulo, 52)) + '</span>' : '');
}

/** Barra horizontal para el bloque de antigüedad. */
function notif_barra_(etiqueta, valor, porcentaje, color, destacar) {
  const pct = Math.max(2, Math.min(100, Math.round(porcentaje)));
  const colEtiqueta = destacar ? EMAIL_COLORS_.ambarTexto : EMAIL_COLORS_.textoSuave;
  const colValor    = destacar ? '#A8452F' : EMAIL_COLORS_.textoFuerte;
  return '<tr>' +
    '<td width="26%" style="padding:6px 10px 6px 0; font-size:13px; color:' + colEtiqueta +
      ';' + (destacar ? ' font-weight:700;' : '') + '">' + escHtml_(etiqueta) + '</td>' +
    '<td width="56%" style="padding:6px 0;">' +
      '<table width="100%" border="0" cellpadding="0" cellspacing="0"><tr>' +
        '<td bgcolor="' + color + '" width="' + pct + '%" style="height:9px; font-size:0;' +
        ' line-height:0; border-radius:3px;">&nbsp;</td><td>&nbsp;</td>' +
      '</tr></table></td>' +
    '<td align="right" style="padding:6px 0; font-size:' + (destacar ? '15' : '14') +
      'px; color:' + colValor + '; font-weight:' + (destacar ? '800' : '700') + ';">' +
      valor + '</td></tr>';
}

/** Fila etiqueta / valor grande, para los bloques de cifras. */
function notif_filaCifra_(etiqueta, valor, nota, tonoNota) {
  const colores = { verde: '#2C6B47', ambar: EMAIL_COLORS_.ambarTexto, rojo: '#A8452F' };
  const notaHtml = nota
    ? '<span style="color:' + (colores[tonoNota] || EMAIL_COLORS_.textoSuave) +
      '; font-size:12px; font-weight:700;">&nbsp;' + escHtml_(nota) + '</span>'
    : '';
  return '<tr>' +
    '<td width="42%" style="padding:9px 14px 9px 0; border-bottom:1px solid ' +
      EMAIL_COLORS_.bordeSuave + '; font-size:14px; color:' + EMAIL_COLORS_.textoSuave +
      '; vertical-align:top;">' + escHtml_(etiqueta) + '</td>' +
    '<td style="padding:9px 0; border-bottom:1px solid ' + EMAIL_COLORS_.bordeSuave +
      '; font-size:14px; color:' + EMAIL_COLORS_.textoFuerte + ';">' +
      '<span style="font-size:19px; font-weight:700;">' + escHtml_(String(valor)) + '</span>' +
      notaHtml + '</td></tr>';
}

/** Línea "+N más" cuando una tabla se recorta. */
function notif_masFilas_(sobrantes, sustantivo) {
  if (sobrantes <= 0) return '';
  return '<p style="margin:11px 0 0 0; font-size:12px; color:#8A948A;">+ ' + sobrantes +
         ' ' + escHtml_(sustantivo) + ' más. Ver el detalle completo en SEADB.</p>';
}

function notif_botonera_(botones) {
  return '<table border="0" cellpadding="0" cellspacing="0"><tr>' + botones + '</tr></table>';
}


// =========================================================================
// CORREO 1 — DIGEST OPERATIVO (lunes y jueves)
// =========================================================================

function notif_enviarDigestOperativo() {
  try {
    const hoy = notif_hoy_();
    const d = notif_ordenesAbiertas_();
    const pendientes = d.vencidas.length + d.limite.length + d.pausa.length;

    if (pendientes === 0) {
      Logger.log('notif: digest operativo sin pendientes; no se envía.');
      return { enviado: false, razon: 'vacio' };
    }

    // Delta contra el envío anterior: cuáles OTs vencidas son nuevas.
    const previo = notif_detalleAnterior_('DIGEST_OPERATIVO');
    const antes = (previo && previo.vencidas) || [];
    const folios = d.vencidas.map(function (o) { return o.folio; });
    const nuevas = folios.filter(function (f) { return antes.indexOf(f) < 0; });

    let filas = '';

    // ── Vencidas ──────────────────────────────────────────────────────────
    if (d.vencidas.length) {
      const visibles = d.vencidas.slice(0, NOTIF_CONFIG.MAX_FILAS_TABLA);
      const tabla = notif_tabla_(
        ['Orden', 'Límite', 'Atraso'],
        visibles.map(function (o) {
          return [
            notif_celdaCliente_(o.cliente, o.folio + ' · ' + o.nom +
              (o.personal ? ' · ' + o.personal : '')),
            '<span style="font-size:12px; color:' + EMAIL_COLORS_.textoSuave + ';">' +
              notif_fmtFecha_(o.limite) + (o.limiteEstimado ? '<br><em>estimado</em>' : '') + '</span>',
            notif_chip_(Math.abs(o.dias) + (Math.abs(o.dias) === 1 ? ' día' : ' días'), 'rojo')
          ];
        }), [2]);

      const sub = antes.length
        ? (nuevas.length
            ? nuevas.length + (nuevas.length === 1 ? ' nueva' : ' nuevas') + ' desde el envío anterior.'
            : 'Sin órdenes nuevas desde el envío anterior.')
        : '';

      filas += notif_seccion_(
        etiquetaEmail_('Vencidas · ' + d.vencidas.length) +
        (sub ? '<p style="margin:-6px 0 12px 0; font-size:12px; color:#8A948A;">' +
               escHtml_(sub) + '</p>' : '') +
        tabla +
        notif_masFilas_(d.vencidas.length - visibles.length, 'vencidas'));
    }

    // ── En límite ─────────────────────────────────────────────────────────
    if (d.limite.length) {
      const tabla = notif_tabla_(
        ['Orden', 'Entrega'],
        d.limite.slice(0, NOTIF_CONFIG.MAX_FILAS_TABLA).map(function (o) {
          return [
            notif_celdaCliente_(o.cliente, o.folio + ' · ' + o.nom),
            notif_chip_(o.dias === 0 ? 'hoy' : (o.dias === 1 ? 'mañana' : 'en ' + o.dias + ' días'), 'ambar')
          ];
        }), [1]);
      filas += notif_seccion_(
        notif_separador_() +
        etiquetaEmail_('En límite · ' + d.limite.length) + tabla, 18);
    }

    // ── En pausa ──────────────────────────────────────────────────────────
    if (d.pausa.length) {
      const tabla = notif_tabla_(
        ['Orden', 'Detenida'],
        d.pausa.slice(0, NOTIF_CONFIG.MAX_FILAS_TABLA).map(function (o) {
          const dias = o.diasPausa === null ? '—' : o.diasPausa + ' días';
          return [
            notif_celdaCliente_(o.cliente, o.folio + (o.motivo ? ' · ' + o.motivo : '')),
            notif_chip_(dias, (o.diasPausa !== null && o.diasPausa > NOTIF_CONFIG.DIAS_ANTIGUO) ? 'rojo' : 'ambar')
          ];
        }), [1]);
      filas += notif_seccion_(
        notif_separador_() +
        etiquetaEmail_('En pausa · ' + d.pausa.length) + tabla, 18);
    }

    filas += notif_seccion_(
      notif_botonera_(botonEmail_('Abrir SEADB', NOTIF_CONFIG.URL_SEADB, 'primario')), 24);

    const titulo = pendientes + (pendientes === 1 ? ' orden requiere' : ' órdenes requieren') + ' acción';
    const html = envolturaEmail_(
      titulo + ' · ' + d.vencidas.length + ' vencidas',
      encabezadoEmail_({
        kicker:    'Tablero operativo',
        titulo:    titulo,
        subtitulo: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "EEEE d 'de' MMMM, HH:mm") +
                   ' · ' + d.total + ' OTs abiertas'
      }) +
      filas +
      pieEmail_([
        'Enviado lunes y jueves a las 07:30 sólo cuando hay órdenes que requieren acción.',
        'Fuente: hoja ' + CONFIG.SHEET_OT + ' · Sistema SEA — ' + CONFIG.COMPANY_NAME
      ])
    );

    const asunto = 'Operación · ' + d.vencidas.length + ' vencidas, ' +
                   d.limite.length + ' en límite, ' + d.pausa.length + ' en pausa';

    const texto = 'Órdenes que requieren acción: ' + pendientes + '\n' +
      'Vencidas: ' + d.vencidas.length + ' · En límite: ' + d.limite.length +
      ' · En pausa: ' + d.pausa.length + '\n\n' +
      d.vencidas.map(function (o) {
        return '- ' + o.folio + ' | ' + o.cliente + ' | ' + Math.abs(o.dias) + ' días de atraso';
      }).join('\n') +
      '\n\nDetalle completo: ' + NOTIF_CONFIG.URL_SEADB;

    return notif_enviar_({
      tipo:  'DIGEST_OPERATIVO',
      clave: Utilities.formatDate(hoy, CONFIG.TIMEZONE, 'yyyy-MM-dd'),
      destinatarios: NOTIF_CONFIG.DESTINATARIOS.OPERATIVO,
      asunto: asunto, html: html, texto: texto,
      detalle: { vencidas: folios, limite: d.limite.length, pausa: d.pausa.length }
    });

  } catch (err) {
    Logger.log('notif_enviarDigestOperativo error: ' + err.message + '\n' + (err.stack || ''));
    return { enviado: false, razon: 'excepcion', error: String(err) };
  }
}


// =========================================================================
// CORREO 2 — RESUMEN SEMANAL DE DIRECCIÓN (viernes)
// =========================================================================

function notif_enviarResumenDireccion() {
  try {
    const hoy = notif_hoy_();
    const lunes = notif_lunesDe_(hoy);
    const lunesPrevio = new Date(lunes.getTime()); lunesPrevio.setDate(lunesPrevio.getDate() - 7);
    const domingo = new Date(lunes.getTime());     domingo.setDate(domingo.getDate() + 6);

    const ordenes  = notif_ordenesAbiertas_();
    const cartera  = notif_carteraInformes_();

    // ── Movimiento de la semana ───────────────────────────────────────────
    const filasOT  = notif_leerHoja_(CONFIG.SHEET_OT);
    const filasInf = notif_leerHoja_(CONFIG.SHEET_INFORMES);

    function enRango(f, desde, hasta) {
      return f && f.getTime() >= desde.getTime() && f.getTime() <= hasta.getTime();
    }

    let otNuevas = 0, otEntregadas = 0;
    filasOT.forEach(function (fila) {
      if (!String(fila[CO.OT] || '').trim()) return;
      if (enRango(notif_parseFecha_(fila[CO.FECHA]), lunes, domingo))       otNuevas++;
      if (enRango(notif_parseFecha_(fila[CO.FECHA_REAL]), lunes, domingo))  otEntregadas++;
    });

    let infSemana = 0, infPrevia = 0;
    const empresas = {};
    filasInf.forEach(function (fila) {
      const f = notif_parseFecha_(fila[CI.TIMESTAMP]);
      if (!f) return;
      if (enRango(f, lunes, domingo)) {
        infSemana++;
        const e = notif_normalizar_(fila[CI.CLIENTE]);
        if (e) empresas[e] = true;
      } else if (enRango(f, lunesPrevio, new Date(lunes.getTime() - 86400000))) {
        infPrevia++;
      }
    });

    const delta = infSemana - infPrevia;
    const notaDelta = delta === 0 ? 'igual que la semana anterior'
                    : (delta > 0 ? '▲ +' + delta : '▼ ' + delta) + ' vs semana anterior';

    let filas = '';

    // ── La semana en números ──────────────────────────────────────────────
    filas += notif_seccion_(
      etiquetaEmail_('La semana en números') +
      '<table width="100%" border="0" cellpadding="0" cellspacing="0">' +
      notif_filaCifra_('Informes registrados', infSemana, notaDelta, delta >= 0 ? 'verde' : 'ambar') +
      notif_filaCifra_('Órdenes de trabajo nuevas', otNuevas) +
      notif_filaCifra_('Órdenes entregadas', otEntregadas,
        (otNuevas > otEntregadas ? 'entraron ' + otNuevas + ', salieron ' + otEntregadas : ''),
        'ambar') +
      notif_filaCifra_('Empresas atendidas', Object.keys(empresas).length) +
      '</table>');

    // ── Estado de la operación ────────────────────────────────────────────
    const filasEstado = [
      ['Órdenes vencidas', notif_chip_(String(ordenes.vencidas.length), ordenes.vencidas.length ? 'rojo' : 'verde')],
      ['En límite · entregan en menos de ' + NOTIF_CONFIG.DIAS_LIMITE + ' días',
       notif_chip_(String(ordenes.limite.length), ordenes.limite.length ? 'ambar' : 'verde')],
      ['En pausa', notif_chip_(String(ordenes.pausa.length), ordenes.pausa.length ? 'rojo' : 'verde')],
      ['Con holgura', '<span style="font-weight:700;">' + ordenes.holgura.length + '</span>']
    ].map(function (p) {
      return ['<span style="font-size:14px;">' + escHtml_(p[0]) + '</span>', p[1]];
    });

    filas += notif_seccion_(
      notif_separador_() + etiquetaEmail_('Estado de la operación') +
      notif_tabla_(['Concepto', 'Total'], filasEstado, [1]), 18);

    // ── Antigüedad de los informes abiertos ───────────────────────────────
    const a = cartera.antiguedad;
    const maxBucket = Math.max(a.d0_30, a.d31_60, a.d61_90, a.d90, 1);
    filas += notif_seccion_(
      notif_separador_() + etiquetaEmail_('Antigüedad de los informes abiertos') +
      '<table width="100%" border="0" cellpadding="0" cellspacing="0">' +
      notif_barra_('0–30 días',  a.d0_30,  a.d0_30  / maxBucket * 100, EMAIL_COLORS_.verde) +
      notif_barra_('31–60 días', a.d31_60, a.d31_60 / maxBucket * 100, EMAIL_COLORS_.verdeTenue) +
      notif_barra_('61–90 días', a.d61_90, a.d61_90 / maxBucket * 100, EMAIL_COLORS_.ambar) +
      notif_barra_('Más de 90 días', a.d90, a.d90 / maxBucket * 100, '#A8452F', true) +
      '</table>' +
      (cartera.abiertos.length
        ? '<p style="margin:14px 0 0 0; font-size:12px; color:' + EMAIL_COLORS_.textoSuave +
          '; line-height:1.6;"><strong style="color:' + EMAIL_COLORS_.ambarTexto + ';">' +
          cartera.atascados + ' informes llevan más de ' + NOTIF_CONFIG.DIAS_ANTIGUO +
          ' días abiertos</strong> — el ' +
          Math.round(cartera.atascados / cartera.abiertos.length * 100) +
          ' % de la cartera.</p>'
        : ''), 18);

    // ── Requiere tu atención ──────────────────────────────────────────────
    const viejos = cartera.abiertos.filter(function (i) { return i.dias !== null; }).slice(0, 5);
    if (viejos.length) {
      filas += notif_seccion_(
        notif_separador_() + etiquetaEmail_('Requiere tu atención') +
        '<p style="margin:-6px 0 12px 0; font-size:12px; color:#8A948A;">Los ' + viejos.length +
        ' informes abiertos más antiguos.</p>' +
        notif_tabla_(['Días', 'Cliente', 'Etapa', 'Responsable'],
          viejos.map(function (i) {
            return [
              '<span style="color:#A8452F; font-weight:800;">' + i.dias + '</span>',
              notif_celdaCliente_(i.cliente, i.num),
              '<span style="font-size:12px;">' + escHtml_(i.etapa) + '</span>',
              '<span style="font-size:12px;">' + escHtml_(notif_corta_(i.responsable, 20)) + '</span>'
            ];
          })) +
        notif_masFilas_(cartera.atascados - viejos.length, 'informes con más de ' +
          NOTIF_CONFIG.DIAS_ANTIGUO + ' días'), 18);
    }

    // ── Carga por responsable ─────────────────────────────────────────────
    const personas = Object.keys(cartera.porResponsable)
      .map(function (k) { return cartera.porResponsable[k]; })
      .sort(function (x, y) { return y.total - x.total; });

    if (personas.length) {
      const totalAsignado = personas.reduce(function (s, p) { return s + p.total; }, 0);
      filas += notif_seccion_(
        notif_separador_() + etiquetaEmail_('Informes abiertos por responsable') +
        '<table width="100%" border="0" cellpadding="0" cellspacing="0">' +
        personas.map(function (p) {
          const pct = Math.round(p.total / totalAsignado * 100);
          return '<tr><td width="42%" style="padding:9px 14px 9px 0; border-bottom:1px solid ' +
            EMAIL_COLORS_.bordeSuave + '; font-size:14px; color:' + EMAIL_COLORS_.textoSuave + ';">' +
            escHtml_(p.nombre) + '</td>' +
            '<td style="padding:9px 0; border-bottom:1px solid ' + EMAIL_COLORS_.bordeSuave +
            '; font-size:14px; color:' + EMAIL_COLORS_.textoFuerte + '; font-weight:600;">' +
            p.total + ' <span style="color:#8A948A; font-weight:400; font-size:12px;">· ' +
            pct + ' %</span></td></tr>';
        }).join('') +
        '</table>', 18);
    }

    filas += notif_seccion_(
      notif_botonera_(botonEmail_('Abrir SEADB', NOTIF_CONFIG.URL_SEADB, 'primario')), 26);

    const semana = notif_claveSemana_(hoy);
    const html = envolturaEmail_(
      infSemana + ' informes nuevos · ' + ordenes.vencidas.length + ' OTs vencidas',
      encabezadoEmail_({
        kicker:    'Resumen semanal de dirección',
        titulo:    'Semana ' + semana.split('-W')[1] + ' · ' + hoy.getFullYear(),
        subtitulo: notif_fmtFecha_(lunes) + ' al ' + notif_fmtFecha_(domingo),
        metas: [
          { etiqueta: 'OTs abiertas',      valor: String(ordenes.total) },
          { etiqueta: 'Informes abiertos', valor: String(cartera.abiertos.length) },
          { etiqueta: 'Vencidas',          valor: String(ordenes.vencidas.length) }
        ]
      }) +
      filas +
      pieEmail_([
        'Resumen generado los viernes a las 17:00.',
        'Fuente: hojas ' + CONFIG.SHEET_OT + ' e ' + CONFIG.SHEET_INFORMES +
        ' · Sistema SEA — ' + CONFIG.COMPANY_NAME
      ])
    );

    const asunto = 'Semana ' + semana.split('-W')[1] + ' · ' + infSemana +
                   ' informes nuevos · ' + ordenes.vencidas.length + ' OTs vencidas · ' +
                   cartera.atascados + ' informes >' + NOTIF_CONFIG.DIAS_ANTIGUO + ' días';

    const texto = 'Resumen semanal — ' + notif_fmtFecha_(lunes) + ' al ' + notif_fmtFecha_(domingo) + '\n\n' +
      'Informes nuevos: ' + infSemana + ' (semana previa: ' + infPrevia + ')\n' +
      'OTs nuevas: ' + otNuevas + ' · OTs entregadas: ' + otEntregadas + '\n' +
      'OTs abiertas: ' + ordenes.total + ' (vencidas ' + ordenes.vencidas.length +
      ', en límite ' + ordenes.limite.length + ', en pausa ' + ordenes.pausa.length + ')\n' +
      'Informes abiertos: ' + cartera.abiertos.length + ' · con más de ' +
      NOTIF_CONFIG.DIAS_ANTIGUO + ' días: ' + cartera.atascados + '\n\n' +
      'Detalle completo: ' + NOTIF_CONFIG.URL_SEADB;

    return notif_enviar_({
      tipo:  'RESUMEN_DIRECCION',
      clave: semana,
      destinatarios: NOTIF_CONFIG.DESTINATARIOS.DIRECCION,
      asunto: asunto, html: html, texto: texto,
      detalle: { informes: infSemana, vencidas: ordenes.vencidas.length, atascados: cartera.atascados }
    });

  } catch (err) {
    Logger.log('notif_enviarResumenDireccion error: ' + err.message + '\n' + (err.stack || ''));
    return { enviado: false, razon: 'excepcion', error: String(err) };
  }
}


// =========================================================================
// CORREO 3 — RENOVACIONES (lunes)
// =========================================================================

function notif_enviarRenovaciones() {
  try {
    const hoy = notif_hoy_();
    const todas = notif_renovaciones_();
    const ventana = todas.filter(function (r) { return r.dias < NOTIF_CONFIG.VENTANA_RENOVACION; });

    if (!ventana.length) {
      Logger.log('notif: sin renovaciones en ventana (' + todas.length +
                 ' en seguimiento); no se envía.');
      return { enviado: false, razon: 'vacio' };
    }

    const vencidas = ventana.filter(function (r) { return r.dias < 0; });
    const proximas = ventana.filter(function (r) { return r.dias >= 0; });

    const tabla = notif_tabla_(
      ['Cliente / sucursal', 'Servicio', 'Vence', 'Estado'],
      ventana.slice(0, NOTIF_CONFIG.MAX_FILAS_TABLA).map(function (r) {
        const etiqueta = r.dias < 0
          ? 'Vencida · ' + Math.abs(r.dias) + ' d'
          : (r.dias === 0 ? 'Vence hoy' : 'En ' + r.dias + ' días');
        return [
          notif_celdaCliente_(r.cliente, r.sucursal || r.rfc),
          '<span style="font-size:12px;">' + escHtml_(r.servicio) +
            '<br><span style="color:#8A948A; font-size:11px;">' +
            (r.ciclo === 1 ? 'anual' : 'bienal') + '</span></span>',
          '<span style="font-size:12px;">' + notif_fmtFecha_(r.proxima) + '</span>',
          notif_chip_(etiqueta, r.dias < 0 ? 'rojo' : 'ambar')
        ];
      }));

    // Panorama de los siguientes 90 días, para planear campañas.
    const siguientes = todas.filter(function (r) {
      return r.dias >= NOTIF_CONFIG.VENTANA_RENOVACION && r.dias < 90;
    }).length;

    let filas = notif_seccion_(
      etiquetaEmail_('Acción esta semana') +
      '<p style="margin:-6px 0 18px 0; font-size:13px; color:' + EMAIL_COLORS_.texto +
      '; line-height:1.6;">Servicios cuyo ciclo normativo termina. ' +
      'Cotizar antes de que el cliente quede sin vigencia.</p>' +
      tabla +
      notif_masFilas_(ventana.length - Math.min(ventana.length, NOTIF_CONFIG.MAX_FILAS_TABLA),
                      'renovaciones'));

    if (siguientes > 0) {
      filas += notif_seccion_(
        '<div style="background-color:' + EMAIL_COLORS_.panel + '; border-left:3px solid ' +
        EMAIL_COLORS_.verde + '; padding:14px 16px; border-radius:0 6px 6px 0;">' +
        etiquetaEmail_('Lo que viene') +
        '<p style="margin:0; font-size:13px; color:' + EMAIL_COLORS_.texto + '; line-height:1.65;">' +
        siguientes + (siguientes === 1 ? ' renovación más entra' : ' renovaciones más entran') +
        ' en ventana dentro de los próximos 90 días.</p></div>', 22);
    }

    filas += notif_seccion_(
      notif_botonera_(botonEmail_('Ver todas en SEADB', NOTIF_CONFIG.URL_SEADB, 'primario')), 22);

    const titulo = ventana.length + (ventana.length === 1 ? ' cliente por contactar' : ' clientes por contactar');
    const html = envolturaEmail_(
      titulo + ' · ' + vencidas.length + ' vencidas',
      encabezadoEmail_({
        kicker:    'Oportunidades de renovación',
        titulo:    titulo,
        subtitulo: 'Corte al ' + notif_fmtFecha_(hoy),
        metas: [
          { etiqueta: 'Vencidas',          valor: String(vencidas.length) },
          { etiqueta: 'Menos de 30 días',  valor: String(proximas.length) },
          { etiqueta: 'En seguimiento',    valor: String(todas.length) }
        ]
      }) +
      filas +
      pieEmail_([
        'Enviado los lunes a las 08:00 sólo cuando hay renovaciones en ventana.',
        'Fuente: hoja ' + CONFIG.SHEET_INFORMES + ' · Sistema SEA — ' + CONFIG.COMPANY_NAME
      ])
    );

    const asunto = 'Renovaciones · ' + vencidas.length + ' vencidas, ' +
                   proximas.length + ' en menos de ' + NOTIF_CONFIG.VENTANA_RENOVACION + ' días';

    const texto = titulo + '\n\n' + ventana.map(function (r) {
      return '- ' + r.cliente + (r.sucursal ? ' / ' + r.sucursal : '') + ' | ' + r.servicio +
             ' | vence ' + notif_fmtFecha_(r.proxima) +
             ' | ' + (r.dias < 0 ? 'VENCIDA hace ' + Math.abs(r.dias) + ' días' : 'en ' + r.dias + ' días');
    }).join('\n') + '\n\nDetalle completo: ' + NOTIF_CONFIG.URL_SEADB;

    return notif_enviar_({
      tipo:  'RENOVACIONES',
      clave: notif_claveSemana_(hoy),
      destinatarios: NOTIF_CONFIG.DESTINATARIOS.RENOVACIONES,
      asunto: asunto, html: html, texto: texto,
      detalle: { enVentana: ventana.length, vencidas: vencidas.length, rastreadas: todas.length }
    });

  } catch (err) {
    Logger.log('notif_enviarRenovaciones error: ' + err.message + '\n' + (err.stack || ''));
    return { enviado: false, razon: 'excepcion', error: String(err) };
  }
}


// =========================================================================
// INSTALACIÓN DE TRIGGERS
// =========================================================================

/**
 * Crea los tres triggers de tiempo. Borra primero los que ya existan de estas
 * funciones, así que se puede volver a ejecutar sin duplicar.
 *
 * Las horas se interpretan en la zona horaria DEL PROYECTO de Apps Script,
 * que no necesariamente es la de CONFIG.TIMEZONE. Verificar con notif_estado().
 */
function configurarTriggersNotificaciones() {
  const manejadas = [
    'notif_enviarDigestOperativo',
    'notif_enviarResumenDireccion',
    'notif_enviarRenovaciones'
  ];

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (manejadas.indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
      Logger.log('Trigger anterior eliminado: ' + t.getHandlerFunction());
    }
  });

  ScriptApp.newTrigger('notif_enviarDigestOperativo').timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).nearMinute(30).create();
  ScriptApp.newTrigger('notif_enviarDigestOperativo').timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(7).nearMinute(30).create();

  ScriptApp.newTrigger('notif_enviarResumenDireccion').timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(17).create();

  ScriptApp.newTrigger('notif_enviarRenovaciones').timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();

  Logger.log('✅ Triggers de notificaciones configurados:');
  Logger.log('   Digest operativo    → lunes y jueves 07:30');
  Logger.log('   Resumen dirección   → viernes 17:00');
  Logger.log('   Renovaciones        → lunes 08:00');
  Logger.log('   Verifica en: Activadores, en el menú lateral del editor.');
}

/** Quita los triggers de este módulo sin tocar los demás del proyecto. */
function eliminarTriggersNotificaciones() {
  const manejadas = [
    'notif_enviarDigestOperativo',
    'notif_enviarResumenDireccion',
    'notif_enviarRenovaciones'
  ];
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (manejadas.indexOf(t.getHandlerFunction()) >= 0) { ScriptApp.deleteTrigger(t); n++; }
  });
  Logger.log('Triggers de notificaciones eliminados: ' + n);
}


// =========================================================================
// DIAGNÓSTICO Y PRUEBAS — ejecutar a mano desde el editor
// =========================================================================

/** Revisa configuración, zona horaria, datos y triggers. No envía nada. */
function notif_estado() {
  Logger.log('=== ESTADO DE NOTIFICACIONES SEA ===');
  Logger.log('Activo            : ' + (notif_activo_() ? 'sí' : 'NO (NOTIF_ACTIVO=false)'));
  const prueba = notif_dryRun_();
  Logger.log('Modo prueba       : ' + (prueba ? 'SÍ → todo va a ' + prueba : 'no'));

  const tzProyecto = Session.getScriptTimeZone();
  Logger.log('Zona del proyecto : ' + tzProyecto);
  Logger.log('CONFIG.TIMEZONE   : ' + CONFIG.TIMEZONE);
  if (tzProyecto !== CONFIG.TIMEZONE) {
    Logger.log('   ⚠ Los triggers usan la zona DEL PROYECTO. Si no coincide con');
    Logger.log('     la operación, ajústala en Configuración del proyecto.');
  }

  const ordenes = notif_ordenesAbiertas_();
  Logger.log('--- Órdenes de trabajo ---');
  Logger.log('  abiertas: ' + ordenes.total + ' | vencidas: ' + ordenes.vencidas.length +
             ' | en límite: ' + ordenes.limite.length + ' | en pausa: ' + ordenes.pausa.length +
             ' | con holgura: ' + ordenes.holgura.length + ' | sin fecha: ' + ordenes.sinFecha.length);
  if (ordenes.sinFecha.length) {
    Logger.log('  ⚠ Sin fecha límite ni de visita (no entran al digest): ' +
               ordenes.sinFecha.map(function (o) { return o.folio; }).join(', '));
  }

  const cartera = notif_carteraInformes_();
  Logger.log('--- Informes ---');
  Logger.log('  registros: ' + cartera.totalRegistros + ' | abiertos: ' + cartera.abiertos.length +
             ' | con más de ' + NOTIF_CONFIG.DIAS_ANTIGUO + ' días: ' + cartera.atascados);
  Logger.log('  por etapa: ' + JSON.stringify(cartera.porEtapa));

  const renov = notif_renovaciones_();
  const ventana = renov.filter(function (r) { return r.dias < NOTIF_CONFIG.VENTANA_RENOVACION; });
  Logger.log('--- Renovaciones ---');
  Logger.log('  rastreadas: ' + renov.length + ' | en ventana: ' + ventana.length);
  if (renov.length && !ventana.length) {
    const prox = renov.filter(function (r) { return r.dias >= 0; })[0];
    if (prox) Logger.log('  próxima: ' + notif_fmtFecha_(prox.proxima) +
                         ' (' + prox.dias + ' días) — ' + prox.cliente);
  }

  Logger.log('--- Destinatarios ---');
  Logger.log('  digest    : ' + NOTIF_CONFIG.DESTINATARIOS.OPERATIVO.join(', '));
  Logger.log('  dirección : ' + NOTIF_CONFIG.DESTINATARIOS.DIRECCION.join(', '));
  Logger.log('  ventas    : ' + NOTIF_CONFIG.DESTINATARIOS.RENOVACIONES.join(', '));

  const triggers = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction().indexOf('notif_enviar') === 0;
  });
  Logger.log('--- Triggers instalados: ' + triggers.length + ' ---');
  triggers.forEach(function (t) { Logger.log('  ' + t.getHandlerFunction()); });

  Logger.log('--- Cuota de correo ---');
  Logger.log('  restante hoy: ' + MailApp.getRemainingDailyQuota() + ' destinatarios');
}

/**
 * Arma los tres correos y escribe su asunto y tamaño en el Log, sin enviar
 * ni tocar la bitácora. Es la forma segura de revisar antes de programar.
 */
function notif_previsualizar() {
  const activoOriginal = notif_prop_(NOTIF_CONFIG.PROP_ACTIVO, 'true');
  try {
    PropertiesService.getScriptProperties().setProperty(NOTIF_CONFIG.PROP_ACTIVO, 'false');
    Logger.log('=== PREVISUALIZACIÓN (no se envía nada) ===');
    [
      ['Digest operativo',  notif_enviarDigestOperativo],
      ['Resumen dirección', notif_enviarResumenDireccion],
      ['Renovaciones',      notif_enviarRenovaciones]
    ].forEach(function (par) {
      const r = par[1]();
      Logger.log(par[0] + ' → ' + (r.razon === 'vacio'
        ? 'no se enviaría: no hay nada que reportar'
        : 'se enviaría (bloqueado por la previsualización)'));
    });
    Logger.log('Para ver el contenido real, usa notif_probarEnvio() con NOTIF_DRY_RUN puesto.');
  } finally {
    PropertiesService.getScriptProperties().setProperty(NOTIF_CONFIG.PROP_ACTIVO, activoOriginal);
  }
}

/**
 * Envía los tres correos al buzón de NOTIF_DRY_RUN, marcados [PRUEBA].
 * No bloquea los envíos reales posteriores: las pruebas se registran aparte.
 */
function notif_probarEnvio() {
  const prueba = notif_dryRun_();
  if (!prueba) {
    Logger.log('⚠ Pon primero la propiedad ' + NOTIF_CONFIG.PROP_DRY_RUN +
               ' con tu correo, en Configuración del proyecto → Propiedades del script.');
    return;
  }
  Logger.log('Enviando los tres correos de prueba a ' + prueba + '…');
  Logger.log('  digest      : ' + JSON.stringify(notif_enviarDigestOperativo()));
  Logger.log('  dirección   : ' + JSON.stringify(notif_enviarResumenDireccion()));
  Logger.log('  renovaciones: ' + JSON.stringify(notif_enviarRenovaciones()));
}
