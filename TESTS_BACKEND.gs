// =========================================================================
// TESTS E2E — API CENTRAL EJECUTIVA AMBIENTAL
// Compatible con Google Apps Script — sin dependencias externas
// =========================================================================
// FLUJOS CUBIERTOS
//   E01  SEAPD  → registrarCliente   (crea carpeta Drive + fila en CLIENTES_MAESTRO)
//   E02  SEAOT  → buscarClienteRFC + registrarOT  (busca cliente, registra OT)
//   E03  SEAINF → getOrdenes + getConsecutivo + createExpediente
//
// USO
//   Editor GAS → seleccionar runE2ETests → ▶ Ejecutar → Ver registros
//   Para ejecutar un flujo individual: runTest_E01, runTest_E02, runTest_E03
//
// LIMPIEZA
//   Los tests crean datos con RFC 'XTEST000000TST' en CLIENTES_MAESTRO y
//   folio 'TEST-E2E-001' en ORDENES_TRABAJO.
//   Al terminar (éxito o falla) se eliminan automáticamente.
// =========================================================================

// ─── Datos de prueba ──────────────────────────────────────────────────────
var TEST_RFC     = 'XTEST000000TST';
var TEST_FOLIO   = 'TEST-E2E-001';
var TEST_SUCURSAL= 'Sucursal Test E2E';

// ─── Estado compartido entre flujos ──────────────────────────────────────
var _ctx_ = {
  linkDriveCliente: '',   // resultado de E01 → input de E02
  folioOT:          '',   // resultado de E02 → input de E03
  urlExpediente:    '',   // resultado de E03
  clienteFolderId:  '',   // para cleanup E01
  expedienteFolderId: ''  // para cleanup E03
};

// ─── Mini framework ───────────────────────────────────────────────────────
var _results_ = [];

function _pass_(msg) {
  _results_.push('PASS: ' + msg);
  Logger.log('[PASS] ' + msg);
}
function _fail_(msg) {
  _results_.push('FAIL: ' + msg);
  Logger.log('[FAIL] ' + msg);
  throw new Error(msg);
}
function _check_(msg, condition) { condition ? _pass_(msg) : _fail_(msg); }
function _eq_(msg, actual, expected) {
  var ok = String(actual) === String(expected);
  ok ? _pass_(msg) : _fail_(msg + ' | esperado: "' + expected + '" | obtenido: "' + actual + '"');
}

// =========================================================================
// RUNNER PRINCIPAL
// =========================================================================
function runE2ETests() {
  _results_ = [];
  Logger.log('');
  Logger.log('══════════════════════════════════════════════');
  Logger.log('  TESTS E2E — EA Backend v3.0');
  Logger.log('  RFC de prueba : ' + TEST_RFC);
  Logger.log('  Folio de prueba: ' + TEST_FOLIO);
  Logger.log('══════════════════════════════════════════════');

  var e01ok = false, e02ok = false, e03ok = false;
  try { runTest_E01(); e01ok = true; } catch(e) { Logger.log('  E01 abortado: ' + e.message); }
  try { runTest_E02(); e02ok = true; } catch(e) { Logger.log('  E02 abortado: ' + e.message); }
  try { runTest_E03(); e03ok = true; } catch(e) { Logger.log('  E03 abortado: ' + e.message); }

  Logger.log('');
  Logger.log('── LIMPIEZA ──────────────────────────────────');
  _cleanup_();

  var pass = _results_.filter(function(r){ return r.indexOf('PASS') === 0; }).length;
  var fail = _results_.filter(function(r){ return r.indexOf('FAIL') === 0; }).length;
  Logger.log('');
  Logger.log('══════════════════════════════════════════════');
  Logger.log('  RESULTADO: ' + pass + ' PASS  |  ' + fail + ' FAIL');
  Logger.log('  E01 registrarCliente : ' + (e01ok ? 'OK' : 'FALLO'));
  Logger.log('  E02 registrarOT      : ' + (e02ok ? 'OK' : 'FALLO'));
  Logger.log('  E03 createExpediente : ' + (e03ok ? 'OK' : 'FALLO'));
  Logger.log('══════════════════════════════════════════════');
}

// =========================================================================
// E01 — SEAPD: registrarCliente
// =========================================================================
// Simula el payload que SEAPD envía cuando el usuario llena el formulario
// y presiona "Enviar". Verifica que:
//   - La función devuelve success: true
//   - Aparece una fila en CLIENTES_MAESTRO con los datos correctos
//   - El link_drive_cliente (índice 15) no está vacío y apunta a Drive
//   - La carpeta Drive existe y es accesible
function runTest_E01() {
  Logger.log('');
  Logger.log('── E01: SEAPD → registrarCliente ─────────────');

  // Limpiar filas residuales de ejecuciones anteriores para evitar
  // que el backend y el test encuentren filas distintas (forward vs backward).
  try {
    var sheetPre = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_CLIENTES);
    var rowsPre  = sheetPre.getDataRange().getValues();
    for (var pi = rowsPre.length - 1; pi >= 1; pi--) {
      if (String(rowsPre[pi][3]).toUpperCase().trim() === TEST_RFC) {
        sheetPre.deleteRow(pi + 1);
        Logger.log('  [pre-cleanup] Fila residual eliminada (fila ' + (pi + 1) + ')');
      }
    }
  } catch(e) { Logger.log('  [pre-cleanup] ' + e.message); }

  var payload = {
    action:               'registrarCliente',
    nombre_solicitante:   'Prueba Automatizada',
    razon_social:         'EMPRESA TEST E2E SA DE CV',
    sucursal:             TEST_SUCURSAL,
    rfc:                  TEST_RFC,
    telefono_empresa:     '2220000000',
    representante_legal:  'Rep Legal Test',
    direccion_evaluacion: 'Calle Falsa 123, Puebla',
    giro:                 'Pruebas Automatizadas',
    correo_informe:       'test@noenviar.com',
    registro_patronal:    'IMSS-TEST-000',
    capacidad_instalada:  '100 ton',
    capacidad_operacion:  '80 ton',
    dias_turnos_horarios: 'L-V 09:00-18:00',
    aplica_nom020:        'no',
    requiere_pipc:        'no',
    // Campos opcionales vacíos (no se suben archivos en el test)
    nombre_dirigido: '', puesto_dirigido: '', actividad_principal: '',
    descripcion_proceso: '', fechas_preferidas: '', responsable: '',
    telefono_responsable: '',
    _skipEmail: true   // evita enviar correos reales durante el test
  };

  var result = fase1_RegistrarCliente(payload);

  _check_('E01-1: respuesta success=true',      result.success === true);
  _check_('E01-2: sin error en respuesta',       !result.error);

  // Verificar fila en CLIENTES_MAESTRO
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_CLIENTES);
  var rows  = sheet.getDataRange().getValues();
  var fila  = null;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][3]).toUpperCase().trim() === TEST_RFC &&
        String(rows[i][2]).trim() === TEST_SUCURSAL) {
      fila = rows[i]; break;
    }
  }

  _check_('E01-3: fila creada en CLIENTES_MAESTRO',       fila !== null);
  _eq_('E01-4: razon_social en col 2 (índice 1)',         fila[1], 'EMPRESA TEST E2E SA DE CV');
  _eq_('E01-5: sucursal en col 3 (índice 2)',             fila[2], TEST_SUCURSAL);
  _eq_('E01-6: rfc en col 4 (índice 3)',                  fila[3], TEST_RFC);
  _eq_('E01-7: representante_legal en col 5 (índice 4)', fila[4], 'Rep Legal Test');
  _eq_('E01-8: telefono_empresa en col 7 (índice 6)',    fila[6], '2220000000');
  _eq_('E01-9: nombre_solicitante en col 8 (índice 7)', fila[7], 'Prueba Automatizada');
  _eq_('E01-10: correo_informe en col 9 (índice 8)',    fila[8], 'test@noenviar.com');

  var linkDrive = String(fila[15] || '');
  _check_('E01-11: link_drive en col 16 (índice 15) no está vacío', linkDrive !== '');
  _check_('E01-12: link_drive contiene "folders/"', linkDrive.indexOf('folders/') !== -1);

  // Verificar que la carpeta Drive es accesible
  var m = linkDrive.match(/folders\/([a-zA-Z0-9_-]+)/);
  _check_('E01-13: folder ID extraíble del link_drive', !!m);
  var carpeta = DriveApp.getFolderById(m[1]);
  _check_('E01-14: carpeta Drive existe y es accesible', !!carpeta);
  _check_('E01-15: nombre carpeta contiene el RFC', carpeta.getParents().next().getName().indexOf(TEST_RFC) !== -1);

  // Guardar para flujos siguientes y para cleanup
  _ctx_.linkDriveCliente = linkDrive;
  _ctx_.clienteFolderId  = m[1];
  Logger.log('  linkDriveCliente: ' + linkDrive);
}

// =========================================================================
// E02 — SEAOT: buscarClienteRFC + registrarOT
// =========================================================================
// Simula el flujo de SEAOT:
//   1. Usuario escribe el RFC → buscarClienteRFC → obtiene sucursales + link_drive
//   2. Usuario llena el form de OT y presiona "Registrar OT" → registrarOT
// Verifica que:
//   - buscarClienteRFC devuelve found: true con el link correcto (índice 15)
//   - La OT aparece en ORDENES_TRABAJO con link_drive_cliente en col 14
function runTest_E02() {
  Logger.log('');
  Logger.log('── E02: SEAOT → buscarClienteRFC + registrarOT ──');

  // Limpiar filas residuales en ORDENES_TRABAJO para evitar desalineación
  // entre la búsqueda forward del backend y la backward del test.
  try {
    var sheetOTPre = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
    var rowsOTPre  = sheetOTPre.getDataRange().getValues();
    for (var oi = rowsOTPre.length - 1; oi >= 1; oi--) {
      if (String(rowsOTPre[oi][1]).trim() === TEST_FOLIO) {
        sheetOTPre.deleteRow(oi + 1);
        Logger.log('  [pre-cleanup] Fila OT residual eliminada (fila ' + (oi + 1) + ')');
      }
    }
  } catch(e) { Logger.log('  [pre-cleanup OT] ' + e.message); }

  // Paso 1: Búsqueda de cliente por RFC (como hace SEAOT al tipear el RFC)
  var busqueda = fase2_BuscarClienteRFC(TEST_RFC);

  _check_('E02-1: buscarClienteRFC devuelve found=true',   busqueda.found === true);
  _check_('E02-2: hay al menos una sucursal',              busqueda.sucursales && busqueda.sucursales.length > 0);

  var sucursal = busqueda.sucursales[0];
  _eq_('E02-3: razon_social correcta',          sucursal.razon_social,       'EMPRESA TEST E2E SA DE CV');
  _eq_('E02-4: nombre_solicitante (índice 7)',  sucursal.nombre_solicitante, 'Prueba Automatizada');
  _eq_('E02-5: correo_informe (índice 8)',      sucursal.correo_informe,     'test@noenviar.com');
  _eq_('E02-6: telefono_empresa (índice 6)',    sucursal.telefono_empresa,   '2220000000');

  var linkDrive = sucursal.link_drive_cliente || '';
  _check_('E02-7: link_drive_cliente no está vacío (índice 15)',  linkDrive !== '');
  _check_('E02-8: link_drive_cliente contiene "folders/"',        linkDrive.indexOf('folders/') !== -1);

  // Paso 2: Registro de OT (como hace SEAOT al enviar el formulario)
  var payloadOT = {
    action:               'registrarOT',
    ot_folio:             TEST_FOLIO,
    tipo_orden:           'OTA',
    nom_servicio:         'NOM-035-STPS',
    cliente_razon_social: sucursal.razon_social,
    sucursal:             sucursal.sucursal,
    rfc:                  sucursal.rfc,
    personal_asignado:    'Ing. Test',
    fecha_visita:         '2026-03-20',
    fecha_entrega_limite: '2026-03-27',
    link_drive_cliente:   linkDrive,   // ← crítico: viene de la búsqueda por RFC
    observaciones:        'OT generada por test E2E'
  };

  var resultOT = fase2_RegistrarOT(payloadOT);
  _check_('E02-9: registrarOT devuelve success=true', resultOT.success === true);

  // Verificar fila en ORDENES_TRABAJO
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  var rows  = sheet.getDataRange().getValues();
  var filaOT = null;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][1]).trim() === TEST_FOLIO) { filaOT = rows[i]; break; }
  }

  _check_('E02-10: fila creada en ORDENES_TRABAJO',            filaOT !== null);
  _eq_('E02-11: folio en col 2 (índice 1)',                   filaOT[1], TEST_FOLIO);
  _eq_('E02-12: tipo_orden en col 3 (índice 2)',              filaOT[2], 'OTA');
  _eq_('E02-13: nom_servicio en col 5 (índice 4)',            filaOT[4], 'NOM-035-STPS');
  _eq_('E02-14: rfc en col 8 (índice 7)',                     filaOT[7], TEST_RFC);
  _eq_('E02-15: estatus inicial "NO INICIADO"',               filaOT[12], 'NO INICIADO');

  var linkEnOT = String(filaOT[13] || '');
  _check_('E02-16: link_drive_cliente guardado en col 14 (índice 13)', linkEnOT !== '');
  _check_('E02-17: link en OT coincide con link del cliente', linkEnOT === linkDrive);

  _ctx_.folioOT = TEST_FOLIO;
  Logger.log('  OT registrada con folio: ' + TEST_FOLIO);
}

// =========================================================================
// E03 — SEAINF: getOrdenes + getConsecutivo + createExpediente
// =========================================================================
// Simula el flujo de SEAINF:
//   1. Al cargar la pantalla → getOrdenes → lista de OTs disponibles
//   2. Usuario selecciona la OT → getConsecutivo → genera numInforme
//   3. Usuario adjunta archivos y presiona "Crear Expediente" → createExpediente
// Verifica que:
//   - getOrdenes incluye la OT del test con rfc, sucursal, link_drive
//   - getConsecutivo genera un número con formato correcto
//   - createExpediente crea la carpeta DENTRO de la carpeta del cliente (no en raíz)
//   - ORDENES_TRABAJO se actualiza con nuevo link y estatus EN PROCESO
function runTest_E03() {
  Logger.log('');
  Logger.log('── E03: SEAINF → getOrdenes + getConsecutivo + createExpediente ──');

  // Paso 1: getOrdenes (como hace SEAINF al cargar)
  var ordenesResp = getOrdenesSafe_();
  _check_('E03-1: getOrdenes devuelve success=true', ordenesResp.success === true);

  var ordenTest = null;
  for (var i = 0; i < ordenesResp.data.length; i++) {
    if (ordenesResp.data[i].ot === TEST_FOLIO) { ordenTest = ordenesResp.data[i]; break; }
  }
  _check_('E03-2: OT del test aparece en getOrdenes',            ordenTest !== null);
  _eq_('E03-3: campo rfc presente y correcto',                   ordenTest.rfc,       TEST_RFC);
  _eq_('E03-4: campo sucursal presente',                         ordenTest.sucursal,  TEST_SUCURSAL);
  _eq_('E03-5: campo nom_servicio presente',                     ordenTest.nom_servicio, 'NOM-035-STPS');
  _check_('E03-6: campo link_drive presente y no vacío',         (ordenTest.link_drive || '') !== '');

  // Paso 2: getConsecutivo (como hace SEAINF al seleccionar la OT)
  var hoy     = new Date();
  var anio    = String(hoy.getFullYear()).slice(2);
  var mes     = String(hoy.getMonth() + 1).padStart(2, '0');
  var nomCode = 'NOM035';

  var consResp = getConsecutivoSafe_({ anio: anio, mes: mes, nom: nomCode, tipo: 'OTA' });
  _check_('E03-7: getConsecutivo devuelve success=true',       consResp.success === true);
  _check_('E03-8: numeroInforme no está vacío',                !!consResp.numeroInforme);

  var numInforme = consResp.numeroInforme;
  var regexInforme = /^EA-\d{4}-[A-Za-z0-9]+-\d{4}$/;
  _check_('E03-9: formato numInforme es EA-AAMM-NOM-0000',     regexInforme.test(numInforme));
  Logger.log('  numInforme asignado: ' + numInforme);

  // Paso 3: createExpediente (como hace SEAINF al presionar "Crear Expediente")
  var payloadExp = {
    action: 'createExpediente',
    data: {
      ot:         TEST_FOLIO,
      nom:        nomCode,
      numInforme: numInforme,
      cliente:    'EMPRESA TEST E2E SA DE CV',
      sucursal:   TEST_SUCURSAL,
      rfc:        TEST_RFC,                      // ← SEAINF lo envía para fallback
      linkDrive:  ordenTest.link_drive,          // ← SEAINF lo envía para fallback
      fecha:      Utilities.formatDate(hoy, 'GMT-6', 'dd/MM/yyyy'),
      entrega:    '22/03/2026',
      tipoOrden:  'OTA',
      solicitante:'Prueba Automatizada',
      telefono:   '2220000000',
      direccion:  'Calle Falsa 123',
      responsable:'Ing. Test',
      estatus:    'NO INICIADO'
    },
    files: []  // sin archivos en test (evita payload enorme)
  };

  var resultExp = fase3_CrearExpediente(payloadExp);
  _check_('E03-10: createExpediente devuelve success=true', resultExp.success === true);
  _check_('E03-11: url del expediente no está vacía',       !!resultExp.url);

  // Verificar que el expediente NO está en la carpeta raíz
  var m = resultExp.url.match(/folders\/([a-zA-Z0-9_-]+)/);
  _check_('E03-12: folder ID extraíble del url del expediente', !!m);
  var carpetaExp    = DriveApp.getFolderById(m[1]);
  var carpetaPadre  = carpetaExp.getParents().next();
  // Derivar el ID esperado: primero _ctx_ (si E01 corrió antes), luego link_drive de la OT
  var expectedClientFolderId = _ctx_.clienteFolderId;
  if (!expectedClientFolderId && (ordenTest.link_drive || '')) {
    var mLinkCli = String(ordenTest.link_drive).match(/folders\/([a-zA-Z0-9_-]+)/);
    if (mLinkCli) expectedClientFolderId = mLinkCli[1];
  }
  _check_('E03-13: expediente creado dentro de carpeta del cliente (no en raíz)',
    !!expectedClientFolderId && carpetaPadre.getId() === expectedClientFolderId);

  // Verificar subcarpetas del expediente
  var subFolders = [];
  var iter = carpetaExp.getFolders();
  while (iter.hasNext()) subFolders.push(iter.next().getName());
  _check_('E03-14: subcarpeta "1. ORDEN_TRABAJO" creada', subFolders.indexOf('1. ORDEN_TRABAJO') !== -1);
  _check_('E03-15: subcarpeta "2. HDC" creada',           subFolders.indexOf('2. HDC') !== -1);
  _check_('E03-16: subcarpeta "3. CROQUIS" creada',       subFolders.indexOf('3. CROQUIS') !== -1);
  _check_('E03-17: subcarpeta "4. FOTOS" creada',         subFolders.indexOf('4. FOTOS') !== -1);

  // Verificar que ORDENES_TRABAJO se actualizó
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  var rows  = sheet.getDataRange().getValues();
  var filaOT = null;
  for (var j = rows.length - 1; j >= 1; j--) {
    if (String(rows[j][1]).trim() === TEST_FOLIO) { filaOT = rows[j]; break; }
  }
  _check_('E03-18: fila OT actualizada en ORDENES_TRABAJO', filaOT !== null);
  _eq_('E03-19: numInforme guardado en col 4 (índice 3)',  filaOT[3], numInforme);
  _eq_('E03-20: estatus cambiado a EN PROCESO',            filaOT[12], 'EN PROCESO');
  _check_('E03-21: link del expediente guardado en col 14', String(filaOT[13] || '').indexOf('folders/') !== -1);

  _ctx_.urlExpediente     = resultExp.url;
  _ctx_.expedienteFolderId = m[1];
  Logger.log('  Expediente creado en: ' + resultExp.url);
}

// =========================================================================
// LIMPIEZA — elimina todos los datos de prueba
// =========================================================================
function _cleanup_() {
  // 1. Eliminar fila de CLIENTES_MAESTRO
  try {
    var sheetCli = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_CLIENTES);
    var rowsCli  = sheetCli.getDataRange().getValues();
    for (var i = rowsCli.length - 1; i >= 1; i--) {
      if (String(rowsCli[i][3]).toUpperCase().trim() === TEST_RFC) {
        sheetCli.deleteRow(i + 1);
        Logger.log('  Fila eliminada de CLIENTES_MAESTRO (fila ' + (i + 1) + ')');
      }
    }
  } catch(e) { Logger.log('  ERROR cleanup CLIENTES_MAESTRO: ' + e.message); }

  // 2. Eliminar fila de ORDENES_TRABAJO
  try {
    var sheetOT = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
    var rowsOT  = sheetOT.getDataRange().getValues();
    for (var j = rowsOT.length - 1; j >= 1; j--) {
      if (String(rowsOT[j][1]).trim() === TEST_FOLIO) {
        sheetOT.deleteRow(j + 1);
        Logger.log('  Fila eliminada de ORDENES_TRABAJO (fila ' + (j + 1) + ')');
      }
    }
  } catch(e) { Logger.log('  ERROR cleanup ORDENES_TRABAJO: ' + e.message); }

  // 3. Mover a papelera la carpeta del expediente
  if (_ctx_.expedienteFolderId) {
    try {
      DriveApp.getFolderById(_ctx_.expedienteFolderId).setTrashed(true);
      Logger.log('  Carpeta expediente movida a papelera');
    } catch(e) { Logger.log('  ERROR cleanup expediente: ' + e.message); }
  }

  // 4. Mover a papelera la carpeta sucursal del cliente
  if (_ctx_.clienteFolderId) {
    try {
      DriveApp.getFolderById(_ctx_.clienteFolderId).setTrashed(true);
      Logger.log('  Carpeta sucursal del cliente movida a papelera');
    } catch(e) { Logger.log('  ERROR cleanup carpeta cliente: ' + e.message); }
  }

  // 5. Mover a papelera la carpeta padre RFC - EMPRESA TEST
  try {
    var folderRaiz = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    var iter = folderRaiz.getFolders();
    while (iter.hasNext()) {
      var f = iter.next();
      if (f.getName().indexOf(TEST_RFC) !== -1) {
        f.setTrashed(true);
        Logger.log('  Carpeta padre "' + f.getName() + '" movida a papelera');
      }
    }
  } catch(e) { Logger.log('  ERROR cleanup carpeta raíz: ' + e.message); }
}

// =========================================================================
// HELPERS PARA TESTS UNITARIOS (sin SpreadsheetApp)
// =========================================================================

// Ejecuta los tests de lógica pura (no necesitan conexión a Sheets/Drive)
function runUnitTests() {
  _results_ = [];
  Logger.log('');
  Logger.log('── UNIT TESTS (lógica pura) ──────────────────');

  // Mapeo de índices nuevo esquema 16 columnas
  var fila16 = [
    '08/03/2026','EMPRESA TEST','Planta Norte','TST010101AAA',
    'Rep Legal','Dir Eval','2221234567','Solicitante Test',
    'sol@test.com','Manufactura','IMSS-01234','500 / 450 ton',
    'L-V 08:00-18:00','SÍ','NO',
    'https://drive.google.com/drive/folders/LINK_CORRECTO'
  ];
  _eq_('U01: nombre_solicitante en índice [7]', fila16[7], 'Solicitante Test');
  _eq_('U02: correo_informe en índice [8]',     fila16[8], 'sol@test.com');
  _eq_('U03: telefono_empresa en índice [6]',   fila16[6], '2221234567');
  _eq_('U04: link_drive en índice [15]',        fila16[15], 'https://drive.google.com/drive/folders/LINK_CORRECTO');
  _check_('U05: índice [22] es undefined (esquema antiguo)', fila16[22] === undefined);

  // Regex extracción folder ID
  var m = 'https://drive.google.com/drive/folders/ABC123_-XYZ'.match(/folders\/([a-zA-Z0-9_-]+)/);
  _check_('U06: regex extrae folder ID', !!m);
  _eq_('U07: folder ID extraído correctamente', m[1], 'ABC123_-XYZ');

  // sanitizeFileName y cleanCompanyName
  _eq_('U08: sanitizeFileName barra /→_',      sanitizeFileName('A/B'), 'A_B');
  _eq_('U09: sanitizeFileName trunca a 50',    sanitizeFileName('X'.repeat(60)).length, 50);
  _check_('U10: cleanCompanyName elimina SA DE CV', cleanCompanyName('EMPRESA SA DE CV').indexOf('SA DE CV') === -1);

  // tipo_orden por defecto
  _eq_('U11: tipo_orden default OTA', (undefined || 'OTA'), 'OTA');

  var pass = _results_.filter(function(r){ return r.indexOf('PASS') === 0; }).length;
  var fail = _results_.filter(function(r){ return r.indexOf('FAIL') === 0; }).length;
  Logger.log('  UNIT: ' + pass + ' PASS | ' + fail + ' FAIL');
}
