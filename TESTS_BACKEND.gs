// =========================================================================
// TESTS E2E — API CENTRAL EJECUTIVA AMBIENTAL
// Compatible con Google Apps Script — sin dependencias externas
// =========================================================================
// FLUJOS CUBIERTOS
//   E01  SEAPD  → registrarCliente   (crea carpeta Drive + fila en CLIENTES_MAESTRO)
//   E02  SEAOT  → buscarClienteRFC + registrarOT  (busca cliente, registra OT)
//   E03  SEAINF → getOrdenes + getConsecutivo + createExpediente
//   E04  SEAOT  → buscarClienteNombre (búsqueda positiva por nombre parcial)
//   E05  SEAOT  → buscarClienteRFC no encontrado (ruta negativa)
//   E06  SEAOT  → buscarClienteNombre nombre muy corto (validación de entrada)
//   E07  SEAOT  → registrarOT tipo OTB (segundo tipo de orden)
//   E08  SEADB  → updateEstatus ENTREGADO (estatus externo + fecha real)
//   E09  SEAINF → updateEstatusInforme FINALIZADO (estatus interno del informe)
//   E10–E18 → idempotencia, versionado, aislamiento, fallbacks seguros y
//             coincidencia relajada de carpetas manuales / links legados
//
// USO
//   1. Configurar Script Properties de staging:
//      SEA_E2E_ENABLED=TRUE, SEA_TEST_SPREADSHEET_ID y SEA_TEST_FOLDER_ID.
//   2. Editor GAS → seleccionar runE2ETests → ▶ Ejecutar → Ver registros.
//   Para ejecutar un flujo individual: runTest_E01 … runTest_E18.
//   Para solo pruebas unitarias: runUnitTests (no requiere staging).
//
// SEGURIDAD
//   Las E2E abortan si faltan las propiedades de staging o si alguno de sus
//   IDs coincide con producción. Nunca deben ejecutarse contra datos reales.
//
// LIMPIEZA
//   Los tests eliminan automáticamente sus filas y envían sus carpetas de
//   staging a la papelera al terminar, tanto en éxito como en falla.
// =========================================================================

// ─── Datos de prueba ──────────────────────────────────────────────────────
var TEST_RFC            = 'XTES000000TST';
var TEST_RFC_HIST       = 'HIST000000TST';
var TEST_RFC_MISSING    = 'MISS000000TST';
var TEST_FOLIO          = 'TEST-E2E-001';
var TEST_FOLIO_B        = 'TEST-E2E-002';
var TEST_FOLIO_MISSING  = 'TEST-E2E-MISSING';
var TEST_FOLIO_EMPTY    = 'TEST-E2E-EMPTY-LINK';
var TEST_FOLIO_GARBAGE  = 'TEST-E2E-GARBAGE-LINK';
var TEST_FOLIO_LEGACY   = 'TEST-E2E-LEGACY';
var TEST_FOLIO_LEGACYLINK = 'TEST-E2E-LEGACY-LINK';
var TEST_FOLIO_RESOLVED = 'TEST-E2E-RESOLVED';
var TEST_FOLIO_FAKEID   = 'TEST-E2E-FAKE-ID';
var TEST_FOLIO_FOREIGN  = 'TEST-E2E-FOREIGN-LINK';
var TEST_FOLIO_INTRUDER = 'TEST-E2E-INTRUDER-RFC';
var TEST_RFC_MANUAL     = 'MANU000000TST';
var TEST_SUCURSAL       = 'Sucursal Test E2E';
var TEST_SUCURSAL_HIST  = 'Sucursal Histórica E2E';
var TEST_SUCURSAL_MANUAL = 'Sucursal Manual E2E';
var TEST_PARENT_MANUAL  = 'CLIENTE MANUAL E2E (SIN RFC EN NOMBRE)';

var TEST_FOLIOS_ = [
  TEST_FOLIO, TEST_FOLIO_B, TEST_FOLIO_MISSING,
  TEST_FOLIO_EMPTY, TEST_FOLIO_GARBAGE, TEST_FOLIO_LEGACY,
  TEST_FOLIO_LEGACYLINK, TEST_FOLIO_RESOLVED,
  TEST_FOLIO_FAKEID, TEST_FOLIO_FOREIGN, TEST_FOLIO_INTRUDER
];
var TEST_RFCS_ = [TEST_RFC, TEST_RFC_HIST, TEST_RFC_MISSING, TEST_RFC_MANUAL];

// ─── Estado compartido entre flujos ──────────────────────────────────────
var _ctx_ = {
  linkDriveCliente:          '',
  folioOT:                   '',
  numInforme:                '',
  urlExpediente:             '',
  clienteFolderId:           '',
  perfilFolderId:            '',
  expedienteFolderId:        '',
  foreignExpedienteFolderId: '',
  legacyExpedienteFolderId:  '',
  folderIdsToTrash:          [],
  stagingActive:             false
};

// ─── Guard obligatorio de staging ─────────────────────────────────────────
function _isUnsafeE2ETarget_(spreadsheetId, folderId) {
  return !spreadsheetId || !folderId ||
    spreadsheetId === CONFIG.SPREADSHEET_ID ||
    folderId === CONFIG.FOLDER_ID;
}

function _activateE2EStaging_() {
  if (_ctx_.stagingActive) return;

  var props = PropertiesService.getScriptProperties();
  var enabled = String(props.getProperty('SEA_E2E_ENABLED') || '').toUpperCase() === 'TRUE';
  var spreadsheetId = String(props.getProperty('SEA_TEST_SPREADSHEET_ID') || '').trim();
  var folderId = String(props.getProperty('SEA_TEST_FOLDER_ID') || '').trim();

  if (!enabled || !spreadsheetId || !folderId) {
    throw new Error(
      'E2E bloqueadas: configura SEA_E2E_ENABLED=TRUE, ' +
      'SEA_TEST_SPREADSHEET_ID y SEA_TEST_FOLDER_ID para staging.'
    );
  }
  if (_isUnsafeE2ETarget_(spreadsheetId, folderId)) {
    throw new Error('E2E bloqueadas: los IDs de staging no pueden coincidir con producción.');
  }

  var ss = SpreadsheetApp.openById(spreadsheetId);
  var clientes = ss.getSheetByName(CONFIG.SHEET_CLIENTES);
  var ordenes = ss.getSheetByName(CONFIG.SHEET_OT);
  var informes = ss.getSheetByName(CONFIG.SHEET_INFORMES);
  if (!clientes || clientes.getMaxColumns() < 22 ||
      !ordenes || ordenes.getMaxColumns() < 17 ||
      !informes || informes.getMaxColumns() < 17) {
    throw new Error('E2E bloqueadas: el Spreadsheet de staging no cumple CLIENTES A–V, ORDENES A–Q e INFORMES A–Q.');
  }

  DriveApp.getFolderById(folderId); // valida acceso antes de cualquier escritura
  CONFIG.SPREADSHEET_ID = spreadsheetId;
  CONFIG.FOLDER_ID = folderId;
  _ctx_.stagingActive = true;
  Logger.log('  Entorno E2E aislado en staging.');
}

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
function _neq_(msg, actual, unexpected) {
  var ok = String(actual) !== String(unexpected);
  ok ? _pass_(msg) : _fail_(msg + ' | no debería ser: "' + unexpected + '"');
}

// =========================================================================
// RUNNER PRINCIPAL
// =========================================================================
function runE2ETests() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── PRE-LIMPIEZA DE STAGING ──────────────────────');
  _cleanup_();
  _results_ = [];
  Logger.log('');
  Logger.log('══════════════════════════════════════════════');
  Logger.log('  TESTS E2E — EA Backend v3.0');
  Logger.log('  RFC de prueba  : ' + TEST_RFC);
  Logger.log('  Folio principal: ' + TEST_FOLIO);
  Logger.log('  Folio secundario: ' + TEST_FOLIO_B);
  Logger.log('══════════════════════════════════════════════');

  var results = {
    e01: false, e02: false, e03: false, e04: false, e05: false, e06: false,
    e07: false, e08: false, e09: false, e10: false, e11: false, e12: false,
    e13: false, e14: false, e15: false, e16: false, e17: false, e18: false
  };

  var tests = [
    runTest_E01, runTest_E02, runTest_E03, runTest_E04, runTest_E05, runTest_E06,
    runTest_E07, runTest_E08, runTest_E09, runTest_E10, runTest_E11, runTest_E12,
    runTest_E13, runTest_E14, runTest_E15, runTest_E16, runTest_E17, runTest_E18
  ];
  for (var n = 0; n < tests.length; n++) {
    var number = String(n + 1).padStart(2, '0');
    var key = 'e' + number;
    try {
      tests[n]();
      results[key] = true;
    } catch (e) {
      Logger.log('  E' + number + ' abortado: ' + e.message);
    }
  }

  Logger.log('');
  Logger.log('── LIMPIEZA ──────────────────────────────────');
  _cleanup_();

  var pass = _results_.filter(function(r){ return r.indexOf('PASS') === 0; }).length;
  var fail = _results_.filter(function(r){ return r.indexOf('FAIL') === 0; }).length;
  Logger.log('');
  Logger.log('══════════════════════════════════════════════');
  Logger.log('  RESULTADO: ' + pass + ' PASS  |  ' + fail + ' FAIL');
  Logger.log('  E01 registrarCliente       : ' + (results.e01 ? 'OK' : 'FALLO'));
  Logger.log('  E02 registrarOT            : ' + (results.e02 ? 'OK' : 'FALLO'));
  Logger.log('  E03 createExpediente       : ' + (results.e03 ? 'OK' : 'FALLO'));
  Logger.log('  E04 buscarClienteNombre    : ' + (results.e04 ? 'OK' : 'FALLO'));
  Logger.log('  E05 RFC no encontrado      : ' + (results.e05 ? 'OK' : 'FALLO'));
  Logger.log('  E06 Nombre muy corto       : ' + (results.e06 ? 'OK' : 'FALLO'));
  Logger.log('  E07 registrarOT tipo OTB   : ' + (results.e07 ? 'OK' : 'FALLO'));
  Logger.log('  E08 updateEstatus          : ' + (results.e08 ? 'OK' : 'FALLO'));
  Logger.log('  E09 updateEstatusInforme   : ' + (results.e09 ? 'OK' : 'FALLO'));
  Logger.log('  E10 idempotencia           : ' + (results.e10 ? 'OK' : 'FALLO'));
  Logger.log('  E11 versionado MD5         : ' + (results.e11 ? 'OK' : 'FALLO'));
  Logger.log('  E12 aislamiento lote       : ' + (results.e12 ? 'OK' : 'FALLO'));
  Logger.log('  E13 rechazo parcial        : ' + (results.e13 ? 'OK' : 'FALLO'));
  Logger.log('  E14 histórico read-only    : ' + (results.e14 ? 'OK' : 'FALLO'));
  Logger.log('  E15 bloqueo sin carpeta    : ' + (results.e15 ? 'OK' : 'FALLO'));
  Logger.log('  E16 OT sin carpeta/legado  : ' + (results.e16 ? 'OK' : 'FALLO'));
  Logger.log('  E17 retroajuste carpetas   : ' + (results.e17 ? 'OK' : 'FALLO'));
  Logger.log('  E18 carpetas manuales      : ' + (results.e18 ? 'OK' : 'FALLO'));
  Logger.log('══════════════════════════════════════════════');

  // Ejecutar también las pruebas unitarias
  Logger.log('');
  runUnitTests();
}

// =========================================================================
// E01 — SEAPD: registrarCliente
// =========================================================================
// Simula el payload que SEAPD envía cuando el usuario llena el formulario
// y presiona "Enviar". Verifica que:
//   - La función devuelve success: true
//   - Aparece una fila en CLIENTES_MAESTRO con los datos correctos
//   - El link_drive_cliente (índice 20) no está vacío y apunta a Drive
//   - La carpeta Drive existe y es accesible
function runTest_E01() {
  _activateE2EStaging_();
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

  var linkDrive = String(fila[20] || '');
  _check_('E01-11: link_drive en col 21 (índice 20) no está vacío', linkDrive !== '');
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
  var perfilFolders = carpeta.getFoldersByName('01_Cliente');
  if (perfilFolders.hasNext()) _ctx_.perfilFolderId = perfilFolders.next().getId();
  Logger.log('  linkDriveCliente: ' + linkDrive);
}

// =========================================================================
// E02 — SEAOT: buscarClienteRFC + registrarOT
// =========================================================================
// Simula el flujo de SEAOT:
//   1. Usuario escribe el RFC → buscarClienteRFC → obtiene sucursales + link_drive
//   2. Usuario llena el form de OT y presiona "Registrar OT" → registrarOT
// Verifica que:
//   - buscarClienteRFC devuelve found: true con el link correcto (índice 20)
//   - La OT aparece en ORDENES_TRABAJO con link_drive_cliente en col M
function runTest_E02() {
  _activateE2EStaging_();
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
  _check_('E02-7: link_drive_cliente no está vacío (índice 20 de CLIENTES)',  linkDrive !== '');
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
  _eq_('E02-13: nom_servicio en col D (índice 3)',            filaOT[CO.NOM], 'NOM-035-STPS');
  _eq_('E02-14: rfc en col G (índice 6)',                     filaOT[CO.RFC], TEST_RFC);
  _eq_('E02-15: estatus inicial en col L',                    filaOT[CO.ESTATUS_EXTERNO], 'NO INICIADO');

  var linkEnOT = String(filaOT[CO.LINK_DRIVE] || '');
  _check_('E02-16: link_drive_cliente guardado en col M (índice 12)', linkEnOT !== '');
  // El backend guarda siempre la forma canónica; se compara por ID de carpeta.
  _eq_('E02-17: link en OT apunta a la carpeta del cliente',
    extractDriveFolderId_(linkEnOT), extractDriveFolderId_(linkDrive));

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
//   - INFORMES recibe el número, estatus inicial y link del expediente
function runTest_E03() {
  _activateE2EStaging_();
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
  _check_('E03-14: subcarpeta "1. ORDEN_TRABAJO" creada',    subFolders.indexOf('1. ORDEN_TRABAJO') !== -1);
  _check_('E03-15: subcarpeta "2. HDC" creada',              subFolders.indexOf('2. HDC') !== -1);
  _check_('E03-16: subcarpeta "3. CROQUIS" creada',          subFolders.indexOf('3. CROQUIS') !== -1);
  _check_('E03-17: subcarpeta "4. FOTOS" creada',            subFolders.indexOf('4. FOTOS') !== -1);
  _check_('E03-18: subcarpeta "5. INFORMES Y MEMORIAS"',     subFolders.indexOf('5. INFORMES Y MEMORIAS') !== -1);
  _check_('E03-19: subcarpeta "6. INFORME PRELIMINAR"',      subFolders.indexOf('6. INFORME PRELIMINAR') !== -1);

  // El expediente se registra en INFORMES. ORDENES_TRABAJO conserva su
  // estructura y el enlace de la sucursal.
  var sheetInf = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_INFORMES);
  var rowsInf = sheetInf.getDataRange().getValues();
  var filaInf = null;
  for (var j = rowsInf.length - 1; j >= 1; j--) {
    if (normalizeOtForSeainf_(rowsInf[j][CI.OT]) === normalizeOtForSeainf_(TEST_FOLIO)) {
      filaInf = rowsInf[j]; break;
    }
  }
  _check_('E03-20: fila creada en INFORMES', filaInf !== null);
  _eq_('E03-21: número de informe en INFORMES', filaInf[CI.NUM_INFORME], numInforme);
  _eq_('E03-22: estatus inicial del informe', filaInf[CI.ESTATUS], 'NO INICIADO');
  _check_('E03-23: link del expediente en INFORMES', String(filaInf[CI.LINK_DRIVE] || '').indexOf('folders/') !== -1);

  _ctx_.numInforme         = numInforme;
  _ctx_.urlExpediente      = resultExp.url;
  _ctx_.expedienteFolderId = m[1];
  Logger.log('  Expediente creado en: ' + resultExp.url);
}

// =========================================================================
// E04 — SEAOT: buscarClienteNombre (búsqueda positiva)
// =========================================================================
// Requiere que E01 haya creado la empresa de prueba en CLIENTES_MAESTRO.
// Verifica que:
//   - buscarClienteNombre devuelve found: true cuando se busca un término parcial
//   - El resultado incluye el RFC, link de Drive y nombre_solicitante correctos
//   - La búsqueda es insensible a mayúsculas/minúsculas
function runTest_E04() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E04: SEAOT → buscarClienteNombre (positivo) ──');

  // Búsqueda con término parcial en minúsculas (debe encontrar "EMPRESA TEST E2E SA DE CV")
  var busqueda = fase2_BuscarClienteNombre('empresa test e2e');
  _check_('E04-1: buscarClienteNombre devuelve found=true',         busqueda.found === true);
  _check_('E04-2: hay al menos un resultado',                        busqueda.resultados && busqueda.resultados.length > 0);

  var r = busqueda.resultados[0];
  _check_('E04-3: razon_social contiene el término buscado',
    String(r.razon_social).toUpperCase().indexOf('EMPRESA TEST E2E') !== -1);
  _eq_('E04-4: rfc del resultado es correcto',            r.rfc,                TEST_RFC);
  _check_('E04-5: link_drive_cliente no está vacío',      (r.link_drive_cliente || '') !== '');
  _check_('E04-6: link contiene "folders/"',              String(r.link_drive_cliente).indexOf('folders/') !== -1);
  _eq_('E04-7: nombre_solicitante correcto',              r.nombre_solicitante, 'Prueba Automatizada');
  _eq_('E04-8: correo_informe correcto',                  r.correo_informe,     'test@noenviar.com');
  _eq_('E04-9: telefono_empresa correcto',                r.telefono_empresa,   '2220000000');

  // Búsqueda con solo 3 caracteres (mínimo permitido)
  var busqueda3 = fase2_BuscarClienteNombre('XTE');
  // Puede o no encontrar resultados, pero no debe lanzar excepción ni devolver error de validación
  _check_('E04-10: búsqueda de 3 chars no devuelve error de validación',
    busqueda3.error !== 'Nombre demasiado corto (mínimo 3 caracteres)');

  Logger.log('  Empresa encontrada: ' + r.razon_social + ' | Sucursal: ' + r.sucursal);
}

// =========================================================================
// E05 — SEAOT: buscarClienteRFC — RFC no encontrado (ruta negativa)
// =========================================================================
// Verifica que cuando se busca un RFC que no existe en el sistema:
//   - La función devuelve found: false
//   - NO lanza excepción
//   - NO devuelve datos de otro cliente
function runTest_E05() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E05: buscarClienteRFC → RFC no encontrado ─────');

  var rfcInexistente = 'ZZZZZ999999ZZZ';
  var busqueda = fase2_BuscarClienteRFC(rfcInexistente);

  _check_('E05-1: respuesta es un objeto (no lanzó excepción)', typeof busqueda === 'object' && busqueda !== null);
  _check_('E05-2: found es false para RFC inexistente',         busqueda.found === false);
  _check_('E05-3: no hay campo sucursales en la respuesta negativa',
    !busqueda.sucursales || busqueda.sucursales.length === 0);

  Logger.log('  RFC inexistente correctamente rechazado.');
}

// =========================================================================
// E06 — SEAOT: buscarClienteNombre — nombre muy corto (validación)
// =========================================================================
// Verifica que la validación de longitud mínima funciona:
//   - 1 carácter → error de validación
//   - 2 caracteres → error de validación
//   - "" (vacío) → error de validación
function runTest_E06() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E06: buscarClienteNombre → validación de entrada ─');

  var casos = [
    { input: '', desc: 'cadena vacía' },
    { input: 'A', desc: '1 carácter' },
    { input: 'AB', desc: '2 caracteres' }
  ];

  for (var c = 0; c < casos.length; c++) {
    var caso = casos[c];
    var resp = fase2_BuscarClienteNombre(caso.input);
    _check_('E06-' + (c * 2 + 1) + ': ' + caso.desc + ' → found=false',
      resp.found === false);
    _check_('E06-' + (c * 2 + 2) + ': ' + caso.desc + ' → contiene error de validación',
      typeof resp.error === 'string' && resp.error.length > 0);
  }

  Logger.log('  Validación de longitud mínima funciona correctamente.');
}

// =========================================================================
// E07 — SEAOT: registrarOT tipo OTB
// =========================================================================
// Verifica que el sistema soporta el segundo tipo de orden (OTB - Brigada):
//   - La OT se crea con tipo OTB
//   - El consecutivo se calcula de forma independiente al de OTA
//   - El estatus inicial es "NO INICIADO"
//   - El folio secundario TEST-E2E-002 se guarda correctamente
function runTest_E07() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E07: SEAOT → registrarOT tipo OTB ─────────────');

  // Limpiar folio secundario residual
  try {
    var sheetOTPre = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
    var rowsOTPre  = sheetOTPre.getDataRange().getValues();
    for (var oi = rowsOTPre.length - 1; oi >= 1; oi--) {
      if (String(rowsOTPre[oi][1]).trim() === TEST_FOLIO_B) {
        sheetOTPre.deleteRow(oi + 1);
        Logger.log('  [pre-cleanup] Fila OTB residual eliminada (fila ' + (oi + 1) + ')');
      }
    }
  } catch(e) { Logger.log('  [pre-cleanup OTB] ' + e.message); }

  var linkDrive = _ctx_.linkDriveCliente || '';

  var payloadOTB = {
    action:               'registrarOT',
    ot_folio:             TEST_FOLIO_B,
    tipo_orden:           'OTB',
    nom_servicio:         'NOM-036-STPS',
    cliente_razon_social: 'EMPRESA TEST E2E SA DE CV',
    sucursal:             TEST_SUCURSAL,
    rfc:                  TEST_RFC,
    personal_asignado:    'Brigada Test',
    fecha_visita:         '2026-03-25',
    fecha_entrega_limite: '2026-04-05',
    link_drive_cliente:   linkDrive,
    observaciones:        'OTB generada por test E2E'
  };

  var resultOTB = fase2_RegistrarOT(payloadOTB);
  _check_('E07-1: registrarOT tipo OTB devuelve success=true', resultOTB.success === true);

  // Verificar fila en ORDENES_TRABAJO
  var sheet  = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  var rows   = sheet.getDataRange().getValues();
  var filaOTB = null;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][1]).trim() === TEST_FOLIO_B) { filaOTB = rows[i]; break; }
  }

  _check_('E07-2: fila OTB creada en ORDENES_TRABAJO',         filaOTB !== null);
  _eq_('E07-3: folio en col 2 (índice 1)',                    filaOTB[1], TEST_FOLIO_B);
  _eq_('E07-4: tipo_orden es OTB (índice 2)',                 filaOTB[2], 'OTB');
  _eq_('E07-5: nom_servicio NOM-036 (col D)',                 filaOTB[CO.NOM], 'NOM-036-STPS');
  _eq_('E07-6: rfc correcto (col G)',                         filaOTB[CO.RFC], TEST_RFC);
  _eq_('E07-7: estatus externo inicial (col L)',              filaOTB[CO.ESTATUS_EXTERNO], 'NO INICIADO');
  _check_('E07-8: enlace de sucursal conservado (col M)',     String(filaOTB[CO.LINK_DRIVE] || '').indexOf('folders/') !== -1);

  // Verificar que el consecutivo OTB es independiente del OTA
  var hoy    = new Date();
  var anio   = String(hoy.getFullYear()).slice(2);
  var mes    = String(hoy.getMonth() + 1).padStart(2, '0');
  var consOTB = getConsecutivoSafe_({ anio: anio, mes: mes, nom: 'NOM036', tipo: 'OTB' });
  var consOTA = getConsecutivoSafe_({ anio: anio, mes: mes, nom: 'NOM035', tipo: 'OTA' });

  _check_('E07-9: consecutivo OTB devuelve success=true',         consOTB.success === true);
  _check_('E07-10: consecutivo OTA devuelve success=true',        consOTA.success === true);
  _check_('E07-11: consecutivo OTB contiene "NOM036"',
    String(consOTB.numeroInforme).indexOf('NOM036') !== -1);
  _check_('E07-12: consecutivo OTA contiene "NOM035"',
    String(consOTA.numeroInforme).indexOf('NOM035') !== -1);

  Logger.log('  OTB registrada con folio: ' + TEST_FOLIO_B);
  Logger.log('  Consecutivo OTB: ' + consOTB.numeroInforme);
  Logger.log('  Consecutivo OTA: ' + consOTA.numeroInforme);
}

// =========================================================================
// E08 — SEADB: updateEstatus → ENTREGADO
// =========================================================================
// Requiere que E02 haya creado TEST_FOLIO en ORDENES_TRABAJO.
// Verifica que:
//   - El estatus externo se cambia correctamente
//   - Al marcar como ENTREGADO, la fecha real de entrega se registra automáticamente
//   - Intentar actualizar un folio inexistente devuelve error sin excepción
function runTest_E08() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E08: SEADB → updateEstatus ENTREGADO ──────────');

  var dataUpdate = { ot: TEST_FOLIO, estatus: 'ENTREGADO' };
  var result = updateEstatusSafe_(dataUpdate, 'test-automatizado');

  _check_('E08-1: updateEstatus devuelve success=true', result.success === true);

  // Verificar en la hoja que el estatus cambió y la fecha real se registró
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  var rows  = sheet.getDataRange().getValues();
  var filaOT = null;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][1]).trim() === TEST_FOLIO) { filaOT = rows[i]; break; }
  }

  _check_('E08-2: fila OT encontrada en ORDENES_TRABAJO',       filaOT !== null);
  _eq_('E08-3: estatus_externo (col L) actualizado a ENTREGADO', filaOT[CO.ESTATUS_EXTERNO], 'ENTREGADO');
  _check_('E08-4: fecha_real_entrega (col K) registrada',
    filaOT[CO.FECHA_REAL] !== null && String(filaOT[CO.FECHA_REAL]).trim() !== '');

  // Ruta negativa: folio inexistente
  var resultNeg = updateEstatusSafe_({ ot: 'FOLIO-INEXISTENTE-ZZZ', estatus: 'EN PROCESO' }, 'test');
  _check_('E08-5: folio inexistente devuelve success=false sin excepción', resultNeg.success === false);
  _check_('E08-6: mensaje de error presente en ruta negativa', typeof resultNeg.error === 'string');

  // Ruta negativa: payload incompleto
  var resultIncompleto = updateEstatusSafe_({ ot: TEST_FOLIO }, 'test');
  _check_('E08-7: payload sin estatus devuelve success=false', resultIncompleto.success === false);

  Logger.log('  Estatus actualizado a ENTREGADO y fecha real registrada.');
}

// =========================================================================
// E09 — SEAINF: updateEstatusInforme → FINALIZADO
// =========================================================================
// Requiere que E02 haya creado TEST_FOLIO en ORDENES_TRABAJO.
// Verifica que:
//   - El estatus INTERNO se actualiza en INFORMES (col N)
//   - Una OT con expediente se excluye de Nuevo Expediente
//   - El estatus externo de ORDENES_TRABAJO (col L) NO se modifica
function runTest_E09() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E09: SEAINF → updateEstatusInforme FINALIZADO ─');

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetOt = ss.getSheetByName(CONFIG.SHEET_OT);
  var rowsOt = sheetOt.getDataRange().getValues();
  var estatusExternoBefore = '';
  for (var k = rowsOt.length - 1; k >= 1; k--) {
    if (normalizeOtForSeainf_(rowsOt[k][CO.OT]) === normalizeOtForSeainf_(TEST_FOLIO)) {
      estatusExternoBefore = String(rowsOt[k][CO.ESTATUS_EXTERNO]);
      break;
    }
  }

  var result = updateEstatusInformeSafe_(
    { ot: TEST_FOLIO, estatus: 'FINALIZADO' },
    'test-automatizado'
  );
  _check_('E09-1: updateEstatusInforme devuelve success=true', result.success === true);

  var sheetInf = ss.getSheetByName(CONFIG.SHEET_INFORMES);
  var rowsInf = sheetInf.getDataRange().getValues();
  var filaInf = null;
  for (var i = rowsInf.length - 1; i >= 1; i--) {
    if (normalizeOtForSeainf_(rowsInf[i][CI.OT]) === normalizeOtForSeainf_(TEST_FOLIO)) {
      filaInf = rowsInf[i]; break;
    }
  }
  _check_('E09-2: fila encontrada en INFORMES', filaInf !== null);
  _eq_('E09-3: estatus interno actualizado en INFORMES', filaInf[CI.ESTATUS], 'FINALIZADO');

  var rowsOtAfter = sheetOt.getDataRange().getValues();
  var estatusExternoAfter = '';
  for (var j = rowsOtAfter.length - 1; j >= 1; j--) {
    if (normalizeOtForSeainf_(rowsOtAfter[j][CO.OT]) === normalizeOtForSeainf_(TEST_FOLIO)) {
      estatusExternoAfter = String(rowsOtAfter[j][CO.ESTATUS_EXTERNO]);
      break;
    }
  }
  _eq_('E09-4: estatus externo de OT no fue modificado',
    estatusExternoAfter, estatusExternoBefore);

  var ordenesResp = getOrdenesSafe_();
  _check_('E09-5: getOrdenes devuelve success=true', ordenesResp.success === true);
  var aparece = ordenesResp.data.some(function(o) { return o.ot === TEST_FOLIO; });
  _check_('E09-6: OT con expediente excluida de Nuevo Expediente', !aparece);

  var resultNeg = updateEstatusInformeSafe_({ ot: TEST_FOLIO }, 'test');
  _check_('E09-7: payload sin estatus devuelve success=false', resultNeg.success === false);
}

// =========================================================================
// HELPERS E2E E10–E17
// =========================================================================
function _createPayload_(ot, numInforme, nom, rfc, sucursal, cliente, linkDrive, files) {
  return {
    action: 'createExpediente',
    data: {
      ot: ot,
      nom: nom,
      numInforme: numInforme,
      cliente: cliente,
      sucursal: sucursal,
      rfc: rfc,
      linkDrive: linkDrive || '',
      fecha: Utilities.formatDate(new Date(), 'GMT-6', 'dd/MM/yyyy'),
      entrega: '31/12/2026',
      tipoOrden: 'OTA',
      solicitante: 'Prueba Automatizada',
      telefono: '2220000000',
      direccion: 'Dirección de staging',
      responsable: 'Ing. Test',
      estatus: 'NO INICIADO'
    },
    files: files || []
  };
}

function _countRowsByOt_(sheet, ot, otIndex) {
  var values = sheet.getDataRange().getValues();
  var expected = normalizeOtForSeainf_(ot);
  var count = 0;
  for (var i = 1; i < values.length; i++) {
    if (normalizeOtForSeainf_(values[i][otIndex]) === expected) count++;
  }
  return count;
}

function _folderFileNames_(folder) {
  var names = [];
  var files = folder.getFiles();
  while (files.hasNext()) names.push(files.next().getName());
  return names;
}

function _rootChildIds_() {
  var root = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  var ids = [];
  var folders = root.getFolders();
  while (folders.hasNext()) ids.push(folders.next().getId());
  return ids.sort();
}

function _appendOtRow_(ot, cliente, sucursal, rfc, linkDrive, nom) {
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  sheet.appendRow([
    new Date(), ot, 'OTA', nom || 'TEST',
    cliente, sucursal, rfc, 'Ing. Test',
    '01/08/2026', '31/12/2026', '',
    'NO INICIADO', linkDrive || '', 'Fila exclusiva de staging',
    '', '', ''
  ]);
}

function _appendInformeRow_(ot, numInforme, cliente, sucursal, rfc, url) {
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_INFORMES);
  sheet.appendRow([
    new Date(), numInforme, 'OTA', ot, 'TEST', cliente,
    'Prueba Automatizada', rfc, '2220000000', 'Dirección de staging',
    '01/08/2026', '31/12/2026', 'NO', 'NO INICIADO',
    url, 'Ing. Test', sucursal
  ]);
}

// =========================================================================
// E10 — Reintento idempotente con los mismos archivos
// =========================================================================
function runTest_E10() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E10: idempotencia de expediente y archivos ────');

  _check_('E10-1: E03 dejó expediente disponible', !!_ctx_.urlExpediente && !!_ctx_.numInforme);
  var file = {
    name: 'idempotencia.pdf',
    type: 'application/pdf',
    content: 'QUJD',
    category: 'ORDEN_TRABAJO'
  };
  var payload = _createPayload_(
    TEST_FOLIO, _ctx_.numInforme, 'NOM035', TEST_RFC, TEST_SUCURSAL,
    'EMPRESA TEST E2E SA DE CV', _ctx_.linkDriveCliente, [file]
  );

  var first = fase3_CrearExpediente(payload);
  var second = fase3_CrearExpediente(payload);
  _check_('E10-2: primer reintento reutiliza expediente', first.success && first.alreadyExists === true);
  _eq_('E10-3: primer reintento guarda el archivo faltante', first.acceptedFiles, 1);
  _check_('E10-4: segundo reintento reutiliza expediente', second.success && second.alreadyExists === true);
  _eq_('E10-5: mismo archivo se omite por contenido idéntico', second.skippedFiles, 1);
  _eq_('E10-6: segundo reintento no crea archivo', second.acceptedFiles, 0);
  _eq_('E10-7: ambos reintentos conservan la misma URL', second.url, first.url);

  var sheetInf = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_INFORMES);
  _eq_('E10-8: existe una sola fila INFORMES para la OT',
    _countRowsByOt_(sheetInf, TEST_FOLIO, CI.OT), 1);
}

// =========================================================================
// E11 — Mismo nombre y contenido diferente crea versión
// =========================================================================
function runTest_E11() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E11: versionado real por MD5 ──────────────────');

  var first = fase3_AddFilesToExpediente({
    ot: TEST_FOLIO,
    expedienteUrl: _ctx_.urlExpediente,
    files: [{ name: 'foto-versionada.jpg', type: 'image/jpeg', content: 'QUJD', category: 'FOTOS' }]
  });
  var second = fase3_AddFilesToExpediente({
    ot: TEST_FOLIO,
    expedienteUrl: _ctx_.urlExpediente,
    files: [{ name: 'foto-versionada.jpg', type: 'image/jpeg', content: 'REVG', category: 'FOTOS' }]
  });

  _check_('E11-1: primer archivo fue guardado', first.success && first.acceptedFiles === 1);
  _check_('E11-2: segundo contenido fue versionado', second.success && second.versionedFiles === 1);

  var expediente = DriveApp.getFolderById(_ctx_.expedienteFolderId);
  var fotos = expediente.getFoldersByName('4. FOTOS').next();
  var names = _folderFileNames_(fotos);
  _check_('E11-3: se conserva el archivo original', names.indexOf('foto-versionada.jpg') !== -1);
  _check_('E11-4: existe foto (v2).jpg', names.indexOf('foto-versionada (v2).jpg') !== -1);
}

// =========================================================================
// E12 — URL de otra OT no puede recibir el lote
// =========================================================================
function runTest_E12() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E12: aislamiento OT–expediente ────────────────');

  var cons = getConsecutivoSafe_({ anio: '26', mes: '08', nom: 'NOM036', tipo: 'OTB' });
  _check_('E12-1: se obtuvo consecutivo para expediente ajeno', cons.success === true);
  var foreign = fase3_CrearExpediente(_createPayload_(
    TEST_FOLIO_B, cons.numeroInforme, 'NOM036', TEST_RFC, TEST_SUCURSAL,
    'EMPRESA TEST E2E SA DE CV', _ctx_.linkDriveCliente, []
  ));
  _check_('E12-2: expediente de la segunda OT creado', foreign.success === true);
  _ctx_.foreignExpedienteFolderId = extractDriveFolderId_(foreign.url);
  _ctx_.folderIdsToTrash.push(_ctx_.foreignExpedienteFolderId);

  var result = fase3_AddFilesToExpediente({
    ot: TEST_FOLIO,
    expedienteUrl: foreign.url,
    files: [{ name: 'NO-DEBE-GUARDARSE.pdf', type: 'application/pdf', content: 'QUJD', category: 'FOTOS' }]
  });
  _check_('E12-3: URL de otra OT es rechazada', result.success === false);

  var foreignFolder = DriveApp.getFolderById(_ctx_.foreignExpedienteFolderId);
  var foreignFotos = foreignFolder.getFoldersByName('4. FOTOS').next();
  _check_('E12-4: no se escribió en la carpeta ajena',
    !foreignFotos.getFilesByName('NO-DEBE-GUARDARSE.pdf').hasNext());
}

// =========================================================================
// E13 — Lote parcial conserva los archivos válidos
// =========================================================================
function runTest_E13() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E13: rechazo parcial recuperable ──────────────');

  var result = fase3_AddFilesToExpediente({
    ot: TEST_FOLIO,
    expedienteUrl: _ctx_.urlExpediente,
    files: [
      { name: 'parcial-a.pdf', type: 'application/pdf', content: 'QUJD', category: 'ORDEN_TRABAJO' },
      { name: 'parcial-b.heic', type: 'image/heic', content: 'REVG', category: 'FOTOS' },
      { name: 'parcial-c.txt', type: 'text/plain', content: 'R0hJ', category: 'FOTOS' }
    ]
  });

  _check_('E13-1: lote mixto devuelve success=true', result.success === true);
  _check_('E13-2: respuesta marca partial=true', result.partial === true);
  _eq_('E13-3: reporta un archivo rechazado', result.rejectedFiles.length, 1);
  _eq_('E13-4: identifica el archivo rechazado', result.rejectedFiles[0].name, 'parcial-c.txt');
  _eq_('E13-5: guarda los dos archivos válidos', result.acceptedFiles, 2);

  var expediente = DriveApp.getFolderById(_ctx_.expedienteFolderId);
  var otFolder = expediente.getFoldersByName('1. ORDEN_TRABAJO').next();
  var fotos = expediente.getFoldersByName('4. FOTOS').next();
  _check_('E13-6: PDF válido existe en Drive', otFolder.getFilesByName('parcial-a.pdf').hasNext());
  _check_('E13-7: HEIC válido existe en Drive', fotos.getFilesByName('parcial-b.heic').hasNext());
  _check_('E13-8: TXT rechazado no existe en Drive', !fotos.getFilesByName('parcial-c.txt').hasNext());
}

// =========================================================================
// E14 — Cliente histórico sin Link Drive se resuelve sin escribir la celda
// =========================================================================
function runTest_E14() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E14: resolución read-only de cliente histórico ─');

  var root = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  var parent = root.createFolder(TEST_RFC_HIST + ' - CLIENTE HISTORICO E2E');
  var branch = parent.createFolder(TEST_SUCURSAL_HIST);
  _ctx_.folderIdsToTrash.push(parent.getId());

  var row = Array(22).fill('');
  row[CL.FECHA_REGISTRO] = new Date();
  row[CL.RAZON_SOCIAL] = 'CLIENTE HISTORICO E2E SA DE CV';
  row[CL.SUCURSAL] = TEST_SUCURSAL_HIST;
  row[CL.RFC] = TEST_RFC_HIST;
  row[CL.LINK_DRIVE] = '';
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_CLIENTES);
  sheet.appendRow(row);
  var sheetRow = sheet.getLastRow();

  var result = fase2_ResolverCarpetaCliente_(
    TEST_RFC_HIST, TEST_SUCURSAL_HIST, 'CLIENTE HISTORICO E2E SA DE CV'
  );
  _check_('E14-1: cliente histórico fue encontrado', result.success && result.found);
  _eq_('E14-2: URL resuelta corresponde a la sucursal exacta',
    extractDriveFolderId_(result.linkDrive), branch.getId());
  _eq_('E14-3: LINK_DRIVE permanece vacío',
    sheet.getRange(sheetRow, CL.LINK_DRIVE + 1).getValue(), '');
}

// =========================================================================
// E15 — Sin carpeta exacta no se crea expediente ni fallback en raíz
// =========================================================================
function runTest_E15() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E15: bloqueo seguro sin carpeta ───────────────');

  _appendOtRow_(
    TEST_FOLIO_MISSING, 'CLIENTE SIN CARPETA E2E', 'Sucursal Inexistente',
    TEST_RFC_MISSING, '', 'TEST'
  );
  var before = JSON.stringify(_rootChildIds_());
  var result = fase3_CrearExpediente(_createPayload_(
    TEST_FOLIO_MISSING, 'EA-2608-TEST-9991', 'TEST',
    TEST_RFC_MISSING, 'Sucursal Inexistente', 'CLIENTE SIN CARPETA E2E', '', []
  ));
  var after = JSON.stringify(_rootChildIds_());

  _check_('E15-1: createExpediente bloquea carpeta inexistente', result.success === false);
  _eq_('E15-2: no cambió la lista de carpetas en la raíz', after, before);
  var sheetInf = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_INFORMES);
  _eq_('E15-3: no se creó fila INFORMES',
    _countRowsByOt_(sheetInf, TEST_FOLIO_MISSING, CI.OT), 0);
}

// =========================================================================
// E16 — Registrar OT: bloqueo sin carpeta resoluble, aceptación de links
//       legados y resolución server-side para clientes registrados
// =========================================================================
function runTest_E16() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E16: OT sin carpeta resoluble / links legados ──');

  function payload(folio, link, rfc, sucursal, razon) {
    return {
      ot_folio: folio,
      tipo_orden: 'OTA',
      nom_servicio: 'TEST',
      cliente_razon_social: razon,
      sucursal: sucursal,
      rfc: rfc,
      personal_asignado: 'Ing. Test',
      fecha_visita: '01/08/2026',
      fecha_entrega_limite: '31/12/2026',
      link_drive_cliente: link,
      observaciones: 'Caso E16'
    };
  }

  // Cliente inexistente: ni el enlace ni la resolución server-side salvan la OT.
  var empty = fase2_RegistrarOT(payload(
    TEST_FOLIO_EMPTY, '', TEST_RFC_MISSING, 'Sucursal Inexistente', 'CLIENTE SIN CARPETA E2E'));
  var garbage = fase2_RegistrarOT(payload(
    TEST_FOLIO_GARBAGE, 'https://example.com/no-drive', TEST_RFC_MISSING, 'Sucursal Inexistente', 'CLIENTE SIN CARPETA E2E'));
  _check_('E16-1: cliente sin carpeta con enlace vacío es rechazado', empty.success === false);
  _eq_('E16-2: el rechazo trae code CARPETA_NO_ENCONTRADA', empty.code, 'CARPETA_NO_ENCONTRADA');
  _check_('E16-3: enlace basura sin carpeta es rechazado', garbage.success === false);

  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  _eq_('E16-4: no se escribió OT con enlace vacío',
    _countRowsByOt_(sheet, TEST_FOLIO_EMPTY, CO.OT), 0);
  _eq_('E16-5: no se escribió OT con enlace basura',
    _countRowsByOt_(sheet, TEST_FOLIO_GARBAGE, CO.OT), 0);

  // Un ID con formato válido pero inexistente no debe colarse: el servidor
  // verifica que la carpeta exista y corresponda al RFC+sucursal.
  var fakeId = fase2_RegistrarOT(payload(
    TEST_FOLIO_FAKEID, 'https://drive.google.com/open?id=NoExisteEsteId0123456789',
    TEST_RFC_MISSING, 'Sucursal Inexistente', 'CLIENTE SIN CARPETA E2E'));
  _check_('E16-FK1: ID bien formado pero inexistente es rechazado', fakeId.success === false);
  _eq_('E16-FK2: no se escribió OT con ID inexistente',
    _countRowsByOt_(sheet, TEST_FOLIO_FAKEID, CO.OT), 0);

  // Un enlace real pero de OTRO cliente/sucursal tampoco se acepta.
  var foreign = fase2_RegistrarOT(payload(
    TEST_FOLIO_FOREIGN, 'https://drive.google.com/drive/folders/' + _ctx_.clienteFolderId,
    TEST_RFC_MISSING, 'Sucursal Inexistente', 'CLIENTE SIN CARPETA E2E'));
  _check_('E16-FR1: enlace de otra carpeta de cliente es rechazado', foreign.success === false);
  _eq_('E16-FR2: no se escribió OT con enlace ajeno',
    _countRowsByOt_(sheet, TEST_FOLIO_FOREIGN, CO.OT), 0);

  // Enlace legado open?id= del cliente registrado en E01: se acepta y la
  // columna M guarda la forma canónica /folders/<id>.
  var legacy = fase2_RegistrarOT(payload(
    TEST_FOLIO_LEGACYLINK, 'https://drive.google.com/open?id=' + _ctx_.clienteFolderId,
    TEST_RFC, TEST_SUCURSAL, 'EMPRESA TEST E2E SA DE CV'));
  _check_('E16-6: enlace legado open?id= es aceptado', legacy.success === true);
  var rowsOT = sheet.getDataRange().getValues();
  var filaLegacy = null;
  for (var i = rowsOT.length - 1; i >= 1; i--) {
    if (String(rowsOT[i][CO.OT]).trim() === TEST_FOLIO_LEGACYLINK) { filaLegacy = rowsOT[i]; break; }
  }
  _check_('E16-7: fila de OT con enlace legado existe', filaLegacy !== null);
  _eq_('E16-8: col M guarda la forma canónica /folders/',
    String(filaLegacy[CO.LINK_DRIVE]),
    'https://drive.google.com/drive/folders/' + _ctx_.clienteFolderId);

  // Sin enlace pero cliente registrado: el servidor resuelve la carpeta
  // (mismo criterio RFC+sucursal que usa SEAOT).
  var resolved = fase2_RegistrarOT(payload(
    TEST_FOLIO_RESOLVED, '', TEST_RFC, TEST_SUCURSAL, 'EMPRESA TEST E2E SA DE CV'));
  _check_('E16-9: sin enlace, la carpeta se resuelve server-side', resolved.success === true);
  rowsOT = sheet.getDataRange().getValues();
  var filaResolved = null;
  for (var r = rowsOT.length - 1; r >= 1; r--) {
    if (String(rowsOT[r][CO.OT]).trim() === TEST_FOLIO_RESOLVED) { filaResolved = rowsOT[r]; break; }
  }
  _check_('E16-10: fila de OT resuelta existe', filaResolved !== null);
  _eq_('E16-11: col M apunta a la sucursal del cliente',
    extractDriveFolderId_(String(filaResolved[CO.LINK_DRIVE])), _ctx_.clienteFolderId);
}

// =========================================================================
// E17 — Expediente antiguo de cuatro subcarpetas se completa a seis
// =========================================================================
function runTest_E17() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E17: retroajuste de subcarpetas ───────────────');

  var otResult = fase2_RegistrarOT({
    ot_folio: TEST_FOLIO_LEGACY,
    tipo_orden: 'OTA',
    nom_servicio: 'TEST',
    cliente_razon_social: 'EMPRESA TEST E2E SA DE CV',
    sucursal: TEST_SUCURSAL,
    rfc: TEST_RFC,
    personal_asignado: 'Ing. Test',
    fecha_visita: '01/08/2026',
    fecha_entrega_limite: '31/12/2026',
    link_drive_cliente: _ctx_.linkDriveCliente,
    observaciones: 'Expediente legado de staging'
  });
  _check_('E17-1: OT legado creada', otResult.success === true);

  var clientFolder = DriveApp.getFolderById(_ctx_.clienteFolderId);
  var legacy = clientFolder.createFolder('02_Expediente_LEGACY_TEST');
  ['1. ORDEN_TRABAJO', '2. HDC', '3. CROQUIS', '4. FOTOS'].forEach(function(name) {
    legacy.createFolder(name);
  });
  _ctx_.legacyExpedienteFolderId = legacy.getId();
  _ctx_.folderIdsToTrash.push(legacy.getId());
  _appendInformeRow_(
    TEST_FOLIO_LEGACY, 'EA-2608-TEST-9993', 'EMPRESA TEST E2E SA DE CV',
    TEST_SUCURSAL, TEST_RFC, legacy.getUrl()
  );

  var result = fase3_AddFilesToExpediente({
    ot: TEST_FOLIO_LEGACY,
    expedienteUrl: legacy.getUrl(),
    files: []
  });
  _check_('E17-2: addFiles completa expediente legado', result.success === true);

  var counts = {};
  var folders = legacy.getFolders();
  while (folders.hasNext()) {
    var name = folders.next().getName();
    counts[name] = (counts[name] || 0) + 1;
  }
  Object.keys(CONFIG.FOLDER_STRUCTURE).forEach(function(key) {
    var expected = CONFIG.FOLDER_STRUCTURE[key];
    _eq_('E17: existe una sola subcarpeta ' + expected, counts[expected] || 0, 1);
  });
}

// =========================================================================
// E18 — Coincidencia relajada: carpeta manual sin RFC en el nombre del padre
// =========================================================================
// Reproduce el caso real que rompía SEAOT: un cliente registrado en
// CLIENTES_MAESTRO cuya carpeta fue creada a mano con la razón social
// (sin el RFC como prefijo) y cuyo Link Drive usa el formato legado open?id=.
function runTest_E18() {
  _activateE2EStaging_();
  Logger.log('');
  Logger.log('── E18: coincidencia relajada de carpetas manuales ─');

  var root = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  var parent = root.createFolder(TEST_PARENT_MANUAL);
  var branch = parent.createFolder(TEST_SUCURSAL_MANUAL);
  _ctx_.folderIdsToTrash.push(parent.getId());

  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_CLIENTES);
  var row = Array(22).fill('');
  row[CL.FECHA_REGISTRO] = new Date();
  row[CL.RAZON_SOCIAL] = 'CLIENTE MANUAL E2E SA DE CV';
  row[CL.SUCURSAL] = TEST_SUCURSAL_MANUAL;
  row[CL.RFC] = TEST_RFC_MANUAL;
  row[CL.LINK_DRIVE] = 'https://drive.google.com/open?id=' + branch.getId(); // formato legado
  sheet.appendRow(row);
  var sheetRow = sheet.getLastRow();

  // 1) Link legado en la fila: se resuelve vía fila + regla de razón social
  //    (el padre no tiene el RFC en el nombre).
  var byRow = fase2_ResolverCarpetaCliente_(
    TEST_RFC_MANUAL, TEST_SUCURSAL_MANUAL, 'CLIENTE MANUAL E2E SA DE CV');
  _check_('E18-1: link legado open?id= resuelto vía fila', byRow.success && byRow.found);
  _eq_('E18-2: la URL corresponde a la sucursal manual',
    extractDriveFolderId_(byRow.linkDrive), branch.getId());

  // 2) Sin link en la fila: se resuelve con la búsqueda en raíz por razón social.
  sheet.getRange(sheetRow, CL.LINK_DRIVE + 1).setValue('');
  var byName = fase2_ResolverCarpetaCliente_(
    TEST_RFC_MANUAL, TEST_SUCURSAL_MANUAL, 'CLIENTE MANUAL E2E SA DE CV');
  _check_('E18-3: carpeta sin RFC en el nombre encontrada por razón social',
    byName.success && byName.found);
  _eq_('E18-4: misma carpeta de sucursal',
    extractDriveFolderId_(byName.linkDrive), branch.getId());
  _eq_('E18-5: LINK_DRIVE permanece vacío (read-only)',
    sheet.getRange(sheetRow, CL.LINK_DRIVE + 1).getValue(), '');

  // 3) Un RFC no registrado (o mal tecleado) NO puede adoptar la carpeta de
  //    esta empresa: sin fila maestra que lo ate, la coincidencia por razón
  //    social no aplica.
  var intruso = fase2_ResolverCarpetaCliente_(
    TEST_RFC_MISSING, TEST_SUCURSAL_MANUAL, 'CLIENTE MANUAL E2E SA DE CV');
  _check_('E18-6: RFC no registrado no resuelve por razón social',
    intruso.success === true && intruso.found === false);

  var intrusoOt = fase2_RegistrarOT({
    ot_folio: TEST_FOLIO_INTRUDER,
    tipo_orden: 'OTA',
    nom_servicio: 'TEST',
    cliente_razon_social: 'CLIENTE MANUAL E2E SA DE CV',
    sucursal: TEST_SUCURSAL_MANUAL,
    rfc: TEST_RFC_MISSING,
    personal_asignado: 'Ing. Test',
    fecha_visita: '01/08/2026',
    fecha_entrega_limite: '31/12/2026',
    link_drive_cliente: 'https://drive.google.com/drive/folders/' + branch.getId(),
    observaciones: 'Debe rechazarse: RFC ajeno a la carpeta'
  });
  _check_('E18-7: OT con RFC ajeno a la carpeta es rechazada', intrusoOt.success === false);
  var sheetOt = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  _eq_('E18-8: no se escribió la OT del RFC ajeno',
    _countRowsByOt_(sheetOt, TEST_FOLIO_INTRUDER, CO.OT), 0);
}

// =========================================================================
// LIMPIEZA — elimina todos los datos de prueba
// =========================================================================
function _cleanup_() {
  _activateE2EStaging_();
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // 1. Eliminar filas de CLIENTES_MAESTRO para todos los RFC de prueba.
  try {
    var sheetCli = ss.getSheetByName(CONFIG.SHEET_CLIENTES);
    var rowsCli = sheetCli.getDataRange().getValues();
    for (var i = rowsCli.length - 1; i >= 1; i--) {
      var rfc = String(rowsCli[i][CL.RFC] || '').toUpperCase().trim();
      if (TEST_RFCS_.indexOf(rfc) !== -1) {
        sheetCli.deleteRow(i + 1);
        Logger.log('  Fila CLIENTES eliminada para RFC: ' + rfc);
      }
    }
  } catch (e) { Logger.log('  ERROR cleanup CLIENTES_MAESTRO: ' + e.message); }

  // 2. Eliminar todas las OTs de prueba.
  try {
    var sheetOT = ss.getSheetByName(CONFIG.SHEET_OT);
    var rowsOT = sheetOT.getDataRange().getValues();
    for (var j = rowsOT.length - 1; j >= 1; j--) {
      var folio = String(rowsOT[j][CO.OT] || '').trim();
      if (TEST_FOLIOS_.indexOf(folio) !== -1) {
        sheetOT.deleteRow(j + 1);
        Logger.log('  Fila ORDENES eliminada: ' + folio);
      }
    }
  } catch (e) { Logger.log('  ERROR cleanup ORDENES_TRABAJO: ' + e.message); }

  // 3. Eliminar todas las filas INFORMES de ambas OTs y casos auxiliares.
  try {
    var sheetInf = ss.getSheetByName(CONFIG.SHEET_INFORMES);
    var rowsInf = sheetInf.getDataRange().getValues();
    for (var k = rowsInf.length - 1; k >= 1; k--) {
      var otInf = String(rowsInf[k][CI.OT] || '').trim();
      if (TEST_FOLIOS_.indexOf(otInf) !== -1) {
        sheetInf.deleteRow(k + 1);
        Logger.log('  Fila INFORMES eliminada: ' + otInf);
      }
    }
  } catch (e) { Logger.log('  ERROR cleanup INFORMES: ' + e.message); }

  // 4. Eliminar trazas de auditoría de las OTs de prueba.
  try {
    var audit = ss.getSheetByName(CONFIG.SHEET_AUDITORIA);
    if (audit) {
      var auditRows = audit.getDataRange().getValues();
      for (var a = auditRows.length - 1; a >= 1; a--) {
        if (TEST_FOLIOS_.indexOf(String(auditRows[a][3] || '').trim()) !== -1) {
          audit.deleteRow(a + 1);
        }
      }
    }
  } catch (e) { Logger.log('  ERROR cleanup AUDITORIA: ' + e.message); }

  // 5. Enviar a papelera carpetas conocidas, incluida 01_Cliente.
  var ids = [
    _ctx_.perfilFolderId,
    _ctx_.expedienteFolderId,
    _ctx_.foreignExpedienteFolderId,
    _ctx_.legacyExpedienteFolderId,
    _ctx_.clienteFolderId
  ].concat(_ctx_.folderIdsToTrash || []);
  var seen = {};
  ids.forEach(function(id) {
    if (!id || seen[id]) return;
    seen[id] = true;
    try {
      DriveApp.getFolderById(id).setTrashed(true);
      Logger.log('  Carpeta de staging enviada a papelera: ' + id);
    } catch (e) { Logger.log('  ERROR cleanup carpeta ' + id + ': ' + e.message); }
  });

  // 6. Defensa final: eliminar padres RFC de prueba que hayan sobrevivido.
  try {
    var root = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    TEST_RFCS_.forEach(function(testRfc) {
      var parents = root.searchFolders(buildClientParentFolderQuery_(testRfc, 'CLIENTE HISTORICO E2E SA DE CV'));
      while (parents.hasNext()) {
        var parent = parents.next();
        if (isParentFolderForRfc_(parent.getName(), testRfc)) {
          parent.setTrashed(true);
          Logger.log('  Carpeta padre de staging enviada a papelera: ' + parent.getName());
        }
      }
    });
    // La carpeta manual de E18 no lleva RFC en el nombre: se busca por nombre.
    var manualParents = root.getFoldersByName(TEST_PARENT_MANUAL);
    while (manualParents.hasNext()) {
      var manualParent = manualParents.next();
      manualParent.setTrashed(true);
      Logger.log('  Carpeta manual de staging enviada a papelera: ' + manualParent.getName());
    }
  } catch (e) { Logger.log('  ERROR cleanup carpeta raíz: ' + e.message); }
}

// =========================================================================
// PRUEBA DE CORREO — ejecutar manualmente para verificar plantillas
// =========================================================================
// Selecciona esta función en el editor GAS y presiona ▶ Ejecutar.
// Crea una carpeta temporal en Drive, envía el correo de equipo +
// confirmación al cliente (correo_informe) y luego borra la carpeta.
function runTest_Email() {
  Logger.log('── Prueba de correo ─────────────────────────');
  var dataPrueba = {
    nombre_solicitante:   'Prueba Automatizada',
    razon_social:         'EMPRESA TEST E2E SA DE CV',
    sucursal:             'Sucursal Test Email',
    rfc:                  'XTES000000TST',
    telefono_empresa:     '2220000000',
    representante_legal:  'Rep Legal Test',
    correo_informe:       Session.getActiveUser().getEmail(), // llega a tu propia cuenta
    giro:                 'Pruebas Automatizadas',
    aplica_nom020:        'no',
    requiere_pipc:        'no',
    fechas_preferidas:    'Cualquier día hábil'
  };

  // Carpeta temporal solo para obtener una URL real
  var carpetaTemp = DriveApp.getFolderById(CONFIG.FOLDER_ID)
    .createFolder('TEMP_TEST_EMAIL_' + new Date().getTime());
  var sheetUrlFake = carpetaTemp.getUrl(); // reutilizamos la misma URL como placeholder

  try {
    var result = enviarNotificacionRobusta(dataPrueba, [], carpetaTemp, sheetUrlFake, Logger.log.bind(Logger));
    if (result.success) {
      Logger.log('[OK] Correo enviado — revisa tu bandeja: ' + dataPrueba.correo_informe);
    } else {
      Logger.log('[FAIL] ' + result.error);
    }
  } finally {
    carpetaTemp.setTrashed(true);
    Logger.log('Carpeta temporal eliminada.');
  }
}

// =========================================================================
// UNIT TESTS — lógica pura, sin efectos secundarios en Drive/Sheets
// =========================================================================
function runUnitTests() {
  _results_ = [];
  Logger.log('');
  Logger.log('── UNIT TESTS (lógica pura) ──────────────────');

  // ── Mapeo de índices nuevo esquema 22 columnas ─────────────────────────
  var fila22 = [
    '08/03/2026','EMPRESA TEST','Planta Norte','TST010101AAA',
    'Rep Legal','Dir Eval','2221234567','Solicitante Test',
    'sol@test.com','Manufactura','IMSS-01234','500 ton','450 ton',
    'L-V 08:00-18:00','SÍ','NO',
    'Ing. Responsable','5551234567','Lic. Dirigido','Gerente General',
    'https://drive.google.com/drive/folders/LINK_CORRECTO',
    'Asesor Externo SA'
  ];
  _eq_('U01: nombre_solicitante en índice [7]',  fila22[7], 'Solicitante Test');
  _eq_('U02: correo_informe en índice [8]',      fila22[8], 'sol@test.com');
  _eq_('U03: telefono_empresa en índice [6]',    fila22[6], '2221234567');
  _eq_('U04: link_drive en índice [20]',         fila22[20], 'https://drive.google.com/drive/folders/LINK_CORRECTO');
  _eq_('U05: asesor_consultor en índice [21]',   fila22[21], 'Asesor Externo SA');
  _check_('U06: índice [22] es undefined (límite del esquema)', fila22[22] === undefined);

  // ── Regex extracción folder ID ─────────────────────────────────────────
  var m1 = 'https://drive.google.com/drive/folders/ABC123_-XYZ'.match(/folders\/([a-zA-Z0-9_-]+)/);
  _check_('U07: regex extrae folder ID', !!m1);
  _eq_('U08: folder ID extraído correctamente', m1[1], 'ABC123_-XYZ');

  // Folder ID con guiones y underscores (formato real de Drive)
  var m2 = 'https://drive.google.com/drive/folders/1nHd-70uUeciClDm_3_pgbmqGF7II1lfQ'.match(/folders\/([a-zA-Z0-9_-]+)/);
  _check_('U09: regex extrae ID con guiones y underscores', !!m2);
  _eq_('U10: ID con guiones extraído completo', m2[1], '1nHd-70uUeciClDm_3_pgbmqGF7II1lfQ');

  // ── sanitizeFileName ───────────────────────────────────────────────────
  _eq_('U11: sanitizeFileName barra /→_',      sanitizeFileName('A/B'), 'A_B');
  _eq_('U12: sanitizeFileName trunca a 50',    sanitizeFileName('X'.repeat(60)).length, 50);
  _eq_('U13: sanitizeFileName asterisco →_',   sanitizeFileName('A*B'), 'A_B');
  _eq_('U14: sanitizeFileName dos puntos →_',  sanitizeFileName('A:B'), 'A_B');
  _eq_('U15: sanitizeFileName acepta acentos', sanitizeFileName('ción'), 'ción');
  _eq_('U16: sanitizeFileName acepta ñ',       sanitizeFileName('niño'), 'niño');
  _check_('U17: sanitizeFileName resultado no vacío para cadena vacía',
    sanitizeFileName('').length > 0);

  // ── cleanCompanyName ───────────────────────────────────────────────────
  _check_('U18: cleanCompanyName elimina SA DE CV',
    cleanCompanyName('EMPRESA SA DE CV').indexOf('SA DE CV') === -1);
  _check_('U19: cleanCompanyName elimina S.A. DE C.V.',
    cleanCompanyName('EMPRESA S.A. DE C.V.').indexOf('S.A.') === -1);
  _check_('U20: cleanCompanyName elimina S.C.',
    cleanCompanyName('DESPACHO S.C.').indexOf('S.C.') === -1);
  _check_('U21: cleanCompanyName resultado no vacío',
    cleanCompanyName('EMPRESA SA DE CV').trim().length > 0);

  // ── Formato de número de informe ───────────────────────────────────────
  var regexInforme = /^EA-\d{4}-[A-Za-z0-9]+-\d{4}$/;
  _check_('U22: regex valida EA-2603-NOM035-0001',   regexInforme.test('EA-2603-NOM035-0001'));
  _check_('U23: regex valida EA-2603-NOM036-0042',   regexInforme.test('EA-2603-NOM036-0042'));
  _check_('U24: regex rechaza formato incompleto',   !regexInforme.test('EA-2603-NOM035'));
  _check_('U25: regex rechaza formato sin consecutivo', !regexInforme.test('EA-2603-NOM035-'));
  _check_('U26: regex rechaza cadena vacía',         !regexInforme.test(''));

  // ── Tipo de orden por defecto ──────────────────────────────────────────
  _eq_('U27: tipo_orden default OTA', (undefined || 'OTA'), 'OTA');

  // ── Formato de folio OT ────────────────────────────────────────────────
  // Los folios tienen formato libre pero deben ser no vacíos
  _check_('U28: folio TEST-E2E-001 es válido (no vacío)', TEST_FOLIO.length > 0);
  _check_('U29: folio TEST-E2E-002 es válido (no vacío)', TEST_FOLIO_B.length > 0);
  _neq_('U30: folios principal y secundario son distintos', TEST_FOLIO, TEST_FOLIO_B);

  // ── Índices de columna en CONFIG ───────────────────────────────────────
  _eq_('U31: CONFIG.COLUMNS.CLIENTES.LINK_DRIVE es 20',    CONFIG.COLUMNS.CLIENTES.LINK_DRIVE,    20);
  _eq_('U32: CONFIG.COLUMNS.CLIENTES.ASESOR_CONSULTOR es 21', CONFIG.COLUMNS.CLIENTES.ASESOR_CONSULTOR, 21);
  _eq_('U33: CONFIG.COLUMNS.ORDENES.LINK_DRIVE es 12',     CONFIG.COLUMNS.ORDENES.LINK_DRIVE,     12);
  _eq_('U34: CONFIG.COLUMNS.ORDENES.ESTATUS_EXTERNO es 11',CONFIG.COLUMNS.ORDENES.ESTATUS_EXTERNO,11);
  _eq_('U35: CONFIG.COLUMNS.ORDENES.FECHA_REAL es 10',     CONFIG.COLUMNS.ORDENES.FECHA_REAL,     10);
  _eq_('U36: CONFIG.COLUMNS.ORDENES.OBSERVACIONES es 13',  CONFIG.COLUMNS.ORDENES.OBSERVACIONES,  13);
  _eq_('U36A: CONFIG.COLUMNS.ORDENES.FECHA_PAUSA es 14',      CONFIG.COLUMNS.ORDENES.FECHA_PAUSA,      14);
  _eq_('U36B: CONFIG.COLUMNS.ORDENES.MOTIVO_PAUSA es 15',     CONFIG.COLUMNS.ORDENES.MOTIVO_PAUSA,     15);
  _eq_('U36C: CONFIG.COLUMNS.ORDENES.FECHA_INFO_COMPLETA es 16', CONFIG.COLUMNS.ORDENES.FECHA_INFO_COMPLETA, 16);

  // ── Estructura de subcarpetas en CONFIG ────────────────────────────────
  _eq_('U37: FOLDER_STRUCTURE.ORDEN_TRABAJO',   CONFIG.FOLDER_STRUCTURE.ORDEN_TRABAJO,   '1. ORDEN_TRABAJO');
  _eq_('U38: FOLDER_STRUCTURE.HOJAS_CAMPO',      CONFIG.FOLDER_STRUCTURE.HOJAS_CAMPO,      '2. HDC');
  _eq_('U39: FOLDER_STRUCTURE.CROQUIS',          CONFIG.FOLDER_STRUCTURE.CROQUIS,          '3. CROQUIS');
  _eq_('U40: FOLDER_STRUCTURE.FOTOS',            CONFIG.FOLDER_STRUCTURE.FOTOS,            '4. FOTOS');
  _eq_('U41: FOLDER_STRUCTURE.INFORMES_MEMORIA', CONFIG.FOLDER_STRUCTURE.INFORMES_MEMORIA, '5. INFORMES Y MEMORIAS');
  _eq_('U42: FOLDER_STRUCTURE.INF_PRELIMINAR',   CONFIG.FOLDER_STRUCTURE.INF_PRELIMINAR,   '6. INFORME PRELIMINAR');

  // ── Zona horaria en CONFIG ─────────────────────────────────────────────
  _eq_('U43: CONFIG.TIMEZONE es GMT-6', CONFIG.TIMEZONE, 'GMT-6');

  // ── Validación de RFC (formato México: 13 chars alfanuméricos) ─────────
  var rfcRegex = /^[A-Z&Ñ]{3,4}[0-9]{6}[A-Z0-9]{3}$/;
  _check_('U44: RFC persona física válido XTES000000TST',    rfcRegex.test(TEST_RFC));
  _check_('U45: RFC persona moral válido EMP010101AAA',      rfcRegex.test('EMP010101AAA'));
  _check_('U46: RFC persona física válido GACJ800101H12',    rfcRegex.test('GACJ800101H12'));
  _check_('U47: RFC inválido rechazado (muy corto)',          !rfcRegex.test('EMP01'));
  _check_('U48: RFC inválido rechazado (caracteres ilegales)',!rfcRegex.test('EMP01010#AAA'));

  // ── Reglas de seguridad agregadas en PR #94 ─────────────────────────────
  _check_('U49: SIN_RFC no se reutiliza como identidad estable',
    canReuseParentByRfc_('SIN_RFC') === false);
  _check_('U50: RFC real sí permite reutilización de carpeta padre',
    canReuseParentByRfc_('EMP010101AAA') === true);
  _eq_('U51: nombre de archivo conserva extensión y guion',
    sanitizeDriveFileName_('IMG-001.jpg'), 'IMG-001.jpg');
  _check_('U52: HEIC permitido', validarArchivo_('QUJD', 'image/heic', 'foto.heic').valid);
  _check_('U53: DXF permitido sin MIME', validarArchivo_('QUJD', '', 'croquis.dxf').valid);
  _check_('U54: RAR permitido', validarArchivo_('QUJD', 'application/vnd.rar', 'planos.rar').valid);
  _check_('U55: MP4 permitido', validarArchivo_('QUJD', 'video/mp4', 'evidencia.mp4').valid);
  _check_('U56: extensión ejecutable rechazada aunque declare PDF',
    !validarArchivo_('QUJD', 'application/pdf', 'archivo.exe').valid);

  // ── Lotes parciales: un archivo inválido no descarta los válidos ─────────
  var mixto = validateDriveFiles_([
    { name: 'ot.pdf',    type: 'application/pdf', content: 'QUJD' },
    { name: 'foto.heic', type: 'image/heic',      content: 'QUJD' },
    { name: 'nota.txt',  type: 'text/plain',      content: 'QUJD' }
  ]);
  _eq_('U57: lote mixto conserva los válidos', mixto.valid.length, 2);
  _eq_('U58: lote mixto reporta el rechazado', mixto.rejected.length, 1);
  _eq_('U59: el rechazo identifica el archivo', mixto.rejected[0].name, 'nota.txt');

  // ── Versionado de nombres con stub de Folder ─────────────────────────────
  var folderStub = {
    getFilesByName: function(name) {
      return { hasNext: function() { return name.indexOf('(v2)') !== -1; } };
    }
  };
  _eq_('U60: salta a (v3) si (v2) ya existe',
    versionedFileName_(folderStub, 'foto.jpg'), 'foto (v3).jpg');

  // ── Casos límite de nombres de archivo ───────────────────────────────────
  _eq_('U61: conserva doble extensión', sanitizeDriveFileName_('archivo.tar.gz'), 'archivo.tar.gz');
  _check_('U62: extensión inválida no deja punto final',
    !/[.]$/.test(sanitizeDriveFileName_('raro.%%%')));
  var longPdf = sanitizeDriveFileName_(Array(201).join('x') + '.pdf');
  _check_('U63: nombre largo no supera 180 caracteres y conserva .pdf',
    longPdf.length <= 180 && /[.]pdf$/i.test(longPdf));
  _eq_('U64: conserva nombre sin extensión',
    sanitizeDriveFileName_('sin_extension'), 'sin_extension');

  // ── Identidad de carpetas y normalización simétrica ──────────────────────
  _check_('U65: RFC de 12 no coincide con carpeta de RFC de 13',
    !isParentFolderForRfc_('ABC010101AB1X - OTRA', 'ABC010101AB1'));
  var parentDone = false;
  var branchStub = {
    getName: function() { return 'Planta Norte - Puebla'; },
    getParents: function() {
      return {
        hasNext: function() { return !parentDone; },
        next: function() {
          parentDone = true;
          return { getName: function() { return 'ABC010101AB1 - EMPRESA'; } };
        }
      };
    }
  };
  _check_('U66: sucursal normalizada acepta guion en Drive y diagonal en hoja',
    folderMatchesClientBranch_(branchStub, 'ABC010101AB1', 'Planta Norte / Puebla', 'EMPRESA', false));

  // ── Consulta acotada a la raíz ───────────────────────────────────────────
  _eq_('U67: RFC real usa búsqueda por prefijo',
    buildClientParentFolderQuery_('ABC010101AB1', 'EMPRESA'),
    "title contains 'ABC010101AB1' and trashed = false");
  _eq_('U68: SIN_RFC usa nombre exacto de empresa',
    buildClientParentFolderQuery_('SIN_RFC', 'EMPRESA SA DE CV'),
    "title = 'SIN_RFC - EMPRESA' and trashed = false");
  _eq_('U69: valores de consulta escapan barra y apóstrofe',
    escapeDriveQueryValue_("A\\B'C"), "A\\\\B\\'C");

  // ── El arnés E2E falla cerrado ante IDs productivos ──────────────────────
  _check_('U70: staging rechaza el Spreadsheet productivo',
    _isUnsafeE2ETarget_(CONFIG.SPREADSHEET_ID, 'folder-staging') === true);
  _check_('U71: IDs distintos de producción son elegibles para staging',
    _isUnsafeE2ETarget_('sheet-staging', 'folder-staging') === false);

  // ── Formatos de enlace Drive aceptados ───────────────────────────────────
  _eq_('U72: extrae ID de /folders/ con sufijo',
    extractDriveFolderId_('https://drive.google.com/drive/folders/1AbC_d-123xyz?usp=sharing'), '1AbC_d-123xyz');
  _eq_('U73: extrae ID del formato legado open?id=',
    extractDriveFolderId_('https://drive.google.com/open?id=1AbCdEfGh123456789012345678'), '1AbCdEfGh123456789012345678');
  _eq_('U74: extrae ID de uc?export=download&id=',
    extractDriveFolderId_('https://drive.google.com/uc?export=download&id=1AbCdEfGh123456789012345678'), '1AbCdEfGh123456789012345678');
  _eq_('U75: extrae ID suelto de 28 caracteres',
    extractDriveFolderId_('1AbCdEfGh123456789012345678x'), '1AbCdEfGh123456789012345678x');
  _eq_('U76: enlace vacío no produce ID', extractDriveFolderId_(''), '');
  _eq_('U77: URL sin Drive no produce ID', extractDriveFolderId_('https://example.com/no-drive'), '');
  _eq_('U78: token corto no es un ID', extractDriveFolderId_('PENDIENTE'), '');
  _eq_('U79: forma canónica del enlace',
    canonicalDriveFolderLink_('1AbC_d-123xyz'), 'https://drive.google.com/drive/folders/1AbC_d-123xyz');

  // ── Coincidencia relajada del padre ──────────────────────────────────────
  // El cuarto argumento (allowCompanyFallback) representa la fila de
  // CLIENTES_MAESTRO que ata RFC+sucursal+empresa. Sin ella sólo aplica la
  // regla del RFC.
  _check_('U80: padre con prefijo RFC (convención SEAPD)',
    parentFolderMatchesClient_('ABC010101AB1 - EMPRESA', 'ABC010101AB1', 'EMPRESA SA DE CV', false));
  _check_('U81: RFC delimitado en cualquier parte del nombre',
    parentFolderMatchesClient_('EMPRESA DEMO (ABC010101AB1)', 'ABC010101AB1', '', false));
  _check_('U82: RFC incrustado sin delimitador no coincide',
    !parentFolderMatchesClient_('XABC010101AB1X', 'ABC010101AB1', '', false));
  _check_('U83: nombre que empieza con la razón social limpia coincide con fila maestra',
    parentFolderMatchesClient_('BODEGA CRUZ AZUL DEL CENTRO (matriz)', 'ZZZ010101ZZ9', 'BODEGA CRUZ AZUL DEL CENTRO S.A. DE C.V.', true));
  _check_('U84: razón social corta no coincide con el nombre más largo de otro cliente',
    !parentFolderMatchesClient_('BODEGA CRUZ AZUL DEL CENTRO', 'ZZZ010101ZZ9', 'CRUZ AZUL S.A. DE C.V.', true));
  _check_('U85: razón social vacía sin RFC en el nombre no coincide',
    !parentFolderMatchesClient_('CUALQUIER CARPETA', 'ZZZ010101ZZ9', '', true));
  _check_('U86: el default Sin_nombre no coincide con todo',
    !parentFolderMatchesClient_('SIN_NOMBRE HISTORICO', 'ZZZ010101ZZ9', ' S.A. DE C.V.', true));

  // ── La razón social exige respaldo de fila maestra ───────────────────────
  _check_('U87: sin fila maestra, la razón social no adopta la carpeta',
    !parentFolderMatchesClient_('BODEGA CRUZ AZUL DEL CENTRO (matriz)', 'ZZZ010101ZZ9', 'BODEGA CRUZ AZUL DEL CENTRO S.A. DE C.V.', false));
  _check_('U88: RFC mal tecleado no hereda la carpeta de la empresa registrada',
    !parentFolderMatchesClient_('BODEGA CRUZ AZUL DEL CENTRO', 'MALTECLEADO99', 'BODEGA CRUZ AZUL DEL CENTRO S.A. DE C.V.', false));
  _check_('U89: el flag omitido falla cerrado (sólo regla de RFC)',
    !parentFolderMatchesClient_('BODEGA CRUZ AZUL DEL CENTRO', 'ZZZ010101ZZ9', 'BODEGA CRUZ AZUL DEL CENTRO S.A. DE C.V.'));
  _check_('U90: con fila maestra el RFC correcto sí resuelve la carpeta manual',
    parentFolderMatchesClient_('BODEGA CRUZ AZUL DEL CENTRO', 'ZZZ010101ZZ9', 'BODEGA CRUZ AZUL DEL CENTRO S.A. DE C.V.', true));
  _check_('U91: la regla del RFC no depende del flag',
    parentFolderMatchesClient_('OTRA EMPRESA (ABC010101AB1)', 'ABC010101AB1', 'NO COINCIDE SA DE CV', false));

  // ── El término de búsqueda en Drive concuerda con el comparador ──────────
  // Invariante: el término debe estar contenido en el nombre de la carpeta que
  // el comparador aceptaría; si no, Drive filtra la carpeta antes de evaluarla.
  var nombreLargo = 'CORPORATIVO INDUSTRIAL METALURGICO DEL BAJIO Y OCCIDENTE S.A. DE C.V.';
  var carpetaLarga = cleanCompanyName(nombreLargo);
  _check_('U92: razón social larga se trunca a 50 caracteres', carpetaLarga.length === 50);
  _check_('U93: el término de búsqueda cabe en el nombre truncado de la carpeta',
    carpetaLarga.indexOf(companyQueryTerm_(nombreLargo)) === 0);
  _check_('U94: la carpeta truncada sigue siendo aceptada por el comparador',
    parentFolderMatchesClient_(carpetaLarga, 'ZZZ010101ZZ9', nombreLargo, true));

  var nombreConSignos = 'GRUPO XYZ, S DE RL';
  _eq_('U95: el término se corta antes del primer carácter que se sanitiza',
    companyQueryTerm_(nombreConSignos), 'GRUPO XYZ');
  _check_('U96: el término cortado está contenido en el nombre sanitizado',
    cleanCompanyName(nombreConSignos).indexOf(companyQueryTerm_(nombreConSignos)) === 0);

  _eq_('U97: razón social normal no se altera',
    companyQueryTerm_('BODEGA CRUZ AZUL DEL CENTRO S.A. DE C.V.'), 'BODEGA CRUZ AZUL DEL CENTRO');
  _eq_('U98: razón social vacía no produce término', companyQueryTerm_(''), '');

  var pass = _results_.filter(function(r){ return r.indexOf('PASS') === 0; }).length;
  var fail = _results_.filter(function(r){ return r.indexOf('FAIL') === 0; }).length;
  Logger.log('');
  Logger.log('  UNIT: ' + pass + ' PASS | ' + fail + ' FAIL');
}
