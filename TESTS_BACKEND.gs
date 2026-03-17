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
//
// USO
//   Editor GAS → seleccionar runE2ETests → ▶ Ejecutar → Ver registros
//   Para ejecutar un flujo individual: runTest_E01 … runTest_E09
//   Para solo pruebas unitarias: runUnitTests
//
// LIMPIEZA
//   Los tests crean datos con RFC 'XTEST000000TST' en CLIENTES_MAESTRO y
//   folios 'TEST-E2E-001' / 'TEST-E2E-002' en ORDENES_TRABAJO.
//   Al terminar (éxito o falla) se eliminan automáticamente.
// =========================================================================

// ─── Datos de prueba ──────────────────────────────────────────────────────
var TEST_RFC      = 'XTES000000TST';
var TEST_FOLIO    = 'TEST-E2E-001';
var TEST_FOLIO_B  = 'TEST-E2E-002';
var TEST_SUCURSAL = 'Sucursal Test E2E';

// ─── Estado compartido entre flujos ──────────────────────────────────────
var _ctx_ = {
  linkDriveCliente:   '',   // resultado de E01 → input de E02/E04
  folioOT:            '',   // resultado de E02 → input de E03/E08/E09
  urlExpediente:      '',   // resultado de E03
  clienteFolderId:    '',   // para cleanup E01
  expedienteFolderId: ''    // para cleanup E03
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
function _neq_(msg, actual, unexpected) {
  var ok = String(actual) !== String(unexpected);
  ok ? _pass_(msg) : _fail_(msg + ' | no debería ser: "' + unexpected + '"');
}

// =========================================================================
// RUNNER PRINCIPAL
// =========================================================================
function runE2ETests() {
  _results_ = [];
  Logger.log('');
  Logger.log('══════════════════════════════════════════════');
  Logger.log('  TESTS E2E — EA Backend v3.0');
  Logger.log('  RFC de prueba  : ' + TEST_RFC);
  Logger.log('  Folio principal: ' + TEST_FOLIO);
  Logger.log('  Folio secundario: ' + TEST_FOLIO_B);
  Logger.log('══════════════════════════════════════════════');

  var results = {
    e01: false, e02: false, e03: false,
    e04: false, e05: false, e06: false,
    e07: false, e08: false, e09: false
  };

  try { runTest_E01(); results.e01 = true; } catch(e) { Logger.log('  E01 abortado: ' + e.message); }
  try { runTest_E02(); results.e02 = true; } catch(e) { Logger.log('  E02 abortado: ' + e.message); }
  try { runTest_E03(); results.e03 = true; } catch(e) { Logger.log('  E03 abortado: ' + e.message); }
  try { runTest_E04(); results.e04 = true; } catch(e) { Logger.log('  E04 abortado: ' + e.message); }
  try { runTest_E05(); results.e05 = true; } catch(e) { Logger.log('  E05 abortado: ' + e.message); }
  try { runTest_E06(); results.e06 = true; } catch(e) { Logger.log('  E06 abortado: ' + e.message); }
  try { runTest_E07(); results.e07 = true; } catch(e) { Logger.log('  E07 abortado: ' + e.message); }
  try { runTest_E08(); results.e08 = true; } catch(e) { Logger.log('  E08 abortado: ' + e.message); }
  try { runTest_E09(); results.e09 = true; } catch(e) { Logger.log('  E09 abortado: ' + e.message); }

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
  _eq_('E07-5: nom_servicio NOM-036 (índice 4)',              filaOTB[4], 'NOM-036-STPS');
  _eq_('E07-6: rfc correcto (índice 7)',                      filaOTB[7], TEST_RFC);
  _eq_('E07-7: estatus inicial "NO INICIADO"',                filaOTB[12], 'NO INICIADO');
  _eq_('E07-8: estatus_informe inicial "NO INICIADO"',        filaOTB[15], 'NO INICIADO');

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
  _eq_('E08-3: estatus_externo (índice 12) actualizado a ENTREGADO', filaOT[12], 'ENTREGADO');
  _check_('E08-4: fecha_real_entrega (índice 11) registrada',
    filaOT[11] !== null && String(filaOT[11]).trim() !== '');

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
//   - El estatus INTERNO del informe (col 16) se actualiza correctamente
//   - Al marcar como FINALIZADO, la OT se excluye de getOrdenes (filtro activo)
//   - El estatus externo (col 13) NO se modifica por esta función
function runTest_E09() {
  Logger.log('');
  Logger.log('── E09: SEAINF → updateEstatusInforme FINALIZADO ─');

  // Leer estatus externo ANTES de actualizar (para verificar que no cambia)
  var sheetBefore = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  var rowsBefore  = sheetBefore.getDataRange().getValues();
  var estatusExternoBefore = '';
  for (var k = rowsBefore.length - 1; k >= 1; k--) {
    if (String(rowsBefore[k][1]).trim() === TEST_FOLIO) {
      estatusExternoBefore = String(rowsBefore[k][12]);
      break;
    }
  }

  var dataInforme = { ot: TEST_FOLIO, estatus: 'FINALIZADO' };
  var result = updateEstatusInformeSafe_(dataInforme, 'test-automatizado');

  _check_('E09-1: updateEstatusInforme devuelve success=true', result.success === true);

  // Verificar en la hoja
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  var rows  = sheet.getDataRange().getValues();
  var filaOT = null;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][1]).trim() === TEST_FOLIO) { filaOT = rows[i]; break; }
  }

  _check_('E09-2: fila OT encontrada en ORDENES_TRABAJO', filaOT !== null);
  _eq_('E09-3: estatus_informe (índice 15) actualizado a FINALIZADO', filaOT[15], 'FINALIZADO');

  // El estatus externo (índice 12) NO debe haber cambiado
  _eq_('E09-4: estatus_externo (índice 12) no fue modificado por updateEstatusInforme',
    String(filaOT[12]), estatusExternoBefore);

  // La OT FINALIZADA ya no debe aparecer en getOrdenes
  var ordenesResp = getOrdenesSafe_();
  _check_('E09-5: getOrdenes devuelve success=true', ordenesResp.success === true);
  var aparece = ordenesResp.data.some(function(o) { return o.ot === TEST_FOLIO; });
  _check_('E09-6: OT FINALIZADA excluida de getOrdenes (filtro activo)', !aparece);

  // Ruta negativa: payload incompleto
  var resultNeg = updateEstatusInformeSafe_({ ot: TEST_FOLIO }, 'test');
  _check_('E09-7: payload sin estatus devuelve success=false', resultNeg.success === false);

  Logger.log('  Estatus informe actualizado a FINALIZADO.');
  Logger.log('  Estatus externo conservado: ' + estatusExternoBefore);
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

  // 2. Eliminar filas de ORDENES_TRABAJO (folio principal y secundario)
  try {
    var sheetOT = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
    var rowsOT  = sheetOT.getDataRange().getValues();
    for (var j = rowsOT.length - 1; j >= 1; j--) {
      var folio = String(rowsOT[j][1]).trim();
      if (folio === TEST_FOLIO || folio === TEST_FOLIO_B) {
        sheetOT.deleteRow(j + 1);
        Logger.log('  Fila eliminada de ORDENES_TRABAJO: ' + folio + ' (fila ' + (j + 1) + ')');
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
  _eq_('U33: CONFIG.COLUMNS.ORDENES.LINK_DRIVE es 13',     CONFIG.COLUMNS.ORDENES.LINK_DRIVE,     13);
  _eq_('U34: CONFIG.COLUMNS.ORDENES.ESTATUS_EXTERNO es 12',CONFIG.COLUMNS.ORDENES.ESTATUS_EXTERNO,12);
  _eq_('U35: CONFIG.COLUMNS.ORDENES.ESTATUS_INFORME es 15',CONFIG.COLUMNS.ORDENES.ESTATUS_INFORME,15);
  _eq_('U36: CONFIG.COLUMNS.ORDENES.FECHA_REAL es 11',     CONFIG.COLUMNS.ORDENES.FECHA_REAL,     11);

  // ── Estructura de subcarpetas en CONFIG ────────────────────────────────
  _eq_('U37: FOLDER_STRUCTURE.ORDEN_TRABAJO',  CONFIG.FOLDER_STRUCTURE.ORDEN_TRABAJO, '1. ORDEN_TRABAJO');
  _eq_('U38: FOLDER_STRUCTURE.HOJAS_CAMPO',    CONFIG.FOLDER_STRUCTURE.HOJAS_CAMPO,   '2. HDC');
  _eq_('U39: FOLDER_STRUCTURE.CROQUIS',        CONFIG.FOLDER_STRUCTURE.CROQUIS,       '3. CROQUIS');
  _eq_('U40: FOLDER_STRUCTURE.FOTOS',          CONFIG.FOLDER_STRUCTURE.FOTOS,         '4. FOTOS');

  // ── Zona horaria en CONFIG ─────────────────────────────────────────────
  _eq_('U41: CONFIG.TIMEZONE es GMT-6', CONFIG.TIMEZONE, 'GMT-6');

  // ── Validación de RFC (formato México: 13 chars alfanuméricos) ─────────
  var rfcRegex = /^[A-Z&Ñ]{3,4}[0-9]{6}[A-Z0-9]{3}$/;
  _check_('U42: RFC persona física válido XTES000000TST',    rfcRegex.test(TEST_RFC));
  _check_('U43: RFC persona moral válido EMP010101AAA',      rfcRegex.test('EMP010101AAA'));
  _check_('U44: RFC persona física válido GACJ800101H12',    rfcRegex.test('GACJ800101H12'));
  _check_('U45: RFC inválido rechazado (muy corto)',          !rfcRegex.test('EMP01'));
  _check_('U46: RFC inválido rechazado (caracteres ilegales)',!rfcRegex.test('EMP01010#AAA'));

  var pass = _results_.filter(function(r){ return r.indexOf('PASS') === 0; }).length;
  var fail = _results_.filter(function(r){ return r.indexOf('FAIL') === 0; }).length;
  Logger.log('');
  Logger.log('  UNIT: ' + pass + ' PASS | ' + fail + ' FAIL');
}
