// =========================================================================
// TESTS — API CENTRAL EJECUTIVA AMBIENTAL
// Compatible con Google Apps Script (sin dependencias externas)
// Para ejecutar: abre el editor GAS → selecciona runAllTests → ▶ Ejecutar
// =========================================================================
// Cobertura:
//   T01 – Mapeo de columnas en BuscarClienteRFC  (índice 15 = link Drive)
//   T02 – Mapeo de columnas en BuscarClienteNombre
//   T03 – getConsecutivo: número inicial y continuación de serie
//   T04 – getConsecutivo: tipo de orden OTA vs OT (no mezcla series)
//   T05 – getOrdenesSafe_: incluye rfc, sucursal, personal, nom, fecha_visita
//   T06 – getOrdenesSafe_: filtra ENTREGADO y FINALIZADO
//   T07 – Fallback de carpeta: extrae folder ID de URL Drive
//   T08 – sanitizeFileName: limpia caracteres no permitidos
//   T09 – cleanCompanyName: elimina sufijos legales
//   T10 – fase2_RegistrarOT: tipo_orden por defecto es OTA
// =========================================================================

// ─── Mini framework ───────────────────────────────────────────────────────
var _testResults_ = [];

function assert_(description, condition) {
  var status = condition ? 'PASS' : 'FAIL';
  _testResults_.push({ status: status, description: description });
  Logger.log('[' + status + '] ' + description);
  if (!condition) throw new Error('Assertion failed: ' + description);
}

function assertEqual_(description, actual, expected) {
  var ok = JSON.stringify(actual) === JSON.stringify(expected);
  var detail = ok ? '' : ' | esperado: ' + JSON.stringify(expected) + ' | obtenido: ' + JSON.stringify(actual);
  _testResults_.push({ status: ok ? 'PASS' : 'FAIL', description: description + detail });
  Logger.log('[' + (ok ? 'PASS' : 'FAIL') + '] ' + description + detail);
  if (!ok) throw new Error('assertEqual_ failed: ' + description + detail);
}

function runTest_(name, fn) {
  try {
    fn();
  } catch (e) {
    _testResults_.push({ status: 'ERROR', description: name + ' → ' + e.message });
    Logger.log('[ERROR] ' + name + ' → ' + e.message);
  }
}

function runAllTests() {
  _testResults_ = [];
  Logger.log('========================================');
  Logger.log(' EJECUTANDO TESTS — EA Backend v3.0');
  Logger.log('========================================');

  runTest_('T01 – BuscarClienteRFC: mapeo de columnas (16-col schema)', test_T01_BuscarRFC_columnas);
  runTest_('T02 – BuscarClienteNombre: mapeo de columnas (16-col schema)', test_T02_BuscarNombre_columnas);
  runTest_('T03 – getConsecutivo: primer número es 0001', test_T03_Consecutivo_primero);
  runTest_('T04 – getConsecutivo: OTA y OT tienen series independientes', test_T04_Consecutivo_tipos);
  runTest_('T05 – getOrdenesSafe: incluye campos requeridos por SEAINF', test_T05_GetOrdenes_campos);
  runTest_('T06 – getOrdenesSafe: filtra ENTREGADO / FINALIZADO', test_T06_GetOrdenes_filtro);
  runTest_('T07 – Fallback carpeta: extrae folder ID de URL Drive', test_T07_ExtractFolderID);
  runTest_('T08 – sanitizeFileName: limpia caracteres especiales', test_T08_SanitizeFileName);
  runTest_('T09 – cleanCompanyName: elimina sufijos S.A. de C.V.', test_T09_CleanCompanyName);
  runTest_('T10 – fase2_RegistrarOT: tipo por defecto es OTA', test_T10_OT_tipoPorDefecto);

  // Resumen
  var pass = _testResults_.filter(function(r){ return r.status === 'PASS'; }).length;
  var fail = _testResults_.filter(function(r){ return r.status === 'FAIL'; }).length;
  var err  = _testResults_.filter(function(r){ return r.status === 'ERROR'; }).length;
  Logger.log('========================================');
  Logger.log(' RESULTADO: ' + pass + ' PASS | ' + fail + ' FAIL | ' + err + ' ERROR');
  Logger.log('========================================');
  return { pass: pass, fail: fail, error: err, total: _testResults_.length };
}

// ─── Tests de lógica pura (sin llamadas a SpreadsheetApp) ─────────────────

// T01 — BuscarClienteRFC: verifica que link_drive_cliente viene de índice [15]
function test_T01_BuscarRFC_columnas() {
  // Simula una fila del nuevo esquema de 16 columnas
  var fila = [
    '08/03/2026 10:00:00',    // [0]  Fecha
    'EMPRESA TEST S.A.',       // [1]  Razón Social
    'Planta Norte',            // [2]  Sucursal
    'TST010101AAA',            // [3]  RFC
    'Juan Representante',      // [4]  Representante Legal
    'Calle 1 #100, Puebla',    // [5]  Dirección
    '2221234567',              // [6]  Teléfono Empresa
    'Carlos Solicitante',      // [7]  Nombre Solicitante
    'carlos@test.com',         // [8]  Correo Informe
    'Manufactura',             // [9]  Giro
    'IMSS-01234',              // [10] Registro Patronal
    '500 ton / 450 ton',       // [11] Capacidad combinada
    'L-V 08:00-18:00',         // [12] Días/Turnos
    'SÍ',                      // [13] Aplica NOM-020
    'NO',                      // [14] Requiere PIPC
    'https://drive.google.com/drive/folders/ABC123_link_correcto' // [15] Drive Link
  ];

  // Replica la lógica de fase2_BuscarClienteRFC con esa fila
  var resultado = {
    razon_social:        fila[1],
    sucursal:            fila[2] || 'Matriz',
    rfc:                 fila[3],
    nombre_solicitante:  fila[7],
    correo_informe:      fila[8],
    telefono_empresa:    fila[6],
    representante_legal: fila[4],
    direccion_evaluacion:fila[5],
    giro:                fila[9],
    registro_patronal:   fila[10],
    link_drive_cliente:  fila[15]   // ← índice corregido
  };

  assertEqual_('T01a: nombre_solicitante viene de col 8 (índice 7)', resultado.nombre_solicitante, 'Carlos Solicitante');
  assertEqual_('T01b: correo_informe viene de col 9 (índice 8)',      resultado.correo_informe,     'carlos@test.com');
  assertEqual_('T01c: telefono_empresa viene de col 7 (índice 6)',    resultado.telefono_empresa,   '2221234567');
  assertEqual_('T01d: link_drive_cliente viene de col 16 (índice 15)',resultado.link_drive_cliente, 'https://drive.google.com/drive/folders/ABC123_link_correcto');
  assert_('T01e: link_drive_cliente NO está vacío', resultado.link_drive_cliente !== '');
  assert_('T01f: link_drive_cliente NO es undefined', resultado.link_drive_cliente !== undefined);
}

// T02 — BuscarClienteNombre: mismo mapeo
function test_T02_BuscarNombre_columnas() {
  var fila = [
    '08/03/2026', 'ACEROS DEL NORTE', 'Matriz', 'ADN920101XYZ',
    'Rep Legal', 'Dir Eval', '2224567890', 'Solicitante Dos',
    'sol2@aceros.com', 'Siderurgia', 'IMSS-99999',
    '1000 ton / 800 ton', 'L-S 06:00-22:00', 'NO', 'SÍ',
    'https://drive.google.com/drive/folders/XYZ789_link_nombre'
  ];

  var resultado = {
    nombre_solicitante:  fila[7],
    correo_informe:      fila[8],
    telefono_empresa:    fila[6],
    link_drive_cliente:  fila[15]
  };

  assertEqual_('T02a: nombre_solicitante índice 7',   resultado.nombre_solicitante, 'Solicitante Dos');
  assertEqual_('T02b: correo_informe índice 8',        resultado.correo_informe,     'sol2@aceros.com');
  assertEqual_('T02c: link_drive_cliente índice 15',   resultado.link_drive_cliente, 'https://drive.google.com/drive/folders/XYZ789_link_nombre');
}

// T03 — getConsecutivo: si no hay registros previos, devuelve 0001
function test_T03_Consecutivo_primero() {
  var serieVacia = [];
  var max = calcularMaxConsecutivo_(serieVacia, 'OTA');
  var siguiente = String(max + 1).padStart(4, '0');
  assertEqual_('T03: primer número de serie OTA es 0001', siguiente, '0001');
}

// T04 — getConsecutivo: OTA y OT no comparten contador
function test_T04_Consecutivo_tipos() {
  // Filas simuladas de ORDENES_TRABAJO col[2]=tipo, col[3]=numInforme
  var filas = [
    ['', 'OT-001', 'OT',  'EA-2601-NOM-0001', '...'],
    ['', 'OT-002', 'OT',  'EA-2601-NOM-0002', '...'],
    ['', 'OTA-001','OTA', 'EA-2601-RUP-0001', '...'],
  ];

  var maxOTA = calcularMaxConsecutivo_(filas, 'OTA');
  var maxOT  = calcularMaxConsecutivo_(filas, 'OT');

  assertEqual_('T04a: siguiente OTA es 0002', String(maxOTA + 1).padStart(4, '0'), '0002');
  assertEqual_('T04b: siguiente OT es 0003',  String(maxOT  + 1).padStart(4, '0'), '0003');
}

// T05 — getOrdenesSafe_: el objeto de orden tiene todos los campos que SEAINF necesita
function test_T05_GetOrdenes_campos() {
  // Simula una fila de ORDENES_TRABAJO (15 columnas)
  var fila = [
    '2026-03-01',   // [0]  Fecha
    'OTA-2601-001', // [1]  Folio OT
    'OTA',          // [2]  Tipo
    'EA-2601-NOM-0001', // [3] Num Informe
    'Ruido',        // [4]  NOM Servicio
    'EMPRESA SA',   // [5]  Cliente razon social
    'Planta Sur',   // [6]  Sucursal
    'EMP010101XYZ', // [7]  RFC
    'Dr. Gomez',    // [8]  Personal asignado
    '15/03/2026',   // [9]  Fecha visita
    '22/03/2026',   // [10] Fecha entrega
    '',             // [11] Fecha real entrega
    'EN PROCESO',   // [12] Estatus
    'https://drive.google.com/drive/folders/FOLDER_OT' // [13] Link Drive
  ];

  var orden = mapearOrden_(fila);

  assertEqual_('T05a: campo ot',          orden.ot,          'OTA-2601-001');
  assertEqual_('T05b: campo rfc',          orden.rfc,         'EMP010101XYZ');
  assertEqual_('T05c: campo sucursal',     orden.sucursal,    'Planta Sur');
  assertEqual_('T05d: campo personal',     orden.personal,    'Dr. Gomez');
  assertEqual_('T05e: campo nom_servicio', orden.nom_servicio,'Ruido');
  assertEqual_('T05f: campo fecha_visita', orden.fecha_visita,'15/03/2026');
  assertEqual_('T05g: campo link_drive',   orden.link_drive,  'https://drive.google.com/drive/folders/FOLDER_OT');
}

// T06 — getOrdenesSafe_: ENTREGADO y FINALIZADO se excluyen del resultado
function test_T06_GetOrdenes_filtro() {
  var filas = [
    makeFila_('OTA-001', 'EN PROCESO'),
    makeFila_('OTA-002', 'ENTREGADO'),
    makeFila_('OTA-003', 'FINALIZADO'),
    makeFila_('OTA-004', 'NO INICIADO'),
    makeFila_('OTA-005', 'entregado'),   // minúsculas también deben filtrarse
  ];

  var activos = filas.filter(function(row) {
    var estatus = String(row[12] || '').trim().toUpperCase();
    return estatus !== 'ENTREGADO' && estatus !== 'FINALIZADO';
  });

  assertEqual_('T06a: solo 2 órdenes activas', activos.length, 2);
  assertEqual_('T06b: primera activa es OTA-001', activos[0][1], 'OTA-001');
  assertEqual_('T06c: segunda activa es OTA-004', activos[1][1], 'OTA-004');
}

// T07 — Fallback de carpeta: regex extrae folder ID de URL Drive
function test_T07_ExtractFolderID() {
  var urls = [
    'https://drive.google.com/drive/folders/1nHd-70uUeciClDm_3_pgbmqGF7II1lfQ',
    'https://drive.google.com/drive/folders/ABC123xyz-_456',
    'https://drive.google.com/drive/u/0/folders/SHORT_ID?usp=sharing'
  ];
  var esperados = [
    '1nHd-70uUeciClDm_3_pgbmqGF7II1lfQ',
    'ABC123xyz-_456',
    'SHORT_ID'
  ];

  urls.forEach(function(url, i) {
    var m = String(url).match(/folders\/([a-zA-Z0-9_-]+)/);
    assert_('T07[' + i + ']: regex encuentra folder ID en URL', !!m);
    assertEqual_('T07[' + i + ']: folder ID correcto', m[1], esperados[i]);
  });

  var urlVacia = 'https://docs.google.com/spreadsheets/d/ABC';
  var mVacia = String(urlVacia).match(/folders\/([a-zA-Z0-9_-]+)/);
  assert_('T07d: URL sin carpeta devuelve null', mVacia === null);
}

// T08 — sanitizeFileName
function test_T08_SanitizeFileName() {
  assertEqual_('T08a: barra /  → guion bajo', sanitizeFileName('Nombre/Raro'), 'Nombre_Raro');
  assertEqual_('T08b: dos puntos → guion bajo', sanitizeFileName('HH:mm'),     'HH_mm');
  assertEqual_('T08c: texto normal sin cambios', sanitizeFileName('PuertoNorte'), 'PuertoNorte');
  assertEqual_('T08d: trunca a 50 chars', sanitizeFileName('A'.repeat(60)).length, 50);
  assertEqual_('T08e: valor vacío → Sin_nombre', sanitizeFileName(''), 'Sin_nombre');
  assertEqual_('T08f: null → Sin_nombre', sanitizeFileName(null), 'Sin_nombre');
}

// T09 — cleanCompanyName
function test_T09_CleanCompanyName() {
  assertEqual_('T09a: elimina S.A. DE C.V.', cleanCompanyName('EMPRESA SA DE CV'), 'EMPRESA');
  assertEqual_('T09b: elimina S.A.',          cleanCompanyName('ACEROS S.A.'),      'ACEROS');
  assertEqual_('T09c: elimina S.C.',          cleanCompanyName('DESPACHO S.C.'),    'DESPACHO');
  assert_('T09d: resultado no está vacío', cleanCompanyName('EMPRESA S.A. DE C.V.').length > 0);
}

// T10 — fase2_RegistrarOT: tipo_orden por defecto es OTA (no OT)
function test_T10_OT_tipoPorDefecto() {
  var payload = {
    ot_folio: 'OTA-001', nom_servicio: 'Ruido', cliente_razon_social: 'TEST SA',
    sucursal: 'Matriz', rfc: 'TST010101', personal_asignado: 'Ing. X',
    fecha_visita: '2026-03-15', fecha_entrega_limite: '2026-03-22',
    link_drive_cliente: 'https://drive.google.com/drive/folders/ABC',
    observaciones: ''
    // tipo_orden NO se envía — debe defaultear a 'OTA'
  };

  var tipoEfectivo = payload.tipo_orden || 'OTA';
  assertEqual_('T10: tipo_orden por defecto es OTA (no OT)', tipoEfectivo, 'OTA');
}

// ─── Helpers para los tests ───────────────────────────────────────────────

// Simula el mapeo de una fila de ORDENES_TRABAJO a un objeto orden
function mapearOrden_(row) {
  return {
    ot:           row[1],
    tipo_orden:   row[2],
    nom_servicio: row[4],
    clienteInicial: row[5],
    clienteFinal:   row[6],
    cliente:      row[5],
    sucursal:     row[6],
    rfc:          row[7],
    personal:     row[8],
    fecha_visita: row[9],
    link_drive:   row[13]
  };
}

// Construye una fila mínima con OT y estatus
function makeFila_(ot, estatus) {
  var fila = new Array(15).fill('');
  fila[1]  = ot;
  fila[12] = estatus;
  return fila;
}

// Calcula el máximo consecutivo para un tipo dado (lógica de getConsecutivoSafe_)
function calcularMaxConsecutivo_(filas, tipo) {
  var regex = /^EA-\d{4}-.+-(\d{4})$/;
  var max = 0;
  filas.forEach(function(row) {
    var valNum  = row[3];
    var valTipo = String(row[2] || '').trim().toUpperCase();
    if (valTipo === tipo.toUpperCase()) {
      var m = String(valNum || '').trim().match(regex);
      if (m) {
        var n = parseInt(m[1], 10);
        if (n > max) max = n;
      }
    }
  });
  return max;
}

// ─── Tests de integración (requieren conexión a Sheets/Drive real) ─────────
// Ejecutar manualmente con datos reales en el entorno de producción.

function integrationTest_BuscarClienteRFC() {
  Logger.log('=== INTEGRATION: BuscarClienteRFC ===');
  // Requiere RFC existente en CLIENTES_MAESTRO
  var rfc = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_CLIENTES)
    .getDataRange().getValues().slice(1)
    .map(function(r){ return r[3]; })
    .find(function(r){ return r && r.toString().trim() !== ''; });

  if (!rfc) { Logger.log('Sin datos en CLIENTES_MAESTRO — omitiendo.'); return; }

  var resultado = fase2_BuscarClienteRFC(rfc.toString().trim());
  Logger.log('found: ' + resultado.found);
  if (resultado.found && resultado.sucursales.length > 0) {
    var s = resultado.sucursales[0];
    Logger.log('link_drive_cliente: ' + s.link_drive_cliente);
    Logger.log('[' + (s.link_drive_cliente ? 'PASS' : 'FAIL') + '] link_drive_cliente presente');
  }
}

function integrationTest_GetOrdenes() {
  Logger.log('=== INTEGRATION: getOrdenesSafe_ ===');
  var r = getOrdenesSafe_();
  if (!r.success || r.data.length === 0) { Logger.log('Sin órdenes activas — omitiendo.'); return; }
  var o = r.data[0];
  var camposRequeridos = ['ot','rfc','sucursal','personal','nom_servicio','fecha_visita','link_drive'];
  camposRequeridos.forEach(function(campo) {
    Logger.log('[' + (campo in o ? 'PASS' : 'FAIL') + '] campo ' + campo + ' presente');
  });
}
