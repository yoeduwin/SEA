// =========================================================================
//  MIGRACIÓN DE DATOS ANTERIORES — Ejecutiva Ambiental
//  Archivo independiente. NO modifica BACKEND_FIXES.gs.
//
//  ANTES DE EJECUTAR:
//  1. Reemplaza SOURCE_SPREADSHEET_ID con el ID de tu Sheet anterior
//     (está en la URL: docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit)
//  2. Reemplaza SOURCE_SHEET_NAME con el nombre exacto de la pestaña
//  3. Reemplaza DEST_SPREADSHEET_ID con el ID del Spreadsheet del sistema
//  4. Ejecuta migrarDatosAnteriores() desde el editor GAS
//  5. Revisa el Logger (Ver → Registros) para el resumen
//
//  Columnas asumidas en el Sheet origen (fila 1 = encabezados):
//  A=# | B=OrdenTrabajo | C=Cotización | D=CLIENTE INICIAL | E=CLIENTE FINAL
//  F=NOM | G=Proveedor | H=Personal | I=FECHA VISITA | ... | N=FECHA INFORME DIGITAL
//  O=FECHA INFORME FÍSICO | P=ESTATUS
// =========================================================================

var SOURCE_SPREADSHEET_ID = 'REEMPLAZA_CON_ID_SHEET_ANTERIOR';
var SOURCE_SHEET_NAME     = 'REEMPLAZA_CON_NOMBRE_PESTANA';    // ej. 'PANEL'
var DEST_SPREADSHEET_ID   = 'REEMPLAZA_CON_ID_SHEET_SISTEMA';  // el mismo que CONFIG.SPREADSHEET_ID

function migrarDatosAnteriores() {

  // ── Abrir hojas ──────────────────────────────────────────────────────────
  var srcSS    = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var srcSheet = srcSS.getSheetByName(SOURCE_SHEET_NAME);
  if (!srcSheet) {
    Logger.log('ERROR: No se encontró la pestaña "' + SOURCE_SHEET_NAME + '"');
    return;
  }

  var destSS       = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  var destOT       = destSS.getSheetByName('ORDENES_TRABAJO');
  var destClientes = destSS.getSheetByName('CLIENTES_MAESTRO');
  if (!destOT || !destClientes) {
    Logger.log('ERROR: No se encontraron las hojas ORDENES_TRABAJO o CLIENTES_MAESTRO.');
    Logger.log('Verifica que DEST_SPREADSHEET_ID sea correcto.');
    return;
  }

  var srcData = srcSheet.getDataRange().getValues();
  var hoy     = new Date();
  var otRows      = [];
  var clientesMap = {};   // key: "RazonSocial||Sucursal" para deduplicar
  var skipped  = 0;
  var migradas = 0;

  for (var i = 1; i < srcData.length; i++) {
    var row = srcData[i];

    // Ignorar filas vacías o con errores (#REF!)
    var folio = String(row[1] || '').trim();
    if (!folio || folio.indexOf('#') === 0) { skipped++; continue; }

    // ── Mapeo de columnas origen ─────────────────────────────────────────
    var cotizacion     = String(row[2]  || '').trim();   // C — Cotización
    var empresa        = String(row[3]  || '').trim();   // D — CLIENTE INICIAL = Razón Social
    var sucursal       = String(row[4]  || '').trim();   // E — CLIENTE FINAL   = Sucursal
    var nomServicio    = String(row[5]  || '').trim();   // F — NOM
    // row[6] = Proveedor — no se migra
    var personal       = String(row[7]  || '').trim();   // H — Personal
    var fechaVisita    = row[8]  || '';                  // I — Fecha Visita
    // row[9..12] = SignatarioReal, Equipos, Informe, col-O — ignorar
    var fechaEntrega   = row[13] || '';                  // N — Fecha Informe Digital
    var fechaFisicoRaw = row[14];                        // O — Fecha Informe Físico
    var fechaFisico    = fechaFisicoRaw
      ? Utilities.formatDate(new Date(fechaFisicoRaw), Session.getScriptTimeZone(), 'dd/MM/yyyy')
      : '';
    var estatusRaw     = row[15] || '';                  // P — Estatus

    // ── Tipo Orden desde el folio (OTB25-xx → OTB, resto → OT) ─────────
    var tipoOrden = folio.toUpperCase().indexOf('OTB') === 0 ? 'OTB' : 'OT';

    // ── Normalizar Estatus ───────────────────────────────────────────────
    var estatusNorm = 'NO INICIADO';
    var estatusStr  = String(estatusRaw).toUpperCase().trim();
    if (estatusStr === 'TRUE' || estatusStr === 'ENTREGADO' || estatusStr === 'ENTREGADA') {
      estatusNorm = 'ENTREGADO';
    } else if (estatusStr === 'EN PROCESO' || estatusStr === 'EN PROGRESO') {
      estatusNorm = 'EN PROCESO';
    }

    // ── Observaciones: Cotización + Fecha Informe Físico ────────────────
    var obs = cotizacion;
    if (fechaFisico) obs += (obs ? ' | Físico: ' : 'Físico: ') + fechaFisico;

    // ── Fila destino ORDENES_TRABAJO (16 columnas) ──────────────────────
    otRows.push([
      hoy,           // Col 1  — Fecha Registro
      folio,         // Col 2  — OT Folio
      tipoOrden,     // Col 3  — Tipo Orden
      '',            // Col 4  — Num Informe (vacío)
      nomServicio,   // Col 5  — Nom Servicio
      empresa,       // Col 6  — Cliente Razón Social
      sucursal,      // Col 7  — Sucursal
      '',            // Col 8  — RFC (pendiente — llenar en Sheets después)
      personal,      // Col 9  — Personal Asignado
      fechaVisita,   // Col 10 — Fecha Visita
      fechaEntrega,  // Col 11 — Fecha Entrega Límite
      '',            // Col 12 — (vacío)
      estatusNorm,   // Col 13 — Estatus Externo
      '',            // Col 14 — Link Drive (no disponible)
      obs,           // Col 15 — Observaciones
      'NO INICIADO'  // Col 16 — Estatus Informe (no existía antes)
    ]);
    migradas++;

    // ── Acumular clientes únicos para CLIENTES_MAESTRO ──────────────────
    if (empresa) {
      var clave = empresa + '||' + sucursal;
      if (!clientesMap[clave]) {
        clientesMap[clave] = { empresa: empresa, sucursal: sucursal };
      }
    }
  }

  // ── Escribir ORDENES_TRABAJO de golpe ────────────────────────────────────
  if (otRows.length > 0) {
    destOT.getRange(destOT.getLastRow() + 1, 1, otRows.length, 16).setValues(otRows);
  }
  Logger.log('✅ ' + migradas + ' órdenes migradas a ORDENES_TRABAJO');

  // ── Escribir CLIENTES_MAESTRO (sin duplicar) ─────────────────────────────
  var clientesExistentes = destClientes.getDataRange().getValues();
  var yaExisten = {};
  for (var j = 1; j < clientesExistentes.length; j++) {
    var k = String(clientesExistentes[j][1] || '').trim()
          + '||'
          + String(clientesExistentes[j][2] || '').trim();
    yaExisten[k] = true;
  }
  var clientesNuevos = 0;
  for (var cl in clientesMap) {
    if (yaExisten[cl]) continue;
    var c = clientesMap[cl];
    destClientes.appendRow([
      hoy, c.empresa, c.sucursal,
      '', '', '', '', '', '', '', '', '', '', '', '', ''  // 16 columnas
    ]);
    clientesNuevos++;
  }

  Logger.log('✅ ' + clientesNuevos + ' clientes nuevos en CLIENTES_MAESTRO');
  Logger.log('⚠️  ' + skipped + ' filas ignoradas (vacías o con errores)');
  Logger.log('──────────────────────────────────────────────────────────────');
  Logger.log('PENDIENTE: Llenar columna RFC (col D) en CLIENTES_MAESTRO');
  Logger.log('           para habilitar búsqueda por RFC en SEAOT.');
}
