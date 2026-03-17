// =========================================================================
//  MIGRACIÓN DE DATOS ANTERIORES — Ejecutiva Ambiental
//  Lee del sistema SEAINF anterior (hoja "Informes") y escribe en el nuevo
//  sistema SEA v3.0 (ORDENES_TRABAJO + CLIENTES_MAESTRO).
//
//  PARA EJECUTAR:
//  1. Copia este archivo en el editor de Google Apps Script del proyecto SEA v3.0
//  2. Ejecuta migrarDatosAnteriores() desde el editor
//  3. Revisa Ver → Registros para el resumen
//
//  Estructura origen — hoja "Informes" (17 columnas A-Q):
//  A[0]=Timestamp  B[1]=NumInforme  C[2]=TipoOrden  D[3]=OT
//  E[4]=NOM  F[5]=Cliente  G[6]=Solicitante  H[7]=RFC
//  I[8]=Telefono  J[9]=Direccion  K[10]=FechaServicio  L[11]=FechaEntrega
//  M[12]=EsCapacitacion  N[13]=Estatus  O[14]=LinkDrive
//  P[15]=Responsable  Q[16]=Sucursal
// =========================================================================

var SOURCE_SPREADSHEET_ID = '1aa2uX6gqHINUP_h-HcNsxJzlwSI1OnSVzqMaFaI_8NE';
var SOURCE_SHEET_NAME     = 'Informes';
var DEST_SPREADSHEET_ID   = '1MoScea4CYg0NCjvPjHqZwV0cKhrd2nxfW8LYhz_4pDo';

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

  // Fila 1 = encabezados, se omite (i empieza en 1)
  for (var i = 1; i < srcData.length; i++) {
    var row = srcData[i];

    // Ignorar filas con OT vacía o con errores (#REF!, #N/A, etc.)
    var ot = String(row[3] || '').trim();
    if (!ot || ot.indexOf('#') === 0) { skipped++; continue; }

    // ── Mapeo de columnas origen ─────────────────────────────────────────
    var timestamp    = row[0]  || hoy;
    var numInforme   = String(row[1]  || '').trim();
    var tipoOrden    = String(row[2]  || 'OT').trim().toUpperCase();
    // ot ya está leído arriba
    var nom          = String(row[4]  || '').trim();
    var cliente      = String(row[5]  || '').trim();
    var solicitante  = String(row[6]  || '').trim();
    var rfc          = String(row[7]  || '').trim();
    var telefono     = String(row[8]  || '').trim();
    var direccion    = String(row[9]  || '').trim();
    var fechaVisita  = row[10] || '';
    var fechaEntrega = row[11] || '';
    var esCapac      = String(row[12] || '').trim().toUpperCase();
    var estatusRaw   = String(row[13] || '').trim().toUpperCase();
    var linkDrive    = String(row[14] || '').trim();
    var responsable  = String(row[15] || '').trim();
    var sucursal     = String(row[16] || '').trim();

    // ── Normalizar TipoOrden ─────────────────────────────────────────────
    if (tipoOrden !== 'OTB') tipoOrden = 'OT';

    // ── Normalizar Estatus Externo (visible al cliente) ──────────────────
    var estatusExterno;
    if (estatusRaw === 'FINALIZADO' || estatusRaw === 'TRUE' ||
        estatusRaw === 'ENTREGADO' || estatusRaw === 'ENTREGADA') {
      estatusExterno = 'ENTREGADO';
    } else if (estatusRaw === 'EN PROCESO' || estatusRaw === 'EN PROGRESO' ||
               estatusRaw === 'PARA REVISION' || estatusRaw === 'PARA IMPRESION') {
      estatusExterno = 'EN PROCESO';
    } else if (estatusRaw === 'CANCELADO') {
      estatusExterno = 'CANCELADO';
    } else {
      estatusExterno = 'NO INICIADO';
    }

    // ── Normalizar Estatus Informe (seguimiento interno) ─────────────────
    var estatusInforme;
    if (estatusRaw === 'FINALIZADO' || estatusRaw === 'TRUE' ||
        estatusRaw === 'ENTREGADO' || estatusRaw === 'ENTREGADA') {
      estatusInforme = 'FINALIZADO';
    } else if (estatusRaw === 'EN PROCESO' || estatusRaw === 'EN PROGRESO') {
      estatusInforme = 'EN PROCESO';
    } else if (estatusRaw === 'PARA REVISION') {
      estatusInforme = 'PARA REVISION';
    } else if (estatusRaw === 'PARA IMPRESION') {
      estatusInforme = 'PARA IMPRESION';
    } else if (estatusRaw === 'CANCELADO') {
      estatusInforme = 'CANCELADO';
    } else {
      estatusInforme = 'NO INICIADO';
    }

    // ── Observaciones: Solicitante + Teléfono + Dirección + Capacitación ─
    var obs = [];
    if (solicitante) obs.push('Solicitante: ' + solicitante);
    if (telefono)    obs.push('Tel: ' + telefono);
    if (direccion)   obs.push('Dir: ' + direccion);
    if (esCapac === 'SI') obs.push('Capacitación: SÍ');
    var observaciones = obs.join(' | ');

    // ── Fila destino ORDENES_TRABAJO (16 columnas, índices 0-15) ─────────
    otRows.push([
      timestamp,      // [0]  FECHA_REGISTRO
      ot,             // [1]  OT
      tipoOrden,      // [2]  TIPO
      numInforme,     // [3]  NUM_INFORME
      nom,            // [4]  NOM
      cliente,        // [5]  CLIENTE (Razón Social)
      sucursal,       // [6]  SUCURSAL
      rfc,            // [7]  RFC
      responsable,    // [8]  PERSONAL
      fechaVisita,    // [9]  FECHA_VISITA
      fechaEntrega,   // [10] FECHA_ENTREGA
      '',             // [11] FECHA_REAL (no existía)
      estatusExterno, // [12] ESTATUS_EXTERNO
      linkDrive,      // [13] LINK_DRIVE
      observaciones,  // [14] OBSERVACIONES
      estatusInforme  // [15] ESTATUS_INFORME
    ]);
    migradas++;

    // ── Acumular clientes únicos para CLIENTES_MAESTRO ───────────────────
    if (cliente) {
      var clave = cliente + '||' + sucursal;
      if (!clientesMap[clave]) {
        clientesMap[clave] = {
          empresa:    cliente,
          sucursal:   sucursal,
          rfc:        rfc,
          direccion:  direccion,
          telefono:   telefono,
          solicitante:solicitante
        };
      }
    }
  }

  // ── Escribir ORDENES_TRABAJO de golpe ─────────────────────────────────
  if (otRows.length > 0) {
    destOT.getRange(destOT.getLastRow() + 1, 1, otRows.length, 16).setValues(otRows);
  }
  Logger.log('✅ ' + migradas + ' órdenes migradas a ORDENES_TRABAJO');

  // ── Escribir CLIENTES_MAESTRO (sin duplicar) ──────────────────────────
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
    // 22 columnas: FECHA, RAZON_SOCIAL, SUCURSAL, RFC, REPRESENTANTE,
    //              DIRECCION, TELEFONO, SOLICITANTE, + 14 vacías
    destClientes.appendRow([
      hoy,          // [0]  FECHA_REGISTRO
      c.empresa,    // [1]  RAZON_SOCIAL
      c.sucursal,   // [2]  SUCURSAL
      c.rfc,        // [3]  RFC
      '',           // [4]  REPRESENTANTE
      c.direccion,  // [5]  DIRECCION
      c.telefono,   // [6]  TELEFONO
      c.solicitante,// [7]  SOLICITANTE
      '', '', '', '', '', '', '', '', '', '', '', '', '', '' // [8-21] vacíos
    ]);
    clientesNuevos++;
  }

  Logger.log('✅ ' + clientesNuevos + ' clientes nuevos en CLIENTES_MAESTRO');
  Logger.log('⚠️  ' + skipped + ' filas ignoradas (vacías o con errores)');
  Logger.log('──────────────────────────────────────────────────────────────');
  Logger.log('PENDIENTE: Revisar columna RFC en CLIENTES_MAESTRO si algún');
  Logger.log('           cliente quedó sin RFC para habilitar búsqueda en SEAOT.');
}
