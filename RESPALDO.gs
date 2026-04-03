// =========================================================================
// RESPALDO Y CONTINUIDAD OPERATIVA — Sistema SEA (Ejecutiva Ambiental)
// =========================================================================
// Propósito: Mitigar el riesgo de punto único de falla (B-06).
//   1. Respaldo automático semanal del Spreadsheet principal a Drive.
//   2. Instrucciones para agregar co-propietarios al proyecto GAS y recursos.
//
// CONFIGURACIÓN INICIAL (ejecutar una sola vez desde el editor GAS):
//   1. Ejecutar: configurarTriggerRespaldo()    → activa el backup semanal
//   2. Ejecutar: crearCarpetaRespaldos()        → crea carpeta destino
//
// CO-PROPIETARIOS (pasos manuales — hacerlo hoy):
//   1. Abrir: https://script.google.com → seleccionar este proyecto GAS
//      → Compartir → agregar segunda cuenta como EDITOR (o Propietario).
//   2. Abrir el Google Spreadsheet de producción → Compartir → segunda cuenta como EDITOR.
//   3. Abrir la carpeta raíz de Drive → Compartir → segunda cuenta como EDITOR.
//   Segunda cuenta recomendada: otra cuenta @gmail.com o @ejecutivambiental.com del equipo.
// =========================================================================

const RESPALDO_CONFIG = {
  // Spreadsheet principal (mismo que CONFIG.SPREADSHEET_ID)
  SPREADSHEET_ID: '1MoScea4CYg0NCjvPjHqZwV0cKhrd2nxfW8LYhz_4pDo',
  // Nombre de la carpeta donde se guardarán los respaldos en Drive
  CARPETA_RESPALDO_NOMBRE: 'SEA_RESPALDOS_AUTOMATICOS',
  // Número máximo de respaldos a conservar (los más antiguos se eliminan)
  MAX_RESPALDOS: 12,
  // Email para notificación de éxito/fallo del respaldo
  EMAIL_RESPALDO: 'eduwin.ejecutiva@gmail.com'
};

// ── crearCarpetaRespaldos ─────────────────────────────────────────────────
/**
 * Crea (si no existe) la carpeta de respaldos en la raíz de Mi Drive.
 * Ejecutar UNA VEZ desde el editor GAS antes de configurar el trigger.
 */
function crearCarpetaRespaldos() {
  const iter = DriveApp.getFoldersByName(RESPALDO_CONFIG.CARPETA_RESPALDO_NOMBRE);
  if (iter.hasNext()) {
    Logger.log('Carpeta de respaldos ya existe: ' + iter.next().getUrl());
    return;
  }
  const carpeta = DriveApp.createFolder(RESPALDO_CONFIG.CARPETA_RESPALDO_NOMBRE);
  Logger.log('✅ Carpeta de respaldos creada: ' + carpeta.getUrl());
  Logger.log('   Guarda este URL en un lugar seguro: ' + carpeta.getUrl());
}

// ── crearRespaldoSemanal ──────────────────────────────────────────────────
/**
 * Crea una copia del Spreadsheet principal en la carpeta de respaldos.
 * Elimina los respaldos más antiguos si se supera MAX_RESPALDOS.
 * Se ejecuta automáticamente según el trigger semanal configurado.
 */
function crearRespaldoSemanal() {
  try {
    // Encontrar carpeta de respaldos
    const iter = DriveApp.getFoldersByName(RESPALDO_CONFIG.CARPETA_RESPALDO_NOMBRE);
    const carpeta = iter.hasNext() ? iter.next() : DriveApp.createFolder(RESPALDO_CONFIG.CARPETA_RESPALDO_NOMBRE);

    // Crear copia con timestamp (exportar como XLSX para no generar un proyecto GAS nuevo)
    const timestamp = Utilities.formatDate(new Date(), 'GMT-6', 'yyyy-MM-dd_HHmm');
    const nombreCopia = `SEA_Respaldo_${timestamp}`;
    const exportUrl = `https://docs.google.com/spreadsheets/d/${RESPALDO_CONFIG.SPREADSHEET_ID}/export?format=xlsx`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error(`Export falló con código HTTP ${response.getResponseCode()}`);
    }
    const blob = response.getBlob().setName(`${nombreCopia}.xlsx`);
    const copia = carpeta.createFile(blob);

    Logger.log('✅ Respaldo creado: ' + copia.getUrl());

    // Limpiar respaldos antiguos si hay más de MAX_RESPALDOS
    const archivos = [];
    const filesIter = carpeta.getFiles();
    while (filesIter.hasNext()) { archivos.push(filesIter.next()); }
    archivos.sort((a, b) => a.getDateCreated() - b.getDateCreated()); // ascendente

    const excedentes = archivos.length - RESPALDO_CONFIG.MAX_RESPALDOS;
    if (excedentes > 0) {
      archivos.slice(0, excedentes).forEach(f => {
        f.setTrashed(true);
        Logger.log('Respaldo antiguo eliminado: ' + f.getName());
      });
    }

    // Notificación de éxito
    GmailApp.sendEmail(
      RESPALDO_CONFIG.EMAIL_RESPALDO,
      '✅ Respaldo SEA exitoso — ' + timestamp,
      `El respaldo semanal del sistema SEA se completó correctamente.\n\n` +
      `Archivo: ${nombreCopia}\n` +
      `URL: ${copia.getUrl()}\n` +
      `Respaldos conservados: ${Math.min(archivos.length, RESPALDO_CONFIG.MAX_RESPALDOS)} (formato .xlsx)\n\n` +
      `Sistema de Respaldo Automático — Ejecutiva Ambiental`
    );

  } catch (e) {
    Logger.log('ERROR en crearRespaldoSemanal: ' + e.message);
    // Notificación de fallo
    try {
      GmailApp.sendEmail(
        RESPALDO_CONFIG.EMAIL_RESPALDO,
        '❌ FALLO en respaldo SEA — ' + new Date().toISOString(),
        `El respaldo automático del sistema SEA FALLÓ.\n\nError: ${e.message}\n\nRevisa el proyecto GAS para diagnosticar.`
      );
    } catch (_) {}
  }
}

// ── configurarTriggerRespaldo ─────────────────────────────────────────────
/**
 * Configura un trigger semanal (lunes a las 3:00 AM GMT-6) para ejecutar
 * crearRespaldoSemanal() automáticamente.
 *
 * EJECUTAR UNA SOLA VEZ desde el editor GAS.
 * Para verificar: Ver → Triggers en el menú del editor GAS.
 */
function configurarTriggerRespaldo() {
  // Eliminar triggers anteriores de esta función para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'crearRespaldoSemanal') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Trigger anterior eliminado.');
    }
  });

  // Crear nuevo trigger semanal (lunes, 3-4 AM)
  ScriptApp.newTrigger('crearRespaldoSemanal')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(3)
    .create();

  Logger.log('✅ Trigger semanal configurado: crearRespaldoSemanal() → Lunes 3 AM GMT-6');
  Logger.log('   Verifica en: Ver → Triggers del editor GAS');
}

// ── verificarEstadoRespaldos ──────────────────────────────────────────────
/**
 * Verifica el estado de la carpeta de respaldos y reporta en Logger.
 * Útil para diagnóstico manual desde el editor GAS.
 */
function verificarEstadoRespaldos() {
  const iter = DriveApp.getFoldersByName(RESPALDO_CONFIG.CARPETA_RESPALDO_NOMBRE);
  if (!iter.hasNext()) {
    Logger.log('⚠️ Carpeta de respaldos no existe. Ejecuta crearCarpetaRespaldos() primero.');
    return;
  }
  const carpeta = iter.next();
  const archivos = [];
  const filesIter = carpeta.getFiles();
  while (filesIter.hasNext()) { archivos.push(filesIter.next()); }
  archivos.sort((a, b) => b.getDateCreated() - a.getDateCreated()); // descendente

  Logger.log('=== Estado de respaldos ===');
  Logger.log('Carpeta: ' + carpeta.getUrl());
  Logger.log('Total respaldos: ' + archivos.length);
  archivos.forEach((f, idx) => {
    Logger.log(`  ${idx + 1}. ${f.getName()} — ${Utilities.formatDate(f.getDateCreated(), 'GMT-6', 'dd/MM/yyyy HH:mm')}`);
  });

  // Verificar triggers activos
  const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'crearRespaldoSemanal');
  Logger.log('Triggers activos para crearRespaldoSemanal: ' + triggers.length);
}
