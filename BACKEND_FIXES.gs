// =========================================================================
// CORRECCIONES PARA EL BACKEND (Google Apps Script)
// Aplicar estos cambios en el editor de Apps Script
// =========================================================================

// =========================================================================
// FIX 1: getOrdenesSafe_ — agregar link_drive al response
// =========================================================================
// PROBLEMA: No devolvía link_drive, por lo que SEAINF nunca recibía el link
//           de la carpeta del cliente al cargar las OTs disponibles.
//
// ANTES (línea ~186):
//   .map(row => ({
//     ot: row[1], clienteInicial: row[5], clienteFinal: row[6], cliente: row[5]
//   }))
//
// DESPUÉS:
function getOrdenesSafe_() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const values = sheet.getDataRange().getDisplayValues();
  const ordenes = values.slice(1).filter(row => {
    const estatus = String(row[12] || '').trim().toUpperCase();
    return estatus !== 'ENTREGADO' && estatus !== 'FINALIZADO';
  }).map(row => ({
    ot: row[1],
    clienteInicial: row[5],
    clienteFinal: row[6],
    cliente: row[5],
    link_drive: row[13]  // ← AGREGADO: devolver el link de la carpeta
  })).filter(orden => orden.ot && orden.ot.trim() !== '');
  return { success: true, data: ordenes };
}


// =========================================================================
// FIX 2: fase3_CrearExpediente — usar linkDrive del payload como fallback
// =========================================================================
// PROBLEMA: Si la columna 14 de ORDENES_TRABAJO estaba vacía (porque al
//           registrar la OT no se tenía el link), el expediente se creaba
//           directamente en CONFIG.FOLDER_ID (raíz) en lugar de buscarlo
//           en la carpeta del cliente.
//
// DESPUÉS (reemplazar la función completa):
function fase3_CrearExpediente(payload) {
  const info = payload.data || {};
  const files = payload.files || [];

  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_OT);
  const values = sheet.getDataRange().getValues();

  let filaOT = -1;
  let linkCarpetaSucursal = '';

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]).trim() === String(info.ot).trim()) {
      filaOT = i + 1;
      linkCarpetaSucursal = values[i][13]; // Link de la columna 14 en la hoja
      break;
    }
  }

  if (filaOT === -1) return { success: false, error: 'OT no encontrada.' };

  // --- INICIO FIX: Cadena de fallback para encontrar la carpeta correcta ---
  // Prioridad: 1) Link en hoja OT → 2) linkDrive del payload → 3) Buscar por RFC → 4) CONFIG.FOLDER_ID
  let carpetaSucursal = null;

  // Intento 1: Desde la hoja ORDENES_TRABAJO (columna 14)
  if (linkCarpetaSucursal) {
    const folderIdMatch = linkCarpetaSucursal.match(/folders\/([a-zA-Z0-9_-]+)/);
    if (folderIdMatch) {
      try { carpetaSucursal = DriveApp.getFolderById(folderIdMatch[1]); } catch(e) {}
    }
  }

  // Intento 2: Desde el payload del frontend (linkDrive enviado por SEAINF)
  if (!carpetaSucursal && info.linkDrive) {
    const folderIdMatch2 = info.linkDrive.match(/folders\/([a-zA-Z0-9_-]+)/);
    if (folderIdMatch2) {
      try { carpetaSucursal = DriveApp.getFolderById(folderIdMatch2[1]); } catch(e) {}
    }
  }

  // Intento 3: Buscar en CLIENTES_MAESTRO por RFC (más reciente primero)
  if (!carpetaSucursal && info.rfc) {
    try {
      const sheetClientes = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_CLIENTES);
      const clientesData = sheetClientes.getDataRange().getValues();
      const rfcBuscado = String(info.rfc).toUpperCase().trim();
      const sucursalBuscada = String(info.sucursal || '').trim();

      for (let i = clientesData.length - 1; i >= 1; i--) {
        const rfcFila = String(clientesData[i][3]).toUpperCase().trim();
        const sucursalFila = String(clientesData[i][2]).trim();

        if (rfcFila === rfcBuscado && (!sucursalBuscada || sucursalFila === sucursalBuscada)) {
          const linkCliente = clientesData[i][22]; // columna "Link Drive (Carpeta Cliente)"
          if (linkCliente) {
            const folderIdMatch3 = linkCliente.match(/folders\/([a-zA-Z0-9_-]+)/);
            if (folderIdMatch3) {
              try { carpetaSucursal = DriveApp.getFolderById(folderIdMatch3[1]); } catch(e) {}
            }
          }
          break;
        }
      }
    } catch(e) {
      Logger.log('Error buscando carpeta por RFC: ' + e.message);
    }
  }

  // Intento 4: Fallback final a la carpeta raíz (último recurso)
  if (!carpetaSucursal) {
    carpetaSucursal = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    Logger.log('ADVERTENCIA: Expediente creado en carpeta raíz. OT: ' + info.ot);
  }
  // --- FIN FIX ---

  const consecutivoMatch = info.numInforme.match(/-(\d{4})$/);
  const consecutivoPrefix = consecutivoMatch ? consecutivoMatch[1] : '0000';
  const nombreCarpetaOT = `02_Expediente_${consecutivoPrefix}_${info.ot}_${info.nom}`;

  // Se crea el expediente técnico DENTRO de la carpeta de la Sucursal
  const carpetaOT = carpetaSucursal.createFolder(nombreCarpetaOT);

  const folders = {
    ORDEN_TRABAJO: carpetaOT.createFolder('1. ORDEN_TRABAJO'),
    HOJAS_CAMPO:   carpetaOT.createFolder('2. HDC'),
    CROQUIS:       carpetaOT.createFolder('3. CROQUIS'),
    FOTOS:         carpetaOT.createFolder('4. FOTOS')
  };

  files.forEach(file => {
    if (!file || !file.content) return;
    try {
      const decoded = Utilities.base64Decode(file.content);
      const blob = Utilities.newBlob(decoded, file.type, file.name);
      const targetFolder = folders[file.category] || carpetaOT;
      targetFolder.createFile(blob);
    } catch (err) { Logger.log('Error archivo: ' + err.message); }
  });

  sheet.getRange(filaOT, 4).setValue(info.numInforme);
  sheet.getRange(filaOT, 13).setValue('EN PROCESO');
  sheet.getRange(filaOT, 14).setValue(carpetaOT.getUrl());

  return { success: true, url: carpetaOT.getUrl() };
}
