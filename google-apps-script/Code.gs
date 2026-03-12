/**
 * Google Apps Script para detectar nuevos archivos en Google Drive
 * y disparar el GitHub Action automáticamente
 * 
 * CONFIGURACIÓN:
 * 1. Reemplaza GITHUB_TOKEN con tu Personal Access Token
 * 2. Reemplaza GITHUB_OWNER con tu usuario de GitHub
 * 3. Reemplaza GITHUB_REPO con el nombre de tu repositorio
 * 4. Reemplaza DRIVE_FOLDER_ID con el ID de tu carpeta de Drive
 */

// ============================================
// CONFIGURACIÓN - EDITAR ESTOS VALORES
// ============================================
const CONFIG = {
  GITHUB_TOKEN: 'ghp_XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX', // Tu GitHub Personal Access Token
  GITHUB_OWNER: 'tu-usuario',                              // Tu usuario de GitHub
  GITHUB_REPO: 'reportes-servicios-generales',             // Nombre del repo
  DRIVE_FOLDER_ID: 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX',    // ID de la carpeta de Drive
  CHECK_INTERVAL_MINUTES: 1                                 // Cada cuántos minutos revisar
};

// ============================================
// NO MODIFICAR DEBAJO DE ESTA LÍNEA
// ============================================

/**
 * Función principal que revisa si hay nuevos archivos
 * Configurar como trigger cada minuto
 */
function checkForNewFiles() {
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  const processedFiles = getProcessedFiles();
  let newFileFound = false;
  let latestFile = null;
  let latestTime = 0;
  
  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    const lastUpdated = file.getLastUpdated().getTime();
    
    // Encontrar el archivo más reciente
    if (lastUpdated > latestTime) {
      latestTime = lastUpdated;
      latestFile = file;
    }
    
    // Verificar si es nuevo o fue modificado
    if (!processedFiles[fileId] || processedFiles[fileId] < lastUpdated) {
      newFileFound = true;
    }
  }
  
  // También buscar archivos xlsx (nuevo formato)
  const xlsxFiles = folder.getFilesByType(MimeType.MICROSOFT_EXCEL_LEGACY);
  while (xlsxFiles.hasNext()) {
    const file = xlsxFiles.next();
    const fileId = file.getId();
    const lastUpdated = file.getLastUpdated().getTime();
    
    if (lastUpdated > latestTime) {
      latestTime = lastUpdated;
      latestFile = file;
    }
    
    if (!processedFiles[fileId] || processedFiles[fileId] < lastUpdated) {
      newFileFound = true;
    }
  }
  
  if (newFileFound && latestFile) {
    Logger.log('📊 Nuevo archivo detectado: ' + latestFile.getName());
    
    // Disparar GitHub Action
    const success = triggerGitHubAction(latestFile.getId(), latestFile.getName());
    
    if (success) {
      // Marcar archivo como procesado
      markFileAsProcessed(latestFile.getId(), latestFile.getLastUpdated().getTime());
      
      // Enviar notificación por email (opcional)
      sendNotification(latestFile.getName());
    }
  } else {
    Logger.log('✅ Sin cambios detectados');
  }
}

/**
 * Dispara el GitHub Action via repository_dispatch
 */
function triggerGitHubAction(fileId, fileName) {
  const url = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}/dispatches`;
  
  const payload = {
    event_type: 'generate-report',
    client_payload: {
      file_id: fileId,
      file_name: fileName,
      triggered_at: new Date().toISOString()
    }
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Accept': 'application/vnd.github.v3+json',
      'Authorization': `Bearer ${CONFIG.GITHUB_TOKEN}`,
      'Content-Type': 'application/json',
      'X-GitHub-Api-Version': '2022-11-28'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code === 204) {
      Logger.log('✅ GitHub Action disparado exitosamente');
      return true;
    } else {
      Logger.log('❌ Error al disparar GitHub Action: ' + code + ' - ' + response.getContentText());
      return false;
    }
  } catch (error) {
    Logger.log('❌ Error de conexión: ' + error.toString());
    return false;
  }
}

/**
 * Obtiene la lista de archivos ya procesados
 */
function getProcessedFiles() {
  const props = PropertiesService.getScriptProperties();
  const data = props.getProperty('processedFiles');
  return data ? JSON.parse(data) : {};
}

/**
 * Marca un archivo como procesado
 */
function markFileAsProcessed(fileId, timestamp) {
  const props = PropertiesService.getScriptProperties();
  const processedFiles = getProcessedFiles();
  processedFiles[fileId] = timestamp;
  props.setProperty('processedFiles', JSON.stringify(processedFiles));
}

/**
 * Envía notificación por email (opcional)
 */
function sendNotification(fileName) {
  try {
    const email = Session.getActiveUser().getEmail();
    const subject = '📊 Nuevo reporte en proceso';
    const body = `
Se detectó un nuevo archivo: ${fileName}

El reporte se está generando automáticamente.
Podrás verlo en: https://${CONFIG.GITHUB_OWNER}.github.io/${CONFIG.GITHUB_REPO}/

Este proceso toma aproximadamente 2-3 minutos.
    `;
    
    MailApp.sendEmail(email, subject, body);
    Logger.log('📧 Notificación enviada a: ' + email);
  } catch (error) {
    Logger.log('⚠️ No se pudo enviar notificación: ' + error.toString());
  }
}

/**
 * Función para probar la conexión con GitHub
 * Ejecutar manualmente para verificar configuración
 */
function testGitHubConnection() {
  const url = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}`;
  
  const options = {
    method: 'GET',
    headers: {
      'Accept': 'application/vnd.github.v3+json',
      'Authorization': `Bearer ${CONFIG.GITHUB_TOKEN}`,
      'X-GitHub-Api-Version': '2022-11-28'
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      Logger.log('✅ Conexión exitosa!');
      Logger.log('📁 Repositorio: ' + data.full_name);
      Logger.log('🔗 URL: ' + data.html_url);
      return true;
    } else {
      Logger.log('❌ Error: ' + code + ' - ' + response.getContentText());
      return false;
    }
  } catch (error) {
    Logger.log('❌ Error de conexión: ' + error.toString());
    return false;
  }
}

/**
 * Función para probar el acceso a la carpeta de Drive
 * Ejecutar manualmente para verificar configuración
 */
function testDriveAccess() {
  try {
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    Logger.log('✅ Acceso a carpeta exitoso!');
    Logger.log('📁 Nombre: ' + folder.getName());
    
    // Listar archivos Excel
    const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
    let count = 0;
    while (files.hasNext()) {
      const file = files.next();
      Logger.log('  📄 ' + file.getName() + ' (ID: ' + file.getId() + ')');
      count++;
    }
    Logger.log('📊 Total archivos Excel: ' + count);
    
    return true;
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    return false;
  }
}

/**
 * Función para resetear la lista de archivos procesados
 * Útil para forzar reprocesamiento
 */
function resetProcessedFiles() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('processedFiles');
  Logger.log('✅ Lista de archivos procesados reseteada');
}

/**
 * Dispara manualmente el proceso para el archivo más reciente
 */
function triggerManually() {
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  let latestFile = null;
  let latestTime = 0;
  
  while (files.hasNext()) {
    const file = files.next();
    const lastUpdated = file.getLastUpdated().getTime();
    if (lastUpdated > latestTime) {
      latestTime = lastUpdated;
      latestFile = file;
    }
  }
  
  if (latestFile) {
    Logger.log('📊 Procesando: ' + latestFile.getName());
    triggerGitHubAction(latestFile.getId(), latestFile.getName());
  } else {
    Logger.log('❌ No se encontraron archivos Excel');
  }
}
