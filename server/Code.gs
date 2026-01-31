
/**
 * @OnlyCurrentDoc
 * @AuthorizationRequired
 * @oauthScopes https://www.googleapis.com/auth/spreadsheets.currentonly, https://www.googleapis.com/auth/drive, https://www.googleapis.com/auth/script.external_request, https://www.googleapis.com/auth/userinfo.email
 */

// --- CONFIGURATION & CONSTANTS ---
// TODO: AFTER DEPLOYING AS WEB APP, PASTE THE URL HERE FOR EMAILS (API ENDPOINT)
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbymPQQO0C8Xf089bjAVIciWNbsr9DmS50odghFp7t_nh5ZqHGFe7HisbaFF-TqMPxPwwQ/exec'; 

// LINK DE ACCESO A LA PLATAFORMA (INTERFAZ VISUAL)
const PLATFORM_URL = 'https://aistudio.google.com/apps/drive/19BXTKPwCakVCf_-twNMayZXTxVDt6IY-?showAssistant=true&showPreview=true&fullscreenApplet=true';

const LOCK_WAIT_MS = 30000;
const SHEET_NAME_REQUESTS = 'Nueva Base Solicitudes';
const SHEET_NAME_MASTERS = 'MAESTROS';
const SHEET_NAME_RELATIONS = 'CDS vs UDEN';
const SHEET_NAME_INTEGRANTES = 'INTEGRANTES';

// DRIVE CONFIGURATION
const ROOT_DRIVE_FOLDER_ID = '1uaett_yH1qZcS-rVr_sUh73mODvX02im';

// ADMIN EMAIL CONFIGURATION
const ADMIN_EMAIL = 'dsanchez@equitel.com.co';

// HEADERS EXACTLY AS PROVIDED IN CSV + JSON COLUMNS
const HEADERS_REQUESTS = [
  "FECHA SOLICITUD", "EMPRESA", "CIUDAD ORIGEN", "CIUDAD DESTINO", "# ORDEN TRABAJO", 
  "# PERSONAS QUE VIAJAN", "CORREO ENCUESTADO", 
  "CÉDULA PERSONA 1", "NOMBRE PERSONA 1", "CÉDULA PERSONA 2", "NOMBRE PERSONA 2", 
  "CÉDULA PERSONA 3", "NOMBRE PERSONA 3", "CÉDULA PERSONA 4", "NOMBRE PERSONA 4", 
  "CÉDULA PERSONA 5", "NOMBRE PERSONA 5", 
  "CENTRO DE COSTOS", "VARIOS CENTROS COSTOS", "NOMBRE CENTRO DE COSTOS (AUTOMÁTICO)", 
  "UNIDAD DE NEGOCIO", "SEDE", "REQUIERE HOSPEDAJE", "NOMBRE HOTEL", "# NOCHES (AUTOMÁTICO)", 
  "FECHA IDA", "FECHA VUELTA", "HORA LLEGADA VUELO IDA", "HORA LLEGADA VUELO VUELTA", 
  "ID RESPUESTA", // Index 29 (Col AD)
  "APROBADO POR ÁREA?", "COSTO COTIZADO PARA VIAJE", "FECHA DE COMPRA DE TIQUETE", 
  "PERSONA QUE TRAMITA EL TIQUETE /HOTEL", "STATUS", "TIPO DE COMPRA DE TKT", 
  "FECHA DEL VUELO", "No RESERVA", "PROVEEDOR", "SERVICIO SOLICITADO", 
  "FECHA DE FACTURA", "# DE FACTURA", "TIPO DE TKT", "Q TKT", "DIAS DE ANTELACION TKT", 
  "VALOR PAGADO A AEROLINEA Y/O HOTEL", "VALOR PAGADO A AVIATUR Y/O IVA", 
  "TOTAL FACTURA", "PRESUPUESTO", "TARJETA DE CREDITO CON LA QUE SE HIZO LA COMPRA", 
  "OBSERVACIONES", "QUIÉN APRUEBA? (AUTOMÁTICO)", "APROBADO POR ÁREA? (AUTOMÁTICO)", 
  "FECHA/HORA (AUTOMÁTICO)", "CORREO DE QUIEN APRUEBA (AUTOMÁTICO)", "FECHASIMPLE_SOLICITUD",
  "OPCIONES (JSON)", "SELECCION (JSON)", "SOPORTES (JSON)", "CORREOS PASAJEROS (JSON)"
];

// --- SETUP FUNCTION ---
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_REQUESTS);
    sheet.getRange(1, 1, 1, HEADERS_REQUESTS.length).setValues([HEADERS_REQUESTS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS_REQUESTS.length).setFontWeight("bold").setBackground("#D71920").setFontColor("white");
  } else {
    // Check if new columns exist, if not, add headers
    const lastCol = sheet.getLastColumn();
    if (lastCol < HEADERS_REQUESTS.length) {
       sheet.getRange(1, 1, 1, HEADERS_REQUESTS.length).setValues([HEADERS_REQUESTS]);
    }
  }

  // Check for the specific relation sheet
  let relSheet = ss.getSheetByName(SHEET_NAME_RELATIONS);
  if (!relSheet) {
    relSheet = ss.insertSheet(SHEET_NAME_RELATIONS);
    relSheet.appendRow(["CENTRO COSTOS", "Descripcion del CC", "UNIDAD DE NEGOCIO"]);
  }

  // Check for INTEGRANTES sheet
  let intSheet = ss.getSheetByName(SHEET_NAME_INTEGRANTES);
  if (!intSheet) {
    intSheet = ss.insertSheet(SHEET_NAME_INTEGRANTES);
    // Updated header structure based on user request
    intSheet.appendRow([
        "Cedula Numero", "Apellidos y Nombres", "correo", "Empresa", "NDC", 
        "Centro de Costo", "Unidad", "sede", "cargo", "jefe Unidad", 
        "aprobador", "correo aprobador"
    ]);
  }
  
  return "Database Setup Complete.";
}

/**
 * Handle GET requests (Email Links)
 */
function doGet(e) {
  if (!e.parameter) return ContentService.createTextOutput("Equitel API Active.");

  const action = e.parameter.action;

  // 1. Handle Approval/Rejection by Approver
  if (action === 'approve') {
    return processApprovalFromEmail(e);
  }

  // 2. Handle Option Selection by Requester
  if (action === 'select') {
    return processOptionSelection(e);
  }
  
  // 3. Handle Standard API Calls (if used via GET)
  if (action) {
    const result = dispatch(action, e.parameter);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.TEXT);
  }

  return ContentService.createTextOutput("Equitel API Active. Use POST for operations.")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Handle POST requests (App API)
 */
function doPost(e) {
  try {
    // Check for postData validity
    if (!e.postData || !e.postData.contents) {
       throw new Error("Empty Request Body");
    }

    const data = JSON.parse(e.postData.contents);
    const result = dispatch(data.action, data.payload);
    
    // Using TEXT MimeType often helps with CORS issues in fetch
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      error: "Invalid Request: " + error.toString() 
    })).setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Main API Dispatcher
 */
function dispatch(action, payload) {
  const isWriteAction = ['createRequest', 'updateRequest', 'uploadSupportFile', 'closeRequest'].includes(action);
  const lock = LockService.getScriptLock();

  let currentUserEmail = '';
  if (payload && payload.userEmail) {
    currentUserEmail = String(payload.userEmail).trim().toLowerCase();
  } else {
    currentUserEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  }

  try {
    if (isWriteAction) {
      const hasLock = lock.tryLock(LOCK_WAIT_MS);
      if (!hasLock) {
        return { 
          success: false, 
          error: 'El sistema está ocupado procesando otra solicitud. Por favor intente de nuevo en unos segundos.' 
        };
      }
    }

    let result;
    switch (action) {
      case 'getCurrentUser':
        result = currentUserEmail;
        break;
      case 'getCostCenterData':
        result = getCostCenterData();
        break;
      case 'getIntegrantesData':
        result = getIntegrantesData();
        break;
      case 'getMyRequests':
        result = getRequestsByEmail(currentUserEmail);
        break;
      case 'getAllRequests':
        if(!isUserAnalyst(currentUserEmail)) {
           result = getRequestsByEmail(currentUserEmail);
        } else {
           result = getAllRequests();
        }
        break;
      case 'createRequest':
        result = createNewRequest(payload);
        break;
      case 'updateRequest':
        result = updateRequestStatus(payload.id, payload.status, payload.payload);
        break;
      case 'uploadSupportFile':
        result = uploadSupportFile(payload.requestId, payload.fileData, payload.fileName, payload.mimeType);
        break;
      case 'closeRequest':
        result = updateRequestStatus(payload.requestId, 'PROCESADO');
        break;
      default:
        return { success: false, error: 'Acción desconocida: ' + action };
    }
    
    return { success: true, data: result };

  } catch (e) {
    console.error("Error in dispatch: " + e.toString());
    return { success: false, error: e.toString() };
  } finally {
    if (isWriteAction) {
      lock.releaseLock();
    }
  }
}

// --- BUSINESS LOGIC ---

function getRequestsByEmail(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Fetch data safely up to the last column containing data
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const emailIdx = HEADERS_REQUESTS.indexOf("CORREO ENCUESTADO");
  
  if (emailIdx === -1) return [];

  const targetEmail = String(email).toLowerCase().trim();

  // Filter in memory
  const userRequests = data.filter(row => {
    // Check if row has data at email index
    if (emailIdx >= row.length) return false;
    const rowEmail = String(row[emailIdx]).toLowerCase().trim();
    return rowEmail === targetEmail;
  });

  return userRequests.map(mapRowToRequest).reverse();
}

function getAllRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  return data.map(mapRowToRequest).reverse();
}

function getCostCenterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return [];

  const sheet = ss.getSheetByName(SHEET_NAME_RELATIONS);
  
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Assuming columns: A=Code, B=Name, C=Business Unit
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  
  return data.map(row => {
    let code = String(row[0]).trim();
    if (code.match(/^\d+$/)) {
      code = code.padStart(4, '0');
    }
    return {
      code: code,
      name: String(row[1]),
      businessUnit: String(row[2])
    };
  }).filter(item => item.code && item.businessUnit);
}

function getIntegrantesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_INTEGRANTES);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // MAPPING based on user structure:
  // A (0): Cedula Numero
  // B (1): Apellidos y Nombres
  // C (2): correo
  // ...
  // K (10): aprobador
  // L (11): correo aprobador
  
  const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

  return data.map(row => ({
    idNumber: String(row[0]).trim(),
    name: String(row[1]),
    email: String(row[2]).toLowerCase().trim(),
    approverName: String(row[10]), // Column K
    approverEmail: String(row[11]).toLowerCase().trim() // Column L
  })).filter(i => i.idNumber && i.name);
}

function createNewRequest(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("No Container Spreadsheet found");
  
  const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
  if (!sheet) throw new Error("Base de datos no inicializada.");

  // --- SEQUENTIAL ID GENERATION LOGIC ---
  const idColIndex = HEADERS_REQUESTS.indexOf("ID RESPUESTA") + 1; // 1-based index
  const lastRow = sheet.getLastRow();
  let nextIdNum = 1;

  if (lastRow > 1) {
    const existingIds = sheet.getRange(2, idColIndex, lastRow - 1, 1).getValues().flat();
    const numericIds = existingIds
      .map(val => {
         // Clean val: remove 'SOL-' if present, to handle both "1" and "SOL-000001"
         const strVal = String(val).replace(/^SOL-/, '');
         return parseInt(strVal, 10);
      })
      .filter(val => !isNaN(val));

    if (numericIds.length > 0) {
      nextIdNum = Math.max(...numericIds) + 1;
    }
  }
  
  const id = `SOL-${nextIdNum.toString().padStart(6, '0')}`; 
  // ---------------------------------------

  // --- FETCH COST CENTER NAME ---
  let ccName = '';
  if (data.costCenter && data.costCenter !== 'VARIOS') {
     const masters = getCostCenterData();
     const ccObj = masters.find(m => m.code == data.costCenter);
     if (ccObj) ccName = ccObj.name;
  }
  data.costCenterName = ccName;
  // ---------------------------------------

  // --- DETERMINING APPROVER (BASED ON 1ST PASSENGER ID) ---
  let approverEmail = ADMIN_EMAIL; // Default
  
  if (data.passengers && data.passengers.length > 0) {
     const firstPassengerId = data.passengers[0].idNumber;
     const integrantes = getIntegrantesData();
     const integrant = integrantes.find(i => i.idNumber === firstPassengerId);
     
     if (integrant && integrant.approverEmail) {
        approverEmail = integrant.approverEmail;
     } else {
        // Fallback to Cost Center Logic (Original logic if not found in Integrantes)
        approverEmail = findApprover(data.costCenter);
     }
  }
  // ---------------------------------------

  // Nights Calculation
  let nights = 0;
  if (data.requiresHotel) {
      if (data.nights && data.nights > 0) {
          nights = data.nights;
      } else if (data.departureDate && data.returnDate) {
          const d1 = new Date(data.departureDate);
          const d2 = new Date(data.returnDate);
          const diffTime = Math.abs(d2 - d1);
          nights = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      }
  }

  const row = new Array(HEADERS_REQUESTS.length).fill('');

  row[0] = new Date(); // FECHA SOLICITUD
  row[1] = data.company; 
  row[2] = data.origin; 
  row[3] = data.destination; 
  row[4] = data.workOrder || ''; 
  row[5] = data.passengers ? data.passengers.length : 1; 
  row[6] = String(data.requesterEmail).toLowerCase().trim(); 

  const p = data.passengers || [];
  for(let i=0; i<5; i++) {
    const baseIdx = 7 + (i*2);
    row[baseIdx] = p[i] ? p[i].idNumber : '';
    row[baseIdx+1] = p[i] ? p[i].name : '';
  }

  row[17] = data.costCenter; 
  row[18] = data.variousCostCenters || '';
  row[19] = ccName; 

  row[20] = data.businessUnit; 
  row[21] = data.site; 
  row[22] = data.requiresHotel ? 'Sí' : 'No';
  row[23] = data.hotelName || '';
  row[24] = nights;

  row[25] = data.departureDate;
  row[26] = data.returnDate || ''; 
  row[27] = data.departureTimePreference || '';
  row[28] = data.returnTimePreference || '';

  row[29] = id; 
  
  // STATUS
  const statusIdx = HEADERS_REQUESTS.indexOf("STATUS");
  if (statusIdx > -1) row[statusIdx] = 'PENDIENTE_OPCIONES';

  // OBSERVACIONES
  const obsIdx = HEADERS_REQUESTS.indexOf("OBSERVACIONES");
  if (obsIdx > -1) row[obsIdx] = data.comments || '';

  // APPROVER
  const approverIdx = HEADERS_REQUESTS.indexOf("CORREO DE QUIEN APRUEBA (AUTOMÁTICO)");
  if (approverIdx > -1) row[approverIdx] = approverEmail;

  // PASSENGER EMAILS JSON
  const emailsIdx = HEADERS_REQUESTS.indexOf("CORREOS PASAJEROS (JSON)");
  if (emailsIdx > -1) {
     const pEmails = data.passengers.map(p => p.email).filter(e => e);
     row[emailsIdx] = JSON.stringify(pEmails);
  }

  sheet.appendRow(row);

  // --- NOTIFY ADMIN ---
  // We attach passengers info to data for the email function
  data.approverEmail = approverEmail;
  sendNewRequestNotification(data, id);

  return id;
}

function updateRequestStatus(id, status, payload) {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
   if (!sheet) throw new Error("Base de datos no encontrada");

   const lastRow = sheet.getLastRow();
   if (lastRow < 2) throw new Error("No hay datos");

   const idIdx = HEADERS_REQUESTS.indexOf("ID RESPUESTA");
   const statusIdx = HEADERS_REQUESTS.indexOf("STATUS");
   const emailIdx = HEADERS_REQUESTS.indexOf("CORREO ENCUESTADO");

   const ids = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().flat();
   const rowIndex = ids.map(String).indexOf(String(id));

   if (rowIndex === -1) throw new Error("ID de solicitud no encontrado: " + id);
   const rowNumber = rowIndex + 2;

   // Update Status
   sheet.getRange(rowNumber, statusIdx + 1).setValue(status);

   // Update Payload Columns (Options or Selection)
   if (payload) {
      if (payload.analystOptions) {
         const optIdx = HEADERS_REQUESTS.indexOf("OPCIONES (JSON)");
         if (optIdx > -1) {
             const val = typeof payload.analystOptions === 'object' ? JSON.stringify(payload.analystOptions) : payload.analystOptions;
             sheet.getRange(rowNumber, optIdx + 1).setValue(val);
         }
      }
      if (payload.selectedOption) {
         const selIdx = HEADERS_REQUESTS.indexOf("SELECCION (JSON)");
         if (selIdx > -1) {
             const val = typeof payload.selectedOption === 'object' ? JSON.stringify(payload.selectedOption) : payload.selectedOption;
             sheet.getRange(rowNumber, selIdx + 1).setValue(val);
         }
      }
   }

   // --- TRIGGERS BASED ON NEW STATUS ---

   // 1. If status became PENDIENTE_SELECCION, email the Requester
   if (status === 'PENDIENTE_SELECCION' && payload && payload.analystOptions) {
      const requesterEmail = sheet.getRange(rowNumber, emailIdx + 1).getValue();
      const fullRequest = mapRowToRequest(sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0]);
      sendOptionsToRequester(requesterEmail, fullRequest, payload.analystOptions);
   }

   // 2. If status became APROBADO or DENEGADO, email Requester + CC Admin with FULL DETAILS
   if (status === 'APROBADO' || status === 'DENEGADO') {
      const fullRequest = mapRowToRequest(sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0]);
      sendDecisionNotification(fullRequest, status);
   }

   return true;
}

function uploadSupportFile(requestId, fileData, fileName, mimeType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
  if (!sheet) throw new Error("Base no encontrada");

  const idIdx = HEADERS_REQUESTS.indexOf("ID RESPUESTA");
  const lastRow = sheet.getLastRow();
  const ids = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().flat();
  const rowIndex = ids.map(String).indexOf(String(requestId));

  if (rowIndex === -1) throw new Error("Solicitud no encontrada");
  const rowNumber = rowIndex + 2;

  // 1. Get Support Column Index
  const supportIdx = HEADERS_REQUESTS.indexOf("SOPORTES (JSON)");
  if (supportIdx === -1) throw new Error("Columna SOPORTES no configurada");

  // 2. Read Existing Data
  const jsonStr = sheet.getRange(rowNumber, supportIdx + 1).getValue();
  let supportData = jsonStr ? JSON.parse(jsonStr) : { folderId: null, folderUrl: null, files: [] };

  // 3. Handle Drive Folder - Added robust error handling and permission check
  let folder;
  try {
      if (supportData.folderId) {
         try {
           folder = DriveApp.getFolderById(supportData.folderId);
         } catch(e) { console.warn("Saved folder ID invalid or inaccessible, searching/creating new."); }
      }

      if (!folder) {
         const rootFolder = DriveApp.getFolderById(ROOT_DRIVE_FOLDER_ID);
         const folderName = `${requestId}`; 
         
         const folders = rootFolder.getFoldersByName(folderName);
         if (folders.hasNext()) {
            folder = folders.next();
         } else {
            folder = rootFolder.createFolder(folderName);
         }
         
         supportData.folderId = folder.getId();
         supportData.folderUrl = folder.getUrl();
      }
  } catch(e) {
      throw new Error("Error accediendo a Google Drive. Por favor verifique los permisos del Script o la ID de la carpeta raíz. Detalles: " + e.toString());
  }

  // 4. Create File from Base64
  const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
  const file = folder.createFile(blob);
  
  // 5. Update JSON
  const newFileEntry = {
    id: file.getId(),
    name: file.getName(),
    url: file.getUrl(),
    mimeType: mimeType,
    date: new Date().toISOString()
  };
  
  supportData.files.push(newFileEntry);

  // 6. Save back to Sheet
  sheet.getRange(rowNumber, supportIdx + 1).setValue(JSON.stringify(supportData));

  return supportData;
}

// --- EMAILS & WORKFLOW ---

function getCCList(request) {
    const requester = request.requesterEmail;
    // Extract passenger emails from passengers array
    const passengerEmails = (request.passengers || [])
        .map(p => p.email)
        .filter(e => e && e.toLowerCase() !== requester.toLowerCase()); // Avoid duplicate if requester is passenger
    
    // Always include Admin in CC if needed, but usually Admin is TO or BCC. 
    // Here we just return passenger list.
    return passengerEmails.join(',');
}

function sendNewRequestNotification(data, requestId) {
    const subject = `Solicitud de Viaje ${requestId} - ${data.requesterEmail} - ${data.company} ${data.site}`;
    const link = PLATFORM_URL; 

    // Build passenger list
    const passengersList = (data.passengers || []).map(p => 
      `<li style="margin-bottom: 2px;">${p.name} <span style="color:#666; font-size:0.9em;">(${p.idNumber})</span></li>`
    ).join('');

    let costCenterDisplay = data.costCenter;
    if (data.costCenter === 'VARIOS' && data.variousCostCenters) {
        costCenterDisplay = `VARIOS (${data.variousCostCenters})`;
    } else if (data.costCenterName) {
        costCenterDisplay = `${data.costCenter} - ${data.costCenterName}`;
    }
    
    const isOneWay = !data.returnDate || data.returnDate === '';

    const htmlBody = `
      <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden; background-color: #fff;">
         <!-- Header -->
         <div style="background-color: #D71920; color: white; padding: 20px; text-align: center;">
            <h2 style="margin: 0; font-size: 24px;">Nueva Solicitud de Viaje/Hospedaje</h2>
            <p style="margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;">ID: <strong>${requestId}</strong></p>
         </div>
         <!-- Content -->
         <div style="padding: 25px;">
            <p style="margin-top: 0; font-size: 14px; color: #666;">
               Se ha registrado un nuevo requerimiento de viaje para <strong>${data.requesterEmail}</strong>.
            </p>
            <!-- Route Box -->
            <div style="background-color: #f8f9fa; border-left: 4px solid #D71920; padding: 15px; margin: 20px 0; border-radius: 4px;">
               <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                     <td width="45%" style="vertical-align: top;">
                        <span style="font-size: 11px; color: #999; text-transform: uppercase; letter-spacing: 0.5px;">Origen</span><br>
                        <strong style="font-size: 16px;">${data.origin}</strong>
                     </td>
                     <td width="10%" style="vertical-align: middle; text-align: center; color: #ccc;">➝</td>
                     <td width="45%" style="text-align: right; vertical-align: top;">
                        <span style="font-size: 11px; color: #999; text-transform: uppercase; letter-spacing: 0.5px;">Destino</span><br>
                        <strong style="font-size: 16px;">${data.destination}</strong>
                     </td>
                  </tr>
               </table>
               <div style="margin-top: 15px; border-top: 1px solid #e0e0e0; padding-top: 10px;">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="50%" style="text-align: center; border-right: ${isOneWay ? 'none' : '1px solid #e0e0e0'}; padding-right: 10px;">
                        <span style="font-size: 10px; color: #999; text-transform: uppercase; letter-spacing: 0.5px;">Fecha Ida</span><br>
                        <span style="font-size: 14px; font-weight: bold; color: #333;">📅 ${data.departureDate}</span>
                        ${data.departureTimePreference ? `<br><span style="font-size: 11px; color: #666;">(${data.departureTimePreference})</span>` : ''}
                      </td>
                      ${!isOneWay ? `
                      <td width="50%" style="text-align: center; padding-left: 10px;">
                        <span style="font-size: 10px; color: #999; text-transform: uppercase; letter-spacing: 0.5px;">Fecha Regreso</span><br>
                        <span style="font-size: 14px; font-weight: bold; color: #333;">📅 ${data.returnDate}</span>
                         ${data.returnTimePreference ? `<br><span style="font-size: 11px; color: #666;">(${data.returnTimePreference})</span>` : ''}
                      </td>` : `
                      <td width="50%" style="text-align: center; padding-left: 10px;">
                         <span style="font-size: 12px; color: #999; font-style: italic;">(Solo Ida)</span>
                      </td>
                      `}
                    </tr>
                  </table>
               </div>
            </div>
            <!-- Details -->
            <div style="margin-bottom: 20px;">
               <h4 style="margin: 0 0 10px 0; color: #333; border-bottom: 1px solid #eee; padding-bottom: 5px;">Detalles del Caso</h4>
               <table width="100%" style="font-size: 14px;">
                  <tr>
                     <td style="padding: 5px 0; color: #666;" width="40%">Empresa / Sede:</td>
                     <td style="padding: 5px 0; font-weight: bold;">${data.company} - ${data.site}</td>
                  </tr>
                  <tr>
                     <td style="padding: 5px 0; color: #666;">Centro de Costos:</td>
                     <td style="padding: 5px 0; font-weight: bold;">${costCenterDisplay}</td>
                  </tr>
                  <tr>
                     <td style="padding: 5px 0; color: #666;">Aprobador:</td>
                     <td style="padding: 5px 0; font-weight: bold; color: #d32f2f;">${data.approverEmail}</td>
                  </tr>
                  <tr>
                     <td style="padding: 5px 0; color: #666;">Hospedaje:</td>
                     <td style="padding: 5px 0; font-weight: bold;">
                        ${data.requiresHotel ? `<span style="color: #0d47a1;">SÍ - ${data.hotelName}</span> <span style="font-size:0.9em;color:#666;">(${data.nights} Noches)</span>` : 'NO'}
                     </td>
                  </tr>
               </table>
            </div>
            ${data.comments ? `
            <div style="background-color: #fff8e1; border: 1px solid #ffe082; color: #5d4037; padding: 12px; border-radius: 4px; font-size: 13px; margin-bottom: 20px;">
               <strong>📝 Observaciones / Notas:</strong>
               <p style="margin: 5px 0 0 0; white-space: pre-wrap;">${data.comments}</p>
            </div>
            ` : ''}
            <div style="background-color: #e3f2fd; color: #0d47a1; padding: 12px; border-radius: 4px; font-size: 13px;">
               <strong>👥 Pasajero(s) (${(data.passengers || []).length}):</strong>
               <ul style="margin: 5px 0 0 0; padding-left: 20px;">
                  ${passengersList}
               </ul>
            </div>
            <div style="text-align: center; margin-top: 30px;">
               <a href="${link}" style="display: inline-block; background-color: #333; color: white; padding: 12px 24px; text-decoration: none; border-radius: 50px; font-weight: bold; font-size: 14px; box-shadow: 0 2px 5px rgba(0,0,0,0.2);">
                  Ingresar a la Plataforma
               </a>
            </div>
         </div>
         <div style="background-color: #f4f4f4; padding: 15px; text-align: center; font-size: 11px; color: #999;">
            TravelMaster Notifications • ${new Date().getFullYear()}
         </div>
      </div>
    `;

    // Construct CC List: Requester + Passengers
    const ccEmails = [data.requesterEmail, getCCList(data)].filter(e => e).join(',');

    try {
        MailApp.sendEmail({
            to: ADMIN_EMAIL,
            cc: ccEmails,
            subject: subject,
            htmlBody: htmlBody
        });
    } catch (e) {
        console.error("Error sending admin notification: " + e.toString());
    }
}

function sendOptionsToRequester(recipient, request, options) {
   const subject = `Solicitud de Viaje ${request.requestId} - ${request.requesterEmail} - ${request.company} ${request.site}`;
   
   let optionsHtml = '';
   options.forEach(opt => {
      const selectLink = `${WEB_APP_URL}?action=select&id=${request.requestId}&optionId=${opt.id}`;
      
      const outFlightNum = opt.outbound.flightNumber ? ` (${opt.outbound.flightNumber})` : '';
      const inFlightNum = (opt.inbound && opt.inbound.flightNumber) ? ` (${opt.inbound.flightNumber})` : '';

      let flightHtml = '';
      flightHtml += `<div><strong>Ida:</strong> ${opt.outbound.airline}${outFlightNum} - ${opt.outbound.flightTime} <br/> <small>${opt.outbound.notes}</small></div>`;
      if (opt.inbound && opt.inbound.airline) {
         flightHtml += `<div style="margin-top:5px;"><strong>Vuelta:</strong> ${opt.inbound.airline}${inFlightNum} - ${opt.inbound.flightTime} <br/> <small>${opt.inbound.notes}</small></div>`;
      }

      optionsHtml += `
      <div style="border: 1px solid #ccc; padding: 15px; margin-bottom: 15px; border-radius: 8px; background-color: #f9f9f9;">
         <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-bottom: 1px solid #eee; padding-bottom: 10px; margin-bottom: 10px;">
           <tr>
             <td align="left" valign="middle">
                <h3 style="margin: 0; color: #D71920;">Opción ${opt.id}</h3>
             </td>
             <td align="right" valign="middle">
                <span style="font-size: 1.2em; font-weight: bold; white-space: nowrap;">$ ${Number(opt.totalPrice).toLocaleString()}</span>
             </td>
           </tr>
         </table>
         ${flightHtml}
         ${opt.hotel ? `<div style="margin-top:10px; color: #0056b3;"><strong>Hotel:</strong> ${opt.hotel}</div>` : ''}
         <div style="margin-top: 15px; text-align: right;">
             <a href="${selectLink}" style="background-color: #D71920; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; font-weight: bold;">Seleccionar esta Opción</a>
         </div>
      </div>
      `;
   });

   const htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px;">
      <h2 style="color: #333;">Opciones Disponibles</h2>
      <p>Se han cargado las opciones de viaje para su solicitud <strong>${request.requestId}</strong> (${request.origin} - ${request.destination}).</p>
      <p>Por favor revise las siguientes alternativas y haga clic en "Seleccionar" en la que prefiera. Una vez seleccionada, se enviará a aprobación.</p>
      <div style="margin-top: 20px;">
        ${optionsHtml}
      </div>
      ${request.comments ? `
        <p style="font-size: 0.85em; color: #666; margin-top: 20px; border-top: 1px solid #eee; padding-top: 10px;">
           <em>Notas originales: ${request.comments}</em>
        </p>
      ` : ''}
      <p style="font-size: 0.9em; color: #777; margin-top: 30px;">Si ninguna opción se ajusta, por favor contacte al analista de viajes.</p>
    </div>
   `;

   // CC Passengers
   const ccList = getCCList(request);

   try {
     MailApp.sendEmail({
       to: recipient,
       cc: ccList,
       subject: subject,
       htmlBody: htmlBody
     });
   } catch (e) {
     console.error("Error sending options email: " + e.toString());
   }
}

function sendDecisionNotification(request, status) {
  const subject = `Solicitud de Viaje ${request.requestId} - ${request.requesterEmail} - ${request.company} ${request.site}`;
  
  const isApproved = status === 'APROBADO';
  const headerColor = isApproved ? '#28a745' : '#dc3545'; 
  const headerTitle = isApproved ? 'SOLICITUD APROBADA' : 'SOLICITUD DENEGADA';
  const headerMsg = isApproved 
     ? 'Su viaje ha sido autorizado. El equipo de compras procederá con la emisión.' 
     : 'Su solicitud ha sido denegada por el aprobador.';

  let costCenterDisplay = request.costCenter;
  if (request.costCenter === 'VARIOS' && request.variousCostCenters) {
      costCenterDisplay = `VARIOS (${request.variousCostCenters})`;
  } else if (request.costCenterName) {
      costCenterDisplay = `${request.costCenter} - ${request.costCenterName}`;
  }

  const passengersList = (request.passengers || []).map(p => 
      `<li style="margin-bottom: 2px;">${p.name} <span style="color:#666; font-size:0.9em;">(CC: ${p.idNumber})</span></li>`
  ).join('');

  const selectedOption = request.selectedOption;
  if (!selectedOption) return; 

  const outFlightNum = selectedOption.outbound.flightNumber ? ` <span style="font-family:monospace; background:#eee; padding:2px 4px; border-radius:3px;">${selectedOption.outbound.flightNumber}</span>` : '';
  const inFlightNum = (selectedOption.inbound && selectedOption.inbound.flightNumber) ? ` <span style="font-family:monospace; background:#eee; padding:2px 4px; border-radius:3px;">${selectedOption.inbound.flightNumber}</span>` : '';

  const isOneWay = !request.returnDate || request.returnDate === '';

  const htmlBody = `
    <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden; background-color: #fff; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
       <div style="background-color: ${headerColor}; color: white; padding: 25px 20px; text-align: center;">
          <h2 style="margin: 0; font-size: 26px; text-transform: uppercase; letter-spacing: 1px;">${headerTitle}</h2>
          <p style="margin: 8px 0 0 0; font-size: 14px; opacity: 0.9;">${headerMsg}</p>
       </div>
       <div style="padding: 30px;">
          <h3 style="color: #444; border-bottom: 2px solid #eee; padding-bottom: 5px; font-size: 16px; margin-top: 0;">1. Resumen de la Solicitud (${request.requestId})</h3>
          <div style="background-color: #f8f9fa; border-radius: 6px; padding: 15px; margin-bottom: 15px;">
             <table width="100%" border="0">
                 <tr>
                    <td width="45%" style="vertical-align: top;">
                        <span style="font-size: 10px; color: #888; text-transform: uppercase;">Origen</span><br>
                        <strong style="font-size: 18px; color: #222;">${request.origin}</strong>
                    </td>
                    <td width="10%" style="text-align: center; vertical-align: middle; color: #ccc; font-size: 20px;">➝</td>
                    <td width="45%" style="text-align: right; vertical-align: top;">
                        <span style="font-size: 10px; color: #888; text-transform: uppercase;">Destino</span><br>
                        <strong style="font-size: 18px; color: #222;">${request.destination}</strong>
                    </td>
                 </tr>
             </table>
             <div style="margin-top: 15px; border-top: 1px solid #e9ecef; padding-top: 10px;">
                 <table width="100%">
                    <tr>
                       <td width="50%" style="font-size: 13px;">
                          <span style="color: #666;">Salida:</span> <strong style="color: #D71920;">${request.departureDate}</strong>
                          ${request.departureTimePreference ? `<br><span style="font-size:11px; color:#888;">Pref: ${request.departureTimePreference}</span>` : ''}
                       </td>
                       <td width="50%" style="font-size: 13px; text-align: right;">
                          ${!isOneWay 
                             ? `<span style="color: #666;">Regreso:</span> <strong style="color: #D71920;">${request.returnDate}</strong>` 
                             : '<span style="color: #888; font-style: italic;">(Solo Ida)</span>'}
                          ${(!isOneWay && request.returnTimePreference) ? `<br><span style="font-size:11px; color:#888;">Pref: ${request.returnTimePreference}</span>` : ''}
                       </td>
                    </tr>
                 </table>
             </div>
          </div>
          <table width="100%" style="font-size: 13px; color: #555; margin-bottom: 15px;">
              <tr>
                 <td style="padding-bottom: 4px;" width="100"><strong>Pasajeros (${request.passengers.length}):</strong></td>
                 <td style="padding-bottom: 4px;">
                    <ul style="margin: 0; padding-left: 15px;">${passengersList}</ul>
                 </td>
              </tr>
              <tr>
                 <td style="padding: 4px 0;"><strong>Centro Costos:</strong></td>
                 <td style="padding: 4px 0;">${costCenterDisplay}</td>
              </tr>
              <tr>
                 <td style="padding: 4px 0;"><strong>Empresa/Sede:</strong></td>
                 <td style="padding: 4px 0;">${request.company} / ${request.site}</td>
              </tr>
              ${request.comments ? `
              <tr>
                 <td style="padding-top: 8px; vertical-align: top;"><strong>Notas Usuario:</strong></td>
                 <td style="padding-top: 8px; font-style: italic; color: #d68b00; background: #fff8e1; padding: 5px;">"${request.comments}"</td>
              </tr>` : ''}
          </table>
          <div style="margin-top: 30px; margin-bottom: 10px;">
             <h3 style="color: ${headerColor}; border-bottom: 2px solid ${headerColor}; padding-bottom: 5px; font-size: 16px; margin-top: 0; margin-bottom: 15px;">
                2. Detalle de la Opción (${selectedOption.id})
             </h3>
             <div style="border: 2px solid ${headerColor}; background-color: #fff; border-radius: 8px; overflow: hidden;">
                <div style="background-color: ${headerColor}; color: white; padding: 10px 15px; text-align: right;">
                    <span style="font-size: 12px; opacity: 0.9; margin-right: 10px;">COSTO TOTAL:</span>
                    <strong style="font-size: 20px;">$ ${Number(selectedOption.totalPrice).toLocaleString()}</strong>
                </div>
                <div style="padding: 15px;">
                    <div style="margin-bottom: 12px; padding-bottom: 12px; border-bottom: 1px dashed #eee;">
                        <div style="font-size: 12px; color: #888; text-transform: uppercase; margin-bottom: 4px;">✈️ Vuelo de Ida (${request.departureDate})</div>
                        <div style="font-size: 15px; color: #333;">
                            <strong>${selectedOption.outbound.airline}</strong> ${outFlightNum}
                        </div>
                        <div style="font-size: 13px; color: #555; margin-top: 2px;">
                            Hora Salida: <strong>${selectedOption.outbound.flightTime}</strong>
                        </div>
                        <div style="font-size: 12px; color: #777; margin-top: 2px;">
                            <em>Nota: ${selectedOption.outbound.notes}</em>
                        </div>
                    </div>
                    ${selectedOption.inbound ? `
                    <div style="margin-bottom: 12px; padding-bottom: 12px; border-bottom: 1px dashed #eee;">
                        <div style="font-size: 12px; color: #888; text-transform: uppercase; margin-bottom: 4px;">✈️ Vuelo de Regreso (${request.returnDate})</div>
                        <div style="font-size: 15px; color: #333;">
                            <strong>${selectedOption.inbound.airline}</strong> ${inFlightNum}
                        </div>
                        <div style="font-size: 13px; color: #555; margin-top: 2px;">
                            Hora Salida: <strong>${selectedOption.inbound.flightTime}</strong>
                        </div>
                        <div style="font-size: 12px; color: #777; margin-top: 2px;">
                            <em>Nota: ${selectedOption.inbound.notes}</em>
                        </div>
                    </div>
                    ` : ''}
                    ${selectedOption.hotel ? `
                    <div style="background-color: #e3f2fd; padding: 10px; border-radius: 4px; margin-top: 10px;">
                        <div style="font-size: 12px; color: #0d47a1; text-transform: uppercase; margin-bottom: 4px;">🏨 Hospedaje Incluido</div>
                        <div style="font-size: 14px; font-weight: bold; color: #000;">${selectedOption.hotel}</div>
                        <div style="font-size: 12px; color: #555;">${request.nights} Noche(s) calculadas</div>
                    </div>
                    ` : '<div style="font-size: 12px; color: #999; margin-top: 10px; font-style: italic;">No incluye hospedaje</div>'}
                </div>
             </div>
          </div>
       </div>
       <div style="background-color: #333; color: #ccc; padding: 15px; text-align: center; font-size: 11px;">
          Gestión de Viajes Corporativos • ${new Date().getFullYear()}
       </div>
    </div>
  `;

  // CC Passengers
  const ccList = [ADMIN_EMAIL, getCCList(request)].filter(e => e).join(',');

  try {
    MailApp.sendEmail({
      to: request.requesterEmail,
      cc: ccList,
      subject: subject,
      htmlBody: htmlBody
    });
  } catch(e) {
    console.error("Error sending decision email", e);
  }
}

function mapRowToRequest(row) {
  const getValue = (headerName) => {
    const idx = HEADERS_REQUESTS.indexOf(headerName);
    if (idx === -1) return ''; // Handles case where header isn't found
    if (idx >= row.length) return ''; // Handles short rows
    const val = row[idx];
    return (val === undefined || val === null) ? '' : val;
  };

  const safeDate = (val) => {
    if (!val) return '';
    if (val instanceof Date) return val.toISOString().split('T')[0];
    const s = String(val);
    if(s.includes('T')) return s.split('T')[0];
    return s;
  };

  // Parse emails JSON if available
  let passengerEmails = [];
  try {
      const pEmailsStr = getValue("CORREOS PASAJEROS (JSON)");
      if (pEmailsStr) passengerEmails = JSON.parse(pEmailsStr);
  } catch(e) {}

  const passengers = [];
  for(let i=1; i<=5; i++) {
     const name = getValue(`NOMBRE PERSONA ${i}`);
     const id = getValue(`CÉDULA PERSONA ${i}`);
     if(name && String(name).trim() !== '') {
       // Try to assign email by index if available
       const pEmail = passengerEmails[i-1] || '';
       passengers.push({ name: String(name), idNumber: String(id), email: pEmail });
     }
  }

  // Parse JSON fields safely
  let analystOptions = []; 
  let selectedOption = null;
  let supportData = undefined;

  try {
    const rawOpt = getValue("OPCIONES (JSON)");
    if (rawOpt && rawOpt !== '') analystOptions = JSON.parse(rawOpt); 
    
    const rawSel = getValue("SELECCION (JSON)");
    if (rawSel && rawSel !== '') selectedOption = JSON.parse(rawSel);

    const rawSup = getValue("SOPORTES (JSON)");
    if (rawSup && rawSup !== '') supportData = JSON.parse(rawSup);

  } catch(e) { console.error("Error parsing JSON columns", e); }

  return {
    requestId: String(getValue("ID RESPUESTA")),
    timestamp: String(getValue("FECHA SOLICITUD")),
    company: String(getValue("EMPRESA")),
    origin: String(getValue("CIUDAD ORIGEN")),
    destination: String(getValue("CIUDAD DESTINO")),
    requesterEmail: String(getValue("CORREO ENCUESTADO")),
    status: String(getValue("STATUS") || 'PENDIENTE_OPCIONES'),
    departureDate: safeDate(getValue("FECHA IDA")),
    returnDate: safeDate(getValue("FECHA VUELTA")),
    passengers: passengers,
    costCenter: String(getValue("CENTRO DE COSTOS")),
    variousCostCenters: String(getValue("VARIOS CENTROS COSTOS") || ''),
    businessUnit: String(getValue("UNIDAD DE NEGOCIO")),
    site: String(getValue("SEDE")),
    requiresHotel: getValue("REQUIERE HOSPEDAJE") === 'Sí',
    hotelName: String(getValue("NOMBRE HOTEL")),
    nights: Number(getValue("# NOCHES (AUTOMÁTICO)")) || 0,
    approverEmail: String(getValue("CORREO DE QUIEN APRUEBA (AUTOMÁTICO)")),
    analystOptions: analystOptions,
    selectedOption: selectedOption,
    comments: String(getValue("OBSERVACIONES") || ''),
    supportData: supportData
  };
}

function findApprover(costCenter) {
  return 'dsanchez@equitel.com.co'; 
}

function isUserAnalyst(email) {
  return email.includes('admin') || email.includes('compras') || email.includes('analista') || email === ADMIN_EMAIL;
}

function processApprovalFromEmail(e) {
  const id = e.parameter.id;
  const decision = e.parameter.decision;
  
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
     try {
        const status = decision === 'approved' ? 'APROBADO' : 'DENEGADO';
        // Pass empty payload as we just update status
        updateRequestStatus(id, status, {}); 
        
        return HtmlService.createHtmlOutput(`
          <div style="font-family: sans-serif; text-align: center; margin-top: 50px;">
            <h1 style="color: ${decision === 'approved' ? 'green' : 'red'}">
              Solicitud ${id} ha sido ${decision === 'approved' ? 'APROBADA' : 'DENEGADA'}
            </h1>
            <p>Se ha notificado al solicitante y al equipo de viajes.</p>
            <p>Puede cerrar esta ventana.</p>
          </div>
        `);
     } catch(err) {
        return HtmlService.createHtmlOutput("Error procesando solicitud: " + err.toString());
     } finally {
       lock.releaseLock();
     }
  } else {
     return HtmlService.createHtmlOutput("El sistema está ocupado. Por favor intente nuevamente en unos segundos.");
  }
}
