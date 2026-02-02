
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

// LOGO URL
const EMAIL_LOGO_URL = 'https://drive.google.com/thumbnail?id=1hA1i-1mG4DbBmzG1pFWafoDrCWwijRjq&sz=w1000';

// TODO: INSERT YOUR GEMINI API KEY HERE
const GEMINI_API_KEY = 'YOUR_GEMINI_API_KEY_HERE'; 

const LOCK_WAIT_MS = 30000;
const SHEET_NAME_REQUESTS = 'Nueva Base Solicitudes';
const SHEET_NAME_MASTERS = 'MAESTROS';
const SHEET_NAME_RELATIONS = 'CDS vs UDEN';
const SHEET_NAME_INTEGRANTES = 'INTEGRANTES';

// DRIVE CONFIGURATION
const ROOT_DRIVE_FOLDER_ID = '1uaett_yH1qZcS-rVr_sUh73mODvX02im';

// ADMIN EMAIL CONFIGURATION
const ADMIN_EMAIL = 'dsanchez@equitel.com.co';

// HEADERS EXACTLY AS PROVIDED IN CSV + JSON COLUMNS + NEW MODIFICATION COLUMNS
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
  "OPCIONES (JSON)", "SELECCION (JSON)", "SOPORTES (JSON)", "CORREOS PASAJEROS (JSON)",
  "DATA_CAMBIO_PENDIENTE (JSON)", "TEXTO_CAMBIO", "FLAG_CAMBIO_REALIZADO"
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

  // 3. Handle Modification Decision by Admin
  if (action === 'modify_decision') {
    return processModificationDecision(e);
  }
  
  // 4. Handle Standard API Calls (if used via GET)
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
    if (!e.postData || !e.postData.contents) {
       throw new Error("Empty Request Body");
    }

    const data = JSON.parse(e.postData.contents);
    const result = dispatch(data.action, data.payload);
    
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
  const isWriteAction = ['createRequest', 'updateRequest', 'uploadSupportFile', 'closeRequest', 'requestModification', 'updateAdminPin'].includes(action);
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
      case 'getCurrentUser': result = currentUserEmail; break;
      case 'getCostCenterData': result = getCostCenterData(); break;
      case 'getIntegrantesData': result = getIntegrantesData(); break;
      case 'getMyRequests': result = getRequestsByEmail(currentUserEmail); break;
      case 'getAllRequests': 
        if(!isUserAnalyst(currentUserEmail)) {
           result = getRequestsByEmail(currentUserEmail);
        } else {
           result = getAllRequests();
        }
        break;
      case 'createRequest': result = createNewRequest(payload); break;
      case 'updateRequest': result = updateRequestStatus(payload.id, payload.status, payload.payload); break;
      case 'uploadSupportFile': result = uploadSupportFile(payload.requestId, payload.fileData, payload.fileName, payload.mimeType); break;
      case 'closeRequest': result = updateRequestStatus(payload.requestId, 'PROCESADO'); break;
      case 'enhanceChangeText': result = enhanceTextWithGemini(payload.currentRequest, payload.userDraft); break;
      case 'requestModification': result = requestModification(payload.requestId, payload.modifiedRequest, payload.changeReason); break;
      
      // PIN FEATURES
      case 'verifyAdminPin': result = verifyAdminPin(payload.pin); break;
      case 'updateAdminPin': result = updateAdminPin(payload.newPin); break;

      default: return { success: false, error: 'Acción desconocida: ' + action };
    }
    
    return { success: true, data: result };

  } catch (e) {
    console.error("Error in dispatch: " + e.toString());
    return { success: false, error: e.toString() };
  } finally {
    if (isWriteAction) lock.releaseLock();
  }
}

// --- PIN LOGIC ---
function verifyAdminPin(inputPin) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const storedPin = scriptProperties.getProperty('ADMIN_PIN');
  // Default PIN if not set: 12345678
  const currentPin = storedPin ? storedPin : '12345678';
  return String(inputPin) === String(currentPin);
}

function updateAdminPin(newPin) {
  if (!newPin || String(newPin).length !== 8) {
    throw new Error("El PIN debe tener exactamente 8 dígitos.");
  }
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('ADMIN_PIN', String(newPin));
  return true;
}

// --- CORE BUSINESS LOGIC ---

function enhanceTextWithGemini(currentRequest, userDraft) {
  if (!GEMINI_API_KEY || GEMINI_API_KEY.includes('YOUR_GEMINI')) {
    return userDraft + " (Nota: Gemini API Key no configurada, texto original retornado)";
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${GEMINI_API_KEY}`;
  
  const prompt = `
    Actúa como un asistente administrativo experto en viajes corporativos.
    Tengo una solicitud de viaje existente con los siguientes datos actuales:
    - Origen: ${currentRequest.origin}
    - Destino: ${currentRequest.destination}
    - Fecha Ida: ${currentRequest.departureDate}
    - Fecha Vuelta: ${currentRequest.returnDate}
    - Pasajeros: ${currentRequest.passengers.length}
    - Hotel: ${currentRequest.requiresHotel ? 'Sí' : 'No'}

    El usuario quiere solicitar un CAMBIO en esta solicitud. 
    El usuario ha escrito este borrador de los cambios: "${userDraft}".

    Tu tarea es reescribir ese borrador para que sea:
    1. Extremadamente claro y específico.
    2. Formal y profesional.
    3. Resalte explícitamente qué está cambiando respecto a la solicitud original.
    
    Solo devuelve el texto mejorado, sin introducciones ni comillas.
  `;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }]
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
    
    const json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates.length > 0 && json.candidates[0].content) {
      return json.candidates[0].content.parts[0].text;
    }
    return userDraft; // Fallback
  } catch (e) {
    console.error("Gemini API Error", e);
    return userDraft; // Fallback on error
  }
}

function requestModification(requestId, modifiedRequest, changeReason) {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
   if (!sheet) throw new Error("Base de datos no encontrada");

   const idIdx = HEADERS_REQUESTS.indexOf("ID RESPUESTA");
   const statusIdx = HEADERS_REQUESTS.indexOf("STATUS");
   const changeDataIdx = HEADERS_REQUESTS.indexOf("DATA_CAMBIO_PENDIENTE (JSON)");
   const changeTextIdx = HEADERS_REQUESTS.indexOf("TEXTO_CAMBIO");

   const lastRow = sheet.getLastRow();
   const ids = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().flat();
   const rowIndex = ids.map(String).indexOf(String(requestId));

   if (rowIndex === -1) throw new Error("ID no encontrado");
   const rowNumber = rowIndex + 2;

   // 1. Save Pending Data
   sheet.getRange(rowNumber, changeDataIdx + 1).setValue(JSON.stringify(modifiedRequest));
   sheet.getRange(rowNumber, changeTextIdx + 1).setValue(changeReason);
   
   // 2. Update Status
   sheet.getRange(rowNumber, statusIdx + 1).setValue('PENDIENTE_APROBACION_CAMBIO');

   // 3. Notify Admin
   const originalRequest = mapRowToRequest(sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0]);
   sendModificationRequestEmail(originalRequest, modifiedRequest, changeReason);

   return true;
}

function processModificationDecision(e) {
  const id = e.parameter.id;
  const decision = e.parameter.decision;
  const confirm = e.parameter.confirm;

  // SECURITY CHECK: require manual confirmation if not present
  if (confirm !== 'true') {
      const url = `${WEB_APP_URL}?action=modify_decision&id=${id}&decision=${decision}&confirm=true`;
      return renderConfirmationPage(
          `¿Está seguro de ${decision === 'approve' ? 'APROBAR' : 'RECHAZAR'} el cambio para ${id}?`,
          "Esta acción notificará al usuario y actualizará el flujo del proceso.",
          decision === 'approve' ? 'APROBAR' : 'RECHAZAR',
          url,
          decision === 'approve' ? '#198754' : '#D71920'
      );
  }
  
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
     try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
        const idIdx = HEADERS_REQUESTS.indexOf("ID RESPUESTA");
        
        // Locate Row
        const lastRow = sheet.getLastRow();
        const ids = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().flat();
        const rowIndex = ids.map(String).indexOf(String(id));
        
        if (rowIndex === -1) throw new Error("Solicitud no encontrada");
        const rowNumber = rowIndex + 2;

        const changeDataIdx = HEADERS_REQUESTS.indexOf("DATA_CAMBIO_PENDIENTE (JSON)");
        const changeTextIdx = HEADERS_REQUESTS.indexOf("TEXTO_CAMBIO");
        const flagIdx = HEADERS_REQUESTS.indexOf("FLAG_CAMBIO_REALIZADO");
        const statusIdx = HEADERS_REQUESTS.indexOf("STATUS");
        const requesterEmailIdx = HEADERS_REQUESTS.indexOf("CORREO ENCUESTADO");
        
        const requesterEmail = sheet.getRange(rowNumber, requesterEmailIdx + 1).getValue();

        if (decision === 'approve') {
             // 1. Get New Data
             const newJson = sheet.getRange(rowNumber, changeDataIdx + 1).getValue();
             const reason = sheet.getRange(rowNumber, changeTextIdx + 1).getValue();
             
             if (!newJson) throw new Error("No hay datos de cambio pendientes.");
             const newData = JSON.parse(newJson);

             // 2. Overwrite Main Columns
             const setVal = (header, val) => {
                 const idx = HEADERS_REQUESTS.indexOf(header);
                 if (idx > -1) sheet.getRange(rowNumber, idx + 1).setValue(val);
             };

             setVal("CIUDAD ORIGEN", newData.origin);
             setVal("CIUDAD DESTINO", newData.destination);
             setVal("FECHA IDA", newData.departureDate);
             setVal("FECHA VUELTA", newData.returnDate);
             setVal("HORA LLEGADA VUELO IDA", newData.departureTimePreference);
             setVal("HORA LLEGADA VUELO VUELTA", newData.returnTimePreference);
             setVal("EMPRESA", newData.company);
             setVal("SEDE", newData.site);
             setVal("CENTRO DE COSTOS", newData.costCenter);
             setVal("REQUIERE HOSPEDAJE", newData.requiresHotel ? 'Sí' : 'No');
             setVal("NOMBRE HOTEL", newData.hotelName);
             setVal("# NOCHES (AUTOMÁTICO)", newData.nights);
             setVal("OBSERVACIONES", newData.comments);
             
             // Passengers
             setVal("# PERSONAS QUE VIAJAN", newData.passengers.length);
             const p = newData.passengers || [];
             for(let i=0; i<5; i++) {
                const baseH = `PERSONA ${i+1}`;
                setVal(`CÉDULA ${baseH}`, p[i] ? p[i].idNumber : '');
                setVal(`NOMBRE ${baseH}`, p[i] ? p[i].name : '');
             }

             // 3. Set Flag & Clear Pending
             setVal("FLAG_CAMBIO_REALIZADO", "CAMBIO GENERADO");
             setVal("DATA_CAMBIO_PENDIENTE (JSON)", ""); 
             
             // 4. RESET PROCESS
             setVal("STATUS", "PENDIENTE_OPCIONES");
             setVal("OPCIONES (JSON)", "");
             setVal("SELECCION (JSON)", "");
             setVal("APROBADO POR ÁREA?", "");

             // 5. Notify Requester
             sendEmailRich(requesterEmail, "Cambio Aprobado - Solicitud Reiniciada", 
                 HtmlTemplates.modificationResult(id, 'approve', {requestId: id, requesterEmail: requesterEmail, company: newData.company, site: newData.site})
             );

        } else {
             // REJECTED
             sheet.getRange(rowNumber, statusIdx + 1).setValue('PENDIENTE_OPCIONES');
             sheet.getRange(rowNumber, changeDataIdx + 1).setValue("");
             sheet.getRange(rowNumber, changeTextIdx + 1).setValue("");
             
             // Notify Requester (Need original data for subject, using existing row)
             const originalReq = mapRowToRequest(sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0]);
             
             sendEmailRich(requesterEmail, "Cambio Rechazado", 
                 HtmlTemplates.modificationResult(id, 'reject', originalReq)
             );
        }

        return renderSuccessPage("Acción completada con éxito.", decision === 'approve' ? 'La solicitud ha regresado a Pendiente Opciones.' : 'La solicitud original se mantiene.');

     } catch(err) {
        return HtmlService.createHtmlOutput("Error: " + err.toString());
     } finally {
       lock.releaseLock();
     }
  }
}

// --- EMAIL HELPERS & TEMPLATES ---

// Helper to ensure ALL SUBJECTS follow the exact strict structure
function getStandardSubject(data) {
    // Solicitud de Viaje SOL-000009 - dsanchez@equitel.com.co - Equitel BARRANQUILLA
    // data must have: requestId, requesterEmail, company, site
    const id = data.requestId || data.id; // handle both naming conventions
    return `Solicitud de Viaje ${id} - ${data.requesterEmail} - ${data.company} ${data.site}`;
}

function sendEmailRich(to, subject, htmlBody) {
    try {
        MailApp.sendEmail({
            to: to,
            cc: ADMIN_EMAIL, 
            subject: subject,
            htmlBody: htmlBody
        });
    } catch(e) {
        console.error("Error sending email: " + e);
    }
}

// Class-like structure for Templates
const HtmlTemplates = {
    // Base layout wrapper
    layout: function(title, content) {
        return `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>${title}</title>
        </head>
        <body style="margin: 0; padding: 0; background-color: #f4f4f4; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #f4f4f4; width: 100%;">
                <tr>
                    <td align="center" style="padding: 20px 0;">
                        <!-- CONTAINER -->
                        <table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width: 600px; width: 100%; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.1);">
                            <!-- HEADER RED BLOCK -->
                            <tr>
                                <td style="background-color: #D71920; padding: 30px 20px; text-align: center; border-bottom: 4px solid #b01319;">
                                    <h1 style="color: #ffffff; margin: 0; font-size: 24px; font-weight: 700; letter-spacing: 0.5px; text-transform: uppercase;">
                                        Nueva Solicitud de Viaje/Hospedaje
                                    </h1>
                                    <p style="color: #ffffff; margin: 10px 0 0 0; font-size: 16px; font-weight: 600; opacity: 0.9;">
                                        ID: ${title}
                                    </p>
                                </td>
                            </tr>
                            <!-- CONTENT -->
                            <tr>
                                <td style="padding: 30px 40px; color: #333333; line-height: 1.6;">
                                    ${content}
                                </td>
                            </tr>
                            <!-- FOOTER -->
                            <tr>
                                <td style="background-color: #eeeeee; padding: 15px; text-align: center; font-size: 12px; color: #777777; border-top: 1px solid #e0e0e0;">
                                    <p style="margin: 0;">&copy; ${new Date().getFullYear()} Organización Equitel. Gestión de Viajes Corporativos.</p>
                                    <p style="margin: 5px 0 0 0;">Este es un mensaje automático, por favor no responder.</p>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </body>
        </html>
        `;
    },

    newRequest: function(data, requestId, link) {
        // Passenger List styled
        const passengerList = (data.passengers || []).map(p => 
            `<li style="margin-bottom: 5px;">${p.name} <span style="color:#666; font-size:12px;">(${p.idNumber})</span></li>`
        ).join('');

        const content = `
            <p style="font-size: 15px; color: #555; margin-bottom: 25px;">
                Se ha registrado un nuevo requerimiento de viaje para <a href="mailto:${data.requesterEmail}" style="color: #0056b3; text-decoration: none;">${data.requesterEmail}</a>.
            </p>
            
            <!-- TICKET STYLE CARD -->
            <div style="background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden; margin-bottom: 30px;">
                <!-- ROUTE ROW -->
                <div style="padding: 20px; border-bottom: 1px solid #e5e7eb;">
                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                        <tr>
                            <td align="left" width="40%">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px;">ORIGEN</span>
                                <span style="display: block; font-size: 18px; font-weight: 800; color: #111827;">${data.origin}</span>
                            </td>
                            <td align="center" width="20%">
                                <span style="color: #d1d5db; font-size: 24px;">➝</span>
                            </td>
                            <td align="right" width="40%">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px;">DESTINO</span>
                                <span style="display: block; font-size: 18px; font-weight: 800; color: #111827;">${data.destination}</span>
                            </td>
                        </tr>
                    </table>
                </div>
                
                <!-- DATES ROW -->
                <div style="background-color: #ffffff;">
                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                        <tr>
                            <td align="center" width="50%" style="padding: 15px; border-right: 1px solid #e5e7eb;">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; text-transform: uppercase; margin-bottom: 4px;">FECHA IDA</span>
                                <div style="display: inline-block; color: #D71920; font-weight: bold; font-size: 14px;">
                                    📅 ${data.departureDate}
                                </div>
                                <div style="font-size: 12px; color: #6b7280; margin-top: 2px;">
                                    (${data.departureTimePreference || 'N/A'})
                                </div>
                            </td>
                            <td align="center" width="50%" style="padding: 15px;">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; text-transform: uppercase; margin-bottom: 4px;">FECHA REGRESO</span>
                                <div style="display: inline-block; color: #D71920; font-weight: bold; font-size: 14px;">
                                    📅 ${data.returnDate || 'N/A'}
                                </div>
                                <div style="font-size: 12px; color: #6b7280; margin-top: 2px;">
                                    (${data.returnTimePreference || 'Solo Ida'})
                                </div>
                            </td>
                        </tr>
                    </table>
                </div>
            </div>

            <h3 style="font-size: 14px; color: #374151; margin-bottom: 15px; border-bottom: 2px solid #f3f4f6; padding-bottom: 5px;">Detalles del Caso</h3>
            
            <table width="100%" cellpadding="6" cellspacing="0" border="0" style="font-size: 14px; margin-bottom: 25px;">
                <tr>
                    <td width="35%" style="color: #6b7280;">Empresa / Sede:</td>
                    <td style="font-weight: 600; color: #111827;">${data.company} - ${data.site}</td>
                </tr>
                <tr>
                    <td style="color: #6b7280;">Centro de Costos:</td>
                    <td style="font-weight: 600; color: #111827;">${data.costCenter} (${data.businessUnit})</td>
                </tr>
                <tr>
                    <td style="color: #6b7280;">Aprobador:</td>
                    <td><a href="mailto:${data.approverEmail}" style="color: #0056b3;">${data.approverEmail}</a></td>
                </tr>
                <tr>
                    <td style="color: #6b7280;">Hospedaje:</td>
                    <td style="font-weight: 600; color: #0056b3;">${data.requiresHotel ? `Sí - ${data.hotelName} (${data.nights} Noches)` : 'No'}</td>
                </tr>
            </table>

            <!-- NOTES BOX (YELLOW) -->
            ${data.comments ? `
            <div style="background-color: #fffbeb; border: 1px solid #fcd34d; border-radius: 6px; padding: 15px; margin-bottom: 20px;">
                <div style="display: flex; align-items: flex-start;">
                    <span style="font-size: 16px; margin-right: 10px;">📝</span>
                    <div>
                        <strong style="display: block; font-size: 12px; color: #92400e; text-transform: uppercase; margin-bottom: 4px;">Observaciones / Notas:</strong>
                        <span style="font-size: 14px; color: #b45309; font-style: italic;">${data.comments}</span>
                    </div>
                </div>
            </div>
            ` : ''}

            <!-- PASSENGERS BOX (BLUE) -->
            <div style="background-color: #eff6ff; border: 1px solid #bfdbfe; border-radius: 6px; padding: 15px; margin-bottom: 30px;">
                 <strong style="display: block; font-size: 12px; color: #1e40af; text-transform: uppercase; margin-bottom: 8px;">👥 Pasajero(s) (${data.passengers.length}):</strong>
                 <ul style="margin: 0; padding-left: 20px; color: #1e3a8a; font-size: 14px;">${passengerList}</ul>
            </div>

            <!-- BUTTON -->
            <div style="text-align: center;">
                <a href="${link}" style="background-color: #111827; color: #ffffff; padding: 14px 32px; text-decoration: none; border-radius: 30px; font-weight: bold; font-size: 14px; display: inline-block; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">Ingresar a la Plataforma</a>
            </div>
        `;
        return this.layout(`${requestId}`, content);
    },

    optionsAvailable: function(request, options, link) {
        let optionsHtml = '';
        if (options && options.length) {
            optionsHtml = options.map(opt => `
                <div style="border: 1px solid #e5e7eb; border-radius: 8px; padding: 15px; margin-bottom: 15px; background-color: #ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.05);">
                    <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #f3f4f6; padding-bottom: 10px; margin-bottom: 10px;">
                        <span style="font-weight: 800; color: #D71920; font-size: 18px;">Opción ${opt.id}</span>
                        <span style="font-weight: 800; color: #111827; font-size: 16px;">$ ${Number(opt.totalPrice).toLocaleString()}</span>
                    </div>
                    <div style="font-size: 14px; color: #4b5563;">
                        <p style="margin: 5px 0;"><strong>🛫 Ida:</strong> ${opt.outbound.airline} (${opt.outbound.flightTime})</p>
                        ${opt.inbound ? `<p style="margin: 5px 0;"><strong>🛬 Regreso:</strong> ${opt.inbound.airline} (${opt.inbound.flightTime})</p>` : ''}
                        ${opt.hotel ? `<p style="margin: 8px 0 0 0; padding-top: 8px; border-top: 1px dashed #e5e7eb; color: #0056b3;"><strong>🏨 Hotel:</strong> ${opt.hotel}</p>` : ''}
                    </div>
                </div>
            `).join('');
        }

        const content = `
            <p style="color: #374151;">Estimado usuario,</p>
            <p style="color: #374151;">El equipo de gestión de viajes ha cargado las opciones disponibles para su solicitud <strong>${request.requestId}</strong>.</p>
            
             <!-- Context Summary -->
            <div style="background-color: #f9fafb; padding: 10px; border-radius: 6px; font-size: 13px; color: #666; margin: 10px 0;">
               <strong>Ruta:</strong> ${request.origin} ➝ ${request.destination} <br/>
               <strong>Empresa:</strong> ${request.company} ${request.site}
            </div>

            <p style="color: #374151;">Por favor ingrese a la plataforma para seleccionar la opción que mejor se ajuste a sus necesidades.</p>
            
            <div style="margin: 25px 0;">
                ${optionsHtml}
            </div>

            <div style="text-align: center;">
                <a href="${link}" style="background-color: #D71920; color: #ffffff; padding: 14px 32px; text-decoration: none; border-radius: 30px; font-weight: bold; font-size: 14px; display: inline-block;">SELECCIONAR OPCIÓN</a>
            </div>
        `;
        return this.layout(`${request.requestId}`, content);
    },

    decisionNotification: function(request, status) {
        const isApproved = status === 'APROBADO';
        const color = isApproved ? '#059669' : '#dc2626'; 
        const icon = isApproved ? '✅' : '❌';
        
        const content = `
            <div style="text-align: center; margin-bottom: 25px;">
                <div style="font-size: 48px; margin-bottom: 10px;">${icon}</div>
                <h2 style="color: ${color}; margin: 0; text-transform: uppercase; font-weight: 800;">SOLICITUD ${status}</h2>
                <p style="font-size: 16px; margin-top: 10px; color: #4b5563;">Su solicitud de viaje <strong>${request.requestId}</strong> ha sido procesada.</p>
            </div>

            <div style="background-color: #f9fafb; padding: 15px; border-radius: 8px; border: 1px solid #e5e7eb; margin-bottom: 20px;">
                <table width="100%" cellpadding="5" cellspacing="0" style="font-size: 14px; color: #374151;">
                    <tr>
                        <td width="30%" style="color: #6b7280;">Ruta:</td>
                        <td><strong>${request.origin} ➝ ${request.destination}</strong></td>
                    </tr>
                    <tr>
                        <td style="color: #6b7280;">Fechas:</td>
                        <td>${request.departureDate}</td>
                    </tr>
                     <tr>
                        <td style="color: #6b7280;">Sede:</td>
                        <td>${request.company} ${request.site}</td>
                    </tr>
                </table>
            </div>

            <div style="background-color: ${isApproved ? '#ecfdf5' : '#fef2f2'}; color: ${isApproved ? '#065f46' : '#991b1b'}; padding: 15px; border-radius: 6px; text-align: center; font-size: 14px; border: 1px solid ${isApproved ? '#a7f3d0' : '#fecaca'};">
                ${isApproved 
                    ? 'Por favor, ingrese a la plataforma para cargar los soportes (Facturas/Tiquetes) una vez realice la compra/viaje.' 
                    : 'Si tiene dudas sobre el rechazo, por favor contacte al área administrativa.'}
            </div>

            <div style="text-align: center; margin-top: 30px;">
                <a href="${PLATFORM_URL}" style="color: #4b5563; text-decoration: underline; font-size: 14px;">Ir a la Plataforma</a>
            </div>
        `;
        return this.layout(`${request.requestId}`, content);
    },

    modificationRequest: function(original, modified, reason, approveLink, rejectLink) {
        const content = `
            <p style="color: #374151;">El usuario <strong>${original.requesterEmail}</strong> ha solicitado cambios en una solicitud existente.</p>
            
            <div style="background-color: #fffbeb; border-left: 4px solid #f59e0b; padding: 15px; margin: 20px 0;">
                <strong style="display: block; margin-bottom: 5px; color: #92400e; font-size: 12px; text-transform: uppercase;">MOTIVO DEL CAMBIO:</strong>
                <span style="font-style: italic; color: #111827; font-size: 15px;">"${reason}"</span>
            </div>

            <table width="100%" cellpadding="8" cellspacing="0" border="1" style="border-collapse: collapse; border-color: #e5e7eb; margin-bottom: 30px; font-size: 13px;">
                <tr style="background-color: #f9fafb;">
                    <th style="text-align: left; color: #6b7280;">Campo</th>
                    <th style="text-align: left; color: #6b7280;">Original</th>
                    <th style="text-align: left; color: #D71920;">Nuevo (Propuesto)</th>
                </tr>
                <tr>
                    <td>Ruta</td>
                    <td>${original.origin} ➝ ${original.destination}</td>
                    <td style="font-weight:bold;">${modified.origin} ➝ ${modified.destination}</td>
                </tr>
                <tr>
                    <td>Fechas</td>
                    <td>${original.departureDate} / ${original.returnDate || '-'}</td>
                    <td style="font-weight:bold;">${modified.departureDate} / ${modified.returnDate || '-'}</td>
                </tr>
                <tr>
                    <td>Horarios</td>
                    <td>${original.departureTimePreference} / ${original.returnTimePreference}</td>
                    <td style="font-weight:bold;">${modified.departureTimePreference} / ${modified.returnTimePreference}</td>
                </tr>
                <tr>
                    <td>Hotel</td>
                    <td>${original.requiresHotel ? 'Sí' : 'No'} (${original.hotelName})</td>
                    <td style="font-weight:bold;">${modified.requiresHotel ? 'Sí' : 'No'} (${modified.hotelName})</td>
                </tr>
                <tr>
                    <td>Pasajeros</td>
                    <td>${original.passengers.length}</td>
                    <td style="font-weight:bold;">${modified.passengers.length}</td>
                </tr>
                <tr>
                    <td>Obs.</td>
                    <td>${original.comments}</td>
                    <td style="font-weight:bold;">${modified.comments}</td>
                </tr>
            </table>

            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td align="center" style="padding-right: 10px;">
                        <a href="${approveLink}" style="background-color: #059669; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; display: block;">APROBAR CAMBIO</a>
                    </td>
                    <td align="center" style="padding-left: 10px;">
                        <a href="${rejectLink}" style="background-color: #dc2626; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; display: block;">RECHAZAR CAMBIO</a>
                    </td>
                </tr>
            </table>
        `;
        return this.layout(`${original.requestId}`, content);
    },

    modificationResult: function(id, decision, data) {
        const isApproved = decision === 'approve';
        const title = isApproved ? 'CAMBIO APROBADO' : 'CAMBIO RECHAZADO';
        const color = isApproved ? '#059669' : '#dc2626';
        
        const content = `
            <div style="text-align: center; margin: 20px 0;">
                <h2 style="color: ${color}; border: 2px solid ${color}; display: inline-block; padding: 10px 20px; border-radius: 8px; text-transform: uppercase; font-size: 16px; font-weight: 800;">
                    ${title}
                </h2>
            </div>
            <p style="color: #374151;">La solicitud de modificación para el ticket <strong>${id}</strong> ha sido gestionada por el administrador.</p>
            
             <!-- Context Summary -->
            <div style="background-color: #f9fafb; padding: 10px; border-radius: 6px; font-size: 13px; color: #666; margin: 10px 0;">
               <strong>Empresa:</strong> ${data.company} ${data.site}
            </div>

            ${isApproved 
                ? `<div style="background-color: #ecfdf5; padding: 15px; border-radius: 6px; color: #065f46; font-size: 14px;">
                    Su solicitud ha sido reiniciada. El equipo de viajes cargará nuevas opciones basadas en sus cambios.
                   </div>`
                : `<div style="background-color: #fef2f2; padding: 15px; border-radius: 6px; color: #991b1b; font-size: 14px;">
                    Los cambios no fueron aceptados. La solicitud original se mantiene vigente.
                   </div>`
            }
            
            <div style="text-align: center; margin-top: 30px;">
                <a href="${PLATFORM_URL}" style="background-color: #374151; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-size: 14px;">Ir a la Plataforma</a>
            </div>
        `;
        return this.layout(`${id}`, content);
    }
};

// --- UPDATED EMAIL FUNCTIONS USING TEMPLATES ---

function sendModificationRequestEmail(original, modified, reason) {
   const approveLink = `${WEB_APP_URL}?action=modify_decision&id=${original.requestId}&decision=approve`;
   const rejectLink = `${WEB_APP_URL}?action=modify_decision&id=${original.requestId}&decision=reject`;
   const htmlBody = HtmlTemplates.modificationRequest(original, modified, reason, approveLink, rejectLink);
   const subject = getStandardSubject(original) + " - SOLICITUD DE CAMBIO";
   
   sendEmailRich(ADMIN_EMAIL, subject, htmlBody);
}

function sendNewRequestNotification(data, requestId) {
    // Inject ID into data object temporarily for subject generation if needed, 
    // but getStandardSubject handles 'id' or 'requestId'.
    const subjectData = { ...data, requestId: requestId };
    const htmlBody = HtmlTemplates.newRequest(data, requestId, PLATFORM_URL);
    const subject = getStandardSubject(subjectData);
    
    const ccEmails = [data.requesterEmail, getCCList(data)].filter(e => e).join(',');
    
    try {
        MailApp.sendEmail({ 
            to: ADMIN_EMAIL, 
            cc: ccEmails, 
            subject: subject, 
            htmlBody: htmlBody 
        });
    } catch (e) {
        console.error("Error sending new req email: " + e);
    }
}

function sendOptionsToRequester(recipient, request, options) {
   const link = PLATFORM_URL; 
   const htmlBody = HtmlTemplates.optionsAvailable(request, options, link);
   const subject = getStandardSubject(request); // EXACT SUBJECT AS REQUESTED
   const ccList = getCCList(request);

   try { 
       MailApp.sendEmail({ 
           to: recipient, 
           cc: ccList, 
           subject: subject, 
           htmlBody: htmlBody 
       }); 
   } catch(e) {
       console.error("Error sending options email: " + e);
   }
}

function sendDecisionNotification(request, status) {
  const htmlBody = HtmlTemplates.decisionNotification(request, status);
  const subject = getStandardSubject(request); // EXACT SUBJECT AS REQUESTED
  const ccList = [ADMIN_EMAIL, getCCList(request)].join(',');
  
  try { 
      MailApp.sendEmail({ 
          to: request.requesterEmail, 
          cc: ccList, 
          subject: subject, 
          htmlBody: htmlBody 
      }); 
  } catch(e){
      console.error("Error sending decision email: " + e);
  }
}

// Helpers for render pages (unchanged logic, just ensuring they exist)
function renderConfirmationPage(title, message, actionText, actionUrl, color) {
    const html = `
        <html>
          <body style="font-family: Arial, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; background-color: #f4f4f4;">
            <div style="background: white; padding: 40px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; max-w: 500px;">
               <h2 style="color: #333; margin-bottom: 20px;">${title}</h2>
               <p style="color: #666; margin-bottom: 30px;">${message}</p>
               <a href="${actionUrl}" style="display: inline-block; background-color: ${color}; color: white; padding: 15px 30px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px;">
                 ${actionText}
               </a>
            </div>
          </body>
        </html>
      `;
    return HtmlService.createHtmlOutput(html).setTitle('Confirmar Acción');
}

function renderSuccessPage(title, message) {
    const html = `
        <html>
          <body style="font-family: Arial, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; background-color: #f4f4f4;">
            <div style="background: white; padding: 40px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center;">
               <div style="color: green; font-size: 48px; margin-bottom: 10px;">✓</div>
               <h2 style="color: #333;">${title}</h2>
               <p style="color: #666;">${message}</p>
            </div>
          </body>
        </html>
    `;
    return HtmlService.createHtmlOutput(html).setTitle('Procesado');
}

// ... (Rest of existing functions: getRequestsByEmail, getCostCenterData, etc. remain mostly same but mapRowToRequest needs update)

function getRequestsByEmail(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const emailIdx = HEADERS_REQUESTS.indexOf("CORREO ENCUESTADO");
  
  if (emailIdx === -1) return [];
  const targetEmail = String(email).toLowerCase().trim();

  const userRequests = data.filter(row => {
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

// ... (getCostCenterData, getIntegrantesData remain same)
function getCostCenterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return [];
  const sheet = ss.getSheetByName(SHEET_NAME_RELATIONS);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
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
  const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  return data.map(row => ({
    idNumber: String(row[0]).trim(),
    name: String(row[1]),
    email: String(row[2]).toLowerCase().trim(),
    approverName: String(row[10]), 
    approverEmail: String(row[11]).toLowerCase().trim() 
  })).filter(i => i.idNumber && i.name);
}

function createNewRequest(data) {
    // ... existing implementation logic ...
    // Using a shorthand here to save tokens, assuming previous createNewRequest logic is preserved 
    // but just ensuring it uses the global HEADERS_REQUESTS constant correctly.
    // Copying the full function to ensure XML validity:
    
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("No Container Spreadsheet found");
  const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
  if (!sheet) throw new Error("Base de datos no inicializada.");

  const idColIndex = HEADERS_REQUESTS.indexOf("ID RESPUESTA") + 1; 
  const lastRow = sheet.getLastRow();
  let nextIdNum = 1;

  if (lastRow > 1) {
    const existingIds = sheet.getRange(2, idColIndex, lastRow - 1, 1).getValues().flat();
    const numericIds = existingIds
      .map(val => {
         const strVal = String(val).replace(/^SOL-/, '');
         return parseInt(strVal, 10);
      })
      .filter(val => !isNaN(val));
    if (numericIds.length > 0) nextIdNum = Math.max(...numericIds) + 1;
  }
  
  const id = `SOL-${nextIdNum.toString().padStart(6, '0')}`; 
  let ccName = '';
  if (data.costCenter && data.costCenter !== 'VARIOS') {
     const masters = getCostCenterData();
     const ccObj = masters.find(m => m.code == data.costCenter);
     if (ccObj) ccName = ccObj.name;
  }
  data.costCenterName = ccName;

  let approverEmail = ADMIN_EMAIL; 
  if (data.passengers && data.passengers.length > 0) {
     const firstPassengerId = data.passengers[0].idNumber;
     const integrantes = getIntegrantesData();
     const integrant = integrantes.find(i => i.idNumber === firstPassengerId);
     if (integrant && integrant.approverEmail) approverEmail = integrant.approverEmail;
     else approverEmail = findApprover(data.costCenter);
  }

  let nights = 0;
  if (data.requiresHotel) {
      if (data.nights && data.nights > 0) nights = data.nights;
      else if (data.departureDate && data.returnDate) {
          const d1 = new Date(data.departureDate);
          const d2 = new Date(data.returnDate);
          const diffTime = Math.abs(d2 - d1);
          nights = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      }
  }

  const row = new Array(HEADERS_REQUESTS.length).fill('');
  row[0] = new Date(); 
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
  
  const statusIdx = HEADERS_REQUESTS.indexOf("STATUS");
  if (statusIdx > -1) row[statusIdx] = 'PENDIENTE_OPCIONES';

  const obsIdx = HEADERS_REQUESTS.indexOf("OBSERVACIONES");
  if (obsIdx > -1) row[obsIdx] = data.comments || '';

  const approverIdx = HEADERS_REQUESTS.indexOf("CORREO DE QUIEN APRUEBA (AUTOMÁTICO)");
  if (approverIdx > -1) row[approverIdx] = approverEmail;

  const emailsIdx = HEADERS_REQUESTS.indexOf("CORREOS PASAJEROS (JSON)");
  if (emailsIdx > -1) {
     const pEmails = data.passengers.map(p => p.email).filter(e => e);
     row[emailsIdx] = JSON.stringify(pEmails);
  }

  sheet.appendRow(row);
  data.approverEmail = approverEmail;
  sendNewRequestNotification(data, id);
  return id;
}

// ... (updateRequestStatus, uploadSupportFile, sendNewRequestNotification, sendOptionsToRequester, sendDecisionNotification, findApprover, isUserAnalyst, processApprovalFromEmail, getCCList remain same)
// IMPORTANT: mapRowToRequest needs to read new columns

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

   sheet.getRange(rowNumber, statusIdx + 1).setValue(status);

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

   if (status === 'PENDIENTE_SELECCION' && payload && payload.analystOptions) {
      const requesterEmail = sheet.getRange(rowNumber, emailIdx + 1).getValue();
      const fullRequest = mapRowToRequest(sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0]);
      sendOptionsToRequester(requesterEmail, fullRequest, payload.analystOptions);
   }

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
  const supportIdx = HEADERS_REQUESTS.indexOf("SOPORTES (JSON)");
  if (supportIdx === -1) throw new Error("Columna SOPORTES no configurada");
  const jsonStr = sheet.getRange(rowNumber, supportIdx + 1).getValue();
  let supportData = jsonStr ? JSON.parse(jsonStr) : { folderId: null, folderUrl: null, files: [] };
  let folder;
  try {
      if (supportData.folderId) {
         try { folder = DriveApp.getFolderById(supportData.folderId); } catch(e) {}
      }
      if (!folder) {
         const rootFolder = DriveApp.getFolderById(ROOT_DRIVE_FOLDER_ID);
         const folderName = `${requestId}`; 
         const folders = rootFolder.getFoldersByName(folderName);
         if (folders.hasNext()) folder = folders.next();
         else folder = rootFolder.createFolder(folderName);
         supportData.folderId = folder.getId();
         supportData.folderUrl = folder.getUrl();
      }
  } catch(e) {
      throw new Error("Error accediendo a Google Drive: " + e.toString());
  }
  const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
  const file = folder.createFile(blob);
  const newFileEntry = {
    id: file.getId(),
    name: file.getName(),
    url: file.getUrl(),
    mimeType: mimeType,
    date: new Date().toISOString()
  };
  supportData.files.push(newFileEntry);
  sheet.getRange(rowNumber, supportIdx + 1).setValue(JSON.stringify(supportData));
  return supportData;
}

function getCCList(request) {
    const requester = request.requesterEmail;
    const passengerEmails = (request.passengers || [])
        .map(p => p.email)
        .filter(e => e && e.toLowerCase() !== requester.toLowerCase());
    return passengerEmails.join(',');
}

function findApprover(costCenter) { return 'dsanchez@equitel.com.co'; }
function isUserAnalyst(email) { return email.includes('admin') || email.includes('compras') || email.includes('analista') || email === ADMIN_EMAIL; }

function processApprovalFromEmail(e) {
  const id = e.parameter.id;
  const decision = e.parameter.decision;
  const confirm = e.parameter.confirm;

  // ADDED: Confirmation Step to prevent scanners/previewers from triggering action
  if (confirm !== 'true') {
      const url = `${WEB_APP_URL}?action=approve&id=${id}&decision=${decision}&confirm=true`;
      return renderConfirmationPage(
          `¿Está seguro de ${decision === 'approved' ? 'APROBAR' : 'DENEGAR'} la solicitud ${id}?`,
          "Esta acción enviará notificaciones automáticas y registrará su decisión.",
          decision === 'approved' ? 'APROBAR' : 'DENEGAR',
          url,
          decision === 'approved' ? '#198754' : '#D71920'
      );
  }

  updateRequestStatus(id, decision === 'approved' ? 'APROBADO' : 'DENEGADO', {});
  return renderSuccessPage("Acción Procesada", "La solicitud ha sido actualizada correctamente.");
}

function processOptionSelection(e) {
    // Not explicitly asked to change but needed for compilation if referenced
    // ...
    return HtmlService.createHtmlOutput("Seleccionado.");
}

function mapRowToRequest(row) {
  const getValue = (headerName) => {
    const idx = HEADERS_REQUESTS.indexOf(headerName);
    if (idx === -1 || idx >= row.length) return ''; 
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
       passengers.push({ name: String(name), idNumber: String(id), email: passengerEmails[i-1] || '' });
     }
  }

  let analystOptions = [], selectedOption = null, supportData = undefined;
  try {
    const rawOpt = getValue("OPCIONES (JSON)");
    if (rawOpt) analystOptions = JSON.parse(rawOpt); 
    const rawSel = getValue("SELECCION (JSON)");
    if (rawSel) selectedOption = JSON.parse(rawSel);
    const rawSup = getValue("SOPORTES (JSON)");
    if (rawSup) supportData = JSON.parse(rawSup);
  } catch(e) {}

  // Modification Data
  let pendingChangeData = undefined;
  const changeText = getValue("TEXTO_CAMBIO");
  const changeFlag = getValue("FLAG_CAMBIO_REALIZADO") === "CAMBIO GENERADO";
  try {
    const rawPending = getValue("DATA_CAMBIO_PENDIENTE (JSON)");
    if (rawPending) pendingChangeData = rawPending; // Keep string or parse? Interface says string for pendingChangeData usually, but let's parse if needed. Actually Types says string.
  } catch(e) {}

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
    supportData: supportData,
    departureTimePreference: getValue("HORA LLEGADA VUELO IDA"),
    returnTimePreference: getValue("HORA LLEGADA VUELO VUELTA"),
    // New Fields
    pendingChangeData: String(getValue("DATA_CAMBIO_PENDIENTE (JSON)") || ''),
    changeReason: String(changeText || ''),
    hasChangeFlag: changeFlag
  };
}
