
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
const GEMINI_API_KEY = 'AIzaSyBBI3PTPaSsspcpA_XkzQjl--Bae101Cbo'; 

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
    const lastCol = sheet.getLastColumn();
    if (lastCol < HEADERS_REQUESTS.length) {
       sheet.getRange(1, 1, 1, HEADERS_REQUESTS.length).setValues([HEADERS_REQUESTS]);
    }
  }

  let relSheet = ss.getSheetByName(SHEET_NAME_RELATIONS);
  if (!relSheet) {
    relSheet = ss.insertSheet(SHEET_NAME_RELATIONS);
    relSheet.appendRow(["CENTRO COSTOS", "Descripcion del CC", "UNIDAD DE NEGOCIO"]);
  }

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

// ... (Rest of PIN, Gemini Logic, Setup, etc. standard functions kept for context if needed, focusing on changes below)

function verifyAdminPin(inputPin) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const storedPin = scriptProperties.getProperty('ADMIN_PIN');
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

function enhanceTextWithGemini(currentRequest, userDraft) {
  if (!GEMINI_API_KEY || GEMINI_API_KEY.includes('YOUR_GEMINI')) return userDraft;
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  
  // Construct a rich prompt with context
  const context = `
    ID Solicitud: ${currentRequest.requestId}
    Solicitante: ${currentRequest.requesterEmail}
    Empresa: ${currentRequest.company}
    Ruta Original: ${currentRequest.origin} -> ${currentRequest.destination}
    Fecha Ida Original: ${currentRequest.departureDate}
    Fecha Regreso Original: ${currentRequest.returnDate || 'N/A'}
  `;

  const prompt = `
    Actúa como un asistente administrativo experto en gestión de viajes corporativos.
    Tu tarea es redactar una JUSTIFICACIÓN FORMAL Y CLARA para un cambio en una solicitud de viaje.
    
    CONTEXTO DE LA SOLICITUD ORIGINAL:
    ${context}

    EL USUARIO DICE (BORRADOR):
    "${userDraft}"

    INSTRUCCIONES:
    1. Redacta un párrafo breve (máximo 3 oraciones) que explique el motivo del cambio de manera profesional.
    2. Usa un tono formal y persuasivo dirigido al aprobador financiero.
    3. Si el usuario menciona cambios de fecha, ruta o pasajeros, inclúyelos explícitamente en la redacción para dar claridad.
    4. NO inventes información que no esté en el borrador del usuario, pero sí puedes inferir que es por "necesidades del servicio" o "ajustes de agenda" si el texto es muy vago.
    5. Devuelve SOLAMENTE el texto final de la justificación, sin saludos ni introducciones.
  `;

  try {
    const payload = {
      contents: [{
        parts: [{ text: prompt }]
      }]
    };

    const response = UrlFetchApp.fetch(url, { 
      method: 'post', 
      contentType: 'application/json', 
      payload: JSON.stringify(payload) 
    });
    
    const json = JSON.parse(response.getContentText());
    return json.candidates?.[0]?.content?.parts?.[0]?.text || userDraft;
  } catch (e) { 
    console.error("Gemini Error: " + e.toString());
    return userDraft; 
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
   sheet.getRange(rowNumber, changeDataIdx + 1).setValue(JSON.stringify(modifiedRequest));
   sheet.getRange(rowNumber, changeTextIdx + 1).setValue(changeReason);
   sheet.getRange(rowNumber, statusIdx + 1).setValue('PENDIENTE_APROBACION_CAMBIO');
   const originalRequest = mapRowToRequest(sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0]);
   sendModificationRequestEmail(originalRequest, modifiedRequest, changeReason);
   return true;
}

function processModificationDecision(e) {
  const id = e.parameter.id;
  const decision = e.parameter.decision;
  const confirm = e.parameter.confirm;

  if (confirm !== 'true') {
      const url = `${WEB_APP_URL}?action=modify_decision&id=${id}&decision=${decision}&confirm=true`;
      return renderConfirmationPage(
          `Confirmar Decisión`,
          `¿Está seguro de <strong>${decision === 'approve' ? 'APROBAR' : 'RECHAZAR'}</strong> el cambio para ${id}?`,
          decision === 'approve' ? 'SÍ, APROBAR' : 'SÍ, RECHAZAR',
          url,
          decision === 'approve' ? '#059669' : '#D71920'
      );
  }
  
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
     try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
        const idIdx = HEADERS_REQUESTS.indexOf("ID RESPUESTA");
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
             const newJson = sheet.getRange(rowNumber, changeDataIdx + 1).getValue();
             if (!newJson) throw new Error("No hay datos pendientes.");
             const newData = JSON.parse(newJson);
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
             setVal("# PERSONAS QUE VIAJAN", newData.passengers.length);
             const p = newData.passengers || [];
             for(let i=0; i<5; i++) {
                setVal(`CÉDULA PERSONA ${i+1}`, p[i] ? p[i].idNumber : '');
                setVal(`NOMBRE PERSONA ${i+1}`, p[i] ? p[i].name : '');
             }
             setVal("FLAG_CAMBIO_REALIZADO", "CAMBIO GENERADO");
             setVal("DATA_CAMBIO_PENDIENTE (JSON)", ""); 
             setVal("STATUS", "PENDIENTE_OPCIONES");
             setVal("OPCIONES (JSON)", "");
             setVal("SELECCION (JSON)", "");
             setVal("APROBADO POR ÁREA?", "");
             sendEmailRich(requesterEmail, "Cambio Aprobado - Solicitud Reiniciada", 
                 HtmlTemplates.modificationResult(id, 'approve', {requestId: id, requesterEmail: requesterEmail, company: newData.company, site: newData.site})
             );
        } else {
             sheet.getRange(rowNumber, statusIdx + 1).setValue('PENDIENTE_OPCIONES');
             sheet.getRange(rowNumber, changeDataIdx + 1).setValue("");
             sheet.getRange(rowNumber, changeTextIdx + 1).setValue("");
             const originalReq = mapRowToRequest(sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0]);
             sendEmailRich(requesterEmail, "Cambio Rechazado", 
                 HtmlTemplates.modificationResult(id, 'reject', originalReq)
             );
        }
        return renderMessagePage("Acción Completada", decision === 'approve' ? 'Cambio aprobado. El proceso se ha reiniciado.' : 'Cambio rechazado.', '#059669');
     } catch(err) {
        return renderMessagePage("Error", err.toString(), '#D71920');
     } finally {
       lock.releaseLock();
     }
  }
}

// --- CORE FUNCTION RE-IMPLEMENTED ---

function processOptionSelection(e) {
    const id = e.parameter.id;
    const optionId = e.parameter.optionId;
    const confirm = e.parameter.confirm;

    // 1. Validation & Confirmation Dialog
    if (confirm !== 'true') {
      const url = `${WEB_APP_URL}?action=select&id=${id}&optionId=${optionId}&confirm=true`;
      return renderConfirmationPage(
          `Confirmar Selección`,
          `¿Está seguro de que desea seleccionar la <strong>Opción ${optionId}</strong> para la solicitud <strong>${id}</strong>?`,
          `SÍ, SELECCIONAR OPCIÓN ${optionId}`,
          url,
          '#D71920'
      );
    }

    const lock = LockService.getScriptLock();
    if (lock.tryLock(10000)) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
            const lastRow = sheet.getLastRow();
            if (lastRow < 2) throw new Error("Base de datos vacía.");
            
            // Find Row
            const idIdx = HEADERS_REQUESTS.indexOf("ID RESPUESTA");
            const ids = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().flat();
            const rowIndex = ids.map(String).indexOf(String(id));

            if (rowIndex === -1) throw new Error(`Solicitud ${id} no encontrada.`);
            const rowNumber = rowIndex + 2;

            // Check Status
            const statusIdx = HEADERS_REQUESTS.indexOf("STATUS");
            const currentStatus = sheet.getRange(rowNumber, statusIdx + 1).getValue();

            // Idempotency: If already selected or later stage, show info but don't error
            if (currentStatus !== 'PENDIENTE_OPCIONES' && currentStatus !== 'PENDIENTE_SELECCION') {
                return renderMessagePage(
                    'Información', 
                    `Esta solicitud ya fue procesada anteriormente (Estado: ${currentStatus}). No se requieren acciones adicionales.`,
                    '#374151'
                );
            }

            // Load Options
            const optIdx = HEADERS_REQUESTS.indexOf("OPCIONES (JSON)");
            const rawOpts = sheet.getRange(rowNumber, optIdx + 1).getValue();
            if (!rawOpts) throw new Error("No hay opciones registradas para esta solicitud.");
            
            const options = JSON.parse(rawOpts);
            const selectedOpt = options.find(o => o.id === optionId);

            if (!selectedOpt) throw new Error(`La opción ${optionId} no existe en el registro.`);

            // Update Sheet
            const selIdx = HEADERS_REQUESTS.indexOf("SELECCION (JSON)");
            sheet.getRange(rowNumber, selIdx + 1).setValue(JSON.stringify(selectedOpt));
            sheet.getRange(rowNumber, statusIdx + 1).setValue('PENDIENTE_APROBACION');

            // Trigger Approval Email
            const rowData = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
            const requestData = mapRowToRequest(rowData);
            
            sendApprovalRequestEmail(requestData);
            
            return renderMessagePage(
                'Selección Exitosa',
                `Ha seleccionado la <strong>Opción ${optionId}</strong>.<br><br>
                 El sistema ha cambiado el estado a <strong>PENDIENTE APROBACIÓN</strong> y ha notificado a <strong>${requestData.approverEmail}</strong>.`,
                '#059669' // Green Theme
            );

        } catch (e) {
            return renderMessagePage("Error", "Ocurrió un error al procesar la selección: " + e.toString(), '#D71920');
        } finally {
            lock.releaseLock();
        }
    }
    return renderMessagePage("Sistema Ocupado", "El sistema está ocupado procesando otra petición. Intente nuevamente en unos segundos.", '#D71920');
}

function processApprovalFromEmail(e) {
  const id = e.parameter.id;
  const decision = e.parameter.decision; // 'approved' or 'denied'
  const confirm = e.parameter.confirm;

  const decisionLabel = decision === 'approved' ? 'APROBAR' : 'DENEGAR';
  const decisionColor = decision === 'approved' ? '#059669' : '#D71920';

  if (confirm !== 'true') {
      const url = `${WEB_APP_URL}?action=approve&id=${id}&decision=${decision}&confirm=true`;
      return renderConfirmationPage(
          `Confirmar Decisión`,
          `¿Está seguro de <strong>${decisionLabel}</strong> la solicitud <strong>${id}</strong>?`,
          `SÍ, ${decisionLabel}`,
          url,
          decisionColor
      );
  }

  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
      try {
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
          if (!sheet) throw new Error("Base de datos no encontrada.");
          
          const lastRow = sheet.getLastRow();
          const idIdx = HEADERS_REQUESTS.indexOf("ID RESPUESTA");
          const statusIdx = HEADERS_REQUESTS.indexOf("STATUS");
          
          // Buscar la fila para verificar estado actual antes de escribir
          const ids = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().flat();
          const rowIndex = ids.map(String).indexOf(String(id));

          if (rowIndex === -1) throw new Error(`Solicitud ${id} no encontrada.`);
          const rowNumber = rowIndex + 2;

          const currentStatus = sheet.getRange(rowNumber, statusIdx + 1).getValue();

          // Chequeo de Idempotencia: Si ya está resuelta, no hacer nada y avisar.
          if (currentStatus === 'APROBADO' || currentStatus === 'DENEGADO' || currentStatus === 'PROCESADO') {
              return renderMessagePage(
                  'Información', 
                  `Esta solicitud ya fue procesada anteriormente.<br><br>Estado actual: <strong>${currentStatus}</strong>.`,
                  '#374151'
              );
          }

          // Proceder con la actualización
          const result = updateRequestStatus(id, decision === 'approved' ? 'APROBADO' : 'DENEGADO', {});
          
          if (result === true) {
              return renderMessagePage(
                  'Decisión Registrada', 
                  `La solicitud <strong>${id}</strong> ha sido actualizada a estado <strong>${decision === 'approved' ? 'APROBADO' : 'DENEGADO'}</strong> exitosamente.`,
                  decisionColor
              );
          } else {
              throw new Error("No se pudo actualizar el estado de la solicitud.");
          }

      } catch (e) {
          console.error(e);
          return renderMessagePage('Error', 'Ocurrió un error al procesar la aprobación: ' + e.toString(), '#D71920');
      } finally {
          lock.releaseLock();
      }
  }
  
  return renderMessagePage("Sistema Ocupado", "El sistema está ocupado procesando otra petición. Intente nuevamente en unos segundos.", '#D71920');
}

// --- EMAIL HELPERS & TEMPLATES ---

function getStandardSubject(data) {
    const id = data.requestId || data.id; 
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

const HtmlTemplates = {
    layout: function(title, content, headerColor, mainTitle) {
        const color = headerColor || '#D71920';
        const titleText = mainTitle || 'Gestión de Viajes';
        
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
                        <table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width: 600px; width: 100%; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.1);">
                            <tr>
                                <td style="background-color: ${color}; padding: 30px 20px; text-align: center; border-bottom: 4px solid rgba(0,0,0,0.1);">
                                    <h1 style="color: #ffffff; margin: 0; font-size: 24px; font-weight: 700; letter-spacing: 0.5px; text-transform: uppercase;">
                                        ${titleText}
                                    </h1>
                                    <p style="color: #ffffff; margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;">
                                        ID: ${title}
                                    </p>
                                </td>
                            </tr>
                            <tr>
                                <td style="padding: 30px 40px; color: #333333; line-height: 1.6;">
                                    ${content}
                                </td>
                            </tr>
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
        const passengerList = (data.passengers || []).map(p => 
            `<li style="margin-bottom: 5px;">${p.name} <span style="color:#666; font-size:12px;">(${p.idNumber})</span></li>`
        ).join('');

        const content = `
            <p style="font-size: 15px; color: #555; margin-bottom: 25px;">
                Se ha registrado un nuevo requerimiento de viaje para <a href="mailto:${data.requesterEmail}" style="color: #0056b3; text-decoration: none;">${data.requesterEmail}</a>.
            </p>
            <div style="background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden; margin-bottom: 30px;">
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
                <tr><td width="35%" style="color: #6b7280;">Empresa / Sede:</td><td style="font-weight: 600; color: #111827;">${data.company} - ${data.site}</td></tr>
                <tr><td style="color: #6b7280;">Centro de Costos:</td><td style="font-weight: 600; color: #111827;">${data.costCenter} (${data.businessUnit})</td></tr>
                <tr><td style="color: #6b7280;">Aprobador:</td><td><a href="mailto:${data.approverEmail}" style="color: #0056b3;">${data.approverEmail}</a></td></tr>
                <tr><td style="color: #6b7280;">Hospedaje:</td><td style="font-weight: 600; color: #0056b3;">${data.requiresHotel ? `Sí - ${data.hotelName} (${data.nights} Noches)` : 'No'}</td></tr>
            </table>
            ${data.comments ? `<div style="background-color: #fffbeb; border: 1px solid #fcd34d; border-radius: 6px; padding: 15px; margin-bottom: 20px;"><strong style="display: block; font-size: 12px; color: #92400e; text-transform: uppercase; margin-bottom: 4px;">Observaciones / Notas:</strong><span style="font-size: 14px; color: #b45309; font-style: italic;">${data.comments}</span></div>` : ''}
            <div style="background-color: #eff6ff; border: 1px solid #bfdbfe; border-radius: 6px; padding: 15px; margin-bottom: 30px;">
                 <strong style="display: block; font-size: 12px; color: #1e40af; text-transform: uppercase; margin-bottom: 8px;">👥 Pasajero(s) (${data.passengers.length}):</strong>
                 <ul style="margin: 0; padding-left: 20px; color: #1e3a8a; font-size: 14px;">${passengerList}</ul>
            </div>
            <div style="text-align: center;">
                <a href="${link}" style="background-color: #111827; color: #ffffff; padding: 14px 32px; text-decoration: none; border-radius: 30px; font-weight: bold; font-size: 14px; display: inline-block; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">Ingresar a la Plataforma</a>
            </div>
        `;
        return this.layout(`${requestId}`, content);
    },

    optionsAvailable: function(request, options, link) {
        let optionsHtml = '';
        if (options && options.length) {
            optionsHtml = options.map(opt => {
                const selectLink = `${WEB_APP_URL}?action=select&id=${request.requestId}&optionId=${opt.id}&confirm=true`;
                return `
                <div style="border: 1px solid #e5e7eb; border-radius: 8px; margin-bottom: 20px; background-color: #ffffff; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 15px; border-bottom: 1px solid #f3f4f6; background-color: #fafafa;">
                        <span style="font-weight: 800; color: #D71920; font-size: 18px;">Opción ${opt.id}</span>
                        <span style="font-weight: 800; color: #111827; font-size: 18px;">$ ${Number(opt.totalPrice).toLocaleString()}</span>
                    </div>
                    <div style="padding: 15px; font-size: 14px; color: #4b5563; line-height: 1.5;">
                        <div style="margin-bottom: 10px;">
                            <strong style="color: #111827; display: block; margin-bottom: 2px;">Ida:</strong>
                            <span style="color: #374151;">${opt.outbound.airline} (${opt.outbound.flightNumber || 'N/A'}) - ${opt.outbound.flightTime}</span>
                            <div style="font-size: 12px; color: #6b7280; font-style: italic;">${opt.outbound.notes || ''}</div>
                        </div>
                        ${opt.inbound ? `
                        <div style="margin-bottom: 10px; padding-top: 10px; border-top: 1px dashed #e5e7eb;">
                            <strong style="color: #111827; display: block; margin-bottom: 2px;">Vuelta:</strong>
                            <span style="color: #374151;">${opt.inbound.airline} (${opt.inbound.flightNumber || 'N/A'}) - ${opt.inbound.flightTime}</span>
                            <div style="font-size: 12px; color: #6b7280; font-style: italic;">${opt.inbound.notes || ''}</div>
                        </div>` : ''}
                        ${opt.hotel ? `<div style="margin-top: 10px; padding-top: 10px; border-top: 1px solid #e5e7eb;"><strong style="color: #0056b3;">Hotel:</strong> <span style="color: #1e3a8a;">${opt.hotel}</span></div>` : ''}
                    </div>
                    <div style="padding: 15px; background-color: #ffffff; border-top: 1px solid #f3f4f6; text-align: right;">
                        <a href="${selectLink}" style="background-color: #D71920; color: #ffffff; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 14px; display: inline-block;">Seleccionar esta Opción</a>
                    </div>
                </div>
            `}).join('');
        }

        const content = `
            <p style="color: #374151; margin-bottom: 10px;">Se han cargado las opciones de viaje para su solicitud <strong>${request.requestId}</strong>.</p>
            <p style="color: #374151; margin-bottom: 20px;">Por favor revise las siguientes alternativas y haga clic en "Seleccionar esta Opción" en la que prefiera. Una vez seleccionada, se enviará a aprobación.</p>
            <div style="margin-bottom: 30px;">${optionsHtml}</div>
            <p style="font-size: 12px; color: #9ca3af; font-style: italic; margin-bottom: 20px;">Si ninguna opción se ajusta, por favor contacte al analista de viajes.</p>
            <div style="text-align: center; border-top: 1px solid #e5e7eb; pt-4;">
                <a href="${link}" style="color: #4b5563; text-decoration: underline; font-size: 12px;">Ingresar a la Plataforma</a>
            </div>
        `;
        return this.layout(`${request.requestId}`, content);
    },

    approvalRequest: function(request, selectedOption, approveLink, rejectLink) {
        const passengerList = (request.passengers || []).map(p => 
            `<li style="margin-bottom: 5px;">${p.name} <span style="color:#666; font-size:12px;">(${p.idNumber})</span></li>`
        ).join('');

        const content = `
            <p style="font-size: 15px; color: #555; margin-bottom: 25px;">
                El usuario <a href="mailto:${request.requesterEmail}" style="color: #0056b3; text-decoration: none;">${request.requesterEmail}</a> ha seleccionado una opción de viaje para la solicitud <strong>${request.requestId}</strong> y requiere su aprobación.
            </p>
            
            <!-- TICKET STYLE CARD (Route + Dates) -->
            <div style="background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden; margin-bottom: 30px;">
                <!-- ROUTE ROW -->
                <div style="padding: 20px; border-bottom: 1px solid #e5e7eb;">
                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                        <tr>
                            <td align="left" width="40%">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px;">ORIGEN</span>
                                <span style="display: block; font-size: 18px; font-weight: 800; color: #111827;">${request.origin}</span>
                            </td>
                            <td align="center" width="20%">
                                <span style="color: #d1d5db; font-size: 24px;">➝</span>
                            </td>
                            <td align="right" width="40%">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px;">DESTINO</span>
                                <span style="display: block; font-size: 18px; font-weight: 800; color: #111827;">${request.destination}</span>
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
                                    📅 ${request.departureDate}
                                </div>
                            </td>
                            <td align="center" width="50%" style="padding: 15px;">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; text-transform: uppercase; margin-bottom: 4px;">FECHA REGRESO</span>
                                <div style="display: inline-block; color: #D71920; font-weight: bold; font-size: 14px;">
                                    📅 ${request.returnDate || 'N/A'}
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
                    <td style="font-weight: 600; color: #111827;">${request.company} - ${request.site}</td>
                </tr>
                <tr>
                    <td style="color: #6b7280;">Centro de Costos:</td>
                    <td style="font-weight: 600; color: #111827;">${request.costCenter} ${request.costCenterName ? ` - ${request.costCenterName}` : ''} (${request.businessUnit})</td>
                </tr>
                <tr>
                    <td style="color: #6b7280;">Solicitante:</td>
                    <td><a href="mailto:${request.requesterEmail}" style="color: #0056b3;">${request.requesterEmail}</a></td>
                </tr>
                <tr>
                    <td style="color: #6b7280;">Hospedaje Requerido:</td>
                    <td style="font-weight: 600; color: #0056b3;">${request.requiresHotel ? `Sí (${request.nights} Noches)` : 'No'}</td>
                </tr>
            </table>

            ${request.comments ? `
            <div style="background-color: #fffbeb; border: 1px solid #fcd34d; border-radius: 6px; padding: 15px; margin-bottom: 20px;">
                <div style="display: flex; align-items: flex-start;">
                    <span style="font-size: 16px; margin-right: 10px;">📝</span>
                    <div>
                        <strong style="display: block; font-size: 12px; color: #92400e; text-transform: uppercase; margin-bottom: 4px;">Observaciones / Notas:</strong>
                        <span style="font-size: 14px; color: #b45309; font-style: italic;">${request.comments}</span>
                    </div>
                </div>
            </div>
            ` : ''}

            <div style="background-color: #eff6ff; border: 1px solid #bfdbfe; border-radius: 6px; padding: 15px; margin-bottom: 30px;">
                 <strong style="display: block; font-size: 12px; color: #1e40af; text-transform: uppercase; margin-bottom: 8px;">👥 Pasajero(s) (${request.passengers.length}):</strong>
                 <ul style="margin: 0; padding-left: 20px; color: #1e3a8a; font-size: 14px;">${passengerList}</ul>
            </div>

            <!-- SELECTED OPTION SECTION -->
            <div style="margin-top: 30px; margin-bottom: 30px;">
                <h3 style="font-size: 16px; font-weight: 800; color: #D71920; margin-bottom: 15px; text-transform: uppercase; border-bottom: 2px solid #D71920; padding-bottom: 5px;">
                    Opción Seleccionada (${selectedOption.id})
                </h3>
                
                <div style="border: 1px solid #d1d5db; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                    <!-- PRICE HEADER -->
                    <div style="background-color: #111827; padding: 15px 20px; display: flex; justify-content: space-between; align-items: center;">
                        <span style="color: #9ca3af; font-size: 14px; font-weight: 600; text-transform: uppercase;">Costo Total</span>
                        <span style="color: #ffffff; font-size: 20px; font-weight: 800;">$ ${Number(selectedOption.totalPrice).toLocaleString()}</span>
                    </div>
                    
                    <!-- DETAILS BODY -->
                    <div style="padding: 20px; background-color: white;">
                         <!-- OUTBOUND -->
                         <div style="margin-bottom: 15px;">
                            <div style="display: flex; align-items: center; margin-bottom: 5px;">
                                <span style="font-size: 18px; margin-right: 8px;">🛫</span>
                                <strong style="color: #374151; font-size: 14px;">Vuelo de Ida</strong>
                            </div>
                            <div style="padding-left: 30px; color: #4b5563; font-size: 14px;">
                                <span style="font-weight: 600; color: #111827;">${selectedOption.outbound.airline}</span> 
                                <span style="margin: 0 5px;">|</span> 
                                <span>${selectedOption.outbound.flightTime}</span>
                                ${selectedOption.outbound.flightNumber ? `<span style="color:#6b7280; font-size:12px;"> (${selectedOption.outbound.flightNumber})</span>` : ''}
                                <div style="font-style: italic; font-size: 13px; color: #6b7280; margin-top: 2px;">${selectedOption.outbound.notes || ''}</div>
                            </div>
                         </div>

                         <!-- INBOUND -->
                         ${selectedOption.inbound ? `
                         <div style="margin-bottom: 15px; padding-top: 15px; border-top: 1px dashed #e5e7eb;">
                            <div style="display: flex; align-items: center; margin-bottom: 5px;">
                                <span style="font-size: 18px; margin-right: 8px;">🛬</span>
                                <strong style="color: #374151; font-size: 14px;">Vuelo de Regreso</strong>
                            </div>
                            <div style="padding-left: 30px; color: #4b5563; font-size: 14px;">
                                <span style="font-weight: 600; color: #111827;">${selectedOption.inbound.airline}</span> 
                                <span style="margin: 0 5px;">|</span> 
                                <span>${selectedOption.inbound.flightTime}</span>
                                ${selectedOption.inbound.flightNumber ? `<span style="color:#6b7280; font-size:12px;"> (${selectedOption.inbound.flightNumber})</span>` : ''}
                                <div style="font-style: italic; font-size: 13px; color: #6b7280; margin-top: 2px;">${selectedOption.inbound.notes || ''}</div>
                            </div>
                         </div>
                         ` : ''}

                         <!-- HOTEL -->
                         ${selectedOption.hotel ? `
                         <div style="margin-top: 15px; padding-top: 15px; border-top: 1px solid #e5e7eb;">
                            <div style="display: flex; align-items: center;">
                                <span style="font-size: 18px; margin-right: 8px;">🏨</span>
                                <div>
                                    <strong style="color: #0056b3; font-size: 14px;">Hotel Incluido:</strong>
                                    <span style="color: #1e3a8a; font-size: 14px; margin-left: 5px; font-weight: 600;">${selectedOption.hotel}</span>
                                </div>
                            </div>
                         </div>
                         ` : ''}
                    </div>
                </div>
            </div>

            <!-- ACTION BUTTONS -->
            <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top: 40px;">
                <tr>
                    <td align="center" style="padding-right: 10px;">
                        <a href="${approveLink}" style="background-color: #059669; color: #ffffff; padding: 16px 32px; text-decoration: none; border-radius: 8px; font-weight: 800; display: inline-block; font-size: 15px; box-shadow: 0 4px 6px rgba(5, 150, 105, 0.2);">✅ APROBAR VIAJE</a>
                    </td>
                    <td align="center" style="padding-left: 10px;">
                        <a href="${rejectLink}" style="background-color: #dc2626; color: #ffffff; padding: 16px 32px; text-decoration: none; border-radius: 8px; font-weight: 800; display: inline-block; font-size: 15px; box-shadow: 0 4px 6px rgba(220, 38, 38, 0.2);">❌ RECHAZAR VIAJE</a>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" align="center" style="padding-top: 15px;">
                        <p style="font-size: 12px; color: #6b7280;">Al hacer clic, se registrará su decisión inmediatamente.</p>
                    </td>
                </tr>
            </table>
        `;
        return this.layout(`${request.requestId}`, content);
    },

    decisionNotification: function(request, status) {
        const isApproved = status === 'APROBADO';
        const color = isApproved ? '#059669' : '#dc2626'; 
        const titleText = isApproved ? 'SOLICITUD APROBADA' : 'SOLICITUD DENEGADA';
        const icon = isApproved ? '✅' : '❌';
        
        // Ensure selectedOption exists (it should if approved/denied via flow)
        const selectedOption = request.selectedOption || { id: 'N/A', totalPrice: 0, outbound: {airline: '-', flightTime: '-'} };

        const passengerList = (request.passengers || []).map(p => 
            `<li style="margin-bottom: 5px;">${p.name} <span style="color:#666; font-size:12px;">(${p.idNumber})</span></li>`
        ).join('');

        const content = `
            <div style="text-align: center; margin-bottom: 25px;">
                <div style="font-size: 48px; margin-bottom: 10px;">${icon}</div>
                <p style="font-size: 16px; margin-top: 10px; color: #4b5563;">
                    Su solicitud de viaje <strong>${request.requestId}</strong> ha sido <strong style="color: ${color};">${status}</strong>.
                </p>
            </div>
            
            <!-- TICKET STYLE CARD (Route + Dates) -->
            <div style="background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden; margin-bottom: 30px;">
                <!-- ROUTE ROW -->
                <div style="padding: 20px; border-bottom: 1px solid #e5e7eb;">
                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                        <tr>
                            <td align="left" width="40%">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px;">ORIGEN</span>
                                <span style="display: block; font-size: 18px; font-weight: 800; color: #111827;">${request.origin}</span>
                            </td>
                            <td align="center" width="20%">
                                <span style="color: #d1d5db; font-size: 24px;">➝</span>
                            </td>
                            <td align="right" width="40%">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px;">DESTINO</span>
                                <span style="display: block; font-size: 18px; font-weight: 800; color: #111827;">${request.destination}</span>
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
                                    📅 ${request.departureDate}
                                </div>
                            </td>
                            <td align="center" width="50%" style="padding: 15px;">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; text-transform: uppercase; margin-bottom: 4px;">FECHA REGRESO</span>
                                <div style="display: inline-block; color: #D71920; font-weight: bold; font-size: 14px;">
                                    📅 ${request.returnDate || 'N/A'}
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
                    <td style="font-weight: 600; color: #111827;">${request.company} - ${request.site}</td>
                </tr>
                <tr>
                    <td style="color: #6b7280;">Centro de Costos:</td>
                    <td style="font-weight: 600; color: #111827;">${request.costCenter} ${request.costCenterName ? ` - ${request.costCenterName}` : ''} (${request.businessUnit})</td>
                </tr>
                <tr>
                    <td style="color: #6b7280;">Hospedaje Requerido:</td>
                    <td style="font-weight: 600; color: #0056b3;">${request.requiresHotel ? `Sí (${request.nights} Noches)` : 'No'}</td>
                </tr>
            </table>

            ${request.comments ? `
            <div style="background-color: #fffbeb; border: 1px solid #fcd34d; border-radius: 6px; padding: 15px; margin-bottom: 20px;">
                <div style="display: flex; align-items: flex-start;">
                    <span style="font-size: 16px; margin-right: 10px;">📝</span>
                    <div>
                        <strong style="display: block; font-size: 12px; color: #92400e; text-transform: uppercase; margin-bottom: 4px;">Observaciones / Notas:</strong>
                        <span style="font-size: 14px; color: #b45309; font-style: italic;">${request.comments}</span>
                    </div>
                </div>
            </div>
            ` : ''}

            <div style="background-color: #eff6ff; border: 1px solid #bfdbfe; border-radius: 6px; padding: 15px; margin-bottom: 30px;">
                 <strong style="display: block; font-size: 12px; color: #1e40af; text-transform: uppercase; margin-bottom: 8px;">👥 Pasajero(s) (${request.passengers.length}):</strong>
                 <ul style="margin: 0; padding-left: 20px; color: #1e3a8a; font-size: 14px;">${passengerList}</ul>
            </div>

            <!-- SELECTED OPTION SECTION -->
            <div style="margin-top: 30px; margin-bottom: 30px;">
                <h3 style="font-size: 16px; font-weight: 800; color: ${color}; margin-bottom: 15px; text-transform: uppercase; border-bottom: 2px solid ${color}; padding-bottom: 5px;">
                    Opción Seleccionada (${selectedOption.id})
                </h3>
                
                <div style="border: 1px solid #d1d5db; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                    <!-- PRICE HEADER -->
                    <div style="background-color: #111827; padding: 15px 20px; display: flex; justify-content: space-between; align-items: center;">
                        <span style="color: #9ca3af; font-size: 14px; font-weight: 600; text-transform: uppercase;">Costo Total</span>
                        <span style="color: #ffffff; font-size: 20px; font-weight: 800;">$ ${Number(selectedOption.totalPrice).toLocaleString()}</span>
                    </div>
                    
                    <!-- DETAILS BODY -->
                    <div style="padding: 20px; background-color: white;">
                         <!-- OUTBOUND -->
                         <div style="margin-bottom: 15px;">
                            <div style="display: flex; align-items: center; margin-bottom: 5px;">
                                <span style="font-size: 18px; margin-right: 8px;">🛫</span>
                                <strong style="color: #374151; font-size: 14px;">Vuelo de Ida</strong>
                            </div>
                            <div style="padding-left: 30px; color: #4b5563; font-size: 14px;">
                                <span style="font-weight: 600; color: #111827;">${selectedOption.outbound.airline}</span> 
                                <span style="margin: 0 5px;">|</span> 
                                <span>${selectedOption.outbound.flightTime}</span>
                                ${selectedOption.outbound.flightNumber ? `<span style="color:#6b7280; font-size:12px;"> (${selectedOption.outbound.flightNumber})</span>` : ''}
                                <div style="font-style: italic; font-size: 13px; color: #6b7280; margin-top: 2px;">${selectedOption.outbound.notes || ''}</div>
                            </div>
                         </div>

                         <!-- INBOUND -->
                         ${selectedOption.inbound ? `
                         <div style="margin-bottom: 15px; padding-top: 15px; border-top: 1px dashed #e5e7eb;">
                            <div style="display: flex; align-items: center; margin-bottom: 5px;">
                                <span style="font-size: 18px; margin-right: 8px;">🛬</span>
                                <strong style="color: #374151; font-size: 14px;">Vuelo de Regreso</strong>
                            </div>
                            <div style="padding-left: 30px; color: #4b5563; font-size: 14px;">
                                <span style="font-weight: 600; color: #111827;">${selectedOption.inbound.airline}</span> 
                                <span style="margin: 0 5px;">|</span> 
                                <span>${selectedOption.inbound.flightTime}</span>
                                ${selectedOption.inbound.flightNumber ? `<span style="color:#6b7280; font-size:12px;"> (${selectedOption.inbound.flightNumber})</span>` : ''}
                                <div style="font-style: italic; font-size: 13px; color: #6b7280; margin-top: 2px;">${selectedOption.inbound.notes || ''}</div>
                            </div>
                         </div>
                         ` : ''}

                         <!-- HOTEL -->
                         ${selectedOption.hotel ? `
                         <div style="margin-top: 15px; padding-top: 15px; border-top: 1px solid #e5e7eb;">
                            <div style="display: flex; align-items: center;">
                                <span style="font-size: 18px; margin-right: 8px;">🏨</span>
                                <div>
                                    <strong style="color: #0056b3; font-size: 14px;">Hotel Incluido:</strong>
                                    <span style="color: #1e3a8a; font-size: 14px; margin-left: 5px; font-weight: 600;">${selectedOption.hotel}</span>
                                </div>
                            </div>
                         </div>
                         ` : ''}
                    </div>
                </div>
            </div>

            <!-- FINAL MESSAGE / INSTRUCTIONS -->
            <div style="background-color: ${isApproved ? '#ecfdf5' : '#fef2f2'}; color: ${isApproved ? '#065f46' : '#991b1b'}; padding: 20px; border-radius: 8px; text-align: center; font-size: 15px; border: 1px solid ${isApproved ? '#a7f3d0' : '#fecaca'}; margin-top: 20px; font-weight: 600;">
                ${isApproved ? 'Por favor, ingrese a la plataforma para cargar los soportes (Facturas/Tiquetes) una vez realice la compra/viaje.' : 'Si tiene dudas sobre el rechazo, por favor contacte al analista administrativo.'}
            </div>

            <div style="text-align: center; margin-top: 40px;">
                <a href="${PLATFORM_URL}" style="background-color: #374151; color: white; padding: 14px 32px; text-decoration: none; border-radius: 30px; font-weight: bold; font-size: 14px; display: inline-block; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">Ingresar a la Plataforma</a>
            </div>
        `;
        return this.layout(`${request.requestId}`, content, color, titleText);
    },

    modificationRequest: function(original, modified, reason, approveLink, rejectLink) {
        // Prepare passenger list from modified request
        const passengerList = (modified.passengers || []).map(p => 
            `<li style="margin-bottom: 5px;">${p.name} <span style="color:#666; font-size:12px;">(${p.idNumber})</span></li>`
        ).join('');

        const content = `
            <p style="font-size: 15px; color: #374151; margin-bottom: 15px;">
                El usuario <strong style="color: #0056b3;">${original.requesterEmail}</strong> ha solicitado cambios sustanciales en la solicitud <strong>${original.requestId}</strong>.
            </p>

            <!-- REASON BOX -->
            <div style="background-color: #fffbeb; border-left: 4px solid #f59e0b; padding: 15px; margin-bottom: 25px; border-radius: 4px; border: 1px solid #fcd34d;">
                <strong style="display: block; margin-bottom: 5px; color: #92400e; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">MOTIVO DEL CAMBIO:</strong>
                <span style="font-style: italic; color: #111827; font-size: 15px; display: block;">"${reason}"</span>
            </div>
            
            <p style="font-size: 14px; color: #6b7280; margin-bottom: 15px; text-transform: uppercase; font-weight: 700;">Detalles de la Solicitud Modificada:</p>

            <!-- TRIP CARD -->
            <div style="background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden; margin-bottom: 30px;">
                <div style="padding: 20px; border-bottom: 1px solid #e5e7eb;">
                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                        <tr>
                            <td align="left" width="40%">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px;">ORIGEN</span>
                                <span style="display: block; font-size: 18px; font-weight: 800; color: #111827;">${modified.origin}</span>
                            </td>
                            <td align="center" width="20%">
                                <span style="color: #d1d5db; font-size: 24px;">➝</span>
                            </td>
                            <td align="right" width="40%">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px;">DESTINO</span>
                                <span style="display: block; font-size: 18px; font-weight: 800; color: #111827;">${modified.destination}</span>
                            </td>
                        </tr>
                    </table>
                </div>
                <div style="background-color: #ffffff;">
                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                        <tr>
                            <td align="center" width="50%" style="padding: 15px; border-right: 1px solid #e5e7eb;">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; text-transform: uppercase; margin-bottom: 4px;">FECHA IDA</span>
                                <div style="display: inline-block; color: #D71920; font-weight: bold; font-size: 14px;">
                                    📅 ${modified.departureDate}
                                </div>
                                <div style="font-size: 12px; color: #6b7280; margin-top: 2px;">
                                    (${modified.departureTimePreference || 'N/A'})
                                </div>
                            </td>
                            <td align="center" width="50%" style="padding: 15px;">
                                <span style="display: block; font-size: 10px; color: #9ca3af; font-weight: 700; text-transform: uppercase; margin-bottom: 4px;">FECHA REGRESO</span>
                                <div style="display: inline-block; color: #D71920; font-weight: bold; font-size: 14px;">
                                    📅 ${modified.returnDate || 'N/A'}
                                </div>
                                <div style="font-size: 12px; color: #6b7280; margin-top: 2px;">
                                    (${modified.returnTimePreference || 'Solo Ida'})
                                </div>
                            </td>
                        </tr>
                    </table>
                </div>
            </div>

            <!-- DETAILS TABLE -->
            <h3 style="font-size: 14px; color: #374151; margin-bottom: 15px; border-bottom: 2px solid #f3f4f6; padding-bottom: 5px;">Detalles del Caso (Modificado)</h3>
            <table width="100%" cellpadding="6" cellspacing="0" border="0" style="font-size: 14px; margin-bottom: 25px;">
                <tr><td width="35%" style="color: #6b7280;">Empresa / Sede:</td><td style="font-weight: 600; color: #111827;">${modified.company} - ${modified.site}</td></tr>
                <tr><td style="color: #6b7280;">Centro de Costos:</td><td style="font-weight: 600; color: #111827;">${modified.costCenter} (${modified.businessUnit})</td></tr>
                <tr><td style="color: #6b7280;">Hospedaje:</td><td style="font-weight: 600; color: #0056b3;">${modified.requiresHotel ? `Sí - ${modified.hotelName} (${modified.nights} Noches)` : 'No'}</td></tr>
            </table>

            ${modified.comments ? `<div style="background-color: #fffbeb; border: 1px solid #fcd34d; border-radius: 6px; padding: 15px; margin-bottom: 20px;"><strong style="display: block; font-size: 12px; color: #92400e; text-transform: uppercase; margin-bottom: 4px;">Observaciones / Notas:</strong><span style="font-size: 14px; color: #b45309; font-style: italic;">${modified.comments}</span></div>` : ''}

            <!-- PASSENGERS -->
            <div style="background-color: #eff6ff; border: 1px solid #bfdbfe; border-radius: 6px; padding: 15px; margin-bottom: 30px;">
                 <strong style="display: block; font-size: 12px; color: #1e40af; text-transform: uppercase; margin-bottom: 8px;">👥 Pasajero(s) (${modified.passengers.length}):</strong>
                 <ul style="margin: 0; padding-left: 20px; color: #1e3a8a; font-size: 14px;">${passengerList}</ul>
            </div>

            <!-- ACTIONS -->
            <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top: 20px;">
                <tr>
                    <td align="center" style="padding-right: 10px;">
                        <a href="${approveLink}" style="background-color: #059669; color: #ffffff; padding: 16px 32px; text-decoration: none; border-radius: 8px; font-weight: 800; display: inline-block; font-size: 14px; box-shadow: 0 4px 6px rgba(5, 150, 105, 0.2);">✅ APROBAR CAMBIO</a>
                    </td>
                    <td align="center" style="padding-left: 10px;">
                        <a href="${rejectLink}" style="background-color: #dc2626; color: #ffffff; padding: 16px 32px; text-decoration: none; border-radius: 8px; font-weight: 800; display: inline-block; font-size: 14px; box-shadow: 0 4px 6px rgba(220, 38, 38, 0.2);">❌ RECHAZAR CAMBIO</a>
                    </td>
                </tr>
            </table>
        `;
        return this.layout(`${original.requestId}`, content, '#F59E0B', 'SOLICITUD DE CAMBIO');
    },

    modificationResult: function(id, decision, data) {
        const isApproved = decision === 'approve';
        const title = isApproved ? 'CAMBIO APROBADO' : 'CAMBIO RECHAZADO';
        const color = isApproved ? '#059669' : '#dc2626';
        const content = `
            <div style="text-align: center; margin: 20px 0;">
                <h2 style="color: ${color}; border: 2px solid ${color}; display: inline-block; padding: 10px 20px; border-radius: 8px; text-transform: uppercase; font-size: 16px; font-weight: 800;">${title}</h2>
            </div>
            <p style="color: #374151;">La solicitud de modificación para el ticket <strong>${id}</strong> ha sido gestionada por el administrador.</p>
            ${isApproved ? `<div style="background-color: #ecfdf5; padding: 15px; border-radius: 6px; color: #065f46; font-size: 14px;">Su solicitud ha sido reiniciada.</div>` : `<div style="background-color: #fef2f2; padding: 15px; border-radius: 6px; color: #991b1b; font-size: 14px;">Los cambios no fueron aceptados.</div>`}
            <div style="text-align: center; margin-top: 30px;">
                <a href="${PLATFORM_URL}" style="background-color: #374151; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-size: 14px;">Ir a la Plataforma</a>
            </div>
        `;
        return this.layout(`${id}`, content);
    }
};

// --- RENDER HELPERS ---

function renderConfirmationPage(title, message, actionText, actionUrl, color) {
    const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
      <title>${title}</title>
      <style>
        body { font-family: 'Segoe UI', system-ui, sans-serif; background-color: #f3f4f6; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; }
        .card { background: white; padding: 40px; border-radius: 12px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); max-width: 450px; width: 90%; text-align: center; transition: all 0.3s ease; }
        .icon { font-size: 48px; margin-bottom: 20px; display: block; }
        h1 { color: #1f2937; font-size: 24px; margin-bottom: 10px; }
        p { color: #4b5563; font-size: 16px; line-height: 1.5; margin-bottom: 30px; }
        .btn { display: inline-block; background-color: ${color}; color: white; padding: 14px 28px; text-decoration: none; border-radius: 8px; font-weight: 600; font-size: 16px; transition: opacity 0.2s; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border: none; cursor: pointer; }
        .btn:hover { opacity: 0.9; }
        .footer { margin-top: 20px; font-size: 12px; color: #9ca3af; }
        
        /* Loader Styles */
        .hidden { display: none; }
        .loader {
          border: 4px solid #f3f3f3;
          border-radius: 50%;
          border-top: 4px solid ${color};
          width: 50px;
          height: 50px;
          -webkit-animation: spin 1s linear infinite; /* Safari */
          animation: spin 1s linear infinite;
          margin: 0 auto 20px auto;
        }
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      </style>
      <script>
        function handleAction() {
            // 1. Hide Confirmation Card
            document.getElementById('confirm-card').classList.add('hidden');
            
            // 2. Show Loading Card
            document.getElementById('loading-card').classList.remove('hidden');
            
            // 3. Trigger Navigation
            window.location.href = "${actionUrl}";
        }
      </script>
    </head>
    <body>
      
      <!-- Confirmation State -->
      <div class="card" id="confirm-card">
        <span class="icon">🤔</span>
        <h1>${title}</h1>
        <p>${message}</p>
        <!-- Changed from <a> to <button> to handle UI state before navigation -->
        <button onclick="handleAction()" class="btn">${actionText}</button>
        <div class="footer">Organización Equitel - Gestión de Viajes</div>
      </div>

      <!-- Loading State (Hidden initially) -->
      <div class="card hidden" id="loading-card">
        <div class="loader"></div>
        <h1>Procesando...</h1>
        <p>Estamos registrando su decisión en el sistema. <br>Por favor espere, esto puede tomar unos segundos.</p>
        <div class="footer">No cierre esta ventana</div>
      </div>

    </body>
    </html>`;
    return HtmlService.createHtmlOutput(html)
        .setTitle(title)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderMessagePage(title, message, color) {
    const icon = color === '#059669' ? '✅' : (color === '#D71920' ? '⚠️' : 'ℹ️');
    const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
      <title>${title}</title>
      <style>
        body { font-family: 'Segoe UI', system-ui, sans-serif; background-color: #f3f4f6; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; }
        .card { background: white; width: 100%; max-width: 480px; border-radius: 12px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); overflow: hidden; }
        .header { background-color: ${color}; padding: 30px 20px; text-align: center; }
        .header h1 { color: white; margin: 0; font-size: 22px; text-transform: uppercase; letter-spacing: 1px; }
        .content { padding: 40px 30px; text-align: center; }
        .icon { font-size: 56px; margin-bottom: 20px; display: block; }
        .message { color: #4b5563; font-size: 16px; line-height: 1.6; margin-bottom: 30px; }
        .btn { display: inline-block; padding: 12px 30px; background-color: #111827; color: white; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 14px; }
        .btn:hover { background-color: #000; }
      </style>
    </head>
    <body>
      <div class="card">
        <div class="header"><h1>${title}</h1></div>
        <div class="content">
          <span class="icon">${icon}</span>
          <p class="message">${message}</p>
          <a href="${PLATFORM_URL}" class="btn">Ir a la Plataforma</a>
        </div>
      </div>
    </body>
    </html>`;
    return HtmlService.createHtmlOutput(html)
        .setTitle(title)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Deprecated old helper, mapped to new one
function renderSuccessPage(title, message) {
    return renderMessagePage(title, message, '#059669');
}

// ... (Existing Functions: getRequestsByEmail, getAllRequests, getCostCenterData, getIntegrantesData, createNewRequest, uploadSupportFile, etc.)

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
    if (code.match(/^\d+$/)) { code = code.padStart(4, '0'); }
    return { code: code, name: String(row[1]), businessUnit: String(row[2]) };
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
  const idColIndex = HEADERS_REQUESTS.indexOf("ID RESPUESTA") + 1; 
  const lastRow = sheet.getLastRow();
  let nextIdNum = 1;
  if (lastRow > 1) {
    const existingIds = sheet.getRange(2, idColIndex, lastRow - 1, 1).getValues().flat();
    const numericIds = existingIds.map(val => {
         const strVal = String(val).replace(/^SOL-/, '');
         return parseInt(strVal, 10);
      }).filter(val => !isNaN(val));
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
  row[0] = new Date(); row[1] = data.company; row[2] = data.origin; row[3] = data.destination; row[4] = data.workOrder || ''; row[5] = data.passengers ? data.passengers.length : 1; row[6] = String(data.requesterEmail).toLowerCase().trim(); 
  const p = data.passengers || [];
  for(let i=0; i<5; i++) {
    const baseIdx = 7 + (i*2);
    row[baseIdx] = p[i] ? p[i].idNumber : '';
    row[baseIdx+1] = p[i] ? p[i].name : '';
  }
  row[17] = data.costCenter; row[18] = data.variousCostCenters || ''; row[19] = ccName; row[20] = data.businessUnit; row[21] = data.site; row[22] = data.requiresHotel ? 'Sí' : 'No'; row[23] = data.hotelName || ''; row[24] = nights; row[25] = data.departureDate; row[26] = data.returnDate || ''; row[27] = data.departureTimePreference || ''; row[28] = data.returnTimePreference || ''; row[29] = id; 
  const statusIdx = HEADERS_REQUESTS.indexOf("STATUS"); if (statusIdx > -1) row[statusIdx] = 'PENDIENTE_OPCIONES';
  const obsIdx = HEADERS_REQUESTS.indexOf("OBSERVACIONES"); if (obsIdx > -1) row[obsIdx] = data.comments || '';
  const approverIdx = HEADERS_REQUESTS.indexOf("CORREO DE QUIEN APRUEBA (AUTOMÁTICO)"); if (approverIdx > -1) row[approverIdx] = approverEmail;
  const emailsIdx = HEADERS_REQUESTS.indexOf("CORREOS PASAJEROS (JSON)");
  if (emailsIdx > -1) { const pEmails = data.passengers.map(p => p.email).filter(e => e); row[emailsIdx] = JSON.stringify(pEmails); }
  sheet.appendRow(row);
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
   sheet.getRange(rowNumber, statusIdx + 1).setValue(status);
   if (payload) {
      if (payload.analystOptions) {
         const optIdx = HEADERS_REQUESTS.indexOf("OPCIONES (JSON)");
         if (optIdx > -1) sheet.getRange(rowNumber, optIdx + 1).setValue(JSON.stringify(payload.analystOptions));
      }
      if (payload.selectedOption) {
         const selIdx = HEADERS_REQUESTS.indexOf("SELECCION (JSON)");
         if (selIdx > -1) sheet.getRange(rowNumber, selIdx + 1).setValue(JSON.stringify(payload.selectedOption));
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
      if (supportData.folderId) { try { folder = DriveApp.getFolderById(supportData.folderId); } catch(e) {} }
      if (!folder) {
         const rootFolder = DriveApp.getFolderById(ROOT_DRIVE_FOLDER_ID);
         const folderName = `${requestId}`; 
         const folders = rootFolder.getFoldersByName(folderName);
         if (folders.hasNext()) folder = folders.next();
         else folder = rootFolder.createFolder(folderName);
         supportData.folderId = folder.getId();
         supportData.folderUrl = folder.getUrl();
      }
  } catch(e) { throw new Error("Error accediendo a Google Drive: " + e.toString()); }
  const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
  const file = folder.createFile(blob);
  const newFileEntry = { id: file.getId(), name: file.getName(), url: file.getUrl(), mimeType: mimeType, date: new Date().toISOString() };
  supportData.files.push(newFileEntry);
  sheet.getRange(rowNumber, supportIdx + 1).setValue(JSON.stringify(supportData));
  return supportData;
}

function getCCList(request) {
    const requester = request.requesterEmail;
    const passengerEmails = (request.passengers || []).map(p => p.email).filter(e => e && e.toLowerCase() !== requester.toLowerCase());
    return passengerEmails.join(',');
}

function findApprover(costCenter) { return 'dsanchez@equitel.com.co'; }
function isUserAnalyst(email) { return email.includes('admin') || email.includes('compras') || email.includes('analista') || email === ADMIN_EMAIL; }

function sendModificationRequestEmail(original, modified, reason) {
   const approveLink = `${WEB_APP_URL}?action=modify_decision&id=${original.requestId}&decision=approve`;
   const rejectLink = `${WEB_APP_URL}?action=modify_decision&id=${original.requestId}&decision=reject`;
   const htmlBody = HtmlTemplates.modificationRequest(original, modified, reason, approveLink, rejectLink);
   const subject = getStandardSubject(original) + " - SOLICITUD DE CAMBIO";
   sendEmailRich(ADMIN_EMAIL, subject, htmlBody);
}

function sendNewRequestNotification(data, requestId) {
    const subjectData = { ...data, requestId: requestId };
    const htmlBody = HtmlTemplates.newRequest(data, requestId, PLATFORM_URL);
    const subject = getStandardSubject(subjectData);
    const ccEmails = [data.requesterEmail, getCCList(data)].filter(e => e).join(',');
    try { MailApp.sendEmail({ to: ADMIN_EMAIL, cc: ccEmails, subject: subject, htmlBody: htmlBody }); } catch (e) { console.error("Error sending new req email: " + e); }
}

function sendOptionsToRequester(recipient, request, options) {
   const link = PLATFORM_URL; 
   const htmlBody = HtmlTemplates.optionsAvailable(request, options, link);
   const subject = getStandardSubject(request); 
   const ccList = getCCList(request);
   try { MailApp.sendEmail({ to: recipient, cc: ccList, subject: subject, htmlBody: htmlBody }); } catch(e) { console.error("Error sending options email: " + e); }
}

function sendDecisionNotification(request, status) {
  const htmlBody = HtmlTemplates.decisionNotification(request, status);
  const subject = getStandardSubject(request); 
  const ccList = [ADMIN_EMAIL, getCCList(request)].join(',');
  try { MailApp.sendEmail({ to: request.requesterEmail, cc: ccList, subject: subject, htmlBody: htmlBody }); } catch(e){ console.error("Error sending decision email: " + e); }
}

function sendApprovalRequestEmail(request) {
    if (!request.selectedOption) { console.error("Cannot send approval request without selected option"); return; }
    const approveLink = `${WEB_APP_URL}?action=approve&id=${request.requestId}&decision=approved`;
    const rejectLink = `${WEB_APP_URL}?action=approve&id=${request.requestId}&decision=denied`;
    const htmlBody = HtmlTemplates.approvalRequest(request, request.selectedOption, approveLink, rejectLink);
    const subject = getStandardSubject(request) + " - SOLICITUD DE APROBACIÓN";
    sendEmailRich(request.approverEmail, subject, htmlBody);
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
    const rawOpt = getValue("OPCIONES (JSON)"); if (rawOpt) analystOptions = JSON.parse(rawOpt); 
    const rawSel = getValue("SELECCION (JSON)"); if (rawSel) selectedOption = JSON.parse(rawSel);
    const rawSup = getValue("SOPORTES (JSON)"); if (rawSup) supportData = JSON.parse(rawSup);
  } catch(e) {}
  let pendingChangeData = undefined;
  const changeText = getValue("TEXTO_CAMBIO");
  const changeFlag = getValue("FLAG_CAMBIO_REALIZADO") === "CAMBIO GENERADO";
  try { const rawPending = getValue("DATA_CAMBIO_PENDIENTE (JSON)"); if (rawPending) pendingChangeData = rawPending; } catch(e) {}

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
    costCenterName: String(getValue("NOMBRE CENTRO DE COSTOS (AUTOMÁTICO)")), // ADDED
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
    pendingChangeData: String(getValue("DATA_CAMBIO_PENDIENTE (JSON)") || ''),
    changeReason: String(changeText || ''),
    hasChangeFlag: changeFlag
  };
}
