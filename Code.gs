const DEFAULT_SPREADSHEET_ID = "1PDKeKnVIAPv7o221z5lnjhqWUTlzyMWVtmZhk-lAE-0";
const EMAIL_LOG_SHEET = "REGISTRO CORREOS GERENCIA";
const EMAIL_LOG_HEADERS = [
  "Fecha de registro",
  "Estado del envío",
  "Destinatario",
  "Asunto",
  "Tipo de comunicación",
  "Prioridad",
  "Fecha y hora del hecho",
  "Sector / ubicación",
  "Responsable",
  "Descripción",
  "Acciones realizadas",
  "Solicitud a Gerencia",
  "Enviado por",
  "Detalle del error",
  "CC"
];

const ALLOWED_SHEETS = new Set([
  "CRONOGRAMA",
  "MOVIMIENTOS",
  "TAREAS TECNICAS",
  "ESTRUCTURA_CAMARAS",
  "CAMARAS",
  "CONTROL DE TEMPERATURAS",
  "CREDENCIALES",
  "SOLICITUD DE UNIFORMES",
  "NOVEDADES",
  "PEDIDO DE UTILES",
  "CAPACITACIONES",
  "PERSONAL"
]);

function doGet(event) {
  try {
    const view = String(event && event.parameter && event.parameter.vista || "").toLowerCase();
    if (view === "correo") {
      return HtmlService
        .createHtmlOutputFromFile("CorreoGerencia")
        .setTitle("Envío Formal de Correo | Puesto Uno")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    if (view === "operaciones") {
      const requestedSection = String(event && event.parameter && event.parameter.seccion || "").toLowerCase();
      const availableSections = new Set([
        "inicio",
        "temperatura",
        "credenciales",
        "uniformes",
        "novedades",
        "utiles",
        "capacitaciones",
        "tareas-tecnicas",
        "admin"
      ]);
      const template = HtmlService.createTemplateFromFile("OperacionesApp");
      template.initialSection = availableSections.has(requestedSection) ? requestedSection : "inicio";
      return HtmlService
        .createHtmlOutput(template.evaluate().getContent())
        .setTitle("Puesto Uno | Módulos operativos")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    requireAuthorizedUser_();
    const sheetName = validateSheetName_(event && event.parameter && event.parameter["pestaña"]);
    const action = String(event && event.parameter && event.parameter.accion || "").toUpperCase();
    if (action !== "LEER") {
      throw new Error("Acción GET no permitida.");
    }

    const sheet = getSpreadsheet_().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`No existe la pestaña "${sheetName}" en la hoja configurada.`);
    }
    return jsonOutput_(sheet.getDataRange().getDisplayValues());
  } catch (error) {
    return jsonOutput_({ ok: false, error: getErrorMessage_(error) });
  }
}

function doPost(event) {
  const lock = LockService.getScriptLock();
  try {
    requireAuthorizedUser_();
    const parameters = event && event.parameter || {};
    const action = String(parameters.accion || "").toUpperCase();
    const body = JSON.parse(event && event.postData && event.postData.contents || "{}");

    if (action === "ENVIAR_CORREO_GERENCIA") {
      if (!body.datos || typeof body.datos !== "object" || Array.isArray(body.datos)) {
        throw new Error("El cuerpo debe incluir los campos del informe en 'datos'.");
      }
      requireAuthorizedEditor_();
      lock.waitLock(10000);
      return jsonOutput_({ ok: true, ...enviarCorreoGerencia_(body.datos) });
    }

    if (!Array.isArray(body.datos)) {
      throw new Error("El cuerpo debe incluir un arreglo 'datos'.");
    }

    const sheetName = validateSheetName_(parameters["pestaña"]);
    const sheet = getSpreadsheet_().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`No existe la pestaña "${sheetName}" en la hoja configurada.`);
    }

    lock.waitLock(10000);
    if (action === "ESCRIBIR") {
      const row = sanitizeRow_(body.datos);
      if (row.length === 0) {
        throw new Error("No se puede guardar una fila vacía.");
      }
      sheet.appendRow(row);
    } else if (action === "ACTUALIZAR_MATRIZ") {
      requireAuthorizedEditor_();
      const matrix = sanitizeMatrix_(body.datos);
      if (matrix.length === 0 || matrix[0].length === 0) {
        throw new Error("No se puede reemplazar la hoja con una matriz vacía.");
      }
      matrix.forEach((row, index) => {
        const sheetRow = index + 1;
        if (sheetRow >= 2 && (sheetRow - 2) % 3 === 0) {
          if (row.length > 0) {
            sheet.getRange(sheetRow, 1, 1, row.length).setValues([row]);
          }
        }
      });
    } else if (action === "ENVIAR_REPORTE_GERENCIA") {
      requireAuthorizedEditor_();
      enviarReporteGerencia_(body.datos);
    } else {
      throw new Error("Acción POST no permitida.");
    }

    return jsonOutput_({ ok: true });
  } catch (error) {
    return jsonOutput_({ ok: false, error: getErrorMessage_(error) });
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function leerDatosParaInterfaz(pestana) {
  requireAuthorizedUser_();
  const sheetName = validateSheetName_(pestana);
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`No existe la pestaña "${sheetName}" en la hoja configurada.`);
  }
  return sheet.getDataRange().getDisplayValues();
}

function escribirDatosDesdeInterfaz(pestana, accion, datos) {
  requireAuthorizedUser_();
  const sheetName = validateSheetName_(pestana);
  const action = String(accion || "").toUpperCase();
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`No existe la pestaña "${sheetName}" en la hoja configurada.`);
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (action === "ESCRIBIR") {
      if (!Array.isArray(datos)) {
        throw new Error("Los datos a guardar deben ser un arreglo.");
      }
      const row = sanitizeRow_(datos);
      if (row.length === 0) {
        throw new Error("No se puede guardar una fila vacía.");
      }
      sheet.appendRow(row);
    } else if (action === "ACTUALIZAR_MATRIZ") {
      requireAuthorizedEditor_();
      if (!Array.isArray(datos)) {
        throw new Error("La matriz a actualizar debe ser un arreglo.");
      }
      const matrix = sanitizeMatrix_(datos);
      if (matrix.length === 0 || matrix[0].length === 0) {
        throw new Error("No se puede reemplazar la hoja con una matriz vacía.");
      }
      matrix.forEach((row, index) => {
        const sheetRow = index + 1;
        if (sheetRow >= 2 && (sheetRow - 2) % 3 === 0 && row.length > 0) {
          sheet.getRange(sheetRow, 1, 1, row.length).setValues([row]);
        }
      });
    } else {
      throw new Error("Acción de escritura no permitida.");
    }
    return { ok: true };
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Puesto Uno")
    .addItem("Enviar informe de tarea técnica a Gerencia", "enviarReporteGerencia")
    .addItem("Abrir registro de correos a Gerencia", "abrirRegistroCorreosGerencia")
    .addToUi();
}

function abrirRegistroCorreosGerencia() {
  const sheet = getOrCreateEmailLogSheet_();
  sheet.activate();
  SpreadsheetApp.flush();
  return `Pestaña "${EMAIL_LOG_SHEET}" lista en "${sheet.getParent().getName()}".`;
}

function autorizarMiCuentaParaCorreoGerencia() {
  const email = Session.getActiveUser().getEmail().trim().toLowerCase();
  if (!email) {
    throw new Error("Google no informó el correo de la cuenta activa. Abrí este proyecto con la cuenta que debe enviar los informes e intentá nuevamente.");
  }

  const properties = PropertiesService.getScriptProperties();
  ["ALLOWED_USERS", "EDITORS"].forEach(propertyName => {
    const configuredUsers = getConfiguredEmailList_(propertyName);
    if (!configuredUsers.includes(email)) {
      configuredUsers.push(email);
      properties.setProperty(propertyName, configuredUsers.join(","));
    }
  });

  const result = {
    email: email,
    allowed: true,
    editor: true
  };
  Logger.log(JSON.stringify(result));
  return result;
}

function enviarReporteGerencia() {
  requireAuthorizedEditor_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName().trim().toUpperCase() !== "TAREAS TECNICAS") {
    throw new Error('Abrí la pestaña "TAREAS TECNICAS" antes de enviar el informe.');
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("Todavía no hay tareas técnicas para informar.");
  }

  const values = sheet.getRange(lastRow, 1, 1, 11).getDisplayValues()[0];
  enviarReporteGerencia_(values);
  SpreadsheetApp.getUi().alert(`Informe de orden Nº ${values[0]} enviado a Gerencia.`);
}

function enviarReporteGerencia_(values) {
  if (!Array.isArray(values) || values.length < 10) {
    throw new Error("El informe requiere al menos las columnas A:J de la tarea técnica.");
  }

  const [
    nroOrden,
    fecha,
    horaLlegada,
    nroAbonado,
    direccion,
    empresa,
    tecnicos,
    trabajoRealizado,
    observaciones,
    solucionado,
    urlFoto
  ] = values.map(value => value === null || value === undefined ? "" : String(value));

  if (!nroOrden.trim()) {
    throw new Error("El número de orden es obligatorio para enviar el informe.");
  }

  const estadoSolucion = solucionado.trim().toUpperCase();
  const estaSolucionado = estadoSolucion === "SI" || estadoSolucion === "SÍ";
  const asuntoEstado = estaSolucionado ? "SOLUCIONADO" : "PENDIENTE";
  const asunto = `REPORTE CCTV: Orden Nº ${nroOrden} [${asuntoEstado}]`;
  const textoEstado = estaSolucionado ? "SÍ - SOLUCIONADO" :
    estadoSolucion === "NO" ? "NO - REQUIERE ACCIÓN" : "No especificado";
  const colorEstado = estaSolucionado ? "#16a34a" :
    estadoSolucion === "NO" ? "#dc2626" : "#475569";
  const foto = crearBloqueFoto_(urlFoto);
  const escape = escapeHtml_;

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:600px;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;">
      <div style="background:#1a365d;color:#fff;padding:20px;text-align:center;">
        <h2 style="margin:0;font-size:20px;">CONSTANCIA DE TAREAS TÉCNICAS</h2>
        <p style="margin:5px 0 0;font-size:14px;">Mantenimiento de sistema de cámaras CCTV</p>
      </div>
      <div style="padding:24px;background:#fff;color:#333;line-height:1.6;">
        <p>Estimada Gerencia:</p>
        <p>Compartimos el reporte de novedades correspondiente a la asistencia técnica:</p>
        <table style="width:100%;border-collapse:collapse;margin:20px 0;font-size:14px;">
          ${reportRow_("Número de orden", nroOrden)}
          ${reportRow_("Estado", textoEstado, colorEstado)}
          ${reportRow_("Servicio / empresa", empresa)}
          ${reportRow_("Número de abonado", nroAbonado)}
          ${reportRow_("Dirección", direccion)}
          ${reportRow_("Fecha de visita", fecha)}
          ${reportRow_("Hora de llegada", horaLlegada)}
          ${reportRow_("Técnico(s)", tecnicos)}
        </table>
        <div style="margin-top:20px;padding:14px;background:#f0fdf4;border-left:4px solid #16a34a;">
          <strong>Trabajo realizado</strong>
          <p style="margin:6px 0 0;white-space:pre-wrap;">${escape(trabajoRealizado)}</p>
        </div>
        <div style="margin-top:14px;padding:14px;background:#fef3c7;border-left:4px solid #d97706;">
          <strong>Observaciones técnicas</strong>
          <p style="margin:6px 0 0;white-space:pre-wrap;">${escape(observaciones || "Sin novedades adicionales.")}</p>
        </div>
        ${foto}
        <p style="margin-top:24px;font-size:11px;color:#64748b;text-align:center;">Informe digital emitido desde Puesto Uno.</p>
      </div>
    </div>`;

  MailApp.sendEmail({
    to: "porteriacooperativa2017@gmail.com",
    subject: asunto,
    htmlBody: htmlBody,
    body: [
      `REPORTE CCTV: Orden Nº ${nroOrden}`,
      `Estado: ${textoEstado}`,
      `Empresa: ${empresa}`,
      `Dirección: ${direccion}`,
      `Trabajo: ${trabajoRealizado}`,
      `Observaciones: ${observaciones}`
    ].join("\n")
  });
}

function enviarCorreoGerencia_(report) {
  const recipients = validateEmailRecipients_(report.destinatarios);
  const toRecipients = recipients.filter(recipient => recipient.tipo === "PARA").map(recipient => recipient.email);
  const ccRecipients = recipients.filter(recipient => recipient.tipo === "CC").map(recipient => recipient.email);
  const recipient = toRecipients.join(", ");
  const cc = ccRecipients.join(", ");
  const subjectLine = String(report.asunto || "").replace(/[\r\n]+/g, " ").trim();
  const allowedTypes = [
    "Informe de situación",
    "Novedad relevante",
    "Incidente operativo",
    "Solicitud de decisión",
    "Seguimiento"
  ];
  const allowedPriorities = ["Normal", "Media", "Alta", "Urgente"];
  const type = String(report.tipo || "").trim();
  const priority = String(report.prioridad || "").trim();
  const incidentDate = String(report.fecha || "").trim();
  const sector = requiredReportField_(report.sector, "sector / ubicación", 160);
  const reporter = requiredReportField_(report.responsable, "responsable que informa", 120);
  const description = requiredReportField_(report.descripcion, "descripción de la situación", 5000);
  const actions = requiredReportField_(report.acciones, "acciones realizadas", 4000);
  const request = requiredReportField_(report.solicitud, "solicitud a Gerencia", 3000);

  if (!subjectLine || subjectLine.length > 150) {
    throw new Error("El asunto es obligatorio y no puede superar los 150 caracteres.");
  }
  if (!allowedTypes.includes(type)) {
    throw new Error("El tipo de comunicación seleccionado no es válido.");
  }
  if (!allowedPriorities.includes(priority)) {
    throw new Error("La prioridad seleccionada no es válida.");
  }
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(incidentDate) || isNaN(new Date(incidentDate).getTime())) {
    throw new Error("Ingresá una fecha y hora válidas para el hecho.");
  }

  const greetingName = String(report.nombreDestinatario || "").trim().slice(0, 100) ||
    guessRecipientName_(toRecipients[0]);
  const greeting = greetingName ? `Estimado/a ${greetingName}:` : "Estimado/a:";
  const priorityColors = {
    Normal: "#2563eb",
    Media: "#d97706",
    Alta: "#ea580c",
    Urgente: "#b91c1c"
  };
  const safe = escapeHtml_;
  const paragraph = value => safe(value).replace(/\r?\n/g, "<br>");
  const detailBlock = (title, value, accent, background) => `
    <section style="margin:16px 0;padding:15px 16px;background:${background};border-left:4px solid ${accent};border-radius:4px;">
      <h3 style="margin:0 0 8px;color:#1e293b;font-size:14px;">${safe(title)}</h3>
      <p style="margin:0;color:#334155;font-size:14px;line-height:1.65;">${paragraph(value)}</p>
    </section>`;
  const subject = `[${priority.toUpperCase()}] Puesto Uno | ${type}: ${subjectLine}`;
  const htmlBody = `
    <div style="max-width:720px;margin:0 auto;border:1px solid #dbe3ee;border-radius:10px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;color:#1e293b;">
      <header style="padding:24px 28px;background:#10243d;color:#fff;">
        <div style="font-size:11px;font-weight:bold;letter-spacing:1.4px;text-transform:uppercase;color:#9fb4cc;">COFARMEN · PUESTO UNO</div>
        <h1 style="margin:9px 0 4px;font-size:22px;line-height:1.3;">Informe formal a Gerencia</h1>
        <p style="margin:0;color:#d2ddea;font-size:13px;">Comunicación operativa para conocimiento, evaluación y seguimiento</p>
      </header>
      <main style="padding:24px 28px;background:#fff;line-height:1.55;">
        <p style="margin:0 0 16px;font-size:14px;">${safe(greeting)}</p>
        <p style="margin:0 0 20px;color:#475569;font-size:14px;">Por medio del presente, se eleva el siguiente informe para su conocimiento y consideración.</p>
        <div style="margin:0 0 20px;padding:14px 16px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:7px;">
          <div style="margin-bottom:12px;font-size:17px;font-weight:bold;color:#0f172a;">${safe(subjectLine)}</div>
          <table role="presentation" style="width:100%;border-collapse:collapse;font-size:13px;">
            ${reportRow_("Tipo de comunicación", type)}
            ${reportRow_("Prioridad", priority, priorityColors[priority])}
            ${reportRow_("Fecha y hora del hecho", incidentDate.replace("T", " "))}
            ${reportRow_("Sector / ubicación", sector)}
            ${reportRow_("Responsable que informa", reporter)}
          </table>
        </div>
        ${detailBlock("1. Resumen y descripción de la situación", description, "#2563eb", "#eff6ff")}
        ${detailBlock("2. Acciones realizadas y estado actual", actions, "#0f766e", "#f0fdfa")}
        ${detailBlock("3. Solicitud o seguimiento requerido", request, "#b45309", "#fffbeb")}
        <p style="margin:22px 0 0;color:#475569;font-size:14px;">Quedamos a disposición para ampliar la información y realizar el seguimiento que se indique.</p>
        <p style="margin:18px 0 0;font-size:14px;">Atentamente,<br><strong>${safe(reporter)}</strong><br><span style="color:#64748b;">Puesto Uno · COFARMEN</span></p>
      </main>
      <footer style="padding:13px 20px;background:#f1f5f9;border-top:1px solid #e2e8f0;text-align:center;color:#64748b;font-size:10px;">
        Informe generado desde el módulo de Envío Formal de Puesto Uno.
      </footer>
    </div>`;
  const plainBody = [
    "COFARMEN · PUESTO UNO",
    "INFORME FORMAL A GERENCIA",
    "",
    greeting,
    "",
    `Asunto: ${subjectLine}`,
    `Tipo: ${type}`,
    `Prioridad: ${priority}`,
    `Fecha y hora del hecho: ${incidentDate.replace("T", " ")}`,
    `Sector / ubicación: ${sector}`,
    `Responsable que informa: ${reporter}`,
    "",
    "1. RESUMEN Y DESCRIPCIÓN DE LA SITUACIÓN",
    description,
    "",
    "2. ACCIONES REALIZADAS Y ESTADO ACTUAL",
    actions,
    "",
    "3. SOLICITUD O SEGUIMIENTO REQUERIDO",
    request,
    "",
    "Atentamente,",
    reporter,
    "Puesto Uno · COFARMEN"
  ].join("\n");

  const sender = Session.getActiveUser().getEmail();
  const logRow = appendEmailLog_([
    new Date(),
    "EN PROCESO",
    recipient,
    subject,
    type,
    priority,
    incidentDate.replace("T", " "),
    sector,
    reporter,
    description,
    actions,
    request,
    sender,
    "",
    cc
  ]);

  try {
    MailApp.sendEmail({
      to: recipient,
      cc: cc,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      name: "Puesto Uno · COFARMEN"
    });
  } catch (error) {
    updateEmailLogStatus_(logRow, "FALLIDO", getErrorMessage_(error));
    throw error;
  }

  try {
    updateEmailLogStatus_(logRow, "ENVIADO", "");
    return { sent: true, logged: true };
  } catch (error) {
    return {
      sent: true,
      logged: false,
      warning: `El correo se envió, pero no se pudo actualizar su estado en la pestaña "${EMAIL_LOG_SHEET}". El registro queda como EN PROCESO; no lo reenvíes.`
    };
  }
}

function appendEmailLog_(values) {
  const sheet = getOrCreateEmailLogSheet_();
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, values.length).setValues([values.map(sanitizeCell_)]);
  return row;
}

function getOrCreateEmailLogSheet_() {
  const spreadsheet = getSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(EMAIL_LOG_SHEET);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(EMAIL_LOG_SHEET);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, EMAIL_LOG_HEADERS.length).setValues([EMAIL_LOG_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, EMAIL_LOG_HEADERS.length).setFontWeight("bold");
  } else {
    const legacyHeaderCount = EMAIL_LOG_HEADERS.length - 1;
    const headers = sheet.getRange(1, 1, 1, legacyHeaderCount).getDisplayValues()[0];
    if (!EMAIL_LOG_HEADERS.slice(0, legacyHeaderCount).every((header, index) => headers[index] === header)) {
      throw new Error(`La pestaña "${EMAIL_LOG_SHEET}" ya existe, pero sus encabezados no coinciden con el formato esperado.`);
    }
    const ccHeader = sheet.getRange(1, EMAIL_LOG_HEADERS.length).getDisplayValues()[0][0];
    if (!ccHeader) {
      sheet.getRange(1, EMAIL_LOG_HEADERS.length).setValue(EMAIL_LOG_HEADERS[legacyHeaderCount]);
      sheet.getRange(1, EMAIL_LOG_HEADERS.length).setFontWeight("bold");
    } else if (ccHeader !== EMAIL_LOG_HEADERS[legacyHeaderCount]) {
      throw new Error(`La columna ${EMAIL_LOG_HEADERS.length} de "${EMAIL_LOG_SHEET}" ya está ocupada; no se pudo agregar el registro de CC.`);
    }
  }

  return sheet;
}

function updateEmailLogStatus_(row, status, errorMessage) {
  const sheet = getSpreadsheet_().getSheetByName(EMAIL_LOG_SHEET);
  if (!sheet) {
    throw new Error(`No se encuentra la pestaña "${EMAIL_LOG_SHEET}" para actualizar el registro.`);
  }
  sheet.getRange(row, 2).setValue(status);
  sheet.getRange(row, 14).setValue(sanitizeCell_(errorMessage));
}

function validateEmailRecipients_(recipients) {
  if (!Array.isArray(recipients) || recipients.length === 0 || recipients.length > 20) {
    throw new Error("Agregá entre 1 y 20 destinatarios e indicá si cada uno es Para o CC.");
  }

  const seen = new Set();
  const validated = recipients.map((recipient, index) => {
    const email = String(recipient && recipient.email || "").trim();
    const type = String(recipient && recipient.tipo || "").trim().toUpperCase();
    if (!email && type === "CC") return null;
    if (!/^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(email)) {
      throw new Error(`Ingresá un correo válido para el destinatario ${index + 1}.`);
    }
    if (type !== "PARA" && type !== "CC") {
      throw new Error(`Seleccioná Para o CC para el destinatario ${index + 1}.`);
    }
    if (seen.has(email.toLowerCase())) {
      throw new Error(`El correo ${email} está repetido en la lista de destinatarios.`);
    }
    seen.add(email.toLowerCase());
    return { email: email, tipo: type };
  }).filter(Boolean);

  if (!validated.some(recipient => recipient.tipo === "PARA")) {
    throw new Error("Agregá al menos un destinatario principal con tipo Para.");
  }
  return validated;
}

function guessRecipientName_(email) {
  const localPart = String(email || "").split("@")[0];
  const tokens = localPart.split(/[._-]+/).filter(Boolean);
  if (tokens.length === 0 || tokens[0].length > 16 || !/^[a-zA-ZÀ-ÿ]+$/.test(tokens[0])) {
    return "";
  }
  return tokens[0].charAt(0).toLocaleUpperCase("es") + tokens[0].slice(1);
}

function obtenerEstadoAccesoCorreo() {
  const email = Session.getActiveUser().getEmail().trim().toLowerCase();
  if (!email) {
    return {
      ok: false,
      error: "Google no pudo identificar la sesión. Abrí el formulario directamente en Google e iniciá sesión con tu cuenta autorizada."
    };
  }

  if (!getConfiguredEmailList_("ALLOWED_USERS").includes(email)) {
    return {
      ok: false,
      email: email,
      error: `La cuenta identificada es ${email}, pero no está incluida en ALLOWED_USERS de este proyecto de Apps Script.`
    };
  }

  if (!getConfiguredEmailList_("EDITORS").includes(email)) {
    return {
      ok: false,
      email: email,
      error: `La cuenta ${email} está autorizada para acceder, pero no está incluida en EDITORS de este proyecto de Apps Script.`
    };
  }

  return { ok: true, email: email };
}

function enviarCorreoGerencia(report) {
  requireAuthorizedUser_();
  requireAuthorizedEditor_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    return enviarCorreoGerencia_(report);
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function requiredReportField_(value, label, maxLength) {
  const text = String(value || "").trim();
  if (!text) {
    throw new Error(`El campo "${label}" es obligatorio.`);
  }
  if (text.length > maxLength) {
    throw new Error(`El campo "${label}" supera el máximo de ${maxLength} caracteres.`);
  }
  return text;
}

function reportRow_(label, value, color) {
  const cellValue = color
    ? `<span style="display:inline-block;padding:3px 9px;color:#fff;background:${color};border-radius:12px;font-weight:bold;">${escapeHtml_(value)}</span>`
    : escapeHtml_(value);
  return `<tr><td style="padding:8px 6px;border-bottom:1px solid #e2e8f0;font-weight:bold;width:40%;color:#475569;">${escapeHtml_(label)}</td><td style="padding:8px 6px;border-bottom:1px solid #e2e8f0;color:#1e293b;">${cellValue}</td></tr>`;
}

function crearBloqueFoto_(value) {
  const url = String(value || "").trim();
  if (!url) {
    return "";
  }

  const driveUrl = url.match(/^https:\/\/drive\.google\.com\/(?:file\/d\/([a-zA-Z0-9_-]+)|open\?id=([a-zA-Z0-9_-]+))/);
  if (!driveUrl) {
    const safeLink = /^https:\/\//i.test(url)
      ? `<a href="${escapeHtml_(url)}" rel="noopener noreferrer">Ver comprobante</a>`
      : escapeHtml_(url);
    return `<p style="margin-top:18px;">Evidencia digital: ${safeLink}</p>`;
  }

  const fileId = driveUrl[1] || driveUrl[2];
  const imageUrl = `https://drive.google.com/uc?export=view&id=${encodeURIComponent(fileId)}`;
  return `<div style="margin-top:18px;text-align:center;"><h3 style="font-size:15px;text-align:left;">Evidencia digital</h3><a href="${escapeHtml_(url)}" rel="noopener noreferrer">Abrir comprobante</a><br><img src="${imageUrl}" alt="Comprobante de la orden" style="max-width:100%;max-height:440px;margin-top:10px;border:1px solid #cbd5e1;"></div>`;
}

function escapeHtml_(value) {
  return String(value === null || value === undefined ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getSpreadsheet_() {
  const configuredId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  return SpreadsheetApp.openById(configuredId || DEFAULT_SPREADSHEET_ID);
}

function requireAuthorizedUser_() {
  const email = Session.getActiveUser().getEmail().trim().toLowerCase();
  const allowedUsers = getConfiguredEmailList_("ALLOWED_USERS");

  if (!email) {
    throw new Error("Google no pudo identificar la sesión. Abrí el formulario directamente en Google e iniciá sesión con tu cuenta autorizada.");
  }
  if (!allowedUsers.includes(email)) {
    throw new Error("Tu cuenta de Google está identificada, pero no está incluida en ALLOWED_USERS.");
  }
}

function requireAuthorizedEditor_() {
  const email = Session.getActiveUser().getEmail().trim().toLowerCase();
  const editors = getConfiguredEmailList_("EDITORS");

  if (!email) {
    throw new Error("Google no pudo identificar la sesión. Abrí el formulario directamente en Google e iniciá sesión con tu cuenta autorizada.");
  }
  if (!editors.includes(email)) {
    throw new Error("Tu cuenta está autorizada para acceder, pero no está incluida en EDITORS.");
  }
}

function getConfiguredEmailList_(propertyName) {
  return String(PropertiesService.getScriptProperties().getProperty(propertyName) || "")
    .split(",")
    .map(user => user.trim().toLowerCase())
    .filter(Boolean);
}

function validateSheetName_(value) {
  const name = String(value || "").trim().toUpperCase();
  if (!ALLOWED_SHEETS.has(name)) {
    throw new Error("Pestaña no autorizada.");
  }
  return name;
}

function sanitizeRow_(row) {
  return row.map(sanitizeCell_);
}

function sanitizeMatrix_(matrix) {
  if (!matrix.every(Array.isArray)) {
    throw new Error("La matriz contiene una fila no válida.");
  }
  return matrix.map(sanitizeRow_);
}

function sanitizeCell_(value) {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  const text = String(value);
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function jsonOutput_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function getErrorMessage_(error) {
  return error && error.message ? error.message : String(error);
}
