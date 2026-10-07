(() => {
  "use strict";

  const MONTHS = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const TO_RECIPIENTS = [
    "Griselda Dolores Castillo Bacilio <griselda.dolores.castillo@eysglobal.co>",
    "Cristian Camilo Jimenez Pena <camilo.jimenez@huawei.com>",
  ];
  const CC_RECIPIENTS = ["sebastian.rolong_costa@manpowercolombia.com", "ali.villa_costa@manpowercolombia.com"];
  const FEEDBACK_KEY = "control-rechazos-fme-feedback-v1";
  const state = { rows: [], messages: [], baseBytes: null, month: 1, year: 0, activeClosed: null };
  const $ = (id) => document.getElementById(id);
  const norm = (value) => value == null ? "" : String(value).trim();
  const normalizeHeader = (value) => norm(value).replace(/\s+/g, " ").normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase("es");
  const normalizeSite = (value) => norm(value).replace(/\s+/g, "").toUpperCase();
  const normalizeDigits = (value) => norm(value).replace(/\D+/g, "");
  const normalizeName = (value) => norm(value).normalize("NFKD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Za-z0-9]+/g, " ").toUpperCase().replace(/\s+/g, " ").trim();
  const escapeHtml = (value) => norm(value).replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[ch]);

  function setNotice(message, type = "success") {
    const notice = $("app-status");
    notice.textContent = message;
    notice.className = `notice visible ${type}`;
  }

  function clearNotice() {
    $("app-status").className = "notice";
    $("app-status").textContent = "";
  }

  function readFeedbacks() {
    try {
      const value = JSON.parse(localStorage.getItem(FEEDBACK_KEY) || "{}");
      return value && typeof value === "object" && !Array.isArray(value) ? value : {};
    } catch (error) {
      throw new Error(`No se pudo leer el feedback local: ${error.message}`);
    }
  }

  function writeFeedbacks(value) {
    try {
      localStorage.setItem(FEEDBACK_KEY, JSON.stringify(value));
    } catch (error) {
      throw new Error(`No se pudo guardar en el almacenamiento local del navegador: ${error.message}`);
    }
  }

  function readWorkbook(fileBytes) {
    return XLSX.read(fileBytes, { type: "array", cellDates: false });
  }

  function sheetRows(workbook, name) {
    const sheet = workbook.Sheets[name];
    if (!sheet) return [];
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true, blankrows: false });
  }

  function recordsFromSheet(workbook, name) {
    const matrix = sheetRows(workbook, name);
    const headers = (matrix[0] || []).map(norm);
    return matrix.slice(1).filter((row) => row.some((value) => value != null && value !== "")).map((row) => {
      const record = {};
      headers.forEach((header, index) => { if (header) record[header] = row[index] ?? null; });
      return record;
    });
  }

  function getField(record, ...names) {
    const index = new Map(Object.keys(record).map((key) => [normalizeHeader(key), key]));
    for (const name of names) {
      const key = index.get(normalizeHeader(name));
      if (key !== undefined) return record[key];
    }
    return null;
  }

  function dateValue(value) {
    if (value == null || value === "") return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (!parsed) return null;
      return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, Math.round(parsed.S));
    }
    const text = norm(value);
    let match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (match) return new Date(+match[3], +match[2] - 1, +match[1], +(match[4] || 0), +(match[5] || 0), +(match[6] || 0));
    match = text.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (match) return new Date(+match[1], +match[2] - 1, +match[3], +(match[4] || 0), +(match[5] || 0), +(match[6] || 0));
    return null;
  }

  function displayDate(value, withTime = false) {
    const date = dateValue(value);
    if (!date) return "";
    const dd = String(date.getDate()).padStart(2, "0");
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const yyyy = date.getFullYear();
    if (!withTime) return `${dd}/${mm}/${yyyy}`;
    return `${dd}/${mm}/${yyyy} ${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}:${String(date.getSeconds()).padStart(2, "0")}`;
  }

  function extractDocument(value) {
    const text = norm(value);
    const match = text.match(/\((\d{6,12})\)\s*$/);
    if (match) return match[1];
    const digits = normalizeDigits(text);
    return text && digits === text.replace(/\s/g, "") && digits.length >= 6 && digits.length <= 12 ? digits : "";
  }

  function extractName(value) {
    return normalizeName(norm(value).replace(/\(\d{6,12}\)\s*$/, "").trim());
  }

  function normalizeMobile(value) {
    const digits = normalizeDigits(value);
    if (!digits) return "";
    if (digits.startsWith("57") && digits.length === 12) return digits;
    if (digits.length === 10 && digits.startsWith("3")) return `57${digits}`;
    return digits;
  }

  function loadFmeDirectory(workbook) {
    const byDoc = new Map();
    const byName = new Map();
    for (const sheetName of workbook.SheetNames) {
      const matrix = sheetRows(workbook, sheetName);
      const headers = (matrix[0] || []).map(normalizeHeader);
      const nameIndex = headers.findIndex((header) => ["nombre", "nombres"].includes(header));
      const docIndex = headers.findIndex((header) => header === "cedula");
      const phoneIndex = headers.findIndex((header) => ["celular", "telefono"].includes(header));
      if (nameIndex < 0 || docIndex < 0 || phoneIndex < 0) continue;
      matrix.slice(1).forEach((row) => {
        const name = norm(row[nameIndex]);
        const doc = normalizeDigits(row[docIndex]);
        const rawPhone = norm(row[phoneIndex]);
        const rec = { "Nombre FME data": name, "Cedula FME": doc, "Celular FME": normalizeMobile(rawPhone), "Celular original": rawPhone, "Origen celular": sheetName };
        if (!name && !doc) return;
        if (doc && !byDoc.has(doc)) byDoc.set(doc, rec);
        const nameKey = normalizeName(name);
        if (nameKey && !byName.has(nameKey)) byName.set(nameKey, rec);
      });
    }
    return { byDoc, byName };
  }

  function findFmeContact(value, directory) {
    const doc = extractDocument(value);
    if (doc && directory.byDoc.has(doc)) return { ...directory.byDoc.get(doc), "Origen match celular": "CEDULA" };
    const name = extractName(value);
    if (name && directory.byName.has(name)) return { ...directory.byName.get(name), "Origen match celular": "NOMBRE" };
    return { "Nombre FME data": "", "Cedula FME": doc, "Celular FME": "", "Celular original": "", "Origen celular": "No encontrado", "Origen match celular": "No encontrado" };
  }

  function loadSiteDirectory(workbook) {
    const byCode = new Map();
    const byNetco = new Map();
    const bySite = new Map();
    const invalid = new Set(["-", "#N/A", "NOAPLICA"]);
    if (workbook.SheetNames.includes("Hoja2")) {
      sheetRows(workbook, "Hoja2").slice(1).forEach((row) => {
        const coord = norm(row[1]);
        if (!coord) return;
        const data = { Coordinador: coord, Zona: norm(row[0]), Codigo: norm(row[3]), CodigoNetco: norm(row[4]), "Origen coordinador": "Hoja2 - CODIGO (columna D)" };
        const code = normalizeSite(row[3]);
        const netco = normalizeSite(row[4]);
        if (code && !invalid.has(code) && !byCode.has(code)) byCode.set(code, data);
        if (netco && !invalid.has(netco) && !byNetco.has(netco)) byNetco.set(netco, { ...data, "Origen coordinador": "Hoja2 - CodigoNetco (columna E)" });
      });
    }
    if (workbook.SheetNames.includes("Hoja1")) {
      sheetRows(workbook, "Hoja1").slice(1).forEach((row) => {
        const site = normalizeSite(row[0]);
        const coord = norm(row[1]);
        if (site && coord && !bySite.has(site)) bySite.set(site, { Coordinador: coord, "Origen coordinador": "Hoja1 - SITIO (columna A)" });
      });
    }
    return { byCode, byNetco, bySite };
  }

  function findCoordinator(site, directory) {
    const key = normalizeSite(site);
    return directory.byCode.get(key) || directory.byNetco.get(key) || directory.bySite.get(key) || { Coordinador: "Por validar", "Origen coordinador": "No encontrado en DATA COORDINADORES" };
  }

  const CATEGORY_RULES = {
    "Evidencia fotográfica": ["foto", "fotograf", "imagen", "evidencia", "no se observa", "no se visualiza", "debe verse", "verse completo", "registro fotográfico"],
    "Intervención física": ["retirar", "quitar", "abrir", "tapa", "breaker", "rack", "rectificador", "modulo", "módulo", "puerta", "soporte", "instalar", "ajustar"],
    "Limpieza / organización": ["limpieza", "limpiar", "basura", "sobrante", "desmalez", "organizado", "organización", "sellamiento"],
    "Edición de formulario / texto": ["eliminar formulario", "borrar formulario", "quitar formulario", "modificar formulario", "agregar texto", "añadir texto", "adicionar texto", "agregar información", "añadir información", "diligenciar", "corregir texto", "modificar texto", "editar formulario", "editar campo"],
    "Documentación / certificación": ["certific", "document", "soporte documental", "acta", "formato"],
    "Información / códigos": ["codigo", "código", "serial", "marquilla", "referencia", "id del rack", "diligenciar", "información", "micode", "azimut"],
    "Hallazgo / condición técnica": ["hallazgo", "alarma", "aterriz", "tierra", "voltaje", "corriente", "bater", "dañado", "falla"],
  };
  const NON_EDITABLE = new Set(["Edición de formulario / texto", "Intervención física", "Limpieza / organización", "Documentación / certificación", "Información / códigos", "Hallazgo / condición técnica"]);

  function classifyRemark(remark) {
    const text = norm(remark).toLocaleLowerCase("es");
    const categories = Object.entries(CATEGORY_RULES).filter(([, words]) => words.some((word) => text.includes(word.toLocaleLowerCase("es")))).map(([category]) => category);
    return categories.length ? categories : ["Revisión manual"];
  }

  function closedManagement(categories) {
    const hasImages = categories.includes("Evidencia fotográfica");
    const hasNoneditable = categories.some((category) => NON_EDITABLE.has(category));
    if (hasImages && !hasNoneditable) return "Enviar evidencias por correo";
    if (hasImages && hasNoneditable) return "Correo mixto: evidencias + solicitud de exclusión";
    return "Solicitud de exclusión";
  }

  function buildPurpose(categories) {
    const hasImages = categories.includes("Evidencia fotográfica");
    const actions = [];
    if (categories.includes("Edición de formulario / texto")) actions.push("eliminar o modificar formularios, agregar texto o actualizar información dentro de la actividad");
    if (categories.includes("Intervención física")) actions.push("realizar verificaciones o ajustes físicos sobre los equipos o elementos del sitio");
    if (categories.includes("Limpieza / organización")) actions.push("ejecutar actividades físicas de limpieza, organización o adecuación");
    if (categories.includes("Documentación / certificación")) actions.push("complementar documentación o soportes dentro de la actividad");
    if (categories.includes("Información / códigos")) actions.push("actualizar o recapturar información técnica, códigos, seriales o marquillas dentro de la actividad");
    if (categories.includes("Hallazgo / condición técnica")) actions.push("realizar verificaciones o correcciones técnicas en sitio");
    if (hasImages && !actions.length) return "El propósito de este mensaje es atender el requerimiento de auditoría relacionado con evidencias fotográficas. Dado que la actividad se encuentra en estado Closed, las imágenes solicitadas no pueden incorporarse nuevamente dentro de la actividad; sin embargo, sí pueden ser remitidas por correo electrónico para su validación. Por lo anterior, se envían por este medio las evidencias solicitadas para su revisión.";
    const actionText = actions.length > 1 ? `${actions.slice(0, -1).join(", ")} y ${actions.at(-1)}` : actions[0] || "realizar las correcciones solicitadas dentro de la actividad";
    if (hasImages) return `El propósito de este mensaje es atender los hallazgos identificados durante la auditoría. Las evidencias fotográficas solicitadas pueden ser remitidas por correo electrónico y se adjuntan por este medio para su validación. No obstante, el rechazo también requiere ${actionText}. Dado que la actividad se encuentra en estado Closed, no es posible efectuar dichas modificaciones o correcciones dentro del sistema. Por lo anterior, agradecemos gestionar la exclusión correspondiente para los aspectos que no pueden ser corregidos en la actividad cerrada, con el fin de evitar un impacto negativo en la medición de los KPI.`;
    return `El propósito de este mensaje es solicitar formalmente la exclusión de la actividad en referencia. Los hallazgos identificados durante la auditoría requieren ${actionText}. Dado que la actividad se encuentra en estado Closed, no es posible realizar estas modificaciones o correcciones dentro del sistema. Por lo anterior, agradecemos gestionar la exclusión correspondiente con el fin de evitar un impacto negativo en la medición de los KPI.`;
  }

  function coordinatorDirectory(workbook) {
    const existing = new Set();
    workbook.SheetNames.forEach((sheet) => {
      const matrix = sheetRows(workbook, sheet);
      const headers = (matrix[0] || []).map(normalizeHeader);
      const col = headers.indexOf("assignment id");
      if (col >= 0) matrix.slice(1).forEach((row) => { const value = norm(row[col]); if (value) existing.add(value); });
    });
    return existing;
  }

  function reportSheet(workbook) {
    for (const preferred of ["WO Result query Custom", "WO Result query"]) {
      if (workbook.SheetNames.includes(preferred)) return preferred;
    }
    const required = ["assignment id", "audit time", "fme id", "audit remark", "audit status"];
    const match = workbook.SheetNames.find((name) => required.every((header) => (sheetRows(workbook, name)[0] || []).map(normalizeHeader).includes(header)));
    if (!match) throw new Error("No encontré una hoja con las columnas esperadas del reporte de auditoría.");
    return match;
  }

  function processRows(reportWorkbook, year, month, existingIds, coordinators, fmes, feedbacks) {
    const output = [];
    const seen = new Set();
    for (const record of recordsFromSheet(reportWorkbook, reportSheet(reportWorkbook))) {
      const status = norm(getField(record, "Audit status"));
      if (status.toLocaleLowerCase("es") !== "reject") continue;
      const auditDate = dateValue(getField(record, "Audit time"));
      if (!auditDate || auditDate.getFullYear() !== year || auditDate.getMonth() + 1 !== month) continue;
      const assignment = norm(getField(record, "Assignment ID"));
      if (!assignment || seen.has(assignment)) continue;
      seen.add(assignment);
      const fme = norm(getField(record, "FME ID")) || "Sin FME";
      const site = norm(getField(record, "Site ID"));
      const task = norm(getField(record, "Task ID"));
      const coordinatorData = findCoordinator(site, coordinators);
      const fmeData = findFmeContact(fme, fmes);
      const remark = norm(getField(record, "Audit remark"));
      const categories = classifyRemark(remark);
      const taskStatus = norm(getField(record, "Task status"));
      const feedback = feedbacks[assignment] || feedbacks[task] || {};
      const row = {
        Coordinador: coordinatorData.Coordinador, "Origen coordinador": coordinatorData["Origen coordinador"],
        "Assignment ID": assignment, "Create time": getField(record, "Create time"), "Assign time": getField(record, "Assign time"),
        "Assign operator": getField(record, "Assign operator"), "Audit time": auditDate, "Task ID": task,
        "Task status": taskStatus, "Site ID": site, Region: norm(getField(record, "Region")), "FM Office": norm(getField(record, "FM Office")),
        "FME ID": fme, "FME Supplier": norm(getField(record, "FME Supplier")), "Audit remark": remark, "Audit status": status,
        "Feedback Remark": norm(getField(record, "Feedback Remark")), "Feedback Qty": getField(record, "Feedback Qty"),
        "Audit type": norm(getField(record, "Audit type")), "Customer ticket": norm(getField(record, "Customer ticket")),
        "Task title": norm(getField(record, "Task title")), "Assign User": norm(getField(record, "Assign User")),
        "Estado contacto": "Pendiente", "Es nuevo": existingIds.has(assignment) ? "No" : "Sí",
        "Es Closed": taskStatus.toLocaleLowerCase("es") === "closed", "Tipo rechazo": categories.join(" + "),
        Categorias: categories, Proposito: buildPurpose(categories), "Gestion Closed": closedManagement(categories),
        "Fecha de feedback": norm(feedback.fecha), "Estado OWS": norm(feedback.estado), "Feedback OWS": norm(feedback.feedback),
        "Tiene feedback OWS": feedback.fecha && feedback.estado && feedback.feedback ? "Sí" : "No",
        "Nombre FME data": fmeData["Nombre FME data"], "Cedula FME": fmeData["Cedula FME"],
        "Celular FME": fmeData["Celular FME"], "Celular original": fmeData["Celular original"],
        "Origen celular": fmeData["Origen celular"],
      };
      output.push(row);
    }
    return output;
  }

  function whatsappUrl(phone, message) {
    const normalized = normalizeMobile(phone);
    return normalized ? `https://wa.me/${normalized}?text=${encodeURIComponent(message)}` : "";
  }

  function groupMessages(rows) {
    const groups = new Map();
    rows.filter((row) => !row["Es Closed"] && row["Es nuevo"] === "Sí").forEach((row) => {
      const fme = row["FME ID"] || "Sin FME";
      if (!groups.has(fme)) groups.set(fme, []);
      groups.get(fme).push(row);
    });
    return [...groups.entries()].sort(([a], [b]) => a.localeCompare(b, "es", { sensitivity: "base" })).map(([fme, items]) => {
      const sourceName = items.find((row) => row["Nombre FME data"])?.["Nombre FME data"] || "";
      const name = sourceName ? sourceName.split(/\s+/)[0].toLocaleLowerCase("es").replace(/^./, (char) => char.toLocaleUpperCase("es")) : fme.replace(/\s*\(\d{6,12}\)\s*$/, "").split(/\s+/)[0];
      const greeting = name ? `Buen día, ${name}.` : "Buen día.";
      const chunks = items.map((row, index) => `${index + 1}. Sitio: ${row["Site ID"]}\nTask ID: ${row["Task ID"]}\nAudit Remark: ${row["Audit remark"]}`);
      const message = `${greeting} Te comparto los rechazos de auditoría pendientes de corrección:\n\n${chunks.join("\n\n")}\n\nPor favor realizar las correcciones indicadas y confirmar cuando queden gestionadas. Gracias.`;
      const phone = items.find((row) => row["Celular FME"])?.["Celular FME"] || "";
      return {
        Coordinador: items[0].Coordinador, FME: fme, "Cantidad rechazos": items.length,
        Sitios: [...new Set(items.map((row) => row["Site ID"]).filter(Boolean))].sort().join(", "),
        Mensaje: message, "Celular FME": phone,
        "Celular visible": phone.startsWith("57") && phone.length === 12 ? phone.slice(2) : phone,
        "WhatsApp URL": whatsappUrl(phone, message),
        "Origen celular": items.find((row) => row["Celular FME"])?.["Origen celular"] || "No encontrado",
      };
    });
  }

  function hasFeedback(row) {
    return row["Tiene feedback OWS"] === "Sí";
  }

  function renderTable(table, headers, rows) {
    const columnClass = (header) => `col-${String(header).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "")}`;
    table.innerHTML = `<thead><tr>${headers.map((header) => `<th class="${columnClass(header)}">${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${rows.map((row) => `<tr>${headers.map((header) => `<td class="${columnClass(header)}">${escapeHtml(row[header])}</td>`).join("")}</tr>`).join("")}</tbody>`;
  }

  function setOptions(select, values, first = null, preserve = false) {
    const previous = preserve ? [...select.selectedOptions].map((option) => option.value) : [];
    select.innerHTML = "";
    if (first !== null) select.add(new Option(first, first));
    values.forEach((value) => select.add(new Option(value, value)));
    [...select.options].forEach((option) => { if (previous.includes(option.value)) option.selected = true; });
  }

  function selectedValues(select) {
    return [...select.selectedOptions].map((option) => option.value);
  }

  function filterRows(ignoreCoordinator = false) {
    const coordinator = $("filter-coordinator").value;
    const fmes = selectedValues($("filter-fme"));
    const sites = selectedValues($("filter-site"));
    const taskStatuses = selectedValues($("filter-task-status"));
    const owsStatuses = selectedValues($("filter-ows-status"));
    const scope = $("filter-scope").value;
    return state.rows.filter((row) => {
      if (!ignoreCoordinator && coordinator !== "Todos" && row.Coordinador !== coordinator) return false;
      if (fmes.length && !fmes.includes("Todos") && !fmes.includes(row["FME ID"])) return false;
      if (sites.length && !sites.includes("Todos") && !sites.includes(row["Site ID"])) return false;
      if (taskStatuses.length && !taskStatuses.includes("Todos") && !taskStatuses.includes(row["Task status"])) return false;
      if (owsStatuses.length && !owsStatuses.includes("Todos") && !owsStatuses.includes(row["Estado OWS"])) return false;
      if (scope === "Sin feedback" && hasFeedback(row)) return false;
      if (scope === "Con feedback" && !hasFeedback(row)) return false;
      if (scope === "Solo nuevos" && row["Es nuevo"] !== "Sí") return false;
      if (scope === "Solo Closed" && !row["Es Closed"]) return false;
      if (scope === "Solo abiertos" && row["Es Closed"]) return false;
      return true;
    });
  }

  function renderDashboard() {
    const filtered = filterRows();
    const columns = ["Coordinador", "FME", "CELULAR", "SITIO", "TASK ID", "FECHA DE FEEDBACK", "ESTADO", "FEEDBACK"];
    renderTable($("dashboard-table"), columns, filtered.map((row) => ({
      Coordinador: row.Coordinador, FME: row["FME ID"], CELULAR: row["Celular original"] || row["Celular FME"],
      SITIO: row["Site ID"], "TASK ID": row["Task ID"], "FECHA DE FEEDBACK": row["Fecha de feedback"],
      ESTADO: row["Estado OWS"] || "Sin feedback", FEEDBACK: row["Feedback OWS"] || "Sin feedback",
    })));
    $("dashboard-count").textContent = `Mostrando ${filtered.length} actividad(es).`;
    const counts = new Map();
    filterRows(true).forEach((row) => counts.set(row.Coordinador || "Por validar", (counts.get(row.Coordinador || "Por validar") || 0) + 1));
    const summary = [...counts.entries()].sort(([a], [b]) => a.localeCompare(b, "es")).map(([name, count]) => ({ "Coordinador reporte": name, Actividades: count }));
    renderTable($("coordinator-summary"), ["Coordinador reporte", "Actividades"], summary);
  }

  function renderRejections() {
    const fields = ["Nuevo", "Task status", "Coordinador", "Task ID", "Site ID", "FME", "Celular", "Audit time", "Fecha feedback", "Estado OWS", "Feedback OWS", "Audit Remark"];
    const rows = state.rows.map((row) => ({
      Nuevo: row["Es nuevo"], "Task status": row["Task status"], Coordinador: row.Coordinador,
      "Task ID": row["Task ID"], "Site ID": row["Site ID"], FME: row["FME ID"],
      Celular: row["Celular original"] || row["Celular FME"], "Audit time": displayDate(row["Audit time"], true),
      "Fecha feedback": row["Fecha de feedback"], "Estado OWS": row["Estado OWS"],
      "Feedback OWS": row["Feedback OWS"], "Audit Remark": row["Audit remark"],
    }));
    renderTable($("rejections-table"), fields, rows);
  }

  function renderMessages() {
    const container = $("messages-list");
    if (!state.messages.length) {
      container.innerHTML = '<p class="notice visible success">No hay actividades nuevas y abiertas que requieran mensaje al FME.</p>';
      return;
    }
    container.innerHTML = state.messages.map((item, index) => `<article class="message-card">
      <button class="message-head message-toggle" type="button" aria-expanded="false" aria-controls="message-details-${index}">
        <span class="message-summary"><strong>${escapeHtml(item.FME)}</strong><small>${item["Cantidad rechazos"]} rechazo(s) · ${escapeHtml(item.Coordinador)}</small></span>
        <span class="message-site">${escapeHtml(item.Sitios)}</span><span class="message-chevron" aria-hidden="true"></span>
      </button>
      <div class="message-details" id="message-details-${index}" hidden>
      <div class="message-fields">
        <label>Celular FME<input readonly value="${escapeHtml(item["Celular visible"] || "No encontrado")}"></label>
        <label>Origen celular<input readonly value="${escapeHtml(item["Origen celular"])}"></label>
        ${item["WhatsApp URL"] ? `<a class="button primary" href="${escapeHtml(item["WhatsApp URL"])}" target="_blank" rel="noopener noreferrer">Abrir WhatsApp</a>` : "<span class=\"muted\">Sin celular</span>"}
      </div>
      <label>Mensaje listo para copiar<textarea rows="8" id="message-${index}">${escapeHtml(item.Mensaje)}</textarea></label>
      </div>
    </article>`).join("");
    container.querySelectorAll(".message-toggle").forEach((button) => button.addEventListener("click", () => {
      const details = $(button.getAttribute("aria-controls"));
      const open = button.getAttribute("aria-expanded") === "true";
      button.setAttribute("aria-expanded", String(!open));
      details.hidden = open;
      button.closest(".message-card").classList.toggle("is-open", !open);
    }));
  }

  function renderFeedbackOptions() {
    const select = $("feedback-activity");
    const current = select.value;
    select.innerHTML = "";
    state.rows.forEach((row, index) => {
      const label = `${row["Task ID"]} · ${row["Site ID"]} · ${row["FME ID"]} · ${row.Coordinador}`;
      select.add(new Option(label, String(index)));
    });
    if (current && Number(current) < state.rows.length) select.value = current;
    renderFeedbackForm();
  }

  function currentFeedbackRow() {
    const index = Number($("feedback-activity").value);
    return Number.isInteger(index) ? state.rows[index] : null;
  }

  function renderFeedbackForm() {
    const row = currentFeedbackRow();
    if (!row) return;
    $("feedback-details").innerHTML = [["Task ID", row["Task ID"]], ["Sitio", row["Site ID"]], ["FME", row["FME ID"]], ["Coordinador", row.Coordinador]].map(([label, value]) => `<div class="detail-card"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`).join("");
    $("feedback-date").value = row["Fecha de feedback"] || dateInputLocal(new Date());
    $("feedback-state").value = ["Pendiente", "Aprobado", "Rechazado", "Otro"].includes(row["Estado OWS"]) ? row["Estado OWS"] : "Pendiente";
    $("feedback-text").value = row["Feedback OWS"] || "";
    $("delete-feedback").hidden = !hasFeedback(row);
  }

  function dateInputLocal(date) {
    const local = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
    return local.toISOString().slice(0, 10);
  }

  function renderClosedOptions() {
    const rows = state.rows.map((row, index) => ({ row, index })).filter(({ row }) => row["Es Closed"]);
    const select = $("closed-activity");
    const previous = select.value;
    select.innerHTML = "";
    rows.forEach(({ row, index }) => select.add(new Option(`${row["Task ID"]} · ${row["Site ID"]} · ${row.Coordinador}`, String(index))));
    if (previous && rows.some(({ index }) => String(index) === previous)) select.value = previous;
    if (!rows.length) {
      $("closed-details").innerHTML = '<p class="notice visible success">No hay actividades Closed rechazadas en este periodo.</p>';
      ["close-date", "email-to", "email-cc", "email-subject", "email-body", "open-outlook", "download-email"].forEach((id) => { $(id).disabled = true; });
      return;
    }
    ["close-date", "email-to", "email-cc", "email-subject", "email-body", "open-outlook", "download-email"].forEach((id) => { $(id).disabled = false; });
    if (!$("close-date").value) $("close-date").value = dateInputLocal(new Date());
    $("email-to").value = TO_RECIPIENTS.join("; ");
    $("email-cc").value = CC_RECIPIENTS.join("; ");
    renderClosedEmail();
  }

  function emailForRow(row) {
    const closeDate = $("close-date").value;
    const closeTime = $("close-time").value || "00:00:00";
    const [year, month, day] = (closeDate || "").split("-").map(Number);
    const [hour, minute, second] = closeTime.split(":").map(Number);
    const closeDateObj = year ? new Date(year, month - 1, day, hour || 0, minute || 0, second || 0) : null;
    const closeText = closeDateObj ? `${closeDateObj.getFullYear()}-${String(closeDateObj.getMonth() + 1).padStart(2, "0")}-${String(closeDateObj.getDate()).padStart(2, "0")} ${String(closeDateObj.getHours()).padStart(2, "0")}:${String(closeDateObj.getMinutes()).padStart(2, "0")}:${String(closeDateObj.getSeconds()).padStart(2, "0")}` : "";
    const audit = row["Audit time"];
    const auditText = audit ? `${displayDate(audit)} ${((audit.getHours() + 11) % 12) + 1}:${String(audit.getMinutes()).padStart(2, "0")}:${String(audit.getSeconds()).padStart(2, "0")} ${audit.getHours() < 12 ? "a. m." : "p. m."}` : "";
    const task = row["Task ID"] || row["Assignment ID"];
    const management = row["Gestion Closed"];
    let subject;
    if (management === "Enviar evidencias por correo") subject = `ENVÍO DE EVIDENCIAS PARA ${task} EN ESTADO CLOSED`;
    else if (management.startsWith("Correo mixto")) subject = `EVIDENCIAS Y SOLICITUD DE EXCLUSIÓN PARA ${task} EN ESTADO CLOSED`;
    else subject = `EXCLUSIÓN PARA ${task} EN ESTADO CLOSED`;
    const body = `Buenos días, cordial saludo.\n\nDe acuerdo a lo mencionado con anterioridad con respecto al cierre de actividades, esta fue cerrada automáticamente por el sistema el día ${closeText}. Posteriormente, fue auditada por parte de GNOC, ${auditText}. Quedando en un estado rechazada.\n\n${task}\nclosed\n${row["Site ID"]}\n${displayDate(audit)} Reject GNOC\n${row["Audit remark"]}\n\n${row.Proposito}`;
    return { subject, body };
  }

  function renderClosedEmail() {
    const index = Number($("closed-activity").value);
    const row = state.rows[index];
    if (!row || !row["Es Closed"]) return;
    $("closed-details").innerHTML = `<div class="detail-grid">${[["Task ID", row["Task ID"]], ["Sitio", row["Site ID"]], ["Coordinador", row.Coordinador], ["Audit time", displayDate(row["Audit time"], true)]].map(([label, value]) => `<div class="detail-card"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`).join("")}</div><label>Audit Remark<textarea rows="5" readonly>${escapeHtml(row["Audit remark"])}</textarea></label><label>Propósito generado según Audit Remark<textarea rows="5" readonly>${escapeHtml(row.Proposito)}</textarea></label>`;
    const email = emailForRow(row);
    $("email-subject").value = email.subject;
    $("email-body").value = email.body;
  }

  function renderMetrics() {
    const rows = state.rows;
    const count = (label, value) => `<div class="metric"><strong>${value}</strong><span>${label}</span></div>`;
    $("metrics").innerHTML = [
      count("Rechazos del periodo", rows.length),
      count("Nuevos", rows.filter((row) => row["Es nuevo"] === "Sí").length),
      count("Sin feedback OWS", rows.filter((row) => !hasFeedback(row)).length),
      count("Para notificar FME", rows.filter((row) => row["Es nuevo"] === "Sí" && !row["Es Closed"]).length),
      count("Closed", rows.filter((row) => row["Es Closed"]).length),
      count("FME con celular", rows.filter((row) => row["Celular FME"]).length),
    ].join("");
  }

  function renderAll() {
    renderMetrics();
    const unique = (field) => [...new Set(state.rows.map((row) => row[field]).filter(Boolean))].sort((a, b) => a.localeCompare(b, "es"));
    setOptions($("filter-coordinator"), unique("Coordinador"), "Todos", true);
    [["filter-fme", "FME ID"], ["filter-site", "Site ID"]].forEach(([id, field]) => setOptions($(id), unique(field), "Todos", true));
    [["filter-task-status", "Task status"], ["filter-ows-status", "Estado OWS"]].forEach(([id, field]) => setOptions($(id), unique(field), "Todos", true));
    renderDashboard();
    renderRejections();
    renderMessages();
    renderFeedbackOptions();
    renderClosedOptions();
    $("results").hidden = false;
  }

  function downloadBlob(filename, content, type) {
    const url = URL.createObjectURL(new Blob([content], { type }));
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function safeSheetName(name) {
    return name.replace(/[\\/*?:[\]]/g, "-").slice(0, 31);
  }

  function writeOutputWorkbook() {
    const workbook = readWorkbook(state.baseBytes);
    const monthLabel = MONTHS[state.month - 1].toLocaleUpperCase("es");
    const detailName = safeSheetName(`RECHAZOS ${monthLabel}`);
    const messagesName = safeSheetName(`MENSAJES FME ${monthLabel}`);
    const closedName = safeSheetName(`CLOSED ${monthLabel}`);
    [detailName, messagesName, closedName].forEach((name) => { if (workbook.SheetNames.includes(name)) { delete workbook.Sheets[name]; workbook.SheetNames = workbook.SheetNames.filter((sheet) => sheet !== name); } });
    const rejectionHeaders = ["Coordinador", "Origen coordinador", "Assignment ID", "Create time", "Assign time", "Assign operator", "Audit time", "Task ID", "Task status", "Site ID", "Region", "FM Office", "FME ID", "FME Supplier", "Audit remark", "Audit status", "Feedback Remark", "Feedback Qty", "Audit type", "Customer ticket", "Task title", "Assign User", "Estado contacto", "Tipo rechazo", "Departamento", "Municipio", "Nombre sitio", "Celular FME", "Fecha de feedback", "Estado OWS", "Feedback OWS", "Tiene feedback OWS"];
    const newRows = state.rows.filter((row) => row["Es nuevo"] === "Sí");
    const detailData = [rejectionHeaders, ...newRows.map((row) => rejectionHeaders.map((header) => row[header] ?? ""))];
    const detailSheet = XLSX.utils.aoa_to_sheet(detailData);
    detailSheet["!cols"] = rejectionHeaders.map((header) => ({ wch: header === "Audit remark" ? 55 : 20 }));
    detailSheet["!autofilter"] = { ref: detailSheet["!ref"] };
    workbook.SheetNames.push(detailName);
    workbook.Sheets[detailName] = detailSheet;
    const messageHeaders = ["Coordinador", "FME", "Celular FME", "Cantidad rechazos", "Sitios", "Mensaje", "Estado contacto"];
    const messageData = [messageHeaders, ...state.messages.map((item) => messageHeaders.map((header) => item[header] ?? ""))];
    const messageSheet = XLSX.utils.aoa_to_sheet(messageData);
    messageSheet["!cols"] = messageHeaders.map((header) => ({ wch: header === "Mensaje" || header === "Sitios" ? 55 : 20 }));
    workbook.SheetNames.push(messagesName);
    workbook.Sheets[messagesName] = messageSheet;
    const closedHeaders = ["Coordinador", "Origen coordinador", "Task ID", "Assignment ID", "Site ID", "FME ID", "Audit time", "Audit remark", "Tipo rechazo", "Gestion Closed", "Fecha cierre OWS", "Estado gestión"];
    const closedRows = state.rows.filter((row) => row["Es Closed"]).map((row) => [...closedHeaders.slice(0, 10).map((header) => row[header] ?? ""), "", "Falta fecha OWS"]);
    const closedSheet = XLSX.utils.aoa_to_sheet([closedHeaders, ...closedRows]);
    closedSheet["!cols"] = closedHeaders.map((header) => ({ wch: header === "Audit remark" ? 50 : 22 }));
    workbook.SheetNames.push(closedName);
    workbook.Sheets[closedName] = closedSheet;
    const output = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    downloadBlob(`RECHAZOS_Maria_actualizado_${MONTHS[state.month - 1].toLocaleLowerCase("es")}_${state.year}.xlsx`, output, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  }

  function downloadMessagesCsv() {
    const fields = ["Coordinador", "FME", "Celular FME", "Cantidad rechazos", "Sitios", "Mensaje", "Estado contacto"];
    const quoteCsv = (value) => `"${norm(value).replace(/"/g, '""')}"`;
    const csv = [fields, ...state.messages.map((item) => fields.map((field) => item[field] ?? ""))].map((row) => row.map(quoteCsv).join(",")).join("\r\n");
    downloadBlob(`mensajes_fme_${MONTHS[state.month - 1].toLocaleLowerCase("es")}_${state.year}.csv`, `\ufeff${csv}`, "text/csv;charset=utf-8");
  }

  function emlValue(value) {
    return norm(value).replace(/\r?\n/g, "\r\n").replace(/^[.]/gm, "..");
  }

  function emailAddresses(value) {
    return norm(value).split(";").map((entry) => {
      const match = entry.match(/<([^>]+)>/);
      return (match ? match[1] : entry).trim();
    }).filter(Boolean).join(";");
  }

  function openOutlookComposer() {
    const parameters = new URLSearchParams({
      to: emailAddresses($("email-to").value),
      cc: emailAddresses($("email-cc").value),
      subject: $("email-subject").value,
      body: $("email-body").value,
    });
    window.open(`https://outlook.office.com/mail/deeplink/compose?${parameters.toString()}`, "_blank", "noopener");
  }

  function downloadEml() {
    const index = Number($("closed-activity").value);
    const row = state.rows[index];
    if (!row) return;
    const to = $("email-to").value.split(";").map((value) => value.trim()).filter(Boolean);
    const cc = $("email-cc").value.split(";").map((value) => value.trim()).filter(Boolean);
    const headers = [`To: ${emlValue(to.join(", "))}`];
    if (cc.length) headers.push(`Cc: ${emlValue(cc.join(", "))}`);
    const subject = btoa(unescape(encodeURIComponent(emlValue($("email-subject").value))));
    headers.push(`Subject: =?UTF-8?B?${subject}?=`, "MIME-Version: 1.0", 'Content-Type: text/plain; charset="utf-8"', "Content-Transfer-Encoding: 8bit", "", emlValue($("email-body").value), "");
    downloadBlob(`${row["Task ID"] || row["Assignment ID"]}_closed.eml`, headers.join("\r\n"), "message/rfc822;charset=utf-8");
  }

  async function processFiles() {
    clearNotice();
    const ids = ["base-file", "report-file", "coordinator-file", "squads-file"];
    const files = ids.map((id) => $(id).files[0]);
    if (files.some((file) => !file)) {
      setNotice("Selecciona los cuatro archivos Excel antes de procesar.", "error");
      return;
    }
    if (!window.XLSX) {
      setNotice("No se pudo cargar la biblioteca Excel. Verifica la conexión a internet y vuelve a intentar.", "error");
      return;
    }
    const button = $("process-button");
    button.disabled = true;
    button.textContent = "Procesando…";
    try {
      const buffers = await Promise.all(files.map((file) => file.arrayBuffer()));
      const baseWorkbook = readWorkbook(buffers[0]);
      const reportWorkbook = readWorkbook(buffers[1]);
      const coordinatorWorkbook = readWorkbook(buffers[2]);
      const squadWorkbook = readWorkbook(buffers[3]);
      const selectedMonth = Number($("month").value);
      const selectedYear = Number($("year").value);
      if (!Number.isInteger(selectedYear) || selectedYear < 2000 || selectedYear > 2100) throw new Error("Ingresa un año válido entre 2000 y 2100.");
      const feedbacks = readFeedbacks();
      const rows = processRows(reportWorkbook, selectedYear, selectedMonth, coordinatorDirectory(baseWorkbook), loadSiteDirectory(coordinatorWorkbook), loadFmeDirectory(squadWorkbook), feedbacks);
      state.rows = rows;
      state.messages = groupMessages(rows);
      state.baseBytes = buffers[0];
      state.month = selectedMonth;
      state.year = selectedYear;
      renderAll();
      setNotice(`Procesamiento terminado: ${rows.length} rechazo(s) para ${MONTHS[selectedMonth - 1]} ${selectedYear}.`, "success");
    } catch (error) {
      setNotice(`No se pudo procesar la información: ${error.message}`, "error");
    } finally {
      button.disabled = false;
      button.textContent = "Procesar archivos";
    }
  }

  function saveCurrentFeedback(event) {
    event.preventDefault();
    const row = currentFeedbackRow();
    const feedbackText = $("feedback-text").value.trim();
    if (!row || !feedbackText) {
      setNotice("Selecciona una actividad y escribe el comentario de feedback de OWS.", "error");
      return;
    }
    try {
      const feedbacks = readFeedbacks();
      const key = row["Assignment ID"] || row["Task ID"];
      feedbacks[key] = { fecha: $("feedback-date").value, estado: $("feedback-state").value, feedback: feedbackText };
      writeFeedbacks(feedbacks);
      row["Fecha de feedback"] = feedbacks[key].fecha;
      row["Estado OWS"] = feedbacks[key].estado;
      row["Feedback OWS"] = feedbacks[key].feedback;
      row["Tiene feedback OWS"] = feedbacks[key].fecha && feedbacks[key].estado && feedbacks[key].feedback ? "Sí" : "No";
      renderAll();
      setNotice("Feedback guardado en este navegador.", "success");
    } catch (error) {
      setNotice(error.message, "error");
    }
  }

  function deleteCurrentFeedback() {
    const row = currentFeedbackRow();
    if (!row) return;
    try {
      const feedbacks = readFeedbacks();
      delete feedbacks[row["Assignment ID"]];
      delete feedbacks[row["Task ID"]];
      writeFeedbacks(feedbacks);
      row["Fecha de feedback"] = "";
      row["Estado OWS"] = "";
      row["Feedback OWS"] = "";
      row["Tiene feedback OWS"] = "No";
      renderAll();
      setNotice("Se eliminó el registro de feedback local.", "success");
    } catch (error) {
      setNotice(error.message, "error");
    }
  }

  function exportFeedback() {
    try {
      downloadBlob("respaldo_feedback_ows.json", JSON.stringify(readFeedbacks(), null, 2), "application/json;charset=utf-8");
    } catch (error) {
      setNotice(error.message, "error");
    }
  }

  async function importFeedback(file) {
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      if (!data || typeof data !== "object" || Array.isArray(data) || Object.values(data).some((item) => !item || typeof item !== "object" || Array.isArray(item))) {
        throw new Error("El archivo no tiene el formato esperado de respaldo de feedback.");
      }
      writeFeedbacks({ ...readFeedbacks(), ...data });
      if (state.rows.length && state.baseBytes) await processFiles();
      setNotice("Respaldo importado en este navegador.", "success");
    } catch (error) {
      setNotice(`No se pudo importar el respaldo: ${error.message}`, "error");
    } finally {
      $("restore-feedback").value = "";
    }
  }

  function initialize() {
    MONTHS.forEach((month, index) => $("month").add(new Option(month, String(index + 1))));
    const now = new Date();
    $("month").value = String(now.getMonth() + 1);
    $("year").value = String(now.getFullYear());
    $("process-button").addEventListener("click", processFiles);
    document.querySelectorAll(".tab").forEach((button) => button.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((tab) => tab.classList.toggle("active", tab === button));
      document.querySelectorAll(".tab-panel").forEach((panel) => { panel.hidden = panel.id !== `tab-${button.dataset.tab}`; });
    }));
    ["filter-coordinator", "filter-fme", "filter-site", "filter-task-status", "filter-ows-status", "filter-scope"].forEach((id) => $(id).addEventListener("change", renderDashboard));
    $("feedback-activity").addEventListener("change", renderFeedbackForm);
    $("feedback-form").addEventListener("submit", saveCurrentFeedback);
    $("delete-feedback").addEventListener("click", deleteCurrentFeedback);
    $("closed-activity").addEventListener("change", renderClosedEmail);
    ["close-date", "close-time"].forEach((id) => $(id).addEventListener("change", renderClosedEmail));
    $("open-outlook").addEventListener("click", openOutlookComposer);
    $("download-email").addEventListener("click", downloadEml);
    $("download-xlsx").addEventListener("click", () => { try { writeOutputWorkbook(); } catch (error) { setNotice(`No se pudo exportar el Excel: ${error.message}`, "error"); } });
    $("download-csv").addEventListener("click", downloadMessagesCsv);
    $("backup-feedback").addEventListener("click", exportFeedback);
    $("restore-feedback").addEventListener("change", (event) => importFeedback(event.target.files[0]));
  }

  document.addEventListener("DOMContentLoaded", initialize);
})();
