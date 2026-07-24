
import * as XLSX from "xlsx";

const app = window.__KONNECT__;
const STORAGE_KEY = "konnect_dashboard_v19_data";

const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];

const openUpdateFlow = $("#openUpdateFlow");
const updateChoiceOverlay = $("#updateChoiceOverlay");
const uploadOverlay = $("#uploadOverlay");
const cancelUpdateChoice = $("#cancelUpdateChoice");
const cancelUpload = $("#cancelUpload");
const selectFileBtn = $("#selectFileBtn");
const fileInput = $("#fileInput");
const dropZone = $("#dropZone");
const uploadTitle = $("#uploadTitle");
const uploadDescription = $("#uploadDescription");
const updateStatus = $("#updateStatus");
const updateValidation = $("#updateValidation");
const applyUpdate = $("#applyUpdate");

let updateType = null;
let pendingPayload = null;

const MONTHS = [
  "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
];
const MONTH_LABELS = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9 ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value == null || value === "") return 0;
  let s = String(value).trim().replace(/\$/g, "").replace(/\s/g, "");
  if (!s) return 0;
  if (s.includes(",") && s.includes(".")) {
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) {
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      s = s.replace(/,/g, "");
    }
  } else if (s.includes(",")) {
    const parts = s.split(",");
    if (parts.at(-1).length <= 2) {
      s = parts.slice(0, -1).join("").replace(/\./g, "") + "." + parts.at(-1);
    } else {
      s = parts.join("");
    }
  }
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function formatMoney(value) {
  return new Intl.NumberFormat("es-MX", {
    style: "currency",
    currency: "MXN",
    maximumFractionDigits: 0
  }).format(Number(value || 0));
}

function formatNumber(value) {
  return new Intl.NumberFormat("es-MX", { maximumFractionDigits: 0 }).format(Number(value || 0));
}

function formatPercent(value) {
  return `${Number(value || 0).toFixed(1)}%`;
}

function cleanText(value, limit = 85) {
  const s = String(value ?? "").replace(/\s+/g, " ").trim();
  return s.length > limit ? `${s.slice(0, limit - 1)}…` : s;
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return value;
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  const s = String(value ?? "").trim();
  if (!s) return null;
  const match = s.match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})/);
  if (match) {
    let year = Number(match[3]);
    if (year < 100) year += 2000;
    const d = new Date(year, Number(match[2]) - 1, Number(match[1]));
    return Number.isNaN(d.valueOf()) ? null : d;
  }
  const d = new Date(s);
  return Number.isNaN(d.valueOf()) ? null : d;
}

function formatDate(value) {
  const d = parseDate(value);
  if (!d) return String(value ?? "");
  return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getFullYear()).slice(-2)}`;
}

function sheetRows(workbook, name) {
  const exact = workbook.SheetNames.find(n => normalizeText(n) === normalizeText(name));
  if (!exact) return null;
  return XLSX.utils.sheet_to_json(workbook.Sheets[exact], { header: 1, raw: true, defval: null });
}

function mapHeaders(row) {
  const map = {};
  row.forEach((value, index) => {
    const key = normalizeText(value);
    if (key) map[key] = index;
  });
  return map;
}

function findHeaderIndex(map, candidates) {
  for (const candidate of candidates) {
    const key = Object.keys(map).find(k => k === candidate || k.includes(candidate));
    if (key != null) return map[key];
  }
  return -1;
}

function sumBy(items, getter) {
  return items.reduce((total, item) => total + Number(getter(item) || 0), 0);
}

function countBy(items, getter) {
  const result = {};
  items.forEach(item => {
    const key = String(getter(item) || "Sin definir").trim() || "Sin definir";
    result[key] = (result[key] || 0) + 1;
  });
  return result;
}

function moneyBy(items, keyGetter, valueGetter) {
  const result = {};
  items.forEach(item => {
    const key = String(keyGetter(item) || "Sin definir").trim() || "Sin definir";
    result[key] = (result[key] || 0) + Number(valueGetter(item) || 0);
  });
  return result;
}

function entriesSorted(map, limit = null) {
  const arr = Object.entries(map).sort((a, b) => b[1] - a[1]);
  return limit ? arr.slice(0, limit) : arr;
}

function classifyBlocker(comment) {
  const t = normalizeText(comment);
  if (/DOCUMENT|INE|CONSTANCIA|CSF|CIEC|CARATULA|COMPROBANTE|EXPEDIENTE|FIRMA/.test(t)) return "Documentación pendiente";
  if (/BROKER|CONSULTORIA|ASESOR|SOCIO/.test(t)) return "Seguimiento con broker";
  if (/CLIENTE|DECISION|CONFIRMACION CLIENTE|PAGO CLIENTE|RESPUESTA CLIENTE/.test(t)) return "Decisión del cliente";
  if (/FINANCIERA|BANCO|SOFOM|CREDITO|COMITE|ANALISIS|RESPUESTA/.test(t)) return "Respuesta de financiera";
  return "Otro";
}

function normalizeStatus(value) {
  const t = normalizeText(value);
  if (t.includes("VIABILIDAD")) return "Viabilidad";
  if (t.includes("INTEGRACION")) return "Integración";
  if (t.includes("ANALISIS")) return "Análisis";
  if (t.includes("AUTORIZ")) return "Autorización";
  if (t.includes("FORMALIZ")) return "Formalización";
  if (t.includes("DISPERS")) return "Dispersión";
  return String(value || "").trim();
}

function bucketCommercialStatus(value) {
  const t = normalizeText(value);
  if (t === "PAGADO") return "Pagado";
  if (t.includes("RECHAZ") || t.includes("BAJA")) return "No viable";
  if (t.includes("DESPU") || t.includes("RECUPER") || t.includes("TOUR VIRTUAL") || t.includes("LLAMADA")) return "Reactivación";
  if (t.includes("EVOLUC")) return "Desarrollo";
  return "Cierre";
}

function closeMonthIndex(value) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return value.getMonth();
  if (typeof value === "number" && value >= 1 && value <= 12) return Math.trunc(value) - 1;
  const text = normalizeText(value);
  if (!text) return null;
  const monthIndex = MONTHS.findIndex(month => text.includes(month));
  if (monthIndex >= 0) return monthIndex;
  const numeric = text.match(/(?:^|\s)(0?[1-9]|1[0-2])(?:$|\s)/);
  if (numeric) return Number(numeric[1]) - 1;
  const parsed = parseDate(value);
  return parsed ? parsed.getMonth() : null;
}

function statusColorClass(value) {
  const t = normalizeText(value);
  if (t.includes("RECHAZ") || t.includes("BAJA") || t.includes("NO VIABLE")) return "status-danger";
  if (t === "PAGADO" || t.includes("CONTRATO FIRMADO")) return "status-success";
  if (t.includes("PAGO PEND") || t.includes("RECABANDO")) return "status-payment";
  if (t.includes("CONTRATO")) return "status-contract";
  if (t.includes("NDA")) return "status-nda";
  if (t.includes("COMISION")) return "status-commission";
  if (t.includes("INTEGRACION") || t.includes("ASESORES KONNECT")) return "status-integration";
  if (t.includes("EVOLUC") || t.includes("TOUR") || t.includes("LLAMADA")) return "status-development";
  return "status-neutral";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizeMembership(value) {
  const t = normalizeText(value);
  if (t.includes("EVOLUCIONA")) return "Evoluciona";
  if (t.includes("DIAMANTE")) return "Diamante";
  if (t.includes("PLAT")) return "Platino";
  if (t.includes("ORO")) return "Oro";
  return "Sin definir";
}

function parseProjectionSheet(rows) {
  const result = {
    projection: [],
    dispersions: [],
    target: 65000000,
    periodMonth: null,
    periodYear: null
  };
  if (!rows) return result;

  for (const row of rows.slice(0, 8)) {
    row.forEach(value => {
      const s = String(value ?? "");
      const targetMatch = s.match(/OBJETIVO\s*\$?\s*([\d,.\s]+)/i);
      if (targetMatch) result.target = toNumber(targetMatch[1]);
      const monthIndex = MONTHS.findIndex(month => normalizeText(s).includes(month));
      const yearMatch = s.match(/20\d{2}/);
      if (monthIndex >= 0) result.periodMonth = monthIndex;
      if (yearMatch) result.periodYear = Number(yearMatch[0]);
    });
  }

  for (let i = 3; i < rows.length; i++) {
    const row = rows[i] || [];
    const leftClient = row[4];
    const leftAmount = toNumber(row[1]);
    const leftLabel = normalizeText(row[0]);
    if (leftClient && !leftLabel.includes("TOTAL")) {
      result.projection.push({
        director: String(row[0] ?? "").trim(),
        amount: leftAmount,
        financial: String(row[2] ?? "").trim(),
        broker: String(row[3] ?? "").trim(),
        client: String(leftClient ?? "").trim()
      });
    }

    const rightClient = row[10];
    const rightAmount = toNumber(row[7]);
    const rightLabel = normalizeText(row[6]);
    if (rightClient && !rightLabel.includes("TOTAL")) {
      result.dispersions.push({
        director: String(row[6] ?? "").trim(),
        amount: rightAmount,
        financial: String(row[8] ?? "").trim(),
        broker: String(row[9] ?? "").trim(),
        client: String(rightClient ?? "").trim()
      });
    }
  }
  return result;
}

function parseHistoricalClosings(rows) {
  if (!rows) return [];
  const sections = [];
  rows.forEach((row, index) => {
    const first = normalizeText(row?.[0]);
    const monthIndex = MONTHS.findIndex(month => first.includes(month));
    if (monthIndex >= 0 && first.includes("2026")) sections.push({ index, monthIndex });
  });

  const result = [];
  sections.forEach((section, idx) => {
    const end = sections[idx + 1]?.index ?? rows.length;
    const candidates = [];
    for (let r = section.index + 1; r < end; r++) {
      const label = normalizeText(rows[r]?.[0]);
      const amount = toNumber(rows[r]?.[1]);
      if (!amount) continue;
      if (
        label.includes("MONTO DISPERSADO EN TOTAL") ||
        label.includes("MONTO DISPERSADO TOTAL") ||
        label.includes("MONTO DISPRSADO TOTAL") ||
        label.includes("TOTAL DISPERSADO")
      ) {
        candidates.push({ label, amount });
      }
    }
    if (candidates.length) {
      result.push({
        monthIndex: section.monthIndex,
        label: MONTH_LABELS[section.monthIndex],
        amount: candidates.at(-1).amount
      });
    }
  });
  return result.sort((a, b) => a.monthIndex - b.monthIndex);
}

function parseOperationalWorkbook(workbook) {
  const pipelineRows = sheetRows(workbook, "PIPELINE");
  const projectionRows = sheetRows(workbook, "PROYECCIÓN") || sheetRows(workbook, "PROYECCION");
  const closureRows = sheetRows(workbook, "CIERRES 2026");
  if (!pipelineRows || !projectionRows || !closureRows) {
    throw new Error("El archivo debe contener las hojas PIPELINE, PROYECCIÓN y CIERRES 2026.");
  }

  const headerRow = pipelineRows.findIndex(row => normalizeText(row?.[0]).includes("FECHA") && normalizeText(row?.[1]).includes("ESTATUS"));
  if (headerRow < 0) throw new Error("No pude localizar los encabezados del Pipeline.");

  const headers = mapHeaders(pipelineRows[headerRow]);
  const idx = {
    date: findHeaderIndex(headers, ["FECHA"]),
    status: findHeaderIndex(headers, ["ESTATUS OPERACION", "ESTATUS"]),
    folio: findHeaderIndex(headers, ["FOLIO"]),
    client: findHeaderIndex(headers, ["CLIENTE"]),
    broker: findHeaderIndex(headers, ["CONSULTORIA"]),
    financial: findHeaderIndex(headers, ["FINANCIERA"]),
    product: findHeaderIndex(headers, ["PRODUCTO"]),
    requested: findHeaderIndex(headers, ["MONTO SOLICITADO"]),
    comment: findHeaderIndex(headers, ["COMENTARIOS"]),
    commentDate: findHeaderIndex(headers, ["FECHA DE COMENTARIO"]),
    granted: findHeaderIndex(headers, ["MONTO OTORGADO"])
  };

  const pipeline = pipelineRows.slice(headerRow + 1).map(row => ({
    date: parseDate(row[idx.date]),
    dateDisplay: formatDate(row[idx.date]),
    status: normalizeStatus(row[idx.status]),
    folio: row[idx.folio] == null ? "" : String(row[idx.folio]).replace(/\.0$/, ""),
    client: String(row[idx.client] ?? "").trim(),
    broker: String(row[idx.broker] ?? "").trim(),
    financial: String(row[idx.financial] ?? "").trim(),
    product: String(row[idx.product] ?? "").trim(),
    requested: toNumber(row[idx.requested]),
    granted: toNumber(row[idx.granted]),
    comment: cleanText(row[idx.comment], 90),
    commentDate: formatDate(row[idx.commentDate])
  })).filter(row => row.client || row.folio || row.status);

  const dated = pipeline.filter(row => row.date);
  const latestDate = dated.sort((a, b) => b.date - a.date)[0]?.date || new Date();
  const currentRows = pipeline.filter(row =>
    row.date &&
    row.date.getFullYear() === latestDate.getFullYear() &&
    row.date.getMonth() === latestDate.getMonth()
  );

  const projection = parseProjectionSheet(projectionRows);
  const historical = parseHistoricalClosings(closureRows);
  const target = projection.target || 65000000;
  const dispersed = sumBy(projection.dispersions, x => x.amount);
  const missing = Math.max(0, target - dispersed);
  const progress = target ? dispersed / target * 100 : 0;

  const stageNames = ["Viabilidad", "Integración", "Análisis", "Autorización", "Formalización"];
  const stages = {};
  stageNames.forEach(stage => {
    const rows = currentRows.filter(row => row.status === stage);
    stages[stage] = {
      rows,
      count: rows.length,
      requested: sumBy(rows, x => x.requested),
      granted: sumBy(rows, x => x.granted)
    };
  });
  stages["Dispersión"] = {
    rows: projection.dispersions,
    count: projection.dispersions.length,
    requested: dispersed,
    granted: dispersed
  };

  const integrationBlockers = countBy(stages["Integración"].rows, x => classifyBlocker(x.comment));
  const projectionByFinancial = moneyBy(projection.projection, x => x.financial, x => x.amount);
  const dispersionByFinancial = moneyBy(projection.dispersions, x => x.financial, x => x.amount);
  const dispersionCountByFinancial = countBy(projection.dispersions, x => normalizeText(x.financial) || 'Sin financiera');

  const previousHistory = historical.filter(x => x.monthIndex < (projection.periodMonth ?? latestDate.getMonth())).at(-1) || historical.at(-1);
  const previousAmount = previousHistory?.amount || 0;
  const previousProgress = target ? previousAmount / target * 100 : 0;

  const tableRow = x => [
    x.dateDisplay || "",
    x.folio || "",
    x.client || "",
    x.broker || "",
    x.financial || "",
    x.product || "",
    formatMoney(x.requested || 0),
    x.comment || ""
  ];
  const integrationTableRow = x => [
    x.dateDisplay || "",
    x.folio || "",
    x.client || "",
    x.financial || "",
    x.product || "",
    formatMoney(x.requested || 0),
    classifyBlocker(x.comment)
  ];
  const finalTableRow = x => [
    x.dateDisplay || "",
    x.folio || "",
    x.client || "",
    x.financial || "",
    x.product || "",
    formatMoney(x.requested || 0),
    formatMoney(x.granted || 0)
  ];

  const views = {
    viabilidad: {
      title: "Operaciones en Viabilidad",
      columns: ["Fecha", "Folio", "Cliente", "Consultoría", "Financiera", "Producto", "Monto solicitado", "Comentario"],
      rows: stages["Viabilidad"].rows.map(tableRow),
      summary: [
        { label: "Operaciones", value: formatNumber(stages["Viabilidad"].count) },
        { label: "Monto solicitado", value: formatMoney(stages["Viabilidad"].requested) }
      ]
    },
    integracion: {
      title: "Operaciones en Integración",
      columns: ["Fecha", "Folio", "Cliente", "Financiera", "Producto", "Monto solicitado", "Motivo"],
      rows: stages["Integración"].rows.map(integrationTableRow),
      summary: [
        { label: "Operaciones", value: formatNumber(stages["Integración"].count) },
        { label: "Monto solicitado", value: formatMoney(stages["Integración"].requested) }
      ]
    },
    analisis: {
      title: "Operaciones en Análisis",
      columns: ["Fecha", "Folio", "Cliente", "Consultoría", "Financiera", "Producto", "Monto solicitado", "Comentario"],
      rows: stages["Análisis"].rows.map(tableRow),
      summary: [
        { label: "Operaciones", value: formatNumber(stages["Análisis"].count) },
        { label: "Monto solicitado", value: formatMoney(stages["Análisis"].requested) }
      ]
    },
    autorizacion: {
      title: "Operaciones en Autorización",
      columns: ["Fecha", "Folio", "Cliente", "Financiera", "Producto", "Monto solicitado", "Monto otorgado"],
      rows: stages["Autorización"].rows.map(finalTableRow),
      summary: [
        { label: "Operaciones", value: formatNumber(stages["Autorización"].count) },
        { label: "Monto solicitado", value: formatMoney(stages["Autorización"].requested) },
        { label: "Monto otorgado", value: formatMoney(stages["Autorización"].granted) }
      ]
    },
    formalizacion: {
      title: "Operaciones en Formalización",
      columns: ["Fecha", "Folio", "Cliente", "Financiera", "Producto", "Monto solicitado", "Monto otorgado"],
      rows: stages["Formalización"].rows.map(finalTableRow),
      summary: [
        { label: "Operaciones", value: formatNumber(stages["Formalización"].count) },
        { label: "Monto solicitado", value: formatMoney(stages["Formalización"].requested) },
        { label: "Monto otorgado", value: formatMoney(stages["Formalización"].granted) }
      ]
    },
    proyeccion: {
      title: "Operaciones en Proyección",
      columns: ["Cliente", "Financiera", "Broker / Consultoría", "Monto"],
      rows: projection.projection.map(x => [x.client, x.financial, x.broker, formatMoney(x.amount)]),
      summary: [
        { label: "Operaciones", value: formatNumber(projection.projection.length) },
        { label: "Potencial", value: formatMoney(sumBy(projection.projection, x => x.amount)) }
      ]
    },
    dispersion: {
      title: "Operaciones en Dispersión",
      columns: ["Cliente", "Financiera", "Broker / Consultoría", "Monto dispersado"],
      rows: projection.dispersions.map(x => [x.client, x.financial, x.broker, formatMoney(x.amount)]),
      summary: [
        { label: "Operaciones", value: formatNumber(projection.dispersions.length) },
        { label: "Monto dispersado", value: formatMoney(dispersed) }
      ]
    }
  };

  return {
    type: "operational",
    importedAt: new Date().toISOString(),
    periodMonth: projection.periodMonth ?? latestDate.getMonth(),
    periodYear: projection.periodYear ?? latestDate.getFullYear(),
    target,
    dispersed,
    missing,
    progress,
    historical,
    previousHistory,
    previousAmount,
    previousProgress,
    stages,
    projection,
    integrationBlockers,
    projectionByFinancial,
    dispersionByFinancial,
    dispersionCountByFinancial,
    views
  };
}


const EXCLUDED_COMMERCIAL_NAMES = ["NANCY", "PEDRO", "ERIKA", "ERICKA"];

function hasExcludedCommercialName(value) {
  const text = normalizeText(value);
  return EXCLUDED_COMMERCIAL_NAMES.some(name => text.includes(name));
}

function isExcludedCommercialRecord(record) {
  const values = Array.isArray(record.sourceValues)
    ? record.sourceValues
    : Object.values(record);
  return values.some(value => hasExcludedCommercialName(value));
}

function hasNoReferral(value) {
  const text = normalizeText(value);
  if (!text) return true;
  return text === "NO" ||
    text === "NA" ||
    text === "N A" ||
    text === "NINGUNO" ||
    text === "NINGUNA" ||
    text.startsWith("NO ") ||
    text.startsWith("NO/") ||
    text.startsWith("NO-");
}

function cleanReferralLabel(value) {
  const cleaned = String(value ?? "")
    .trim()
    .replace(/^\s*SI\s*[\/\-:]?\s*/i, "")
    .trim();
  return cleaned || "Sin identificar";
}

function buildDirectorScope(prospects, targetName) {
  const own = [];
  const referred = [];
  const seen = new Set();

  const pushReferred = (row, source) => {
    const key = `${row.name}|${row.director}|${row.referred}|${targetName}`;
    if (seen.has(key)) return;
    seen.add(key);
    referred.push({ ...row, referralSource: source || "Sin identificar" });
  };

  prospects.forEach(row => {
    const directorText = normalizeText(row.director);
    const referredText = normalizeText(row.referred);
    const directorIsTarget = directorText.includes(targetName);
    const referralTargetsPerson = referredText.includes(targetName);

    if (directorIsTarget && hasNoReferral(row.referred)) {
      own.push(row);
      return;
    }

    if (directorIsTarget && !hasNoReferral(row.referred)) {
      let source = cleanReferralLabel(row.referred);
      if (normalizeText(source).includes(targetName)) source = "Sin identificar";
      pushReferred(row, source);
      return;
    }

    if (!directorIsTarget && referralTargetsPerson) {
      pushReferred(row, row.director || "Sin identificar");
    }
  });

  return {
    name: targetName.charAt(0) + targetName.slice(1).toLowerCase(),
    own,
    referred,
    ownCount: own.length,
    referredCount: referred.length,
    totalCount: own.length + referred.length,
    ownStatuses: countBy(own, row => bucketCommercialStatus(row.status)),
    referredStatuses: countBy(referred, row => bucketCommercialStatus(row.status)),
    referralSources: countBy(referred, row => row.referralSource || "Sin identificar")
  };
}

function parseCommercialWorkbook(workbook) {
  const rows = sheetRows(workbook, "PROSPECTOS MEMBRESIAS");
  if (!rows) throw new Error("El archivo debe contener la hoja PROSPECTOS MEMBRESIAS.");

  const headerRow = rows.findIndex(row => normalizeText(row?.[0]) === "NOMBRE" && normalizeText(row?.[7]).includes("ESTATUS"));
  if (headerRow < 0) throw new Error("No pude localizar los encabezados comerciales.");

  const headers = mapHeaders(rows[headerRow]);
  const idx = {
    name: findHeaderIndex(headers, ["NOMBRE"]),
    program: findHeaderIndex(headers, ["PROGRAMA"]),
    director: findHeaderIndex(headers, ["DIRECTOR COMERCIAL"]),
    location: findHeaderIndex(headers, ["LOCALIDAD"]),
    email: findHeaderIndex(headers, ["CORREO"]),
    phone: findHeaderIndex(headers, ["TELEFONO"]),
    membership: findHeaderIndex(headers, ["MEMBRESIA INTERESADO"]),
    status: findHeaderIndex(headers, ["ESTATUS"]),
    comment: findHeaderIndex(headers, ["COMENTARIOS"]),
    commentDate: findHeaderIndex(headers, ["FECHA DE ULTIMO COMENTARIO"]),
    referred: findHeaderIndex(headers, ["REFERENCIADO"]),
    closeMonth: findHeaderIndex(headers, ["MES PARA CIERRE"])
  };

  const allProspects = rows.slice(headerRow + 1).map(row => ({
    sourceValues: row.map(value => String(value ?? "")),
    name: String(row[idx.name] ?? "").trim(),
    program: String(row[idx.program] ?? "").trim(),
    director: String(row[idx.director] ?? "").trim() || "Sin asignar",
    location: String(row[idx.location] ?? "").trim() || "Sin localidad",
    email: String(row[idx.email] ?? "").trim(),
    phone: String(row[idx.phone] ?? "").trim(),
    membership: String(row[idx.membership] ?? "").trim(),
    status: String(row[idx.status] ?? "").trim(),
    comment: cleanText(row[idx.comment], 90),
    commentDate: formatDate(row[idx.commentDate]),
    referred: idx.referred >= 0 ? String(row[idx.referred] ?? "").trim() : "",
    closeMonthRaw: idx.closeMonth >= 0 ? row[idx.closeMonth] : null,
    closeMonthIndex: idx.closeMonth >= 0 ? closeMonthIndex(row[idx.closeMonth]) : null
  })).filter(row => row.name);

  const prospects = allProspects.filter(row => !isExcludedCommercialRecord(row));
  const excludedCount = allProspects.length - prospects.length;
  const buckets = countBy(prospects, row => bucketCommercialStatus(row.status));
  const open = prospects.filter(row => bucketCommercialStatus(row.status) !== "Pagado");
  const directorsOpen = countBy(open, row => row.director);
  const locationsOpen = countBy(open, row => row.location);
  const memberships = countBy(prospects, row => normalizeMembership(row.membership));

  const now = new Date();
  const currentCloseMonthIndex = now.getMonth();
  const nextCloseMonthIndex = (currentCloseMonthIndex + 1) % 12;
  const currentCloseYear = now.getFullYear();
  const nextCloseYear = currentCloseMonthIndex === 11 ? currentCloseYear + 1 : currentCloseYear;
  const currentClosings = prospects.filter(row => row.closeMonthIndex === currentCloseMonthIndex);
  const nextClosings = prospects.filter(row => row.closeMonthIndex === nextCloseMonthIndex);

  const directorScopes = {
    diego: buildDirectorScope(prospects, "DIEGO"),
    jorge: buildDirectorScope(prospects, "JORGE")
  };

  const rowForTable = row => [
    row.name,
    row.program || row.membership,
    row.director,
    row.location,
    row.status,
    row.phone,
    row.comment
  ];

  const makeView = (title, filter) => {
    const filtered = prospects.filter(filter);
    return {
      title,
      columns: ["Nombre", "Programa / Membresía", "Director", "Localidad", "Estatus", "Teléfono", "Comentario"],
      rows: filtered.map(rowForTable),
      summary: [{ label: "Registros", value: formatNumber(filtered.length) }]
    };
  };

  const ownScopeView = (scope, name) => ({
    title: `${name} · operaciones 100% propias`,
    columns: ["Nombre", "Programa / Membresía", "Director", "Localidad", "Estatus", "Teléfono", "Comentario"],
    rows: scope.own.map(rowForTable),
    summary: [{ label: "Operaciones", value: formatNumber(scope.ownCount) }]
  });

  const referredScopeView = (scope, name) => ({
    title: `${name} · operaciones referenciadas`,
    columns: ["Nombre", "Programa / Membresía", "Referenciada por", "Director registrado", "Localidad", "Estatus", "Comentario"],
    rows: scope.referred.map(row => [
      row.name,
      row.program || row.membership,
      row.referralSource || "Sin identificar",
      row.director,
      row.location,
      row.status,
      row.comment
    ]),
    summary: [{ label: "Operaciones", value: formatNumber(scope.referredCount) }]
  });

  const views = {
    abiertos: makeView("Seguimiento abierto", row => bucketCommercialStatus(row.status) !== "Pagado"),
    pagados: makeView("Membresías pagadas", row => bucketCommercialStatus(row.status) === "Pagado"),
    reactivacion: makeView("Prospectos en reactivación", row => bucketCommercialStatus(row.status) === "Reactivación"),
    desarrollo: makeView("Prospectos en desarrollo", row => bucketCommercialStatus(row.status) === "Desarrollo"),
    cierre: makeView("Prospectos en cierre", row => bucketCommercialStatus(row.status) === "Cierre"),
    no_viable: makeView("Prospectos no viables", row => bucketCommercialStatus(row.status) === "No viable"),
    prioritarios: {
      title: `Cierres prioritarios de ${MONTHS[currentCloseMonthIndex].toLowerCase()}`,
      columns: ["Nombre", "Programa / Membresía", "Director", "Localidad", "Estatus", "Teléfono", "Comentario"],
      rows: currentClosings.map(rowForTable),
      summary: [{ label: "Registros", value: formatNumber(currentClosings.length) }]
    },
    cierres_siguiente: {
      title: `Cierres previstos para ${MONTHS[nextCloseMonthIndex].toLowerCase()}`,
      columns: ["Nombre", "Programa / Membresía", "Director", "Localidad", "Estatus", "Teléfono", "Comentario"],
      rows: nextClosings.map(rowForTable),
      summary: [{ label: "Registros", value: formatNumber(nextClosings.length) }]
    },
    diego_propias: ownScopeView(directorScopes.diego, "Diego"),
    diego_referenciadas: referredScopeView(directorScopes.diego, "Diego"),
    jorge_propias: ownScopeView(directorScopes.jorge, "Jorge"),
    jorge_referenciadas: referredScopeView(directorScopes.jorge, "Jorge")
  };

  return {
    type: "commercial",
    importedAt: new Date().toISOString(),
    prospects,
    excludedCount,
    buckets,
    open,
    directorsOpen,
    locationsOpen,
    memberships,
    currentCloseMonthIndex,
    nextCloseMonthIndex,
    currentCloseYear,
    nextCloseYear,
    currentClosings,
    nextClosings,
    directorScopes,
    views
  };
}


function findMetricCard(sectionId, label) {
  const section = document.getElementById(sectionId);
  if (!section) return null;
  return $$(".metric-card", section).find(card =>
    normalizeText($(".eyebrow", card)?.textContent) === normalizeText(label)
  ) || null;
}

function setMetric(sectionId, label, value, subtle = null) {
  const card = findMetricCard(sectionId, label);
  if (!card) return;
  const metric = $(".metric", card);
  if (metric) metric.textContent = value;
  if (subtle != null) {
    const sub = $(".subtle", card);
    if (sub) sub.textContent = subtle;
  }
}

function buildDonut(items, centerMain, centerSub, size = 230, compact = false) {
  const colors = ["#19d6e2", "#3db7ff", "#6b74ff", "#8f5cff", "#12a7ff", "#3df0c7", "#725cff"];
  const total = items.reduce((s, item) => s + Number(item.value || 0), 0) || 1;
  let start = 0;
  const segments = [];
  const legends = [];
  items.forEach((item, index) => {
    const pct = Number(item.value || 0) / total * 100;
    const end = start + pct;
    const color = colors[index % colors.length];
    segments.push(`${color} ${start.toFixed(2)}% ${end.toFixed(2)}%`);
    legends.push(`
      <div class="legend-item">
        <span class="dot" style="background:${color}"></span>
        <span class="legend-name">${item.name}</span>
        <span class="legend-val">${formatNumber(item.value)}</span>
      </div>
    `);
    start = end;
  });
  const extra = compact ? ' style="grid-template-columns:168px 1fr;gap:18px;align-items:center;justify-content:center;"' : "";
  const centerStyle = compact ? ' style="font-size:28px;"' : "";
  const subStyle = compact ? ' style="font-size:12px;max-width:110px;"' : "";
  const legendStyle = compact ? ' style="gap:8px;"' : "";
  return `
    <div class="donut-wrap"${extra}>
      <div class="donut" style="--size:${size}px; --segments:${segments.join(", ")}">
        <div class="donut-center">
          <div class="center-main"${centerStyle}>${centerMain}</div>
          <div class="center-sub"${subStyle}>${centerSub}</div>
        </div>
      </div>
      <div class="legend"${legendStyle}>${legends.join("")}</div>
    </div>
  `;
}

function buildBars(entries, money = false, maxItems = null) {
  const visible = maxItems ? entries.slice(0, maxItems) : entries;
  const max = Math.max(...visible.map(([, value]) => Number(value || 0)), 1);
  return `<div class="bar-list compact">${
    visible.map(([name, value]) => `
      <div class="bar-row">
        <div class="bar-label">${name}</div>
        <div class="bar-track"><div class="bar-fill" style="--w:${(Number(value || 0) / max * 100).toFixed(1)}%"></div></div>
        <div class="bar-value">${money ? formatMoney(value) : formatNumber(value)}</div>
      </div>
    `).join("")
  }</div>`;
}

function replaceSectionContent(sectionId, title, html) {
  const section = document.getElementById(sectionId);
  if (!section) return;
  const titleNode = $$(".section-title", section).find(node => normalizeText(node.textContent) === normalizeText(title));
  if (!titleNode) return;
  const card = titleNode.closest(".card");
  if (!card) return;
  [...card.children].forEach(child => {
    if (child !== titleNode) child.remove();
  });
  card.insertAdjacentHTML("beforeend", html);
}

function updateHistory(data) {
  const chart = $("#op-01 .compact-history-chart");
  if (!chart) return;
  const max = Math.max(...data.historical.map(x => x.amount), 1);
  chart.innerHTML = `
    <div class="history-grid"></div>
    <div class="history-cols compact-history-cols">
      ${data.historical.map(item => `
        <div class="history-col compact-history-col">
          <div class="history-bar compact-history-bar">
            <div class="history-fill" style="--h:${(item.amount / max * 100).toFixed(2)}%"></div>
          </div>
          <div class="history-amount compact-history-amount">${formatMoney(item.amount)}</div>
          <div class="history-label">${item.label}</div>
        </div>
      `).join("")}
    </div>
  `;
}

function updateOperationalVisual(data) {
  setMetric("op-01", "Meta mensual", formatMoney(data.target));
  setMetric("op-01", "Dispersión actual", formatMoney(data.dispersed), `${data.projection.dispersions.length} operaciones confirmadas.`);
  setMetric("op-01", "Faltante", formatMoney(data.missing));
  setMetric("op-01", "Avance", formatPercent(data.progress));
  updateHistory(data);

  const delta = data.dispersed - data.previousAmount;
  const deltaNode = $("#op-01 .compare-delta");
  if (deltaNode) deltaNode.textContent = `${delta >= 0 ? "+" : "-"}${formatMoney(Math.abs(delta))}`;
  const baseNode = $("#op-01 .compare-base");
  if (baseNode) baseNode.textContent = `Actual ${formatMoney(data.dispersed)} · ${data.previousHistory?.label || "Anterior"} ${formatMoney(data.previousAmount)}`;
  const pctBoxes = $$("#op-01 .compare-pct-box");
  if (pctBoxes[0]) {
    const label = $(".eyebrow", pctBoxes[0]);
    const pct = $(".compare-pct", pctBoxes[0]);
    if (label) label.textContent = data.previousHistory?.label || "Anterior";
    if (pct) pct.textContent = formatPercent(data.previousProgress);
  }
  if (pctBoxes[1]) {
    const pct = $(".compare-pct", pctBoxes[1]);
    if (pct) pct.textContent = formatPercent(data.progress);
  }

  const stageList = ["Viabilidad", "Integración", "Análisis", "Autorización", "Formalización", "Dispersión"];
  const maxCount = Math.max(...stageList.map(name => data.stages[name]?.count || 0), 1);
  $$("#op-02 .status-card").forEach(card => {
    const name = $(".status-name", card)?.textContent.trim();
    const stage = data.stages[name];
    if (!stage) return;
    $(".status-count", card).textContent = formatNumber(stage.count);
    $(".status-money", card).textContent = formatMoney(stage.requested);
    const fill = $(".fill", card);
    if (fill) fill.style.setProperty("--w", `${(stage.count / maxCount * 100).toFixed(1)}%`);
  });
  replaceSectionContent(
    "op-02",
    "Participación del pipeline",
    buildDonut(stageList.map(name => ({ name, value: data.stages[name]?.count || 0 })), formatNumber(sumBy(stageList, name => data.stages[name]?.count || 0)), "Operaciones visibles", 280)
  );

  const blockerEntries = entriesSorted(data.integrationBlockers);
  const integration = data.stages["Integración"];
  setMetric("op-03", "Operaciones", formatNumber(integration.count));
  setMetric("op-03", "Monto solicitado", formatMoney(integration.requested));
  setMetric("op-03", "Causa principal", blockerEntries[0]?.[0] || "Sin clasificar", blockerEntries.length ? formatPercent(blockerEntries[0][1] / Math.max(integration.count, 1) * 100) : "0.0%");
  setMetric("op-03", "Top 3 causas", blockerEntries.length ? formatPercent(blockerEntries.slice(0, 3).reduce((s, x) => s + x[1], 0) / Math.max(integration.count, 1) * 100) : "0.0%");
  setMetric("op-03", "Respuesta financiera", formatNumber(data.integrationBlockers["Respuesta de financiera"] || 0));
  setMetric("op-03", "Documentación pendiente", formatNumber(data.integrationBlockers["Documentación pendiente"] || 0));
  replaceSectionContent(
    "op-03",
    "Participación por causa",
    buildDonut(blockerEntries.map(([name, value]) => ({ name, value })), formatNumber(integration.count), "Causas clasificadas", 250)
  );
  replaceSectionContent("op-03", "Peso por motivo", buildBars(blockerEntries, false));

  const projectionTotal = sumBy(data.projection.projection, x => x.amount);
  setMetric("op-04", "Autorización", formatNumber(data.stages["Autorización"].count), `${formatMoney(data.stages["Autorización"].requested)} solicitado`);
  setMetric("op-04", "Formalización", formatNumber(data.stages["Formalización"].count), `${formatMoney(data.stages["Formalización"].requested)} solicitado`);
  setMetric("op-04", "Proyección", formatNumber(data.projection.projection.length), `${formatMoney(projectionTotal)} potencial`);
  replaceSectionContent("op-04", "Concentración del potencial", buildBars(entriesSorted(data.projectionByFinancial), true));
  const closingAmounts = $$("#op-04 .closing-amount");
  if (closingAmounts[0]) closingAmounts[0].textContent = formatMoney(data.stages["Autorización"].granted);
  if (closingAmounts[1]) closingAmounts[1].textContent = formatMoney(data.stages["Formalización"].granted);
  const potential = $("#op-04 .potential-total-amount");
  if (potential) potential.textContent = formatMoney(projectionTotal);
  replaceSectionContent(
    "op-04",
    "Composición",
    buildDonut([
      { name: "Autorización", value: data.stages["Autorización"].count },
      { name: "Formalización", value: data.stages["Formalización"].count },
      { name: "Proyección", value: data.projection.projection.length }
    ], formatNumber(data.stages["Autorización"].count + data.stages["Formalización"].count + data.projection.projection.length), "Registros visibles", 190)
  );

  setMetric("op-05", "Monto dispersado", formatMoney(data.dispersed));
  setMetric("op-05", "Operaciones", formatNumber(data.projection.dispersions.length));
  setMetric("op-05", "Avance", formatPercent(data.progress));
  setMetric("op-05", "Faltante", formatMoney(data.missing));
  replaceSectionContent("op-05", "Operaciones por financiera", buildBars(entriesSorted(data.dispersionCountByFinancial), false));

  const sortedDisp = [...data.projection.dispersions].sort((a, b) => b.amount - a.amount);
  const principal = entriesSorted(data.dispersionCountByFinancial)[0] || ["Sin financiera", 0];
  setMetric("op-05", "Mayor operación", formatMoney(sortedDisp[0]?.amount || 0), cleanText(sortedDisp[0]?.client || "Sin operación", 50));
  setMetric("op-05", "Financiera principal", principal[0], `${formatNumber(principal[1])} operaciones`);
  setMetric("op-05", "Promedio", formatMoney(data.projection.dispersions.length ? data.dispersed / data.projection.dispersions.length : 0));

  Object.entries(data.views).forEach(([key, value]) => {
    app.viewTables[key] = value;
  });
}

function renderMonthClosingSlide(sectionId, rows, monthIndex, year) {
  const section = document.getElementById(sectionId);
  if (!section) return;
  const monthName = MONTHS[monthIndex] || "MES";
  const monthLabel = monthName.charAt(0) + monthName.slice(1).toLowerCase();
  const titleMonth = $(".month-slide-name", section);
  const chip = $(".month-period-chip", section);
  if (titleMonth) titleMonth.textContent = monthLabel;
  if (chip) chip.textContent = `${monthName} ${year}`;

  const directors = countBy(rows, x => x.director || "Sin asignar");
  const statuses = countBy(rows, x => x.status || "Por definir");
  const topDirector = entriesSorted(directors)[0] || ["—", 0];
  const topStatus = entriesSorted(statuses)[0] || ["—", 0];

  const countNode = $(".month-case-count", section);
  const directorNode = $(".month-top-director", section);
  const directorSub = $(".month-top-director-sub", section);
  const statusNode = $(".month-top-status", section);
  const statusSub = $(".month-top-status-sub", section);
  const tableCount = $(".month-table-count", section);
  if (countNode) countNode.textContent = formatNumber(rows.length);
  if (directorNode) directorNode.textContent = topDirector[0];
  if (directorSub) directorSub.textContent = `${formatNumber(topDirector[1])} casos`;
  if (statusNode) statusNode.textContent = topStatus[0];
  if (statusSub) statusSub.textContent = `${formatNumber(topStatus[1])} casos`;
  if (tableCount) tableCount.textContent = `${formatNumber(rows.length)} registros`;

  const tbody = $(".month-closing-table tbody", section);
  if (tbody) {
    tbody.innerHTML = rows.map(row => `
      <tr>
        <td><div class="priority-person">${escapeHtml(row.name)}</div><div class="priority-location">${escapeHtml(row.location || "—")}</div></td>
        <td>${escapeHtml(row.program || row.membership || "—")}</td>
        <td>${escapeHtml(row.director || "—")}</td>
        <td><span class="month-status-pill ${statusColorClass(row.status)}">${escapeHtml(row.status || "Por definir")}</span></td>
        <td><div class="priority-comment">${escapeHtml(row.comment || "Sin comentario")}</div></td>
      </tr>
    `).join("");
  }

  const wrap = $(".month-closing-table-wrap", section);
  if (wrap) {
    const existing = $(".priority-empty-state", wrap);
    if (existing) existing.remove();
    if (!rows.length) wrap.insertAdjacentHTML("beforeend", '<div class="priority-empty-state">No hay cierres marcados para este mes.</div>');
  }
}


function scopeStatusClass(name) {
  const normalized = normalizeText(name);
  if (normalized === "PAGADO") return "status-success";
  if (normalized === "NO VIABLE") return "status-danger";
  if (normalized === "DESARROLLO") return "status-development";
  if (normalized === "REACTIVACION") return "status-commission";
  if (normalized === "CIERRE") return "status-payment";
  return "status-neutral";
}

function renderScopeStatusPills(statuses) {
  const entries = entriesSorted(statuses || {});
  if (!entries.length) return '<div class="scope-empty">Sin operaciones en esta sección.</div>';
  return entries.map(([name, value]) => `
    <span class="scope-status-pill ${scopeStatusClass(name)}">
      <span>${escapeHtml(name)}</span>
      <strong>${formatNumber(value)}</strong>
    </span>
  `).join("");
}

function renderScopeSourcePills(sources) {
  const entries = entriesSorted(sources || {});
  if (!entries.length) return '<div class="scope-empty">Sin referencias registradas.</div>';
  return entries.map(([name, value]) => `
    <span class="scope-source-pill">
      <span>Referenciadas por ${escapeHtml(name)}</span>
      <strong>${formatNumber(value)}</strong>
    </span>
  `).join("");
}

function renderDirectorScopeSlide(data) {
  const section = document.getElementById("com-scope");
  if (!section) return;

  const excludedNode = $(".scope-excluded-count", section);
  if (excludedNode) excludedNode.textContent = formatNumber(data.excludedCount || 0);

  ["diego", "jorge"].forEach(key => {
    const scope = data.directorScopes?.[key] || {
      name: key,
      ownCount: 0,
      referredCount: 0,
      totalCount: 0,
      ownStatuses: {},
      referredStatuses: {},
      referralSources: {}
    };
    const container = $(`[data-scope-person="${key}"]`, section);
    if (!container) return;

    container.innerHTML = `
      <div class="scope-director-header">
        <div>
          <div class="scope-director-name">${escapeHtml(scope.name || key)}</div>
          <div class="scope-director-sub"><strong>${formatNumber(scope.totalCount || 0)}</strong> operaciones dentro de su alcance</div>
        </div>
        <div class="scope-total-chip">${formatNumber(scope.totalCount || 0)} total</div>
      </div>
      <div class="scope-block-grid">
        <div class="scope-block scope-own">
          <div class="scope-block-top">
            <div>
              <div class="scope-kicker">Cartera directa</div>
              <div class="scope-block-title">100% propias</div>
            </div>
            <div class="scope-count">${formatNumber(scope.ownCount || 0)}</div>
          </div>
          <div class="scope-label">Distribución por estatus</div>
          <div class="scope-status-list">${renderScopeStatusPills(scope.ownStatuses)}</div>
        </div>
        <div class="scope-block scope-referred">
          <div class="scope-block-top">
            <div>
              <div class="scope-kicker">Cartera compartida</div>
              <div class="scope-block-title">Referenciadas</div>
            </div>
            <div class="scope-count">${formatNumber(scope.referredCount || 0)}</div>
          </div>
          <div class="scope-label">Origen de las referencias</div>
          <div class="scope-source-list">${renderScopeSourcePills(scope.referralSources)}</div>
          <div class="scope-label scope-label-status">Estatus</div>
          <div class="scope-status-list">${renderScopeStatusPills(scope.referredStatuses)}</div>
        </div>
      </div>
    `;
  });
}


function updateCommercialVisual(data) {
  const prospectTotal = data.prospectCount ?? data.prospects?.length ?? 0;
  const paid = data.buckets?.["Pagado"] || 0;
  const open = data.openCount ?? data.open?.length ?? 0;

  setMetric("com-01", "Prospectos totales", formatNumber(prospectTotal));
  setMetric("com-01", "Pagados", formatNumber(paid));
  setMetric("com-01", "Seguimiento abierto", formatNumber(open));
  setMetric("com-01", "Cierre prioritario", formatNumber(data.currentClosings?.length || 0));

  replaceSectionContent(
    "com-01",
    "Estado comercial",
    buildDonut(
      ["Pagado", "Cierre", "Desarrollo", "Reactivación", "No viable"].map(name => ({ name, value: data.buckets?.[name] || 0 })),
      formatNumber(prospectTotal),
      "Prospectos considerados",
      260
    )
  );
  replaceSectionContent("com-01", "Directores con seguimiento abierto", buildBars(entriesSorted(data.directorsOpen || {}), false));

  renderDirectorScopeSlide(data);
  renderMonthClosingSlide("com-02", data.currentClosings || [], data.currentCloseMonthIndex ?? new Date().getMonth(), data.currentCloseYear ?? new Date().getFullYear());
  renderMonthClosingSlide("com-03", data.nextClosings || [], data.nextCloseMonthIndex ?? ((new Date().getMonth() + 1) % 12), data.nextCloseYear ?? new Date().getFullYear());

  const topLocation = entriesSorted(data.locationsOpen || {})[0] || ["Sin localidad", 0];
  setMetric("com-04", "Reactivación", formatNumber(data.buckets?.["Reactivación"] || 0));
  setMetric("com-04", "En cierre", formatNumber(data.buckets?.["Cierre"] || 0));
  setMetric("com-04", "Localidad principal", topLocation[0], `${formatNumber(topLocation[1])} casos`);
  setMetric("com-04", "Seguimiento abierto", formatNumber(open));
  replaceSectionContent("com-04", "Localidades con mayor seguimiento", buildBars(entriesSorted(data.locationsOpen || {}, 6), false));
  replaceSectionContent(
    "com-04",
    "Estado de seguimiento abierto",
    buildDonut([
      { name: "Reactivación", value: data.buckets?.["Reactivación"] || 0 },
      { name: "Cierre", value: data.buckets?.["Cierre"] || 0 },
      { name: "Desarrollo", value: data.buckets?.["Desarrollo"] || 0 }
    ], formatNumber((data.buckets?.["Reactivación"] || 0) + (data.buckets?.["Cierre"] || 0) + (data.buckets?.["Desarrollo"] || 0)), "Casos visibles", 168, true)
  );

  Object.entries(data.views || {}).forEach(([key, value]) => {
    app.viewTables[key] = value;
  });
  app.viewGroups.alcance_directores = ["diego_propias", "diego_referenciadas", "jorge_propias", "jorge_referenciadas"];
}

function applyPayload(payload, persist = true) {
  if (payload.type === "operational") updateOperationalVisual(payload);
  if (payload.type === "commercial") updateCommercialVisual(payload);
  if (persist) {
    const existing = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
    existing[payload.type] = payload;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(existing));
  }
  app.scaleRepeated?.();
}

function showValidation(payload, file) {
  const items = [];
  if (payload.type === "operational") {
    items.push(
      ["Archivo", file.name],
      ["Periodo detectado", `${MONTH_LABELS[payload.periodMonth] || ""} ${payload.periodYear || ""}`],
      ["Operaciones Pipeline", formatNumber(Object.values(payload.stages).reduce((s, x) => s + x.count, 0))],
      ["Proyección", formatNumber(payload.projection.projection.length)],
      ["Dispersiones", formatNumber(payload.projection.dispersions.length)],
      ["Histórico 2026", `${payload.historical.length} meses`]
    );
  } else {
    items.push(
      ["Archivo", file.name],
      ["Prospectos considerados", formatNumber(payload.prospects.length)],
      ["Fuera del alcance", formatNumber(payload.excludedCount || 0)],
      ["Seguimiento abierto", formatNumber(payload.open.length)],
      ["Pagados", formatNumber(payload.buckets["Pagado"] || 0)],
      ["Diego · propias / referidas", `${formatNumber(payload.directorScopes?.diego?.ownCount || 0)} / ${formatNumber(payload.directorScopes?.diego?.referredCount || 0)}`],
      ["Jorge · propias / referidas", `${formatNumber(payload.directorScopes?.jorge?.ownCount || 0)} / ${formatNumber(payload.directorScopes?.jorge?.referredCount || 0)}`],
      [`Cierres ${MONTH_LABELS[payload.currentCloseMonthIndex] || "actual"}`, formatNumber(payload.currentClosings?.length || 0)],
      [`Cierres ${MONTH_LABELS[payload.nextCloseMonthIndex] || "siguiente"}`, formatNumber(payload.nextClosings?.length || 0)]
    );
  }
  updateValidation.innerHTML = items.map(([label, value]) => `
    <div class="validation-item"><span>${label}</span><strong>${value}</strong></div>
  `).join("");
  updateValidation.classList.add("visible");
}

async function processFile(file) {
  if (!file) return;
  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    throw new Error("Selecciona un archivo Excel con extensión .xlsx o .xls.");
  }
  updateStatus.className = "update-status";
  updateStatus.textContent = "Analizando archivo…";
  applyUpdate.disabled = true;
  updateValidation.classList.remove("visible");
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  pendingPayload = updateType === "operational"
    ? parseOperationalWorkbook(workbook)
    : parseCommercialWorkbook(workbook);
  showValidation(pendingPayload, file);
  updateStatus.className = "update-status success";
  updateStatus.textContent = "Archivo validado. Revisa el resumen antes de aplicar.";
  applyUpdate.disabled = false;
}

function resetUploadUI() {
  pendingPayload = null;
  fileInput.value = "";
  applyUpdate.disabled = true;
  updateStatus.className = "update-status";
  updateStatus.textContent = "";
  updateValidation.classList.remove("visible");
  updateValidation.innerHTML = "";
}

function closeToStart() {
  updateChoiceOverlay.classList.remove("visible");
  uploadOverlay.classList.remove("visible");
  resetUploadUI();
  app.showStart?.();
}

openUpdateFlow?.addEventListener("click", () => {
  document.getElementById("selectorOverlay")?.classList.remove("visible");
  updateChoiceOverlay.classList.add("visible");
});

cancelUpdateChoice?.addEventListener("click", closeToStart);
cancelUpload?.addEventListener("click", closeToStart);

$$(".update-choice-card").forEach(button => {
  button.addEventListener("click", () => {
    updateType = button.dataset.updateType;
    updateChoiceOverlay.classList.remove("visible");
    uploadOverlay.classList.add("visible");
    uploadTitle.textContent = updateType === "operational"
      ? "Actualizar seguimiento operativo"
      : "Actualizar seguimiento comercial";
    uploadDescription.textContent = updateType === "operational"
      ? "Carga el archivo actualizado del Pipeline operativo."
      : "Carga el archivo actualizado de prospectos comerciales. Se excluirán Nancy, Pedro y Erika/Ericka; además se separará la cartera propia y referenciada de Diego y Jorge.";
    resetUploadUI();
  });
});

selectFileBtn?.addEventListener("click", () => fileInput.click());
fileInput?.addEventListener("change", async event => {
  try {
    await processFile(event.target.files?.[0]);
  } catch (error) {
    updateStatus.className = "update-status error";
    updateStatus.textContent = error.message || "No se pudo procesar el archivo.";
  }
});

["dragenter", "dragover"].forEach(name => {
  dropZone?.addEventListener(name, event => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  });
});
["dragleave", "drop"].forEach(name => {
  dropZone?.addEventListener(name, event => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  });
});
dropZone?.addEventListener("drop", async event => {
  try {
    await processFile(event.dataTransfer.files?.[0]);
  } catch (error) {
    updateStatus.className = "update-status error";
    updateStatus.textContent = error.message || "No se pudo procesar el archivo.";
  }
});

applyUpdate?.addEventListener("click", () => {
  if (!pendingPayload) return;
  applyPayload(pendingPayload, true);
  updateStatus.className = "update-status success";
  updateStatus.textContent = "Presentación actualizada correctamente.";
  applyUpdate.disabled = true;
  setTimeout(closeToStart, 900);
});

// Restore last browser-saved update. V19 uses a new key so previous unfiltered data cannot override this version.
try {
  const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
  if (saved.operational) applyPayload(saved.operational, false);
  if (saved.commercial) {
    applyPayload(saved.commercial, false);
  } else if (window.__KONNECT_V19_INITIAL__) {
    applyPayload(window.__KONNECT_V19_INITIAL__, false);
  }
} catch (error) {
  console.warn("No se pudo restaurar la última actualización:", error);
  if (window.__KONNECT_V19_INITIAL__) applyPayload(window.__KONNECT_V19_INITIAL__, false);
}
