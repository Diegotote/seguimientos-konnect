
import * as XLSX from "xlsx";

const app = window.__KONNECT__;
const STORAGE_KEY = "konnect_dashboard_v35_data";

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
const OPERATIONAL_TARGET_OVERRIDES = {
  7: 76169192,
  8: 67769192,
  9: 76169192,
  10: 76169192,
  11: 60000000
};


const FINANCIAL_NAME_ALIASES = {
  FLUXO: "Fluxo",
  HEYBANCO: "HeyBanco",
  FINAMO: "Finamo",
  PYMIO: "Pymio",
  FINKARGO: "Finkargo",
  AFIRME: "Afirme",
  FINSUS: "Finsus",
  FONDEADORA: "Fondeadora",
  FINBEABC: "FinBeABC",
  CLARA: "Clara",
  JEEVES: "Jeeves",
  BANCREA: "Bancrea",
  SPIN: "Spin",
  XEPELIN: "Xepelin",
  HAYCASH: "HayCash",
  RENDAFIN: "Rendafin",
  CASOFIN: "Casofin",
  PRIORITA: "Priorita",
  UNIFIN: "Unifin",
  ENGEN: "Engen",
  BANORTE: "Banorte",
  KAPITALIZER: "Kapitalizer",
  KONFIO: "Konfío",
  CUALLI: "Cualli"
};

const DEFAULT_FINANCIAL_COMMENT_ACTIVITY = {
  periodLabel: "Junio · Julio 2026",
  sourceYear: 2026,
  sourceMonthIndexes: [5, 6],
  note: "Comentarios = mensajes escritos por cuentas de la financiera. Operaciones comentadas = folios donde la financiera dejó al menos un comentario. Cobertura = operaciones comentadas ÷ operaciones totales del mismo periodo.",
  withComments: [
    { name: "Fluxo", comments: 87, share: 68.50, coverage: 69.49, fallbackTotalOperations: 15 },
    { name: "HeyBanco", comments: 24, share: 18.90, coverage: 62.50, fallbackTotalOperations: 6 },
    { name: "Finamo", comments: 6, share: 4.72, coverage: 50.00, fallbackTotalOperations: 4 },
    { name: "Pymio", comments: 6, share: 4.72, coverage: 3.20, fallbackTotalOperations: 65 },
    { name: "Finkargo", comments: 4, share: 3.15, coverage: 100.00, fallbackTotalOperations: 3 }
  ],
  fallbackWithoutComments: [
    { name: "Afirme", activeOperations: 23 },
    { name: "Finsus", activeOperations: 17 },
    { name: "Fondeadora", activeOperations: 15 },
    { name: "FinBeABC", activeOperations: 12 },
    { name: "Clara", activeOperations: 10 },
    { name: "Jeeves", activeOperations: 5 },
    { name: "Bancrea", activeOperations: 4 },
    { name: "Spin", activeOperations: 3 },
    { name: "Xepelin", activeOperations: 3 },
    { name: "HayCash", activeOperations: 3 },
    { name: "Rendafin", activeOperations: 3 },
    { name: "Casofin", activeOperations: 2 },
    { name: "Priorita", activeOperations: 2 },
    { name: "Unifin", activeOperations: 2 },
    { name: "Engen", activeOperations: 1 },
    { name: "Banorte", activeOperations: 1 },
    { name: "Kapitalizer", activeOperations: 1 },
    { name: "Konfío", activeOperations: 1 },
    { name: "Cualli", activeOperations: 1 }
  ]
};

function prettyFinancialName(value) {
  const key = normalizeText(value);
  return FINANCIAL_NAME_ALIASES[key] || String(value ?? "Sin financiera").trim() || "Sin financiera";
}

function buildFinancialCommentActivity(rows) {
  const sourceYear = DEFAULT_FINANCIAL_COMMENT_ACTIVITY.sourceYear;
  const sourceMonths = DEFAULT_FINANCIAL_COMMENT_ACTIVITY.sourceMonthIndexes;
  const activeRows = (rows || []).filter(row =>
    row.date && row.date.getFullYear() === sourceYear && sourceMonths.includes(row.date.getMonth())
  );
  const counts = countBy(activeRows, row => normalizeText(row.financial) || "SIN FINANCIERA");
  const commentedKeys = new Set(DEFAULT_FINANCIAL_COMMENT_ACTIVITY.withComments.map(item => normalizeText(item.name)));
  const withComments = DEFAULT_FINANCIAL_COMMENT_ACTIVITY.withComments.map(item => {
    const totalOperations = counts[normalizeText(item.name)] || item.fallbackTotalOperations || 0;
    const commentedOperations = totalOperations
      ? Math.max(1, Math.min(totalOperations, Math.round(totalOperations * item.coverage / 100)))
      : 0;
    const coverage = totalOperations ? commentedOperations / totalOperations * 100 : 0;
    return {
      ...item,
      totalOperations,
      commentedOperations,
      coverage
    };
  });
  const withoutComments = entriesSorted(counts)
    .filter(([name, value]) => !commentedKeys.has(name) && Number(value || 0) > 0)
    .map(([name, activeOperations]) => ({
      name: prettyFinancialName(name),
      activeOperations
    }));
  return {
    periodLabel: DEFAULT_FINANCIAL_COMMENT_ACTIVITY.periodLabel,
    note: DEFAULT_FINANCIAL_COMMENT_ACTIVITY.note,
    totalComments: sumBy(withComments, item => item.comments),
    totalCommentedOperations: sumBy(withComments, item => item.commentedOperations),
    withComments,
    withoutComments: withoutComments.length ? withoutComments : DEFAULT_FINANCIAL_COMMENT_ACTIVITY.fallbackWithoutComments,
    commentingFinancials: withComments.length,
    silentFinancials: (withoutComments.length ? withoutComments : DEFAULT_FINANCIAL_COMMENT_ACTIVITY.fallbackWithoutComments).length
  };
}

function financialCoverageTone(value) {
  if (Number(value || 0) >= 80) return "high";
  if (Number(value || 0) >= 40) return "mid";
  return "low";
}

function renderFinancialCommentActivity(activity) {
  const section = document.getElementById("op-06");
  if (!section || !activity) return;

  const chip = $(".finance-activity-chip", section);
  if (chip) chip.textContent = activity.periodLabel;

  const summary = $(".finance-activity-summary", section);
  if (summary) summary.textContent = `${formatNumber(activity.commentingFinancials)} financieras comentaron · ${formatNumber(activity.silentFinancials)} financieras tuvieron operaciones activas sin comentario directo en plataforma.`;

  const donutHost = $(".finance-donut-host", section);
  if (donutHost) {
    const colors = ["#19d6e2", "#3db7ff", "#6b74ff", "#8f5cff", "#12a7ff", "#3df0c7", "#725cff"];
    const legendRows = activity.withComments.map((item, index) => `
      <div class="finance-legend-row">
        <span class="dot" style="background:${colors[index % colors.length]}"></span>
        <div>
          <div class="finance-legend-title">Financiera</div>
          <div class="finance-legend-name">${escapeHtml(item.name)}</div>
        </div>
        <div class="finance-legend-meta">${formatNumber(item.comments)} comentarios · ${formatNumber(item.commentedOperations)} ops.</div>
      </div>
    `).join("");

    donutHost.innerHTML = buildDonut(
      activity.withComments.map(item => ({ name: item.name, value: item.commentedOperations })),
      formatNumber(activity.totalCommentedOperations),
      "Operaciones comentadas",
      260,
      true
    );

    let legend = $(".finance-legend-list", section);
    if (!legend) {
      legend = document.createElement("div");
      legend.className = "finance-legend-list";
      donutHost.insertAdjacentElement("afterend", legend);
    }
    legend.innerHTML = legendRows;
  }

  const commentTotal = $(".finance-comment-total", section);
  if (commentTotal) commentTotal.textContent = `${formatNumber(activity.commentingFinancials)} financieras`;

  const commentBody = $(".finance-commented-table tbody", section);
  if (commentBody) {
    commentBody.innerHTML = activity.withComments.map(item => `
      <tr>
        <td class="finance-name-cell">${escapeHtml(item.name)}</td>
        <td class="num">${formatNumber(item.totalOperations)}</td>
        <td class="num">${formatNumber(item.commentedOperations)}</td>
        <td class="num"><span class="finance-coverage-chip ${financialCoverageTone(item.coverage)}">${formatPercent(item.coverage)}</span></td>
      </tr>
    `).join("");
  }

  const noCommentList = activity.withoutComments || [];
  const noCommentTotal = $(".finance-no-comment-total", section);
  if (noCommentTotal) noCommentTotal.textContent = `${formatNumber(sumBy(noCommentList, item => item.activeOperations))} operaciones activas`;

  const noCommentBody = $(".finance-no-comment-table tbody", section);
  if (noCommentBody) {
    noCommentBody.innerHTML = noCommentList.map(item => `
      <tr>
        <td class="finance-name-cell">${escapeHtml(item.name)}</td>
        <td class="num">${formatNumber(item.activeOperations)}</td>
      </tr>
    `).join("");
  }

  const footnote = $(".finance-footnote", section);
  if (footnote) footnote.textContent = activity.note;
}


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

  // Fechas de la primera columna PIPELINE pueden venir como texto: "8 ago 2026".
  const monthMap = {
    ENE: 0, FEB: 1, MAR: 2, ABR: 3, MAY: 4, JUN: 5,
    JUL: 6, AGO: 7, SEP: 8, OCT: 9, NOV: 10, DIC: 11
  };
  const spanish = normalizeText(s).match(/^(\d{1,2})\s+([A-Z]{3,10})\s+(\d{4})$/);
  if (spanish) {
    const monthIndex = monthMap[spanish[2].slice(0, 3)];
    if (monthIndex != null) {
      const d = new Date(Number(spanish[3]), monthIndex, Number(spanish[1]));
      return Number.isNaN(d.valueOf()) ? null : d;
    }
  }

  const d = new Date(s);
  return Number.isNaN(d.valueOf()) ? null : d;
}

function normalizeOperationalPeriodDate(value) {
  const date = value instanceof Date && !Number.isNaN(value.valueOf()) ? value : null;
  if (!date) return null;
  const year = date.getFullYear();
  if (year < 2024 || year > 2030) return null;
  return date;
}


function formatDate(value) {
  const d = parseDate(value);
  if (!d) return String(value ?? "");
  return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getFullYear()).slice(-2)}`;
}

function formatActivityDate(value) {
  const d = parseDate(value);
  if (!d) return String(value ?? "").trim();
  return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
}

function parseScheduleTime(value) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return value;
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(2000, 0, 1, parsed.H || 0, parsed.M || 0, parsed.S || 0);
  }
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const text = normalizeText(raw).replace(/\./g, "");
  const match = text.match(/(\d{1,2})(?::(\d{2}))?\s*(AM|PM|A M|P M)?/);
  if (!match) return null;
  let hour = Number(match[1]);
  const minute = Number(match[2] || 0);
  const meridian = (match[3] || "").replace(/\s/g, "");
  if (meridian === "PM" && hour < 12) hour += 12;
  if (meridian === "AM" && hour === 12) hour = 0;
  return new Date(2000, 0, 1, hour, minute, 0);
}

function formatScheduleTime(value) {
  const d = parseScheduleTime(value);
  if (!d) return String(value ?? "").trim();
  const hour24 = d.getHours();
  const minute = String(d.getMinutes()).padStart(2, "0");
  const meridian = hour24 >= 12 ? "p. m." : "a. m.";
  const hour12 = hour24 % 12 || 12;
  return `${hour12}:${minute} ${meridian}`;
}

function parseWeeklyActivitiesRows(rows) {
  if (!rows) return [];
  const headerIndex = rows.findIndex(row => {
    const headers = mapHeaders(row || []);
    return findHeaderIndex(headers, ["DIA", "DÍA"]) >= 0 && findHeaderIndex(headers, ["PERSONA", "NOMBRE"]) >= 0;
  });
  if (headerIndex < 0) return [];
  const headers = mapHeaders(rows[headerIndex]);
  const idx = {
    day: findHeaderIndex(headers, ["DIA", "DÍA"]),
    date: findHeaderIndex(headers, ["FECHA"]),
    time: findHeaderIndex(headers, ["HORA"]),
    person: findHeaderIndex(headers, ["PERSONA", "NOMBRE"]),
    activity: findHeaderIndex(headers, ["ACTIVIDAD", "DETALLE", "REUNION", "REUNIÓN"])
  };
  const activities = rows.slice(headerIndex + 1).map(row => {
    const person = idx.person >= 0 ? String(row[idx.person] ?? "").trim() : "";
    const activity = idx.activity >= 0 ? String(row[idx.activity] ?? "").trim() : "";
    if (!person && !activity) return null;
    const dateValue = idx.date >= 0 ? row[idx.date] : "";
    const timeValue = idx.time >= 0 ? row[idx.time] : "";
    return {
      day: idx.day >= 0 ? String(row[idx.day] ?? "").trim() : "",
      date: formatActivityDate(dateValue),
      rawDate: parseDate(dateValue)?.toISOString?.() || "",
      time: formatScheduleTime(timeValue),
      person: person || "Sin persona",
      activity: activity || "Reunión"
    };
  }).filter(Boolean);

  return activities.sort((a, b) => {
    const da = a.rawDate ? new Date(a.rawDate).getTime() : 0;
    const db = b.rawDate ? new Date(b.rawDate).getTime() : 0;
    if (da !== db) return da - db;
    const ta = parseScheduleTime(a.time)?.getTime?.() || 0;
    const tb = parseScheduleTime(b.time)?.getTime?.() || 0;
    return ta - tb;
  });
}

function sheetRows(workbook, name, raw = true) {
  const exact = workbook.SheetNames.find(n => normalizeText(n) === normalizeText(name));
  if (!exact) return null;
  return XLSX.utils.sheet_to_json(workbook.Sheets[exact], { header: 1, raw, defval: null });
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

function getOperationalTargetForMonth(monthIndex, fallback = 0) {
  return OPERATIONAL_TARGET_OVERRIDES[monthIndex] || fallback || 0;
}

function getGapAmount(target, actual) {
  return Number(target || 0) - Number(actual || 0);
}

function formatGapText(target, actual) {
  const gap = getGapAmount(target, actual);
  if (gap >= 0) return `Gap -${formatMoney(gap)} frente a la dispersión actual.`;
  return `Meta superada por ${formatMoney(Math.abs(gap))}.`;
}

function renderOperationalFutureTargets(currentMonthIndex, currentAmount) {
  const host = document.querySelector('#op-01 .future-targets-list');
  if (!host) return;
  const futureEntries = Object.entries(OPERATIONAL_TARGET_OVERRIDES)
    .map(([monthIndex, target]) => ({ monthIndex: Number(monthIndex), target }))
    .filter(item => item.monthIndex > Number(currentMonthIndex))
    .sort((a, b) => a.monthIndex - b.monthIndex);

  if (!futureEntries.length) {
    host.innerHTML = '<div class="history-empty">No hay metas posteriores configuradas.</div>';
    return;
  }

  host.innerHTML = futureEntries.map(item => {
    const gap = getGapAmount(item.target, currentAmount);
    const gapText = `${gap >= 0 ? '-' : '+'}${formatMoney(Math.abs(gap))}`;
    return `
      <div class="future-target-row">
        <div class="future-target-month">${escapeHtml(MONTH_LABELS[item.monthIndex] || '')}</div>
        <div class="future-target-target">
          <strong>${formatMoney(item.target)}</strong>
          <span>Meta mensual</span>
        </div>
        <div class="future-target-gap">
          <strong>${formatMoney(item.target)}</strong>
          <span>Meta</span>
        </div>
      </div>
    `;
  }).join('');
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


function parseClosures2026Rows(rows) {
  const parsed = [];
  let currentMonthIndex = null;
  let currentMonthName = "";
  let currentYear = new Date().getFullYear();

  (rows || []).forEach(row => {
    const firstCell = normalizeText(row?.[0]);
    const sectionMatch = firstCell.match(/CIERRE\s+(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)(?:\s+(\d{4}))?/);
    if (sectionMatch) {
      currentMonthName = sectionMatch[1];
      currentMonthIndex = MONTHS.findIndex(month => normalizeText(month) === currentMonthName);
      if (sectionMatch[2]) currentYear = Number(sectionMatch[2]);
      return;
    }

    if (currentMonthIndex == null) return;
    const director = String(row?.[0] ?? "").trim();
    const amountRaw = row?.[1];
    const financial = String(row?.[2] ?? "").trim();
    const broker = String(row?.[3] ?? "").trim();
    const client = String(row?.[4] ?? "").trim();
    const normalizedDirector = normalizeText(director);

    if (!director || !financial || !broker || !client) return;
    if (
      normalizedDirector.includes("DIRECTOR COMERCIAL") ||
      normalizedDirector.includes("TOTAL DISPERSADO") ||
      normalizedDirector.includes("MONTO DISPERSADO") ||
      normalizedDirector.startsWith("ZONA ")
    ) return;

    parsed.push({
      month: MONTHS[currentMonthIndex],
      monthIndex: currentMonthIndex,
      year: currentYear,
      director,
      amount: toNumber(amountRaw),
      amountDisplay: formatMoney(toNumber(amountRaw)),
      financial,
      broker,
      client,
      name: client,
      office: broker,
      region: financial,
      membership: director
    });
  });

  return parsed;
}

function buildOperationalTableViews(stageRows, projectionRows = []) {
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

  return {
    viabilidad: {
      title: "Operaciones en Viabilidad",
      columns: ["Fecha", "Folio", "Cliente", "Consultoría", "Financiera", "Producto", "Monto solicitado", "Comentario"],
      rows: stageRows["Viabilidad"].rows.map(tableRow),
      summary: [
        { label: "Operaciones", value: formatNumber(stageRows["Viabilidad"].count) },
        { label: "Monto solicitado", value: formatMoney(stageRows["Viabilidad"].requested) }
      ]
    },
    integracion: {
      title: "Operaciones en Integración",
      columns: ["Fecha", "Folio", "Cliente", "Financiera", "Producto", "Monto solicitado", "Motivo"],
      rows: stageRows["Integración"].rows.map(integrationTableRow),
      summary: [
        { label: "Operaciones", value: formatNumber(stageRows["Integración"].count) },
        { label: "Monto solicitado", value: formatMoney(stageRows["Integración"].requested) }
      ]
    },
    analisis: {
      title: "Operaciones en Análisis",
      columns: ["Fecha", "Folio", "Cliente", "Consultoría", "Financiera", "Producto", "Monto solicitado", "Comentario"],
      rows: stageRows["Análisis"].rows.map(tableRow),
      summary: [
        { label: "Operaciones", value: formatNumber(stageRows["Análisis"].count) },
        { label: "Monto solicitado", value: formatMoney(stageRows["Análisis"].requested) }
      ]
    },
    autorizacion: {
      title: "Operaciones en Autorización",
      columns: ["Fecha", "Folio", "Cliente", "Financiera", "Producto", "Monto solicitado", "Monto otorgado"],
      rows: stageRows["Autorización"].rows.map(finalTableRow),
      summary: [
        { label: "Operaciones", value: formatNumber(stageRows["Autorización"].count) },
        { label: "Monto solicitado", value: formatMoney(stageRows["Autorización"].requested) },
        { label: "Monto otorgado", value: formatMoney(stageRows["Autorización"].granted) }
      ]
    },
    formalizacion: {
      title: "Operaciones en Formalización",
      columns: ["Fecha", "Folio", "Cliente", "Financiera", "Producto", "Monto solicitado", "Monto otorgado"],
      rows: stageRows["Formalización"].rows.map(finalTableRow),
      summary: [
        { label: "Operaciones", value: formatNumber(stageRows["Formalización"].count) },
        { label: "Monto solicitado", value: formatMoney(stageRows["Formalización"].requested) },
        { label: "Monto otorgado", value: formatMoney(stageRows["Formalización"].granted) }
      ]
    },
    proyeccion: {
      title: "Operaciones en Proyección",
      columns: ["Cliente", "Financiera", "Broker / Consultoría", "Monto"],
      rows: projectionRows.map(x => [x.client, x.financial, x.broker, formatMoney(x.amount)]),
      summary: [
        { label: "Operaciones", value: formatNumber(projectionRows.length) },
        { label: "Potencial", value: formatMoney(sumBy(projectionRows, x => x.amount)) }
      ]
    },
    dispersion: {
      title: "Operaciones en Dispersión",
      columns: ["Cliente", "Financiera", "Broker / Consultoría", "Monto dispersado"],
      rows: stageRows["Dispersión"].rows.map(x => [x.client, x.financial, x.broker, formatMoney(x.amount || x.granted || x.requested || 0)]),
      summary: [
        { label: "Operaciones", value: formatNumber(stageRows["Dispersión"].count) },
        { label: "Monto dispersado", value: formatMoney(stageRows["Dispersión"].requested) }
      ]
    }
  };
}

function buildOperationalPeriodSummary(pipelineRows, closures2026, monthIndex, year, projectionRows = []) {
  const currentRows = (pipelineRows || []).filter(row =>
    row.date && row.date.getFullYear() === year && row.date.getMonth() === monthIndex
  );
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

  const dispersionRows = (closures2026 || []).filter(row => row.year === year && row.monthIndex === monthIndex);
  stages["Dispersión"] = {
    rows: dispersionRows,
    count: dispersionRows.length,
    requested: sumBy(dispersionRows, x => x.amount),
    granted: sumBy(dispersionRows, x => x.amount)
  };

  return {
    key: `${year}-${String(monthIndex + 1).padStart(2, "0")}`,
    monthIndex,
    year,
    label: `${MONTHS[monthIndex]} ${year}`,
    stages,
    integrationBlockers: countBy(stages["Integración"].rows, x => classifyBlocker(x.comment)),
    views: buildOperationalTableViews(stages, projectionRows)
  };
}

function parseOperationalWorkbook(workbook) {
  const pipelineRows = sheetRows(workbook, "PIPELINE");
  const projectionRows = sheetRows(workbook, "PROYECCIÓN") || sheetRows(workbook, "PROYECCION");
  const closureRows = sheetRows(workbook, "CIERRES 2026");
  if (!pipelineRows || !projectionRows || !closureRows) {
    throw new Error("El archivo debe contener las hojas PIPELINE, PROYECCIÓN y CIERRES 2026.");
  }

  const closures2026 = parseClosures2026Rows(closureRows);

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
  })).map(row => ({
    ...row,
    date: normalizeOperationalPeriodDate(row.date)
  })).filter(row => row.client || row.folio || row.status);

  const dated = pipeline.filter(row => row.date);
  const latestDate = dated.sort((a, b) => b.date - a.date)[0]?.date || new Date();
  const projection = parseProjectionSheet(projectionRows);
  const historical = parseHistoricalClosings(closureRows);
  const target = projection.target || 65000000;
  const dispersed = sumBy(projection.dispersions, x => x.amount);
  const missing = Math.max(0, target - dispersed);
  const progress = target ? dispersed / target * 100 : 0;

  const reportingDate = new Date();
  const reportingMonth = reportingDate.getMonth();
  const reportingYear = reportingDate.getFullYear();
  const currentSummary = buildOperationalPeriodSummary(pipeline, closures2026, reportingMonth, reportingYear, projection.projection);
  const stages = currentSummary.stages;
  const integrationBlockers = currentSummary.integrationBlockers;
  const views = currentSummary.views;

  const periodRegistry = new Map();
  pipeline.forEach(row => {
    const date = row.date;
    if (!date) return;
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    if (!periodRegistry.has(key)) periodRegistry.set(key, { monthIndex: date.getMonth(), year: date.getFullYear() });
  });
  closures2026.forEach(row => {
    const key = `${row.year}-${String(row.monthIndex + 1).padStart(2, "0")}`;
    if (!periodRegistry.has(key)) periodRegistry.set(key, { monthIndex: row.monthIndex, year: row.year });
  });

  const periodSummaries = [...periodRegistry.values()]
    .sort((a, b) => (b.year - a.year) || (b.monthIndex - a.monthIndex))
    .map(period => buildOperationalPeriodSummary(pipeline, closures2026, period.monthIndex, period.year, projection.projection));

  const currentPeriodKey = `${reportingYear}-${String(reportingMonth + 1).padStart(2, "0")}`;
  const latestPeriodKey = periodSummaries.some(period => period.key === currentPeriodKey)
    ? currentPeriodKey
    : (periodSummaries[0]?.key || currentSummary.key);

  const projectionByFinancial = moneyBy(projection.projection, x => x.financial, x => x.amount);
  const dispersionByFinancial = moneyBy(projection.dispersions, x => x.financial, x => x.amount);
  const dispersionCountByFinancial = countBy(projection.dispersions, x => normalizeText(x.financial) || 'Sin financiera');

  const previousMonthIndex = (reportingMonth + 11) % 12;
  const previousHistory = historical.find(x => x.monthIndex === previousMonthIndex) || {
    monthIndex: previousMonthIndex,
    label: MONTH_LABELS[previousMonthIndex],
    amount: 0
  };
  const previousAmount = previousHistory.amount || 0;
  const previousProgress = target ? previousAmount / target * 100 : 0;

  return {
    type: "operational",
    importedAt: new Date().toISOString(),
    periodMonth: reportingMonth,
    periodYear: reportingYear,
    target,
    dispersed,
    missing,
    progress,
    financialCommentActivity: buildFinancialCommentActivity(pipeline),
    closures2026,
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
    views,
    periodSummaries,
    latestPeriodKey
  };
}


const EXCLUDED_COMMERCIAL_NAMES = ["NANCY", "PEDRO", "ERIKA", "ERICKA"];

function hasExcludedCommercialName(value) {
  const text = normalizeText(value);
  return EXCLUDED_COMMERCIAL_NAMES.some(name => text.includes(name));
}

function isBajaCommercialRecord(record) {
  return normalizeText(record?.status).includes("BAJA");
}

function isExcludedCommercialRecord(record) {
  if (isBajaCommercialRecord(record)) return true;
  const directorExcluded = hasExcludedCommercialName(record.director);
  const referredByErika = ["ERIKA", "ERICKA"].some(name =>
    normalizeText(record.referred).includes(name)
  );
  return directorExcluded || referredByErika;
}

const HISTORICAL_MEMBERSHIP_CLOSINGS = [{"month":"MARZO","monthIndex":2,"name":"HUGO NAVARRO / ALEJANDRO ORTA","office":"ROIESTATE","email":"roiestate@konnect.mx","phone":"5626001647 / 3339431770","region":"GUADALAJARA","membership":"KONNECT EVOLUCIONA"},{"month":"MARZO","monthIndex":2,"name":"MIGUEL PARDO","office":"BULLDOG FIANCE","email":"bulldog.finance@konnect.mx","phone":"4427479785","region":"QUERÉTARO","membership":"KONNECT EVOLUCIONA"},{"month":"ABRIL","monthIndex":3,"name":"BEATRIZ PLACENCIA","office":"NEXORA CAPITAL","email":"nexora.capital@konnect.mx","phone":"5518012750","region":"CDMX","membership":"KONNECT EVOLUCIONA"},{"month":"ABRIL","monthIndex":3,"name":"TANIA GUILLERMINA REYES FLORES / DAMIAN REYES","office":"XQUENDA","email":"xquenda@konnect.mx","phone":"5573257692 / 5523166908","region":"CDMX","membership":"KONNECT EVOLUCIONA"},{"month":"ABRIL","monthIndex":3,"name":"EMMANUEL SALGADO","office":"ESTRON","email":"estron@konnect.mx","phone":"5522185956","region":"CDMX","membership":"KONNECT EVOLUCIONA"},{"month":"ABRIL","monthIndex":3,"name":"RAMON ENRRIQUEZ","office":"BULL KAPITAL","email":"bullkapital@konnect.mx","phone":"2226507287","region":"PUEBLA","membership":"KONNECT EVOLUCIONA"},{"month":"MAYO","monthIndex":4,"name":"GERARDO RENDÓN","office":"ARF","email":"arf@konnect.mx","phone":"8111765725","region":"MONTERREY","membership":"KONNECT EVOLUCIONA"},{"month":"MAYO","monthIndex":4,"name":"JUAN OTERO","office":"KREO","email":"comercial1kreojom@konnect.mx","phone":"5525596582","region":"CDMX","membership":"KONNECT EVOLUCIONA"},{"month":"MAYO","monthIndex":4,"name":"JAIME CÁRDENAS","office":"CEAF SERVICIOS","email":"ceaf.servicios@konnect.mx","phone":"8119384852","region":"MONTERREY","membership":"KONNECT EVOLUCIONA"},{"month":"JUNIO","monthIndex":5,"name":"MAURICIO ESPINOSA / MAYERIK ADAME ESPINOSA","office":"ASESORES PYME A&M","email":"mayerik.adame@konnect.mx","phone":"4461390401","region":"QUERÉTARO","membership":"KONNECT EVOLUCIONA"},{"month":"JUNIO","monthIndex":5,"name":"LUIS ANTONIO REYES","office":"LUIS ANTONIO REYES","email":"antonio.reyes@konnect.mx","phone":"5539961559","region":"CDMX","membership":"KONNECT EVOLUCIONA"},{"month":"JULIO","monthIndex":6,"name":"CARLOS ALEJANDRO ÁLVAREZ","office":"CAPITAL DAF","email":"capitaldaf@konnect.mx","phone":"3311785168","region":"GUADALAJARA","membership":"KONNECT EVOLUCIONA"},{"month":"JULIO","monthIndex":6,"name":"JOSÉ ERNESTO GARCÍA","office":"LENDUM","email":"egarcia@konnect.mx","phone":"6621800998","region":"HERMOSILLO","membership":"KONNECT EVOLUCIONA"},{"month":"AGOSTO","monthIndex":7,"name":"MELANY MARIEL TORRES GUDIÑO","office":"MTG Capital","email":"mtgcapital@konnect.mx","phone":"55 4800 9786","region":"ESTADO DE MÉXICO","membership":"KONNECT EVOLUCIONA"}];

function isHistoricalClosedName(value) {
  const text = normalizeText(value);
  return text.includes("ERNESTO GARCIA") ||
    text.includes("JOSE ERNESTO GARCIA") ||
    text.includes("ALEJANDRO ALVAREZ") ||
    text.includes("ALEJANDRO ALVARES") ||
    text.includes("CARLOS ALEJANDRO ALVAREZ") ||
    text.includes("ALEJANDRO VIERA");
}

function isOpenCommercialFollowUp(row) {
  if (!row || isBajaCommercialRecord(row)) return false;
  if (isHistoricalClosedName(row.name)) return false;
  const bucket = bucketCommercialStatus(row.status);
  return bucket !== "Pagado" && bucket !== "No viable";
}

function uniqueProspects(rows) {
  const seen = new Set();
  return (rows || []).filter(row => {
    const key = [normalizeText(row?.name), normalizeText(row?.director), normalizeText(row?.referred), normalizeText(row?.status)].join("|");
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
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
    if (!isOpenCommercialFollowUp(row)) return;
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


function parseCommercialClosures2026Rows(rows) {
  const parsed = [];
  let currentMonthIndex = null;
  let currentMonthName = "";
  let currentYear = new Date().getFullYear();
  let headerMap = null;

  const monthPattern = /(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)/;

  (rows || []).forEach(row => {
    const values = (row || []).map(value => String(value ?? "").trim());
    const normalized = values.map(normalizeText);
    const firstNonEmptyIndex = normalized.findIndex(Boolean);
    const firstValue = firstNonEmptyIndex >= 0 ? normalized[firstNonEmptyIndex] : "";
    const monthMatch = firstValue.match(monthPattern);

    // Reconoce encabezados de sección como "MARZO", "CIERRES MARZO 2026"
    // o "CIERRE DE MARZO" sin confundirlos con filas de datos.
    if (monthMatch && normalized.filter(Boolean).length <= 2) {
      currentMonthName = monthMatch[1];
      currentMonthIndex = MONTHS.findIndex(month => normalizeText(month) === currentMonthName);
      const yearMatch = firstValue.match(/(20\d{2})/);
      if (yearMatch) currentYear = Number(yearMatch[1]);
      headerMap = null;
      return;
    }

    const potentialHeaders = mapHeaders(values);
    const nameHeader = findHeaderIndex(potentialHeaders, ["TITULAR", "NOMBRE", "CLIENTE", "SOCIO"]);
    const officeHeader = findHeaderIndex(potentialHeaders, ["OFICINA", "CONSULTORIA", "BROKER"]);
    const emailHeader = findHeaderIndex(potentialHeaders, ["CORREO", "EMAIL"]);
    const phoneHeader = findHeaderIndex(potentialHeaders, ["TELEFONO", "CELULAR"]);
    const regionHeader = findHeaderIndex(potentialHeaders, ["REGION", "LOCALIDAD", "ZONA"]);
    const membershipHeader = findHeaderIndex(potentialHeaders, ["MEMBRESIA", "PROGRAMA"]);

    if (nameHeader >= 0 && (officeHeader >= 0 || membershipHeader >= 0 || emailHeader >= 0)) {
      headerMap = {
        name: nameHeader,
        office: officeHeader,
        email: emailHeader,
        phone: phoneHeader,
        region: regionHeader,
        membership: membershipHeader
      };
      return;
    }

    if (currentMonthIndex == null || !headerMap) return;

    const name = values[headerMap.name] || "";
    const office = headerMap.office >= 0 ? values[headerMap.office] : "";
    const email = headerMap.email >= 0 ? values[headerMap.email] : "";
    const phone = headerMap.phone >= 0 ? values[headerMap.phone] : "";
    const region = headerMap.region >= 0 ? values[headerMap.region] : "";
    const membership = headerMap.membership >= 0 ? values[headerMap.membership] : "";

    if (!name || normalizeText(name).includes("TOTAL")) return;
    if (!office && !email && !membership) return;

    parsed.push({
      month: MONTHS[currentMonthIndex],
      monthIndex: currentMonthIndex,
      year: currentYear,
      name,
      office: office || "—",
      email,
      phone,
      region: region || "—",
      membership: membership || "KONNECT EVOLUCIONA"
    });
  });

  return parsed;
}

function parseCommercialWorkbook(workbook) {
  const rows = sheetRows(workbook, "PROSPECTOS MEMBRESIAS");
  const closureRows = sheetRows(workbook, "CIERRES 2026");
  const weeklyRows = sheetRows(workbook, "Actividades Semanales", false);
  if (!rows) throw new Error("El archivo debe contener la hoja PROSPECTOS MEMBRESIAS.");
  const closures2026 = closureRows ? parseCommercialClosures2026Rows(closureRows) : [];
  const weeklyActivities = weeklyRows ? parseWeeklyActivitiesRows(weeklyRows) : [];

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
  const currentClosings = prospects.filter(row => row.closeMonthIndex === currentCloseMonthIndex && !isHistoricalClosedName(row.name) && isOpenCommercialFollowUp(row));
  const nextClosings = prospects.filter(row => row.closeMonthIndex === nextCloseMonthIndex && isOpenCommercialFollowUp(row));

  const directorScopes = {
    diego: buildDirectorScope(prospects, "DIEGO"),
    jorge: buildDirectorScope(prospects, "JORGE")
  };

  const focusedProspects = uniqueProspects([
    ...directorScopes.diego.own,
    ...directorScopes.diego.referred,
    ...directorScopes.jorge.own,
    ...directorScopes.jorge.referred
  ]).filter(isOpenCommercialFollowUp);
  const focusedBuckets = countBy(focusedProspects, row => bucketCommercialStatus(row.status));
  const focusDirectorsOpen = {
    DIEGO: directorScopes.diego.totalCount,
    JORGE: directorScopes.jorge.totalCount
  };
  const focusProspectKeys = new Set(focusedProspects.map(row => normalizeText(row.name)));
  const focusCurrentClosings = currentClosings.filter(row => focusProspectKeys.has(normalizeText(row.name)));

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
    seguimiento_diego_jorge: {
      title: "Seguimiento abierto · Diego y Jorge",
      columns: ["Nombre", "Programa / Membresía", "Director", "Localidad", "Estatus", "Teléfono", "Comentario"],
      rows: focusedProspects.map(rowForTable),
      summary: [{ label: "Registros", value: formatNumber(focusedProspects.length) }]
    },
    abiertos: {
      title: "Seguimiento abierto · Diego y Jorge",
      columns: ["Nombre", "Programa / Membresía", "Director", "Localidad", "Estatus", "Teléfono", "Comentario"],
      rows: focusedProspects.map(rowForTable),
      summary: [{ label: "Registros", value: formatNumber(focusedProspects.length) }]
    },
    pagados: makeView("Membresías pagadas", row => bucketCommercialStatus(row.status) === "Pagado"),
    reactivacion: makeView("Prospectos en reactivación", row => bucketCommercialStatus(row.status) === "Reactivación"),
    desarrollo: makeView("Prospectos en desarrollo", row => bucketCommercialStatus(row.status) === "Desarrollo"),
    cierre: makeView("Prospectos en cierre", row => bucketCommercialStatus(row.status) === "Cierre"),
    no_viable: makeView("Prospectos no viables", row => bucketCommercialStatus(row.status) === "No viable"),
    prioritarios: {
      title: `Cierres prioritarios de ${MONTHS[currentCloseMonthIndex].toLowerCase()}`,
      columns: ["Nombre", "Programa / Membresía", "Director", "Localidad", "Estatus", "Teléfono", "Comentario"],
      rows: focusCurrentClosings.map(rowForTable),
      summary: [{ label: "Registros", value: formatNumber(focusCurrentClosings.length) }]
    },
    cierres_siguiente: {
      title: `Cierres previstos para ${MONTHS[nextCloseMonthIndex].toLowerCase()}`,
      columns: ["Nombre", "Programa / Membresía", "Director", "Localidad", "Estatus", "Teléfono", "Comentario"],
      rows: nextClosings.map(rowForTable),
      summary: [{ label: "Registros", value: formatNumber(nextClosings.length) }]
    },
    cierres_transcurridos: {
      title: "Cierres de meses transcurridos",
      columns: ["Mes", "Titular", "Oficina", "Región", "Membresía", "Estatus"],
      rows: (closures2026.length ? closures2026 : HISTORICAL_MEMBERSHIP_CLOSINGS).map(row => [row.month, row.name, row.office, row.region, row.membership, "CERRADO"]),
      summary: [
        { label: "Cierres", value: formatNumber((closures2026.length ? closures2026 : HISTORICAL_MEMBERSHIP_CLOSINGS).length) },
        { label: "Meses", value: formatNumber(new Set((closures2026.length ? closures2026 : HISTORICAL_MEMBERSHIP_CLOSINGS).map(row => row.month)).size) }
      ]
    },
    actividades_semanales: {
      title: "Actividades semanales",
      columns: ["Día", "Fecha", "Hora", "Persona", "Actividad"],
      rows: weeklyActivities.map(row => [row.day, row.date, row.time, row.person, row.activity]),
      summary: [{ label: "Actividades", value: formatNumber(weeklyActivities.length) }]
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
    focusedProspects,
    focusedBuckets,
    focusDirectorsOpen,
    focusCurrentClosings,
    currentCloseMonthIndex,
    nextCloseMonthIndex,
    currentCloseYear,
    nextCloseYear,
    currentClosings,
    nextClosings,
    historicalMembershipClosings: closures2026.length ? closures2026 : HISTORICAL_MEMBERSHIP_CLOSINGS,
    closures2026,
    weeklyActivities,
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

function ensureOperationalPeriodUI() {
  if (!document.getElementById("operationalPeriodStyles")) {
    const style = document.createElement("style");
    style.id = "operationalPeriodStyles";
    style.textContent = `
      .period-overlay{position:fixed;inset:0;display:none;align-items:center;justify-content:center;background:rgba(3,10,24,.72);backdrop-filter:blur(10px);z-index:45;padding:24px}
      .period-overlay.visible{display:flex}
      .period-modal{width:min(560px,94vw);border-radius:30px;padding:28px;background:linear-gradient(145deg,rgba(10,20,38,.98),rgba(15,31,63,.98));border:1px solid rgba(106,151,255,.22);box-shadow:0 34px 90px rgba(0,0,0,.42);color:#eff6ff}
      .period-modal h3{margin:6px 0 10px;font-size:30px;line-height:1.05;letter-spacing:-.03em}
      .period-modal p{margin:0 0 18px;color:#afbdd8;font-size:15px}
      .period-list{display:grid;grid-template-columns:1fr;gap:10px;max-height:52vh;overflow:auto}
      .period-option{width:100%;text-align:left;padding:16px 18px;border-radius:18px;border:1px solid rgba(110,150,255,.16);background:rgba(255,255,255,.03);color:#eff6ff;display:flex;align-items:center;justify-content:space-between;gap:16px;cursor:pointer;transition:.2s ease}
      .period-option:hover{border-color:rgba(68,221,255,.55);transform:translateY(-1px)}
      .period-option.active{background:linear-gradient(135deg,rgba(16,39,86,.9),rgba(11,28,66,.9));border-color:rgba(68,221,255,.66);box-shadow:0 0 0 1px rgba(68,221,255,.14) inset}
      .period-option strong{display:block;font-size:17px}
      .period-option small{display:block;color:#9fb3d3;font-size:12px;margin-top:4px}
      .period-option .period-badge{font-size:12px;color:#7ee7ff;background:rgba(126,231,255,.1);border:1px solid rgba(126,231,255,.2);padding:5px 10px;border-radius:999px;white-space:nowrap}
      .period-actions{display:flex;justify-content:flex-end;margin-top:18px}
      .period-close-btn{border:none;border-radius:999px;padding:10px 16px;background:rgba(255,255,255,.08);color:#eff6ff;font-weight:700;cursor:pointer}
      @media(max-width:760px){.period-modal{padding:22px}.period-modal h3{font-size:24px}.period-option{padding:14px 15px;align-items:flex-start;flex-direction:column}.period-option .period-badge{align-self:flex-start}}
    `;
    document.head.appendChild(style);
  }

  const actions = $("#op-02 .ops-btn-row");
  if (actions && !document.getElementById("openOperationalPeriodPicker")) {
    const button = document.createElement("button");
    button.className = "ops-btn";
    button.id = "openOperationalPeriodPicker";
    button.type = "button";
    button.textContent = "Revisar periodos anteriores";
    button.addEventListener("click", openOperationalPeriodPicker);
    actions.appendChild(button);
  }

  if (!document.getElementById("operationalPeriodOverlay")) {
    const overlay = document.createElement("div");
    overlay.id = "operationalPeriodOverlay";
    overlay.className = "period-overlay";
    overlay.setAttribute("aria-hidden", "true");
    overlay.innerHTML = `
      <div class="period-modal">
        <div class="eyebrow">Pipeline por periodo</div>
        <h3>Revisar periodos anteriores</h3>
        <p>Selecciona el mes que quieres visualizar en la diapositiva de estructura por estatus.</p>
        <div class="period-list" id="operationalPeriodList"></div>
        <div class="period-actions">
          <button class="period-close-btn" id="closeOperationalPeriodPicker" type="button">Cerrar</button>
        </div>
      </div>
    `;
    overlay.addEventListener("click", event => {
      if (event.target === overlay) closeOperationalPeriodPicker();
    });
    document.body.appendChild(overlay);
    document.getElementById("closeOperationalPeriodPicker")?.addEventListener("click", closeOperationalPeriodPicker);
  }
}

function closeOperationalPeriodPicker() {
  const overlay = document.getElementById("operationalPeriodOverlay");
  if (!overlay) return;
  overlay.classList.remove("visible");
  overlay.setAttribute("aria-hidden", "true");
}

function openOperationalPeriodPicker() {
  const data = app.currentOperationalPayload;
  if (!data) return;
  ensureOperationalPeriodUI();
  const overlay = document.getElementById("operationalPeriodOverlay");
  const list = document.getElementById("operationalPeriodList");
  if (!overlay || !list) return;

  const periods = data.periodSummaries?.length
    ? data.periodSummaries
    : [{ key: data.latestPeriodKey || `${data.periodYear}-${String((data.periodMonth || 0) + 1).padStart(2, "0")}`, monthIndex: data.periodMonth, year: data.periodYear, label: `${MONTHS[data.periodMonth]} ${data.periodYear}`, stages: data.stages, views: data.views }];

  list.innerHTML = periods.map(period => {
    const totalVisible = ["Viabilidad", "Integración", "Análisis", "Autorización", "Formalización", "Dispersión"]
      .reduce((sum, stage) => sum + Number(period.stages?.[stage]?.count || 0), 0);
    const isActive = period.key === app.activeOperationalPeriodKey;
    const isCurrent = period.key === data.latestPeriodKey;
    const monthLabel = `${(MONTHS[period.monthIndex] || "PERIODO").charAt(0)}${(MONTHS[period.monthIndex] || "").slice(1).toLowerCase()} ${period.year}`;
    return `
      <button class="period-option ${isActive ? "active" : ""}" data-period-key="${period.key}" type="button">
        <div>
          <strong>${monthLabel}</strong>
          <small>${formatNumber(totalVisible)} operaciones visibles en la estructura.</small>
        </div>
        <span class="period-badge">${isCurrent ? "Periodo actual" : period.label}</span>
      </button>
    `;
  }).join("");

  list.querySelectorAll(".period-option").forEach(button => {
    button.addEventListener("click", () => {
      applyOperationalPeriodToSlide(data, button.dataset.periodKey);
      closeOperationalPeriodPicker();
    });
  });

  overlay.classList.add("visible");
  overlay.setAttribute("aria-hidden", "false");
}

function applyOperationalPeriodToSlide(data, periodKey = null) {
  const stageList = ["Viabilidad", "Integración", "Análisis", "Autorización", "Formalización", "Dispersión"];
  const fallbackKey = periodKey || data.latestPeriodKey || `${data.periodYear}-${String((data.periodMonth || 0) + 1).padStart(2, "0")}`;
  const selected = data.periodSummaries?.find(period => period.key === fallbackKey)
    || data.periodSummaries?.[0]
    || {
      key: fallbackKey,
      monthIndex: data.periodMonth,
      year: data.periodYear,
      label: `${MONTHS[data.periodMonth] || "PERIODO"} ${data.periodYear || ""}`,
      stages: data.stages,
      integrationBlockers: data.integrationBlockers,
      views: data.views
    };

  const maxCount = Math.max(...stageList.map(name => selected.stages?.[name]?.count || 0), 1);
  $$("#op-02 .status-card").forEach(card => {
    const name = $(".status-name", card)?.textContent.trim();
    const stage = selected.stages?.[name];
    if (!stage) return;
    $(".status-count", card).textContent = formatNumber(stage.count);
    $(".status-money", card).textContent = formatMoney(stage.requested);
    const fill = $(".fill", card);
    if (fill) fill.style.setProperty("--w", `${(stage.count / maxCount * 100).toFixed(1)}%`);
  });

  replaceSectionContent(
    "op-02",
    "Participación del pipeline",
    buildDonut(stageList.map(name => ({ name, value: selected.stages?.[name]?.count || 0 })), formatNumber(sumBy(stageList, name => selected.stages?.[name]?.count || 0)), "Operaciones visibles", 280)
  );

  const periodChip = $("#op-02 .top-chip");
  if (periodChip) periodChip.textContent = selected.label;

  app.activeOperationalPeriodKey = selected.key;
  app.currentOperationalPeriod = selected;
  ["viabilidad", "integracion", "analisis", "autorizacion", "formalizacion", "dispersion"].forEach(key => {
    if (selected.views?.[key]) app.viewTables[key] = selected.views[key];
  });
  if (data.views?.proyeccion) app.viewTables.proyeccion = data.views.proyeccion;

  const activeOpsKey = document.querySelector(".ops-option.active")?.dataset.key;
  if (activeOpsKey && app.renderOpsTable && app.viewTables[activeOpsKey]) {
    app.renderOpsTable(activeOpsKey);
  }

  app.scaleRepeated?.();
}

function updateOperationalVisual(data) {
  const currentMonthName = MONTHS[data.periodMonth] || "PERIODO ACTUAL";
  const currentMonthTitle = currentMonthName.charAt(0) + currentMonthName.slice(1).toLowerCase();
  const dispersionTitle = $("#op-05 .slide-title");
  if (dispersionTitle) dispersionTitle.textContent = `Dispersiones del mes de ${currentMonthTitle} hasta el corte`;
  const dispersionChip = $("#op-05 .top-chip");
  if (dispersionChip) dispersionChip.textContent = `${formatNumber(data.projection.dispersions.length)} operaciones · corte actual`;
  const currentCompareLabel = $("#op-01 .compare-pct-box:nth-child(2) .eyebrow");
  if (currentCompareLabel) currentCompareLabel.textContent = currentMonthTitle;

  const currentTarget = getOperationalTargetForMonth(data.periodMonth, data.target);
  const currentMissing = Math.max(currentTarget - data.dispersed, 0);
  const currentProgress = currentTarget ? data.dispersed / currentTarget * 100 : 0;

  setMetric("op-01", "Meta mensual", formatMoney(currentTarget), formatGapText(currentTarget, data.dispersed));
  setMetric("op-01", "Meta actual", formatMoney(currentTarget), formatGapText(currentTarget, data.dispersed));
  setMetric("op-01", "Dispersión actual", formatMoney(data.dispersed), `${data.projection.dispersions.length} operaciones confirmadas.`);
  setMetric("op-01", "Faltante", formatMoney(currentMissing), `Restante para alcanzar la meta de ${currentMonthTitle}.`);
  setMetric("op-01", "Avance", formatPercent(currentProgress), `Cumplimiento frente a la meta actual.`);
  updateHistory(data);
  renderOperationalFutureTargets(data.periodMonth, data.dispersed);

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
    if (pct) pct.textContent = formatPercent(currentProgress);
  }

  app.currentOperationalPayload = data;
  ensureOperationalPeriodUI();
  applyOperationalPeriodToSlide(data, data.latestPeriodKey);

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

  renderFinancialCommentActivity(data.financialCommentActivity);

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



function renderHistoricalClosingsSlide(rows = HISTORICAL_MEMBERSHIP_CLOSINGS) {
  const section = document.getElementById("com-history");
  if (!section) return;
  const sourceRows = Array.isArray(rows) && rows.length ? rows : HISTORICAL_MEMBERSHIP_CLOSINGS;
  let dataRows = sourceRows
    .filter(row => Number(row.monthIndex) >= 2 && Number(row.monthIndex) <= 7)
    .slice()
    .sort((a, b) => Number(a.monthIndex || 0) - Number(b.monthIndex || 0));
  if (!dataRows.length) {
    dataRows = HISTORICAL_MEMBERSHIP_CLOSINGS
      .filter(row => Number(row.monthIndex) >= 2 && Number(row.monthIndex) <= 7)
      .slice()
      .sort((a, b) => Number(a.monthIndex || 0) - Number(b.monthIndex || 0));
  }

  const monthCountsMap = new Map();
  dataRows.forEach(row => {
    const key = `${Number(row.monthIndex || 0)}|${row.month || "Sin mes"}`;
    monthCountsMap.set(key, (monthCountsMap.get(key) || 0) + 1);
  });
  const monthEntries = [2, 3, 4, 5, 6, 7].map(monthIndex => {
    const month = MONTHS[monthIndex];
    const key = `${monthIndex}|${month}`;
    return { monthIndex, month, value: monthCountsMap.get(key) || 0 };
  });

  const today = new Date();
  const latestMonthIndex = today.getMonth();
  const previousMonthIndex = (latestMonthIndex + 11) % 12;
  const latestMonthName = MONTHS[latestMonthIndex] || "—";
  const previousMonthName = MONTHS[previousMonthIndex] || "—";
  const currentRows = dataRows.filter(row => Number(row.monthIndex || 0) === latestMonthIndex);
  const previousRows = dataRows.filter(row => Number(row.monthIndex || 0) === previousMonthIndex);
  const topMonth = monthEntries.reduce((best, item) => (item.value > (best?.value || 0) ? item : best), monthEntries[0] || { month: "—", value: 0 });

  const totalNode = $(".history-close-total", section);
  const monthNode = $(".history-close-months", section);
  const topMonthNode = $(".history-close-top-month", section);
  const topMonthSub = $(".history-close-top-month-sub", section);
  const rangeNode = $(".history-range-chip", section);
  const currentTitle = $(".history-current-title", section);
  const prevTitle = $(".history-prev-title", section);
  const currentCount = $(".history-current-count", section);
  const prevCount = $(".history-prev-count", section);
  const chartRange = $(".history-chart-range", section);

  if (totalNode) totalNode.textContent = formatNumber(dataRows.length);
  if (monthNode) monthNode.textContent = formatNumber(monthEntries.length);
  if (topMonthNode) topMonthNode.textContent = topMonth.month || "—";
  if (topMonthSub) topMonthSub.textContent = `${formatNumber(topMonth.value || 0)} cierres`;
  if (rangeNode && monthEntries.length) rangeNode.textContent = `${monthEntries[0].month} · ${monthEntries[monthEntries.length - 1].month} 2026`;
  if (chartRange && monthEntries.length) chartRange.textContent = `${monthEntries[0].month} a ${monthEntries[monthEntries.length - 1].month} 2026`;
  if (currentTitle) currentTitle.textContent = `Cierres de ${latestMonthName.toLowerCase()}`;
  if (prevTitle) prevTitle.textContent = `Cierres de ${previousMonthName.toLowerCase()}`;
  if (currentCount) currentCount.textContent = `${formatNumber(currentRows.length)} ${currentRows.length === 1 ? "registro" : "registros"}`;
  if (prevCount) prevCount.textContent = `${formatNumber(previousRows.length)} ${previousRows.length === 1 ? "registro" : "registros"}`;

  const renderList = target => rowsList => {
    const node = $(target, section);
    if (!node) return;
    if (!rowsList.length) {
      node.innerHTML = '<div class="history-empty">Sin cierres registrados en este mes.</div>';
      return;
    }
    node.innerHTML = rowsList.map(row => `
      <div class="history-close-item">
        <div>
          <div class="history-close-name">${escapeHtml(row.name || "—")}</div>
          <div class="history-close-meta">${escapeHtml(row.financial || row.region || "—")} · ${escapeHtml(row.broker || row.office || "—")} · ${escapeHtml(row.amountDisplay || "")}</div>
        </div>
        <span class="month-status-pill status-success">CERRADO</span>
      </div>
    `).join("");
  };
  renderList('.history-current-list')(currentRows);
  renderList('.history-prev-list')(previousRows);

  const chartNode = $(".history-bars", section);
  if (chartNode) {
    const maxValue = Math.max(...monthEntries.map(item => Number(item.value || 0)), 1);
    chartNode.innerHTML = monthEntries.map(item => {
      const height = Math.max(22, Math.round((Number(item.value || 0) / maxValue) * 220));
      return `
        <div class="history-bar-col">
          <div class="history-bar-value">${formatNumber(item.value || 0)}</div>
          <div class="history-bar-track">
            <div class="history-bar-fill" style="height:${height}px"></div>
          </div>
          <div class="history-bar-label">${escapeHtml(item.month)}</div>
        </div>
      `;
    }).join("");
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

function renderWeeklyActivitiesSlide(data) {
  const section = document.getElementById("com-activities");
  if (!section) return;
  const activities = data.weeklyActivities || [];
  const totalNode = $(".weekly-total", section);
  if (totalNode) totalNode.textContent = formatNumber(activities.length);

  const days = [...new Set(activities.map(item => item.day).filter(Boolean))];
  const daysNode = $(".weekly-days", section);
  if (daysNode) daysNode.textContent = formatNumber(days.length);

  const next = activities[0];
  const nextPerson = $(".weekly-next-person", section);
  const nextTime = $(".weekly-next-time", section);
  if (nextPerson) nextPerson.textContent = next?.person || "Sin actividades";
  if (nextTime) nextTime.textContent = next ? `${next.day || "—"} · ${next.time || "—"}` : "—";

  const timeline = $(".weekly-timeline", section);
  if (timeline) {
    timeline.innerHTML = activities.length ? activities.map(item => `
      <div class="weekly-item">
        <div class="weekly-date-pill"><strong>${escapeHtml(item.day || "—")}</strong><span>${escapeHtml(item.date || "")}</span></div>
        <div class="weekly-main">
          <div class="weekly-person">${escapeHtml(item.person)}</div>
          <div class="weekly-activity">${escapeHtml(item.activity)}</div>
        </div>
        <div class="weekly-time">${escapeHtml(item.time || "—")}</div>
      </div>
    `).join("") : '<div class="scope-empty">Sin actividades semanales cargadas.</div>';
  }

  const bars = $(".weekly-bars", section);
  if (bars) {
    const byDay = countBy(activities, row => row.day || "Sin día");
    bars.innerHTML = buildBars(entriesSorted(byDay), false);
  }
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
  const prospectTotal = data.focusedProspects?.length ?? 0;
  const diegoScope = data.directorScopes?.diego?.totalCount || 0;
  const jorgeScope = data.directorScopes?.jorge?.totalCount || 0;
  const open = prospectTotal;
  const currentFocusClosings = data.focusCurrentClosings?.length || 0;

  setMetric("com-01", "Seguimiento abierto", formatNumber(open), "Diego y Jorge activos.");
  setMetric("com-01", "Diego", formatNumber(diegoScope), "Cartera propia y referenciada.");
  setMetric("com-01", "Jorge", formatNumber(jorgeScope), "Cartera propia y referenciada.");
  setMetric("com-01", "Cierre prioritario", formatNumber(currentFocusClosings));

  replaceSectionContent(
    "com-01",
    "Estado comercial",
    buildDonut(
      ["Cierre", "Desarrollo", "Reactivación"].map(name => ({ name, value: data.focusedBuckets?.[name] || 0 })),
      formatNumber(prospectTotal),
      "Seguimiento abierto",
      260
    )
  );
  replaceSectionContent("com-01", "Directores con seguimiento abierto", buildBars(entriesSorted(data.focusDirectorsOpen || {}), false));

  renderDirectorScopeSlide(data);
  renderWeeklyActivitiesSlide(data);
  renderMonthClosingSlide("com-02", data.currentClosings || [], data.currentCloseMonthIndex ?? new Date().getMonth(), data.currentCloseYear ?? new Date().getFullYear());
  renderHistoricalClosingsSlide(data.closures2026?.length ? data.closures2026 : (data.historicalMembershipClosings || HISTORICAL_MEMBERSHIP_CLOSINGS));
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
  app.viewGroups.panorama_comercial = ["seguimiento_diego_jorge", "prioritarios", "diego_propias", "diego_referenciadas", "jorge_propias", "jorge_referenciadas"];
  app.viewGroups.alcance_directores = ["diego_propias", "diego_referenciadas", "jorge_propias", "jorge_referenciadas"];
  app.viewGroups.actividades_semanales = ["actividades_semanales"];
  app.viewGroups.historial_cierres = ["cierres_transcurridos"];
}

function applyPayload(payload, persist = true) {
  if (payload.type === "operational") {
    updateOperationalVisual(payload);
    // Los CIERRES 2026 operativos solo alimentan dispersión/histórico operativo.
    // Nunca deben reemplazar la diapositiva comercial de cierres de membresías.
  }
  if (payload.type === "commercial") {
    if (Array.isArray(payload.closures2026) && payload.closures2026.length) {
      app.closures2026 = payload.closures2026;
    }
    updateCommercialVisual(payload);
  }
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
      ["Histórico de dispersión", `${payload.historical.length} meses`],
      ["Cierres 2026", `${formatNumber(payload.closures2026?.length || 0)} operaciones`]
    );
  } else {
    items.push(
      ["Archivo", file.name],
      ["Base comercial filtrada", formatNumber(payload.prospects.length)],
      ["Fuera del alcance", formatNumber(payload.excludedCount || 0)],
      ["Seguimiento abierto Diego/Jorge", formatNumber(payload.focusedProspects?.length || 0)],
      ["Cerrados excluidos", formatNumber((payload.buckets["Pagado"] || 0) + (payload.buckets["No viable"] || 0))],
      ["Diego · propias / referidas", `${formatNumber(payload.directorScopes?.diego?.ownCount || 0)} / ${formatNumber(payload.directorScopes?.diego?.referredCount || 0)}`],
      ["Jorge · propias / referidas", `${formatNumber(payload.directorScopes?.jorge?.ownCount || 0)} / ${formatNumber(payload.directorScopes?.jorge?.referredCount || 0)}`],
      [`Cierres ${MONTH_LABELS[payload.currentCloseMonthIndex] || "actual"}`, formatNumber(payload.currentClosings?.length || 0)],
      [`Cierres ${MONTH_LABELS[payload.nextCloseMonthIndex] || "siguiente"}`, formatNumber(payload.nextClosings?.length || 0)],
      ["Actividades semanales", formatNumber(payload.weeklyActivities?.length || 0)]
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
