import React, { useState, useMemo, useRef, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import _ from "lodash";
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area, PieChart, Pie, Cell,
  ScatterChart, Scatter, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, LabelList,
} from "recharts";
import {
  LayoutGrid, Upload, FileSpreadsheet, BarChart3, FileText, Settings as SettingsIcon,
  Plus, ArrowRight, ArrowLeft, Search, ChevronUp, ChevronDown, Trash2, Copy, Eye,
  Download, Sparkles, AlertTriangle, CheckCircle2, X, Sun, Moon, Globe, Link2,
  Table2, TrendingUp, TrendingDown, Minus, GripVertical, ZoomIn, ZoomOut, Loader2,
  Building2, Palette, Type as TypeIcon, ChevronRight, Info, ShieldAlert, EyeOff,
} from "lucide-react";

/* ============================== DESIGN TOKENS ============================== */
const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap');`;

const GlobalStyle = () => (
  <style>{`
    ${FONT_IMPORT}
    .dv-root{
      --ink:#0B1220; --ink-2:#131C2E; --ink-3:#1B2740;
      --paper:#F5F6FA; --surface:#FFFFFF; --border:#E3E6EC; --border-2:#EEF0F4;
      --text:#101828; --text-2:#5B6472; --text-3:#96A0AF;
      --blue:#2F5FED; --blue-dim:#EAEFFE; --blue-ink:#1C3FA8;
      --teal:#0EA894; --teal-dim:#E4F8F4;
      --amber:#DD9B2E; --amber-dim:#FBF1DE;
      --rose:#DD5350; --rose-dim:#FCEAEA;
      --violet:#7C5CFA; --violet-dim:#EFEAFE;
      --sans:'Inter',system-ui,sans-serif; --serif:'Fraunces',serif; --mono:'IBM Plex Mono',monospace;
      font-family:var(--sans); color:var(--text); background:var(--paper);
    }
    .dv-root[data-theme="dark"]{
      --paper:#0D1320; --surface:#121A2B; --border:#232D42; --border-2:#1B2437;
      --text:#EDF0F5; --text-2:#9CA6B8; --text-3:#5F6B80;
      --blue-dim:#182448; --teal-dim:#0E2A28; --amber-dim:#2E2412; --rose-dim:#2E1919; --violet-dim:#211C3C;
    }
    .dv-root [dir="rtl"]{direction:rtl;}
    .dv-serif{font-family:var(--serif);}
    .dv-mono{font-family:var(--mono);}
    .dv-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;}
    .dv-btn{display:inline-flex;align-items:center;gap:8px;font-weight:600;font-size:13.5px;border-radius:10px;padding:10px 16px;transition:all .15s ease;cursor:pointer;border:1px solid transparent;white-space:nowrap;}
    .dv-btn:active{transform:translateY(1px);}
    .dv-btn-primary{background:var(--blue);color:#fff;}
    .dv-btn-primary:hover{background:var(--blue-ink);}
    .dv-btn-ghost{background:transparent;color:var(--text-2);border-color:var(--border);}
    .dv-btn-ghost:hover{background:var(--border-2);color:var(--text);}
    .dv-btn-dark{background:var(--ink);color:#fff;}
    .dv-btn-dark:hover{background:var(--ink-3);}
    .dv-btn-sm{padding:6px 11px;font-size:12.5px;border-radius:8px;}
    .dv-input{width:100%;background:var(--surface);border:1px solid var(--border);border-radius:9px;padding:9px 11px;font-size:13.5px;color:var(--text);font-family:var(--sans);}
    .dv-input:focus{outline:2px solid var(--blue);outline-offset:1px;}
    .dv-label{font-size:11.5px;font-weight:600;color:var(--text-2);text-transform:uppercase;letter-spacing:.04em;margin-bottom:6px;display:block;}
    .dv-sidebar-link{display:flex;align-items:center;gap:11px;padding:9px 12px;border-radius:9px;font-size:13.5px;font-weight:500;color:#9BA7C4;cursor:pointer;transition:all .12s;}
    .dv-sidebar-link:hover{background:rgba(255,255,255,.06);color:#fff;}
    .dv-sidebar-link.active{background:var(--blue);color:#fff;}
    .dv-pill{display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:700;padding:3px 9px;border-radius:99px;letter-spacing:.02em;}
    .dv-scrollbar::-webkit-scrollbar{width:8px;height:8px;}
    .dv-scrollbar::-webkit-scrollbar-thumb{background:var(--border);border-radius:8px;}
    .dv-fade-in{animation:dvFade .35s ease both;}
    @keyframes dvFade{from{opacity:0;transform:translateY(6px);}to{opacity:1;transform:translateY(0);}}
    @media (prefers-reduced-motion: reduce){ .dv-fade-in{animation:none;} }
    .dv-cell{width:16px;height:11px;border-radius:2px;background:#2A3A5C;}
    @keyframes cellRise{0%{transform:scaleY(.3);opacity:.4;}100%{transform:scaleY(1);opacity:1;}}
    .dv-report-page{background:#fff;color:#101828;box-shadow:0 1px 3px rgba(16,24,40,.08), 0 12px 32px rgba(16,24,40,.10);}
    .dv-split{display:grid;grid-template-columns:var(--split-w,320px) 1fr;min-height:0;}
    @media (max-width: 880px){
      .dv-split{grid-template-columns:1fr;grid-template-rows:auto 1fr;}
      .dv-split-side{max-height:46vh;border-right:none !important;border-bottom:1px solid var(--border);}
    }
    .dv-tpl-thumb{width:36px;height:26px;border-radius:4px;flex-shrink:0;overflow:hidden;position:relative;border:1px solid rgba(0,0,0,.06);}
    @media print{
      @page{ margin:0; }
      body *{visibility:hidden;}
      #dv-print-area, #dv-print-area *{visibility:visible;}
      #dv-print-area{position:absolute !important;left:0;top:0;width:auto !important;transform:none !important;}
      .dv-report-page{box-shadow:none !important;page-break-after:always;break-after:page;page-break-inside:avoid;break-inside:avoid;overflow:hidden;}
      /* Recharts keeps its hover tooltip cursor/active-dot in the DOM at all times, only toggling
         visibility:hidden while idle. Our blanket "make everything in the print area visible" rule
         above accidentally un-hides that resting cursor, which is what shows up as a stray little
         square near the first data point. Force those specific Recharts internals back off. */
      #dv-print-area .recharts-tooltip-wrapper,
      #dv-print-area .recharts-tooltip-cursor,
      #dv-print-area .recharts-active-dot,
      #dv-print-area .recharts-brush{visibility:hidden !important;display:none !important;}
    }
  `}</style>
);

/* ============================== DATA HELPERS ============================== */
const uid = () => Math.random().toString(36).slice(2, 10);

// Currencies available in Settings → Default currency. Symbols are shown before the amount
// everywhere a value is formatted as "currency" (KPI cards, chart axes/tooltips, report totals).
const CURRENCIES = [
  { code: "USD", symbol: "$", label: "US Dollar (USD)" },
  { code: "EUR", symbol: "€", label: "Euro (EUR)" },
  { code: "GBP", symbol: "£", label: "British Pound (GBP)" },
  { code: "SAR", symbol: "ر.س", label: "Saudi Riyal (SAR)" },
  { code: "EGP", symbol: "ج.م", label: "Egyptian Pound (EGP)" },
];
const currencySymbol = (code) => (CURRENCIES.find((c) => c.code === code) || CURRENCIES[0]).symbol;
// Read by formatValue() below. Kept as a simple module-level value (rather than threading a
// "currency" prop through every KPI/insight/chart call site) and kept in sync from Settings via
// a useEffect in the root App component.
let currentCurrency = "USD";
const setCurrentCurrency = (code) => { currentCurrency = code; };
let currentDateFormat = "MM/DD/YYYY";
const setCurrentDateFormat = (f) => { currentDateFormat = f; };
function formatDateBySetting(date) {
  const d = date instanceof Date ? date : new Date(date);
  if (isNaN(d)) return "";
  const pad = (n) => String(n).padStart(2, "0");
  const mm = pad(d.getMonth() + 1), dd = pad(d.getDate()), yyyy = d.getFullYear();
  if (currentDateFormat === "DD/MM/YYYY") return `${dd}/${mm}/${yyyy}`;
  if (currentDateFormat === "YYYY-MM-DD") return `${yyyy}-${mm}-${dd}`;
  return `${mm}/${dd}/${yyyy}`;
}

// Excel's date serial epoch (Dec 30 1899). Converts a raw serial number to a real Date.
function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  const d = new Date(utcDays * 86400 * 1000);
  const frac = serial - Math.floor(serial);
  if (frac > 0) d.setUTCSeconds(Math.round(frac * 86400));
  return d;
}
function toISODate(d) {
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
}
const DATE_NAME_HINT = /date|_dt$|created|updated|timestamp|\bdob\b/i;

// Spreadsheet exports (especially CSV, or Excel cells with no date format applied) often carry
// dates as bare serial numbers like 45123. detectColumnType alone can't tell those apart from a
// genuine numeric column, so this pass looks at the column NAME too: if it reads like a date field
// and every value falls in a plausible Excel-serial range, we treat it as a date up front.
function normalizeSerialDateColumns(rows, columns) {
  const serialCols = columns.filter((col) => {
    if (!DATE_NAME_HINT.test(col)) return false;
    const vals = rows.map((r) => r[col]).filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
    if (vals.length < 3) return false;
    return vals.every((v) => {
      const n = Number(String(v).trim());
      return !isNaN(n) && n > 20000 && n < 60000 && Number.isFinite(n);
    });
  });
  if (!serialCols.length) return rows;
  return rows.map((r) => {
    const next = { ...r };
    serialCols.forEach((col) => {
      const v = next[col];
      if (v !== null && v !== undefined && String(v).trim() !== "") {
        next[col] = toISODate(excelSerialToDate(Number(String(v).trim())));
      }
    });
    return next;
  });
}

function detectColumnType(values, columnName = "") {
  const nonEmpty = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
  if (nonEmpty.length === 0) return "text";
  const sample = nonEmpty.slice(0, 60).map((v) => String(v).trim());
  const boolSet = new Set(["true", "false", "yes", "no", "y", "n"]);
  if (sample.every((v) => boolSet.has(v.toLowerCase())) && new Set(sample.map((v) => v.toLowerCase())).size <= 2) return "boolean";
  if (sample.every((v) => /^-?[$€£]\s?[\d,]+(\.\d+)?$/.test(v))) return "currency";
  if (sample.every((v) => /^-?\d+(\.\d+)?\s?%$/.test(v))) return "percentage";
  const dateHits = sample.filter((v) => {
    if (/^\d{4}-\d{1,2}-\d{1,2}/.test(v) || /^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(v) || /^\d{1,2}-\d{1,2}-\d{2,4}/.test(v) || /^[A-Za-z]{3,9}\s\d{1,2},?\s\d{4}/.test(v)) {
      return !isNaN(Date.parse(v));
    }
    return false;
  });
  if (dateHits.length / sample.length > 0.75) return "date";
  if (sample.every((v) => v !== "" && !isNaN(Number(v.replace(/,/g, ""))))) return "number";
  return "text";
}

function parseNumeric(v, type) {
  if (v === null || v === undefined) return null;
  let s = String(v).trim();
  if (s === "") return null;
  if (type === "percentage") s = s.replace("%", "");
  s = s.replace(/[$€£,]/g, "");
  const n = Number(s);
  return isNaN(n) ? null : n;
}

// Personal/sensitive data detection — by column name (email, phone, customer name, address...)
// and, as a backstop, by sampling actual values against email/phone shapes so an oddly-named
// column ("Contact", "رقم") still gets caught. Flagged columns are excluded from KPIs, insights,
// chart fields, and report tables by default, and masked in the data preview.
const PII_NAME_PATTERN = /e[-\s]?mail|phone|mobile|whatsapp|tel(?:ephone)?|contact.?(no|number)?|customer.?name|client.?name|full.?name|^name$|address|ssn|national.?id|passport|\bid.?number\b/i;
function looksLikeEmail(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v).trim()); }
function looksLikePhone(v) { const s = String(v).trim().replace(/[\s\-()]/g, ""); return /^\+?\d{7,15}$/.test(s) && s.replace(/\D/g, "").length >= 7; }
function detectSensitiveColumn(colName, values) {
  if (PII_NAME_PATTERN.test(colName)) return true;
  const nonEmpty = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== "").slice(0, 25);
  if (!nonEmpty.length) return false;
  const emailHits = nonEmpty.filter(looksLikeEmail).length;
  if (emailHits / nonEmpty.length > 0.6) return true;
  const phoneHits = nonEmpty.filter(looksLikePhone).length;
  if (phoneHits / nonEmpty.length > 0.6) return true;
  return false;
}
function maskValue(v) {
  const s = String(v);
  if (looksLikeEmail(s)) { const [user, domain] = s.split("@"); return `${user.slice(0, 1)}•••@${domain}`; }
  if (looksLikePhone(s)) return s.replace(/\d(?=\d{2})/g, "•");
  return s.length <= 2 ? "••" : s[0] + "•".repeat(Math.min(s.length - 1, 6));
}

function buildSchema(rows, columns) {
  return columns.map((col) => {
    const values = rows.map((r) => r[col]);
    const type = detectColumnType(values, col);
    const nonEmpty = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
    const unique = new Set(nonEmpty.map((v) => String(v).trim()));
    const sensitive = detectSensitiveColumn(col, values);
    return { name: col, type, missing: values.length - nonEmpty.length, unique: unique.size, sensitive };
  });
}

function computeQuality(rows, schema) {
  const totalCells = rows.length * schema.length;
  const missing = schema.reduce((a, c) => a + c.missing, 0);
  const seen = new Set();
  let duplicates = 0;
  rows.forEach((r) => {
    const key = JSON.stringify(r);
    if (seen.has(key)) duplicates++;
    else seen.add(key);
  });
  return {
    totalRows: rows.length,
    totalCols: schema.length,
    missing,
    missingPct: totalCells ? Math.round((missing / totalCells) * 1000) / 10 : 0,
    duplicates,
    numericCols: schema.filter((c) => ["number", "currency", "percentage"].includes(c.type)).length,
    dateCols: schema.filter((c) => c.type === "date").length,
    textCols: schema.filter((c) => c.type === "text").length,
  };
}

function computeKpis(rows, schema) {
  const numericCols = schema.filter((c) => ["number", "currency", "percentage"].includes(c.type));
  const dateCol = schema.find((c) => c.type === "date");
  const kpis = [];
  numericCols.slice(0, 6).forEach((col) => {
    const vals = rows.map((r) => parseNumeric(r[col.name], col.type)).filter((v) => v !== null);
    if (!vals.length) return;
    const sum = _.sum(vals);
    const avg = sum / vals.length;
    kpis.push({ id: uid(), field: col.name, type: col.type, label: `Total ${col.name}`, value: sum, format: col.type, metric: "sum" });
    kpis.push({ id: uid(), field: col.name, type: col.type, label: `Average ${col.name}`, value: avg, format: col.type, metric: "avg" });
    if (dateCol) {
      const withDates = rows
        .map((r) => ({ d: Date.parse(r[dateCol.name]), v: parseNumeric(r[col.name], col.type) }))
        .filter((x) => !isNaN(x.d) && x.v !== null)
        .sort((a, b) => a.d - b.d);
      if (withDates.length >= 2) {
        const first = withDates[0].v, last = withDates[withDates.length - 1].v;
        if (first !== 0) {
          const growth = ((last - first) / Math.abs(first)) * 100;
          kpis.push({ id: uid(), field: col.name, type: col.type, label: `${col.name} Growth`, value: growth, format: "percentage", metric: "growth" });
        }
      }
    }
  });
  return kpis;
}

// Computes one KPI value on demand from a user-chosen {field, metric} spec — used by the
// report's KPI card editor so people can pick exactly which column/metric combo to show,
// not just whatever computeKpis() happened to auto-generate.
function computeSingleKpiValue(rows, schema, field, metric) {
  const col = schema.find((s) => s.name === field);
  if (!col) return null;
  if (metric === "growth") {
    const dateCol = schema.find((s) => s.type === "date");
    if (!dateCol) return null;
    const withDates = rows.map((r) => ({ d: Date.parse(r[dateCol.name]), v: parseNumeric(r[field], col.type) }))
      .filter((x) => !isNaN(x.d) && x.v !== null).sort((a, b) => a.d - b.d);
    if (withDates.length < 2) return null;
    const first = withDates[0].v, last = withDates[withDates.length - 1].v;
    if (first === 0) return null;
    return ((last - first) / Math.abs(first)) * 100;
  }
  const vals = rows.map((r) => parseNumeric(r[field], col.type)).filter((v) => v !== null);
  if (!vals.length) return null;
  if (metric === "avg") return _.mean(vals);
  if (metric === "count") return vals.length;
  if (metric === "min") return _.min(vals);
  if (metric === "max") return _.max(vals);
  return _.sum(vals);
}
function kpiSpecFormat(schema, field, metric) {
  if (metric === "growth") return "percentage";
  if (metric === "count") return "number";
  return (schema.find((s) => s.name === field) || {}).type || "number";
}

function pearson(xs, ys) {
  const n = xs.length;
  if (n < 3) return 0;
  const mx = _.mean(xs), my = _.mean(ys);
  let num = 0, dx = 0, dy = 0;
  for (let i = 0; i < n; i++) {
    num += (xs[i] - mx) * (ys[i] - my);
    dx += (xs[i] - mx) ** 2;
    dy += (ys[i] - my) ** 2;
  }
  const denom = Math.sqrt(dx * dy);
  return denom === 0 ? 0 : num / denom;
}

function computeInsights(rows, schema) {
  const insights = [];
  const numericCols = schema.filter((c) => ["number", "currency", "percentage"].includes(c.type));
  const dateCol = schema.find((c) => c.type === "date");
  const categoryCols = schema.filter((c) => c.type === "text" && c.unique >= 2 && c.unique <= 20);

  if (categoryCols.length && numericCols.length) {
    const cat = categoryCols[0], num = numericCols[0];
    const grouped = _.groupBy(rows, (r) => r[cat.name]);
    const sums = Object.entries(grouped).map(([k, rs]) => ({
      key: k, sum: _.sumBy(rs, (r) => parseNumeric(r[num.name], num.type) || 0),
    })).filter(x => x.key && x.key !== 'undefined');
    if (sums.length >= 2) {
      const sorted = _.orderBy(sums, "sum", "desc");
      insights.push({ type: "positive", text: `"${sorted[0].key}" leads in ${num.name}, contributing ${formatValue(sorted[0].sum, num.type)} — the top-performing ${cat.name.toLowerCase()}.` });
      const bottom = sorted[sorted.length - 1];
      insights.push({ type: "warning", text: `"${bottom.key}" has the lowest ${num.name} at ${formatValue(bottom.sum, num.type)}, ${sorted[0].sum !== 0 ? Math.round((1 - bottom.sum / sorted[0].sum) * 100) : 0}% below the leader.` });
    }
  }

  if (dateCol && numericCols.length) {
    const num = numericCols[0];
    const withDates = rows.map((r) => ({ d: Date.parse(r[dateCol.name]), v: parseNumeric(r[num.name], num.type) }))
      .filter((x) => !isNaN(x.d) && x.v !== null).sort((a, b) => a.d - b.d);
    if (withDates.length >= 2) {
      const first = withDates[0].v, last = withDates[withDates.length - 1].v;
      const pct = first !== 0 ? Math.round(((last - first) / Math.abs(first)) * 1000) / 10 : 0;
      insights.push({
        type: pct >= 0 ? "positive" : "negative",
        text: `${num.name} ${pct >= 0 ? "increased" : "decreased"} by ${Math.abs(pct)}% from the start to the end of the observed period.`,
      });
    }
  }

  if (numericCols.length >= 2) {
    for (let i = 0; i < numericCols.length - 1; i++) {
      for (let j = i + 1; j < numericCols.length; j++) {
        const a = numericCols[i], b = numericCols[j];
        const pairs = rows.map((r) => [parseNumeric(r[a.name], a.type), parseNumeric(r[b.name], b.type)]).filter((p) => p[0] !== null && p[1] !== null);
        if (pairs.length < 5) continue;
        const r = pearson(pairs.map((p) => p[0]), pairs.map((p) => p[1]));
        if (Math.abs(r) >= 0.5) {
          insights.push({ type: "info", text: `${a.name} and ${b.name} show a ${r > 0 ? "positive" : "negative"} correlation (r = ${r.toFixed(2)}).` });
        }
      }
    }
  }
  if (!insights.length) insights.push({ type: "info", text: "No strong patterns were detected in this dataset yet — try a dataset with a date or category column alongside numeric values for deeper insights." });
  return insights;
}

function formatValue(v, type) {
  if (v === null || v === undefined || isNaN(v)) return "—";
  if (type === "currency") return currencySymbol(currentCurrency) + Number(v).toLocaleString(undefined, { maximumFractionDigits: 0 });
  if (type === "percentage") return Number(v).toFixed(1) + "%";
  return Number(v).toLocaleString(undefined, { maximumFractionDigits: 1 });
}

function buildRecommendations(schema) {
  const dateCol = schema.find((c) => c.type === "date");
  const numericCols = schema.filter((c) => ["number", "currency", "percentage"].includes(c.type));
  const categoryCols = schema.filter((c) => c.type === "text" && c.unique >= 2 && c.unique <= 20);
  const recs = [];

  if (dateCol && numericCols[0]) {
    recs.push({ id: uid(), type: "line", title: `${numericCols[0].name} Over Time`, subtitle: `Trend of ${numericCols[0].name} across ${dateCol.name}`, explanation: "A time series is best read as a line — it makes direction and momentum immediately visible.", xField: dateCol.name, yField: numericCols[0].name, groupBy: "", aggregation: "sum" });
  }
  if (categoryCols[0] && numericCols[0]) {
    recs.push({ id: uid(), type: "bar", title: `${numericCols[0].name} by ${categoryCols[0].name}`, subtitle: `Comparison across ${categoryCols[0].unique} ${categoryCols[0].name.toLowerCase()} values`, explanation: "Bars make it easy to rank and compare discrete categories side by side.", xField: categoryCols[0].name, yField: numericCols[0].name, groupBy: "", aggregation: "sum" });
    if (categoryCols[0].unique <= 6) {
      recs.push({ id: uid(), type: "donut", title: `Share of ${numericCols[0].name} by ${categoryCols[0].name}`, subtitle: "Proportional breakdown", explanation: "With 6 or fewer segments, a donut chart communicates share-of-total at a glance.", xField: categoryCols[0].name, yField: numericCols[0].name, groupBy: "", aggregation: "sum" });
    }
  }
  if (numericCols.length >= 2) {
    recs.push({ id: uid(), type: "scatter", title: `${numericCols[0].name} vs ${numericCols[1].name}`, subtitle: "Relationship between two numeric variables", explanation: "A scatter plot reveals whether two numeric fields move together.", xField: numericCols[0].name, yField: numericCols[1].name, groupBy: "", aggregation: "sum" });
  }
  if (dateCol && categoryCols[0] && numericCols[0] && categoryCols[0].unique <= 6) {
    recs.push({ id: uid(), type: "stacked-bar", title: `${numericCols[0].name} by ${categoryCols[0].name} Over Time`, subtitle: "Multiple categories tracked across time", explanation: "Stacking by category over time shows both the total trend and each segment's contribution.", xField: dateCol.name, yField: numericCols[0].name, groupBy: categoryCols[0].name, aggregation: "sum" });
  }
  return recs.map((r) => ({
    ...r, filters: { dateFrom: "", dateTo: "", categories: [], numMin: "", numMax: "", topN: "", bottomN: "", compareYears: [] },
    appearance: { legend: true, dataLabels: false, gridlines: true, orientation: "vertical", numberFormat: "number", colorPalette: "classic", fontSize: 12 },
    addedToReport: false,
  }));
}

const CHART_COLORS = ["#2F5FED", "#0EA894", "#DD9B2E", "#7C5CFA", "#DD5350", "#2AA8D8", "#B98BF0", "#4CA37A"];

const COLOR_PALETTES = {
  classic: { label: "Classic", colors: ["#2F5FED", "#0EA894", "#DD9B2E", "#7C5CFA", "#DD5350", "#2AA8D8", "#B98BF0", "#4CA37A"] },
  vibrant: { label: "Vibrant", colors: ["#FF4D6D", "#FFB020", "#00C2A8", "#5B5FEF", "#FF7A45", "#22C55E", "#E040FB", "#00B4D8"] },
  pastel: { label: "Pastel", colors: ["#A7C7FF", "#B8E6D9", "#FFD9A0", "#D4C1F9", "#FFB6B9", "#9FE7F5", "#C9E4A5", "#F6C6EA"] },
  corporate: { label: "Corporate", colors: ["#1F3A5F", "#3D6A94", "#7A93AC", "#B0A99F", "#8C7B6B", "#5C7A6A", "#A3B18A", "#4A4E69"] },
  sunset: { label: "Sunset", colors: ["#F94144", "#F3722C", "#F8961E", "#F9C74F", "#90BE6D", "#43AA8B", "#577590", "#277DA1"] },
  mono: { label: "Monochrome Blue", colors: ["#0B2E6B", "#1D4BA0", "#2F5FED", "#5C82F0", "#89A4F4", "#B6C7F8", "#DDE5FC", "#3D6AC9"] },
};
const getPalette = (id) => (COLOR_PALETTES[id] || COLOR_PALETTES.classic).colors;

const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const QUARTERS = { Q1: [0, 2], Q2: [3, 5], Q3: [6, 8], Q4: [9, 11] };

function aggregateChart(rows, schema, cfg) {
  const xType = (schema.find((s) => s.name === cfg.xField) || {}).type;
  const yType = (schema.find((s) => s.name === cfg.yField) || {}).type || "number";

  // "Compare years" mode replaces the normal X-axis pivot with one series per selected year,
  // bucketed by month, so e.g. 2023 vs 2024 plot as two aligned lines/bars across Jan–Dec.
  if (xType === "date" && cfg.filters.compareYears?.length) {
    const years = cfg.filters.compareYears;
    const passesOtherFilters = (r) => {
      if (cfg.filters.categories?.length && !cfg.filters.categories.includes(String(r[cfg.xField]))) return false;
      if (cfg.filters.numMin !== "" && parseNumeric(r[cfg.yField], yType) < Number(cfg.filters.numMin)) return false;
      if (cfg.filters.numMax !== "" && parseNumeric(r[cfg.yField], yType) > Number(cfg.filters.numMax)) return false;
      return true;
    };
    const yearRows = rows.filter((r) => {
      const d = new Date(r[cfg.xField]);
      return !isNaN(d) && years.includes(String(d.getFullYear())) && passesOtherFilters(r);
    });
    const aggVals = (vals) => {
      const nums = vals.filter((v) => v !== null);
      if (!nums.length) return 0;
      if (cfg.aggregation === "avg") return _.mean(nums);
      if (cfg.aggregation === "count") return nums.length;
      if (cfg.aggregation === "min") return _.min(nums);
      if (cfg.aggregation === "max") return _.max(nums);
      return _.sum(nums);
    };
    const data = MONTH_NAMES.map((m, mi) => {
      const row = { name: m };
      years.forEach((y) => {
        const subset = yearRows.filter((r) => { const d = new Date(r[cfg.xField]); return d.getMonth() === mi && String(d.getFullYear()) === y; });
        row[y] = aggVals(subset.map((r) => parseNumeric(r[cfg.yField], yType)));
      });
      return row;
    });
    return { data, series: years };
  }

  let filtered = rows.filter((r) => {
    if (cfg.filters.categories?.length && !cfg.filters.categories.includes(String(r[cfg.xField]))) return false;
    if (xType === "date" && (cfg.filters.dateFrom || cfg.filters.dateTo)) {
      const d = Date.parse(r[cfg.xField]);
      if (cfg.filters.dateFrom && d < Date.parse(cfg.filters.dateFrom)) return false;
      if (cfg.filters.dateTo && d > Date.parse(cfg.filters.dateTo)) return false;
    }
    if (cfg.filters.numMin !== "" && parseNumeric(r[cfg.yField], yType) < Number(cfg.filters.numMin)) return false;
    if (cfg.filters.numMax !== "" && parseNumeric(r[cfg.yField], yType) > Number(cfg.filters.numMax)) return false;
    return true;
  });

  if (cfg.type === "scatter") {
    // A date field can't be parsed as a plain number — fall back to its timestamp so scatter
    // still plots something sensible if a date ever ends up as X here, instead of silently
    // dropping every point.
    const xVal = (r) => xType === "date" ? Date.parse(r[cfg.xField]) : parseNumeric(r[cfg.xField], xType);
    const data = filtered.map((r) => ({ x: xVal(r), y: parseNumeric(r[cfg.yField], yType) })).filter((p) => p.x !== null && !isNaN(p.x) && p.y !== null);
    return { data, series: [] };
  }

  const agg = (vals) => {
    const nums = vals.filter((v) => v !== null);
    if (!nums.length) return 0;
    if (cfg.aggregation === "avg") return _.mean(nums);
    if (cfg.aggregation === "count") return nums.length;
    if (cfg.aggregation === "min") return _.min(nums);
    if (cfg.aggregation === "max") return _.max(nums);
    return _.sum(nums);
  };

  let data, series = [];
  if (cfg.groupBy) {
    const byX = _.groupBy(filtered, (r) => r[cfg.xField]);
    const groupKeys = _.uniq(filtered.map((r) => String(r[cfg.groupBy]))).slice(0, 8);
    series = groupKeys;
    data = Object.entries(byX).map(([xKey, rs]) => {
      const row = { name: xKey };
      groupKeys.forEach((gk) => {
        const sub = rs.filter((r) => String(r[cfg.groupBy]) === gk);
        row[gk] = agg(sub.map((r) => parseNumeric(r[cfg.yField], yType)));
      });
      return row;
    });
    if (xType === "date") data = _.orderBy(data, (d) => Date.parse(d.name), "asc");
  } else {
    const byX = _.groupBy(filtered, (r) => r[cfg.xField]);
    data = Object.entries(byX).map(([xKey, rs]) => ({ name: xKey, value: agg(rs.map((r) => parseNumeric(r[cfg.yField], yType))) }));
    if (xType === "date") data = _.orderBy(data, (d) => Date.parse(d.name), "asc");
    else data = _.orderBy(data, "value", "desc");
  }

  if (cfg.filters.topN) data = _.orderBy(data, cfg.groupBy ? series[0] : "value", "desc").slice(0, Number(cfg.filters.topN));
  if (cfg.filters.bottomN) data = _.orderBy(data, cfg.groupBy ? series[0] : "value", "asc").slice(0, Number(cfg.filters.bottomN));

  return { data, series };
}

/* ============================== SMALL UI ATOMS ============================== */
const Badge = ({ children, tone = "blue" }) => (
  <span className="dv-pill" style={{ background: `var(--${tone}-dim)`, color: `var(--${tone === "ink" ? "text" : tone})` }}>{children}</span>
);

const IconBox = ({ icon: Icon, tone = "blue" }) => (
  <div style={{ width: 38, height: 38, borderRadius: 10, background: `var(--${tone}-dim)`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
    <Icon size={18} color={`var(--${tone})`} strokeWidth={2.2} />
  </div>
);

const EmptyState = ({ icon: Icon, title, subtitle, action }) => (
  <div className="dv-fade-in" style={{ textAlign: "center", padding: "56px 24px", color: "var(--text-2)" }}>
    <div style={{ width: 56, height: 56, borderRadius: 16, background: "var(--blue-dim)", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px" }}>
      <Icon size={24} color="var(--blue)" />
    </div>
    <div style={{ fontWeight: 700, fontSize: 15, color: "var(--text)", marginBottom: 6 }}>{title}</div>
    <div style={{ fontSize: 13.5, maxWidth: 360, margin: "0 auto 18px" }}>{subtitle}</div>
    {action}
  </div>
);

// Global confirmation dialog — used anywhere a destructive action (delete file / report / section)
// needs a deliberate second step instead of deleting on the first click.
function ConfirmDialog({ state, onCancel }) {
  if (!state) return null;
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,18,32,.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 20 }} onClick={onCancel}>
      <div className="dv-card dv-fade-in" style={{ padding: 22, maxWidth: 340, width: "100%" }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", gap: 10, alignItems: "flex-start", marginBottom: 16 }}>
          <div style={{ width: 34, height: 34, borderRadius: 9, background: "var(--rose-dim)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
            <AlertTriangle size={16} color="var(--rose)" />
          </div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 4 }}>{state.title}</div>
            <div style={{ fontSize: 12.5, color: "var(--text-2)", lineHeight: 1.5 }}>{state.message}</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
          <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={onCancel}>Cancel</button>
          <button className="dv-btn dv-btn-sm" style={{ background: "var(--rose)", color: "#fff" }} onClick={() => { state.onConfirm(); onCancel(); }}>Delete</button>
        </div>
      </div>
    </div>
  );
}

const Stepper = ({ steps, current }) => (
  <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 28, flexWrap: "wrap" }}>
    {steps.map((s, i) => (
      <React.Fragment key={s}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <div style={{
            width: 24, height: 24, borderRadius: 999, fontSize: 11.5, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center",
            background: i < current ? "var(--teal)" : i === current ? "var(--blue)" : "var(--border-2)",
            color: i <= current ? "#fff" : "var(--text-3)",
          }}>{i < current ? <CheckCircle2 size={13} /> : i + 1}</div>
          <span style={{ fontSize: 13, fontWeight: i === current ? 700 : 500, color: i === current ? "var(--text)" : "var(--text-3)" }}>{s}</span>
        </div>
        {i < steps.length - 1 && <div style={{ width: 24, height: 1, background: "var(--border)" }} />}
      </React.Fragment>
    ))}
  </div>
);

/* ============================== LANDING PAGE ============================== */
function LandingPage({ onStart }) {
  return (
    <div style={{ background: "var(--ink)", minHeight: "100%", color: "#fff" }} className="dv-fade-in">
      <div style={{ maxWidth: 1180, margin: "0 auto", padding: "26px 28px 0" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
            <div style={{ width: 30, height: 30, borderRadius: 8, background: "linear-gradient(135deg,#2F5FED,#7C5CFA)", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <BarChart3 size={16} color="#fff" />
            </div>
            <span className="dv-serif" style={{ fontSize: 18, fontWeight: 600 }}>DataVision AI</span>
          </div>
          <button className="dv-btn dv-btn-primary" onClick={onStart}>Start Analyzing <ArrowRight size={15} /></button>
        </div>
      </div>

      <div style={{ maxWidth: 980, margin: "0 auto", padding: "72px 28px 40px", textAlign: "center" }}>
        <div style={{ display: "inline-flex", alignItems: "center", gap: 7, background: "rgba(255,255,255,.07)", border: "1px solid rgba(255,255,255,.12)", borderRadius: 999, padding: "6px 13px", fontSize: 12, fontWeight: 600, color: "#B9C2DB", marginBottom: 26 }}>
          <Sparkles size={13} color="#7C5CFA" /> Spreadsheets in. Boardroom-ready reports out.
        </div>
        <h1 className="dv-serif" style={{ fontSize: "clamp(34px,5.6vw,58px)", lineHeight: 1.08, fontWeight: 600, letterSpacing: "-0.01em", marginBottom: 20 }}>
          Turn Your Spreadsheets Into<br />Powerful Reports
        </h1>
        <p style={{ fontSize: 17, color: "#9BA7C4", maxWidth: 560, margin: "0 auto 34px", lineHeight: 1.6 }}>
          Upload Excel or connect Google Sheets and automatically transform your data into charts, KPIs and professional PDF reports.
        </p>
        <div style={{ display: "flex", gap: 12, justifyContent: "center", flexWrap: "wrap", marginBottom: 64 }}>
          <button className="dv-btn dv-btn-primary" style={{ padding: "12px 22px", fontSize: 14.5 }} onClick={onStart}>Start Analyzing <ArrowRight size={16} /></button>
          <button className="dv-btn" style={{ padding: "12px 22px", fontSize: 14.5, background: "rgba(255,255,255,.06)", color: "#fff", border: "1px solid rgba(255,255,255,.15)" }} onClick={onStart}>
            <Upload size={15} /> Upload Your Data
          </button>
        </div>

        {/* Signature element: cells assembling into a chart */}
        <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "center", gap: 16, padding: "20px 0 8px" }}>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(6,16px)", gridAutoRows: 11, gap: 3 }}>
            {Array.from({ length: 24 }).map((_, i) => (
              <div key={i} className="dv-cell" style={{ background: i % 7 === 0 ? "#2F5FED" : i % 5 === 0 ? "#0EA894" : "#26314C", animation: `cellRise .5s ease ${(i % 6) * 0.05}s both` }} />
            ))}
          </div>
          <ArrowRight size={20} color="#5F6B80" />
          <div style={{ display: "flex", alignItems: "flex-end", gap: 6, height: 66 }}>
            {[28, 44, 34, 58, 40, 66].map((h, i) => (
              <div key={i} style={{ width: 14, height: h, borderRadius: "4px 4px 0 0", background: CHART_COLORS[i % 3 === 0 ? 0 : i % 3 === 1 ? 1 : 3], animation: `cellRise .5s ease ${0.3 + i * 0.06}s both`, transformOrigin: "bottom" }} />
            ))}
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1080, margin: "0 auto", padding: "0 28px 90px", display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))", gap: 16 }}>
        {[
          { icon: Sparkles, title: "AI Data Analysis", body: "Automatic column detection, KPI extraction and plain-language insights generated straight from your dataset.", tone: "#7C5CFA" },
          { icon: BarChart3, title: "Interactive Charts", body: "Chart types recommended based on your data's shape — then fully customizable with a live preview.", tone: "#2F5FED" },
          { icon: FileText, title: "Professional PDF Reports", body: "Drag together KPIs, charts and narrative into a branded, print-ready report in minutes.", tone: "#0EA894" },
        ].map((f) => (
          <div key={f.title} style={{ background: "rgba(255,255,255,.04)", border: "1px solid rgba(255,255,255,.10)", borderRadius: 16, padding: 24, textAlign: "left" }}>
            <div style={{ width: 40, height: 40, borderRadius: 11, background: "rgba(255,255,255,.08)", display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 16 }}>
              <f.icon size={19} color={f.tone} />
            </div>
            <div style={{ fontWeight: 700, fontSize: 15.5, marginBottom: 8 }}>{f.title}</div>
            <div style={{ fontSize: 13.5, color: "#9BA7C4", lineHeight: 1.55 }}>{f.body}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ============================== APP SHELL ============================== */
// Real (if partial) EN/AR translation dictionary — covers the persistent app chrome (sidebar,
// top bar, dashboard, settings). Deeper workflow screens (Chart Editor, Report Builder controls)
// aren't translated yet — that's a much larger follow-up, called out honestly in Settings below,
// rather than pretending the whole app is localized when only direction was flipping before.
const TRANSLATIONS = {
  nav: {
    dashboard: { en: "Dashboard", ar: "لوحة التحكم" },
    "new-analysis": { en: "New Analysis", ar: "تحليل جديد" },
    files: { en: "My Files", ar: "ملفاتي" },
    charts: { en: "Charts", ar: "الشارتات" },
    reports: { en: "Reports", ar: "التقارير" },
    settings: { en: "Settings", ar: "الإعدادات" },
  },
  topbar: {
    dashboard: { title: { en: "Dashboard", ar: "لوحة التحكم" }, sub: { en: "Your data analysis workspace", ar: "مساحة عمل تحليل بياناتك" } },
    "new-analysis": { title: { en: "New Analysis", ar: "تحليل جديد" }, sub: { en: "Import, preview and analyze your dataset", ar: "استورد بياناتك، عاينها، وحلّلها" } },
    charts: { title: { en: "Charts", ar: "الشارتات" }, sub: { en: "Recommended and customized visualizations", ar: "الرسومات المقترحة والمخصصة" } },
    files: { title: { en: "My Files", ar: "ملفاتي" }, sub: { en: "Datasets you've uploaded", ar: "مجموعات البيانات اللي رفعتها" } },
    reports: { title: { en: "My Reports", ar: "تقاريري" }, sub: { en: "Reports you've generated", ar: "التقارير اللي عملتها" } },
    "report-builder": { title: { en: "Report Builder", ar: "إنشاء التقرير" }, sub: { en: "Assemble your report", ar: "ابني تقريرك" } },
    "report-preview": { title: { en: "Report Preview", ar: "معاينة التقرير" }, sub: { en: "Exactly how your PDF will look", ar: "شكل ملف الـ PDF بالظبط" } },
    settings: { title: { en: "Settings", ar: "الإعدادات" }, sub: { en: "Preferences and defaults", ar: "التفضيلات والإعدادات الافتراضية" } },
  },
  dash: {
    welcome: { en: "Welcome back", ar: "أهلاً بيك" },
    subtitle: { en: "Here's what's happening with your data.", ar: "ده اللي بيحصل في بياناتك." },
    newAnalysis: { en: "New Analysis", ar: "تحليل جديد" },
    totalAnalyses: { en: "Total Analyses", ar: "إجمالي التحليلات" },
    totalReports: { en: "Total Reports", ar: "إجمالي التقارير" },
    filesUploaded: { en: "Files Uploaded", ar: "الملفات المرفوعة" },
    chartsCreated: { en: "Charts Created", ar: "الشارتات المُنشأة" },
    recentFiles: { en: "Recent Files", ar: "أحدث الملفات" },
    recentReports: { en: "Recent Reports", ar: "أحدث التقارير" },
    noFilesTitle: { en: "No files yet", ar: "مفيش ملفات لسه" },
    noFilesSub: { en: "Upload your first dataset to start analyzing your data.", ar: "ارفع أول ملف بيانات عشان تبدأ التحليل." },
    uploadData: { en: "Upload data", ar: "ارفع بيانات" },
    noReportsTitle: { en: "No reports yet", ar: "مفيش تقارير لسه" },
    noReportsSub: { en: "Create your first report from an analysis.", ar: "اعمل أول تقرير من تحليل." },
    analyzed: { en: "Analyzed", ar: "اتحلل" },
    pending: { en: "Pending", ar: "في الانتظار" },
  },
  settings: {
    company: { en: "Company", ar: "الشركة" },
    companyName: { en: "Company name", ar: "اسم الشركة" },
    defaultCurrency: { en: "Default currency", ar: "العملة الافتراضية" },
    defaultDateFormat: { en: "Default date format", ar: "صيغة التاريخ الافتراضية" },
    defaultTemplate: { en: "Default report template", ar: "قالب التقرير الافتراضي" },
    appearance: { en: "Appearance", ar: "المظهر" },
    light: { en: "Light", ar: "فاتح" },
    dark: { en: "Dark", ar: "غامق" },
    language: { en: "Language", ar: "اللغة" },
    languageNote: { en: "Core navigation and settings are translated. Deeper workflow screens (chart & report editors) are still English — full coverage is on the roadmap.", ar: "التنقل الأساسي والإعدادات متاحة بالعربي. شاشات العمل التفصيلية (محرر الشارتات والتقارير) لسه بالإنجليزي — هنكمّل الترجمة تباعًا." },
    save: { en: "Save Settings", ar: "حفظ الإعدادات" },
    saved: { en: "Saved ✓", ar: "اتحفظت ✓" },
    savedNote: { en: "Applies to new dates, currency labels, and this session's default report company/template. Nothing is persisted after a page reload yet.", ar: "بيتطبق على التواريخ الجديدة، رمز العملة، وشركة/قالب التقرير الافتراضي لهذه الجلسة. لسه مفيش حفظ دائم بعد تحديث الصفحة." },
  },
};
function t(lang, path) {
  const node = path.split(".").reduce((acc, k) => acc?.[k], TRANSLATIONS);
  return node?.[lang] || node?.en || path;
}

const NAV_ITEMS = [
  { id: "dashboard", label: "Dashboard", icon: LayoutGrid },
  { id: "new-analysis", label: "New Analysis", icon: Plus },
  { id: "files", label: "My Files", icon: FileSpreadsheet },
  { id: "charts", label: "Charts", icon: BarChart3 },
  { id: "reports", label: "Reports", icon: FileText },
  { id: "settings", label: "Settings", icon: SettingsIcon },
];

function Sidebar({ route, setRoute, lang }) {
  return (
    <div style={{ width: 224, background: "var(--ink)", flexShrink: 0, display: "flex", flexDirection: "column", padding: "18px 12px" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "6px 10px 22px" }}>
        <div style={{ width: 28, height: 28, borderRadius: 8, background: "linear-gradient(135deg,#2F5FED,#7C5CFA)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <BarChart3 size={15} color="#fff" />
        </div>
        <span className="dv-serif" style={{ color: "#fff", fontSize: 15.5, fontWeight: 600 }}>DataVision AI</span>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
        {NAV_ITEMS.map((it) => (
          <div key={it.id} className={`dv-sidebar-link ${route === it.id || (route === "chart-builder" && it.id === "charts") || (["report-builder","report-preview"].includes(route) && it.id === "reports") || (route.startsWith("new-analysis") && it.id === "new-analysis") ? "active" : ""}`} onClick={() => setRoute(it.id)}>
            <it.icon size={16} strokeWidth={2.2} /> {t(lang, `nav.${it.id}`)}
          </div>
        ))}
      </div>
      <div style={{ marginTop: "auto", padding: "14px 10px", borderTop: "1px solid rgba(255,255,255,.08)" }}>
        <div style={{ fontSize: 11, color: "#5F6B80" }}>MVP Prototype · v0.1</div>
      </div>
    </div>
  );
}

function TopBar({ title, subtitle, theme, setTheme, lang, setLang }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "16px 28px", borderBottom: "1px solid var(--border)", background: "var(--surface)" }}>
      <div>
        <div style={{ fontSize: 17, fontWeight: 700 }}>{title}</div>
        {subtitle && <div style={{ fontSize: 12.5, color: "var(--text-2)", marginTop: 1 }}>{subtitle}</div>}
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => setLang(lang === "en" ? "ar" : "en")} title="Toggle language (EN/AR)">
          <Globe size={14} /> {lang === "en" ? "EN" : "AR"}
        </button>
        <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => setTheme(theme === "light" ? "dark" : "light")} title="Toggle theme">
          {theme === "light" ? <Moon size={14} /> : <Sun size={14} />}
        </button>
      </div>
    </div>
  );
}

/* ============================== DASHBOARD ============================== */
function Dashboard({ files, reports, setRoute, lang }) {
  return (
    <div className="dv-fade-in" style={{ padding: 28, maxWidth: 1140, margin: "0 auto" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 22 }}>
        <div>
          <div style={{ fontSize: 20, fontWeight: 800 }}>{t(lang, "dash.welcome")}</div>
          <div style={{ fontSize: 13.5, color: "var(--text-2)" }}>{t(lang, "dash.subtitle")}</div>
        </div>
        <button className="dv-btn dv-btn-primary" onClick={() => setRoute("new-analysis")}><Plus size={15} /> {t(lang, "dash.newAnalysis")}</button>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 14, marginBottom: 24 }}>
        {[
          { label: t(lang, "dash.totalAnalyses"), value: files.filter(f => f.analyzed).length, icon: Sparkles, tone: "violet" },
          { label: t(lang, "dash.totalReports"), value: reports.length, icon: FileText, tone: "teal" },
          { label: t(lang, "dash.filesUploaded"), value: files.length, icon: FileSpreadsheet, tone: "blue" },
          { label: t(lang, "dash.chartsCreated"), value: files.reduce((a, f) => a + (f.chartCount || 0), 0), icon: BarChart3, tone: "amber" },
        ].map((k) => (
          <div key={k.label} className="dv-card" style={{ padding: 18 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 14 }}>
              <IconBox icon={k.icon} tone={k.tone} />
              <div style={{ fontSize: 12.5, color: "var(--text-2)", fontWeight: 600 }}>{k.label}</div>
            </div>
            <div className="dv-mono" style={{ fontSize: 26, fontWeight: 700 }}>{k.value}</div>
          </div>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.1fr .9fr", gap: 16 }}>
        <div className="dv-card" style={{ padding: 20 }}>
          <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 14 }}>{t(lang, "dash.recentFiles")}</div>
          {files.length === 0 ? (
            <EmptyState icon={FileSpreadsheet} title={t(lang, "dash.noFilesTitle")} subtitle={t(lang, "dash.noFilesSub")} action={<button className="dv-btn dv-btn-primary dv-btn-sm" onClick={() => setRoute("new-analysis")}>{t(lang, "dash.uploadData")}</button>} />
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {files.slice(0, 5).map((f) => (
                <div key={f.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "9px 0", borderBottom: "1px solid var(--border-2)" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <IconBox icon={FileSpreadsheet} tone="blue" />
                    <div>
                      <div style={{ fontSize: 13.5, fontWeight: 600 }}>{f.name}</div>
                      <div style={{ fontSize: 11.5, color: "var(--text-3)" }}>{f.rowCount} rows · {f.colCount} cols</div>
                    </div>
                  </div>
                  <Badge tone={f.analyzed ? "teal" : "amber"}>{f.analyzed ? t(lang, "dash.analyzed") : t(lang, "dash.pending")}</Badge>
                </div>
              ))}
            </div>
          )}
        </div>
        <div className="dv-card" style={{ padding: 20 }}>
          <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 14 }}>{t(lang, "dash.recentReports")}</div>
          {reports.length === 0 ? (
            <EmptyState icon={FileText} title={t(lang, "dash.noReportsTitle")} subtitle={t(lang, "dash.noReportsSub")} />
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {reports.slice(0, 5).map((r) => (
                <div key={r.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "9px 0", borderBottom: "1px solid var(--border-2)" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <IconBox icon={FileText} tone="teal" />
                    <div>
                      <div style={{ fontSize: 13.5, fontWeight: 600 }}>{r.title}</div>
                      <div style={{ fontSize: 11.5, color: "var(--text-3)" }}>{r.template} template</div>
                    </div>
                  </div>
                  <Badge tone="blue">{r.status}</Badge>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ============================== NEW ANALYSIS WIZARD ============================== */
function NewAnalysis({ dataset, setDataset, onFinish, setRoute, onAddToReport, onCustomizeChart }) {
  const [step, setStep] = useState(dataset ? (dataset.schema ? 2 : 1) : 0);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [sheetUrl, setSheetUrl] = useState("");
  const [search, setSearch] = useState("");
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState("asc");
  const [page, setPage] = useState(1);
  const [hiddenCols, setHiddenCols] = useState([]);
  const [revealedCols, setRevealedCols] = useState([]); // columns the user chose to un-mask in this preview only
  const [pendingWorkbook, setPendingWorkbook] = useState(null); // { wb, fileName } — set when an uploaded Excel file has more than one sheet
  const fileInputRef = useRef();
  const pageSize = 8;

  const steps = ["Import", "Preview & Quality", "Analysis", "Chart Recommendations"];

  const finishImport = (rawRows, columns, sheetName, fileName) => {
    if (!rawRows.length || !columns.length) {
      setError("This sheet appears to be empty. Please choose a sheet that contains data, or upload a different file.");
      setLoading(false);
      return;
    }
    const rows = normalizeSerialDateColumns(rawRows, columns);
    const schema = buildSchema(rows, columns);
    setDataset({ name: fileName, sheetName: sheetName || "Sheet1", rows, columns, schema, id: uid() });
    setPendingWorkbook(null);
    setLoading(false);
    setStep(1);
  };

  const loadSheet = (wb, sheetName, fileName) => {
    const ws = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false, dateNF: "yyyy-mm-dd" });
    const columns = json.length ? Object.keys(json[0]) : [];
    finishImport(json, columns, sheetName, fileName);
  };

  // Public Google Sheets can be read straight from the browser with no API key and no backend,
  // using Google's own "visualization query" endpoint in JSONP mode (a <script> tag load, not a
  // fetch()) — fetch() is blocked by CORS for this endpoint, but script-tag loading isn't, since
  // it's the exact mechanism Google Charts itself has used for years to embed live sheet data.
  // This only works for sheets shared as "Anyone with the link can view" — private sheets need
  // real OAuth, which is out of scope for a no-backend prototype.
  const connectGoogleSheet = (url) => {
    setError(""); setLoading(true);
    const idMatch = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!idMatch) {
      setError("That doesn't look like a Google Sheets URL. Copy the link from your browser's address bar while viewing the sheet.");
      setLoading(false);
      return;
    }
    const sheetId = idMatch[1];
    const gidMatch = url.match(/[#&?]gid=(\d+)/);
    const gid = gidMatch ? gidMatch[1] : "0";
    const callbackName = `__dvGSheetCb_${Date.now()}`;
    const cleanup = (script) => { delete window[callbackName]; script.remove(); };
    const timer = setTimeout(() => {
      if (window[callbackName]) {
        setError("Couldn't load this sheet. Make sure sharing is set to \"Anyone with the link can view\", then try again.");
        setLoading(false);
        delete window[callbackName];
      }
    }, 10000);

    window[callbackName] = (resp) => {
      clearTimeout(timer);
      try {
        const table = resp.table;
        if (!table || !table.cols || !table.cols.length) throw new Error("empty");
        const columns = table.cols.map((c, i) => (c.label && c.label.trim()) || (c.id && c.id.trim()) || `Column ${i + 1}`);
        const rows = table.rows.map((r) => {
          const obj = {};
          columns.forEach((name, i) => {
            const cell = r.c && r.c[i];
            let val = "";
            if (cell) {
              // Prefer the formatted string ("f") over the raw value ("v") — dates and currency
              // come through pre-formatted from Sheets, same as we do for Excel with raw:false.
              if (cell.f !== undefined && cell.f !== null) val = cell.f;
              else if (cell.v !== undefined && cell.v !== null) val = cell.v;
            }
            obj[name] = val;
          });
          return obj;
        });
        finishImport(rows, columns, `Sheet (gid ${gid})`, `Google Sheet — ${sheetId.slice(0, 8)}…`);
        cleanup(script);
      } catch (e) {
        setError("This sheet loaded but appears to be empty, or its sharing setting blocks reading it publicly.");
        setLoading(false);
        cleanup(script);
      }
    };

    const script = document.createElement("script");
    // tqx sub-parameters are colon-separated key:value pairs joined by semicolons — using "="
    // here (an easy mistake) makes Google silently ignore our callback name and fall back to its
    // own default handler, which doesn't exist on our page, so nothing ever fires.
    const tqx = encodeURIComponent(`out:json;responseHandler:${callbackName}`);
    script.src = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=${tqx}&gid=${gid}`;
    script.onerror = () => {
      clearTimeout(timer);
      setError("Couldn't reach this sheet. Double-check the link and that it's shared as \"Anyone with the link can view.\"");
      setLoading(false);
      cleanup(script);
    };
    document.body.appendChild(script);
  };

  const parseFile = (file) => {
    setError(""); setLoading(true); setPendingWorkbook(null);
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["csv", "xlsx", "xls"].includes(ext)) {
      setError("Unsupported format. Please upload a .csv, .xlsx or .xls file.");
      setLoading(false);
      return;
    }
    if (ext === "csv") {
      Papa.parse(file, {
        header: true, skipEmptyLines: true,
        complete: (res) => {
          const columns = res.meta.fields || [];
          finishImport(res.data, columns, file.name.replace(/\.csv$/i, ""), file.name);
        },
        error: () => { setError("Failed to read this CSV file. Please check the format and try again."); setLoading(false); },
      });
    } else {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
          if (wb.SheetNames.length > 1) {
            // Multiple sheets — let the person pick which one to analyze instead of silently
            // reading just the first sheet.
            setPendingWorkbook({ wb, fileName: file.name });
            setLoading(false);
            return;
          }
          // raw:false renders each cell using its Excel number format, so dates/currency/percentages
          // come through as readable strings ("8/29/2026", "$1,234", "12%") instead of raw serials —
          // which is what our column-type detector expects.
          loadSheet(wb, wb.SheetNames[0], file.name);
        } catch (err) {
          setError("We couldn't parse this Excel file. It may be corrupted or in an unsupported format.");
          setLoading(false);
        }
      };
      reader.onerror = () => { setError("Failed to read this file."); setLoading(false); };
      reader.readAsArrayBuffer(file);
    }
  };

  const quality = useMemo(() => dataset ? computeQuality(dataset.rows, dataset.schema) : null, [dataset]);

  const filteredRows = useMemo(() => {
    if (!dataset) return [];
    let rows = dataset.rows;
    if (search) {
      const s = search.toLowerCase();
      rows = rows.filter((r) => dataset.columns.some((c) => String(r[c]).toLowerCase().includes(s)));
    }
    if (sortCol) {
      rows = _.orderBy(rows, (r) => {
        const v = r[sortCol];
        const n = Number(String(v).replace(/[$€%,]/g, ""));
        return isNaN(n) ? v : n;
      }, sortDir);
    }
    return rows;
  }, [dataset, search, sortCol, sortDir]);

  const pagedRows = filteredRows.slice((page - 1) * pageSize, page * pageSize);
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / pageSize));

  // Sensitive columns (email, phone, customer name...) are excluded from every downstream
  // calculation — KPIs, insights, and chart recommendations only ever see the safe subset.
  const safeSchema = useMemo(() => dataset ? dataset.schema.filter((s) => !s.sensitive) : [], [dataset]);
  const kpis = useMemo(() => dataset ? computeKpis(dataset.rows, safeSchema) : [], [dataset, safeSchema]);
  const insights = useMemo(() => dataset ? computeInsights(dataset.rows, safeSchema) : [], [dataset, safeSchema]);
  const recommendations = useMemo(() => dataset ? buildRecommendations(safeSchema) : [], [dataset, safeSchema]);

  const runAnalyze = () => {
    setLoading(true);
    setTimeout(() => { setStep(2); setLoading(false); }, 550);
  };

  const goToCharts = () => {
    onFinish({ kpis, insights, recommendations });
    setStep(3);
  };

  return (
    <div className="dv-fade-in" style={{ padding: 28, maxWidth: 1080, margin: "0 auto" }}>
      <Stepper steps={steps} current={step} />

      {step === 0 && (
        <div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 20 }}>
            <div className="dv-card" style={{ padding: 26, textAlign: "center", cursor: "pointer" }} onClick={() => fileInputRef.current?.click()}
              onDragOver={(e) => e.preventDefault()} onDrop={(e) => { e.preventDefault(); if (e.dataTransfer.files[0]) parseFile(e.dataTransfer.files[0]); }}>
              <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={(e) => e.target.files[0] && parseFile(e.target.files[0])} />
              <div style={{ width: 52, height: 52, borderRadius: 14, background: "var(--blue-dim)", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
                <Upload size={22} color="var(--blue)" />
              </div>
              <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 5 }}>Upload a file</div>
              <div style={{ fontSize: 12.5, color: "var(--text-2)", marginBottom: 14 }}>Drag & drop or click to browse</div>
              <div style={{ display: "flex", gap: 6, justifyContent: "center" }}>
                {[".xlsx", ".xls", ".csv"].map((f) => <Badge key={f} tone="blue">{f}</Badge>)}
              </div>
            </div>
            <div className="dv-card" style={{ padding: 26 }}>
              <div style={{ width: 52, height: 52, borderRadius: 14, background: "var(--teal-dim)", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
                <Link2 size={22} color="var(--teal)" />
              </div>
              <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 5, textAlign: "center" }}>Connect Google Sheets</div>
              <div style={{ fontSize: 12.5, color: "var(--text-2)", marginBottom: 14, textAlign: "center" }}>Paste a shared Google Sheets URL</div>
              <input className="dv-input" placeholder="https://docs.google.com/spreadsheets/d/..." value={sheetUrl} onChange={(e) => setSheetUrl(e.target.value)} style={{ marginBottom: 10 }} />
              <button className="dv-btn dv-btn-dark" style={{ width: "100%", justifyContent: "center" }} onClick={() => sheetUrl.trim() && connectGoogleSheet(sheetUrl.trim())} disabled={!sheetUrl.trim() || loading}>
                {loading ? <><Loader2 size={14} style={{ animation: "spin 1s linear infinite" }} /> Connecting…</> : "Connect Google Sheet"}
              </button>
              <div style={{ fontSize: 10.5, color: "var(--text-3)", marginTop: 8, textAlign: "center" }}>Sharing must be set to "Anyone with the link can view". To pick a specific tab, open that tab in Sheets first, then copy its URL.</div>
            </div>
          </div>
          {pendingWorkbook && (
            <div className="dv-card dv-fade-in" style={{ padding: 18, marginBottom: 14 }}>
              <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 4 }}>This file has {pendingWorkbook.wb.SheetNames.length} sheets</div>
              <div style={{ fontSize: 12.5, color: "var(--text-2)", marginBottom: 12 }}>Choose which one to analyze — you can always upload again to pick a different sheet.</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {pendingWorkbook.wb.SheetNames.map((name) => (
                  <button key={name} className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => { setLoading(true); loadSheet(pendingWorkbook.wb, name, pendingWorkbook.fileName); }}>
                    <FileSpreadsheet size={13} /> {name}
                  </button>
                ))}
              </div>
            </div>
          )}
          {error && (
            <div className="dv-card" style={{ padding: 14, borderColor: "var(--rose)", background: "var(--rose-dim)", display: "flex", gap: 10, fontSize: 13, color: "var(--text)" }}>
              <AlertTriangle size={16} color="var(--rose)" style={{ flexShrink: 0, marginTop: 1 }} /> {error}
            </div>
          )}
          {loading && <div style={{ display: "flex", alignItems: "center", gap: 8, color: "var(--text-2)", fontSize: 13, marginTop: 12 }}><Loader2 size={15} className="dv-spin" style={{ animation: "spin 1s linear infinite" }} /> Reading your file…</div>}
        </div>
      )}

      {step === 1 && dataset && (
        <div>
          <div className="dv-card" style={{ padding: 20, marginBottom: 16 }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14, marginBottom: 4 }}>
              {[["File name", dataset.name], ["Sheet name", dataset.sheetName], ["Rows", dataset.rows.length], ["Columns", dataset.columns.length]].map(([l, v]) => (
                <div key={l}><div className="dv-label">{l}</div><div className="dv-mono" style={{ fontSize: 14, fontWeight: 600 }}>{v}</div></div>
              ))}
            </div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(6,1fr)", gap: 10, marginBottom: 16 }}>
            {[
              ["Rows", quality.totalRows, "blue"], ["Columns", quality.totalCols, "blue"],
              ["Missing", `${quality.missingPct}%`, quality.missingPct > 5 ? "rose" : "teal"],
              ["Duplicates", quality.duplicates, quality.duplicates > 0 ? "amber" : "teal"],
              ["Numeric cols", quality.numericCols, "violet"], ["Date cols", quality.dateCols, "violet"],
            ].map(([l, v, tone]) => (
              <div key={l} className="dv-card" style={{ padding: 12 }}>
                <div style={{ fontSize: 10.5, color: "var(--text-2)", fontWeight: 600, marginBottom: 4 }}>{l}</div>
                <div className="dv-mono" style={{ fontSize: 17, fontWeight: 700, color: `var(--${tone})` }}>{v}</div>
              </div>
            ))}
          </div>

          {(quality.missing > 0 || quality.duplicates > 0) && (
            <div className="dv-card" style={{ padding: 14, marginBottom: 16, background: "var(--amber-dim)", borderColor: "var(--amber)" }}>
              <div style={{ display: "flex", gap: 8, alignItems: "flex-start" }}>
                <AlertTriangle size={16} color="var(--amber)" style={{ flexShrink: 0, marginTop: 1 }} />
                <div style={{ fontSize: 13 }}>
                  <b>Data quality notes:</b> {quality.missing > 0 && `${quality.missing} missing value${quality.missing === 1 ? "" : "s"} detected. `}
                  {quality.duplicates > 0 && `${quality.duplicates} duplicate row${quality.duplicates === 1 ? "" : "s"} found.`}
                </div>
              </div>
            </div>
          )}

          {dataset.schema.some((s) => s.sensitive) && (
            <div className="dv-card" style={{ padding: 14, marginBottom: 16, background: "var(--rose-dim)", borderColor: "var(--rose)" }}>
              <div style={{ display: "flex", gap: 8, alignItems: "flex-start" }}>
                <Info size={16} color="var(--rose)" style={{ flexShrink: 0, marginTop: 1 }} />
                <div style={{ fontSize: 13 }}>
                  <b>Personal data detected and protected:</b> {dataset.schema.filter((s) => s.sensitive).map((s) => s.name).join(", ")}. These columns are masked below and automatically left out of KPIs, insights, charts, and report tables. Click the shield icon next to a column name if this was flagged by mistake.
                </div>
              </div>
            </div>
          )}

          <div className="dv-card" style={{ padding: 16, marginBottom: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 10 }}>
              <div style={{ position: "relative" }}>
                <Search size={14} color="var(--text-3)" style={{ position: "absolute", left: 10, top: 10 }} />
                <input className="dv-input" style={{ paddingLeft: 30, width: 220 }} placeholder="Search rows…" value={search} onChange={(e) => { setSearch(e.target.value); setPage(1); }} />
              </div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {dataset.columns.map((c) => (
                  <button key={c} className="dv-btn dv-btn-sm" style={{ background: hiddenCols.includes(c) ? "transparent" : "var(--blue-dim)", color: hiddenCols.includes(c) ? "var(--text-3)" : "var(--blue)", border: "1px solid var(--border)" }}
                    onClick={() => setHiddenCols((h) => h.includes(c) ? h.filter((x) => x !== c) : [...h, c])}>{c}</button>
                ))}
              </div>
            </div>
            <div style={{ fontSize: 11, color: "var(--text-3)", marginBottom: 8, display: "flex", alignItems: "center", gap: 5 }}>
              <Info size={11} /> Wrong type detected? Change the dropdown next to any column name below.
            </div>
            <div className="dv-scrollbar" style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                <thead>
                  <tr>
                    {dataset.columns.filter((c) => !hiddenCols.includes(c)).map((c) => {
                      const colSchema = dataset.schema.find((s) => s.name === c);
                      const type = colSchema?.type;
                      const isSensitive = colSchema?.sensitive;
                      return (
                        <th key={c} onClick={() => { setSortCol(c); setSortDir(sortCol === c && sortDir === "asc" ? "desc" : "asc"); }}
                          style={{ textAlign: "left", padding: "8px 10px", borderBottom: "1px solid var(--border)", cursor: "pointer", whiteSpace: "nowrap", color: "var(--text-2)", fontWeight: 700 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                            {c} {sortCol === c && (sortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
                            <select value={type} onClick={(e) => e.stopPropagation()} onChange={(e) => {
                              const newType = e.target.value;
                              setDataset((d) => ({ ...d, schema: d.schema.map((s) => s.name === c ? { ...s, type: newType } : s) }));
                            }} style={{ fontWeight: 500, color: "var(--text-2)", fontSize: 10, border: "1px solid var(--border)", borderRadius: 5, background: "var(--surface)", padding: "1px 3px" }} title="Override detected type">
                              {["date", "number", "currency", "percentage", "text", "boolean"].map((t) => <option key={t} value={t}>{t}</option>)}
                            </select>
                            <button title={isSensitive ? "Marked as personal data — click to unmark" : "Mark as personal data (excludes from analysis)"}
                              onClick={(e) => { e.stopPropagation(); setDataset((d) => ({ ...d, schema: d.schema.map((s) => s.name === c ? { ...s, sensitive: !s.sensitive } : s) })); }}
                              style={{ border: "none", background: "none", padding: 0, cursor: "pointer", display: "flex", opacity: isSensitive ? 1 : 0.35 }}>
                              <ShieldAlert size={12} color={isSensitive ? "var(--rose)" : "var(--text-3)"} />
                            </button>
                            {isSensitive && (
                              <button title={revealedCols.includes(c) ? "Hide values again" : "Reveal values (preview only)"} onClick={(e) => { e.stopPropagation(); setRevealedCols((r) => r.includes(c) ? r.filter((x) => x !== c) : [...r, c]); }}
                                style={{ border: "none", background: "none", padding: 0, cursor: "pointer", display: "flex" }}>
                                {revealedCols.includes(c) ? <EyeOff size={12} color="var(--text-3)" /> : <Eye size={12} color="var(--text-3)" />}
                              </button>
                            )}
                          </div>
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {pagedRows.map((r, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid var(--border-2)" }}>
                      {dataset.columns.filter((c) => !hiddenCols.includes(c)).map((c) => {
                        const isEmpty = String(r[c]).trim() === "";
                        const isSensitive = dataset.schema.find((s) => s.name === c)?.sensitive;
                        const masked = isSensitive && !revealedCols.includes(c) && !isEmpty;
                        return (
                          <td key={c} className="dv-mono" style={{ padding: "8px 10px", whiteSpace: "nowrap", color: isEmpty ? "var(--rose)" : masked ? "var(--text-3)" : "var(--text)" }}>
                            {isEmpty ? "missing" : masked ? maskValue(r[c]) : String(r[c])}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 12, fontSize: 12, color: "var(--text-2)" }}>
              <div>Showing {pagedRows.length ? (page - 1) * pageSize + 1 : 0}–{Math.min(page * pageSize, filteredRows.length)} of {filteredRows.length} rows</div>
              <div style={{ display: "flex", gap: 6 }}>
                <button className="dv-btn dv-btn-ghost dv-btn-sm" disabled={page === 1} onClick={() => setPage((p) => p - 1)}>Prev</button>
                <button className="dv-btn dv-btn-ghost dv-btn-sm" disabled={page === totalPages} onClick={() => setPage((p) => p + 1)}>Next</button>
              </div>
            </div>
          </div>

          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <button className="dv-btn dv-btn-ghost" onClick={() => setStep(0)}><ArrowLeft size={15} /> Back</button>
            <button className="dv-btn dv-btn-primary" onClick={runAnalyze} disabled={loading}>{loading ? <><Loader2 size={15} style={{ animation: "spin 1s linear infinite" }} /> Analyzing…</> : <>Analyze Data <Sparkles size={15} /></>}</button>
          </div>
        </div>
      )}

      {step === 2 && dataset && (
        <div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 12, marginBottom: 22 }}>
            {kpis.slice(0, 8).map((k) => (
              <div key={k.id} className="dv-card" style={{ padding: 16 }}>
                <div style={{ fontSize: 11.5, color: "var(--text-2)", fontWeight: 600, marginBottom: 8 }}>{k.label}</div>
                <div className="dv-mono" style={{ fontSize: 22, fontWeight: 700, display: "flex", alignItems: "center", gap: 6 }}>
                  {formatValue(k.value, k.format)}
                  {k.metric === "growth" && (k.value >= 0 ? <TrendingUp size={15} color="var(--teal)" /> : <TrendingDown size={15} color="var(--rose)" />)}
                </div>
              </div>
            ))}
          </div>

          <div className="dv-card" style={{ padding: 20, marginBottom: 20 }}>
            <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 14 }}>Automatic Insights</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {insights.map((ins, i) => (
                <div key={i} style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
                  {ins.type === "positive" && <TrendingUp size={16} color="var(--teal)" style={{ flexShrink: 0, marginTop: 1 }} />}
                  {ins.type === "negative" && <TrendingDown size={16} color="var(--rose)" style={{ flexShrink: 0, marginTop: 1 }} />}
                  {ins.type === "warning" && <AlertTriangle size={16} color="var(--amber)" style={{ flexShrink: 0, marginTop: 1 }} />}
                  {ins.type === "info" && <Info size={16} color="var(--blue)" style={{ flexShrink: 0, marginTop: 1 }} />}
                  <div style={{ fontSize: 13.5, lineHeight: 1.5 }}>{ins.text}</div>
                </div>
              ))}
            </div>
          </div>

          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <button className="dv-btn dv-btn-ghost" onClick={() => setStep(1)}><ArrowLeft size={15} /> Back</button>
            <button className="dv-btn dv-btn-primary" onClick={goToCharts}>See Chart Recommendations <ArrowRight size={15} /></button>
          </div>
        </div>
      )}

      {step === 3 && (
        <ChartRecommendationsStep recommendations={recommendations} dataset={dataset} setRoute={setRoute} onAddToReport={onAddToReport} onCustomizeChart={onCustomizeChart} />
      )}
    </div>
  );
}

function ChartRecommendationsStep({ recommendations, dataset, setRoute, onAddToReport, onCustomizeChart }) {
  const [added, setAdded] = useState([]);
  return (
    <div>
      {recommendations.length === 0 ? (
        <EmptyState icon={BarChart3} title="No chart recommendations yet" subtitle="Analyze your dataset to generate chart recommendations." />
      ) : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(300px,1fr))", gap: 16, marginBottom: 20 }}>
          {recommendations.map((rec) => (
            <div key={rec.id} className="dv-card" style={{ padding: 16 }}>
              <div style={{ height: 140, marginBottom: 10 }}>
                <MiniChartPreview cfg={rec} rows={dataset.rows} schema={dataset.schema} />
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 4 }}>
                <Badge tone="blue">{rec.type}</Badge>
              </div>
              <div style={{ fontWeight: 700, fontSize: 14 }}>{rec.title}</div>
              <div style={{ fontSize: 12, color: "var(--text-2)", marginBottom: 8 }}>{rec.explanation}</div>
              <div style={{ display: "flex", gap: 8 }}>
                <button className="dv-btn dv-btn-sm" style={{ flex: 1, justifyContent: "center", background: added.includes(rec.id) ? "var(--teal-dim)" : "var(--blue)", color: added.includes(rec.id) ? "var(--teal)" : "#fff" }}
                  onClick={() => { onAddToReport(rec); setAdded((a) => [...a, rec.id]); }}>
                  {added.includes(rec.id) ? <><CheckCircle2 size={13} /> Added</> : "Add to Report"}
                </button>
                <button className="dv-btn dv-btn-ghost dv-btn-sm" style={{ flex: 1, justifyContent: "center" }} onClick={() => onCustomizeChart(rec.id)}>Customize</button>
              </div>
            </div>
          ))}
        </div>
      )}
      <div style={{ display: "flex", justifyContent: "space-between" }}>
        <button className="dv-btn dv-btn-ghost" onClick={() => setRoute("dashboard")}>Back to Dashboard</button>
        <button className="dv-btn dv-btn-primary" onClick={() => setRoute("report-builder")}>Go to Report Builder <ArrowRight size={15} /></button>
      </div>
    </div>
  );
}

function MiniChartPreview({ cfg, rows, schema }) {
  const { data, series } = useMemo(() => aggregateChart(rows, schema, cfg), [rows, schema, cfg]);
  const palette = getPalette(cfg.appearance?.colorPalette);
  if (!data.length) return <div style={{ height: "100%", display: "flex", alignItems: "center", justifyContent: "center", color: "var(--text-3)", fontSize: 12 }}>No data</div>;
  if (cfg.type === "scatter") return (
    <ResponsiveContainer><ScatterChart><CartesianGrid stroke="var(--border-2)" /><XAxis dataKey="x" tick={{ fontSize: 10 }} /><YAxis dataKey="y" tick={{ fontSize: 10 }} /><Scatter data={data} fill={palette[0]} /></ScatterChart></ResponsiveContainer>
  );
  if (cfg.type === "donut" || cfg.type === "pie") return (
    <ResponsiveContainer><PieChart><Pie data={data} dataKey="value" nameKey="name" innerRadius={cfg.type === "donut" ? 32 : 0} outerRadius={55}>
      {data.map((_, i) => <Cell key={i} fill={palette[i % palette.length]} />)}
    </Pie><Tooltip /></PieChart></ResponsiveContainer>
  );
  if (cfg.type === "line") return (
    <ResponsiveContainer><LineChart data={data}><CartesianGrid stroke="var(--border-2)" /><XAxis dataKey="name" tick={{ fontSize: 9 }} hide={data.length > 8} /><YAxis tick={{ fontSize: 10 }} /><Line type="monotone" dataKey="value" stroke={palette[0]} strokeWidth={2} dot={false} /></LineChart></ResponsiveContainer>
  );
  if (cfg.type === "area") return (
    <ResponsiveContainer><AreaChart data={data}><CartesianGrid stroke="var(--border-2)" /><XAxis dataKey="name" tick={{ fontSize: 9 }} hide={data.length > 8} /><YAxis tick={{ fontSize: 10 }} /><Area type="monotone" dataKey="value" stroke={palette[0]} fill={palette[0]} fillOpacity={0.25} /></AreaChart></ResponsiveContainer>
  );
  return (
    <ResponsiveContainer><BarChart data={data}><CartesianGrid stroke="var(--border-2)" /><XAxis dataKey="name" tick={{ fontSize: 9 }} hide={data.length > 8} /><YAxis tick={{ fontSize: 10 }} />
      {series.length ? series.map((s, i) => <Bar key={s} dataKey={s} stackId={cfg.type === "stacked-bar" ? "a" : undefined} fill={palette[i % palette.length]} />) : <Bar dataKey="value" fill={palette[0]} radius={[4, 4, 0, 0]} />}
    </BarChart></ResponsiveContainer>
  );
}

/* ============================== CHARTS LIBRARY / BUILDER ============================== */
function ChartsPage({ charts, setCharts, dataset, onAddToReport, editingId, setEditingId }) {
  if (editingId) {
    const chart = charts.find((c) => c.id === editingId);
    if (chart) return <ChartBuilder chart={chart} dataset={dataset} onSave={(updated) => { setCharts((cs) => cs.map((c) => c.id === updated.id ? updated : c)); setEditingId(null); }} onCancel={() => setEditingId(null)} onAddToReport={onAddToReport} />;
  }
  if (!dataset) return <div style={{ padding: 28 }}><EmptyState icon={BarChart3} title="No dataset loaded" subtitle="Run a New Analysis first to generate chart recommendations you can customize here." /></div>;
  return (
    <div className="dv-fade-in" style={{ padding: 28, maxWidth: 1140, margin: "0 auto" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
        <div style={{ fontWeight: 700, fontSize: 17 }}>Charts</div>
        <button className="dv-btn dv-btn-primary dv-btn-sm" onClick={() => {
          const safeSchema = dataset.schema.filter((s) => !s.sensitive);
          // Default X to a date/category field and Y to a numeric field where possible — picking
          // the first two raw columns blindly can land on two numeric fields (or worse, a
          // near-unique currency column as X), which produces a cluttered, hard-to-read chart.
          const xDefault = (safeSchema.find((s) => s.type === "date") || safeSchema.find((s) => s.type === "text") || safeSchema[0])?.name;
          const yDefault = (safeSchema.find((s) => ["number", "currency", "percentage"].includes(s.type) && s.name !== xDefault) || safeSchema.find((s) => s.name !== xDefault) || safeSchema[0])?.name;
          const blank = { id: uid(), type: "bar", title: "New Chart", subtitle: "", explanation: "", xField: xDefault, yField: yDefault, groupBy: "", aggregation: "sum", filters: { dateFrom: "", dateTo: "", categories: [], numMin: "", numMax: "", topN: "", bottomN: "", compareYears: [] }, appearance: { legend: true, dataLabels: false, gridlines: true, orientation: "vertical", numberFormat: "number", colorPalette: "classic", fontSize: 12 }, addedToReport: false };
          setCharts((cs) => [...cs, blank]); setEditingId(blank.id);
        }}><Plus size={14} /> New Chart</button>
      </div>
      {charts.length === 0 ? <EmptyState icon={BarChart3} title="No charts yet" subtitle="Analyze your dataset to generate chart recommendations." /> : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(300px,1fr))", gap: 16 }}>
          {charts.map((c) => (
            <div key={c.id} className="dv-card" style={{ padding: 16 }}>
              <div style={{ height: 140, marginBottom: 10 }}><MiniChartPreview cfg={c} rows={dataset.rows} schema={dataset.schema} /></div>
              <div style={{ display: "flex", gap: 6, marginBottom: 6 }}><Badge tone="blue">{c.type}</Badge>{c.addedToReport && <Badge tone="teal">In report</Badge>}</div>
              <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 8 }}>{c.title}</div>
              <div style={{ display: "flex", gap: 8 }}>
                <button className="dv-btn dv-btn-primary dv-btn-sm" style={{ flex: 1, justifyContent: "center" }} onClick={() => onAddToReport(c)}>Add to Report</button>
                <button className="dv-btn dv-btn-ghost dv-btn-sm" style={{ flex: 1, justifyContent: "center" }} onClick={() => setEditingId(c.id)}>Customize</button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ChartBuilder({ chart, dataset, onSave, onCancel, onAddToReport }) {
  const [cfg, setCfg] = useState(chart);
  const set = (path, val) => setCfg((c) => {
    const next = _.cloneDeep(c);
    _.set(next, path, val);
    return next;
  });
  const { data, series } = useMemo(() => aggregateChart(dataset.rows, dataset.schema, cfg), [dataset, cfg]);
  const safeChartSchema = dataset.schema.filter((c) => !c.sensitive);
  const numericCols = safeChartSchema.filter((c) => ["number", "currency", "percentage"].includes(c.type)).map((c) => c.name);
  const categoryCols = safeChartSchema.filter((c) => c.type === "text").map((c) => c.name);
  const xCatCols = safeChartSchema.filter((c) => ["date", "text", "boolean"].includes(c.type)).map((c) => c.name);
  const allCols = safeChartSchema.map((c) => c.name);
  const xColType = dataset.schema.find((s) => s.name === cfg.xField)?.type;
  const uniqueXVals = useMemo(() => _.uniq(dataset.rows.map((r) => String(r[cfg.xField]))).slice(0, 30), [dataset, cfg.xField]);
  const xUniqueCount = useMemo(() => new Set(dataset.rows.map((r) => String(r[cfg.xField]))).size, [dataset, cfg.xField]);
  const availableYears = useMemo(() => {
    if (xColType !== "date") return [];
    const years = new Set(dataset.rows.map((r) => { const d = new Date(r[cfg.xField]); return isNaN(d) ? null : d.getFullYear(); }).filter((y) => y !== null));
    return _.orderBy(Array.from(years));
  }, [dataset, cfg.xField, xColType]);
  const [quickYear, setQuickYear] = useState("");
  const [quickQuarter, setQuickQuarter] = useState("");
  const chartRef = useRef(null);

  const downloadPNG = () => {
    const svgEl = chartRef.current?.querySelector("svg");
    if (!svgEl) return;
    let svgStr = new XMLSerializer().serializeToString(svgEl);
    if (!svgStr.includes("xmlns=")) svgStr = svgStr.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
    const url = URL.createObjectURL(new Blob([svgStr], { type: "image/svg+xml;charset=utf-8" }));

    // Recharts renders its <Legend> as an HTML element sitting next to the <svg>, not inside it —
    // so exporting just the SVG silently drops the legend. We redraw a matching legend row
    // ourselves onto the export canvas using the same series/colors the chart is using.
    const palette = getPalette(cfg.appearance.colorPalette);
    const legendItems = !cfg.appearance.legend ? [] :
      (cfg.type === "pie" || cfg.type === "donut") ? data.map((d, i) => ({ name: d.name, color: palette[i % palette.length] })) :
      series.length ? series.map((s, i) => ({ name: s, color: palette[i % palette.length] })) :
      [{ name: cfg.yField, color: palette[0] }];

    const img = new Image();
    img.onload = () => {
      const scale = 2;
      const w = svgEl.clientWidth || 700, h = svgEl.clientHeight || 400;
      const legendRowH = legendItems.length ? 30 : 0;
      const canvas = document.createElement("canvas");
      canvas.width = w * scale; canvas.height = (h + legendRowH) * scale;
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#ffffff"; ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.scale(scale, scale);
      ctx.drawImage(img, 0, 0, w, h);
      if (legendItems.length) {
        ctx.font = "12px Inter, sans-serif";
        ctx.textBaseline = "middle";
        let x = 12, y = h + legendRowH / 2;
        legendItems.forEach((item) => {
          const textW = ctx.measureText(item.name).width;
          if (x + 16 + textW > w - 10) { x = 12; }
          ctx.fillStyle = item.color;
          ctx.fillRect(x, y - 5, 10, 10);
          ctx.fillStyle = "#333333";
          ctx.fillText(item.name, x + 15, y);
          x += 15 + textW + 18;
        });
      }
      URL.revokeObjectURL(url);
      canvas.toBlob((blob) => {
        const link = document.createElement("a");
        link.download = `${(cfg.title || "chart").replace(/[^a-z0-9]+/gi, "_")}.png`;
        link.href = URL.createObjectURL(blob);
        link.click();
      });
    };
    img.src = url;
  };

  return (
    <div className="dv-fade-in dv-split" style={{ "--split-w": "320px", height: "100%" }}>
      <div className="dv-scrollbar dv-split-side" style={{ borderRight: "1px solid var(--border)", padding: 20, overflowY: "auto" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Chart Editor</div>
          <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={onCancel}><X size={14} /></button>
        </div>

        <Section title="Chart Type">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6 }}>
            {["bar", "line", "area", "pie", "donut", "scatter", "stacked-bar", "grouped-bar"].map((t) => (
              <button key={t} onClick={() => {
                if (t === "scatter" && xColType !== "number" && xColType !== "currency" && xColType !== "percentage") {
                  const firstNumeric = numericCols.find((c) => c !== cfg.yField) || numericCols[0];
                  if (firstNumeric) { setCfg((c) => ({ ...c, type: t, xField: firstNumeric })); return; }
                }
                set("type", t);
              }} className="dv-btn dv-btn-sm" style={{ justifyContent: "center", background: cfg.type === t ? "var(--blue)" : "var(--surface)", color: cfg.type === t ? "#fff" : "var(--text-2)", border: "1px solid var(--border)" }}>{t}</button>
            ))}
          </div>
        </Section>

        <Section title="Data">
          <Field label="X-Axis">
            <select className="dv-input" value={cfg.xField} onChange={(e) => set("xField", e.target.value)}>
              {cfg.type === "scatter" ? (
                <>
                  {numericCols.length > 0 && <optgroup label="Recommended (numeric)">{numericCols.map((c) => <option key={c} value={c}>{c}</option>)}</optgroup>}
                  {xCatCols.length > 0 && <optgroup label="Other (won't plot as a number)">{xCatCols.map((c) => <option key={c} value={c}>{c}</option>)}</optgroup>}
                </>
              ) : (
                <>
                  {xCatCols.length > 0 && <optgroup label="Recommended (date / category)">{xCatCols.map((c) => <option key={c} value={c}>{c}</option>)}</optgroup>}
                  {numericCols.length > 0 && <optgroup label="Numeric (can create too many groups)">{numericCols.map((c) => <option key={c} value={c}>{c}</option>)}</optgroup>}
                </>
              )}
            </select>
            {cfg.type !== "scatter" && (numericCols.includes(cfg.xField) || xUniqueCount > 20) && (
              <div style={{ fontSize: 11, color: "var(--amber)", display: "flex", gap: 5, alignItems: "flex-start", marginTop: 6, lineHeight: 1.4 }}>
                <AlertTriangle size={12} style={{ flexShrink: 0, marginTop: 1 }} />
                <span>"{cfg.xField}" has {xUniqueCount} unique values{numericCols.includes(cfg.xField) ? " and looks like a continuous number" : ""} — {cfg.type} works best with a date or a low-cardinality category as X. Try switching X-Axis, or use Scatter for two numeric fields.</span>
              </div>
            )}
            {cfg.type === "scatter" && xColType === "date" && (
              <div style={{ fontSize: 11, color: "var(--amber)", display: "flex", gap: 5, alignItems: "flex-start", marginTop: 6, lineHeight: 1.4 }}>
                <AlertTriangle size={12} style={{ flexShrink: 0, marginTop: 1 }} />
                <span>Scatter needs two numeric fields. "{cfg.xField}" is a date — pick a numeric X, or use Line/Area to plot a trend over time instead.</span>
              </div>
            )}
          </Field>
          <Field label="Y-Axis">
            <select className="dv-input" value={cfg.yField} onChange={(e) => set("yField", e.target.value)}>{[...numericCols, ...allCols.filter(c=>!numericCols.includes(c))].map((c) => <option key={c} value={c}>{c}</option>)}</select>
          </Field>
          {["bar", "line", "stacked-bar", "grouped-bar"].includes(cfg.type) && (
            <Field label="Group By (optional)">
              <select className="dv-input" value={cfg.groupBy} onChange={(e) => set("groupBy", e.target.value)}><option value="">None</option>{categoryCols.map((c) => <option key={c} value={c}>{c}</option>)}</select>
            </Field>
          )}
          <Field label="Aggregation">
            <select className="dv-input" value={cfg.aggregation} onChange={(e) => set("aggregation", e.target.value)}>
              {["sum", "avg", "count", "min", "max"].map((a) => <option key={a} value={a}>{a}</option>)}
            </select>
          </Field>
        </Section>

        <Section title="Filters">
          {xColType === "date" && (
            <>
              <Field label="Quick range">
                <div style={{ display: "flex", gap: 6 }}>
                  <select className="dv-input" value={quickYear} onChange={(e) => {
                    const y = e.target.value; setQuickYear(y);
                    if (!y) return;
                    const q = quickQuarter && QUARTERS[quickQuarter];
                    const from = q ? new Date(Date.UTC(y, q[0], 1)) : new Date(Date.UTC(y, 0, 1));
                    const to = q ? new Date(Date.UTC(y, q[1] + 1, 0)) : new Date(Date.UTC(y, 11, 31));
                    set("filters.dateFrom", from.toISOString().slice(0, 10));
                    set("filters.dateTo", to.toISOString().slice(0, 10));
                  }}>
                    <option value="">Year…</option>
                    {availableYears.map((y) => <option key={y} value={y}>{y}</option>)}
                  </select>
                  <select className="dv-input" value={quickQuarter} onChange={(e) => {
                    const q = e.target.value; setQuickQuarter(q);
                    if (!quickYear) return;
                    const range = q && QUARTERS[q];
                    const from = range ? new Date(Date.UTC(quickYear, range[0], 1)) : new Date(Date.UTC(quickYear, 0, 1));
                    const to = range ? new Date(Date.UTC(quickYear, range[1] + 1, 0)) : new Date(Date.UTC(quickYear, 11, 31));
                    set("filters.dateFrom", from.toISOString().slice(0, 10));
                    set("filters.dateTo", to.toISOString().slice(0, 10));
                  }}>
                    <option value="">Full year</option>
                    {Object.keys(QUARTERS).map((q) => <option key={q} value={q}>{q}</option>)}
                  </select>
                </div>
                <div style={{ fontSize: 10.5, color: "var(--text-3)", marginTop: 4 }}>Fills in the From/To fields below — pick a year (and optionally a quarter) as a shortcut.</div>
              </Field>
              <div style={{ display: "flex", gap: 8 }}>
                <Field label="From"><input type="date" className="dv-input" value={cfg.filters.dateFrom} onChange={(e) => set("filters.dateFrom", e.target.value)} /></Field>
                <Field label="To"><input type="date" className="dv-input" value={cfg.filters.dateTo} onChange={(e) => set("filters.dateTo", e.target.value)} /></Field>
              </div>
              {availableYears.length > 1 && (
                <Field label="Compare years (overrides From/To — plots each year Jan–Dec)">
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                    {availableYears.map((y) => {
                      const ys = String(y);
                      const active = cfg.filters.compareYears.includes(ys);
                      return (
                        <button key={y} className="dv-btn dv-btn-sm" style={{ border: `1.5px solid ${active ? "var(--blue)" : "var(--border)"}`, background: active ? "var(--blue-dim)" : "var(--surface)", color: "var(--text)" }}
                          onClick={() => set("filters.compareYears", active ? cfg.filters.compareYears.filter((x) => x !== ys) : [...cfg.filters.compareYears, ys])}>
                          {y}
                        </button>
                      );
                    })}
                    {cfg.filters.compareYears.length > 0 && <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => set("filters.compareYears", [])}>Clear</button>}
                  </div>
                </Field>
              )}
            </>
          )}
          <Field label="Category filter">
            <select multiple className="dv-input" style={{ height: 70 }} value={cfg.filters.categories} onChange={(e) => set("filters.categories", Array.from(e.target.selectedOptions, (o) => o.value))}>
              {uniqueXVals.map((v) => <option key={v} value={v}>{v}</option>)}
            </select>
          </Field>
          <div style={{ display: "flex", gap: 8 }}>
            <Field label="Numeric min"><input type="number" className="dv-input" value={cfg.filters.numMin} onChange={(e) => set("filters.numMin", e.target.value)} /></Field>
            <Field label="Numeric max"><input type="number" className="dv-input" value={cfg.filters.numMax} onChange={(e) => set("filters.numMax", e.target.value)} /></Field>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <Field label="Top N"><input type="number" className="dv-input" value={cfg.filters.topN} onChange={(e) => set("filters.topN", e.target.value)} /></Field>
            <Field label="Bottom N"><input type="number" className="dv-input" value={cfg.filters.bottomN} onChange={(e) => set("filters.bottomN", e.target.value)} /></Field>
          </div>
        </Section>

        <Section title="Appearance">
          <Field label="Chart title"><input className="dv-input" value={cfg.title} onChange={(e) => set("title", e.target.value)} /></Field>
          <Field label="Subtitle"><input className="dv-input" value={cfg.subtitle} onChange={(e) => set("subtitle", e.target.value)} /></Field>
          <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 6 }}>
            {[["legend", "Legend"], ["dataLabels", "Data labels"], ["gridlines", "Gridlines"]].map(([k, l]) => (
              <label key={k} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}>
                <input type="checkbox" checked={cfg.appearance[k]} onChange={(e) => set(`appearance.${k}`, e.target.checked)} /> {l}
              </label>
            ))}
          </div>
          <Field label="Orientation">
            <select className="dv-input" value={cfg.appearance.orientation} onChange={(e) => set("appearance.orientation", e.target.value)}><option value="vertical">Vertical (Column)</option><option value="horizontal">Horizontal (Bar)</option></select>
          </Field>
          <Field label="Number format">
            <select className="dv-input" value={cfg.appearance.numberFormat} onChange={(e) => set("appearance.numberFormat", e.target.value)}><option value="number">Number</option><option value="currency">Currency</option><option value="percentage">Percentage</option></select>
          </Field>
          <Field label="Color palette">
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {Object.entries(COLOR_PALETTES).map(([id, p]) => (
                <button key={id} onClick={() => set("appearance.colorPalette", id)} className="dv-btn dv-btn-sm"
                  style={{ justifyContent: "flex-start", gap: 8, border: `1.5px solid ${cfg.appearance.colorPalette === id ? "var(--blue)" : "var(--border)"}`, background: cfg.appearance.colorPalette === id ? "var(--blue-dim)" : "var(--surface)", color: "var(--text)" }}>
                  <div style={{ display: "flex", gap: 3 }}>{p.colors.slice(0, 5).map((c, i) => <span key={i} style={{ width: 11, height: 11, borderRadius: 3, background: c, display: "inline-block" }} />)}</div>
                  {p.label}
                </button>
              ))}
            </div>
          </Field>
          <Field label={`Chart font size — ${cfg.appearance.fontSize || 12}px`}>
            <input type="range" min="9" max="18" step="1" value={cfg.appearance.fontSize || 12} onChange={(e) => set("appearance.fontSize", Number(e.target.value))} style={{ width: "100%" }} />
          </Field>
        </Section>

        <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
          <button className="dv-btn dv-btn-primary" style={{ flex: 1, justifyContent: "center" }} onClick={() => onSave(cfg)}>Save Chart</button>
          <button className="dv-btn dv-btn-ghost" onClick={() => { onAddToReport(cfg); onSave(cfg); }}>Add to Report</button>
        </div>
      </div>

      <div style={{ padding: 28, background: "var(--paper)", overflowY: "auto" }}>
        <div className="dv-card" style={{ padding: 26, maxWidth: 720, margin: "0 auto" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
            <div>
              <div style={{ fontWeight: 700, fontSize: 17 }}>{cfg.title}</div>
              {cfg.subtitle && <div style={{ fontSize: 13, color: "var(--text-2)" }}>{cfg.subtitle}</div>}
            </div>
            <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={downloadPNG}><Download size={13} /> PNG</button>
          </div>
          <div ref={chartRef} style={{ height: 380, marginTop: 14 }}>
            <FullChart cfg={cfg} data={data} series={series} />
          </div>
        </div>
      </div>
    </div>
  );
}

const Section = ({ title, children }) => (
  <div style={{ marginBottom: 20, paddingBottom: 18, borderBottom: "1px solid var(--border-2)" }}>
    <div style={{ fontSize: 12, fontWeight: 700, color: "var(--text-2)", textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 10 }}>{title}</div>
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>{children}</div>
  </div>
);
const Field = ({ label, children }) => <div style={{ flex: 1 }}><span className="dv-label">{label}</span>{children}</div>;

function FullChart({ cfg, data, series }) {
  if (!data.length) return <div style={{ height: "100%", display: "flex", alignItems: "center", justifyContent: "center", color: "var(--text-3)" }}>No data matches the current filters</div>;
  const palette = getPalette(cfg.appearance.colorPalette);
  const fs = cfg.appearance.fontSize || 12;
  const tick = { fontSize: fs };
  const grid = cfg.appearance.gridlines ? <CartesianGrid stroke="var(--border-2)" /> : null;
  const legend = cfg.appearance.legend ? <Legend wrapperStyle={{ fontSize: fs }} /> : null;
  const labelFmt = (v) => cfg.appearance.numberFormat === "currency" ? currencySymbol(currentCurrency) + Number(v).toLocaleString() : cfg.appearance.numberFormat === "percentage" ? v + "%" : Number(v).toLocaleString();
  const dl = cfg.appearance.dataLabels;

  if (cfg.type === "scatter") return (
    <ResponsiveContainer><ScatterChart>{grid}<XAxis dataKey="x" name={cfg.xField} tick={tick} /><YAxis dataKey="y" name={cfg.yField} tick={tick} /><Tooltip cursor={{ strokeDasharray: "3 3" }} />{legend}
      <Scatter data={data} fill={palette[0]}>{dl && <LabelList dataKey="y" position="top" style={{ fontSize: fs - 1, fill: "var(--text-2)" }} formatter={labelFmt} />}</Scatter>
    </ScatterChart></ResponsiveContainer>
  );
  if (cfg.type === "pie" || cfg.type === "donut") {
    // Custom label renderer: short text (value only — the name is already in the legend/tooltip),
    // positioned outside the ring along its own slice angle so labels fan out instead of stacking.
    const RADIAN = Math.PI / 180;
    const renderLabel = (props) => {
      const { cx, cy, midAngle, outerRadius: or_, index, value } = props;
      const r = or_ + 16;
      const x = cx + r * Math.cos(-midAngle * RADIAN);
      const y = cy + r * Math.sin(-midAngle * RADIAN);
      return <text x={x} y={y} fill={palette[index % palette.length]} textAnchor={x > cx ? "start" : "end"} dominantBaseline="central" fontSize={fs} fontWeight={600}>{labelFmt(value)}</text>;
    };
    return (
      <ResponsiveContainer><PieChart margin={{ top: 22, right: 70, bottom: 22, left: 70 }}>
        <Pie data={data} dataKey="value" nameKey="name" innerRadius={cfg.type === "donut" ? 62 : 0} outerRadius={104} label={dl ? renderLabel : false} labelLine={dl ? { stroke: "var(--border)" } : false} minAngle={3}>
          {data.map((_, i) => <Cell key={i} fill={palette[i % palette.length]} />)}
        </Pie><Tooltip formatter={labelFmt} />{legend}
      </PieChart></ResponsiveContainer>
    );
  }
  if (cfg.type === "line") return (
    <ResponsiveContainer><LineChart data={data}>{grid}<XAxis dataKey="name" tick={tick} /><YAxis tick={tick} tickFormatter={labelFmt} /><Tooltip formatter={labelFmt} />{legend}
      {series.length ? series.map((s, i) => (
        <Line key={s} type="monotone" dataKey={s} stroke={palette[i % palette.length]} strokeWidth={2.5} dot={false}>
          {dl && <LabelList dataKey={s} position="top" style={{ fontSize: fs - 1, fill: palette[i % palette.length] }} formatter={labelFmt} />}
        </Line>
      )) : (
        <Line type="monotone" dataKey="value" stroke={palette[0]} strokeWidth={2.5} dot={{ r: 3 }}>
          {dl && <LabelList dataKey="value" position="top" style={{ fontSize: fs - 1, fill: palette[0] }} formatter={labelFmt} />}
        </Line>
      )}
    </LineChart></ResponsiveContainer>
  );
  if (cfg.type === "area") return (
    <ResponsiveContainer><AreaChart data={data}>{grid}<XAxis dataKey="name" tick={tick} /><YAxis tick={tick} tickFormatter={labelFmt} /><Tooltip formatter={labelFmt} />{legend}
      <Area type="monotone" dataKey="value" stroke={palette[0]} fill={palette[0]} fillOpacity={0.25} strokeWidth={2.5}>
        {dl && <LabelList dataKey="value" position="top" style={{ fontSize: fs - 1, fill: "var(--text-2)" }} formatter={labelFmt} />}
      </Area>
    </AreaChart></ResponsiveContainer>
  );
  return (
    <ResponsiveContainer><BarChart data={data} layout={cfg.appearance.orientation === "horizontal" ? "vertical" : "horizontal"}>{grid}
      {cfg.appearance.orientation === "horizontal" ? <><XAxis type="number" tick={tick} tickFormatter={labelFmt} /><YAxis type="category" dataKey="name" tick={tick} width={90} /></> : <><XAxis dataKey="name" tick={tick} /><YAxis tick={tick} tickFormatter={labelFmt} /></>}
      <Tooltip formatter={labelFmt} />{legend}
      {series.length ? series.map((s, i) => (
        <Bar key={s} dataKey={s} stackId={cfg.type === "stacked-bar" ? "a" : undefined} fill={palette[i % palette.length]} radius={[3, 3, 0, 0]}>
          {dl && <LabelList dataKey={s} position={cfg.appearance.orientation === "horizontal" ? "right" : "top"} style={{ fontSize: fs - 1, fill: "var(--text-2)" }} formatter={labelFmt} />}
        </Bar>
      )) : (
        <Bar dataKey="value" fill={palette[0]} radius={[3, 3, 0, 0]}>
          {dl && <LabelList dataKey="value" position={cfg.appearance.orientation === "horizontal" ? "right" : "top"} style={{ fontSize: fs - 1, fill: "var(--text-2)" }} formatter={labelFmt} />}
        </Bar>
      )}
    </BarChart></ResponsiveContainer>
  );
}

/* ============================== REPORT BUILDER ============================== */
const SECTION_LIBRARY = [
  { type: "cover", label: "Cover Page", icon: FileText },
  { type: "summary", label: "Executive Summary", icon: Sparkles },
  { type: "kpi", label: "KPI Cards", icon: TrendingUp },
  { type: "chart", label: "Chart", icon: BarChart3 },
  { type: "text", label: "Text Block", icon: TypeIcon },
  { type: "table", label: "Data Table", icon: Table2 },
  { type: "insights", label: "Key Insights", icon: Sparkles },
];

const TEMPLATES = [
  { id: "executive", name: "Executive", desc: "Clean corporate design", accent: "#2F5FED" },
  { id: "sales", name: "Sales", desc: "Optimized for revenue reports", accent: "#0EA894" },
  { id: "marketing", name: "Marketing", desc: "Campaigns, leads, conversion", accent: "#7C5CFA" },
  { id: "financial", name: "Financial", desc: "Optimized for financial metrics", accent: "#DD9B2E" },
  { id: "simple", name: "Simple", desc: "Minimal clean report", accent: "#5B6472" },
];

const ACCENT_SWATCHES = ["#2F5FED", "#0EA894", "#7C5CFA", "#DD9B2E", "#DD5350", "#1F3A5F", "#2AA8D8", "#5B6472"];

// Logo is absolutely positioned within the fixed-size cover page div, so moving/resizing it never
// affects the page's layout flow or dimensions — keeps print pagination exactly as designed.
const LOGO_POSITIONS = {
  "top-left": { top: 40, left: 40 },
  "top-center": { top: 40, left: "50%", transform: "translateX(-50%)" },
  "top-right": { top: 40, right: 40 },
  "middle-left": { top: "50%", left: 40, transform: "translateY(-50%)" },
  "center": { top: "50%", left: "50%", transform: "translate(-50%,-50%)" },
  "middle-right": { top: "50%", right: 40, transform: "translateY(-50%)" },
  "bottom-left": { bottom: 40, left: 40 },
  "bottom-center": { bottom: 40, left: "50%", transform: "translateX(-50%)" },
  "bottom-right": { bottom: 40, right: 40 },
};
const LOGO_GRID = [["top-left", "top-center", "top-right"], ["middle-left", "center", "middle-right"], ["bottom-left", "bottom-center", "bottom-right"]];

function getReportTheme(report) {
  const t = TEMPLATES.find((x) => x.id === report.template) || TEMPLATES[0];
  const d = report.design || {};
  return {
    name: t.name,
    accent: d.accentColor || t.accent,
    headingFont: d.headingFont || "serif",
    bodyFontSize: d.bodyFontSize || 13,
  };
}

function ReportBuilder({ report, setReport, dataset, analysis, charts, setRoute, askConfirm }) {
  const [zoom, setZoom] = useState(0.34);
  const addElement = (type) => {
    const el = { id: uid(), type, title: SECTION_LIBRARY.find((s) => s.type === type)?.label, config: {} };
    // Chart sections start unassigned — the user must explicitly pick which chart shows here,
    // rather than silently defaulting to whichever chart happens to be first in the list.
    if (type === "chart") el.config.chartId = null;
    if (type === "kpi") el.config.kpis = (analysis?.kpis || []).slice(0, 4).map((k) => ({ id: uid(), field: k.field, metric: k.metric, label: k.label, format: k.format }));
    if (type === "text") el.config.body = "Add your commentary here…";
    if (type === "summary") el.config.body = "";
    if (type === "table") {
      const safeCols = dataset ? dataset.schema.filter((s) => !s.sensitive).map((s) => s.name) : [];
      el.config = { columns: safeCols.slice(0, 8), rowLimit: 10, sortBy: "", sortDir: "desc" };
    }
    setReport((r) => ({ ...r, elements: [...r.elements, el] }));
  };
  const move = (id, dir) => setReport((r) => {
    const idx = r.elements.findIndex((e) => e.id === id);
    const swap = idx + dir;
    if (swap < 0 || swap >= r.elements.length) return r;
    const els = [...r.elements];
    [els[idx], els[swap]] = [els[swap], els[idx]];
    return { ...r, elements: els };
  });
  const remove = (id) => setReport((r) => ({ ...r, elements: r.elements.filter((e) => e.id !== id) }));
  const updateEl = (id, patch) => setReport((r) => ({ ...r, elements: r.elements.map((e) => e.id === id ? { ...e, ...patch } : e) }));
  const setDesign = (patch) => setReport((r) => ({ ...r, design: { ...(r.design || {}), ...patch } }));
  const setLogo = (patch) => setReport((r) => ({ ...r, logo: { ...(r.logo || { dataUrl: null, size: 64, position: "top-left" }), ...patch } }));
  const handleLogoUpload = (file) => {
    if (!file || !file.type.startsWith("image/")) return;
    const reader = new FileReader();
    reader.onload = (e) => setLogo({ dataUrl: e.target.result });
    reader.readAsDataURL(file);
  };

  if (!dataset) return <div style={{ padding: 28 }}><EmptyState icon={FileText} title="No dataset loaded" subtitle="Run a New Analysis first, then come back to build your report." /></div>;

  const theme = getReportTheme(report);
  const pageW = report.orientation === "landscape" ? 1123 : 794;
  const pageH = report.orientation === "landscape" ? 794 : 1123;

  return (
    <div className="dv-fade-in" style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 24px", borderBottom: "1px solid var(--border)", gap: 12, flexWrap: "wrap" }}>
        <input className="dv-input" style={{ fontSize: 16, fontWeight: 700, border: "none", padding: "4px 0", width: 300, background: "transparent" }} value={report.title} onChange={(e) => setReport((r) => ({ ...r, title: e.target.value }))} />
        <div style={{ display: "flex", gap: 8 }}>
          <button className="dv-btn dv-btn-ghost" onClick={() => setRoute("report-preview")}><Eye size={15} /> Full Preview</button>
          <button className="dv-btn dv-btn-primary" onClick={() => setRoute("report-preview")}><Download size={15} /> Export PDF</button>
        </div>
      </div>

      <div className="dv-split" style={{ "--split-w": "420px", flex: 1, minHeight: 0 }}>
        <div className="dv-scrollbar dv-split-side" style={{ overflowY: "auto", padding: 20, borderRight: "1px solid var(--border)" }}>
          <Section title="Template">
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {TEMPLATES.map((t) => (
                <button key={t.id} onClick={() => setReport((r) => ({ ...r, template: t.id }))} className="dv-btn dv-btn-sm" style={{ border: `1.5px solid ${report.template === t.id ? t.accent : "var(--border)"}`, background: report.template === t.id ? t.accent + "18" : "var(--surface)", color: "var(--text)", alignItems: "center", height: "auto", padding: "7px 10px", gap: 8 }}>
                  <svg width="34" height="24" viewBox="0 0 34 24" className="dv-tpl-thumb">
                    <rect width="34" height="24" fill="#fff" />
                    <rect x="3" y="3" width="13" height="2.5" fill={t.accent} rx="1" />
                    <rect x="3" y="8" width="28" height="1" fill="#E3E6EC" />
                    <rect x="3" y="12" width="6" height="9" fill={t.accent} opacity="0.2" />
                    <rect x="11" y="15" width="6" height="6" fill={t.accent} opacity="0.4" />
                    <rect x="19" y="9" width="6" height="12" fill={t.accent} opacity="0.6" />
                  </svg>
                  <span style={{ display: "flex", flexDirection: "column", alignItems: "flex-start" }}>
                    <span style={{ fontWeight: 700 }}>{t.name}</span><span style={{ fontSize: 10, color: "var(--text-2)", fontWeight: 400 }}>{t.desc}</span>
                  </span>
                </button>
              ))}
            </div>
            <Field label="Report title"><input className="dv-input" value={report.reportTitle} onChange={(e) => setReport((r) => ({ ...r, reportTitle: e.target.value }))} /></Field>
            <Field label="Company name"><input className="dv-input" value={report.company} onChange={(e) => setReport((r) => ({ ...r, company: e.target.value }))} /></Field>
            <Field label="Author name"><input className="dv-input" value={report.author} onChange={(e) => setReport((r) => ({ ...r, author: e.target.value }))} /></Field>
            <Field label="Orientation">
              <select className="dv-input" value={report.orientation} onChange={(e) => setReport((r) => ({ ...r, orientation: e.target.value }))}><option value="portrait">Portrait</option><option value="landscape">Landscape</option></select>
            </Field>
          </Section>

          <Section title="Design">
            <Field label="Heading font">
              <div style={{ display: "flex", gap: 6 }}>
                <button className="dv-btn dv-btn-sm" style={{ flex: 1, justifyContent: "center", fontFamily: "'Fraunces',serif", border: `1.5px solid ${theme.headingFont === "serif" ? "var(--blue)" : "var(--border)"}` }} onClick={() => setDesign({ headingFont: "serif" })}>Serif</button>
                <button className="dv-btn dv-btn-sm" style={{ flex: 1, justifyContent: "center", fontFamily: "'Inter',sans-serif", border: `1.5px solid ${theme.headingFont === "sans" ? "var(--blue)" : "var(--border)"}` }} onClick={() => setDesign({ headingFont: "sans" })}>Sans</button>
              </div>
            </Field>
            <Field label={`Body text size — ${theme.bodyFontSize}px`}>
              <input type="range" min="10" max="16" step="1" value={theme.bodyFontSize} onChange={(e) => setDesign({ bodyFontSize: Number(e.target.value) })} style={{ width: "100%" }} />
            </Field>
            <Field label="Accent color">
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                {ACCENT_SWATCHES.map((c) => (
                  <button key={c} onClick={() => setDesign({ accentColor: c })} style={{ width: 24, height: 24, borderRadius: 7, background: c, border: theme.accent === c ? "2px solid var(--text)" : "2px solid transparent", cursor: "pointer" }} />
                ))}
                <input type="color" value={theme.accent} onChange={(e) => setDesign({ accentColor: e.target.value })} style={{ width: 28, height: 24, border: "1px solid var(--border)", borderRadius: 6, padding: 0, background: "none", cursor: "pointer" }} />
                {report.design?.accentColor && <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => setDesign({ accentColor: null })}>Reset</button>}
              </div>
            </Field>
          </Section>

          <Section title="Branding">
            <Field label="Logo (cover page)">
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                {report.logo?.dataUrl ? (
                  <img src={report.logo.dataUrl} alt="Logo preview" style={{ width: 44, height: 44, objectFit: "contain", border: "1px solid var(--border)", borderRadius: 8, background: "#fff" }} />
                ) : (
                  <div style={{ width: 44, height: 44, borderRadius: 8, border: "1px dashed var(--border)", display: "flex", alignItems: "center", justifyContent: "center" }}><Building2 size={16} color="var(--text-3)" /></div>
                )}
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  <label className="dv-btn dv-btn-ghost dv-btn-sm" style={{ cursor: "pointer" }}>
                    <Upload size={12} /> {report.logo?.dataUrl ? "Replace" : "Upload"}
                    <input type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => handleLogoUpload(e.target.files[0])} />
                  </label>
                  {report.logo?.dataUrl && <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => setLogo({ dataUrl: null })}>Remove</button>}
                </div>
              </div>
            </Field>
            {report.logo?.dataUrl && (
              <>
                <Field label={`Logo size — ${report.logo.size || 64}px`}>
                  <input type="range" min="32" max="180" step="4" value={report.logo.size || 64} onChange={(e) => setLogo({ size: Number(e.target.value) })} style={{ width: "100%" }} />
                </Field>
                <Field label="Position on cover page">
                  <div style={{ display: "inline-grid", gridTemplateColumns: "repeat(3, 28px)", gridTemplateRows: "repeat(3, 28px)", gap: 4 }}>
                    {LOGO_GRID.flat().map((pos) => (
                      <button key={pos} title={pos} onClick={() => setLogo({ position: pos })}
                        style={{ width: 28, height: 28, borderRadius: 6, border: `1.5px solid ${(report.logo.position || "top-left") === pos ? "var(--blue)" : "var(--border)"}`, background: (report.logo.position || "top-left") === pos ? "var(--blue-dim)" : "var(--surface)", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
                        <div style={{ width: 6, height: 6, borderRadius: 2, background: (report.logo.position || "top-left") === pos ? "var(--blue)" : "var(--text-3)" }} />
                      </button>
                    ))}
                  </div>
                </Field>
              </>
            )}
          </Section>

          <div className="dv-label" style={{ marginBottom: 8 }}>Add section</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 18 }}>
            {SECTION_LIBRARY.map((s) => (
              <button key={s.type} className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => addElement(s.type)}><s.icon size={13} /> {s.label}</button>
            ))}
          </div>

          <div className="dv-label" style={{ marginBottom: 8 }}>Sections in this report</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {report.elements.length === 0 && <EmptyState icon={LayoutGrid} title="Your report is empty" subtitle="Add sections above to start building your report." />}
            {report.elements.map((el, i) => (
              <div key={el.id} className="dv-card" style={{ padding: 12 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 7, minWidth: 0 }}>
                    <GripVertical size={13} color="var(--text-3)" style={{ flexShrink: 0 }} />
                    <input className="dv-input" style={{ border: "none", fontWeight: 700, fontSize: 13, padding: "2px 0", width: 140 }} value={el.title} onChange={(e) => updateEl(el.id, { title: e.target.value })} />
                    <Badge tone="blue">{el.type}</Badge>
                  </div>
                  <div style={{ display: "flex", gap: 3, flexShrink: 0 }}>
                    <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => move(el.id, -1)}><ChevronUp size={12} /></button>
                    <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => move(el.id, 1)}><ChevronDown size={12} /></button>
                    <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => askConfirm("Remove section?", `"${el.title}" will be removed from this report.`, () => remove(el.id))}><Trash2 size={12} color="var(--rose)" /></button>
                  </div>
                </div>
                <ElementEditor el={el} updateEl={updateEl} dataset={dataset} analysis={analysis} charts={charts} />
              </div>
            ))}
          </div>
        </div>

        <div className="dv-scrollbar" style={{ background: "var(--ink)", overflowY: "auto", padding: 24, display: "flex", flexDirection: "column", alignItems: "center", gap: 18 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, position: "sticky", top: 0, zIndex: 2, background: "rgba(11,18,32,.85)", backdropFilter: "blur(6px)", padding: "6px 12px", borderRadius: 999 }}>
            <span style={{ fontSize: 11, color: "#9BA7C4", fontWeight: 600 }}>Live preview</span>
            <button className="dv-btn dv-btn-ghost dv-btn-sm" style={{ color: "#fff" }} onClick={() => setZoom((z) => Math.max(0.2, z - 0.06))}><ZoomOut size={13} /></button>
            <span style={{ fontSize: 11, color: "#9BA7C4", width: 34, textAlign: "center" }}>{Math.round(zoom * 100)}%</span>
            <button className="dv-btn dv-btn-ghost dv-btn-sm" style={{ color: "#fff" }} onClick={() => setZoom((z) => Math.min(0.7, z + 0.06))}><ZoomIn size={13} /></button>
          </div>
          {(() => {
            const pageCount = report.elements.length || 1;
            const totalH = pageCount * pageH + (pageCount - 1) * 24;
            return (
              <div style={{ width: pageW * zoom, height: totalH * zoom }}>
                <div style={{ transform: `scale(${zoom})`, transformOrigin: "top left", width: pageW }}>
                  <ReportPages report={report} dataset={dataset} analysis={analysis} charts={charts} theme={theme} pageW={pageW} pageH={pageH} />
                </div>
              </div>
            );
          })()}
        </div>
      </div>
    </div>
  );
}

function ElementEditor({ el, updateEl, dataset, analysis, charts }) {
  if (el.type === "chart") {
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <select className="dv-input" value={el.config.chartId || ""} onChange={(e) => {
          const chosen = charts.find((c) => c.id === e.target.value);
          updateEl(el.id, { config: { ...el.config, chartId: e.target.value }, title: chosen ? chosen.title : el.title });
        }}>
          <option value="">Select a chart…</option>
          {charts.map((c) => <option key={c.id} value={c.id}>{c.title}</option>)}
        </select>
        <textarea className="dv-input" rows={2} placeholder="Optional paragraph under this chart (e.g. context or takeaway)…" value={el.config.caption || ""} onChange={(e) => updateEl(el.id, { config: { ...el.config, caption: e.target.value } })} />
      </div>
    );
  }
  if (el.type === "text") {
    return <textarea className="dv-input" rows={3} value={el.config.body} onChange={(e) => updateEl(el.id, { config: { ...el.config, body: e.target.value } })} />;
  }
  if (el.type === "kpi") {
    return <KpiEditor el={el} updateEl={updateEl} dataset={dataset} />;
  }
  if (el.type === "table") {
    return <TableEditor el={el} updateEl={updateEl} dataset={dataset} />;
  }
  if (el.type === "insights") {
    return <div style={{ fontSize: 12, color: "var(--text-2)" }}>{(analysis?.insights || []).length} automatic insight(s) will be listed here</div>;
  }
  if (el.type === "summary") {
    return <textarea className="dv-input" rows={3} placeholder="Write a short overview of overall performance, key findings and recommendations…" value={el.config.body || ""} onChange={(e) => updateEl(el.id, { config: { ...el.config, body: e.target.value } })} />;
  }
  return <div style={{ fontSize: 12, color: "var(--text-2)" }}>Cover page with report title, company and author</div>;
}

// Lets someone pick exactly which KPI cards appear in this section — column + metric + an
// optional custom label — rather than being stuck with whatever computeKpis() auto-generated.
function KpiEditor({ el, updateEl, dataset }) {
  const [field, setField] = useState("");
  const [metric, setMetric] = useState("sum");
  const [label, setLabel] = useState("");
  const numericCols = dataset.schema.filter((s) => !s.sensitive && ["number", "currency", "percentage"].includes(s.type));
  const hasDateCol = dataset.schema.some((s) => s.type === "date");
  const kpis = el.config.kpis || [];

  const addKpi = () => {
    if (!field) return;
    const autoLabel = label.trim() || `${metric === "growth" ? "Growth of " : metric.charAt(0).toUpperCase() + metric.slice(1) + " "}${field}`;
    const spec = { id: uid(), field, metric, label: autoLabel, format: kpiSpecFormat(dataset.schema, field, metric) };
    updateEl(el.id, { config: { ...el.config, kpis: [...kpis, spec] } });
    setField(""); setLabel("");
  };
  const removeKpi = (id) => updateEl(el.id, { config: { ...el.config, kpis: kpis.filter((k) => k.id !== id) } });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      {kpis.length === 0 && <div style={{ fontSize: 11.5, color: "var(--text-3)" }}>No KPI cards chosen yet — add one below.</div>}
      {kpis.map((k) => (
        <div key={k.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: "var(--paper)", borderRadius: 8, padding: "6px 10px" }}>
          <span style={{ fontSize: 12 }}>{k.label} <span style={{ color: "var(--text-3)" }}>({k.metric} of {k.field})</span></span>
          <button onClick={() => removeKpi(k.id)} style={{ border: "none", background: "none", cursor: "pointer", display: "flex" }}><X size={12} color="var(--rose)" /></button>
        </div>
      ))}
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center", borderTop: "1px dashed var(--border)", paddingTop: 8, marginTop: 2 }}>
        <select className="dv-input" style={{ flex: "1 1 110px" }} value={field} onChange={(e) => setField(e.target.value)}>
          <option value="">Column…</option>
          {numericCols.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
        </select>
        <select className="dv-input" style={{ flex: "0 0 84px" }} value={metric} onChange={(e) => setMetric(e.target.value)}>
          {["sum", "avg", "min", "max", "count"].map((m) => <option key={m} value={m}>{m}</option>)}
          {hasDateCol && <option value="growth">growth</option>}
        </select>
        <input className="dv-input" style={{ flex: "1 1 100px" }} placeholder="Label (optional)" value={label} onChange={(e) => setLabel(e.target.value)} />
        <button className="dv-btn dv-btn-primary dv-btn-sm" onClick={addKpi} disabled={!field}><Plus size={12} /></button>
      </div>
    </div>
  );
}

// Lets someone pick which columns show, how many rows, and a sort order for the report's data
// table — instead of always dumping the first 8 columns / first 10 rows.
function TableEditor({ el, updateEl, dataset }) {
  const safeCols = dataset.schema.filter((s) => !s.sensitive).map((s) => s.name);
  const catCols = dataset.schema.filter((s) => !s.sensitive && ["text", "boolean", "date"].includes(s.type)).map((s) => s.name);
  const numCols = dataset.schema.filter((s) => !s.sensitive && ["number", "currency", "percentage"].includes(s.type)).map((s) => s.name);
  const selected = el.config.columns || safeCols.slice(0, 8);
  const rowLimit = el.config.rowLimit ?? 10;
  const sortBy = el.config.sortBy || "";
  const sortDir = el.config.sortDir || "desc";
  const groupBy = el.config.groupBy || "";
  const aggregation = el.config.aggregation || "sum";
  const excludedCount = dataset.schema.filter((s) => s.sensitive).length;

  const toggleCol = (name) => {
    const next = selected.includes(name) ? selected.filter((c) => c !== name) : [...selected, name];
    updateEl(el.id, { config: { ...el.config, columns: next } });
  };
  const setCfg = (patch) => updateEl(el.id, { config: { ...el.config, ...patch } });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      <div>
        <span className="dv-label">Group by (optional)</span>
        <select className="dv-input" value={groupBy} onChange={(e) => setCfg({ groupBy: e.target.value })}>
          <option value="">No grouping — show raw rows</option>
          {catCols.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
        {groupBy && <div style={{ fontSize: 10.5, color: "var(--text-3)", marginTop: 4 }}>Shows one row per "{groupBy}", with numeric columns below summarized by the aggregation you pick.</div>}
      </div>
      {groupBy && (
        <div>
          <span className="dv-label">Aggregation</span>
          <select className="dv-input" value={aggregation} onChange={(e) => setCfg({ aggregation: e.target.value })}>
            {["sum", "avg", "count", "min", "max"].map((a) => <option key={a} value={a}>{a}</option>)}
          </select>
        </div>
      )}
      <div>
        <div className="dv-label" style={{ marginBottom: 5 }}>Columns ({selected.length} of {safeCols.length})</div>
        <div className="dv-scrollbar" style={{ display: "flex", flexWrap: "wrap", gap: 5, maxHeight: 96, overflowY: "auto" }}>
          {(groupBy ? numCols : safeCols).map((c) => (
            <button key={c} onClick={() => toggleCol(c)} className="dv-btn dv-btn-sm"
              style={{ border: `1.5px solid ${selected.includes(c) ? "var(--blue)" : "var(--border)"}`, background: selected.includes(c) ? "var(--blue-dim)" : "var(--surface)", color: "var(--text)" }}>
              {c}
            </button>
          ))}
        </div>
        {groupBy && <div style={{ fontSize: 10.5, color: "var(--text-3)", marginTop: 4 }}>Only numeric columns can be aggregated — "{groupBy}" is always shown as the first column.</div>}
        {excludedCount > 0 && <div style={{ fontSize: 10.5, color: "var(--text-3)", marginTop: 4 }}>{excludedCount} personal-data column{excludedCount === 1 ? "" : "s"} always excluded.</div>}
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <div style={{ flex: 1 }}>
          <span className="dv-label">Rows</span>
          <select className="dv-input" value={rowLimit} onChange={(e) => setCfg({ rowLimit: Number(e.target.value) })}>
            {[5, 10, 20, 50, 100].map((n) => <option key={n} value={n}>{n} rows</option>)}
            <option value={0}>All rows</option>
          </select>
        </div>
        <div style={{ flex: 1 }}>
          <span className="dv-label">Sort by</span>
          <select className="dv-input" value={sortBy} onChange={(e) => setCfg({ sortBy: e.target.value })}>
            <option value="">Original order</option>
            {(groupBy ? [groupBy, ...numCols] : safeCols).map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        {sortBy && (
          <div style={{ flex: "0 0 74px" }}>
            <span className="dv-label">Dir</span>
            <select className="dv-input" value={sortDir} onChange={(e) => setCfg({ sortDir: e.target.value })}>
              <option value="asc">Asc</option>
              <option value="desc">Desc</option>
            </select>
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================== REPORT PREVIEW / PDF ============================== */
function ReportPreview({ report, dataset, analysis, charts, setRoute }) {
  const [zoom, setZoom] = useState(0.62);
  const theme = getReportTheme(report);
  const pageW = report.orientation === "landscape" ? 1123 : 794;
  const pageH = report.orientation === "landscape" ? 794 : 1123;

  return (
    <div className="dv-fade-in" style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 24px", borderBottom: "1px solid var(--border)" }}>
        <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => setRoute("report-builder")}><ArrowLeft size={14} /> Edit report</button>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => setZoom((z) => Math.max(0.3, z - 0.1))}><ZoomOut size={14} /></button>
          <span style={{ fontSize: 12, color: "var(--text-2)", width: 40, textAlign: "center" }}>{Math.round(zoom * 100)}%</span>
          <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => setZoom((z) => Math.min(1.2, z + 0.1))}><ZoomIn size={14} /></button>
          <button className="dv-btn dv-btn-primary dv-btn-sm" onClick={() => window.print()}><Download size={14} /> Export PDF</button>
        </div>
      </div>
      <div className="dv-scrollbar" style={{ flex: 1, overflow: "auto", background: "var(--ink)", padding: 30, display: "flex", flexDirection: "column", alignItems: "center", gap: 24 }}>
        <div id="dv-print-area" style={{ transform: `scale(${zoom})`, transformOrigin: "top center" }}>
          <ReportPages report={report} dataset={dataset} analysis={analysis} charts={charts} theme={theme} pageW={pageW} pageH={pageH} />
        </div>
      </div>
    </div>
  );
}

function ReportPages({ report, dataset, analysis, charts, theme, pageW, pageH }) {
  const cover = report.elements.find((e) => e.type === "cover");
  const rest = report.elements.filter((e) => e.type !== "cover");
  const pad = 52;
  const headingClass = theme.headingFont === "serif" ? "dv-serif" : "";
  const headingFamily = theme.headingFont === "serif" ? "'Fraunces',serif" : "'Inter',sans-serif";

  return (
    <>
      {cover && (
        <div className="dv-report-page" style={{ width: pageW, height: pageH, marginBottom: 24, padding: pad, display: "flex", flexDirection: "column", justifyContent: "space-between", position: "relative", overflow: "hidden" }}>
          {report.logo?.dataUrl ? (
            <img src={report.logo.dataUrl} alt="Logo" style={{ position: "absolute", width: report.logo.size || 64, height: "auto", ...LOGO_POSITIONS[report.logo.position || "top-left"] }} />
          ) : (
            <div style={{ width: 34, height: 34, borderRadius: 9, background: theme.accent, display: "flex", alignItems: "center", justifyContent: "center", position: "absolute", ...LOGO_POSITIONS[report.logo?.position || "top-left"] }}><BarChart3 size={18} color="#fff" /></div>
          )}
          <div style={{ marginTop: 90 }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: theme.accent, letterSpacing: ".08em", textTransform: "uppercase", marginBottom: 14 }}>{theme.name} Report</div>
            <div style={{ fontFamily: headingFamily, fontSize: 40, fontWeight: 600, lineHeight: 1.15, marginBottom: 20 }}>{report.reportTitle || report.title}</div>
            <div style={{ fontSize: 14, color: "#5B6472" }}>{report.company}</div>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#5B6472", borderTop: "1px solid #E3E6EC", paddingTop: 16 }}>
            <span>Prepared by {report.author || "—"}</span>
            <span>{formatDateBySetting(new Date())}</span>
          </div>
        </div>
      )}
      {rest.map((el, i) => (
        <div key={el.id} className="dv-report-page" style={{ width: pageW, minHeight: pageH, marginBottom: 24, padding: pad, display: "flex", flexDirection: "column" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 20, borderBottom: `2px solid ${theme.accent}`, paddingBottom: 10 }}>
            <div style={{ fontFamily: headingFamily, fontSize: 21, fontWeight: 600 }}>{el.title}</div>
            <div style={{ fontSize: 10, color: "#96A0AF" }}>{report.reportTitle}</div>
          </div>
          <ReportElementBody el={el} dataset={dataset} analysis={analysis} charts={charts} theme={theme} />
          <div style={{ marginTop: "auto", paddingTop: 16, display: "flex", justifyContent: "space-between", fontSize: 10, color: "#96A0AF" }}>
            <span>{report.company}</span><span>Page {i + 2}</span>
          </div>
        </div>
      ))}
    </>
  );
}

function ReportElementBody({ el, dataset, analysis, charts, theme }) {
  const bfs = theme.bodyFontSize || 13;
  if (el.type === "kpi") {
    // Specs the user picked in the KPI editor (column + metric + label) are computed live from
    // the dataset here, so edits show up immediately without needing analysis to be recomputed.
    const specs = (el.config.kpis && el.config.kpis.length)
      ? el.config.kpis
      : (analysis?.kpis || []).slice(0, 4).map((k) => ({ id: k.id, field: k.field, metric: k.metric, label: k.label, format: k.format }));
    const cards = specs.map((s) => ({ ...s, value: computeSingleKpiValue(dataset.rows, dataset.schema, s.field, s.metric) }));
    if (!cards.length) return <div style={{ fontSize: bfs, color: "#96A0AF" }}>No KPI cards chosen — pick a column and metric in the editor on the left.</div>;
    return (
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14 }}>
        {cards.map((k) => (
          <div key={k.id} style={{ border: "1px solid #E3E6EC", borderRadius: 10, padding: 14 }}>
            <div style={{ fontSize: bfs - 2.5, color: "#5B6472", fontWeight: 600, marginBottom: 6 }}>{k.label}</div>
            <div className="dv-mono" style={{ fontSize: bfs + 7, fontWeight: 700, color: theme.accent }}>{formatValue(k.value, k.format)}</div>
          </div>
        ))}
      </div>
    );
  }
  if (el.type === "chart") {
    const chart = charts.find((c) => c.id === el.config.chartId);
    if (!chart) return <div style={{ color: "#96A0AF", fontSize: bfs }}>No chart selected — pick one in the field below.</div>;
    const { data, series } = aggregateChart(dataset.rows, dataset.schema, chart);
    const isPie = chart.type === "pie" || chart.type === "donut";
    return (
      <div style={{ breakInside: "avoid", pageBreakInside: "avoid" }}>
        <div style={{ height: isPie ? 320 : 260 }}><FullChart cfg={chart} data={data} series={series} /></div>
        {el.config.caption && <div style={{ fontSize: bfs - 1, lineHeight: 1.6, color: "#5B6472", marginTop: 12 }}>{el.config.caption}</div>}
      </div>
    );
  }
  if (el.type === "text") return <div style={{ fontSize: bfs, lineHeight: 1.7, color: "#344054" }}>{el.config.body}</div>;
  if (el.type === "insights") return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      {(analysis?.insights || []).map((ins, i) => <div key={i} style={{ fontSize: bfs - 0.5, lineHeight: 1.6, paddingLeft: 14, borderLeft: `3px solid ${theme.accent}` }}>{ins.text}</div>)}
    </div>
  );
  if (el.type === "table") {
    const allSafeCols = dataset.schema.filter((s) => !s.sensitive).map((s) => s.name);
    const groupBy = el.config.groupBy && allSafeCols.includes(el.config.groupBy) ? el.config.groupBy : "";
    const limit = el.config.rowLimit ?? 10;
    const sortBy = el.config.sortBy;
    const sortDir = el.config.sortDir || "desc";

    if (groupBy) {
      // Grouped mode: one row per unique value of "groupBy", numeric columns rolled up by the
      // chosen aggregation (sum/avg/count/min/max) — e.g. total Quantity per RegionManager.
      const numCols = dataset.schema.filter((s) => !s.sensitive && ["number", "currency", "percentage"].includes(s.type)).map((s) => s.name);
      const cols = [groupBy, ...(el.config.columns?.length ? el.config.columns.filter((c) => numCols.includes(c)) : numCols)];
      const agg = el.config.aggregation || "sum";
      const aggregate = (vals) => {
        const nums = vals.filter((v) => v !== null);
        if (!nums.length) return null;
        if (agg === "avg") return _.mean(nums);
        if (agg === "count") return nums.length;
        if (agg === "min") return _.min(nums);
        if (agg === "max") return _.max(nums);
        return _.sum(nums);
      };
      const grouped = _.groupBy(dataset.rows, (r) => r[groupBy]);
      let rows = Object.entries(grouped).map(([key, rs]) => {
        const row = { [groupBy]: key };
        cols.slice(1).forEach((c) => {
          const colType = dataset.schema.find((s) => s.name === c)?.type;
          row[c] = aggregate(rs.map((r) => parseNumeric(r[c], colType)));
        });
        return row;
      });
      if (sortBy) rows = _.orderBy(rows, (r) => typeof r[sortBy] === "number" ? r[sortBy] : String(r[sortBy] ?? "").toLowerCase(), sortDir);
      else rows = _.orderBy(rows, (r) => (r[cols[1]] ?? 0), sortDir);
      if (limit > 0) rows = rows.slice(0, limit);
      return (
        <div style={{ overflow: "hidden" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: Math.max(9, bfs - 2.5) }}>
            <thead><tr>{cols.map((c) => <th key={c} style={{ textAlign: "left", padding: "6px 8px", background: "#F5F6FA", fontWeight: 700 }}>{c}</th>)}</tr></thead>
            <tbody>{rows.map((r, i) => <tr key={i}>{cols.map((c) => <td key={c} style={{ padding: "6px 8px", borderBottom: "1px solid #EEF0F4" }}>{typeof r[c] === "number" ? formatValue(r[c], dataset.schema.find((s) => s.name === c)?.type) : String(r[c] ?? "")}</td>)}</tr>)}</tbody>
          </table>
        </div>
      );
    }

    // Never let a stale sensitive column sneak back in even if it was selected before being
    // flagged — always intersect with the current safe list.
    const cols = (el.config.columns?.length ? el.config.columns : allSafeCols.slice(0, 8)).filter((c) => allSafeCols.includes(c));
    let rows = dataset.rows;
    if (sortBy) {
      const sortType = dataset.schema.find((s) => s.name === sortBy)?.type;
      rows = _.orderBy(rows, (r) => sortType === "date" ? Date.parse(r[sortBy]) : (parseNumeric(r[sortBy], sortType) ?? String(r[sortBy]).toLowerCase()), sortDir);
    }
    if (limit > 0) rows = rows.slice(0, limit);
    return (
      <div style={{ overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: Math.max(9, bfs - 2.5) }}>
          <thead><tr>{cols.map((c) => <th key={c} style={{ textAlign: "left", padding: "6px 8px", background: "#F5F6FA", fontWeight: 700 }}>{c}</th>)}</tr></thead>
          <tbody>{rows.map((r, i) => <tr key={i}>{cols.map((c) => <td key={c} style={{ padding: "6px 8px", borderBottom: "1px solid #EEF0F4" }}>{String(r[c])}</td>)}</tr>)}</tbody>
        </table>
      </div>
    );
  }
  if (el.type === "summary") return <div style={{ fontSize: bfs, lineHeight: 1.7, color: "#344054" }}>{el.config.body || "Write your executive summary in the editor on the left."}</div>;
  return null;
}

/* ============================== FILES / REPORTS / SETTINGS ============================== */
function FilesPage({ files, setFiles, setRoute, askConfirm }) {
  return (
    <div className="dv-fade-in" style={{ padding: 28, maxWidth: 1080, margin: "0 auto" }}>
      <div style={{ fontWeight: 700, fontSize: 17, marginBottom: 18 }}>My Files</div>
      {files.length === 0 ? <EmptyState icon={FileSpreadsheet} title="No files yet" subtitle="Upload your first dataset to start analyzing your data." action={<button className="dv-btn dv-btn-primary dv-btn-sm" onClick={() => setRoute("new-analysis")}>Upload data</button>} /> : (
        <div className="dv-card">
          {files.map((f, i) => (
            <div key={f.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 18px", borderBottom: i < files.length - 1 ? "1px solid var(--border-2)" : "none" }}>
              <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
                <IconBox icon={FileSpreadsheet} tone="blue" />
                <div>
                  <div style={{ fontWeight: 600, fontSize: 13.5 }}>{f.name}</div>
                  <div style={{ fontSize: 11.5, color: "var(--text-3)" }}>{f.rowCount} rows · {f.colCount} cols · uploaded {f.uploadDate}</div>
                </div>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => setRoute("new-analysis")}>Open</button>
                <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => setRoute("new-analysis")}>Analyze</button>
                <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => askConfirm("Delete file?", `"${f.name}" will be removed from My Files. This can't be undone.`, () => setFiles((fs) => fs.filter((x) => x.id !== f.id)))}>
                  <Trash2 size={13} color="var(--rose)" />
                </button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ReportsPage({ reports, setReports, setRoute, askConfirm }) {
  return (
    <div className="dv-fade-in" style={{ padding: 28, maxWidth: 1080, margin: "0 auto" }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 18 }}>
        <div style={{ fontWeight: 700, fontSize: 17 }}>My Reports</div>
        <button className="dv-btn dv-btn-primary dv-btn-sm" onClick={() => setRoute("report-builder")}><Plus size={14} /> New Report</button>
      </div>
      <div className="dv-card" style={{ padding: 16, marginBottom: 18, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, background: "var(--blue-dim)", borderColor: "var(--blue)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <IconBox icon={FileText} tone="blue" />
          <div style={{ fontSize: 13, color: "var(--text)" }}>Continue editing your current draft report</div>
        </div>
        <button className="dv-btn dv-btn-primary dv-btn-sm" onClick={() => setRoute("report-builder")}>Open Report Builder <ArrowRight size={13} /></button>
      </div>
      {reports.length === 0 ? <EmptyState icon={FileText} title="No saved reports yet" subtitle="Reports you build get listed here once you preview or export them." /> : (
        <div className="dv-card">
          {reports.map((r, i) => (
            <div key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 18px", borderBottom: i < reports.length - 1 ? "1px solid var(--border-2)" : "none" }}>
              <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
                <IconBox icon={FileText} tone="teal" />
                <div>
                  <div style={{ fontWeight: 600, fontSize: 13.5 }}>{r.title}</div>
                  <div style={{ fontSize: 11.5, color: "var(--text-3)" }}>{r.template} · {r.elements.length} sections</div>
                </div>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => setRoute("report-preview")}>View</button>
                <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => setRoute("report-builder")}>Edit</button>
                <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => askConfirm("Delete report?", `"${r.title}" will be permanently deleted. This can't be undone.`, () => setReports((rs) => rs.filter((x) => x.id !== r.id)))}>
                  <Trash2 size={13} color="var(--rose)" />
                </button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function SettingsPage({ settings, setSettings, theme, setTheme, lang, setLang, setReport }) {
  const [justSaved, setJustSaved] = useState(false);
  const handleSave = () => {
    // Concretely apply Settings to the app: push company/template into the report currently
    // being worked on (previously these two were completely disconnected from each other).
    setReport((r) => ({ ...r, company: settings.company, template: settings.template }));
    setJustSaved(true);
    setTimeout(() => setJustSaved(false), 2500);
  };
  return (
    <div className="dv-fade-in" style={{ padding: 28, maxWidth: 720, margin: "0 auto" }}>
      <div style={{ fontWeight: 700, fontSize: 17, marginBottom: 18 }}>{t(lang, "topbar.settings.title")}</div>
      <div className="dv-card" style={{ padding: 20, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 14, display: "flex", gap: 7, alignItems: "center" }}><Building2 size={15} /> {t(lang, "settings.company")}</div>
        <div style={{ display: "flex", gap: 14, marginBottom: 12 }}>
          <Field label={t(lang, "settings.companyName")}><input className="dv-input" value={settings.company} onChange={(e) => setSettings((s) => ({ ...s, company: e.target.value }))} /></Field>
          <Field label={t(lang, "settings.defaultCurrency")}>
            <select className="dv-input" value={settings.currency} onChange={(e) => setSettings((s) => ({ ...s, currency: e.target.value }))}>
              {CURRENCIES.map((c) => <option key={c.code} value={c.code}>{c.label}</option>)}
            </select>
          </Field>
        </div>
        <div style={{ display: "flex", gap: 14, marginBottom: 16 }}>
          <Field label={t(lang, "settings.defaultDateFormat")}>
            <select className="dv-input" value={settings.dateFormat} onChange={(e) => setSettings((s) => ({ ...s, dateFormat: e.target.value }))}><option>MM/DD/YYYY</option><option>DD/MM/YYYY</option><option>YYYY-MM-DD</option></select>
          </Field>
          <Field label={t(lang, "settings.defaultTemplate")}>
            <select className="dv-input" value={settings.template} onChange={(e) => setSettings((s) => ({ ...s, template: e.target.value }))}>{TEMPLATES.map((tpl) => <option key={tpl.id} value={tpl.id}>{tpl.name}</option>)}</select>
          </Field>
        </div>
        <button className="dv-btn dv-btn-sm" style={{ background: justSaved ? "var(--teal)" : "var(--blue)", color: "#fff" }} onClick={handleSave}>
          {justSaved ? <CheckCircle2 size={13} /> : <Download size={13} style={{ transform: "rotate(180deg)" }} />} {justSaved ? t(lang, "settings.saved") : t(lang, "settings.save")}
        </button>
        <div style={{ fontSize: 10.5, color: "var(--text-3)", marginTop: 8 }}>{t(lang, "settings.savedNote")}</div>
      </div>
      <div className="dv-card" style={{ padding: 20, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 14, display: "flex", gap: 7, alignItems: "center" }}><Palette size={15} /> {t(lang, "settings.appearance")}</div>
        <div style={{ display: "flex", gap: 10 }}>
          <button className="dv-btn dv-btn-sm" style={{ border: `1.5px solid ${theme === "light" ? "var(--blue)" : "var(--border)"}` }} onClick={() => setTheme("light")}><Sun size={13} /> {t(lang, "settings.light")}</button>
          <button className="dv-btn dv-btn-sm" style={{ border: `1.5px solid ${theme === "dark" ? "var(--blue)" : "var(--border)"}` }} onClick={() => setTheme("dark")}><Moon size={13} /> {t(lang, "settings.dark")}</button>
        </div>
      </div>
      <div className="dv-card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 14, display: "flex", gap: 7, alignItems: "center" }}><Globe size={15} /> {t(lang, "settings.language")}</div>
        <div style={{ display: "flex", gap: 10 }}>
          <button className="dv-btn dv-btn-sm" style={{ border: `1.5px solid ${lang === "en" ? "var(--blue)" : "var(--border)"}` }} onClick={() => setLang("en")}>English</button>
          <button className="dv-btn dv-btn-sm" style={{ border: `1.5px solid ${lang === "ar" ? "var(--blue)" : "var(--border)"}` }} onClick={() => setLang("ar")}>العربية</button>
        </div>
        <div style={{ fontSize: 11.5, color: "var(--text-3)", marginTop: 10 }}>{t(lang, "settings.languageNote")}</div>
      </div>
    </div>
  );
}

/* ============================== ROOT APP ============================== */
export default function DataVisionApp() {
  const [route, setRoute] = useState("landing");
  const [theme, setTheme] = useState("light");
  const [lang, setLang] = useState("en");
  const [dataset, setDataset] = useState(null);
  const [analysis, setAnalysis] = useState(null);
  const [charts, setCharts] = useState([]);
  const [editingChartId, setEditingChartId] = useState(null);
  const [files, setFiles] = useState([]);
  const [reports, setReports] = useState([]);
  const [report, setReport] = useState({
    id: uid(), title: "Untitled Report", reportTitle: "Quarterly Business Report", company: "Acme Inc.", author: "",
    template: "executive", orientation: "portrait", status: "draft",
    design: { headingFont: "serif", bodyFontSize: 13, accentColor: null },
    logo: { dataUrl: null, size: 64, position: "top-left" },
    elements: [
      { id: uid(), type: "cover", title: "Cover Page", config: {} },
    ],
  });
  const [settings, setSettings] = useState({ company: "Acme Inc.", currency: "USD", dateFormat: "MM/DD/YYYY", template: "executive" });
  const [confirmState, setConfirmState] = useState(null);
  const askConfirm = (title, message, onConfirm) => setConfirmState({ title, message, onConfirm });

  // Keep the module-level currency (read by formatValue everywhere) in sync with Settings.
  React.useEffect(() => { setCurrentCurrency(settings.currency); }, [settings.currency]);
  React.useEffect(() => { setCurrentDateFormat(settings.dateFormat); }, [settings.dateFormat]);

  const handleFinishAnalysis = useCallback((result) => {
    setAnalysis(result);
    setCharts(result.recommendations);
    if (dataset) {
      setFiles((fs) => {
        const exists = fs.find((f) => f.id === dataset.id);
        const entry = { id: dataset.id, name: dataset.name, rowCount: dataset.rows.length, colCount: dataset.columns.length, analyzed: true, uploadDate: formatDateBySetting(new Date()), chartCount: result.recommendations.length };
        return exists ? fs.map((f) => f.id === dataset.id ? entry : f) : [entry, ...fs];
      });
    }
  }, [dataset]);

  const addChartToReport = (chart) => {
    setCharts((cs) => cs.map((c) => c.id === chart.id ? { ...c, addedToReport: true } : c));
    setReport((r) => {
      // If this exact chart is already placed somewhere in the report, just refresh its
      // title/content instead of inserting a second copy.
      const already = r.elements.find((e) => e.type === "chart" && e.config?.chartId === chart.id);
      if (already) {
        return { ...r, elements: r.elements.map((e) => e.id === already.id ? { ...e, title: chart.title } : e) };
      }
      return { ...r, elements: [...r.elements, { id: uid(), type: "chart", title: chart.title, config: { chartId: chart.id } }] };
    });
  };

  const saveReportSnapshot = () => setReports((rs) => {
    const exists = rs.find((r) => r.id === report.id);
    const entry = { ...report, status: "draft" };
    return exists ? rs.map((r) => r.id === report.id ? entry : r) : [entry, ...rs];
  });

  const goRoute = (r) => {
    if (r === "report-builder" || r === "report-preview") saveReportSnapshot();
    setEditingChartId(null);
    setRoute(r);
  };

  const onCustomizeChart = (chartId) => { setEditingChartId(chartId); setRoute("charts"); };

  const titles = {
    dashboard: [t(lang, "topbar.dashboard.title"), t(lang, "topbar.dashboard.sub")],
    "new-analysis": [t(lang, "topbar.new-analysis.title"), t(lang, "topbar.new-analysis.sub")],
    charts: [t(lang, "topbar.charts.title"), t(lang, "topbar.charts.sub")],
    files: [t(lang, "topbar.files.title"), t(lang, "topbar.files.sub")],
    reports: [t(lang, "topbar.reports.title"), t(lang, "topbar.reports.sub")],
    "report-builder": [t(lang, "topbar.report-builder.title"), t(lang, "topbar.report-builder.sub")],
    "report-preview": [t(lang, "topbar.report-preview.title"), t(lang, "topbar.report-preview.sub")],
    settings: [t(lang, "topbar.settings.title"), t(lang, "topbar.settings.sub")],
  };

  if (route === "landing") {
    return <div className="dv-root" data-theme={theme}><GlobalStyle /><LandingPage onStart={() => setRoute("new-analysis")} /></div>;
  }

  const isFullBleed = (route === "charts" && editingChartId) || route === "report-builder";
  const isPreview = route === "report-preview";

  return (
    <div className="dv-root" data-theme={theme} dir={lang === "ar" ? "rtl" : "ltr"} style={{ height: "100vh", minHeight: 640, display: "flex" }}>
      <GlobalStyle />
      <style>{`@keyframes spin{from{transform:rotate(0)}to{transform:rotate(360deg)}}`}</style>
      <ConfirmDialog state={confirmState} onCancel={() => setConfirmState(null)} />
      <Sidebar route={route} setRoute={goRoute} lang={lang} />
      <div style={{ flex: 1, display: "flex", flexDirection: "column", minWidth: 0 }}>
        {!isFullBleed && !isPreview && <TopBar title={titles[route]?.[0] || ""} subtitle={titles[route]?.[1]} theme={theme} setTheme={setTheme} lang={lang} setLang={setLang} />}
        <div className="dv-scrollbar" style={{ flex: 1, overflowY: isFullBleed ? "hidden" : "auto", minHeight: 0 }}>
          {route === "dashboard" && <Dashboard files={files} reports={reports} setRoute={goRoute} lang={lang} />}
          {route === "new-analysis" && <NewAnalysis dataset={dataset} setDataset={setDataset} onFinish={handleFinishAnalysis} setRoute={goRoute} onAddToReport={addChartToReport} onCustomizeChart={onCustomizeChart} />}
          {route === "charts" && <ChartsPage charts={charts} setCharts={setCharts} dataset={dataset} onAddToReport={addChartToReport} editingId={editingChartId} setEditingId={setEditingChartId} />}
          {route === "files" && <FilesPage files={files} setFiles={setFiles} setRoute={goRoute} askConfirm={askConfirm} />}
          {route === "reports" && <ReportsPage reports={reports} setReports={setReports} setRoute={goRoute} askConfirm={askConfirm} />}
          {route === "report-builder" && <ReportBuilder report={report} setReport={setReport} dataset={dataset} analysis={analysis} charts={charts} setRoute={goRoute} askConfirm={askConfirm} />}
          {route === "report-preview" && <ReportPreview report={report} dataset={dataset} analysis={analysis} charts={charts} setRoute={goRoute} />}
          {route === "settings" && <SettingsPage settings={settings} setSettings={setSettings} theme={theme} setTheme={setTheme} lang={lang} setLang={setLang} setReport={setReport} />}
        </div>
      </div>
    </div>
  );
}
