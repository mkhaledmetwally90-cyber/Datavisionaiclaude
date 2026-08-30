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
  Building2, Palette, Type as TypeIcon, ChevronRight, Info,
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
    @media print{
      body *{visibility:hidden;}
      #dv-print-area, #dv-print-area *{visibility:visible;}
      #dv-print-area{position:absolute;left:0;top:0;width:100%;}
      .dv-report-page{box-shadow:none !important;page-break-after:always;}
    }
  `}</style>
);

/* ============================== DATA HELPERS ============================== */
const uid = () => Math.random().toString(36).slice(2, 10);

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

function buildSchema(rows, columns) {
  return columns.map((col) => {
    const values = rows.map((r) => r[col]);
    const type = detectColumnType(values, col);
    const nonEmpty = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
    const unique = new Set(nonEmpty.map((v) => String(v).trim()));
    return { name: col, type, missing: values.length - nonEmpty.length, unique: unique.size };
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
  if (type === "currency") return "$" + Number(v).toLocaleString(undefined, { maximumFractionDigits: 0 });
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
    ...r, filters: { dateFrom: "", dateTo: "", categories: [], numMin: "", numMax: "", topN: "", bottomN: "" },
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

function aggregateChart(rows, schema, cfg) {
  const xType = (schema.find((s) => s.name === cfg.xField) || {}).type;
  const yType = (schema.find((s) => s.name === cfg.yField) || {}).type || "number";
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
    const data = filtered.map((r) => ({ x: parseNumeric(r[cfg.xField], xType), y: parseNumeric(r[cfg.yField], yType) })).filter((p) => p.x !== null && p.y !== null);
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
const NAV_ITEMS = [
  { id: "dashboard", label: "Dashboard", icon: LayoutGrid },
  { id: "new-analysis", label: "New Analysis", icon: Plus },
  { id: "files", label: "My Files", icon: FileSpreadsheet },
  { id: "charts", label: "Charts", icon: BarChart3 },
  { id: "reports", label: "Reports", icon: FileText },
  { id: "settings", label: "Settings", icon: SettingsIcon },
];

function Sidebar({ route, setRoute }) {
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
            <it.icon size={16} strokeWidth={2.2} /> {it.label}
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
function Dashboard({ files, reports, setRoute }) {
  const t = { en: { title: "Dashboard", sub: "Your data analysis workspace" }, ar: { title: "لوحة التحكم", sub: "مساحة عمل تحليل البيانات" } };
  return (
    <div className="dv-fade-in" style={{ padding: 28, maxWidth: 1140, margin: "0 auto" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 22 }}>
        <div>
          <div style={{ fontSize: 20, fontWeight: 800 }}>Welcome back</div>
          <div style={{ fontSize: 13.5, color: "var(--text-2)" }}>Here's what's happening with your data.</div>
        </div>
        <button className="dv-btn dv-btn-primary" onClick={() => setRoute("new-analysis")}><Plus size={15} /> New Analysis</button>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 14, marginBottom: 24 }}>
        {[
          { label: "Total Analyses", value: files.filter(f => f.analyzed).length, icon: Sparkles, tone: "violet" },
          { label: "Total Reports", value: reports.length, icon: FileText, tone: "teal" },
          { label: "Files Uploaded", value: files.length, icon: FileSpreadsheet, tone: "blue" },
          { label: "Charts Created", value: files.reduce((a, f) => a + (f.chartCount || 0), 0), icon: BarChart3, tone: "amber" },
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
          <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 14 }}>Recent Files</div>
          {files.length === 0 ? (
            <EmptyState icon={FileSpreadsheet} title="No files yet" subtitle="Upload your first dataset to start analyzing your data." action={<button className="dv-btn dv-btn-primary dv-btn-sm" onClick={() => setRoute("new-analysis")}>Upload data</button>} />
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
                  <Badge tone={f.analyzed ? "teal" : "amber"}>{f.analyzed ? "Analyzed" : "Pending"}</Badge>
                </div>
              ))}
            </div>
          )}
        </div>
        <div className="dv-card" style={{ padding: 20 }}>
          <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 14 }}>Recent Reports</div>
          {reports.length === 0 ? (
            <EmptyState icon={FileText} title="No reports yet" subtitle="Create your first report from an analysis." />
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
  const fileInputRef = useRef();
  const pageSize = 8;

  const steps = ["Import", "Preview & Quality", "Analysis", "Chart Recommendations"];

  const parseFile = (file) => {
    setError(""); setLoading(true);
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["csv", "xlsx", "xls"].includes(ext)) {
      setError("Unsupported format. Please upload a .csv, .xlsx or .xls file.");
      setLoading(false);
      return;
    }
    const finish = (rawRows, columns, sheetName) => {
      if (!rawRows.length || !columns.length) {
        setError("This file appears to be empty. Please upload a spreadsheet that contains data.");
        setLoading(false);
        return;
      }
      const rows = normalizeSerialDateColumns(rawRows, columns);
      const schema = buildSchema(rows, columns);
      setDataset({ name: file.name, sheetName: sheetName || "Sheet1", rows, columns, schema, id: uid() });
      setLoading(false);
      setStep(1);
    };
    if (ext === "csv") {
      Papa.parse(file, {
        header: true, skipEmptyLines: true,
        complete: (res) => {
          const columns = res.meta.fields || [];
          finish(res.data, columns, file.name.replace(/\.csv$/i, ""));
        },
        error: () => { setError("Failed to read this CSV file. Please check the format and try again."); setLoading(false); },
      });
    } else {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
          const sheetName = wb.SheetNames[0];
          const ws = wb.Sheets[sheetName];
          // raw:false renders each cell using its Excel number format, so dates/currency/percentages
          // come through as readable strings ("8/29/2026", "$1,234", "12%") instead of raw serials —
          // which is what our column-type detector expects.
          const json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false, dateNF: "yyyy-mm-dd" });
          const columns = json.length ? Object.keys(json[0]) : [];
          finish(json, columns, sheetName);
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

  const kpis = useMemo(() => dataset ? computeKpis(dataset.rows, dataset.schema) : [], [dataset]);
  const insights = useMemo(() => dataset ? computeInsights(dataset.rows, dataset.schema) : [], [dataset]);
  const recommendations = useMemo(() => dataset ? buildRecommendations(dataset.schema) : [], [dataset]);

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
              <button className="dv-btn dv-btn-dark" style={{ width: "100%", justifyContent: "center" }} onClick={() => setError("__gsheet__")}>Connect Google Sheet</button>
            </div>
          </div>
          {error === "__gsheet__" && (
            <div className="dv-card" style={{ padding: 14, borderColor: "var(--amber)", background: "var(--amber-dim)", display: "flex", gap: 10, fontSize: 13, color: "var(--text)", marginBottom: 14 }}>
              <Info size={16} color="var(--amber)" style={{ flexShrink: 0, marginTop: 1 }} />
              <div>Live Google Sheets import requires a connected backend (Google Sheets API) that isn't available in this prototype. The UI and data model are ready for it — for now, please export your sheet as .csv or .xlsx and upload it on the left.</div>
            </div>
          )}
          {error && error !== "__gsheet__" && (
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
            <div className="dv-scrollbar" style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                <thead>
                  <tr>
                    {dataset.columns.filter((c) => !hiddenCols.includes(c)).map((c) => {
                      const type = dataset.schema.find((s) => s.name === c)?.type;
                      return (
                        <th key={c} onClick={() => { setSortCol(c); setSortDir(sortCol === c && sortDir === "asc" ? "desc" : "asc"); }}
                          style={{ textAlign: "left", padding: "8px 10px", borderBottom: "1px solid var(--border)", cursor: "pointer", whiteSpace: "nowrap", color: "var(--text-2)", fontWeight: 700 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                            {c} {sortCol === c && (sortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
                            <span style={{ fontWeight: 500, color: "var(--text-3)", fontSize: 10 }}>· {type}</span>
                          </div>
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {pagedRows.map((r, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid var(--border-2)" }}>
                      {dataset.columns.filter((c) => !hiddenCols.includes(c)).map((c) => (
                        <td key={c} className="dv-mono" style={{ padding: "8px 10px", whiteSpace: "nowrap", color: String(r[c]).trim() === "" ? "var(--rose)" : "var(--text)" }}>
                          {String(r[c]).trim() === "" ? "missing" : String(r[c])}
                        </td>
                      ))}
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
          const blank = { id: uid(), type: "bar", title: "New Chart", subtitle: "", explanation: "", xField: dataset.columns[0], yField: dataset.columns[1] || dataset.columns[0], groupBy: "", aggregation: "sum", filters: { dateFrom: "", dateTo: "", categories: [], numMin: "", numMax: "", topN: "", bottomN: "" }, appearance: { legend: true, dataLabels: false, gridlines: true, orientation: "vertical", numberFormat: "number", colorPalette: "classic", fontSize: 12 }, addedToReport: false };
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
  const numericCols = dataset.schema.filter((c) => ["number", "currency", "percentage"].includes(c.type)).map((c) => c.name);
  const categoryCols = dataset.schema.filter((c) => c.type === "text").map((c) => c.name);
  const allCols = dataset.columns;
  const xColType = dataset.schema.find((s) => s.name === cfg.xField)?.type;
  const uniqueXVals = useMemo(() => _.uniq(dataset.rows.map((r) => String(r[cfg.xField]))).slice(0, 30), [dataset, cfg.xField]);

  return (
    <div className="dv-fade-in" style={{ display: "grid", gridTemplateColumns: "320px 1fr", height: "100%" }}>
      <div className="dv-scrollbar" style={{ borderRight: "1px solid var(--border)", padding: 20, overflowY: "auto" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Chart Editor</div>
          <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={onCancel}><X size={14} /></button>
        </div>

        <Section title="Chart Type">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6 }}>
            {["bar", "line", "area", "pie", "donut", "scatter", "stacked-bar", "grouped-bar"].map((t) => (
              <button key={t} onClick={() => set("type", t)} className="dv-btn dv-btn-sm" style={{ justifyContent: "center", background: cfg.type === t ? "var(--blue)" : "var(--surface)", color: cfg.type === t ? "#fff" : "var(--text-2)", border: "1px solid var(--border)" }}>{t}</button>
            ))}
          </div>
        </Section>

        <Section title="Data">
          <Field label="X-Axis">
            <select className="dv-input" value={cfg.xField} onChange={(e) => set("xField", e.target.value)}>{allCols.map((c) => <option key={c} value={c}>{c}</option>)}</select>
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
            <div style={{ display: "flex", gap: 8 }}>
              <Field label="From"><input type="date" className="dv-input" value={cfg.filters.dateFrom} onChange={(e) => set("filters.dateFrom", e.target.value)} /></Field>
              <Field label="To"><input type="date" className="dv-input" value={cfg.filters.dateTo} onChange={(e) => set("filters.dateTo", e.target.value)} /></Field>
            </div>
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
          <div style={{ fontWeight: 700, fontSize: 17 }}>{cfg.title}</div>
          {cfg.subtitle && <div style={{ fontSize: 13, color: "var(--text-2)", marginBottom: 14 }}>{cfg.subtitle}</div>}
          <div style={{ height: 380, marginTop: 14 }}>
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
  const labelFmt = (v) => cfg.appearance.numberFormat === "currency" ? "$" + Number(v).toLocaleString() : cfg.appearance.numberFormat === "percentage" ? v + "%" : Number(v).toLocaleString();
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

function ReportBuilder({ report, setReport, dataset, analysis, charts, setRoute }) {
  const [zoom, setZoom] = useState(0.34);
  const addElement = (type) => {
    const el = { id: uid(), type, title: SECTION_LIBRARY.find((s) => s.type === type)?.label, config: {} };
    // Chart sections start unassigned — the user must explicitly pick which chart shows here,
    // rather than silently defaulting to whichever chart happens to be first in the list.
    if (type === "chart") el.config.chartId = null;
    if (type === "kpi") el.config.kpiIds = (analysis?.kpis || []).slice(0, 4).map((k) => k.id);
    if (type === "text") el.config.body = "Add your commentary here…";
    if (type === "summary") el.config.body = "";
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

      <div style={{ flex: 1, display: "grid", gridTemplateColumns: "420px 1fr", minHeight: 0 }}>
        <div className="dv-scrollbar" style={{ overflowY: "auto", padding: 20, borderRight: "1px solid var(--border)" }}>
          <Section title="Template">
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {TEMPLATES.map((t) => (
                <button key={t.id} onClick={() => setReport((r) => ({ ...r, template: t.id }))} className="dv-btn dv-btn-sm" style={{ border: `1.5px solid ${report.template === t.id ? t.accent : "var(--border)"}`, background: report.template === t.id ? t.accent + "18" : "var(--surface)", color: "var(--text)", flexDirection: "column", alignItems: "flex-start", height: "auto", padding: "7px 10px" }}>
                  <span style={{ fontWeight: 700 }}>{t.name}</span><span style={{ fontSize: 10, color: "var(--text-2)", fontWeight: 400 }}>{t.desc}</span>
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
                    <button className="dv-btn dv-btn-ghost dv-btn-sm" onClick={() => remove(el.id)}><Trash2 size={12} color="var(--rose)" /></button>
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
      <select className="dv-input" value={el.config.chartId || ""} onChange={(e) => {
        const chosen = charts.find((c) => c.id === e.target.value);
        updateEl(el.id, { config: { ...el.config, chartId: e.target.value }, title: chosen ? chosen.title : el.title });
      }}>
        <option value="">Select a chart…</option>
        {charts.map((c) => <option key={c.id} value={c.id}>{c.title}</option>)}
      </select>
    );
  }
  if (el.type === "text") {
    return <textarea className="dv-input" rows={3} value={el.config.body} onChange={(e) => updateEl(el.id, { config: { ...el.config, body: e.target.value } })} />;
  }
  if (el.type === "kpi") {
    return <div style={{ fontSize: 12, color: "var(--text-2)" }}>{(analysis?.kpis || []).slice(0, 4).map((k) => k.label).join(" · ") || "No KPIs available"}</div>;
  }
  if (el.type === "table") {
    return <div style={{ fontSize: 12, color: "var(--text-2)" }}>Shows the first 10 rows of "{dataset.name}" ({dataset.columns.length} columns)</div>;
  }
  if (el.type === "insights") {
    return <div style={{ fontSize: 12, color: "var(--text-2)" }}>{(analysis?.insights || []).length} automatic insight(s) will be listed here</div>;
  }
  if (el.type === "summary") {
    return <textarea className="dv-input" rows={3} placeholder="Write a short overview of overall performance, key findings and recommendations…" value={el.config.body || ""} onChange={(e) => updateEl(el.id, { config: { ...el.config, body: e.target.value } })} />;
  }
  return <div style={{ fontSize: 12, color: "var(--text-2)" }}>Cover page with report title, company and author</div>;
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
        <div className="dv-report-page" style={{ width: pageW, height: pageH, marginBottom: 24, padding: pad, display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
          <div>
            <div style={{ width: 34, height: 34, borderRadius: 9, background: theme.accent, display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 40 }}><BarChart3 size={18} color="#fff" /></div>
            <div style={{ fontSize: 12, fontWeight: 700, color: theme.accent, letterSpacing: ".08em", textTransform: "uppercase", marginBottom: 14 }}>{theme.name} Report</div>
            <div style={{ fontFamily: headingFamily, fontSize: 40, fontWeight: 600, lineHeight: 1.15, marginBottom: 20 }}>{report.reportTitle || report.title}</div>
            <div style={{ fontSize: 14, color: "#5B6472" }}>{report.company}</div>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#5B6472", borderTop: "1px solid #E3E6EC", paddingTop: 16 }}>
            <span>Prepared by {report.author || "—"}</span>
            <span>{new Date().toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" })}</span>
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
    const kpis = (analysis?.kpis || []).filter((k) => (el.config.kpiIds || []).includes(k.id));
    return (
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14 }}>
        {(kpis.length ? kpis : (analysis?.kpis || []).slice(0, 4)).map((k) => (
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
    return <div style={{ height: isPie ? 320 : 260 }}><FullChart cfg={chart} data={data} series={series} /></div>;
  }
  if (el.type === "text") return <div style={{ fontSize: bfs, lineHeight: 1.7, color: "#344054" }}>{el.config.body}</div>;
  if (el.type === "insights") return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      {(analysis?.insights || []).map((ins, i) => <div key={i} style={{ fontSize: bfs - 0.5, lineHeight: 1.6, paddingLeft: 14, borderLeft: `3px solid ${theme.accent}` }}>{ins.text}</div>)}
    </div>
  );
  if (el.type === "table") return (
    <div style={{ overflow: "hidden" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: Math.max(9, bfs - 2.5) }}>
        <thead><tr>{dataset.columns.slice(0, 8).map((c) => <th key={c} style={{ textAlign: "left", padding: "6px 8px", background: "#F5F6FA", fontWeight: 700 }}>{c}</th>)}</tr></thead>
        <tbody>{dataset.rows.slice(0, 10).map((r, i) => <tr key={i}>{dataset.columns.slice(0, 8).map((c) => <td key={c} style={{ padding: "6px 8px", borderBottom: "1px solid #EEF0F4" }}>{String(r[c])}</td>)}</tr>)}</tbody>
      </table>
    </div>
  );
  if (el.type === "summary") return <div style={{ fontSize: bfs, lineHeight: 1.7, color: "#344054" }}>{el.config.body || "Write your executive summary in the editor on the left."}</div>;
  return null;
}

/* ============================== FILES / REPORTS / SETTINGS ============================== */
function FilesPage({ files, setRoute }) {
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
                <button className="dv-btn dv-btn-ghost dv-btn-sm"><Trash2 size={13} color="var(--rose)" /></button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ReportsPage({ reports, setRoute }) {
  return (
    <div className="dv-fade-in" style={{ padding: 28, maxWidth: 1080, margin: "0 auto" }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 18 }}>
        <div style={{ fontWeight: 700, fontSize: 17 }}>My Reports</div>
        <button className="dv-btn dv-btn-primary dv-btn-sm" onClick={() => setRoute("report-builder")}><Plus size={14} /> New Report</button>
      </div>
      {reports.length === 0 ? <EmptyState icon={FileText} title="No reports yet" subtitle="Create your first report from an analysis." action={<button className="dv-btn dv-btn-primary dv-btn-sm" onClick={() => setRoute("report-builder")}>Build a report</button>} /> : (
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
                <button className="dv-btn dv-btn-ghost dv-btn-sm"><Copy size={13} /></button>
                <button className="dv-btn dv-btn-ghost dv-btn-sm"><Trash2 size={13} color="var(--rose)" /></button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function SettingsPage({ settings, setSettings, theme, setTheme, lang, setLang }) {
  return (
    <div className="dv-fade-in" style={{ padding: 28, maxWidth: 720, margin: "0 auto" }}>
      <div style={{ fontWeight: 700, fontSize: 17, marginBottom: 18 }}>Settings</div>
      <div className="dv-card" style={{ padding: 20, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 14, display: "flex", gap: 7, alignItems: "center" }}><Building2 size={15} /> Company</div>
        <div style={{ display: "flex", gap: 14, marginBottom: 12 }}>
          <Field label="Company name"><input className="dv-input" value={settings.company} onChange={(e) => setSettings((s) => ({ ...s, company: e.target.value }))} /></Field>
          <Field label="Default currency">
            <select className="dv-input" value={settings.currency} onChange={(e) => setSettings((s) => ({ ...s, currency: e.target.value }))}><option>USD</option><option>EUR</option><option>GBP</option><option>SAR</option></select>
          </Field>
        </div>
        <div style={{ display: "flex", gap: 14 }}>
          <Field label="Default date format">
            <select className="dv-input" value={settings.dateFormat} onChange={(e) => setSettings((s) => ({ ...s, dateFormat: e.target.value }))}><option>MM/DD/YYYY</option><option>DD/MM/YYYY</option><option>YYYY-MM-DD</option></select>
          </Field>
          <Field label="Default report template">
            <select className="dv-input" value={settings.template} onChange={(e) => setSettings((s) => ({ ...s, template: e.target.value }))}>{TEMPLATES.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}</select>
          </Field>
        </div>
      </div>
      <div className="dv-card" style={{ padding: 20, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 14, display: "flex", gap: 7, alignItems: "center" }}><Palette size={15} /> Appearance</div>
        <div style={{ display: "flex", gap: 10 }}>
          <button className="dv-btn dv-btn-sm" style={{ border: `1.5px solid ${theme === "light" ? "var(--blue)" : "var(--border)"}` }} onClick={() => setTheme("light")}><Sun size={13} /> Light</button>
          <button className="dv-btn dv-btn-sm" style={{ border: `1.5px solid ${theme === "dark" ? "var(--blue)" : "var(--border)"}` }} onClick={() => setTheme("dark")}><Moon size={13} /> Dark</button>
        </div>
      </div>
      <div className="dv-card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 14, display: "flex", gap: 7, alignItems: "center" }}><Globe size={15} /> Language</div>
        <div style={{ display: "flex", gap: 10 }}>
          <button className="dv-btn dv-btn-sm" style={{ border: `1.5px solid ${lang === "en" ? "var(--blue)" : "var(--border)"}` }} onClick={() => setLang("en")}>English</button>
          <button className="dv-btn dv-btn-sm" style={{ border: `1.5px solid ${lang === "ar" ? "var(--blue)" : "var(--border)"}` }} onClick={() => setLang("ar")}>العربية</button>
        </div>
        <div style={{ fontSize: 11.5, color: "var(--text-3)", marginTop: 10 }}>Full UI translation is architected but not yet complete in this prototype — the toggle flips text direction and the top bar label.</div>
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
    elements: [
      { id: uid(), type: "cover", title: "Cover Page", config: {} },
    ],
  });
  const [settings, setSettings] = useState({ company: "Acme Inc.", currency: "USD", dateFormat: "MM/DD/YYYY", template: "executive" });

  const handleFinishAnalysis = useCallback((result) => {
    setAnalysis(result);
    setCharts(result.recommendations);
    if (dataset) {
      setFiles((fs) => {
        const exists = fs.find((f) => f.id === dataset.id);
        const entry = { id: dataset.id, name: dataset.name, rowCount: dataset.rows.length, colCount: dataset.columns.length, analyzed: true, uploadDate: new Date().toLocaleDateString(), chartCount: result.recommendations.length };
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
    dashboard: ["Dashboard", "Your data analysis workspace"],
    "new-analysis": ["New Analysis", "Import, preview and analyze your dataset"],
    charts: ["Charts", "Recommended and customized visualizations"],
    files: ["My Files", "Datasets you've uploaded"],
    reports: ["My Reports", "Reports you've generated"],
    "report-builder": ["Report Builder", "Assemble your report"],
    "report-preview": ["Report Preview", "Exactly how your PDF will look"],
    settings: ["Settings", "Preferences and defaults"],
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
      <Sidebar route={route} setRoute={goRoute} />
      <div style={{ flex: 1, display: "flex", flexDirection: "column", minWidth: 0 }}>
        {!isFullBleed && !isPreview && <TopBar title={titles[route]?.[0] || ""} subtitle={titles[route]?.[1]} theme={theme} setTheme={setTheme} lang={lang} setLang={setLang} />}
        <div className="dv-scrollbar" style={{ flex: 1, overflowY: isFullBleed ? "hidden" : "auto", minHeight: 0 }}>
          {route === "dashboard" && <Dashboard files={files} reports={reports} setRoute={goRoute} />}
          {route === "new-analysis" && <NewAnalysis dataset={dataset} setDataset={setDataset} onFinish={handleFinishAnalysis} setRoute={goRoute} onAddToReport={addChartToReport} onCustomizeChart={onCustomizeChart} />}
          {route === "charts" && <ChartsPage charts={charts} setCharts={setCharts} dataset={dataset} onAddToReport={addChartToReport} editingId={editingChartId} setEditingId={setEditingChartId} />}
          {route === "files" && <FilesPage files={files} setRoute={goRoute} />}
          {route === "reports" && <ReportsPage reports={reports} setRoute={goRoute} />}
          {route === "report-builder" && <ReportBuilder report={report} setReport={setReport} dataset={dataset} analysis={analysis} charts={charts} setRoute={goRoute} />}
          {route === "report-preview" && <ReportPreview report={report} dataset={dataset} analysis={analysis} charts={charts} setRoute={goRoute} />}
          {route === "settings" && <SettingsPage settings={settings} setSettings={setSettings} theme={theme} setTheme={setTheme} lang={lang} setLang={setLang} />}
        </div>
      </div>
    </div>
  );
}
