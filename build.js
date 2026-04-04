#!/usr/bin/env node
/**
 * FAO Ukraine KPI Dashboard Builder
 * Reads the SharePoint CSV export and generates a self-contained HTML dashboard.
 *
 * Usage:
 *   node build.js                     # uses OneDrive paths
 *   node build.js --ci                # uses ./data/... and ./public/index.html
 *   node build.js --input x.csv --output out.html
 */

import fs   from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);

// ── CLI argument parsing ──────────────────────────────────────────────────────

const args = process.argv.slice(2);
const hasFlag  = (f) => args.includes(f);
const flagVal  = (f) => { const i = args.indexOf(f); return i !== -1 ? args[i + 1] : null; };

const CI = hasFlag('--ci');

const DEFAULT_INPUT  = CI
  ? path.resolve(__dirname, 'data', 'Procurement_Tracking_FAO_Ukraine.csv')
  : '/Users/francesco/Library/CloudStorage/OneDrive-FoodandAgricultureOrganization/Ukraine/KPI dashboard/data/Procurement_Tracking_FAO_Ukraine.csv';

const DEFAULT_OUTPUT = CI
  ? path.resolve(__dirname, 'public', 'index.html')
  : '/Users/francesco/Library/CloudStorage/OneDrive-FoodandAgricultureOrganization/Ukraine/KPI dashboard/index.html';

const INPUT_PATH  = flagVal('--input')  || DEFAULT_INPUT;
const OUTPUT_PATH = flagVal('--output') || DEFAULT_OUTPUT;

// ── CSV parser ────────────────────────────────────────────────────────────────
// Handles RFC-4180 quoting, including fields that contain commas, newlines,
// or escaped double-quotes (represented as "" inside a quoted field).

function parseCSV(text) {
  const rows   = [];
  let   col    = 0;
  let   inQuote = false;
  let   field  = '';
  let   row    = [];

  // Normalise line endings
  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  for (let i = 0; i < text.length; i++) {
    const ch   = text[i];
    const next = text[i + 1];

    if (inQuote) {
      if (ch === '"' && next === '"') {
        field += '"';
        i++;
      } else if (ch === '"') {
        inQuote = false;
      } else {
        field += ch;
      }
    } else {
      if (ch === '"') {
        inQuote = true;
      } else if (ch === ',') {
        row.push(field);
        field = '';
        col++;
      } else if (ch === '\n') {
        row.push(field);
        rows.push(row);
        row   = [];
        field = '';
        col   = 0;
      } else {
        field += ch;
      }
    }
  }
  // Final field / row
  if (field || row.length) {
    row.push(field);
    if (row.some(c => c !== '')) rows.push(row);
  }

  if (rows.length < 2) return [];

  const headers = rows[0];
  return rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, idx) => { obj[h.trim()] = (r[idx] ?? '').trim(); });
    return obj;
  });
}

// ── Field-type parsers ────────────────────────────────────────────────────────

function parseLookup(raw) {
  if (!raw) return '';
  try {
    const v = raw.trim();
    if (v.startsWith('{')) {
      const obj = JSON.parse(v);
      return obj.Value ?? '';
    }
    if (v.startsWith('[')) {
      const arr = JSON.parse(v);
      return arr.map(x => x.Value ?? '').filter(Boolean).join('; ');
    }
  } catch (_) { /* fall through */ }
  return raw;
}

function parseMultiLookup(raw) {
  if (!raw) return '';
  try {
    const v = raw.trim();
    if (v.startsWith('[')) {
      const arr = JSON.parse(v);
      return arr.map(x => x.Value ?? '').filter(Boolean).join('; ');
    }
    if (v.startsWith('{')) {
      const obj = JSON.parse(v);
      return obj.Value ?? '';
    }
  } catch (_) { /* fall through */ }
  return raw;
}

function parseUser(raw) {
  if (!raw) return '';
  try {
    const v = raw.trim();
    if (v.startsWith('{')) {
      const obj = JSON.parse(v);
      let name = obj.DisplayName ?? '';
      // Strip suffixes like " (FAOUA)", " (CSLP)", " (REU)"
      name = name.replace(/\s*\([^)]+\)\s*$/, '').trim();
      // "Surname, Firstname" → take surname
      if (name.includes(',')) {
        name = name.split(',')[0].trim();
      }
      return name;
    }
  } catch (_) { /* fall through */ }
  return raw;
}

function parseFloat2(raw) {
  if (!raw) return null;
  const n = parseFloat(String(raw).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

function parseDate(raw) {
  if (!raw) return null;
  // Accept YYYY-MM-DD or ISO datetime
  const m = String(raw).match(/^(\d{4}-\d{2}-\d{2})/);
  if (!m) return null;
  const d = new Date(m[1] + 'T00:00:00Z');
  return isNaN(d.getTime()) ? null : d;
}

function dateDiffDays(a, b) {
  if (!a || !b) return null;
  return Math.round((b - a) / 86400000);
}

function fmtDate(d) {
  if (!d) return '';
  return d.toISOString().slice(0, 10);
}

// ── Data loading & normalisation ──────────────────────────────────────────────

function loadData(csvPath) {
  console.log(`Reading CSV: ${csvPath}`);
  const raw = fs.readFileSync(csvPath, 'utf8');
  const rows = parseCSV(raw);
  console.log(`  Parsed ${rows.length} rows`);

  return rows.map(r => {
    const buyer         = parseUser(r['Buyer']);
    const prValue       = parseFloat2(r['PRValue']);
    const cumulativePO  = parseFloat2(r['CumulativePO_x0024_']);
    const savings       = parseFloat2(r['Savings']);
    const method        = parseLookup(r['SollicitationMethod']);
    const stage         = parseLookup(r['ProcurementStage']);
    const planRaw       = parseLookup(r['PartofProcurementPlan']);
    const awardBasis    = parseLookup(r['AwardBasis']);
    const marketCat     = parseMultiLookup(r['MarketCategory']);
    const projRef       = parseMultiLookup(r['ProjectReference']);

    const prReceived    = parseDate(r['PRReceived']);
    const poDate        = parseDate(r['POIssuancedate']);
    const solIssued     = parseDate(r['SollicitationIssued']);
    const solClosed     = parseDate(r['SolicitationClosed']);
    const created       = parseDate(r['Created']);
    const modified      = parseDate(r['Modified']);
    const dateAssigned  = parseDate(r['PRAssigned']);
    const dateClosed    = parseDate(r['DateClosed']);

    const cycleTime     = dateDiffDays(prReceived, poDate);
    const year          = prReceived ? prReceived.getUTCFullYear() : (poDate ? poDate.getUTCFullYear() : null);

    // Competitive vs Direct
    const COMPETITIVE_METHODS = ['ITB', 'RFQ', 'RFP', 'BAFO', 'EOI'];
    const methodUpper = (method || '').toUpperCase();
    const isCompetitive = COMPETITIVE_METHODS.some(m => methodUpper.includes(m));

    // Plan compliance bucket
    let planBucket;
    if (/yes.*part.*procurement.*plan/i.test(planRaw)) planBucket = 'Planned';
    else if (/not planned/i.test(planRaw))             planBucket = 'Unplanned';
    else                                               planBucket = 'N/A';

    return {
      id:            r['ID'] || r['ItemInternalId'] || '',
      title:         r['Title'] || r['Description'] || '',
      prgrms:        r['PRGRMS_x0023_'] || '',
      buyer,
      stage,
      status:        (r['Status'] || '').trim(),
      prValue,
      cumulativePO,
      savings,
      method,
      marketCat,
      projRef,
      awardBasis,
      planRaw,
      planBucket,
      isCompetitive,
      prReceived,
      poDate,
      solIssued,
      solClosed,
      created,
      modified,
      dateAssigned,
      dateClosed,
      cycleTime,
      year,
      poNumber: r['PO_x0023_'] || '',
      buyingUnit: parseLookup(r['BuyingUnit']),
    };
  });
}

// ── KPI computations ──────────────────────────────────────────────────────────

function computeKPIs(rows) {
  // ── KPI 1: Cycle Time ──────────────────────────────────────────────────────
  const cycleByMethod = {};
  for (const r of rows) {
    if (r.cycleTime === null || r.cycleTime < 0) continue;
    const m = r.method || 'Unknown';
    if (!cycleByMethod[m]) cycleByMethod[m] = [];
    cycleByMethod[m].push(r.cycleTime);
  }
  const kpi1 = Object.entries(cycleByMethod).map(([method, times]) => {
    const avg = times.reduce((a, b) => a + b, 0) / times.length;
    return {
      method,
      avg: Math.round(avg),
      count: times.length,
      color: avg <= 30 ? '#4CAF50' : avg <= 90 ? '#FF9800' : '#F44336',
    };
  }).sort((a, b) => b.count - a.count);

  // Overall cycle time stats
  const allCycles = rows.filter(r => r.cycleTime !== null && r.cycleTime >= 0).map(r => r.cycleTime);
  const avgCycle  = allCycles.length ? Math.round(allCycles.reduce((a, b) => a + b, 0) / allCycles.length) : 0;

  // ── KPI 2: Savings ────────────────────────────────────────────────────────
  let sumPos = 0, sumNeg = 0;
  const savingsRows = rows.filter(r => r.savings !== null);
  for (const r of savingsRows) {
    if (r.savings > 0) sumPos += r.savings;
    else if (r.savings < 0) sumNeg += r.savings;
  }
  const top10Savings = [...savingsRows]
    .sort((a, b) => Math.abs(b.savings) - Math.abs(a.savings))
    .slice(0, 10)
    .map(r => ({
      id:      r.id,
      title:   r.title.slice(0, 60) + (r.title.length > 60 ? '…' : ''),
      savings: r.savings,
      buyer:   r.buyer,
    }));

  const kpi2 = { sumPos, sumNeg, net: sumPos + sumNeg, top10: top10Savings };

  // ── KPI 3: Competitive vs Direct ──────────────────────────────────────────
  let compCount = 0, compValue = 0, dirCount = 0, dirValue = 0;
  for (const r of rows) {
    if (r.isCompetitive) {
      compCount++;
      if (r.prValue) compValue += r.prValue;
    } else {
      dirCount++;
      if (r.prValue) dirValue += r.prValue;
    }
  }
  const kpi3 = { compCount, compValue, dirCount, dirValue };

  // ── KPI 4: Plan Compliance ────────────────────────────────────────────────
  const planCounts = { Planned: 0, Unplanned: 0, 'N/A': 0 };
  const planValues = { Planned: 0, Unplanned: 0, 'N/A': 0 };
  for (const r of rows) {
    planCounts[r.planBucket]++;
    if (r.prValue) planValues[r.planBucket] += r.prValue;
  }
  const kpi4 = { planCounts, planValues };

  // ── KPI 5: Buyer Workload ─────────────────────────────────────────────────
  const buyerMap = {};
  for (const r of rows) {
    const b = r.buyer || 'Unknown';
    if (!buyerMap[b]) buyerMap[b] = { count: 0, value: 0 };
    buyerMap[b].count++;
    if (r.prValue) buyerMap[b].value += r.prValue;
  }
  const kpi5 = Object.entries(buyerMap)
    .map(([buyer, d]) => ({ buyer, ...d }))
    .sort((a, b) => b.count - a.count);

  // ── Pipeline: count per stage ─────────────────────────────────────────────
  const pipelineMap = {};
  for (const r of rows) {
    const s = r.stage || 'Unknown';
    if (!pipelineMap[s]) pipelineMap[s] = { count: 0, value: 0 };
    pipelineMap[s].count++;
    if (r.prValue) pipelineMap[s].value += r.prValue;
  }
  // Sort by a logical procurement progression
  const STAGE_ORDER = [
    'New PR', 'Assigned', 'Solicitation Issued', 'Offers Received',
    'Evaluation', 'Award Recommendation', 'PO Issued', 'Closed', 'Cancelled',
  ];
  const pipeline = Object.entries(pipelineMap)
    .map(([stage, d]) => ({ stage, ...d }))
    .sort((a, b) => {
      const ia = STAGE_ORDER.findIndex(s => a.stage.toLowerCase().includes(s.toLowerCase()));
      const ib = STAGE_ORDER.findIndex(s => b.stage.toLowerCase().includes(s.toLowerCase()));
      if (ia !== -1 && ib !== -1) return ia - ib;
      if (ia !== -1) return -1;
      if (ib !== -1) return 1;
      return b.count - a.count;
    });

  // ── Year list ─────────────────────────────────────────────────────────────
  const years = [...new Set(rows.map(r => r.year).filter(Boolean))].sort();

  // ── Buyer list ────────────────────────────────────────────────────────────
  const buyers = [...new Set(rows.map(r => r.buyer).filter(Boolean))].sort();

  // ── Stage list ────────────────────────────────────────────────────────────
  const stages = [...new Set(rows.map(r => r.stage).filter(Boolean))].sort();

  // ── Last modified date ────────────────────────────────────────────────────
  const modDates = rows.map(r => r.modified).filter(Boolean);
  const lastUpdated = modDates.length
    ? fmtDate(new Date(Math.max(...modDates.map(d => d.getTime()))))
    : fmtDate(new Date());

  // ── Serialisable rows (dates as strings) ─────────────────────────────────
  const serialRows = rows.map(r => ({
    ...r,
    prReceived:   fmtDate(r.prReceived),
    poDate:       fmtDate(r.poDate),
    solIssued:    fmtDate(r.solIssued),
    solClosed:    fmtDate(r.solClosed),
    created:      fmtDate(r.created),
    modified:     fmtDate(r.modified),
    dateAssigned: fmtDate(r.dateAssigned),
    dateClosed:   fmtDate(r.dateClosed),
  }));

  return { kpi1, kpi2, kpi3, kpi4, kpi5, pipeline, years, buyers, stages, lastUpdated, rows: serialRows, avgCycle };
}

// ── HTML generator ────────────────────────────────────────────────────────────

function generateHTML(data) {
  const dataJson = JSON.stringify(data);

  return `<!DOCTYPE html>
<html lang="en" class="light">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>FAO Ukraine – Procurement KPI Dashboard</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js" crossorigin="anonymous"></script>
  <script src="https://cdn.tailwindcss.com"></script>
  <script>
    tailwind.config = {
      darkMode: 'class',
      theme: {
        extend: {
          colors: {
            fao: {
              blue:  '#009FDA',
              navy:  '#003366',
              green: '#4CAF50',
              amber: '#FF9800',
              red:   '#F44336',
            }
          }
        }
      }
    };
  </script>
  <style>
    :root {
      --fao-blue:  #009FDA;
      --fao-navy:  #003366;
      --fao-green: #4CAF50;
      --fao-amber: #FF9800;
      --fao-red:   #F44336;
    }
    body { font-family: 'Segoe UI', Arial, sans-serif; }
    .dark body { background: #0f172a; color: #e2e8f0; }
    canvas { max-height: 320px; }
    .chart-card { position: relative; }
    .export-btn {
      position: absolute; top: 0.5rem; right: 0.5rem;
      font-size: 0.7rem; padding: 2px 8px;
      background: rgba(0,159,218,0.1); border: 1px solid #009FDA;
      color: #009FDA; border-radius: 4px; cursor: pointer;
      transition: background 0.2s;
    }
    .export-btn:hover { background: rgba(0,159,218,0.25); }
    .kpi-card { transition: transform 0.15s; }
    .kpi-card:hover { transform: translateY(-2px); }
    /* Modal */
    #modal-overlay {
      display: none; position: fixed; inset: 0;
      background: rgba(0,0,0,0.5); z-index: 50;
      align-items: center; justify-content: center;
    }
    #modal-overlay.open { display: flex; }
    #modal-box {
      background: white; border-radius: 12px;
      max-width: 90vw; width: 700px; max-height: 85vh;
      overflow-y: auto; padding: 1.5rem; position: relative;
    }
    .dark #modal-box { background: #1e293b; color: #e2e8f0; }
    #modal-close {
      position: absolute; top: 0.75rem; right: 1rem;
      font-size: 1.4rem; cursor: pointer; color: #64748b;
    }
    table.drill { width: 100%; border-collapse: collapse; font-size: 0.8rem; }
    table.drill th { background: #003366; color: white; padding: 6px 10px; text-align: left; }
    table.drill td { padding: 5px 10px; border-bottom: 1px solid #e2e8f0; }
    .dark table.drill td { border-color: #334155; }
    table.drill tr:hover td { background: rgba(0,159,218,0.08); }
    /* Filter bar */
    select.filter-sel {
      padding: 5px 10px; border: 1px solid #cbd5e1;
      border-radius: 6px; font-size: 0.85rem;
      background: white; cursor: pointer;
    }
    .dark select.filter-sel { background: #1e293b; border-color: #475569; color: #e2e8f0; }
  </style>
</head>
<body class="bg-gray-50 dark:bg-slate-900 text-slate-800 dark:text-slate-200 min-h-screen">

<!-- ════════════════ HEADER ════════════════ -->
<header class="bg-[#003366] text-white px-6 py-4 shadow-lg">
  <div class="max-w-screen-xl mx-auto flex items-center justify-between flex-wrap gap-3">
    <div class="flex items-center gap-4">
      <!-- FAO logo placeholder -->
      <div class="w-10 h-10 rounded-full bg-[#009FDA] flex items-center justify-center font-bold text-white text-sm select-none">FAO</div>
      <div>
        <h1 class="text-xl font-bold leading-tight">Ukraine Procurement KPI Dashboard</h1>
        <p class="text-xs text-blue-200">Last updated: <span id="last-updated"></span></p>
      </div>
    </div>
    <div class="flex items-center gap-3">
      <!-- Filter bar -->
      <select id="filter-year"  class="filter-sel" title="Filter by year">
        <option value="">All Years</option>
      </select>
      <select id="filter-buyer" class="filter-sel" title="Filter by buyer">
        <option value="">All Buyers</option>
      </select>
      <select id="filter-stage" class="filter-sel" title="Filter by stage">
        <option value="">All Stages</option>
      </select>
      <button id="reset-filters" class="text-xs px-3 py-1 bg-white/10 hover:bg-white/20 rounded border border-white/30 transition">Reset</button>
      <!-- Dark mode toggle -->
      <button id="theme-toggle" class="ml-2 px-3 py-1 bg-white/10 hover:bg-white/20 rounded border border-white/30 text-xs transition">Dark</button>
    </div>
  </div>
</header>

<!-- ════════════════ KPI CARDS ROW ════════════════ -->
<main class="max-w-screen-xl mx-auto px-4 py-6 space-y-8">

  <section id="kpi-cards" class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-4">
    <!-- injected by JS -->
  </section>

  <!-- ════ CHARTS GRID ════ -->
  <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">

    <!-- KPI 1 – Cycle Time -->
    <div class="chart-card bg-white dark:bg-slate-800 rounded-xl shadow p-4">
      <button class="export-btn" onclick="exportChart('chart-cycle')">PNG</button>
      <h2 class="font-semibold text-sm mb-3 text-slate-600 dark:text-slate-300 uppercase tracking-wide">KPI 1 – Avg Cycle Time by Method (days)</h2>
      <canvas id="chart-cycle"></canvas>
    </div>

    <!-- KPI 2 – Savings -->
    <div class="chart-card bg-white dark:bg-slate-800 rounded-xl shadow p-4">
      <button class="export-btn" onclick="exportChart('chart-savings')">PNG</button>
      <h2 class="font-semibold text-sm mb-3 text-slate-600 dark:text-slate-300 uppercase tracking-wide">KPI 2 – Top 10 Savings (USD)</h2>
      <canvas id="chart-savings"></canvas>
    </div>

    <!-- KPI 3 – Competitive vs Direct -->
    <div class="chart-card bg-white dark:bg-slate-800 rounded-xl shadow p-4">
      <button class="export-btn" onclick="exportChart('chart-comp')">PNG</button>
      <h2 class="font-semibold text-sm mb-3 text-slate-600 dark:text-slate-300 uppercase tracking-wide">KPI 3 – Competitive vs Direct (count)</h2>
      <canvas id="chart-comp"></canvas>
    </div>

    <!-- KPI 4 – Plan Compliance -->
    <div class="chart-card bg-white dark:bg-slate-800 rounded-xl shadow p-4">
      <button class="export-btn" onclick="exportChart('chart-plan')">PNG</button>
      <h2 class="font-semibold text-sm mb-3 text-slate-600 dark:text-slate-300 uppercase tracking-wide">KPI 4 – Plan Compliance</h2>
      <canvas id="chart-plan"></canvas>
    </div>

    <!-- KPI 5 – Buyer Workload -->
    <div class="chart-card bg-white dark:bg-slate-800 rounded-xl shadow p-4 lg:col-span-2">
      <button class="export-btn" onclick="exportChart('chart-buyer')">PNG</button>
      <h2 class="font-semibold text-sm mb-3 text-slate-600 dark:text-slate-300 uppercase tracking-wide">KPI 5 – Buyer Workload</h2>
      <canvas id="chart-buyer"></canvas>
    </div>

    <!-- Pipeline -->
    <div class="chart-card bg-white dark:bg-slate-800 rounded-xl shadow p-4 lg:col-span-2">
      <button class="export-btn" onclick="exportChart('chart-pipeline')">PNG</button>
      <h2 class="font-semibold text-sm mb-3 text-slate-600 dark:text-slate-300 uppercase tracking-wide">Pipeline – Items per Procurement Stage</h2>
      <canvas id="chart-pipeline"></canvas>
    </div>

  </div>
</main>

<!-- ════════════════ FOOTER ════════════════ -->
<footer class="text-center text-xs text-slate-400 py-6 border-t border-slate-200 dark:border-slate-700 mt-4">
  Data source: FAO Ukraine Procurement Tracking SharePoint List | Refreshed daily via GitHub Actions
</footer>

<!-- ════════════════ MODAL ════════════════ -->
<div id="modal-overlay" role="dialog" aria-modal="true">
  <div id="modal-box">
    <span id="modal-close" onclick="closeModal()" aria-label="Close">&times;</span>
    <h3 id="modal-title" class="font-bold text-base mb-4 text-[#003366] dark:text-[#009FDA]"></h3>
    <div id="modal-content"></div>
  </div>
</div>

<script>
// ════════════════════════════════════════════════════════════════════
// Embedded data
// ════════════════════════════════════════════════════════════════════
const DASHBOARD_DATA = ${dataJson};

// ════════════════════════════════════════════════════════════════════
// State
// ════════════════════════════════════════════════════════════════════
let activeFilters = { year: '', buyer: '', stage: '' };
const chartRegistry = {};   // id → Chart instance

// ════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════
function fmt(n, dec = 0) {
  if (n === null || n === undefined) return '–';
  return n.toLocaleString('en-US', { maximumFractionDigits: dec, minimumFractionDigits: dec });
}
function fmtUSD(n) {
  if (n === null || n === undefined) return '–';
  return '$' + fmt(n, 0);
}
function isDark() { return document.documentElement.classList.contains('dark'); }
function gridColor() { return isDark() ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.07)'; }
function textColor() { return isDark() ? '#94a3b8' : '#64748b'; }

// ════════════════════════════════════════════════════════════════════
// Filter helpers
// ════════════════════════════════════════════════════════════════════
function filteredRows() {
  return DASHBOARD_DATA.rows.filter(r => {
    if (activeFilters.year  && String(r.year)  !== activeFilters.year)  return false;
    if (activeFilters.buyer && r.buyer !== activeFilters.buyer) return false;
    if (activeFilters.stage && r.stage !== activeFilters.stage) return false;
    return true;
  });
}

function recomputeKPIs(rows) {
  // KPI 1
  const cycleByMethod = {};
  for (const r of rows) {
    if (r.cycleTime === null || r.cycleTime < 0) continue;
    const m = r.method || 'Unknown';
    if (!cycleByMethod[m]) cycleByMethod[m] = [];
    cycleByMethod[m].push(r.cycleTime);
  }
  const kpi1 = Object.entries(cycleByMethod).map(([method, times]) => {
    const avg = times.reduce((a, b) => a + b, 0) / times.length;
    return { method, avg: Math.round(avg), count: times.length,
      color: avg <= 30 ? '#4CAF50' : avg <= 90 ? '#FF9800' : '#F44336' };
  }).sort((a, b) => b.count - a.count);

  const allCycles = rows.filter(r => r.cycleTime !== null && r.cycleTime >= 0).map(r => r.cycleTime);
  const avgCycle = allCycles.length ? Math.round(allCycles.reduce((a, b) => a + b, 0) / allCycles.length) : 0;

  // KPI 2
  let sumPos = 0, sumNeg = 0;
  const savingsRows = rows.filter(r => r.savings !== null);
  for (const r of savingsRows) {
    if (r.savings > 0) sumPos += r.savings;
    else if (r.savings < 0) sumNeg += r.savings;
  }
  const top10 = [...savingsRows].sort((a, b) => Math.abs(b.savings) - Math.abs(a.savings)).slice(0, 10)
    .map(r => ({ id: r.id, title: r.title.slice(0, 60) + (r.title.length > 60 ? '…' : ''), savings: r.savings, buyer: r.buyer }));
  const kpi2 = { sumPos, sumNeg, net: sumPos + sumNeg, top10 };

  // KPI 3
  let compCount = 0, compValue = 0, dirCount = 0, dirValue = 0;
  for (const r of rows) {
    if (r.isCompetitive) { compCount++; if (r.prValue) compValue += r.prValue; }
    else                  { dirCount++;  if (r.prValue) dirValue  += r.prValue; }
  }
  const kpi3 = { compCount, compValue, dirCount, dirValue };

  // KPI 4
  const planCounts = { Planned: 0, Unplanned: 0, 'N/A': 0 };
  const planValues = { Planned: 0, Unplanned: 0, 'N/A': 0 };
  for (const r of rows) { planCounts[r.planBucket]++; if (r.prValue) planValues[r.planBucket] += r.prValue; }
  const kpi4 = { planCounts, planValues };

  // KPI 5
  const buyerMap = {};
  for (const r of rows) {
    const b = r.buyer || 'Unknown';
    if (!buyerMap[b]) buyerMap[b] = { count: 0, value: 0 };
    buyerMap[b].count++;
    if (r.prValue) buyerMap[b].value += r.prValue;
  }
  const kpi5 = Object.entries(buyerMap).map(([buyer, d]) => ({ buyer, ...d })).sort((a, b) => b.count - a.count);

  // Pipeline
  const pipelineMap = {};
  for (const r of rows) {
    const s = r.stage || 'Unknown';
    if (!pipelineMap[s]) pipelineMap[s] = { count: 0, value: 0 };
    pipelineMap[s].count++;
    if (r.prValue) pipelineMap[s].value += r.prValue;
  }
  const STAGE_ORDER = ['New PR','Assigned','Solicitation Issued','Offers Received','Evaluation','Award Recommendation','PO Issued','Closed','Cancelled'];
  const pipeline = Object.entries(pipelineMap).map(([stage, d]) => ({ stage, ...d }))
    .sort((a, b) => {
      const ia = STAGE_ORDER.findIndex(s => a.stage.toLowerCase().includes(s.toLowerCase()));
      const ib = STAGE_ORDER.findIndex(s => b.stage.toLowerCase().includes(s.toLowerCase()));
      if (ia !== -1 && ib !== -1) return ia - ib;
      if (ia !== -1) return -1; if (ib !== -1) return 1;
      return b.count - a.count;
    });

  return { kpi1, kpi2, kpi3, kpi4, kpi5, pipeline, avgCycle };
}

// ════════════════════════════════════════════════════════════════════
// KPI Cards
// ════════════════════════════════════════════════════════════════════
function renderCards(kpis, totalRows) {
  const { kpi1, kpi2, kpi3, kpi4, avgCycle } = kpis;
  const compPct = (kpi3.compCount + kpi3.dirCount)
    ? Math.round(100 * kpi3.compCount / (kpi3.compCount + kpi3.dirCount)) : 0;
  const planPct = (totalRows)
    ? Math.round(100 * kpi4.planCounts.Planned / totalRows) : 0;

  const cycleColor = avgCycle <= 30 ? 'text-green-600' : avgCycle <= 90 ? 'text-amber-500' : 'text-red-500';

  const cards = [
    { label: 'Total PRs',    value: fmt(totalRows),         sub: 'in selection',              color: 'bg-[#003366]', icon: '📋' },
    { label: 'Avg Cycle',    value: avgCycle + 'd',         sub: 'PR → PO issuance',          color: avgCycle <= 30 ? 'bg-green-600' : avgCycle <= 90 ? 'bg-amber-500' : 'bg-red-500', icon: '⏱' },
    { label: 'Net Savings',  value: fmtUSD(kpi2.net),       sub: 'vs estimated value',        color: kpi2.net >= 0 ? 'bg-green-600' : 'bg-red-500', icon: '💰' },
    { label: 'Competitive',  value: compPct + '%',          sub: fmt(kpi3.compCount) + ' of ' + fmt(kpi3.compCount + kpi3.dirCount), color: 'bg-[#009FDA]', icon: '🏆' },
    { label: 'Planned',      value: planPct + '%',          sub: fmt(kpi4.planCounts.Planned) + ' planned PRs', color: 'bg-indigo-600', icon: '📅' },
    { label: 'Total Value',  value: fmtUSD(kpi3.compValue + kpi3.dirValue), sub: 'sum PRValue', color: 'bg-slate-700', icon: '💵' },
  ];

  document.getElementById('kpi-cards').innerHTML = cards.map(c => \`
    <div class="kpi-card \${c.color} text-white rounded-xl shadow p-4 flex flex-col gap-1">
      <span class="text-xs opacity-70 font-medium uppercase tracking-wide">\${c.icon} \${c.label}</span>
      <span class="text-2xl font-extrabold">\${c.value}</span>
      <span class="text-xs opacity-60">\${c.sub}</span>
    </div>
  \`).join('');
}

// ════════════════════════════════════════════════════════════════════
// Chart helpers
// ════════════════════════════════════════════════════════════════════
function makeOrUpdate(id, config) {
  if (chartRegistry[id]) { chartRegistry[id].destroy(); }
  const ctx = document.getElementById(id).getContext('2d');
  chartRegistry[id] = new Chart(ctx, config);
}

// ════════════════════════════════════════════════════════════════════
// Render all charts
// ════════════════════════════════════════════════════════════════════
function renderCharts(kpis) {
  const { kpi1, kpi2, kpi3, kpi4, kpi5, pipeline } = kpis;

  // ── Chart 1: Cycle Time horizontal bar ───────────────────────────
  makeOrUpdate('chart-cycle', {
    type: 'bar',
    data: {
      labels: kpi1.map(d => d.method),
      datasets: [{
        label: 'Avg days',
        data: kpi1.map(d => d.avg),
        backgroundColor: kpi1.map(d => d.color),
        borderRadius: 4,
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => \` \${ctx.parsed.x} days (n=\${kpi1[ctx.dataIndex].count})\` } },
      },
      scales: {
        x: { grid: { color: gridColor() }, ticks: { color: textColor() }, title: { display: true, text: 'Days', color: textColor() } },
        y: { grid: { display: false }, ticks: { color: textColor() } },
      },
      onClick: (_e, elements) => {
        if (!elements.length) return;
        const method = kpi1[elements[0].index].method;
        const rows = filteredRows().filter(r => r.method === method && r.cycleTime !== null);
        openModal(\`Cycle Time – \${method}\`, buildDrillTable(rows, ['id','title','buyer','prReceived','poDate','cycleTime']));
      }
    }
  });

  // ── Chart 2: Savings bar (top 10) ────────────────────────────────
  makeOrUpdate('chart-savings', {
    type: 'bar',
    data: {
      labels: kpi2.top10.map(d => d.title),
      datasets: [{
        label: 'Savings (USD)',
        data: kpi2.top10.map(d => d.savings),
        backgroundColor: kpi2.top10.map(d => d.savings >= 0 ? '#4CAF50' : '#F44336'),
        borderRadius: 4,
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ' ' + fmtUSD(ctx.parsed.x) } },
      },
      scales: {
        x: { grid: { color: gridColor() }, ticks: { color: textColor(), callback: v => fmtUSD(v) } },
        y: { grid: { display: false }, ticks: { color: textColor(), font: { size: 10 } } },
      },
      onClick: (_e, elements) => {
        if (!elements.length) return;
        const item = kpi2.top10[elements[0].index];
        const rows = filteredRows().filter(r => r.id === item.id);
        openModal(\`Savings detail – ID \${item.id}\`, buildDrillTable(rows, ['id','title','buyer','method','prValue','cumulativePO','savings']));
      }
    }
  });

  // ── Chart 3: Competitive vs Direct – doughnut ────────────────────
  makeOrUpdate('chart-comp', {
    type: 'doughnut',
    data: {
      labels: ['Competitive', 'Direct/Other'],
      datasets: [{
        data: [kpi3.compCount, kpi3.dirCount],
        backgroundColor: ['#009FDA', '#FF9800'],
        hoverOffset: 6,
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'bottom', labels: { color: textColor() } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = kpi3.compCount + kpi3.dirCount;
              const pct = total ? Math.round(100 * ctx.parsed / total) : 0;
              const val = ctx.dataIndex === 0 ? kpi3.compValue : kpi3.dirValue;
              return \` \${ctx.parsed} PRs (\${pct}%) – \${fmtUSD(val)}\`;
            }
          }
        }
      },
      onClick: (_e, elements) => {
        if (!elements.length) return;
        const isComp = elements[0].index === 0;
        const rows = filteredRows().filter(r => r.isCompetitive === isComp);
        openModal(isComp ? 'Competitive PRs' : 'Direct / Other PRs', buildDrillTable(rows, ['id','title','buyer','method','prValue','stage']));
      }
    }
  });

  // ── Chart 4: Plan Compliance – pie ───────────────────────────────
  const planLabels = ['Planned', 'Unplanned', 'N/A'];
  makeOrUpdate('chart-plan', {
    type: 'pie',
    data: {
      labels: planLabels,
      datasets: [{
        data: planLabels.map(l => kpi4.planCounts[l]),
        backgroundColor: ['#4CAF50', '#F44336', '#90a4ae'],
        hoverOffset: 6,
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'bottom', labels: { color: textColor() } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = planLabels.reduce((a, l) => a + kpi4.planCounts[l], 0);
              const pct = total ? Math.round(100 * ctx.parsed / total) : 0;
              return \` \${ctx.parsed} (\${pct}%) – \${fmtUSD(kpi4.planValues[planLabels[ctx.dataIndex]])}\`;
            }
          }
        }
      },
      onClick: (_e, elements) => {
        if (!elements.length) return;
        const bucket = planLabels[elements[0].index];
        const rows = filteredRows().filter(r => r.planBucket === bucket);
        openModal(\`Plan Compliance – \${bucket}\`, buildDrillTable(rows, ['id','title','buyer','method','prValue','stage']));
      }
    }
  });

  // ── Chart 5: Buyer Workload – grouped bar ────────────────────────
  const TOP_N = Math.min(12, kpi5.length);
  const top5 = kpi5.slice(0, TOP_N);
  makeOrUpdate('chart-buyer', {
    type: 'bar',
    data: {
      labels: top5.map(d => d.buyer),
      datasets: [
        { label: '# PRs', data: top5.map(d => d.count), backgroundColor: '#009FDA', borderRadius: 4, yAxisID: 'y' },
        { label: 'Value (USD)', data: top5.map(d => d.value), backgroundColor: '#003366', borderRadius: 4, yAxisID: 'y2' },
      ]
    },
    options: {
      responsive: true,
      plugins: { legend: { labels: { color: textColor() } } },
      scales: {
        x: { grid: { display: false }, ticks: { color: textColor() } },
        y:  { position: 'left',  grid: { color: gridColor() }, ticks: { color: textColor() }, title: { display: true, text: 'Count', color: textColor() } },
        y2: { position: 'right', grid: { display: false }, ticks: { color: textColor(), callback: v => fmtUSD(v) }, title: { display: true, text: 'USD', color: textColor() } },
      },
      onClick: (_e, elements) => {
        if (!elements.length) return;
        const buyer = top5[elements[0].index].buyer;
        const rows = filteredRows().filter(r => r.buyer === buyer);
        openModal(\`Workload – \${buyer}\`, buildDrillTable(rows, ['id','title','method','prValue','stage','prReceived','poDate','cycleTime']));
      }
    }
  });

  // ── Chart 6: Pipeline stages ─────────────────────────────────────
  makeOrUpdate('chart-pipeline', {
    type: 'bar',
    data: {
      labels: pipeline.map(d => d.stage),
      datasets: [
        { label: '# Items', data: pipeline.map(d => d.count), backgroundColor: '#009FDA', borderRadius: 4, yAxisID: 'y' },
        { label: 'Value (USD)', data: pipeline.map(d => d.value), backgroundColor: '#FF9800', borderRadius: 4, yAxisID: 'y2' },
      ]
    },
    options: {
      responsive: true,
      plugins: { legend: { labels: { color: textColor() } } },
      scales: {
        x: { grid: { display: false }, ticks: { color: textColor() } },
        y:  { position: 'left',  grid: { color: gridColor() }, ticks: { color: textColor() }, title: { display: true, text: 'Count', color: textColor() } },
        y2: { position: 'right', grid: { display: false }, ticks: { color: textColor(), callback: v => fmtUSD(v) }, title: { display: true, text: 'USD', color: textColor() } },
      },
      onClick: (_e, elements) => {
        if (!elements.length) return;
        const stage = pipeline[elements[0].index].stage;
        const rows = filteredRows().filter(r => r.stage === stage);
        openModal(\`Pipeline – \${stage}\`, buildDrillTable(rows, ['id','title','buyer','method','prValue','prReceived','poDate']));
      }
    }
  });
}

// ════════════════════════════════════════════════════════════════════
// Drill-down modal
// ════════════════════════════════════════════════════════════════════
const COL_LABEL = {
  id: 'ID', title: 'Description', buyer: 'Buyer', method: 'Method',
  prValue: 'PR Value', cumulativePO: 'Cumul. PO', savings: 'Savings',
  stage: 'Stage', prReceived: 'PR Received', poDate: 'PO Date',
  cycleTime: 'Cycle (d)', isCompetitive: 'Competitive?', planBucket: 'Plan',
};
function colFmt(key, val) {
  if (val === null || val === undefined || val === '') return '–';
  if (['prValue','cumulativePO','savings'].includes(key)) return fmtUSD(val);
  if (key === 'isCompetitive') return val ? 'Yes' : 'No';
  return String(val);
}
function buildDrillTable(rows, cols) {
  if (!rows.length) return '<p class="text-sm text-slate-500">No records found.</p>';
  const head = cols.map(c => \`<th>\${COL_LABEL[c] || c}</th>\`).join('');
  const body = rows.map(r =>
    '<tr>' + cols.map(c => \`<td>\${colFmt(c, r[c])}</td>\`).join('') + '</tr>'
  ).join('');
  return \`<table class="drill"><thead><tr>\${head}</tr></thead><tbody>\${body}</tbody></table>\`;
}
function openModal(title, html) {
  document.getElementById('modal-title').textContent = title;
  document.getElementById('modal-content').innerHTML = html;
  document.getElementById('modal-overlay').classList.add('open');
}
function closeModal() {
  document.getElementById('modal-overlay').classList.remove('open');
}
document.getElementById('modal-overlay').addEventListener('click', e => {
  if (e.target === document.getElementById('modal-overlay')) closeModal();
});
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

// ════════════════════════════════════════════════════════════════════
// Export PNG
// ════════════════════════════════════════════════════════════════════
function exportChart(id) {
  const chart = chartRegistry[id];
  if (!chart) return;
  const a = document.createElement('a');
  a.href = chart.toBase64Image('image/png', 1);
  a.download = id + '.png';
  a.click();
}

// ════════════════════════════════════════════════════════════════════
// Dark / Light mode toggle
// ════════════════════════════════════════════════════════════════════
document.getElementById('theme-toggle').addEventListener('click', () => {
  const html = document.documentElement;
  html.classList.toggle('dark');
  document.getElementById('theme-toggle').textContent = html.classList.contains('dark') ? 'Light' : 'Dark';
  // Re-render charts so grid/text colors update
  const rows = filteredRows();
  const kpis = recomputeKPIs(rows);
  renderCharts(kpis);
});

// ════════════════════════════════════════════════════════════════════
// Filters
// ════════════════════════════════════════════════════════════════════
function populateFilters() {
  const ysel = document.getElementById('filter-year');
  DASHBOARD_DATA.years.forEach(y => {
    const o = document.createElement('option'); o.value = y; o.textContent = y; ysel.appendChild(o);
  });
  const bsel = document.getElementById('filter-buyer');
  DASHBOARD_DATA.buyers.forEach(b => {
    const o = document.createElement('option'); o.value = b; o.textContent = b; bsel.appendChild(o);
  });
  const ssel = document.getElementById('filter-stage');
  DASHBOARD_DATA.stages.forEach(s => {
    const o = document.createElement('option'); o.value = s; o.textContent = s; ssel.appendChild(o);
  });
}

function applyFilters() {
  activeFilters.year  = document.getElementById('filter-year').value;
  activeFilters.buyer = document.getElementById('filter-buyer').value;
  activeFilters.stage = document.getElementById('filter-stage').value;
  refresh();
}

['filter-year','filter-buyer','filter-stage'].forEach(id =>
  document.getElementById(id).addEventListener('change', applyFilters)
);
document.getElementById('reset-filters').addEventListener('click', () => {
  ['filter-year','filter-buyer','filter-stage'].forEach(id => { document.getElementById(id).value = ''; });
  activeFilters = { year: '', buyer: '', stage: '' };
  refresh();
});

// ════════════════════════════════════════════════════════════════════
// Main refresh
// ════════════════════════════════════════════════════════════════════
function refresh() {
  const rows = filteredRows();
  const kpis = recomputeKPIs(rows);
  renderCards(kpis, rows.length);
  renderCharts(kpis);
}

// ════════════════════════════════════════════════════════════════════
// Init
// ════════════════════════════════════════════════════════════════════
document.getElementById('last-updated').textContent = DASHBOARD_DATA.lastUpdated;
populateFilters();
refresh();
</script>
</body>
</html>`;
}

// ── Entry point ───────────────────────────────────────────────────────────────

if (!fs.existsSync(INPUT_PATH)) {
  console.error(`ERROR: Input file not found: ${INPUT_PATH}`);
  process.exit(1);
}

const rows = loadData(INPUT_PATH);
if (!rows.length) {
  console.error('ERROR: No rows parsed from CSV.');
  process.exit(1);
}

const data = computeKPIs(rows);
console.log(`  KPIs computed. Last updated: ${data.lastUpdated}`);
console.log(`  Rows: ${data.rows.length} | Buyers: ${data.buyers.length} | Stages: ${data.stages.length}`);

const html = generateHTML(data);
const outDir = path.dirname(OUTPUT_PATH);
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(OUTPUT_PATH, html, 'utf8');
console.log(`Dashboard written to: ${OUTPUT_PATH} (${(html.length / 1024).toFixed(1)} KB)`);
