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

  // Monthly savings — group by YYYY-MM using poDate or prReceived
  const monthlyMap = {};
  for (const r of rows) {
    const d = r.poDate || r.prReceived;
    if (!d || r.savings === null) continue;
    const month = fmtDate(d).slice(0, 7); // YYYY-MM
    if (!monthlyMap[month]) monthlyMap[month] = 0;
    monthlyMap[month] += r.savings;
  }
  const monthlySavings = Object.entries(monthlyMap)
    .sort(([a],[b]) => a.localeCompare(b))
    .map(([month, total]) => ({ month, total: Math.round(total) }));
  const avgMonthlySavings = monthlySavings.length
    ? Math.round(monthlySavings.reduce((s,m) => s + m.total, 0) / monthlySavings.length)
    : 0;
  kpi2.monthly = monthlySavings;
  kpi2.avgMonthly = avgMonthlySavings;

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

  // Projects (from projRef multi-lookup field, split by ';')
  const projectSet = new Set();
  for (const r of rows) {
    if (r.projRef) {
      r.projRef.split(';').map(p => p.trim()).filter(Boolean).forEach(p => projectSet.add(p));
    }
  }
  const projects = [...projectSet].sort();

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

  return { kpi1, kpi2, kpi3, kpi4, kpi5, pipeline, years, buyers, stages, projects, lastUpdated, rows: serialRows, avgCycle };
}


// ── HTML generator ────────────────────────────────────────────────────────────

function generateHTML(data) {
  const dataJson = JSON.stringify(data);
  return `<!DOCTYPE html>
<html lang="en" data-theme="light">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>FAO Ukraine – Procurement KPI Dashboard</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js" crossorigin="anonymous"><\/script>
  <style>
    :root {
      --bg:#f1f5f9; --card:#fff; --text:#1a2e44; --muted:#64748b;
      --border:#e2e8f0; --navy:#003366; --blue:#009FDA;
      --green:#22c55e; --amber:#f59e0b; --red:#ef4444;
      --shadow:0 2px 10px rgba(0,0,0,0.07);
    }
    [data-theme="dark"] {
      --bg:#0f172a; --card:#1e293b; --text:#e2e8f0; --muted:#94a3b8;
      --border:#334155; --shadow:0 2px 10px rgba(0,0,0,0.3);
    }
    *{box-sizing:border-box;margin:0;padding:0;}
    body{font-family:'Segoe UI',Arial,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;transition:background .25s,color .25s;}

    /* ── Header ── */
    .hdr{background:var(--navy);color:#fff;padding:10px 24px;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;}
    .hdr-left{display:flex;align-items:center;gap:14px;}
    .hdr-title{font-size:16px;font-weight:700;line-height:1.2;}
    .hdr-sub{font-size:10px;opacity:.6;margin-top:2px;}
    /* Theme toggle */
    .toggle-wrap{display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;}
    .toggle-wrap input{display:none;}
    .toggle-track{width:40px;height:22px;background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.3);border-radius:11px;position:relative;transition:background .25s;}
    .toggle-track::after{content:'';position:absolute;top:3px;left:3px;width:14px;height:14px;border-radius:50%;background:#fff;transition:transform .25s;}
    input:checked + .toggle-track{background:var(--blue);border-color:var(--blue);}
    input:checked + .toggle-track::after{transform:translateX(18px);}

    /* ── Filter bar ── */
    .filter-bar{background:var(--card);border-bottom:1px solid var(--border);padding:10px 24px;display:flex;gap:18px;flex-wrap:wrap;align-items:flex-start;}
    .filter-group{display:flex;flex-direction:column;gap:5px;}
    .filter-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.7px;color:var(--muted);}
    .pills{display:flex;flex-wrap:wrap;gap:4px;}
    .pill{padding:3px 12px;border-radius:20px;font-size:11.5px;font-weight:500;border:1.5px solid var(--border);background:var(--card);color:var(--text);cursor:pointer;transition:all .13s;white-space:nowrap;}
    .pill:hover{border-color:var(--blue);color:var(--blue);}
    .pill.on{background:var(--navy);border-color:var(--navy);color:#fff;}

    /* ── KPI strip ── */
    .kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px;}
    .kpi-card{background:var(--card);border-radius:12px;padding:16px 20px;box-shadow:var(--shadow);border-left:4px solid var(--blue);display:flex;flex-direction:column;gap:3px;}
    .kpi-lbl{font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);font-weight:700;}
    .kpi-val{font-size:28px;font-weight:800;line-height:1.1;}
    .kpi-sub{font-size:11px;color:var(--muted);}

    /* ── Charts ── */
    .content{max-width:1200px;margin:0 auto;padding:20px 20px;}
    .grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px;}
    .full{grid-column:1/-1;}
    .chart-card{background:var(--card);border-radius:12px;padding:18px 20px;box-shadow:var(--shadow);position:relative;}
    .chart-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);margin-bottom:14px;}
    .export-btn{position:absolute;top:12px;right:12px;font-size:10px;padding:2px 8px;background:rgba(0,159,218,.1);border:1px solid #009FDA;color:#009FDA;border-radius:4px;cursor:pointer;font-family:inherit;}
    .export-btn:hover{background:rgba(0,159,218,.25);}
    canvas{max-height:300px;}

    /* ── Buyer workload cards ── */
    .buyer-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:10px;}
    .buyer-card{background:var(--card);border-radius:10px;padding:14px 16px;box-shadow:var(--shadow);cursor:pointer;transition:transform .12s;}
    .buyer-card:hover{transform:translateY(-2px);}
    .buyer-name{font-weight:700;font-size:13px;color:var(--navy);margin-bottom:8px;}
    [data-theme="dark"] .buyer-name{color:var(--blue);}
    .buyer-bar-bg{background:var(--border);border-radius:4px;height:5px;overflow:hidden;margin:5px 0;}
    .buyer-bar-fg{height:5px;border-radius:4px;background:var(--blue);}
    .buyer-stats{display:flex;justify-content:space-between;font-size:11px;color:var(--muted);}

    /* ── Modal ── */
    #modal-ov{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:100;align-items:center;justify-content:center;}
    #modal-ov.open{display:flex;}
    #modal-box{background:var(--card);border-radius:14px;max-width:90vw;width:750px;max-height:88vh;overflow-y:auto;padding:24px;position:relative;}
    #modal-close{position:absolute;top:12px;right:16px;font-size:20px;cursor:pointer;color:var(--muted);background:none;border:none;}
    #modal-title{font-weight:700;font-size:15px;margin-bottom:16px;color:var(--navy);}
    [data-theme="dark"] #modal-title{color:var(--blue);}
    table.drill{width:100%;border-collapse:collapse;font-size:.8rem;}
    table.drill th{background:var(--navy);color:#fff;padding:7px 10px;text-align:left;font-size:11px;}
    table.drill td{padding:6px 10px;border-bottom:1px solid var(--border);}
    table.drill tr:hover td{background:rgba(0,159,218,.06);}

    footer{text-align:center;font-size:11px;color:var(--muted);padding:20px;border-top:1px solid var(--border);margin-top:8px;}

    @media(max-width:768px){
      .kpi-row{grid-template-columns:1fr 1fr;}
      .grid2{grid-template-columns:1fr;}
      .full{grid-column:auto;}
    }
  </style>
</head>
<body>

<!-- ── HEADER ── -->
<header class="hdr">
  <div class="hdr-left">
    <!-- FAO logo SVG (simplified emblem) -->
    <svg width="38" height="38" viewBox="0 0 38 38" fill="none" xmlns="http://www.w3.org/2000/svg">
      <circle cx="19" cy="19" r="18" stroke="white" stroke-width="1.2" fill="none"/>
      <!-- Wheat left -->
      <line x1="9" y1="28" x2="12" y2="12" stroke="white" stroke-width="1"/>
      <ellipse cx="10.5" cy="26" rx="2.2" ry="1.2" fill="white" transform="rotate(-20 10.5 26)"/>
      <ellipse cx="10" cy="23" rx="2.2" ry="1.2" fill="white" transform="rotate(-15 10 23)"/>
      <ellipse cx="11" cy="20" rx="2" ry="1.1" fill="white" transform="rotate(-10 11 20)"/>
      <!-- Wheat right -->
      <line x1="29" y1="28" x2="26" y2="12" stroke="white" stroke-width="1"/>
      <ellipse cx="27.5" cy="26" rx="2.2" ry="1.2" fill="white" transform="rotate(20 27.5 26)"/>
      <ellipse cx="28" cy="23" rx="2.2" ry="1.2" fill="white" transform="rotate(15 28 23)"/>
      <ellipse cx="27" cy="20" rx="2" ry="1.1" fill="white" transform="rotate(10 27 20)"/>
      <!-- Globe circle -->
      <circle cx="19" cy="21" r="6" stroke="white" stroke-width="1" fill="none"/>
      <line x1="13" y1="21" x2="25" y2="21" stroke="white" stroke-width=".8"/>
      <ellipse cx="19" cy="21" rx="3" ry="6" stroke="white" stroke-width=".8" fill="none"/>
      <!-- FAO text -->
      <text x="19" y="12" text-anchor="middle" font-size="5.5" font-weight="bold" fill="white" font-family="Arial">FAO</text>
    </svg>
    <div>
      <div class="hdr-title">Ukraine Procurement KPI Dashboard</div>
      <div class="hdr-sub">Last updated: <span id="last-updated"></span> &nbsp;·&nbsp; Food and Agriculture Organization of the UN</div>
    </div>
  </div>
  <label class="toggle-wrap" title="Toggle dark mode">
    <span style="font-size:13px;">☀️</span>
    <input type="checkbox" id="theme-chk"/>
    <div class="toggle-track"></div>
    <span style="font-size:13px;">🌙</span>
  </label>
</header>

<!-- ── FILTER BAR ── -->
<div class="filter-bar" id="filter-bar">
  <div class="filter-group">
    <div class="filter-label">📅 Year</div>
    <div class="pills" id="pills-year"></div>
  </div>
  <div class="filter-group">
    <div class="filter-label">👤 Buyer</div>
    <div class="pills" id="pills-buyer"></div>
  </div>
  <div class="filter-group">
    <div class="filter-label">📁 Project</div>
    <div class="pills" id="pills-project"></div>
  </div>
</div>

<!-- ── MAIN ── -->
<div class="content">

  <!-- KPI summary strip -->
  <div class="kpi-row" id="kpi-row"></div>

  <!-- Charts row 1: Cycle Time + Competitive -->
  <div class="grid2">
    <div class="chart-card">
      <button class="export-btn" onclick="exportChart('chart-cycle')">PNG</button>
      <div class="chart-title">KPI 1 – Avg Cycle Time by Method (days)</div>
      <canvas id="chart-cycle"></canvas>
    </div>
    <div class="chart-card">
      <button class="export-btn" onclick="exportChart('chart-comp')">PNG</button>
      <div class="chart-title">KPI 3 – Competitive vs Direct</div>
      <canvas id="chart-comp"></canvas>
    </div>
  </div>

  <!-- Charts row 2: Savings over time (full width) -->
  <div class="grid2">
    <div class="chart-card full">
      <button class="export-btn" onclick="exportChart('chart-savings')">PNG</button>
      <div class="chart-title">KPI 2 – Savings per Month (USD) with Average</div>
      <canvas id="chart-savings"></canvas>
    </div>
  </div>

  <!-- Charts row 3: Plan Compliance + Buyer Workload -->
  <div class="grid2" style="margin-bottom:16px;">
    <div class="chart-card">
      <button class="export-btn" onclick="exportChart('chart-plan')">PNG</button>
      <div class="chart-title">KPI 4 – Plan Compliance</div>
      <canvas id="chart-plan"></canvas>
    </div>
    <div class="chart-card">
      <div class="chart-title">KPI 5 – Buyer Workload</div>
      <div class="buyer-grid" id="buyer-grid"></div>
    </div>
  </div>

  <!-- Pipeline (full width) -->
  <div class="grid2">
    <div class="chart-card full">
      <button class="export-btn" onclick="exportChart('chart-pipeline')">PNG</button>
      <div class="chart-title">Pipeline – Items per Procurement Stage</div>
      <canvas id="chart-pipeline"></canvas>
    </div>
  </div>

</div>

<footer>Data source: FAO Ukraine Procurement Tracking SharePoint List &nbsp;·&nbsp; Refreshed daily via GitHub Actions</footer>

<!-- ── MODAL ── -->
<div id="modal-ov" role="dialog" aria-modal="true">
  <div id="modal-box">
    <button id="modal-close" onclick="closeModal()" aria-label="Close">&times;</button>
    <div id="modal-title"></div>
    <div id="modal-content"></div>
  </div>
</div>

<script>
const DASHBOARD_DATA = ${dataJson};

// ── State ──────────────────────────────────────────────────────────
let AF = { year:'', buyer:'', project:'' };
const CR = {};

// ── Helpers ────────────────────────────────────────────────────────
function fmt(n,d=0){ return n==null?'–':n.toLocaleString('en-US',{maximumFractionDigits:d,minimumFractionDigits:d}); }
function fmtUSD(n){ return n==null?'–':'$'+fmt(n,0); }
function isDark(){ return document.documentElement.getAttribute('data-theme')==='dark'; }
function gc(){ return isDark()?'rgba(255,255,255,.07)':'rgba(0,0,0,.07)'; }
function tc(){ return isDark()?'#94a3b8':'#64748b'; }

// ── Filter rows ────────────────────────────────────────────────────
function filteredRows(){
  return DASHBOARD_DATA.rows.filter(r=>{
    if(AF.year    && String(r.year)!==AF.year) return false;
    if(AF.buyer   && r.buyer!==AF.buyer)       return false;
    if(AF.project && !(r.projRef||'').includes(AF.project)) return false;
    return true;
  });
}

// ── Recompute KPIs client-side ─────────────────────────────────────
function recompute(rows){
  // KPI 1
  const cm={};
  for(const r of rows){
    if(r.cycleTime===null||r.cycleTime<0) continue;
    const m=r.method||'Unknown';
    if(!cm[m]) cm[m]=[];
    cm[m].push(r.cycleTime);
  }
  const kpi1=Object.entries(cm).map(([method,t])=>{
    const avg=Math.round(t.reduce((a,b)=>a+b,0)/t.length);
    return{method,avg,count:t.length,color:avg<=30?'#22c55e':avg<=90?'#f59e0b':'#ef4444'};
  }).sort((a,b)=>b.count-a.count);
  const allC=rows.filter(r=>r.cycleTime!==null&&r.cycleTime>=0).map(r=>r.cycleTime);
  const avgCycle=allC.length?Math.round(allC.reduce((a,b)=>a+b,0)/allC.length):0;

  // KPI 2
  let sumPos=0,sumNeg=0;
  const sRows=rows.filter(r=>r.savings!==null);
  for(const r of sRows){ if(r.savings>0)sumPos+=r.savings; else if(r.savings<0)sumNeg+=r.savings; }
  // Monthly savings
  const mm={};
  for(const r of rows){
    const d=r.poDate||r.prReceived; if(!d||r.savings===null) continue;
    const mo=d.slice(0,7);
    if(!mm[mo]) mm[mo]=0; mm[mo]+=r.savings;
  }
  const monthly=Object.entries(mm).sort(([a],[b])=>a.localeCompare(b)).map(([month,total])=>({month,total:Math.round(total)}));
  const avgMonthly=monthly.length?Math.round(monthly.reduce((s,m)=>s+m.total,0)/monthly.length):0;
  const kpi2={sumPos,sumNeg,net:sumPos+sumNeg,monthly,avgMonthly};

  // KPI 3
  let cc=0,cv=0,dc=0,dv=0;
  for(const r of rows){ if(r.isCompetitive){cc++;if(r.prValue)cv+=r.prValue;}else{dc++;if(r.prValue)dv+=r.prValue;} }
  const kpi3={compCount:cc,compValue:cv,dirCount:dc,dirValue:dv};

  // KPI 4
  const pc={Planned:0,Unplanned:0,'N/A':0},pv={Planned:0,Unplanned:0,'N/A':0};
  for(const r of rows){ pc[r.planBucket]++; if(r.prValue)pv[r.planBucket]+=r.prValue; }
  const kpi4={planCounts:pc,planValues:pv};

  // KPI 5
  const bm={};
  for(const r of rows){ const b=r.buyer||'Unknown'; if(!bm[b])bm[b]={count:0,value:0}; bm[b].count++; if(r.prValue)bm[b].value+=r.prValue; }
  const kpi5=Object.entries(bm).map(([buyer,d])=>({buyer,...d})).sort((a,b)=>b.count-a.count);

  // Pipeline
  const pm={};
  for(const r of rows){ const s=r.stage||'Unknown'; if(!pm[s])pm[s]={count:0,value:0}; pm[s].count++; if(r.prValue)pm[s].value+=r.prValue; }
  const SO=['New PR','Assigned','Solicitation Issued','Offers Received','Evaluation','Award Recommendation','PO Issued','Closed','Cancelled'];
  const pipeline=Object.entries(pm).map(([stage,d])=>({stage,...d}))
    .sort((a,b)=>{ const ia=SO.findIndex(s=>a.stage.toLowerCase().includes(s.toLowerCase())); const ib=SO.findIndex(s=>b.stage.toLowerCase().includes(s.toLowerCase())); if(ia!==-1&&ib!==-1)return ia-ib; if(ia!==-1)return-1; if(ib!==-1)return 1; return b.count-a.count; });

  return{kpi1,kpi2,kpi3,kpi4,kpi5,pipeline,avgCycle};
}

// ── KPI Cards ──────────────────────────────────────────────────────
function renderCards(K,total){
  const {kpi2,kpi3,kpi4,avgCycle}=K;
  const compPct=(kpi3.compCount+kpi3.dirCount)?Math.round(100*kpi3.compCount/(kpi3.compCount+kpi3.dirCount)):0;
  const cycleCol=avgCycle<=30?'var(--green)':avgCycle<=90?'var(--amber)':'var(--red)';
  const netCol=kpi2.net>=0?'var(--green)':'var(--red)';
  const cards=[
    {lbl:'Total PRs',     val:fmt(total),        sub:'in selection',         col:'var(--navy)'},
    {lbl:'Avg Cycle Time',val:avgCycle+'d',       sub:'PR \u2192 PO issuance',     col:cycleCol},
    {lbl:'Net Savings',   val:fmtUSD(kpi2.net),  sub:'vs estimated value',   col:netCol},
    {lbl:'Competitive',   val:compPct+'%',        sub:fmt(kpi3.compCount)+' of '+fmt(kpi3.compCount+kpi3.dirCount)+' PRs', col:'var(--blue)'},
  ];
  document.getElementById('kpi-row').innerHTML=cards.map(c=>\`
    <div class="kpi-card" style="border-left-color:\${c.col}">
      <div class="kpi-lbl">\${c.lbl}</div>
      <div class="kpi-val" style="color:\${c.col}">\${c.val}</div>
      <div class="kpi-sub">\${c.sub}</div>
    </div>\`).join('');
}

// ── Buyer workload cards ───────────────────────────────────────────
function renderBuyers(kpi5){
  const maxC=kpi5[0]?.count||1;
  document.getElementById('buyer-grid').innerHTML=kpi5.map(b=>\`
    <div class="buyer-card" onclick="openBuyerDrill('\${b.buyer.replace(/'/g,"\\\\'")}')">
      <div class="buyer-name">\${b.buyer}</div>
      <div class="buyer-bar-bg"><div class="buyer-bar-fg" style="width:\${Math.round(100*b.count/maxC)}%"></div></div>
      <div class="buyer-stats"><span>\${b.count} PRs</span><span>\${fmtUSD(b.value)}</span></div>
    </div>\`).join('');
}

// ── Chart helpers ─────────────────────────────────────────────────
function mou(id,cfg){
  if(CR[id]){CR[id].destroy();}
  CR[id]=new Chart(document.getElementById(id).getContext('2d'),cfg);
}

// ── Render charts ─────────────────────────────────────────────────
function renderCharts(K){
  const{kpi1,kpi2,kpi3,kpi4,pipeline}=K;

  // Chart 1: Cycle time horizontal bar
  mou('chart-cycle',{type:'bar',data:{
    labels:kpi1.map(d=>d.method),
    datasets:[{label:'Avg days',data:kpi1.map(d=>d.avg),backgroundColor:kpi1.map(d=>d.color),borderRadius:4}]
  },options:{indexAxis:'y',responsive:true,plugins:{legend:{display:false},
    tooltip:{callbacks:{label:ctx=>\` \${ctx.parsed.x}d (n=\${kpi1[ctx.dataIndex].count})\`}}},
    scales:{x:{grid:{color:gc()},ticks:{color:tc()},title:{display:true,text:'Days',color:tc()}},y:{grid:{display:false},ticks:{color:tc()}}},
    onClick:(_e,el)=>{ if(!el.length)return; const m=kpi1[el[0].index].method;
      openModal(\`Cycle Time \u2013 \${m}\`,drillTable(filteredRows().filter(r=>r.method===m&&r.cycleTime!==null),['id','title','buyer','prReceived','poDate','cycleTime'])); }
  }});

  // Chart 2: Monthly savings bar + average line
  const labels=kpi2.monthly.map(m=>{
    const[y,mo]=m.month.split('-'); return new Date(+y,+mo-1).toLocaleString('en',{month:'short',year:'2-digit'});
  });
  const avgArr=kpi2.monthly.map(()=>kpi2.avgMonthly);
  mou('chart-savings',{type:'bar',data:{
    labels,
    datasets:[
      {type:'bar',label:'Monthly Savings (USD)',data:kpi2.monthly.map(m=>m.total),
       backgroundColor:kpi2.monthly.map(m=>m.total>=0?'rgba(34,197,94,.7)':'rgba(239,68,68,.7)'),borderRadius:3,order:2},
      {type:'line',label:'Monthly Average',data:avgArr,borderColor:'#f59e0b',borderWidth:2,borderDash:[5,4],
       pointRadius:0,fill:false,order:1,tension:0}
    ]
  },options:{responsive:true,plugins:{legend:{labels:{color:tc()}}},
    scales:{x:{grid:{color:gc()},ticks:{color:tc(),maxRotation:45}},
            y:{grid:{color:gc()},ticks:{color:tc(),callback:v=>fmtUSD(v)}}},
    onClick:(_e,el)=>{ if(!el.length)return; const item=kpi2.monthly[el[0].index];
      openModal(\`Savings \u2013 \${item.month}\`,drillTable(filteredRows().filter(r=>{
        const d=(r.poDate||r.prReceived||''); return d.startsWith(item.month)&&r.savings!==null;
      }),['id','title','buyer','method','prValue','cumulativePO','savings'])); }
  }});

  // Chart 3: Competitive vs Direct doughnut
  mou('chart-comp',{type:'doughnut',data:{
    labels:['Competitive','Direct/Other'],
    datasets:[{data:[kpi3.compCount,kpi3.dirCount],backgroundColor:['#009FDA','#f59e0b'],hoverOffset:6}]
  },options:{responsive:true,plugins:{legend:{position:'bottom',labels:{color:tc()}},
    tooltip:{callbacks:{label:ctx=>{
      const tot=kpi3.compCount+kpi3.dirCount;
      const pct=tot?Math.round(100*ctx.parsed/tot):0;
      const val=ctx.dataIndex===0?kpi3.compValue:kpi3.dirValue;
      return \` \${ctx.parsed} PRs (\${pct}%) \u2013 \${fmtUSD(val)}\`;
    }}}},
    onClick:(_e,el)=>{ if(!el.length)return; const ic=el[0].index===0;
      openModal(ic?'Competitive PRs':'Direct / Other PRs',drillTable(filteredRows().filter(r=>r.isCompetitive===ic),['id','title','buyer','method','prValue','stage'])); }
  }});

  // Chart 4: Plan compliance pie
  const pLabels=['Planned','Unplanned','N/A'];
  mou('chart-plan',{type:'pie',data:{
    labels:pLabels,
    datasets:[{data:pLabels.map(l=>kpi4.planCounts[l]),backgroundColor:['#22c55e','#ef4444','#94a3b8'],hoverOffset:6}]
  },options:{responsive:true,plugins:{legend:{position:'bottom',labels:{color:tc()}},
    tooltip:{callbacks:{label:ctx=>{
      const tot=pLabels.reduce((a,l)=>a+kpi4.planCounts[l],0);
      const pct=tot?Math.round(100*ctx.parsed/tot):0;
      return \` \${ctx.parsed} (\${pct}%) \u2013 \${fmtUSD(kpi4.planValues[pLabels[ctx.dataIndex]])}\`;
    }}}},
    onClick:(_e,el)=>{ if(!el.length)return; const b=pLabels[el[0].index];
      openModal(\`Plan \u2013 \${b}\`,drillTable(filteredRows().filter(r=>r.planBucket===b),['id','title','buyer','method','prValue','stage'])); }
  }});

  // Chart 5: Pipeline horizontal bar
  mou('chart-pipeline',{type:'bar',data:{
    labels:pipeline.map(d=>d.stage),
    datasets:[{label:'# Items',data:pipeline.map(d=>d.count),backgroundColor:'#009FDA',borderRadius:4}]
  },options:{indexAxis:'y',responsive:true,
    plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>\` \${ctx.parsed.x} items  |  \${fmtUSD(pipeline[ctx.dataIndex].value)}\`}}},
    scales:{x:{grid:{color:gc()},ticks:{color:tc()}},y:{grid:{display:false},ticks:{color:tc()}}},
    onClick:(_e,el)=>{ if(!el.length)return; const s=pipeline[el[0].index].stage;
      openModal(\`Pipeline \u2013 \${s}\`,drillTable(filteredRows().filter(r=>r.stage===s),['id','title','buyer','method','prValue','prReceived','poDate'])); }
  }});
}

// ── Drill table ────────────────────────────────────────────────────
const CL={id:'ID',title:'Description',buyer:'Buyer',method:'Method',prValue:'PR Value',
  cumulativePO:'Cumul. PO',savings:'Savings',stage:'Stage',prReceived:'PR Received',
  poDate:'PO Date',cycleTime:'Cycle (d)',isCompetitive:'Competitive?',planBucket:'Plan'};
function cf(k,v){
  if(v===null||v===undefined||v==='')return'\u2013';
  if(['prValue','cumulativePO','savings'].includes(k))return fmtUSD(v);
  if(k==='isCompetitive')return v?'Yes':'No';
  return String(v);
}
function drillTable(rows,cols){
  if(!rows.length)return'<p style="font-size:13px;color:var(--muted)">No records found.</p>';
  return\`<table class="drill"><thead><tr>\${cols.map(c=>\`<th>\${CL[c]||c}</th>\`).join('')}</tr></thead><tbody>\${
    rows.map(r=>'<tr>'+cols.map(c=>\`<td>\${cf(c,r[c])}</td>\`).join('')+'</tr>').join('')
  }</tbody></table>\`;
}
function openModal(title,html){
  document.getElementById('modal-title').textContent=title;
  document.getElementById('modal-content').innerHTML=html;
  document.getElementById('modal-ov').classList.add('open');
}
function closeModal(){ document.getElementById('modal-ov').classList.remove('open'); }
document.getElementById('modal-ov').addEventListener('click',e=>{ if(e.target===document.getElementById('modal-ov'))closeModal(); });
document.addEventListener('keydown',e=>{ if(e.key==='Escape')closeModal(); });

// ── Export PNG ─────────────────────────────────────────────────────
function exportChart(id){
  const c=CR[id]; if(!c)return;
  const a=document.createElement('a');
  a.href=c.toBase64Image('image/png',1); a.download=id+'.png'; a.click();
}

// ── Theme toggle ───────────────────────────────────────────────────
document.getElementById('theme-chk').addEventListener('change',function(){
  document.documentElement.setAttribute('data-theme',this.checked?'dark':'light');
  const rows=filteredRows(); const K=recompute(rows);
  renderCharts(K);
});

// ── Slicer pills ───────────────────────────────────────────────────
function makePills(containerId,items,filterKey){
  const wrap=document.getElementById(containerId);
  // "All" pill
  const all=document.createElement('button');
  all.className='pill on'; all.textContent='All';
  all.onclick=()=>{ AF[filterKey]=''; wrap.querySelectorAll('.pill').forEach(p=>p.classList.remove('on')); all.classList.add('on'); refresh(); };
  wrap.appendChild(all);
  items.forEach(item=>{
    const p=document.createElement('button');
    p.className='pill'; p.textContent=item;
    p.onclick=()=>{
      AF[filterKey]=item;
      wrap.querySelectorAll('.pill').forEach(p2=>p2.classList.remove('on'));
      p.classList.add('on');
      refresh();
    };
    wrap.appendChild(p);
  });
}

// ── Buyer drill (from workload card) ──────────────────────────────
function openBuyerDrill(buyer){
  openModal(\`Workload \u2013 \${buyer}\`,drillTable(filteredRows().filter(r=>r.buyer===buyer),['id','title','method','prValue','stage','prReceived','poDate','cycleTime']));
}

// ── Main refresh ───────────────────────────────────────────────────
function refresh(){
  const rows=filteredRows();
  const K=recompute(rows);
  renderCards(K,rows.length);
  renderCharts(K);
  renderBuyers(K.kpi5);
}

// ── Init ───────────────────────────────────────────────────────────
document.getElementById('last-updated').textContent=DASHBOARD_DATA.lastUpdated;
makePills('pills-year',  DASHBOARD_DATA.years,    'year');
makePills('pills-buyer', DASHBOARD_DATA.buyers,   'buyer');
makePills('pills-project',DASHBOARD_DATA.projects,'project');
refresh();
<\/script>
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
console.log(`  Rows: ${data.rows.length} | Buyers: ${data.buyers.length} | Stages: ${data.stages.length} | Projects: ${data.projects.length}`);

const html = generateHTML(data);
const outDir = path.dirname(OUTPUT_PATH);
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(OUTPUT_PATH, html, 'utf8');
console.log(`Dashboard written to: ${OUTPUT_PATH} (${(html.length / 1024).toFixed(1)} KB)`);
