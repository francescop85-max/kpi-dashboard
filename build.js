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

// ── Method consolidation ──────────────────────────────────────────────────────

const EXCLUDED_METHODS = new Set([
  'Call for Application (704)',
  'EOI (502)',
  'EOI (507)',
  'Invitation for Proposal (507)',
  'RFI (502)',
  'Contract Amendment',
  'Other',
]);

// Maps raw solicitation method → consolidated category (null = exclude row)
function consolidateMethod(raw) {
  if (!raw || EXCLUDED_METHODS.has(raw)) return null;
  if (/LTA|UN Award/i.test(raw))               return 'LTA';
  if (/\b(ITB|RFP|BAFO|RFQ)\b/i.test(raw))    return 'Formal Solicitation';
  if (/micro purchase|direct|very low value|re-utilisation/i.test(raw)) return 'Informal Solicitation';
  return null; // unmapped → exclude
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
    const rawMethod     = parseLookup(r['SollicitationMethod']);
    const method        = consolidateMethod(rawMethod);
    const stage         = parseLookup(r['ProcurementStage']);
    const planRaw       = parseLookup(r['PartofProcurementPlan']);
    const awardBasis    = parseLookup(r['AwardBasis']);
    const marketCat     = parseMultiLookup(r['MarketCategory']);
    const projRef       = parseMultiLookup(r['ProjectReference']);

    if (method === null) return null; // excluded method — drop row

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

    // Competitive vs Direct — LTA and any record with competitive award basis count as competitive
    const awardBasisCompetitive = /competitive/i.test(awardBasis);
    const isCompetitive = method === 'Formal Solicitation' || method === 'LTA' || awardBasisCompetitive;
    const isDirect      = !isCompetitive;

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
      isDirect,
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
  }).filter(r => r && !/cancelled/i.test(r.stage));
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
  }).sort((a, b) => b.avg - a.avg);

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

  // ── KPI 3: Competitive vs Direct Procurement vs Other ─────────────────────
  let compCount = 0, compValue = 0, dirCount = 0, dirValue = 0, otherCount = 0, otherValue = 0;
  for (const r of rows) {
    if (r.isCompetitive) {
      compCount++;
      if (r.prValue) compValue += r.prValue;
    } else if (r.isDirect) {
      dirCount++;
      if (r.prValue) dirValue += r.prValue;
    } else {
      otherCount++;
      if (r.prValue) otherValue += r.prValue;
    }
  }
  const kpi3 = { compCount, compValue, dirCount, dirValue, otherCount, otherValue };

  // ── KPI 4: Plan Compliance ────────────────────────────────────────────────
  const planCounts = { Planned: 0, Unplanned: 0, 'N/A': 0 };
  const planValues = { Planned: 0, Unplanned: 0, 'N/A': 0 };
  for (const r of rows) {
    planCounts[r.planBucket]++;
    if (r.prValue) planValues[r.planBucket] += r.prValue;
  }
  const kpi4 = { planCounts, planValues };

  // ── KPI 5: Team Workload vs Target ────────────────────────────────────────
  const WL_TARGET = 15;
  const WL_EXCL = new Set(['Francesco Perini', 'Adrian Horvath', 'Weng', 'Unknown']);
  const wlExcluded = b => WL_EXCL.has(b) || /weng/i.test(b);

  const wlMoMap = {};
  for (const r of rows) {
    const b = r.buyer || 'Unknown';
    if (wlExcluded(b)) continue;
    const d = r.prReceived;
    if (!d) continue;
    const mo = fmtDate(d).slice(0, 7);
    if (!wlMoMap[mo]) wlMoMap[mo] = {};
    wlMoMap[mo][b] = (wlMoMap[mo][b] || 0) + 1;
  }
  const wlMonthly = Object.entries(wlMoMap)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([month, bc]) => {
      const nb = Object.keys(bc).length;
      const tot = Object.values(bc).reduce((a, x) => a + x, 0);
      return { month, avg: nb ? parseFloat((tot / nb).toFixed(1)) : 0, numBuyers: nb, total: tot };
    });

  const CLOSED_RE = /closed|cancelled|po.issued/i;
  const activeBM = {};
  for (const r of rows) {
    const b = r.buyer || 'Unknown';
    if (wlExcluded(b) || CLOSED_RE.test(r.stage || '')) continue;
    activeBM[b] = (activeBM[b] || 0) + 1;
  }
  const nActive = Object.keys(activeBM).length || 1;
  const totalActive = Object.values(activeBM).reduce((a, x) => a + x, 0);
  const kpi5 = {
    wlMonthly, currentAvg: parseFloat((totalActive / nActive).toFixed(1)),
    numActiveBuyers: nActive, totalActive,
    monthsOver: wlMonthly.filter(m => m.avg > WL_TARGET).length,
    wlTarget: WL_TARGET, activeBuyerMap: activeBM,
  };

  // ── KPI 6: PR Assigned → Solicitation Issued ─────────────────────────────
  const assignToSolByMethod = {};
  for (const r of rows) {
    const days = dateDiffDays(r.dateAssigned, r.solIssued);
    if (days === null || days < 0) continue;
    const m = r.method || 'Unknown';
    if (!assignToSolByMethod[m]) assignToSolByMethod[m] = [];
    assignToSolByMethod[m].push(days);
  }
  const kpi6 = Object.entries(assignToSolByMethod).map(([method, times]) => {
    const avg = Math.round(times.reduce((a, b) => a + b, 0) / times.length);
    return { method, avg, count: times.length };
  }).sort((a, b) => b.avg - a.avg);
  const allA2S = Object.values(assignToSolByMethod).flat();
  const avgAssignToSol = allA2S.length
    ? Math.round(allA2S.reduce((a, b) => a + b, 0) / allA2S.length)
    : 0;

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

  // ── Buyer list (exclude admin/mgmt names from filter) ─────────────────────
  const BUYER_FILTER_EXCL = /perini|horvath|weng/i;
  const buyers = [...new Set(rows.map(r => r.buyer).filter(b => b && !BUYER_FILTER_EXCL.test(b)))].sort();

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

  // Solicitation methods list
  const methods = [...new Set(rows.map(r => r.method).filter(Boolean))].sort();

  // ── Last modified date ────────────────────────────────────────────────────
  const modDates = rows.map(r => r.modified).filter(Boolean);
  const lastUpdated = modDates.length
    ? fmtDate(new Date(Math.max(...modDates.map(d => d.getTime()))))
    : fmtDate(new Date());

  // ── KPI 3 trend: monthly competitive mix ────────────────────────────────
  const k3map = {};
  for (const r of rows) {
    const d = r.prReceived || r.poDate; if (!d) continue;
    const month = fmtDate(d).slice(0, 7);
    if (!k3map[month]) k3map[month] = { formal: 0, lta: 0, dir: 0 };
    if (r.method === 'LTA') k3map[month].lta++;
    else if (r.isCompetitive) k3map[month].formal++;
    else k3map[month].dir++;
  }
  const kpi3Trend = Object.entries(k3map).sort(([a],[b]) => a.localeCompare(b)).map(([month, d]) => {
    const total = d.formal + d.lta + d.dir;
    return { month, formal: d.formal, lta: d.lta, dir: d.dir, total,
      formalPct: total ? Math.round(100*d.formal/total) : 0,
      ltaPct:    total ? Math.round(100*d.lta/total)    : 0,
      dirPct:    total ? Math.round(100*d.dir/total)    : 0,
      compPct:   total ? Math.round(100*(d.formal+d.lta)/total) : 0 };
  });

  // ── KPI 4 trend: monthly plan compliance mix ─────────────────────────────
  const k4map = {};
  for (const r of rows) {
    const d = r.prReceived || r.poDate; if (!d) continue;
    const month = fmtDate(d).slice(0, 7);
    if (!k4map[month]) k4map[month] = { Planned: 0, Unplanned: 0, 'N/A': 0 };
    k4map[month][r.planBucket]++;
  }
  const kpi4Trend = Object.entries(k4map).sort(([a],[b]) => a.localeCompare(b)).map(([month, d]) => {
    const total = d.Planned + d.Unplanned + d['N/A'];
    return { month, planned: d.Planned, unplanned: d.Unplanned, na: d['N/A'], total,
      plannedPct:   total ? Math.round(100*d.Planned/total)   : 0,
      unplannedPct: total ? Math.round(100*d.Unplanned/total) : 0,
      naPct:        total ? Math.round(100*d['N/A']/total)    : 0 };
  });

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

  return { kpi1, kpi2, kpi3, kpi4, kpi5, kpi6, pipeline, years, buyers, stages, methods, projects, lastUpdated, rows: serialRows, avgCycle, avgAssignToSol, kpi3Trend, kpi4Trend };
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
  <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js" crossorigin="anonymous"><\/script>
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

    /* Multi-check dropdown */
    .mcd-wrap { position: relative; display: inline-block; }
    .mcd-btn {
      padding: 4px 14px 4px 10px; border-radius: 20px; font-size: 11.5px; font-weight: 600;
      border: 1.5px solid var(--border); background: var(--card); color: var(--text);
      cursor: pointer; white-space: nowrap; display: flex; align-items: center; gap: 6px;
      font-family: inherit;
    }
    .mcd-btn:hover, .mcd-btn.open { border-color: var(--blue); color: var(--blue); }
    .mcd-btn.has-selection { background: var(--navy); border-color: var(--navy); color: #fff; }
    .mcd-panel {
      display: none; position: absolute; top: calc(100% + 4px); left: 0; z-index: 200;
      background: var(--card); border: 1.5px solid var(--border); border-radius: 10px;
      box-shadow: 0 8px 24px rgba(0,0,0,.15); min-width: 220px; max-width: 320px;
      padding: 8px 0;
    }
    .mcd-wrap.open .mcd-panel { display: block; }
    .mcd-search {
      display: block; width: calc(100% - 16px); margin: 0 8px 6px;
      padding: 5px 10px; border: 1px solid var(--border); border-radius: 6px;
      font-size: 12px; background: var(--bg); color: var(--text); font-family: inherit;
    }
    .mcd-search:focus { outline: none; border-color: var(--blue); }
    .mcd-list { max-height: 220px; overflow-y: auto; }
    .mcd-item {
      display: flex; align-items: center; gap: 8px; padding: 5px 12px;
      font-size: 12px; cursor: pointer; color: var(--text);
    }
    .mcd-item:hover { background: rgba(0,159,218,.08); }
    .mcd-item input[type=checkbox] { accent-color: var(--navy); width: 13px; height: 13px; cursor: pointer; }
    .mcd-clear {
      display: block; width: calc(100% - 16px); margin: 6px 8px 0;
      padding: 4px; font-size: 11px; text-align: center; cursor: pointer;
      color: var(--blue); border: none; background: none; font-family: inherit;
      border-top: 1px solid var(--border); padding-top: 8px;
    }
    .mcd-clear:hover { text-decoration: underline; }

    /* KPI narrative */
    .kpi-narrative { font-size: 11px; color: var(--muted); margin-top: 4px; line-height: 1.5; font-style: italic; }

    /* ── KPI strip ── */
    .kpi-row{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:18px;}
    .kpi-card{background:var(--card);border-radius:12px;padding:16px 20px;box-shadow:var(--shadow);border-left:4px solid var(--blue);display:flex;flex-direction:column;gap:3px;}
    .kpi-lbl{font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);font-weight:700;}
    .kpi-val{font-size:28px;font-weight:800;line-height:1.1;}
    .kpi-sub{font-size:11px;color:var(--muted);}

    /* ── Charts ── */
    .content{max-width:1200px;margin:0 auto;padding:20px 20px;}
    .grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px;}
    .full{grid-column:1/-1;}
    .chart-card{background:var(--card);border-radius:12px;padding:18px 20px;box-shadow:var(--shadow);position:relative;}
    .chart-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);margin-bottom:10px;}
    .kpi-desc{font-size:11.5px;color:#475569;line-height:1.5;margin-bottom:12px;padding-bottom:10px;border-bottom:1px solid #e2e8f0;}
    [data-theme="dark"] .kpi-desc{color:#94a3b8;border-bottom-color:#334155;}
    .disclaimer{font-size:10.5px;color:#64748b;line-height:1.55;font-style:italic;background:#f8fafc;border:1px solid #e2e8f0;border-left:3px solid #94a3b8;padding:7px 10px;border-radius:4px;margin-top:12px;}
    [data-theme="dark"] .disclaimer{background:#1e293b;color:#94a3b8;border-color:#334155;border-left-color:#64748b;}
    .export-btn{position:absolute;top:12px;right:12px;font-size:10px;padding:2px 8px;background:rgba(0,159,218,.1);border:1px solid #009FDA;color:#009FDA;border-radius:4px;cursor:pointer;font-family:inherit;}
    .export-btn:hover{background:rgba(0,159,218,.25);}
    canvas{max-height:300px;}

    /* ── Buyer workload cards ── */
    #workload-summary [data-theme="dark"] div[style*="--navy"]{color:var(--blue);}

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

    .print-btn{padding:5px 14px;font-size:12px;font-weight:600;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.4);color:#fff;border-radius:6px;cursor:pointer;font-family:inherit;transition:background .15s;}
    .print-btn:hover{background:rgba(255,255,255,.28);}

    @media(max-width:768px){
      .kpi-row{grid-template-columns:1fr 1fr;}
      .grid2{grid-template-columns:1fr;}
      .full{grid-column:auto;}
    }
    @media print{
      .filter-bar,.export-btn,.print-btn,.toggle-wrap,#modal-ov{display:none!important;}
      .hdr{background:#003366!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
      .content{padding:10px;}
      canvas{max-height:220px!important;}
      .grid2{gap:10px;}
      body{background:#fff;}
    }
  </style>
</head>
<body>

<!-- ── HEADER ── -->
<header class="hdr">
  <div class="hdr-left">
    <div>
      <div class="hdr-title">Ukraine Procurement KPI Dashboard</div>
      <div class="hdr-sub">Food and Agriculture Organization of the United Nations · Last updated: <span id="last-updated"></span></div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:12px;">
    <button class="print-btn" onclick="openAbout()" style="background:rgba(255,255,255,.12);border-color:rgba(255,255,255,.35);">ℹ About</button>
    <button class="print-btn" onclick="generateReport()" style="background:rgba(0,159,218,.25);border-color:#009FDA;">📊 Management Report</button>
    <button class="print-btn" onclick="window.print()">🖨️ Print / PDF</button>
    <label class="toggle-wrap" title="Toggle dark mode">
      <span style="font-size:13px;">☀️</span>
      <input type="checkbox" id="theme-chk"/>
      <div class="toggle-track"></div>
      <span style="font-size:13px;">🌙</span>
    </label>
  </div>
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
    <div class="filter-label">🔖 Method</div>
    <div class="mcd-wrap" id="mcd-method" data-filterkey="methods" data-labelall="All Methods">
      <button class="mcd-btn" onclick="toggleMcd('mcd-method')">All Methods <span>▾</span></button>
      <div class="mcd-panel">
        <input class="mcd-search" type="search" placeholder="Search…" oninput="filterMcdItems('mcd-method',this.value)"/>
        <div class="mcd-list" id="mcd-method-list"></div>
        <button class="mcd-clear" onclick="clearMcd('mcd-method')">Clear selection</button>
      </div>
    </div>
  </div>
  <div class="filter-group">
    <div class="filter-label">📁 Project</div>
    <div class="mcd-wrap" id="mcd-project" data-filterkey="projects" data-labelall="All Projects">
      <button class="mcd-btn" onclick="toggleMcd('mcd-project')">All Projects <span>▾</span></button>
      <div class="mcd-panel">
        <input class="mcd-search" type="search" placeholder="Search…" oninput="filterMcdItems('mcd-project',this.value)"/>
        <div class="mcd-list" id="mcd-project-list"></div>
        <button class="mcd-clear" onclick="clearMcd('mcd-project')">Clear selection</button>
      </div>
    </div>
  </div>
</div>

<!-- ── MAIN ── -->
<div class="content">

  <!-- KPI summary strip -->
  <div class="kpi-row" id="kpi-row"></div>

  <!-- KPI 1 Trend: Cycle Time over Time -->
  <div class="grid2">
    <div class="chart-card full">
      <button class="export-btn" onclick="exportChart('chart-cycle-trend')">PNG</button>
      <div class="chart-title" style="padding-right:60px;">KPI 1 – Cycle Time Trend (avg days per month)</div><div style="font-size:10px;font-weight:400;color:var(--blue);margin-bottom:8px;">● Click a point to see records</div>
      <div class="kpi-desc">Shows how the average procurement cycle time (PR received to PO issued) evolves month by month. Points are colour-coded: <span style="color:#22c55e;font-weight:600">green ≤30d</span>, <span style="color:#f59e0b;font-weight:600">amber ≤90d</span>, <span style="color:#ef4444;font-weight:600">red &gt;90d</span>. Use the filter below to focus on a specific solicitation method.</div>
      <div style="margin:8px 0 12px;">
        <div class="filter-label" style="margin-bottom:5px;">Filter by Solicitation Method</div>
        <div class="pills" id="pills-kpi1method"></div>
      </div>
      <canvas id="chart-cycle-trend"></canvas>
    </div>
  </div>

  <!-- KPI 2: Savings over time (full width) -->
  <div class="grid2">
    <div class="chart-card full">
      <button class="export-btn" onclick="exportChart('chart-savings')">PNG</button>
      <div class="chart-title">KPI 2 – Savings per Month (USD) with Average</div>
      <div class="kpi-desc">Tracks the cumulative procurement savings achieved each month, calculated as the difference between the estimated PR value and the final PO value. Positive figures indicate money saved against the initial budget estimate.</div>
      <canvas id="chart-savings"></canvas>
      <div class="disclaimer">&#9432; Savings are calculated as the gap between the requisitioner's estimated value and the final contracted amount. This assumes the original estimate was realistic and market-facing. Where estimates were conservative or inflated, the savings figure may not reflect actual market efficiency.</div>
    </div>
  </div>

  <!-- KPI 3 + KPI 4 side by side -->
  <div class="grid2">
    <div class="chart-card">
      <button class="export-btn" onclick="exportChart('chart-comp')">PNG</button>
      <div class="chart-title">KPI 3 – Competitive vs Direct</div>
      <div class="kpi-desc">Shows the proportion of PRs awarded through competitive solicitation (ITB, RFP, BAFO, RFQ) or processed under an existing LTA (itself competitively awarded) versus direct/informal procurement.</div>
      <canvas id="chart-comp"></canvas>
      <div class="disclaimer">&#9432; <strong>Other</strong> includes PRs where the solicitation method does not map to a competitive or direct procurement category — typically older records created before the Award Basis field was introduced, or methods not yet classified in the system.</div>
    </div>
    <div class="chart-card">
      <button class="export-btn" onclick="exportChart('chart-plan')">PNG</button>
      <div class="chart-title">KPI 4 – Plan Compliance</div>
      <div class="kpi-desc">Tracks whether each PR was anticipated in FAO's procurement planning module. Only PRs recorded after plan tracking began are included — earlier records (N/A) are excluded from the calculation.</div>
      <canvas id="chart-plan"></canvas>
      <div class="disclaimer">&#9432; Only PRs with a Planned or Unplanned classification are included. Records pre-dating the introduction of procurement plan tracking are excluded from this KPI.</div>
    </div>
  </div>

  <!-- KPI 3 + KPI 4 trends -->
  <div class="grid2">
    <div class="chart-card">
      <button class="export-btn" onclick="exportChart('chart-comp-trend')">PNG</button>
      <div class="chart-title" style="padding-right:60px;">KPI 3 – Competitive Mix Trend (% per month) <span id="kpi3-trend-badge"></span></div>
      <div style="font-size:10px;font-weight:400;color:var(--blue);margin-bottom:8px;">● Click a bar to see records</div>
      <div class="kpi-desc">Shows how the monthly mix of Formal Solicitation, LTA, and Informal procurement evolved over time as a percentage of PRs received.</div>
      <canvas id="chart-comp-trend"></canvas>
    </div>
    <div class="chart-card">
      <button class="export-btn" onclick="exportChart('chart-plan-trend')">PNG</button>
      <div class="chart-title" style="padding-right:60px;">KPI 4 – Plan Compliance Trend (% per month) <span id="kpi4-trend-badge"></span></div>
      <div style="font-size:10px;font-weight:400;color:var(--blue);margin-bottom:8px;">● Click a bar to see records</div>
      <div class="kpi-desc">Shows how the monthly share of Planned, Unplanned, and N/A PRs evolved over time.</div>
      <canvas id="chart-plan-trend"></canvas>
    </div>
  </div>

  <!-- KPI 5: Team Workload (full width) -->
  <div class="grid2">
    <div class="chart-card full">
      <button class="export-btn" onclick="exportChart('chart-workload')">PNG</button>
      <div class="chart-title" style="padding-right:60px;">KPI 5 – Team Workload vs Target (15 PRs / buyer)</div><div style="font-size:10px;font-weight:400;color:var(--blue);margin-bottom:8px;">● Click a bar to see records</div>
      <div class="kpi-desc">Monitors the average number of active Purchase Requests managed per procurement officer over time, compared to a standard workload target of 15 PRs per buyer. Bars in red indicate quarters where the team average exceeded the target — a key indicator for staffing adequacy.</div>
      <div id="workload-summary"></div>
      <canvas id="chart-workload"></canvas>
    </div>
  </div>

  <!-- KPI 6 Trend: Prep Time over Time -->
  <div class="grid2">
    <div class="chart-card full">
      <button class="export-btn" onclick="exportChart('chart-prep-trend')">PNG</button>
      <div class="chart-title" style="padding-right:60px;">KPI 6 – Prep Time Trend (avg days per month)</div><div style="font-size:10px;font-weight:400;color:var(--blue);margin-bottom:8px;">● Click a point to see records</div>
      <div class="kpi-desc">Shows how the average preparation time (PR assigned → solicitation issued) evolves month by month. Points are colour-coded: <span style="color:#22c55e;font-weight:600">green ≤30d</span>, <span style="color:#f59e0b;font-weight:600">amber ≤45d</span>, <span style="color:#ef4444;font-weight:600">red &gt;45d</span>. Use the filter below to focus on a specific solicitation method.</div>
      <div style="margin:8px 0 12px;">
        <div class="filter-label" style="margin-bottom:5px;">Filter by Solicitation Method</div>
        <div class="pills" id="pills-kpi6method"></div>
      </div>
      <canvas id="chart-prep-trend"></canvas>
    </div>
  </div>

</div>

<footer>Data source: FAO Ukraine Procurement Tracking SharePoint List &nbsp;·&nbsp; Refreshed daily via GitHub Actions</footer>

<!-- ── MODAL ── -->
<!-- ── ABOUT MODAL ── -->
<div id="about-ov" role="dialog" aria-modal="true" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:200;align-items:center;justify-content:center;">
  <div style="background:var(--card);border-radius:16px;max-width:92vw;width:820px;max-height:90vh;overflow-y:auto;padding:32px 36px;position:relative;box-shadow:0 20px 60px rgba(0,0,0,.25);">
    <button onclick="closeAbout()" aria-label="Close" style="position:absolute;top:14px;right:18px;font-size:22px;cursor:pointer;color:var(--muted);background:none;border:none;">&times;</button>
    <!-- Purpose -->
    <div style="display:flex;align-items:center;gap:14px;margin-bottom:20px;">
      <div style="background:var(--navy);border-radius:10px;padding:10px 14px;flex-shrink:0;">
        <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
      </div>
      <div>
        <div style="font-size:18px;font-weight:800;color:var(--navy);letter-spacing:-.3px;">About this Dashboard</div>
        <div style="font-size:11px;color:var(--muted);margin-top:2px;">FAO Ukraine · Procurement Performance</div>
      </div>
    </div>
    <p style="font-size:13.5px;line-height:1.7;color:var(--text);margin-bottom:26px;padding-bottom:22px;border-bottom:1px solid var(--border);">
      This dashboard monitors the procurement performance of the FAO Ukraine office, tracking key efficiency and compliance indicators across all Purchase Requests (PRs) managed by the procurement team. It is designed to support management oversight, identify bottlenecks, and demonstrate accountability in the use of project funds.
    </p>
    <!-- KPI grid -->
    <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.9px;color:var(--muted);margin-bottom:14px;">Performance Indicators</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">
      ${[
        ['1','Procurement Cycle Time','Measures the average number of days between receipt of a Purchase Request and issuance of the Purchase Order. A shorter cycle time reflects a more efficient procurement process. The target is 60 days. Tracked monthly by solicitation method to identify where delays concentrate.'],
        ['2','Procurement Savings','Tracks the financial savings achieved by the procurement team, calculated as the difference between the requisitioner\'s estimated budget and the final contracted amount. Positive savings indicate effective market engagement and negotiation. Reported monthly with cumulative totals.'],
        ['3','Competitive vs Direct Procurement','Shows what proportion of PRs were processed through a competitive solicitation process (formal tender or existing LTA) versus direct/informal procurement. A higher share of competitive procurement reflects stronger value-for-money practices and compliance with FAO procurement rules.'],
        ['4','Procurement Plan Compliance','Measures whether PRs were anticipated in FAO\'s annual procurement plan. A higher share of planned procurement signals stronger budget forecasting and reduces reactive, emergency procurement. Tracked only for the period since plan monitoring was introduced.'],
        ['5','Team Workload','Monitors the average number of active PRs managed per procurement officer, compared to a standard target of 15 PRs per buyer. This KPI supports staffing decisions and helps flag periods of excessive workload that may affect processing quality or speed.'],
        ['6','Solicitation Preparation Time','Measures the average number of days between assignment of a PR to a procurement officer and issuance of the solicitation document. This isolates the internal preparation phase of the process and helps identify bottlenecks in document drafting and approvals.'],
      ].map(([n,title,desc])=>`
      <div style="border:1px solid var(--border);border-radius:10px;padding:16px 18px;display:flex;gap:14px;align-items:flex-start;">
        <div style="background:var(--navy);color:#fff;font-size:13px;font-weight:800;border-radius:8px;min-width:30px;height:30px;display:flex;align-items:center;justify-content:center;flex-shrink:0;">${n}</div>
        <div>
          <div style="font-size:12px;font-weight:700;color:var(--text);margin-bottom:5px;">KPI ${n} – ${title}</div>
          <div style="font-size:11.5px;color:var(--muted);line-height:1.6;">${desc}</div>
        </div>
      </div>`).join('')}
    </div>
    <div style="margin-top:22px;padding-top:16px;border-top:1px solid var(--border);font-size:11px;color:var(--muted);text-align:center;">
      Data source: FAO Ukraine Procurement Tracking SharePoint List &nbsp;·&nbsp; Refreshed daily
    </div>
  </div>
</div>

<div id="modal-ov" role="dialog" aria-modal="true">
  <div id="modal-box">
    <button id="modal-close" onclick="closeModal()" aria-label="Close">&times;</button>
    <div id="modal-title"></div>
    <div id="modal-content"></div>
  </div>
</div>

<script>
const DASHBOARD_DATA = ${dataJson};

// ── Plugin setup ───────────────────────────────────────────────────
Chart.register(ChartDataLabels);
Chart.defaults.set('plugins.datalabels', { display: false });

// ── State ──────────────────────────────────────────────────────────
let AF = { year: '', buyer: '', methods: new Set(), projects: new Set() };
const WL_EXCL_C = new Set(['Francesco Perini', 'Adrian Horvath', 'Weng', 'Unknown']);
const wlExclC = b => WL_EXCL_C.has(b) || /weng/i.test(b);
const CR = {};

// ── Helpers ────────────────────────────────────────────────────────
function fmt(n,d=0){ return n==null?'\u2013':n.toLocaleString('en-US',{maximumFractionDigits:d,minimumFractionDigits:d}); }
function fmtUSD(n){ return n==null?'\u2013':'$'+fmt(n,0); }
function isDark(){ return document.documentElement.getAttribute('data-theme')==='dark'; }
function gc(){ return isDark()?'rgba(255,255,255,.07)':'rgba(0,0,0,.07)'; }
function tc(){ return isDark()?'#94a3b8':'#64748b'; }

// ── Filter rows ────────────────────────────────────────────────────
function filteredRows() {
  return DASHBOARD_DATA.rows.filter(r => {
    if (AF.year   && String(r.year) !== String(AF.year))   return false;
    if (AF.buyer  && r.buyer !== AF.buyer)          return false;
    if (AF.methods.size  && !AF.methods.has(r.method || ''))  return false;
    if (AF.projects.size) {
      const rp = (r.projRef || '').split(';').map(p => p.trim()).filter(Boolean);
      if (!rp.some(p => AF.projects.has(p))) return false;
    }
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
    return{method,avg,count:t.length};
  }).sort((a,b)=>b.avg-a.avg);
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
  let cc=0,cv=0,dc=0,dv=0,oc=0,ov=0;
  for(const r of rows){
    if(r.isCompetitive){cc++;if(r.prValue)cv+=r.prValue;}
    else if(r.isDirect){dc++;if(r.prValue)dv+=r.prValue;}
    else{oc++;if(r.prValue)ov+=r.prValue;}
  }
  const kpi3={compCount:cc,compValue:cv,dirCount:dc,dirValue:dv,otherCount:oc,otherValue:ov};

  // KPI 4
  const pc={Planned:0,Unplanned:0,'N/A':0},pv={Planned:0,Unplanned:0,'N/A':0};
  for(const r of rows){ pc[r.planBucket]++; if(r.prValue)pv[r.planBucket]+=r.prValue; }
  const kpi4={planCounts:pc,planValues:pv};

  // KPI 5: Team Workload
  const wlMoMapC={};
  for(const r of rows){
    const b=r.buyer||'Unknown'; if(wlExclC(b))continue;
    const d=r.prReceived; if(!d)continue;
    const mo=d.slice(0,7);
    if(!wlMoMapC[mo])wlMoMapC[mo]={};
    wlMoMapC[mo][b]=(wlMoMapC[mo][b]||0)+1;
  }
  const wlMonthlyC=Object.entries(wlMoMapC).sort(([a],[b])=>a.localeCompare(b))
    .map(([month,bc])=>{ const nb=Object.keys(bc).length; const tot=Object.values(bc).reduce((a,x)=>a+x,0); return{month,avg:nb?parseFloat((tot/nb).toFixed(1)):0,numBuyers:nb,total:tot}; });
  const CLOSED_REC=/closed|cancelled|po.issued/i;
  const aBMC={};
  for(const r of rows){ const b=r.buyer||'Unknown'; if(wlExclC(b)||CLOSED_REC.test(r.stage||''))continue; aBMC[b]=(aBMC[b]||0)+1; }
  const nActC=Object.keys(aBMC).length||1;
  const totActC=Object.values(aBMC).reduce((a,x)=>a+x,0);
  const kpi5={wlMonthly:wlMonthlyC,currentAvg:parseFloat((totActC/nActC).toFixed(1)),numActiveBuyers:nActC,totalActive:totActC,monthsOver:wlMonthlyC.filter(m=>m.avg>15).length,wlTarget:15,activeBuyerMap:aBMC};

  // Pipeline
  const pm={};
  for(const r of rows){ const s=r.stage||'Unknown'; if(!pm[s])pm[s]={count:0,value:0}; pm[s].count++; if(r.prValue)pm[s].value+=r.prValue; }
  const SO=['New PR','Assigned','Solicitation Issued','Offers Received','Evaluation','Award Recommendation','PO Issued','Closed','Cancelled'];
  const pipeline=Object.entries(pm).map(([stage,d])=>({stage,...d}))
    .sort((a,b)=>{ const ia=SO.findIndex(s=>a.stage.toLowerCase().includes(s.toLowerCase())); const ib=SO.findIndex(s=>b.stage.toLowerCase().includes(s.toLowerCase())); if(ia!==-1&&ib!==-1)return ia-ib; if(ia!==-1)return-1; if(ib!==-1)return 1; return b.count-a.count; });

  // KPI 6: Assigned → Sol Issued
  const a2sm={};
  for(const r of rows){
    if(!r.dateAssigned||!r.solIssued) continue;
    const days=Math.round((new Date(r.solIssued)-new Date(r.dateAssigned))/86400000);
    if(days<0) continue;
    const m=r.method||'Unknown'; if(!a2sm[m]) a2sm[m]=[]; a2sm[m].push(days);
  }
  const kpi6=Object.entries(a2sm).map(([method,t])=>({method,avg:Math.round(t.reduce((a,b)=>a+b,0)/t.length),count:t.length})).sort((a,b)=>b.avg-a.avg);
  const avgA2S=Object.values(a2sm).flat().reduce((s,v,_,arr)=>s+v/arr.length,0)|0;

  // KPI 1 trend: monthly avg cycle time by method
  const ctMap={};
  for(const r of rows){
    if(r.cycleTime===null||r.cycleTime<0||!r.prReceived) continue;
    const month=r.prReceived.slice(0,7);
    const m=r.method||'Unknown';
    if(!ctMap[month]) ctMap[month]={};
    if(!ctMap[month][m]) ctMap[month][m]={sum:0,count:0};
    ctMap[month][m].sum+=r.cycleTime; ctMap[month][m].count++;
  }
  const ctMonths=Object.keys(ctMap).sort();
  const kpi1Trend={
    months:ctMonths,
    data:ctMonths.map(month=>{
      const bm={}; let ts=0,cnt=0;
      for(const[m,d] of Object.entries(ctMap[month])){bm[m]=Math.round(d.sum/d.count);ts+=d.sum;cnt+=d.count;}
      return{month,byMethod:bm,avg:cnt?Math.round(ts/cnt):null};
    }),
  };

  // KPI 6 trend: monthly avg prep time (PR assigned → sol issued) by method
  const ptMap={};
  for(const r of rows){
    if(!r.dateAssigned||!r.solIssued) continue;
    const days=Math.round((new Date(r.solIssued)-new Date(r.dateAssigned))/86400000);
    if(days<0) continue;
    const month=r.dateAssigned.slice(0,7);
    const m=r.method||'Unknown';
    if(!ptMap[month]) ptMap[month]={};
    if(!ptMap[month][m]) ptMap[month][m]={sum:0,count:0};
    ptMap[month][m].sum+=days; ptMap[month][m].count++;
  }
  const ptMonths=Object.keys(ptMap).sort();
  const kpi6Trend={
    months:ptMonths,
    data:ptMonths.map(month=>{
      const bm={}; let ts=0,cnt=0;
      for(const[m,d] of Object.entries(ptMap[month])){bm[m]=Math.round(d.sum/d.count);ts+=d.sum;cnt+=d.count;}
      return{month,byMethod:bm,avg:cnt?Math.round(ts/cnt):null};
    }),
  };

  // KPI 3 trend
  const k3mc={};
  for(const r of rows){
    const d=r.prReceived||r.poDate; if(!d) continue;
    const month=d.slice(0,7);
    if(!k3mc[month]) k3mc[month]={formal:0,lta:0,dir:0};
    if(r.method==='LTA') k3mc[month].lta++;
    else if(r.isCompetitive) k3mc[month].formal++;
    else k3mc[month].dir++;
  }
  const kpi3Trend=Object.entries(k3mc).sort(([a],[b])=>a.localeCompare(b)).map(([month,d])=>{
    const total=d.formal+d.lta+d.dir;
    return{month,formal:d.formal,lta:d.lta,dir:d.dir,total,
      formalPct:total?Math.round(100*d.formal/total):0,
      ltaPct:total?Math.round(100*d.lta/total):0,
      dirPct:total?Math.round(100*d.dir/total):0,
      compPct:total?Math.round(100*(d.formal+d.lta)/total):0};
  });
  // KPI 4 trend
  const k4mc={};
  for(const r of rows){
    const d=r.prReceived||r.poDate; if(!d) continue;
    const month=d.slice(0,7);
    if(!k4mc[month]) k4mc[month]={Planned:0,Unplanned:0,'N/A':0};
    k4mc[month][r.planBucket]++;
  }
  const kpi4Trend=Object.entries(k4mc).sort(([a],[b])=>a.localeCompare(b)).map(([month,d])=>{
    const total=d.Planned+d.Unplanned+d['N/A'];
    return{month,planned:d.Planned,unplanned:d.Unplanned,na:d['N/A'],total,
      plannedPct:total?Math.round(100*d.Planned/total):0,
      unplannedPct:total?Math.round(100*d.Unplanned/total):0,
      naPct:total?Math.round(100*d['N/A']/total):0};
  });

  return{kpi1,kpi2,kpi3,kpi4,kpi5,pipeline,avgCycle,kpi6,avgAssignToSol:avgA2S,kpi1Trend,kpi6Trend,kpi3Trend,kpi4Trend};
}

// ── KPI Cards ──────────────────────────────────────────────────────
function renderCards(K, total) {
  const { kpi1, kpi2, kpi3, kpi4, avgCycle, kpi6, avgAssignToSol } = K;
  const compTotal = kpi3.compCount + kpi3.dirCount + kpi3.otherCount;
  // Exclude unclassified (Other) from the competitive % — those records pre-date the Award Basis field
  const compBase = kpi3.compCount + kpi3.dirCount;
  const compPct = compBase ? Math.round(100 * kpi3.compCount / compBase) : 0;
  const dirPct  = compBase ? Math.round(100 * kpi3.dirCount  / compBase) : 0;

  const target = 60; const cycDiff = avgCycle - target;
  const fastestM = [...kpi1].sort((a,b)=>a.avg-b.avg)[0];
  const compTarget = 70;

  const narr = [
    (() => { const closed=filteredRows().filter(r=>/closed|po issued/i.test(r.stage||'')).length; return \`\${total-closed} active, \${closed} closed or completed.\`; })(),
    cycDiff>0 ? \`\${cycDiff}d above the \${target}-day benchmark.\${fastestM?' Fastest: '+fastestM.method+' ('+fastestM.avg+'d).':''}\`
              : \`Within target — \${Math.abs(cycDiff)}d below the \${target}-day benchmark.\${fastestM?' Fastest: '+fastestM.method+' ('+fastestM.avg+'d).':''}\`,
    \`\${fmtUSD(kpi2.sumPos)} positive, \${fmtUSD(Math.abs(kpi2.sumNeg))} negative. Avg monthly: \${fmtUSD(kpi2.avgMonthly)}.\`,
    compPct>=compTarget ? \`\${compPct}% competitive (target \${compTarget}%, excl. \${kpi3.otherCount} unclassified). Direct: \${dirPct}% (\${kpi3.dirCount} PRs, \${fmtUSD(kpi3.dirValue)}).\`
                        : \`\${compPct}% competitive — below \${compTarget}% target (excl. \${kpi3.otherCount} unclassified). Direct: \${dirPct}% (\${kpi3.dirCount} PRs).\`,
    avgAssignToSol > 0 ? \`Avg \${avgAssignToSol} days from PR assignment to solicitation launch.\${kpi6[0]?' Slowest: '+kpi6.slice(-1)[0].method+' ('+kpi6.slice(-1)[0].avg+'d).':''}\`
                       : 'No data available for this period.',
  ];

  // Status colours: red = target missed, amber = borderline, blue/navy = on target
  const cycCol = avgCycle > target ? (avgCycle > target * 1.25 ? 'var(--red)' : 'var(--amber)') : 'var(--blue)';
  const savCol = kpi2.net < 0 ? 'var(--red)' : kpi2.net === 0 ? 'var(--amber)' : 'var(--blue)';
  const cmpCol = compPct < 60 ? 'var(--red)' : compPct < compTarget ? 'var(--amber)' : 'var(--blue)';
  const prepCol = avgAssignToSol > 45 ? 'var(--red)' : avgAssignToSol > 30 ? 'var(--amber)' : 'var(--blue)';
  const bgOf = col => col==='var(--red)'?'rgba(239,68,68,.07)':col==='var(--amber)'?'rgba(245,158,11,.07)':'';
  const cards = [
    { lbl:'Total PRs',       val:fmt(total),           sub:'in selection',                  col:'var(--navy)' },
    { lbl:'Avg Cycle Time',  val:avgCycle+'d',          sub:\`target \${target}d\`,             col:cycCol },
    { lbl:'Net Savings',     val:fmtUSD(kpi2.net),     sub:'vs estimated value',             col:savCol },
    { lbl:'Competitive',     val:compPct+'%',           sub:fmt(kpi3.compCount)+' of '+fmt(compBase)+' classified PRs', col:cmpCol },
    { lbl:'Avg Prep Time',   val:avgAssignToSol+'d',    sub:'PR assigned → solicitation',    col:prepCol },
  ];
  document.getElementById('kpi-row').innerHTML = cards.map((c,i) => \`
    <div class="kpi-card" style="border-left-color:\${c.col};\${bgOf(c.col)?'background:'+bgOf(c.col)+';':''}">
      <div class="kpi-lbl">\${c.lbl}</div>
      <div class="kpi-val" style="color:\${c.col}">\${c.val}</div>
      <div class="kpi-sub">\${c.sub}</div>
      <div class="kpi-narrative">\${narr[i]}</div>
    </div>\`).join('');
}

// ── Team workload chart ────────────────────────────────────────────
function renderWorkload(kpi5){
  const{wlMonthly,currentAvg,numActiveBuyers,monthsOver,wlTarget}=kpi5;
  const clr=currentAvg>wlTarget?'var(--red)':currentAvg>wlTarget*0.8?'var(--amber)':'var(--green)';
  document.getElementById('workload-summary').innerHTML=\`
    <div style="display:flex;gap:20px;flex-wrap:wrap;margin-bottom:8px;align-items:flex-end;">
      <div style="text-align:center;">
        <div style="font-size:28px;font-weight:800;color:\${clr};line-height:1">\${currentAvg}</div>
        <div style="font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-top:2px">PRs/buyer (active)</div>
      </div>
      <div style="text-align:center;">
        <div style="font-size:28px;font-weight:800;color:var(--muted);line-height:1">\${wlTarget}</div>
        <div style="font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-top:2px">Standard target</div>
      </div>
      <div style="text-align:center;">
        <div style="font-size:28px;font-weight:800;color:var(--amber);line-height:1">\${monthsOver}</div>
        <div style="font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-top:2px">Months over target</div>
      </div>
      <div style="text-align:center;">
        <div style="font-size:28px;font-weight:800;color:var(--navy);line-height:1">\${numActiveBuyers}</div>
        <div style="font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-top:2px">Active buyers</div>
      </div>
    </div>\`;
  // Aggregate monthly → quarterly
  const qMap={};
  wlMonthly.forEach(m=>{
    const[y,mo]=m.month.split('-');
    const q=\`\${y} Q\${Math.ceil(+mo/3)}\`;
    if(!qMap[q])qMap[q]={total:0,buyers:0,months:0};
    qMap[q].total+=m.total; qMap[q].buyers+=m.numBuyers; qMap[q].months++;
  });
  const wlQ=Object.entries(qMap).map(([q,d])=>({q,avg:d.months?parseFloat((d.total/(d.buyers/d.months)).toFixed(1)):0,numBuyers:Math.round(d.buyers/d.months)}));
  mou('chart-workload',{type:'bar',data:{
    labels:wlQ.map(d=>d.q),
    datasets:[
      {type:'bar',label:'Avg PRs/buyer',data:wlQ.map(d=>d.avg),
       backgroundColor:wlQ.map(d=>d.avg>wlTarget?'rgba(239,68,68,.75)':'rgba(0,128,198,.65)'),borderRadius:3,order:2},
      {type:'line',label:\`Target (\${wlTarget})\`,data:wlQ.map(()=>wlTarget),
       borderColor:'#f59e0b',borderWidth:2,borderDash:[6,4],pointRadius:0,fill:false,order:1,tension:0},
    ]
  },options:{responsive:true,plugins:{legend:{labels:{color:tc()}},
    tooltip:{callbacks:{label:ctx=>ctx.dataset.type==='bar'?\` \${ctx.parsed.y} PRs/buyer (~\${wlQ[ctx.dataIndex].numBuyers} buyers)\`:\` Target: \${wlTarget}\`}}},
    scales:{x:{grid:{color:gc()},ticks:{color:tc()}},
            y:{grid:{color:gc()},ticks:{color:tc()},title:{display:true,text:'Avg PRs / buyer',color:tc()},suggestedMin:0}},
    onHover:(e,el)=>{ e.native.target.style.cursor=el.length&&el[0].datasetIndex===0?'pointer':'default'; },
    onClick:(_e,el)=>{
      if(!el.length||el[0].datasetIndex!==0) return;
      const quarter=wlQ[el[0].index].q;
      const[qy,qq]=quarter.split(' Q');
      const qMonths=['01','02','03','04','05','06','07','08','09','10','11','12']
        .slice((+qq-1)*3,(+qq-1)*3+3).map(m=>\`\${qy}-\${m}\`);
      const rows=filteredRows().filter(r=>{
        const b=r.buyer||'Unknown';
        if(wlExclC(b)) return false;
        const d=r.prReceived||'';
        return qMonths.some(m=>d.startsWith(m));
      });
      openModal(\`Workload – \${quarter}\`,drillTable(rows,['id','title','buyer','method','prReceived','stage','cycleTime']));
    },
  }});
}

// ── Trend badge helper ────────────────────────────────────────────
function setTrendBadge(id, regLine, trendUp){
  const el=document.getElementById(id); if(!el) return;
  const delta=Math.abs(regLine[regLine.length-1]-regLine[0]);
  const sign=trendUp?'+':'-';
  const color=trendUp?'#22c55e':'#ef4444';
  el.innerHTML=\`<span style="font-size:11px;font-weight:700;color:\${color};margin-left:8px;padding:2px 8px;border-radius:10px;background:\${color}18;">\${trendUp?'↑':'↓'} \${sign}\${delta}% over period</span>\`;
}

// ── KPI 3 trend ───────────────────────────────────────────────────
function renderComp3Trend(trend){
  if(!trend||!trend.length) return;
  const labels=trend.map(d=>{ const[y,mo]=d.month.split('-'); return new Date(+y,+mo-1).toLocaleString('en',{month:'short',year:'2-digit'}); });
  const compPcts=trend.map(d=>d.compPct);
  // Linear regression over competitive %
  const n=compPcts.length;
  const xs=compPcts.map((_,i)=>i);
  const mx=xs.reduce((a,b)=>a+b,0)/n;
  const my=compPcts.reduce((a,b)=>a+b,0)/n;
  const slope=xs.reduce((a,x,i)=>a+(x-mx)*(compPcts[i]-my),0)/xs.reduce((a,x)=>a+(x-mx)**2,0);
  const intercept=my-slope*mx;
  const regLine=xs.map(x=>Math.round(Math.min(100,Math.max(0,intercept+slope*x))));
  const trendUp=slope>0;
  const trendColor=trendUp?'#22c55e':'#ef4444';
  setTrendBadge('kpi3-trend-badge',regLine,trendUp);
  mou('chart-comp-trend',{type:'bar',data:{labels,datasets:[
    {type:'bar',label:'Competitive',data:compPcts,backgroundColor:'rgba(0,51,102,.75)',stack:'s',order:2,yAxisID:'y'},
    {type:'bar',label:'Direct',data:trend.map(d=>d.dirPct),backgroundColor:'rgba(158,202,225,.75)',stack:'s',order:2,yAxisID:'y'},
    {type:'line',label:'Target (70%)',data:trend.map(()=>70),borderColor:'#f59e0b',borderWidth:1.5,borderDash:[6,4],pointRadius:0,fill:false,order:0,tension:0,yAxisID:'y2'},
    {type:'line',label:\`Trend (\${trendUp?'↑ improving':'↓ declining'})\`,data:regLine,
     borderColor:trendColor,borderWidth:2.5,pointRadius:0,fill:false,order:1,tension:0,yAxisID:'y2'},
  ]},options:{responsive:true,
    plugins:{legend:{labels:{color:tc()}},datalabels:{display:false},
      tooltip:{callbacks:{label:ctx=>{
        if(ctx.dataset.yAxisID==='y2') return\` \${ctx.dataset.label}: \${ctx.parsed.y}%\`;
        const count=ctx.datasetIndex===0?trend[ctx.dataIndex].formal+trend[ctx.dataIndex].lta:trend[ctx.dataIndex].dir;
        return\` \${ctx.dataset.label}: \${ctx.parsed.y}% (\${count} PRs)\`;
      }}}},
    scales:{
      x:{stacked:true,grid:{color:gc()},ticks:{color:tc(),maxRotation:45}},
      y:{stacked:true,grid:{color:gc()},ticks:{color:tc(),callback:v=>v+'%'},min:0,max:100,title:{display:true,text:'% of PRs',color:tc()}},
      y2:{display:false,min:0,max:100},
    },
    onHover:(e,el)=>{ e.native.target.style.cursor=el.length&&el[0].datasetIndex<2?'pointer':'default'; },
    onClick:(_e,el)=>{
      if(!el.length||el[0].datasetIndex>1) return;
      const month=trend[el[0].index].month;
      const isComp=el[0].datasetIndex===0;
      const rows=filteredRows().filter(r=>(r.prReceived||r.poDate||'').startsWith(month)&&(isComp?r.isCompetitive:!r.isCompetitive));
      openModal(\`KPI 3 – \${isComp?'Competitive':'Direct'} – \${month}\`,drillTable(rows,['id','title','buyer','method','prReceived','prValue','stage']));
    },
  }});
}

// ── KPI 4 trend ───────────────────────────────────────────────────
function renderPlan4Trend(trend){
  if(!trend||!trend.length) return;
  // Exclude months where no plan data exists (all N/A)
  const tracked=trend.filter(d=>d.planned+d.unplanned>0);
  if(!tracked.length) return;
  const labels=tracked.map(d=>{ const[y,mo]=d.month.split('-'); return new Date(+y,+mo-1).toLocaleString('en',{month:'short',year:'2-digit'}); });
  // Recompute % from Planned+Unplanned only (exclude N/A from denominator)
  const pPct=tracked.map(d=>{ const t=d.planned+d.unplanned; return t?Math.round(100*d.planned/t):0; });
  const uPct=tracked.map(d=>{ const t=d.planned+d.unplanned; return t?Math.round(100*d.unplanned/t):0; });
  // Linear regression over planned %
  const n4=pPct.length;
  const xs4=pPct.map((_,i)=>i);
  const mx4=xs4.reduce((a,b)=>a+b,0)/n4;
  const my4=pPct.reduce((a,b)=>a+b,0)/n4;
  const slope4=xs4.reduce((a,x,i)=>a+(x-mx4)*(pPct[i]-my4),0)/xs4.reduce((a,x)=>a+(x-mx4)**2,0);
  const intercept4=my4-slope4*mx4;
  const regLine4=xs4.map(x=>Math.round(Math.min(100,Math.max(0,intercept4+slope4*x))));
  const trendUp4=slope4>0;
  const trendColor4=trendUp4?'#22c55e':'#ef4444';
  setTrendBadge('kpi4-trend-badge',regLine4,trendUp4);
  mou('chart-plan-trend',{type:'bar',data:{labels,datasets:[
    {type:'bar',label:'Planned',data:pPct,backgroundColor:'rgba(0,51,102,.75)',stack:'s',order:2,yAxisID:'y'},
    {type:'bar',label:'Unplanned',data:uPct,backgroundColor:'rgba(192,57,43,.75)',stack:'s',order:2,yAxisID:'y'},
    {type:'line',label:'Target (70%)',data:tracked.map(()=>70),borderColor:'#f59e0b',borderWidth:1.5,borderDash:[6,4],pointRadius:0,fill:false,order:0,tension:0,yAxisID:'y2'},
    {type:'line',label:\`Trend (\${trendUp4?'↑ improving':'↓ declining'})\`,data:regLine4,
     borderColor:trendColor4,borderWidth:2.5,pointRadius:0,fill:false,order:1,tension:0,yAxisID:'y2'},
  ]},options:{responsive:true,
    plugins:{legend:{labels:{color:tc()}},datalabels:{display:false},
      tooltip:{callbacks:{label:ctx=>{
        if(ctx.dataset.yAxisID==='y2') return\` \${ctx.dataset.label}: \${ctx.parsed.y}%\`;
        const count=ctx.datasetIndex===0?tracked[ctx.dataIndex].planned:tracked[ctx.dataIndex].unplanned;
        return\` \${ctx.dataset.label}: \${ctx.parsed.y}% (\${count} PRs)\`;
      }}}},
    scales:{
      x:{stacked:true,grid:{color:gc()},ticks:{color:tc(),maxRotation:45}},
      y:{stacked:true,grid:{color:gc()},ticks:{color:tc(),callback:v=>v+'%'},min:0,max:100,title:{display:true,text:'% of tracked PRs',color:tc()}},
      y2:{display:false,min:0,max:100},
    },
    onHover:(e,el)=>{ e.native.target.style.cursor=el.length&&el[0].datasetIndex<2?'pointer':'default'; },
    onClick:(_e,el)=>{
      if(!el.length||el[0].datasetIndex>1) return;
      const month=tracked[el[0].index].month;
      const b=el[0].datasetIndex===0?'Planned':'Unplanned';
      const rows=filteredRows().filter(r=>(r.prReceived||r.poDate||'').startsWith(month)&&r.planBucket===b);
      openModal(\`KPI 4 – \${b} – \${month}\`,drillTable(rows,['id','title','buyer','method','prReceived','prValue','stage']));
    },
  }});
}

// ── Chart helpers ─────────────────────────────────────────────────
function mou(id,cfg){
  if(CR[id]){CR[id].destroy();}
  CR[id]=new Chart(document.getElementById(id).getContext('2d'),cfg);
}

// ── KPI 1 trend state ─────────────────────────────────────────────
let lastKpi1Trend=null;
let kpi1MethodSel=new Set();
const KPI1_PALETTE=['#009FDA','#003366','#22c55e','#8b5cf6','#ec4899','#06b6d4','#84cc16','#f97316','#a16207','#0f766e'];

function renderCycleTrend(trend){
  if(!trend||!trend.months.length) return;
  const methods=[...kpi1MethodSel];
  const labels=trend.months.map(m=>{
    const[y,mo]=m.split('-');
    return new Date(+y,+mo-1).toLocaleString('en',{month:'short',year:'2-digit'});
  });
  let datasets;
  if(!methods.length){
    const pts=trend.data.map(d=>d.avg);
    datasets=[{label:'All Methods (avg)',data:pts,
      borderColor:'#009FDA',backgroundColor:'rgba(0,159,218,0.08)',
      borderWidth:2,pointRadius:5,tension:0.3,spanGaps:true,fill:true,
      pointBackgroundColor:pts.map(v=>v===null?'transparent':v<=30?'#22c55e':v<=90?'#f59e0b':'#ef4444'),
      pointBorderColor:pts.map(v=>v===null?'transparent':v<=30?'#22c55e':v<=90?'#f59e0b':'#ef4444')}];
  } else {
    datasets=methods.map((method,i)=>{
      const color=KPI1_PALETTE[i%KPI1_PALETTE.length];
      const pts=trend.data.map(d=>d.byMethod[method]??null);
      return{label:method,data:pts,borderColor:color,backgroundColor:color+'18',
        borderWidth:2,pointRadius:4,tension:0.3,spanGaps:true,fill:false,
        pointBackgroundColor:pts.map(v=>v===null?'transparent':v<=30?'#22c55e':v<=90?'#f59e0b':'#ef4444'),
        pointBorderColor:pts.map(v=>v===null?'transparent':v<=30?'#22c55e':v<=90?'#f59e0b':'#ef4444')};
    });
  }
  datasets.push({label:'Target (60d)',data:trend.months.map(()=>60),
    borderColor:'#f59e0b',borderWidth:2,borderDash:[6,4],pointRadius:0,fill:false,tension:0});
  mou('chart-cycle-trend',{type:'line',data:{labels,datasets},options:{responsive:true,
    plugins:{
      legend:{labels:{color:tc()}},
      datalabels:{display:false},
      tooltip:{callbacks:{label:ctx=>ctx.parsed.y!==null?' '+Math.round(ctx.parsed.y)+'d':' No data'}},
    },
    scales:{
      x:{grid:{color:gc()},ticks:{color:tc(),maxRotation:45}},
      y:{grid:{color:gc()},ticks:{color:tc(),callback:v=>v+'d'},
         title:{display:true,text:'Avg Cycle Time (days)',color:tc()},suggestedMin:0},
    },
    onHover:(e,el)=>{ e.native.target.style.cursor=el.length?'pointer':'default'; },
    onClick:(_e,el)=>{
      if(!el.length) return;
      const month=trend.months[el[0].index];
      const selMethods=[...kpi1MethodSel];
      let rows=filteredRows().filter(r=>(r.prReceived||'').startsWith(month)&&r.cycleTime!==null&&r.cycleTime>=0);
      if(selMethods.length) rows=rows.filter(r=>selMethods.includes(r.method));
      const lbl=selMethods.length?selMethods.join(', '):'All Methods';
      openModal(\`Cycle Time – \${lbl} – \${month}\`,drillTable(rows,['id','title','buyer','method','prReceived','poDate','cycleTime','stage']));
    },
  }});
}

// ── KPI 6 trend state ─────────────────────────────────────────────
let lastKpi6Trend=null;
let kpi6MethodSel=new Set();

// Working-day targets per method (converted to approx. calendar days for chart axis)
const PREP_TARGETS={
  'Formal Solicitation': {wd:7, cd:10, color:'#003366', dash:[8,4]},
  'Informal Solicitation':{wd:5, cd:7,  color:'#009FDA', dash:[5,3]},
  'LTA':                  {wd:3, cd:4,  color:'#22c55e', dash:[3,3]},
};

function renderPrepTrend(trend){
  if(!trend||!trend.months.length) return;
  const selMethods=[...kpi6MethodSel];
  const labels=trend.months.map(m=>{
    const[y,mo]=m.split('-');
    return new Date(+y,+mo-1).toLocaleString('en',{month:'short',year:'2-digit'});
  });

  // Determine active targets and point colour thresholds
  const activeMethods=selMethods.length?selMethods:Object.keys(PREP_TARGETS);
  // When single method selected use its target for colour coding; else use strictest (Formal)
  const primaryTarget=selMethods.length===1&&PREP_TARGETS[selMethods[0]]
    ? PREP_TARGETS[selMethods[0]]
    : null;
  const ptColor=v=>{
    if(v===null) return'transparent';
    const t=primaryTarget?primaryTarget.cd:10; // default to formal threshold when mixed
    return v<=t?'#22c55e':v<=t*1.5?'#f59e0b':'#ef4444';
  };

  let datasets;
  if(!selMethods.length){
    const pts=trend.data.map(d=>d.avg);
    datasets=[{label:'All Methods (avg)',data:pts,
      borderColor:'#0080c6',backgroundColor:'rgba(0,128,198,0.08)',
      borderWidth:2,pointRadius:5,tension:0.3,spanGaps:true,fill:true,
      pointBackgroundColor:pts.map(ptColor),pointBorderColor:pts.map(ptColor)}];
  } else {
    datasets=selMethods.map((method,i)=>{
      const color=KPI1_PALETTE[i%KPI1_PALETTE.length];
      const pts=trend.data.map(d=>d.byMethod[method]??null);
      return{label:method,data:pts,borderColor:color,backgroundColor:color+'18',
        borderWidth:2,pointRadius:4,tension:0.3,spanGaps:true,fill:false,
        pointBackgroundColor:pts.map(ptColor),pointBorderColor:pts.map(ptColor)};
    });
  }

  // Add target line(s) only when a method is selected (not for "All")
  const targetMethods=selMethods.length?selMethods:[];
  targetMethods.forEach(method=>{
    const t=PREP_TARGETS[method]; if(!t) return;
    datasets.push({label:\`Target \${method} (\${t.wd} wd)\`,data:trend.months.map(()=>t.cd),
      borderColor:t.color,borderWidth:1.5,borderDash:t.dash,pointRadius:0,fill:false,tension:0,order:0});
  });

  mou('chart-prep-trend',{type:'line',data:{labels,datasets},options:{responsive:true,
    plugins:{
      legend:{labels:{color:tc(),boxHeight:2}},
      datalabels:{display:false},
      tooltip:{callbacks:{label:ctx=>{
        if(ctx.dataset.borderDash) return\` \${ctx.dataset.label}: \${ctx.parsed.y}d\`;
        return ctx.parsed.y!==null?\` \${Math.round(ctx.parsed.y)}d\`:' No data';
      }}},
    },
    scales:{
      x:{grid:{color:gc()},ticks:{color:tc(),maxRotation:45}},
      y:{grid:{color:gc()},ticks:{color:tc(),callback:v=>v+'d'},
         title:{display:true,text:'Avg Prep Time (calendar days)',color:tc()},suggestedMin:0},
    },
    onHover:(e,el)=>{ e.native.target.style.cursor=el.length?'pointer':'default'; },
    onClick:(_e,el)=>{
      if(!el.length) return;
      const month=trend.months[el[0].index];
      let rows=filteredRows().filter(r=>{
        if(!r.dateAssigned||!r.solIssued) return false;
        const days=Math.round((new Date(r.solIssued)-new Date(r.dateAssigned))/86400000);
        return days>=0&&(r.dateAssigned||'').startsWith(month);
      });
      if(selMethods.length) rows=rows.filter(r=>selMethods.includes(r.method));
      const lbl=selMethods.length?selMethods.join(', '):'All Methods';
      openModal(\`Prep Time – \${lbl} – \${month}\`,drillTable(rows,['id','title','buyer','method','dateAssigned','solIssued','stage']));
    },
  }});
}

// ── Render charts ─────────────────────────────────────────────────
function renderCharts(K){
  const{kpi1,kpi2,kpi3,kpi4,kpi6,pipeline,kpi1Trend,kpi6Trend,kpi3Trend,kpi4Trend}=K;
  lastKpi1Trend=kpi1Trend;
  lastKpi6Trend=kpi6Trend;

  // Chart 2: Monthly savings bar + average line
  const labels=kpi2.monthly.map(m=>{
    const[y,mo]=m.month.split('-'); return new Date(+y,+mo-1).toLocaleString('en',{month:'short',year:'2-digit'});
  });
  const avgArr=kpi2.monthly.map(()=>kpi2.avgMonthly);
  mou('chart-savings',{type:'bar',data:{
    labels,
    datasets:[
      {type:'bar',label:'Monthly Savings (USD)',data:kpi2.monthly.map(m=>m.total),
       backgroundColor:kpi2.monthly.map(m=>m.total>=0?'rgba(0,128,198,.7)':'rgba(192,57,43,.7)'),borderRadius:3,order:2},
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

  // Chart 3: Competitive vs Direct vs Other — 3-slice doughnut
  const c3Labels=['Competitive','Direct Procurement','Other'];
  const c3Tot=kpi3.compCount+kpi3.dirCount+kpi3.otherCount;
  mou('chart-comp',{type:'doughnut',data:{
    labels:c3Labels,
    datasets:[{data:[kpi3.compCount,kpi3.dirCount,kpi3.otherCount],
      backgroundColor:['#003366','#009FDA','#9ecae1'],hoverOffset:6}]
  },options:{responsive:true,plugins:{legend:{position:'bottom',labels:{color:tc()}},
    tooltip:{callbacks:{label:ctx=>{
      const vals=[kpi3.compValue,kpi3.dirValue,kpi3.otherValue];
      const pct=c3Tot?Math.round(100*ctx.parsed/c3Tot):0;
      return \` \${ctx.parsed} PRs (\${pct}%) – \${fmtUSD(vals[ctx.dataIndex])}\`;
    }}},
    datalabels:{
      display:true,
      color:'#fff',
      font:{weight:'bold',size:12},
      textAlign:'center',
      formatter:(value,ctx)=>{
        const pct=c3Tot?Math.round(100*value/c3Tot):0;
        return pct>=5?\`\${c3Labels[ctx.dataIndex]}\\n\${pct}%\`:'';
      }
    }},
    onClick:(_e,el)=>{ if(!el.length)return;
      const cats=[r=>r.isCompetitive, r=>r.isDirect&&!r.isCompetitive, r=>!r.isCompetitive&&!r.isDirect];
      const lbls=['Competitive PRs','Direct Procurement PRs','Other PRs'];
      const i=el[0].index;
      openModal(lbls[i],drillTable(filteredRows().filter(cats[i]),['id','title','buyer','method','prValue','stage']));
    }
  }});

  // Chart 4: Plan compliance pie — exclude N/A (pre-dates plan tracking)
  const pLabels=['Planned','Unplanned'];
  const pTot=pLabels.reduce((a,l)=>a+kpi4.planCounts[l],0);
  mou('chart-plan',{type:'pie',data:{
    labels:pLabels,
    datasets:[{data:pLabels.map(l=>kpi4.planCounts[l]),backgroundColor:['#003366','#c0392b'],hoverOffset:6}]
  },options:{responsive:true,plugins:{
    legend:{position:'bottom',labels:{color:tc()}},
    tooltip:{callbacks:{label:ctx=>{
      const pct=pTot?Math.round(100*ctx.parsed/pTot):0;
      return \` \${ctx.parsed} (\${pct}%) \u2013 \${fmtUSD(kpi4.planValues[pLabels[ctx.dataIndex]])}\`;
    }}},
    datalabels:{
      display:true,
      color:'#fff',
      font:{weight:'bold',size:12},
      textAlign:'center',
      formatter:(value,ctx)=>{
        const pct=pTot?Math.round(100*value/pTot):0;
        return pct>=5?\`\${pLabels[ctx.dataIndex]}\\n\${pct}%\`:'';
      }
    }},
    onClick:(_e,el)=>{ if(!el.length)return; const b=pLabels[el[0].index];
      openModal(\`Plan \u2013 \${b}\`,drillTable(filteredRows().filter(r=>r.planBucket===b),['id','title','buyer','method','prValue','stage'])); }
  }});

  renderCycleTrend(kpi1Trend);
  renderPrepTrend(kpi6Trend);
  renderComp3Trend(kpi3Trend);
  renderPlan4Trend(kpi4Trend);
}

// ── Drill table ────────────────────────────────────────────────────
const CL={id:'ID',title:'Description',buyer:'Buyer',method:'Method',prValue:'PR Value',
  cumulativePO:'Cumul. PO',savings:'Savings',stage:'Stage',prReceived:'PR Received',
  poDate:'PO Date',cycleTime:'Cycle (d)',isCompetitive:'Competitive?',planBucket:'Plan',
  dateAssigned:'PR Assigned',solIssued:'Sol. Issued'};
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
// ── Management Report Generator ───────────────────────────────────
function generateReport(){
  const rows=filteredRows();
  const K=recompute(rows);
  const avgCycle=K.avgCycle,kpi2=K.kpi2,kpi3=K.kpi3,kpi4=K.kpi4,kpi5=K.kpi5,avgAssignToSol=K.avgAssignToSol,kpi3Trend=K.kpi3Trend,kpi4Trend=K.kpi4Trend;
  const now=new Date();
  const reportDate=now.toLocaleDateString('en-GB',{day:'numeric',month:'long',year:'numeric'});
  const reportMonth=now.toLocaleDateString('en-GB',{month:'long',year:'numeric'});

  // ── Derived values ────────────────────────────────────────────────
  const compBase=kpi3.compCount+kpi3.dirCount;
  const compPct=compBase?Math.round(100*kpi3.compCount/compBase):0;
  const planBase=kpi4.planCounts.Planned+kpi4.planCounts.Unplanned;
  const planPct=planBase?Math.round(100*kpi4.planCounts.Planned/planBase):0;
  const cycleTgt=60, wlTgt=15, compTgt=70, planTgt=70;

  // Trend directions
  const k3last=kpi3Trend&&kpi3Trend.length>1?kpi3Trend[kpi3Trend.length-1].compPct:null;
  const k3first=kpi3Trend&&kpi3Trend.length>1?kpi3Trend[0].compPct:null;
  const k3improving=k3last!==null&&k3last>k3first;
  const k4last=kpi4Trend&&kpi4Trend.length>1?kpi4Trend.filter(d=>d.planned+d.unplanned>0).slice(-1)[0]:null;
  const k4first=kpi4Trend&&kpi4Trend.length>1?kpi4Trend.filter(d=>d.planned+d.unplanned>0)[0]:null;
  const k4lastPct=k4last?Math.round(100*k4last.planned/(k4last.planned+k4last.unplanned)):null;
  const k4firstPct=k4first?Math.round(100*k4first.planned/(k4first.planned+k4first.unplanned)):null;
  const k4improving=k4lastPct!==null&&k4firstPct!==null&&k4lastPct>k4firstPct;

  // RAG status helpers
  const cycleRag=avgCycle<=cycleTgt?'green':avgCycle<=cycleTgt*1.25?'amber':'red';
  const savRag=kpi2.net>0?'green':kpi2.net===0?'amber':'red';
  const compRag=compPct>=compTgt?'green':compPct>=compTgt-10?'amber':'red';
  const planRag=planPct>=planTgt?'green':planPct>=planTgt-15?'amber':'red';
  const wlRag=kpi5.currentAvg<=wlTgt?'green':kpi5.currentAvg<=wlTgt*1.2?'amber':'red';
  const prepRag=avgAssignToSol<=10?'green':avgAssignToSol<=15?'amber':'red';
  const ragColor={green:'#16a34a',amber:'#d97706',red:'#dc2626'};
  const ragBg={green:'#f0fdf4',amber:'#fffbeb',red:'#fef2f2'};
  const ragLabel={green:'On Target',amber:'Monitor',red:'Action Required'};

  // Active filter context
  const filterCtx=[
    AF.year?\`Year: \${AF.year}\`:'',
    AF.buyer?\`Buyer: \${AF.buyer}\`:'',
    AF.methods.size?\`Methods: \${[...AF.methods].join(', ')}\`:'',
    AF.projects.size?\`Projects: \${[...AF.projects].join(', ')}\`:'',
  ].filter(Boolean);
  const scopeNote=filterCtx.length?\`Filtered view — \${filterCtx.join(' | ')}\`:'Full portfolio — all periods and buyers';

  // ── Narrative paragraphs ──────────────────────────────────────────
  const p1=(()=>{
    const cycleMsg=avgCycle<=cycleTgt
      ?\`The average procurement cycle time stands at <strong>\${avgCycle} days</strong>, \${cycleTgt-avgCycle} days below the \${cycleTgt}-day target — indicating a well-paced procurement process.\`
      :\`The average procurement cycle time stands at <strong>\${avgCycle} days</strong>, \${avgCycle-cycleTgt} days above the \${cycleTgt}-day target — warranting attention to pipeline efficiency.\`;
    const savMsg=kpi2.net>0
      ?\`Procurement savings for the period total <strong>\${fmtUSD(kpi2.net)}</strong> against estimated values, reflecting effective market engagement and negotiation by the team.\`
      :\`The net savings position for the period stands at <strong>\${fmtUSD(kpi2.net)}</strong>, suggesting that final contracted values exceeded initial estimates in some cases.\`;
    return \`\${cycleMsg} \${savMsg}\`;
  })();

  const p2=(()=>{
    const compMsg=compPct>=compTgt
      ?\`Competitive procurement compliance is strong at <strong>\${compPct}%</strong> (target: \${compTgt}%), \${k3improving?'with an improving trend over the reporting period':'though the trend over the period shows a gradual decline that should be monitored'}.\`
      :\`Competitive procurement compliance stands at <strong>\${compPct}%</strong>, below the \${compTgt}% target. \${k3improving?'The trend is improving, suggesting recent corrective actions are taking effect.':'The declining trend requires management attention and a review of procurement modalities.'}\`;
    const planMsg=planBase>0?(planPct>=planTgt
      ?\`Procurement plan adherence is strong at <strong>\${planPct}%</strong> of tracked PRs (target: \${planTgt}%), \${k4improving?'with a positive trajectory':'though the trend merits continued attention'}.\`
      :\`Procurement plan adherence stands at <strong>\${planPct}%</strong> (target: \${planTgt}%), indicating that a significant share of procurement activity remains reactive. \${k4improving?'Encouragingly, the trend is improving.':'Strengthening the annual planning process is recommended.'}\`)
      :'Plan compliance tracking was recently introduced; data will become more meaningful over time.';
    return \`\${compMsg} \${planMsg}\`;
  })();

  const p3=(()=>{
    const wlMsg=kpi5.currentAvg<=wlTgt
      ?\`Team workload is within the standard target, with an average of <strong>\${kpi5.currentAvg} active PRs per buyer</strong> against the \${wlTgt}-PR benchmark across \${kpi5.numActiveBuyers} active procurement officers.\`
      :\`Team workload is above the standard target at <strong>\${kpi5.currentAvg} PRs per buyer</strong> (target: \${wlTgt}), with \${kpi5.numActiveBuyers} active procurement officers currently managing \${kpi5.totalActive} open PRs. This level may affect processing quality and timeliness.\`;
    const prepMsg=avgAssignToSol>0
      ?\`The average solicitation preparation time is <strong>\${avgAssignToSol} days</strong> from PR assignment to solicitation issuance.\`
      :'';
    return \`\${wlMsg} \${prepMsg?prepMsg+' ':''}\`.trim();
  })();

  // ── Key observations ──────────────────────────────────────────────
  const obs=[];
  if(cycleRag==='green') obs.push({icon:'✓',col:'#16a34a',text:\`Cycle time (\${avgCycle}d) is \${cycleTgt-avgCycle}d below the \${cycleTgt}-day target.\`});
  if(cycleRag==='amber') obs.push({icon:'⚠',col:'#d97706',text:\`Cycle time (\${avgCycle}d) slightly exceeds target — review highest-delay methods.\`});
  if(cycleRag==='red')   obs.push({icon:'✗',col:'#dc2626',text:\`Cycle time (\${avgCycle}d) significantly above target (\${cycleTgt}d) — immediate process review recommended.\`});
  if(savRag==='green')   obs.push({icon:'✓',col:'#16a34a',text:\`Net savings of \${fmtUSD(kpi2.net)} demonstrate strong value-for-money.\`});
  if(savRag==='red')     obs.push({icon:'✗',col:'#dc2626',text:\`Negative net savings (\${fmtUSD(kpi2.net)}) — review estimate accuracy and negotiation outcomes.\`});
  if(compRag==='green')  obs.push({icon:'✓',col:'#16a34a',text:\`Competitive procurement at \${compPct}% meets the \${compTgt}% target.\${k3improving?' Trend is improving.':''}\`});
  if(compRag!=='green')  obs.push({icon:compRag==='red'?'✗':'⚠',col:ragColor[compRag],text:\`Competitive procurement at \${compPct}% is below the \${compTgt}% target.\${k3improving?' Trend improving.':' Trend declining — review modality selection.'}\`});
  if(planBase>0){
    if(planRag==='green') obs.push({icon:'✓',col:'#16a34a',text:\`Plan compliance at \${planPct}% meets the \${planTgt}% target.\${k4improving?' Upward trend confirmed.':''}\`});
    if(planRag!=='green') obs.push({icon:planRag==='red'?'✗':'⚠',col:ragColor[planRag],text:\`Plan compliance at \${planPct}% is below \${planTgt}% — strengthen annual procurement planning.\`});
  }
  if(wlRag==='red')      obs.push({icon:'✗',col:'#dc2626',text:\`Team workload (\${kpi5.currentAvg} PRs/buyer) exceeds the \${wlTgt}-PR standard — consider staffing review.\`});
  if(kpi5.monthsOver>0)  obs.push({icon:'⚠',col:'#d97706',text:\`Workload exceeded target in \${kpi5.monthsOver} months over the reporting period.\`});

  // ── Recommended actions (only when issues exist) ──────────────────
  const actions=[];
  if(cycleRag!=='green') actions.push('Conduct a root-cause analysis of the highest-delay solicitation methods and introduce process milestones to track progress.');
  if(compRag!=='green'&&!k3improving) actions.push('Review procurement modality selection criteria. Ensure LTA utilisation is maximised before resorting to direct procurement.');
  if(planRag!=='green') actions.push('Strengthen the annual procurement planning cycle: engage programme teams earlier in the planning process to increase PR predictability.');
  if(wlRag!=='green') actions.push('Assess current staffing levels against pipeline volume. Consider temporary support or workload redistribution across the procurement team.');
  if(savRag==='red') actions.push('Review the quality of budget estimates at PR submission stage. Engage requesters on realistic market pricing to improve savings accuracy.');

  // ── KPI card builder ──────────────────────────────────────────────
  const kpiCard=(num,title,value,target,rag,finding)=>\`
    <div style="border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;">
      <div style="background:#003366;padding:10px 14px;display:flex;align-items:center;gap:10px;">
        <div style="background:rgba(255,255,255,.15);color:#fff;font-size:12px;font-weight:800;border-radius:6px;width:24px;height:24px;display:flex;align-items:center;justify-content:center;flex-shrink:0;">\${num}</div>
        <div style="color:#fff;font-size:11.5px;font-weight:700;">\${title}</div>
        <div style="margin-left:auto;background:\${ragBg[rag]};color:\${ragColor[rag]};font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;white-space:nowrap;">\${ragLabel[rag]}</div>
      </div>
      <div style="padding:12px 14px;">
        <div style="display:flex;align-items:baseline;gap:8px;margin-bottom:6px;">
          <div style="font-size:26px;font-weight:800;color:\${ragColor[rag]};line-height:1;">\${value}</div>
          <div style="font-size:11px;color:#64748b;">target: \${target}</div>
        </div>
        <div style="font-size:11.5px;color:#475569;line-height:1.5;border-top:1px solid #f1f5f9;padding-top:8px;">\${finding}</div>
      </div>
    </div>\`;

  // ── Build HTML ────────────────────────────────────────────────────
  const html=\`<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>Procurement Performance Report – \${reportMonth}</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0;}
    body{font-family:'Segoe UI',Arial,sans-serif;background:#f8fafc;color:#1a2e44;font-size:13px;}
    .page{max-width:900px;margin:0 auto;padding:40px 48px;}
    h2{font-size:15px;font-weight:700;color:#003366;margin-bottom:14px;padding-bottom:8px;border-bottom:2px solid #003366;}
    p{font-size:13px;line-height:1.75;color:#334155;margin-bottom:12px;}
    .kpi-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:32px;}
    .obs-item{display:flex;gap:10px;margin-bottom:10px;align-items:flex-start;}
    .obs-icon{font-size:14px;font-weight:700;width:20px;flex-shrink:0;margin-top:1px;}
    .action-item{display:flex;gap:10px;margin-bottom:10px;}
    .action-num{background:#003366;color:#fff;font-size:10px;font-weight:700;border-radius:50%;width:18px;height:18px;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:1px;}
    .tag{display:inline-block;font-size:10px;font-weight:600;padding:2px 8px;border-radius:10px;margin-right:4px;}
    @media print{
      body{background:#fff;}
      .no-print{display:none!important;}
      .page{padding:20px 28px;}
    }
  </style>
</head>
<body>
<div class="page">

  <!-- Print button -->
  <div class="no-print" style="text-align:right;margin-bottom:20px;">
    <button onclick="window.print()" style="background:#003366;color:#fff;border:none;padding:8px 20px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;">🖨️ Save as PDF / Print</button>
  </div>

  <!-- Letterhead -->
  <div style="background:#003366;border-radius:12px;padding:28px 32px;margin-bottom:32px;color:#fff;">
    <div style="font-size:11px;letter-spacing:1.5px;opacity:.7;text-transform:uppercase;margin-bottom:4px;">Food and Agriculture Organization of the United Nations</div>
    <div style="font-size:22px;font-weight:800;letter-spacing:-.3px;margin-bottom:2px;">Ukraine Procurement Unit</div>
    <div style="font-size:15px;font-weight:400;opacity:.85;margin-bottom:20px;">Procurement Performance Report</div>
    <div style="display:flex;gap:24px;flex-wrap:wrap;">
      <div><div style="font-size:10px;opacity:.6;text-transform:uppercase;letter-spacing:.8px;">Date of Issue</div><div style="font-size:13px;font-weight:600;margin-top:3px;">\${reportDate}</div></div>
      <div><div style="font-size:10px;opacity:.6;text-transform:uppercase;letter-spacing:.8px;">Portfolio Scope</div><div style="font-size:13px;font-weight:600;margin-top:3px;">\${rows.length} Purchase Requests</div></div>
      <div><div style="font-size:10px;opacity:.6;text-transform:uppercase;letter-spacing:.8px;">Data as of</div><div style="font-size:13px;font-weight:600;margin-top:3px;">\${DASHBOARD_DATA.lastUpdated}</div></div>
      \${filterCtx.length?\`<div style="border-left:1px solid rgba(255,255,255,.2);padding-left:20px;"><div style="font-size:10px;opacity:.6;text-transform:uppercase;letter-spacing:.8px;">Active Filters</div><div style="font-size:12px;margin-top:3px;opacity:.9;">\${filterCtx.join(' · ')}</div></div>\`:''}
    </div>
  </div>

  <!-- Executive Summary -->
  <h2>Executive Summary</h2>
  <p>\${p1}</p>
  <p>\${p2}</p>
  <p>\${p3}</p>

  <!-- KPI Overview -->
  <h2 style="margin-top:28px;">Key Performance Indicators</h2>
  <div class="kpi-grid">
    \${kpiCard('1','Procurement Cycle Time',avgCycle+'d',cycleTgt+'d',cycleRag,
      avgCycle<=cycleTgt
        ?\`Cycle time is \${cycleTgt-avgCycle} days below target. \${K.kpi1.length?\`Fastest method: \${[...K.kpi1].sort((a,b)=>a.avg-b.avg)[0].method} (\${[...K.kpi1].sort((a,b)=>a.avg-b.avg)[0].avg}d).\`:''}\`
        :\`Cycle time exceeds target by \${avgCycle-cycleTgt} days. \${K.kpi1.length?\`Slowest method: \${K.kpi1[0].method} (\${K.kpi1[0].avg}d).\`:''}\`
    )}
    \${kpiCard('2','Procurement Savings',fmtUSD(kpi2.net),'Positive',savRag,
      \`Positive: \${fmtUSD(kpi2.sumPos)} · Negative: \${fmtUSD(Math.abs(kpi2.sumNeg))} · Avg monthly: \${fmtUSD(kpi2.avgMonthly)}.\`
    )}
    \${kpiCard('3','Competitive Procurement',compPct+'%',compTgt+'%',compRag,
      \`\${kpi3.compCount} competitive PRs out of \${compBase} classified. LTA: \${kpi3.otherCount} PRs. Trend: \${k3improving?'↑ improving':'↓ declining'}.\`
    )}
    \${kpiCard('4','Plan Compliance',planBase?planPct+'%':'N/A',planTgt+'%',planBase?planRag:'amber',
      planBase
        ?\`\${kpi4.planCounts.Planned} planned, \${kpi4.planCounts.Unplanned} unplanned out of \${planBase} tracked PRs. Trend: \${k4improving?'↑ improving':'↓ declining'}.\`
        :'Insufficient tracked data for this period.'
    )}
    \${kpiCard('5','Team Workload',kpi5.currentAvg+' PRs/buyer',wlTgt+' PRs/buyer',wlRag,
      \`\${kpi5.numActiveBuyers} active buyers managing \${kpi5.totalActive} open PRs. Target exceeded in \${kpi5.monthsOver} month(s).\`
    )}
    \${kpiCard('6','Solicitation Prep Time',avgAssignToSol?avgAssignToSol+'d':'—','7–10 wd',prepRag,
      avgAssignToSol?\`Average \${avgAssignToSol} calendar days from PR assignment to solicitation launch.\`:'No data available for this period.'
    )}
  </div>

  <!-- Key Observations -->
  <h2>Key Observations</h2>
  <div style="margin-bottom:28px;">
    \${obs.map(o=>\`<div class="obs-item"><div class="obs-icon" style="color:\${o.col}">\${o.icon}</div><div style="font-size:12.5px;color:#334155;line-height:1.6;">\${o.text}</div></div>\`).join('')}
  </div>

  \${actions.length?\`
  <!-- Recommended Actions -->
  <h2>Recommended Actions</h2>
  <div style="margin-bottom:28px;">
    \${actions.map((a,i)=>\`<div class="action-item"><div class="action-num">\${i+1}</div><div style="font-size:12.5px;color:#334155;line-height:1.6;">\${a}</div></div>\`).join('')}
  </div>\`:''}

  <!-- Footer -->
  <div style="border-top:1px solid #e2e8f0;padding-top:16px;margin-top:8px;display:flex;justify-content:space-between;align-items:center;">
    <div style="font-size:10.5px;color:#94a3b8;">FAO Ukraine Procurement Unit · Generated \${reportDate} · Data source: SharePoint Procurement Tracking List</div>
    <div style="font-size:10.5px;color:#94a3b8;font-style:italic;">INTERNAL USE</div>
  </div>

</div>
</body>
</html>\`;

  const w=window.open('','_blank');
  w.document.write(html);
  w.document.close();
}

function openAbout(){ const el=document.getElementById('about-ov'); el.style.display='flex'; }
function closeAbout(){ document.getElementById('about-ov').style.display='none'; }
document.getElementById('about-ov').addEventListener('click',e=>{ if(e.target===document.getElementById('about-ov'))closeAbout(); });
document.getElementById('modal-ov').addEventListener('click',e=>{ if(e.target===document.getElementById('modal-ov'))closeModal(); });
document.addEventListener('keydown',e=>{ if(e.key==='Escape'){ closeModal(); closeAbout(); } });

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

// ── Multi-check dropdown ───────────────────────────────────────────
const MCD_STATE = {}; // id to Set of selected values

function initMcd(wrapId, items, filterKey, labelAll) {
  MCD_STATE[wrapId] = new Set();
  const list = document.getElementById(wrapId + '-list');
  items.forEach(item => {
    const div = document.createElement('div');
    div.className = 'mcd-item';
    div.dataset.val = item;
    div.innerHTML = \`<input type="checkbox" data-val="\${item.replace(/"/g,'&quot;')}"/><span>\${item}</span>\`;
    div.querySelector('input').addEventListener('change', function() {
      if (this.checked) MCD_STATE[wrapId].add(item);
      else MCD_STATE[wrapId].delete(item);
      AF[filterKey] = MCD_STATE[wrapId];
      updateMcdBtn(wrapId, labelAll);
      refresh();
    });
    list.appendChild(div);
  });
  // Close on outside click
  document.addEventListener('click', e => {
    const wrap = document.getElementById(wrapId);
    if (!wrap.contains(e.target)) wrap.classList.remove('open');
  });
}

function toggleMcd(wrapId) {
  const wrap = document.getElementById(wrapId);
  // Close all others
  document.querySelectorAll('.mcd-wrap.open').forEach(w => { if (w.id !== wrapId) w.classList.remove('open'); });
  wrap.classList.toggle('open');
}

function filterMcdItems(wrapId, query) {
  const q = query.toLowerCase();
  document.querySelectorAll(\`#\${wrapId}-list .mcd-item\`).forEach(item => {
    item.style.display = item.dataset.val.toLowerCase().includes(q) ? '' : 'none';
  });
}

function clearMcd(wrapId) {
  MCD_STATE[wrapId].clear();
  document.querySelectorAll(\`#\${wrapId}-list input[type=checkbox]\`).forEach(cb => cb.checked = false);
  const wrapEl = document.getElementById(wrapId);
  const filterKey = wrapEl.dataset.filterkey;
  AF[filterKey] = MCD_STATE[wrapId];
  updateMcdBtn(wrapId, wrapEl.dataset.labelall);
  refresh();
}

function updateMcdBtn(wrapId, labelAll) {
  const wrap = document.getElementById(wrapId);
  const btn = wrap.querySelector('.mcd-btn');
  const sel = MCD_STATE[wrapId];
  if (sel.size === 0) {
    btn.textContent = '';
    btn.innerHTML = labelAll + ' <span>\u25be</span>';
    btn.classList.remove('has-selection');
  } else {
    btn.textContent = '';
    btn.innerHTML = sel.size + ' selected <span>\u25be</span>';
    btn.classList.add('has-selection');
  }
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
  renderWorkload(K.kpi5);
}

// ── Init ───────────────────────────────────────────────────────────
document.getElementById('last-updated').textContent = DASHBOARD_DATA.lastUpdated;
makePills('pills-year',  DASHBOARD_DATA.years,  'year');
makePills('pills-buyer', DASHBOARD_DATA.buyers, 'buyer');
initMcd('mcd-method',  DASHBOARD_DATA.methods,  'methods',  'All Methods');
initMcd('mcd-project', DASHBOARD_DATA.projects, 'projects', 'All Projects');
// KPI 6 method filter – multi-select (local to trend chart only)
(function(){
  const wrap=document.getElementById('pills-kpi6method');
  const all=document.createElement('button');
  all.className='pill on'; all.textContent='All';
  all.onclick=()=>{
    kpi6MethodSel.clear();
    wrap.querySelectorAll('.pill').forEach(p=>p.classList.remove('on'));
    all.classList.add('on');
    if(lastKpi6Trend) renderPrepTrend(lastKpi6Trend);
  };
  wrap.appendChild(all);
  DASHBOARD_DATA.methods.forEach(method=>{
    const p=document.createElement('button');
    p.className='pill'; p.textContent=method;
    p.onclick=()=>{
      const wasOn=kpi6MethodSel.has(method);
      kpi6MethodSel.clear();
      wrap.querySelectorAll('.pill').forEach(q=>q.classList.remove('on'));
      if(!wasOn){ kpi6MethodSel.add(method); p.classList.add('on'); }
      else all.classList.add('on');
      if(lastKpi6Trend) renderPrepTrend(lastKpi6Trend);
    };
    wrap.appendChild(p);
  });
})();
// KPI 1 method filter – multi-select (local to trend chart only)
(function(){
  const wrap=document.getElementById('pills-kpi1method');
  const all=document.createElement('button');
  all.className='pill on'; all.textContent='All';
  all.onclick=()=>{
    kpi1MethodSel.clear();
    wrap.querySelectorAll('.pill').forEach(p=>p.classList.remove('on'));
    all.classList.add('on');
    if(lastKpi1Trend) renderCycleTrend(lastKpi1Trend);
  };
  wrap.appendChild(all);
  DASHBOARD_DATA.methods.forEach(method=>{
    const p=document.createElement('button');
    p.className='pill'; p.textContent=method;
    p.onclick=()=>{
      const wasOn=kpi1MethodSel.has(method);
      if(wasOn) kpi1MethodSel.delete(method);
      else kpi1MethodSel.add(method);
      p.classList.toggle('on',!wasOn);
      if(kpi1MethodSel.size===0) all.classList.add('on');
      else all.classList.remove('on');
      if(lastKpi1Trend) renderCycleTrend(lastKpi1Trend);
    };
    wrap.appendChild(p);
  });
})();
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
