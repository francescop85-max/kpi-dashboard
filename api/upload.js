import { put } from '@vercel/blob';

// ── Parsing helpers (mirrors build.js) ───────────────────────────────────────

const EXCLUDED_METHODS = new Set([
  'Call for Application (704)', 'EOI (502)', 'EOI (507)',
  'Invitation for Proposal (507)', 'RFI (502)', 'Contract Amendment', 'Other',
]);

function consolidateMethod(raw) {
  if (!raw || EXCLUDED_METHODS.has(raw)) return null;
  if (/LTA|UN Award/i.test(raw))            return 'LTA';
  if (/\b(ITB|RFP|BAFO|RFQ)\b/i.test(raw)) return 'Formal Solicitation';
  if (/micro purchase|direct|very low value|re-utilisation/i.test(raw)) return 'Informal Solicitation';
  return null;
}

function parseLookup(raw) {
  if (!raw) return '';
  try {
    const v = raw.trim();
    if (v.startsWith('{')) { const o = JSON.parse(v); return o.Value ?? ''; }
    if (v.startsWith('[')) { const a = JSON.parse(v); return a.map(x => x.Value ?? '').filter(Boolean).join('; '); }
  } catch (_) {}
  return raw;
}

function parseMultiLookup(raw) {
  if (!raw) return '';
  try {
    const v = raw.trim();
    if (v.startsWith('[')) { const a = JSON.parse(v); return a.map(x => x.Value ?? '').filter(Boolean).join('; '); }
    if (v.startsWith('{')) { const o = JSON.parse(v); return o.Value ?? ''; }
  } catch (_) {}
  return raw;
}

function parseUser(raw) {
  if (!raw) return '';
  try {
    const v = raw.trim();
    if (v.startsWith('{')) {
      const o = JSON.parse(v);
      let name = (o.DisplayName ?? '').replace(/\s*\([^)]+\)\s*$/, '').trim();
      if (name.includes(',')) name = name.split(',')[0].trim();
      return name;
    }
  } catch (_) {}
  return raw;
}

function parseFloat2(raw) {
  if (!raw) return null;
  const n = parseFloat(String(raw).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

function parseDate(raw) {
  if (!raw) return null;
  const m = String(raw).match(/^(\d{4}-\d{2}-\d{2})/);
  if (!m) return null;
  const d = new Date(m[1] + 'T00:00:00Z');
  return isNaN(d.getTime()) ? null : d;
}

function fmtDate(d) {
  if (!d) return '';
  return d.toISOString().slice(0, 10);
}

function parseCSV(text) {
  const rows = [];
  let col = 0, inQuote = false, field = '', row = [];
  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  for (let i = 0; i < text.length; i++) {
    const ch = text[i], next = text[i + 1];
    if (inQuote) {
      if (ch === '"' && next === '"') { field += '"'; i++; }
      else if (ch === '"') inQuote = false;
      else field += ch;
    } else {
      if (ch === '"') inQuote = true;
      else if (ch === ',') { row.push(field); field = ''; col++; }
      else if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; col = 0; }
      else field += ch;
    }
  }
  if (field || row.length) { row.push(field); if (row.some(c => c !== '')) rows.push(row); }
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, idx) => { obj[h.trim()] = (r[idx] ?? '').trim(); });
    return obj;
  });
}

function processCSV(csvText) {
  const rawRows = parseCSV(csvText);
  const rows = rawRows.map(r => {
    const buyer        = parseUser(r['Buyer']);
    const prValue      = parseFloat2(r['PRValue']);
    const cumulativePO = parseFloat2(r['CumulativePO_x0024_']);
    const savings      = parseFloat2(r['Savings']);
    const rawMethod    = parseLookup(r['SollicitationMethod']);
    const method       = consolidateMethod(rawMethod);
    const stage        = parseLookup(r['ProcurementStage']);
    const planRaw      = parseLookup(r['PartofProcurementPlan']);
    const awardBasis   = parseLookup(r['AwardBasis']);
    const marketCat    = parseMultiLookup(r['MarketCategory']);
    const projRef      = parseMultiLookup(r['ProjectReference']);
    if (method === null) return null;
    const prReceived      = parseDate(r['PRReceived']);
    const poDate          = parseDate(r['POIssuancedate']);
    const solIssued       = parseDate(r['SollicitationIssued']);
    const solClosed       = parseDate(r['SolicitationClosed']);
    const created         = parseDate(r['Created']);
    const modified        = parseDate(r['Modified']);
    const dateAssigned    = parseDate(r['PRAssigned']);
    const dateClosed      = parseDate(r['DateClosed']);
    const techOfferShared = parseDate(r['TechnicalOfferShared']);
    const tcoClearance    = parseDate(r['TCOClearance']);
    const ltoClearance    = parseDate(r['LTOClearance']);
    const awardRec        = parseDate(r['AwardRecommandation']);
    const cycleTime = prReceived && poDate ? Math.round((poDate - prReceived) / 86400000) : null;
    const year = prReceived ? prReceived.getUTCFullYear() : (poDate ? poDate.getUTCFullYear() : null);
    const awardBasisCompetitive = /competitive/i.test(awardBasis);
    const isCompetitive = method === 'Formal Solicitation' || method === 'LTA' || awardBasisCompetitive;
    const isDirect = !isCompetitive;
    let planBucket;
    if (/yes.*part.*procurement.*plan/i.test(planRaw))  planBucket = 'Planned';
    else if (/not planned/i.test(planRaw))               planBucket = 'Unplanned';
    else                                                  planBucket = 'N/A';
    return {
      id: r['ID'] || r['ItemInternalId'] || '',
      title: r['Title'] || r['Description'] || '',
      prgrms: r['PRGRMS_x0023_'] || '',
      buyer, stage, status: (r['Status'] || '').trim(),
      prValue, cumulativePO, savings, method, marketCat, projRef,
      awardBasis, planRaw, planBucket, isCompetitive, isDirect,
      prReceived:   fmtDate(prReceived),   poDate:       fmtDate(poDate),
      solIssued:    fmtDate(solIssued),    solClosed:    fmtDate(solClosed),
      created:      fmtDate(created),      modified:     fmtDate(modified),
      dateAssigned: fmtDate(dateAssigned), dateClosed:   fmtDate(dateClosed),
      techOfferShared: fmtDate(techOfferShared), tcoClearance: fmtDate(tcoClearance),
      ltoClearance: fmtDate(ltoClearance), awardRec: fmtDate(awardRec),
      cycleTime, year,
      poNumber:   r['PO_x0023_'] || '',
      buyingUnit: parseLookup(r['BuyingUnit']),
    };
  }).filter(r => r && !/cancelled/i.test(r.stage || ''));

  const years   = [...new Set(rows.map(r => r.year).filter(Boolean))].sort();
  const BUYER_EXCL = /perini|horvath|weng/i;
  const buyers  = [...new Set(rows.map(r => r.buyer).filter(b => b && !BUYER_EXCL.test(b)))].sort();
  const methods = [...new Set(rows.map(r => r.method).filter(Boolean))].sort();
  const projectSet = new Set();
  rows.forEach(r => {
    if (r.projRef) r.projRef.split(';').map(p => p.trim()).filter(Boolean).forEach(p => projectSet.add(p));
  });
  const projects = [...projectSet].sort();
  const modDates = rows.map(r => r.modified).filter(Boolean).sort();
  const lastUpdated = modDates.length ? modDates[modDates.length - 1] : new Date().toISOString().slice(0, 10);

  return { rows, years, buyers, methods, projects, lastUpdated };
}

// ── Handler ───────────────────────────────────────────────────────────────────

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-upload-secret');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ error: 'Method not allowed' });
  }

  // Auth
  const secret = process.env.UPLOAD_SECRET;
  if (secret) {
    const provided = req.headers['x-upload-secret'] || '';
    if (provided !== secret) return res.status(401).json({ error: 'Unauthorized' });
  }

  // Read body
  const chunks = [];
  for await (const chunk of req) chunks.push(Buffer.from(chunk));
  const csvText = Buffer.concat(chunks).toString('utf8');
  if (!csvText.trim()) return res.status(400).json({ error: 'Empty body' });

  try {
    const data = processCSV(csvText);
    await put('dashboard-data.json', JSON.stringify(data), {
      access: 'public',
      contentType: 'application/json',
      addRandomSuffix: false,
      allowOverwrite: true,
    });
    return res.status(200).json({ ok: true, rows: data.rows.length, lastUpdated: data.lastUpdated });
  } catch (err) {
    console.error('Upload error:', err);
    return res.status(500).json({ error: err.message });
  }
}
