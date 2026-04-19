import { put } from '@vercel/blob';

// ── Parsing helpers for SharePoint "Export to CSV" format ─────────────────────

const EXCLUDED_METHODS = new Set([
  'Call for Application (704)', 'EOI (502)', 'EOI (507)',
  'Invitation for Proposal (507)', 'RFI (502)', 'Contract Amendment', 'Other',
]);

function consolidateMethod(raw) {
  if (!raw) return null;
  // raw may be a joined string like "ITB; BAFO" or "RFQ; UN Award"
  const parts = raw.split(';').map(s => s.trim());
  // Check exclusion (single-value only)
  if (parts.length === 1 && EXCLUDED_METHODS.has(parts[0])) return null;
  // LTA / UN Award
  if (/LTA|UN Award/i.test(raw)) return 'LTA';
  // Formal solicitation
  if (/(^|[^a-z])(ITB|RFP|BAFO|RFQ)([^a-z]|$)/i.test(raw)) return 'Formal Solicitation';
  // Informal
  if (/micro purchase|direct procurement|very low value|re-utilisation/i.test(raw)) return 'Informal Solicitation';
  // Catch remaining excluded patterns
  if (/EOI|RFI|invitation for proposal|call for application|contract amendment/i.test(raw)) return null;
  return null;
}

// Parse SharePoint simple JSON arrays: ["Value1","Value2"] → "Value1; Value2"
function parseLookup(raw) {
  if (!raw) return '';
  const v = raw.trim();
  try {
    if (v.startsWith('[')) {
      const arr = JSON.parse(v);
      return arr.filter(x => x != null && x !== '').join('; ');
    }
    if (v.startsWith('{')) {
      const obj = JSON.parse(v);
      return obj.Value ?? '';
    }
  } catch (_) {}
  return raw;
}

// Buyer: plain text "Surname, Firstname (ORG)" → "Surname"
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
  // Plain text
  let name = raw.replace(/\s*\([^)]+\)\s*$/, '').trim();
  if (name.includes(',')) name = name.split(',')[0].trim();
  return name;
}

// Parse European-format numbers: "$6.150.000,00" or "($88.140,3)"
function parseFloat2(raw) {
  if (!raw) return null;
  let s = String(raw).trim();
  const neg = s.startsWith('(') && s.endsWith(')');
  s = s.replace(/[$\s()]/g, '');
  if (!s) return null;
  if (s.includes(',')) {
    // European: . = thousands separator, , = decimal
    s = s.replace(/\./g, '').replace(',', '.');
  } else {
    // No comma — remove dots (thousands only)
    s = s.replace(/\./g, '');
  }
  const n = parseFloat(s);
  if (isNaN(n)) return null;
  return neg ? -n : n;
}

// Parse DD/MM/YYYY or YYYY-MM-DD dates
function parseDate(raw) {
  if (!raw) return null;
  const s = String(raw).trim();
  // DD/MM/YYYY HH:MM or DD/MM/YYYY
  const m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m1) {
    const d = new Date(`${m1[3]}-${m1[2]}-${m1[1]}T00:00:00Z`);
    return isNaN(d.getTime()) ? null : d;
  }
  // YYYY-MM-DD
  const m2 = s.match(/^(\d{4}-\d{2}-\d{2})/);
  if (m2) {
    const d = new Date(m2[1] + 'T00:00:00Z');
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function fmtDate(d) {
  if (!d) return '';
  return d.toISOString().slice(0, 10);
}

function parseCSV(text) {
  const rows = [];
  let col = 0, inQuote = false, field = '', row = [];
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
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
    // ── Column mapping: SharePoint Export → normalised ───────────────────────
    const buyer        = parseUser(r['Buyer'] || '');
    const prValue      = parseFloat2(r['PR Value']);
    const cumulativePO = parseFloat2(r['Cumulative PO $']);
    const savings      = parseFloat2(r['Savings']);
    const rawMethod    = parseLookup(r['Solicitation Method'] || r['SollicitationMethod'] || '');
    const method       = consolidateMethod(rawMethod);
    const stage        = r['Procurement Stage'] || parseLookup(r['ProcurementStage'] || '');
    const planRaw      = r['Part of Procurement Plan'] || parseLookup(r['PartofProcurementPlan'] || '');
    const awardBasis   = parseLookup(r['Award Basis'] || r['AwardBasis'] || '');
    const marketCat    = parseLookup(r['Market Category'] || r['MarketCategory'] || '');
    const projRef      = parseLookup(r['Project Reference'] || r['ProjectReference'] || '');

    if (method === null) return null;

    const prReceived      = parseDate(r['PR Received']             || r['PRReceived']);
    const poDate          = parseDate(r['PO Issuance date']         || r['POIssuancedate']);
    const solIssued       = parseDate(r['Solicitation Issued']      || r['SollicitationIssued']);
    const solClosed       = parseDate(r['Solicitation Closed']      || r['SolicitationClosed']);
    const created         = parseDate(r['Data/ora creazione']       || r['Created']);
    const modified        = parseDate(r['Data/ora modifica']        || r['Modified']);
    const dateAssigned    = parseDate(r['PR Assigned']              || r['PRAssigned']);
    const dateClosed      = parseDate(r['Date Closed']              || r['DateClosed']);
    const techOfferShared = parseDate(r['Technical Offer Shared']   || r['TechnicalOfferShared']);
    const tcoClearance    = parseDate(r['TCO Clearance']            || r['TCOClearance']);
    const ltoClearance    = parseDate(r['LTO Clearance']            || r['LTOClearance']);
    const awardRec        = parseDate(r['Award Recommandation ']    || r['Award Recommandation'] || r['AwardRecommandation']);

    const cycleTime = prReceived && poDate ? Math.round((poDate - prReceived) / 86400000) : null;
    const year = prReceived ? prReceived.getUTCFullYear() : (poDate ? poDate.getUTCFullYear() : null);
    const awardBasisCompetitive = /competitive/i.test(awardBasis);
    const isCompetitive = method === 'Formal Solicitation' || method === 'LTA' || awardBasisCompetitive;
    const isDirect = !isCompetitive;

    let planBucket;
    if (/yes.*part.*procurement.*plan/i.test(planRaw))  planBucket = 'Planned';
    else if (/not planned/i.test(planRaw))               planBucket = 'Unplanned';
    else                                                  planBucket = 'N/A';

    // Use GRMS number as ID (new format) or SharePoint ID (old format)
    const id = r['PR GRMS %23'] || r['ID'] || r['ItemInternalId'] || '';

    return {
      id, title: r['Description'] || r['Title'] || '',
      prgrms: r['PR GRMS %23'] || r['PRGRMS_x0023_'] || '',
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
      poNumber:   r['PO %23'] || r['PO_x0023_'] || '',
      buyingUnit: parseLookup(r['Buying Unit'] || r['BuyingUnit'] || ''),
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

  const secret = process.env.UPLOAD_SECRET;
  if (secret) {
    const provided = req.headers['x-upload-secret'] || '';
    if (provided !== secret) return res.status(401).json({ error: 'Unauthorized' });
  }

  const chunks = [];
  for await (const chunk of req) chunks.push(Buffer.from(chunk));
  const csvText = Buffer.concat(chunks).toString('utf8');
  if (!csvText.trim()) return res.status(400).json({ error: 'Empty body' });

  try {
    const data = processCSV(csvText);
    console.log(`Parsed: ${data.rows.length} rows | methods: ${JSON.stringify(data.methods)} | lastUpdated: ${data.lastUpdated}`);
    if (data.rows.length === 0) {
      return res.status(400).json({ error: 'CSV parsed but produced 0 valid rows. Check it is the SharePoint procurement tracking list export.' });
    }
    await put('dashboard-data.json', JSON.stringify(data), {
      access: 'private',
      contentType: 'application/json',
      addRandomSuffix: false,
      allowOverwrite: true,
    });
    console.log(`Saved to Blob: ${data.rows.length} rows`);
    return res.status(200).json({ ok: true, rows: data.rows.length, lastUpdated: data.lastUpdated });
  } catch (err) {
    console.error('Upload error:', err);
    return res.status(500).json({ error: err.message });
  }
}
