/* =========================================================
   Collection Tracker Dashboard — script.js
   =========================================================
   Data source: Public Google Sheet (read-only via gviz API)
   Two tabs fetched:
     • AY 26-27  — school-level deal / collection data
     • RECON     — individual collection transactions
   ========================================================= */

// ─────────────────────────────────────────────────────────
// CONFIG  (adjust column numbers here if sheet changes)
// Column numbers below are 1-indexed Excel positions
// (i.e., A=1, B=2, … Z=26, AA=27, … BN=66, etc.)
// ─────────────────────────────────────────────────────────
const CFG = {
  SHEET_ID: '1gleyDo0eF7z5FITmCp_ozK_E-gdaUxklOY4-WX78CAg',
  AY_GID:   '1351790268',

  // AY 26-27 columns (1-indexed Excel)
  AY: {
    SAP_ID:      9,   // I  – SAP School ID
    SCHOOL_NAME: 12,  // L  – School Name
    TRUST_NAME:  13,  // M  – Trust / Vendor Name
    STATE:       22,  // V  – State
    TOTAL_DV:    66,  // BN – Total Deal Value
    TOTAL_COLL:  67,  // BO – Total Collection done
    SCHOOL_INV:  639, // XO – School Invoice (invoicing done)
    // Payment schedule: date+amount pairs, VX–WO (1-indexed 596–613)
    PAY_START:   596, // VX – first date column
    CORE_END:    605, // WG – last Core column (5 pairs: VX-VY … WF-WG)
    CSR_START:   606, // WH – first CSR column
    PAY_END:     613, // WO – last CSR column (4 pairs: WH-WI … WN-WO)
    TOTAL_PAY:   614, // WP – Total Payment (scheduled)
    POC:         616, // WR – Collection POC
  },

  // RECON columns (1-indexed Excel)
  RECON: {
    SALE_YEAR:  2,  // B
    SCHOOL_ID:  3,  // C
    // These will be discovered by header name; set -1 as fallback
    COLL_DATE:  -1,
    COLL_AMT:   -1,
  }
};

// ─────────────────────────────────────────────────────────
// GLOBAL STATE
// ─────────────────────────────────────────────────────────
let schools  = [];   // processed school objects from AY tab
let recon    = [];   // processed rows from RECON tab
let debugLog = [];   // for debug panel

// ─────────────────────────────────────────────────────────
// UTILITY
// ─────────────────────────────────────────────────────────
function num(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = parseFloat(String(v).replace(/[₹,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

function fmt(n) {
  if (n === 0) return '—';
  // Format in Indian numbering system
  return '₹' + Math.round(n).toLocaleString('en-IN');
}

function pct(a, b) {
  if (!b || b === 0) return '0.0%';
  return ((a / b) * 100).toFixed(1) + '%';
}

function pctNum(a, b) {
  if (!b || b === 0) return 0;
  return (a / b) * 100;
}

// Parse a gviz cell value.  gviz returns dates as "Date(y,m,d)" strings.
function parseGvizValue(cell) {
  if (!cell || cell.v === null || cell.v === undefined) return '';

  // Date cell: v = "Date(2026,2,1)" (months 0-based)
  if (typeof cell.v === 'string' && cell.v.startsWith('Date(')) {
    const m = cell.v.match(/Date\((\d+),(\d+),(\d+)\)/);
    if (m) return new Date(+m[1], +m[2], +m[3]);
  }

  // Formatted string available → prefer it for dates already rendered as strings
  if (cell.f && typeof cell.v === 'number' && cell.v > 40000 && cell.v < 60000) {
    // Likely an Excel serial date — use formatted value
    return new Date(cell.f);
  }

  return cell.v;
}

function isDate(v) { return v instanceof Date && !isNaN(v.getTime()); }

// Parse various date string formats into Date objects
function parseDate(v) {
  if (!v) return null;
  if (isDate(v)) return v;
  const s = String(v).trim();

  // DD/MM/YYYY or DD-MM-YYYY
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);

  // YYYY-MM-DD (ISO)
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);

  // Native fallback
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function today0() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
}

function daysBetween(a, b) {
  return Math.floor((b - a) / 86400000);
}

function fmtDate(d) {
  if (!d || !isDate(d)) return '—';
  return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
}

// ─────────────────────────────────────────────────────────
// FETCH — gviz JSONP  (works from file:// without a server)
// Uses <script> injection to bypass CORS entirely.
// ─────────────────────────────────────────────────────────
let _jsonpSeq = 0;

function fetchGviz(params, tq) {
  return new Promise((resolve, reject) => {
    const cbName = `_gvizCb${++_jsonpSeq}`;

    // Build URL — JSONP callback is set via tqx responseHandler
    let tqx = `out:json;responseHandler:${cbName}`;
    let url  = `https://docs.google.com/spreadsheets/d/${CFG.SHEET_ID}/gviz/tq?tqx=${encodeURIComponent(tqx)}`;
    url += '&' + new URLSearchParams(params).toString();
    if (tq) url += '&tq=' + encodeURIComponent(tq);

    // Timeout safety
    const timer = setTimeout(() => {
      cleanup();
      reject(new Error('gviz request timed out (15s)'));
    }, 15000);

    function cleanup() {
      clearTimeout(timer);
      delete window[cbName];
      const s = document.getElementById(cbName);
      if (s) s.remove();
    }

    window[cbName] = (data) => {
      cleanup();
      if (data.status === 'error') {
        const msg = (data.errors || []).map(e => e.message).join('; ');
        return reject(new Error('gviz error: ' + msg));
      }
      resolve(data.table);
    };

    const script = document.createElement('script');
    script.id  = cbName;
    script.src = url;
    script.onerror = () => { cleanup(); reject(new Error('Failed to load gviz script (sheet may not be public)')); };
    document.head.appendChild(script);
  });
}

// Build a gviz TQ SELECT string for a list of 1-indexed column numbers
function selectTQ(colNums) {
  return 'select ' + colNums.map(n => `Col${n}`).join(',');
}

// Small delay helper (avoids hammering the API with simultaneous JSONP tags)
function delay(ms) { return new Promise(r => setTimeout(r, ms)); }

// ─────────────────────────────────────────────────────────
// LOAD & PROCESS — AY 26-27 TAB
// ─────────────────────────────────────────────────────────
async function loadAY() {
  const ay = CFG.AY;
  const payCols = [];
  for (let c = ay.PAY_START; c <= ay.PAY_END; c++) payCols.push(c);

  const selectedCols = [
    ay.SAP_ID, ay.SCHOOL_NAME, ay.TRUST_NAME, ay.STATE,
    ay.TOTAL_DV, ay.TOTAL_COLL, ay.SCHOOL_INV,
    ...payCols,
    ay.TOTAL_PAY, ay.POC
  ];

  // ── Strategy: try targeted select first; fall back to full sheet.
  // IMPORTANT: colMap must reflect which strategy succeeded.
  let table;
  let useFullSheet = false;

  try {
    table = await fetchGviz({ gid: CFG.AY_GID, headers: 1 }, selectTQ(selectedCols));
    // Verify we actually got the right number of columns back
    if (!table || !table.cols || table.cols.length < selectedCols.length - 5) {
      throw new Error(`Expected ~${selectedCols.length} cols, got ${table?.cols?.length}`);
    }
    debugLog.push(`✓ AY targeted select OK — ${table.cols.length} cols, ${(table.rows||[]).length} rows`);
  } catch (e) {
    debugLog.push(`✗ Targeted select failed: ${e.message} — falling back to full sheet`);
    await delay(600);
    table = await fetchGviz({ gid: CFG.AY_GID, headers: 1 });
    useFullSheet = true;
    debugLog.push(`✓ AY full sheet fetched — ${table.cols.length} cols, ${(table.rows||[]).length} rows`);
  }

  const colLabels = table.cols.map(c => (c.label || c.id || '').trim());
  debugLog.push(`\nFirst 10 AY column labels: ${colLabels.slice(0, 10).join(' | ')}`);

  // ── Build colMap: Excel col number → index in this response's table.cols
  //   • Targeted select → columns are in the order we listed in selectedCols (0-based idx)
  //   • Full sheet      → column I (Excel #9) is at index 8 (excCol - 1)
  const colMap = {};
  if (useFullSheet) {
    selectedCols.forEach(excCol => { colMap[excCol] = excCol - 1; });
  } else {
    selectedCols.forEach((excCol, idx) => { colMap[excCol] = idx; });
  }

  // ── Also try to locate key columns by header label (more robust than hard-coded numbers)
  const labelMap = {};
  colLabels.forEach((lbl, idx) => { if (lbl) labelMap[lbl.toLowerCase()] = idx; });

  const findByLabel = (...keywords) => {
    for (const kw of keywords) {
      const idx = Object.keys(labelMap).find(k => k.includes(kw.toLowerCase()));
      if (idx !== undefined) return labelMap[idx];
    }
    return -1;
  };

  // Override with label-based discovery for key columns when available
  const labelOverrides = {
    [ay.SAP_ID]:      findByLabel('sap', 'school id'),
    [ay.SCHOOL_NAME]: findByLabel('school name'),
    [ay.TRUST_NAME]:  findByLabel('trust', 'vendor'),
    [ay.STATE]:       findByLabel('state'),
    [ay.TOTAL_DV]:    findByLabel('total dv', 'deal value'),
    [ay.TOTAL_COLL]:  findByLabel('total collection', 'collection done'),
    [ay.SCHOOL_INV]:  findByLabel('school invoice', 'invoicing'),
    [ay.POC]:         findByLabel('poc', 'collection poc', 'responsible'),
  };
  Object.entries(labelOverrides).forEach(([excCol, idx]) => {
    if (idx >= 0) {
      debugLog.push(`  Label override: Col${excCol} → idx ${idx} ("${colLabels[idx]}")`);
      colMap[Number(excCol)] = idx;
    }
  });

  function getCellVal(row, excCol) {
    const idx = colMap[excCol];
    if (idx === undefined || idx < 0 || idx >= row.c.length) return '';
    const cell = row.c[idx];
    return cell ? parseGvizValue(cell) : '';
  }

  // ── Log a sample row for debugging
  if (table.rows && table.rows.length > 0) {
    const sample = table.rows[0];
    debugLog.push(`\nSample row[0] (first 10 cells): ${
      (sample.c || []).slice(0, 10).map((c,i) => `[${i}]=${c?.v ?? 'null'}`).join('  ')
    }`);
    debugLog.push(`  SAP_ID(idx ${colMap[ay.SAP_ID]}): ${getCellVal(sample, ay.SAP_ID)}`);
    debugLog.push(`  Name  (idx ${colMap[ay.SCHOOL_NAME]}): ${getCellVal(sample, ay.SCHOOL_NAME)}`);
    debugLog.push(`  State (idx ${colMap[ay.STATE]}): ${getCellVal(sample, ay.STATE)}`);
    debugLog.push(`  DV    (idx ${colMap[ay.TOTAL_DV]}): ${getCellVal(sample, ay.TOTAL_DV)}`);
  }

  // ── Process rows
  const result = [];
  for (const row of (table.rows || [])) {
    if (!row || !row.c) continue;

    const sapId     = String(getCellVal(row, ay.SAP_ID) || '').trim();
    const name      = String(getCellVal(row, ay.SCHOOL_NAME) || '').trim();
    const trust     = String(getCellVal(row, ay.TRUST_NAME) || '').trim();
    const state     = String(getCellVal(row, ay.STATE) || '').trim();
    const dv        = num(getCellVal(row, ay.TOTAL_DV));
    const collected = num(getCellVal(row, ay.TOTAL_COLL));
    const invoice   = num(getCellVal(row, ay.SCHOOL_INV));
    const poc       = String(getCellVal(row, ay.POC) || '').trim();

    // Skip only rows with absolutely nothing useful
    if (!sapId && !name && !trust && dv === 0 && invoice === 0 && collected === 0) continue;

    // Payment schedule — pairs of (date, amount)
    const payments = [];
    for (let c = ay.PAY_START; c <= ay.PAY_END; c += 2) {
      const dateVal = getCellVal(row, c);
      const amtVal  = getCellVal(row, c + 1);
      const amt     = num(amtVal);
      const date    = parseDate(dateVal);
      const type    = c < ay.CSR_START ? 'Core' : 'CSR';
      if (date || amt > 0) {
        payments.push({ date, amount: amt, type, colNum: c });
      }
    }

    const displayName = name || trust || sapId || 'Unknown';

    result.push({
      sapId, name: displayName, trust, state: state || 'Unknown',
      dv, collected, invoice, poc: poc || 'Unknown',
      due: Math.max(0, invoice - collected),
      payments
    });
  }

  debugLog.push(`\nAY rows processed: ${result.length}`);
  return result;
}

// ─────────────────────────────────────────────────────────
// LOAD & PROCESS — RECON TAB
// ─────────────────────────────────────────────────────────
async function loadRECON() {
  await delay(800);
  const table = await fetchGviz({ sheet: 'RECON', headers: 1 });
  debugLog.push(`\n✓ RECON tab fetched — ${table.cols.length} cols, ${(table.rows||[]).length} rows`);

  const colLabels = table.cols.map(c => (c.label || c.id || '').trim());
  debugLog.push(`RECON column labels: ${colLabels.join(' | ')}`);

  // Fixed column positions (1-indexed Excel → 0-indexed JS)
  // Col B (2)  = Sale Year
  // Col C (3)  = School ID
  // Col J (10) = Amount collected
  // Col K (11) = Collection date
  // Col S (19) = Status (only include rows where value = "received")
  const yearIdx   = CFG.RECON.SALE_YEAR - 1;  // B → 1
  const schoolIdx = CFG.RECON.SCHOOL_ID - 1;  // C → 2
  const amtIdx    = 9;   // J → index 9
  const dateIdx   = 10;  // K → index 10
  const statusIdx = 18;  // S → index 18

  debugLog.push(`RECON fixed cols → year:${yearIdx} school:${schoolIdx} amount(J):${amtIdx} date(K):${dateIdx} status(S):${statusIdx}`);

  // Log first data row for verification
  if (table.rows && table.rows.length > 0) {
    const s = table.rows[0];
    const preview = (s.c || []).map((c, i) => `[${i}]=${c?.v ?? 'null'}`).join('  ');
    debugLog.push(`RECON row[0] preview: ${preview}`);
  }

  const result = [];
  for (const row of (table.rows || [])) {
    if (!row || !row.c) continue;

    const getCellStr = (idx) => {
      if (idx < 0 || idx >= row.c.length) return '';
      const v = parseGvizValue(row.c[idx]);
      return v !== null && v !== undefined ? String(v).trim() : '';
    };

    // Filter: AY 26-27 only
    const saleYear = getCellStr(yearIdx);
    if (saleYear && !/(26.*27|2026)/.test(saleYear)) continue;

    // Filter: status must be "received" (case-insensitive)
    const status = getCellStr(statusIdx);
    if (!status || status.toLowerCase() !== 'received') continue;

    const schoolId = getCellStr(schoolIdx);
    const collDate = parseDate(parseGvizValue(row.c[dateIdx] || {}));
    const amount   = num(parseGvizValue(row.c[amtIdx] || {}));

    if (!collDate && amount === 0) continue;

    result.push({ schoolId, collDate, amount });
  }

  debugLog.push(`RECON rows after filter (26-27 + received): ${result.length}`);
  return result;
}

// ─────────────────────────────────────────────────────────
// SCHOOL LOOKUP MAP  (sapId → school info)
// ─────────────────────────────────────────────────────────
function buildLookup(schoolList) {
  const map = {};
  for (const s of schoolList) {
    if (s.sapId) map[s.sapId] = s;
    if (s.trust) map[s.trust.toLowerCase()] = s;  // fallback by trust name
    if (s.name)  map[s.name.toLowerCase()] = s;   // fallback by school name
  }
  return map;
}

// ─────────────────────────────────────────────────────────
// FILTER HELPERS
// ─────────────────────────────────────────────────────────
function getSelected(id) {
  const el = document.getElementById(id);
  if (!el) return [];
  const vals = [...el.selectedOptions].map(o => o.value);
  return vals.includes('ALL') ? [] : vals;
}

function populateSelect(id, values) {
  const el = document.getElementById(id);
  if (!el) return;
  const prev = getSelected(id);
  el.innerHTML = '<option value="ALL" selected>All</option>';
  for (const v of values) {
    const opt = document.createElement('option');
    opt.value = opt.textContent = v;
    if (prev.includes(v)) opt.selected = true;
    el.appendChild(opt);
  }
}

function unique(arr, key) {
  return [...new Set(arr.map(r => r[key]).filter(Boolean))].sort();
}

function filterSchools(stateFilter, pocFilter) {
  let list = schools;
  if (stateFilter.length) list = list.filter(s => stateFilter.includes(s.state));
  if (pocFilter.length)   list = list.filter(s => pocFilter.includes(s.poc));
  return list;
}

// ─────────────────────────────────────────────────────────
// RENDER HELPERS
// ─────────────────────────────────────────────────────────
function renderKPIs(id, cards) {
  document.getElementById(id).innerHTML = cards.map(c => `
    <div class="kpi-card ${c.cls || ''}">
      <div class="kpi-value">${c.value}</div>
      <div class="kpi-label">${c.label}</div>
    </div>`).join('');
}

function setTbody(tableId, html) {
  const tbody = document.querySelector(`#${tableId} tbody`);
  tbody.innerHTML = html || '<tr class="empty-row"><td colspan="99">No data for the selected filters.</td></tr>';
}

function setTfoot(tableId, html) {
  const tfoot = document.querySelector(`#${tableId} tfoot`);
  if (tfoot) tfoot.innerHTML = html;
}

// ─────────────────────────────────────────────────────────
// TAB 1 — STATE OVERVIEW
// ─────────────────────────────────────────────────────────
function renderOverview(stateFilter) {
  const filtered = filterSchools(stateFilter, []);

  const totDV   = filtered.reduce((s, r) => s + r.dv, 0);
  const totInv  = filtered.reduce((s, r) => s + r.invoice, 0);
  const totColl = filtered.reduce((s, r) => s + r.collected, 0);
  const totDue  = filtered.reduce((s, r) => s + r.due, 0);

  renderKPIs('overall-kpis', [
    { label: 'Total Deal Value',       value: fmt(totDV),   cls: 'blue-border'  },
    { label: 'Invoicing Done',         value: fmt(totInv),  cls: 'blue-border'  },
    { label: '% Invoiced (of DV)',     value: pct(totInv, totDV)               },
    { label: 'Total Collection',       value: fmt(totColl), cls: 'green-border' },
    { label: 'Due Against Invoicing',  value: fmt(totDue),  cls: 'red-border'   },
  ]);

  // Aggregate by state
  const byState = {};
  for (const s of filtered) {
    if (!byState[s.state]) byState[s.state] = { dv: 0, inv: 0, coll: 0 };
    byState[s.state].dv   += s.dv;
    byState[s.state].inv  += s.invoice;
    byState[s.state].coll += s.collected;
  }

  let rows = '';
  for (const st of Object.keys(byState).sort()) {
    const d   = byState[st];
    const due = Math.max(0, d.inv - d.coll);
    rows += `<tr>
      <td>${st}</td>
      <td class="num">${fmt(d.dv)}</td>
      <td class="num">${fmt(d.inv)}</td>
      <td class="num">${pct(d.inv, d.dv)}</td>
      <td class="num pos">${fmt(d.coll)}</td>
      <td class="due">${fmt(due)}</td>
    </tr>`;
  }

  setTbody('state-table', rows);
  setTfoot('state-table', `<tr>
    <td><strong>Total</strong></td>
    <td class="num"><strong>${fmt(totDV)}</strong></td>
    <td class="num"><strong>${fmt(totInv)}</strong></td>
    <td class="num"><strong>${pct(totInv, totDV)}</strong></td>
    <td class="num pos"><strong>${fmt(totColl)}</strong></td>
    <td class="due"><strong>${fmt(totDue)}</strong></td>
  </tr>`);
}

function applyOverviewFilter() {
  renderOverview(getSelected('ov-state-filter'));
  renderTop10(document.getElementById('top10-state-filter').value);
}

function clearOverviewFilter() {
  const sel = document.getElementById('ov-state-filter');
  [...sel.options].forEach(o => { o.selected = o.value === 'ALL'; });
  renderOverview([]);
  renderTop10('ALL');
}

// ─────────────────────────────────────────────────────────
// TAB 1 — TOP 10 BY DUES
// ─────────────────────────────────────────────────────────
function renderTop10(stateFilter) {
  const list = stateFilter === 'ALL'
    ? [...schools]
    : schools.filter(s => s.state === stateFilter);

  const top = list.sort((a, b) => b.due - a.due).slice(0, 10);

  const rows = top.map((s, i) => `<tr>
    <td>${i + 1}</td>
    <td>${s.name}</td>
    <td>${s.state}</td>
    <td>${s.poc}</td>
    <td class="num">${fmt(s.dv)}</td>
    <td class="num">${fmt(s.invoice)}</td>
    <td class="num pos">${fmt(s.collected)}</td>
    <td class="due">${fmt(s.due)}</td>
  </tr>`).join('');

  setTbody('top10-table', rows);
}

// ─────────────────────────────────────────────────────────
// TAB 2 — DETAILED ANALYSIS
// ─────────────────────────────────────────────────────────
function renderDetailed(stateFilter, pocFilter, dateFrom, dateTo) {
  const today = today0();
  const filtered = filterSchools(stateFilter, pocFilter);

  // KPIs
  const totDV   = filtered.reduce((s, r) => s + r.dv, 0);
  const totInv  = filtered.reduce((s, r) => s + r.invoice, 0);
  const totColl = filtered.reduce((s, r) => s + r.collected, 0);
  const totDue  = filtered.reduce((s, r) => s + r.due, 0);

  // "Total due till date" = sum of all past-due scheduled payments within date filter
  let totDueTillDate = 0;
  for (const s of filtered) {
    for (const p of s.payments) {
      if (!p.date) continue;
      if (dateFrom && p.date < dateFrom) continue;
      if (dateTo   && p.date > dateTo)   continue;
      if (p.date <= today) totDueTillDate += p.amount;
    }
  }

  renderKPIs('det-kpis', [
    { label: 'Total Clients',         value: filtered.length.toString()            },
    { label: 'Total Deal Value',      value: fmt(totDV),  cls: 'blue-border'       },
    { label: 'Invoicing Done',        value: fmt(totInv), cls: 'blue-border'       },
    { label: '% Invoiced',            value: pct(totInv, totDV)                    },
    { label: 'Collection Done',       value: fmt(totColl),cls: 'green-border'      },
    { label: 'Due vs Invoice',        value: fmt(totDue), cls: 'red-border'        },
    { label: 'Total Due Till Today',  value: fmt(totDueTillDate), cls: 'red-border'},
  ]);

  // ── Client-wise table
  const clientRows = filtered
    .sort((a, b) => b.due - a.due)
    .map(s => {
      let dueTillDate = 0;
      for (const p of s.payments) {
        if (!p.date) continue;
        if (dateFrom && p.date < dateFrom) continue;
        if (dateTo   && p.date > dateTo)   continue;
        if (p.date <= today) dueTillDate += p.amount;
      }
      return `<tr>
        <td>${s.name}</td>
        <td>${s.state}</td>
        <td>${s.poc}</td>
        <td class="num">${fmt(s.dv)}</td>
        <td class="num">${fmt(s.invoice)}</td>
        <td class="num">${pct(s.invoice, s.dv)}</td>
        <td class="num pos">${fmt(s.collected)}</td>
        <td class="due">${fmt(s.due)}</td>
        <td class="due">${fmt(dueTillDate)}</td>
      </tr>`;
    }).join('');

  setTbody('client-table', clientRows);

  // ── POC-wise aging table
  const agingEntries = [];
  for (const s of filtered) {
    for (const p of s.payments) {
      if (!p.date || p.amount <= 0) continue;
      if (dateFrom && p.date < dateFrom) continue;
      if (dateTo   && p.date > dateTo)   continue;
      const daysOverdue = daysBetween(p.date, today);
      agingEntries.push({ poc: s.poc, name: s.name, type: p.type, date: p.date, amount: p.amount, daysOverdue });
    }
  }

  // Sort: POC asc, then daysOverdue desc
  agingEntries.sort((a, b) => {
    if (a.poc < b.poc) return -1;
    if (a.poc > b.poc) return 1;
    return b.daysOverdue - a.daysOverdue;
  });

  const agingRows = agingEntries.map(e => {
    let badge;
    if (e.daysOverdue <= 0)       badge = '<span class="badge on-time">Upcoming</span>';
    else if (e.daysOverdue <= 30) badge = `<span class="badge warn">${e.daysOverdue}d overdue</span>`;
    else if (e.daysOverdue <= 60) badge = `<span class="badge danger">${e.daysOverdue}d overdue</span>`;
    else                          badge = `<span class="badge critical">${e.daysOverdue}d overdue</span>`;

    return `<tr>
      <td>${e.poc}</td>
      <td>${e.name}</td>
      <td>${e.type}</td>
      <td>${fmtDate(e.date)}</td>
      <td class="num">${fmt(e.amount)}</td>
      <td class="num">${e.daysOverdue > 0 ? e.daysOverdue : '—'}</td>
      <td>${badge}</td>
    </tr>`;
  }).join('');

  setTbody('aging-table', agingRows);
}

function applyDetailedFilter() {
  const states = getSelected('det-state');
  const pocs   = getSelected('det-poc');
  const from   = document.getElementById('det-from').value ? new Date(document.getElementById('det-from').value) : null;
  const to     = document.getElementById('det-to').value   ? new Date(document.getElementById('det-to').value)   : null;
  renderDetailed(states, pocs, from, to);
}

function clearDetailedFilter() {
  ['det-state', 'det-poc'].forEach(id => {
    const el = document.getElementById(id);
    [...el.options].forEach(o => { o.selected = o.value === 'ALL'; });
  });
  document.getElementById('det-from').value = '';
  document.getElementById('det-to').value   = '';
  renderDetailed([], [], null, null);
}

// ─────────────────────────────────────────────────────────
// TAB 3 — RECENT COLLECTIONS (last 7 days from RECON)
// ─────────────────────────────────────────────────────────
function renderRecent(stateFilter) {
  const today      = today0();
  const cutoff     = new Date(today);
  cutoff.setDate(cutoff.getDate() - 7);

  // Build lookup for enriching RECON rows with school info
  const lookup = buildLookup(schools);

  const enriched = recon
    .filter(r => r.collDate && r.collDate >= cutoff && r.collDate <= today)
    .map(r => {
      // Match by exact ID, then by lowercase name/trust
      const school = lookup[r.schoolId]
        || lookup[(r.schoolId || '').toLowerCase()]
        || null;
      return {
        date:  r.collDate,
        name:  school ? school.name  : (r.schoolId || 'Unknown'),
        state: school ? school.state : 'Unknown',
        poc:   school ? school.poc   : 'Unknown',
        amount: r.amount
      };
    })
    .filter(r => {
      if (!stateFilter.length) return true;
      return stateFilter.includes(r.state);
    })
    .sort((a, b) => b.date - a.date);

  const totAmt = enriched.reduce((s, r) => s + r.amount, 0);

  renderKPIs('rec-summary', [
    { label: 'Collections (7 days)',  value: enriched.length.toString()    },
    { label: 'Total Amount',          value: fmt(totAmt), cls: 'green-border' },
  ]);

  const rows = enriched.map(r => `<tr>
    <td>${fmtDate(r.date)}</td>
    <td>${r.name}</td>
    <td>${r.state}</td>
    <td>${r.poc}</td>
    <td class="num pos">${fmt(r.amount)}</td>
  </tr>`).join('');

  setTbody('recent-table', rows);
}

function applyRecentFilter()  { renderRecent(getSelected('rec-state-filter')); }
function clearRecentFilter()  {
  const el = document.getElementById('rec-state-filter');
  [...el.options].forEach(o => { o.selected = o.value === 'ALL'; });
  renderRecent([]);
}

// ─────────────────────────────────────────────────────────
// DEBUG PANEL
// ─────────────────────────────────────────────────────────
function toggleDebug() {
  const panel = document.getElementById('debug-panel');
  panel.classList.toggle('hidden');
  document.getElementById('debug-content').textContent = debugLog.join('\n');
}

// ─────────────────────────────────────────────────────────
// TAB NAVIGATION
// ─────────────────────────────────────────────────────────
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(s => s.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
  });
});

// ─────────────────────────────────────────────────────────
// INIT
// ─────────────────────────────────────────────────────────
async function init() {
  debugLog = [];

  const loadingEl = document.getElementById('loading');
  const appEl     = document.getElementById('app');
  const msgEl     = document.getElementById('loading-msg');

  loadingEl.classList.remove('hidden');
  appEl.classList.add('hidden');
  loadingEl.innerHTML = '<div class="spinner"></div><p id="loading-msg">Connecting to Google Sheets…</p>';

  try {
    document.getElementById('loading-msg').textContent = 'Loading AY 26-27 data…';
    schools = await loadAY();

    document.getElementById('loading-msg').textContent = 'Loading RECON data…';
    try {
      recon = await loadRECON();
    } catch (reconErr) {
      debugLog.push(`\n⚠ RECON fetch failed: ${reconErr.message}`);
      debugLog.push('  Recent Collections tab will be empty.');
      recon = [];
    }

    // Populate all filter dropdowns
    const states = unique(schools, 'state');
    const pocs   = unique(schools, 'poc');

    populateSelect('ov-state-filter', states);
    populateSelect('top10-state-filter', states);
    populateSelect('det-state', states);
    populateSelect('det-poc', pocs);
    populateSelect('rec-state-filter', states);

    // top10 state filter is a single-select with 'ALL'
    const top10Sel = document.getElementById('top10-state-filter');
    top10Sel.innerHTML = '<option value="ALL">All States</option>';
    states.forEach(s => {
      const o = document.createElement('option');
      o.value = o.textContent = s;
      top10Sel.appendChild(o);
    });
    top10Sel.onchange = (e) => renderTop10(e.target.value);

    // Initial renders
    renderOverview([]);
    renderTop10('ALL');
    renderDetailed([], [], null, null);
    renderRecent([]);

    document.getElementById('last-updated').textContent =
      'Last updated: ' + new Date().toLocaleString('en-IN');

    loadingEl.classList.add('hidden');
    appEl.classList.remove('hidden');

  } catch (err) {
    debugLog.push('\nFATAL: ' + err.message);
    loadingEl.innerHTML = `
      <div class="error-state">
        <h3>Failed to load data</h3>
        <p>${err.message}</p>
        <p>Check that the Google Sheet is publicly accessible (Anyone with the link can view).</p>
        <pre>${debugLog.join('\n')}</pre>
        <button onclick="init()">Retry</button>
      </div>`;
  }
}

// Start
init();
