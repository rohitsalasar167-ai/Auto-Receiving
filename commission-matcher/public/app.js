/* ═══════════════════════════════════════
   CommissionMatcher — app.js
   ═══════════════════════════════════════ */

/* ── State ── */
let masterData = null, masterCols = [], masterMapping = {};
let stmtData   = null, stmtCols   = [], stmtMapping   = {};
let matchResults = [];

/* ── Field definitions ── */
const MASTER_FIELDS = [
  { key: 'cno',       label: 'C.No (Control No)',               req: true  },
  { key: 'instno',    label: 'Inst No  (0 = base, >0 = endo)', req: true  },
  { key: 'polno',     label: 'Policy No',                       req: false },
  { key: 'endno',     label: 'Endorsement No',                  req: false },
  { key: 'custname',  label: 'Customer Name',                   req: false },
  { key: 'insurer',   label: 'Insurer Name',                    req: false },
  { key: 'dept',      label: 'Department',                      req: false },
  { key: 'ptype',     label: 'Policy Type',                     req: false },
  { key: 'odprem',    label: 'OD Premium',                      req: false },
  { key: 'tpprem',    label: 'TP Premium',                      req: false },
  { key: 'totprem',   label: 'Total Premium',                   req: false },
  { key: 'sumins',    label: 'Sum Insured',                     req: false },
  { key: 'startdate', label: 'Policy Start Date',               req: false },
  { key: 'enddate',   label: 'Policy End Date',                 req: false },
  { key: 'brokdue',   label: 'Brok Due (Commission Due)',       req: false },
];

const STMT_FIELDS = [
  { key: 'polno',      label: 'Policy No',          req: false },
  { key: 'endno',      label: 'Endorsement No',     req: false },
  { key: 'custname',   label: 'Customer Name',      req: false },
  { key: 'insurer',    label: 'Insurer Name',       req: false },
  { key: 'odprem',     label: 'OD Premium',         req: false },
  { key: 'tpprem',     label: 'TP Premium',         req: false },
  { key: 'totprem',    label: 'Total Premium',      req: false },
  { key: 'startdate',  label: 'Policy Start Date',  req: false },
  { key: 'enddate',    label: 'Policy End Date',    req: false },
  { key: 'commission', label: 'Commission Amount',  req: false },
];

/* ── Tab switching ── */
const TAB_IDS = ['master', 'match', 'results', 'samples', 'rules'];

function switchTab(name) {
  TAB_IDS.forEach(id => {
    const sec = document.getElementById('tab-' + id);
    const btn = document.querySelector(`.nav-item[data-tab="${id}"]`);
    if (sec) sec.classList.toggle('active', id === name);
    if (btn) btn.classList.toggle('active', id === name);
  });

  if (name === 'match') {
    const hasMaster = !!masterData;
    document.getElementById('match-no-master').style.display  = hasMaster ? 'none'  : 'block';
    document.getElementById('match-main-area').style.display  = hasMaster ? 'block' : 'none';
  }

  if (name === 'results') {
    const hasResults = matchResults.length > 0;
    document.getElementById('results-empty').style.display   = hasResults ? 'none'  : 'block';
    document.getElementById('results-content').style.display = hasResults ? 'block' : 'none';
    if (hasResults) renderResultsTable();
  }
}

document.querySelectorAll('.nav-item[data-tab]').forEach(btn => {
  btn.addEventListener('click', () => switchTab(btn.dataset.tab));
});

/* ── Helpers ── */
function norm(s) {
  return String(s === null || s === undefined ? '' : s)
    .toLowerCase().replace(/[\s\-\/\.\,]+/g, '').trim();
}

function esc(s) {
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function toNum(v) {
  const n = parseFloat(String(v === null || v === undefined ? '' : v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? null : n;
}

function premClose(a, b) {
  const na = toNum(a), nb = toNum(b);
  if (na === null || nb === null) return false;
  if (na === 0 && nb === 0) return true;
  if (na === 0 || nb === 0) return false;
  return Math.abs(na - nb) / Math.max(Math.abs(na), Math.abs(nb)) < 0.02;
}

function isZeroBrokDue(v) {
  const n = toNum(v);
  return n !== null && Math.abs(n) <= 5;
}

function nameMatch(a, b) {
  const na = norm(a), nb = norm(b);
  if (!na || !nb) return false;
  if (na === nb) return true;
  const words = na.split(/\s+/).filter(w => w.length > 2);
  if (!words.length) return false;
  const hits = words.filter(w => nb.includes(w)).length;
  return hits / words.length >= 0.6;
}

function isEndoRow(v) {
  const s = norm(v);
  return !(!s || s === '0' || s === 'na' || s === 'nil' || s === 'none');
}

function extractYear(v) {
  if (!v && v !== 0) return null;
  if (typeof v === 'number' && v > 10000) {
    try {
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      const y = d.getUTCFullYear();
      return (y > 1900 && y < 2100) ? y : null;
    } catch { return null; }
  }
  const s = String(v);
  const m1 = s.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/);
  if (m1) return parseInt(m1[3]);
  const m2 = s.match(/(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (m2) return parseInt(m2[1]);
  const m3 = s.match(/\b(20\d{2}|19\d{2})\b/);
  if (m3) return parseInt(m3[1]);
  return null;
}

function yearMatch(mStart, mEnd, sStart, sEnd) {
  const mSY = extractYear(mStart), mEY = extractYear(mEnd);
  const sSY = extractYear(sStart), sEY = extractYear(sEnd);
  if ((mSY === null && mEY === null) || (sSY === null && sEY === null)) return true;
  if (mSY !== null && sSY !== null && mSY === sSY) return true;
  if (mEY !== null && sEY !== null && mEY === sEY) return true;
  if (mSY !== null && sEY !== null && mSY === sEY) return true;
  if (mEY !== null && sSY !== null && mEY === sSY) return true;
  return false;
}

/* ── Auto-detect column mapping ── */
const HINTS = {
  cno:       ['c.no','cno','control no','ctrl no'],
  instno:    ['inst no','instno','inst'],
  polno:     ['policy no','policy number','pol no','polno','policy'],
  endno:     ['endorsement no','endorsement','endorse','end no','endno','endt no','endt'],
  custname:  ['customer name','customer','insured name','insured','client'],
  insurer:   ['insurer name','insurer','insurance company','ins co'],
  dept:      ['department','dept','branch','class of business'],
  ptype:     ['policy type','type','product','plan'],
  odprem:    ['od premium','od prem','own damage','od'],
  tpprem:    ['tp premium','tp prem','third party','tp'],
  totprem:   ['total premium','total prem','net premium','gross premium','premium','tot prem'],
  sumins:    ['sum insured','sum assured','si'],
  startdate: ['policy start date','start date','inception date','from date','issue date'],
  enddate:   ['policy end date','end date','expiry date','to date','expiry'],
  brokdue:   ['brok due','brokerage due','commission due','brok'],
  commission:['commission amount','commission','comm','net comm','brokerage'],
};

function autoDetect(cols, fields) {
  const m = {};
  fields.forEach(f => {
    const hw = HINTS[f.key] || [];
    const found = cols.find(c => {
      const cn = norm(c);
      return hw.some(h => cn === norm(h) || cn.includes(norm(h)));
    });
    if (found) m[f.key] = found;
  });
  return m;
}

function buildMappingGrid(gridId, fields, cols, mapping, storeKey) {
  const grid = document.getElementById(gridId);
  grid.innerHTML = '';
  fields.forEach(f => {
    const wrap = document.createElement('div');
    const opts = cols.map(c =>
      `<option value="${esc(c)}"${(mapping[f.key] || '') === c ? ' selected' : ''}>${esc(c)}</option>`
    ).join('');
    const req = f.req ? ' <span style="color:var(--red)">*</span>' : '';
    wrap.innerHTML = `
      <div>
        <label>${esc(f.label)}${req}</label>
        <select onchange="window['${storeKey}']['${f.key}']=this.value">
          <option value="">-- not mapped --</option>
          ${opts}
        </select>
      </div>`;
    grid.appendChild(wrap);
  });
}

/* ── Master File ── */
window.masterMapping = masterMapping;
window.stmtMapping   = stmtMapping;

function onMasterFileSelected(inp) {
  const file = inp.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'binary' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!rows.length) { alert('No data found in file'); return; }
      masterCols = Object.keys(rows[0]);
      masterData = rows;
      masterMapping = autoDetect(masterCols, MASTER_FIELDS);
      window.masterMapping = masterMapping;
      buildMappingGrid('master-mapping-grid', MASTER_FIELDS, masterCols, masterMapping, 'masterMapping');
      document.getElementById('master-upload-title').textContent = file.name;
      document.getElementById('master-dropzone').classList.add('has-file');
      document.getElementById('master-mapping-area').style.display = 'block';
      document.getElementById('master-ready-area').style.display   = 'none';
    } catch (err) { alert('Error reading file: ' + err.message); }
  };
  reader.readAsBinaryString(file);
}

function saveMasterMapping() {
  document.querySelectorAll('#master-mapping-grid select').forEach((s, i) => {
    if (s.value) masterMapping[MASTER_FIELDS[i].key] = s.value;
  });
  if (!masterMapping.cno)    { alert('Please map the C.No column');    return; }
  if (!masterMapping.instno) { alert('Please map the Inst No column'); return; }

  document.getElementById('master-mapping-area').style.display = 'none';
  document.getElementById('master-ready-area').style.display   = 'block';

  const base = masterData.filter(r => {
    const v = r[masterMapping.instno];
    return norm(v) === '0' || v === 0 || norm(v) === '';
  }).length;
  const skip = masterMapping.brokdue
    ? masterData.filter(r => isZeroBrokDue(r[masterMapping.brokdue])).length : 0;

  document.getElementById('stat-total').textContent = masterData.length.toLocaleString();
  document.getElementById('stat-base').textContent  = base.toLocaleString();
  document.getElementById('stat-endo').textContent  = (masterData.length - base).toLocaleString();
  document.getElementById('stat-skip').textContent  = skip.toLocaleString();

  const badge = document.getElementById('nav-master-badge');
  badge.textContent = masterData.length.toLocaleString();
  badge.style.display = 'inline-block';

  renderPreviewTable('master-preview-table', masterData.slice(0, 6), masterCols.slice(0, 7));
}

function resetMaster() {
  masterData = null; masterCols = []; masterMapping = {}; window.masterMapping = {};
  document.getElementById('master-ready-area').style.display   = 'none';
  document.getElementById('master-mapping-area').style.display = 'none';
  document.getElementById('master-dropzone').classList.remove('has-file');
  document.getElementById('master-upload-title').textContent   = 'Drop your master Excel here';
  document.getElementById('master-file-input').value           = '';
  document.getElementById('nav-master-badge').style.display    = 'none';
}

/* ── Statement File ── */
function onStmtFileSelected(inp) {
  const file = inp.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'binary' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!rows.length) { alert('No data found in file'); return; }
      stmtCols = Object.keys(rows[0]);
      stmtData = rows;
      stmtMapping = autoDetect(stmtCols, STMT_FIELDS);
      window.stmtMapping = stmtMapping;
      buildMappingGrid('stmt-mapping-grid', STMT_FIELDS, stmtCols, stmtMapping, 'stmtMapping');
      document.getElementById('stmt-upload-title').textContent = file.name;
      document.getElementById('stmt-dropzone').classList.add('has-file');
      document.getElementById('stmt-mapping-area').style.display = 'block';
    } catch (err) { alert('Error reading file: ' + err.message); }
  };
  reader.readAsBinaryString(file);
}

/* ── Core Match Logic ── */
function matchOneRow(sRow) {
  const sPolno    = norm(sRow[stmtMapping.polno]);
  const sEndno    = norm(sRow[stmtMapping.endno]);
  const sCust     = String(sRow[stmtMapping.custname] || '');
  const sTotPrem  = sRow[stmtMapping.totprem];
  const sOdPrem   = sRow[stmtMapping.odprem];
  const sStart    = sRow[stmtMapping.startdate];
  const sEnd      = sRow[stmtMapping.enddate];
  const stmtIsEnd = isEndoRow(sRow[stmtMapping.endno]);

  const pool = masterData.filter(m => {
    const mInst   = m[masterMapping.instno];
    const mIsBase = norm(mInst) === '0' || mInst === 0 || norm(mInst) === '';
    if (stmtIsEnd && mIsBase)  return false;
    if (!stmtIsEnd && !mIsBase) return false;
    if (masterMapping.brokdue && isZeroBrokDue(m[masterMapping.brokdue])) return false;
    return true;
  });

  let best = null, bestLevel = 0, bestOn = [];

  for (const m of pool) {
    const mPolno   = norm(m[masterMapping.polno]);
    const mEndno   = norm(m[masterMapping.endno]);
    const mCust    = String(m[masterMapping.custname] || '');
    const mTotPrem = m[masterMapping.totprem];
    const mOdPrem  = m[masterMapping.odprem];
    const mStart   = m[masterMapping.startdate];
    const mEnd     = m[masterMapping.enddate];

    const polHit  = !!(sPolno && mPolno && sPolno === mPolno);
    const endHit  = !!(stmtIsEnd && sEndno && mEndno && sEndno === mEndno);
    const premHit = premClose(sTotPrem, mTotPrem) || premClose(sOdPrem, mOdPrem);
    const custHit = nameMatch(sCust, mCust);
    const dateOk  = yearMatch(mStart, mEnd, sStart, sEnd);

    if (!dateOk) continue;

    let level = 0, on = [];
    if      (polHit && custHit && premHit) { level = 100; on = ['PolNo', 'CustName', 'Premium', 'Year']; }
    else if (polHit && premHit)             { level = 80;  on = ['PolNo', 'Premium', 'Year']; }
    else if (polHit && endHit)              { level = 70;  on = ['PolNo', 'EndoNo', 'Year']; }
    else if (endHit && premHit)             { level = 65;  on = ['EndoNo', 'Premium', 'Year']; }
    else if (polHit)                        { level = 50;  on = ['PolNo', 'Year']; }
    else if (custHit && premHit && !mPolno) { level = 45;  on = ['CustName', 'Premium', 'Year']; }

    if (level > bestLevel) { bestLevel = level; best = m; bestOn = on; }
  }

  const status = bestLevel === 100 ? '100' : bestLevel >= 45 ? 'partial' : 'unmatched';

  return {
    _stmtRow:    sRow,
    _cno:        best ? String(best[masterMapping.cno]    ?? '') : '',
    _instno:     best ? String(best[masterMapping.instno] ?? '') : '',
    _matchLevel: bestLevel,
    _matchStatus: status,
    _matchedOn:  bestOn.join(' + '),
  };
}

/* ── Run Matching ── */
async function runMatching() {
  if (!masterData || !stmtData) { alert('Upload both files first'); return; }

  document.querySelectorAll('#stmt-mapping-grid select').forEach((s, i) => {
    if (s.value) stmtMapping[STMT_FIELDS[i].key] = s.value;
  });

  document.getElementById('stmt-mapping-area').style.display  = 'none';
  document.getElementById('matching-progress').style.display  = 'block';

  const log  = document.getElementById('match-log');
  const fill = document.getElementById('prog-fill');
  const lbl  = document.getElementById('prog-label');

  const addLog = (msg, cls = '') => {
    const span = document.createElement('span');
    span.className = cls; span.textContent = msg + '\n';
    log.appendChild(span); log.scrollTop = log.scrollHeight;
  };

  matchResults = [];
  const total = stmtData.length;
  const skip  = masterMapping.brokdue
    ? masterData.filter(r => isZeroBrokDue(r[masterMapping.brokdue])).length : 0;

  addLog(`Processing ${total} statement rows…`, 'log-ok');
  addLog(`Master pool: ${masterData.length} records  (${skip} excluded — Brok Due ≈ 0)`, 'log-warn');

  let c100 = 0, cP = 0, cU = 0;
  const BATCH = 10;

  for (let i = 0; i < total; i += BATCH) {
    const batch = stmtData.slice(i, Math.min(i + BATCH, total));
    fill.style.width = Math.round((i / total) * 100) + '%';
    lbl.textContent  = `Row ${i + 1} – ${Math.min(i + BATCH, total)} of ${total}`;
    for (const row of batch) {
      const r = matchOneRow(row);
      matchResults.push(r);
      if (r._matchStatus === '100')      c100++;
      else if (r._matchStatus === 'partial') cP++;
      else                               cU++;
    }
    await new Promise(res => setTimeout(res, 5));
  }

  fill.style.width = '100%';
  lbl.textContent  = 'Done!';
  addLog(`100% matched: ${c100}  |  Partial: ${cP}  |  Unmatched: ${cU}`, 'log-ok');
  if (cU > 0) addLog(`${cU} rows need manual review`, 'log-warn');

  const badge = document.getElementById('nav-results-badge');
  badge.textContent    = matchResults.length.toLocaleString();
  badge.style.display  = 'inline-block';

  setTimeout(() => {
    document.getElementById('matching-progress').style.display = 'none';
    switchTab('results');
  }, 900);
}

/* ── Render Results Table ── */
function renderResultsTable() {
  const filter = document.getElementById('results-filter').value;
  const rows   = matchResults.filter(r => filter === 'all' || r._matchStatus === filter);

  document.getElementById('r-total').textContent    = matchResults.length;
  document.getElementById('r-100').textContent      = matchResults.filter(r => r._matchStatus === '100').length;
  document.getElementById('r-partial').textContent  = matchResults.filter(r => r._matchStatus === 'partial').length;
  document.getElementById('r-unmatched').textContent = matchResults.filter(r => r._matchStatus === 'unmatched').length;

  const visCols = stmtCols.slice(0, 5);
  const th = [...visCols, 'C.No', 'Inst No', 'Matched On', 'Status']
    .map(c => `<th>${esc(String(c))}</th>`).join('');

  const tb = rows.slice(0, 300).map(r => {
    const cells = visCols.map(c => `<td>${esc(String(r._stmtRow[c] ?? ''))}</td>`).join('');
    let badge = '';
    if      (r._matchStatus === '100')      badge = '<span class="badge green">100% Match</span>';
    else if (r._matchStatus === 'partial')  badge = '<span class="badge amber">Partial</span>';
    else                                    badge = '<span class="badge red">Unmatched</span>';
    const matchedOn = r._matchedOn
      ? `<span style="font-size:11px;font-family:'DM Mono',monospace;color:var(--text3)">${esc(r._matchedOn)}</span>`
      : '—';
    return `<tr>${cells}<td><strong>${esc(r._cno)}</strong></td><td>${esc(r._instno)}</td><td>${matchedOn}</td><td>${badge}</td></tr>`;
  }).join('');

  document.getElementById('results-table').innerHTML =
    `<table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table>`;
}

/* ── Download Results ── */
function downloadResults() {
  if (!matchResults.length) { alert('No results to download'); return; }
  try {
    const wb = XLSX.utils.book_new();

    const buildRows = arr => arr.map(r => {
      const obj = {};
      stmtCols.forEach(c => { obj[c] = r._stmtRow[c] !== undefined ? r._stmtRow[c] : ''; });
      obj['C.No']         = r._cno;
      obj['Inst No']      = r._instno;
      obj['Matched On']   = r._matchedOn || '';
      obj['Match Status'] = r._matchStatus === '100' ? '100% Match'
                          : r._matchStatus === 'partial' ? 'Partial Match' : 'Unmatched';
      return obj;
    });

    const addSheet = (rows, name) => {
      const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ Note: 'No data' }]);
      XLSX.utils.book_append_sheet(wb, ws, name);
    };

    addSheet(buildRows(matchResults),                                        'All Results');
    addSheet(buildRows(matchResults.filter(r => r._matchStatus === '100')),  '100% Matched');
    addSheet(buildRows(matchResults.filter(r => r._matchStatus === 'partial')), 'Partial Match');
    addSheet(buildRows(matchResults.filter(r => r._matchStatus === 'unmatched')), 'Unmatched');

    XLSX.writeFile(wb, 'commission_matched_' + new Date().toISOString().slice(0, 10) + '.xlsx');
  } catch (err) {
    alert('Download error: ' + err.message);
    console.error(err);
  }
}

/* ── Render Preview Table ── */
function renderPreviewTable(containerId, rows, cols) {
  const wrap = document.getElementById(containerId);
  const th   = cols.map(c => `<th>${esc(String(c))}</th>`).join('');
  const tb   = rows.map(r =>
    `<tr>${cols.map(c => `<td>${esc(String(r[c] ?? ''))}</td>`).join('')}</tr>`
  ).join('');
  wrap.innerHTML = `<table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table>`;
}

/* ── Sample Template Downloads ── */
function downloadMasterTemplate() {
  try {
    const wb  = XLSX.utils.book_new();
    const hdr = ['C.No','Inst No','Policy No','Endorsement No','Customer Name','Insurer Name',
                 'Department','Policy Type','OD Premium','TP Premium','Total Premium','Sum Insured',
                 'Policy Start Date','Policy End Date','Brok Due'];
    const data = [
      hdr,
      ['CNT001', 0,  'MOT/2024/001234', '',         'Rajesh Kumar',      'New India Assurance', 'Motor',  'Package Policy',  8500, 2200, 10700,   500000, '01/04/2024', '31/03/2025', 1070],
      ['CNT002', 1,  'MOT/2024/001234', 'END/001',  'Rajesh Kumar',      'New India Assurance', 'Motor',  'Endorsement',      500,    0,   500,        '', '15/06/2024', '31/03/2025',   50],
      ['CNT003', 0,  'FIR/2024/005678', '',         'Priya Sharma',      'Oriental Insurance',  'Fire',   'IAR',                0,    0, 25000, 2000000, '01/07/2024', '30/06/2025', 2500],
      ['CNT004', 0,  'HLT/2024/009012', '',         'Mehta Enterprises', 'Star Health',         'Health', 'Group Mediclaim',    0,    0, 45000, 1000000, '01/04/2024', '31/03/2025', 4500],
      ['CNT005', 0,  'MOT/2024/007890', '',         'Anita Desai',       'Bajaj Allianz',       'Motor',  'Third Party',        0, 3500,  3500,       0, '15/05/2024', '14/05/2025',  200],
      ['CNT006', 2,  'HLT/2024/009012', 'END/002',  'Mehta Enterprises', 'Star Health',         'Health', 'Endorsement',        0,    0,  5000,   50000, '01/08/2024', '31/03/2025',  500],
      ['CNT007', 0,  'MAR/2024/003344', '',         'Cargo Movers Ltd',  'ICICI Lombard',       'Marine', 'Cargo',              0,    0, 18000,  750000, '01/06/2024', '31/05/2025', 1800],
      ['CNT008', 0,  'TRV/2024/008877', '',         'Sunita Patel',      'HDFC Ergo',           'Travel', 'International',      0,    0,  3200,       0, '10/09/2024', '25/09/2024',  320],
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = hdr.map(() => ({ wch: 20 }));
    XLSX.utils.book_append_sheet(wb, ws, 'Master Data');
    XLSX.writeFile(wb, 'master_policy_template.xlsx');
  } catch (err) { alert('Download error: ' + err.message); }
}

function downloadStmtTemplate() {
  try {
    const wb  = XLSX.utils.book_new();
    const hdr = ['Policy No','Endorsement No','Customer Name','Insurer Name',
                 'OD Premium','TP Premium','Total Premium',
                 'Policy Start Date','Policy End Date','Commission Amount'];
    const data = [
      hdr,
      ['MOT/2024/001234', '',        'Rajesh Kumar',      'New India Assurance',  8500, 2200, 10700, '01/04/2024', '31/03/2025', 1070],
      ['MOT/2024/001234', 'END/001', 'Rajesh Kumar',      'New India Assurance',   500,    0,   500, '15/06/2024', '31/03/2025',   50],
      ['FIR/2024/005678', '',        'Priya Sharma',      'Oriental Insurance',      0,    0, 25000, '01/07/2024', '30/06/2025', 2500],
      ['HLT/2024/009012', '',        'Mehta Enterprises', 'Star Health',             0,    0, 45000, '01/04/2024', '31/03/2025', 4500],
      ['MOT/2024/007890', '',        'Anita Desai',       'Bajaj Allianz',           0, 3500,  3500, '15/05/2024', '14/05/2025',  200],
      ['HLT/2024/009012', 'END/002', 'Mehta Enterprises', 'Star Health',             0,    0,  5000, '01/08/2024', '31/03/2025',  500],
      ['MAR/2024/003344', '',        'Cargo Movers Ltd',  'ICICI Lombard',           0,    0, 18000, '01/06/2024', '31/05/2025', 1800],
      ['TRV/2024/008877', '',        'Sunita Patel',      'HDFC Ergo',               0,    0,  3200, '10/09/2024', '25/09/2024',  320],
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = hdr.map(() => ({ wch: 22 }));
    XLSX.utils.book_append_sheet(wb, ws, 'Commission Statement');
    XLSX.writeFile(wb, 'commission_statement_template.xlsx');
  } catch (err) { alert('Download error: ' + err.message); }
}
