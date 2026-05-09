/* ═══════════════════════════════════════════════
   CommissionMatcher v5 — app.js
   New in v5:
   - Editable default rule with Reset button
   - Policy No Transform modes (icici, strip-prefix, extract-segment, none)
   - Inst No Mapping modes (endo-seq, direct, ignore)
   - Extended rules tab: Transform, InstNo, Score, Advanced sections
   - ICICI default: strip 4225/ prefix, concat core+renewal as master key
   - Endo sequence → Inst No numeric part matching
   ── Carried from v4 ──
   - Share % support, OD/TP/Total any-match
   - Deep name cleaning, fuzzy token matching
   - Per-insurer rules, consolidated statement
═══════════════════════════════════════════════ */

const SK_MASTER   = 'cm_master_data_v4';
const SK_MASTER_M = 'cm_master_mapping_v4';
const SK_RULES    = 'cm_insurer_rules_v5';
const SK_CONSO    = 'cm_consolidated_v4';
const SK_DEF_RULE = 'cm_default_rule_v5';

/* ── Built-in default rule (reflects your ICICI manual matching) ── */
const FACTORY_DEFAULT_RULE = {
  polnoMode:      'contains',   // ICICI: stmt polno substring matches master polno
  polnoTransform: 'icici',      // strip insurer code, concat core+renewal
  instnoMode:     'endo-seq',   // extract numeric from 4th segment → match Inst No
  suffixN:        10,
  premTol:        2,
  dateCheck:      'no',         // ICICI motor policies: date check off by default
  useShare:       true,
  nameStrip:      true,
  fuzzyPct:       60,
  minScore:       45,
  scorePolCustPrem: 100,
  scorePolPrem:     95,
  scorePolOnly:     80,
  scoreEndoOnly:    80,
  scorePolEndo:     80,
  scoreCustPrem:    65,
  scorePolCust:     60,
  skipZeroBrok:     true,
  skipBaseForEndo:  true,
  notes: 'Auto-configured from ICICI template: strip 4225/ prefix, match core+renewal segment against master policy no',
};

let defaultRule = {...FACTORY_DEFAULT_RULE};

let masterData    = null, masterCols = [], masterMapping = {};
let stmtData      = null, stmtCols   = [], stmtMapping   = {};
let matchResults  = [];
let insurerRules  = [];
let consoData     = [];

window.masterMapping = masterMapping;
window.stmtMapping   = stmtMapping;

/* ── Field defs ── */
const MASTER_FIELDS = [
  {key:'cno',       label:'C.No (Control No)',         req:true},
  {key:'instno',    label:'Inst No (0=base, >0=endo)', req:true},
  {key:'polno',     label:'Policy No',                 req:false},
  {key:'endno',     label:'Endorsement No',            req:false},
  {key:'custname',  label:'Customer Name',             req:false},
  {key:'insurer',   label:'Insurer Name',              req:false},
  {key:'dept',      label:'Department',                req:false},
  {key:'ptype',     label:'Policy Type',               req:false},
  {key:'odprem',    label:'OD Premium',                req:false},
  {key:'tpprem',    label:'TP Premium',                req:false},
  {key:'totprem',   label:'Total Premium',             req:false},
  {key:'sharepct',  label:'Share % (your share)',      req:false},
  {key:'sumins',    label:'Sum Insured',               req:false},
  {key:'startdate', label:'Policy Start Date',         req:false},
  {key:'enddate',   label:'Policy End Date',           req:false},
  {key:'brokdue',   label:'Brok Due (Commission Due)', req:false},
];
const STMT_FIELDS = [
  {key:'polno',      label:'Policy No',         req:false},
  {key:'endno',      label:'Endorsement No',    req:false},
  {key:'custname',   label:'Customer Name',     req:false},
  {key:'insurer',    label:'Insurer Name',      req:false},
  {key:'odprem',     label:'OD Premium',        req:false},
  {key:'tpprem',     label:'TP Premium',        req:false},
  {key:'totprem',    label:'Total Premium',     req:false},
  {key:'startdate',  label:'Policy Start Date', req:false},
  {key:'enddate',    label:'Policy End Date',   req:false},
  {key:'commission', label:'Commission Amount', req:false},
];

/* ═══ HELPERS ═══ */
function norm(s){return String(s===null||s===undefined?'':s).toLowerCase().replace(/[\s\-\/\.\,]+/g,'').trim();}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function toNum(v){const n=parseFloat(String(v===null||v===undefined?'':v).replace(/[^0-9.\-]/g,''));return isNaN(n)?null:n;}

/* Share-aware premium: returns 100% equivalent of master premium */
function fullPrem(val, sharePct){
  const n = toNum(val);
  if(n === null) return null;
  const s = toNum(sharePct);
  if(s === null || s <= 0 || s >= 100) return n; // no valid share → use as-is
  return n / (s / 100);
}

/* Check any of OD/TP/Total against statement, using share % */
function premiumMatch(mRow, sOD, sTP, sTotal, tolPct, useShare){
  const share = useShare ? mRow._share : null;
  const mOD   = fullPrem(mRow._odprem,  share);
  const mTP   = fullPrem(mRow._tpprem,  share);
  const mTot  = fullPrem(mRow._totprem, share);

  const close = (a, b) => {
    if(a===null||b===null) return false;
    if(a===0&&b===0) return true;
    if(a===0||b===0) return false;
    return Math.abs(a-b)/Math.max(Math.abs(a),Math.abs(b))*100 <= tolPct;
  };

  if(close(mOD,  toNum(sOD)))   return true;
  if(close(mTP,  toNum(sTP)))   return true;
  if(close(mTot, toNum(sTotal))) return true;
  return false;
}

function isZeroBrokDue(v){const n=toNum(v);return n!==null&&Math.abs(n)<=5;}
function isEndoRow(v){const s=norm(v);return !(!s||s==='0'||s==='na'||s==='nil'||s==='none');}

function extractYear(v){
  if(!v&&v!==0) return null;
  if(typeof v==='number'&&v>10000){
    try{const d=new Date(Math.round((v-25569)*86400*1000));const y=d.getUTCFullYear();return(y>1900&&y<2100)?y:null;}catch{return null;}
  }
  const s=String(v);
  const m1=s.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/);if(m1)return parseInt(m1[3]);
  const m2=s.match(/(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);if(m2)return parseInt(m2[1]);
  const m3=s.match(/\b(20\d{2}|19\d{2})\b/);if(m3)return parseInt(m3[1]);
  return null;
}

function yearMatch(mStart,mEnd,sStart,sEnd){
  const mSY=extractYear(mStart),mEY=extractYear(mEnd);
  const sSY=extractYear(sStart),sEY=extractYear(sEnd);
  if((mSY===null&&mEY===null)||(sSY===null&&sEY===null)) return true;
  if(mSY!==null&&sSY!==null&&mSY===sSY) return true;
  if(mEY!==null&&sEY!==null&&mEY===sEY) return true;
  if(mSY!==null&&sEY!==null&&mSY===sEY) return true;
  if(mEY!==null&&sSY!==null&&mEY===sSY) return true;
  return false;
}

/* ── Name cleaning ── */
const NAME_STRIP_RE = /\b(mr|mrs|ms|miss|dr|prof|shri|smt|km|master|m\/s|m\.s\.|messrs|ltd|limited|pvt|private|llp|llc|inc|corp|co|and\s+co|sons|brothers|enterprises|industries|traders|agency|services|group|holdings|international|india|bharat|national|works|company|firm)\b\.?/gi;

function cleanName(s){
  return String(s||'')
    .replace(NAME_STRIP_RE,'')
    .replace(/[^a-z0-9\s]/gi,' ')
    .replace(/\s+/g,' ')
    .trim()
    .toLowerCase();
}

function nameMatch(a, b, thresholdPct=60){
  const ca = cleanName(a), cb = cleanName(b);
  if(!ca||!cb) return 0;
  if(ca===cb) return 100;

  const wa = ca.split(/\s+/).filter(w=>w.length>1);
  const wb = cb.split(/\s+/).filter(w=>w.length>1);
  if(!wa.length||!wb.length) return 0;

  // token overlap both ways
  const fwd = wa.filter(w=>cb.includes(w)).length / wa.length;
  const rev = wb.filter(w=>ca.includes(w)).length / wb.length;
  const score = Math.round(Math.max(fwd, rev) * 100);
  return score >= thresholdPct ? score : 0;
}

/* ── Policy No Transform ── */
function transformPolno(rawPolno, transform){
  // rawPolno is the normalized (lowercased, no spaces) statement polno
  // We need to work on the original (un-normalized) polno for segment splitting
  return rawPolno; // Will be handled pre-normalized below
}

/* Transform statement policy no before matching (raw, un-normalized string) */
function transformStmtPolno(rawStmt, transform){
  if(!rawStmt) return {key: norm(rawStmt), endoSeq: null};
  const s = String(rawStmt).trim();
  const parts = s.split('/');

  if(transform === 'icici'){
    // ICICI format: 4225/CORE/RENEWAL/ENDO → master key = norm(CORE + RENEWAL)
    // e.g. 4225/1000512039/00/0000 → '100051203900'
    // Also handles 4226/..., 4193i/..., etc with letter suffixes
    // For ICICI-style (4-digit or 4-digit+letter code first segment):
    if(parts.length >= 3){
      // Check if first segment is a short insurer code (≤6 chars) and second is the policy core
      const first = parts[0];
      const second = parts[1];
      const third = parts.length >= 3 ? parts[2] : '';
      const fourth = parts.length >= 4 ? parts[3] : null;
      // ICICI numeric codes: 4225, 4226 → strip first segment, concat 2nd+3rd
      if(/^\d{3,5}[a-z]?$/i.test(first) && /^\d{6,}$/.test(second)){
        const key = norm(second + third);
        // endo seq = 4th segment stripped of leading zeros → integer
        const endoRaw = fourth ? parseInt(fourth, 10) : 0;
        return {key, endoSeq: endoRaw};
      }
    }
    // Fallback: use full normalized polno
    return {key: norm(s), endoSeq: null};
  }

  if(transform === 'strip-prefix'){
    // Remove first segment (insurer code), use rest joined without separator
    if(parts.length >= 2){
      return {key: norm(parts.slice(1).join('')), endoSeq: null};
    }
    return {key: norm(s), endoSeq: null};
  }

  if(transform === 'extract-core'){
    // Take segments 2 and 3 (0-indexed: 1 and 2) — the core policy number + renewal
    if(parts.length >= 3){
      return {key: norm(parts[1] + parts[2]), endoSeq: null};
    }
    return {key: norm(s), endoSeq: null};
  }

  // 'none' or default: use full polno as-is
  return {key: norm(s), endoSeq: null};
}

/* Extract the numeric sequence from endo field for Inst No matching */
function extractEndoSeq(stmtPolno, stmtEndno){
  // From statement policy no 4th segment: 4225/CORE/RENEWAL/ENDO → parseInt(ENDO)
  const polParts = String(stmtPolno||'').split('/');
  if(polParts.length >= 4){
    const seq = parseInt(polParts[3], 10);
    if(!isNaN(seq)) return seq;
  }
  // Try endorsement number 4th segment
  const endParts = String(stmtEndno||'').split('/');
  if(endParts.length >= 4){
    const seq = parseInt(endParts[3], 10);
    if(!isNaN(seq)) return seq;
  }
  return 0;
}

/* Check if Inst No in master matches the endo sequence from statement */
function instnoMatch(masterInstno, endoSeq, instnoMode){
  if(instnoMode === 'ignore') return true; // don't check inst
  const inst = String(masterInstno||'').trim();
  if(inst === '0' || inst === '') return endoSeq === 0; // base policy
  if(instnoMode === 'endo-seq'){
    // Extract numeric part from master inst no: E1→1, M12→12, E32→32
    const m = inst.match(/(\d+)$/);
    if(m) return parseInt(m[1], 10) === endoSeq;
    return false;
  }
  if(instnoMode === 'direct'){
    // Compare full inst no string with endo seq string
    return norm(inst) === norm(String(endoSeq));
  }
  return true;
}

/* ── Policy No matching modes ── */
function polnoHit(sPolno, mPolno, mode, suffixN=8){
  if(!sPolno||!mPolno) return false;
  if(mode==='exact')    return sPolno===mPolno;
  if(mode==='contains') return mPolno.includes(sPolno)||sPolno.includes(mPolno);
  if(mode==='suffix'){
    const n=parseInt(suffixN)||8;
    return sPolno.slice(-n)===mPolno.slice(-n);
  }
  return false;
}

/* ── Auto-detect ── */
const HINTS={
  cno:['c.no','cno','control no','ctrl no'],
  instno:['inst no','instno','inst'],
  polno:['policy no','policy number','pol no','polno','policy'],
  endno:['endorsement no','endorsement','endorse','end no','endno','endt no','endt'],
  custname:['customer name','customer','insured name','insured','client'],
  insurer:['insurer name','insurer','insurance company','ins co'],
  dept:['department','dept','branch'],
  ptype:['policy type','type','product','plan'],
  odprem:['od premium','od prem','own damage','od'],
  tpprem:['tp premium','tp prem','third party','tp'],
  totprem:['total premium','total prem','net premium','gross premium','premium','tot prem'],
  sharepct:['share %','share%','share pct','our share','share','broker share','co-insurer share'],
  sumins:['sum insured','sum assured','si'],
  startdate:['policy start date','start date','inception date','from date','issue date'],
  enddate:['policy end date','end date','expiry date','to date','expiry'],
  brokdue:['brok due','brokerage due','commission due','brok'],
  commission:['commission amount','commission','comm','net comm','brokerage'],
};
function autoDetect(cols,fields){
  const m={};
  fields.forEach(f=>{
    const hw=HINTS[f.key]||[];
    const found=cols.find(c=>{const cn=norm(c);return hw.some(h=>cn===norm(h)||cn.includes(norm(h)));});
    if(found)m[f.key]=found;
  });
  return m;
}
function buildMappingGrid(gridId,fields,cols,mapping,storeVar){
  const g=document.getElementById(gridId);g.innerHTML='';
  fields.forEach(f=>{
    const d=document.createElement('div');
    const opts=cols.map(c=>`<option value="${esc(c)}"${(mapping[f.key]||'')===c?' selected':''}>${esc(c)}</option>`).join('');
    const req=f.req?` <span class="req">*</span>`:'';
    d.innerHTML=`<div><label>${esc(f.label)}${req}</label><select onchange="window['${storeVar}']['${f.key}']=this.value"><option value="">-- not mapped --</option>${opts}</select></div>`;
    g.appendChild(d);
  });
}

/* ═══ TABS ═══ */
const TABS=['master','rules','match','results','conso','samples','howrules'];
function switchTab(name){
  TABS.forEach(id=>{
    const s=document.getElementById('tab-'+id);
    const b=document.querySelector(`.nav-item[data-tab="${id}"]`);
    if(s)s.classList.toggle('active',id===name);
    if(b)b.classList.toggle('active',id===name);
  });
  if(name==='match'){const has=!!masterData;document.getElementById('match-no-master').style.display=has?'none':'block';document.getElementById('match-main-area').style.display=has?'block':'none';}
  if(name==='results'){const has=matchResults.length>0;document.getElementById('results-empty').style.display=has?'none':'block';document.getElementById('results-content').style.display=has?'block':'none';if(has)renderResultsTable();}
  if(name==='conso')renderConsoPage();
}
document.querySelectorAll('.nav-item[data-tab]').forEach(b=>b.addEventListener('click',()=>switchTab(b.dataset.tab)));

function showToast(msg,d=2800){const t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),d);}

/* ═══ MASTER DATA ═══ */
function onMasterFileSelected(inp){
  const file=inp.files[0];if(!file)return;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'binary'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
      if(!rows.length){alert('No data found');return;}
      masterCols=Object.keys(rows[0]);masterData=rows;
      masterMapping=autoDetect(masterCols,MASTER_FIELDS);window.masterMapping=masterMapping;
      buildMappingGrid('master-mapping-grid',MASTER_FIELDS,masterCols,masterMapping,'masterMapping');
      document.getElementById('master-upload-title').textContent=file.name;
      document.getElementById('master-dropzone').classList.add('has-file');
      document.getElementById('master-mapping-area').style.display='block';
      document.getElementById('master-ready-area').style.display='none';
    }catch(err){alert('Error: '+err.message);}
  };
  reader.readAsBinaryString(file);
}

function saveMasterMapping(){
  document.querySelectorAll('#master-mapping-grid select').forEach((s,i)=>{if(s.value)masterMapping[MASTER_FIELDS[i].key]=s.value;});
  if(!masterMapping.cno){alert('Please map C.No');return;}
  if(!masterMapping.instno){alert('Please map Inst No');return;}
  // Pre-compute helper fields on each row for faster matching
  masterData.forEach(r=>{
    r._odprem  = r[masterMapping.odprem]  ?? '';
    r._tpprem  = r[masterMapping.tpprem]  ?? '';
    r._totprem = r[masterMapping.totprem] ?? '';
    r._share   = r[masterMapping.sharepct]?? '';
    r._polno   = norm(r[masterMapping.polno]  ?? '');
    r._endno   = norm(r[masterMapping.endno]  ?? '');
    r._cust    = String(r[masterMapping.custname] ?? '');
    r._start   = r[masterMapping.startdate] ?? '';
    r._end     = r[masterMapping.enddate]   ?? '';
    r._instno  = r[masterMapping.instno];
    r._isBase  = norm(r._instno)==='0'||r._instno===0||norm(r._instno)==='';
    r._cno     = String(r[masterMapping.cno] ?? '');
    r._brok    = r[masterMapping.brokdue] ?? '';
  });
  try{localStorage.setItem(SK_MASTER,JSON.stringify(masterData));localStorage.setItem(SK_MASTER_M,JSON.stringify(masterMapping));}catch(e){console.warn('localStorage save failed');}
  showMasterReady();showToast('✅ Master data saved to browser storage');
}

function showMasterReady(){
  document.getElementById('master-mapping-area').style.display='none';
  document.getElementById('master-upload-area').style.display='none';
  document.getElementById('master-ready-area').style.display='block';
  document.getElementById('master-saved-banner').style.display='none';
  const base=masterData.filter(r=>r._isBase).length;
  const skip=masterData.filter(r=>isZeroBrokDue(r._brok)).length;
  document.getElementById('stat-total').textContent=masterData.length.toLocaleString();
  document.getElementById('stat-base').textContent=base.toLocaleString();
  document.getElementById('stat-endo').textContent=(masterData.length-base).toLocaleString();
  document.getElementById('stat-skip').textContent=skip.toLocaleString();
  const badge=document.getElementById('nav-master-badge');badge.textContent=masterData.length.toLocaleString();badge.style.display='inline-block';
  document.getElementById('storage-status').innerHTML='<span class="sdot green"></span> Master saved ('+masterData.length+' records)';
  renderPreviewTable('master-preview-table',masterData.slice(0,6),masterCols.slice(0,8));
}

function resetMaster(){
  if(!confirm('Remove master data from browser?'))return;
  masterData=null;masterCols=[];masterMapping={};window.masterMapping={};
  localStorage.removeItem(SK_MASTER);localStorage.removeItem(SK_MASTER_M);
  document.getElementById('master-ready-area').style.display='none';
  document.getElementById('master-mapping-area').style.display='none';
  document.getElementById('master-upload-area').style.display='block';
  document.getElementById('master-saved-banner').style.display='none';
  document.getElementById('master-dropzone').classList.remove('has-file');
  document.getElementById('master-upload-title').textContent='Drop your master Excel here';
  document.getElementById('master-file-input').value='';
  document.getElementById('nav-master-badge').style.display='none';
  document.getElementById('storage-status').innerHTML='<span class="sdot gray"></span> No data saved';
  showToast('Master data cleared');
}

/* ═══ INSURER RULES ═══ */
function loadRules(){
  try{const r=localStorage.getItem(SK_RULES);if(r)insurerRules=JSON.parse(r);}catch{}
  try{const d=localStorage.getItem(SK_DEF_RULE);if(d)defaultRule={...FACTORY_DEFAULT_RULE,...JSON.parse(d)};}catch{}
  renderRulesTable();
  renderDefaultRuleForm();
}
function saveRules(){try{localStorage.setItem(SK_RULES,JSON.stringify(insurerRules));}catch{}}
function saveDefaultRule(){try{localStorage.setItem(SK_DEF_RULE,JSON.stringify(defaultRule));}catch{}}
function resetDefaultRule(){if(!confirm('Reset default rule to factory settings?'))return;defaultRule={...FACTORY_DEFAULT_RULE};saveDefaultRule();renderDefaultRuleForm();showToast('Default rule reset to factory settings');}

function renderDefaultRuleForm(){
  const fields=[
    ['def-polno-mode','polnoMode'],['def-polno-transform','polnoTransform'],['def-instno-mode','instnoMode'],
    ['def-prem-tol','premTol'],['def-date','dateCheck'],['def-share','useShare'],
    ['def-namestrip','nameStrip'],['def-fuzzy','fuzzyPct'],['def-min-score','minScore'],
    ['def-suffix-n','suffixN'],
    ['def-score-pol-cust-prem','scorePolCustPrem'],['def-score-pol-prem','scorePolPrem'],
    ['def-score-pol-only','scorePolOnly'],['def-score-endo-only','scoreEndoOnly'],
    ['def-score-pol-endo','scorePolEndo'],['def-score-cust-prem','scoreCustPrem'],
    ['def-score-pol-cust','scorePolCust'],
    ['def-skip-zero-brok','skipZeroBrok'],['def-skip-base-endo','skipBaseForEndo'],
    ['def-notes','notes'],
  ];
  fields.forEach(([elId,key])=>{
    const el=document.getElementById(elId);
    if(!el) return;
    const v=defaultRule[key];
    if(el.type==='checkbox') el.checked=!!v;
    else el.value=v!==undefined?String(v):'';
  });
  toggleDefSuffix();
}

function saveDefaultRuleFromForm(){
  const r={};
  const get=(id,type)=>{const e=document.getElementById(id);if(!e)return undefined;return type==='num'?parseFloat(e.value):type==='bool'?e.value==='yes'||e.checked:e.value;};
  r.polnoMode=get('def-polno-mode');r.polnoTransform=get('def-polno-transform');r.instnoMode=get('def-instno-mode');
  r.suffixN=get('def-suffix-n','num');r.premTol=get('def-prem-tol','num');r.dateCheck=get('def-date');
  r.useShare=get('def-share','bool');r.nameStrip=get('def-namestrip','bool');
  r.fuzzyPct=get('def-fuzzy','num');r.minScore=get('def-min-score','num');
  r.scorePolCustPrem=get('def-score-pol-cust-prem','num');r.scorePolPrem=get('def-score-pol-prem','num');
  r.scorePolOnly=get('def-score-pol-only','num');r.scoreEndoOnly=get('def-score-endo-only','num');
  r.scorePolEndo=get('def-score-pol-endo','num');r.scoreCustPrem=get('def-score-cust-prem','num');
  r.scorePolCust=get('def-score-pol-cust','num');
  r.skipZeroBrok=get('def-skip-zero-brok','bool');r.skipBaseForEndo=get('def-skip-base-endo','bool');
  r.notes=get('def-notes');
  defaultRule=r;saveDefaultRule();showToast('✅ Default rule saved');
}

function toggleDefSuffix(){
  const v=document.getElementById('def-polno-mode');
  if(v) document.getElementById('def-suffix-wrap').style.display=v.value==='suffix'?'block':'none';
}
function saveRules(){try{localStorage.setItem(SK_RULES,JSON.stringify(insurerRules));}catch{}}

// Show/hide suffix-n field
document.getElementById('r-polno-mode').addEventListener('change',function(){
  document.getElementById('r-suffix-wrap').style.display=this.value==='suffix'?'block':'none';
});

function addInsurerRule(){
  const insurer=document.getElementById('r-insurer').value.trim();
  if(!insurer){alert('Insurer name required');return;}
  const rule={
    insurer,
    polnoMode:   document.getElementById('r-polno-mode').value,
    suffixN:     parseInt(document.getElementById('r-suffix-n').value)||8,
    premTol:     parseFloat(document.getElementById('r-prem-tol').value)||2,
    dateCheck:   document.getElementById('r-date').value,
    useShare:    document.getElementById('r-share').value==='yes',
    nameStrip:   document.getElementById('r-namestrip').value==='yes',
    fuzzyPct:    parseInt(document.getElementById('r-fuzzy').value)||60,
    minScore:    parseInt(document.getElementById('r-min-score').value)||45,
    primary:     document.getElementById('r-primary').value,
    secondary:   document.getElementById('r-secondary').value,
    dept:        document.getElementById('r-dept').value.trim(),
    notes:       document.getElementById('r-notes').value.trim(),
  };
  const idx=insurerRules.findIndex(r=>norm(r.insurer)===norm(insurer));
  if(idx>=0)insurerRules[idx]=rule;else insurerRules.push(rule);
  saveRules();renderRulesTable();
  document.getElementById('r-insurer').value='';document.getElementById('r-notes').value='';document.getElementById('r-dept').value='';
  showToast('✅ Rule saved for '+insurer);
}

function deleteRule(i){insurerRules.splice(i,1);saveRules();renderRulesTable();showToast('Rule deleted');}

function renderRulesTable(){
  const empty=document.getElementById('rules-empty');
  const wrap=document.getElementById('rules-table-wrap');
  const hint=document.getElementById('rules-count-hint');
  const badge=document.getElementById('nav-rules-badge');
  if(!insurerRules.length){empty.style.display='block';wrap.style.display='none';hint.textContent='— none yet';badge.style.display='none';return;}
  empty.style.display='none';wrap.style.display='block';
  hint.textContent=`— ${insurerRules.length} rule${insurerRules.length>1?'s':''} saved`;
  badge.textContent=insurerRules.length;badge.style.display='inline-block';
  document.getElementById('rules-tbody').innerHTML=insurerRules.map((r,i)=>`
    <tr>
      <td><strong>${esc(r.insurer)}</strong></td>
      <td>${esc(r.polnoMode)}${r.polnoMode==='suffix'?` (${r.suffixN}c)`:''}</td>
      <td>${r.premTol}%</td><td>${r.dateCheck}</td>
      <td>${r.useShare?'Yes':'No'}</td><td>${r.nameStrip?'Yes':'No'}</td>
      <td>${r.fuzzyPct}%</td><td>${r.minScore}</td>
      <td>${esc(r.primary)}</td><td>${esc(r.secondary)}</td>
      <td style="color:var(--text2)">${esc(r.dept||'—')}</td>
      <td style="color:var(--text2)">${esc(r.notes||'—')}</td>
      <td><button class="btn btn-ghost btn-sm" style="color:var(--red)" onclick="deleteRule(${i})">✕</button></td>
    </tr>`).join('');
}

function getRule(insurerName){
  if(!insurerName)return null;
  const n=norm(insurerName);
  return insurerRules.find(r=>n.includes(norm(r.insurer))||norm(r.insurer).includes(n))||null;
}

/* ═══ STATEMENT LOAD ═══ */
function onStmtFileSelected(inp){
  const file=inp.files[0];if(!file)return;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'binary'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
      if(!rows.length){alert('No data found');return;}
      stmtCols=Object.keys(rows[0]);stmtData=rows;
      stmtMapping=autoDetect(stmtCols,STMT_FIELDS);window.stmtMapping=stmtMapping;
      buildMappingGrid('stmt-mapping-grid',STMT_FIELDS,stmtCols,stmtMapping,'stmtMapping');
      document.getElementById('stmt-upload-title').textContent=file.name;
      document.getElementById('stmt-dropzone').classList.add('has-file');
      document.getElementById('stmt-mapping-area').style.display='block';
    }catch(err){alert('Error: '+err.message);}
  };
  reader.readAsBinaryString(file);
}

/* ═══ MATCHING ENGINE ═══ */
function matchOneRow(sRow){
  const sPolno  = norm(sRow[stmtMapping.polno]);
  const sEndno  = norm(sRow[stmtMapping.endno]);
  const sCust   = String(sRow[stmtMapping.custname]||'');
  const sOD     = sRow[stmtMapping.odprem];
  const sTP     = sRow[stmtMapping.tpprem];
  const sTotal  = sRow[stmtMapping.totprem];
  const sStart  = sRow[stmtMapping.startdate];
  const sEnd    = sRow[stmtMapping.enddate];
  const sIns    = String(sRow[stmtMapping.insurer]||'');
  const stmtEnd = isEndoRow(sRow[stmtMapping.endno]);

  const rule      = getRule(sIns);
  const premTol   = rule ? rule.premTol   : 2;
  const dateReq   = rule ? rule.dateCheck==='yes' : true;
  const polnoMode = rule ? rule.polnoMode : 'exact';
  const suffixN   = rule ? rule.suffixN   : 8;
  const useShare  = rule ? rule.useShare  : true;  // default: always use share %
  const doStrip   = rule ? rule.nameStrip : true;  // default: always strip
  const fuzzyPct  = rule ? rule.fuzzyPct  : 60;
  const minScore  = rule ? rule.minScore  : 45;

  const pool = masterData.filter(m=>{
    if(stmtEnd && m._isBase)  return false;
    if(!stmtEnd && !m._isBase) return false;
    if(isZeroBrokDue(m._brok)) return false;
    return true;
  });

  let best=null, bestScore=0, bestOn=[];

  for(const m of pool){
    const polHit  = polnoMode!=='ignore' ? polnoHit(sPolno, m._polno, polnoMode, suffixN) : false;
    const endHit  = !!(stmtEnd && sEndno && m._endno && sEndno===m._endno);
    const premHit = premiumMatch(m, sOD, sTP, sTotal, premTol, useShare);
    const dateOk  = dateReq ? yearMatch(m._start, m._end, sStart, sEnd) : true;
    const custScr = doStrip ? nameMatch(sCust, m._cust, fuzzyPct) : nameMatch(sCust, m._cust, fuzzyPct);
    const custHit = custScr > 0;

    if(!dateOk) continue;

    let score=0, on=[];

    /* ── Scoring rules in priority order ── */
    if(polHit && custHit && premHit){
      score=100; on=['PolNo','CustName','Premium'];
    } else if(polHit && premHit){
      score=95; on=['PolNo','Premium'];
    } else if(polHit && endHit){
      score=80; on=['PolNo','EndoNo'];
    } else if(stmtEnd && endHit && premHit){
      score=78; on=['EndoNo','Premium'];
    } else if(stmtEnd && endHit){
      score=80; on=['EndoNo'];   // endo-only = 80 for endo rows
    } else if(!stmtEnd && polHit){
      score=80; on=['PolNo'];    // polno-only = 80 for base rows
    } else if(custHit && premHit){
      score=65; on=['CustName','Premium'];
    } else if(polHit && custHit){
      score=60; on=['PolNo','CustName'];
    }

    // Boost if customer name was a strong match
    if(custScr>=90 && on.includes('CustName')) score=Math.min(score+3, 100);

    if(score > bestScore){ bestScore=score; best=m; bestOn=on; }
  }

  // Determine status
  const status = bestScore>=95 ? '100' : bestScore>=minScore ? 'partial' : 'unmatched';
  const dateYr = dateReq ? '+Year' : '';
  const onStr  = bestOn.map(x=>x==='Premium'?'Prem':x).join(' + ') + (bestOn.length&&dateYr?' '+dateYr:'');

  return{
    _stmtRow:   sRow,
    _cno:       best ? best._cno              : '',
    _instno:    best ? String(best._instno??''): '',
    _score:     bestScore,
    _matchStatus: status,
    _matchedOn: onStr,
    _insurer:   sIns,
    _ruleApplied: rule ? rule.insurer : 'default',
  };
}

async function runMatching(){
  if(!masterData||!stmtData){alert('Upload both files first');return;}
  document.querySelectorAll('#stmt-mapping-grid select').forEach((s,i)=>{if(s.value)stmtMapping[STMT_FIELDS[i].key]=s.value;});
  document.getElementById('stmt-mapping-area').style.display='none';
  document.getElementById('matching-progress').style.display='block';

  const log=document.getElementById('match-log');
  const fill=document.getElementById('prog-fill');
  const lbl=document.getElementById('prog-label');
  const addLog=(msg,cls='')=>{const sp=document.createElement('span');sp.className=cls;sp.textContent=msg+'\n';log.appendChild(sp);log.scrollTop=log.scrollHeight;};

  matchResults=[];
  const total=stmtData.length;
  const skip=masterData.filter(r=>isZeroBrokDue(r._brok)).length;
  addLog(`Processing ${total} statement rows…`,'log-ok');
  addLog(`Master: ${masterData.length} records (${skip} excluded — Brok Due ≈ 0)`,'log-warn');
  if(insurerRules.length) addLog(`Custom rules: ${insurerRules.map(r=>r.insurer).join(', ')}`,'log-ok');
  addLog('Share % applied · Name cleaning ON · OD/TP/Total any-match · Fuzzy names ON','log-ok');

  let c100=0,cP=0,cU=0;
  for(let i=0;i<total;i+=10){
    const batch=stmtData.slice(i,Math.min(i+10,total));
    fill.style.width=Math.round(i/total*100)+'%';
    lbl.textContent=`Row ${i+1}–${Math.min(i+10,total)} of ${total}`;
    for(const row of batch){
      const r=matchOneRow(row);matchResults.push(r);
      if(r._matchStatus==='100')c100++;
      else if(r._matchStatus==='partial')cP++;
      else cU++;
    }
    await new Promise(res=>setTimeout(res,4));
  }
  fill.style.width='100%';
  lbl.textContent='Done!';
  addLog(`100% matched: ${c100}  |  Partial: ${cP}  |  Unmatched: ${cU}`,'log-ok');
  if(cU>0) addLog(`${cU} rows need manual review`,'log-warn');

  const badge=document.getElementById('nav-results-badge');
  badge.textContent=matchResults.length;badge.style.display='inline-block';
  setTimeout(()=>{document.getElementById('matching-progress').style.display='none';switchTab('results');},900);
}

/* ═══ RESULTS TABLE ═══ */
function renderResultsTable(){
  const filter=document.getElementById('results-filter').value;
  const rows=matchResults.filter(r=>filter==='all'||r._matchStatus===filter);
  document.getElementById('r-total').textContent=matchResults.length;
  document.getElementById('r-100').textContent=matchResults.filter(r=>r._matchStatus==='100').length;
  document.getElementById('r-partial').textContent=matchResults.filter(r=>r._matchStatus==='partial').length;
  document.getElementById('r-unmatched').textContent=matchResults.filter(r=>r._matchStatus==='unmatched').length;

  const visCols=stmtCols.slice(0,5);
  const th=[...visCols,'C.No','Inst No','Score','Matched On','Rule','Status'].map(c=>`<th>${esc(String(c))}</th>`).join('');
  const tb=rows.slice(0,400).map(r=>{
    const cells=visCols.map(c=>`<td>${esc(String(r._stmtRow[c]??''))}</td>`).join('');
    const badge=r._matchStatus==='100'?'<span class="badge green">100% Match</span>':r._matchStatus==='partial'?'<span class="badge amber">Partial</span>':'<span class="badge red">Unmatched</span>';
    const scorePill=`<span class="score-pill${r._score<80?' amber':''}">${r._score}</span>`;
    const mo=r._matchedOn?`<span style="font-size:11px;font-family:'DM Mono',monospace;color:var(--text3)">${esc(r._matchedOn)}</span>`:'—';
    const rl=`<span style="font-size:11px;color:var(--text3)">${esc(r._ruleApplied)}</span>`;
    return `<tr>${cells}<td><strong>${esc(r._cno)}</strong></td><td>${esc(r._instno)}</td><td>${scorePill}</td><td>${mo}</td><td>${rl}</td><td>${badge}</td></tr>`;
  }).join('');
  document.getElementById('results-table').innerHTML=`<table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table>`;
}

function downloadResults(){
  if(!matchResults.length){alert('No results');return;}
  try{
    const wb=XLSX.utils.book_new();
    const build=arr=>arr.map(r=>{
      const obj={};
      stmtCols.forEach(c=>{obj[c]=r._stmtRow[c]!==undefined?r._stmtRow[c]:'';});
      obj['C.No']=r._cno;obj['Inst No']=r._instno;
      obj['Match Score']=r._score;
      obj['Matched On']=r._matchedOn||'';
      obj['Rule Applied']=r._ruleApplied||'default';
      obj['Match Status']=r._matchStatus==='100'?'100% Match':r._matchStatus==='partial'?'Partial Match':'Unmatched';
      return obj;
    });
    const add=(rows,name)=>{XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows.length?rows:[{Note:'No data'}]),name);};
    add(build(matchResults),'All Results');
    add(build(matchResults.filter(r=>r._matchStatus==='100')),'100% Matched');
    add(build(matchResults.filter(r=>r._matchStatus==='partial')),'Partial Match');
    add(build(matchResults.filter(r=>r._matchStatus==='unmatched')),'Unmatched');
    XLSX.writeFile(wb,'commission_matched_'+new Date().toISOString().slice(0,10)+'.xlsx');
  }catch(err){alert('Download error: '+err.message);}
}

/* ═══ CONSOLIDATED ═══ */
function loadConso(){try{const c=localStorage.getItem(SK_CONSO);if(c)consoData=JSON.parse(c);}catch{}}
function saveConso(){try{localStorage.setItem(SK_CONSO,JSON.stringify(consoData));}catch{alert('Storage full — try clearing old consolidated data.');}}

function addToConsolidated(){
  if(!matchResults.length){alert('No results to add');return;}
  const label=document.getElementById('conso-stmt-label').value.trim()||'Statement '+new Date().toLocaleDateString('en-IN');
  const flatRows=matchResults.map(r=>{
    const obj={};
    stmtCols.forEach(c=>{obj[c]=r._stmtRow[c]!==undefined?r._stmtRow[c]:'';});
    obj['C.No']=r._cno;obj['Inst No']=r._instno;
    obj['Match Score']=r._score;
    obj['Matched On']=r._matchedOn||'';
    obj['Rule Applied']=r._ruleApplied||'default';
    obj['Match Status']=r._matchStatus==='100'?'100% Match':r._matchStatus==='partial'?'Partial Match':'Unmatched';
    obj['_Statement']=label;
    return obj;
  });
  consoData.push({label,date:new Date().toISOString(),rows:flatRows});
  saveConso();
  const total=consoData.reduce((s,b)=>s+b.rows.length,0);
  const badge=document.getElementById('nav-conso-badge');badge.textContent=total;badge.style.display='inline-block';
  document.getElementById('conso-stmt-label').value='';
  showToast(`✅ ${flatRows.length} rows added to Consolidated`);
  switchTab('conso');
}

function renderConsoPage(){
  const allRows=consoData.flatMap(b=>b.rows);
  const has=allRows.length>0;
  document.getElementById('conso-empty').style.display=has?'none':'block';
  document.getElementById('conso-content').style.display=has?'block':'none';
  if(!has)return;
  document.getElementById('c-total').textContent=allRows.length.toLocaleString();
  document.getElementById('c-batches').textContent=consoData.length;
  document.getElementById('c-100').textContent=allRows.filter(r=>r['Match Status']==='100% Match').length.toLocaleString();
  document.getElementById('c-review').textContent=allRows.filter(r=>r['Match Status']!=='100% Match').length.toLocaleString();
  const badge=document.getElementById('nav-conso-badge');badge.textContent=allRows.length;badge.style.display='inline-block';
  document.getElementById('conso-batch-list').innerHTML=consoData.map((b,i)=>`
    <div class="batch-item">
      <div><div class="batch-label">${esc(b.label)}</div>
        <div class="batch-meta">${b.rows.length} rows · ${new Date(b.date).toLocaleDateString('en-IN',{day:'numeric',month:'short',year:'numeric'})}</div>
      </div>
      <div class="batch-actions">
        <span class="badge green">${b.rows.filter(r=>r['Match Status']==='100% Match').length} matched</span>
        <span class="badge amber">${b.rows.filter(r=>r['Match Status']==='Partial Match').length} partial</span>
        <span class="badge red">${b.rows.filter(r=>r['Match Status']==='Unmatched').length} unmatched</span>
        <button class="btn btn-ghost btn-sm" style="color:var(--red)" onclick="removeBatch(${i})">✕ Remove</button>
      </div>
    </div>`).join('');
  renderConsoTable();
}

function renderConsoTable(){
  const filter=document.getElementById('conso-filter').value;
  const allRows=consoData.flatMap(b=>b.rows);
  const rows=filter==='all'?allRows:allRows.filter(r=>{
    if(filter==='100')return r['Match Status']==='100% Match';
    if(filter==='partial')return r['Match Status']==='Partial Match';
    if(filter==='unmatched')return r['Match Status']==='Unmatched';
    return true;
  });
  if(!rows.length){document.getElementById('conso-table').innerHTML='<div style="padding:20px;text-align:center;color:var(--text2);font-size:13px">No rows</div>';return;}
  const cols=Object.keys(rows[0]).filter(c=>!c.startsWith('_')).slice(0,9);
  const th=[...cols,'Match Status'].map(c=>`<th>${esc(String(c))}</th>`).join('');
  const tb=rows.slice(0,500).map(r=>{
    const cells=cols.map(c=>`<td>${esc(String(r[c]??''))}</td>`).join('');
    const s=r['Match Status'];
    const b=s==='100% Match'?'<span class="badge green">100%</span>':s==='Partial Match'?'<span class="badge amber">Partial</span>':'<span class="badge red">Unmatched</span>';
    return `<tr>${cells}<td>${b}</td></tr>`;
  }).join('');
  document.getElementById('conso-table').innerHTML=`<table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table>`;
}

function removeBatch(i){if(!confirm(`Remove "${consoData[i].label}"?`))return;consoData.splice(i,1);saveConso();renderConsoPage();showToast('Batch removed');}
function clearConso(){if(!confirm('Clear ALL consolidated data?'))return;consoData=[];saveConso();document.getElementById('nav-conso-badge').style.display='none';renderConsoPage();showToast('Cleared');}

function downloadConso(){
  const allRows=consoData.flatMap(b=>b.rows);
  if(!allRows.length){alert('No data');return;}
  try{
    const wb=XLSX.utils.book_new();
    const clean=rows=>rows.map(r=>{const o={};Object.keys(r).filter(k=>!k.startsWith('_')).forEach(k=>{o[k]=r[k];});return o;});
    const add=(rows,name)=>{XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows.length?rows:[{Note:'No data'}]),name);};
    add(clean(allRows),'All Consolidated');
    add(clean(allRows.filter(r=>r['Match Status']==='100% Match')),'100% Matched');
    add(clean(allRows.filter(r=>r['Match Status']==='Partial Match')),'Partial Match');
    add(clean(allRows.filter(r=>r['Match Status']==='Unmatched')),'Unmatched');
    consoData.forEach(b=>{add(clean(b.rows),b.label.slice(0,28).replace(/[\\\/\?\*\[\]]/g,'_'));});
    XLSX.writeFile(wb,'consolidated_commission_'+new Date().toISOString().slice(0,10)+'.xlsx');
  }catch(err){alert('Download error: '+err.message);}
}

/* ═══ PREVIEW ═══ */
function renderPreviewTable(cid,rows,cols){
  const w=document.getElementById(cid);
  const th=cols.map(c=>`<th>${esc(String(c))}</th>`).join('');
  const tb=rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(String(r[c]??''))}</td>`).join('')}</tr>`).join('');
  w.innerHTML=`<table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table>`;
}

/* ═══ SAMPLE TEMPLATES ═══ */
function downloadMasterTemplate(){
  try{
    const wb=XLSX.utils.book_new();
    const hdr=['C.No','Inst No','Policy No','Endorsement No','Customer Name','Insurer Name','Department','Policy Type','OD Premium','TP Premium','Total Premium','Share %','Sum Insured','Policy Start Date','Policy End Date','Brok Due'];
    const data=[hdr,
      ['CNT001',0,'MOT/2024/001234','','Rajesh Kumar','New India Assurance','Motor','Package Policy',8500,2200,10700,100,500000,'01/04/2024','31/03/2025',1070],
      ['CNT002',1,'MOT/2024/001234','END/001','Rajesh Kumar','New India Assurance','Motor','Endorsement',500,0,500,100,'','15/06/2024','31/03/2025',50],
      ['CNT003',0,'FIR/2024/005678','','M/s Priya Sharma & Co','Oriental Insurance','Fire','IAR',0,0,15000,60,2000000,'01/07/2024','30/06/2025',1500],
      ['CNT004',0,'HLT/2024/009012','','Mehta Enterprises Pvt Ltd','Star Health','Health','Group Mediclaim',0,0,27000,60,1000000,'01/04/2024','31/03/2025',2700],
      ['CNT005',0,'MOT/2024/007890','','Mrs Anita Desai','Bajaj Allianz','Motor','Third Party',0,3500,3500,100,0,'15/05/2024','14/05/2025',200],
      ['CNT006',2,'HLT/2024/009012','END/002','Mehta Enterprises Pvt Ltd','Star Health','Health','Endorsement',0,0,3000,60,50000,'01/08/2024','31/03/2025',300],
      ['CNT007',0,'MAR/2024/003344','','Cargo Movers Ltd','ICICI Lombard','Marine','Cargo',0,0,10800,60,750000,'01/06/2024','31/05/2025',1080],
      ['CNT008',0,'TRV/2024/008877','','Mr Sunil Patel','HDFC Ergo','Travel','International',0,0,3200,100,0,'10/09/2024','25/09/2024',320],
    ];
    const ws=XLSX.utils.aoa_to_sheet(data);ws['!cols']=hdr.map(()=>({wch:20}));
    XLSX.utils.book_append_sheet(wb,ws,'Master Data');
    XLSX.writeFile(wb,'master_policy_template.xlsx');
  }catch(err){alert('Error: '+err.message);}
}

function downloadStmtTemplate(){
  try{
    const wb=XLSX.utils.book_new();
    const hdr=['Policy No','Endorsement No','Customer Name','Insurer Name','OD Premium','TP Premium','Total Premium','Policy Start Date','Policy End Date','Commission Amount'];
    const data=[hdr,
      ['MOT/2024/001234','','Rajesh Kumar','New India Assurance',8500,2200,10700,'01/04/2024','31/03/2025',1070],
      ['MOT/2024/001234','END/001','Rajesh Kumar','New India Assurance',500,0,500,'15/06/2024','31/03/2025',50],
      ['FIR/2024/005678','','Priya Sharma & Co','Oriental Insurance',0,0,25000,'01/07/2024','30/06/2025',2500],
      ['HLT/2024/009012','','Mehta Enterprises','Star Health',0,0,45000,'01/04/2024','31/03/2025',4500],
      ['MOT/2024/007890','','Anita Desai','Bajaj Allianz',0,3500,3500,'15/05/2024','14/05/2025',200],
      ['HLT/2024/009012','END/002','Mehta Enterprises','Star Health',0,0,5000,'01/08/2024','31/03/2025',500],
    ];
    const ws=XLSX.utils.aoa_to_sheet(data);ws['!cols']=hdr.map(()=>({wch:22}));
    XLSX.utils.book_append_sheet(wb,ws,'Commission Statement');
    XLSX.writeFile(wb,'commission_statement_template.xlsx');
  }catch(err){alert('Error: '+err.message);}
}

/* ═══ INIT ═══ */
function init(){
  loadRules();
  loadConso();
  if(consoData.length){const t=consoData.reduce((s,b)=>s+b.rows.length,0);const b=document.getElementById('nav-conso-badge');b.textContent=t;b.style.display='inline-block';}
  try{
    const saved=localStorage.getItem(SK_MASTER);
    const savedM=localStorage.getItem(SK_MASTER_M);
    if(saved&&savedM){
      masterData=JSON.parse(saved);masterMapping=JSON.parse(savedM);window.masterMapping=masterMapping;
      masterCols=Object.keys(masterData[0]||{});
      // Re-attach computed fields (they're stored as raw data)
      masterData.forEach(r=>{
        r._odprem  = r[masterMapping.odprem]  ?? '';
        r._tpprem  = r[masterMapping.tpprem]  ?? '';
        r._totprem = r[masterMapping.totprem] ?? '';
        r._share   = r[masterMapping.sharepct]?? '';
        r._polno   = norm(r[masterMapping.polno]  ?? '');
        r._endno   = norm(r[masterMapping.endno]  ?? '');
        r._cust    = String(r[masterMapping.custname] ?? '');
        r._start   = r[masterMapping.startdate] ?? '';
        r._end     = r[masterMapping.enddate]   ?? '';
        r._instno  = r[masterMapping.instno];
        r._isBase  = norm(r._instno)==='0'||r._instno===0||norm(r._instno)==='';
        r._cno     = String(r[masterMapping.cno] ?? '');
        r._brok    = r[masterMapping.brokdue] ?? '';
      });
      document.getElementById('master-upload-area').style.display='none';
      document.getElementById('master-saved-banner').style.display='block';
      document.getElementById('master-saved-info').textContent=masterData.length.toLocaleString()+' records loaded.';
      showMasterReady();
    }
  }catch(e){console.warn('Could not restore master data',e);}
}

init();
