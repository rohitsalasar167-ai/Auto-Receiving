/* ═══════════════════════════════════════════════
   CommissionMatcher v5 — app.js
   Full rewrite with:
   - Editable default rule (persisted, resettable)
   - Per-insurer rules with 20+ configurable fields
   - Policy No transform: ICICI strip, suffix, contains, segment-extract
   - Inst No mode: endo-seq matching (4th segment → E1/M2 etc)
   - Score weights: each field has its own weight, fully configurable
   - Premium: both % tolerance AND absolute value tolerance
   - Zero policy no fallback: use name+premium+date when polno=0
   - Pool filter: skip zero brok, skip base for endo, dept filter
   - Boost system: strong name, exact polno, exact premium all boost
   - Confidence labels: 100% / High / Medium / Low / Unmatched
═══════════════════════════════════════════════ */

/* ── Storage keys ── */
const SK_MASTER   = 'cm_master_v5';
const SK_MASTER_M = 'cm_master_map_v5';
const SK_RULES    = 'cm_rules_v5';
const SK_DEFRULE  = 'cm_defrule_v5';
const SK_CONSO    = 'cm_conso_v5';

/* ── State ── */
let masterData = null, masterCols = [], masterMapping = {};
let stmtData   = null, stmtCols   = [], stmtMapping   = {};
let matchResults = [];
let insurerRules = [];
let consoData    = [];
let defaultRule  = null; // loaded from storage or factory

window.masterMapping = masterMapping;
window.stmtMapping   = stmtMapping;

/* ══════════════════════════════════════════════
   FACTORY DEFAULT RULE
   Based on analysis of ICICI template data:
   - Policy No in statement: 4225/1000512039/00/0000
   - Master policy no:       100051203900 (strip prefix, concat core+renewal)
   - Inst No: endo seq from 4th segment (0001→E1, 0002→E2/M2)
══════════════════════════════════════════════ */
const FACTORY_DEFAULT = {
  // ── Policy No ──
  polnoMode:        'contains',   // exact | contains | suffix | ignore
  polnoTransform:   'icici',      // none | icici | strip-prefix | extract-segment
  suffixN:          10,
  polnoWeight:      40,           // score weight 0-100

  // ── Endorsement No ──
  endnoWeight:      30,           // score weight for endo match

  // ── Inst No matching ──
  instnoMode:       'endo-seq',   // none | endo-seq | direct
  instnoRequired:   false,        // if true, inst no must match to count as hit

  // ── Customer Name ──
  nameStrip:        true,
  fuzzyPct:         60,
  nameWeight:       15,           // score weight

  // ── Premium ──
  premTolPct:       2,            // % tolerance
  premTolAbs:       50,           // absolute ₹ tolerance (whichever is larger wins)
  useShare:         true,
  premWeight:       20,           // score weight

  // ── Date / Year ──
  dateCheck:        'no',         // yes | no | optional
  dateWeight:       10,

  // ── Sum Insured ──
  sumInsCheck:      false,
  sumInsTolPct:     5,

  // ── Zero Policy No handling ──
  zeroPolicyFallback: true,       // if polno=0 in master, match using name+prem+date

  // ── Pool filters ──
  skipZeroBrok:     true,
  skipBaseForEndo:  true,
  skipEndoForBase:  true,
  deptFilter:       '',           // filter master pool by dept

  // ── Score thresholds ──
  scoreFullMatch:   95,           // score >= this → "100% Match"
  scorePartial:     45,           // score >= this → "Partial"
  // custom score combos (all configurable)
  scorePolCustPrem: 100,
  scorePolPrem:     95,
  scorePolCust:     70,
  scoreEndoPrem:    85,
  scoreEndoCust:    75,
  scoreEndoOnly:    80,
  scorePolOnly:     80,
  scoreCustPrem:    65,
  scoreCustOnly:    40,
  scorePremOnly:    30,

  // ── Boosts ──
  boostExactPolno:  5,            // bonus if polno is exact match (not just contains)
  boostStrongName:  5,            // bonus if name score ≥ 85
  boostExactPrem:   3,            // bonus if premium matches exactly (0% difference)
  boostYearMatch:   3,            // bonus when year matches and dateCheck='optional'

  // ── Notes ──
  notes: 'Auto-configured for ICICI Lombard motor/health statements. Strips insurer code (4225/4226), matches core+renewal segment against master policy no. Endo sequence from 4th segment maps to Inst No (E1, M2 etc).',
};

/* ── Field definitions ── */
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
  {key:'endodate',  label:'Endo Date (endorsement start date)', req:false},
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

/* ══════════════════════════════════════════════
   HELPERS — Rewritten to match Excel logic exactly
══════════════════════════════════════════════ */
function norm(s){ return String(s===null||s===undefined?'':s).toLowerCase().replace(/[\s\-\/\.\,]+/g,'').trim(); }
function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function toNum(v){ const n=parseFloat(String(v===null||v===undefined?'':v).replace(/[^0-9.\-]/g,'')); return isNaN(n)?null:n; }

/* ── Share-aware premium ── */
function fullPrem(val, sharePct){
  const n=toNum(val); if(n===null) return null;
  const s=toNum(sharePct);
  if(s===null||s<=0||s>=100) return n;
  return n/(s/100);
}

/* ── Premium match: % OR absolute tolerance ── */
function premiumMatch(mRow, sOD, sTP, sTotal, tolPct, tolAbs, useShare){
  const share = useShare ? mRow._share : null;
  const mOD  = fullPrem(mRow._odprem,  share);
  const mTP  = fullPrem(mRow._tpprem,  share);
  const mTot = fullPrem(mRow._totprem, share);

  let exactMatch = false;

  const close=(a,b)=>{
    if(a===null||b===null) return false;
    if(a===0&&b===0){ exactMatch=true; return true; }
    if(a===0||b===0) return false;
    const diff=Math.abs(a-b);
    const pctOk = diff/Math.max(Math.abs(a),Math.abs(b))*100<=tolPct;
    const absOk = diff<=tolAbs;
    if(pctOk&&diff===0) exactMatch=true;
    return pctOk||absOk;
  };

  const hit = close(mOD,toNum(sOD)) || close(mTP,toNum(sTP)) || close(mTot,toNum(sTotal));
  return {hit, exact: exactMatch};
}

function isZeroBrokDue(v){ const n=toNum(v); return n!==null&&Math.abs(n)<=5; }
function isEndoRow(v){ const s=norm(v); return !(!s||s==='0'||s==='na'||s==='nil'||s==='none'); }
function isZeroPolno(v){ const s=norm(v); return !s||s==='0'||s==='na'||s==='nil'||s==='none'; }

function extractYear(v){
  if(!v&&v!==0) return null;
  if(typeof v==='number'&&v>10000){
    try{ const d=new Date(Math.round((v-25569)*86400*1000)); const y=d.getUTCFullYear(); return(y>1900&&y<2100)?y:null; }catch{ return null; }
  }
  const s=String(v);
  const m1=s.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/); if(m1) return parseInt(m1[3]);
  const m2=s.match(/(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/); if(m2) return parseInt(m2[1]);
  const m3=s.match(/\b(20\d{2}|19\d{2})\b/); if(m3) return parseInt(m3[1]);
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
  return String(s||'').replace(NAME_STRIP_RE,'').replace(/[^a-z0-9\s]/gi,' ').replace(/\s+/g,' ').trim().toLowerCase();
}

function nameMatch(a, b, thresholdPct=60){
  const ca=cleanName(a), cb=cleanName(b);
  if(!ca||!cb) return 0;
  if(ca===cb) return 100;
  const wa=ca.split(/\s+/).filter(w=>w.length>1);
  const wb=cb.split(/\s+/).filter(w=>w.length>1);
  if(!wa.length||!wb.length) return 0;
  const fwd=wa.filter(w=>cb.includes(w)).length/wa.length;
  const rev=wb.filter(w=>ca.includes(w)).length/wb.length;
  const score=Math.round(Math.max(fwd,rev)*100);
  return score>=thresholdPct?score:0;
}

/* ══════════════════════════════════════════════
   POLICY NO TRANSFORM (ICICI & others)
   ICICI: 4225/1000512039/00/0000
     → strip first segment (insurer code)
     → concat core (seg2) + renewal (seg3) = 100051203900
     → endo seq from seg4: 0001 → 1 (matches Inst No E1, M1 etc)
══════════════════════════════════════════════ */
function transformStmtPolno(raw, transform){
  if(!raw) return {key:'', endoSeq:0, isEndo:false};
  const s=String(raw).trim();
  const parts=s.split('/');

  if(transform==='icici'){
    // Format: CODE/CORE/RENEWAL/ENDOSEQ (4+ segments)
    // Also handles 5-seg: CODE/X/CORE/RENEWAL/ENDOSEQ
    if(parts.length>=4){
      const first=parts[0];
      // Check if first is short insurer code (≤6 chars, alphanumeric)
      if(/^[\d]{3,5}[a-z]?$/i.test(first)){
        // Check if next part is the core (long numeric)
        let coreIdx=1;
        // Handle 5-segment: 4016/X/261930683/03/013
        if(parts.length>=5 && /^[a-z]{1,2}$/i.test(parts[1])) coreIdx=2;
        const core=parts[coreIdx]||'';
        const renewal=parts[coreIdx+1]||'';
        const endoPart=parts[parts.length-1]||'0';
        const endoSeq=parseInt(endoPart,10)||0;
        const key=norm(core+renewal);
        return {key, endoSeq, isEndo: endoSeq>0};
      }
    }
    // Some ICICI-format polnos are already clean (no prefix)
    return {key:norm(s), endoSeq:0, isEndo:false};
  }

  if(transform==='strip-prefix'){
    if(parts.length>=2) return {key:norm(parts.slice(1).join('')), endoSeq:0, isEndo:false};
    return {key:norm(s), endoSeq:0, isEndo:false};
  }

  if(transform==='extract-segment'){
    // Take middle segments (skip first and last)
    if(parts.length>=3) return {key:norm(parts.slice(1,-1).join('')), endoSeq:0, isEndo:false};
    return {key:norm(s), endoSeq:0, isEndo:false};
  }

  // 'none' — use as-is
  return {key:norm(s), endoSeq:0, isEndo:false};
}

/* Get endo sequence from stmt polno or endno 4th segment */
function getStmtEndoSeq(stmtPolno, stmtEndno){
  const tryParts=(raw)=>{
    const parts=String(raw||'').split('/');
    if(parts.length>=4){ const n=parseInt(parts[parts.length-1],10); if(!isNaN(n)&&n>0) return n; }
    return null;
  };
  return tryParts(stmtPolno) || tryParts(stmtEndno) || 0;
}

/* Check if master Inst No matches endo sequence */
function instnoMatchesSeq(masterInstno, endoSeq, mode){
  if(mode==='none') return true;
  const inst=String(masterInstno||'').trim();
  if(endoSeq===0) return inst==='0'||inst==='';
  if(mode==='endo-seq'){
    // E1→1, E12→12, M4→4, L2E1→1 (last number)
    const m=inst.match(/(\d+)$/);
    return m ? parseInt(m[1],10)===endoSeq : false;
  }
  if(mode==='direct') return norm(inst)===norm(String(endoSeq));
  return true;
}

/* ── Policy No matching ── */
function polnoHit(sKey, mPolno, mode, suffixN=8){
  if(!sKey||!mPolno) return {hit:false, exact:false};
  if(mode==='exact'){
    const hit=sKey===mPolno;
    return {hit, exact:hit};
  }
  if(mode==='contains'){
    const hit=mPolno.includes(sKey)||sKey.includes(mPolno);
    const exact=sKey===mPolno;
    return {hit, exact};
  }
  if(mode==='suffix'){
    const n=parseInt(suffixN)||8;
    const hit=sKey.slice(-n)===mPolno.slice(-n);
    return {hit, exact:sKey===mPolno};
  }
  return {hit:false, exact:false};
}

/* ══════════════════════════════════════════════
   AUTO-DETECT COLUMN MAPPING
══════════════════════════════════════════════ */
const HINTS={
  cno:['c.no','cno','control no','ctrl no','c no'],
  instno:['inst no','instno','inst','installation no'],
  polno:['policy no','policy number','pol no','polno','policy'],
  endno:['endorsement no','endorsement','endorse','end no','endno','endt no','endt'],
  custname:['customer name','customer','insured name','insured','client','party name'],
  insurer:['insurer name','insurer','insurance company','ins co','company'],
  dept:['department','dept','branch','line of business','lob'],
  ptype:['policy type','type','product','plan','class'],
  odprem:['od premium','od prem','own damage','od'],
  tpprem:['tp premium','tp prem','third party','tp'],
  totprem:['total premium','total prem','net premium','gross premium','premium','tot prem','net prem'],
  sharepct:['share %','share%','share pct','our share','share','broker share'],
  sumins:['sum insured','sum assured','si','tsi'],
  startdate:['policy start date','start date','inception date','from date','issue date','risk start'],
  endodate: ['endo date','endorsement date','endt date','endo start','end date of endo'],
  enddate:  ['policy end date','end date','expiry date','to date','expiry','risk end'],
  brokdue:['brok due','brokerage due','commission due','brok','brokerage'],
  commission:['commission amount','commission','comm','net comm','brokerage','brok amount'],
};

function autoDetect(cols,fields){
  const m={};
  fields.forEach(f=>{
    const hw=HINTS[f.key]||[];
    const found=cols.find(c=>{ const cn=norm(c); return hw.some(h=>cn===norm(h)||cn.includes(norm(h))); });
    if(found) m[f.key]=found;
  });
  return m;
}

function buildMappingGrid(gridId,fields,cols,mapping,storeVar){
  const g=document.getElementById(gridId); g.innerHTML='';
  fields.forEach(f=>{
    const d=document.createElement('div');
    const opts=cols.map(c=>`<option value="${esc(c)}"${(mapping[f.key]||'')===c?' selected':''}>${esc(c)}</option>`).join('');
    const req=f.req?` <span class="req">*</span>`:'';
    d.innerHTML=`<div><label>${esc(f.label)}${req}</label><select onchange="window['${storeVar}']['${f.key}']=this.value"><option value="">-- not mapped --</option>${opts}</select></div>`;
    g.appendChild(d);
  });
}

/* ══════════════════════════════════════════════
   TABS
══════════════════════════════════════════════ */
const TABS=['master','rules','match','results','conso','samples','howrules'];
function switchTab(name){
  TABS.forEach(id=>{
    const s=document.getElementById('tab-'+id);
    const b=document.querySelector(`.nav-item[data-tab="${id}"]`);
    if(s) s.classList.toggle('active',id===name);
    if(b) b.classList.toggle('active',id===name);
  });
  if(name==='match'){ const has=!!masterData; document.getElementById('match-no-master').style.display=has?'none':'block'; document.getElementById('match-main-area').style.display=has?'block':'none'; }
  if(name==='results'){ const has=matchResults.length>0; document.getElementById('results-empty').style.display=has?'none':'block'; document.getElementById('results-content').style.display=has?'block':'none'; if(has) renderResultsTable(); }
  if(name==='conso') renderConsoPage();
}
document.querySelectorAll('.nav-item[data-tab]').forEach(b=>b.addEventListener('click',()=>switchTab(b.dataset.tab)));

function showToast(msg,d=2800){ const t=document.getElementById('toast'); t.textContent=msg; t.classList.add('show'); setTimeout(()=>t.classList.remove('show'),d); }

/* ══════════════════════════════════════════════
   MASTER DATA
══════════════════════════════════════════════ */
function onMasterFileSelected(inp){
  const file=inp.files[0]; if(!file) return;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'binary'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
      if(!rows.length){ alert('No data found'); return; }
      masterCols=Object.keys(rows[0]); masterData=rows;
      masterMapping=autoDetect(masterCols,MASTER_FIELDS); window.masterMapping=masterMapping;
      buildMappingGrid('master-mapping-grid',MASTER_FIELDS,masterCols,masterMapping,'masterMapping');
      document.getElementById('master-upload-title').textContent=file.name;
      document.getElementById('master-dropzone').classList.add('has-file');
      document.getElementById('master-mapping-area').style.display='block';
      document.getElementById('master-ready-area').style.display='none';
    }catch(err){ alert('Error: '+err.message); }
  };
  reader.readAsBinaryString(file);
}

function precomputeMaster(){
  masterData.forEach(r=>{
    r._odprem  = r[masterMapping.odprem]  ??'';
    r._tpprem  = r[masterMapping.tpprem]  ??'';
    r._totprem = r[masterMapping.totprem] ??'';
    r._share   = r[masterMapping.sharepct]??'';
    r._polno   = String(r[masterMapping.polno]  ??'');
    r._endno   = String(r[masterMapping.endno]  ??'');
    r._cust    = String(r[masterMapping.custname]??'');
    r._start   = r[masterMapping.startdate]??'';
    r._end     = r[masterMapping.enddate]  ??'';
    r._endodate= r[masterMapping.endodate] ?? r[masterMapping.startdate] ??''; // Endo Date (AD in sheet)
    r._instno  = r[masterMapping.instno];
    r._isBase  = norm(r._instno)==='0'||r._instno===0||norm(r._instno)==='';
    r._cno     = String(r[masterMapping.cno]??'');
    r._brok    = r[masterMapping.brokdue]  ??'';
    r._dept    = String(r[masterMapping.dept]??'').toLowerCase();
    r._sumins  = r[masterMapping.sumins]   ??'';
    r._zeroPol = !r._polno || r._polno==='0';
    // Pre-compute fresh keys (matching Excel's freshening logic)
    r._freshPol  = freshPolno(r._polno);
    r._freshEndo = freshPolno(r._endno);
    r._freshCust = freshCustName(r._cust);
    r._cc        = String(r._cno) + String(r._instno??''); // CC = Cno+Inst concatenated
  });
  // Rebuild lookup maps
  masterMaps = null;
  buildMasterMaps();
}

function saveMasterMapping(){
  document.querySelectorAll('#master-mapping-grid select').forEach((s,i)=>{ if(s.value) masterMapping[MASTER_FIELDS[i].key]=s.value; });
  if(!masterMapping.cno){ alert('Please map C.No'); return; }
  if(!masterMapping.instno){ alert('Please map Inst No'); return; }
  precomputeMaster();
  try{ localStorage.setItem(SK_MASTER,JSON.stringify(masterData)); localStorage.setItem(SK_MASTER_M,JSON.stringify(masterMapping)); }catch(e){ console.warn('localStorage save failed'); }
  showMasterReady(); showToast('✅ Master data saved');
}

function showMasterReady(){
  document.getElementById('master-mapping-area').style.display='none';
  document.getElementById('master-upload-area').style.display='none';
  document.getElementById('master-ready-area').style.display='block';
  document.getElementById('master-saved-banner').style.display='none';
  const base=masterData.filter(r=>r._isBase).length;
  const skip=masterData.filter(r=>isZeroBrokDue(r._brok)).length;
  const zeroPol=masterData.filter(r=>r._zeroPol).length;
  document.getElementById('stat-total').textContent=masterData.length.toLocaleString();
  document.getElementById('stat-base').textContent=base.toLocaleString();
  document.getElementById('stat-endo').textContent=(masterData.length-base).toLocaleString();
  document.getElementById('stat-skip').textContent=skip.toLocaleString();
  document.getElementById('stat-zeropol').textContent=zeroPol.toLocaleString();
  const badge=document.getElementById('nav-master-badge'); badge.textContent=masterData.length.toLocaleString(); badge.style.display='inline-block';
  document.getElementById('storage-status').innerHTML='<span class="sdot green"></span> Master saved ('+masterData.length+' records)';
  renderPreviewTable('master-preview-table',masterData.slice(0,6),masterCols.slice(0,8));
}

function resetMaster(){
  if(!confirm('Remove master data from browser?')) return;
  masterData=null; masterCols=[]; masterMapping={}; window.masterMapping={};
  localStorage.removeItem(SK_MASTER); localStorage.removeItem(SK_MASTER_M);
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

/* ══════════════════════════════════════════════
   DEFAULT RULE — LOAD / SAVE / RENDER
══════════════════════════════════════════════ */
function loadDefaultRule(){
  try{
    const saved=localStorage.getItem(SK_DEFRULE);
    defaultRule = saved ? {...FACTORY_DEFAULT,...JSON.parse(saved)} : {...FACTORY_DEFAULT};
  }catch{ defaultRule={...FACTORY_DEFAULT}; }
}

function saveDefaultRuleToStorage(){
  try{ localStorage.setItem(SK_DEFRULE,JSON.stringify(defaultRule)); }catch{}
}

function resetDefaultRule(){
  if(!confirm('Reset default rule to factory settings?\nThis will undo all your changes to the default rule.')) return;
  defaultRule={...FACTORY_DEFAULT};
  saveDefaultRuleToStorage();
  renderDefaultRuleForm();
  showToast('🔄 Default rule reset to factory settings');
}

function saveDefaultRuleFromForm(){
  const g=(id,type='str')=>{
    const el=document.getElementById(id); if(!el) return undefined;
    if(type==='num') return parseFloat(el.value)||0;
    if(type==='int') return parseInt(el.value)||0;
    if(type==='bool') return el.value==='yes';
    if(type==='check') return el.checked;
    return el.value;
  };

  defaultRule = {
    polnoMode:         g('dr-polno-mode'),
    polnoTransform:    g('dr-polno-transform'),
    suffixN:           g('dr-suffix-n','int'),
    polnoWeight:       g('dr-polno-weight','int'),
    endnoWeight:       g('dr-endno-weight','int'),
    instnoMode:        g('dr-instno-mode'),
    instnoRequired:    g('dr-instno-req','bool'),
    nameStrip:         g('dr-namestrip','bool'),
    fuzzyPct:          g('dr-fuzzy','int'),
    nameWeight:        g('dr-name-weight','int'),
    premTolPct:        g('dr-prem-tol-pct','num'),
    premTolAbs:        g('dr-prem-tol-abs','num'),
    useShare:          g('dr-share','bool'),
    premWeight:        g('dr-prem-weight','int'),
    dateCheck:         g('dr-date'),
    dateWeight:        g('dr-date-weight','int'),
    sumInsCheck:       g('dr-sumins-check','bool'),
    sumInsTolPct:      g('dr-sumins-tol','num'),
    zeroPolicyFallback:g('dr-zeropol','bool'),
    skipZeroBrok:      g('dr-skip-zerobrok','bool'),
    skipBaseForEndo:   g('dr-skip-base-endo','bool'),
    skipEndoForBase:   g('dr-skip-endo-base','bool'),
    deptFilter:        g('dr-dept-filter'),
    scoreFullMatch:    g('dr-score-full','int'),
    scorePartial:      g('dr-score-partial','int'),
    scorePolCustPrem:  g('dr-sc-pol-cust-prem','int'),
    scorePolPrem:      g('dr-sc-pol-prem','int'),
    scorePolCust:      g('dr-sc-pol-cust','int'),
    scoreEndoPrem:     g('dr-sc-endo-prem','int'),
    scoreEndoCust:     g('dr-sc-endo-cust','int'),
    scoreEndoOnly:     g('dr-sc-endo-only','int'),
    scorePolOnly:      g('dr-sc-pol-only','int'),
    scoreCustPrem:     g('dr-sc-cust-prem','int'),
    scoreCustOnly:     g('dr-sc-cust-only','int'),
    scorePremOnly:     g('dr-sc-prem-only','int'),
    boostExactPolno:   g('dr-boost-exact-pol','int'),
    boostStrongName:   g('dr-boost-name','int'),
    boostExactPrem:    g('dr-boost-prem','int'),
    boostYearMatch:    g('dr-boost-year','int'),
    notes:             g('dr-notes'),
  };
  saveDefaultRuleToStorage();
  showToast('✅ Default rule saved');
}

function renderDefaultRuleForm(){
  if(!defaultRule) return;
  const s=(id,val)=>{ const el=document.getElementById(id); if(el) el.value=String(val??''); };
  const b=(id,val)=>{ const el=document.getElementById(id); if(el) el.value=val?'yes':'no'; };

  s('dr-polno-mode',         defaultRule.polnoMode);
  s('dr-polno-transform',    defaultRule.polnoTransform);
  s('dr-suffix-n',           defaultRule.suffixN);
  s('dr-polno-weight',       defaultRule.polnoWeight);
  s('dr-endno-weight',       defaultRule.endnoWeight);
  s('dr-instno-mode',        defaultRule.instnoMode);
  b('dr-instno-req',         defaultRule.instnoRequired);
  b('dr-namestrip',          defaultRule.nameStrip);
  s('dr-fuzzy',              defaultRule.fuzzyPct);
  s('dr-name-weight',        defaultRule.nameWeight);
  s('dr-prem-tol-pct',       defaultRule.premTolPct);
  s('dr-prem-tol-abs',       defaultRule.premTolAbs);
  b('dr-share',              defaultRule.useShare);
  s('dr-prem-weight',        defaultRule.premWeight);
  s('dr-date',               defaultRule.dateCheck);
  s('dr-date-weight',        defaultRule.dateWeight);
  b('dr-sumins-check',       defaultRule.sumInsCheck);
  s('dr-sumins-tol',         defaultRule.sumInsTolPct);
  b('dr-zeropol',            defaultRule.zeroPolicyFallback);
  b('dr-skip-zerobrok',      defaultRule.skipZeroBrok);
  b('dr-skip-base-endo',     defaultRule.skipBaseForEndo);
  b('dr-skip-endo-base',     defaultRule.skipEndoForBase);
  s('dr-dept-filter',        defaultRule.deptFilter);
  s('dr-score-full',         defaultRule.scoreFullMatch);
  s('dr-score-partial',      defaultRule.scorePartial);
  s('dr-sc-pol-cust-prem',   defaultRule.scorePolCustPrem);
  s('dr-sc-pol-prem',        defaultRule.scorePolPrem);
  s('dr-sc-pol-cust',        defaultRule.scorePolCust);
  s('dr-sc-endo-prem',       defaultRule.scoreEndoPrem);
  s('dr-sc-endo-cust',       defaultRule.scoreEndoCust);
  s('dr-sc-endo-only',       defaultRule.scoreEndoOnly);
  s('dr-sc-pol-only',        defaultRule.scorePolOnly);
  s('dr-sc-cust-prem',       defaultRule.scoreCustPrem);
  s('dr-sc-cust-only',       defaultRule.scoreCustOnly);
  s('dr-sc-prem-only',       defaultRule.scorePremOnly);
  s('dr-boost-exact-pol',    defaultRule.boostExactPolno);
  s('dr-boost-name',         defaultRule.boostStrongName);
  s('dr-boost-prem',         defaultRule.boostExactPrem);
  s('dr-boost-year',         defaultRule.boostYearMatch);
  s('dr-notes',              defaultRule.notes||'');

  // show/hide suffix field
  const pm=document.getElementById('dr-polno-mode');
  if(pm){ const sw=document.getElementById('dr-suffix-wrap'); if(sw) sw.style.display=pm.value==='suffix'?'block':'none'; }
}

/* ══════════════════════════════════════════════
   PER-INSURER RULES
══════════════════════════════════════════════ */
function loadRules(){
  try{ const r=localStorage.getItem(SK_RULES); if(r) insurerRules=JSON.parse(r); }catch{}
  loadDefaultRule();
  renderRulesTable();
  renderDefaultRuleForm();
}

function saveRules(){ try{ localStorage.setItem(SK_RULES,JSON.stringify(insurerRules)); }catch{} }

function getEffectiveRule(insurerName){
  // Merge insurer-specific overrides onto the default rule
  if(insurerName){
    const n=norm(insurerName);
    const override=insurerRules.find(r=>n.includes(norm(r.insurer))||norm(r.insurer).includes(n));
    if(override) return {...defaultRule,...override};
  }
  return {...defaultRule};
}

function addInsurerRule(){
  const insurer=document.getElementById('r-insurer').value.trim();
  if(!insurer){ alert('Insurer name required'); return; }

  const g=(id,type='str')=>{
    const el=document.getElementById(id); if(!el) return undefined;
    if(type==='num') return parseFloat(el.value);
    if(type==='int') return parseInt(el.value);
    if(type==='bool') return el.value==='yes';
    return el.value||undefined;
  };

  const rule={
    insurer,
    polnoMode:       g('r-polno-mode'),
    polnoTransform:  g('r-polno-transform'),
    suffixN:         g('r-suffix-n','int'),
    polnoWeight:     g('r-polno-weight','int'),
    endnoWeight:     g('r-endno-weight','int'),
    instnoMode:      g('r-instno-mode'),
    instnoRequired:  g('r-instno-req','bool'),
    nameStrip:       g('r-namestrip','bool'),
    fuzzyPct:        g('r-fuzzy','int'),
    nameWeight:      g('r-name-weight','int'),
    premTolPct:      g('r-prem-tol-pct','num'),
    premTolAbs:      g('r-prem-tol-abs','num'),
    useShare:        g('r-share','bool'),
    premWeight:      g('r-prem-weight','int'),
    dateCheck:       g('r-date'),
    dateWeight:      g('r-date-weight','int'),
    zeroPolicyFallback: g('r-zeropol','bool'),
    skipZeroBrok:    g('r-skip-zerobrok','bool'),
    skipBaseForEndo: g('r-skip-base-endo','bool'),
    skipEndoForBase: g('r-skip-endo-base','bool'),
    deptFilter:      g('r-dept-filter'),
    scoreFullMatch:  g('r-score-full','int'),
    scorePartial:    g('r-score-partial','int'),
    scorePolCustPrem:g('r-sc-pol-cust-prem','int'),
    scorePolPrem:    g('r-sc-pol-prem','int'),
    scorePolCust:    g('r-sc-pol-cust','int'),
    scoreEndoPrem:   g('r-sc-endo-prem','int'),
    scoreEndoCust:   g('r-sc-endo-cust','int'),
    scoreEndoOnly:   g('r-sc-endo-only','int'),
    scorePolOnly:    g('r-sc-pol-only','int'),
    scoreCustPrem:   g('r-sc-cust-prem','int'),
    scoreCustOnly:   g('r-sc-cust-only','int'),
    scorePremOnly:   g('r-sc-prem-only','int'),
    boostExactPolno: g('r-boost-exact-pol','int'),
    boostStrongName: g('r-boost-name','int'),
    boostExactPrem:  g('r-boost-prem','int'),
    boostYearMatch:  g('r-boost-year','int'),
    notes:           g('r-notes'),
  };

  // Remove undefined fields so defaults still apply
  Object.keys(rule).forEach(k=>{ if(rule[k]===undefined||rule[k]===''||rule[k]===null) delete rule[k]; });

  const idx=insurerRules.findIndex(r=>norm(r.insurer)===norm(insurer));
  if(idx>=0) insurerRules[idx]=rule; else insurerRules.push(rule);
  saveRules(); renderRulesTable();
  document.getElementById('r-insurer').value='';
  showToast('✅ Rule saved for '+insurer);
}

function deleteRule(i){ insurerRules.splice(i,1); saveRules(); renderRulesTable(); showToast('Rule deleted'); }

function editRule(i){
  const r=insurerRules[i];
  const s=(id,val)=>{ const el=document.getElementById(id); if(el&&val!==undefined) el.value=String(val); };
  const b=(id,val)=>{ const el=document.getElementById(id); if(el&&val!==undefined) el.value=val?'yes':'no'; };
  s('r-insurer', r.insurer);
  s('r-polno-mode',r.polnoMode); s('r-polno-transform',r.polnoTransform);
  s('r-suffix-n',r.suffixN); s('r-polno-weight',r.polnoWeight); s('r-endno-weight',r.endnoWeight);
  s('r-instno-mode',r.instnoMode); b('r-instno-req',r.instnoRequired);
  b('r-namestrip',r.nameStrip); s('r-fuzzy',r.fuzzyPct); s('r-name-weight',r.nameWeight);
  s('r-prem-tol-pct',r.premTolPct); s('r-prem-tol-abs',r.premTolAbs);
  b('r-share',r.useShare); s('r-prem-weight',r.premWeight);
  s('r-date',r.dateCheck); s('r-date-weight',r.dateWeight);
  b('r-zeropol',r.zeroPolicyFallback);
  b('r-skip-zerobrok',r.skipZeroBrok); b('r-skip-base-endo',r.skipBaseForEndo); b('r-skip-endo-base',r.skipEndoForBase);
  s('r-dept-filter',r.deptFilter);
  s('r-score-full',r.scoreFullMatch); s('r-score-partial',r.scorePartial);
  s('r-sc-pol-cust-prem',r.scorePolCustPrem); s('r-sc-pol-prem',r.scorePolPrem);
  s('r-sc-pol-cust',r.scorePolCust); s('r-sc-endo-prem',r.scoreEndoPrem);
  s('r-sc-endo-cust',r.scoreEndoCust); s('r-sc-endo-only',r.scoreEndoOnly);
  s('r-sc-pol-only',r.scorePolOnly); s('r-sc-cust-prem',r.scoreCustPrem);
  s('r-sc-cust-only',r.scoreCustOnly); s('r-sc-prem-only',r.scorePremOnly);
  s('r-boost-exact-pol',r.boostExactPolno); s('r-boost-name',r.boostStrongName);
  s('r-boost-prem',r.boostExactPrem); s('r-boost-year',r.boostYearMatch);
  s('r-notes',r.notes);
  document.getElementById('r-insurer').scrollIntoView({behavior:'smooth'});
}

function renderRulesTable(){
  const empty=document.getElementById('rules-empty');
  const wrap=document.getElementById('rules-table-wrap');
  const hint=document.getElementById('rules-count-hint');
  const badge=document.getElementById('nav-rules-badge');
  if(!insurerRules.length){ empty.style.display='block'; wrap.style.display='none'; hint.textContent='— none yet'; badge.style.display='none'; return; }
  empty.style.display='none'; wrap.style.display='block';
  hint.textContent=`— ${insurerRules.length} rule${insurerRules.length>1?'s':''} saved`;
  badge.textContent=insurerRules.length; badge.style.display='inline-block';
  document.getElementById('rules-tbody').innerHTML=insurerRules.map((r,i)=>`
    <tr>
      <td><strong>${esc(r.insurer)}</strong></td>
      <td>${esc(r.polnoMode||'—')}</td>
      <td>${esc(r.polnoTransform||'—')}</td>
      <td>${r.premTolPct??'—'}% / ₹${r.premTolAbs??'—'}</td>
      <td>${r.dateCheck||'—'}</td>
      <td>${r.instnoMode||'—'}</td>
      <td>${r.useShare===false?'No':r.useShare===true?'Yes':'—'}</td>
      <td>${r.fuzzyPct??'—'}%</td>
      <td>${r.scoreFullMatch??'—'} / ${r.scorePartial??'—'}</td>
      <td style="color:var(--text3);font-size:11px">${esc(r.notes||'—')}</td>
      <td>
        <button class="btn btn-ghost btn-sm" onclick="editRule(${i})">✎</button>
        <button class="btn btn-ghost btn-sm" style="color:var(--red)" onclick="deleteRule(${i})">✕</button>
      </td>
    </tr>`).join('');
}

/* ══════════════════════════════════════════════
   STATEMENT LOAD
══════════════════════════════════════════════ */
function onStmtFileSelected(inp){
  const file=inp.files[0]; if(!file) return;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'binary'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
      if(!rows.length){ alert('No data found'); return; }
      stmtCols=Object.keys(rows[0]); stmtData=rows;
      stmtMapping=autoDetect(stmtCols,STMT_FIELDS); window.stmtMapping=stmtMapping;
      buildMappingGrid('stmt-mapping-grid',STMT_FIELDS,stmtCols,stmtMapping,'stmtMapping');
      document.getElementById('stmt-upload-title').textContent=file.name;
      document.getElementById('stmt-dropzone').classList.add('has-file');
      document.getElementById('stmt-mapping-area').style.display='block';
    }catch(err){ alert('Error: '+err.message); }
  };
  reader.readAsBinaryString(file);
}

/* ══════════════════════════════════════════════
   EXCEL-FAITHFUL MATCHING ENGINE v5
   Ported from Auto_Statement_Receiving.xlsx logic

   MATCHING WATERFALL (mirrors Excel IFERROR chain):
   PRIMARY (Cno1):
     1a. StmtFreshEndoNo    → Dump.FreshEndoNo           exact endo match
     1b. StmtPolNo+PBST     → Dump.pol&pbst              polno+premium combo
     1c. StmtFreshPolNo     → Dump.FreshPolNo             fresh polno
     1d. StmtPolNo+Date     → Dump.PolNo&StartDate + PBST tolerance ±5%
   FALLBACK (Cno2 when Primary blank):
     2a. StmtPolNo+PBST     → Dump.PolNo+OD              polno+OD
     2b. StmtPolNo+PBST     → Dump.PolNo+TP              polno+TP
     2c. StmtCustName+PBST  → Dump.CustName+PBST         name+premium
     2d. StmtPolNo+Date     → Dump.PolNo&StartDate        polno+date
     2e. StmtCustName+Date  → Dump.CustName+Date          name+date
     2f. StmtFreshPolNo     → Dump.PolNo (col E)          polno any
     2g. StmtFreshCustName  → Dump.FreshCustName          name only (last resort)

   SCORING (Excel AU): (PolNo+EndoNo)/200*70 + (Start+EndoDate+End)/300*15 + PBST/100*15
   REMARKS: Zero Commission | Contra | Matched | Partial | ENF
══════════════════════════════════════════════ */

/* ── Excel date serial: days since 1899-12-30 ── */
function toExcelSerial(v){
  if(!v && v!==0) return null;
  if(typeof v === 'number' && v > 10000) return Math.round(v); // already serial
  let d;
  if(v instanceof Date) d = v;
  else {
    const s = String(v).trim();
    const m1 = s.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    const m2 = s.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
    if(m1)      d = new Date(Date.UTC(+m1[1], +m1[2]-1, +m1[3]));
    else if(m2) d = new Date(Date.UTC(+m2[3], +m2[2]-1, +m2[1]));
    else        d = new Date(s);
  }
  if(d instanceof Date && !isNaN(d.getTime())){
    return Math.floor(d.getTime() / 86400000 + 25569);
  }
  return null;
}

/* ── FreshPolNo = LEFT(SUBSTITUTE(polno,":",""), 21) ── */
function freshPolno(raw){
  if(!raw) return '';
  return String(raw).replace(/:/g,'').substring(0,21).trim();
}

/* ── FreshCustName: UPPER, strip titles & special chars, collapse spaces ── */
const CUST_STRIP = ['MR.','MRS.','MS.','MISS.','DR.','SHRI.','SMT.','KUMAR.',
  'MR','MRS','MS','MISS','DR','SHRI','SMT','KUMAR','PVT.','PVT','PRIVATE',
  'LIMITED','LTD.','LTD'];

function freshCustName(raw){
  if(!raw) return '';
  let s = String(raw).toUpperCase();
  CUST_STRIP.forEach(t => { s = s.split(t).join(''); });
  s = s.replace(/[^A-Z0-9\s&]/g,'');  // remove . / * etc
  return s.replace(/\s+/g,' ').trim();
}

/* ── Build all composite keys for a statement row ── */
function buildStmtKeys(sRow){
  const rawPol = String(sRow[stmtMapping.polno]  ||'');
  const rawEnd = String(sRow[stmtMapping.endno]  ||'');
  const rawCst = String(sRow[stmtMapping.custname]||'');
  const pbst   = toNum(sRow[stmtMapping.totprem]) ?? toNum(sRow[stmtMapping.odprem]) ?? 0;
  const od     = toNum(sRow[stmtMapping.odprem])  ?? 0;
  const tp     = toNum(sRow[stmtMapping.tpprem])  ?? 0;
  const startV = sRow[stmtMapping.startdate];

  const fp  = freshPolno(rawPol);
  const fe  = freshPolno(rawEnd);   // same LEFT-21 logic
  const fc  = freshCustName(rawCst);
  const ds  = toExcelSerial(startV);
  const dsStr = ds !== null ? String(ds) : '';

  return {
    rawPol, rawEnd, rawCst, pbst, od, tp, startV,
    fp, fe, fc, ds,
    polPbst:  fp + String(Math.round(pbst)),   // FreshPolNo + ROUND(PBST,0)
    polDate:  fp + dsStr,                       // FreshPolNo + ExcelDateSerial
    custPbst: fc + String(Math.round(pbst)),   // FreshCustName + ROUND(PBST,0)
    custDate: fc + dsStr,                       // FreshCustName + ExcelDateSerial
    polOD:    fp + String(Math.round(od)),      // FreshPolNo + ROUND(OD,0)
    polTP:    fp + String(Math.round(tp)),      // FreshPolNo + ROUND(TP,0)
  };
}

/* ── Pre-compute lookup maps from masterData (called after upload) ── */
let masterMaps = null;
function buildMasterMaps(){
  masterMaps = {
    byFreshEndo: new Map(),
    byPolPbst:   new Map(),
    byFreshPol:  new Map(),
    byPolDate:   new Map(),
    byPolOD:     new Map(),
    byPolTP:     new Map(),
    byCustPbst:  new Map(),
    byCustDate:  new Map(),
    byFreshCust: new Map(),
    byCC:        new Map(),
  };

  const add = (map, key, row) => {
    if(key === null || key === undefined) return;
    const k = String(key).trim();
    if(!k) return;
    if(!map.has(k)) map.set(k, []);
    map.get(k).push(row);
  };

  masterData.forEach(m => {
    const fp  = m._freshPol;
    const fe  = m._freshEndo;
    const fc  = m._freshCust;
    const od  = toNum(m._odprem)  ?? 0;
    const tp  = toNum(m._tpprem)  ?? 0;
    const pbt = toNum(m._totprem) ?? 0;
    const ds  = toExcelSerial(m._start);
    const dsStr = ds !== null ? String(ds) : '';

    add(masterMaps.byFreshEndo,  fe,                          m);
    add(masterMaps.byPolPbst,    fp + String(Math.round(pbt)),m);
    add(masterMaps.byFreshPol,   fp,                          m);
    if(dsStr){ add(masterMaps.byPolDate,  fp + dsStr,         m); }
    add(masterMaps.byPolOD,      fp + String(Math.round(od)), m);
    add(masterMaps.byPolTP,      fp + String(Math.round(tp)), m);
    add(masterMaps.byCustPbst,   fc + String(Math.round(pbt)),m);
    add(masterMaps.byCustPbst,   fc + String(Math.round(od)), m);
    add(masterMaps.byCustPbst,   fc + String(Math.round(tp)), m);
    if(dsStr){ add(masterMaps.byCustDate, fc + dsStr,         m); }
    add(masterMaps.byFreshCust,  fc,                          m);
    if(m._cc) add(masterMaps.byCC, m._cc,                     m);
  });
}

/* ── Lookup helper ── */
function lookup(map, key, filterFn){
  if(!key && key!==0) return null;
  const rows = map.get(String(key).trim());
  if(!rows || rows.length===0) return null;
  return filterFn ? (rows.find(filterFn)||null) : rows[0];
}

/* ── PBST Score: 8-way premium check (Excel AT formula) ── */
function pbstScore(m, sk, rule){
  const sPbst = sk.pbst;
  if(!sPbst && sPbst!==0) return 0;
  const tol = rule.premTolPct ?? 2;

  const within = (a, b) => {
    if(a===null||b===null) return false;
    if(a===0&&b===0) return true;
    if(a===0||b===0) return false;
    return Math.abs((a-b)/b*100) <= tol;
  };

  const share  = toNum(m._share);
  const mPBST  = toNum(m._totprem) ?? 0;
  const mOD    = toNum(m._odprem)  ?? 0;
  const mTP    = toNum(m._tpprem)  ?? 0;
  const safe   = (share&&share>0&&share<100) ? share : 100;

  // Year diff for multi-year TP: OD + yearDiff * TP
  const sy = extractYear(m._start), ey = extractYear(m._end);
  const ydiff = (sy&&ey) ? Math.max(0,ey-sy) : 0;

  // Group sum (SUMIFS same CC group)
  let grpPbst = null;
  if(m._cc && masterMaps?.byCC?.has(m._cc)){
    grpPbst = masterMaps.byCC.get(m._cc).reduce((s,r)=>s+(toNum(r._totprem)||0), 0);
  }

  return (
    within(mPBST,              sPbst) ||  // a) PBST vs PBST
    within(mOD,                sPbst) ||  // b) OD vs PBST
    within(mTP,                sPbst) ||  // c) TP vs PBST
    within(mOD + ydiff*mTP,    sPbst) ||  // d) multi-year OD+yearDiff*TP
    (grpPbst!==null && within(grpPbst, mPBST)) || // e) group sum
    within(mTP/(safe/100),     sPbst) ||  // f) share-adj TP
    within(mOD/(safe/100),     sPbst) ||  // g) share-adj OD
    within(mPBST/(safe/100),   sPbst)     // h) share-adj PBST
  ) ? 100 : 0;
}

/* ── Date scores (Excel AQ/AR/AS) ── */
function calcDateScores(m, sk, stmtEndDate){
  const instIsBase = (String(m._instno||'')==='0'||m._instno===0||String(m._instno||'')==='');
  const eq = (a, b) => { const sa=toExcelSerial(a), sb=toExcelSerial(b); return sa!==null&&sb!==null&&sa===sb; };

  // AQ: if base → check StmtStart==MasterStart; if endo → always 100
  const startScore = instIsBase ? (eq(m._start, sk.startV) ? 100 : 0) : 100;

  // AR: if endo → check StmtStart==MasterEndoDate; if base → always 100
  const endoDateScore = (!instIsBase) ? (eq(m._endodate, sk.startV) ? 100 : 0) : 100;

  // AS: StmtEnd == MasterEnd
  const endDateScore = eq(m._end, stmtEndDate) ? 100 : 0;

  return { startScore, endoDateScore, endDateScore };
}

/* ── Main total score (Excel AU) ── */
function calcScore(pol, endo, start, endoDate, endDate, pbst){
  return ((pol+endo)/200*70) + ((start+endoDate+endDate)/300*15) + (pbst/100*15);
}

/* ── Core matching function ── */
function matchOneRow(sRow){
  if(!masterMaps) buildMasterMaps();

  const sIns  = String(sRow[stmtMapping.insurer]||'');
  const rule  = getEffectiveRule(sIns);
  const sComm = toNum(sRow[stmtMapping.commission]);
  const sEnd  = sRow[stmtMapping.enddate];
  const sk    = buildStmtKeys(sRow);
  const notZ  = m => !isZeroBrokDue(m._brok);

  // ── 1. Zero Commission ──
  if(sComm !== null && Math.abs(sComm) <= 1){
    return mk(sRow, null, 0, 'zero-commission', 'Zero Commission', sIns);
  }

  // ── 2. Contra: same polno+startDate group sums to ~0 ──
  if(sk.fp && sk.ds !== null && stmtData){
    const grp = stmtData.filter(r => {
      return freshPolno(String(r[stmtMapping.polno]||''))===sk.fp &&
             toExcelSerial(r[stmtMapping.startdate])===sk.ds;
    });
    if(grp.length > 1){
      const sumC = grp.reduce((s,r)=>(s+(toNum(r[stmtMapping.commission])||0)),0);
      if(Math.abs(sumC)<0.01) return mk(sRow, null, 0, 'contra', 'Contra', sIns);
    }
  }

  // ── PRIMARY MATCH (Cno1 chain) ──
  let best =
    lookup(masterMaps.byFreshEndo, sk.fe,      notZ)                          || // 1a
    lookup(masterMaps.byPolPbst,   sk.polPbst,  notZ)                         || // 1b
    lookup(masterMaps.byFreshPol,  sk.fp,       notZ)                         || // 1c
    (() => {                                                                       // 1d polno+date with ±5% pbst
      const cands = masterMaps.byPolDate.get(sk.polDate);
      if(!cands) return null;
      return cands.find(m => notZ(m) && (() => {
        const mp = toNum(m._totprem); if(!mp||!sk.pbst) return true;
        return Math.abs((mp-sk.pbst)/sk.pbst*100) <= 5;
      })()) || null;
    })();

  // ── FALLBACK MATCH (Cno2, only if primary blank) ──
  if(!best){
    best =
      lookup(masterMaps.byPolOD,    sk.polOD,    notZ) ||  // 2a polno+OD
      lookup(masterMaps.byPolTP,    sk.polTP,    notZ) ||  // 2b polno+TP
      lookup(masterMaps.byCustPbst, sk.custPbst, notZ) ||  // 2c name+pbst
      lookup(masterMaps.byPolDate,  sk.polDate,  notZ) ||  // 2d polno+date
      lookup(masterMaps.byCustDate, sk.custDate, notZ) ||  // 2e name+date
      lookup(masterMaps.byFreshPol, sk.fp,       notZ) ||  // 2f polno
      lookup(masterMaps.byFreshCust,sk.fc,       notZ);    // 2g name only
  }

  if(!best) return mk(sRow, null, 5, 'unmatched', 'ENF', sIns);

  // ── SCORING ──
  // AO: PolNo Score — if master polno is '0' or blank → 100; if matches stmt → 100
  const polScore = (!best._freshPol||best._freshPol==='0')
    ? 100 : (best._freshPol===sk.fp ? 100 : 0);

  // AP: EndoNo Score — if master endno '0'/blank → 100; if matches → 100
  const fe2 = best._freshEndo;
  const endoScore = (!fe2||fe2==='0'||fe2==='main policy')
    ? 100 : (fe2===sk.fe ? 100 : 0);

  // AQ/AR/AS: Date scores
  const ds = calcDateScores(best, sk, sEnd);

  // AT: PBST score
  const pbst100 = pbstScore(best, sk, rule);

  // AU: Total
  const total = calcScore(polScore, endoScore, ds.startScore, ds.endoDateScore, ds.endDateScore, pbst100);

  // ── REMARKS (AW) ──
  let status, remark;
  if((total>84 && pbst100===100) || ((polScore===100||endoScore===100) && pbst100===100)){
    status='100'; remark='Matched';
  } else if(total>50){
    status='partial'; remark='Partial('+Math.round(total)+')';
  } else {
    status='unmatched'; remark='ENF('+Math.round(total)+')';
  }

  const on=[];
  if(polScore===100)      on.push('Pol');
  if(endoScore===100)     on.push('Endo');
  if(pbst100===100)       on.push('PBST');
  if(ds.startScore===100) on.push('Start');
  if(ds.endoDateScore===100) on.push('EndoDate');
  if(ds.endDateScore===100)  on.push('End');

  return {
    _stmtRow:sRow,
    _cno:         String(best._cno||''),
    _instno:      String(best._instno??''),
    _score:       Math.round(total*10)/10,
    _polScore:    polScore,
    _endoScore:   endoScore,
    _pbstScore:   pbst100,
    _matchStatus: status,
    _matchedOn:   on.join('+') || remark,
    _remark:      remark,
    _insurer:     sIns,
    _ruleKey:     sIns ? norm(sIns).slice(0,20) : 'default',
  };
}

function mk(sRow, best, score, status, remark, sIns){
  return {_stmtRow:sRow,_cno:best?String(best._cno||''):'',_instno:best?String(best._instno??''):'',
    _score:score,_polScore:0,_endoScore:0,_pbstScore:0,
    _matchStatus:status,_matchedOn:remark,_remark:remark,_insurer:sIns,_ruleKey:'default'};
}

function extractYear(v){
  if(!v&&v!==0) return null;
  if(typeof v==='number'&&v>10000){ try{ const d=new Date(Math.round((v-25569)*86400*1000));const y=d.getUTCFullYear();return(y>1900&&y<2100)?y:null; }catch{return null;} }
  const s=String(v);
  const m1=s.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/); if(m1) return parseInt(m1[3]);
  const m2=s.match(/(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/); if(m2) return parseInt(m2[1]);
  const m3=s.match(/\b(20\d{2}|19\d{2})\b/); if(m3) return parseInt(m3[1]);
  return null;
}

async function runMatching(){
  if(!masterData||!stmtData){ alert('Upload both files first'); return; }
  document.querySelectorAll('#stmt-mapping-grid select').forEach((s,i)=>{ if(s.value) stmtMapping[STMT_FIELDS[i].key]=s.value; });
  document.getElementById('stmt-mapping-area').style.display='none';
  document.getElementById('matching-progress').style.display='block';

  const log=document.getElementById('match-log');
  const fill=document.getElementById('prog-fill');
  const lbl=document.getElementById('prog-label');
  const addLog=(msg,cls='')=>{ const sp=document.createElement('span'); sp.className=cls; sp.textContent=msg+'\n'; log.appendChild(sp); log.scrollTop=log.scrollHeight; };

  // Reset & rebuild lookup maps
  masterMaps = null;
  buildMasterMaps();

  matchResults=[];
  const total=stmtData.length;
  addLog(`Processing ${total} statement rows…`,'log-ok');
  addLog(`Master: ${masterData.length} records in ${masterMaps.byFreshPol.size} unique policy nos`,'log-ok');
  addLog(`Lookup maps built — 9 composite-key indexes ready`,'log-ok');
  addLog(`Engine: Excel-faithful waterfall (1a→1b→1c→1d → 2a→2b→2c→2d→2e→2f→2g)`,'log-ok');
  if(insurerRules.length) addLog(`Custom insurer rules: ${insurerRules.map(r=>r.insurer).join(', ')}`,'log-ok');

  let cMatch=0, cPartial=0, cENF=0, cZero=0, cContra=0;
  for(let i=0;i<total;i+=10){
    const batch=stmtData.slice(i,Math.min(i+10,total));
    fill.style.width=Math.round(i/total*100)+'%';
    lbl.textContent=`Row ${i+1}–${Math.min(i+10,total)} of ${total}`;
    for(const row of batch){
      const r=matchOneRow(row); matchResults.push(r);
      if(r._matchStatus==='100')            cMatch++;
      else if(r._matchStatus==='partial')   cPartial++;
      else if(r._matchStatus==='zero-commission') cZero++;
      else if(r._matchStatus==='contra')    cContra++;
      else                                  cENF++;
    }
    await new Promise(res=>setTimeout(res,4));
  }
  fill.style.width='100%'; lbl.textContent='Done!';
  addLog(`✅ Matched: ${cMatch} | Partial: ${cPartial} | ENF: ${cENF} | Zero Comm: ${cZero} | Contra: ${cContra}`,'log-ok');
  if(cENF>0) addLog(`⚠ ${cENF} rows need manual review (ENF)`,'log-warn');

  const badge=document.getElementById('nav-results-badge');
  badge.textContent=matchResults.length; badge.style.display='inline-block';
  setTimeout(()=>{ document.getElementById('matching-progress').style.display='none'; switchTab('results'); },900);
}

/* ══════════════════════════════════════════════
   RESULTS TABLE
══════════════════════════════════════════════ */
function renderResultsTable(){
  const filter=document.getElementById('results-filter').value;
  const rows=matchResults.filter(r=>{
    if(filter==='all') return true;
    if(filter==='100') return r._matchStatus==='100';
    if(filter==='partial') return r._matchStatus==='partial';
    if(filter==='unmatched') return r._matchStatus==='unmatched';
    if(filter==='zero-commission') return r._matchStatus==='zero-commission';
    if(filter==='contra') return r._matchStatus==='contra';
    return true;
  });

  const c100=matchResults.filter(r=>r._matchStatus==='100').length;
  const cP=matchResults.filter(r=>r._matchStatus==='partial').length;
  const cU=matchResults.filter(r=>r._matchStatus==='unmatched').length;
  const cZ=matchResults.filter(r=>r._matchStatus==='zero-commission').length;
  const cC=matchResults.filter(r=>r._matchStatus==='contra').length;

  document.getElementById('r-total').textContent=matchResults.length;
  document.getElementById('r-100').textContent=c100;
  document.getElementById('r-partial').textContent=cP;
  document.getElementById('r-unmatched').textContent=cU;
  const rz=document.getElementById('r-zero'); if(rz) rz.textContent=cZ;
  const rc=document.getElementById('r-contra'); if(rc) rc.textContent=cC;

  const visCols=stmtCols.slice(0,5);
  const th=[...visCols,'C.No','Inst No','Score','PolScore','EndoScore','PBST','Matched On','Rule','Status'].map(c=>`<th>${esc(String(c))}</th>`).join('');
  const tb=rows.slice(0,400).map(r=>{
    const cells=visCols.map(c=>`<td>${esc(String(r._stmtRow[c]??''))}</td>`).join('');
    const st = r._matchStatus;
    const badge =
      st==='100'?'<span class="badge green">✓ Matched</span>':
      st==='partial'?'<span class="badge amber">Partial</span>':
      st==='zero-commission'?'<span class="badge" style="background:rgba(99,102,241,.15);color:#818cf8">Zero Comm</span>':
      st==='contra'?'<span class="badge" style="background:rgba(168,85,247,.15);color:#c084fc">Contra</span>':
      '<span class="badge red">ENF</span>';
    const sc = r._score;
    const scoreClass = sc>=85?'green':sc>=50?'':'amber';
    const scorePill=`<span class="score-pill ${scoreClass}">${sc}</span>`;
    const polPill=`<span class="score-pill ${r._polScore===100?'green':'amber'}">${r._polScore}</span>`;
    const endoPill=`<span class="score-pill ${r._endoScore===100?'green':'amber'}">${r._endoScore}</span>`;
    const pbstPill=`<span class="score-pill ${r._pbstScore===100?'green':'amber'}">${r._pbstScore}</span>`;
    const mo=r._matchedOn?`<span class="match-on-tag">${esc(r._matchedOn)}</span>`:'—';
    const rl=`<span class="rule-tag">${esc(r._ruleKey)}</span>`;
    return `<tr>${cells}<td><strong>${esc(r._cno)}</strong></td><td>${esc(r._instno)}</td><td>${scorePill}</td><td>${polPill}</td><td>${endoPill}</td><td>${pbstPill}</td><td>${mo}</td><td>${rl}</td><td>${badge}</td></tr>`;
  }).join('');
  document.getElementById('results-table').innerHTML=`<table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table>`;
}

function downloadResults(){
  if(!matchResults.length){ alert('No results'); return; }
  try{
    const wb=XLSX.utils.book_new();
    const build=arr=>arr.map(r=>{
      const obj={};
      stmtCols.forEach(c=>{ obj[c]=r._stmtRow[c]!==undefined?r._stmtRow[c]:''; });
      obj['C.No']=r._cno; obj['Inst No']=r._instno;
      obj['Match Score']=r._score;
      obj['Matched On']=r._matchedOn||'';
      obj['Rule Applied']=r._ruleKey||'default';
      obj['Match Status']=r._matchStatus==='100'?'100% Match':r._matchStatus==='partial'?'Partial Match':'Unmatched';
      return obj;
    });
    const add=(rows,name)=>{ XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows.length?rows:[{Note:'No data'}]),name); };
    add(build(matchResults),'All Results');
    add(build(matchResults.filter(r=>r._matchStatus==='100')),'100% Matched');
    add(build(matchResults.filter(r=>r._matchStatus==='partial')),'Partial Match');
    add(build(matchResults.filter(r=>r._matchStatus==='unmatched')),'Unmatched');
    XLSX.writeFile(wb,'commission_matched_'+new Date().toISOString().slice(0,10)+'.xlsx');
  }catch(err){ alert('Download error: '+err.message); }
}

/* ══════════════════════════════════════════════
   CONSOLIDATED
══════════════════════════════════════════════ */
function loadConso(){ try{ const c=localStorage.getItem(SK_CONSO); if(c) consoData=JSON.parse(c); }catch{} }
function saveConso(){ try{ localStorage.setItem(SK_CONSO,JSON.stringify(consoData)); }catch{ alert('Storage full'); } }

function addToConsolidated(){
  if(!matchResults.length){ alert('No results to add'); return; }
  const label=document.getElementById('conso-stmt-label').value.trim()||'Statement '+new Date().toLocaleDateString('en-IN');
  const flatRows=matchResults.map(r=>{
    const obj={};
    stmtCols.forEach(c=>{ obj[c]=r._stmtRow[c]!==undefined?r._stmtRow[c]:''; });
    obj['C.No']=r._cno; obj['Inst No']=r._instno;
    obj['Match Score']=r._score;
    obj['Matched On']=r._matchedOn||'';
    obj['Match Status']=r._matchStatus==='100'?'100% Match':r._matchStatus==='partial'?'Partial Match':'Unmatched';
    obj['_Statement']=label;
    return obj;
  });
  consoData.push({label,date:new Date().toISOString(),rows:flatRows});
  saveConso();
  const total=consoData.reduce((s,b)=>s+b.rows.length,0);
  const badge=document.getElementById('nav-conso-badge'); badge.textContent=total; badge.style.display='inline-block';
  document.getElementById('conso-stmt-label').value='';
  showToast(`✅ ${flatRows.length} rows added`);
  switchTab('conso');
}

function renderConsoPage(){
  const allRows=consoData.flatMap(b=>b.rows);
  const has=allRows.length>0;
  document.getElementById('conso-empty').style.display=has?'none':'block';
  document.getElementById('conso-content').style.display=has?'block':'none';
  if(!has) return;
  document.getElementById('c-total').textContent=allRows.length.toLocaleString();
  document.getElementById('c-batches').textContent=consoData.length;
  document.getElementById('c-100').textContent=allRows.filter(r=>r['Match Status']==='100% Match').length.toLocaleString();
  document.getElementById('c-review').textContent=allRows.filter(r=>r['Match Status']!=='100% Match').length.toLocaleString();
  const badge=document.getElementById('nav-conso-badge'); badge.textContent=allRows.length; badge.style.display='inline-block';
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
    if(filter==='100') return r['Match Status']==='100% Match';
    if(filter==='partial') return r['Match Status']==='Partial Match';
    if(filter==='unmatched') return r['Match Status']==='Unmatched';
    return true;
  });
  if(!rows.length){ document.getElementById('conso-table').innerHTML='<div style="padding:20px;text-align:center;color:var(--text2)">No rows</div>'; return; }
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

function removeBatch(i){ if(!confirm(`Remove "${consoData[i].label}"?`)) return; consoData.splice(i,1); saveConso(); renderConsoPage(); showToast('Batch removed'); }
function clearConso(){ if(!confirm('Clear ALL consolidated data?')) return; consoData=[]; saveConso(); document.getElementById('nav-conso-badge').style.display='none'; renderConsoPage(); showToast('Cleared'); }

function downloadConso(){
  const allRows=consoData.flatMap(b=>b.rows);
  if(!allRows.length){ alert('No data'); return; }
  try{
    const wb=XLSX.utils.book_new();
    const clean=rows=>rows.map(r=>{ const o={}; Object.keys(r).filter(k=>!k.startsWith('_')).forEach(k=>{ o[k]=r[k]; }); return o; });
    const add=(rows,name)=>{ XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows.length?rows:[{Note:'No data'}]),name); };
    add(clean(allRows),'All Consolidated');
    add(clean(allRows.filter(r=>r['Match Status']==='100% Match')),'100% Matched');
    add(clean(allRows.filter(r=>r['Match Status']==='Partial Match')),'Partial Match');
    add(clean(allRows.filter(r=>r['Match Status']==='Unmatched')),'Unmatched');
    consoData.forEach(b=>{ add(clean(b.rows),b.label.slice(0,28).replace(/[\\\/\?\*\[\]]/g,'_')); });
    XLSX.writeFile(wb,'consolidated_commission_'+new Date().toISOString().slice(0,10)+'.xlsx');
  }catch(err){ alert('Download error: '+err.message); }
}

/* ══════════════════════════════════════════════
   PREVIEW
══════════════════════════════════════════════ */
function renderPreviewTable(cid,rows,cols){
  const w=document.getElementById(cid);
  const th=cols.map(c=>`<th>${esc(String(c))}</th>`).join('');
  const tb=rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(String(r[c]??''))}</td>`).join('')}</tr>`).join('');
  w.innerHTML=`<table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table>`;
}

/* ══════════════════════════════════════════════
   SAMPLE TEMPLATES
══════════════════════════════════════════════ */
function downloadMasterTemplate(){
  try{
    const wb=XLSX.utils.book_new();
    const hdr=['C.No','Inst No','Policy No','Endorsement No','Customer Name','Insurer Name','Department','Policy Type','OD Premium','TP Premium','Total Premium','Share %','Sum Insured','Policy Start Date','Policy End Date','Brok Due'];
    const data=[hdr,
      ['CNT001',0,'100051203900','','Gaurav Fatesaria','ICICI LOMBARD',  'MOTOR','Package',24922,0,24922,100,500000,'16/03/2026','15/03/2029',5106],
      ['CNT001','E1','100051203900','100051203900-E1','Gaurav Fatesaria','ICICI LOMBARD','MOTOR','Endorsement',500,0,500,100,'','20/04/2026','15/03/2029',102],
      ['CNT002',0,'100048387000','','Kousik Roy',       'ICICI LOMBARD','MOTOR','Package',23256,0,23256,100,400000,'02/03/2026','01/03/2027',4765],
      ['CNT003',0,'MOT/2024/001234','','Rajesh Kumar',  'New India Assurance','MOTOR','Package',8500,2200,10700,100,500000,'01/04/2024','31/03/2025',1070],
      ['CNT004',0,'FIR/2024/005678','','M/s Priya Sharma & Co','Oriental Insurance','FIRE','IAR',0,0,25000,60,2000000,'01/07/2024','30/06/2025',2500],
      ['CNT005',0,'HLT/2024/009012','','Mehta Enterprises Pvt Ltd','Star Health','HEALTH','Group Mediclaim',0,0,45000,60,1000000,'01/04/2024','31/03/2025',4500],
    ];
    const ws=XLSX.utils.aoa_to_sheet(data); ws['!cols']=hdr.map(()=>({wch:20}));
    XLSX.utils.book_append_sheet(wb,ws,'Master Data');
    XLSX.writeFile(wb,'master_policy_template.xlsx');
  }catch(err){ alert('Error: '+err.message); }
}

function downloadStmtTemplate(){
  try{
    const wb=XLSX.utils.book_new();
    const hdr=['Policy No','Endorsement No','Customer Name','Insurer Name','OD Premium','TP Premium','Total Premium','Policy Start Date','Policy End Date','Commission Amount'];
    const data=[hdr,
      ['4225/1000512039/00/0000','4225/1000512039/00/0000','GAURAV FATESARIA','ICICI LOMBARD',24922,0,24922,'16/03/2026','15/03/2029',5106],
      ['4225/1000512039/00/0001','4225/1000512039/00/0001','GAURAV FATESARIA','ICICI LOMBARD',500,0,500,'20/04/2026','15/03/2029',102],
      ['4225/1000483870/00/0000','4225/1000483870/00/0000','KOUSIK ROY',      'ICICI LOMBARD',23256,0,23256,'02/03/2026','01/03/2027',4765],
      ['MOT/2024/001234','','Rajesh Kumar','New India Assurance',8500,2200,10700,'01/04/2024','31/03/2025',1070],
      ['FIR/2024/005678','','Priya Sharma & Co','Oriental Insurance',0,0,25000,'01/07/2024','30/06/2025',2500],
    ];
    const ws=XLSX.utils.aoa_to_sheet(data); ws['!cols']=hdr.map(()=>({wch:22}));
    XLSX.utils.book_append_sheet(wb,ws,'Commission Statement');
    XLSX.writeFile(wb,'commission_statement_template.xlsx');
  }catch(err){ alert('Error: '+err.message); }
}

/* ══════════════════════════════════════════════
   RULE FORM HELPERS (show/hide suffix)
══════════════════════════════════════════════ */
function onPolnoModeChange(prefix){
  const v=document.getElementById(prefix+'-polno-mode');
  const w=document.getElementById(prefix+'-suffix-wrap');
  if(v&&w) w.style.display=v.value==='suffix'?'block':'none';
}

/* Export/Import rules as JSON */
function exportRules(){
  const data=JSON.stringify({defaultRule, insurerRules}, null, 2);
  const blob=new Blob([data],{type:'application/json'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download='commission_rules_'+new Date().toISOString().slice(0,10)+'.json';
  a.click();
}

function importRulesFile(inp){
  const file=inp.files[0]; if(!file) return;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const d=JSON.parse(e.target.result);
      if(d.defaultRule){ defaultRule={...FACTORY_DEFAULT,...d.defaultRule}; saveDefaultRuleToStorage(); renderDefaultRuleForm(); }
      if(d.insurerRules&&Array.isArray(d.insurerRules)){ insurerRules=d.insurerRules; saveRules(); renderRulesTable(); }
      showToast('✅ Rules imported successfully');
    }catch(err){ alert('Import error: '+err.message); }
  };
  reader.readAsText(file);
}

/* ══════════════════════════════════════════════
   INIT
══════════════════════════════════════════════ */
function init(){
  loadRules();  // also loads defaultRule and renders forms
  loadConso();
  if(consoData.length){ const t=consoData.reduce((s,b)=>s+b.rows.length,0); const b=document.getElementById('nav-conso-badge'); b.textContent=t; b.style.display='inline-block'; }
  try{
    const saved=localStorage.getItem(SK_MASTER);
    const savedM=localStorage.getItem(SK_MASTER_M);
    if(saved&&savedM){
      masterData=JSON.parse(saved); masterMapping=JSON.parse(savedM); window.masterMapping=masterMapping;
      masterCols=Object.keys(masterData[0]||{});
      precomputeMaster();
      document.getElementById('master-upload-area').style.display='none';
      document.getElementById('master-saved-banner').style.display='block';
      document.getElementById('master-saved-info').textContent=masterData.length.toLocaleString()+' records loaded.';
      showMasterReady();
    }
  }catch(e){ console.warn('Could not restore master data',e); }
}

init();
