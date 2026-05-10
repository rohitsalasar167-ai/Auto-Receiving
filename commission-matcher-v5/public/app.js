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
   HELPERS
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
  enddate:['policy end date','end date','expiry date','to date','expiry','risk end'],
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
    r._polno   = norm(r[masterMapping.polno]  ??'');
    r._endno   = norm(r[masterMapping.endno]  ??'');
    r._cust    = String(r[masterMapping.custname]??'');
    r._start   = r[masterMapping.startdate]??'';
    r._end     = r[masterMapping.enddate]  ??'';
    r._instno  = r[masterMapping.instno];
    r._isBase  = norm(r._instno)==='0'||r._instno===0||norm(r._instno)==='';
    r._cno     = String(r[masterMapping.cno]??'');
    r._brok    = r[masterMapping.brokdue]  ??'';
    r._dept    = String(r[masterMapping.dept]??'').toLowerCase();
    r._sumins  = r[masterMapping.sumins]   ??'';
    r._zeroPol = isZeroPolno(r._polno);
  });
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
   CORE MATCHING ENGINE
══════════════════════════════════════════════ */
function matchOneRow(sRow){
  const rawPolno = String(sRow[stmtMapping.polno]||'');
  const rawEndno = String(sRow[stmtMapping.endno]||'');
  const sCust    = String(sRow[stmtMapping.custname]||'');
  const sOD      = sRow[stmtMapping.odprem];
  const sTP      = sRow[stmtMapping.tpprem];
  const sTotal   = sRow[stmtMapping.totprem];
  const sStart   = sRow[stmtMapping.startdate];
  const sEnd     = sRow[stmtMapping.enddate];
  const sIns     = String(sRow[stmtMapping.insurer]||'');

  const rule = getEffectiveRule(sIns);

  // Transform statement policy no
  const transformed = transformStmtPolno(rawPolno, rule.polnoTransform);
  const sPolnoKey  = transformed.key;
  const endoSeqFromPolno = transformed.endoSeq;

  // Determine if this is an endorsement row
  const stmtEndoSeq = getStmtEndoSeq(rawPolno, rawEndno);
  const stmtIsEndo  = isEndoRow(rawEndno) || endoSeqFromPolno>0 || stmtEndoSeq>0;

  const sEndnoNorm = norm(rawEndno);

  // Build master pool
  const deptF = (rule.deptFilter||'').toLowerCase().trim();
  const pool = masterData.filter(m=>{
    if(rule.skipZeroBrok && isZeroBrokDue(m._brok)) return false;
    if(rule.skipBaseForEndo && stmtIsEndo && m._isBase) return false;
    if(rule.skipEndoForBase && !stmtIsEndo && !m._isBase) return false;
    if(deptF && m._dept && !m._dept.includes(deptF)) return false;
    return true;
  });

  let best=null, bestScore=0, bestOn=[], bestDetails={};

  for(const m of pool){
    // ── Policy No ──
    const mPolnoZero = m._zeroPol;
    let polHit=false, polExact=false;
    if(!mPolnoZero && rule.polnoMode!=='ignore' && sPolnoKey){
      const r=polnoHit(sPolnoKey, m._polno, rule.polnoMode, rule.suffixN);
      polHit=r.hit; polExact=r.exact;
    }

    // ── Endorsement No ──
    let endHit=false;
    if(stmtIsEndo && sEndnoNorm && m._endno && sEndnoNorm===m._endno) endHit=true;
    // Also try: statement polno key against master endno
    if(!endHit && stmtIsEndo && sPolnoKey && m._endno && sPolnoKey===m._endno) endHit=true;

    // ── Inst No ──
    let instOk=true;
    if(rule.instnoMode!=='none'){
      const seq = stmtEndoSeq || endoSeqFromPolno;
      instOk = instnoMatchesSeq(m._instno, seq, rule.instnoMode);
    }
    if(rule.instnoRequired && !instOk) continue;

    // ── Customer Name ──
    const custScr = nameMatch(sCust, m._cust, rule.fuzzyPct);
    const custHit = custScr>0;

    // ── Premium ──
    const premRes = premiumMatch(m, sOD, sTP, sTotal, rule.premTolPct, rule.premTolAbs, rule.useShare);
    const premHit = premRes.hit;

    // ── Date / Year ──
    let dateOk=true, yearHit=false;
    if(rule.dateCheck==='yes'){
      dateOk=yearMatch(m._start, m._end, sStart, sEnd);
    } else if(rule.dateCheck==='optional'){
      yearHit=yearMatch(m._start, m._end, sStart, sEnd);
    }
    if(!dateOk) continue;

    // ── Sum Insured (optional) ──
    if(rule.sumInsCheck){
      const mSI=toNum(m._sumins), sSI=toNum(sRow[stmtMapping?.sumins]);
      if(mSI&&sSI){ const diff=Math.abs(mSI-sSI)/Math.max(mSI,sSI)*100; if(diff>rule.sumInsTolPct) continue; }
    }

    // ── Zero Polno fallback ──
    // If master polno is 0/empty, skip polno check and rely on name+prem
    let usingZeroFallback = false;
    if(mPolnoZero && rule.zeroPolicyFallback){
      usingZeroFallback = true;
      // Don't count polHit positively for zero-polno rows
    }

    /* ═══════════════════════════════════════
       SCORE COMPUTATION
       Priority order — first match wins
    ═══════════════════════════════════════ */
    let score=0, on=[];

    if(!usingZeroFallback){
      if(polHit && custHit && premHit){        score=rule.scorePolCustPrem; on=['PolNo','Name','Prem']; }
      else if(polHit && premHit){              score=rule.scorePolPrem;     on=['PolNo','Prem']; }
      else if(polHit && custHit){              score=rule.scorePolCust;     on=['PolNo','Name']; }
      else if(endHit && premHit && custHit){   score=rule.scoreEndoPrem+5;  on=['EndoNo','Prem','Name']; }
      else if(endHit && premHit){              score=rule.scoreEndoPrem;    on=['EndoNo','Prem']; }
      else if(endHit && custHit){              score=rule.scoreEndoCust;    on=['EndoNo','Name']; }
      else if(stmtIsEndo && endHit){           score=rule.scoreEndoOnly;    on=['EndoNo']; }
      else if(!stmtIsEndo && polHit){          score=rule.scorePolOnly;     on=['PolNo']; }
      else if(custHit && premHit){             score=rule.scoreCustPrem;    on=['Name','Prem']; }
      else if(custHit){                        score=rule.scoreCustOnly;    on=['Name']; }
      else if(premHit){                        score=rule.scorePremOnly;    on=['Prem']; }
    } else {
      // Zero polno: only name+prem+date
      if(custHit && premHit){   score=rule.scoreCustPrem;  on=['Name','Prem','(zeroPol)']; }
      else if(custHit){         score=rule.scoreCustOnly;  on=['Name','(zeroPol)']; }
    }

    if(score===0) continue;

    // ── Boosts ──
    if(polExact && on.includes('PolNo'))    score=Math.min(score+rule.boostExactPolno, 100);
    if(custScr>=85 && on.includes('Name'))  score=Math.min(score+rule.boostStrongName, 100);
    if(premRes.exact && on.includes('Prem'))score=Math.min(score+rule.boostExactPrem, 100);
    if(yearHit)                             score=Math.min(score+rule.boostYearMatch, 100);
    if(instOk && rule.instnoMode!=='none')  score=Math.min(score+2, 100); // small inst bonus

    if(score>bestScore){ bestScore=score; best=m; bestOn=[...on]; bestDetails={polExact,custScr,premRes,instOk,yearHit}; }
  }

  const status = bestScore>=rule.scoreFullMatch ? '100' : bestScore>=rule.scorePartial ? 'partial' : 'unmatched';
  const onStr = bestOn.join(' + ') + (bestDetails.yearHit?' +Year':'');

  return {
    _stmtRow:     sRow,
    _cno:         best ? best._cno : '',
    _instno:      best ? String(best._instno??'') : '',
    _score:       bestScore,
    _matchStatus: status,
    _matchedOn:   onStr,
    _insurer:     sIns,
    _ruleKey:     sIns ? norm(sIns).slice(0,20) : 'default',
    _debug:{
      polnoKey: sPolnoKey,
      endoSeq: stmtEndoSeq||endoSeqFromPolno,
      isEndo: stmtIsEndo,
    }
  };
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

  matchResults=[];
  const total=stmtData.length;
  const skip=masterData.filter(r=>isZeroBrokDue(r._brok)).length;
  addLog(`Processing ${total} statement rows…`,'log-ok');
  addLog(`Master: ${masterData.length} records (${skip} excluded — zero brok due)`,'log-warn');
  addLog(`Policy No transform: ${defaultRule.polnoTransform} | Inst No mode: ${defaultRule.instnoMode}`,'log-ok');
  addLog(`Premium tol: ${defaultRule.premTolPct}% or ₹${defaultRule.premTolAbs} | Date check: ${defaultRule.dateCheck}`,'log-ok');
  if(insurerRules.length) addLog(`Custom insurer rules: ${insurerRules.map(r=>r.insurer).join(', ')}`,'log-ok');

  let c100=0,cP=0,cU=0;
  for(let i=0;i<total;i+=10){
    const batch=stmtData.slice(i,Math.min(i+10,total));
    fill.style.width=Math.round(i/total*100)+'%';
    lbl.textContent=`Row ${i+1}–${Math.min(i+10,total)} of ${total}`;
    for(const row of batch){
      const r=matchOneRow(row); matchResults.push(r);
      if(r._matchStatus==='100') c100++;
      else if(r._matchStatus==='partial') cP++;
      else cU++;
    }
    await new Promise(res=>setTimeout(res,4));
  }
  fill.style.width='100%';
  lbl.textContent='Done!';
  addLog(`✅ 100% matched: ${c100}  |  Partial: ${cP}  |  Unmatched: ${cU}`,'log-ok');
  if(cU>0) addLog(`⚠ ${cU} rows need manual review`,'log-warn');

  const badge=document.getElementById('nav-results-badge');
  badge.textContent=matchResults.length; badge.style.display='inline-block';
  setTimeout(()=>{ document.getElementById('matching-progress').style.display='none'; switchTab('results'); },900);
}

/* ══════════════════════════════════════════════
   RESULTS TABLE
══════════════════════════════════════════════ */
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
    const badge=r._matchStatus==='100'?'<span class="badge green">✓ 100%</span>':r._matchStatus==='partial'?'<span class="badge amber">Partial</span>':'<span class="badge red">Unmatched</span>';
    const scoreClass=r._score>=95?'green':r._score>=70?'':'amber';
    const scorePill=`<span class="score-pill ${scoreClass}">${r._score}</span>`;
    const mo=r._matchedOn?`<span class="match-on-tag">${esc(r._matchedOn)}</span>`:'—';
    const rl=`<span class="rule-tag">${esc(r._ruleKey)}</span>`;
    return `<tr>${cells}<td><strong>${esc(r._cno)}</strong></td><td>${esc(r._instno)}</td><td>${scorePill}</td><td>${mo}</td><td>${rl}</td><td>${badge}</td></tr>`;
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
