/* ═══════════════════════════════════════════════════════════
   CommissionMatcher v6 — app.js
   New in v6:
   ✓ Policy No trimmer for Stmt & Master (left/right N chars)
   ✓ Premium mode: Total / OD / TP / OD+TP / Any
   ✓ Insurer must-match guard (configurable)
   ✓ 12-step editable search waterfall with per-step scores
   ✓ Date mode: Start / End / Either / Both
   ✓ Endo Date fallback to Start/End if blank/mismatch
   ✓ Zero polno → fully rely on name+prem+date
   ✓ Contra: commission sum=0 for same polno → Contra
     If Cno found → show as Partial Match with real score
   ✓ Matched On & Step columns in results
═══════════════════════════════════════════════════════════ */

const SK_MASTER   = 'cm_master_v6';
const SK_MASTER_M = 'cm_master_map_v6';
const SK_RULES    = 'cm_rules_v6';
const SK_DEFRULE  = 'cm_defrule_v6';
const SK_CONSO    = 'cm_conso_v6';

let masterData = null, masterCols = [], masterMapping = {};
let stmtData   = null, stmtCols   = [], stmtMapping   = {};
let matchResults = [], insurerRules = [], consoData = [];
let defaultRule  = null, masterMaps = null;

window.masterMapping = masterMapping;
window.stmtMapping   = stmtMapping;

/* ═══════════════════════════════════════════════════════════
   FACTORY DEFAULT RULE
═══════════════════════════════════════════════════════════ */
const FACTORY_DEFAULT = {
  stmtPolTrimDir:   'left',   // left | right
  stmtPolTrimN:     21,
  masterPolTrimDir: 'left',
  masterPolTrimN:   21,
  removeColon:      true,
  requireInsurerMatch: true,
  insurerFuzzy:     true,
  premMode:         'any',    // any | total | od | tp | od_tp
  premTolPct:       2,
  premTolAbs:       50,
  useShare:         true,
  multiYearTP:      true,
  nameStrip:        true,
  fuzzyPct:         60,
  dateMode:         'either', // none | start | end | either | both
  endoDateFallback: true,
  zeroPolicyFallback: true,
  skipZeroBrok:     true,
  contraCheck:      true,
  contraThreshold:  1,
  scoreFullMatch:   85,
  scorePartial:     45,
  steps: [
    {id:'s1',  label:'Pol + Endo + Prem + Name + Date', on:true,  score:100},
    {id:'s2',  label:'Pol + Endo + Prem + Name',        on:true,  score:97},
    {id:'s3',  label:'Pol + Prem + Name + Date',        on:true,  score:95},
    {id:'s4',  label:'Pol + Prem + Name',               on:true,  score:92},
    {id:'s5',  label:'Endo + Prem + Name + Date',       on:true,  score:90},
    {id:'s6',  label:'Endo + Prem + Name',              on:true,  score:87},
    {id:'s7',  label:'Pol + Prem',                      on:true,  score:85},
    {id:'s8',  label:'Endo + Prem',                     on:true,  score:82},
    {id:'s9',  label:'Name + Prem + Date',              on:true,  score:75},
    {id:'s10', label:'Name + Prem',                     on:true,  score:65},
    {id:'s11', label:'Pol only',                        on:true,  score:55},
    {id:'s12', label:'Endo only',                       on:true,  score:50},
  ],
  notes: 'Default v6: 12-step waterfall. Left-21 trim both sides. Any-premium 2% tol. Insurer must match (fuzzy). Date: either start or end.',
};

/* ── Field definitions ── */
const MASTER_FIELDS = [
  {key:'cno',      label:'C.No (Control No)',           req:true},
  {key:'instno',   label:'Inst No (0=base)',             req:true},
  {key:'polno',    label:'Policy No',                    req:false},
  {key:'endno',    label:'Endorsement No',               req:false},
  {key:'custname', label:'Customer Name',                req:false},
  {key:'insurer',  label:'Insurer Name',                 req:false},
  {key:'dept',     label:'Department',                   req:false},
  {key:'ptype',    label:'Policy Type',                  req:false},
  {key:'odprem',   label:'OD Premium',                   req:false},
  {key:'tpprem',   label:'TP Premium',                   req:false},
  {key:'totprem',  label:'Total Premium (PBST)',         req:false},
  {key:'sharepct', label:'Share %',                      req:false},
  {key:'sumins',   label:'Sum Insured',                  req:false},
  {key:'startdate',label:'Policy Start Date',            req:false},
  {key:'endodate', label:'Endo Date',                    req:false},
  {key:'enddate',  label:'Policy End Date',              req:false},
  {key:'brokdue',  label:'Brok Due',                     req:false},
];
const STMT_FIELDS = [
  {key:'polno',     label:'Policy No',         req:false},
  {key:'endno',     label:'Endorsement No',    req:false},
  {key:'custname',  label:'Customer Name',     req:false},
  {key:'insurer',   label:'Insurer Name',      req:false},
  {key:'odprem',    label:'OD Premium',        req:false},
  {key:'tpprem',    label:'TP Premium',        req:false},
  {key:'totprem',   label:'Total Premium',     req:false},
  {key:'startdate', label:'Policy Start Date', req:false},
  {key:'enddate',   label:'Policy End Date',   req:false},
  {key:'commission',label:'Commission Amount', req:false},
];

/* ═══════════════════════════════════════════════════════════
   HELPERS
═══════════════════════════════════════════════════════════ */
const norm = s => String(s==null?'':s).toLowerCase().replace(/[\s\-\/\.\,]+/g,'').trim();
const esc  = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const toNum= v => { const n=parseFloat(String(v==null?'':v).replace(/[^0-9.\-]/g,'')); return isNaN(n)?null:n; };

function toExcelSerial(v){
  if(!v&&v!==0) return null;
  if(typeof v==='number'&&v>10000) return Math.round(v);
  let d = v instanceof Date ? v : null;
  if(!d){
    const s=String(v).trim();
    const m1=s.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    const m2=s.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
    if(m1) d=new Date(Date.UTC(+m1[1],+m1[2]-1,+m1[3]));
    else if(m2) d=new Date(Date.UTC(+m2[3],+m2[2]-1,+m2[1]));
    else d=new Date(s);
  }
  return (d&&!isNaN(d.getTime()))?Math.floor(d.getTime()/86400000+25569):null;
}

function extractYear(v){
  if(!v&&v!==0) return null;
  if(typeof v==='number'&&v>10000){ try{ const d=new Date(Math.round((v-25569)*86400*1000)); const y=d.getUTCFullYear(); return(y>1900&&y<2100)?y:null;}catch{return null;} }
  const s=String(v);
  const m1=s.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/); if(m1) return parseInt(m1[3]);
  const m2=s.match(/(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/); if(m2) return parseInt(m2[1]);
  const m3=s.match(/\b(20\d{2}|19\d{2})\b/); if(m3) return parseInt(m3[1]);
  return null;
}

/* ── Policy No trimmer ── */
function trimPolno(raw, dir, n, removeColon){
  if(!raw) return '';
  let s=String(raw).trim();
  if(removeColon) s=s.replace(/:/g,'');
  if(!n||n===0) return s;
  return dir==='right' ? s.slice(-n) : s.substring(0,n);
}

/* ── Customer name cleaning ── */
const STRIP_WORDS=['MR.','MRS.','MS.','MISS.','DR.','SHRI.','SMT.','KUMAR.','M/S.',
  'MR','MRS','MS','MISS','DR','SHRI','SMT','KUMAR','M/S',
  'PVT.','PVT','PRIVATE','LIMITED','LTD.','LTD','LLC','LLP','INC','CORP'];

function freshCustName(raw){
  if(!raw) return '';
  let s=String(raw).toUpperCase();
  STRIP_WORDS.forEach(t=>{ s=s.split(t).join(''); });
  s=s.replace(/[^A-Z0-9\s&]/g,'');
  return s.replace(/\s+/g,' ').trim();
}

function nameSim(a,b,thr){
  const ca=freshCustName(a), cb=freshCustName(b);
  if(!ca||!cb) return 0;
  if(ca===cb) return 100;
  const wa=ca.split(/\s+/).filter(w=>w.length>1);
  const wb=cb.split(/\s+/).filter(w=>w.length>1);
  if(!wa.length||!wb.length) return 0;
  const fwd=wa.filter(w=>cb.includes(w)).length/wa.length;
  const rev=wb.filter(w=>ca.includes(w)).length/wb.length;
  const sc=Math.round(Math.max(fwd,rev)*100);
  return sc>=(thr||60)?sc:0;
}

/* ── Insurer fuzzy match ── */
function insurerOk(sIns,mIns,fuzzy){
  if(!sIns||!mIns) return true;
  const sn=norm(sIns), mn=norm(mIns);
  if(sn===mn) return true;
  if(!fuzzy) return false;
  return sn.includes(mn)||mn.includes(sn)||sn.slice(0,8)===mn.slice(0,8);
}

/* ── Premium hit ── */
function premHit(m,sk,rule){
  const tol=rule.premTolPct??2, tolA=rule.premTolAbs??50;
  const shr=(rule.useShare&&toNum(m._share)>0&&toNum(m._share)<100)?toNum(m._share)/100:1;
  const mOD=(toNum(m._odprem)??0)/shr, mTP=(toNum(m._tpprem)??0)/shr, mTot=(toNum(m._totprem)??0)/shr;
  const sy=extractYear(m._start),ey=extractYear(m._end),yd=(sy&&ey)?Math.max(0,ey-sy):0;

  const close=(a,b)=>{
    if(a===null||b===null) return false;
    if(Math.abs(a)<0.01&&Math.abs(b)<0.01) return true;
    const diff=Math.abs(a-b);
    return diff/Math.max(Math.abs(a),Math.abs(b),0.01)*100<=tol||diff<=tolA;
  };

  const mode=rule.premMode||'any';
  if(mode==='total') return close(mTot,sk.tot)||close(mTot/shr,sk.tot);
  if(mode==='od')    return close(mOD,sk.od)||close(mOD,sk.tot);
  if(mode==='tp')    return close(mTP,sk.tp)||close(mTP,sk.tot);
  if(mode==='od_tp') return close(mOD,sk.od)||close(mTP,sk.tp);

  // 'any' — 10-way check
  const grp=m._cc&&masterMaps?.byCC?.has(m._cc)?masterMaps.byCC.get(m._cc).reduce((s,r)=>s+(toNum(r._totprem)||0),0):null;
  return close(mTot,sk.tot)||close(mOD,sk.tot)||close(mTP,sk.tot)||
    close(mOD+yd*mTP,sk.tot)||(grp!==null&&close(grp,mTot))||
    close(mTP/shr,sk.tot)||close(mOD/shr,sk.tot)||close(mTot/shr,sk.tot)||
    close(mOD,sk.od)||close(mTP,sk.tp);
}

/* ── Date hit ── */
function dateHit(m,sk,rule,sEnd){
  const mode=rule.dateMode||'either';
  if(mode==='none') return true;
  const eq=(a,b)=>{ const sa=toExcelSerial(a),sb=toExcelSerial(b); return sa!=null&&sb!=null&&sa===sb; };
  const instBase=(String(m._instno||'')==='0'||m._instno===0||String(m._instno||'')==='');

  let startOk=false;
  if(!instBase&&m._endodate){
    startOk=eq(m._endodate,sk.startV);
    if(!startOk&&rule.endoDateFallback) startOk=eq(m._start,sk.startV);
  } else {
    startOk=eq(m._start,sk.startV);
    if(!startOk&&rule.endoDateFallback&&m._endodate) startOk=eq(m._endodate,sk.startV);
  }
  const endOk=eq(m._end,sEnd);

  if(mode==='start')  return startOk;
  if(mode==='end')    return endOk;
  if(mode==='either') return startOk||endOk;
  if(mode==='both')   return startOk&&endOk;
  return true;
}

/* ── Polno match (trimmed both sides) ── */
function polMatch(sp,mp){
  if(!sp||!mp) return false;
  if(sp===mp) return true;
  return mp.includes(sp)||sp.includes(mp);
}

/* ── Endo match ── */
function endMatch(se,me){
  if(!se||!me||se==='0'||me==='0') return false;
  return se===me||me.includes(se)||se.includes(me);
}

/* ── Zero brok check ── */
function isZeroBrokDue(v){ const n=toNum(v); return n!==null&&Math.abs(n)<=5; }

/* ═══════════════════════════════════════════════════════════
   AUTO-DETECT
═══════════════════════════════════════════════════════════ */
const HINTS={
  cno:['c.no','cno','control no','ctrl no'],
  instno:['inst no','instno','inst'],
  polno:['policy no','policy number','pol no','polno','policy'],
  endno:['endorsement no','endorsement','end no','endno','endt no'],
  custname:['customer name','customer','insured name','insured','client','party name'],
  insurer:['insurer name','insurer','insurance company','ins co','company'],
  dept:['department','dept','branch','lob'],
  ptype:['policy type','type','product','plan'],
  odprem:['od premium','od prem','own damage','od'],
  tpprem:['tp premium','tp prem','third party','tp'],
  totprem:['total premium','total prem','net premium','gross premium','premium','pbst','tot prem'],
  sharepct:['share %','share%','share pct','our share','share'],
  sumins:['sum insured','sum assured','si','tsi'],
  startdate:['policy start date','start date','inception date','from date','risk start'],
  endodate:['endo date','endorsement date','endt date','risk date'],
  enddate:['policy end date','end date','expiry date','to date','expiry','risk end'],
  brokdue:['brok due','brokerage due','commission due','brok'],
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

/* ═══════════════════════════════════════════════════════════
   TABS
═══════════════════════════════════════════════════════════ */
const TABS=['master','rules','match','results','conso','samples','howrules'];
function switchTab(name){
  TABS.forEach(id=>{
    const s=document.getElementById('tab-'+id);
    const b=document.querySelector(`.nav-item[data-tab="${id}"]`);
    if(s) s.classList.toggle('active',id===name);
    if(b) b.classList.toggle('active',id===name);
  });
  if(name==='match'){ const h=!!masterData; document.getElementById('match-no-master').style.display=h?'none':'block'; document.getElementById('match-main-area').style.display=h?'block':'none'; }
  if(name==='results'){ const h=matchResults.length>0; document.getElementById('results-empty').style.display=h?'none':'block'; document.getElementById('results-content').style.display=h?'block':'none'; if(h) renderResultsTable(); }
  if(name==='conso') renderConsoPage();
}
document.querySelectorAll('.nav-item[data-tab]').forEach(b=>b.addEventListener('click',()=>switchTab(b.dataset.tab)));
function showToast(msg,d=2800){ const t=document.getElementById('toast'); t.textContent=msg; t.classList.add('show'); setTimeout(()=>t.classList.remove('show'),d); }

/* ═══════════════════════════════════════════════════════════
   MASTER DATA
═══════════════════════════════════════════════════════════ */
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
  const rule=defaultRule||FACTORY_DEFAULT;
  masterData.forEach(r=>{
    r._odprem   = r[masterMapping.odprem]   ??'';
    r._tpprem   = r[masterMapping.tpprem]   ??'';
    r._totprem  = r[masterMapping.totprem]  ??'';
    r._share    = r[masterMapping.sharepct] ??'';
    r._rawPol   = String(r[masterMapping.polno]    ??'');
    r._rawEndo  = String(r[masterMapping.endno]    ??'');
    r._cust     = String(r[masterMapping.custname] ??'');
    r._insurer  = String(r[masterMapping.insurer]  ??'');
    r._start    = r[masterMapping.startdate]??'';
    r._end      = r[masterMapping.enddate]  ??'';
    r._endodate = r[masterMapping.endodate] ??'';
    r._instno   = r[masterMapping.instno];
    r._isBase   = (String(r._instno||'')==='0'||r._instno===0||String(r._instno||'')==='');
    r._cno      = String(r[masterMapping.cno]??'');
    r._brok     = r[masterMapping.brokdue]  ??'';
    r._zeroPol  = !r._rawPol||r._rawPol==='0';
    r._polTrim  = trimPolno(r._rawPol,  rule.masterPolTrimDir||'left', rule.masterPolTrimN||21, rule.removeColon!==false).toUpperCase();
    r._endoTrim = trimPolno(r._rawEndo, rule.masterPolTrimDir||'left', rule.masterPolTrimN||21, rule.removeColon!==false).toUpperCase();
    r._custClean= freshCustName(r._cust);
    r._cc       = r._cno+String(r._instno??'');
    r._startSerial=toExcelSerial(r._start);
  });
  masterMaps=null; buildMasterMaps();
}

function saveMasterMapping(){
  document.querySelectorAll('#master-mapping-grid select').forEach((s,i)=>{ if(s.value) masterMapping[MASTER_FIELDS[i].key]=s.value; });
  if(!masterMapping.cno){ alert('Please map C.No'); return; }
  if(!masterMapping.instno){ alert('Please map Inst No'); return; }
  precomputeMaster();
  try{ localStorage.setItem(SK_MASTER,JSON.stringify(masterData)); localStorage.setItem(SK_MASTER_M,JSON.stringify(masterMapping)); }catch(e){ console.warn('save failed'); }
  showMasterReady(); showToast('✅ Master data saved');
}

function showMasterReady(){
  document.getElementById('master-mapping-area').style.display='none';
  document.getElementById('master-upload-area').style.display='none';
  document.getElementById('master-ready-area').style.display='block';
  document.getElementById('master-saved-banner').style.display='none';
  const base=masterData.filter(r=>r._isBase).length;
  const skip=masterData.filter(r=>isZeroBrokDue(r._brok)).length;
  const zp=masterData.filter(r=>r._zeroPol).length;
  document.getElementById('stat-total').textContent=masterData.length.toLocaleString();
  document.getElementById('stat-base').textContent=base.toLocaleString();
  document.getElementById('stat-endo').textContent=(masterData.length-base).toLocaleString();
  document.getElementById('stat-skip').textContent=skip.toLocaleString();
  document.getElementById('stat-zeropol').textContent=zp.toLocaleString();
  const badge=document.getElementById('nav-master-badge'); badge.textContent=masterData.length.toLocaleString(); badge.style.display='inline-block';
  document.getElementById('storage-status').innerHTML='<span class="sdot green"></span> Master saved ('+masterData.length+' records)';
  renderPreviewTable('master-preview-table',masterData.slice(0,6),masterCols.slice(0,8));
}

function resetMaster(){
  if(!confirm('Remove master data?')) return;
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

/* ═══════════════════════════════════════════════════════════
   LOOKUP MAPS
═══════════════════════════════════════════════════════════ */
function buildMasterMaps(){
  masterMaps={byPol:new Map(),byEndo:new Map(),byCust:new Map(),
    byPolPbst:new Map(),byPolDate:new Map(),byPolOD:new Map(),byPolTP:new Map(),
    byCustPbst:new Map(),byCustDate:new Map(),byCC:new Map()};

  const add=(map,key,row)=>{
    if(!key&&key!==0) return;
    const k=String(key).trim(); if(!k) return;
    if(!map.has(k)) map.set(k,[]); map.get(k).push(row);
  };

  masterData.forEach(m=>{
    const pt=m._polTrim, et=m._endoTrim, fc=m._custClean;
    const od=toNum(m._odprem)??0, tp=toNum(m._tpprem)??0, pbt=toNum(m._totprem)??0;
    const ds=m._startSerial, dsStr=ds!=null?String(ds):'';
    add(masterMaps.byPol,    pt,m);
    add(masterMaps.byEndo,   et,m);
    add(masterMaps.byCust,   fc,m);
    add(masterMaps.byPolPbst,pt+String(Math.round(pbt)),m);
    add(masterMaps.byPolPbst,pt+String(Math.round(od)),m);
    add(masterMaps.byPolPbst,pt+String(Math.round(tp)),m);
    if(dsStr){ add(masterMaps.byPolDate,pt+dsStr,m); add(masterMaps.byCustDate,fc+dsStr,m); }
    add(masterMaps.byPolOD,  pt+String(Math.round(od)),m);
    add(masterMaps.byPolTP,  pt+String(Math.round(tp)),m);
    add(masterMaps.byCustPbst,fc+String(Math.round(pbt)),m);
    add(masterMaps.byCustPbst,fc+String(Math.round(od)),m);
    add(masterMaps.byCustPbst,fc+String(Math.round(tp)),m);
    if(m._cc) add(masterMaps.byCC,m._cc,m);
  });
}

function lkp(map,key,fn){
  if(!key&&key!==0) return null;
  const rows=map.get(String(key).trim());
  if(!rows||!rows.length) return null;
  return fn?(rows.find(fn)||null):rows[0];
}

/* ═══════════════════════════════════════════════════════════
   DEFAULT RULE — SAVE / LOAD / RENDER
═══════════════════════════════════════════════════════════ */
function mergeSteps(saved){
  return FACTORY_DEFAULT.steps.map(fs=>{
    const ps=saved&&saved.find(s=>s.id===fs.id);
    return ps?{...fs,...ps}:{...fs};
  });
}

function loadDefaultRule(){
  try{
    const saved=localStorage.getItem(SK_DEFRULE);
    if(saved){
      const p=JSON.parse(saved);
      defaultRule={...FACTORY_DEFAULT,...p,steps:mergeSteps(p.steps)};
    } else { defaultRule={...FACTORY_DEFAULT,steps:mergeSteps(null)}; }
  }catch{ defaultRule={...FACTORY_DEFAULT,steps:mergeSteps(null)}; }
}

function saveDefaultRuleToStorage(){ try{ localStorage.setItem(SK_DEFRULE,JSON.stringify(defaultRule)); }catch{} }

function resetDefaultRule(){
  if(!confirm('Reset default rule to factory settings?')) return;
  defaultRule={...FACTORY_DEFAULT,steps:mergeSteps(null)};
  saveDefaultRuleToStorage(); renderDefaultRuleForm();
  showToast('🔄 Default rule reset to factory');
}

function gv(id,type='str'){
  const el=document.getElementById(id); if(!el) return undefined;
  if(type==='int') return parseInt(el.value)||0;
  if(type==='num') return parseFloat(el.value)||0;
  if(type==='bool') return el.value==='yes';
  return el.value;
}

function saveDefaultRuleFromForm(){
  defaultRule.stmtPolTrimDir  = gv('dr-stmt-dir');
  defaultRule.stmtPolTrimN    = gv('dr-stmt-n','int');
  defaultRule.masterPolTrimDir= gv('dr-mst-dir');
  defaultRule.masterPolTrimN  = gv('dr-mst-n','int');
  defaultRule.removeColon     = gv('dr-removecolon','bool');
  defaultRule.requireInsurerMatch=gv('dr-ins-match','bool');
  defaultRule.insurerFuzzy    = gv('dr-ins-fuzzy','bool');
  defaultRule.premMode        = gv('dr-prem-mode');
  defaultRule.premTolPct      = gv('dr-prem-tol-pct','num');
  defaultRule.premTolAbs      = gv('dr-prem-tol-abs','num');
  defaultRule.useShare        = gv('dr-share','bool');
  defaultRule.multiYearTP     = gv('dr-multiyear','bool');
  defaultRule.nameStrip       = gv('dr-namestrip','bool');
  defaultRule.fuzzyPct        = gv('dr-fuzzy','int');
  defaultRule.dateMode        = gv('dr-date-mode');
  defaultRule.endoDateFallback= gv('dr-endo-fallback','bool');
  defaultRule.zeroPolicyFallback=gv('dr-zeropol','bool');
  defaultRule.skipZeroBrok    = gv('dr-skip-brok','bool');
  defaultRule.contraCheck     = gv('dr-contra','bool');
  defaultRule.contraThreshold = gv('dr-contra-thr','num');
  defaultRule.scoreFullMatch  = gv('dr-score-full','int');
  defaultRule.scorePartial    = gv('dr-score-partial','int');
  defaultRule.notes           = gv('dr-notes');
  defaultRule.steps = defaultRule.steps.map(st=>{
    const onEl=document.getElementById('dr-on-'+st.id);
    const scEl=document.getElementById('dr-sc-'+st.id);
    return {...st, on:onEl?onEl.checked:st.on, score:scEl?(parseInt(scEl.value)||st.score):st.score};
  });
  saveDefaultRuleToStorage(); showToast('✅ Default rule saved');
}

function sv(id,val){ const el=document.getElementById(id); if(el) el.value=String(val??''); }
function sbv(id,val){ const el=document.getElementById(id); if(el) el.value=val?'yes':'no'; }

function renderDefaultRuleForm(){
  if(!defaultRule) return;
  sv('dr-stmt-dir',  defaultRule.stmtPolTrimDir);  sv('dr-stmt-n',     defaultRule.stmtPolTrimN);
  sv('dr-mst-dir',   defaultRule.masterPolTrimDir); sv('dr-mst-n',      defaultRule.masterPolTrimN);
  sbv('dr-removecolon',defaultRule.removeColon);    sbv('dr-ins-match',  defaultRule.requireInsurerMatch);
  sbv('dr-ins-fuzzy',defaultRule.insurerFuzzy);     sv('dr-prem-mode',   defaultRule.premMode);
  sv('dr-prem-tol-pct',defaultRule.premTolPct);     sv('dr-prem-tol-abs',defaultRule.premTolAbs);
  sbv('dr-share',    defaultRule.useShare);          sbv('dr-multiyear',  defaultRule.multiYearTP);
  sbv('dr-namestrip',defaultRule.nameStrip);         sv('dr-fuzzy',       defaultRule.fuzzyPct);
  sv('dr-date-mode', defaultRule.dateMode);          sbv('dr-endo-fallback',defaultRule.endoDateFallback);
  sbv('dr-zeropol',  defaultRule.zeroPolicyFallback);sbv('dr-skip-brok',  defaultRule.skipZeroBrok);
  sbv('dr-contra',   defaultRule.contraCheck);       sv('dr-contra-thr',  defaultRule.contraThreshold);
  sv('dr-score-full',defaultRule.scoreFullMatch);    sv('dr-score-partial',defaultRule.scorePartial);
  sv('dr-notes',     defaultRule.notes||'');
  renderStepTable('dr-steps-tbody', defaultRule.steps, 'dr-on-', 'dr-sc-');
}

function renderStepTable(tbodyId, steps, onPfx, scPfx){
  const tb=document.getElementById(tbodyId); if(!tb) return;
  tb.innerHTML=steps.map((st,i)=>`
    <tr>
      <td class="step-num">${i+1}</td>
      <td class="step-label">${esc(st.label)}</td>
      <td><label class="toggle-label"><input type="checkbox" id="${onPfx+st.id}" ${st.on?'checked':''}><span class="toggle-pill"></span></label></td>
      <td><input type="number" id="${scPfx+st.id}" value="${st.score}" min="0" max="100" class="score-input"></td>
    </tr>`).join('');
}

/* ═══════════════════════════════════════════════════════════
   PER-INSURER RULES
═══════════════════════════════════════════════════════════ */
function loadRules(){
  try{ const r=localStorage.getItem(SK_RULES); if(r) insurerRules=JSON.parse(r); }catch{}
  loadDefaultRule(); renderRulesTable(); renderDefaultRuleForm();
}
function saveRules(){ try{ localStorage.setItem(SK_RULES,JSON.stringify(insurerRules)); }catch{} }
function getEffectiveRule(ins){
  if(ins){ const n=norm(ins); const ov=insurerRules.find(r=>n.includes(norm(r.insurer))||norm(r.insurer).includes(n)); if(ov) return {...defaultRule,...ov,steps:ov.steps||defaultRule.steps}; }
  return defaultRule;
}

function addInsurerRule(){
  const insurer=document.getElementById('r-insurer').value.trim(); if(!insurer){ alert('Insurer name required'); return; }
  const rule={insurer,
    stmtPolTrimDir:gv('r-stmt-dir'),stmtPolTrimN:gv('r-stmt-n','int'),masterPolTrimDir:gv('r-mst-dir'),masterPolTrimN:gv('r-mst-n','int'),
    removeColon:gv('r-removecolon','bool'),requireInsurerMatch:gv('r-ins-match','bool'),insurerFuzzy:gv('r-ins-fuzzy','bool'),
    premMode:gv('r-prem-mode'),premTolPct:gv('r-prem-tol-pct','num'),premTolAbs:gv('r-prem-tol-abs','num'),
    useShare:gv('r-share','bool'),multiYearTP:gv('r-multiyear','bool'),
    nameStrip:gv('r-namestrip','bool'),fuzzyPct:gv('r-fuzzy','int'),
    dateMode:gv('r-date-mode'),endoDateFallback:gv('r-endo-fallback','bool'),
    zeroPolicyFallback:gv('r-zeropol','bool'),skipZeroBrok:gv('r-skip-brok','bool'),
    contraCheck:gv('r-contra','bool'),contraThreshold:gv('r-contra-thr','num'),
    scoreFullMatch:gv('r-score-full','int'),scorePartial:gv('r-score-partial','int'),notes:gv('r-notes'),
  };
  Object.keys(rule).forEach(k=>{ if(rule[k]===undefined||rule[k]==='') delete rule[k]; });
  const idx=insurerRules.findIndex(r=>norm(r.insurer)===norm(insurer));
  if(idx>=0) insurerRules[idx]=rule; else insurerRules.push(rule);
  saveRules(); renderRulesTable(); document.getElementById('r-insurer').value=''; showToast('✅ Rule saved for '+insurer);
}

function deleteRule(i){ insurerRules.splice(i,1); saveRules(); renderRulesTable(); showToast('Rule deleted'); }

function editRule(i){
  const r=insurerRules[i];
  sv('r-insurer',r.insurer); sv('r-stmt-dir',r.stmtPolTrimDir); sv('r-stmt-n',r.stmtPolTrimN);
  sv('r-mst-dir',r.masterPolTrimDir); sv('r-mst-n',r.masterPolTrimN);
  sbv('r-removecolon',r.removeColon); sbv('r-ins-match',r.requireInsurerMatch); sbv('r-ins-fuzzy',r.insurerFuzzy);
  sv('r-prem-mode',r.premMode); sv('r-prem-tol-pct',r.premTolPct); sv('r-prem-tol-abs',r.premTolAbs);
  sbv('r-share',r.useShare); sbv('r-multiyear',r.multiYearTP);
  sbv('r-namestrip',r.nameStrip); sv('r-fuzzy',r.fuzzyPct);
  sv('r-date-mode',r.dateMode); sbv('r-endo-fallback',r.endoDateFallback);
  sbv('r-zeropol',r.zeroPolicyFallback); sbv('r-skip-brok',r.skipZeroBrok);
  sbv('r-contra',r.contraCheck); sv('r-contra-thr',r.contraThreshold);
  sv('r-score-full',r.scoreFullMatch); sv('r-score-partial',r.scorePartial); sv('r-notes',r.notes);
  document.getElementById('r-insurer').scrollIntoView({behavior:'smooth'});
}

function renderRulesTable(){
  const empty=document.getElementById('rules-empty'),wrap=document.getElementById('rules-table-wrap');
  const hint=document.getElementById('rules-count-hint'),badge=document.getElementById('nav-rules-badge');
  if(!insurerRules.length){ empty.style.display='block';wrap.style.display='none';hint.textContent='— none yet';badge.style.display='none';return; }
  empty.style.display='none';wrap.style.display='block';
  hint.textContent=`— ${insurerRules.length} rule${insurerRules.length>1?'s':''} saved`;
  badge.textContent=insurerRules.length;badge.style.display='inline-block';
  document.getElementById('rules-tbody').innerHTML=insurerRules.map((r,i)=>`
    <tr><td><strong>${esc(r.insurer)}</strong></td>
    <td>${r.stmtPolTrimDir||'—'} ${r.stmtPolTrimN||'—'}c / ${r.masterPolTrimDir||'—'} ${r.masterPolTrimN||'—'}c</td>
    <td>${esc(r.premMode||'inherit')}</td><td>${r.premTolPct??'—'}%</td>
    <td>${esc(r.dateMode||'inherit')}</td>
    <td>${r.requireInsurerMatch===false?'No':r.requireInsurerMatch?'Yes':'inherit'}</td>
    <td>${r.scoreFullMatch??'—'}/${r.scorePartial??'—'}</td>
    <td style="color:var(--text3);font-size:11px">${esc(r.notes||'—')}</td>
    <td><button class="btn btn-ghost btn-sm" onclick="editRule(${i})">✎</button> <button class="btn btn-ghost btn-sm" style="color:var(--red)" onclick="deleteRule(${i})">✕</button></td>
    </tr>`).join('');
}

/* ═══════════════════════════════════════════════════════════
   STATEMENT LOAD
═══════════════════════════════════════════════════════════ */
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

/* ═══════════════════════════════════════════════════════════
   MATCHING ENGINE v6
═══════════════════════════════════════════════════════════ */
function buildSK(sRow, rule){
  const rawPol=String(sRow[stmtMapping.polno]||'');
  const rawEnd=String(sRow[stmtMapping.endno]||'');
  const rawCst=String(sRow[stmtMapping.custname]||'');
  const sIns  =String(sRow[stmtMapping.insurer]||'');
  const tot   =toNum(sRow[stmtMapping.totprem])??0;
  const od    =toNum(sRow[stmtMapping.odprem])??0;
  const tp    =toNum(sRow[stmtMapping.tpprem])??0;
  const startV=sRow[stmtMapping.startdate];
  const endV  =sRow[stmtMapping.enddate];
  const comm  =toNum(sRow[stmtMapping.commission]);

  const pt =trimPolno(rawPol,rule.stmtPolTrimDir||'left',rule.stmtPolTrimN||21,rule.removeColon!==false).toUpperCase();
  const et =trimPolno(rawEnd,rule.stmtPolTrimDir||'left',rule.stmtPolTrimN||21,rule.removeColon!==false).toUpperCase();
  const fc =freshCustName(rawCst);
  const ds =toExcelSerial(startV);
  const dsStr=ds!=null?String(ds):'';

  return {rawPol,rawEnd,rawCst,sIns,tot,od,tp,startV,endV,comm,pt,et,fc,ds,dsStr,
    polPbst:pt+String(Math.round(tot)), polDate:pt+dsStr,
    custPbst:fc+String(Math.round(tot)), custDate:fc+dsStr,
    polOD:pt+String(Math.round(od)), polTP:pt+String(Math.round(tp))};
}

function passFilter(m,sk,rule){
  if(rule.skipZeroBrok&&isZeroBrokDue(m._brok)) return false;
  if(m._zeroPol&&!rule.zeroPolicyFallback) return false;
  if(rule.requireInsurerMatch&&sk.sIns&&m._insurer&&!insurerOk(sk.sIns,m._insurer,rule.insurerFuzzy)) return false;
  return true;
}

function evalStep(id,sig){
  const {pol,endo,prem,name,date,zp}=sig;
  const p=pol||zp; // zero-polno rows pass polno check automatically
  if(id==='s1')  return p&&endo&&prem&&name&&date;
  if(id==='s2')  return p&&endo&&prem&&name;
  if(id==='s3')  return p&&prem&&name&&date;
  if(id==='s4')  return p&&prem&&name;
  if(id==='s5')  return endo&&prem&&name&&date;
  if(id==='s6')  return endo&&prem&&name;
  if(id==='s7')  return p&&prem;
  if(id==='s8')  return endo&&prem;
  if(id==='s9')  return name&&prem&&date;
  if(id==='s10') return name&&prem;
  if(id==='s11') return pol&&!zp;
  if(id==='s12') return endo;
  return false;
}

function getCands(sk,rule){
  const seen=new Set(),out=[];
  const add=rows=>{ if(!rows) return; rows.forEach(m=>{ if(!seen.has(m)){seen.add(m);out.push(m);} }); };
  // High-confidence maps first
  add(masterMaps.byPolPbst.get(sk.polPbst));
  add(masterMaps.byEndo.get(sk.et));
  add(masterMaps.byPolDate.get(sk.polDate));
  add(masterMaps.byPol.get(sk.pt));
  add(masterMaps.byPolOD.get(sk.polOD));
  add(masterMaps.byPolTP.get(sk.polTP));
  add(masterMaps.byCustPbst.get(sk.custPbst));
  add(masterMaps.byCustDate.get(sk.custDate));
  add(masterMaps.byCust.get(sk.fc));
  return out;
}

function runSteps(m,sk,rule,sEnd){
  const pol  = polMatch(sk.pt, m._polTrim);
  const endo = endMatch(sk.et, m._endoTrim);
  const prem = premHit(m,sk,rule);
  const name = nameSim(sk.fc,m._custClean,rule.fuzzyPct||60)>0;
  const date = dateHit(m,sk,rule,sEnd);
  const zp   = m._zeroPol&&rule.zeroPolicyFallback;
  const sig  = {pol,endo,prem,name,date,zp};
  const steps= rule.steps||FACTORY_DEFAULT.steps;
  for(const st of steps){
    if(!st.on) continue;
    if(evalStep(st.id,sig)) return {score:st.score,label:st.label,on:sig};
  }
  return null;
}

function matchOneRow(sRow){
  if(!masterMaps) buildMasterMaps();
  const rule=getEffectiveRule(String(sRow[stmtMapping.insurer]||''));
  const sk  =buildSK(sRow,rule);
  const sEnd=sRow[stmtMapping.enddate];

  // ── Zero Commission ──
  if(sk.comm!==null&&Math.abs(sk.comm)<=1) return mk(sRow,null,0,'zero-commission','Zero Commission',sk.sIns,'—');

  // ── Contra ──
  if(rule.contraCheck&&sk.pt&&stmtData){
    const grp=stmtData.filter(r=>trimPolno(String(r[stmtMapping.polno]||''),rule.stmtPolTrimDir||'left',rule.stmtPolTrimN||21,rule.removeColon!==false).toUpperCase()===sk.pt);
    if(grp.length>1){
      const sumC=grp.reduce((s,r)=>s+(toNum(r[stmtMapping.commission])||0),0);
      if(Math.abs(sumC)<=(rule.contraThreshold??1)){
        // Try to find Cno anyway
        const cands=getCands(sk,rule).filter(m=>passFilter(m,sk,rule));
        for(const m of cands){
          const hit=runSteps(m,sk,rule,sEnd);
          if(hit) return {...mk(sRow,m,hit.score,'contra','Contra',sk.sIns,hit.label),_matchedOn:buildOn(hit.on),_score:hit.score};
        }
        return mk(sRow,null,0,'contra','Contra',sk.sIns,'—');
      }
    }
  }

  // ── Main waterfall ──
  const cands=getCands(sk,rule).filter(m=>passFilter(m,sk,rule));
  let best=null,bestScore=0,bestHit=null;
  for(const m of cands){
    const hit=runSteps(m,sk,rule,sEnd);
    if(hit&&hit.score>bestScore){ bestScore=hit.score; best=m; bestHit=hit; }
  }

  if(!best) return mk(sRow,null,5,'unmatched','ENF',sk.sIns,'—');

  const status=bestScore>=(rule.scoreFullMatch??85)?'100':bestScore>=(rule.scorePartial??45)?'partial':'unmatched';
  return {
    _stmtRow:sRow, _cno:String(best._cno||''), _instno:String(best._instno??''),
    _score:bestScore, _matchedOn:buildOn(bestHit.on), _stepLabel:bestHit.label,
    _matchStatus:status, _insurer:sk.sIns, _ruleKey:sk.sIns?norm(sk.sIns).slice(0,18):'default',
  };
}

function buildOn(sig){
  if(!sig) return '—';
  const p=[]; if(sig.pol||sig.zp) p.push('Pol'); if(sig.endo) p.push('Endo');
  if(sig.prem) p.push('Prem'); if(sig.name) p.push('Name'); if(sig.date) p.push('Date');
  return p.join('+');
}

function mk(sRow,m,score,status,mo,ins,step){
  return {_stmtRow:sRow,_cno:m?String(m._cno||''):'',_instno:m?String(m._instno??''):'',
    _score:score,_matchedOn:mo,_stepLabel:step||'',_matchStatus:status,_insurer:ins,_ruleKey:'default'};
}

/* ═══════════════════════════════════════════════════════════
   RUN MATCHING
═══════════════════════════════════════════════════════════ */
async function runMatching(){
  if(!masterData||!stmtData){ alert('Upload both files first'); return; }
  document.querySelectorAll('#stmt-mapping-grid select').forEach((s,i)=>{ if(s.value) stmtMapping[STMT_FIELDS[i].key]=s.value; });
  document.getElementById('stmt-mapping-area').style.display='none';
  document.getElementById('matching-progress').style.display='block';

  const log=document.getElementById('match-log'),fill=document.getElementById('prog-fill'),lbl=document.getElementById('prog-label');
  const addLog=(msg,cls='')=>{ const sp=document.createElement('span'); sp.className=cls; sp.textContent=msg+'\n'; log.appendChild(sp); log.scrollTop=log.scrollHeight; };

  masterMaps=null; buildMasterMaps();
  matchResults=[];
  const total=stmtData.length, rule=defaultRule||FACTORY_DEFAULT;

  addLog(`Processing ${total} statement rows…`,'log-ok');
  addLog(`Master: ${masterData.length} records | ${masterMaps.byPol.size} unique polno keys`,'log-ok');
  addLog(`Stmt trim: ${rule.stmtPolTrimDir} ${rule.stmtPolTrimN}c | Master trim: ${rule.masterPolTrimDir} ${rule.masterPolTrimN}c`,'log-ok');
  addLog(`Prem mode: ${rule.premMode} ±${rule.premTolPct}% | Date: ${rule.dateMode} | Insurer: ${rule.requireInsurerMatch?'required':'skip'}`,'log-ok');
  addLog(`Steps active: ${(rule.steps||[]).filter(s=>s.on).length}/12`,'log-ok');
  if(insurerRules.length) addLog(`Insurer overrides: ${insurerRules.map(r=>r.insurer).join(', ')}`,'log-ok');

  let cM=0,cP=0,cU=0,cZ=0,cC=0;
  for(let i=0;i<total;i+=10){
    const batch=stmtData.slice(i,Math.min(i+10,total));
    fill.style.width=Math.round(i/total*100)+'%';
    lbl.textContent=`Row ${i+1}–${Math.min(i+10,total)} of ${total}`;
    for(const row of batch){
      const r=matchOneRow(row); matchResults.push(r);
      if(r._matchStatus==='100') cM++;
      else if(r._matchStatus==='partial') cP++;
      else if(r._matchStatus==='zero-commission') cZ++;
      else if(r._matchStatus==='contra') cC++;
      else cU++;
    }
    await new Promise(res=>setTimeout(res,4));
  }
  fill.style.width='100%'; lbl.textContent='Done!';
  addLog(`✅ Matched: ${cM} | Partial: ${cP} | ENF: ${cU} | Zero Comm: ${cZ} | Contra: ${cC}`,'log-ok');
  if(cU>0) addLog(`⚠ ${cU} rows unmatched — check polno trim, insurer name, premium mode`,'log-warn');

  const badge=document.getElementById('nav-results-badge');
  badge.textContent=matchResults.length; badge.style.display='inline-block';
  setTimeout(()=>{ document.getElementById('matching-progress').style.display='none'; switchTab('results'); },900);
}

/* ═══════════════════════════════════════════════════════════
   RESULTS
═══════════════════════════════════════════════════════════ */
function renderResultsTable(){
  const filter=document.getElementById('results-filter').value;
  const rows=matchResults.filter(r=>filter==='all'||r._matchStatus===filter);
  const cnt=s=>matchResults.filter(r=>r._matchStatus===s).length;
  document.getElementById('r-total').textContent=matchResults.length;
  document.getElementById('r-100').textContent=cnt('100');
  document.getElementById('r-partial').textContent=cnt('partial');
  document.getElementById('r-unmatched').textContent=cnt('unmatched');
  const rz=document.getElementById('r-zero'); if(rz) rz.textContent=cnt('zero-commission');
  const rc=document.getElementById('r-contra'); if(rc) rc.textContent=cnt('contra');

  const visCols=stmtCols.slice(0,4);
  const th=[...visCols,'C.No','Inst No','Score','Matched On','Step','Status'].map(c=>`<th>${esc(String(c))}</th>`).join('');
  const tb=rows.slice(0,400).map(r=>{
    const cells=visCols.map(c=>`<td title="${esc(String(r._stmtRow[c]??''))}">${esc(String(r._stmtRow[c]??'').substring(0,22))}</td>`).join('');
    const st=r._matchStatus;
    const badge2=st==='100'?'<span class="badge green">✓ Matched</span>':st==='partial'?'<span class="badge amber">Partial</span>':
      st==='zero-commission'?'<span class="badge" style="background:rgba(99,102,241,.15);color:#818cf8">Zero Comm</span>':
      st==='contra'?'<span class="badge" style="background:rgba(168,85,247,.15);color:#c084fc">Contra</span>':
      '<span class="badge red">ENF</span>';
    const sc=r._score, spCls=sc>=85?'green':sc>=50?'':'amber';
    const sp=`<span class="score-pill ${spCls}">${sc}</span>`;
    const mo=r._matchedOn?`<span class="match-on-tag">${esc(r._matchedOn)}</span>`:'—';
    const sl=r._stepLabel?`<span class="rule-tag">${esc(r._stepLabel.replace(/ \+ /g,'+').replace(/Prem/g,'Pr').replace(/Name/g,'N').replace(/Date/g,'D'))}</span>`:'—';
    return `<tr>${cells}<td><strong>${esc(r._cno)}</strong></td><td>${esc(r._instno)}</td><td>${sp}</td><td>${mo}</td><td>${sl}</td><td>${badge2}</td></tr>`;
  }).join('');
  document.getElementById('results-table').innerHTML=`<table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table>`;
}

function downloadResults(){
  if(!matchResults.length){ alert('No results'); return; }
  try{
    const wb=XLSX.utils.book_new();
    const build=arr=>arr.map(r=>{
      const obj={}; stmtCols.forEach(c=>{ obj[c]=r._stmtRow[c]!==undefined?r._stmtRow[c]:''; });
      obj['C.No']=r._cno; obj['Inst No']=r._instno; obj['Match Score']=r._score;
      obj['Matched On']=r._matchedOn||''; obj['Match Step']=r._stepLabel||'';
      obj['Rule Applied']=r._ruleKey||'default';
      const st=r._matchStatus;
      obj['Match Status']=st==='100'?'Matched':st==='partial'?'Partial Match':st==='zero-commission'?'Zero Commission':st==='contra'?'Contra':'ENF';
      return obj;
    });
    const add=(rows,name)=>{ XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows.length?rows:[{Note:'No data'}]),name); };
    add(build(matchResults),'All Results');
    add(build(matchResults.filter(r=>r._matchStatus==='100')),'Matched');
    add(build(matchResults.filter(r=>r._matchStatus==='partial')),'Partial Match');
    add(build(matchResults.filter(r=>r._matchStatus==='unmatched')),'ENF');
    add(build(matchResults.filter(r=>r._matchStatus==='zero-commission')),'Zero Commission');
    add(build(matchResults.filter(r=>r._matchStatus==='contra')),'Contra');
    XLSX.writeFile(wb,'commission_matched_'+new Date().toISOString().slice(0,10)+'.xlsx');
  }catch(err){ alert('Download error: '+err.message); }
}

/* ═══════════════════════════════════════════════════════════
   CONSOLIDATED
═══════════════════════════════════════════════════════════ */
function loadConso(){ try{ const c=localStorage.getItem(SK_CONSO); if(c) consoData=JSON.parse(c); }catch{} }
function saveConso(){ try{ localStorage.setItem(SK_CONSO,JSON.stringify(consoData)); }catch{ alert('Storage full'); } }
function addToConsolidated(){
  if(!matchResults.length){ alert('No results to add'); return; }
  const label=document.getElementById('conso-stmt-label').value.trim()||'Statement '+new Date().toLocaleDateString('en-IN');
  const flatRows=matchResults.map(r=>{
    const obj={}; stmtCols.forEach(c=>{ obj[c]=r._stmtRow[c]!==undefined?r._stmtRow[c]:''; });
    obj['C.No']=r._cno; obj['Inst No']=r._instno; obj['Match Score']=r._score;
    obj['Matched On']=r._matchedOn||''; obj['Match Step']=r._stepLabel||'';
    const st=r._matchStatus;
    obj['Match Status']=st==='100'?'Matched':st==='partial'?'Partial Match':st==='zero-commission'?'Zero Commission':st==='contra'?'Contra':'ENF';
    obj['_Statement']=label; return obj;
  });
  consoData.push({label,date:new Date().toISOString(),rows:flatRows}); saveConso();
  const total=consoData.reduce((s,b)=>s+b.rows.length,0);
  const badge=document.getElementById('nav-conso-badge'); badge.textContent=total; badge.style.display='inline-block';
  document.getElementById('conso-stmt-label').value=''; showToast(`✅ ${flatRows.length} rows added`); switchTab('conso');
}
function renderConsoPage(){
  const allRows=consoData.flatMap(b=>b.rows), has=allRows.length>0;
  document.getElementById('conso-empty').style.display=has?'none':'block';
  document.getElementById('conso-content').style.display=has?'block':'none';
  if(!has) return;
  document.getElementById('c-total').textContent=allRows.length.toLocaleString();
  document.getElementById('c-batches').textContent=consoData.length;
  document.getElementById('c-100').textContent=allRows.filter(r=>r['Match Status']==='Matched').length.toLocaleString();
  document.getElementById('c-review').textContent=allRows.filter(r=>r['Match Status']!=='Matched').length.toLocaleString();
  const badge=document.getElementById('nav-conso-badge'); badge.textContent=allRows.length; badge.style.display='inline-block';
  document.getElementById('conso-batch-list').innerHTML=consoData.map((b,i)=>`
    <div class="batch-item"><div><div class="batch-label">${esc(b.label)}</div><div class="batch-meta">${b.rows.length} rows · ${new Date(b.date).toLocaleDateString('en-IN',{day:'numeric',month:'short',year:'numeric'})}</div></div>
    <div class="batch-actions"><span class="badge green">${b.rows.filter(r=>r['Match Status']==='Matched').length} matched</span>
    <span class="badge amber">${b.rows.filter(r=>r['Match Status']==='Partial Match').length} partial</span>
    <span class="badge red">${b.rows.filter(r=>r['Match Status']==='ENF').length} ENF</span>
    <button class="btn btn-ghost btn-sm" style="color:var(--red)" onclick="removeBatch(${i})">✕ Remove</button></div></div>`).join('');
  renderConsoTable();
}
function renderConsoTable(){
  const filter=document.getElementById('conso-filter').value, allRows=consoData.flatMap(b=>b.rows);
  const rows=filter==='all'?allRows:allRows.filter(r=>{ if(filter==='100') return r['Match Status']==='Matched'; if(filter==='partial') return r['Match Status']==='Partial Match'; if(filter==='unmatched') return r['Match Status']==='ENF'; return true; });
  if(!rows.length){ document.getElementById('conso-table').innerHTML='<div style="padding:20px;text-align:center;color:var(--text2)">No rows</div>'; return; }
  const cols=Object.keys(rows[0]).filter(c=>!c.startsWith('_')).slice(0,9);
  const th=[...cols,'Match Status'].map(c=>`<th>${esc(String(c))}</th>`).join('');
  const tb=rows.slice(0,500).map(r=>{ const cells=cols.map(c=>`<td>${esc(String(r[c]??''))}</td>`).join(''); const s=r['Match Status'];
    const b2=s==='Matched'?'<span class="badge green">✓</span>':s==='Partial Match'?'<span class="badge amber">Partial</span>':'<span class="badge red">ENF</span>';
    return `<tr>${cells}<td>${b2}</td></tr>`; }).join('');
  document.getElementById('conso-table').innerHTML=`<table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table>`;
}
function removeBatch(i){ if(!confirm(`Remove "${consoData[i].label}"?`)) return; consoData.splice(i,1); saveConso(); renderConsoPage(); showToast('Batch removed'); }
function clearConso(){ if(!confirm('Clear ALL consolidated data?')) return; consoData=[]; saveConso(); document.getElementById('nav-conso-badge').style.display='none'; renderConsoPage(); showToast('Cleared'); }
function downloadConso(){
  const allRows=consoData.flatMap(b=>b.rows); if(!allRows.length){ alert('No data'); return; }
  try{
    const wb=XLSX.utils.book_new();
    const clean=rows=>rows.map(r=>{ const o={}; Object.keys(r).filter(k=>!k.startsWith('_')).forEach(k=>{ o[k]=r[k]; }); return o; });
    const add=(rows,name)=>{ XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows.length?rows:[{Note:'No data'}]),name); };
    add(clean(allRows),'All Consolidated'); add(clean(allRows.filter(r=>r['Match Status']==='Matched')),'Matched');
    add(clean(allRows.filter(r=>r['Match Status']==='Partial Match')),'Partial Match');
    add(clean(allRows.filter(r=>r['Match Status']==='ENF')),'ENF');
    consoData.forEach(b=>{ add(clean(b.rows),b.label.slice(0,28).replace(/[\\\/\?\*\[\]]/g,'_')); });
    XLSX.writeFile(wb,'consolidated_commission_'+new Date().toISOString().slice(0,10)+'.xlsx');
  }catch(err){ alert('Download error: '+err.message); }
}

/* ── Preview ── */
function renderPreviewTable(cid,rows,cols){
  const w=document.getElementById(cid);
  const th=cols.map(c=>`<th>${esc(String(c))}</th>`).join('');
  const tb=rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(String(r[c]??''))}</td>`).join('')}</tr>`).join('');
  w.innerHTML=`<table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table>`;
}

/* ── Templates ── */
function downloadMasterTemplate(){
  try{
    const wb=XLSX.utils.book_new();
    const hdr=['C.No','Inst No','Policy No','Endorsement No','Customer Name','Insurer Name','Department','Policy Type','OD Premium','TP Premium','Total Premium','Share %','Sum Insured','Policy Start Date','Endo Date','Policy End Date','Brok Due'];
    const data=[hdr,
      ['415880',0,'100051203900','','GAURAV FATESARIA','ICICI LOMBARD','HEALTH','FAMILY FLOATER',0,0,24922,100,500000,'16/03/2026','','15/03/2029',5106],
      ['415880','E1','100051203900','100051203900','GAURAV FATESARIA','ICICI LOMBARD','HEALTH','Endorsement',500,0,500,100,'','20/04/2026','20/04/2026','15/03/2029',102],
      ['168522','E12','5006/134308872/00/015','5006/134308872/00/015','TECHNOELECTRIC ENGG CO','HDFC ERGO','MARINE','CARGO',258,0,258,100,750000,'08/06/2017','26/05/2025','07/03/2025',40],
      ['400123',0,'MOT/2024/001234','','Rajesh Kumar','New India Assurance','MOTOR','Package',8500,2200,10700,100,500000,'01/04/2024','','31/03/2025',1070],
    ];
    const ws=XLSX.utils.aoa_to_sheet(data); ws['!cols']=hdr.map(()=>({wch:20}));
    XLSX.utils.book_append_sheet(wb,ws,'Master Data'); XLSX.writeFile(wb,'master_policy_template.xlsx');
  }catch(err){ alert('Error: '+err.message); }
}
function downloadStmtTemplate(){
  try{
    const wb=XLSX.utils.book_new();
    const hdr=['Policy No','Endorsement No','Customer Name','Insurer Name','OD Premium','TP Premium','Total Premium','Policy Start Date','Policy End Date','Commission Amount'];
    const data=[hdr,
      ['4225/1000512039/00/0000','4225/1000512039/00/0000','GAURAV FATESARIA','ICICI LOMBARD',0,0,24922,'16/03/2026','15/03/2029',5106],
      ['4225/1000512039/00/0001','4225/1000512039/00/0001','GAURAV FATESARIA','ICICI LOMBARD',500,0,500,'20/04/2026','15/03/2029',102],
      ['5006/134308872/00/015','5006/134308872/00/015','TECHNOELECTRIC ENGG CO','HDFC ERGO',258,0,258,'26/05/2025','07/03/2025',40],
      ['MOT/2024/001234','','Rajesh Kumar','New India Assurance',8500,2200,10700,'01/04/2024','31/03/2025',1070],
    ];
    const ws=XLSX.utils.aoa_to_sheet(data); ws['!cols']=hdr.map(()=>({wch:22}));
    XLSX.utils.book_append_sheet(wb,ws,'Commission Statement'); XLSX.writeFile(wb,'commission_statement_template.xlsx');
  }catch(err){ alert('Error: '+err.message); }
}

/* ── Export/Import ── */
function exportRules(){
  const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([JSON.stringify({defaultRule,insurerRules},null,2)],{type:'application/json'}));
  a.download='commission_rules_v6_'+new Date().toISOString().slice(0,10)+'.json'; a.click();
}
function importRulesFile(inp){
  const file=inp.files[0]; if(!file) return;
  const reader=new FileReader();
  reader.onload=e=>{ try{
    const d=JSON.parse(e.target.result);
    if(d.defaultRule){ defaultRule={...FACTORY_DEFAULT,...d.defaultRule,steps:mergeSteps(d.defaultRule.steps)}; saveDefaultRuleToStorage(); renderDefaultRuleForm(); }
    if(d.insurerRules&&Array.isArray(d.insurerRules)){ insurerRules=d.insurerRules; saveRules(); renderRulesTable(); }
    showToast('✅ Rules imported');
  }catch(err){ alert('Import error: '+err.message); } };
  reader.readAsText(file);
}

/* ═══════════════════════════════════════════════════════════
   INIT
═══════════════════════════════════════════════════════════ */
function init(){
  loadRules(); loadConso();
  if(consoData.length){ const t=consoData.reduce((s,b)=>s+b.rows.length,0); const b=document.getElementById('nav-conso-badge'); b.textContent=t; b.style.display='inline-block'; }
  try{
    const saved=localStorage.getItem(SK_MASTER), savedM=localStorage.getItem(SK_MASTER_M);
    if(saved&&savedM){
      masterData=JSON.parse(saved); masterMapping=JSON.parse(savedM); window.masterMapping=masterMapping;
      masterCols=Object.keys(masterData[0]||{}); precomputeMaster();
      document.getElementById('master-upload-area').style.display='none';
      document.getElementById('master-saved-banner').style.display='block';
      document.getElementById('master-saved-info').textContent=masterData.length.toLocaleString()+' records loaded.';
      showMasterReady();
    }
  }catch(e){ console.warn('Could not restore master data',e); }
}
init();
