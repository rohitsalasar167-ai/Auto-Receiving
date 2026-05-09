/* ═══════════════════════════════════════════
   CommissionMatcher v3 — app.js
   Features:
   - Master data persisted in localStorage
   - Per-insurer custom match rules (localStorage)
   - Consolidated statement (localStorage)
   - Improved matching with insurer rules
   ═══════════════════════════════════════════ */

/* ── Storage Keys ── */
const SK_MASTER   = 'cm_master_data';
const SK_MASTER_M = 'cm_master_mapping';
const SK_RULES    = 'cm_insurer_rules';
const SK_CONSO    = 'cm_consolidated';

/* ── State ── */
let masterData    = null;
let masterCols    = [];
let masterMapping = {};
let stmtData      = null;
let stmtCols      = [];
let stmtMapping   = {};
let matchResults  = [];
let insurerRules  = [];   // [{insurer, polnoMatch, premTol, dateCheck, primary, secondary, notes}]
let consoData     = [];   // [{label, date, rows:[...matchResult flat objects]}]

/* ── Field definitions ── */
const MASTER_FIELDS = [
  {key:'cno',       label:'C.No (Control No)',              req:true},
  {key:'instno',    label:'Inst No (0=base, >0=endo)',      req:true},
  {key:'polno',     label:'Policy No',                      req:false},
  {key:'endno',     label:'Endorsement No',                 req:false},
  {key:'custname',  label:'Customer Name',                  req:false},
  {key:'insurer',   label:'Insurer Name',                   req:false},
  {key:'dept',      label:'Department',                     req:false},
  {key:'ptype',     label:'Policy Type',                    req:false},
  {key:'odprem',    label:'OD Premium',                     req:false},
  {key:'tpprem',    label:'TP Premium',                     req:false},
  {key:'totprem',   label:'Total Premium',                  req:false},
  {key:'sumins',    label:'Sum Insured',                    req:false},
  {key:'startdate', label:'Policy Start Date',              req:false},
  {key:'enddate',   label:'Policy End Date',                req:false},
  {key:'brokdue',   label:'Brok Due (Commission Due)',      req:false},
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

/* ── Helpers ── */
function norm(s){return String(s===null||s===undefined?'':s).toLowerCase().replace(/[\s\-\/\.\,]+/g,'').trim();}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function toNum(v){const n=parseFloat(String(v===null||v===undefined?'':v).replace(/[^0-9.\-]/g,''));return isNaN(n)?null:n;}

function premClose(a,b,tolPct=2){
  const na=toNum(a),nb=toNum(b);
  if(na===null||nb===null)return false;
  if(na===0&&nb===0)return true;
  if(na===0||nb===0)return false;
  return Math.abs(na-nb)/Math.max(Math.abs(na),Math.abs(nb))*100 <= tolPct;
}

function isZeroBrokDue(v){const n=toNum(v);return n!==null&&Math.abs(n)<=5;}

function nameMatch(a,b){
  const na=norm(a),nb=norm(b);
  if(!na||!nb)return false;
  if(na===nb)return true;
  const words=na.split(/\s+/).filter(w=>w.length>2);
  if(!words.length)return false;
  return words.filter(w=>nb.includes(w)).length/words.length>=0.6;
}

function isEndoRow(v){const s=norm(v);return !(!s||s==='0'||s==='na'||s==='nil'||s==='none');}

function extractYear(v){
  if(!v&&v!==0)return null;
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
  if((mSY===null&&mEY===null)||(sSY===null&&sEY===null))return true;
  if(mSY!==null&&sSY!==null&&mSY===sSY)return true;
  if(mEY!==null&&sEY!==null&&mEY===sEY)return true;
  if(mSY!==null&&sEY!==null&&mSY===sEY)return true;
  if(mEY!==null&&sSY!==null&&mEY===sSY)return true;
  return false;
}

function polnoMatch(sPolno,mPolno,mode='exact'){
  if(!sPolno||!mPolno)return false;
  if(mode==='exact')return sPolno===mPolno;
  if(mode==='contains')return mPolno.includes(sPolno)||sPolno.includes(mPolno);
  return false;
}

function showToast(msg,duration=2800){
  const t=document.getElementById('toast');
  t.textContent=msg;t.classList.add('show');
  setTimeout(()=>t.classList.remove('show'),duration);
}

/* ── Auto-detect column hints ── */
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
  const grid=document.getElementById(gridId);grid.innerHTML='';
  fields.forEach(f=>{
    const d=document.createElement('div');
    const opts=cols.map(c=>`<option value="${esc(c)}"${(mapping[f.key]||'')===c?' selected':''}>${esc(c)}</option>`).join('');
    const req=f.req?` <span style="color:var(--red)">*</span>`:'';
    d.innerHTML=`<div><label>${esc(f.label)}${req}</label><select onchange="window['${storeVar}']['${f.key}']=this.value"><option value="">-- not mapped --</option>${opts}</select></div>`;
    grid.appendChild(d);
  });
}

/* ── Tab switching ── */
const TABS=['master','rules','match','results','conso','samples','howrules'];

function switchTab(name){
  TABS.forEach(id=>{
    const sec=document.getElementById('tab-'+id);
    const btn=document.querySelector(`.nav-item[data-tab="${id}"]`);
    if(sec)sec.classList.toggle('active',id===name);
    if(btn)btn.classList.toggle('active',id===name);
  });
  if(name==='match'){
    const has=!!masterData;
    document.getElementById('match-no-master').style.display=has?'none':'block';
    document.getElementById('match-main-area').style.display=has?'block':'none';
  }
  if(name==='results'){
    const has=matchResults.length>0;
    document.getElementById('results-empty').style.display=has?'none':'block';
    document.getElementById('results-content').style.display=has?'block':'none';
    if(has)renderResultsTable();
  }
  if(name==='conso'){
    renderConsoPage();
  }
}

document.querySelectorAll('.nav-item[data-tab]').forEach(btn=>{
  btn.addEventListener('click',()=>switchTab(btn.dataset.tab));
});

/* ════════════════════════════════
   MASTER DATA
════════════════════════════════ */
window.masterMapping=masterMapping;
window.stmtMapping=stmtMapping;

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
      masterMapping=autoDetect(masterCols,MASTER_FIELDS);
      window.masterMapping=masterMapping;
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

  // Save to localStorage
  try{
    localStorage.setItem(SK_MASTER,JSON.stringify(masterData));
    localStorage.setItem(SK_MASTER_M,JSON.stringify(masterMapping));
  }catch(e){console.warn('localStorage save failed',e);}

  showMasterReady();
  showToast('✅ Master data saved to browser storage');
}

function showMasterReady(){
  document.getElementById('master-mapping-area').style.display='none';
  document.getElementById('master-upload-area').style.display='none';
  document.getElementById('master-ready-area').style.display='block';
  document.getElementById('master-saved-banner').style.display='none';

  const base=masterData.filter(r=>{const v=r[masterMapping.instno];return norm(v)==='0'||v===0||norm(v)==='';}).length;
  const skip=masterMapping.brokdue?masterData.filter(r=>isZeroBrokDue(r[masterMapping.brokdue])).length:0;
  document.getElementById('stat-total').textContent=masterData.length.toLocaleString();
  document.getElementById('stat-base').textContent=base.toLocaleString();
  document.getElementById('stat-endo').textContent=(masterData.length-base).toLocaleString();
  document.getElementById('stat-skip').textContent=skip.toLocaleString();

  const badge=document.getElementById('nav-master-badge');
  badge.textContent=masterData.length.toLocaleString();badge.style.display='inline-block';

  const dot=document.getElementById('storage-status');
  dot.innerHTML='<span class="sdot green"></span> Master data saved ('+masterData.length+' records)';

  renderPreviewTable('master-preview-table',masterData.slice(0,6),masterCols.slice(0,7));
}

function resetMaster(){
  if(!confirm('Remove master data from browser? You will need to re-upload.'))return;
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

/* ════════════════════════════════
   INSURER RULES
════════════════════════════════ */
function loadRules(){
  try{const r=localStorage.getItem(SK_RULES);if(r)insurerRules=JSON.parse(r);}catch{}
  renderRulesTable();
}

function saveRules(){
  try{localStorage.setItem(SK_RULES,JSON.stringify(insurerRules));}catch{}
}

function addInsurerRule(){
  const insurer=document.getElementById('rule-insurer').value.trim();
  if(!insurer){alert('Insurer name is required');return;}
  const rule={
    insurer,
    polnoMatch:document.getElementById('rule-polno').value,
    premTol:parseFloat(document.getElementById('rule-prem-tol').value)||2,
    dateCheck:document.getElementById('rule-date').value,
    primary:document.getElementById('rule-primary').value,
    secondary:document.getElementById('rule-secondary').value,
    notes:document.getElementById('rule-notes').value.trim(),
  };
  // Replace if exists
  const idx=insurerRules.findIndex(r=>norm(r.insurer)===norm(insurer));
  if(idx>=0)insurerRules[idx]=rule;else insurerRules.push(rule);
  saveRules();renderRulesTable();
  document.getElementById('rule-insurer').value='';
  document.getElementById('rule-notes').value='';
  showToast('✅ Rule saved for '+insurer);
}

function deleteRule(idx){
  insurerRules.splice(idx,1);saveRules();renderRulesTable();showToast('Rule deleted');
}

function renderRulesTable(){
  const empty=document.getElementById('rules-empty');
  const wrap=document.getElementById('rules-table-wrap');
  const hint=document.getElementById('rules-count-hint');
  const badge=document.getElementById('nav-rules-badge');

  if(!insurerRules.length){
    empty.style.display='block';wrap.style.display='none';
    hint.textContent='— none yet, using defaults for all insurers';
    badge.style.display='none';return;
  }
  empty.style.display='none';wrap.style.display='block';
  hint.textContent=`— ${insurerRules.length} rule${insurerRules.length>1?'s':''} saved`;
  badge.textContent=insurerRules.length;badge.style.display='inline-block';

  const tbody=document.getElementById('rules-tbody');
  tbody.innerHTML=insurerRules.map((r,i)=>`
    <tr>
      <td><strong>${esc(r.insurer)}</strong></td>
      <td>${esc(r.polnoMatch)}</td>
      <td>${r.premTol}%</td>
      <td>${r.dateCheck==='yes'?'Yes':'No'}</td>
      <td>${esc(r.primary)}</td>
      <td>${esc(r.secondary)}</td>
      <td style="color:var(--text2)">${esc(r.notes||'—')}</td>
      <td><button class="btn btn-ghost btn-sm" style="color:var(--red)" onclick="deleteRule(${i})">✕ Delete</button></td>
    </tr>`).join('');
}

function getInsurerRule(insurerName){
  if(!insurerName)return null;
  const n=norm(insurerName);
  return insurerRules.find(r=>n.includes(norm(r.insurer))||norm(r.insurer).includes(n))||null;
}

/* ════════════════════════════════
   STATEMENT LOAD
════════════════════════════════ */
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
      stmtMapping=autoDetect(stmtCols,STMT_FIELDS);
      window.stmtMapping=stmtMapping;
      buildMappingGrid('stmt-mapping-grid',STMT_FIELDS,stmtCols,stmtMapping,'stmtMapping');
      document.getElementById('stmt-upload-title').textContent=file.name;
      document.getElementById('stmt-dropzone').classList.add('has-file');
      document.getElementById('stmt-mapping-area').style.display='block';
    }catch(err){alert('Error: '+err.message);}
  };
  reader.readAsBinaryString(file);
}

/* ════════════════════════════════
   MATCHING ENGINE
════════════════════════════════ */
function matchOneRow(sRow){
  const sPolno  =norm(sRow[stmtMapping.polno]);
  const sEndno  =norm(sRow[stmtMapping.endno]);
  const sCust   =String(sRow[stmtMapping.custname]||'');
  const sTotPrem=sRow[stmtMapping.totprem];
  const sOdPrem =sRow[stmtMapping.odprem];
  const sStart  =sRow[stmtMapping.startdate];
  const sEnd    =sRow[stmtMapping.enddate];
  const sInsurer=String(sRow[stmtMapping.insurer]||'');
  const stmtIsEnd=isEndoRow(sRow[stmtMapping.endno]);

  // Get insurer-specific rule
  const rule=getInsurerRule(sInsurer);
  const premTol   =rule?rule.premTol:2;
  const dateReq   =rule?rule.dateCheck==='yes':true;
  const polnoMode =rule?rule.polnoMatch:'exact';

  const pool=masterData.filter(m=>{
    const mInst=m[masterMapping.instno];
    const mIsBase=norm(mInst)==='0'||mInst===0||norm(mInst)==='';
    if(stmtIsEnd&&mIsBase)return false;
    if(!stmtIsEnd&&!mIsBase)return false;
    if(masterMapping.brokdue&&isZeroBrokDue(m[masterMapping.brokdue]))return false;
    return true;
  });

  let best=null,bestLevel=0,bestOn=[];

  for(const m of pool){
    const mPolno  =norm(m[masterMapping.polno]);
    const mEndno  =norm(m[masterMapping.endno]);
    const mCust   =String(m[masterMapping.custname]||'');
    const mTotPrem=m[masterMapping.totprem];
    const mOdPrem =m[masterMapping.odprem];
    const mStart  =m[masterMapping.startdate];
    const mEnd    =m[masterMapping.enddate];

    const polHit =polnoMode!=='ignore'?polnoMatch(sPolno,mPolno,polnoMode):false;
    const endHit =!!(stmtIsEnd&&sEndno&&mEndno&&sEndno===mEndno);
    const premHit=premClose(sTotPrem,mTotPrem,premTol)||premClose(sOdPrem,mOdPrem,premTol);
    const custHit=nameMatch(sCust,mCust);
    const dateOk =dateReq?yearMatch(mStart,mEnd,sStart,sEnd):true;

    if(!dateOk)continue;

    // If custom rule defines primary/secondary fields, use those for scoring
    let level=0,on=[];
    if(rule&&rule.primary!=='polno'){
      // Custom primary field logic
      const primaryHit=
        rule.primary==='endno'?endHit:
        rule.primary==='custname'?custHit:polHit;
      const secondaryHit=
        rule.secondary==='totprem'?premHit:
        rule.secondary==='polno'?polHit:
        rule.secondary==='custname'?custHit:false;

      if(primaryHit&&custHit&&premHit){level=100;on=[rule.primary,'CustName','Premium',dateReq?'Year':''];}
      else if(primaryHit&&premHit){level=80;on=[rule.primary,'Premium',dateReq?'Year':''];}
      else if(primaryHit&&secondaryHit){level=70;on=[rule.primary,rule.secondary,dateReq?'Year':''];}
      else if(primaryHit){level=50;on=[rule.primary,dateReq?'Year':''];}
    } else {
      // Default logic
      if(polHit&&custHit&&premHit){level=100;on=['PolNo','CustName','Premium',dateReq?'Year':''];}
      else if(polHit&&premHit)    {level=80; on=['PolNo','Premium',dateReq?'Year':''];}
      else if(polHit&&endHit)     {level=70; on=['PolNo','EndoNo',dateReq?'Year':''];}
      else if(endHit&&premHit)    {level=65; on=['EndoNo','Premium',dateReq?'Year':''];}
      else if(polHit)             {level=50; on=['PolNo',dateReq?'Year':''];}
      else if(custHit&&premHit&&!mPolno){level=45;on=['CustName','Premium',dateReq?'Year':''];}
    }
    on=on.filter(Boolean);
    if(level>bestLevel){bestLevel=level;best=m;bestOn=on;}
  }

  const status=bestLevel===100?'100':bestLevel>=45?'partial':'unmatched';
  return{
    _stmtRow:sRow,
    _cno:     best?String(best[masterMapping.cno]??''):'',
    _instno:  best?String(best[masterMapping.instno]??''):'',
    _matchLevel:bestLevel,
    _matchStatus:status,
    _matchedOn:bestOn.join(' + '),
    _insurer:sInsurer,
    _ruleApplied:rule?rule.insurer:'default',
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
  const skip=masterMapping.brokdue?masterData.filter(r=>isZeroBrokDue(r[masterMapping.brokdue])).length:0;
  addLog(`Processing ${total} statement rows…`,'log-ok');
  addLog(`Master: ${masterData.length} records (${skip} excluded — Brok Due ≈ 0)`,'log-warn');
  if(insurerRules.length)addLog(`Custom rules active for: ${insurerRules.map(r=>r.insurer).join(', ')}`,'log-ok');

  let c100=0,cP=0,cU=0;
  for(let i=0;i<total;i+=10){
    const batch=stmtData.slice(i,Math.min(i+10,total));
    fill.style.width=Math.round((i/total)*100)+'%';
    lbl.textContent=`Row ${i+1}–${Math.min(i+10,total)} of ${total}`;
    for(const row of batch){
      const r=matchOneRow(row);matchResults.push(r);
      if(r._matchStatus==='100')c100++;
      else if(r._matchStatus==='partial')cP++;
      else cU++;
    }
    await new Promise(res=>setTimeout(res,5));
  }
  fill.style.width='100%';
  lbl.textContent='Done!';
  addLog(`100% matched: ${c100}  |  Partial: ${cP}  |  Unmatched: ${cU}`,'log-ok');
  if(cU>0)addLog(`${cU} rows need manual review`,'log-warn');

  const badge=document.getElementById('nav-results-badge');
  badge.textContent=matchResults.length;badge.style.display='inline-block';

  setTimeout(()=>{
    document.getElementById('matching-progress').style.display='none';
    switchTab('results');
  },900);
}

/* ════════════════════════════════
   RESULTS TABLE
════════════════════════════════ */
function renderResultsTable(){
  const filter=document.getElementById('results-filter').value;
  const rows=matchResults.filter(r=>filter==='all'||r._matchStatus===filter);
  document.getElementById('r-total').textContent=matchResults.length;
  document.getElementById('r-100').textContent=matchResults.filter(r=>r._matchStatus==='100').length;
  document.getElementById('r-partial').textContent=matchResults.filter(r=>r._matchStatus==='partial').length;
  document.getElementById('r-unmatched').textContent=matchResults.filter(r=>r._matchStatus==='unmatched').length;

  const visCols=stmtCols.slice(0,5);
  const th=[...visCols,'C.No','Inst No','Matched On','Rule','Status'].map(c=>`<th>${esc(String(c))}</th>`).join('');
  const tb=rows.slice(0,300).map(r=>{
    const cells=visCols.map(c=>`<td>${esc(String(r._stmtRow[c]??''))}</td>`).join('');
    const badge=r._matchStatus==='100'?'<span class="badge green">100% Match</span>':r._matchStatus==='partial'?'<span class="badge amber">Partial</span>':'<span class="badge red">Unmatched</span>';
    const moText=r._matchedOn?`<span style="font-size:11px;font-family:'DM Mono',monospace;color:var(--text3)">${esc(r._matchedOn)}</span>`:'—';
    const ruleText=`<span style="font-size:11px;color:var(--text3)">${esc(r._ruleApplied)}</span>`;
    return `<tr>${cells}<td><strong>${esc(r._cno)}</strong></td><td>${esc(r._instno)}</td><td>${moText}</td><td>${ruleText}</td><td>${badge}</td></tr>`;
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
      obj['Matched On']=r._matchedOn||'';
      obj['Rule Applied']=r._ruleApplied||'default';
      obj['Match Status']=r._matchStatus==='100'?'100% Match':r._matchStatus==='partial'?'Partial Match':'Unmatched';
      return obj;
    });
    const addSheet=(rows,name)=>{const ws=XLSX.utils.json_to_sheet(rows.length?rows:[{Note:'No data'}]);XLSX.utils.book_append_sheet(wb,ws,name);};
    addSheet(build(matchResults),'All Results');
    addSheet(build(matchResults.filter(r=>r._matchStatus==='100')),'100% Matched');
    addSheet(build(matchResults.filter(r=>r._matchStatus==='partial')),'Partial Match');
    addSheet(build(matchResults.filter(r=>r._matchStatus==='unmatched')),'Unmatched');
    XLSX.writeFile(wb,'commission_matched_'+new Date().toISOString().slice(0,10)+'.xlsx');
  }catch(err){alert('Download error: '+err.message);}
}

/* ════════════════════════════════
   CONSOLIDATED STATEMENT
════════════════════════════════ */
function loadConso(){
  try{const c=localStorage.getItem(SK_CONSO);if(c)consoData=JSON.parse(c);}catch{}
}

function saveConso(){
  try{localStorage.setItem(SK_CONSO,JSON.stringify(consoData));}catch(e){alert('Storage full — try clearing old consolidated data.');}
}

function addToConsolidated(){
  if(!matchResults.length){alert('No results to add');return;}
  const label=document.getElementById('conso-stmt-label').value.trim()||
    'Statement '+new Date().toLocaleDateString('en-IN');

  // Flatten results for storage
  const flatRows=matchResults.map(r=>{
    const obj={};
    stmtCols.forEach(c=>{obj[c]=r._stmtRow[c]!==undefined?r._stmtRow[c]:'';});
    obj['C.No']=r._cno;obj['Inst No']=r._instno;
    obj['Matched On']=r._matchedOn||'';
    obj['Rule Applied']=r._ruleApplied||'default';
    obj['Match Status']=r._matchStatus==='100'?'100% Match':r._matchStatus==='partial'?'Partial Match':'Unmatched';
    obj['_Statement']=label;
    return obj;
  });

  consoData.push({label,date:new Date().toISOString(),rows:flatRows});
  saveConso();

  const badge=document.getElementById('nav-conso-badge');
  const total=consoData.reduce((s,b)=>s+b.rows.length,0);
  badge.textContent=total;badge.style.display='inline-block';

  document.getElementById('conso-stmt-label').value='';
  showToast(`✅ ${flatRows.length} rows added to Consolidated`);
  switchTab('conso');
}

function renderConsoPage(){
  const allRows=consoData.flatMap(b=>b.rows);
  const hasData=allRows.length>0;
  document.getElementById('conso-empty').style.display=hasData?'none':'block';
  document.getElementById('conso-content').style.display=hasData?'block':'none';
  if(!hasData)return;

  document.getElementById('c-total').textContent=allRows.length.toLocaleString();
  document.getElementById('c-batches').textContent=consoData.length;
  document.getElementById('c-100').textContent=allRows.filter(r=>r['Match Status']==='100% Match').length.toLocaleString();
  document.getElementById('c-review').textContent=allRows.filter(r=>r['Match Status']!=='100% Match').length.toLocaleString();

  const badge=document.getElementById('nav-conso-badge');
  badge.textContent=allRows.length;badge.style.display='inline-block';

  // Batch list
  const batchList=document.getElementById('conso-batch-list');
  batchList.innerHTML=consoData.map((b,i)=>`
    <div class="batch-item">
      <div>
        <div class="batch-label">${esc(b.label)}</div>
        <div class="batch-meta">${b.rows.length} rows · Added ${new Date(b.date).toLocaleDateString('en-IN',{day:'numeric',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'})}</div>
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
  const cols=Object.keys(rows[0]).filter(c=>!c.startsWith('_')).slice(0,8);
  const th=[...cols,'Match Status'].map(c=>`<th>${esc(String(c))}</th>`).join('');
  const tb=rows.slice(0,500).map(r=>{
    const cells=cols.map(c=>`<td>${esc(String(r[c]??''))}</td>`).join('');
    const status=r['Match Status'];
    const badge=status==='100% Match'?'<span class="badge green">100% Match</span>':status==='Partial Match'?'<span class="badge amber">Partial</span>':'<span class="badge red">Unmatched</span>';
    return `<tr>${cells}<td>${badge}</td></tr>`;
  }).join('');
  document.getElementById('conso-table').innerHTML=`<table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table>`;
}

function removeBatch(idx){
  if(!confirm(`Remove "${consoData[idx].label}" from consolidated?`))return;
  consoData.splice(idx,1);saveConso();renderConsoPage();showToast('Batch removed');
}

function clearConso(){
  if(!confirm('Clear ALL consolidated data? This cannot be undone.'))return;
  consoData=[];saveConso();
  document.getElementById('nav-conso-badge').style.display='none';
  renderConsoPage();showToast('Consolidated data cleared');
}

function downloadConso(){
  const allRows=consoData.flatMap(b=>b.rows);
  if(!allRows.length){alert('No data to download');return;}
  try{
    const wb=XLSX.utils.book_new();
    const cleanRows=allRows.map(r=>{const o={};Object.keys(r).filter(k=>!k.startsWith('_')).forEach(k=>{o[k]=r[k];});return o;});
    const addSheet=(rows,name)=>{const ws=XLSX.utils.json_to_sheet(rows.length?rows:[{Note:'No data'}]);XLSX.utils.book_append_sheet(wb,ws,name);};
    addSheet(cleanRows,'All Consolidated');
    addSheet(cleanRows.filter(r=>r['Match Status']==='100% Match'),'100% Matched');
    addSheet(cleanRows.filter(r=>r['Match Status']==='Partial Match'),'Partial Match');
    addSheet(cleanRows.filter(r=>r['Match Status']==='Unmatched'),'Unmatched');
    // Per-statement sheets
    consoData.forEach(b=>{
      const name=b.label.slice(0,28).replace(/[\\\/\?\*\[\]]/g,'_');
      const clean=b.rows.map(r=>{const o={};Object.keys(r).filter(k=>!k.startsWith('_')).forEach(k=>{o[k]=r[k];});return o;});
      addSheet(clean,name);
    });
    XLSX.writeFile(wb,'consolidated_commission_'+new Date().toISOString().slice(0,10)+'.xlsx');
  }catch(err){alert('Download error: '+err.message);}
}

/* ════════════════════════════════
   PREVIEW TABLE
════════════════════════════════ */
function renderPreviewTable(cid,rows,cols){
  const w=document.getElementById(cid);
  const th=cols.map(c=>`<th>${esc(String(c))}</th>`).join('');
  const tb=rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(String(r[c]??''))}</td>`).join('')}</tr>`).join('');
  w.innerHTML=`<table><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table>`;
}

/* ════════════════════════════════
   SAMPLE TEMPLATES
════════════════════════════════ */
function downloadMasterTemplate(){
  try{
    const wb=XLSX.utils.book_new();
    const hdr=['C.No','Inst No','Policy No','Endorsement No','Customer Name','Insurer Name','Department','Policy Type','OD Premium','TP Premium','Total Premium','Sum Insured','Policy Start Date','Policy End Date','Brok Due'];
    const data=[hdr,
      ['CNT001',0,'MOT/2024/001234','','Rajesh Kumar','New India Assurance','Motor','Package Policy',8500,2200,10700,500000,'01/04/2024','31/03/2025',1070],
      ['CNT002',1,'MOT/2024/001234','END/001','Rajesh Kumar','New India Assurance','Motor','Endorsement',500,0,500,'','15/06/2024','31/03/2025',50],
      ['CNT003',0,'FIR/2024/005678','','Priya Sharma','Oriental Insurance','Fire','IAR',0,0,25000,2000000,'01/07/2024','30/06/2025',2500],
      ['CNT004',0,'HLT/2024/009012','','Mehta Enterprises','Star Health','Health','Group Mediclaim',0,0,45000,1000000,'01/04/2024','31/03/2025',4500],
      ['CNT005',0,'MOT/2024/007890','','Anita Desai','Bajaj Allianz','Motor','Third Party',0,3500,3500,0,'15/05/2024','14/05/2025',200],
      ['CNT006',2,'HLT/2024/009012','END/002','Mehta Enterprises','Star Health','Health','Endorsement',0,0,5000,50000,'01/08/2024','31/03/2025',500],
      ['CNT007',0,'MAR/2024/003344','','Cargo Movers Ltd','ICICI Lombard','Marine','Cargo',0,0,18000,750000,'01/06/2024','31/05/2025',1800],
      ['CNT008',0,'TRV/2024/008877','','Sunita Patel','HDFC Ergo','Travel','International',0,0,3200,0,'10/09/2024','25/09/2024',320],
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
      ['FIR/2024/005678','','Priya Sharma','Oriental Insurance',0,0,25000,'01/07/2024','30/06/2025',2500],
      ['HLT/2024/009012','','Mehta Enterprises','Star Health',0,0,45000,'01/04/2024','31/03/2025',4500],
      ['MOT/2024/007890','','Anita Desai','Bajaj Allianz',0,3500,3500,'15/05/2024','14/05/2025',200],
    ];
    const ws=XLSX.utils.aoa_to_sheet(data);ws['!cols']=hdr.map(()=>({wch:22}));
    XLSX.utils.book_append_sheet(wb,ws,'Commission Statement');
    XLSX.writeFile(wb,'commission_statement_template.xlsx');
  }catch(err){alert('Error: '+err.message);}
}

/* ════════════════════════════════
   INIT — load from localStorage
════════════════════════════════ */
function init(){
  // Load insurer rules
  loadRules();

  // Load consolidated
  loadConso();
  if(consoData.length){
    const total=consoData.reduce((s,b)=>s+b.rows.length,0);
    const badge=document.getElementById('nav-conso-badge');
    badge.textContent=total;badge.style.display='inline-block';
  }

  // Load master data
  try{
    const saved=localStorage.getItem(SK_MASTER);
    const savedM=localStorage.getItem(SK_MASTER_M);
    if(saved&&savedM){
      masterData=JSON.parse(saved);
      masterMapping=JSON.parse(savedM);
      window.masterMapping=masterMapping;
      masterCols=Object.keys(masterData[0]||{});

      document.getElementById('master-upload-area').style.display='none';
      document.getElementById('master-saved-banner').style.display='block';
      document.getElementById('master-saved-info').textContent=
        `${masterData.length.toLocaleString()} records loaded.`;

      showMasterReady();
    }
  }catch(e){console.warn('Could not load master from storage',e);}
}

init();
