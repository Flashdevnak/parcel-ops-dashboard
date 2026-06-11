
window.PARCEL_OPS_BUILD = 'v21-chartjs-and-redraw-fix';
console.info('Parcel Operations Dashboard build:', window.PARCEL_OPS_BUILD);


const $ = (id) => document.getElementById(id);
const page = document.body.dataset.page || 'overview';
const els = {
  fileInput:$('fileInput'), chooseFileBtn:$('chooseFileBtn'), fileName:$('fileName'), reportTime:$('reportTime'), statusCard:$('statusCard'), controls:$('controls'), content:$('content'),
  searchInput:$('searchInput'), baseHubFilter:$('baseHubFilter'), modeFilter:$('modeFilter'), agingFilter:$('agingFilter'), hubFilter:$('hubFilter'), actionFilter:$('actionFilter'), copyLineMode:$('copyLineMode'), copyLineLimit:$('copyLineLimit'), copyReportFocusGroup:$('copyReportFocusGroup'), copyTrackMode:$('copyTrackMode'),
  managerSearch:$('managerSearch'), managerChecklist:$('managerChecklist'), managerSelectedCount:$('managerSelectedCount'), selectVisibleManagers:$('selectVisibleManagers'), clearManagers:$('clearManagers'), resetBtn:$('resetBtn'), copyBtn:$('copyBtn'), exportBtn:$('exportBtn'), exportXlsxBtn:$('exportXlsxBtn'), toast:$('toast'),
  summaryContext:$('summaryContext'), kpiGrid:$('kpiGrid'), hubCount:$('hubCount'), hubSummaryTable:$('hubSummaryTable'), branchSummaryTable:$('branchSummaryTable'), fdCount:$('fdCount'), lhOtherTable:$('lhOtherTable'), lhOtherCount:$('lhOtherCount'),
  rowCount:$('rowCount'), detailTable:$('detailTable'), showAllCols:$('showAllCols'), cleanThCodes:$('cleanThCodes'), editMode:$('editMode'), compactRows:$('compactRows'), sortFilter:$('sortFilter'), editCount:$('editCount'), prevPage:$('prevPage'), nextPage:$('nextPage'), pageInfo:$('pageInfo'),
  exportContext:$('exportContext'), exportXlsxBtn2:$('exportXlsxBtn2'), exportBtn2:$('exportBtn2'), copyBtn2:$('copyBtn2'), clearDataBtn:$('clearDataBtn'),
  chartContext:$('chartContext'), statusTypeCount:$('statusTypeCount'), branchAffectedCount:$('branchAffectedCount'), lhHubAffectedCount:$('lhHubAffectedCount'), topStatusCount:$('topStatusCount'), topStatusLabel:$('topStatusLabel'), statusSummaryTable:$('statusSummaryTable'), fdBranchStatusTable:$('fdBranchStatusTable'), lhHubStatusTable:$('lhHubStatusTable'), statusChart:$('statusChart'), branchChart:$('branchChart'), lhChart:$('lhChart')
};
const AGING_ORDER = ['อยู่ในคลังไม่เกิน 6 ชั่วโมง','อยู่ในคลังไม่เกิน 12 ชั่วโมง','อยู่ในคลังไม่เกิน 22 ชั่วโมง','อยู่ในคลังระหว่าง 22–24 ชั่วโมง','อยู่ในคลังเกิน 24 ชั่วโมง','อยู่ในคลังมากกว่า 48 ชั่วโมง','ตรวจสอบข้อมูลเวลามาถึง'];
const AGING_SHORT = ['≤6 ชม.','≤12 ชม.','≤22 ชม.','22–24 ชม.','>24 ชม.','>48 ชม.','ตรวจสอบ'];
const FOCUS_COLS = ['เลขพัสดุ','เวลาอยู่ภายในฮับ','ชั่วโมงในคลัง','เวลาที่ใช้คำนวณ','เวลาถึงตามจริง','เวลาที่ดำเนินการล่าสุด','สถานะพัสดุ','การดำเนินการล่าสุด','HUBปลายทาง','สาขาปลายทาง','น้ำหนัก','ประเภทลูกค้า','ผู้ดำเนินการล่าสุด','หมายเลขโทรศัพท์ผู้จัดการสาขาที่ดำเนินการครั้งสุดท้าย'];
const STORAGE_KEY = 'parcelOpsDashboardGitReadyFinal';
const DB_NAME = 'ParcelOpsDashboardStorage';
const DB_STORE = 'workspaces';
const DB_ID = 'current-workspace-git-ready-final';
let state = { headers:[], rows:[], filtered:[], managerValues:[], selectedManagers:new Set(), page:1, pageSize:80, edits:0, fileName:'' };
let chartRefs = {status:null, branch:null, lh:null};
let saveTimer = null;

init();
async function init(){
  setDefaultReportTime();
  bind();
  if(new URLSearchParams(location.search).has('reset')){
    await wipeAllClientData(false);
    if(els.fileName) els.fileName.textContent='ยังไม่ได้เลือกไฟล์ Excel';
    if(els.controls) els.controls.hidden=true;
    if(els.content) els.content.hidden=true;
    setStatus('เริ่มหน้าใหม่แล้ว','เลือกไฟล์ Excel เพื่อเริ่มตรวจงาน');
    history.replaceState(null,'',location.pathname);
    return;
  }
  await restoreFromSession();
}
function bind(){
  els.chooseFileBtn?.addEventListener('click',()=>els.fileInput?.click());
  els.clearDataBtn?.addEventListener('click',clearStoredData);
  els.fileInput?.addEventListener('change',handleFile);
  [els.reportTime,els.searchInput,els.baseHubFilter,els.modeFilter,els.agingFilter,els.hubFilter,els.actionFilter,els.sortFilter,els.copyLineMode,els.copyLineLimit,els.copyTrackMode].forEach(el=>el?.addEventListener('input',()=>{state.page=1;applyFilters();saveLightState();}));
  [els.showAllCols,els.cleanThCodes,els.compactRows].forEach(el=>el?.addEventListener('change',()=>{renderAll();saveLightState();}));
  els.editMode?.addEventListener('change',renderDetail);
  els.managerSearch?.addEventListener('input',renderManagerChecklist);
  els.managerChecklist?.addEventListener('change',onManagerChange);
  els.selectVisibleManagers?.addEventListener('click',()=>{document.querySelectorAll('#managerChecklist input').forEach(cb=>state.selectedManagers.add(cb.value));state.page=1;renderManagerChecklist();applyFilters();});
  els.clearManagers?.addEventListener('click',()=>{state.selectedManagers.clear();state.page=1;renderManagerChecklist();applyFilters();});
  els.resetBtn?.addEventListener('click',resetFilters);
  els.copyBtn?.addEventListener('click',copySummary); els.copyBtn2?.addEventListener('click',copySummary);
  els.exportBtn?.addEventListener('click',exportCsv); els.exportBtn2?.addEventListener('click',exportCsv);
  els.exportXlsxBtn?.addEventListener('click',exportWorkbook); els.exportXlsxBtn2?.addEventListener('click',exportWorkbook);
  els.prevPage?.addEventListener('click',()=>{state.page=Math.max(1,state.page-1);renderDetail();});
  els.nextPage?.addEventListener('click',()=>{state.page+=1;renderDetail();});
  els.detailTable?.addEventListener('blur',onCellEdit,true);
  els.detailTable?.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();e.target.blur();}});
  document.querySelectorAll('[data-capture-target]').forEach(btn=>btn.addEventListener('click',()=>captureTarget(btn.dataset.captureTarget, btn.dataset.captureName || 'capture')));
  document.querySelectorAll('[data-copy-table]').forEach(btn=>btn.addEventListener('click',()=>{try{copyTableForLine(btn.dataset.copyTable, btn.dataset.copyTitle || 'ตารางรายงาน');}catch(e){console.error('copyTableForLine failed',e);toast(`คัดลอกไม่สำเร็จ: ${e?.message || 'ตรวจ Console'}`);}}));
  document.querySelectorAll('[data-copy-records-report]').forEach(btn=>btn.addEventListener('click',()=>{try{copyRecordsReportForLine();}catch(e){console.error('copyRecordsReportForLine failed',e);toast(`คัดลอกไม่สำเร็จ: ${e?.message || 'ตรวจ Console'}`);}}));
  document.querySelectorAll('input[name="copyReportFocus"]').forEach(cb=>cb.addEventListener('change',onReportFocusCheckboxChange));
  bindNavigationPersistence();
  bindSpaNavigation();
  window.addEventListener('pagehide', persistWorkspaceNow);
  document.addEventListener('visibilitychange',()=>{ if(document.visibilityState==='hidden') persistWorkspaceNow(); });
}
function setDefaultReportTime(){ const d = new Date(); d.setSeconds(0,0); if(els.reportTime && !els.reportTime.value) els.reportTime.value = toLocalInputValue(d); }


const PAGE_HASH_MAP = {
  '#overview':'overview',
  '#status':'status',
  '#fd':'fd',
  '#lh':'lh',
  '#records':'records',
  '#export':'export'
};
function currentPageKey(){
  return PAGE_HASH_MAP[location.hash] || pageKeyFromPath(location.pathname) || 'overview';
}
function pageKeyFromPath(path){
  const file = String(path || '').split('/').pop() || 'index.html';
  if(file === 'status.html') return 'status';
  if(file === 'fd.html') return 'fd';
  if(file === 'lh.html') return 'lh';
  if(file === 'records.html') return 'records';
  if(file === 'export.html') return 'export';
  return 'overview';
}
function setActivePage(page){
  const key = page || currentPageKey();
  const panels = document.querySelectorAll('[data-page-panel]');
  panels.forEach(panel=>{
    panel.classList.toggle('active', panel.dataset.pagePanel === key);
  });
  document.querySelectorAll('a[href^="#"]').forEach(a=>{
    const linkKey = PAGE_HASH_MAP[a.getAttribute('href')] || 'overview';
    a.classList.toggle('active', linkKey === key);
    if(linkKey === key) a.setAttribute('aria-current','page'); else a.removeAttribute('aria-current');
  });
  try{ sessionStorage.setItem(STORAGE_KEY + '_activePage', key); }catch(e){}
  if(state.rows && state.rows.length){
    try{ renderAll(); }catch(e){ console.error('render after SPA page switch failed', e); toast(`เปลี่ยนเมนูไม่สำเร็จ: ${e?.message || 'ตรวจ Console'}`); }
  }
}
function bindSpaNavigation(){
  document.querySelectorAll('a[href^="#"]').forEach(a=>{
    if(a.dataset.spaBound === '1') return;
    a.dataset.spaBound = '1';
    a.dataset.pageLink = '1';
    a.addEventListener('click', (ev)=>{
      const href = a.getAttribute('href') || '#overview';
      const key = PAGE_HASH_MAP[href] || 'overview';
      ev.preventDefault();
      try{ persistWorkspaceNow(); }catch(e){ console.warn('persist before SPA click failed', e); }
      if(location.hash !== href){
        history.pushState(null, '', href);
      }
      setActivePage(key);
      if(key === 'status' && state.rows && state.rows.length){
        requestAnimationFrame(()=>setTimeout(renderStatusVisuals, 150));
      }
      window.scrollTo({top:0, behavior:'auto'});
    });
  });
  window.addEventListener('popstate',()=>setActivePage(currentPageKey()));
  setActivePage(currentPageKey());
}
function currentWorkspacePayload(){
  return {
    headers:state.headers,
    rows:state.rows,
    fileName:state.fileName,
    selectedManagers:[...state.selectedManagers],
    edits:state.edits,
    controls:readControls(),
    savedAt:new Date().toISOString()
  };
}
function persistWorkspaceNow(){
  if(!state.rows || !state.rows.length) return;
  const payload = currentWorkspacePayload();

  // metadata เล็ก ๆ ไว้บอกว่ามีข้อมูลจำอยู่
  try{
    localStorage.setItem(STORAGE_KEY + '_meta', JSON.stringify({
      fileName:payload.fileName,
      rows:payload.rows.length,
      headers:payload.headers.length,
      selectedManagers:payload.selectedManagers,
      edits:payload.edits,
      controls:payload.controls,
      savedAt:payload.savedAt
    }));
  }catch(e){}

  // sessionStorage ใช้สำหรับเปลี่ยนหน้าใน tab เดียวกัน ถ้าไฟล์ไม่ใหญ่มาก
  try{ sessionStorage.setItem(STORAGE_KEY, JSON.stringify(payload)); }catch(e){}

  // IndexedDB เป็นตัวหลัก เก็บไฟล์ใหญ่ได้ดีกว่า localStorage
  try{ dbPut(DB_ID, payload).catch(err=>console.warn('async page navigation save failed', err)); }catch(e){}
}
function bindNavigationPersistence(){
  document.querySelectorAll('a[href$=".html"]').forEach(a=>{
    if(a.dataset.persistBound === '1') return;
    a.dataset.persistBound = '1';
    a.addEventListener('click', async (ev)=>{
      const href = a.getAttribute('href');
      if(!href || href.startsWith('http') || href.startsWith('#')) return;
      if(state.rows && state.rows.length){
        ev.preventDefault();
        try{
          clearTimeout(saveTimer);
          persistWorkspaceNow();
          await saveSession();
        }catch(e){
          console.warn('save before navigation failed', e);
        }
        location.href = href;
      }
    });
  });
}

async function handleFile(e){
  const file = e.target.files?.[0]; if(!file) return;
  setStatus('กำลังอ่านไฟล์', file.name);
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf,{type:'array',cellDates:true});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:''});
  const headerIdx = findHeaderRow(aoa);
  const headers = aoa[headerIdx].map(h=>cleanText(h));
  const rows=[];
  for(let i=headerIdx+1;i<aoa.length;i++){
    const arr=aoa[i]; if(!arr || arr.every(v=>cleanText(v)==='')) continue;
    const r={}; headers.forEach((h,idx)=>r[h]=arr[idx]??''); r.__id = `${Date.now()}_${i}_${Math.random().toString(16).slice(2)}`; rows.push(r);
  }
  state.headers=headers; state.rows=rows; state.fileName=file.name; state.selectedManagers.clear(); state.edits=0; state.page=1;
  els.fileName.textContent=file.name;
  setupComputed(); populateFilters(); applyFilters(); persistWorkspaceNow(); await saveSession();
  if(els.fileInput) els.fileInput.value='';
  setStatus('อ่านไฟล์แล้ว', `${num(rows.length)} รายการ / ${num(headers.length)} คอลัมน์ • ใช้ร่วมกันทุกหน้าแล้ว`);
}
function findHeaderRow(aoa){
  let best=0, score=-1;
  aoa.slice(0,30).forEach((row,i)=>{ const s=(row||[]).map(cleanText).join('|'); let p=0; ['เลขพัสดุ','สถานะพัสดุ','HUBปลายทาง','สาขาปลายทาง','เวลาที่ดำเนินการล่าสุด'].forEach(k=>{ if(s.includes(k)) p+=2; }); if((row||[]).filter(x=>cleanText(x)).length>score+p){best=i;score=(row||[]).filter(x=>cleanText(x)).length+p;} });
  return best;
}
function setupComputed(){
  const report = getReportDate();
  state.rows.forEach(r=>{
    const calc = pickCalcDate(r);
    r['เวลาที่ใช้คำนวณ'] = calc ? formatDateTime(calc) : '';
    r['ชั่วโมงในคลัง'] = calc ? round2((report - calc)/36e5) : '';
    r['เวลาอยู่ภายในฮับ'] = agingLabel(Number(r['ชั่วโมงในคลัง']));
    r.__aging = r['เวลาอยู่ภายในฮับ'];
    r.__managerKey = managerKey(r);
    r.__search = Object.keys(r).filter(k=>!k.startsWith('__')).map(k=>String(r[k]??'')).join(' ').toLowerCase();
  });
}
function pickCalcDate(r){ const f = parseDateValue(r['เวลาถึงตามจริง']); const k = parseDateValue(r['เวลาที่ดำเนินการล่าสุด']); if(cleanText(r['เวลาถึงตามจริง'])==='--') return k; return f || k; }
function parseDateValue(v){
  if(v instanceof Date && !isNaN(v)) return v;
  let s=cleanText(v); if(!s || s==='--') return null;
  const m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if(m){ let y=+m[3]; if(y<100) y+=2000; return new Date(y,+m[2]-1,+m[1],+m[4],+m[5],+(m[6]||0)); }
  const d = new Date(s); return isNaN(d)?null:d;
}
function agingLabel(h){ if(!isFinite(h)) return AGING_ORDER[6]; if(h<=6) return AGING_ORDER[0]; if(h<=12) return AGING_ORDER[1]; if(h<=22) return AGING_ORDER[2]; if(h<=24) return AGING_ORDER[3]; if(h<=48) return AGING_ORDER[4]; return AGING_ORDER[5]; }
function managerKey(r){ const vals=['ผู้ดำเนินการล่าสุด','หมายเลขโทรศัพท์ผู้จัดการสาขาที่ดำเนินการครั้งสุดท้าย','รหัสผู้จัดการสาขาที่ดำเนินการครั้งสุดท้าย','ผู้จัดการสาขาที่ดำเนินการครั้งสุดท้าย'].map(k=>cleanText(r[k])).filter(Boolean); return vals.length ? unique(vals).join(' | ') : '(ว่าง)'; }
function populateFilters(){
  if(els.controls) els.controls.hidden=false; if(els.content) els.content.hidden=false;
  fillSelect(els.agingFilter, ['all',...AGING_ORDER], v=>v==='all'?'ทุกช่วงเวลา':v);
  fillSelect(els.hubFilter, ['all',...unique(state.rows.map(r=>cleanText(r['HUBปลายทาง'])).filter(Boolean)).sort(thSort)], v=>v==='all'?'ทุก HUB':displayValue(v));
  fillSelect(els.baseHubFilter, ['all',...unique(state.rows.map(r=>cleanText(r['HUBปลายทาง'])).filter(Boolean)).sort(thSort)], v=>v==='all'?'ยังไม่เลือก':displayValue(v));
  fillSelect(els.actionFilter, ['all',...unique(state.rows.map(r=>cleanText(r['การดำเนินการล่าสุด'])).filter(Boolean)).sort(thSort)], v=>v==='all'?'ทุกสถานะ':displayValue(v));
  state.managerValues = unique(state.rows.map(r=>r.__managerKey)).sort(thSort);
  renderManagerChecklist();
}
function fillSelect(el, values, labelFn){ if(!el) return; const old=el.value; el.innerHTML=values.map(v=>`<option value="${esc(v)}">${esc(labelFn(v))}</option>`).join(''); if(values.includes(old)) el.value=old; }
function applyFilters(){
  setupComputed();
  const q=cleanText(els.searchInput?.value).toLowerCase(); const base=els.baseHubFilter?.value||'all'; const mode=els.modeFilter?.value||'all'; const aging=els.agingFilter?.value||'all'; const hub=els.hubFilter?.value||'all'; const action=els.actionFilter?.value||'all';
  state.filtered = state.rows.filter(r=>{
    if(q && !r.__search.includes(q)) return false;
    if(state.selectedManagers.size && !state.selectedManagers.has(r.__managerKey)) return false;
    if(aging!=='all' && r.__aging!==aging) return false;
    if(hub!=='all' && cleanText(r['HUBปลายทาง'])!==hub) return false;
    if(action!=='all' && cleanText(r['การดำเนินการล่าสุด'])!==action) return false;
    if(mode==='fd-base' && (base==='all' || cleanText(r['HUBปลายทาง'])!==base)) return false;
    if(mode==='lh-other' && (base==='all' || cleanText(r['HUBปลายทาง'])===base)) return false;
    if(mode==='missing-manager' && r.__managerKey!=='(ว่าง)') return false;
    return true;
  });
  sortFilteredRows();
  renderAll(); saveSession();
}
function renderAll(){ renderContext(); renderKpi(); renderSummaryTables(); renderDetail(); renderStatusVisuals(); }
function renderContext(){
  const base = (els.baseHubFilter?.value==='all'||!els.baseHubFilter) ? 'ยังไม่เลือก' : displayValue(els.baseHubFilter.value);
  const text = `เวลารายงาน ${formatDateTime(getReportDate())} • สาขา ${base} • ${num(state.filtered.length)} รายการ`;
  if(els.summaryContext) els.summaryContext.textContent=text;
  if(els.exportContext) els.exportContext.textContent=`${num(state.filtered.length)} รายการตามตัวกรอง`;
  if(els.chartContext) els.chartContext.textContent=`${num(state.filtered.length)} รายการ • ${base}`;
}
function renderKpi(){
  if(!els.kpiGrid) return; const total=state.filtered.length; const count=a=>state.filtered.filter(r=>r.__aging===a).length; const base=els.baseHubFilter?.value||'all';
  const fd = base==='all'?0:state.filtered.filter(r=>cleanText(r['HUBปลายทาง'])===base).length;
  const lh = base==='all'?0:state.filtered.filter(r=>cleanText(r['HUBปลายทาง'])!==base).length;
  const kpis=[['รายการทั้งหมด',total,'ตามเงื่อนไขปัจจุบัน','ink'],['ไม่เกิน 6 ชม.',count(AGING_ORDER[0]),'อยู่ในกรอบเริ่มต้น','ok'],['ไม่เกิน 12 ชม.',count(AGING_ORDER[1]),'ควรติดตามต่อเนื่อง','blue'],['ไม่เกิน 22 ชม.',count(AGING_ORDER[2]),'ใกล้ช่วงเร่งตรวจ','amber'],['22–24 ชม.',count(AGING_ORDER[3]),'เฝ้าระวังใกล้เกินวัน','amber'],['เกิน 24 ชม.',count(AGING_ORDER[4]),'ต้องเร่งตรวจสอบ','red'],['มากกว่า 48 ชม.',count(AGING_ORDER[5]),'ความเสี่ยงสูง','hot'],['FD / LH',`${num(fd)} / ${num(lh)}`,base==='all'?'เลือกสาขาหรือ HUB ที่ต้องการตรวจ':'FD ตรงปลายทาง / LH ปลายทางอื่น','violet']];
  els.kpiGrid.innerHTML=kpis.map(([l,v,h,c])=>`<article class="kpi ${c}"><span>${esc(l)}</span><b>${esc(v)}</b><small>${esc(h)}</small></article>`).join('');
}
function renderSummaryTables(){
  const base = els.baseHubFilter?.value || 'all';
  const fdRows = getFdRowsForSelectedBase();
  const lhRows = getLhRowsForSelectedBase();

  let hubRows;
  if(base === 'all'){
    hubRows = summarize(state.filtered, 'HUBปลายทาง');
  }else{
    const fdBranchRows = summarize(fdRows, 'สาขาปลายทาง').map(x=>({...x,name:`FD: ${x.name}`,raw:`FD: ${x.raw}`}));
    const lhHubRows = summarize(lhRows, 'HUBปลายทาง').map(x=>({...x,name:`LH: ${x.name}`,raw:`LH: ${x.raw}`}));
    hubRows = [...fdBranchRows, ...lhHubRows].sort((a,b)=>b.total-a.total||thSort(a.name,b.name));
  }

  const branchRows = summarize(fdRows, 'สาขาปลายทาง');
  const lhOtherRows = summarize(lhRows, 'HUBปลายทาง');

  if(els.hubCount){
    els.hubCount.textContent = base === 'all'
      ? `${num(hubRows.length)} HUB`
      : `${num(fdRows.length)} FD / ${num(lhRows.length)} LH`;
  }
  if(els.fdCount) els.fdCount.textContent = `${num(fdRows.length)} รายการ / ${num(branchRows.length)} สาขา`;
  if(els.lhOtherCount) els.lhOtherCount.textContent = `${num(lhRows.length)} รายการ / ${num(lhOtherRows.length)} HUB`;

  renderSummaryTable(
    els.hubSummaryTable,
    hubRows,
    base === 'all' ? 'HUB ปลายทาง' : 'FD สาขา + LH HUB อื่น'
  );
  renderSummaryTable(els.branchSummaryTable, branchRows, 'FD / สาขาปลายทาง');
  renderSummaryTable(els.lhOtherTable, lhOtherRows, 'HUB ปลายทางอื่น');
}
function summarize(rows,key){
  const map=new Map();
  rows.forEach(r=>{
    const raw=cleanText(r[key])||'(ว่าง)';
    const groupKey=normalizeGroupKey(raw);
    if(!map.has(groupKey)) map.set(groupKey,{name:displayValue(raw),raw,total:0,counts:Object.fromEntries(AGING_ORDER.map(a=>[a,0]))});
    const o=map.get(groupKey);
    o.total++;
    o.counts[r.__aging]=(o.counts[r.__aging]||0)+1;
  });
  return [...map.values()].sort((a,b)=>b.total-a.total||thSort(a.name,b.name));
}
function normalizeGroupKey(v){
  return cleanHubName(v)
    .replace(/\s+/g,' ')
    .replace(/\s*-\s*/g,'-')
    .replace(/\s*_\s*/g,'_')
    .replace(/[‐‑‒–—]/g,'-')
    .trim()
    .toUpperCase();
}
function cleanHubName(v){
  return cleanText(v)
    .replace(/^\([^)]+\)\s*/,'')
    .replace(/\s+/g,' ')
    .replace(/\s*-\s*/g,'-')
    .trim();
}

function renderSummaryTable(table,rows,firstLabel){ if(!table) return; table.innerHTML=`<thead><tr><th>${esc(firstLabel)}</th>${AGING_SHORT.slice(0,6).map(x=>`<th class="num">${x}</th>`).join('')}<th class="num">รวม</th></tr></thead><tbody>${rows.length?rows.map(r=>`<tr><td title="${esc(displayValue(r.raw))}">${esc(r.name)}</td>${AGING_ORDER.slice(0,6).map(a=>`<td class="num">${num(r.counts[a]||0)}</td>`).join('')}<td class="num total">${num(r.total)}</td></tr>`).join(''):`<tr><td colspan="8" class="empty">ไม่มีข้อมูลตามเงื่อนไข</td></tr>`}</tbody>`; }
function renderStatusVisuals(){
  if(!els.statusChart && !els.statusSummaryTable && !els.fdBranchStatusTable && !els.lhHubStatusTable) return;

  const base = els.baseHubFilter?.value || 'all';
  const fdRows = base === 'all' ? [] : state.filtered.filter(r => cleanText(r['HUBปลายทาง']) === base);
  const lhRows = base === 'all' ? [] : state.filtered.filter(r => cleanText(r['HUBปลายทาง']) !== base);

  // ใช้คอลัมน์ "การดำเนินการล่าสุด" เท่านั้น ไม่ใช้ "สถานะพัสดุ"
  const actionRows = summarize(state.filtered, 'การดำเนินการล่าสุด');
  const fdBranchRows = summarize(fdRows, 'สาขาปลายทาง');
  const lhHubRows = summarize(lhRows, 'HUBปลายทาง');

  if(els.statusTypeCount) els.statusTypeCount.textContent = num(actionRows.length);
  if(els.branchAffectedCount) els.branchAffectedCount.textContent = num(fdBranchRows.length);
  if(els.lhHubAffectedCount) els.lhHubAffectedCount.textContent = num(lhHubRows.length);
  if(els.topStatusCount) els.topStatusCount.textContent = num(actionRows[0]?.total || 0);
  if(els.topStatusLabel) els.topStatusLabel.textContent = actionRows[0]?.name || '-';

  renderSummaryTable(els.statusSummaryTable, actionRows, 'สถานะดำเนินการล่าสุด');
  renderSummaryTableWithEmpty(els.fdBranchStatusTable, fdBranchRows, 'สาขาปลายทาง FD', base === 'all' ? 'กรุณาเลือกสาขา / HUB ที่ต้องการตรวจก่อน' : 'ไม่มีข้อมูล FD ตามเงื่อนไข');
  renderSummaryTableWithEmpty(els.lhHubStatusTable, lhHubRows, 'HUB ปลายทางอื่น LH', base === 'all' ? 'กรุณาเลือกสาขา / HUB ที่ต้องการตรวจก่อน' : 'ไม่มีข้อมูล LH ตามเงื่อนไข');

  if(typeof Chart === 'undefined'){
    console.warn('Chart.js ยังไม่โหลด จึงยังไม่วาดกราฟ');
    return;
  }
  requestAnimationFrame(()=>setTimeout(()=>{
    drawStatusChart(actionRows);
    drawBranchChart(fdBranchRows.slice(0,15));
    drawLhChart(lhHubRows.slice(0,15));
  }, 80));
}
function renderSummaryTableWithEmpty(table, rows, firstLabel, emptyText){
  if(!table) return;
  if(rows.length){ renderSummaryTable(table, rows, firstLabel); return; }
  table.innerHTML = `<thead><tr><th>${esc(firstLabel)}</th>${AGING_SHORT.slice(0,6).map(x=>`<th class="num">${x}</th>`).join('')}<th class="num">รวม</th></tr></thead><tbody><tr><td colspan="8" class="empty">${esc(emptyText)}</td></tr></tbody>`;
}
function summarizeSimple(rows,key){
  const map=new Map();
  rows.forEach(r=>{ const raw=cleanText(r[key]) || '(ว่าง)'; map.set(raw,(map.get(raw)||0)+1); });
  return [...map.entries()].map(([raw,count])=>({raw,name:displayValue(raw),count})).sort((a,b)=>b.count-a.count||thSort(a.name,b.name));
}

const valueLabelPlugin = {
  id:'valueLabelPlugin',
  afterDatasetsDraw(chart){
    const {ctx, chartArea} = chart;
    const dataset = chart.data.datasets[0];
    if(!dataset) return;
    const meta = chart.getDatasetMeta(0);
    ctx.save();
    ctx.font = '700 12px system-ui, -apple-system, Segoe UI, sans-serif';
    ctx.fillStyle = '#1f2937';
    ctx.textBaseline = 'middle';
    meta.data.forEach((bar, i) => {
      const value = Number(dataset.data[i] || 0);
      if(!value) return;
      const label = compactNum(value);
      const props = bar.getProps(['x','y','base'], true);
      if(chart.options.indexAxis === 'y'){
        const x = Math.min(props.x + 8, chartArea.right - 42);
        ctx.textAlign = 'left';
        ctx.fillText(label, x, props.y);
      }else{
        ctx.textAlign = 'center';
        const y = Math.max(props.y - 10, chartArea.top + 10);
        ctx.fillText(label, props.x, y);
      }
    });
    ctx.restore();
  }
};
function compactNum(v){
  const n = Number(v || 0);
  if(n >= 1000000) return (n/1000000).toFixed(n%1000000===0?0:1)+'M';
  if(n >= 1000) return (n/1000).toFixed(n%1000===0?0:1)+'k';
  return num(n);
}
function chartOptionsWithLabels(options){
  return {...options, plugins:{...(options.plugins||{}), valueLabelPlugin:{}}};
}
function chartValue(r){ return Number(r.total ?? r.count ?? 0); }
function compactChartRows(rows, limit=12){
  const clean = rows.filter(r => chartValue(r) > 0);
  const head = clean.slice(0, limit);
  const rest = clean.slice(limit);
  if(rest.length){
    head.push({name:'อื่น ๆ', raw:'อื่น ๆ', total:rest.reduce((s,r)=>s+chartValue(r),0), count:rest.reduce((s,r)=>s+chartValue(r),0)});
  }
  return head;
}
function drawStatusChart(rows){
  if(!els.statusChart) return;
  const top = compactChartRows(rows, 12);
  chartRefs.status?.destroy?.();
  const h = Math.max(360, top.length * 38 + 90);
  els.statusChart.parentElement.style.height = h + 'px';
  chartRefs.status = new Chart(els.statusChart, {
    type:'bar',
    data:{ labels: top.map(r=>shortLabel(r.name,46)), datasets:[{ label:'จำนวนชิ้น', data: top.map(chartValue), backgroundColor:'rgba(255,184,28,.72)', borderColor:'rgba(225,138,0,.95)', borderWidth:1, borderRadius:10, maxBarThickness:30 }]},
    plugins:[valueLabelPlugin],
    options:{ indexAxis:'y', maintainAspectRatio:false, plugins:{ legend:{display:false}, tooltip:{callbacks:{title:(items)=>top[items[0].dataIndex]?.name||''}} }, scales:{ x:{ beginAtZero:true, ticks:{ color:'#475569', precision:0 }, grid:{ color:'rgba(100,116,139,.14)' } }, y:{ ticks:{ color:'#475569', autoSkip:false }, grid:{display:false} } } }
  });
}
function drawBranchChart(rows){
  if(!els.branchChart) return;
  const top = compactChartRows(rows, 15);
  chartRefs.branch?.destroy?.();
  els.branchChart.parentElement.style.height = Math.max(380, top.length * 34 + 90) + 'px';
  chartRefs.branch = new Chart(els.branchChart, {
    type:'bar',
    data:{ labels: top.map(r=>shortLabel(r.name,44)), datasets:[{ label:'จำนวนชิ้น', data: top.map(chartValue), backgroundColor:'rgba(255,184,28,.68)', borderColor:'rgba(225,138,0,.95)', borderWidth:1, borderRadius:10 }]},
    plugins:[valueLabelPlugin],
    options:{ indexAxis:'y', maintainAspectRatio:false, plugins:{ legend:{display:false}, tooltip:{callbacks:{title:(items)=>top[items[0].dataIndex]?.name||''}} }, scales:{ x:{ beginAtZero:true, ticks:{ color:'#475569', precision:0 }, grid:{ color:'rgba(100,116,139,.14)' } }, y:{ ticks:{ color:'#475569', autoSkip:false }, grid:{display:false} } } }
  });
}
function drawLhChart(rows){
  if(!els.lhChart) return;
  const top = compactChartRows(rows, 15);
  chartRefs.lh?.destroy?.();
  els.lhChart.parentElement.style.height = Math.max(380, top.length * 34 + 90) + 'px';
  chartRefs.lh = new Chart(els.lhChart, {
    type:'bar',
    data:{ labels: top.map(r=>shortLabel(r.name,44)), datasets:[{ label:'จำนวนชิ้น', data: top.map(chartValue), backgroundColor:'rgba(37,99,235,.58)', borderColor:'rgba(29,78,216,.95)', borderWidth:1, borderRadius:10 }]},
    plugins:[valueLabelPlugin],
    options:{ indexAxis:'y', maintainAspectRatio:false, plugins:{ legend:{display:false}, tooltip:{callbacks:{title:(items)=>top[items[0].dataIndex]?.name||''}} }, scales:{ x:{ beginAtZero:true, ticks:{ color:'#475569', precision:0 }, grid:{ color:'rgba(100,116,139,.14)' } }, y:{ ticks:{ color:'#475569', autoSkip:false }, grid:{display:false} } } }
  });
}
function renderDetail(){
  if(!els.detailTable) return; const cols=getDisplayColumns(); els.detailTable.classList.toggle('compact', !!els.compactRows?.checked); const max=Math.max(1,Math.ceil(state.filtered.length/state.pageSize)); if(state.page>max) state.page=max; const start=(state.page-1)*state.pageSize; const rows=state.filtered.slice(start,start+state.pageSize);
  if(els.rowCount) els.rowCount.textContent=`${num(state.filtered.length)} รายการ`; if(els.pageInfo) els.pageInfo.textContent=`หน้า ${num(state.page)} / ${num(max)}`; if(els.prevPage) els.prevPage.disabled=state.page<=1; if(els.nextPage) els.nextPage.disabled=state.page>=max; if(els.editCount) els.editCount.textContent=state.edits?`แก้ไขแล้ว ${num(state.edits)} ช่อง`:'ยังไม่มีการแก้ไข';
  const editable=!!els.editMode?.checked; els.detailTable.innerHTML=`<thead><tr>${cols.map(c=>headerHtml(c)).join('')}</tr></thead><tbody>${rows.length?rows.map(r=>`<tr>${cols.map(c=>cellHtml(c,r[c],r,editable)).join('')}</tr>`).join(''):`<tr><td colspan="${cols.length}" class="empty">ไม่มีข้อมูลตามเงื่อนไข</td></tr>`}</tbody>`;
}
function getDisplayColumns(){ const base=FOCUS_COLS.filter(c=>state.headers.includes(c)||isComputed(c)); if(els.showAllCols?.checked) return unique([...base,...state.headers]); return base; }
function cellHtml(col,val,row,editable){
  const shown=displayValue(String(val??''));
  const cls=[colClass(col), /ชั่วโมง|น้ำหนัก|COD|จำนวน|ราคา/i.test(col)?'num':'', col==='เลขพัสดุ'?'tracking-cell':'', col==='เวลาอยู่ภายในฮับ'?'aging-cell':'', (col.includes('เวลา')||col.includes('วันที่'))?'time-cell':''].filter(Boolean).join(' ');
  const body = col==='เวลาอยู่ภายในฮับ' ? `<span class="aging-badge ${agingClass(shown)}">${esc(shown)}</span>` : esc(shown);
  return `<td class="${cls} ${editable?'editable':''}" title="${esc(shown)}" ${editable?'contenteditable="true"':''} data-id="${row.__id}" data-col="${esc(col)}">${body}</td>`;
}
function onCellEdit(e){ const td=e.target.closest('td[data-id][data-col]'); if(!td) return; const r=state.rows.find(x=>x.__id===td.dataset.id); if(!r) return; const next=td.textContent.trim(); if(String(r[td.dataset.col]??'')!==next){ r[td.dataset.col]=next; state.edits++; setupComputed(); applyFilters(); saveSession(); } }
function renderManagerChecklist(){ if(!els.managerChecklist) return; const q=cleanText(els.managerSearch?.value).toLowerCase(); const counts=new Map(); state.rows.forEach(r=>counts.set(r.__managerKey,(counts.get(r.__managerKey)||0)+1)); const values=state.managerValues.filter(v=>!q||v.toLowerCase().includes(q)||displayValue(v).toLowerCase().includes(q)); els.managerChecklist.innerHTML=values.map(v=>`<label class="check"><input type="checkbox" value="${esc(v)}" ${state.selectedManagers.has(v)?'checked':''}><span title="${esc(displayValue(v))}">${esc(compactManagerLabel(v))}</span><b>${num(counts.get(v)||0)}</b></label>`).join('')||'<div class="empty">ไม่พบค่าที่ค้นหา</div>'; updateManagerCount(); }
function onManagerChange(e){ const cb=e.target.closest('input[type=checkbox]'); if(!cb) return; cb.checked?state.selectedManagers.add(cb.value):state.selectedManagers.delete(cb.value); state.page=1; updateManagerCount(); applyFilters(); }
function updateManagerCount(){ if(!els.managerSelectedCount) return; els.managerSelectedCount.textContent = state.selectedManagers.size ? `เลือก ${num(state.selectedManagers.size)} ค่า` : 'ทุกค่า'; }
function resetFilters(){ ['searchInput','managerSearch'].forEach(k=>{if(els[k]) els[k].value='';}); ['baseHubFilter','modeFilter','agingFilter','hubFilter','actionFilter','sortFilter'].forEach(k=>{if(els[k]) els[k].value=(k==='sortFilter'?'hours-desc':'all');}); state.selectedManagers.clear(); state.page=1; renderManagerChecklist(); applyFilters(); }
function copySummary(){
  if(!state.rows.length){toast('ยังไม่มีข้อมูล');return;}
  const base=selectedBaseText();
  const lines=['Parcel Operations Dashboard',`เวลารายงาน: ${formatDateTime(getReportDate())}`,`สาขา ${base}`,`เงื่อนไขผู้ดำเนินการล่าสุด: ${managerSummaryText(true)}`,`จำนวนรายการ: ${num(state.filtered.length)}`,...AGING_ORDER.slice(0,6).map(a=>`${a}: ${num(state.filtered.filter(r=>r.__aging===a).length)}`)];
  navigator.clipboard?.writeText(lines.join('\n')); toast('คัดลอกสรุปแล้ว');
}
function managerSummaryText(includeBase=false){
  const items = state.selectedManagers.size ? unique([...state.selectedManagers].map(summarySafeManager).filter(Boolean)) : ['ทุกค่า'];
  if(includeBase){ const baseRaw = els.baseHubFilter?.value; const base = baseRaw && baseRaw !== 'all' ? displayValue(baseRaw) : ''; if(base && !items.includes(base)) items.push(base); }
  if(!items.length) return state.selectedManagers.size ? `เลือก ${num(state.selectedManagers.size)} ค่า` : 'ทุกค่า';
  return items.length>8 ? `${items.slice(0,8).join(', ')} และอีก ${num(items.length-8)} ค่า` : items.join(', ');
}
function selectedBaseText(){ return els.baseHubFilter?.value==='all'?'ยังไม่เลือก':displayValue(els.baseHubFilter.value); }
function summarySafeManager(v){ return String(v).split('|').map(x=>cleanText(x)).filter(x=>x && !/^0\d{8,10}$/.test(x) && !/(นาย|นาง|นางสาว|คุณ)\s/.test(x)).join(' | '); }
function exportCsv(){ if(!state.filtered.length){toast('ไม่มีข้อมูลให้ส่งออก');return;} const cols=getExportColumns(); const csv=[cols,...state.filtered.map(r=>cols.map(c=>exportValue(r,c)))].map(row=>row.map(csvCell).join(',')).join('\n'); downloadBlob('\ufeff'+csv,`parcel_detail_${dateStamp()}.csv`,'text/csv;charset=utf-8;'); toast('ส่งออก CSV แล้ว'); }
function exportWorkbook(){
  if(typeof XLSX==='undefined'){toast('ไม่พบตัวสร้างไฟล์ Excel');return;} if(!state.filtered.length){toast('ไม่มีข้อมูลให้ส่งออก');return;}
  const wb=XLSX.utils.book_new(); const base=els.baseHubFilter?.value||'all'; const fdRows=base==='all'?[]:state.filtered.filter(r=>cleanText(r['HUBปลายทาง'])===base); const lhRows=base==='all'?[]:state.filtered.filter(r=>cleanText(r['HUBปลายทาง'])!==base);
  addSheet(wb,'Summary',buildSummary(fdRows,lhRows)); addSheet(wb,'Detail_All',rowsToAoa(state.filtered)); addSheet(wb,'FD_Selected',rowsToAoa(fdRows)); addSheet(wb,'LH_Other',rowsToAoa(lhRows)); addSheet(wb,'By_HUB',summaryToAoa(summarize(state.filtered,'HUBปลายทาง'),'HUB ปลายทาง')); addSheet(wb,'By_Branch_FD',summaryToAoa(summarize(fdRows,'สาขาปลายทาง'),'สาขาปลายทาง')); addSheet(wb,'By_Last_Action',simpleSummaryAoa(summarizeSimple(state.filtered,'การดำเนินการล่าสุด'),'สถานะดำเนินการล่าสุด')); addSheet(wb,'By_Aging',agingToAoa(state.filtered));
  XLSX.writeFile(wb,`parcel_operations_report_${dateStamp()}.xlsx`); toast('ส่งออก Excel แล้ว');
}
function buildSummary(fdRows,lhRows){ const base=selectedBaseText(); const aoa=[['Parcel Operations Dashboard'],['เวลารายงาน',formatDateTime(getReportDate())],['สาขา',base],['จำนวนรายการ',state.filtered.length],['FD เฉพาะสาขา/HUB ที่เลือก',fdRows.length],['LH เฉพาะปลายทางอื่น',lhRows.length],['จำนวนสาขาที่ติด',unique(state.filtered.map(r=>cleanText(r['สาขาปลายทาง'])).filter(Boolean)).length],['เงื่อนไขผู้ดำเนินการล่าสุด',managerSummaryText(true)],[],['กลุ่มเวลา','จำนวน']]; AGING_ORDER.slice(0,6).forEach(a=>aoa.push([a,state.filtered.filter(r=>r.__aging===a).length])); aoa.push([],['สถานะดำเนินการล่าสุด','จำนวน']); summarizeSimple(state.filtered,'การดำเนินการล่าสุด').forEach(r=>aoa.push([r.name,r.count])); return aoa; }
function rowsToAoa(rows){ const cols=getExportColumns(); return [cols,...rows.map(r=>cols.map(c=>exportValue(r,c)))]; }
function getExportColumns(){ return unique([...FOCUS_COLS.filter(c=>state.headers.includes(c)||isComputed(c)),...state.headers]); }
function exportValue(r,c){ return isComputed(c)?r[c]??'':r[c]??''; }
function summaryToAoa(rows,label){ return [[label,...AGING_SHORT.slice(0,6),'รวม'],...rows.map(r=>[r.name,...AGING_ORDER.slice(0,6).map(a=>r.counts[a]||0),r.total])]; }
function simpleSummaryAoa(rows,label){ return [[label,'จำนวน'], ...rows.map(r=>[r.name,r.count])]; }
function agingToAoa(rows){ return [['กลุ่มเวลา','จำนวน'],...AGING_ORDER.slice(0,6).map(a=>[a,rows.filter(r=>r.__aging===a).length])]; }
function addSheet(wb,name,aoa){ const ws=XLSX.utils.aoa_to_sheet(aoa); ws['!cols']=aoa[0]?.map((_,i)=>({wch:i===0?28:18}))||[]; XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31)); }
async function saveSession(){
  const payload = currentWorkspacePayload();
  try{
    await dbPut(DB_ID, payload);

    // localStorage เก็บเฉพาะ meta เพื่อไม่ชน quota เวลาไฟล์ Excel ใหญ่
    localStorage.setItem(STORAGE_KEY + '_meta', JSON.stringify({
      fileName:state.fileName,
      rows:state.rows.length,
      headers:state.headers.length,
      selectedManagers:[...state.selectedManagers],
      edits:state.edits,
      controls:readControls(),
      savedAt:payload.savedAt
    }));

    // sessionStorage สำรองเฉพาะ tab เดียวกัน ถ้าไฟล์ใหญ่เกินจะเงียบ ไม่ทำให้ระบบล้ม
    try{ sessionStorage.setItem(STORAGE_KEY, JSON.stringify(payload)); }catch(sessionFull){ console.warn('sessionStorage full, IndexedDB is primary', sessionFull); }
  }catch(e){
    console.warn('IndexedDB save failed', e);
    try{
      sessionStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
      setStatus('ใช้โหมดสำรองข้อมูลข้ามหน้า','IndexedDB ใช้ไม่ได้ แต่บันทึกผ่าน sessionStorage ในแท็บนี้แล้ว');
    }catch(localError){
      console.warn('sessionStorage backup failed', localError);
      setStatus('บันทึกข้อมูลข้ามหน้าไม่สำเร็จ','เบราว์เซอร์บล็อกพื้นที่จัดเก็บข้อมูล กรุณาอย่าเปิดแบบ private/incognito');
    }
  }
}
function saveLightState(){ clearTimeout(saveTimer); saveTimer=setTimeout(()=>saveSession(),250); }
async function restoreFromSession(){
  try{
    let data = null;

    // 1) ตอนเปลี่ยนเมนูใน tab เดียวกัน ใช้ sessionStorage ก่อน เพื่อกลับมาทันที
    try{
      const session = sessionStorage.getItem(STORAGE_KEY);
      data = session ? JSON.parse(session) : null;
    }catch(e){ data = null; }

    // 2) ถ้าไม่มี ให้ใช้ IndexedDB เป็นตัวหลัก
    if(!data || !data.rows || !data.rows.length){
      data = await dbGet(DB_ID);
    }

    // 3) fallback เก่า เผื่อคนใช้ build เดิมที่เคยเก็บใน localStorage
    if(!data || !data.rows || !data.rows.length){
      try{
        const legacy = localStorage.getItem(STORAGE_KEY);
        data = legacy ? JSON.parse(legacy) : null;
      }catch(e){ data = null; }
    }

    if(!data || !data.rows || !data.rows.length){
      setStatus('ยังไม่มีข้อมูลที่บันทึกไว้','เลือกไฟล์ Excel เพื่อเริ่มใช้งาน • build v21');
      return;
    }

    state.headers = data.headers || [];
    state.rows = data.rows || [];
    state.fileName = data.fileName || '';
    state.selectedManagers = new Set(data.selectedManagers || []);
    state.edits = data.edits || 0;
    state.page = 1;

    if(els.fileName) els.fileName.textContent = state.fileName || 'ใช้ข้อมูลร่วมจากหน้าก่อนหน้า';

    setupComputed();
    populateFilters();

    try{
      writeControls(data.controls || {});
    }catch(controlError){
      console.warn('Restore controls failed but workspace is kept', controlError);
      try{ restoreReportFocusCheckboxes(data.controls?.copyFocus || 'all'); }catch(e){}
    }

    applyFilters();
    setStatus('กำลังใช้ข้อมูลที่บันทึกไว้',`${num(state.rows.length)} รายการ • ${state.fileName||'ข้อมูลจากเบราว์เซอร์'} • ใช้ข้อมูลร่วมกันทุกหน้าแล้ว`);
  }catch(e){
    console.warn('Restore failed', e);
    setStatus('โหลดข้อมูลเดิมไม่สำเร็จ','กรุณากดล้างข้อมูลแล้วเลือกไฟล์ใหม่ หรืออย่าเปิดแบบ private/incognito');
  }
}
async function clearStoredData(){
  const ok = confirm('ต้องการล้างข้อมูลทั้งหมดและเริ่มหน้าเว็บใหม่ใช่ไหม?\nระบบจะล้างไฟล์ Excel ที่จำไว้, ตัวกรอง, IndexedDB, localStorage, sessionStorage, CacheStorage และพาไปหน้าแรก');
  if(!ok) return;

  setStatus('กำลังล้างข้อมูลทั้งหมด','กำลังล้างข้อมูลที่เบราว์เซอร์จำไว้และรีเซ็ตหน้าเว็บ');
  await wipeAllClientData(true);

  state = { headers:[], rows:[], filtered:[], managerValues:[], selectedManagers:new Set(), page:1, pageSize:80, edits:0, fileName:'' };
  if(els.fileName) els.fileName.textContent='ยังไม่ได้เลือกไฟล์ Excel';
  if(els.fileInput) els.fileInput.value='';
  if(els.controls) els.controls.hidden=true;
  if(els.content) els.content.hidden=true;
  toast('ล้างข้อมูลทั้งหมดแล้ว');

  // บังคับออกจากหน้าปัจจุบัน กลับหน้าแรก พร้อม cache-busting query
  const basePath = location.pathname.replace(/[^\/]*$/, 'index.html');
  const nextUrl = `${location.origin}${basePath}?reset=${Date.now()}`;
  setTimeout(() => location.replace(nextUrl), 350);
}

async function wipeAllClientData(includeServiceWorker=false){
  try{ await dbDelete(DB_ID); }catch(e){ console.warn('Current IndexedDB delete failed', e); }

  // ล้างฐานข้อมูลเก่าทุกเวอร์ชันของโปรเจกต์นี้
  try{
    const dbs = await indexedDB.databases?.();
    if(dbs && dbs.length){
      await Promise.all(dbs
        .filter(d => d.name && (String(d.name).includes('ParcelOpsDashboard') || String(d.name).includes('ParcelOps')))
        .map(d => new Promise(resolve => {
          const req = indexedDB.deleteDatabase(d.name);
          req.onsuccess = req.onerror = req.onblocked = () => resolve();
        })));
    }else{
      indexedDB.deleteDatabase(DB_NAME);
    }
  }catch(e){ console.warn('IndexedDB full wipe failed', e); }

  // เว็บนี้ใช้ storage ของตัวเองอย่างเดียว จึงล้างทั้ง origin ได้เลย
  try{ localStorage.clear(); }catch(e){}
  try{ sessionStorage.clear(); }catch(e){}

  // ล้าง CacheStorage เช่นไฟล์ JS/CSS/HTML ที่เคย cache ไว้
  try{
    if('caches' in window){
      const keys = await caches.keys();
      await Promise.all(keys.map(k => caches.delete(k)));
    }
  }catch(e){ console.warn('CacheStorage delete failed', e); }

  // กันกรณีเคยมี service worker จาก hosting/template อื่น
  if(includeServiceWorker){
    try{
      if('serviceWorker' in navigator){
        const regs = await navigator.serviceWorker.getRegistrations();
        await Promise.all(regs.map(reg => reg.unregister()));
      }
    }catch(e){ console.warn('Service worker unregister failed', e); }
  }
}

function openDashboardDb(){ return new Promise((resolve,reject)=>{ const req=indexedDB.open(DB_NAME,1); req.onupgradeneeded=()=>{ const db=req.result; if(!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE); }; req.onsuccess=()=>resolve(req.result); req.onerror=()=>reject(req.error); }); }
async function dbPut(key,value){ const db=await openDashboardDb(); return new Promise((resolve,reject)=>{ const tx=db.transaction(DB_STORE,'readwrite'); tx.objectStore(DB_STORE).put(value,key); tx.oncomplete=()=>{db.close();resolve();}; tx.onerror=()=>{db.close();reject(tx.error);}; }); }
async function dbGet(key){ const db=await openDashboardDb(); return new Promise((resolve,reject)=>{ const tx=db.transaction(DB_STORE,'readonly'); const req=tx.objectStore(DB_STORE).get(key); req.onsuccess=()=>resolve(req.result); req.onerror=()=>reject(req.error); tx.oncomplete=()=>db.close(); }); }
async function dbDelete(key){ const db=await openDashboardDb(); return new Promise((resolve,reject)=>{ const tx=db.transaction(DB_STORE,'readwrite'); tx.objectStore(DB_STORE).delete(key); tx.oncomplete=()=>{db.close();resolve();}; tx.onerror=()=>{db.close();reject(tx.error);}; }); }
function sortFilteredRows(){ const mode = els.sortFilter?.value || 'hours-desc'; const rows = state.filtered; if(mode==='original') return; const byHours = (a,b)=>Number(a['ชั่วโมงในคลัง']||0)-Number(b['ชั่วโมงในคลัง']||0); const byLatest = (a,b)=>(parseDateValue(a['เวลาที่ดำเนินการล่าสุด'])?.getTime()||0)-(parseDateValue(b['เวลาที่ดำเนินการล่าสุด'])?.getTime()||0); const byTrack = (a,b)=>thSort(cleanText(a['เลขพัสดุ']), cleanText(b['เลขพัสดุ'])); rows.sort((a,b)=>{ if(mode==='hours-desc') return byHours(b,a) || byLatest(a,b); if(mode==='hours-asc') return byHours(a,b) || byLatest(a,b); if(mode==='latest-desc') return byLatest(b,a) || byHours(b,a); if(mode==='latest-asc') return byLatest(a,b) || byHours(b,a); if(mode==='tracking-asc') return byTrack(a,b); return 0; }); }
function headerHtml(col){ const m = headerMeta(col); return `<th class="${colClass(col)} ${/ชั่วโมง|น้ำหนัก|COD|จำนวน|ราคา/i.test(col)?'num':''}"><span>${esc(m.title)}</span>${m.sub?`<small>${esc(m.sub)}</small>`:''}</th>`; }
function headerMeta(c){ const map={'เลขพัสดุ':['เลขพัสดุ','Tracking'],'เวลาอยู่ภายในฮับ':['กลุ่มเวลา','Aging'],'ชั่วโมงในคลัง':['ชั่วโมงค้าง','Hours'],'เวลาที่ใช้คำนวณ':['เวลาที่ใช้คำนวณ','Calc time'],'เวลาถึงตามจริง':['เวลาถึงจริง','Arrival'],'เวลาที่ดำเนินการล่าสุด':['เวลาดำเนินการ','Last action time'],'สถานะพัสดุ':['สถานะพัสดุ','Parcel status'],'การดำเนินการล่าสุด':['งานล่าสุด','Last action'],'HUBปลายทาง':['HUB ปลายทาง','Destination HUB'],'สาขาปลายทาง':['สาขาปลายทาง','Destination branch'],'น้ำหนัก':['น้ำหนัก','Weight'],'ประเภทลูกค้า':['ประเภทลูกค้า','Customer type'],'ผู้ดำเนินการล่าสุด':['ผู้ดำเนินการ','Operator'],'หมายเลขโทรศัพท์ผู้จัดการสาขาที่ดำเนินการครั้งสุดท้าย':['ผู้ดำเนินการ/เบอร์/รหัส','Operator / code']}; const x=map[c]||[shortHeader(c),'']; return {title:x[0],sub:x[1]}; }
function colClass(col){ const map={'เลขพัสดุ':'col-track','เวลาอยู่ภายในฮับ':'col-aging','ชั่วโมงในคลัง':'col-hours','เวลาที่ใช้คำนวณ':'col-calc','เวลาถึงตามจริง':'col-arrival','เวลาที่ดำเนินการล่าสุด':'col-lasttime','สถานะพัสดุ':'col-status','การดำเนินการล่าสุด':'col-action','HUBปลายทาง':'col-hub','สาขาปลายทาง':'col-branch','น้ำหนัก':'col-weight','ประเภทลูกค้า':'col-customer','ผู้ดำเนินการล่าสุด':'col-operator','หมายเลขโทรศัพท์ผู้จัดการสาขาที่ดำเนินการครั้งสุดท้าย':'col-manager'}; return map[col] || 'col-other'; }
function agingClass(v){ if(String(v).includes('มากกว่า 48')) return 'age-hot'; if(String(v).includes('เกิน 24')) return 'age-red'; if(String(v).includes('22–24')) return 'age-amber'; if(String(v).includes('ไม่เกิน 22')) return 'age-orange'; if(String(v).includes('ไม่เกิน 12')) return 'age-blue'; if(String(v).includes('ไม่เกิน 6')) return 'age-green'; return 'age-gray'; }
function readControls(){ return {reportTime:els.reportTime?.value, search:els.searchInput?.value, base:els.baseHubFilter?.value, mode:els.modeFilter?.value, aging:els.agingFilter?.value, hub:els.hubFilter?.value, action:els.actionFilter?.value, copyMode:els.copyLineMode?.value, copyLimit:els.copyLineLimit?.value, copyFocus:getSelectedReportFocuses().join(','), copyTrackMode:els.copyTrackMode?.value, sort:els.sortFilter?.value, showAll:els.showAllCols?.checked, clean:els.cleanThCodes?.checked, compact:els.compactRows?.checked}; }
function writeControls(c){ if(c.reportTime&&els.reportTime) els.reportTime.value=c.reportTime; if(c.search&&els.searchInput) els.searchInput.value=c.search; [['baseHubFilter','base'],['modeFilter','mode'],['agingFilter','aging'],['hubFilter','hub'],['actionFilter','action'],['copyLineMode','copyMode'],['copyLineLimit','copyLimit'],['copyTrackMode','copyTrackMode'],['sortFilter','sort']].forEach(([id,k])=>{if(c[k]&&els[id]) els[id].value=c[k];});
  restoreReportFocusCheckboxes(c.copyFocus || 'all'); if(els.showAllCols) els.showAllCols.checked=!!c.showAll; if(els.cleanThCodes) els.cleanThCodes.checked=c.clean!==false; if(els.compactRows) els.compactRows.checked=c.compact!==false; }
function displayValue(v){ let s=cleanText(v); if(els.cleanThCodes?.checked!==false) s=s.replace(/\s*\(TH\d{6,}\)\s*/g,' ').replace(/\s+/g,' ').trim(); return s; }
function compactManagerLabel(v){ const s=displayValue(v); return s.length>70?s.slice(0,67)+'...':s; }
function isComputed(c){ return ['เวลาที่ใช้คำนวณ','ชั่วโมงในคลัง','เวลาอยู่ภายในฮับ'].includes(c); }
function cleanText(v){ return String(v??'').replace(/\u00a0/g,' ').trim(); }
function unique(arr){ return [...new Set(arr.filter(v=>v!==undefined&&v!==null&&String(v).trim()!==''))]; }
function thSort(a,b){ return String(a).localeCompare(String(b),'th',{numeric:true,sensitivity:'base'}); }
function esc(s){ return String(s??'').replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m])); }
function num(n){ return Number(n||0).toLocaleString('th-TH'); }
function percent(n,total){ return total?`${((n/total)*100).toFixed(1)}%`:'0.0%'; }
function round2(n){ return Math.round(n*100)/100; }
function getReportDate(){ return els.reportTime?.value ? new Date(els.reportTime.value) : new Date(); }
function toLocalInputValue(d){ const pad=n=>String(n).padStart(2,'0'); return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`; }
function formatDateTime(d){ if(!d||isNaN(d)) return ''; const pad=n=>String(n).padStart(2,'0'); return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; }
function dateStamp(){ const d=new Date(); const p=n=>String(n).padStart(2,'0'); return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`; }
function shortHeader(c){ return c.replace('หมายเลขโทรศัพท์ผู้จัดการสาขาที่ดำเนินการครั้งสุดท้าย','ผู้ดำเนินการ/เบอร์/รหัส').replace('เวลาที่ดำเนินการล่าสุด','เวลาดำเนินการล่าสุด'); }
function shortLabel(s,max=36){ const t=String(s||''); return t.length>max?t.slice(0,max-1)+'…':t; }
function csvCell(v){ const s=String(v??''); return /[",\n]/.test(s)?`"${s.replace(/"/g,'""')}"`:s; }
function downloadBlob(content,filename,type){ const blob=new Blob([content],{type}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename; a.click(); URL.revokeObjectURL(a.href); }
function setStatus(title,msg){ if(els.statusCard) els.statusCard.innerHTML=`<b>${esc(title)}</b><span>${esc(msg||'')}</span>`; }
function toast(msg){ if(!els.toast) return; els.toast.textContent=msg; els.toast.hidden=false; clearTimeout(toast.t); toast.t=setTimeout(()=>els.toast.hidden=true,2200); }
function palette(i,alpha=1){ const colors=[[197,127,54],[95,118,141],[94,123,97],[155,77,72],[115,100,137],[170,116,55],[109,90,69],[127,114,142]]; const c=colors[i % colors.length]; return `rgba(${c[0]},${c[1]},${c[2]},${alpha})`; }

function timestampFile(){
  const d = new Date();
  const pad = n => String(n).padStart(2,'0');
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}
function formatHours(v){
  const n = Number(v || 0);
  if(!Number.isFinite(n)) return '0';
  return n >= 10 ? n.toFixed(1) : n.toFixed(2);
}
function trackingNo(r){
  const candidates = ['เลขพัสดุ','หมายเลขพัสดุ','tracking','Tracking','TRACKING','track','Track','เลขที่พัสดุ','waybill','Waybill','barcode','Barcode'];
  for(const k of candidates){
    const v = cleanText(r?.[k]);
    if(v) return v;
  }
  const found = Object.keys(r || {}).find(k => /track|พัสดุ|waybill|barcode/i.test(k) && cleanText(r[k]));
  return found ? cleanText(r[found]) : '';
}
function shortAgingLabel(label){
  const text = String(label || '');
  if(AGING_SHORT.includes(text)) return text;
  if(text.includes('ไม่เกิน 6')) return '≤6 ชม.';
  if(text.includes('ไม่เกิน 12')) return '≤12 ชม.';
  if(text.includes('ไม่เกิน 22')) return '≤22 ชม.';
  if(text.includes('22') && text.includes('24')) return '22–24 ชม.';
  if(text.includes('มากกว่า 48')) return '>48 ชม.';
  if(text.includes('เกิน 24')) return '>24 ชม.';
  return 'ตรวจสอบ';
}

async function captureTarget(id, name){
  const target = document.getElementById(id);
  if(!target){ toast('ไม่พบส่วนที่ต้องการบันทึก'); return; }

  const captureEl = getBestCaptureTarget(target);
  toast('กำลังสร้างภาพ...');

  // ตาราง: วาดภาพเองด้วย Canvas ก่อนเลย เสถียรกว่า html2canvas มาก
  const table = captureEl.querySelector?.('table') || (captureEl.tagName === 'TABLE' ? captureEl : null);
  if(table){
    try{
      const ok = await captureTableCanvas(table, captureEl, name || id);
      if(ok){ toast('บันทึกภาพแล้ว'); return; }
    }catch(tableErr){
      console.error('Table canvas capture failed', tableErr);
    }
  }

  // กราฟ: ดึงจาก canvas ของ Chart.js โดยตรง
  const chartCanvas = captureEl.querySelector?.('canvas');
  if(chartCanvas){
    try{
      const ok = await captureChartCanvas(chartCanvas, captureEl, name || id);
      if(ok){ toast('บันทึกภาพแล้ว'); return; }
    }catch(chartErr){
      console.error('Chart canvas capture failed', chartErr);
    }
  }

  // วิธีสำรองท้ายสุด: html2canvas
  try{
    if(typeof html2canvas === 'undefined') throw new Error('html2canvas not loaded');
    const canvas = await html2canvas(captureEl, {
      backgroundColor:'#ffffff',
      scale:2,
      useCORS:true,
      allowTaint:true,
      logging:false,
      removeContainer:true
    });
    if(!canvas || !canvas.width || !canvas.height) throw new Error('empty canvas');
    downloadCanvas(canvas, `${name || id}_${timestampFile()}.png`);
    toast('บันทึกภาพแล้ว');
  }catch(err){
    console.error('All capture methods failed', err);
    toast('บันทึกภาพไม่สำเร็จ: ใช้ปุ่มคัดลอกไลน์งาน หรือ Export Excel แทนชั่วคราว');
  }
}
function getBestCaptureTarget(target){
  const cards = target.querySelectorAll?.('.table-card,.chart-card,.card,.panel,.capture-card');
  if(cards && cards.length === 1) return cards[0];

  const table = target.querySelector?.('table');
  if(table){
    const card = table.closest('.table-card,.card,.panel,.capture-card');
    if(card) return card;
    return table.closest('.table-wrap') || table;
  }

  const chart = target.querySelector?.('canvas');
  if(chart){
    const card = chart.closest('.chart-card,.card,.panel,.capture-card');
    if(card) return card;
  }

  return target;
}
function downloadCanvas(canvas, filename){
  const a = document.createElement('a');
  a.download = filename;
  try{
    a.href = canvas.toDataURL('image/png');
  }catch(e){
    console.error('toDataURL failed', e);
    return false;
  }
  document.body.appendChild(a);
  a.click();
  a.remove();
  return true;
}
async function captureTableCanvas(table, container, name){
  const rows = [...table.rows].map(tr => [...tr.cells].map(td => cleanText(td.innerText)));
  if(!rows.length) return false;

  const title = safeCaptureTitle(container, 'รายงานตาราง');
  const subtitle = safeCaptureSubtitle(container);
  const colCount = Math.max(...rows.map(r => r.length));
  const font = 'Arial, Tahoma, sans-serif';
  const scale = 2;
  const padX = 20;
  const padY = 18;
  const titleH = subtitle ? 70 : 48;
  const rowH = 38;

  const measureCanvas = document.createElement('canvas');
  const mctx = measureCanvas.getContext('2d');
  mctx.font = `14px ${font}`;

  const colWidths = [];
  for(let c=0;c<colCount;c++){
    let maxText = '';
    rows.forEach(r => { if(String(r[c]||'').length > maxText.length) maxText = String(r[c]||''); });
    const measured = Math.ceil(mctx.measureText(maxText).width) + 34;
    colWidths[c] = Math.min(Math.max(measured, c === 0 ? 300 : 95), c === 0 ? 520 : 150);
  }

  const tableW = colWidths.reduce((a,b)=>a+b,0);
  const width = Math.max(760, tableW + padX*2);
  const height = titleH + rows.length * rowH + padY;

  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(width * scale);
  canvas.height = Math.ceil(height * scale);
  const ctx = canvas.getContext('2d');
  ctx.scale(scale, scale);

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0,0,width,height);

  ctx.fillStyle = '#0f172a';
  ctx.font = `700 22px ${font}`;
  ctx.textBaseline = 'alphabetic';
  ctx.fillText(title, padX, 32);

  if(subtitle){
    ctx.fillStyle = '#64748b';
    ctx.font = `500 13px ${font}`;
    ctx.fillText(subtitle, padX, 54);
  }

  let y = titleH;
  rows.forEach((row, rIdx) => {
    let x = padX;
    const isHeader = rIdx === 0;
    ctx.fillStyle = isHeader ? '#eef2f7' : (rIdx % 2 ? '#ffffff' : '#f8fafc');
    ctx.fillRect(padX, y, tableW, rowH);

    for(let c=0;c<colCount;c++){
      const w = colWidths[c];

      ctx.strokeStyle = '#dbe3ee';
      ctx.lineWidth = 1;
      ctx.strokeRect(x, y, w, rowH);

      const raw = row[c] || '';
      const isNumCol = c > 0;
      const isTotalCol = c === colCount - 1;
      const maxChars = Math.max(4, Math.floor((w - 20) / (isNumCol ? 8 : 7.5)));
      const text = truncateForCell(raw, maxChars);

      ctx.fillStyle = isHeader ? '#0f172a' : (isTotalCol ? '#9a3412' : '#334155');
      ctx.font = `${isHeader || isTotalCol ? '700' : '500'} 14px ${font}`;
      ctx.textAlign = isNumCol ? 'right' : 'left';
      ctx.textBaseline = 'middle';
      ctx.fillText(text, isNumCol ? x + w - 10 : x + 10, y + rowH/2 + 1);

      x += w;
    }
    y += rowH;
  });

  return downloadCanvas(canvas, `${name}_${timestampFile()}.png`);
}
async function captureChartCanvas(chartCanvas, container, name){
  const title = safeCaptureTitle(container, 'กราฟรายงาน');
  const subtitle = safeCaptureSubtitle(container);
  const scale = 2;
  const pad = 22;
  const titleH = subtitle ? 76 : 54;
  const width = Math.max(720, chartCanvas.width / (window.devicePixelRatio || 1) + pad*2);
  const height = Math.max(420, chartCanvas.height / (window.devicePixelRatio || 1) + titleH + pad);

  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(width * scale);
  canvas.height = Math.ceil(height * scale);
  const ctx = canvas.getContext('2d');
  ctx.scale(scale, scale);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0,0,width,height);

  ctx.fillStyle = '#0f172a';
  ctx.font = '700 22px Arial, Tahoma, sans-serif';
  ctx.fillText(title, pad, 32);
  if(subtitle){
    ctx.fillStyle = '#64748b';
    ctx.font = '500 13px Arial, Tahoma, sans-serif';
    ctx.fillText(subtitle, pad, 56);
  }
  ctx.drawImage(chartCanvas, pad, titleH, width-pad*2, height-titleH-pad);
  return downloadCanvas(canvas, `${name}_${timestampFile()}.png`);
}

function safeCaptureTitle(container, fallback){
  try{ return captureTitle(container) || fallback || 'Parcel Operations'; }
  catch(e){ console.warn('captureTitle failed', e); return fallback || 'Parcel Operations'; }
}
function safeCaptureSubtitle(container){
  try{ return captureSubtitle(container) || ''; }
  catch(e){ console.warn('captureSubtitle failed', e); return ''; }
}

function captureTitle(container){
  return cleanText(container.querySelector?.('h1,h2,h3,.card-title,.section-title')?.innerText)
    || cleanText(container.closest?.('.card,.table-card,.chart-card')?.querySelector?.('h1,h2,h3,.card-title,.section-title')?.innerText)
    || 'Parcel Operations';
}
function captureSubtitle(container){
  const text = cleanText(container.querySelector?.('.muted,.subtitle,.card-subtitle')?.innerText);
  return text.length > 120 ? text.slice(0,117) + '...' : text;
}
function truncateForCell(text, max){
  const s = String(text ?? '');
  if(s.length <= max) return s;
  return s.slice(0, Math.max(1,max-1)) + '…';
}

function getBestCaptureTarget(target){
  const cards = target.querySelectorAll?.('.table-card,.chart-card,.card,.panel,.capture-card');
  if(cards && cards.length === 1) return cards[0];

  const table = target.querySelector?.('table');
  if(table){
    const card = table.closest('.table-card,.card,.panel,.capture-card');
    if(card) return card;
    return table.closest('.table-wrap') || table;
  }

  const chart = target.querySelector?.('canvas');
  if(chart){
    const card = chart.closest('.chart-card,.card,.panel,.capture-card');
    if(card) return card;
  }

  return target;
}
function downloadCanvas(canvas, filename){
  const a = document.createElement('a');
  a.download = filename;
  a.href = canvas.toDataURL('image/png');
  document.body.appendChild(a);
  a.click();
  a.remove();
}
async function captureTableFallback(target, name){
  const table = target.querySelector?.('table') || (target.tagName === 'TABLE' ? target : null);
  if(!table) return false;

  const rows = [...table.rows].map(tr => [...tr.cells].map(td => cleanText(td.innerText)));
  if(!rows.length) return false;

  const title = cleanText(target.querySelector?.('h1,h2,h3,.card-title')?.innerText) || cleanText(target.previousElementSibling?.innerText) || 'รายงาน';
  const colCount = Math.max(...rows.map(r => r.length));
  const colWidths = [];
  for(let c=0;c<colCount;c++){
    const maxLen = Math.max(...rows.map(r => String(r[c]||'').length), 6);
    colWidths[c] = Math.min(Math.max(maxLen * 8 + 28, c === 0 ? 260 : 92), c === 0 ? 430 : 130);
  }
  const width = Math.max(720, colWidths.reduce((a,b)=>a+b,0) + 32);
  const rowH = 34;
  const headerH = 58;
  const height = headerH + rows.length * rowH + 24;

  let x = 16;
  let y = headerH;
  const escXml = s => String(s ?? '').replace(/[&<>"]/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));
  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
    <rect width="100%" height="100%" fill="#ffffff"/>
    <text x="16" y="34" font-family="Arial, Tahoma, sans-serif" font-size="20" font-weight="700" fill="#0f172a">${escXml(title)}</text>`;

  rows.forEach((row, rIdx) => {
    x = 16;
    const bg = rIdx === 0 ? '#eef2f7' : (rIdx % 2 ? '#ffffff' : '#f8fafc');
    const fg = rIdx === 0 ? '#0f172a' : '#334155';
    const weight = rIdx === 0 ? '700' : '500';

    for(let c=0;c<colCount;c++){
      const w = colWidths[c];
      svg += `<rect x="${x}" y="${y}" width="${w}" height="${rowH}" fill="${bg}" stroke="#dbe3ee" stroke-width="1"/>`;
      const text = escXml(truncateForCell(row[c] || '', Math.floor((w-18)/7.2)));
      const alignRight = c > 0;
      const tx = alignRight ? x + w - 10 : x + 10;
      const anchor = alignRight ? 'end' : 'start';
      const fill = c === colCount-1 && rIdx > 0 ? '#9a3412' : fg;
      const fw = c === colCount-1 || rIdx === 0 ? '700' : weight;
      svg += `<text x="${tx}" y="${y+22}" font-family="Arial, Tahoma, sans-serif" font-size="14" font-weight="${fw}" fill="${fill}" text-anchor="${anchor}">${text}</text>`;
      x += w;
    }
    y += rowH;
  });
  svg += `</svg>`;

  const blob = new Blob([svg], {type:'image/svg+xml;charset=utf-8'});
  const url = URL.createObjectURL(blob);
  const img = new Image();
  img.decoding = 'async';
  img.src = url;
  await img.decode();

  const canvas = document.createElement('canvas');
  canvas.width = width * 2;
  canvas.height = height * 2;
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0,0,canvas.width,canvas.height);
  ctx.scale(2,2);
  ctx.drawImage(img,0,0);
  URL.revokeObjectURL(url);

  downloadCanvas(canvas, `${name}_${timestampFile()}.png`);
  return true;
}
function truncateForCell(text, max){
  const s = String(text ?? '');
  if(s.length <= max) return s;
  return s.slice(0, Math.max(1,max-1)) + '…';
}

function getBestCaptureTarget(target){
  // ใช้ target ที่ปุ่มกำหนดเป็นหลัก เพื่อลดโอกาส html2canvas หา element ใน cloned document ไม่เจอ
  // แต่ถ้า target เป็น wrapper ใหญ่และมีการ์ดเดียว ให้จับการ์ดนั้นแทน
  const cards = target.querySelectorAll?.('.table-card,.chart-card,.card,.panel,.capture-card');
  if(cards && cards.length === 1) return cards[0];

  const table = target.querySelector?.('table');
  if(table){
    const card = table.closest('.table-card,.card,.panel,.capture-card');
    if(card) return card;
    return table.closest('.table-wrap') || table;
  }

  const chart = target.querySelector?.('canvas');
  if(chart){
    const card = chart.closest('.chart-card,.card,.panel,.capture-card');
    if(card) return card;
  }

  return target;
}

function getBestCaptureTarget(target){
  const table = target.querySelector?.('table');
  if(table){
    const card = table.closest('.table-card,.card,.panel,.capture-card');
    if(card) return card;
    const wrap = table.closest('.table-wrap');
    if(wrap) return wrap;
    return table;
  }
  const chart = target.querySelector?.('canvas');
  if(chart){
    const card = chart.closest('.chart-card,.card,.panel,.capture-card');
    if(card) return card;
  }
  return target;
}



function currentSortedRows(){
  const mode = els.sortFilter?.value || 'hours-desc';
  const rows = [...state.filtered];
  const hoursOf = r => Number(r['ชั่วโมงในคลัง'] ?? r.hours ?? 0);
  const latestOf = r => parseDateValue(r['เวลาที่ดำเนินการล่าสุด'])?.getTime?.() || parseDateValue(r.lastActionAt)?.getTime?.() || 0;
  const byTrack = (a,b)=>thSort(trackingNo(a),trackingNo(b));
  if(mode==='original') rows.sort((a,b)=>(a.__idx||0)-(b.__idx||0));
  else if(mode==='hours-asc') rows.sort((a,b)=>hoursOf(a)-hoursOf(b) || latestOf(a)-latestOf(b));
  else if(mode==='latest-desc') rows.sort((a,b)=>latestOf(b)-latestOf(a) || hoursOf(b)-hoursOf(a));
  else if(mode==='latest-asc') rows.sort((a,b)=>latestOf(a)-latestOf(b) || hoursOf(b)-hoursOf(a));
  else if(mode==='tracking-asc') rows.sort(byTrack);
  else rows.sort((a,b)=>hoursOf(b)-hoursOf(a) || latestOf(a)-latestOf(b));
  return rows;
}

function getCopySettings(){
  const limitRaw = els.copyLineLimit?.value || 'all';
  const limit = limitRaw === 'all' ? Infinity : Number(limitRaw || 10);
  return {
    mode: els.copyLineMode?.value || 'aging',
    focuses: getSelectedReportFocuses(),
    trackMode: els.copyTrackMode?.value || 'all',
    limit,
    limitLabel: limitRaw === 'all' ? 'ทั้งหมด' : `Top ${limit}`
  };
}
function getSelectedReportFocuses(){
  const checked = [...document.querySelectorAll('input[name="copyReportFocus"]:checked')].map(x=>x.value);
  return checked.length ? checked : ['all'];
}
function onReportFocusCheckboxChange(e){
  const boxes = [...document.querySelectorAll('input[name="copyReportFocus"]')];
  const allBox = boxes.find(x=>x.value==='all');
  if(e?.target?.value === 'all' && e.target.checked){
    boxes.filter(x=>x.value!=='all').forEach(x=>x.checked=false);
  }else if(e?.target?.value !== 'all' && e?.target?.checked){
    if(allBox) allBox.checked=false;
  }
  if(!boxes.some(x=>x.checked) && allBox) allBox.checked=true;
  saveLightState();
}
function restoreReportFocusCheckboxes(value){
  const values = String(value || 'all').split(',').map(x=>x.trim()).filter(Boolean);
  const set = new Set(values.length ? values : ['all']);
  const boxes = [...document.querySelectorAll('input[name="copyReportFocus"]')];
  boxes.forEach(x=>x.checked=set.has(x.value));
  const allBox = boxes.find(x=>x.value==='all');
  if(!boxes.some(x=>x.checked) && allBox) allBox.checked=true;
  if(allBox?.checked) boxes.filter(x=>x.value!=='all').forEach(x=>x.checked=false);
}
function copyModeLabel(mode){
  if(mode === 'count') return 'จำนวนชิ้น';
  if(mode === 'both') return 'เวลาในคลัง + จำนวนชิ้น';
  return 'เวลาในคลัง';
}
function focusLabel(focus){
  const map = {
    all:'ทุกช่วงเวลา',
    le6:'≤6 ชม.',
    le12:'≤12 ชม.',
    le22:'≤22 ชม.',
    between22_24:'22–24 ชม.',
    over24:'>24 ชม.',
    over48:'>48 ชม.',
    critical:'22–24 / >24 / >48 ชม.'
  };
  return map[focus] || focus;
}
function focusLabels(focuses){
  return normalizeFocuses(focuses).map(focusLabel).join(' + ');
}
function normalizeFocuses(focuses){
  const arr = Array.isArray(focuses) ? focuses : [focuses || 'all'];
  if(arr.includes('all')) return ['all'];
  return arr.length ? arr : ['all'];
}
function selectedBucketLabels(focuses){
  const f = normalizeFocuses(focuses);
  if(f.includes('all')) return AGING_SHORT.slice(0,6);
  const map = {
    le6:'≤6 ชม.',
    le12:'≤12 ชม.',
    le22:'≤22 ชม.',
    between22_24:'22–24 ชม.',
    over24:'>24 ชม.',
    over48:'>48 ชม.',
    critical:null
  };
  let labels = [];
  f.forEach(x => {
    if(x === 'critical') labels.push('22–24 ชม.','>24 ชม.','>48 ชม.');
    else if(map[x]) labels.push(map[x]);
  });
  return [...new Set(labels)].filter(Boolean);
}
function criticalCount(map){ return (map['22–24 ชม.']||0) + (map['>24 ชม.']||0) + (map['>48 ชม.']||0); }
function over24Count(map){ return (map['>24 ชม.']||0) + (map['>48 ชม.']||0); }
function focusMatch(row, focus){
  if(focus === 'le6') return row.buckets?.['≤6 ชม.']?.count > 0 || row.bucketMap?.['≤6 ชม.'] > 0;
  if(focus === 'le12') return row.buckets?.['≤12 ชม.']?.count > 0 || row.bucketMap?.['≤12 ชม.'] > 0;
  if(focus === 'le22') return row.buckets?.['≤22 ชม.']?.count > 0 || row.bucketMap?.['≤22 ชม.'] > 0;
  if(focus === 'between22_24') return row.buckets?.['22–24 ชม.']?.count > 0 || row.bucketMap?.['22–24 ชม.'] > 0;
  if(focus === 'over24') return (row.over24 || 0) > 0;
  if(focus === 'over48') return (row.over48 || 0) > 0;
  if(focus === 'critical') return (row.critical || 0) > 0;
  return (row.total || 0) > 0;
}
function filterReportRows(rows, focuses){
  const f = normalizeFocuses(focuses);
  if(f.includes('all')) return rows.filter(r => (r.total || 0) > 0);
  return rows.filter(r => f.some(x=>focusMatch(r,x)));
}
function reportPriority(row, focuses){
  const f = normalizeFocuses(focuses);
  if(f.includes('over48')) return row.over48 || 0;
  if(f.includes('over24')) return row.over24 || 0;
  if(f.includes('between22_24')) return row.buckets?.['22–24 ชม.']?.count || 0;
  if(f.includes('critical')) return row.critical || 0;
  return row.total || 0;
}
function reportSorter(focuses){
  return (a,b) => {
    const an = a.name || a.hub || '';
    const bn = b.name || b.hub || '';
    const fa = normalizeFocuses(focuses);
    if(!fa.includes('all')){
      return reportPriority(b,fa)-reportPriority(a,fa) || (b.maxHours||0)-(a.maxHours||0) || (b.total||0)-(a.total||0) || thSort(an,bn);
    }
    return (b.total||0)-(a.total||0) || (b.maxHours||0)-(a.maxHours||0) || thSort(an,bn);
  };
}
function formatBucketLine(label, value){
  return `${label} = ${num(value)} ชิ้น`;
}
function formatBucketInline(label, value){
  return `${label} = ${num(value)} ชิ้น`;
}
function formatTotalAndMax(total, maxHours){
  return `รวม ${num(total)} ชิ้น | ค้างนานสุด ${formatHours(maxHours)} ชม.`;
}


function copyTableForLine(tableId, title){
  const settings = getCopySettings();
  const base = selectedBaseText();
  const groupKey = inferCopyGroupKey(tableId, title);
  const sourceRows = copySourceRowsForTable(tableId);
  let groups;
  const overviewHubRows = tableId === 'hubSummaryTable' && getSelectedBaseRaw() !== 'all' ? getLhRowsForSelectedBase() : sourceRows;
  if(tableId === 'hubSummaryTable' && getSelectedBaseRaw() !== 'all'){
    groups = [
      ...buildGroupedTrackingReportFromRows(getFdRowsForSelectedBase(), 'สาขาปลายทาง', settings).map(g=>({...g,name:`FD: ${g.name}`})),
      ...buildGroupedTrackingReportFromRows(getLhRowsForSelectedBase(), 'HUBปลายทาง', settings).map(g=>({...g,name:`LH: ${g.name}`}))
    ].sort(reportSorter(settings.focuses));
  }else{
    groups = buildGroupedTrackingReportFromRows(overviewHubRows, groupKey, settings);
  }

  const lines = [];
  lines.push(`📌 ${title}`);
  lines.push(`เวลารายงาน ${formatDateTime(getReportDate())}`);
  lines.push(`สาขา ${base}`);
  lines.push(`ข้อมูลที่ใช้คัดลอก ${copyScopeLabel(tableId)}`);
  lines.push(`รวม ${num(tableId === 'hubSummaryTable' && getSelectedBaseRaw() !== 'all' ? overviewHubRows.length : sourceRows.length)} ชิ้น`);
  lines.push(`กลุ่มเวลา ${focusLabels(settings.focuses)}`);
  lines.push(`รูปแบบ ${copyModeLabel(settings.mode)}`);
  lines.push(`เลขพัสดุ ${settings.trackMode === 'all' ? 'แสดงทั้งหมด' : 'แสดงตัวอย่าง 5 เลข/ช่วงเวลา'}`);

  appendGroupedReport(lines, groups, settings, '🏷 รายการหลัก');



  if(settings.trackMode === 'all'){
    lines.push('');
    lines.push('หมายเหตุ: โหมดเลขพัสดุทั้งหมดอาจทำให้ข้อความยาวมาก หากวางใน LINE ไม่ครบ ให้ใช้ Export Excel แทน');
  }

  copyText(lines.join('\n'));
}

function appendGroupedReport(lines, groups, settings, heading, includeTotals=true){
  if(!groups.length){
    lines.push('');
    lines.push('ไม่มีข้อมูลตามเงื่อนไข');
    return;
  }

  const shown = groups.slice(0, settings.limit);
  const rest = groups.slice(settings.limit);
  const totals = sumBucketTotals(groups);

  if(includeTotals && settings.mode !== 'count'){
    lines.push('');
    lines.push('⏱ สรุปเวลาในคลัง');
    selectedBucketLabels(settings.focuses).forEach(label => {
      if((totals[label] || 0) > 0) lines.push(`- ${formatBucketLine(label, totals[label])}`);
    });
  }

  lines.push('');
  lines.push(`${heading} ${settings.limitLabel} (${num(groups.length)} รายการ)`);

  shown.forEach((g, idx) => {
    lines.push(`${idx+1}) ${g.name}`);
    if(settings.mode === 'count'){
      lines.push(`   รวม ${num(g.total)} ชิ้น`);
    }else{
      const agingText = selectedBucketLabels(settings.focuses)
        .filter(label => g.buckets[label]?.count > 0)
        .map(label => formatBucketInline(label, g.buckets[label].count))
        .join(' | ');
      if(settings.mode === 'both') lines.push(`   รวม ${num(g.total)} ชิ้น`);
      lines.push(`   เวลาในคลัง ${agingText || '-'}`);
    }

    selectedBucketLabels(settings.focuses).forEach(label => {
      const b = g.buckets[label];
      if(!b || b.count <= 0) return;
      lines.push(`   - ${formatBucketLine(label, b.count)}`);
      lines.push(`     เลขพัสดุ ${formatTracksForReport(b.tracks) || '-'}`);
    });
  });

  if(rest.length){
    const restTotal = rest.reduce((s,g)=>s+g.total,0);
    lines.push(`อื่น ๆ ${num(rest.length)} รายการ — รวม ${num(restTotal)} ชิ้น`);
  }
}

function copyRecordsReportForLine(){
  const settings = getCopySettings();
  const recordsSettings = {...settings};
  if(normalizeFocuses(recordsSettings.focuses).includes('all')){
    recordsSettings.focuses = ['between22_24','over24','over48'];
  }

  const groups = buildGroupedTrackingReportFromRows(state.filtered, 'HUBปลายทาง', recordsSettings);
  const base = selectedBaseText();

  const lines = [];
  lines.push('📦 เลขพัสดุค้างตาม HUB ปลายทาง');
  lines.push(`เวลารายงาน ${formatDateTime(getReportDate())}`);
  lines.push(`สาขา ${base}`);
  lines.push(`รวมตามเงื่อนไข ${num(state.filtered.length)} ชิ้น`);
  lines.push(`รายงานเฉพาะ ${focusLabels(recordsSettings.focuses)}`);
  lines.push(`เลขพัสดุ ${recordsSettings.trackMode === 'all' ? 'แสดงทั้งหมด' : 'แสดงตัวอย่าง 5 เลข/ช่วงเวลา'}`);
  lines.push('เรียงตาม งานค้างนาน/เสี่ยงสูงก่อน');
  appendGroupedReport(lines, groups, recordsSettings, 'รายการ HUB ค้าง', false);

  if(recordsSettings.trackMode === 'all'){
    lines.push('');
    lines.push('หมายเหตุ: โหมดเลขพัสดุทั้งหมดอาจยาวมาก หากวางใน LINE ไม่ครบ ให้ใช้ Export Excel แทน');
  }

  copyText(lines.join('\n'));
}

function getSelectedBaseRaw(){
  return els.baseHubFilter?.value || 'all';
}
function getFdRowsForSelectedBase(){
  const base = getSelectedBaseRaw();
  if(base === 'all') return [];
  return state.filtered.filter(r => cleanText(r['HUBปลายทาง']) === base);
}
function getLhRowsForSelectedBase(){
  const base = getSelectedBaseRaw();
  if(base === 'all') return [];
  return state.filtered.filter(r => cleanText(r['HUBปลายทาง']) !== base);
}
function copySourceRowsForTable(tableId){
  if(tableId === 'branchSummaryTable' || tableId === 'fdBranchStatusTable') return getFdRowsForSelectedBase();
  if(tableId === 'lhOtherTable' || tableId === 'lhHubStatusTable') return getLhRowsForSelectedBase();
  return state.filtered;
}
function copyScopeLabel(tableId){
  if(tableId === 'branchSummaryTable' || tableId === 'fdBranchStatusTable') return 'เฉพาะ FD / สาขาปลายทางของ HUB ที่เลือก';
  if(tableId === 'lhOtherTable' || tableId === 'lhHubStatusTable') return 'เฉพาะ LH / HUB ปลายทางอื่น';
  if(tableId === 'hubSummaryTable') return getSelectedBaseRaw() === 'all' ? 'ภาพรวม HUB ปลายทางทั้งหมด' : 'ภาพรวมรวม FD รายสาขาของ HUB ที่เลือก + LH / HUB อื่น';
  if(tableId === 'statusSummaryTable') return 'สรุปตามสถานะดำเนินการล่าสุด';
  return 'ตามตัวกรองปัจจุบัน';
}

function inferCopyGroupKey(tableId, title){
  const t = `${tableId || ''} ${title || ''}`;
  if(t.includes('สถานะ') || tableId === 'statusSummaryTable') return 'การดำเนินการล่าสุด';
  if(t.includes('สาขา') || tableId === 'branchSummaryTable' || tableId === 'fdBranchStatusTable') return 'สาขาปลายทาง';
  return 'HUBปลายทาง';
}
function buildGroupedTrackingReport(groupKey, settings){
  return buildGroupedTrackingReportFromRows(state.filtered, groupKey, settings);
}
function buildGroupedTrackingReportFromRows(rows, groupKey, settings){
  const labels = selectedBucketLabels(settings.focuses);
  const sampleLimit = settings.trackMode === 'all' ? Infinity : 5;
  const map = new Map();

  rows.forEach(r => {
    const label = shortAgingLabel(r.__aging || r['เวลาอยู่ภายในฮับ'] || r.aging?.label || '');
    if(!labels.includes(label)) return;

    const raw = cleanText(r[groupKey]) || 'ไม่ระบุ';
    const name = groupKey.includes('HUB') || groupKey.includes('สาขา') ? cleanHubName(raw) : displayValue(raw);
    const key = normalizeGroupKey(name);
    if(!map.has(key)){
      map.set(key, {
        name,total:0,maxHours:0,
        buckets:Object.fromEntries(AGING_SHORT.slice(0,6).map(x => [x,{count:0,tracks:[],maxHours:0}]))
      });
    }

    const g = map.get(key);
    const hours = Number(r['ชั่วโมงในคลัง'] ?? r.hours ?? 0);
    const bucket = g.buckets[label] || (g.buckets[label] = {count:0,tracks:[],maxHours:0});
    g.total += 1;
    g.maxHours = Math.max(g.maxHours, hours || 0);
    bucket.count += 1;
    bucket.maxHours = Math.max(bucket.maxHours, hours || 0);

    const tr = trackingNo(r);
    if(tr && bucket.tracks.length < sampleLimit) bucket.tracks.push(tr);
  });

  return [...map.values()].map(g => ({
    ...g,
    critical:criticalCount(Object.fromEntries(Object.entries(g.buckets).map(([k,v]) => [k, v.count]))),
    over24:over24Count(Object.fromEntries(Object.entries(g.buckets).map(([k,v]) => [k, v.count]))),
    over48:g.buckets['>48 ชม.']?.count || 0
  })).filter(g => g.total > 0).sort(reportSorter(settings.focuses));
}
function sumBucketTotals(groups){
  const out = Object.fromEntries(AGING_SHORT.slice(0,6).map(x => [x,0]));
  groups.forEach(g => Object.entries(g.buckets).forEach(([label,b]) => out[label] = (out[label] || 0) + (b.count || 0)));
  return out;
}
function formatTracksForReport(tracks){
  if(!tracks || !tracks.length) return '';
  return tracks.join(', ');
}

async function copyText(text){
  const value = String(text ?? '');
  try{
    if(navigator.clipboard?.writeText && window.isSecureContext){
      await navigator.clipboard.writeText(value);
      toast('คัดลอกไลน์งานแล้ว');
      return;
    }
  }catch(e){
    console.warn('navigator.clipboard failed, fallback copy', e);
  }
  fallbackCopy(value);
}
function fallbackCopy(text){
  const ta=document.createElement('textarea');
  ta.value=text;
  ta.setAttribute('readonly','');
  ta.style.position='fixed';
  ta.style.left='-9999px';
  ta.style.top='0';
  ta.style.width='1px';
  ta.style.height='1px';
  document.body.appendChild(ta);
  ta.focus();
  ta.select();
  ta.setSelectionRange(0, ta.value.length);
  let ok=false;
  try{ ok=document.execCommand('copy'); }catch(e){ console.warn('execCommand copy failed', e); }
  ta.remove();
  if(ok) toast('คัดลอกไลน์งานแล้ว');
  else toast('คัดลอกไม่สำเร็จ เบราว์เซอร์บล็อก clipboard');
}
function normalizeCopyText(v){
  return String(v ?? '').replace(/\s+/g,' ').trim();
}




window.ParcelOpsSelfCheck = function(){
  const checks = [];
  const total = state.filtered.length;
  const sumTotal = (rows) => rows.reduce((s,r)=>s+(r.total||r.count||0),0);
  try{
    const hubTotal = sumTotal(summarize(state.filtered,'HUBปลายทาง'));
    const actionTotal = sumTotal(summarize(state.filtered,'การดำเนินการล่าสุด'));
    checks.push({name:'filtered rows loaded', pass: total >= 0, value: total});
    checks.push({name:'hub summary equals filtered rows', pass: hubTotal === total, value: hubTotal});
    checks.push({name:'latest action summary equals filtered rows', pass: actionTotal === total, value: actionTotal});
    const base = els.baseHubFilter?.value || 'all';
    if(base !== 'all'){
      const fd = state.filtered.filter(r => cleanText(r['HUBปลายทาง']) === base).length;
      const lh = state.filtered.filter(r => cleanText(r['HUBปลายทาง']) !== base).length;
      checks.push({name:'FD + LH equals filtered rows', pass: fd + lh === total, value:`FD ${fd} + LH ${lh} = ${fd+lh}`});
    }
    const bucketTotal = AGING_ORDER.slice(0,6).reduce((s,a)=>s+state.filtered.filter(r=>r.__aging===a).length,0);
    checks.push({name:'aging buckets equal filtered rows', pass: bucketTotal === total, value: bucketTotal});
  }catch(e){
    checks.push({name:'self check exception', pass:false, value:String(e)});
  }
  console.table(checks);
  return checks;
};

window.ParcelOpsDebugCheck = function(){
  const names = ['timestampFile','formatHours','trackingNo','shortAgingLabel','getCopySettings','formatBucketLine','selectedBucketLabels','buildGroupedTrackingReport','copyTableForLine','copyRecordsReportForLine','captureTarget','captureTableCanvas'];
  const result = names.map(name => ({name, exists: typeof window[name] === 'function' || typeof eval(name) === 'function'}));
  console.table(result);
  return result;
};


window.ParcelOpsStorageDebug = async function(){
  let dbData = null, sessionData = null, meta = null;
  try{ sessionData = JSON.parse(sessionStorage.getItem(STORAGE_KEY) || 'null'); }catch(e){}
  try{ meta = JSON.parse(localStorage.getItem(STORAGE_KEY + '_meta') || 'null'); }catch(e){}
  try{ dbData = await dbGet(DB_ID); }catch(e){ console.warn('db debug failed', e); }
  const result = {
    build: window.PARCEL_OPS_BUILD || 'unknown',
    currentRows: state.rows?.length || 0,
    currentFile: state.fileName || '',
    sessionRows: sessionData?.rows?.length || 0,
    sessionFile: sessionData?.fileName || '',
    indexedDbRows: dbData?.rows?.length || 0,
    indexedDbFile: dbData?.fileName || '',
    metaRows: meta?.rows || 0,
    metaFile: meta?.fileName || '',
    page: location.pathname
  };
  console.table(result);
  return result;
};


window.getCopySettings = getCopySettings;
window.copyTableForLine = copyTableForLine;
window.copyRecordsReportForLine = copyRecordsReportForLine;


window.ParcelOpsChartDebug = function(){
  const result = {
    build: window.PARCEL_OPS_BUILD || 'unknown',
    chartLoaded: typeof Chart !== 'undefined',
    statusCanvas: !!els.statusChart,
    branchCanvas: !!els.branchChart,
    lhCanvas: !!els.lhChart,
    statusPanelActive: !!document.querySelector('[data-page-panel="status"].active'),
    rows: state.rows?.length || 0,
    filtered: state.filtered?.length || 0
  };
  console.table(result);
  return result;
};

