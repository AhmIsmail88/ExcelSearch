let activeJob = null;
async function postJSON(url, data){
  const response = await fetch(url, {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  const result = await response.json();
  if(!response.ok || result.error) throw new Error(result.error || 'تعذر تنفيذ العملية');
  return result;
}
async function runJob(url, data){
  if(activeJob) throw new Error('انتظر المهمة الحالية أو ألغها أولًا.');
  activeJob = 'starting';
  const panel = document.getElementById('jobPanel');
  document.getElementById('jobMessage').textContent = 'بدء العملية...';
  document.getElementById('cancelJobBtn').disabled = false;
  panel.classList.remove('hidden');
  try{
    const start = await postJSON(url, {...data, async:true});
    activeJob = start.job_id;
    if(!activeJob) return start;
    while(true){
      const response = await fetch('/api/jobs/' + activeJob);
      const job = await response.json();
      if(!response.ok) throw new Error(job.error || 'تعذر متابعة العملية');
      document.getElementById('jobMessage').textContent = job.message || 'جاري التنفيذ...';
      if(job.state === 'done') return job.result;
      if(['error','cancelled'].includes(job.state)) throw new Error(job.error);
      await new Promise(resolve => setTimeout(resolve,700));
    }
  }finally{ activeJob = null; panel.classList.add('hidden'); }
}
document.getElementById('cancelJobBtn').addEventListener('click',async()=>{
  if(!activeJob || activeJob==='starting') return;
  document.getElementById('cancelJobBtn').disabled = true;
  await postJSON('/api/jobs/'+activeJob+'/cancel',{});
  document.getElementById('jobMessage').textContent = 'طلب الإلغاء مسجل؛ ينتظر نقطة توقف آمنة...';
});
let savedFilters={filters:[],draft:[]};
try{ savedFilters=JSON.parse(localStorage.getItem('excel-search-filters')||sessionStorage.getItem('excel-search-filters')||'null')||savedFilters; }catch(_){}
savedFilters.filters=Array.isArray(savedFilters.filters)?savedFilters.filters.filter(f=>f && typeof f.col==='string').slice(0,12):[];
savedFilters.draft=Array.isArray(savedFilters.draft)?savedFilters.draft.filter(f=>f && typeof f.col==='string').slice(0,12):savedFilters.filters;
let state = {offset: 0, limit: 20, total: 0, query: '', localColumns: [], columns: [], filters: Array.isArray(savedFilters.filters)?savedFilters.filters:[], dbId: null};
function readFilterRows(area){
  return [...area.querySelectorAll('.filter-row')].map(row=>({col:row.querySelector('.filter-col').value,scope:row.querySelector('.filter-col').value===''?'all':'column',op:row.querySelector('.filter-op').value,value:row.querySelector('.filter-value').value.trim()}));
}
function rememberSearchFilters(){
  const draft=readFilterRows(document.getElementById('filterRows'));
  const pending=JSON.stringify(draft)!==JSON.stringify(state.filters);
  document.getElementById('filterStatus').textContent=pending?'شروط معدّلة — اضغط تطبيق الفلاتر':state.filters.length?`${state.filters.length} شرط مطبّق — محفوظ عند التبديل وإعادة الفتح`:'بدون شروط — اضغط إضافة شرط للفلترة';
  try{const data=JSON.stringify({filters:state.filters,draft});localStorage.setItem('excel-search-filters',data);sessionStorage.setItem('excel-search-filters',data);}catch(_){}
}
function appendFilterRow(area,columns,value={},changed=()=>{}){
  if(area.children.length>=12) return;
  const row=document.createElement('div');row.className='filter-row';
  const choices=[...columns];if(value.col && !choices.includes(value.col)) choices.unshift(value.col);
  row.innerHTML=`<select class="filter-col" aria-label="عمود الفلترة"><option value="">أي خلية في الصف — كل الأعمدة</option>${choices.map(c=>`<option value="${escapeHTML(c)}">${fieldLabel(c)}${columns.includes(c)?'':' — غير موجود في هذا الملف'}</option>`).join('')}</select>
    <select class="filter-op" aria-label="شرط الفلترة"><option value="contains">يحتوي على</option><option value="not_contains">لا يحتوي على — استبعاد</option><option value="equals">يساوي</option><option value="empty">فارغ</option><option value="not_empty">غير فارغ</option></select>
    <input class="filter-value" aria-label="قيمة الفلتر" placeholder="مثال: الأعمال الميكانيكية">
    <button class="link-btn" type="button" aria-label="حذف الشرط">حذف</button>`;
  if(value.col) row.querySelector('.filter-col').value=value.col;
  row.querySelector('.filter-op').value=value.op||'contains';
  row.querySelector('input').value=value.value||'';
  const refresh=()=>{
    const op=row.querySelector('.filter-op'),all=row.querySelector('.filter-col').value==='';
    [...op.options].forEach(o=>o.disabled=all&&!['contains','not_contains'].includes(o.value));
    if(all&&!['contains','not_contains'].includes(op.value))op.value='not_contains';
    row.querySelector('input').disabled=['empty','not_empty'].includes(op.value);
  };
  refresh();row.addEventListener('change',()=>{refresh();changed();});row.addEventListener('input',changed);
  row.addEventListener('keydown',e=>{if(e.key==='Enter' && area.id==='filterRows'){e.preventDefault();applyFilters();}});
  row.querySelector('button').addEventListener('click',()=>{row.remove();changed();if(area.id==='filterRows') applyFilters();});
  area.appendChild(row);return row;
}
function refreshFilterOptions(area,columns,initial=[],changed=()=>{}){
  const draft=area.children.length?readFilterRows(area):initial;
  area.replaceChildren();draft.forEach(f=>appendFilterRow(area,columns,f,changed));
}


function showSetup(isFirstTime){
  document.getElementById('setupOverlay').classList.remove('hidden');
  document.getElementById('setupCloseBtn').classList.toggle('hidden', isFirstTime);
  document.getElementById('setupTitle').textContent = isFirstTime ? 'إعداد أول مرة' : 'إضافة ملف جديد';
  document.getElementById('nameInput').value = '';
  document.getElementById('excelPathInput').value = '';
  document.getElementById('uniqueColInput').value = '';
  document.getElementById('headerRowInput').value = '';
  document.getElementById('lastColumnInput').value = '';
  document.getElementById('setupError').classList.remove('show');
}
function hideSetup(){ document.getElementById('setupOverlay').classList.add('hidden'); }

document.getElementById('setupCloseBtn').addEventListener('click', hideSetup);

let statsRequestId=0;
async function loadStats(){
  const ticket=++statsRequestId;
  const r = await fetch('/api/stats' + (state.dbId?'?'+new URLSearchParams({db_id:state.dbId}):''));
  const s = await r.json();
  if(ticket!==statsRequestId) return;
  if(!s.configured){
    document.getElementById('mainScreen').classList.add('hidden');
    showSetup(!s.has_any_db);
    return;
  }
  hideSetup();
  document.getElementById('mainScreen').classList.remove('hidden');
  document.getElementById('rowCount').textContent = s.row_count;
  document.getElementById('lastSync').textContent = s.last_sync || '—';
  document.getElementById('activeFileName').textContent = s.display_name || '—';
  state.localColumns = s.local_columns || [];
  state.columns = s.search_columns || [];
  state.activeColumns = state.columns;
  state.dbId = s.db_id;
  if(state.loadedDbId !== s.db_id){
    state.loadedDbId = s.db_id;
    state.offset = 0;
  }
  await loadFileList();
  if(document.getElementById('allFilesToggle').checked) state.columns=state.allColumns || state.columns;
  refreshFilterOptions(document.getElementById('filterRows'),state.columns,savedFilters.draft||state.filters,rememberSearchFilters);
  savedFilters.draft=[];
  rememberSearchFilters();
  runSearch();
}

async function loadFileList(){
  const r = await fetch('/api/databases');
  const data = await r.json();
  state.allColumns=[...new Set(data.databases.flatMap(db=>db.columns || []))];
  const el = document.getElementById('fileDropdown');
  let html = '';
  for(const db of data.databases){
    html += `<div class="file-item ${db.id===state.dbId ? 'active':''}" data-id="${db.id}">
      <span class="fname">${escapeHTML(db.name)}</span>
      <span class="fcount">${db.row_count ?? ''}</span>
      <span class="fdel" data-del="${db.id}" title="حذف">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M3 6h18M8 6V4h8v2m-9 0 1 14h8l1-14"/></svg>
      </span>
    </div>`;
  }
  html += '<hr><button class="file-add-btn" id="addFileBtn">+ إضافة ملف جديد</button>';
  el.innerHTML = html;

  el.querySelectorAll('.file-item').forEach(item => {
    item.addEventListener('click', async (e) => {
      if(e.target.closest('.fdel')) return;
      const id = item.dataset.id;
      await fetch('/api/switch', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({db_id: id})});
      el.classList.remove('show');
      state.dbId=id;
      loadStats();
    });
  });
  el.querySelectorAll('.fdel').forEach(btn => {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      if(!confirm('هيتم حذف هذا الملف من الأرشيف نهائياً. متأكد؟')) return;
      await fetch('/api/delete_db', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({db_id: btn.dataset.del})});
      el.classList.remove('show');
      if(state.dbId===btn.dataset.del) state.dbId=null;
      loadStats();
    });
  });
  document.getElementById('addFileBtn').addEventListener('click', () => {
    el.classList.remove('show');
    showSetup(false);
  });
}

document.getElementById('fileSwitchBtn').addEventListener('click', (e) => {
  e.stopPropagation();
  document.getElementById('fileDropdown').classList.toggle('show');
});
document.addEventListener('click', (e) => {
  const dd = document.getElementById('fileDropdown');
  if(dd && !e.target.closest('.file-switcher')) dd.classList.remove('show');
});

function escapeHTML(value){ return String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function fieldLabel(name){ return escapeHTML(name.replaceAll('_', ' ')); }
function cellDirection(value){ return /\p{Script=Arabic}/u.test(String(value ?? '')) ? 'rtl' : 'ltr'; }

function renderResults(rows, total){
  const area = document.getElementById('resultsArea');
  document.getElementById('resultMeta').textContent = total ? `${total} نتيجة` : '';
  if(!rows.length){
    area.innerHTML = '<div class="empty">مفيش نتائج مطابقة. جرّب كلمة تانية.</div>';
    document.getElementById('pager').classList.add('hidden');
    return;
  }
  let html = '<div class="results-grid">';
  for(const row of rows) html += renderCardHTML(row);
  html += '</div>';
  area.innerHTML = html;
  attachEditHandlers(area);

  const pager = document.getElementById('pager');
  pager.classList.remove('hidden');
  document.getElementById('prevBtn').disabled = state.offset === 0;
  document.getElementById('nextBtn').disabled = state.offset + state.limit >= total;
}

let searchTimer;
document.getElementById('searchInput').addEventListener('input', (e) => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => {
    state.query = e.target.value.trim();
    state.offset = 0;
    runSearch();
  }, 300);
});
document.getElementById('allFilesToggle').addEventListener('change', () => {
  state.offset = 0;
  const allFiles = document.getElementById('allFilesToggle').checked;
  state.columns=allFiles?state.allColumns:state.activeColumns;
  refreshFilterOptions(document.getElementById('filterRows'),state.columns,[],rememberSearchFilters);
  rememberSearchFilters();
  runSearch();
});

document.getElementById('addFilterBtn').addEventListener('click',()=>{
  const row=appendFilterRow(document.getElementById('filterRows'),state.columns,{},rememberSearchFilters);
  row?.querySelector('input').focus();rememberSearchFilters();
});
function applyFilters(){
  const filters = readFilterRows(document.getElementById('filterRows'));
  if(filters.some(f => ['contains','not_contains','equals'].includes(f.op) && !f.value)){
    document.getElementById('resultMeta').textContent = 'اكتب قيمة لكل شرط أو اختار فارغ / غير فارغ.';
    return;
  }
  state.filters = filters;
  rememberSearchFilters();
  state.offset = 0;
  runSearch();
}
document.getElementById('applyFiltersBtn').addEventListener('click', applyFilters);
document.getElementById('clearFiltersBtn').addEventListener('click', () => {
  document.getElementById('filterRows').replaceChildren();
  state.filters = []; state.offset = 0; rememberSearchFilters(); runSearch();
});

let searchRequestId = 0;
async function runSearch(){
  const requestId = ++searchRequestId;
  const dbId = state.dbId;
  try {
  const area = document.getElementById('resultsArea');
  const allFiles = document.getElementById('allFilesToggle').checked;
  document.getElementById('pager').classList.toggle('hidden', allFiles);
  area.innerHTML = '<div class="loading">بيدوّر...</div>';
  const q = state.query;

  if(allFiles){
    if(!q && !state.filters.length){ area.innerHTML = '<div class="empty">اكتب كلمة أو طبّق فلترًا للبحث في كل الملفات.</div>'; document.getElementById('resultMeta').textContent=''; return; }
    const r = await fetch('/api/search_all?' + new URLSearchParams({q,filters:JSON.stringify(state.filters)}).toString());
    const data = await r.json();
    if(requestId !== searchRequestId || dbId!==state.dbId) return;
    if(data.error) throw new Error(data.error);
    renderGroupedResults(data.groups, q);
    if(data.skipped?.length) document.getElementById('resultMeta').textContent += ` — مستبعد لغياب أعمدة الفلتر: ${data.skipped.join('، ')}`;
    return;
  }

  if(state.filters.some(f=>f.scope!=='all'&&!state.columns.includes(f.col))){ area.innerHTML='<div class="empty">الفلاتر محفوظة، لكن أحد أعمدتها غير موجود في هذا الملف. اختر العمود المقابل أو احذف الشرط لتطبيق البحث.</div>'; document.getElementById('pager').classList.add('hidden'); return; }
  const params = new URLSearchParams({q, limit: state.limit, offset: state.offset, filters: JSON.stringify(state.filters), db_id:state.dbId});
  const r = await fetch('/api/search?' + params.toString());
  const data = await r.json();
    if(requestId !== searchRequestId || dbId!==state.dbId) return;
  if(data.error){ area.innerHTML = `<div class="empty">${escapeHTML(data.error)}</div>`; return; }
  state.total = data.total;
  renderResults(data.rows, data.total);
  } catch(error){
    if(requestId === searchRequestId) document.getElementById('resultsArea').textContent = error.message || 'تعذر تحميل النتائج. حاول مرة أخرى.';
  }
}

function renderCardHTML(row){
  const protectedCols = row._protected || [];
  let html = `<div class="rcard"><span class="tab">#${row.id}</span>`;
  for(const k in row){
    if(k === 'id' || k === '_protected') continue;
    const isLocal = state.localColumns.includes(k);
    if(!row[k] && !isLocal) continue;
    const isProtected = protectedCols.includes(k);
    let cls = 'fvalue' + (isLocal ? ' local-col' : '') + (isProtected ? ' protected-col' : '');
    const lockIcon = isProtected
      ? `<span class="unlock-btn" data-db="${escapeHTML(state.dbId)}" data-id="${row.id}" data-col="${escapeHTML(k)}" title="إلغاء الحماية عن الخلية دي">🔒</span>` : '';
    html += `<div class="rfield"><span class="flabel">${fieldLabel(k)}</span><span class="fval-wrap">${lockIcon}<span dir="${cellDirection(row[k])}" class="${cls}" data-id="${row.id}" data-col="${escapeHTML(k)}" data-db="${escapeHTML(state.dbId)}" data-local="${isLocal}">${escapeHTML(row[k])}</span></span></div>`;
  }
  html += '</div>';
  return html;
}

function attachEditHandlers(container){
  container.querySelectorAll('.unlock-btn').forEach(btn => {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const rowId = btn.dataset.id, col = btn.dataset.col;
      await fetch('/api/clear_override', {method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({db_id:btn.dataset.db,id: Number(rowId), col})});
      btn.remove();
      const span = container.querySelector(`.fvalue[data-id="${rowId}"][data-col="${CSS.escape(col)}"]`);
      if(span) span.classList.remove('protected-col');
    });
  });
  container.querySelectorAll('.fvalue').forEach(span => {
    span.addEventListener('click', function onClick(){
      if(span.querySelector('textarea')) return;
      const rowId = span.dataset.id, col = span.dataset.col, isLocal = span.dataset.local === 'true';
      const oldVal = span.textContent;
      span.innerHTML = '';
      const input = document.createElement('textarea');
      input.className = 'fvalue-input'; input.value = oldVal;
      input.dir = cellDirection(oldVal);
      input.addEventListener('input',()=>{input.dir=cellDirection(input.value);});
      span.appendChild(input);
      input.focus(); input.select();

      let finished = false;
      const save = async () => {
        if(finished) return;
        finished = true;
        const newVal = input.value;
        span.classList.add('saving');
        span.textContent = newVal;
        try{
          const r = await fetch('/api/update_cell', {method:'POST', headers:{'Content-Type':'application/json'},
            body: JSON.stringify({db_id:span.dataset.db,id: Number(rowId), col, value: newVal})});
          const res = await r.json();
          if(res.error){ alert(res.error); span.textContent = oldVal; }
          else if(!isLocal && !span.classList.contains('protected-col')){
            span.classList.add('protected-col');
            const wrap = span.closest('.fval-wrap');
            if(wrap && !wrap.querySelector('.unlock-btn')){
              const lock = document.createElement('span');
              lock.className = 'unlock-btn'; lock.title = 'إلغاء الحماية عن الخلية دي'; lock.textContent = '🔒';
              lock.dataset.id = rowId; lock.dataset.col = col; lock.dataset.db = span.dataset.db;
              wrap.insertBefore(lock, span);
              attachEditHandlers(wrap);
            }
          }
        } catch(e){ span.textContent = oldVal; }
        span.classList.remove('saving');
        span.dir = cellDirection(span.textContent);
      };
      input.addEventListener('blur', save);
      input.addEventListener('keydown', (e) => {
        if(e.key === 'Enter' && (e.ctrlKey || e.metaKey)){ e.preventDefault(); input.blur(); }
        if(e.key === 'Escape'){ finished = true; span.textContent = oldVal; }
      });
    });
  });
}

function renderGroupedResults(groups, q){
  const area = document.getElementById('resultsArea');
  document.getElementById('resultMeta').textContent = groups.length ? `ظهرت في ${groups.length} ملف` : '';
  if(!groups.length){
    area.innerHTML = '<div class="empty">مفيش نتائج مطابقة في أي ملف.</div>';
    return;
  }
  let html = '';
  for(const g of groups){
    html += `<div class="file-group">
      <div class="file-group-header"><b>${escapeHTML(g.name)}</b><span>${g.total} نتيجة — عرض أول ${g.rows.length}</span><button class="link-btn" data-open-file="${escapeHTML(g.db_id)}">فتح الملف</button></div>
      <div class="results-grid">`;
    for(const row of g.rows) html += renderCardHTML(row);
    html += '</div></div>';
  }
  area.innerHTML = html;
  area.querySelectorAll('[data-open-file]').forEach(btn=>btn.addEventListener('click',()=>{
    state.dbId=btn.dataset.openFile;
    document.getElementById('allFilesToggle').checked=false;
    loadStats();
  }));
}

document.getElementById('prevBtn').addEventListener('click', () => {
  state.offset = Math.max(0, state.offset - state.limit);
  runSearch();
});
document.getElementById('nextBtn').addEventListener('click', () => {
  state.offset += state.limit;
  runSearch();
});

document.getElementById('syncBtn').addEventListener('click', async () => {
  const btn = document.getElementById('syncBtn');
  btn.disabled = true; btn.classList.add('spinning');
  const banner = document.getElementById('syncBanner');
  banner.className = 'banner';
  try{
    const s = await runJob('/api/sync',{db_id:state.dbId});
    if(s.error){
      banner.textContent = s.error; banner.classList.add('err','show');
    } else if(s.skipped){
      banner.textContent = 'الملف لم يتغيّر — القاعدة محدّثة بالفعل.'; banner.classList.add('ok','show');
    } else {
      banner.textContent = `تمت المزامنة: ${s.added} صف جديد، ${s.updated} اتحدّث، ${s.removed} اتشال.`;
      banner.classList.add('ok','show');
    }
  } catch(e){
    banner.textContent = e.message || 'حصل خطأ أثناء المزامنة.'; banner.classList.add('err','show');
  }
  btn.disabled = false; btn.classList.remove('spinning');
  loadStats();
  setTimeout(() => banner.classList.remove('show'), 6000);
});

document.getElementById('exportBtn').addEventListener('click', () => {
  const params = new URLSearchParams({q: state.query, filters: JSON.stringify(state.filters), db_id:state.dbId,all_files:document.getElementById('allFilesToggle').checked?'1':'0'});
  window.location = '/api/export?' + params.toString();
});

// ------------------ إضافة عمود جديد ------------------
document.getElementById('addColumnBtn').addEventListener('click', () => {
  document.getElementById('newColNameInput').value = '';
  document.getElementById('newColDefaultInput').value = '';
  document.getElementById('addColumnError').classList.remove('show');
  document.getElementById('addColumnOverlay').classList.remove('hidden');
});
document.getElementById('addColumnCloseBtn').addEventListener('click', () => {
  document.getElementById('addColumnOverlay').classList.add('hidden');
});
document.getElementById('addColumnConfirmBtn').addEventListener('click', async () => {
  const btn = document.getElementById('addColumnConfirmBtn');
  const err = document.getElementById('addColumnError');
  err.classList.remove('show');
  const name = document.getElementById('newColNameInput').value.trim();
  const def = document.getElementById('newColDefaultInput').value.trim();
  if(!name){ err.textContent = 'اكتب اسم العمود.'; err.classList.add('show'); return; }
  btn.disabled = true; btn.textContent = 'بيضيف...';
  try{
    const r = await fetch('/api/add_column', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({db_id:state.dbId,name, default: def})});
    const res = await r.json();
    if(res.error){ err.textContent = res.error; err.classList.add('show'); }
    else{
      document.getElementById('addColumnOverlay').classList.add('hidden');
      loadStats();
    }
  } catch(e){
    err.textContent = e.message || 'حصل خطأ غير متوقع.'; err.classList.add('show');
  }
  btn.disabled = false; btn.textContent = 'إضافة العمود';
});

// ------------------ مقارنة ملفين ------------------
const compareColumns={A:[],B:[]};
let savedComparison={};
try{savedComparison=JSON.parse(localStorage.getItem('excel-comparison-settings')||'{}')||{};}catch(_){}
let comparisonRestored=false;
function rememberComparison(){
  const settings={};
  for(const id of ['compareFileA','compareFileB','compareKeyA','compareKeyB','compareFieldA','compareFieldB','compareMode','compareThreshold','compareQuery']) settings[id]=document.getElementById(id).value;
  for(const side of ['A','B']) settings['filters'+side]=readFilterRows(document.getElementById('compareFilters'+side));
  Object.assign(settings,quantitySettings());
  settings.compoundKeys=compoundKeySettings();
  try{localStorage.setItem('excel-comparison-settings',JSON.stringify(settings));}catch(_){}
}
document.getElementById('compareOverlay').addEventListener('input',rememberComparison);
document.getElementById('compareOverlay').addEventListener('change',rememberComparison);
for(const side of ['A','B']) document.getElementById('addCompareFilter'+side).addEventListener('click',()=>{
  appendFilterRow(document.getElementById('compareFilters'+side),compareColumns[side],{},rememberComparison);rememberComparison();
});
document.getElementById('useSearchFilters').addEventListener('click',()=>{
  for(const side of ['A','B']){
    const area=document.getElementById('compareFilters'+side);area.replaceChildren();
    state.filters.forEach(f=>appendFilterRow(area,compareColumns[side],f,rememberComparison));
  }
  rememberComparison();
});
document.getElementById('copyCompareFilters').addEventListener('click',()=>{
  const filters=readFilterRows(document.getElementById('compareFiltersA'));
  const area=document.getElementById('compareFiltersB');area.replaceChildren();
  filters.forEach(f=>appendFilterRow(area,compareColumns.B,f,rememberComparison));rememberComparison();
});
document.getElementById('compareBtn').addEventListener('click', openCompareOverlay);
document.getElementById('compareCloseBtn').addEventListener('click', () => {
  document.getElementById('compareOverlay').classList.add('hidden');
});
document.getElementById('compareResultsCloseBtn').addEventListener('click', async () => {
  if(hasUnsavedReviewNotes() && !confirm('توجد ملاحظات لم تحفظ. إغلاق دون حفظها؟'))return;
  if(lastCompareResult) await postJSON('/api/compare_close',{comparison_id:lastCompareResult.comparison_id});
  lastCompareResult = null;
  comparisonPageRequest++;
  document.getElementById('compareResultsOverlay').classList.add('hidden');
});

async function openCompareOverlay(){
  document.getElementById('compareError').classList.remove('show');
  const r = await fetch('/api/databases');
  const data = await r.json();
  const selA = document.getElementById('compareFileA');
  const selB = document.getElementById('compareFileB');
  const previousA=selA.value||savedComparison.compareFileA||state.dbId, previousB=selB.value||savedComparison.compareFileB;
  const opts = data.databases.map(d => `<option value="${d.id}">${escapeHTML(d.name)} (${d.row_count ?? 0})</option>`).join('');
  selA.innerHTML = opts;
  selB.innerHTML = opts;
  if(data.databases.length > 1) selB.selectedIndex = 1;
  if(data.databases.some(d=>d.id===previousA)) selA.value=previousA;
  if(data.databases.some(d=>d.id===previousB)) selB.value=previousB;
  if(selA.value===selB.value && data.databases.length>1) selB.value=data.databases.find(d=>d.id!==selA.value).id;
  await Promise.all([loadKeyOptions('compareFileA','compareKeyA'), loadKeyOptions('compareFileB','compareKeyB')]);
  if(!comparisonRestored){
    for(const id of ['compareKeyA','compareKeyB','compareFieldA','compareFieldB','compareMode','compareThreshold','compareQuery']){
      const el=document.getElementById(id),value=savedComparison[id];
      if(typeof value==='string' && (el.tagName!=='SELECT'||[...el.options].some(o=>o.value===value))) el.value=value;
    }
    for(const side of ['A','B']) refreshFilterOptions(document.getElementById('compareFilters'+side),compareColumns[side],Array.isArray(savedComparison['filters'+side])?savedComparison['filters'+side]:[],rememberComparison);
    comparisonRestored=true;
    restoreQuantitySettings(savedComparison);
    (Array.isArray(savedComparison.compoundKeys)?savedComparison.compoundKeys:[]).filter(p=>p&&typeof p==='object').slice(0,5).forEach(addCompoundKey);
    document.getElementById('compareMode').dispatchEvent(new Event('change'));
  }
  document.getElementById('compareOverlay').classList.remove('hidden');
}

async function loadKeyOptions(fileSelectId, keySelectId){
  const dbId = document.getElementById(fileSelectId).value;
  const r = await fetch('/api/columns?' + new URLSearchParams({db_id: dbId}).toString());
  const data = await r.json();
  if(document.getElementById(fileSelectId).value!==dbId) return;
  const side=fileSelectId.endsWith('A')?'A':'B';
  compareColumns[side]=data.columns||[];
  for(const id of [keySelectId,'compareField'+side]){
    const select=document.getElementById(id),previous=select.value;
    select.innerHTML=(id===keySelectId?'':'<option value="">كل الأعمدة المشتركة</option>') + compareColumns[side].map(c=>`<option value="${escapeHTML(c)}">${fieldLabel(c)}</option>`).join('');
    if(compareColumns[side].includes(previous)) select.value=previous;
  }
  refreshFilterOptions(document.getElementById('compareFilters'+side),compareColumns[side],[],rememberComparison);
  refreshQuantityColumns(side);
  refreshCompoundColumns(side);
}
document.getElementById('compareFileA').addEventListener('change', () => loadKeyOptions('compareFileA','compareKeyA'));
document.getElementById('compareFileB').addEventListener('change', () => loadKeyOptions('compareFileB','compareKeyB'));

document.getElementById('runCompareBtn').addEventListener('click', async () => {
  const btn = document.getElementById('runCompareBtn');
  const err = document.getElementById('compareError');
  err.classList.remove('show');
  const payload = {
    db_id_a: document.getElementById('compareFileA').value,
    db_id_b: document.getElementById('compareFileB').value,
    key_col_a: document.getElementById('compareKeyA').value,
    key_col_b: document.getElementById('compareKeyB').value,
    query: document.getElementById('compareQuery').value.trim(),
    mode: document.getElementById('compareMode').value,
    threshold: Number(document.getElementById('compareThreshold').value),
    title:document.getElementById('comparisonTitle').value.trim(),
    filters_a:readFilterRows(document.getElementById('compareFiltersA')),
    filters_b:readFilterRows(document.getElementById('compareFiltersB')),
    field_a:document.getElementById('compareMode').value==='key'?document.getElementById('compareFieldA').value:'',
    field_b:document.getElementById('compareMode').value==='key'?document.getElementById('compareFieldB').value:'',
  };
  if(payload.db_id_a === payload.db_id_b){ err.textContent = 'اختار ملفين مختلفين.'; err.classList.add('show'); return; }
  if(payload.mode==='key'){
    const extra=compoundKeySettings();
    payload.key_columns_a=[payload.key_col_a,...extra.map(p=>p.a)];payload.key_columns_b=[payload.key_col_b,...extra.map(p=>p.b)];
    if(extra.some(p=>!compareColumns.A.includes(p.a)||!compareColumns.B.includes(p.b))){err.textContent='اختر كل أعمدة الربط الإضافية في الملفين.';err.classList.add('show');return;}
  }
  const quantity=quantitySettings();
  if(quantity.quantityMode){
    Object.assign(payload,{quantity_mode:quantity.quantityMode,quantity_pairs:quantity.quantityPairs,
      display_columns_a:quantity.displayA,display_columns_b:quantity.displayB,field_a:'',field_b:''});
    if(!quantity.quantityPairs.length || quantity.quantityPairs.some(p=>!compareColumns.A.includes(p.a)||!compareColumns.B.includes(p.b))){
      err.textContent='أضف أعمدة السنوات واختر العمود المقابل في كل ملف.';err.classList.add('show');return;
    }
  }
  rememberComparison();
  btn.disabled = true; btn.textContent = 'بيقارن...';
  try{
    const result = await runJob('/api/compare',payload);
    if(result.error){ err.textContent = result.error; err.classList.add('show'); }
    else{
      document.getElementById('compareOverlay').classList.add('hidden');
      showCompareResults(result, payload);
    }
  } catch(e){
    err.textContent = e.message || 'حصل خطأ غير متوقع.'; err.classList.add('show');
  }
  btn.disabled = false; btn.textContent = 'قارن الملفين';
});

let lastCompareResult = null, lastComparePayload = null;
let comparisonPageRequest = 0;
document.getElementById('compareMode').addEventListener('change',()=>{
  const similar = document.getElementById('compareMode').value === 'similarity';
  document.getElementById('thresholdField').classList.toggle('hidden',!similar);
  document.getElementById('compareValueFields').classList.toggle('hidden',similar);
  document.querySelectorAll('.compare-key-label').forEach(el=>el.textContent=similar?'عمود الوصف المراد مقارنة تشابهه':'عمود المفتاح لربط الصفوف');
  updateQuantityVisibility();
  document.getElementById('compoundKeyOptions').classList.toggle('hidden',similar);
});
function formatQtyNumber(value){
  if(value===null||value===undefined) return '—';
  const n=Number(value);
  if(!isFinite(n)) return '—';
  return String(Math.round(n*1000)/1000);
}
function formatQtyDelta(value){
  if(value===null||value===undefined) return '—';
  const n=Number(value);
  if(!isFinite(n)) return '—';
  const rounded=Math.round(n*1000)/1000;
  if(rounded===0) return '0';
  return (rounded>0?'+':'')+String(rounded);
}
function qtyDeltaClass(value){
  if(value===null||value===undefined) return 'qty-zero';
  const n=Number(value);
  if(!isFinite(n)||n===0) return 'qty-zero';
  return n>0?'qty-pos':'qty-neg';
}
function renderQtyHTML(qty){
  let rows='';
  for(const col of qty.cols){
    const a=(col.a_num!==null && col.a_num!==undefined)?formatQtyNumber(col.a_num):(col.a||'—');
    const b=(col.b_num!==null && col.b_num!==undefined)?formatQtyNumber(col.b_num):(col.b||'—');
    rows+=`<tr><td class="qname">${fieldLabel(col.name||'')}</td><td>${escapeHTML(a)}</td><td>${escapeHTML(b)}</td><td class="qty-delta ${qtyDeltaClass(col.delta)}">${formatQtyDelta(col.delta)}</td></tr>`;
  }
  rows+=`<tr class="qtotal"><td>الإجمالي</td><td>${formatQtyNumber(qty.a_total)}</td><td>${formatQtyNumber(qty.b_total)}</td><td class="qty-delta ${qtyDeltaClass(qty.delta_total)}">${formatQtyDelta(qty.delta_total)}</td></tr>`;
  return `<div class="qty-summary"><div class="qty-head"><b>الكميات</b><span>الفرق = الثاني ناقص الأول</span></div>${qty.status?`<p class="quantity-status">${escapeHTML(qty.status)}</p><p class="edit-hint">الإجمالي للقيم الرقمية الموجودة في الأعمدة المختارة فقط. الفراغ لا يعني صفرًا.</p>`:''}<table><thead><tr><th>العمود</th><th>الأول</th><th>الثاني</th><th>الفرق</th></tr></thead><tbody>${rows}</tbody></table></div>`;
}
function showCompareResults(result, payload){
  lastCompareResult = result; lastComparePayload = payload;
  document.getElementById('compareResultsTitle').textContent = result.session?.title || `مقارنة: ${result.name_a} ⇔ ${result.name_b}`;
  const labels = result.mode === 'similarity' ? {suggested:'اقتراحات للمراجعة',accepted:'مقبول يدويًا',rejected:'مرفوض'} : {changed:'مختلف',identical:'متطابق',only_a:'فقط في الأول',only_b:'فقط في الثاني'};
  if(result.mode==='key' && result.quantity_mode==='total'){labels.identical='إجمالي متساوٍ';labels.changed='إجمالي مختلف / يحتاج مراجعة';}
  document.getElementById('compareSummary').innerHTML = Object.entries(labels).map(([key,label])=>`<div class="compare-chip identical">${label}: ${result.counts[key] || 0}</div>`).join('');
  const qtyWrap = document.getElementById('compareQtyToggleWrap');
  const qtyBox = document.getElementById('compareQtyOnly');
  if(result.has_qty){
    const qtyTotal = Object.keys(labels).reduce((sum,key)=>sum+((result.qty_counts||{})[key]||0),0);
    document.getElementById('compareSummary').innerHTML += `<div class="compare-chip changed">بنود بها فرق كميات: ${qtyTotal}</div>`;
    qtyWrap.classList.remove('hidden');
    qtyBox.checked = !!result.qty_only;
    document.getElementById('compareQtyCount').textContent = (result.qty_counts||{})[result.section] || 0;
  }else{
    qtyWrap.classList.add('hidden');
    qtyBox.checked = false;
  }
  document.getElementById('compareNotice').textContent = `النطاق بعد الفلترة: الأول ${result.filtered_count_a} من ${result.source_count_a}، الثاني ${result.filtered_count_b} من ${result.source_count_b}. النتائج تخص النطاق المختار فقط. صفوف مستبعدة لفراغ ${result.mode==='similarity'?'الوصف':'المفتاح'}: الأول ${result.blank_a}، الثاني ${result.blank_b}. ` + (result.mode==='similarity' ? `النسب لتشابه النص فقط. ${result.without_suggestion_a} صفًا في نطاق البحث بالملف الأول بدون اقتراح. قرارات المراجعة محفوظة تلقائيًا على الجهاز.` : 'المفاتيح متطابقة حرفيًا بعد إزالة المسافات الطرفية.');
  if(result.quantity_mode) document.getElementById('compareNotice').textContent+=' مقارنة الكميات: '+({yearly:'سنة بسنة مع الإجمالي',total:'حسب الإجمالي؛ توزيع السنوات يظهر للتوضيح',custom:'حسب أزواج السنوات المختارة'}[result.quantity_mode])+'. أعمدة العرض لا تؤثر في حكم مقارنة الكميات.';
  if(result.session) document.getElementById('compareNotice').textContent+=' لقطة محفوظة بتاريخ '+new Date(result.session.created_at).toLocaleString('ar-SA')+'؛ تغييرات ملفات المصدر اللاحقة لا تدخل هذه الجلسة.';
  document.getElementById('compareTabs').innerHTML = Object.entries(labels).map(([key,label])=>`<button class="compare-tab ${key===result.section?'active':''}" data-tab="${key}">${label} (${result.counts[key]||0})</button>`).join('');
  document.querySelectorAll('.compare-tab').forEach(btn=>btn.addEventListener('click',()=>loadComparisonPage(btn.dataset.tab,0)));
  const body = document.getElementById('compareBody');
  body.innerHTML = result.items.length ? result.items.map(item=>{
    let html = `<div class="diff-card"><div class="diff-key">${escapeHTML(item.key)} ${item.score==null?'':`— تشابه نصي ${item.score}%`}</div>`;
    html += renderReviewFields(item,result);
    if(item.qty && item.qty.cols && item.qty.cols.length) html += renderQtyHTML(item.qty);
    return html+'</div>';
  }).join('') : '<div class="empty">لا توجد نتائج في هذا القسم.</div>';
  body.querySelectorAll('[data-decision]').forEach(btn=>btn.addEventListener('click',async()=>{
    const buttons=[...body.querySelectorAll('[data-decision]')];buttons.forEach(b=>b.disabled=true);
    try{
      // Preserve all edited notes on this page before re-rendering after a decision.
      for(const input of body.querySelectorAll('.review-note')){
        if(input.value===input.defaultValue)continue;
        const current=result.items.find(item=>item.id===Number(input.dataset.note));
        await postJSON('/api/compare_decision',{comparison_id:result.comparison_id,id:current.id,decision:current.status,note:input.value});
        input.defaultValue=input.value;
      }
      const note=body.querySelector(`[data-note="${btn.dataset.result}"]`).value;
      await postJSON('/api/compare_decision',{comparison_id:result.comparison_id,id:Number(btn.dataset.result),decision:btn.dataset.decision,note});
      document.getElementById('reviewSaveStatus').textContent='تم حفظ القرار والملاحظة على الجهاز.';
      await loadComparisonPage(result.section,result.offset,true);
    }catch(e){ alert(e.message); }
    finally{buttons.forEach(b=>b.disabled=false);}
  }));
  document.getElementById('comparePrev').disabled = result.offset===0;
  document.getElementById('compareNext').disabled = result.offset+result.limit>=result.total;
  document.getElementById('comparePageLabel').textContent = result.total?`${result.offset+1}–${Math.min(result.offset+result.items.length,result.total)} من ${result.total}`:'0 نتيجة';
  document.getElementById('compareResultsOverlay').classList.remove('hidden');
}
async function loadComparisonPage(section,offset,saved=false){
  if(!saved && hasUnsavedReviewNotes() && !confirm('توجد ملاحظات لم تحفظ. الانتقال دون حفظها؟'))return;
  const ticket=++comparisonPageRequest;
  try{
    const result=await postJSON('/api/compare_page',{comparison_id:lastCompareResult.comparison_id,section,offset,qty_only:document.getElementById('compareQtyOnly').checked});
    if(ticket===comparisonPageRequest) showCompareResults(result,lastComparePayload);
  }catch(e){alert(e.message);}
}
document.getElementById('comparePrev').addEventListener('click',()=>loadComparisonPage(lastCompareResult.section,Math.max(0,lastCompareResult.offset-20)));
document.getElementById('compareNext').addEventListener('click',()=>loadComparisonPage(lastCompareResult.section,lastCompareResult.offset+20));
document.getElementById('compareQtyOnly').addEventListener('change',()=>{
  if(!lastCompareResult) return;
  loadComparisonPage(lastCompareResult.section,0);
});
document.getElementById('compareExportBtn').addEventListener('click', async () => {
  if(!lastCompareResult) return;
  try{
    const r = await fetch('/api/compare_export',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({comparison_id:lastCompareResult.comparison_id})});
    if(!r.ok) throw new Error((await r.json()).error || 'تعذر التصدير');
    const url = URL.createObjectURL(await r.blob());
    const a = document.createElement('a');a.href=url;a.download='تقرير_المقارنة.xlsx';a.click();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
  }catch(e){alert(e.message);}
});

document.getElementById('setupBtn').addEventListener('click', async () => {
  const btn = document.getElementById('setupBtn');
  const err = document.getElementById('setupError');
  err.classList.remove('show');
  const name = document.getElementById('nameInput').value.trim();
  const excel_path = document.getElementById('excelPathInput').value.trim();
  const unique_col = document.getElementById('uniqueColInput').value.trim();
  const header_row = document.getElementById('headerRowInput').value.trim();
  const last_column = document.getElementById('lastColumnInput').value.trim();
  if(!excel_path){ err.textContent = 'اكتب مسار الملف أولاً.'; err.classList.add('show'); return; }
  btn.disabled = true; btn.textContent = 'جاري الفهرسة...';
  try{
    const s = await runJob('/api/setup',{name,excel_path,unique_col,header_row,last_column});
    if(s.error){ err.textContent = s.error; err.classList.add('show'); }
    else { state.dbId=s.db_id; loadStats(); }
  } catch(e){
    err.textContent = e.message || 'حصل خطأ غير متوقع.'; err.classList.add('show');
  }
  btn.disabled = false; btn.textContent = 'ابدأ الفهرسة';
});

loadStats();
