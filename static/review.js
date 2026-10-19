function compoundKeySettings(){
  return [...document.querySelectorAll('#compoundKeyPairs .compound-key')].map(row=>({a:row.querySelector('.compound-a').value,b:row.querySelector('.compound-b').value}));
}
function addCompoundKey(pair={}){
  const area=document.getElementById('compoundKeyPairs');if(area.children.length>=5)return;
  const row=document.createElement('div');row.className='quantity-pair compound-key';
  row.innerHTML='<select class="compound-a" aria-label="عمود ربط إضافي في الأول"></select><span>مقابل</span><select class="compound-b" aria-label="عمود ربط إضافي في الثاني"></select><button type="button" class="link-btn">حذف</button>';
  for(const side of ['a','b'])quantitySelectOptions(row.querySelector('.compound-'+side),compareColumns[side.toUpperCase()],pair[side]);
  row.querySelector('button').onclick=()=>{row.remove();rememberComparison();};area.appendChild(row);
}
function refreshCompoundColumns(side){
  for(const row of document.querySelectorAll('.compound-key')){
    const select=row.querySelector('.compound-'+side.toLowerCase());quantitySelectOptions(select,compareColumns[side],select.value);
  }
}
document.getElementById('addCompoundKey').onclick=()=>{addCompoundKey();rememberComparison();};
function renderReviewFields(item,result){
  const marked=parts=>parts.map(part=>part.changed?`<mark>${escapeHTML(part.text)}</mark>`:escapeHTML(part.text)).join('');
  let html=`<div class="review-file-head"><b>${escapeHTML(result.name_a)} — صف ${item.row_id_a??'—'}</b><b>${escapeHTML(result.name_b)} — صف ${item.row_id_b??'—'}</b></div>`;
  for(const field of item.review_fields||[]){
    html+=`<div class="review-pair">${['a','b'].map(side=>`<div class="review-cell review-${side}"><b>${field[side+'_column']?fieldLabel(field[side+'_column']):'لا يوجد عمود مقابل'}</b><div dir="${cellDirection(field[side])}" class="review-value">${marked(field[side+'_segments'])||'<span class="edit-hint">فارغ</span>'}</div></div>`).join('')}</div>`;
  }
  html+=`<details class="review-details"><summary>عرض كل بيانات الصفين</summary><div class="compare-grid"><div>${renderCardHTML({...item.row_a,id:item.row_id_a??'—'})}</div><div>${renderCardHTML({...item.row_b,id:item.row_id_b??'—'})}</div></div></details>`;
  if(result.mode==='similarity'){
    const label={suggested:'لم تتم المراجعة',accepted:'مقبول يدويًا',rejected:'مرفوض'}[item.status];
    html+=`<p class="review-state">الحالة: ${label} — التمييز اللوني يوضح اختلاف الكلمات، وليس حكمًا على تكافؤ المواصفات.</p><label class="review-note-label">ملاحظة المراجعة<textarea class="review-note" maxlength="2000" data-note="${item.id}" dir="auto" placeholder="سبب القبول أو الرفض...">${escapeHTML(item.note||'')}</textarea></label><div class="filter-actions"><button class="link-btn" data-decision="accepted" data-result="${item.id}">قبول التطابق</button><button class="link-btn" data-decision="rejected" data-result="${item.id}">رفض</button><button class="link-btn" data-decision="suggested" data-result="${item.id}">إعادة للمراجعة</button><button class="link-btn" data-decision="${item.status}" data-result="${item.id}">حفظ الملاحظة</button></div>`;
  }
  return html;
}
function hasUnsavedReviewNotes(){return [...document.querySelectorAll('.review-note')].some(el=>el.value!==el.defaultValue);}
document.getElementById('savedComparisonsBtn').onclick=loadSavedComparisons;
document.getElementById('setupSavedComparisonsBtn').onclick=loadSavedComparisons;
document.getElementById('savedComparisonsClose').onclick=()=>document.getElementById('savedComparisonsOverlay').classList.add('hidden');
async function loadSavedComparisons(){
  const overlay=document.getElementById('savedComparisonsOverlay'),area=document.getElementById('savedComparisonsList');
  overlay.classList.remove('hidden');area.textContent='تحميل المقارنات...';
  try{
    const response=await fetch('/api/comparisons'),data=await response.json();
    if(!response.ok)throw Error(data.error||'تعذر تحميل المقارنات');
    area.innerHTML=data.sessions.length?data.sessions.map(s=>`<article class="saved-comparison"><b>${escapeHTML(s.title)}</b><p>${s.created_at?escapeHTML(new Date(s.created_at).toLocaleString('ar-SA')):''}</p><p>${s.error?escapeHTML(s.error):`مقبول: ${s.counts.accepted||0} · مرفوض: ${s.counts.rejected||0} · للمراجعة: ${s.counts.suggested||0} · مختلف: ${s.counts.changed||0}`}</p><div class="filter-actions"><button class="link-btn" data-resume="${escapeHTML(s.id)}" ${s.error?'disabled':''}>استكمال المراجعة</button><button class="link-btn" data-rename="${escapeHTML(s.id)}">تغيير الاسم</button><button class="link-btn" data-delete-session="${escapeHTML(s.id)}">حذف الجلسة</button></div></article>`).join(''):'لا توجد مقارنات محفوظة حتى الآن.';
    area.querySelectorAll('[data-resume]').forEach(btn=>btn.onclick=async()=>{
      try{const result=await postJSON('/api/compare_page',{comparison_id:btn.dataset.resume});overlay.classList.add('hidden');showCompareResults(result,null);}catch(e){alert(e.message);}
    });
    area.querySelectorAll('[data-rename]').forEach(btn=>btn.onclick=async()=>{
      const title=prompt('اسم المقارنة',data.sessions.find(s=>s.id===btn.dataset.rename).title);if(title===null)return;
      try{await postJSON('/api/compare_rename',{comparison_id:btn.dataset.rename,title});await loadSavedComparisons();}catch(e){alert(e.message);}
    });
    area.querySelectorAll('[data-delete-session]').forEach(btn=>btn.onclick=async()=>{
      if(!confirm('حذف لقطة المقارنة وقراراتها نهائيًا؟ ملفات Excel الأصلية لن تتغير.'))return;
      try{await postJSON('/api/compare_delete',{comparison_id:btn.dataset.deleteSession});await loadSavedComparisons();}catch(e){alert(e.message);}
    });
  }catch(e){area.textContent=e.message;}
}
