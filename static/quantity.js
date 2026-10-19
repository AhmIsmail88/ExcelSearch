function quantitySettings(){
  return {quantityMode:document.getElementById('quantityMode').value,
    quantityPairs:[...document.querySelectorAll('#quantityPairs .quantity-pair')].map(row=>({a:row.querySelector('.quantity-a').value,b:row.querySelector('.quantity-b').value})),
    displayA:[...document.querySelectorAll('#displayColumnsA input:checked')].map(el=>el.value),
    displayB:[...document.querySelectorAll('#displayColumnsB input:checked')].map(el=>el.value)};
}
function quantitySelectOptions(select, columns, value){
  const choices=[...columns];if(value && !choices.includes(value))choices.unshift(value);
  select.innerHTML='<option value="">اختر العمود</option>'+choices.map(c=>`<option value="${escapeHTML(c)}">${fieldLabel(c)}${columns.includes(c)?'':' — غير موجود في الملف'}</option>`).join('');
  select.value=value||'';
}
function addQuantityPair(pair={}){
  const area=document.getElementById('quantityPairs');if(area.children.length>=12)return;
  const row=document.createElement('div');row.className='quantity-pair';
  row.innerHTML='<select class="quantity-a" aria-label="كمية الملف الأول"></select><span>مقابل</span><select class="quantity-b" aria-label="كمية الملف الثاني"></select><button type="button" class="link-btn">حذف</button>';
  quantitySelectOptions(row.querySelector('.quantity-a'),compareColumns.A,pair.a);
  quantitySelectOptions(row.querySelector('.quantity-b'),compareColumns.B,pair.b);
  row.querySelector('button').onclick=()=>{row.remove();rememberComparison();};area.appendChild(row);
}
function refreshQuantityColumns(side, selected){
  const area=document.getElementById('displayColumns'+side);
  selected=selected||[...area.querySelectorAll('input:checked')].map(el=>el.value);
  const cols=[...compareColumns[side],...selected.filter(c=>!compareColumns[side].includes(c))];
  area.innerHTML=cols.map(c=>`<label><input type="checkbox" value="${escapeHTML(c)}" ${selected.includes(c)?'checked':''}>${fieldLabel(c)}${compareColumns[side].includes(c)?'':' — غير موجود في الملف'}</label>`).join('');
  for(const row of document.querySelectorAll('#quantityPairs .quantity-pair')){
    const select=row.querySelector('.quantity-'+side.toLowerCase());quantitySelectOptions(select,compareColumns[side],select.value);
  }
}
function updateQuantityVisibility(){
  const quantity=!!document.getElementById('quantityMode').value;
  document.getElementById('quantityOptions').classList.toggle('hidden',!quantity);
  document.getElementById('compareValueFields').classList.toggle('hidden',quantity||document.getElementById('compareMode').value==='similarity');
}
function restoreQuantitySettings(saved){
  const mode=document.getElementById('quantityMode');
  if(['','yearly','total','custom'].includes(saved.quantityMode))mode.value=saved.quantityMode;
  document.getElementById('quantityPairs').replaceChildren();
  (Array.isArray(saved.quantityPairs)?saved.quantityPairs:[]).filter(p=>p&&typeof p==='object').slice(0,12).forEach(addQuantityPair);
  for(const side of ['A','B'])refreshQuantityColumns(side,Array.isArray(saved['display'+side])?saved['display'+side].filter(c=>typeof c==='string'):[]);
  updateQuantityVisibility();
}
document.getElementById('addQuantityPair').onclick=()=>{addQuantityPair();rememberComparison();};
document.getElementById('quantityMode').onchange=()=>{updateQuantityVisibility();rememberComparison();};
document.getElementById('suggestQuantityPairs').onclick=()=>{
  const normalized=c=>c.replace(/[أإآ]/g,'ا').replace(/ة/g,'ه').replace(/ى/g,'ي').replace(/[_\s]+/g,' ').trim().toLowerCase();
  const pairs=compareColumns.A.filter(c=>/سنه|كميه|year|quantity|qty/.test(normalized(c))).map(a=>({a,b:compareColumns.B.find(b=>normalized(a)===normalized(b))})).filter(p=>p.b);
  if(!pairs.length){addQuantityPair();}else{
    document.getElementById('quantityPairs').replaceChildren();pairs.slice(0,12).forEach(addQuantityPair);
  }
  rememberComparison();
};
