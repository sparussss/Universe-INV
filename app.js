const $ = (s, root=document) => root.querySelector(s);
const $$ = (s, root=document) => [...root.querySelectorAll(s)];
const money = new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',minimumFractionDigits:2});

const state = {
  products: new Map(),
  images: new Map(),
  items: [],
  insertAt: null,
  stream: null,
  installPrompt: null
};

const sample = {lotNo:'133685', artNo:'PT-37499', price:3092, unit:'PC', descriptions:['2.10Y750','1-AMCTOCT25x9-10.75ct','6-CDMRD(B2)-0.05ct']};
state.products.set(sample.lotNo,sample);

function normalizeKey(v){ return String(v ?? '').trim(); }
function normalizeArticle(v){ return normalizeKey(v).toUpperCase(); }
function field(row, names){
  const keys = Object.keys(row);
  for(const name of names){ const k=keys.find(x=>x.trim().toUpperCase()===name); if(k) return row[k]; }
  return '';
}
function setStatus(id,text,isError=false){ const el=$(id); el.textContent=text; el.style.background=isError?'#fff1f0':'#f4f7fb'; el.style.color=isError?'#b42318':'#4d596f'; }
function updateHeader(){
  $('#productCount').textContent=state.products.size;
  $('#imageCount').textContent=[...state.images.values()].reduce((a,b)=>a+b.length,0);
  $('#invoiceCount').textContent=state.items.length;
  $('#headerTotal').textContent=formatCurrency(calcTotals().total);
}
function formatCurrency(v){ const c=$('#currency').value || 'USD'; return new Intl.NumberFormat('en-US',{style:'currency',currency:c,minimumFractionDigits:2}).format(v||0); }

$('#invoiceDate').value = new Date().toISOString().slice(0,10);

$$('.tab').forEach(btn=>btn.addEventListener('click',()=>{
  $$('.tab').forEach(x=>x.classList.remove('active')); $$('.tab-panel').forEach(x=>x.classList.remove('active'));
  btn.classList.add('active'); $('#'+btn.dataset.tab).classList.add('active');
  if(btn.dataset.tab==='preview') renderPreview();
}));

$('#excelInput').addEventListener('change', async e=>{
  const file=e.target.files[0]; if(!file) return;
  if(typeof XLSX==='undefined'){ setStatus('#excelStatus','Excel 解析程式未載入。首次開啟 PWA 時請連接網絡，重新載入後再匯入。',true); return; }
  try{
    const data=await file.arrayBuffer();
    const wb=XLSX.read(data,{type:'array',cellDates:false});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{defval:'',raw:false});
    const imported=new Map(); const invalid=[];
    rows.forEach((row,i)=>{
      const lot=normalizeKey(field(row,['LOTNO','LOT NO','LOT.NO.']));
      const art=normalizeKey(field(row,['ARTNO','ARTICLE NO','ART NO']));
      const rawPrice=field(row,['PRICE','U PRICE','UPRICE']);
      const price=Number(String(rawPrice).replace(/[^0-9.-]/g,''));
      if(!lot||!art||!Number.isFinite(price)){ if(Object.values(row).some(Boolean)) invalid.push(i+2); return; }
      const descriptions=[];
      for(let n=1;n<=7;n++){ const d=normalizeKey(field(row,[`DESC${n}`,`DESC ${n}`])); if(d) descriptions.push(d); }
      const unit=normalizeKey(field(row,['UNIT']))||'PC';
      imported.set(lot,{lotNo:lot,artNo:art,price,unit,descriptions});
    });
    state.products=imported;
    setStatus('#excelStatus',`已匯入 ${imported.size} 件貨品${invalid.length?`；${invalid.length} 行資料未能讀取`:''}。檔案：${file.name}`);
    updateHeader();
  }catch(err){ console.error(err); setStatus('#excelStatus','匯入失敗：'+err.message,true); }
});

$('#imageInput').addEventListener('change', async e=>{
  const files=[...e.target.files]; if(!files.length) return;
  for(const file of files){
    const base=file.name.replace(/\.[^.]+$/,'').trim();
    const match=base.match(/^(.+?)(?:\s+([^\s].*))?$/);
    const artNo=normalizeArticle(match?.[1]);
    const variant=(match?.[2]||'Default').trim();
    const url=URL.createObjectURL(file);
    if(!state.images.has(artNo)) state.images.set(artNo,[]);
    const list=state.images.get(artNo);
    const idx=list.findIndex(x=>x.variant.toUpperCase()===variant.toUpperCase());
    const rec={artNo,variant,url,fileName:file.name};
    if(idx>=0) list[idx]=rec; else list.push(rec);
  }
  renderImageLibrary(); updateHeader();
  setStatus('#imageStatus',`已匯入 / 更新 ${files.length} 張圖片。圖片以 ARTNO 檔名配對。`);
});

function renderImageLibrary(){
  const box=$('#imageLibrary'); box.innerHTML='';
  [...state.images.entries()].slice(0,20).forEach(([art,list])=>list.forEach(img=>{
    const chip=document.createElement('div'); chip.className='image-chip';
    chip.innerHTML=`<img src="${img.url}"><span>${art} · ${img.variant}</span>`; box.appendChild(chip);
  }));
}
function getImages(artNo){ return state.images.get(normalizeArticle(artNo))||[]; }
function getSelectedImage(item){
  const imgs=getImages(item.artNo); if(!imgs.length) return null;
  return imgs.find(x=>x.variant===item.imageVariant)||imgs.find(x=>x.variant.toLowerCase()==='default')||imgs[0];
}
function placeholderSvg(art){ return `data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="240" height="240"><rect width="100%" height="100%" fill="#edf1f5"/><text x="50%" y="48%" text-anchor="middle" font-family="Arial" font-size="18" fill="#64748b">${art}</text><text x="50%" y="60%" text-anchor="middle" font-family="Arial" font-size="13" fill="#94a3b8">No image</text></svg>`)}`; }

$('#addLotBtn').addEventListener('click',()=>addByLot($('#lotInput').value));
$('#lotInput').addEventListener('keydown',e=>{if(e.key==='Enter') addByLot(e.currentTarget.value)});
function addByLot(raw){
  const lot=normalizeKey(raw); if(!lot) return;
  const product=state.products.get(lot);
  if(!product){ setStatus('#addMessage',`找不到 LOTNO ${lot}。請檢查 Excel 或重新輸入。`,true); return; }
  if(state.items.some(x=>x.lotNo===lot) && !confirm(`LOTNO ${lot} 已在 Invoice 內，仍要再次加入嗎？`)) return;
  const rate=Number($('#salesRate').value)||0;
  const unitPrice=Math.ceil(product.price*rate);
  const imgs=getImages(product.artNo);
  const defaultVariant=(imgs.find(x=>x.variant.toLowerCase()==='default')||imgs[0])?.variant||'Default';
  const item={id:crypto.randomUUID?.()||String(Date.now()+Math.random()),...product,qty:1,unitPrice,imageVariant:defaultVariant};
  if(state.insertAt===null) state.items.push(item); else {state.items.splice(state.insertAt,0,item); state.insertAt=null;}
  $('#lotInput').value=''; setStatus('#addMessage',`已加入 ${product.artNo} / LOTNO ${lot} / ${formatCurrency(unitPrice)}`);
  renderItems();
}

function renderItems(){
  const box=$('#invoiceItems'); box.innerHTML='';
  if(!state.items.length){box.className='invoice-items empty-state';box.textContent='尚未加入貨品。';updateTotals();return}
  box.className='invoice-items';
  state.items.forEach((item,index)=>{
    const node=$('#itemTemplate').content.firstElementChild.cloneNode(true); node.dataset.id=item.id;
    $('.item-artno',node).textContent=item.artNo; $('.item-lot',node).textContent=`LOTNO ${item.lotNo}`;
    $('.item-desc',node).textContent=item.descriptions.join('\n');
    $('.item-price-note',node).textContent=`${item.price}u × ${Number($('#salesRate').value||0)} → ${formatCurrency(item.unitPrice)}`;
    const img=getSelectedImage(item); $('.item-thumb',node).src=img?.url||placeholderSvg(item.artNo);
    const sel=$('.variant-select',node); const imgs=getImages(item.artNo);
    if(imgs.length){imgs.forEach(x=>{const o=document.createElement('option');o.value=x.variant;o.textContent=`圖片：${x.variant}`;o.selected=x.variant===item.imageVariant;sel.appendChild(o)})}
    else{sel.innerHTML='<option>沒有圖片</option>';sel.disabled=true}
    $('.qty-input',node).value=item.qty; $('.price-input',node).value=item.unitPrice;
    sel.addEventListener('change',e=>{item.imageVariant=e.target.value;renderItems()});
    $('.qty-input',node).addEventListener('change',e=>{item.qty=Math.max(1,Number(e.target.value)||1);updateTotals()});
    $('.price-input',node).addEventListener('change',e=>{item.unitPrice=Math.max(0,Math.ceil(Number(e.target.value)||0));updateTotals()});
    $('.delete-item',node).addEventListener('click',()=>{if(confirm(`刪除 ${item.artNo} / LOTNO ${item.lotNo}？`)){state.items.splice(index,1);renderItems()}});
    $('.insert-above',node).addEventListener('click',()=>prepareInsert(index)); $('.insert-below',node).addEventListener('click',()=>prepareInsert(index+1));
    node.addEventListener('dragstart',e=>e.dataTransfer.setData('text/plain',String(index)));
    node.addEventListener('dragover',e=>e.preventDefault());
    node.addEventListener('drop',e=>{e.preventDefault();const from=Number(e.dataTransfer.getData('text/plain'));const moved=state.items.splice(from,1)[0];let to=index;if(from<to)to--;state.items.splice(to,0,moved);renderItems()});
    box.appendChild(node);
  }); updateTotals();
}
function prepareInsert(index){ state.insertAt=index; $('#lotInput').focus(); setStatus('#addMessage',`下一件貨會插入第 ${index+1} 行。請掃 Barcode 或輸入 LOTNO。`); }
$('#clearInvoiceBtn').addEventListener('click',()=>{if(state.items.length&&confirm('確定清空目前 Invoice 草稿？')){state.items=[];renderItems()}});
$('#discountAmount').addEventListener('input',updateTotals); $('#currency').addEventListener('change',()=>{renderItems();renderPreview()}); $('#salesRate').addEventListener('change',()=>{state.items.forEach(x=>x.unitPrice=Math.ceil(x.price*(Number($('#salesRate').value)||0)));renderItems()});
function calcTotals(){ const qty=state.items.reduce((a,x)=>a+x.qty,0); const subtotal=state.items.reduce((a,x)=>a+x.qty*x.unitPrice,0); const discount=Math.max(0,Number($('#discountAmount').value)||0); return {qty,subtotal,discount,total:Math.max(0,subtotal-discount)}; }
function updateTotals(){ const t=calcTotals(); $('#totalQty').textContent=t.qty; $('#subtotal').textContent=formatCurrency(t.subtotal);$('#discountDisplay').textContent=formatCurrency(t.discount);$('#grandTotal').textContent=formatCurrency(t.total);updateHeader(); }

function numberToWords(n){
  n=Math.floor(n); if(n===0)return 'ZERO';
  const ones=['','ONE','TWO','THREE','FOUR','FIVE','SIX','SEVEN','EIGHT','NINE','TEN','ELEVEN','TWELVE','THIRTEEN','FOURTEEN','FIFTEEN','SIXTEEN','SEVENTEEN','EIGHTEEN','NINETEEN'];
  const tens=['','','TWENTY','THIRTY','FORTY','FIFTY','SIXTY','SEVENTY','EIGHTY','NINETY'];
  const under1000=x=>{let s='';if(x>=100){s+=ones[Math.floor(x/100)]+' HUNDRED ';x%=100}if(x>=20){s+=tens[Math.floor(x/10)]+' ';x%=10}if(x>0)s+=ones[x]+' ';return s.trim()};
  const scales=[[1e9,'BILLION'],[1e6,'MILLION'],[1e3,'THOUSAND'],[1,'']]; let out=[];
  for(const [v,name] of scales){if(n>=v){const part=Math.floor(n/v);n%=v;out.push(under1000(part)+(name?' '+name:''))}}return out.join(' ');
}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function renderPreview(){
  const t=calcTotals();
  const rows=state.items.map((x,i)=>`<tr><td>${i+1}</td><td><strong>Lot.No. : ${esc(x.lotNo)}</strong><br>${esc(x.artNo)}</td><td>${x.descriptions.map(esc).join('<br>')}</td><td class="num">${x.qty}</td><td>${esc(x.unit)}</td><td class="num">${formatCurrency(x.unitPrice)}</td><td class="num">${formatCurrency(x.qty*x.unitPrice)}</td></tr>`).join('');
  $('#invoiceDocument').innerHTML=`
    <div class="doc-title">Sales Invoice</div>
    <div class="doc-grid"><div class="doc-meta">No. : <strong>${esc($('#invoiceNo').value)}</strong><br>Invoice Date : ${esc($('#invoiceDate').value)}<br>Shipment Method : ${esc($('#shipmentMethod').value)}<br>Currency : ${esc($('#currency').value)}<br><br>Customer : <strong>${esc($('#customerName').value)}</strong><br>${esc($('#customerAddress').value).replace(/\n/g,'<br>')}</div><div class="doc-meta"><strong>Vender's Banker</strong><br>The Hong Kong &amp; Shanghai Banking Corporation Ltd.<br>Address : 41 Ma Tau Wai Road,Hung Hom,Kowloon,Hong Kong<br>A/C # : 012-593570-001<br>A/C Name : Universe Gems &amp; Jewellery Co.</div></div>
    <table class="doc-table"><thead><tr><th>No.</th><th>Article No.</th><th>Description</th><th>Quantity</th><th>Unit</th><th class="num">Unit Price</th><th class="num">Amount</th></tr><tr><th colspan="7">F.O.B. Value</th></tr></thead><tbody>${rows||'<tr><td colspan="7">No items</td></tr>'}</tbody></table>
    <div class="doc-footer"><div class="doc-totals"><div><span>Total Quantity :</span><strong>${t.qty}</strong></div><div><span>Sub Total:</span><strong>${formatCurrency(t.subtotal)}</strong></div>${t.discount?`<div><span>DISCOUNT AMOUNT</span><strong>(${formatCurrency(t.discount)})</strong></div>`:''}<div class="total"><span>Total : (${esc($('#currency').value)})</span><strong>${formatCurrency(t.total)}</strong></div></div><p><strong>Total Amount :</strong> ${esc($('#currency').value)} ${numberToWords(t.total)}</p><p><strong>Remark :</strong> ${esc($('#remark').value)}</p><br><div style="display:flex;justify-content:space-between"><span>Vender Signature : __________________</span><span>Accept By : __________________</span></div><p><strong>UNIVERSE GEMS &amp; JEWELLERY CO.</strong></p></div>`;
}
$('#refreshPreviewBtn').addEventListener('click',renderPreview); $('#printBtn').addEventListener('click',()=>{renderPreview();window.print()});

$('#scanBtn').addEventListener('click',startScanner); $('#closeScannerBtn').addEventListener('click',stopScanner);
async function startScanner(){
  const dlg=$('#scannerDialog'); dlg.showModal(); $('#scannerStatus').textContent='正在啟動相機…';
  try{
    if(!navigator.mediaDevices?.getUserMedia) throw new Error('此瀏覽器不支援相機存取');
    state.stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:'environment'}}});
    const video=$('#scannerVideo'); video.srcObject=state.stream; await video.play();
    if(!('BarcodeDetector' in window)) throw new Error('此 iOS / Safari 版本不支援網頁 BarcodeDetector，請使用手動 LOTNO');
    const detector=new BarcodeDetector({formats:['code_128','code_39','ean_13','ean_8','upc_a','upc_e','itf']});
    $('#scannerStatus').textContent='請將 Barcode 對準畫面中央。';
    const loop=async()=>{if(!dlg.open)return;try{const codes=await detector.detect(video);if(codes.length){const val=codes[0].rawValue;stopScanner();$('#lotInput').value=val;addByLot(val);return}}catch{}requestAnimationFrame(loop)};loop();
  }catch(err){$('#scannerStatus').textContent=err.message;}
}
function stopScanner(){state.stream?.getTracks().forEach(t=>t.stop());state.stream=null;$('#scannerDialog').close()}

window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();state.installPrompt=e;$('#installBtn').classList.remove('hidden')});
$('#installBtn').addEventListener('click',async()=>{if(state.installPrompt){state.installPrompt.prompt();await state.installPrompt.userChoice;state.installPrompt=null;$('#installBtn').classList.add('hidden')}});
if('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js').catch(console.error);
renderItems(); renderPreview(); updateHeader();
