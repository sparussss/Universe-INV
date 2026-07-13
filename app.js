const $=(s,root=document)=>root.querySelector(s);
const $$=(s,root=document)=>[...root.querySelectorAll(s)];

const state={
  products:new Map(), customers:new Map(), images:new Map(), imageFiles:[], items:[], soldLots:new Map(),
  stockRows:null, stockHeaderRow:-1, stockLotCol:-1, stockFileName:'jmsdata.xls',
  insertAt:null, scanner:null, scannerRunning:false, scannerPaused:false, scannerTransitioning:false, scannerCancelRequested:false,
  nextSequence:1, lastScan:{value:'',time:0}, installPrompt:null
};
const sample={lotNo:'133685',artNo:'PT-37499',price:3092,unit:'PC',desc2:'1-AMCTOCT25x9-10.75ct',descriptions:['2.10Y750','1-AMCTOCT25x9-10.75ct','6-CDMRD(B2)-0.05ct']};
const sampleCustomer={code:'JP1221',company:'ICHIMARU JEWELRY CO., LTD',address:'31-21, 2-CHOME YUSHIMA, BUNKYO-KU, TOKYO\n113-0034, JAPAN',rate:0.34,terms:'COD'};
state.products.set(sample.lotNo,sample); state.customers.set(sampleCustomer.code,sampleCustomer);

function normalizeKey(v){return String(v??'').trim()}
function normalizeArticle(v){return normalizeKey(v).toUpperCase()}
function normalizeCustomerCode(v){return String(v??'').replace(/\s+/g,'').trim().toUpperCase()}
function normalizeLot(v){return String(v??'').replace(/\s+/g,'').trim()}
function field(row,names){const keys=Object.keys(row);for(const name of names){const k=keys.find(x=>x.trim().toUpperCase()===name);if(k)return row[k]}return''}
function setStatus(id,text,type=''){const el=$(id);el.textContent=text;el.className='notice'+(type?' '+type:'')}
function formatCurrency(v){const c=$('#currency').value||'USD';return new Intl.NumberFormat('en-US',{style:'currency',currency:c,minimumFractionDigits:2}).format(v||0)}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function updateHeader(){const t=calcTotals();$('#productCount').textContent=state.products.size;$('#customerCount').textContent=state.customers.size;$('#invoiceCount').textContent=state.items.length;$('#headerTotal').textContent=formatCurrency(t.total)}
$('#invoiceDate').value=new Date().toISOString().slice(0,10);

$$('.tab').forEach(btn=>btn.addEventListener('click',()=>{
  $$('.tab').forEach(x=>x.classList.remove('active')); $$('.tab-panel').forEach(x=>x.classList.remove('active'));
  btn.classList.add('active'); $('#'+btn.dataset.tab).classList.add('active');
  if(btn.dataset.tab==='invoice')renderSelectedCustomer(); if(btn.dataset.tab==='preview')renderPreview();
}));

async function readWorkbook(file){
  if(typeof XLSX==='undefined')throw new Error('Excel 解析程式未載入。請連接網絡並重新開啟。');
  const data=await file.arrayBuffer(); return XLSX.read(data,{type:'array',cellDates:false});
}
function findHeader(rows,required){
  for(let i=0;i<Math.min(rows.length,30);i++){
    const values=(rows[i]||[]).map(v=>normalizeKey(v).toUpperCase().replace(/[ .]/g,''));
    if(required.every(r=>values.includes(r)))return i;
  }
  return -1;
}
$('#excelInput').addEventListener('change',async e=>{
  const file=e.target.files[0]; if(!file)return;
  try{
    const wb=await readWorkbook(file), ws=wb.Sheets[wb.SheetNames[0]];
    const aoa=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});
    const hi=findHeader(aoa,['LOTNO','ARTNO','PRICE']);
    if(hi<0)throw new Error('找不到 LOTNO / ARTNO / PRICE 標題列。');
    const headers=(aoa[hi]||[]).map(v=>normalizeKey(v));
    const lotCol=headers.findIndex(x=>x.toUpperCase().replace(/[ .]/g,'')==='LOTNO');
    const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
    const map=new Map(); let invalid=0;
    for(const row of rows){
      const lot=normalizeLot(field(row,['LOTNO','LOT NO','LOT.NO.'])); if(!lot)continue;
      const art=normalizeArticle(field(row,['ARTNO','ART NO','ARTICLE NO']));
      const price=Number(field(row,['PRICE','U PRICE','UPRICE']));
      if(!art||!Number.isFinite(price)){invalid++;continue}
      const rawDescriptions=[]; for(let i=1;i<=6;i++){rawDescriptions.push(normalizeKey(field(row,[`DESC${i}`,`DESCRIPTION${i}`])))}
      const descriptions=rawDescriptions.filter(Boolean), desc2=rawDescriptions[1]||'';
      map.set(lot,{lotNo:lot,artNo:art,price,unit:normalizeKey(field(row,['UNIT']))||'PC',desc2,descriptions});
    }
    if(!map.size)throw new Error('找不到有效貨品。');
    state.products=map; state.stockRows=aoa; state.stockHeaderRow=hi; state.stockLotCol=lotCol; state.stockFileName=file.name;
    state.soldLots.clear();
    setStatus('#excelStatus',`已匯入 ${file.name}：${map.size} 件可售貨品${invalid?`；略過 ${invalid} 行無效資料`:''}。`,'ok');
    rebuildImageIndex(); updateHeader(); persistLight();
  }catch(err){setStatus('#excelStatus',`匯入失敗：${err.message}`,'error')}
});

$('#customerExcelInput').addEventListener('change',async e=>{
  const file=e.target.files[0];if(!file)return;
  try{
    const wb=await readWorkbook(file),ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});
    const map=new Map();let started=false;
    for(const r of rows){
      const code=normalizeCustomerCode(r[0]), company=normalizeKey(r[1]);
      if(!started){if(code.includes('CUSTOMER')||company.toUpperCase().includes('COMPANY')){started=true;continue}if(!code||!company)continue;started=true}
      if(!code||!company)continue;
      const rawRate=r[11], parsed=Number(rawRate), rate=rawRate===''||rawRate==null||!Number.isFinite(parsed)?0.34:parsed;
      map.set(code,{code,company,address:[r[2],r[3],r[4]].map(normalizeKey).filter(Boolean).join('\n'),rate,terms:normalizeKey(r[10]),contact:normalizeKey(r[9]),phone:normalizeKey(r[6]),email:normalizeKey(r[8])});
    }
    if(!map.size)throw new Error('找不到有效 Customer Code / Company 資料。');
    state.customers=map;setStatus('#customerExcelStatus',`已匯入 ${file.name}：${map.size} 位客戶。空白 Sales Rate 已按 0.34 處理。`,'ok');updateHeader();
  }catch(err){setStatus('#customerExcelStatus',`匯入失敗：${err.message}`,'error')}
});

function searchCustomers(query){const q=normalizeKey(query).toUpperCase(),qc=normalizeCustomerCode(query);if(!q)return[];return [...state.customers.values()].filter(c=>c.code.includes(qc)||c.company.toUpperCase().includes(q)).slice(0,12)}
function renderCustomerMatches(){const box=$('#customerMatches'),matches=searchCustomers($('#customerSearch').value);box.innerHTML='';if(!matches.length){box.innerHTML='<div class="notice">找不到客戶。</div>';return}matches.forEach(c=>{const b=document.createElement('button');b.className='customer-match';b.innerHTML=`<span><strong>${esc(c.code)} · ${esc(c.company)}</strong><small>${esc(c.address).replace(/\n/g,' · ')}</small></span><span>Rate ${c.rate}</span>`;b.addEventListener('click',()=>selectCustomer(c));box.appendChild(b)})}
function selectCustomer(c){$('#customerCode').value=c.code;$('#customerName').value=c.company;$('#customerAddress').value=c.address;$('#salesRate').value=c.rate;$('#customerTerms').value=c.terms||'';$('#customerMatches').innerHTML='';$('#customerSearch').value='';repriceItems();renderSelectedCustomer();persistLight()}
$('#findCustomerBtn').addEventListener('click',renderCustomerMatches);$('#customerSearch').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();renderCustomerMatches()}});$('#customerSearch').addEventListener('input',()=>{$('#customerSearch').value.length>=2?renderCustomerMatches():$('#customerMatches').innerHTML=''})
function renderSelectedCustomer(){$('#selectedCustomerSummary').innerHTML=`<strong>${esc($('#customerCode').value)} · ${esc($('#customerName').value)}</strong><span>Sales Rate ${esc($('#salesRate').value)} · ${esc($('#currency').value)}</span>`}

function parseImageName(name){
  const base=name.replace(/\.[^.]+$/,'').trim();
  const clean=base.replace(/\s+\(\d+\)$/,'').trim();
  const firstSpace=clean.search(/\s/);
  const art=normalizeArticle(firstSpace<0?clean:clean.slice(0,firstSpace));
  const variant=firstSpace<0?'Default':clean.slice(firstSpace).trim()||'Default';
  return {art,variant};
}
$('#imageFolderInput').addEventListener('change',e=>{
  state.imageFiles=[...e.target.files].filter(f=>f.type.startsWith('image/')||/\.(jpe?g|png|webp|heic)$/i.test(f.name));
  rebuildImageIndex();
});
function rebuildImageIndex(){
  for(const imgs of state.images.values())for(const x of imgs)URL.revokeObjectURL(x.url);
  state.images=new Map(); let matched=0,ignored=0;
  const needed=new Set([...state.products.values()].map(p=>normalizeArticle(p.artNo)));
  for(const file of state.imageFiles){
    const {art,variant}=parseImageName(file.name); if(!needed.has(art)){ignored++;continue}
    const arr=state.images.get(art)||[]; const url=URL.createObjectURL(file);
    const existing=arr.findIndex(x=>x.variant.toUpperCase()===variant.toUpperCase());
    if(existing>=0){URL.revokeObjectURL(arr[existing].url);arr.splice(existing,1)}
    arr.push({variant,fileName:file.name,url}); arr.sort((a,b)=>a.variant==='Default'?-1:b.variant==='Default'?1:a.variant.localeCompare(b.variant));
    state.images.set(art,arr);matched++;
  }
  state.items.forEach(item=>{if(!item.imageManual)item.imageVariant=chooseImageVariant(item,getImages(item.artNo))});
  if(state.imageFiles.length)setStatus('#imageStatus',`已選擇 Folder：${state.imageFiles.length} 張圖片；其中 ${matched} 張符合目前倉存，${ignored} 張略過。圖片沒有複製進 PWA。`,'ok');
  renderItems();
}
function renderImageLibrary(){}
function getImages(art){return state.images.get(normalizeArticle(art))||[]}
function getSelectedImage(item){const imgs=getImages(item.artNo);return imgs.find(x=>x.variant===item.imageVariant)||imgs.find(x=>x.variant==='Default')||imgs[0]}
function placeholderSvg(text){return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="200" height="200"><rect width="100%" height="100%" fill="#eef2f6"/><text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" font-family="Arial" font-size="18" fill="#64748b">${text}</text></svg>`)}`}

const STONE_IMAGE_MAP={
  'SKY BT':'SKY BT','GPS':'G.AM','GAM':'G.AM','AMCT':'AMCT','ALEX':'ALEX',
  'QAM':'AM','LBT':'LBT','BTO':'BT','YCT':'CT','GPD':'PD','RQZ':'RQZ','MG':'MG',
  'PTQ':'PTR','AQ':'AQ','PAM':'P.AM','MCT':'MCT','RGT':'RGT','TZ':'TZ','IO':'IO',
  'GTQ':'GTR','GGT':'GGT','GSA':'GSA','WSA':'WSA','ZSA':'ZSA','BSA':'BSA',
  'PSA':'PSA','OSA':'OSA','YSA':'YSA','SSU':'SSU','RRU':'RRU','GEM':'GEM',
  'ZSP':'ZSP','DIA':'DIA','SQZ':'SQZ','LQZ':'LQZ','KU':'KU','GT':'GT',
  'AG':'AG','AMZ':'AMZ','BCH':'BCH','BO':'BO','GMA':'GMA','LAB':'LAB','LAP':'LAP',
  'MOON':'MOON','OPAL':'OPAL','WPL':'WPL','RCH':'RCH','TE':'TE','TQ':'TQ'
};
const STONE_BREAKDOWN_CODES=Object.keys(STONE_IMAGE_MAP).sort((a,b)=>b.length-a.length);
function normalizeStoneToken(v){return String(v||'').toUpperCase().replace(/[._\s-]+/g,'').replace(/[^A-Z0-9+]/g,'')}
function extractStoneTargets(desc2){
  const raw=String(desc2||'').toUpperCase();
  const core=(raw.split('-')[1]||raw).trim();
  const parts=core.split('+');
  const out=[];
  for(const partRaw of parts){
    const compact=partRaw.replace(/\s+/g,'');
    const code=STONE_BREAKDOWN_CODES.find(c=>compact.startsWith(c.replace(/\s+/g,'')));
    if(code){const mapped=STONE_IMAGE_MAP[code];if(mapped&&!out.includes(mapped))out.push(mapped)}
  }
  if(!out.length){
    const compact=raw.replace(/\s+/g,'');
    for(const code of STONE_BREAKDOWN_CODES){
      if(compact.includes(code.replace(/\s+/g,''))){const mapped=STONE_IMAGE_MAP[code];if(mapped&&!out.includes(mapped))out.push(mapped)}
    }
  }
  return out;
}
function variantComparable(v){
  return normalizeStoneToken(String(v||'')
    .replace(/\s+\(\d+\)$/,'')
    .replace(/\((18K[RYW]|14K[RYW]|9K[RYW]|10\d{2}|12\d{2}|REG|CK)\)/gi,''));
}
function chooseImageVariant(product,imgs){
  if(!imgs.length)return 'Default';
  const targets=extractStoneTargets(product.desc2||product.descriptions?.[1]||'');
  if(targets.length){
    const ordered=targets.map(normalizeStoneToken).join('+');
    const targetSet=new Set(targets.map(normalizeStoneToken));
    let best=null,bestScore=-1;
    for(const img of imgs){
      const comp=variantComparable(img.variant);
      const parts=comp.split('+').filter(Boolean);
      const partSet=new Set(parts);
      let score=0;
      if(comp===ordered)score=100;
      else if(parts.length===targetSet.size&&[...targetSet].every(x=>partSet.has(x)))score=90;
      else if([...targetSet].every(x=>comp.includes(x)))score=75;
      else if(comp===normalizeStoneToken(targets[0]))score=70;
      else if(comp.includes(normalizeStoneToken(targets[0])))score=55;
      if(score>bestScore){bestScore=score;best=img}
    }
    if(best&&bestScore>0)return best.variant;
  }
  return imgs.find(x=>x.variant==='Default')?.variant||imgs[0].variant;
}

$('#addLotBtn').addEventListener('click',()=>addByLot($('#lotInput').value));$('#lotInput').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();addByLot(e.target.value)}});
function addByLot(raw){
  const lot=normalizeLot(raw);if(!lot)return setStatus('#addMessage','請輸入 LOTNO。','error');
  if(state.soldLots.has(lot)){const s=state.soldLots.get(lot);return setStatus('#addMessage',`LOTNO ${lot} 已售出：${s.invoiceNo} / ${s.customerCode}。`,'error')}
  const product=state.products.get(lot);if(!product)return setStatus('#addMessage',`找不到 LOTNO ${lot}。請檢查 Remaining Stock。`,'error');
  if(state.items.some(x=>x.lotNo===lot))return setStatus('#addMessage',`LOTNO ${lot} 已經在這張 Invoice。`,'error');
  const rate=Number($('#salesRate').value)||0,unitPrice=Math.ceil(product.price*rate),imgs=getImages(product.artNo);
  const item={id:crypto.randomUUID?crypto.randomUUID():String(Date.now()+Math.random()),sequence:state.nextSequence++,...product,qty:1,unitPrice,imageVariant:chooseImageVariant(product,imgs)};
  state.insertAt===null?state.items.push(item):state.items.splice(state.insertAt,0,item);state.insertAt=null;$('#lotInput').value='';
  setStatus('#addMessage',`已加入 ${product.artNo} / LOTNO ${lot} / ${formatCurrency(unitPrice)}`,'ok');renderItems();persistLight();
}
function renderItems(){
  const box=$('#invoiceItems');box.innerHTML='';
  $('#invoiceListCount').textContent=`共 ${state.items.length} 件`;
  if(!state.items.length){box.className='invoice-items empty-state';box.textContent='尚未加入貨品。';$('#scrollNewestBtn').classList.add('hidden');updateTotals();return}
  box.className='invoice-items scrollable-items';
  $('#scrollNewestBtn').classList.toggle('hidden',state.items.length<=1);
  const view=[...state.items].map((item,sourceIndex)=>({item,sourceIndex})).reverse();
  view.forEach(({item,sourceIndex})=>{
    const node=$('#itemTemplate').content.firstElementChild.cloneNode(true);node.dataset.id=item.id;
    $('.item-seq',node).textContent=`#${item.sequence||sourceIndex+1}`;
    $('.item-artno',node).textContent=item.artNo;$('.item-lot',node).textContent=`LOTNO ${item.lotNo}`;$('.item-desc',node).textContent=item.descriptions.join('\n');
    $('.item-price-note',node).textContent=`${item.price}u × ${Number($('#salesRate').value||0)} → ${formatCurrency(item.unitPrice)}`;
    $('.item-thumb',node).src=getSelectedImage(item)?.url||placeholderSvg(item.artNo);
    const sel=$('.variant-select',node),imgs=getImages(item.artNo);
    if(imgs.length)imgs.forEach(x=>{const o=document.createElement('option');o.value=x.variant;o.textContent=x.variant;o.selected=x.variant===item.imageVariant;sel.appendChild(o)});else{const o=document.createElement('option');o.textContent='No image';sel.appendChild(o);sel.disabled=true}
    sel.addEventListener('change',()=>{item.imageVariant=sel.value;item.imageManual=true;renderItems();persistLight()});
    const qty=$('.qty-input',node),price=$('.price-input',node);qty.value=item.qty;price.value=item.unitPrice;
    qty.addEventListener('change',()=>{item.qty=Math.max(1,Number(qty.value)||1);updateTotals();persistLight()});price.addEventListener('change',()=>{item.unitPrice=Math.max(0,Math.ceil(Number(price.value)||0));updateTotals();persistLight()});
    $('.insert-above',node).addEventListener('click',()=>{state.insertAt=sourceIndex;$('#lotInput').focus();setStatus('#addMessage',`下一件會插入 #${item.sequence||sourceIndex+1} 之前。`,'ok')});
    $('.insert-below',node).addEventListener('click',()=>{state.insertAt=sourceIndex+1;$('#lotInput').focus();setStatus('#addMessage',`下一件會插入 #${item.sequence||sourceIndex+1} 之後。`,'ok')});
    $('.delete-item',node).addEventListener('click',()=>{if(confirm(`刪除 ${item.artNo} / LOTNO ${item.lotNo}？`)){state.items.splice(sourceIndex,1);renderItems();persistLight()}});
    node.addEventListener('dragstart',e=>e.dataTransfer.setData('text/plain',String(sourceIndex)));node.addEventListener('dragover',e=>e.preventDefault());node.addEventListener('drop',e=>{e.preventDefault();const from=Number(e.dataTransfer.getData('text/plain'));const [m]=state.items.splice(from,1);const target=sourceIndex>from?sourceIndex-1:sourceIndex;state.items.splice(target,0,m);renderItems();persistLight()});
    box.appendChild(node);
  });
  box.scrollTop=0;
  updateTotals();
}
$('#scrollNewestBtn').addEventListener('click',()=>{$('#invoiceItems').scrollTo({top:0,behavior:'smooth'})});
$('#clearInvoiceBtn').addEventListener('click',()=>{if(state.items.length&&confirm('清空目前 Invoice 草稿？')){state.items=[];state.nextSequence=1;renderItems();persistLight()}});
$('#salesRate').addEventListener('change',()=>{repriceItems();renderSelectedCustomer();persistLight()});$('#currency').addEventListener('change',()=>{renderItems();renderSelectedCustomer();persistLight()});$('#discountAmount').addEventListener('input',()=>{updateTotals();persistLight()});
function repriceItems(){const rate=Number($('#salesRate').value)||0;state.items.forEach(x=>x.unitPrice=Math.ceil(x.price*rate));renderItems()}
function calcTotals(){const qty=state.items.reduce((a,x)=>a+x.qty,0),subtotal=state.items.reduce((a,x)=>a+x.qty*x.unitPrice,0),discount=Math.max(0,Number($('#discountAmount').value)||0);return{qty,subtotal,discount,total:Math.max(0,subtotal-discount)}}
function updateTotals(){const t=calcTotals();$('#totalQty').textContent=t.qty;$('#subtotal').textContent=formatCurrency(t.subtotal);$('#discountDisplay').textContent=formatCurrency(t.discount);$('#grandTotal').textContent=formatCurrency(t.total);updateHeader()}
function numberToWords(n){n=Math.floor(n);if(n===0)return'ZERO';const ones=['','ONE','TWO','THREE','FOUR','FIVE','SIX','SEVEN','EIGHT','NINE','TEN','ELEVEN','TWELVE','THIRTEEN','FOURTEEN','FIFTEEN','SIXTEEN','SEVENTEEN','EIGHTEEN','NINETEEN'],tens=['','','TWENTY','THIRTY','FORTY','FIFTY','SIXTY','SEVENTY','EIGHTY','NINETY'];const under1000=x=>{let s='';if(x>=100){s+=ones[Math.floor(x/100)]+' HUNDRED ';x%=100}if(x>=20){s+=tens[Math.floor(x/10)]+' ';x%=10}if(x>0)s+=ones[x]+' ';return s.trim()};let out=[];for(const[v,name]of [[1e9,'BILLION'],[1e6,'MILLION'],[1e3,'THOUSAND'],[1,'']])if(n>=v){const part=Math.floor(n/v);n%=v;out.push(under1000(part)+(name?' '+name:''))}return out.join(' ')}
function renderPreview(){const t=calcTotals(),rows=state.items.map((x,i)=>`<tr><td>${i+1}</td><td><strong>Lot.No. : ${esc(x.lotNo)}</strong><br>${esc(x.artNo)}</td><td>${x.descriptions.map(esc).join('<br>')}</td><td class="num">${x.qty}</td><td>${esc(x.unit)}</td><td class="num">${formatCurrency(x.unitPrice)}</td><td class="num">${formatCurrency(x.qty*x.unitPrice)}</td></tr>`).join('');$('#invoiceDocument').innerHTML=`<div class="letterhead"><h2>UNIVERSE GEMS &amp; JEWELLERY CO.</h2><p>UNIT 11-12, 10/F., FU HANG INDUSTRIAL BUILDING, NO. 1 HOK YUEN STREET EAST,<br>HUNG HOM, KOWLOON, HONG KONG · TEL : (852) 2363 5409 · FAX : (852) 2765 0343</p></div><div class="doc-title">Sales Invoice</div><div class="doc-grid"><div class="doc-meta">No. : <strong>${esc($('#invoiceNo').value)}</strong><br>Invoice Date : ${esc($('#invoiceDate').value)}<br>Shipment Method : ${esc($('#shipmentMethod').value)}<br>Currency : ${esc($('#currency').value)}<br><br>Customer Code : ${esc($('#customerCode').value)}<br>Customer : <strong>${esc($('#customerName').value)}</strong><br>${esc($('#customerAddress').value).replace(/\n/g,'<br>')}</div><div class="doc-meta"><strong>Vender's Banker</strong><br>The Hong Kong &amp; Shanghai Banking Corporation Ltd.<br>Address : 41 Ma Tau Wai Road,Hung Hom,Kowloon,Hong Kong<br>A/C # : 012-593570-001<br>A/C Name : Universe Gems &amp; Jewellery Co.</div></div><table class="doc-table"><thead><tr><th>No.</th><th>Article No.</th><th>Description</th><th>Quantity</th><th>Unit</th><th class="num">Unit Price</th><th class="num">Amount</th></tr><tr><th colspan="7">F.O.B. Value</th></tr></thead><tbody>${rows||'<tr><td colspan="7">No items</td></tr>'}</tbody></table><div class="doc-footer"><div class="doc-totals"><div><span>Total Quantity :</span><strong>${t.qty}</strong></div><div><span>Sub Total:</span><strong>${formatCurrency(t.subtotal)}</strong></div>${t.discount?`<div><span>DISCOUNT AMOUNT</span><strong>(${formatCurrency(t.discount)})</strong></div>`:''}<div class="total"><span>Total : (${esc($('#currency').value)})</span><strong>${formatCurrency(t.total)}</strong></div></div><p><strong>Total Amount :</strong> ${esc($('#currency').value)} ${numberToWords(t.total)}</p><p><strong>Remark :</strong> ${esc($('#remark').value)}</p><br><div style="display:flex;justify-content:space-between"><span>Vender Signature : __________________</span><span>Accept By : __________________</span></div><p><strong>UNIVERSE GEMS &amp; JEWELLERY CO.</strong></p></div>`}
$('#refreshPreviewBtn').addEventListener('click',renderPreview);$('#printBtn').addEventListener('click',()=>{renderPreview();window.print()});

function sanitizeFileName(s){return String(s||'').replace(/[^a-zA-Z0-9._-]+/g,'_')}
function exportRemainingStock(invoiceNo){
  if(!state.stockRows||state.stockHeaderRow<0||state.stockLotCol<0)throw new Error('請先匯入真實倉存 Excel。');
  const output=[];
  for(let i=0;i<state.stockRows.length;i++){
    const row=state.stockRows[i]||[];
    if(i<=state.stockHeaderRow){output.push(row);continue}
    const lot=normalizeLot(row[state.stockLotCol]);
    if(!lot||state.products.has(lot))output.push(row);
  }
  const ws=XLSX.utils.aoa_to_sheet(output),wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Remaining Stock');
  const stamp=new Date().toISOString().replace(/[-:T]/g,'').slice(0,12);
  const fn=`Remaining_Stock_${sanitizeFileName(invoiceNo)}_${state.products.size}pcs_${stamp}.xlsx`;
  XLSX.writeFile(wb,fn,{compression:true});
  return fn;
}
$('#confirmInvoiceBtn').addEventListener('click',()=>{
  if(!state.items.length)return alert('Invoice 沒有貨品。');
  const invoiceNo=normalizeKey($('#invoiceNo').value)||'Invoice';
  if(!confirm(`Confirm ${invoiceNo}？\n\n完成後會把 ${state.items.length} 個 LOTNO 從目前倉存移除，並匯出新的 Remaining Stock Excel。`))return;
  const snapshot=[...state.items];
  for(const item of snapshot){state.products.delete(item.lotNo);state.soldLots.set(item.lotNo,{invoiceNo,customerCode:normalizeCustomerCode($('#customerCode').value),date:$('#invoiceDate').value})}
  try{
    const fileName=exportRemainingStock(invoiceNo);
    state.items=[];state.nextSequence=1;$('#discountAmount').value='0';renderItems();rebuildImageIndex();persistLight();renderPreview();
    alert(`Invoice 已 Confirm。\n已匯出：${fileName}\n目前剩餘 ${state.products.size} 件貨。`);
  }catch(err){
    for(const item of snapshot){state.products.set(item.lotNo,{lotNo:item.lotNo,artNo:item.artNo,price:item.price,unit:item.unit,descriptions:item.descriptions});state.soldLots.delete(item.lotNo)}
    alert(`未能匯出 Remaining Stock：${err.message}`);
  }
});

$('#scanBtn').addEventListener('click',startScanner);$('#closeScannerBtn').addEventListener('click',stopScanner);$('#pauseScannerBtn').addEventListener('click',toggleScannerPause);$('#torchBtn').addEventListener('click',toggleTorch);
for(const z of [1,2,3,4])$(`#zoom${z}Btn`).addEventListener('click',()=>setZoom(z));
function cleanScannedLot(value){return normalizeLot(value)}function vibrateSuccess(){try{navigator.vibrate?.([80,40,80])}catch{}}
function setScannerControlsDisabled(disabled){['#scanBtn','#pauseScannerBtn','#torchBtn','#zoom1Btn','#zoom2Btn','#zoom3Btn','#zoom4Btn'].forEach(x=>$(x).disabled=disabled);$('#closeScannerBtn').disabled=false}
async function startScanner(){
  if(state.scannerTransitioning||state.scannerRunning)return;
  state.scannerTransitioning=true;state.scannerCancelRequested=false;setScannerControlsDisabled(true);
  const dlg=$('#scannerDialog');if(!dlg.open)dlg.showModal();$('#scannerStatus').textContent='正在啟動後置相機…';
  if(typeof Html5Qrcode==='undefined'){state.scannerTransitioning=false;setScannerControlsDisabled(false);return $('#scannerStatus').textContent='掃描程式未載入。請連接網絡並重新開啟。'}
  try{
    if(state.scanner){try{await state.scanner.clear()}catch{}}
    $('#reader').innerHTML='';
    state.scanner=new Html5Qrcode('reader',{verbose:false});
    const formats=[Html5QrcodeSupportedFormats.CODE_128,Html5QrcodeSupportedFormats.CODE_39,Html5QrcodeSupportedFormats.EAN_13,Html5QrcodeSupportedFormats.EAN_8,Html5QrcodeSupportedFormats.UPC_A,Html5QrcodeSupportedFormats.UPC_E,Html5QrcodeSupportedFormats.ITF];
    await state.scanner.start({facingMode:{ideal:'environment'}},{fps:24,qrbox:(w,h)=>({width:Math.floor(w*.72),height:Math.max(54,Math.min(82,Math.floor(h*.12)))}),aspectRatio:1.7778,formatsToSupport:formats,experimentalFeatures:{useBarCodeDetectorIfSupported:true}},onBarcodeSuccess,()=>{});
    if(state.scannerCancelRequested){try{await state.scanner.stop();await state.scanner.clear()}catch{}state.scanner=null;state.scannerRunning=false;return}
    state.scannerRunning=true;state.scannerPaused=false;$('#scannerStatus').textContent='請把小條碼橫向填滿中央細長框。';
    setScannerControlsDisabled(false);$('#pauseScannerBtn').disabled=false;
    const caps=state.scanner.getRunningTrackCapabilities?.();
    if(!caps?.torch)$('#torchBtn').disabled=true;
    if(caps?.zoom){
      for(const z of [1,2,3,4])$(`#zoom${z}Btn`).disabled=z>(caps.zoom.max||1);
      await setZoom(Math.min(4,caps.zoom.max||1));
    }else for(const z of [1,2,3,4])$(`#zoom${z}Btn`).disabled=true;
    try{await state.scanner.applyVideoConstraints({advanced:[{focusMode:'continuous',exposureMode:'continuous'}]})}catch{}
  }catch(err){
    state.scannerRunning=false;setScannerControlsDisabled(false);$('#closeScannerBtn').disabled=false;
    $('#scannerStatus').textContent=`相機無法啟動：${err?.message||err}`;
  }finally{state.scannerTransitioning=false}
}
async function onBarcodeSuccess(decodedText){const value=cleanScannedLot(decodedText),now=Date.now();if(!value)return;if(state.lastScan.value===value&&now-state.lastScan.time<1800)return;state.lastScan={value,time:now};$('#lotInput').value=value;const before=state.items.length;addByLot(value);if(state.items.length>before){vibrateSuccess();$('#scannerStatus').textContent=`✓ 已加入 LOTNO ${value}，可繼續掃下一件。`}else $('#scannerStatus').textContent=$('#addMessage').textContent}
async function setZoom(value){if(!state.scannerRunning||state.scannerTransitioning)return;try{await state.scanner.applyVideoConstraints({advanced:[{zoom:value}]});for(const z of [1,2,3,4])$(`#zoom${z}Btn`).classList.toggle('active',z===value);$('#scannerStatus').textContent=`已切換 ${value}×。請等相機清晰對焦。`}catch{$('#scannerStatus').textContent=`此裝置不支援 ${value}× 網頁相機變焦。`}}
async function toggleScannerPause(){if(!state.scannerRunning||!state.scanner||state.scannerTransitioning)return;try{if(state.scannerPaused){state.scanner.resume();state.scannerPaused=false;$('#pauseScannerBtn').textContent='暫停'}else{state.scanner.pause(true);state.scannerPaused=true;$('#pauseScannerBtn').textContent='繼續'}}catch{}}
async function toggleTorch(){if(state.scannerTransitioning)return;try{const current=$('#torchBtn').dataset.on==='1';await state.scanner.applyVideoConstraints({advanced:[{torch:!current}]});$('#torchBtn').dataset.on=current?'0':'1';$('#torchBtn').textContent=current?'🔦 手電筒':'🔦 關燈'}catch{$('#scannerStatus').textContent='此裝置不支援網頁控制手電筒。'}}
async function stopScanner(){
  state.scannerCancelRequested=true;
  if($('#scannerDialog').open)$('#scannerDialog').close();
  if(state.scannerTransitioning)return;
  state.scannerTransitioning=true;setScannerControlsDisabled(true);
  try{if(state.scannerRunning&&state.scanner)await state.scanner.stop();if(state.scanner)await state.scanner.clear()}catch{}
  state.scanner=null;state.scannerRunning=false;state.scannerPaused=false;$('#reader').innerHTML='';
  state.scannerTransitioning=false;setScannerControlsDisabled(false);
}
$('#photoScanBtn').addEventListener('click',()=>$('#photoScanInput').click());
$('#photoScanInput').addEventListener('change',async e=>{
  const file=e.target.files[0];if(!file)return;
  if(typeof Html5Qrcode==='undefined')return setStatus('#addMessage','掃描程式未載入。','error');
  let scanner;
  try{
    $('#photoReader').innerHTML='';
    scanner=new Html5Qrcode('photoReader',{verbose:false});
    const result=await scanner.scanFile(file,true);
    const lot=cleanScannedLot(result);$('#lotInput').value=lot;addByLot(lot);
  }catch{setStatus('#addMessage','照片內未能解讀 Barcode。可先在相片 App 裁切／放大條碼後再選取。','error')}
  finally{e.target.value='';try{await scanner?.clear()}catch{}$('#photoReader').innerHTML=''}
});

function spokenToDigits(text){
  const direct=String(text||'').replace(/\D/g,'');if(direct)return direct;
  const map={'零':'0','〇':'0','洞':'0','一':'1','壹':'1','幺':'1','二':'2','兩':'2','贰':'2','三':'3','參':'3','四':'4','肆':'4','五':'5','伍':'5','六':'6','陸':'6','七':'7','柒':'7','八':'8','捌':'8','九':'9','玖':'9'};
  return [...String(text||'')].map(c=>map[c]||'').join('');
}
$('#voiceLotBtn').addEventListener('click',()=>{
  const SR=window.SpeechRecognition||window.webkitSpeechRecognition;
  if(!SR){
    $('#lotInput').focus();
    setStatus('#addMessage','此 iPhone 網頁未提供直接語音辨識。已開啟文字鍵盤，請按鍵盤上的咪高峰說出 LOTNO。','ok');
    return;
  }
  const rec=new SR();rec.lang='zh-HK';rec.interimResults=false;rec.maxAlternatives=1;
  $('#voiceLotBtn').disabled=true;setStatus('#addMessage','正在聆聽 LOTNO…請逐個數字讀出。','ok');
  rec.onresult=e=>{const raw=e.results?.[0]?.[0]?.transcript||'';const lot=spokenToDigits(raw);$('#lotInput').value=lot;if(lot)setStatus('#addMessage',`已辨識 LOTNO ${lot}，請按「加入」確認。`,'ok');else setStatus('#addMessage',`未能從「${raw}」取得數字，請再試。`,'error')};
  rec.onerror=()=>setStatus('#addMessage','語音輸入失敗，請使用鍵盤咪高峰或手動輸入。','error');
  rec.onend=()=>{$('#voiceLotBtn').disabled=false};
  try{rec.start()}catch{$('#voiceLotBtn').disabled=false}
});

function persistLight(){try{localStorage.setItem('ui-v05',JSON.stringify({customerCode:$('#customerCode').value,customerName:$('#customerName').value,customerAddress:$('#customerAddress').value,salesRate:$('#salesRate').value,currency:$('#currency').value,shipmentMethod:$('#shipmentMethod').value,customerTerms:$('#customerTerms').value,invoiceNo:$('#invoiceNo').value,invoiceDate:$('#invoiceDate').value,discount:$('#discountAmount').value,remark:$('#remark').value,items:state.items,soldLots:[...state.soldLots.entries()]}))}catch{}}
function restoreLight(){try{const d=JSON.parse(localStorage.getItem('ui-v05')||'null');if(!d)return;for(const[id,key]of[['#customerCode','customerCode'],['#customerName','customerName'],['#customerAddress','customerAddress'],['#salesRate','salesRate'],['#currency','currency'],['#shipmentMethod','shipmentMethod'],['#customerTerms','customerTerms'],['#invoiceNo','invoiceNo'],['#invoiceDate','invoiceDate'],['#discountAmount','discount'],['#remark','remark']])if(d[key]!==undefined)$(id).value=d[key];if(Array.isArray(d.items)){state.items=d.items;let max=0;state.items.forEach((x,i)=>{if(!x.sequence)x.sequence=i+1;max=Math.max(max,x.sequence||0)});state.nextSequence=max+1}if(Array.isArray(d.soldLots))state.soldLots=new Map(d.soldLots)}catch{}}
['#customerCode','#customerName','#customerAddress','#shipmentMethod','#customerTerms','#invoiceNo','#invoiceDate','#remark'].forEach(id=>$(id).addEventListener('change',persistLight));
window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();state.installPrompt=e;$('#installBtn').classList.remove('hidden')});$('#installBtn').addEventListener('click',async()=>{if(state.installPrompt){state.installPrompt.prompt();state.installPrompt=null;$('#installBtn').classList.add('hidden')}});if('serviceWorker'in navigator)window.addEventListener('load',()=>navigator.serviceWorker.register('./sw.js').catch(()=>{}));
restoreLight();renderSelectedCustomer();renderItems();updateHeader();
