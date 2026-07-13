const $=(s,root=document)=>root.querySelector(s);
const $$=(s,root=document)=>[...root.querySelectorAll(s)];

const state={
  products:new Map(),customers:new Map(),images:new Map(),imageFiles:[],items:[],soldLots:new Map(),
  stockRows:null,stockHeaderRow:-1,stockLotCol:-1,stockFileName:'jmsdata.xls',
  insertAt:null,scanner:null,scannerRunning:false,scannerPaused:false,scannerTransitioning:false,scannerCancelRequested:false,
  nextSequence:1,lastScan:{value:'',time:0},feedbackTimer:null,installPrompt:null
};

function normalizeKey(v){return String(v??'').trim()}
function normalizeArticle(v){return normalizeKey(v).toUpperCase()}
function normalizeCustomerCode(v){return String(v??'').replace(/\s+/g,'').trim().toUpperCase()}
function normalizeLot(v){return String(v??'').replace(/\D+/g,'').trim()}
function field(row,names){const keys=Object.keys(row);for(const name of names){const k=keys.find(x=>x.trim().toUpperCase()===name);if(k)return row[k]}return''}
function setStatus(id,text,type=''){const el=$(id);el.textContent=text;el.className='notice'+(type?' '+type:'');el.classList.remove('hidden')}
function formatCurrency(v){const c=$('#currency').value||'USD';try{return new Intl.NumberFormat('en-US',{style:'currency',currency:c,minimumFractionDigits:2}).format(v||0)}catch{return `$${Number(v||0).toFixed(2)}`}}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function sanitizeFileName(s){return String(s||'').replace(/[^a-zA-Z0-9._-]+/g,'_')}
function downloadBlob(blob,fileName){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=fileName;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000)}
function updateHeader(){const t=calcTotals();$('#productCount').textContent=state.products.size;$('#customerCount').textContent=state.customers.size;$('#invoiceCount').textContent=state.items.length;$('#headerTotal').textContent=formatCurrency(t.total);$('#scannerItemCount').textContent=state.items.length}
$('#invoiceDate').value=new Date().toISOString().slice(0,10);

$$('.tab').forEach(btn=>btn.addEventListener('click',()=>{
  $$('.tab').forEach(x=>x.classList.remove('active'));$$('.tab-panel').forEach(x=>x.classList.remove('active'));
  btn.classList.add('active');$('#'+btn.dataset.tab).classList.add('active');
  if(btn.dataset.tab==='invoice')renderSelectedCustomer();if(btn.dataset.tab==='preview')renderPreview();
}));

async function readWorkbook(file){
  if(typeof XLSX==='undefined')throw new Error('Excel 解析程式未載入。請連接網絡後重新開啟。');
  return XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false});
}
function findHeader(rows,required){for(let i=0;i<Math.min(rows.length,30);i++){const values=(rows[i]||[]).map(v=>normalizeKey(v).toUpperCase().replace(/[ .]/g,''));if(required.every(r=>values.includes(r)))return i}return-1}

$('#excelInput').addEventListener('change',async e=>{
  const file=e.target.files[0];if(!file)return;
  try{
    const wb=await readWorkbook(file),ws=wb.Sheets[wb.SheetNames[0]],aoa=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});
    const hi=findHeader(aoa,['LOTNO','ARTNO','PRICE']);if(hi<0)throw new Error('找不到 LOTNO / ARTNO / PRICE 標題列。');
    const headers=(aoa[hi]||[]).map(v=>normalizeKey(v)),lotCol=headers.findIndex(x=>x.toUpperCase().replace(/[ .]/g,'')==='LOTNO');
    const rows=XLSX.utils.sheet_to_json(ws,{defval:''}),map=new Map();let invalid=0;
    for(const row of rows){
      const lot=String(field(row,['LOTNO','LOT NO','LOT.NO.'])??'').replace(/\.0$/,'').replace(/\s+/g,'').trim();if(!lot)continue;
      const art=normalizeArticle(field(row,['ARTNO','ART NO','ARTICLE NO'])),price=Number(field(row,['PRICE','U PRICE','UPRICE']));
      if(!art||!Number.isFinite(price)){invalid++;continue}
      const raw=[];for(let i=1;i<=6;i++)raw.push(normalizeKey(field(row,[`DESC${i}`,`DESCRIPTION${i}`])));
      map.set(lot,{lotNo:lot,artNo:art,price,unit:normalizeKey(field(row,['UNIT']))||'PC',desc2:raw[1]||'',descriptions:raw.filter(Boolean)});
    }
    if(!map.size)throw new Error('找不到有效貨品。');
    state.products=map;state.stockRows=aoa;state.stockHeaderRow=hi;state.stockLotCol=lotCol;state.stockFileName=file.name;state.soldLots.clear();
    setStatus('#excelStatus',`已匯入 ${file.name}：${map.size} 件貨品${invalid?`；${invalid} 行略過`:''}。`,'ok');rebuildImageIndex();updateHeader();persistLight();
  }catch(err){setStatus('#excelStatus',`匯入失敗：${err.message}`,'error')}
});

$('#customerExcelInput').addEventListener('change',async e=>{
  const file=e.target.files[0];if(!file)return;
  try{
    const wb=await readWorkbook(file),ws=wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true}),map=new Map();let started=false;
    for(const r of rows){
      const code=normalizeCustomerCode(r[0]),company=normalizeKey(r[1]);
      if(!started){if(code.includes('CUSTOMER')||company.toUpperCase().includes('COMPANY')){started=true;continue}if(!code||!company)continue;started=true}
      if(!code||!company)continue;
      const parsed=Number(r[11]),rate=r[11]===''||r[11]==null||!Number.isFinite(parsed)?0.34:parsed;
      map.set(code,{code,company,address:[r[2],r[3],r[4]].map(normalizeKey).filter(Boolean).join('\n'),rate,terms:normalizeKey(r[10]),contact:normalizeKey(r[9]),phone:normalizeKey(r[6]),email:normalizeKey(r[8])});
    }
    if(!map.size)throw new Error('找不到有效 Customer Code / Company 資料。');
    state.customers=map;setStatus('#customerExcelStatus',`已匯入 ${file.name}：${map.size} 位客戶。`,'ok');updateHeader();
  }catch(err){setStatus('#customerExcelStatus',`匯入失敗：${err.message}`,'error')}
});

function searchCustomers(query){const q=normalizeKey(query).toUpperCase(),qc=normalizeCustomerCode(query);if(!q)return[];return[...state.customers.values()].filter(c=>c.code.includes(qc)||c.company.toUpperCase().includes(q)).slice(0,12)}
function renderCustomerMatches(){const box=$('#customerMatches'),matches=searchCustomers($('#customerSearch').value);box.innerHTML='';if(!matches.length){box.innerHTML='<div class="notice">找不到客戶。</div>';return}matches.forEach(c=>{const b=document.createElement('button');b.className='customer-match';b.innerHTML=`<span><strong>${esc(c.code)} · ${esc(c.company)}</strong><small>${esc(c.address).replace(/\n/g,' · ')}</small></span><span>Rate ${c.rate}</span>`;b.addEventListener('click',()=>selectCustomer(c));box.appendChild(b)})}
function selectCustomer(c){$('#customerCode').value=c.code;$('#customerName').value=c.company;$('#customerAddress').value=c.address;$('#salesRate').value=c.rate;$('#currency').value=$('#currency').value||'USD';$('#customerTerms').value=c.terms||'';$('#customerMatches').innerHTML='';$('#customerSearch').value='';repriceItems();renderSelectedCustomer();persistLight()}
$('#findCustomerBtn').addEventListener('click',renderCustomerMatches);$('#customerSearch').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();renderCustomerMatches()}});$('#customerSearch').addEventListener('input',()=>{$('#customerSearch').value.length>=2?renderCustomerMatches():$('#customerMatches').innerHTML=''})
function renderSelectedCustomer(){const code=normalizeCustomerCode($('#customerCode').value),name=normalizeKey($('#customerName').value);if(!code&&!name){$('#selectedCustomerSummary').textContent='尚未選擇客戶。';return}$('#selectedCustomerSummary').innerHTML=`<strong>${esc(code)}${code&&name?' · ':''}${esc(name)}</strong><span>Sales Rate ${esc($('#salesRate').value)} · ${esc($('#currency').value||'')}</span>`}

function parseImageName(name){
  const base=name.replace(/\.[^.]+$/,'').trim(),m=base.match(/\s+\((\d+)\)$/),duplicateIndex=m?Number(m[1]):0,clean=m?base.slice(0,m.index).trim():base;
  const firstSpace=clean.search(/\s/),art=normalizeArticle(firstSpace<0?clean:clean.slice(0,firstSpace)),variant=firstSpace<0?'Default':clean.slice(firstSpace).trim()||'Default';
  return{art,variant,variantKey:variant.toUpperCase(),duplicateIndex};
}
$('#imageFolderInput').addEventListener('change',e=>{state.imageFiles=[...e.target.files].filter(f=>f.type.startsWith('image/')||/\.(jpe?g|png|webp|heic)$/i.test(f.name));rebuildImageIndex()});
function rebuildImageIndex(){
  for(const imgs of state.images.values())for(const x of imgs)URL.revokeObjectURL(x.url);state.images=new Map();let matched=0,ignored=0;
  const needed=new Set([...state.products.values()].map(p=>normalizeArticle(p.artNo)));
  for(const file of state.imageFiles){
    const p=parseImageName(file.name);if(!needed.has(p.art)){ignored++;continue}matched++;
    const arr=state.images.get(p.art)||[],existing=arr.findIndex(x=>x.variantKey===p.variantKey);
    if(existing>=0){if(p.duplicateIndex>=arr[existing].duplicateIndex)continue;URL.revokeObjectURL(arr[existing].url);arr.splice(existing,1)}
    arr.push({...p,fileName:file.name,file,url:URL.createObjectURL(file)});
    arr.sort((a,b)=>a.variant==='Default'?-1:b.variant==='Default'?1:a.variant.localeCompare(b.variant,undefined,{numeric:true,sensitivity:'base'}));state.images.set(p.art,arr);
  }
  state.items.forEach(item=>{if(!item.imageManual)item.imageVariant=chooseImageVariant(item,getImages(item.artNo))});
  if(state.imageFiles.length)setStatus('#imageStatus',`已選擇 Folder：${state.imageFiles.length} 張圖片；其中 ${matched} 張符合目前倉存，${ignored} 張略過。`,'ok');renderItems();
}
function getImages(art){return state.images.get(normalizeArticle(art))||[]}
function getSelectedImage(item){const imgs=getImages(item.artNo),wanted=String(item.imageVariant||'').trim().toUpperCase();return imgs.find(x=>x.variantKey===wanted)||imgs.find(x=>x.variant==='Default')||imgs[0]}
function placeholderSvg(text){return`data:image/svg+xml;charset=UTF-8,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="200" height="200"><rect width="100%" height="100%" fill="#eef2f6"/><text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" font-family="Arial" font-size="18" fill="#64748b">${text}</text></svg>`)}`}

const STONE_IMAGE_MAP={'SKY BT':'SKY BT','GPS':'GAM','GAM':'GAM','AMCT':'AMCT','QAM':'AM','LBT':'L.BT','BTO':'BT','YCT':'CT','GPD':'PD','RQZ':'RQZ','MG':'MG','PTQ':'PTR','AQ':'AQ','PAM':'PAM','MCT':'MCT','RGT':'RGT','TZ':'TZ','IO':'IO','GTQ':'GTR','GGT':'GGT','GSA':'GSA','WSA':'WSA','ZSA':'ZSA','BSA':'BSA','PSA':'PSA','OSA':'OSA','YSA':'YSA','SQZ':'SQZ','OPAL':'OPAL','BO':'BO'};
const STONE_CODES=Object.keys(STONE_IMAGE_MAP).sort((a,b)=>b.length-a.length);
function normStone(v){return String(v||'').toUpperCase().replace(/[._\s-]+/g,'').replace(/[^A-Z0-9+]/g,'')}
function extractStoneTargets(desc2){const raw=String(desc2||'').toUpperCase(),core=(raw.split('-')[1]||raw).trim(),out=[];for(const part of core.split('+')){const compact=part.replace(/\s+/g,''),code=STONE_CODES.find(c=>compact.startsWith(c.replace(/\s+/g,'')));if(code){const mapped=STONE_IMAGE_MAP[code];if(mapped&&!out.includes(mapped))out.push(mapped)}}return out}
function comparableVariant(v){return normStone(String(v||'').replace(/\s+\(\d+\)$/,'').replace(/\((18K[RYW]|14K[RYW]|9K[RYW]|10\d{2}|12\d{2}|REG|CK)\)/gi,''))}
function chooseImageVariant(product,imgs){
  if(!imgs.length)return'Default';const targets=extractStoneTargets(product.desc2||product.descriptions?.[1]||'');
  if(targets.length){const ordered=targets.map(normStone).join('+'),set=new Set(targets.map(normStone));let best=null,score=-1;for(const img of imgs){const comp=comparableVariant(img.variant),parts=comp.split('+').filter(Boolean),pSet=new Set(parts);let s=0;if(comp===ordered)s=100;else if(parts.length===set.size&&[...set].every(x=>pSet.has(x)))s=90;else if([...set].every(x=>comp.includes(x)))s=75;else if(comp===normStone(targets[0]))s=70;else if(comp.includes(normStone(targets[0])))s=50;if(s>score){score=s;best=img}}if(best&&score>0)return best.variant}
  return imgs.find(x=>x.variant==='Default')?.variant||imgs[0].variant;
}

function showAddMessage(text,type='ok'){setStatus('#addMessage',text,type);setTimeout(()=>$('#addMessage').classList.add('hidden'),2200)}
function addByLot(raw,source='manual'){
  const lot=normalizeLot(raw);if(!lot)return{status:'notfound',lot:'',message:'請輸入 LOTNO。'};
  if(state.soldLots.has(lot)){const s=state.soldLots.get(lot);return{status:'duplicate',lot,message:`已售出 ${s.invoiceNo}`}}
  const product=state.products.get(lot);if(!product)return{status:'notfound',lot,message:'找不到'};
  if(state.items.some(x=>x.lotNo===lot))return{status:'duplicate',lot,message:'已加入'};
  const rate=Number($('#salesRate').value);if(!Number.isFinite(rate))return{status:'notfound',lot,message:'請先選擇客戶'};
  const item={id:crypto.randomUUID?crypto.randomUUID():String(Date.now()+Math.random()),sequence:state.nextSequence++,...product,qty:1,unitPrice:Math.ceil(product.price*rate),imageVariant:chooseImageVariant(product,getImages(product.artNo))};
  state.insertAt===null?state.items.push(item):state.items.splice(state.insertAt,0,item);state.insertAt=null;$('#lotInput').value='';renderItems();persistLight();
  if(source!=='scanner')showAddMessage(`已加入 ${product.artNo} / LOTNO ${lot} / ${formatCurrency(item.unitPrice)}`,'ok');
  return{status:'success',lot,message:product.artNo};
}
$('#addLotBtn').addEventListener('click',()=>{const r=addByLot($('#lotInput').value);if(r.status!=='success')showAddMessage(`${r.lot?`LOTNO ${r.lot} `:''}${r.message}`,'error')});
$('#lotInput').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();$('#addLotBtn').click()}});$('#manualModeBtn').addEventListener('click',()=>{$('#lotInput').focus()});

function renderItems(){
  const box=$('#invoiceItems');box.innerHTML='';$('#invoiceListCount').textContent=`共 ${state.items.length} 件`;
  if(!state.items.length){box.className='invoice-items empty-state';box.textContent='尚未加入貨品。';$('#scrollNewestBtn').classList.add('hidden');updateTotals();return}
  box.className='invoice-items scrollable-items';$('#scrollNewestBtn').classList.toggle('hidden',state.items.length<=1);
  [...state.items].map((item,sourceIndex)=>({item,sourceIndex})).reverse().forEach(({item,sourceIndex})=>{
    const node=$('#itemTemplate').content.firstElementChild.cloneNode(true);node.dataset.id=item.id;$('.item-seq',node).textContent=`#${item.sequence||sourceIndex+1}`;$('.item-artno',node).textContent=item.artNo;$('.item-lot',node).textContent=`LOTNO ${item.lotNo}`;$('.item-desc',node).textContent=item.descriptions.join('\n');$('.item-thumb',node).src=getSelectedImage(item)?.url||placeholderSvg(item.artNo);
    const sel=$('.variant-select',node),imgs=getImages(item.artNo);if(imgs.length)imgs.forEach(x=>{const o=document.createElement('option');o.value=x.variant;o.textContent=x.variant;o.selected=x.variant===item.imageVariant;sel.appendChild(o)});else{const o=document.createElement('option');o.textContent='No image';sel.appendChild(o);sel.disabled=true}
    sel.addEventListener('change',()=>{item.imageVariant=sel.value;item.imageManual=true;renderItems();persistLight()});
    const qty=$('.qty-input',node),price=$('.price-input',node);qty.value=item.qty;price.value=item.unitPrice;qty.addEventListener('change',()=>{item.qty=Math.max(1,Number(qty.value)||1);updateTotals();persistLight()});price.addEventListener('change',()=>{item.unitPrice=Math.max(0,Math.ceil(Number(price.value)||0));updateTotals();persistLight()});
    $('.insert-above',node).addEventListener('click',()=>{state.insertAt=sourceIndex;$('#lotInput').focus();showAddMessage(`下一件會插入 #${item.sequence} 之前。`)});$('.insert-below',node).addEventListener('click',()=>{state.insertAt=sourceIndex+1;$('#lotInput').focus();showAddMessage(`下一件會插入 #${item.sequence} 之後。`)});$('.delete-item',node).addEventListener('click',()=>{if(confirm(`刪除 ${item.artNo} / LOTNO ${item.lotNo}？`)){state.items.splice(sourceIndex,1);renderItems();persistLight()}});
    node.addEventListener('dragstart',e=>e.dataTransfer.setData('text/plain',String(sourceIndex)));node.addEventListener('dragover',e=>e.preventDefault());node.addEventListener('drop',e=>{e.preventDefault();const from=Number(e.dataTransfer.getData('text/plain'));const[m]=state.items.splice(from,1),target=sourceIndex>from?sourceIndex-1:sourceIndex;state.items.splice(target,0,m);renderItems();persistLight()});box.appendChild(node);
  });
  box.scrollTop=0;updateTotals();
}
$('#scrollNewestBtn').addEventListener('click',()=>$('#invoiceItems').scrollTo({top:0,behavior:'smooth'}));$('#clearInvoiceBtn').addEventListener('click',()=>{if(state.items.length&&confirm('清空目前 Invoice 草稿？')){state.items=[];state.nextSequence=1;renderItems();persistLight()}});
$('#salesRate').addEventListener('change',()=>{repriceItems();renderSelectedCustomer();persistLight()});$('#currency').addEventListener('change',()=>{renderItems();renderSelectedCustomer();persistLight()});$('#discountAmount').addEventListener('input',()=>{updateTotals();persistLight()});
function repriceItems(){const rate=Number($('#salesRate').value);if(Number.isFinite(rate))state.items.forEach(x=>x.unitPrice=Math.ceil(x.price*rate));renderItems()}
function calcTotals(){const qty=state.items.reduce((a,x)=>a+x.qty,0),subtotal=state.items.reduce((a,x)=>a+x.qty*x.unitPrice,0),discount=Math.max(0,Number($('#discountAmount').value)||0);return{qty,subtotal,discount,total:Math.max(0,subtotal-discount)}}
function updateTotals(){const t=calcTotals();$('#totalQty').textContent=t.qty;$('#subtotal').textContent=formatCurrency(t.subtotal);$('#discountDisplay').textContent=formatCurrency(t.discount);$('#grandTotal').textContent=formatCurrency(t.total);updateHeader()}
function numberToWords(n){n=Math.floor(n);if(n===0)return'ZERO';const ones=['','ONE','TWO','THREE','FOUR','FIVE','SIX','SEVEN','EIGHT','NINE','TEN','ELEVEN','TWELVE','THIRTEEN','FOURTEEN','FIFTEEN','SIXTEEN','SEVENTEEN','EIGHTEEN','NINETEEN'],tens=['','','TWENTY','THIRTY','FORTY','FIFTY','SIXTY','SEVENTY','EIGHTY','NINETY'];const under=x=>{let s='';if(x>=100){s+=ones[Math.floor(x/100)]+' HUNDRED ';x%=100}if(x>=20){s+=tens[Math.floor(x/10)]+' ';x%=10}if(x>0)s+=ones[x]+' ';return s.trim()};const out=[];for(const[v,name]of[[1e9,'BILLION'],[1e6,'MILLION'],[1e3,'THOUSAND'],[1,'']])if(n>=v){const p=Math.floor(n/v);n%=v;out.push(under(p)+(name?' '+name:''))}return out.join(' ')}

function renderPreview(){
  const t=calcTotals(),rows=state.items.map((x,i)=>`<tr><td>${i+1}</td><td><strong>Lot.No. : ${esc(x.lotNo)}</strong><br>${esc(x.artNo)}</td><td>${x.descriptions.map(esc).join('<br>')}</td><td class="num">${x.qty}</td><td>${esc(x.unit)}</td><td class="num">${formatCurrency(x.unitPrice)}</td><td class="num">${formatCurrency(x.qty*x.unitPrice)}</td></tr>`).join('');
  $('#invoiceDocument').innerHTML=`<div class="letterhead"><h2>UNIVERSE GEMS &amp; JEWELLERY CO.</h2><p>UNIT 11-12, 10/F., FU HANG INDUSTRIAL BUILDING, NO. 1 HOK YUEN STREET EAST,<br>HUNG HOM, KOWLOON, HONG KONG · TEL : (852) 2363 5409 · FAX : (852) 2765 0343</p></div><div class="doc-title">Sales Invoice</div><div class="doc-grid"><div class="doc-meta">No. : <strong>${esc($('#invoiceNo').value)}</strong><br>Invoice Date : ${esc($('#invoiceDate').value)}<br>Shipment Method : ${esc($('#shipmentMethod').value)}<br>Currency : ${esc($('#currency').value)}<br><br>Customer Code : ${esc($('#customerCode').value)}<br>Customer : <strong>${esc($('#customerName').value)}</strong><br>${esc($('#customerAddress').value).replace(/\n/g,'<br>')}</div><div class="doc-meta"><strong>Vender's Banker</strong><br>The Hong Kong &amp; Shanghai Banking Corporation Ltd.<br>Address : 41 Ma Tau Wai Road,Hung Hom,Kowloon,Hong Kong<br>A/C # : 012-593570-001<br>A/C Name : Universe Gems &amp; Jewellery Co.</div></div><table class="doc-table"><thead><tr><th>No.</th><th>Article No.</th><th>Description</th><th>Quantity</th><th>Unit</th><th>Unit Price</th><th>Amount</th></tr></thead><tbody>${rows||'<tr><td colspan="7">No items</td></tr>'}</tbody></table><div class="doc-footer"><div class="doc-totals"><div><span>Total Quantity :</span><strong>${t.qty}</strong></div><div><span>Sub Total:</span><strong>${formatCurrency(t.subtotal)}</strong></div><div><span>Discount:</span><strong>${formatCurrency(t.discount)}</strong></div><div class="total"><span>Total : (${esc($('#currency').value)})</span><strong>${formatCurrency(t.total)}</strong></div></div><p><strong>Total Amount :</strong> ${esc($('#currency').value)} ${numberToWords(t.total)}</p><p><strong>Remark :</strong> ${esc($('#remark').value)}</p></div>`;
}
$('#refreshPreviewBtn').addEventListener('click',renderPreview);

async function fileToBase64(file){const bytes=new Uint8Array(await file.arrayBuffer());let bin='';const chunk=0x8000;for(let i=0;i<bytes.length;i+=chunk)bin+=String.fromCharCode(...bytes.subarray(i,i+chunk));return btoa(bin)}
function imageExtension(file){const m=file.name.match(/\.(jpe?g|png)$/i);return m?(m[1].toLowerCase()==='jpg'?'jpeg':m[1].toLowerCase()):'jpeg'}
function setCellBorder(cell){cell.border={top:{style:'thin',color:{argb:'FFB8BEC8'}},bottom:{style:'thin',color:{argb:'FFB8BEC8'}}}}
async function buildInvoiceWorkbook(){
  if(typeof ExcelJS==='undefined')throw new Error('Excel 輸出程式未載入。');if(!state.items.length)throw new Error('Invoice 沒有貨品。');
  const wb=new ExcelJS.Workbook();wb.creator='Universe Invoice PWA';wb.created=new Date();const ws=wb.addWorksheet('Sales Invoice',{pageSetup:{paperSize:9,orientation:'portrait',fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.25,right:.25,top:.3,bottom:.3,header:.1,footer:.1}}});ws.views=[{showGridLines:false}];
  const widths=[5,16,16,16,12,8,8,13,13];['A','B','C','D','E','F','G','H','I'].forEach((c,i)=>ws.getColumn(c).width=widths[i]);
  ws.mergeCells('A1:I2');const title=ws.getCell('A1');title.value='UNIVERSE GEMS & JEWELLERY CO.';title.font={name:'Arial',size:23,bold:true,color:{argb:'FF17365D'}};title.alignment={horizontal:'center',vertical:'middle'};ws.getRow(1).height=28;ws.getRow(2).height=24;
  ws.mergeCells('A3:I3');ws.getCell('A3').value='UNIT 11-12, 10/F., FU HANG INDUSTRIAL BUILDING, NO. 1 HOK YUEN STREET EAST, HUNG HOM, KOWLOON, HONG KONG';ws.getCell('A3').alignment={horizontal:'center'};ws.getCell('A3').font={name:'Arial',size:9};
  ws.mergeCells('A4:I4');ws.getCell('A4').value='TEL : (852) 2363 5409     FAX : (852) 2765 0343';ws.getCell('A4').alignment={horizontal:'center'};ws.getCell('A4').font={name:'Arial',size:9};
  ws.getRow(5).height=6;for(let c=1;c<=9;c++)ws.getCell(5,c).border={bottom:{style:'medium',color:{argb:'FF000000'}}};
  ws.mergeCells('A6:D6');ws.getCell('A6').value='Sales Invoice';ws.getCell('A6').font={name:'Arial',size:18,bold:true};
  const left=[['No.',normalizeKey($('#invoiceNo').value)],['Invoice Date',$('#invoiceDate').value],['Shipment Method',$('#shipmentMethod').value],['Currency',$('#currency').value],['Customer Code',normalizeCustomerCode($('#customerCode').value)],['Customer',$('#customerName').value],['Address',$('#customerAddress').value]];
  let r=7;for(const[k,v]of left){ws.getCell(r,1).value=k+' :';ws.getCell(r,1).font={bold:k==='Customer'};ws.mergeCells(r,2,r,4);ws.getCell(r,2).value=v;ws.getCell(r,2).alignment={wrapText:true};r+=k==='Address'?2:1}
  ws.mergeCells('F6:I6');ws.getCell('F6').value="Vender's Banker";ws.getCell('F6').font={bold:true,size:12};const bank=['The Hong Kong & Shanghai Banking Corporation Ltd.','Address : 41 Ma Tau Wai Road,Hung Hom,Kowloon,Hong Kong','A/C # : 012-593570-001','A/C Name : Universe Gems & Jewellery Co.'];bank.forEach((v,i)=>{ws.mergeCells(7+i,6,7+i,9);ws.getCell(7+i,6).value=v;ws.getCell(7+i,6).alignment={wrapText:true}});
  const headerRow=15;const headers=['No.','Article No.','Description','','Image','Quantity','Unit','Unit Price','Amount'];headers.forEach((v,i)=>{const c=ws.getCell(headerRow,i+1);c.value=v;c.font={bold:true,size:9};c.alignment={horizontal:i>=5?'center':'left'};setCellBorder(c)});ws.mergeCells(headerRow,3,headerRow,4);ws.getCell(headerRow,3).value='Description';ws.getRow(headerRow).height=20;ws.pageSetup.printTitlesRow='1:15';
  let row=headerRow+1,itemCount=0;
  for(const item of state.items){
    const lines=Math.max(4,item.descriptions.length+1),start=row,end=row+lines-1;itemCount++;
    for(const range of [[start,1,end,1],[start,2,end,2],[start,3,end,4],[start,5,end,5],[start,6,end,6],[start,7,end,7],[start,8,end,8],[start,9,end,9]])ws.mergeCells(...range);
    ws.getCell(start,1).value=itemCount;ws.getCell(start,2).value=`Lot.No. : ${item.lotNo}\n${item.artNo}`;ws.getCell(start,3).value=item.descriptions.join('\n');ws.getCell(start,6).value=item.qty;ws.getCell(start,7).value=item.unit;ws.getCell(start,8).value=item.unitPrice;ws.getCell(start,9).value=item.qty*item.unitPrice;
    for(const col of[1,2,3,5,6,7,8,9]){const c=ws.getCell(start,col);c.alignment={vertical:'top',wrapText:true,horizontal:col>=6?'center':'left'};c.font={name:'Arial',size:9};c.border={bottom:{style:'thin',color:{argb:'FFB8BEC8'}}}}
    ws.getCell(start,8).numFmt='$#,##0.00';ws.getCell(start,9).numFmt='$#,##0.00';for(let rr=start;rr<=end;rr++)ws.getRow(rr).height=15;
    const img=getSelectedImage(item);if(img?.file&&/\.(jpe?g|png)$/i.test(img.file.name)){try{const id=wb.addImage({base64:await fileToBase64(img.file),extension:imageExtension(img.file)});ws.addImage(id,{tl:{col:4.08,row:start-1+.12},br:{col:4.92,row:end-.12},editAs:'oneCell'})}catch{}}
    row=end+1;if(itemCount%8===0&&itemCount<state.items.length)ws.getRow(row-1).addPageBreak();
  }
  const t=calcTotals();row+=1;ws.mergeCells(row,1,row,6);ws.getCell(row,1).value=`Total Quantity : ${t.qty}`;ws.getCell(row,1).alignment={horizontal:'right'};ws.getCell(row,7).value='Sub Total:';ws.getCell(row,8).value=t.subtotal;ws.mergeCells(row,8,row,9);ws.getCell(row,8).numFmt='$#,##0.00';row++;
  ws.getCell(row,7).value='Discount:';ws.getCell(row,8).value=t.discount;ws.mergeCells(row,8,row,9);ws.getCell(row,8).numFmt='$#,##0.00';row++;
  ws.getCell(row,7).value=`Total (${($('#currency').value||'USD')}):`;ws.getCell(row,7).font={bold:true};ws.getCell(row,8).value=t.total;ws.mergeCells(row,8,row,9);ws.getCell(row,8).numFmt='$#,##0.00';ws.getCell(row,8).font={bold:true};row+=2;
  ws.mergeCells(row,1,row,9);ws.getCell(row,1).value=`Total Amount : ${$('#currency').value||'USD'} ${numberToWords(t.total)}`;ws.getCell(row,1).font={bold:true};row+=2;ws.mergeCells(row,1,row,9);ws.getCell(row,1).value=`Remark : ${$('#remark').value||''}`;row+=3;ws.mergeCells(row,1,row,4);ws.getCell(row,1).value='Vender Signature :';ws.mergeCells(row,6,row,9);ws.getCell(row,6).value='Accept By :';
  ws.headerFooter.oddFooter='&RPage &P of &N';ws.pageSetup.printArea=`A1:I${row+2}`;return wb;
}
async function exportInvoiceExcel(){const no=normalizeKey($('#invoiceNo').value)||'Invoice',wb=await buildInvoiceWorkbook(),buffer=await wb.xlsx.writeBuffer(),fn=`${sanitizeFileName(no)}.xlsx`;downloadBlob(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),fn);return fn}
$('#exportInvoiceBtn').addEventListener('click',async()=>{try{await exportInvoiceExcel()}catch(e){alert(`未能輸出 Invoice Excel：${e.message}`)}});

function buildRemainingStockBlob(invoiceNo){
  if(!state.stockRows||state.stockHeaderRow<0||state.stockLotCol<0)throw new Error('請先匯入真實倉存 Excel。');const output=[];
  for(let i=0;i<state.stockRows.length;i++){const row=state.stockRows[i]||[];if(i<=state.stockHeaderRow){output.push(row);continue}const lot=String(row[state.stockLotCol]??'').replace(/\.0$/,'').replace(/\s+/g,'');if(!lot||state.products.has(lot))output.push(row)}
  const ws=XLSX.utils.aoa_to_sheet(output),wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Remaining Stock');return new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array',compression:true})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
$('#confirmInvoiceBtn').addEventListener('click',async()=>{
  if(!state.items.length)return alert('Invoice 沒有貨品。');const invoiceNo=normalizeKey($('#invoiceNo').value)||'Invoice';if(!confirm(`Confirm ${invoiceNo}？`))return;const snapshot=[...state.items];
  try{
    const invoiceWb=await buildInvoiceWorkbook(),invoiceBuffer=await invoiceWb.xlsx.writeBuffer();for(const item of snapshot){state.products.delete(item.lotNo);state.soldLots.set(item.lotNo,{invoiceNo,customerCode:normalizeCustomerCode($('#customerCode').value),date:$('#invoiceDate').value})}
    const remainBlob=buildRemainingStockBlob(invoiceNo),stamp=new Date().toISOString().replace(/[-:T]/g,'').slice(0,12);downloadBlob(new Blob([invoiceBuffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),`${sanitizeFileName(invoiceNo)}.xlsx`);setTimeout(()=>downloadBlob(remainBlob,`Remaining_Stock_${sanitizeFileName(invoiceNo)}_${state.products.size}pcs_${stamp}.xlsx`),450);
    state.items=[];state.nextSequence=1;$('#discountAmount').value='0';renderItems();rebuildImageIndex();persistLight();renderPreview();alert(`Invoice 已 Confirm。\n已建立 Invoice Excel 及 Remaining Stock Excel。`);
  }catch(err){for(const item of snapshot){state.products.set(item.lotNo,{lotNo:item.lotNo,artNo:item.artNo,price:item.price,unit:item.unit,desc2:item.desc2,descriptions:item.descriptions});state.soldLots.delete(item.lotNo)}alert(`Confirm 失敗：${err.message}`)}
});

$('#scanBtn').addEventListener('click',startScanner);$('#closeScannerBtn').addEventListener('click',stopScanner);$('#pauseScannerBtn').addEventListener('click',toggleScannerPause);$('#torchBtn').addEventListener('click',toggleTorch);for(const z of[1,2,3,4])$(`#zoom${z}Btn`).addEventListener('click',()=>setZoom(z));
function setScannerControlsDisabled(disabled){['#scanBtn','#pauseScannerBtn','#torchBtn','#zoom1Btn','#zoom2Btn','#zoom3Btn','#zoom4Btn'].forEach(x=>$(x).disabled=disabled);$('#closeScannerBtn').disabled=false}
async function startScanner(){
  if(state.scannerTransitioning||state.scannerRunning)return;state.scannerTransitioning=true;state.scannerCancelRequested=false;setScannerControlsDisabled(true);$('#scannerError').classList.add('hidden');const dlg=$('#scannerDialog');if(!dlg.open)dlg.showModal();
  if(typeof Html5Qrcode==='undefined'){state.scannerTransitioning=false;setScannerControlsDisabled(false);return setStatus('#scannerError','掃描程式未載入。','error')}
  try{
    if(state.scanner){try{await state.scanner.clear()}catch{}}$('#reader').innerHTML='';const formats=[Html5QrcodeSupportedFormats.CODE_128,Html5QrcodeSupportedFormats.CODE_39,Html5QrcodeSupportedFormats.EAN_13,Html5QrcodeSupportedFormats.EAN_8,Html5QrcodeSupportedFormats.UPC_A,Html5QrcodeSupportedFormats.UPC_E,Html5QrcodeSupportedFormats.ITF];state.scanner=new Html5Qrcode('reader',{verbose:false,formatsToSupport:formats});
    const cameras=await Html5Qrcode.getCameras();if(!cameras?.length)throw new Error('找不到相機。');const pattern=/(back|rear|environment|後置|背面)/i,cam=cameras.find(c=>pattern.test(c.label||''))||cameras[cameras.length-1];
    await state.scanner.start(cam.id,{fps:24,qrbox:(w,h)=>({width:Math.floor(w*.72),height:Math.max(48,Math.min(70,Math.floor(h*.09)))}),aspectRatio:1.7778,disableFlip:true,experimentalFeatures:{useBarCodeDetectorIfSupported:true}},onBarcodeSuccess,()=>{});
    state.scannerRunning=true;state.scannerPaused=false;setScannerControlsDisabled(false);const caps=state.scanner.getRunningTrackCapabilities?.();if(!caps?.torch)$('#torchBtn').disabled=true;if(caps?.zoom){for(const z of[1,2,3,4])$(`#zoom${z}Btn`).disabled=z>(caps.zoom.max||1);await setZoom(Math.min(3,caps.zoom.max||1))}else for(const z of[1,2,3,4])$(`#zoom${z}Btn`).disabled=true;try{await state.scanner.applyVideoConstraints({advanced:[{focusMode:'continuous',exposureMode:'continuous'}]})}catch{}
  }catch(err){state.scannerRunning=false;setScannerControlsDisabled(false);setStatus('#scannerError',`相機無法啟動：${err?.message||err}`,'error')}finally{state.scannerTransitioning=false}
}
function showScannerFeedback(result){clearTimeout(state.feedbackTimer);const el=$('#scannerFeedback');el.className=`scanner-feedback ${result.status}`;el.classList.remove('hidden');$('strong',el).textContent=result.lot||'—';$('span',el).textContent=result.status==='success'?'':result.message;const delay=result.status==='notfound'?2000:result.status==='duplicate'?1500:1100;state.feedbackTimer=setTimeout(()=>el.classList.add('hidden'),delay);$('#scannerItemCount').textContent=state.items.length}
async function onBarcodeSuccess(decodedText){const value=normalizeLot(decodedText),now=Date.now();if(!value)return;if(state.lastScan.value===value&&now-state.lastScan.time<1800)return;state.lastScan={value,time:now};const result=addByLot(value,'scanner');showScannerFeedback(result)}
async function setZoom(value){if(!state.scannerRunning||state.scannerTransitioning)return;try{await state.scanner.applyVideoConstraints({advanced:[{zoom:value}]});for(const z of[1,2,3,4])$(`#zoom${z}Btn`).classList.toggle('active',z===value)}catch{}}
async function toggleScannerPause(){if(!state.scannerRunning||!state.scanner||state.scannerTransitioning)return;try{if(state.scannerPaused){state.scanner.resume();state.scannerPaused=false;$('#pauseScannerBtn').textContent='暫停'}else{state.scanner.pause(true);state.scannerPaused=true;$('#pauseScannerBtn').textContent='繼續'}}catch{}}
async function toggleTorch(){if(state.scannerTransitioning)return;try{const current=$('#torchBtn').dataset.on==='1';await state.scanner.applyVideoConstraints({advanced:[{torch:!current}]});$('#torchBtn').dataset.on=current?'0':'1';$('#torchBtn').textContent=current?'🔦':'關燈'}catch{}}
async function stopScanner(){state.scannerCancelRequested=true;if($('#scannerDialog').open)$('#scannerDialog').close();if(state.scannerTransitioning)return;state.scannerTransitioning=true;setScannerControlsDisabled(true);try{if(state.scannerRunning&&state.scanner)await state.scanner.stop();if(state.scanner)await state.scanner.clear()}catch{}finally{state.scanner=null;state.scannerRunning=false;state.scannerPaused=false;state.scannerTransitioning=false;setScannerControlsDisabled(false)}}

function spokenToDigits(raw){const map={'零':'0','〇':'0','一':'1','么':'1','二':'2','兩':'2','三':'3','四':'4','五':'5','六':'6','七':'7','八':'8','九':'9'};return String(raw||'').split('').map(c=>/\d/.test(c)?c:(map[c]||'')).join('')}
$('#voiceLotBtn').addEventListener('click',()=>{
  const SR=window.SpeechRecognition||window.webkitSpeechRecognition;if(!SR){$('#lotInput').setAttribute('inputmode','text');$('#lotInput').focus();showAddMessage('請使用 iPhone 鍵盤咪高峰說出 LOTNO。');return}
  const rec=new SR();rec.lang='zh-HK';rec.interimResults=false;rec.maxAlternatives=1;$('#voiceLotBtn').disabled=true;rec.onresult=e=>{const raw=e.results?.[0]?.[0]?.transcript||'',lot=spokenToDigits(raw);$('#lotInput').value=lot;if(lot){const result=addByLot(lot,'voice');if(result.status!=='success')showAddMessage(`LOTNO ${lot} ${result.message}`,'error')}else showAddMessage('未能辨識 LOTNO。','error')};rec.onerror=()=>showAddMessage('語音輸入失敗。','error');rec.onend=()=>{$('#voiceLotBtn').disabled=false};try{rec.start()}catch{$('#voiceLotBtn').disabled=false}
});

function persistLight(){try{localStorage.setItem('ui-v07',JSON.stringify({customerCode:$('#customerCode').value,customerName:$('#customerName').value,customerAddress:$('#customerAddress').value,salesRate:$('#salesRate').value,currency:$('#currency').value,shipmentMethod:$('#shipmentMethod').value,customerTerms:$('#customerTerms').value,invoiceNo:$('#invoiceNo').value,invoiceDate:$('#invoiceDate').value,discount:$('#discountAmount').value,remark:$('#remark').value,items:state.items,soldLots:[...state.soldLots.entries()]}))}catch{}}
function restoreLight(){try{const d=JSON.parse(localStorage.getItem('ui-v07')||'null');if(!d)return;for(const[id,key]of[['#customerCode','customerCode'],['#customerName','customerName'],['#customerAddress','customerAddress'],['#salesRate','salesRate'],['#currency','currency'],['#shipmentMethod','shipmentMethod'],['#customerTerms','customerTerms'],['#invoiceNo','invoiceNo'],['#invoiceDate','invoiceDate'],['#discountAmount','discount'],['#remark','remark']])if(d[key]!==undefined)$(id).value=d[key];if(Array.isArray(d.items)){state.items=d.items;let max=0;state.items.forEach((x,i)=>{if(!x.sequence)x.sequence=i+1;max=Math.max(max,x.sequence||0)});state.nextSequence=max+1}if(Array.isArray(d.soldLots))state.soldLots=new Map(d.soldLots)}catch{}}
['#customerCode','#customerName','#customerAddress','#shipmentMethod','#customerTerms','#invoiceNo','#invoiceDate','#remark'].forEach(id=>$(id).addEventListener('change',persistLight));
window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();state.installPrompt=e;$('#installBtn').classList.remove('hidden')});$('#installBtn').addEventListener('click',async()=>{if(state.installPrompt){state.installPrompt.prompt();state.installPrompt=null;$('#installBtn').classList.add('hidden')}});if('serviceWorker'in navigator)window.addEventListener('load',()=>navigator.serviceWorker.register('./sw.js').catch(()=>{}));
restoreLight();renderSelectedCustomer();renderItems();updateHeader();renderPreview();
