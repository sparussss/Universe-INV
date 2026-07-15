const $=(s,r=document)=>r.querySelector(s),$$=(s,r=document)=>[...r.querySelectorAll(s)];
const state={products:new Map(),customers:new Map(),imageFiles:new Map(),items:[],stockRows:[],stockHeaders:[],stoneAliases:new Map(),stoneMappingName:'',articleMap:new Map(),articleMappingName:'',invoiceTemplateBuffer:null,invoiceTemplateName:'',scanner:null,scannerBusy:false,scannerRunning:false,scannerZoom:{min:1,max:1,step:1,current:1}};
const norm=v=>String(v??'').trim(),normCode=v=>String(v??'').replace(/\s+/g,'').toUpperCase(),normArt=v=>norm(v).toUpperCase();
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const today=()=>{const d=new Date(),y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,'0'),day=String(d.getDate()).padStart(2,'0');return `${y}-${m}-${day}`};$('#invoiceDate').value=today();
function invoiceYear(){return String(new Date().getFullYear()).slice(-2)}
function invoiceSequenceKey(yy=invoiceYear()){return `universeInvoiceSeq_${yy}`}
function getNextInvoiceSequence(yy=invoiceYear()){
  const saved=Number(localStorage.getItem(invoiceSequenceKey(yy))||1);
  return Number.isInteger(saved)&&saved>0?saved:1;
}
function formatInvoiceNo(seq=getNextInvoiceSequence(),yy=invoiceYear()){
  return `INV${yy}${String(seq).padStart(4,'0')}`;
}
function setDefaultInvoiceNo(force=false){
  const el=$('#invoiceNo');
  if(force||!norm(el.value))el.value=formatInvoiceNo();
}
function advanceInvoiceSequence(confirmedNo){
  const yy=invoiceYear();
  const m=String(confirmedNo||'').toUpperCase().match(/^INV(\d{2})(\d{4})$/);
  let next=getNextInvoiceSequence(yy)+1;
  if(m&&m[1]===yy)next=Math.max(next,Number(m[2])+1);
  localStorage.setItem(invoiceSequenceKey(yy),String(next));
  $('#invoiceNo').value=formatInvoiceNo(next,yy);
}
setDefaultInvoiceNo();
function status(id,msg,type=''){const el=$(id);el.textContent=msg;el.className='notice'+(type?' '+type:'')}
function setImportCollapsed(key,collapsed=true){const card=document.querySelector(`[data-import-card="${key}"]`);if(!card)return;card.classList.toggle('collapsed',collapsed);const btn=card.querySelector('.import-toggle');if(btn){btn.textContent=collapsed?'展開':'收合';btn.setAttribute('aria-expanded',String(!collapsed))}}
$$('.import-toggle').forEach(btn=>btn.addEventListener('click',()=>{const card=btn.closest('.import-card');setImportCollapsed(card?.dataset.importCard,!card.classList.contains('collapsed'))}));
function field(row,names){const keys=Object.keys(row);for(const n of names){const k=keys.find(x=>x.trim().toUpperCase()===n);if(k)return row[k]}return''}
function fmt(v){return new Intl.NumberFormat('en-US',{style:'currency',currency:$('#currency').value||'USD',minimumFractionDigits:2}).format(Number(v)||0)}
function totals(){const qty=state.items.reduce((a,x)=>a+x.qty,0),sub=state.items.reduce((a,x)=>a+x.qty*x.unitPrice,0),discount=Math.max(0,Number($('#discountAmount').value)||0);return{qty,sub,discount,total:Math.max(0,sub-discount)}}
function updateTotals(){const t=totals();$('#totalQty').textContent=t.qty;$('#subtotal').textContent=fmt(t.sub);$('#discountDisplay').textContent=fmt(t.discount);$('#grandTotal').textContent=fmt(t.total);$('#productCount').textContent=state.products.size;$('#customerCount').textContent=state.customers.size;$('#invoiceCount').textContent=state.items.length;$('#headerTotal').textContent=fmt(t.total)}
$$('.tab').forEach(b=>b.onclick=()=>{$$('.tab').forEach(x=>x.classList.remove('active'));$$('.tab-panel').forEach(x=>x.classList.remove('active'));b.classList.add('active');$('#'+b.dataset.tab).classList.add('active');if(b.dataset.tab==='invoice')renderCustomerSummary();if(b.dataset.tab==='preview')renderPreview()});
async function readWB(file){if(typeof XLSX==='undefined')throw new Error('Excel 程式未載入');return XLSX.read(await file.arrayBuffer(),{type:'array'})}
$('#stockInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const wb=await readWB(f),ws=wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(ws,{defval:''});state.stockRows=rows;state.stockHeaders=Object.keys(rows[0]||{});const map=new Map();for(const r of rows){const lot=norm(field(r,['LOTNO']));const art=normArt(field(r,['ARTNO']));const price=Number(field(r,['PRICE']));if(!lot||!art||!Number.isFinite(price))continue;const desc=[];for(let i=1;i<=6;i++){const v=norm(field(r,[`DESC${i}`]));if(v)desc.push(v)}map.set(lot,{lotNo:lot,artNo:art,price,unit:norm(field(r,['UNIT']))||'PC',article:norm(field(r,['ARTICLE']))||'',descriptions:desc,desc2:norm(field(r,['DESC2']))})}state.products=map;status('#stockStatus',`已匯入 ${f.name}：${map.size} 件貨品。`,'ok');setImportCollapsed('stock',true);updateTotals()}catch(err){status('#stockStatus','匯入失敗：'+err.message,'error');setImportCollapsed('stock',false)}};
$('#customerInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const wb=await readWB(f),ws=wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});const map=new Map();for(const r of rows){const code=normCode(r[0]),company=norm(r[1]);if(!code||!company||code.includes('CUSTOMER'))continue;const raw=r[11],num=Number(raw),rate=(raw===''||!Number.isFinite(num))?0.34:num;map.set(code,{code,company,address:[r[2],r[3],r[4]].map(norm).filter(Boolean).join('\n'),rate,terms:norm(r[10])})}state.customers=map;status('#customerStatus',`已匯入 ${f.name}：${map.size} 位客戶。`,'ok');setImportCollapsed('customer',true);updateTotals()}catch(err){status('#customerStatus','匯入失敗：'+err.message,'error');setImportCollapsed('customer',false)}};
const FALLBACK_STONE_ALIASES=new Map([
  ['SKY BTO','SKY BT'],['SKY BT','SKY BT'],['QAM','AM'],['YCT','CT'],['BTO','BT'],['LBT','L.BT'],['MG','MG'],['AQ','AQ'],['PAM','P.AM'],['PTQ','PTR'],['GPD','PD'],['GPS','G.AM'],['GAM','G.AM'],['TZ','TZ']
]);
const FALLBACK_ARTICLE_MAP=new Map([
  ['RG','RING /w SEMI-PRECIOUS'],
  ['ER','EARRING /w SEMI-PRECIOUS'],
  ['PT','PENDANT /w SEMI-PRECIOUS'],
  ['BR','BROOCH /w SEMI-PRECIOUS'],
  ['NL','NECKLACE /w SEMI-PRECIOUS'],
  ['BL','BRACELET /w SEMI-PRECIOUS'],
  ['BG','BANGLE /w SEMI-PRECIOUS']
]);
function activeArticleMap(){return state.articleMap.size?state.articleMap:FALLBACK_ARTICLE_MAP}
function articleDescriptionFor(item){
  const prefix=normArt(item?.artNo).split('-')[0].split('.')[0];
  return state.articleMap.size?(state.articleMap.get(prefix)||''):'';
}

function activeStoneAliases(){return state.stoneAliases.size?state.stoneAliases:FALLBACK_STONE_ALIASES}
$('#stoneMappingInput').onchange=async e=>{
  const f=e.target.files[0];if(!f)return;
  try{
    const wb=await readWB(f);
    const sheetName=wb.SheetNames.find(n=>n.trim().toUpperCase()==='STONE LIST')||wb.SheetNames[0];
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:''});
    const headerIndex=rows.findIndex(r=>r.some(v=>String(v).trim().toUpperCase()==='BREAKDOWN')&&r.some(v=>String(v).trim().toUpperCase()==='QUOTATION'));
    if(headerIndex<0)throw new Error('找不到 BREAKDOWN / QUOTATION 欄位');
    const header=rows[headerIndex].map(v=>String(v).trim().toUpperCase());
    const bCol=header.indexOf('BREAKDOWN'),qCol=header.indexOf('QUOTATION');
    const aliases=new Map();
    for(const r of rows.slice(headerIndex+1)){
      const breakdown=norm(r[bCol]),quotation=norm(r[qCol]);
      if(!breakdown||!quotation)continue;
      for(const code of breakdown.split(/[,，]/).map(norm).filter(Boolean))aliases.set(code.toUpperCase(),quotation.toUpperCase());
    }
    if(!aliases.size)throw new Error('對照表沒有有效資料');
    state.stoneAliases=aliases;state.stoneMappingName=f.name;
    status('#stoneMappingStatus',`已匯入 ${f.name}：${aliases.size} 個石種代碼對照。`,'ok');setImportCollapsed('stone',true);
    for(const item of state.items){item.imageVariant=chooseVariant(item)}
    renderItems();
  }catch(err){status('#stoneMappingStatus','匯入失敗：'+(err.message||err),'error');setImportCollapsed('stone',false)}
};
$('#articleMappingInput').onchange=async e=>{
  const f=e.target.files[0];if(!f)return;
  try{
    const wb=await readWB(f),ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    const map=new Map();
    for(const row of rows){
      const prefix=normCode(row[0]);
      const description=norm(row[1]);
      if(!prefix||!description||prefix==='PREFIX')continue;
      map.set(prefix,description);
    }
    if(!map.size)throw new Error('找不到 Prefix / Article Description 對照。');
    state.articleMap=map;state.articleMappingName=f.name;
    status('#articleMappingStatus',`已匯入 ${f.name}：${map.size} 個 Article 對照。`,'ok');setImportCollapsed('article',true);
  }catch(err){status('#articleMappingStatus','匯入失敗：'+(err.message||err),'error');setImportCollapsed('article',false)}
};

$('#invoiceTemplateInput').onchange=async e=>{
  const f=e.target.files[0];if(!f)return;
  try{
    const buf=await f.arrayBuffer();
    if(typeof ExcelJS==='undefined')throw new Error('Excel 範本程式未載入');
    const test=new ExcelJS.Workbook();await test.xlsx.load(buf.slice(0));
    if(!test.worksheets.length)throw new Error('範本沒有工作表');
    state.invoiceTemplateBuffer=buf;state.invoiceTemplateName=f.name;
    status('#invoiceTemplateStatus',`已匯入 ${f.name}；匯出 Excel Invoice 時會套用此範本。`,'ok');setImportCollapsed('template',true);
  }catch(err){state.invoiceTemplateBuffer=null;status('#invoiceTemplateStatus','匯入失敗：'+(err.message||err),'error');setImportCollapsed('template',false)}
};

function parseImage(file){const stem=file.name.replace(/\.[^.]+$/,'').trim();const arts=[...new Set([...state.products.values()].map(x=>x.artNo))].sort((a,b)=>b.length-a.length);const art=arts.find(a=>stem.toUpperCase()===a||stem.toUpperCase().startsWith(a+' '));if(!art)return null;let variant=stem.slice(art.length).trim().replace(/\s*\(\d+\)$/,'').trim()||'Default';const dup=(stem.match(/\((\d+)\)$/)||[])[1];return{art,variant,dup:dup?Number(dup):0,file}}
$('#imageFolderInput').onchange=e=>{const map=new Map();for(const f of e.target.files){const p=parseImage(f);if(!p)continue;const key=p.art+'|'+p.variant.toUpperCase(),existing=map.get(key);if(!existing||p.dup<existing.dup)map.set(key,p)}state.imageFiles=new Map();for(const p of map.values()){const arr=state.imageFiles.get(p.art)||[];arr.push({variant:p.variant,url:URL.createObjectURL(p.file),fileName:p.file.name,file:p.file});state.imageFiles.set(p.art,arr)}for(const arr of state.imageFiles.values())arr.sort((a,b)=>a.variant==='Default'?-1:b.variant==='Default'?1:a.variant.localeCompare(b.variant));status('#imageStatus',`已選擇圖片 Folder：${e.target.files.length} 張圖片，配對 ${state.imageFiles.size} 個款號。`,'ok');setImportCollapsed('images',true);renderItems()};
function searchCustomers(q){const s=norm(q).toUpperCase(),c=normCode(q);return [...state.customers.values()].filter(x=>x.code.includes(c)||x.company.toUpperCase().includes(s)).slice(0,10)}
function showMatches(){const box=$('#customerMatches'),m=searchCustomers($('#customerSearch').value);box.innerHTML='';if(!m.length){box.innerHTML='<div class="notice">找不到客戶。</div>';return}m.forEach(c=>{const b=document.createElement('button');b.className='customer-match';b.innerHTML=`<span><strong>${esc(c.code)} · ${esc(c.company)}</strong><small>${esc(c.address).replace(/\n/g,' · ')}</small></span><span>${c.rate}</span>`;b.onclick=()=>selectCustomer(c);box.appendChild(b)})}
function selectCustomer(c){$('#customerCode').value=c.code;$('#customerName').value=c.company;$('#customerAddress').value=c.address;$('#salesRate').value=c.rate;$('#customerTerms').value=c.terms;$('#customerMatches').innerHTML='';$('#customerSearch').value='';reprice();renderCustomerSummary()}
$('#findCustomerBtn').onclick=showMatches;
let customerSearchTimer=null;
$('#customerSearch').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();showMatches()}};
$('#customerSearch').oninput=e=>{
  clearTimeout(customerSearchTimer);
  const q=norm(e.target.value);
  if(q.length<2){$('#customerMatches').innerHTML='';return}
  customerSearchTimer=setTimeout(showMatches,120);
};
function renderCustomerSummary(){const code=$('#customerCode').value,name=$('#customerName').value;$('#selectedCustomerSummary').innerHTML=code||name?`<strong>${esc(code)} · ${esc(name)}</strong><span>Sales Rate ${esc($('#salesRate').value)} · ${esc($('#currency').value)}</span>`:'尚未選擇客戶。'}
function chooseVariant(p){
  const imgs=state.imageFiles.get(p.artNo)||[];if(!imgs.length)return'Default';
  const d=(p.desc2||'').toUpperCase();
  const hits=[];
  for(const [code,variant] of [...activeStoneAliases().entries()].sort((a,b)=>b[0].length-a[0].length)){
    const pos=d.indexOf(code.toUpperCase());if(pos>=0)hits.push({pos,variant:String(variant).toUpperCase()});
  }
  hits.sort((a,b)=>a.pos-b.pos);
  const ordered=[...new Set(hits.map(x=>x.variant))];
  const candidates=[];
  if(ordered.length>1){candidates.push(ordered.join('+'));candidates.push([...ordered].reverse().join('+'))}
  candidates.push(...ordered);
  for(const wanted of candidates){const hit=imgs.find(x=>x.variant.toUpperCase()===wanted);if(hit)return hit.variant}
  return imgs.find(x=>x.variant==='Default')?.variant||imgs[0].variant
}
function normalizeLotInput(raw){
  const map={'零':'0','〇':'0','一':'1','二':'2','兩':'2','两':'2','三':'3','四':'4','五':'5','六':'6','七':'7','八':'8','九':'9'};
  return String(raw??'')
    .replace(/[零〇一二兩两三四五六七八九]/g,ch=>map[ch]||ch)
    .replace(/[^0-9A-Za-z]/g,'')
    .toUpperCase();
}
function refocusLotInput(selectAll=false){
  const input=$('#lotInput');
  requestAnimationFrame(()=>{
    input.focus({preventScroll:true});
    if(selectAll)input.select();
    else{const n=input.value.length;try{input.setSelectionRange(n,n)}catch{}}
  });
}
function addLot(raw){
  const lot=normalizeLotInput(raw);
  if(!lot){status('#addMessage','請輸入 LOTNO。','error');refocusLotInput();return false}
  const p=state.products.get(lot);
  if(!p){status('#addMessage',`找不到 LOTNO ${lot}。`,'error');refocusLotInput(true);return false}
  if(state.items.some(x=>x.lotNo===lot)){status('#addMessage',`LOTNO ${lot} 已在 Invoice。`,'error');refocusLotInput(true);return false}
  const rate=Number($('#salesRate').value)||0;
  state.items.push({...p,id:Date.now()+Math.random(),seq:state.items.length+1,qty:1,unitPrice:Math.ceil(p.price*rate),imageVariant:chooseVariant(p)});
  $('#lotInput').value='';
  status('#addMessage',`已加入 ${p.artNo} / LOTNO ${lot}`,'ok');
  renderItems();
  setTimeout(()=>{$('#invoiceItems').scrollTo({top:0,behavior:'smooth'});refocusLotInput()},50);
  return true;
}
$('#addLotBtn').onclick=()=>addLot($('#lotInput').value);
$('#lotInput').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();addLot(e.target.value)}};
$('#lotInput').oninput=e=>{e.target.value=e.target.value.replace(/[，,。\.\-–—_]/g,' ')};
$('#lotInput').onfocus=e=>setTimeout(()=>e.target.select(),50);
function getImg(item){const arr=state.imageFiles.get(item.artNo)||[];return arr.find(x=>x.variant===item.imageVariant)||arr[0]}
function placeholder(t){return `data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="200" height="200"><rect width="100%" height="100%" fill="#f1f5f9"/><text x="50%" y="50%" text-anchor="middle" dominant-baseline="middle" font-family="Arial" font-size="18" fill="#64748b">${t}</text></svg>`)}`}
function renderItems(){const box=$('#invoiceItems');box.innerHTML='';if(!state.items.length){box.className='invoice-items empty-state';box.textContent='尚未加入貨品。';updateTotals();return}box.className='invoice-items';[...state.items].reverse().forEach(item=>{const node=$('#itemTemplate').content.firstElementChild.cloneNode(true);$('.item-seq',node).textContent=item.seq;$('.item-artno',node).textContent=item.artNo;$('.item-lot',node).textContent=`LOTNO ${item.lotNo}`;$('.item-desc',node).textContent=item.descriptions.join('\n');$('.item-price-note',node).textContent=`${item.price}u × ${Number($('#salesRate').value)||0} → ${fmt(item.unitPrice)}`;$('.item-thumb',node).src=getImg(item)?.url||placeholder(item.artNo);const sel=$('.variant-select',node),arr=state.imageFiles.get(item.artNo)||[];if(arr.length){arr.forEach(x=>{const o=document.createElement('option');o.value=x.variant;o.textContent='圖片：'+x.variant;o.selected=x.variant===item.imageVariant;sel.appendChild(o)});sel.onchange=e=>{item.imageVariant=e.target.value;renderItems()}}else{sel.innerHTML='<option>沒有圖片</option>';sel.disabled=true}$('.qty-input',node).value=item.qty;$('.price-input',node).value=item.unitPrice;$('.qty-input',node).onchange=e=>{item.qty=Math.max(1,Number(e.target.value)||1);updateTotals()};$('.price-input',node).onchange=e=>{item.unitPrice=Math.max(0,Math.ceil(Number(e.target.value)||0));updateTotals()};$('.delete-item',node).onclick=()=>{if(confirm(`刪除 ${item.artNo}？`)){state.items=state.items.filter(x=>x.id!==item.id);state.items.forEach((x,i)=>x.seq=i+1);renderItems()}};box.appendChild(node)});updateTotals()}
$('#scrollLatestBtn').onclick=()=>$('#invoiceItems').scrollTo({top:0,behavior:'smooth'});$('#clearInvoiceBtn').onclick=()=>{if(confirm('清空目前 Invoice？')){state.items=[];renderItems()}};
function reprice(){const r=Number($('#salesRate').value)||0;state.items.forEach(x=>x.unitPrice=Math.ceil(x.price*r));renderItems()}$('#salesRate').onchange=reprice;$('#currency').onchange=()=>{renderItems();renderCustomerSummary()};$('#discountAmount').oninput=updateTotals;
function words(n){return String(Math.floor(n))}
function numberToWords(value){
  let n=Math.floor(Number(value)||0);
  if(n===0)return 'ZERO';
  if(n<0)return 'MINUS '+numberToWords(Math.abs(n));
  const ones=['','ONE','TWO','THREE','FOUR','FIVE','SIX','SEVEN','EIGHT','NINE','TEN','ELEVEN','TWELVE','THIRTEEN','FOURTEEN','FIFTEEN','SIXTEEN','SEVENTEEN','EIGHTEEN','NINETEEN'];
  const tens=['','','TWENTY','THIRTY','FORTY','FIFTY','SIXTY','SEVENTY','EIGHTY','NINETY'];
  const underThousand=x=>{
    const parts=[];
    if(x>=100){parts.push(ones[Math.floor(x/100)]+' HUNDRED');x%=100}
    if(x>=20){parts.push(tens[Math.floor(x/10)]);x%=10}
    if(x>0)parts.push(ones[x]);
    return parts.join(' ');
  };
  const scales=[
    [1_000_000_000,'BILLION'],
    [1_000_000,'MILLION'],
    [1_000,'THOUSAND'],
    [1,'']
  ];
  const parts=[];
  for(const [size,label] of scales){
    if(n>=size){
      const chunk=Math.floor(n/size);
      n%=size;
      const text=underThousand(chunk);
      if(text)parts.push(label?`${text} ${label}`:text);
    }
  }
  return parts.join(' ');
}
function currencyWords(code){
  return ({USD:'US DOLLARS',EUR:'EUROS',JPY:'JAPANESE YEN',HKD:'HONG KONG DOLLARS'})[String(code||'').toUpperCase()]||String(code||'').toUpperCase();
}
function renderPreview(){const t=totals(),rows=state.items.map((x,i)=>`<tr><td>${i+1}</td><td><strong>Lot.No. : ${esc(x.lotNo)}</strong><br>${esc(x.artNo)}</td><td>${x.descriptions.map(esc).join('<br>')}</td><td class="num">${x.qty}</td><td>${esc(x.unit)}</td><td class="num">${fmt(x.unitPrice)}</td><td class="num">${fmt(x.qty*x.unitPrice)}</td></tr>`).join('');$('#invoiceDocument').innerHTML=`<div class="letterhead"><h2>UNIVERSE GEMS &amp; JEWELLERY CO.</h2><p>UNIT 11-12, 10/F., FU HANG INDUSTRIAL BUILDING, NO. 1 HOK YUEN STREET EAST,<br>HUNG HOM, KOWLOON, HONG KONG · TEL : (852) 2363 5409 · FAX : (852) 2765 0343</p></div><div class="doc-title">Sales Invoice</div><div class="doc-grid"><div class="doc-meta">No. : <strong>${esc($('#invoiceNo').value)}</strong><br>Invoice Date : ${esc($('#invoiceDate').value)}<br>Shipment Method : ${esc($('#shipmentMethod').value)}<br>Currency : ${esc($('#currency').value)}<br><br>Customer : <strong>${esc($('#customerName').value)}</strong><br>${esc($('#customerAddress').value).replace(/\n/g,'<br>')}</div><div class="doc-meta"><strong>Vender's Banker</strong><br>The Hong Kong &amp; Shanghai Banking Corporation Ltd.<br>Address : 41 Ma Tau Wai Road,Hung Hom,Kowloon,Hong Kong<br>A/C # : 012-593570-001<br>A/C Name : Universe Gems &amp; Jewellery Co.</div></div><table class="doc-table"><thead><tr><th>No.</th><th>Article No.</th><th>Description</th><th>Quantity</th><th>Unit</th><th class="num">Unit Price</th><th class="num">Amount</th></tr><tr><th colspan="7">F.O.B. Value</th></tr></thead><tbody>${rows}</tbody></table><div class="doc-footer"><div class="doc-totals"><div><span>Total Quantity :</span><strong>${t.qty}</strong></div><div><span>Sub Total:</span><strong>${fmt(t.sub)}</strong></div><div><span>Discount:</span><strong>${fmt(t.discount)}</strong></div><div class="total"><span>Total : (${esc($('#currency').value)})</span><strong>${fmt(t.total)}</strong></div></div><p><strong>Remark :</strong> ${esc($('#remark').value)}</p></div>`}

function setExcelExportStatus(message,type=''){
  const el=$('#excelExportStatus');
  if(!el)return;
  el.textContent=message;
  el.className='notice'+(type?' '+type:'');
}
function downloadBlob(blob,fileName){
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;a.download=fileName;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),3000);
}
async function imageFileToJpegDataUrl(file,maxSide=620,quality=.82){
  if(!file)return null;
  const source=URL.createObjectURL(file);
  try{
    const img=await new Promise((resolve,reject)=>{const i=new Image();i.onload=()=>resolve(i);i.onerror=reject;i.src=source});
    const scale=Math.min(1,maxSide/Math.max(img.naturalWidth||img.width,img.naturalHeight||img.height));
    const w=Math.max(1,Math.round((img.naturalWidth||img.width)*scale));
    const h=Math.max(1,Math.round((img.naturalHeight||img.height)*scale));
    const canvas=document.createElement('canvas');canvas.width=w;canvas.height=h;
    const ctx=canvas.getContext('2d');ctx.fillStyle='#ffffff';ctx.fillRect(0,0,w,h);ctx.drawImage(img,0,0,w,h);
    return canvas.toDataURL('image/jpeg',quality);
  }finally{URL.revokeObjectURL(source)}
}
async function imageFileToJpegAsset(file,maxSide=700,quality=.84){
  if(!file)return null;
  const source=URL.createObjectURL(file);
  try{
    const img=await new Promise((resolve,reject)=>{const i=new Image();i.onload=()=>resolve(i);i.onerror=reject;i.src=source});
    const scale=Math.min(1,maxSide/Math.max(img.naturalWidth||img.width,img.naturalHeight||img.height));
    const width=Math.max(1,Math.round((img.naturalWidth||img.width)*scale));
    const height=Math.max(1,Math.round((img.naturalHeight||img.height)*scale));
    const canvas=document.createElement('canvas');canvas.width=width;canvas.height=height;
    const ctx=canvas.getContext('2d');ctx.fillStyle='#ffffff';ctx.fillRect(0,0,width,height);ctx.drawImage(img,0,0,width,height);
    return {base64:canvas.toDataURL('image/jpeg',quality),width,height};
  }finally{URL.revokeObjectURL(source)}
}
function applyThinBorder(cell){cell.border={top:{style:'thin',color:{argb:'FFD1D5DB'}},left:{style:'thin',color:{argb:'FFD1D5DB'}},bottom:{style:'thin',color:{argb:'FFD1D5DB'}},right:{style:'thin',color:{argb:'FFD1D5DB'}}}}

function cloneStyle(v){try{return JSON.parse(JSON.stringify(v||{}))}catch{return v||{}}}
function excelColNumber(letter){let n=0;for(const ch of String(letter).toUpperCase())n=n*26+(ch.charCodeAt(0)-64);return n}
function excelColPixels(ws,colNo){const width=Number(ws.getColumn(colNo).width)||8.43;return Math.max(12,Math.round(width*7+5))}
function excelRowPixels(ws,rowNo){const points=Number(ws.getRow(rowNo).height)||15;return Math.max(8,points*96/72)}
function imageAnchorCol(ws,startColNo,endColNo,offsetPx){let col=startColNo-1,remain=Math.max(0,offsetPx);for(let c=startColNo;c<=endColNo;c++){const px=excelColPixels(ws,c);if(remain<=px)return col+remain/px;remain-=px;col+=1}return endColNo}
function imageAnchorRow(ws,startRow,endRow,offsetPx){let row=startRow-1,remain=Math.max(0,offsetPx);for(let r=startRow;r<=endRow;r++){const px=excelRowPixels(ws,r);if(remain<=px)return row+remain/px;remain-=px;row+=1}return endRow}
function rowRangeHeightPoints(ws,start,end){let total=0;for(let r=start;r<=end;r++)total+=Number(ws.getRow(r).height)||15;return total}
function copyTemplateRowStyle(ws,sourceRow,targetRow){
  const src=ws.getRow(sourceRow),dst=ws.getRow(targetRow);dst.height=src.height;
  for(let c=1;c<=12;c++){
    const s=src.getCell(c),d=dst.getCell(c);
    d.style=cloneStyle(s.style);d.numFmt=s.numFmt;d.alignment=cloneStyle(s.alignment);d.border=cloneStyle(s.border);d.fill=cloneStyle(s.fill);d.font=cloneStyle(s.font);d.protection=cloneStyle(s.protection);
  }
}
async function exportInvoiceFromTemplate(){
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.load(state.invoiceTemplateBuffer.slice(0));

  const mapWs=wb.getWorksheet('Template Map');
  const ws=wb.getWorksheet('Invoice Template')||wb.worksheets.find(s=>s.name!=='Template Map')||wb.worksheets[0];
  if(!ws)throw new Error('範本沒有 Invoice 工作表');

  const map=new Map();
  if(mapWs){
    mapWs.eachRow((row,rowNo)=>{
      if(rowNo===1)return;
      const field=norm(row.getCell(2).value);
      const target=norm(row.getCell(3).value);
      const format=norm(row.getCell(5).value);
      const notes=norm(row.getCell(6).value);
      if(field&&target)map.set(field.toLowerCase(),{field,target,format,notes});
    });
  }
  const getMap=(name,fallback='')=>map.get(name.toLowerCase())?.target||fallback;
  const mapFormat=(name)=>map.get(name.toLowerCase())?.format||'';

  // Header fields are now controlled by the Template Map sheet.
  const setMapped=(field,value)=>{
    const address=getMap(field);
    if(!address)return;
    const cell=ws.getCell(address.split(':')[0]);
    cell.value=value;
    const fmt=mapFormat(field).toLowerCase();
    if(fmt.includes('date'))cell.numFmt='d mmmm, yyyy';
  };
  const inv=norm($('#invoiceNo').value)||formatInvoiceNo();
  const invoiceDateText=norm($('#invoiceDate').value);
  const invoiceDate=invoiceDateText?new Date(`${invoiceDateText}T00:00:00`):new Date();
  setMapped('Invoice No.',inv);
  setMapped('Invoice Date',invoiceDate);
  setMapped('Shipment Method',norm($('#shipmentMethod').value));
  setMapped('Currency',norm($('#currency').value)||'USD');
  setMapped('Company',norm($('#customerName').value));
  setMapped('Customer Code',norm($('#customerCode').value));
  setMapped('Payment Term',norm($('#customerTerms').value));

  const addressLines=norm($('#customerAddress').value).split(/\r?\n/).map(norm).filter(Boolean);
  setMapped('Ship To',addressLines[0]||'');
  setMapped('Country',addressLines.slice(1).join(', '));

  // Locate the item table and footer from the map / visible labels.
  let headerRow=0;
  for(let r=1;r<=Math.min(ws.rowCount,80);r++){
    if(norm(ws.getCell(`A${r}`).value).toLowerCase()==='no.') {headerRow=r;break}
  }
  const firstItemRow=headerRow?headerRow+2:19;
  const footerQtyCell=getMap('Total Quantity','F25').split(':')[0];
  const footerBaseRow=Number((footerQtyCell.match(/\d+/)||['25'])[0]);
  const originalFooterEnd=ws.rowCount;
  const originalContentRows=Math.max(1,footerBaseRow-firstItemRow-1);
  const baseContentRows=Math.min(5,originalContentRows);
  const separatorSourceRow=firstItemRow+baseContentRows;
  const columnCount=Math.max(9,ws.columnCount||9);

  const captureCell=(cell)=>({
    value:cell.value,style:cloneStyle(cell.style),numFmt:cell.numFmt,
    alignment:cloneStyle(cell.alignment),border:cloneStyle(cell.border),
    fill:cloneStyle(cell.fill),font:cloneStyle(cell.font),protection:cloneStyle(cell.protection)
  });
  const contentStyle=[];
  for(let c=1;c<=columnCount;c++)contentStyle.push(captureCell(ws.getRow(firstItemRow).getCell(c)));
  const separatorStyle=[];
  for(let c=1;c<=columnCount;c++)separatorStyle.push(captureCell(ws.getRow(separatorSourceRow).getCell(c)));
  const contentHeight=ws.getRow(firstItemRow).height||18;
  const separatorHeight=ws.getRow(separatorSourceRow).height||8;

  const footerRows=[];
  for(let r=footerBaseRow;r<=originalFooterEnd;r++){
    const row=[];
    for(let c=1;c<=columnCount;c++)row.push(captureCell(ws.getRow(r).getCell(c)));
    footerRows.push({height:ws.getRow(r).height,row});
  }

  // Preserve all drawings already embedded in the imported template, including the company letterhead.
  for(let r=firstItemRow;r<=originalFooterEnd;r++){
    try{ws.unMergeCells(`D${r}:E${r+4}`)}catch{}
  }

  const itemPlans=state.items.map(item=>{
    const lines=[articleDescriptionFor(item),...(item.descriptions||[])].map(norm).filter(Boolean);
    const contentRows=Math.max(5,lines.length);
    return {item,lines,contentRows,totalRows:contentRows+1};
  });
  const totalItemRows=itemPlans.reduce((s,x)=>s+x.totalRows,0);
  const footerStart=firstItemRow+totalItemRows;
  const requiredEnd=footerStart+footerRows.length-1;
  const clearEnd=Math.max(originalFooterEnd,requiredEnd+2);

  for(let r=firstItemRow;r<=clearEnd;r++){
    for(let c=1;c<=columnCount;c++)ws.getRow(r).getCell(c).value=null;
    ws.getRow(r).height=undefined;
  }

  const applyCaptured=(cell,src,includeValue=false)=>{
    cell.style=cloneStyle(src.style);cell.numFmt=src.numFmt;
    cell.alignment=cloneStyle(src.alignment);cell.border=cloneStyle(src.border);
    cell.fill=cloneStyle(src.fill);cell.font=cloneStyle(src.font);
    cell.protection=cloneStyle(src.protection);
    if(includeValue)cell.value=src.value;
  };
  const applyRowStyle=(rowNo,styleRow,height)=>{
    const row=ws.getRow(rowNo);row.height=height;
    for(let c=1;c<=columnCount;c++)applyCaptured(row.getCell(c),styleRow[c-1],false);
  };
  const colLetter=(field,fallback)=>{
    const target=getMap(field,fallback);
    const m=target.match(/[A-Z]+/i);return (m?m[0]:fallback).toUpperCase();
  };
  const noCol='A';
  const lotCol=colLetter('Lot No.','B');
  const artCol=colLetter('ARTNO',lotCol);
  const descCol=colLetter('Article','C');
  const imageTarget=getMap('Product Image','D:E merged item block');
  const imageMatch=imageTarget.match(/([A-Z]+)\s*:\s*([A-Z]+)/i);
  const imageStartCol=(imageMatch?imageMatch[1]:'D').toUpperCase();
  const imageEndCol=(imageMatch?imageMatch[2]:'E').toUpperCase();
  const qtyCol=colLetter('Qty','F');
  const unitCol=colLetter('Unit','G');
  const unitPriceCol=colLetter('Unit Price','H');
  const amountCol=colLetter('Amount','I');

  let rowCursor=firstItemRow;
  let missingImages=0;
  const pageHeightPts=841.89; // A4 portrait
  const marginTopPts=25.2,marginBottomPts=25.2;
  const repeatedHeaderPts=rowRangeHeightPoints(ws,1,Math.max(1,firstItemRow-1));
  const pageBodyCapacityPts=Math.max(220,pageHeightPts-marginTopPts-marginBottomPts-repeatedHeaderPts);
  let pageUsedPts=0;
  let pageItemCount=0;
  const maxItemsPerPage=10;

  for(let i=0;i<itemPlans.length;i++){
    const {item,lines,contentRows,totalRows}=itemPlans[i];
    const start=rowCursor,contentEnd=start+contentRows-1,separatorRow=contentEnd+1;
    const itemHeightPts=contentRows*contentHeight+separatorHeight;
    const pageIsFullByCount=pageItemCount>=maxItemsPerPage;
    const pageIsFullByHeight=pageUsedPts>0&&pageUsedPts+itemHeightPts>pageBodyCapacityPts;
    if(pageUsedPts>0&&(pageIsFullByCount||pageIsFullByHeight)){
      try{ws.getRow(start).addPageBreak()}catch{}
      pageUsedPts=0;
      pageItemCount=0;
    }
    for(let r=start;r<=contentEnd;r++)applyRowStyle(r,contentStyle,contentHeight);
    applyRowStyle(separatorRow,separatorStyle,separatorHeight);

    ws.getCell(`${noCol}${start}`).value=i+1;
    ws.getCell(`${noCol}${start}`).alignment={...cloneStyle(ws.getCell(`${noCol}${start}`).alignment),horizontal:'center',vertical:'middle'};
    ws.getCell(`${lotCol}${start}`).value=`Lot.No. : ${item.lotNo}`;
    ws.getCell(`${artCol}${start+1}`).value=item.artNo;
    ws.getCell(`${lotCol}${start}`).font={...cloneStyle(ws.getCell(`${lotCol}${start}`).font),bold:true};
    ws.getCell(`${artCol}${start+1}`).font={...cloneStyle(ws.getCell(`${artCol}${start+1}`).font),bold:true};

    for(let r=0;r<contentRows;r++){
      const cell=ws.getCell(`${descCol}${start+r}`);
      cell.value=lines[r]||'';
      cell.alignment={...cloneStyle(cell.alignment),vertical:'middle',wrapText:false};
    }

    try{ws.mergeCells(`${imageStartCol}${start}:${imageEndCol}${start+4}`)}catch{}
    ws.getCell(`${imageStartCol}${start}`).value=null;
    ws.getCell(`${imageStartCol}${start}`).alignment={horizontal:'center',vertical:'middle'};

    ws.getCell(`${qtyCol}${start}`).value=item.qty;
    ws.getCell(`${qtyCol}${start}`).numFmt='0';
    ws.getCell(`${unitCol}${start}`).value=item.unit;
    ws.getCell(`${unitPriceCol}${start}`).value=item.unitPrice;
    ws.getCell(`${unitPriceCol}${start}`).numFmt='$#,##0.00';
    ws.getCell(`${amountCol}${start}`).value=item.qty*item.unitPrice;
    ws.getCell(`${amountCol}${start}`).numFmt='$#,##0.00';

    const selected=getImg(item);
    if(selected?.file){
      try{
        const asset=await imageFileToJpegAsset(selected.file,620,.84);
        const imageId=wb.addImage({base64:asset.base64,extension:'jpeg'});
        const imageStartColNo=excelColNumber(imageStartCol),imageEndColNo=excelColNumber(imageEndCol);
        const imageEndRow=start+4;
        let boxW=0,boxH=0;
        for(let c=imageStartColNo;c<=imageEndColNo;c++)boxW+=excelColPixels(ws,c);
        for(let r=start;r<=imageEndRow;r++)boxH+=excelRowPixels(ws,r);
        const pad=6,maxW=Math.max(20,boxW-pad*2),maxH=Math.max(20,boxH-pad*2);
        const scale=Math.min(maxW/asset.width,maxH/asset.height,1);
        const width=Math.max(20,Math.round(asset.width*scale));
        const height=Math.max(20,Math.round(asset.height*scale));
        const xOffset=(boxW-width)/2,yOffset=(boxH-height)/2;
        ws.addImage(imageId,{tl:{col:imageAnchorCol(ws,imageStartColNo,imageEndColNo,xOffset),row:imageAnchorRow(ws,start,imageEndRow,yOffset)},ext:{width,height},editAs:'oneCell'});
      }catch{missingImages++}
    }else missingImages++;

    rowCursor+=totalRows;pageUsedPts+=itemHeightPts;pageItemCount+=1;
    setExcelExportStatus(`正在依 Template Map 建立 Excel… ${i+1}/${state.items.length}`);
  }

  for(let offset=0;offset<footerRows.length;offset++){
    const targetRow=footerStart+offset,captured=footerRows[offset],row=ws.getRow(targetRow);
    row.height=captured.height;
    for(let c=1;c<=columnCount;c++)applyCaptured(row.getCell(c),captured.row[c-1],true);
  }

  const t=totals();
  const shiftedAddress=(field,fallback)=>{
    const addr=getMap(field,fallback).split(':')[0];
    const m=addr.match(/^([A-Z]+)(\d+)$/i);if(!m)return fallback;
    return `${m[1].toUpperCase()}${footerStart+(Number(m[2])-footerBaseRow)}`;
  };
  const totalQtyAddr=shiftedAddress('Total Quantity','F25');
  const subAddr=shiftedAddress('Sub Total','I25');
  const discountAddr=shiftedAddress('Discount','I28');
  const totalAddr=shiftedAddress('Total','I30');
  ws.getCell(totalQtyAddr).value=t.qty;ws.getCell(totalQtyAddr).numFmt='0';
  ws.getCell(subAddr).value=t.sub;ws.getCell(subAddr).numFmt='$#,##0.00';
  ws.getCell(discountAddr).value=t.discount;ws.getCell(discountAddr).numFmt='$#,##0.00';
  ws.getCell(totalAddr).value=t.total;ws.getCell(totalAddr).numFmt='$#,##0.00';

  // Fill text fields in the footer by label, so changing rows in the template remains safe.
  const findLabelRow=(text)=>{
    const needle=text.toLowerCase();
    for(let r=footerStart;r<=requiredEnd;r++)for(let c=1;c<=columnCount;c++){
      if(norm(ws.getRow(r).getCell(c).value).toLowerCase().includes(needle))return {r,c};
    }
    return null;
  };
  const amountLabel=findLabelRow('total amount');
  if(amountLabel){
    const amountCell=ws.getRow(amountLabel.r).getCell(Math.min(columnCount,amountLabel.c+1));
    amountCell.value=`${currencyWords($('#currency').value)} ${numberToWords(t.total)}`;
    amountCell.alignment={...cloneStyle(amountCell.alignment),vertical:'middle',wrapText:false};
  }
  const remarkLabel=findLabelRow('remark');
  if(remarkLabel)ws.getRow(remarkLabel.r).getCell(Math.min(columnCount,remarkLabel.c+1)).value=norm($('#remark').value);

  const footerHeightPts=footerRows.reduce((sum,x)=>sum+(Number(x.height)||15),0);
  if(pageUsedPts>0&&pageUsedPts+footerHeightPts>pageBodyCapacityPts){try{ws.getRow(footerStart).addPageBreak()}catch{}}
  // Uniform alignment requested for the complete Invoice sheet.
  for(let r=1;r<=requiredEnd;r++)for(let c=1;c<=columnCount;c++){
    const cell=ws.getRow(r).getCell(c);
    cell.alignment={...cloneStyle(cell.alignment),vertical:'middle',wrapText:false};
  }
  ws.pageSetup=ws.pageSetup||{};
  ws.pageSetup.paperSize=9;ws.pageSetup.orientation='portrait';ws.pageSetup.fitToPage=true;
  ws.pageSetup.fitToWidth=1;ws.pageSetup.fitToHeight=0;ws.pageSetup.printArea=`A1:I${requiredEnd}`;
  ws.pageSetup.scale=undefined;
  ws.pageSetup.printTitlesRow=`1:${Math.max(1,firstItemRow-1)}`;
  ws.pageSetup.horizontalCentered=true;ws.pageSetup.verticalCentered=false;
  ws.pageSetup.margins={left:.35,right:.35,top:.35,bottom:.35,header:.15,footer:.15};
  ws.headerFooter=ws.headerFooter||{};ws.headerFooter.oddFooter='Page &P of &N';

  if(mapWs)wb.removeWorksheet(mapWs.id);
  ws.name='Invoice';
  const buffer=await wb.xlsx.writeBuffer();
  downloadBlob(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),`${inv}.xlsx`);
  setExcelExportStatus(`已依 Template Map 輸出 ${state.invoiceTemplateName}${missingImages?`；${missingImages} 款沒有圖片`:''}。`,'ok');
}
async function exportInvoiceExcel(){
  if(!state.items.length){alert('Invoice 沒有貨品。');return}
  if(typeof ExcelJS==='undefined'){setExcelExportStatus('Excel 輸出程式未載入，請連接網絡後重新開啟。','error');return}
  const btn=$('#exportExcelBtn');btn.disabled=true;setExcelExportStatus('正在建立 Excel Invoice…');
  try{
    if(state.invoiceTemplateBuffer){await exportInvoiceFromTemplate();return}
    const wb=new ExcelJS.Workbook();
    wb.creator='Universe Invoice PWA';wb.created=new Date();
    const ws=wb.addWorksheet('Sales Invoice',{pageSetup:{paperSize:9,orientation:'portrait',fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.25,right:.25,top:.35,bottom:.35,header:.15,footer:.15}}});
    ws.views=[{showGridLines:false}];
    ws.columns=[
      {key:'no',width:5},{key:'image',width:15},{key:'article',width:17},{key:'description',width:34},
      {key:'qty',width:9},{key:'unit',width:8},{key:'unitPrice',width:14},{key:'amount',width:14}
    ];
    const merge=(range,value,size=10,bold=false,align='left')=>{ws.mergeCells(range);const c=ws.getCell(range.split(':')[0]);c.value=value;c.font={name:'Arial',size,bold};c.alignment={vertical:'middle',horizontal:align,wrapText:true};return c};
    merge('A1:H1','UNIVERSE GEMS & JEWELLERY CO.',17,true,'center');ws.getRow(1).height=24;
    merge('A2:H2','UNIT 11-12, 10/F., FU HANG INDUSTRIAL BUILDING,',9,false,'center');
    merge('A3:H3','NO. 1 HOK YUEN STREET EAST, HUNG HOM, KOWLOON, HONG KONG',9,false,'center');
    merge('A4:H4','TEL : (852) 2363 5409     FAX : (852) 2765 0343',9,false,'center');
    ws.getRow(5).height=7;
    merge('A6:D6','Sales Invoice',16,true,'left');
    merge('E6:H6',`No. : ${norm($('#invoiceNo').value)}`,11,true,'right');
    merge('A7:D7',`Invoice Date : ${norm($('#invoiceDate').value)}`,10);
    merge('E7:H7',`Currency : ${norm($('#currency').value)}`,10,false,'right');
    merge('A8:D8',`Shipment Method : ${norm($('#shipmentMethod').value)}`,10);
    merge('E8:H8',`Customer Code : ${norm($('#customerCode').value)}`,10,false,'right');
    merge('A9:D11',`Customer : ${norm($('#customerName').value)}\n${norm($('#customerAddress').value)}`,10,true);
    merge('E9:H11',"Vender's Banker\nThe Hong Kong & Shanghai Banking Corporation Ltd.\nAddress : 41 Ma Tau Wai Road, Hung Hom, Kowloon, Hong Kong\nA/C # : 012-593570-001\nA/C Name : Universe Gems & Jewellery Co.",9,false,'left');
    [9,10,11].forEach(r=>ws.getRow(r).height=20);
    const headerRow=13;
    const headers=['No.','Picture','Article No.','Description','Quantity','Unit','Unit Price','Amount'];
    headers.forEach((h,i)=>{const c=ws.getCell(headerRow,i+1);c.value=h;c.font={name:'Arial',size:10,bold:true};c.alignment={horizontal:'center',vertical:'middle',wrapText:true};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE5E7EB'}};applyThinBorder(c)});
    ws.getRow(headerRow).height=24;
    ws.mergeCells(`A${headerRow+1}:H${headerRow+1}`);const fob=ws.getCell(headerRow+1,1);fob.value='F.O.B. Value';fob.font={name:'Arial',size:10,bold:true};fob.alignment={vertical:'middle'};applyThinBorder(fob);ws.getRow(headerRow+1).height=21;
    let row=headerRow+2;
    let missingImages=0;
    for(let i=0;i<state.items.length;i++){
      const item=state.items[i],start=row,end=row+5;
      for(let r=start;r<=end;r++)ws.getRow(r).height=18;
      ws.mergeCells(`A${start}:A${end}`);ws.getCell(`A${start}`).value=i+1;
      ws.getCell(`A${start}`).alignment={horizontal:'center',vertical:'middle'};
      ws.mergeCells(`B${start}:B${end}`);
      ws.mergeCells(`C${start}:C${end}`);ws.getCell(`C${start}`).value=`Lot.No. : ${item.lotNo}\n${item.artNo}`;ws.getCell(`C${start}`).font={name:'Arial',size:10,bold:true};ws.getCell(`C${start}`).alignment={vertical:'top',wrapText:true};
      ws.mergeCells(`D${start}:D${end}`);ws.getCell(`D${start}`).value=[articleDescriptionFor(item),...item.descriptions].filter(Boolean).join('\n');ws.getCell(`D${start}`).alignment={vertical:'top',wrapText:true};ws.getCell(`D${start}`).font={name:'Arial',size:10};
      ws.mergeCells(`E${start}:E${end}`);ws.getCell(`E${start}`).value=item.qty;ws.getCell(`E${start}`).alignment={horizontal:'center',vertical:'middle'};
      ws.mergeCells(`F${start}:F${end}`);ws.getCell(`F${start}`).value=item.unit;ws.getCell(`F${start}`).alignment={horizontal:'center',vertical:'middle'};
      ws.mergeCells(`G${start}:G${end}`);ws.getCell(`G${start}`).value=item.unitPrice;ws.getCell(`G${start}`).numFmt='$#,##0.00';ws.getCell(`G${start}`).alignment={horizontal:'right',vertical:'middle'};
      ws.mergeCells(`H${start}:H${end}`);ws.getCell(`H${start}`).value={formula:`E${start}*G${start}`,result:item.qty*item.unitPrice};ws.getCell(`H${start}`).numFmt='$#,##0.00';ws.getCell(`H${start}`).alignment={horizontal:'right',vertical:'middle'};
      for(let r=start;r<=end;r++)for(let c=1;c<=8;c++)applyThinBorder(ws.getCell(r,c));
      const selected=getImg(item);
      if(selected?.file){
        try{
          const dataUrl=await imageFileToJpegDataUrl(selected.file);
          const imageId=wb.addImage({base64:dataUrl,extension:'jpeg'});
          ws.addImage(imageId,{tl:{col:1.08,row:start-1+.12},br:{col:1.92,row:end-.12},editAs:'oneCell'});
        }catch{missingImages++}
      }else missingImages++;
      row=end+1;
      setExcelExportStatus(`正在建立 Excel Invoice… ${i+1}/${state.items.length}`);
    }
    const t=totals();
    ws.mergeCells(`A${row}:F${row}`);ws.getCell(`A${row}`).value='Total Quantity';ws.getCell(`A${row}`).font={bold:true};ws.getCell(`G${row}`).value=t.qty;ws.getCell(`G${row}`).font={bold:true};ws.getCell(`G${row}`).alignment={horizontal:'right'};
    row++;
    ws.mergeCells(`A${row}:F${row}`);ws.getCell(`A${row}`).value='Sub Total';ws.getCell(`A${row}`).font={bold:true};ws.getCell(`G${row}`).value=t.sub;ws.getCell(`G${row}`).numFmt='$#,##0.00';ws.getCell(`G${row}`).font={bold:true};ws.getCell(`G${row}`).alignment={horizontal:'right'};
    row++;
    ws.mergeCells(`A${row}:F${row}`);ws.getCell(`A${row}`).value='Discount Amount';ws.getCell(`G${row}`).value=t.discount;ws.getCell(`G${row}`).numFmt='$#,##0.00';ws.getCell(`G${row}`).alignment={horizontal:'right'};
    row++;
    ws.mergeCells(`A${row}:F${row}`);ws.getCell(`A${row}`).value=`Total : (${norm($('#currency').value)})`;ws.getCell(`A${row}`).font={bold:true,size:12};ws.getCell(`G${row}`).value=t.total;ws.getCell(`G${row}`).numFmt='$#,##0.00';ws.getCell(`G${row}`).font={bold:true,size:12};ws.getCell(`G${row}`).alignment={horizontal:'right'};
    row+=2;
    merge(`A${row}:H${row}`,`Remark : ${norm($('#remark').value)}`,10);
    row+=2;merge(`A${row}:D${row}`,'Vender Signature : ______________________',10);merge(`E${row}:H${row}`,'Accept By : ______________________',10,false,'right');
    ws.headerFooter.oddFooter='Page &P of &N';
    ws.pageSetup.printArea=`A1:H${row}`;
    ws.autoFilter={from:{row:headerRow,column:1},to:{row:headerRow,column:8}};
    const buffer=await wb.xlsx.writeBuffer();
    const inv=norm($('#invoiceNo').value)||formatInvoiceNo();
    downloadBlob(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),`${inv}.xlsx`);
    setExcelExportStatus(`Excel Invoice 已輸出${missingImages?`；${missingImages} 款沒有嵌入圖片`:''}。`,'ok');
  }catch(err){console.error(err);setExcelExportStatus('Excel 輸出失敗：'+(err.message||err),'error')}
  finally{btn.disabled=false}
}

$('#refreshPreviewBtn').onclick=renderPreview;$('#exportExcelBtn').onclick=exportInvoiceExcel;
function updateZoomButtons(){
  $$('.zoom-btn').forEach(btn=>{
    const z=Number(btn.dataset.zoom);
    const supported=z>=state.scannerZoom.min-0.001&&z<=state.scannerZoom.max+0.001;
    btn.disabled=!supported||!state.scannerRunning;
    btn.classList.toggle('active',Math.abs(z-state.scannerZoom.current)<0.05);
  });
}
async function setScannerZoom(requested){
  if(!state.scannerRunning||!state.scanner)return;
  const min=state.scannerZoom.min,max=state.scannerZoom.max,step=state.scannerZoom.step||0.1;
  const clamped=Math.max(min,Math.min(max,requested));
  const zoom=Math.round(clamped/step)*step;
  try{
    await state.scanner.applyVideoConstraints({advanced:[{zoom}]});
    const applied=Number(state.scanner.getRunningTrackSettings?.().zoom);
    state.scannerZoom.current=Number.isFinite(applied)?applied:zoom;
    updateZoomButtons();
    const shown=state.scannerZoom.current;
    $('#scannerStatus').textContent=`後鏡頭 ${shown.toFixed(shown%1?1:0)}×，請把 Barcode 放在掃描框內。`;
  }catch(err){
    $('#scannerStatus').textContent='此倍率未能套用：'+(err.message||err);
  }
}
$$('.zoom-btn').forEach(btn=>btn.onclick=()=>setScannerZoom(Number(btn.dataset.zoom)));
async function startScanner(){
  if(state.scannerBusy)return;
  state.scannerBusy=true;
  if(!$('#scannerDialog').open)$('#scannerDialog').showModal();
  $('#scannerStatus').textContent='正在啟動後鏡頭…';state.scannerZoom={min:1,max:1,step:1,current:1};updateZoomButtons();
  const config={fps:15,qrbox:(w,h)=>({width:Math.floor(w*.82),height:Math.max(70,Math.floor(h*.18))})};
  const onSuccess=txt=>{if(addLot(txt)){navigator.vibrate?.(80);$('#scannerStatus').textContent='已讀取 '+txt}};
  try{
    if(typeof Html5Qrcode==='undefined')throw new Error('掃描程式未載入');
    if(state.scannerRunning&&state.scanner)await state.scanner.stop();
    state.scannerRunning=false;
    $('#reader').innerHTML='';
    state.scanner=new Html5Qrcode('reader');
    const cams=await Html5Qrcode.getCameras();
    if(!cams.length)throw new Error('找不到相機');
    const rearPattern=/back|rear|environment|後置|背面|後鏡/i;
    const rearCams=cams.filter(x=>rearPattern.test(x.label||''));
    const cam=rearCams[rearCams.length-1]||cams[cams.length-1];
    try{
      await state.scanner.start(cam.id,config,onSuccess,()=>{});
    }catch(firstErr){
      try{await state.scanner.clear()}catch{}
      $('#reader').innerHTML='';
      state.scanner=new Html5Qrcode('reader');
      await state.scanner.start({facingMode:'environment'},config,onSuccess,()=>{});
    }
    state.scannerRunning=true;
    const caps=state.scanner.getRunningTrackCapabilities?.()||{};
    const settings=state.scanner.getRunningTrackSettings?.()||{};
    if(caps.zoom){
      state.scannerZoom={min:Number(caps.zoom.min??1),max:Number(caps.zoom.max??1),step:Number(caps.zoom.step??0.1),current:Number(settings.zoom??caps.zoom.min??1)};
      const preferred=Math.min(2,state.scannerZoom.max);
      if(preferred>state.scannerZoom.current+0.05)await setScannerZoom(preferred);
      else updateZoomButtons();
      $('#scannerStatus').textContent=`已使用後鏡頭；支援 ${state.scannerZoom.min}×–${state.scannerZoom.max}×。`;
    }else{
      state.scannerZoom={min:1,max:1,step:1,current:1};
      updateZoomButtons();
      $('#scannerStatus').textContent='已使用後鏡頭；Safari 未提供相機 Zoom 控制。';
    }
  }catch(err){
    state.scannerRunning=false;
    $('#scannerStatus').textContent='後鏡頭無法啟動：'+(err.message||err);
  }finally{
    state.scannerBusy=false;
  }
}
async function stopScanner(){if(state.scannerBusy)return;state.scannerBusy=true;try{if(state.scannerRunning&&state.scanner)await state.scanner.stop()}catch{}finally{state.scannerRunning=false;$('#reader').innerHTML='';state.scannerZoom={min:1,max:1,step:1,current:1};updateZoomButtons();$('#scannerDialog').close();state.scannerBusy=false}}
$('#scanBtn').onclick=startScanner;$('#closeScannerBtn').onclick=stopScanner;
function exportRemaining(){if(!state.items.length)return alert('Invoice 沒有貨品。');const sold=new Set(state.items.map(x=>x.lotNo));const remain=state.stockRows.filter(r=>!sold.has(norm(field(r,['LOTNO']))));const ws=XLSX.utils.json_to_sheet(remain,{header:state.stockHeaders});const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Remaining Stock');const inv=norm($('#invoiceNo').value)||formatInvoiceNo();XLSX.writeFile(wb,`Remaining_Stock_${inv}_${remain.length}pcs_${today().replaceAll('-','')}.xlsx`);for(const lot of sold)state.products.delete(lot);state.stockRows=remain;state.items=[];advanceInvoiceSequence(inv);renderItems();status('#addMessage',`Invoice ${inv} 已 Confirm，下一張為 ${$('#invoiceNo').value}。`,'ok');status('#stockStatus',`目前 Remaining Stock：${state.products.size} 件。`,'ok')}
$('#confirmInvoiceBtn').onclick=()=>{if(confirm('Confirm Invoice 並匯出 Remaining Stock Excel？'))exportRemaining()};
renderCustomerSummary();renderItems();if('serviceWorker'in navigator)navigator.serviceWorker.register('./sw.js').catch(()=>{});
