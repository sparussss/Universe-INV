const $=(s,r=document)=>r.querySelector(s),$$=(s,r=document)=>[...r.querySelectorAll(s)];
const state={products:new Map(),customers:new Map(),imageFiles:new Map(),items:[],stockRows:[],stockHeaders:[],scanner:null,scannerBusy:false};
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
function field(row,names){const keys=Object.keys(row);for(const n of names){const k=keys.find(x=>x.trim().toUpperCase()===n);if(k)return row[k]}return''}
function fmt(v){return new Intl.NumberFormat('en-US',{style:'currency',currency:$('#currency').value||'USD',minimumFractionDigits:2}).format(Number(v)||0)}
function totals(){const qty=state.items.reduce((a,x)=>a+x.qty,0),sub=state.items.reduce((a,x)=>a+x.qty*x.unitPrice,0),discount=Math.max(0,Number($('#discountAmount').value)||0);return{qty,sub,discount,total:Math.max(0,sub-discount)}}
function updateTotals(){const t=totals();$('#totalQty').textContent=t.qty;$('#subtotal').textContent=fmt(t.sub);$('#discountDisplay').textContent=fmt(t.discount);$('#grandTotal').textContent=fmt(t.total);$('#productCount').textContent=state.products.size;$('#customerCount').textContent=state.customers.size;$('#invoiceCount').textContent=state.items.length;$('#headerTotal').textContent=fmt(t.total)}
$$('.tab').forEach(b=>b.onclick=()=>{$$('.tab').forEach(x=>x.classList.remove('active'));$$('.tab-panel').forEach(x=>x.classList.remove('active'));b.classList.add('active');$('#'+b.dataset.tab).classList.add('active');if(b.dataset.tab==='invoice')renderCustomerSummary();if(b.dataset.tab==='preview')renderPreview()});
async function readWB(file){if(typeof XLSX==='undefined')throw new Error('Excel 程式未載入');return XLSX.read(await file.arrayBuffer(),{type:'array'})}
$('#stockInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const wb=await readWB(f),ws=wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(ws,{defval:''});state.stockRows=rows;state.stockHeaders=Object.keys(rows[0]||{});const map=new Map();for(const r of rows){const lot=norm(field(r,['LOTNO']));const art=normArt(field(r,['ARTNO']));const price=Number(field(r,['PRICE']));if(!lot||!art||!Number.isFinite(price))continue;const desc=[];for(let i=1;i<=6;i++){const v=norm(field(r,[`DESC${i}`]));if(v)desc.push(v)}map.set(lot,{lotNo:lot,artNo:art,price,unit:norm(field(r,['UNIT']))||'PC',descriptions:desc,desc2:norm(field(r,['DESC2']))})}state.products=map;status('#stockStatus',`已匯入 ${f.name}：${map.size} 件貨品。`,'ok');updateTotals()}catch(err){status('#stockStatus','匯入失敗：'+err.message,'error')}};
$('#customerInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const wb=await readWB(f),ws=wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});const map=new Map();for(const r of rows){const code=normCode(r[0]),company=norm(r[1]);if(!code||!company||code.includes('CUSTOMER'))continue;const raw=r[11],num=Number(raw),rate=(raw===''||!Number.isFinite(num))?0.34:num;map.set(code,{code,company,address:[r[2],r[3],r[4]].map(norm).filter(Boolean).join('\n'),rate,terms:norm(r[10])})}state.customers=map;status('#customerStatus',`已匯入 ${f.name}：${map.size} 位客戶。`,'ok');updateTotals()}catch(err){status('#customerStatus','匯入失敗：'+err.message,'error')}};
function parseImage(file){const stem=file.name.replace(/\.[^.]+$/,'').trim();const arts=[...new Set([...state.products.values()].map(x=>x.artNo))].sort((a,b)=>b.length-a.length);const art=arts.find(a=>stem.toUpperCase()===a||stem.toUpperCase().startsWith(a+' '));if(!art)return null;let variant=stem.slice(art.length).trim().replace(/\s*\(\d+\)$/,'').trim()||'Default';const dup=(stem.match(/\((\d+)\)$/)||[])[1];return{art,variant,dup:dup?Number(dup):0,file}}
$('#imageFolderInput').onchange=e=>{const map=new Map();for(const f of e.target.files){const p=parseImage(f);if(!p)continue;const key=p.art+'|'+p.variant.toUpperCase(),existing=map.get(key);if(!existing||p.dup<existing.dup)map.set(key,p)}state.imageFiles=new Map();for(const p of map.values()){const arr=state.imageFiles.get(p.art)||[];arr.push({variant:p.variant,url:URL.createObjectURL(p.file),fileName:p.file.name});state.imageFiles.set(p.art,arr)}for(const arr of state.imageFiles.values())arr.sort((a,b)=>a.variant==='Default'?-1:b.variant==='Default'?1:a.variant.localeCompare(b.variant));status('#imageStatus',`已選擇圖片 Folder：${e.target.files.length} 張圖片，配對 ${state.imageFiles.size} 個款號。`,'ok');renderItems()};
function searchCustomers(q){const s=norm(q).toUpperCase(),c=normCode(q);return [...state.customers.values()].filter(x=>x.code.includes(c)||x.company.toUpperCase().includes(s)).slice(0,15)}
function showMatches(){const box=$('#customerMatches'),m=searchCustomers($('#customerSearch').value);box.innerHTML='';if(!m.length){box.innerHTML='<div class="notice">找不到客戶。</div>';return}m.forEach(c=>{const b=document.createElement('button');b.className='customer-match';b.innerHTML=`<span><strong>${esc(c.code)} · ${esc(c.company)}</strong><small>${esc(c.address).replace(/\n/g,' · ')}</small></span><span>${c.rate}</span>`;b.onclick=()=>selectCustomer(c);box.appendChild(b)})}
function selectCustomer(c){$('#customerCode').value=c.code;$('#customerName').value=c.company;$('#customerAddress').value=c.address;$('#salesRate').value=c.rate;$('#customerTerms').value=c.terms;$('#customerMatches').innerHTML='';$('#customerSearch').value='';reprice();renderCustomerSummary()}
$('#findCustomerBtn').onclick=showMatches;$('#customerSearch').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();showMatches()}};
function renderCustomerSummary(){const code=$('#customerCode').value,name=$('#customerName').value;$('#selectedCustomerSummary').innerHTML=code||name?`<strong>${esc(code)} · ${esc(name)}</strong><span>Sales Rate ${esc($('#salesRate').value)} · ${esc($('#currency').value)}</span>`:'尚未選擇客戶。'}
function chooseVariant(p){const imgs=state.imageFiles.get(p.artNo)||[];if(!imgs.length)return'Default';const d=(p.desc2||'').toUpperCase();const aliases=[['QAM','AM'],['YCT','CT'],['BTO','BT'],['MG','MG'],['AQ','AQ'],['PAM','P.AM'],['PTQ','PTR'],['LBT','L.BT'],['SKY BTO','SKY BT']];for(const [code,v] of aliases)if(d.includes(code)){const hit=imgs.find(x=>x.variant.toUpperCase()===v.toUpperCase());if(hit)return hit.variant}return imgs.find(x=>x.variant==='Default')?.variant||imgs[0].variant}
function normalizeLotInput(raw){
  const map={'零':'0','〇':'0','一':'1','二':'2','兩':'2','两':'2','三':'3','四':'4','五':'5','六':'6','七':'7','八':'8','九':'9'};
  return String(raw??'')
    .replace(/[零〇一二兩两三四五六七八九]/g,ch=>map[ch]||ch)
    .replace(/[^0-9A-Za-z]/g,'')
    .toUpperCase();
}
function addLot(raw){const lot=normalizeLotInput(raw);if(!lot)return status('#addMessage','請輸入 LOTNO。','error');const p=state.products.get(lot);if(!p)return status('#addMessage',`找不到 LOTNO ${lot}。`,'error');if(state.items.some(x=>x.lotNo===lot))return status('#addMessage',`LOTNO ${lot} 已在 Invoice。`,'error');const rate=Number($('#salesRate').value)||0;state.items.push({...p,id:Date.now()+Math.random(),seq:state.items.length+1,qty:1,unitPrice:Math.ceil(p.price*rate),imageVariant:chooseVariant(p)});$('#lotInput').value='';status('#addMessage',`已加入 ${p.artNo} / LOTNO ${lot}`,'ok');renderItems();setTimeout(()=>$('#invoiceItems').scrollTo({top:0,behavior:'smooth'}),50)}
$('#addLotBtn').onclick=()=>addLot($('#lotInput').value);
$('#lotInput').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();addLot(e.target.value)}};
$('#lotInput').oninput=e=>{e.target.value=normalizeLotInput(e.target.value)};
$('#lotInput').onfocus=e=>setTimeout(()=>e.target.select(),50);
function getImg(item){const arr=state.imageFiles.get(item.artNo)||[];return arr.find(x=>x.variant===item.imageVariant)||arr[0]}
function placeholder(t){return `data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="200" height="200"><rect width="100%" height="100%" fill="#f1f5f9"/><text x="50%" y="50%" text-anchor="middle" dominant-baseline="middle" font-family="Arial" font-size="18" fill="#64748b">${t}</text></svg>`)}`}
function renderItems(){const box=$('#invoiceItems');box.innerHTML='';if(!state.items.length){box.className='invoice-items empty-state';box.textContent='尚未加入貨品。';updateTotals();return}box.className='invoice-items';[...state.items].reverse().forEach(item=>{const node=$('#itemTemplate').content.firstElementChild.cloneNode(true);$('.item-seq',node).textContent=item.seq;$('.item-artno',node).textContent=item.artNo;$('.item-lot',node).textContent=`LOTNO ${item.lotNo}`;$('.item-desc',node).textContent=item.descriptions.join('\n');$('.item-price-note',node).textContent=`${item.price}u × ${Number($('#salesRate').value)||0} → ${fmt(item.unitPrice)}`;$('.item-thumb',node).src=getImg(item)?.url||placeholder(item.artNo);const sel=$('.variant-select',node),arr=state.imageFiles.get(item.artNo)||[];if(arr.length){arr.forEach(x=>{const o=document.createElement('option');o.value=x.variant;o.textContent='圖片：'+x.variant;o.selected=x.variant===item.imageVariant;sel.appendChild(o)});sel.onchange=e=>{item.imageVariant=e.target.value;renderItems()}}else{sel.innerHTML='<option>沒有圖片</option>';sel.disabled=true}$('.qty-input',node).value=item.qty;$('.price-input',node).value=item.unitPrice;$('.qty-input',node).onchange=e=>{item.qty=Math.max(1,Number(e.target.value)||1);updateTotals()};$('.price-input',node).onchange=e=>{item.unitPrice=Math.max(0,Math.ceil(Number(e.target.value)||0));updateTotals()};$('.delete-item',node).onclick=()=>{if(confirm(`刪除 ${item.artNo}？`)){state.items=state.items.filter(x=>x.id!==item.id);state.items.forEach((x,i)=>x.seq=i+1);renderItems()}};box.appendChild(node)});updateTotals()}
$('#scrollLatestBtn').onclick=()=>$('#invoiceItems').scrollTo({top:0,behavior:'smooth'});$('#clearInvoiceBtn').onclick=()=>{if(confirm('清空目前 Invoice？')){state.items=[];renderItems()}};
function reprice(){const r=Number($('#salesRate').value)||0;state.items.forEach(x=>x.unitPrice=Math.ceil(x.price*r));renderItems()}$('#salesRate').onchange=reprice;$('#currency').onchange=()=>{renderItems();renderCustomerSummary()};$('#discountAmount').oninput=updateTotals;
function words(n){return String(Math.floor(n))}
function renderPreview(){const t=totals(),rows=state.items.map((x,i)=>`<tr><td>${i+1}</td><td><strong>Lot.No. : ${esc(x.lotNo)}</strong><br>${esc(x.artNo)}</td><td>${x.descriptions.map(esc).join('<br>')}</td><td class="num">${x.qty}</td><td>${esc(x.unit)}</td><td class="num">${fmt(x.unitPrice)}</td><td class="num">${fmt(x.qty*x.unitPrice)}</td></tr>`).join('');$('#invoiceDocument').innerHTML=`<div class="letterhead"><h2>UNIVERSE GEMS &amp; JEWELLERY CO.</h2><p>UNIT 11-12, 10/F., FU HANG INDUSTRIAL BUILDING, NO. 1 HOK YUEN STREET EAST,<br>HUNG HOM, KOWLOON, HONG KONG · TEL : (852) 2363 5409 · FAX : (852) 2765 0343</p></div><div class="doc-title">Sales Invoice</div><div class="doc-grid"><div class="doc-meta">No. : <strong>${esc($('#invoiceNo').value)}</strong><br>Invoice Date : ${esc($('#invoiceDate').value)}<br>Shipment Method : ${esc($('#shipmentMethod').value)}<br>Currency : ${esc($('#currency').value)}<br><br>Customer : <strong>${esc($('#customerName').value)}</strong><br>${esc($('#customerAddress').value).replace(/\n/g,'<br>')}</div><div class="doc-meta"><strong>Vender's Banker</strong><br>The Hong Kong &amp; Shanghai Banking Corporation Ltd.<br>Address : 41 Ma Tau Wai Road,Hung Hom,Kowloon,Hong Kong<br>A/C # : 012-593570-001<br>A/C Name : Universe Gems &amp; Jewellery Co.</div></div><table class="doc-table"><thead><tr><th>No.</th><th>Article No.</th><th>Description</th><th>Quantity</th><th>Unit</th><th class="num">Unit Price</th><th class="num">Amount</th></tr><tr><th colspan="7">F.O.B. Value</th></tr></thead><tbody>${rows}</tbody></table><div class="doc-footer"><div class="doc-totals"><div><span>Total Quantity :</span><strong>${t.qty}</strong></div><div><span>Sub Total:</span><strong>${fmt(t.sub)}</strong></div><div><span>Discount:</span><strong>${fmt(t.discount)}</strong></div><div class="total"><span>Total : (${esc($('#currency').value)})</span><strong>${fmt(t.total)}</strong></div></div><p><strong>Remark :</strong> ${esc($('#remark').value)}</p></div>`}
$('#refreshPreviewBtn').onclick=renderPreview;$('#printBtn').onclick=()=>{renderPreview();window.print()};
async function startScanner(){
  if(state.scannerBusy)return;
  state.scannerBusy=true;
  if(!$('#scannerDialog').open)$('#scannerDialog').showModal();
  $('#scannerStatus').textContent='正在啟動後鏡頭…';
  const config={fps:15,qrbox:(w,h)=>({width:Math.floor(w*.82),height:Math.max(70,Math.floor(h*.18))})};
  const onSuccess=txt=>{addLot(txt);navigator.vibrate?.(80);$('#scannerStatus').textContent='已讀取 '+txt};
  try{
    if(typeof Html5Qrcode==='undefined')throw new Error('掃描程式未載入');
    if(state.scanner?.isScanning)await state.scanner.stop();
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
    $('#scannerStatus').textContent='已使用後鏡頭，請把 Barcode 放在掃描框內。';
  }catch(err){
    $('#scannerStatus').textContent='後鏡頭無法啟動：'+(err.message||err);
  }finally{
    state.scannerBusy=false;
  }
}
async function stopScanner(){if(state.scannerBusy)return;state.scannerBusy=true;try{if(state.scanner?.isScanning)await state.scanner.stop()}catch{}finally{$('#reader').innerHTML='';$('#scannerDialog').close();state.scannerBusy=false}}
$('#scanBtn').onclick=startScanner;$('#closeScannerBtn').onclick=stopScanner;
function exportRemaining(){if(!state.items.length)return alert('Invoice 沒有貨品。');const sold=new Set(state.items.map(x=>x.lotNo));const remain=state.stockRows.filter(r=>!sold.has(norm(field(r,['LOTNO']))));const ws=XLSX.utils.json_to_sheet(remain,{header:state.stockHeaders});const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Remaining Stock');const inv=norm($('#invoiceNo').value)||formatInvoiceNo();XLSX.writeFile(wb,`Remaining_Stock_${inv}_${remain.length}pcs_${today().replaceAll('-','')}.xlsx`);for(const lot of sold)state.products.delete(lot);state.stockRows=remain;state.items=[];advanceInvoiceSequence(inv);renderItems();status('#addMessage',`Invoice ${inv} 已 Confirm，下一張為 ${$('#invoiceNo').value}。`,'ok');status('#stockStatus',`目前 Remaining Stock：${state.products.size} 件。`,'ok')}
$('#confirmInvoiceBtn').onclick=()=>{if(confirm('Confirm Invoice 並匯出 Remaining Stock Excel？'))exportRemaining()};
renderCustomerSummary();renderItems();if('serviceWorker'in navigator)navigator.serviceWorker.register('./sw.js').catch(()=>{});
