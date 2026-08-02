const $=(s,r=document)=>r.querySelector(s),$$=(s,r=document)=>[...r.querySelectorAll(s)];
const state={products:new Map(),stockCatalog:new Map(),customers:new Map(),imageFiles:new Map(),items:[],stockRows:[],stockAllRows:[],stockHeaders:[],stockRowByLot:new Map(),stockWorkbook:null,stockSheetName:'jmsdata',stockFileName:'jmsdata.xls',stockIntegrityIssues:[],stoneAliases:new Map(),stoneVariantAliases:new Map(),stoneGroups:new Map(),stoneMappingName:'',articleMap:new Map(),articleMappingName:'',invoiceTemplateBuffer:null,invoiceTemplateName:'',documentType:'invoice',packageName:'',sortable:null,scanner:null,scannerBusy:false,scannerRunning:false,scannerZoom:{min:1,max:1,step:1,current:1},fx:{rate:1,date:'',source:'usd',fetching:false},quote:{karat:'18K',kitcoAsk:0,kitcoTime:'',fetching:false,source:''},inventoryHistory:new Map(),documentStore:{invoiceHeaders:[],invoiceItems:[],consignmentHeaders:[],consignmentItems:[],transactions:[]},recall:null,stockSearch:{query:'',types:[],stones:[],statuses:[],filtersOpen:false},editingItemId:null};
function formalItems(){return [...state.items].sort((a,b)=>(Number(a.seq)||0)-(Number(b.seq)||0))}
function displayItems(){return formalItems().reverse()}
function normalizeItemSequence(){state.items=formalItems();state.items.forEach((item,i)=>item.seq=i+1)}
const norm=v=>String(v??'').trim(),normCode=v=>String(v??'').replace(/\s+/g,'').toUpperCase(),normArt=v=>norm(v).toUpperCase();
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const today=()=>{const d=new Date(),y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,'0'),day=String(d.getDate()).padStart(2,'0');return `${y}-${m}-${day}`};$('#invoiceDate').value=today();
function englishInvoiceDate(iso){const m=String(iso||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)return String(iso||'');const months=['January','February','March','April','May','June','July','August','September','October','November','December'];return `${Number(m[3])} ${months[Number(m[2])-1]}, ${m[1]}`;}
function discountDisplay(value){const n=Math.max(0,Number(value)||0);return n>0?`(${fmt(n)})`:fmt(0)}
function invoiceYear(){return String(new Date().getFullYear()).slice(-2)}
function documentPrefix(type=state.documentType){return type==='consignment'?'CON':type==='quotation'?'QUO':'INV'}
function documentSequenceKey(type=state.documentType,yy=invoiceYear()){const key=type==='consignment'?'Consign':type==='quotation'?'Quotation':'Invoice';return `universe${key}Seq_${yy}`}
function getNextDocumentSequence(type=state.documentType,yy=invoiceYear()){
  const saved=Number(localStorage.getItem(documentSequenceKey(type,yy))||1);
  return Number.isInteger(saved)&&saved>0?saved:1;
}
function formatDocumentNo(seq=getNextDocumentSequence(),yy=invoiceYear(),type=state.documentType){
  return `${documentPrefix(type)}${yy}${String(seq).padStart(4,'0')}`;
}
function formatInvoiceNo(seq=getNextDocumentSequence('invoice'),yy=invoiceYear()){return formatDocumentNo(seq,yy,'invoice')}
function setDefaultInvoiceNo(force=false){
  const el=$('#invoiceNo');
  if(force||!norm(el.value)||!String(el.value).toUpperCase().startsWith(documentPrefix()))el.value=formatDocumentNo();
}
function advanceDocumentSequence(confirmedNo,type=state.documentType){
  const yy=invoiceYear(),prefix=documentPrefix(type);
  const m=String(confirmedNo||'').toUpperCase().match(new RegExp(`^${prefix}(\\d{2})(\\d{4})$`));
  let next=getNextDocumentSequence(type,yy)+1;
  if(m&&m[1]===yy)next=Math.max(next,Number(m[2])+1);
  localStorage.setItem(documentSequenceKey(type,yy),String(next));
  $('#invoiceNo').value=formatDocumentNo(next,yy,type);
}
function advanceInvoiceSequence(confirmedNo){advanceDocumentSequence(confirmedNo,'invoice')}
setDefaultInvoiceNo();
let previewTimer=null;
function schedulePreview(){clearTimeout(previewTimer);previewTimer=setTimeout(()=>{renderPreview();const el=$('#previewUpdatedAt');if(el)el.textContent='最後更新：'+new Date().toLocaleTimeString('zh-HK',{hour:'2-digit',minute:'2-digit',second:'2-digit'});},180)}
function status(id,msg,type=''){const el=$(id);el.textContent=msg;el.className='notice'+(type?' '+type:'')}
function documentLabels(type=state.documentType){if(type==='consignment')return{type:'consignment',title:'Sales Consign',short:'Consignment',date:'Date',no:'Consign No.',items:'Consignment 貨品',confirm:'Confirm Consignment 並更新 jmsdata',export:'匯出 Excel Consignment'};if(type==='quotation')return{type:'quotation',title:'Quotation',short:'Quotation',date:'Date',no:'Quotation No.',items:'Quotation 貨品',confirm:'Confirm Quotation（不扣庫存）',export:'匯出 Excel Quotation'};return{type:'invoice',title:'Sales Invoice',short:'Invoice',date:'Date',no:'Invoice No.',items:'Invoice 貨品',confirm:'Confirm Invoice 並更新 jmsdata',export:'匯出 Excel Invoice'}}
function updateDocumentTypeUI(){
  const l=documentLabels();
  $('#documentDataHeading').textContent=`客人 ${l.short} 資料`;
  $('#documentNoLabel').textContent=l.no;
  $('#documentDateLabel').textContent=l.date;
  $('#documentItemsHeading').textContent=l.items;
  $('#confirmInvoiceBtn').textContent=l.confirm;
  $('#exportExcelBtn').textContent=l.export;
  const pdfBtn=$('#exportPdfBtn');if(pdfBtn)pdfBtn.textContent=`輸出 PDF ${l.short}`;
  $('#invoiceNo').placeholder=`${documentPrefix()}YY0001`;
  setDefaultInvoiceNo(true);
  if(state.documentType!=='quotation'&&state.quote.karat!=='18K')state.quote.karat='18K';
  $$('input[name="quoteKarat"]').forEach(x=>x.checked=x.value===state.quote.karat);
  updateGoldQuoteUI();syncEffectivePrices();renderItems();renderPreview();
}
$$('input[name="documentType"]').forEach(r=>r.addEventListener('change',e=>{if(state.recall){e.preventDefault();e.target.checked=false;const current=$(`input[name="documentType"][value="${state.recall.type}"]`);if(current)current.checked=true;return alert('Recall 修改期間不能轉換文件類型；請先完成或取消 Recall。')}state.documentType=e.target.value;updateDocumentTypeUI()}));
function setImportCollapsed(key,collapsed=true){const card=document.querySelector(`[data-import-card="${key}"]`);if(!card)return;card.classList.toggle('collapsed',collapsed);const btn=card.querySelector('.import-toggle');if(btn){btn.textContent=collapsed?'展開':'收合';btn.setAttribute('aria-expanded',String(!collapsed))}}
$$('.import-toggle').forEach(btn=>btn.addEventListener('click',()=>{const card=btn.closest('.import-card');setImportCollapsed(card?.dataset.importCard,!card.classList.contains('collapsed'))}));
function field(row,names){const keys=Object.keys(row);for(const n of names){const k=keys.find(x=>x.trim().toUpperCase()===n);if(k)return row[k]}return''}

function fieldKey(row,names){const keys=Object.keys(row||{});for(const n of names){const k=keys.find(x=>x.trim().toUpperCase()===String(n).toUpperCase());if(k)return k}return''}
function setField(row,names,value){let key=fieldKey(row,names);if(!key){key=names[0];if(!state.stockHeaders.includes(key))state.stockHeaders.push(key)}row[key]=value}
const DOCUMENT_STORE_KEY='universeDocumentStore_v1';
const DOCUMENT_SHEETS={invoiceHeaders:'Invoice Header',invoiceItems:'Invoice Items',consignmentHeaders:'Consignment Header',consignmentItems:'Consignment Items',transactions:'Transaction History'};
const HEADER_COLUMNS=['DOCUMENT_NO','REVISION','DOCUMENT_STATUS','DOCUMENT_DATE','CUSTOMER_CODE','CUSTOMER','ADDRESS','SALES_RATE','CURRENCY','FX_RATE','SHIPMENT_METHOD','TERMS','DISCOUNT','REMARK','CREATED_AT','UPDATED_AT'];
const ITEM_COLUMNS=['DOCUMENT_NO','REVISION','ITEM_SEQ','LOTNO','ARTNO','ARTICLE','DESC1','DESC2','DESC3','DESC4','DESC5','DESC6','QTY','UNIT','PRICE','USD_UNIT_PRICE','UNIT_PRICE','CURRENCY','IMAGE_VARIANT','IMAGE_GRAYSCALE','IMAGE_AUTO_MATCHED','CUSTOM_IMAGE_FILE','ORIGINAL_STATUS','ORIGINAL_I','ORIGINAL_J','ORIGINAL_K','ORIGINAL_L','ORIGINAL_M','AFTER_STATUS','AFTER_I','AFTER_J','AFTER_K','AFTER_L','AFTER_M'];
const TRANSACTION_COLUMNS=['TIMESTAMP','DOCUMENT_TYPE','DOCUMENT_NO','REVISION','LOTNO','ARTNO','FROM_STATUS','TO_STATUS','FROM_I','FROM_J','FROM_K','FROM_L','FROM_M','TO_I','TO_J','TO_K','TO_L','TO_M'];
function emptyDocumentStore(){return{invoiceHeaders:[],invoiceItems:[],consignmentHeaders:[],consignmentItems:[],transactions:[]}}
function normalizeDocumentStore(raw){const base=emptyDocumentStore();for(const key of Object.keys(base))base[key]=Array.isArray(raw?.[key])?raw[key]:[];return base}
function loadDocumentStore(){try{state.documentStore=normalizeDocumentStore(JSON.parse(localStorage.getItem(DOCUMENT_STORE_KEY)||'null'))}catch{state.documentStore=emptyDocumentStore()}}
function saveDocumentStore(){try{localStorage.setItem(DOCUMENT_STORE_KEY,JSON.stringify(state.documentStore))}catch(err){console.warn('文件記錄無法寫入本機儲存。',err)}}
function sheetRows(wb,name){const ws=wb?.Sheets?.[name];return ws?XLSX.utils.sheet_to_json(ws,{defval:''}):[]}
function syncDocumentSequencesFromStore(){const yy=invoiceYear();for(const type of ['invoice','consignment']){const prefix=documentPrefix(type);let next=getNextDocumentSequence(type,yy);for(const h of latestDocumentHeaders(type)){const no=norm(docValue(h,'DOCUMENT_NO')).toUpperCase(),m=no.match(new RegExp(`^${prefix}(\d{2})(\d{4})$`));if(m&&m[1]===yy)next=Math.max(next,Number(m[2])+1)}localStorage.setItem(documentSequenceKey(type,yy),String(next))}if(!state.recall&&!state.items.length)setDefaultInvoiceNo(true)}
function importDocumentStoreFromWorkbook(wb){const store=emptyDocumentStore();for(const [key,name] of Object.entries(DOCUMENT_SHEETS))store[key]=sheetRows(wb,name);state.documentStore=normalizeDocumentStore(store);saveDocumentStore();syncDocumentSequencesFromStore();renderRecallResults()}
function storeKeysForType(type){return type==='consignment'?{headers:'consignmentHeaders',items:'consignmentItems'}:{headers:'invoiceHeaders',items:'invoiceItems'}}
function docValue(row,name){return field(row,[name])}
function revisionNumber(row){return Number(docValue(row,'REVISION'))||0}
function latestDocumentHeaders(type){const keys=storeKeysForType(type),latest=new Map();for(const row of state.documentStore[keys.headers]){const no=norm(docValue(row,'DOCUMENT_NO')).toUpperCase();if(!no)continue;const prev=latest.get(no);if(!prev||revisionNumber(row)>revisionNumber(prev)||revisionNumber(row)===revisionNumber(prev)&&String(docValue(row,'UPDATED_AT'))>String(docValue(prev,'UPDATED_AT')))latest.set(no,row)}return[...latest.values()].sort((a,b)=>String(docValue(b,'UPDATED_AT')||docValue(b,'DOCUMENT_DATE')).localeCompare(String(docValue(a,'UPDATED_AT')||docValue(a,'DOCUMENT_DATE'))))}
function findLatestDocument(type,documentNo){const no=norm(documentNo).toUpperCase(),header=latestDocumentHeaders(type).find(x=>norm(docValue(x,'DOCUMENT_NO')).toUpperCase()===no);if(!header)return null;const revision=revisionNumber(header),keys=storeKeysForType(type),items=state.documentStore[keys.items].filter(x=>norm(docValue(x,'DOCUMENT_NO')).toUpperCase()===no&&revisionNumber(x)===revision).sort((a,b)=>(Number(docValue(a,'ITEM_SEQ'))||0)-(Number(docValue(b,'ITEM_SEQ'))||0));return{header,items,revision}}
function numberValue(v){const n=Number(v);return Number.isFinite(n)?n:0}
function inventoryVectorFromRow(row){const i=numberValue(field(row,['POQTY','DELIVERED'])),j=numberValue(field(row,['FQTY','AVAILABLE'])),k=numberValue(field(row,['RQTY','SOLD ON HAND','SOLD_ON_HAND'])),l=numberValue(field(row,['COQTY','CONSIGNMENT']));let m=numberValue(field(row,['TQTY','BALANCE']));if(!m)m=i+j+k+l||1;return{i,j,k,l,m}}
function vectorFromRecord(row,prefix){const p=String(prefix||'').toUpperCase();return{i:numberValue(docValue(row,`${p}_I`)),j:numberValue(docValue(row,`${p}_J`)),k:numberValue(docValue(row,`${p}_K`)),l:numberValue(docValue(row,`${p}_L`)),m:numberValue(docValue(row,`${p}_M`))}}
function inventoryStatusFromVector(v){const positives=[['SOLD_DELIVERED',v.i],['AVAILABLE',v.j],['SOLD_ON_HAND',v.k],['CONSIGNED',v.l]].filter(([,n])=>numberValue(n)>0);return positives.length===1?positives[0][0]:positives.length?'CONFLICT':'UNKNOWN'}
function inventoryVectorForStatus(status,total=1){const m=Math.max(1,numberValue(total)||1),v={i:0,j:0,k:0,l:0,m};if(status==='SOLD_DELIVERED')v.i=m;else if(status==='SOLD_ON_HAND')v.k=m;else if(status==='CONSIGNED')v.l=m;else v.j=m;return v}
function vectorsEqual(a,b){return['i','j','k','l','m'].every(k=>numberValue(a?.[k])===numberValue(b?.[k]))}
function vectorFields(prefix,v){const p=String(prefix).toUpperCase();return{[`${p}_STATUS`]:inventoryStatusFromVector(v),[`${p}_I`]:numberValue(v.i),[`${p}_J`]:numberValue(v.j),[`${p}_K`]:numberValue(v.k),[`${p}_L`]:numberValue(v.l),[`${p}_M`]:numberValue(v.m)}}
function writeInventoryVector(row,v){setField(row,['POQTY','DELIVERED'],numberValue(v.i));setField(row,['FQTY','AVAILABLE'],numberValue(v.j));setField(row,['RQTY','SOLD ON HAND'],numberValue(v.k));setField(row,['COQTY','CONSIGNMENT'],numberValue(v.l));setField(row,['TQTY','BALANCE'],numberValue(v.m))}
function rebuildInventoryMaps(){const available=new Map(),history=new Map(),availableRows=[];for(const [lot,p] of state.stockCatalog){const row=state.stockRowByLot.get(String(lot)),status=inventoryStatusFromVector(inventoryVectorFromRow(row||{}));if(status==='AVAILABLE'){available.set(String(lot),p);if(row)availableRows.push(row)}else history.set(String(lot),{...productSnapshot(p),status,updatedAt:Date.now()})}state.products=available;state.stockRows=availableRows;state.inventoryHistory=history;saveInventoryHistory();renderStockSearch();updateTotals()}
function setInventoryVectorForLot(lot,vector){const key=String(lot),row=state.stockRowByLot.get(key);if(!row)throw new Error(`jmsdata 找不到 LOTNO ${key}`);writeInventoryVector(row,vector);return vector}
function appendTransaction({type,documentNo,revision,item,before,after}){if(vectorsEqual(before,after))return;state.documentStore.transactions.push({TIMESTAMP:new Date().toISOString(),DOCUMENT_TYPE:type==='consignment'?'Consignment':'Invoice',DOCUMENT_NO:documentNo,REVISION:revision,LOTNO:item.lotNo,ARTNO:item.artNo,FROM_STATUS:inventoryStatusFromVector(before),TO_STATUS:inventoryStatusFromVector(after),FROM_I:before.i,FROM_J:before.j,FROM_K:before.k,FROM_L:before.l,FROM_M:before.m,TO_I:after.i,TO_J:after.j,TO_K:after.k,TO_L:after.l,TO_M:after.m})}
function headerRowForCurrentDocument(type,documentNo,revision,createdAt='',documentStatus='CONFIRMED'){const now=new Date().toISOString();return{DOCUMENT_NO:documentNo,REVISION:revision,DOCUMENT_STATUS:documentStatus,DOCUMENT_DATE:norm($('#invoiceDate').value),CUSTOMER_CODE:norm($('#customerCode').value),CUSTOMER:norm($('#customerName').value),ADDRESS:norm($('#customerAddress').value),SALES_RATE:numberValue($('#salesRate').value),CURRENCY:currencyCode(),FX_RATE:currencyCode()==='USD'?1:currentFxRate(),SHIPMENT_METHOD:norm($('#shipmentMethod').value),TERMS:norm($('#customerTerms').value),DISCOUNT:numberValue($('#discountAmount').value),REMARK:norm($('#remark').value),CREATED_AT:createdAt||now,UPDATED_AT:now}}
function itemRowForDocument(item,documentNo,revision,original,after){const desc=[...(item.descriptions||[])];return{DOCUMENT_NO:documentNo,REVISION:revision,ITEM_SEQ:Number(item.seq)||1,LOTNO:item.lotNo,ARTNO:item.artNo,ARTICLE:item.article||'',DESC1:desc[0]||'',DESC2:desc[1]||'',DESC3:desc[2]||'',DESC4:desc[3]||'',DESC5:desc[4]||'',DESC6:desc[5]||'',QTY:Number(item.qty)||1,UNIT:item.unit||'PC',PRICE:Number(item.price)||0,USD_UNIT_PRICE:Number(item.usdUnitPrice)||0,UNIT_PRICE:Number(item.unitPrice)||0,CURRENCY:currencyCode(),IMAGE_VARIANT:item.imageVariant||'',IMAGE_GRAYSCALE:item.imageGrayscale?1:0,IMAGE_AUTO_MATCHED:item.imageAutoMatched?1:0,CUSTOM_IMAGE_FILE:item.customImage?.fileName||'',...vectorFields('ORIGINAL',original),...vectorFields('AFTER',after)}}
function worksheetFromRows(rows,headers){return rows.length?XLSX.utils.json_to_sheet(rows,{header:headers}):XLSX.utils.aoa_to_sheet([headers])}
function preserveWorksheetLayout(target,source){if(!target||!source)return target;for(const prop of ['!cols','!rows','!merges','!autofilter','!freeze','!margins'])if(source[prop])target[prop]=structuredClone(source[prop]);for(const address of Object.keys(target)){if(address[0]==='!')continue;const old=source[address],cell=target[address];if(!old||!cell)continue;if(old.s)cell.s=structuredClone(old.s);if(old.z)cell.z=old.z}return target}
function buildJmsdataWorkbook(){if(!state.stockAllRows.length)throw new Error('尚未匯入 jmsdata。');const wb=XLSX.utils.book_new(),stockName=(state.stockSheetName||'jmsdata').slice(0,31),originalStock=state.stockWorkbook?.Sheets?.[state.stockSheetName],stockWs=preserveWorksheetLayout(XLSX.utils.json_to_sheet(state.stockAllRows,{header:state.stockHeaders}),originalStock);if(state.stockWorkbook?.Props)wb.Props=structuredClone(state.stockWorkbook.Props);if(state.stockWorkbook?.Custprops)wb.Custprops=structuredClone(state.stockWorkbook.Custprops);XLSX.utils.book_append_sheet(wb,stockWs,stockName);const managed=new Set([state.stockSheetName,...Object.values(DOCUMENT_SHEETS)]);if(state.stockWorkbook){for(const name of state.stockWorkbook.SheetNames||[]){if(managed.has(name)||name===state.stockSheetName)continue;XLSX.utils.book_append_sheet(wb,state.stockWorkbook.Sheets[name],name.slice(0,31))}}XLSX.utils.book_append_sheet(wb,worksheetFromRows(state.documentStore.invoiceHeaders,HEADER_COLUMNS),DOCUMENT_SHEETS.invoiceHeaders);XLSX.utils.book_append_sheet(wb,worksheetFromRows(state.documentStore.invoiceItems,ITEM_COLUMNS),DOCUMENT_SHEETS.invoiceItems);XLSX.utils.book_append_sheet(wb,worksheetFromRows(state.documentStore.consignmentHeaders,HEADER_COLUMNS),DOCUMENT_SHEETS.consignmentHeaders);XLSX.utils.book_append_sheet(wb,worksheetFromRows(state.documentStore.consignmentItems,ITEM_COLUMNS),DOCUMENT_SHEETS.consignmentItems);XLSX.utils.book_append_sheet(wb,worksheetFromRows(state.documentStore.transactions,TRANSACTION_COLUMNS),DOCUMENT_SHEETS.transactions);return wb}
function exportJmsdata(){try{const wb=buildJmsdataWorkbook();XLSX.writeFile(wb,'jmsdata.xls',{bookType:'xls'});const el=$('#jmsdataExportStatus');if(el){el.textContent='已輸出 jmsdata.xls；請在「檔案」App 儲存時確認取代舊檔。';el.className='notice ok'}return true}catch(err){console.error(err);alert('匯出 jmsdata 失敗：'+(err.message||err));return false}}
function setRecallLock(locked){$$('input[name="documentType"]').forEach(r=>r.disabled=!!locked);const no=$('#invoiceNo');if(no)no.readOnly=!!locked;$('#cancelRecallBtn')?.classList.toggle('hidden',!locked)}
function clearCurrentDocument({keepCustomer=true}={}){releaseCustomImages();state.items=[];state.editingItemId=null;$('#discountAmount').value=0;$('#remark').value='';if(!keepCustomer){$('#customerCode').value='';$('#customerName').value='';$('#customerAddress').value='';$('#salesRate').value='';$('#customerTerms').value=''}renderItems();renderCustomerSummary();schedulePreview()}
function restoreDocumentFields(header){$('#invoiceNo').value=norm(docValue(header,'DOCUMENT_NO'));$('#invoiceDate').value=norm(docValue(header,'DOCUMENT_DATE'))||today();$('#customerCode').value=norm(docValue(header,'CUSTOMER_CODE'));$('#customerName').value=norm(docValue(header,'CUSTOMER'));$('#customerAddress').value=norm(docValue(header,'ADDRESS'));$('#salesRate').value=numberValue(docValue(header,'SALES_RATE'));$('#currency').value=norm(docValue(header,'CURRENCY'))||'USD';$('#fxRate').value=numberValue(docValue(header,'FX_RATE'))||'';state.fx={rate:numberValue(docValue(header,'FX_RATE'))||1,date:'',source:'manual',fetching:false};$('#shipmentMethod').value=norm(docValue(header,'SHIPMENT_METHOD'));$('#customerTerms').value=norm(docValue(header,'TERMS'));$('#discountAmount').value=numberValue(docValue(header,'DISCOUNT'));$('#remark').value=norm(docValue(header,'REMARK'))}
function recallItemFromRow(row){const lot=norm(docValue(row,'LOTNO')),p=state.stockCatalog.get(lot)||{lotNo:lot,artNo:normArt(docValue(row,'ARTNO')),price:numberValue(docValue(row,'PRICE')),unit:norm(docValue(row,'UNIT'))||'PC',article:norm(docValue(row,'ARTICLE')),descriptions:[1,2,3,4,5,6].map(n=>norm(docValue(row,`DESC${n}`))).filter(Boolean)};const code=norm(docValue(row,'CURRENCY'))||currencyCode(),unitPrice=numberValue(docValue(row,'UNIT_PRICE')),item={...productSnapshot(p),id:Date.now()+Math.random(),seq:numberValue(docValue(row,'ITEM_SEQ'))||1,qty:numberValue(docValue(row,'QTY'))||1,usdUnitPrice:numberValue(docValue(row,'USD_UNIT_PRICE'))||Math.ceil((Number(p.price)||0)*(Number($('#salesRate').value)||0)),currencyPrices:{},quote14kCurrencyPrices:{},unitPrice,imageVariant:norm(docValue(row,'IMAGE_VARIANT'))||chooseImageMatch(p).variant,imageGrayscale:numberValue(docValue(row,'IMAGE_GRAYSCALE'))>0,imageAutoMatched:numberValue(docValue(row,'IMAGE_AUTO_MATCHED'))>0,_originalVector:vectorFromRecord(row,'ORIGINAL')};item.currencyPrices[code]=unitPrice;return item}
function openRecallDocument(type,documentNo){if(!state.stockAllRows.length)return alert('請先匯入最新的 jmsdata.xls，才可安全 Recall 及更新庫存。');const found=findLatestDocument(type,documentNo);if(!found)return alert('找不到這張文件。');releaseCustomImages();state.recall={type,documentNo:norm(docValue(found.header,'DOCUMENT_NO')),baseRevision:found.revision,header:found.header,items:found.items};state.documentType=type;const radio=$(`input[name="documentType"][value="${type}"]`);if(radio)radio.checked=true;updateDocumentTypeUI();restoreDocumentFields(found.header);state.items=found.items.map(recallItemFromRow);normalizeItemSequence();setRecallLock(true);const banner=$('#recallActive');if(banner){banner.classList.remove('hidden');banner.innerHTML=`<strong>正在修改 ${esc(state.recall.documentNo)}</strong><span>Revision ${state.recall.baseRevision} → ${state.recall.baseRevision+1}</span>`}updateFxPanel();syncEffectivePrices();renderItems();renderCustomerSummary();renderPreview();status('#addMessage',`${state.recall.documentNo} 已 Recall，可刪除／新增貨品後重新 Confirm。`,'ok');window.scrollTo({top:$('#invoice').offsetTop-70,behavior:'smooth'})}
function cancelRecall(){if(!state.recall)return;const no=state.recall.documentNo;if(!confirm(`取消修改 ${no}？尚未 Confirm 的變更不會保存。`))return;state.recall=null;setRecallLock(false);$('#recallActive')?.classList.add('hidden');clearCurrentDocument();setDefaultInvoiceNo(true);renderRecallResults();status('#addMessage',`已取消 Recall ${no}；庫存沒有變更。`,'ok')}
function renderRecallResults(){const box=$('#recallResults');if(!box)return;const q=norm($('#recallSearchInput')?.value).toUpperCase(),typeFilter=$('#recallTypeFilter')?.value||'all',types=typeFilter==='all'?['invoice','consignment']:[typeFilter],rows=[];for(const type of types){for(const h of latestDocumentHeaders(type)){const no=norm(docValue(h,'DOCUMENT_NO')),customer=norm(docValue(h,'CUSTOMER')),code=norm(docValue(h,'CUSTOMER_CODE')),date=norm(docValue(h,'DOCUMENT_DATE')),hay=`${no} ${customer} ${code} ${date}`.toUpperCase();if(q&&!hay.includes(q))continue;const found=findLatestDocument(type,no);rows.push({type,h,count:found?.items.length||0})}}rows.sort((a,b)=>String(docValue(b.h,'UPDATED_AT')||docValue(b.h,'DOCUMENT_DATE')).localeCompare(String(docValue(a.h,'UPDATED_AT')||docValue(a.h,'DOCUMENT_DATE'))));if(!rows.length){box.innerHTML='<div class="notice">找不到已保存的 Invoice／Consignment。</div>';return}box.innerHTML='';for(const row of rows.slice(0,100)){const no=norm(docValue(row.h,'DOCUMENT_NO')),documentStatus=norm(docValue(row.h,'DOCUMENT_STATUS')).toUpperCase(),statusText=documentStatus==='CANCELLED'?'Cancelled':'Confirmed',card=document.createElement('div');card.className='recall-result';card.innerHTML=`<div><strong>${esc(no)}</strong><span>${row.type==='consignment'?'Consignment':'Invoice'} · Rev ${revisionNumber(row.h)} · ${esc(docValue(row.h,'DOCUMENT_DATE'))} · ${statusText}</span><small>${esc(docValue(row.h,'CUSTOMER_CODE'))} · ${esc(docValue(row.h,'CUSTOMER'))} · ${row.count} 件</small></div><button type="button">重新開啟</button>`;$('button',card).onclick=()=>openRecallDocument(row.type,no);box.appendChild(card)}}
const INVENTORY_HISTORY_KEY='universeInventoryHistory_v1';
const KITCO_CACHE_KEY='universeKitcoGoldAsk_v1';
const KITCO_GRAPHQL='https://kdb-gw.prod.kitco.com/';
const KITCO_PAGE='https://www.kitco.com/price/precious-metals?Currency=USD&Symbol=GOLD';
const KITCO_QUERY=`fragment MetalFragment on Metal { ID symbol currency name results { ...MetalQuoteFragment } } fragment MetalQuoteFragment on Quote { ID ask bid change changePercentage close high low mid open originalTime timestamp unit } query MetalQuote($symbol:String!,$currency:String!,$timestamp:Int){ GetMetalQuoteV3(symbol:$symbol currency:$currency timestamp:$timestamp){ ...MetalFragment } }`;
const K14_WEIGHT_FACTOR=.83,K14_WEIGHT_STEP=.05,TROY_OZ_G=31.1034768;
function loadInventoryHistory(){try{const raw=JSON.parse(localStorage.getItem(INVENTORY_HISTORY_KEY)||'[]');state.inventoryHistory=new Map((Array.isArray(raw)?raw:[]).filter(x=>x&&x.lotNo).map(x=>[String(x.lotNo),x]))}catch{state.inventoryHistory=new Map()}}
function saveInventoryHistory(){try{localStorage.setItem(INVENTORY_HISTORY_KEY,JSON.stringify([...state.inventoryHistory.values()]))}catch{}}
function productSnapshot(p){return{lotNo:p.lotNo,artNo:p.artNo,price:Number(p.price)||0,unit:p.unit||'PC',article:p.article||'',descriptions:[...(p.descriptions||[])],desc2:p.desc2||''}}
function historyStatusLabel(status){return status==='CONSIGNED'?'Consigned':status==='SOLD_ON_HAND'?'Sold - On Hand':status==='SOLD_DELIVERED'?'Sold - Delivered':'Available'}
function historyStatusClass(status){return status==='CONSIGNED'?'consigned':status==='SOLD_ON_HAND'?'sold-on-hand':status==='SOLD_DELIVERED'?'sold-delivered':'available'}
function recordInventoryMovement(item,statusValue,docNo=''){
  const lot=String(item.lotNo),prev=state.inventoryHistory.get(lot)||productSnapshot(item),docs=Array.isArray(prev.docs)?prev.docs:[];
  docs.push({type:state.documentType,no:docNo,date:norm($('#invoiceDate')?.value),customerCode:norm($('#customerCode')?.value),customer:norm($('#customerName')?.value),at:Date.now()});
  state.inventoryHistory.set(lot,{...prev,...productSnapshot(item),status:statusValue,docs,updatedAt:Date.now()});saveInventoryHistory();
}
function findLatestDocumentForLot(lot,type){const key=storeKeysForType(type),matches=state.documentStore[key.items].filter(x=>String(docValue(x,'LOTNO'))===String(lot));let best=null;for(const item of matches){const no=norm(docValue(item,'DOCUMENT_NO')),revision=revisionNumber(item),header=state.documentStore[key.headers].find(h=>norm(docValue(h,'DOCUMENT_NO'))===no&&revisionNumber(h)===revision);if(!best||String(docValue(header||{},'UPDATED_AT'))>String(best.updatedAt||''))best={documentNo:no,revision,updatedAt:docValue(header||{},'UPDATED_AT')}}return best}
function markInventoryDelivered(lot){const key=String(lot),row=state.stockRowByLot.get(key),p=state.stockCatalog.get(key);if(!row||!p)return alert('jmsdata 找不到這件貨。');const before=inventoryVectorFromRow(row);if(inventoryStatusFromVector(before)!=='SOLD_ON_HAND')return alert('只有 Sold on hand 貨品可以標記 Delivered。');const after=inventoryVectorForStatus('SOLD_DELIVERED',before.m),owner=findLatestDocumentForLot(key,'invoice');setInventoryVectorForLot(key,after);appendTransaction({type:'invoice',documentNo:owner?.documentNo||'',revision:owner?.revision||0,item:p,before,after});saveDocumentStore();rebuildInventoryMaps();exportJmsdata();status('#stockSearchMessage',`${p.artNo} / LOTNO ${key} 已標記 Delivered，並輸出更新後 jmsdata.xls。`,'ok')}
function latestDeliveredTransactionForLot(lot){const key=String(lot),rows=state.documentStore.transactions.filter(row=>String(docValue(row,'LOTNO'))===key&&norm(docValue(row,'TO_STATUS')).toUpperCase()==='SOLD_DELIVERED');rows.sort((a,b)=>String(docValue(a,'TIMESTAMP')).localeCompare(String(docValue(b,'TIMESTAMP'))));return rows[rows.length-1]||null}
function undoInventoryDelivered(lot){const key=String(lot),row=state.stockRowByLot.get(key),p=state.stockCatalog.get(key);if(!row||!p)return alert('jmsdata 找不到這件貨。');const before=inventoryVectorFromRow(row);if(inventoryStatusFromVector(before)!=='SOLD_DELIVERED')return alert('只有 Delivered 貨品可以取消已交貨。');const tx=latestDeliveredTransactionForLot(key);if(!tx)return alert('Transaction History 找不到這件貨交貨前的狀態，為避免錯誤不能自動回復。');const deliveredVector=vectorFromRecord(tx,'TO'),after=vectorFromRecord(tx,'FROM'),afterStatus=inventoryStatusFromVector(after);if(!vectorsEqual(before,deliveredVector))return alert('目前 Delivered 數量與 Transaction History 不一致，為避免覆蓋資料不能自動回復。');if(!['SOLD_ON_HAND','CONSIGNED'].includes(afterStatus))return alert(`交貨前狀態是 ${historyStatusLabel(afterStatus)}，請先檢查 Transaction History。`);const documentType=norm(docValue(tx,'DOCUMENT_TYPE')).toUpperCase()==='CONSIGNMENT'?'consignment':'invoice',documentNo=norm(docValue(tx,'DOCUMENT_NO')),revision=revisionNumber(tx);setInventoryVectorForLot(key,after);appendTransaction({type:documentType,documentNo,revision,item:p,before,after});saveDocumentStore();rebuildInventoryMaps();exportJmsdata();status('#stockSearchMessage',`${p.artNo} / LOTNO ${key} 已取消 Delivered，恢復為 ${historyStatusLabel(afterStatus)}，並輸出更新後 jmsdata.xls。`,'ok')}
function quote14KMode(){return state.documentType==='quotation'&&state.quote.karat==='14K'}
function parse18KDesc1(desc){const raw=norm(desc),m=raw.match(/^(\d+(?:\.\d+)?)\s*([A-Z]+(?:\/[A-Z]+)*)750(.*)$/i);if(!m)return null;return{raw,weight:Number(m[1]),metal:m[2].toUpperCase(),suffix:m[3]||''}}
function roundUpWeight05(v){const n=Number(v)||0;return Math.round(Math.ceil((n-1e-9)/K14_WEIGHT_STEP)*K14_WEIGHT_STEP*100)/100}
function estimate14KWeightFromItem(item){const p=parse18KDesc1((item.descriptions||[])[0]);return p?roundUpWeight05(p.weight*K14_WEIGHT_FACTOR):0}
function effectiveDescriptions(item){const d=[...(item.descriptions||[])];if(!quote14KMode()||!d.length)return d;const p=parse18KDesc1(d[0]);if(!p)return d;const w=roundUpWeight05(p.weight*K14_WEIGHT_FACTOR);d[0]=`${w.toFixed(2)}${p.metal}585${p.suffix} (${p.raw})`;return d}
function goldMetrics(){const ask=Number(state.quote.kitcoAsk)||0;if(!(ask>0))return null;const base=Math.ceil((ask*1.01)/10)*10;const g18=base/TROY_OZ_G*.75*1.12,g14=base/TROY_OZ_G*.585*1.12;return{ask,base,g18,g14}}
function goldDifferenceUsd(item){const m=goldMetrics(),p=parse18KDesc1((item.descriptions||[])[0]);if(!m||!p)return 0;const w14=roundUpWeight05(p.weight*K14_WEIGHT_FACTOR);return Math.max(0,p.weight*m.g18-w14*m.g14)}
function quote14KReady(){return !quote14KMode()||!!goldMetrics()}
function effectiveUsdPrice(item){const base=baseUsdPrice(item);if(!(quote14KMode()&&goldMetrics()))return base;const raw=Math.max(0,base-goldDifferenceUsd(item));return Math.ceil(raw-1e-9)}
function activePriceOverrides(item){item.currencyPrices=item.currencyPrices||{};item.quote14kCurrencyPrices=item.quote14kCurrencyPrices||{};return quote14KMode()?item.quote14kCurrencyPrices:item.currencyPrices}
function saveKitcoCache(){try{if(state.quote.kitcoAsk>0)localStorage.setItem(KITCO_CACHE_KEY,JSON.stringify({ask:state.quote.kitcoAsk,time:state.quote.kitcoTime||'',fetchedAt:Date.now()}))}catch{}}
function loadKitcoCache(){try{const x=JSON.parse(localStorage.getItem(KITCO_CACHE_KEY)||'null');if(x&&Number(x.ask)>0){state.quote.kitcoAsk=Number(x.ask);state.quote.kitcoTime=x.time||'';state.quote.source='cache';return x}}catch{}return null}
function updateGoldQuoteUI(message='',type=''){
  const panel=$('#quotationKaratPanel'),gold=$('#goldQuotePanel');if(!panel)return;
  panel.classList.toggle('hidden',state.documentType!=='quotation');
  gold?.classList.toggle('hidden',!quote14KMode());
  const ask=$('#kitcoAskInput');if(ask&&document.activeElement!==ask)ask.value=state.quote.kitcoAsk>0?String(state.quote.kitcoAsk):'';
  const m=goldMetrics();$('#companyGoldBase').textContent=m?`USD ${m.base.toFixed(0)} / oz`:'—';$('#gold18PerGram').textContent=m?`USD ${m.g18.toFixed(3)}`:'—';$('#gold14PerGram').textContent=m?`USD ${m.g14.toFixed(3)}`:'—';
  if(message){const el=$('#goldQuoteStatus');el.textContent=message;el.className='fx-status'+(type?' '+type:'')}
}
function setKitcoAsk(ask,{source='manual',time='',message=''}={}){const n=Number(ask);if(!(Number.isFinite(n)&&n>0))return false;state.quote.kitcoAsk=n;state.quote.kitcoTime=time||new Date().toLocaleString('zh-HK');state.quote.source=source;saveKitcoCache();syncEffectivePrices({clearCurrentOverride:true});renderItems();updateGoldQuoteUI(message||`${source==='online'?'Kitco 線上 Ask':'使用手動 Kitco Ask'}：USD ${n.toFixed(2)} / oz`,source==='online'?'ok':'warn');schedulePreview();return true}
async function fetchKitcoAsk({silent=false}={}){
  if(state.quote.fetching)return;state.quote.fetching=true;const btn=$('#refreshGoldBtn');if(btn)btn.disabled=true;if(!silent)updateGoldQuoteUI('正在向 Kitco 查詢 Gold Ask…');
  const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),12000);let lastErr='';
  try{
    for(const symbol of ['AU','GOLD','XAU']){
      try{
        const res=await fetch(KITCO_GRAPHQL,{method:'POST',mode:'cors',credentials:'omit',headers:{'Content-Type':'application/json','Accept':'application/json'},body:JSON.stringify({query:KITCO_QUERY,variables:{symbol,currency:'USD',timestamp:Math.floor(Date.now()/1000)},operationName:'MetalQuote'}),signal:controller.signal});
        if(!res.ok)throw new Error(`HTTP ${res.status}`);const data=await res.json();const metal=data?.data?.GetMetalQuoteV3,rawResults=metal?.results,results=Array.isArray(rawResults)?rawResults:(rawResults?[rawResults]:[]);const quotes=results.filter(x=>Number(x?.ask)>0).sort((a,b)=>Number(b.timestamp||0)-Number(a.timestamp||0));const q=quotes[0];if(q&&setKitcoAsk(Number(q.ask),{source:'online',time:q.originalTime||q.timestamp||'',message:`Kitco Gold Ask：USD ${Number(q.ask).toFixed(2)} / oz · 線上更新成功`}))return;
        lastErr=data?.errors?.[0]?.message||'沒有 Ask 資料';
      }catch(err){lastErr=err?.message||String(err)}
    }
    // Browser CORS policies can occasionally block Kitco's GraphQL gateway.
    // As a fallback, read Kitco's public precious-metals page through a CORS relay
    // and extract the GOLD Bid / Ask pair. No customer or document data is sent.
    try{
      const relay=`https://api.allorigins.win/raw?url=${encodeURIComponent(KITCO_PAGE)}`,res=await fetch(relay,{signal:controller.signal});if(!res.ok)throw new Error(`page HTTP ${res.status}`);const text=(await res.text()).replace(/<[^>]+>/g,' ').replace(/\s+/g,' ');const m=text.match(/GOLD.{0,260}?(\d{1,3}(?:,\d{3})*\.\d{2})\s+(\d{1,3}(?:,\d{3})*\.\d{2})/i);if(!m)throw new Error('Kitco 頁面找不到 Gold Ask');const ask=Number(m[2].replaceAll(',',''));if(setKitcoAsk(ask,{source:'online',message:`Kitco Gold Ask：USD ${ask.toFixed(2)} / oz · 頁面備援更新成功`}))return;
    }catch(err){lastErr=`${lastErr}; ${err?.message||err}`}
    throw new Error(lastErr||'Kitco 沒有回傳 Gold Ask');
  }catch(err){const cache=loadKitcoCache();if(cache){syncEffectivePrices({clearCurrentOverride:true});renderItems();updateGoldQuoteUI(`Kitco 線上更新失敗；暫用上次 Ask USD ${Number(cache.ask).toFixed(2)}。可手動輸入最新 Ask。`,'warn')}else updateGoldQuoteUI(`無法線上取得 Kitco Ask：${err.message||err}。請手動輸入 Kitco Ask。`,'error')}
  finally{clearTimeout(timer);state.quote.fetching=false;if(btn)btn.disabled=false}
}
function setQuoteKarat(value){state.quote.karat=value==='14K'?'14K':'18K';$$('input[name="quoteKarat"]').forEach(x=>x.checked=x.value===state.quote.karat);syncEffectivePrices();updateGoldQuoteUI();renderItems();schedulePreview();if(quote14KMode()){if(!state.quote.kitcoAsk)loadKitcoCache();updateGoldQuoteUI();fetchKitcoAsk({silent:!!state.quote.kitcoAsk})}}
loadInventoryHistory();loadDocumentStore();loadKitcoCache();
const FX_API='https://api.frankfurter.dev/v2/rate';
const FX_API_V1='https://api.frankfurter.dev/v1/latest';
function currencyCode(){return norm($('#currency')?.value||'USD').toUpperCase()||'USD'}
// Money is displayed with 2 decimals. For EUR converted with an automatic
// Frankfurter rate, cents are dropped (truncated) from Unit Price. If the FX
// rate is entered manually, EUR Unit Price is rounded to the nearest euro.
function currencyDigits(){return 2}
function roundCurrency(v,code=currencyCode()){const p=100;return Math.round((Number(v)||0)*p+Number.EPSILON)/p}
function eurUnitPrice(v,source=state.fx?.source){const n=Math.max(0,Number(v)||0);return source==='manual'?Math.round(n):Math.floor(n+1e-9)}
function roundUnitPrice(v,code=currencyCode()){const n=Math.max(0,Number(v)||0);return code==='EUR'?eurUnitPrice(n):roundCurrency(n,code)}
function fmt(v,code=currencyCode()){return new Intl.NumberFormat('en-US',{style:'currency',currency:code,minimumFractionDigits:2,maximumFractionDigits:2}).format(Number(v)||0)}
function currencyExcelFormat(code=currencyCode(),negative=false){
  const formats={USD:'$#,##0.00',EUR:'€#,##0.00'};
  const positive=formats[code]||`"${code}" #,##0.00`;
  if(!negative)return positive;
  const body=positive.replace(/^([^#0]*)(.*)$/,'$1$2');
  return `(${body});(${body});${positive}`;
}
function fxCacheKey(code){return `universeFx_USD_${code}`}
function currentFxRate(){if(currencyCode()==='USD')return 1;const v=Number($('#fxRate')?.value);return Number.isFinite(v)&&v>0?v:0}
function fxPricingReady(){return currencyCode()==='USD'||currentFxRate()>0}
function baseUsdPrice(item){const stored=Number(item.usdUnitPrice);return Number.isFinite(stored)&&stored>=0?stored:Math.max(0,Number(item.unitPrice)||0)}
function convertedFromUsd(usd){const code=currencyCode();if(code==='USD')return roundUnitPrice(usd,'USD');const rate=currentFxRate();return rate>0?roundUnitPrice(usd*rate,code):0}
function syncEffectivePrices({clearCurrentOverride=false}={}){
  const code=currencyCode();
  for(const item of state.items){
    const overrides=activePriceOverrides(item);
    if(clearCurrentOverride)delete overrides[code];
    const manual=Number(overrides[code]);
    item.unitPrice=Number.isFinite(manual)&&manual>=0?manual:convertedFromUsd(effectiveUsdPrice(item));
  }
}
function setFxStatus(text,type=''){const el=$('#fxStatus');if(!el)return;el.textContent=text;el.className='fx-status'+(type?' '+type:'')}
function saveFxCache(code,rate,date){try{localStorage.setItem(fxCacheKey(code),JSON.stringify({rate,date,fetchedAt:Date.now()}))}catch{}}
function loadFxCache(code){try{const x=JSON.parse(localStorage.getItem(fxCacheKey(code))||'null');return x&&Number(x.rate)>0?x:null}catch{return null}}
function updateFxPanel(){
  const code=currencyCode(),panel=$('#fxPanel');if(!panel)return;
  $('#fxQuoteCurrency').textContent=code;
  if(code==='USD'){panel.classList.add('hidden');state.fx={rate:1,date:'',source:'usd',fetching:false};syncEffectivePrices();return}
  panel.classList.remove('hidden');
}
function applyFxRate(rate,date='',source='manual',message=''){
  const code=currencyCode(),n=Number(rate);if(code==='USD')return;
  if(!(Number.isFinite(n)&&n>0))return;
  $('#fxRate').value=String(n);
  state.fx={rate:n,date,source,fetching:false};
  syncEffectivePrices({clearCurrentOverride:true});
  renderItems();renderCustomerSummary();schedulePreview();
  if(message)setFxStatus(message,source==='online'?'ok':source==='cache'?'warn':'');
}
async function fetchReferenceFxRate({silent=false}={}){
  const code=currencyCode();updateFxPanel();if(code==='USD')return;
  if(code!=='EUR'){setFxStatus(`目前只支援 USD / EUR。`,'error');return}
  if(state.fx.fetching)return;
  state.fx.fetching=true;$('#refreshFxBtn').disabled=true;
  if(!silent)setFxStatus(`正在取得 Frankfurter v2 USD → ${code} 參考匯率…`);
  const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),12000);
  try{
    let rate=0,date='',sourceLabel='Frankfurter v2';
    try{
      const res=await fetch(`${FX_API}/USD/${encodeURIComponent(code)}`,{cache:'no-store',signal:controller.signal});
      if(!res.ok)throw new Error(`v2 HTTP ${res.status}`);
      const data=await res.json();
      rate=Number(data.rate);date=norm(data.date);
      if(!(Number.isFinite(rate)&&rate>0))throw new Error('v2 回傳匯率無效');
    }catch(v2err){
      const res=await fetch(`${FX_API_V1}?base=USD&symbols=${encodeURIComponent(code)}`,{cache:'no-store',signal:controller.signal});
      if(!res.ok)throw new Error(`v1 HTTP ${res.status}; ${v2err?.message||v2err}`);
      const data=await res.json();
      rate=Number(data?.rates?.[code]);date=norm(data.date);sourceLabel='Frankfurter v1 fallback';
      if(!(Number.isFinite(rate)&&rate>0))throw new Error(`v1 回傳匯率無效; ${v2err?.message||v2err}`);
    }
    saveFxCache(code,rate,date);
    applyFxRate(rate,date,'online',`線上參考匯率 · ${date||'最新工作日'} · ${sourceLabel} · EUR Unit Price 刪除 cents`);
  }catch(err){
    const cached=loadFxCache(code);
    console.warn('Frankfurter FX fetch failed',code,err);
    if(cached){applyFxRate(cached.rate,cached.date,'cache',`未能更新；沿用上次 Frankfurter 匯率 · ${cached.date||new Date(cached.fetchedAt).toLocaleDateString('zh-HK')} · EUR Unit Price 刪除 cents`)}
    else{state.fx={rate:0,date:'',source:'error',fetching:false};$('#fxRate').value='';syncEffectivePrices({clearCurrentOverride:true});renderItems();updateTotals();setFxStatus(`無法取得 Frankfurter ${code} 最新匯率，請手動輸入 FX Rate。`,'error')}
  }finally{clearTimeout(timer);state.fx.fetching=false;$('#refreshFxBtn').disabled=false}
}
async function handleCurrencyChange(){
  const code=currencyCode();updateFxPanel();
  if(code==='USD'){syncEffectivePrices();renderItems();renderCustomerSummary();schedulePreview();return}
  const cached=loadFxCache(code);
  if(cached){applyFxRate(cached.rate,cached.date,'cache',`使用上次 Frankfurter 參考匯率 · ${cached.date||''} · EUR Unit Price 刪除 cents`)}else{$('#fxRate').value='';state.fx={rate:0,date:'',source:'pending',fetching:false};syncEffectivePrices({clearCurrentOverride:true});renderItems();updateTotals();setFxStatus(`正在取得 Frankfurter v2 USD → ${code} 參考匯率…`)}
  await fetchReferenceFxRate({silent:!!cached});
}
function totals(){const code=currencyCode(),qty=state.items.reduce((a,x)=>a+x.qty,0),sub=roundCurrency(state.items.reduce((a,x)=>a+x.qty*(Number(x.unitPrice)||0),0),code),discount=roundCurrency(Math.max(0,Number($('#discountAmount').value)||0),code),total=roundCurrency(Math.max(0,sub-discount),code);return{qty,sub,discount,total}}
function updateTotals(){const t=totals();$('#totalQty').textContent=t.qty;$('#subtotal').textContent=fmt(t.sub);const discountEl=$('#discountDisplay');if(discountEl)discountEl.textContent=discountDisplay(t.discount);$('#grandTotal').textContent=fmt(t.total);$('#productCount').textContent=state.products.size;$('#customerCount').textContent=state.customers.size;$('#invoiceCount').textContent=state.items.length;$('#headerTotal').textContent=fmt(t.total);schedulePreview()}
$$('.tab').forEach(b=>b.onclick=()=>{$$('.tab').forEach(x=>x.classList.remove('active'));$$('.tab-panel').forEach(x=>x.classList.remove('active'));b.classList.add('active');$('#'+b.dataset.tab).classList.add('active');if(b.dataset.tab==='invoice'){renderCustomerSummary();renderRecallResults()}if(b.dataset.tab==='preview')renderPreview();if(b.dataset.tab==='stockSearch')renderStockSearch();requestAnimationFrame(updateBackToTopButton)});
async function readWB(file){if(typeof XLSX==='undefined')throw new Error('Excel 程式未載入');return XLSX.read(await file.arrayBuffer(),{type:'array',cellStyles:true,cellDates:true})}
async function importStockFile(f){
  const wb=await readWB(f),sheetName=wb.SheetNames[0],ws=wb.Sheets[sheetName],rows=XLSX.utils.sheet_to_json(ws,{defval:''});
  state.stockWorkbook=wb;state.stockSheetName=sheetName||'jmsdata';state.stockFileName=f.name||'jmsdata.xls';state.stockHeaders=Object.keys(rows[0]||{});state.stockAllRows=rows;state.stockRowByLot=new Map();state.stockIntegrityIssues=[];
  const fullMap=new Map();
  for(const r of rows){const lot=norm(field(r,['LOTNO'])),art=normArt(field(r,['ARTNO'])),price=Number(field(r,['PRICE']));if(!lot||!art||!Number.isFinite(price))continue;const desc=[];for(let i=1;i<=6;i++){const v=norm(field(r,[`DESC${i}`]));if(v)desc.push(v)}const p={lotNo:lot,artNo:art,price,unit:norm(field(r,['UNIT']))||'PC',article:norm(field(r,['ARTICLE']))||'',descriptions:desc,desc2:norm(field(r,['DESC2']))};fullMap.set(lot,p);state.stockRowByLot.set(lot,r);let vector=inventoryVectorFromRow(r),sum=vector.i+vector.j+vector.k+vector.l;if(sum!==vector.m||!['AVAILABLE','CONSIGNED','SOLD_ON_HAND','SOLD_DELIVERED'].includes(inventoryStatusFromVector(vector)))state.stockIntegrityIssues.push({lotNo:lot,vector});}
  if(!fullMap.size)throw new Error('找不到有效 LOTNO / ARTNO / PRICE');state.stockCatalog=fullMap;importDocumentStoreFromWorkbook(wb);rebuildInventoryMaps();const exportBtn=$('#exportJmsdataBtn');if(exportBtn)exportBtn.disabled=false;renderRecallResults();return fullMap.size;
}
async function importCustomerFile(f){
  const wb=await readWB(f),ws=wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});const map=new Map();
  for(const r of rows){const code=normCode(r[0]),company=norm(r[1]);if(!code||!company||code.includes('CUSTOMER'))continue;const raw=r[11],num=Number(raw),rate=(raw===''||!Number.isFinite(num))?0.34:num;map.set(code,{code,company,address:[r[2],r[3],r[4]].map(norm).filter(Boolean).join('\n'),rate,terms:norm(r[10])})}
  if(!map.size)throw new Error('找不到有效客戶資料');state.customers=map;return map.size;
}
async function importStoneFile(f){
  const wb=await readWB(f);const sheetName=wb.SheetNames.find(n=>n.trim().toUpperCase()==='STONE LIST')||wb.SheetNames[0];const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:''});const headerIndex=rows.findIndex(r=>r.some(v=>String(v).trim().toUpperCase()==='BREAKDOWN')&&r.some(v=>String(v).trim().toUpperCase()==='QUOTATION'));if(headerIndex<0)throw new Error('找不到 BREAKDOWN / QUOTATION 欄位');const header=rows[headerIndex].map(v=>String(v).trim().toUpperCase());const bCol=header.indexOf('BREAKDOWN'),qCol=header.indexOf('QUOTATION'),gCol=header.indexOf('GROUP');const aliases=new Map(),variantAliases=new Map(),groups=new Map();
  for(const r of rows.slice(headerIndex+1)){
    const breakdown=norm(r[bCol]),quotation=norm(r[qCol]),group=gCol>=0?norm(r[gCol]):'';if(!breakdown)continue;
    const codes=breakdown.split(/[,，]/).map(norm).filter(Boolean).map(x=>x.toUpperCase());
    const quotes=quotation.split(/[,，]/).map(norm).filter(Boolean).map(x=>x.toUpperCase());
    for(const code of codes){
      if(quotes.length){const list=variantAliases.get(code)||[];for(const q of quotes)if(!list.includes(q))list.push(q);variantAliases.set(code,list);aliases.set(code,quotes[quotes.length-1])}
      if(group)groups.set(code,group);
    }
  }
  if(!aliases.size)throw new Error('對照表沒有有效資料');state.stoneAliases=aliases;state.stoneVariantAliases=variantAliases;state.stoneGroups=groups;state.stoneMappingName=f.name;renderStockSearch();return aliases.size;
}
async function importArticleFile(f){const wb=await readWB(f),ws=wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});const map=new Map();for(const row of rows){const prefix=normCode(row[0]),description=norm(row[1]);if(!prefix||!description||prefix==='PREFIX')continue;map.set(prefix,description)}if(!map.size)throw new Error('找不到 Prefix / Article Description 對照');state.articleMap=map;state.articleMappingName=f.name;return map.size}
async function importTemplateFile(f){const buf=await f.arrayBuffer();if(typeof ExcelJS==='undefined')throw new Error('Excel 範本程式未載入');const test=new ExcelJS.Workbook();await test.xlsx.load(buf.slice(0));if(!test.worksheets.length)throw new Error('範本沒有工作表');state.invoiceTemplateBuffer=buf;state.invoiceTemplateName=f.name;return true}
function importImageFiles(files){const map=new Map();for(const f of files){if(!String(f.type).startsWith('image/')&&!/\.(jpe?g|png|webp)$/i.test(f.name))continue;const p=parseImage(f);if(!p)continue;const key=p.art+'|'+p.variant.toUpperCase(),existing=map.get(key);if(!existing||p.dup<existing.dup)map.set(key,p)}state.imageFiles=new Map();for(const p of map.values()){const arr=state.imageFiles.get(p.art)||[];arr.push({variant:p.variant,url:URL.createObjectURL(p.file),fileName:p.file.name,file:p.file});state.imageFiles.set(p.art,arr)}for(const arr of state.imageFiles.values())arr.sort((a,b)=>a.variant==='Default'?-1:b.variant==='Default'?1:a.variant.localeCompare(b.variant));return {images:[...files].filter(f=>String(f.type).startsWith('image/')||/\.(jpe?g|png|webp)$/i.test(f.name)).length,matched:state.imageFiles.size}}
$('#stockInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const count=await importStockFile(f),issues=state.stockIntegrityIssues.length;status('#stockStatus',`已匯入 ${f.name}：${count} 件貨品；Available ${state.products.size} 件；文件記錄 ${state.documentStore.invoiceHeaders.length+state.documentStore.consignmentHeaders.length} 筆${issues?`；${issues} 列 I:M 需要檢查`:''}。`,issues?'warn':'ok');setImportCollapsed('stock',true);updateTotals()}catch(err){status('#stockStatus','匯入失敗：'+err.message,'error');setImportCollapsed('stock',false)}};
$('#customerInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const count=await importCustomerFile(f);status('#customerStatus',`已匯入 ${f.name}：${count} 位客戶。`,'ok');setImportCollapsed('customer',true);updateTotals()}catch(err){status('#customerStatus','匯入失敗：'+err.message,'error');setImportCollapsed('customer',false)}};
const FALLBACK_STONE_VARIANTS=new Map([
  ['QAM',['AM']],['LBT',['LBT','L.BT']],['BTO',['BT']],['SKY',['SKY','SKY BT']],['YCT',['CT']],['GPS',['G.AM','GAM']],['GPD',['PD']],['RQZ',['RQZ']],['MG',['MG']],['PTQ',['PTR']],['AQ',['AQ']],['AMCT',['AMCT']],['PAM',['P.AM','PAM']],['MCT',['MCT']],['RGT',['RGT']],['TZ',['TZ']],['IO',['IO']],['GT',['GT']],['GTQ',['GTR']],['ALEX',['ALEX']],['KU',['KU']],['LQZ',['LQZ']],['SQZ',['SQZ']],
  ['BSA',['BSA']],['PSA',['PSA']],['GGT',['GGT']],['OSA',['OSA']],['YSA',['YSA']],['SSU',['SSU']],['RRU',['RRU']],['GEM',['GEM']],['GSA',['GSA']],['WSA',['WSA']],['ZSA',['ZSA']],['ZSP',['ZSP']],['XXX',['MULTI']],['DIA',['DIA']],
  ['WCH',['AG']],['AMZ',['AMZ']],['BCH',['BCH']],['BOX',['BO']],['GMA',['GMA']],['LAB',['LAB']],['LAP',['LAP']],['MOON',['MOON']],['OPAL',['OPAL']],['WPL',['WPL']],['RCH',['RCH']],['TE',['TE']],['TQ',['TQ']]
]);
const FALLBACK_STONE_ALIASES=new Map([...FALLBACK_STONE_VARIANTS].map(([code,variants])=>[code,variants[variants.length-1]]));
const FALLBACK_STONE_GROUPS=new Map([
  ['QAM','Purple/Pink/Rose'],['LBT','Blue'],['BTO','Blue'],['SKY','Blue'],['YCT','Yellow/Orange'],['GPS','Green'],['GPD','Green'],['RQZ','Purple/Pink/Rose'],['MG','Purple/Pink/Rose'],['PTQ','Purple/Pink/Rose'],['AQ','Blue'],['AMCT','Yellow/Orange'],['PAM','Purple/Pink/Rose'],['MCT','Yellow/Orange'],['RGT','Purple/Pink/Rose'],['TZ','Blue'],['IO','Blue'],['GT','Purple/Pink/Rose'],['GTQ','Green'],['ALEX','Green'],['KU','Purple/Pink/Rose'],['LQZ','Yellow/Orange'],['SQZ','Yellow/Orange'],
  ['BSA','Blue'],['PSA','Purple/Pink/Rose'],['GGT','Green'],['OSA','Yellow/Orange'],['YSA','Yellow/Orange'],['SSU','Blue'],['RRU','Purple/Pink/Rose'],['GEM','Green'],['GSA','Green'],['WSA','White'],['ZSA','Black'],['ZSP','Black'],['XXX','Multi'],['DIA','White'],
  ['WCH','White'],['AMZ','Blue'],['BCH','Blue'],['BOX','Black'],['GMA','Green'],['LAB','Black'],['LAP','Blue'],['MOON','White'],['OPAL','White'],['WPL','White'],['RCH','Purple/Pink/Rose'],['TE','Yellow/Orange'],['TQ','Blue']
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
function activeStoneVariantAliases(){return state.stoneVariantAliases.size?state.stoneVariantAliases:FALLBACK_STONE_VARIANTS}
function activeStoneGroups(){return state.stoneGroups.size?state.stoneGroups:FALLBACK_STONE_GROUPS}
function stoneImageAliasesForCode(code){const c=String(code||'').toUpperCase(),list=activeStoneVariantAliases().get(c);if(Array.isArray(list)&&list.length)return list;const single=activeStoneAliases().get(c);return single?[String(single).toUpperCase()]:[]}
function stoneGroupForCode(code){return norm(activeStoneGroups().get(String(code||'').toUpperCase())).toUpperCase()}
$('#stoneMappingInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const count=await importStoneFile(f);status('#stoneMappingStatus',`已匯入 ${f.name}：${count} 個石種代碼對照。`,'ok');setImportCollapsed('stone',true);for(const item of state.items)if(!item.customImage)applyAutoImageMatch(item);renderItems()}catch(err){status('#stoneMappingStatus','匯入失敗：'+(err.message||err),'error');setImportCollapsed('stone',false)}};
$('#articleMappingInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const count=await importArticleFile(f);status('#articleMappingStatus',`已匯入 ${f.name}：${count} 個 Article 對照。`,'ok');setImportCollapsed('article',true);schedulePreview()}catch(err){status('#articleMappingStatus','匯入失敗：'+(err.message||err),'error');setImportCollapsed('article',false)}};
$('#invoiceTemplateInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{await importTemplateFile(f);status('#invoiceTemplateStatus',`已匯入 ${f.name}；匯出文件時會套用此範本。`,'ok');setImportCollapsed('template',true)}catch(err){state.invoiceTemplateBuffer=null;status('#invoiceTemplateStatus','匯入失敗：'+(err.message||err),'error');setImportCollapsed('template',false)}};

function imageArtAliasWithoutSuffixDot(art){
  const value=String(art||'').toUpperCase();
  return /\.[A-Z0-9]+$/.test(value)?value.replace(/\.([A-Z0-9]+)$/,'$1'):'';
}
function imageStemMatchesPrefix(stemUpper,prefix){return stemUpper===prefix||stemUpper.startsWith(prefix+' ')}
function parseImage(file){
  const stem=file.name.replace(/\.[^.]+$/,'').trim(),stemUpper=stem.toUpperCase();
  const arts=[...new Set([...state.stockCatalog.values(),...state.products.values(),...state.inventoryHistory.values()].map(x=>normArt(x.artNo)).filter(Boolean))].sort((a,b)=>b.length-a.length);
  let art='',prefix='';
  for(const candidate of arts){if(imageStemMatchesPrefix(stemUpper,candidate)){art=candidate;prefix=candidate;break}}
  if(!art){
    for(const candidate of arts){const alias=imageArtAliasWithoutSuffixDot(candidate);if(alias&&imageStemMatchesPrefix(stemUpper,alias)){art=candidate;prefix=alias;break}}
  }
  if(!art)return null;
  let variant=stem.slice(prefix.length).trim().replace(/\s*\(\d+\)$/,'').trim()||'Default';
  const dup=(stem.match(/\((\d+)\)$/)||[])[1];
  return{art,variant,dup:dup?Number(dup):0,file}
}
$('#imageFolderInput').onchange=e=>{const result=importImageFiles(e.target.files);for(const item of state.items)if(!item.customImage)applyAutoImageMatch(item);status('#imageStatus',`已選擇圖片 Folder：${result.images} 張圖片，配對 ${result.matched} 個款號。`,'ok');setImportCollapsed('images',true);renderItems()};
$('#exhibitionPackageInput').onchange=async e=>{
  const files=[...e.target.files];if(!files.length)return;
  const root=(files[0].webkitRelativePath||'').split('/')[0]||'Exhibition Package';state.packageName=root;
  const excels=files.filter(f=>/\.(xlsx?|xls)$/i.test(f.name));
  const byName=(re)=>excels.find(f=>re.test(f.name));
  let stock=byName(/^jmsdata(?:\s.*)?\.(xls|xlsx)$/i)||byName(/jmsdata/i);
  let customer=byName(/customer/i);
  let stone=byName(/stone\s*list.*shape.*cutting/i);
  let article=byName(/article\s*mapping/i);
  let template=byName(/invoice\s*(master\s*)?template/i)||byName(/template/i);
  const imageFiles=files.filter(f=>String(f.type).startsWith('image/')||/\.(jpe?g|png|webp)$/i.test(f.name));
  const lines=[];let errors=[];
  status('#packageStatus',`正在讀取 ${root}…`);
  try{
    if(!stock)errors.push('找不到 jmsdata.xls / jmsdata.xlsx');else{const n=await importStockFile(stock);lines.push(`✓ ${stock.name} · ${n} 件貨品 · Available ${state.products.size} · Invoice/Consignment 記錄 ${state.documentStore.invoiceHeaders.length+state.documentStore.consignmentHeaders.length} 筆 · ${new Date(stock.lastModified).toLocaleString('zh-HK')}`)}
    if(!customer)errors.push('找不到客戶 Excel');else{const n=await importCustomerFile(customer);lines.push(`✓ ${customer.name} · ${n} 位客戶`)}
    if(!stone)errors.push('找不到 Stone List & Shape & Cutting.xlsx');else{const n=await importStoneFile(stone);lines.push(`✓ ${stone.name} · ${n} 個石種代碼`)}
    if(!template)errors.push('找不到 Invoice Template.xlsx');else{await importTemplateFile(template);lines.push(`✓ ${template.name} · Template 已載入`)}
    if(article){const n=await importArticleFile(article);lines.push(`✓ ${article.name} · ${n} 個 Article 對照`)}else lines.push('○ Article Mapping 未提供（不顯示 Article）');
    if(!imageFiles.length)errors.push('找不到 Pictures 內的圖片');else{const r=importImageFiles(imageFiles);for(const item of state.items)if(!item.customImage)applyAutoImageMatch(item);lines.push(`✓ Pictures · ${r.images} 張圖片 · 配對 ${r.matched} 個款號`)}
    $('#packageSummary').innerHTML=lines.map(x=>`<div>${esc(x)}</div>`).join('')+(errors.length?`<div class="package-errors">${errors.map(x=>'✕ '+esc(x)).join('<br>')}</div>`:'');
    if(errors.length)status('#packageStatus',`${root} 未完整載入，請補回缺少的資料。`,'error');else{status('#packageStatus',`${root} 已完成匯入，可以開始建立文件。`,'ok');document.querySelector('.advanced-imports').open=false}
    status('#stockStatus',stock?`已由資料包匯入 ${stock.name}：${state.products.size} 件貨品。`:'尚未匯入倉存。',stock?'ok':'error');
    status('#customerStatus',customer?`已由資料包匯入 ${customer.name}：${state.customers.size} 位客戶。`:'尚未匯入客戶。',customer?'ok':'error');
    status('#imageStatus',imageFiles.length?`已由資料包匯入 Pictures：${imageFiles.length} 張圖片，配對 ${state.imageFiles.size} 個款號。`:'尚未匯入圖片。',imageFiles.length?'ok':'error');
    if(stone)status('#stoneMappingStatus',`已由資料包匯入 ${stone.name}。`,'ok');
    if(template)status('#invoiceTemplateStatus',`已由資料包匯入 ${template.name}。`,'ok');
    if(article)status('#articleMappingStatus',`已由資料包匯入 ${article.name}。`,'ok');
    updateTotals();renderItems();
  }catch(err){console.error(err);status('#packageStatus','資料包匯入失敗：'+(err.message||err),'error')}
};

function searchCustomers(q){const s=norm(q).toUpperCase(),c=normCode(q);return [...state.customers.values()].filter(x=>x.code.includes(c)||x.company.toUpperCase().includes(s)).slice(0,10)}
function showMatches(){const box=$('#customerMatches'),m=searchCustomers($('#customerSearch').value);box.innerHTML='';if(!m.length){box.innerHTML='<div class="notice">找不到客戶。</div>';return}m.forEach(c=>{const b=document.createElement('button');b.className='customer-match';b.innerHTML=`<span><strong>${esc(c.code)} · ${esc(c.company)}</strong><small>${esc(c.address).replace(/\n/g,' · ')}</small></span><span>${c.rate}</span>`;b.onclick=()=>selectCustomer(c);box.appendChild(b)})}
function selectCustomer(c){$('#customerCode').value=c.code;$('#customerName').value=c.company;$('#customerAddress').value=c.address;$('#salesRate').value=c.rate;$('#customerTerms').value=c.terms;$('#customerMatches').innerHTML='';$('#customerSearch').value='';reprice();renderCustomerSummary();schedulePreview()}
$('#findCustomerBtn').onclick=showMatches;
let customerSearchTimer=null;
$('#customerSearch').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();showMatches()}};
$('#customerSearch').oninput=e=>{
  clearTimeout(customerSearchTimer);
  const q=norm(e.target.value);
  if(q.length<2){$('#customerMatches').innerHTML='';return}
  customerSearchTimer=setTimeout(showMatches,120);
};
function renderCustomerSummary(){const code=$('#customerCode').value,name=$('#customerName').value,currency=currencyCode(),fx=currency==='USD'?'':(currentFxRate()>0?` · FX 1 USD = ${currentFxRate()} ${currency}`:' · FX 未設定');$('#selectedCustomerSummary').innerHTML=code||name?`<strong>${esc(code)} · ${esc(name)}</strong><span>Sales Rate ${esc($('#salesRate').value)} · ${esc(currency)}${esc(fx)}</span>`:'尚未選擇客戶。'}
function variantContainsStone(variant,wanted){
  const normalizeStonePhrase=v=>String(v||'').toUpperCase().replace(/[^A-Z0-9]+/g,' ').trim().replace(/\s+/g,' ');
  const target=normalizeStonePhrase(wanted);if(!target)return false;
  return String(variant||'').split('+').some(part=>{
    const phrase=normalizeStonePhrase(part.replace(/\([^)]*\)/g,''));return phrase===target||(` ${phrase} `).includes(` ${target} `);
  });
}
const MULTI_COLOR_GROUP_THRESHOLD=3;
const IMAGE_STONE_EXCLUSIONS=new Set(['CDM','DIA']);
function imageStoneCodesForProduct(p){
  return stoneCodesForProduct(p).filter(code=>!IMAGE_STONE_EXCLUSIONS.has(String(code||'').toUpperCase()));
}
function normalDesiredStoneVariants(p){
  const out=[];for(const code of imageStoneCodesForProduct(p)){for(const variant of stoneImageAliasesForCode(code)){const v=String(variant||'').toUpperCase();if(v&&!out.includes(v))out.push(v)}}return out;
}
function desiredStoneVariants(p){return normalDesiredStoneVariants(p)}
function productColorGroups(p){
  const groups=[];for(const code of imageStoneCodesForProduct(p)){const group=stoneGroupForCode(code);if(group&&!groups.includes(group))groups.push(group)}return groups;
}
function isMultiColorProduct(p){
  const groups=productColorGroups(p);return groups.includes('MULTI')||groups.length>=MULTI_COLOR_GROUP_THRESHOLD;
}
function originalMetalToken(p){
  const desc1=String((p?.descriptions||[])[0]||'').toUpperCase().replace(/\s+/g,'');
  const m=desc1.match(/([YWR])(750|585)\b/);
  if(!m)return'';
  return `${m[2]==='750'?'18K':'14K'}${m[1]}`;
}
function imageMetalToken(variant){
  const u=String(variant||'').toUpperCase().replace(/\s+/g,'');
  const m=u.match(/(?:^|[^A-Z0-9])((?:18|14)K[YWR])(?:$|[^A-Z0-9])/);
  return m?m[1]:'';
}
function imageStoneSignature(variant){
  return String(variant||'').toUpperCase().replace(/\([^)]*\)/g,'').trim().replace(/\s*\+\s*/g,'+').replace(/\s+/g,' ');
}
function imageStoneTokenCount(variant){const signature=imageStoneSignature(variant);return signature?signature.split('+').map(x=>x.trim()).filter(Boolean).length:0}
function isMultiImageVariant(variant){return imageStoneSignature(variant).split('+').some(x=>x.trim()==='MULTI')}
function variantMatchesProductStoneCode(variant,code){return stoneImageAliasesForCode(code).some(alias=>variantContainsStone(variant,alias))}
function matchedProductStoneIndexes(variant,codes){const out=[];codes.forEach((code,index)=>{if(variantMatchesProductStoneCode(variant,code))out.push(index)});return out}
function chooseImageMatch(p){
  const imgs=state.imageFiles.get(p?.artNo)||[];
  if(!imgs.length)return{variant:'Default',grayscale:false,stoneMatched:false,colorMatched:false};
  const codes=imageStoneCodesForProduct(p),metal=originalMetalToken(p),multiColor=isMultiColorProduct(p);
  const metalScore=img=>{const t=imageMetalToken(img.variant);return metal&&t===metal?2:!t?1:0};
  const regularImgs=imgs.filter(x=>x.variant!=='Default'&&!isMultiImageVariant(x.variant));
  const scoreRegular=(img,index,requireAll=false)=>{
    const matchedIndexes=matchedProductStoneIndexes(img.variant,codes),matchedCount=matchedIndexes.length;if(!matchedCount||requireAll&&matchedCount!==codes.length)return null;
    const tokenCount=imageStoneTokenCount(img.variant),allMatched=matchedCount===codes.length;
    let stoneScore=matchedCount*1000+(allMatched?5000:0)-Math.abs(tokenCount-matchedCount)*80;
    stoneScore+=Math.max(0,40-(matchedIndexes[0]||0)*5);
    return{img,index,matchedCount,allMatched,stoneScore,metalScore:metalScore(img),tokenCount};
  };
  const sortScored=(a,b)=>b.stoneScore-a.stoneScore||b.metalScore-a.metalScore||a.tokenCount-b.tokenCount||a.img.variant.length-b.img.variant.length||a.index-b.index;
  const exactCombos=regularImgs.map((img,index)=>scoreRegular(img,index)).filter(x=>x&&x.allMatched&&codes.length>1).sort(sortScored);
  if(exactCombos.length){const best=exactCombos[0];return{variant:best.img.variant,grayscale:false,stoneMatched:true,colorMatched:best.metalScore===2}}
  if(multiColor){
    const multi=imgs.filter(x=>isMultiImageVariant(x.variant)).map((img,index)=>({img,index,metalScore:metalScore(img)})).sort((a,b)=>b.metalScore-a.metalScore||a.img.variant.length-b.img.variant.length||a.index-b.index)[0];
    if(multi)return{variant:multi.img.variant,grayscale:false,stoneMatched:true,colorMatched:multi.metalScore===2};
    const fallback=[...imgs].sort((a,b)=>metalScore(b)-metalScore(a)||(a.variant==='Default'?-1:b.variant==='Default'?1:0)||a.variant.length-b.variant.length)[0];
    return{variant:fallback.variant,grayscale:true,stoneMatched:false,colorMatched:metalScore(fallback)===2};
  }
  const normal=regularImgs.map((img,index)=>scoreRegular(img,index)).filter(Boolean).sort(sortScored)[0];
  if(normal)return{variant:normal.img.variant,grayscale:false,stoneMatched:true,colorMatched:normal.metalScore===2};
  if(codes.length){
    const fallback=[...imgs].sort((a,b)=>metalScore(b)-metalScore(a)||(a.variant==='Default'?-1:b.variant==='Default'?1:0)||a.variant.length-b.variant.length)[0];
    return{variant:fallback.variant,grayscale:true,stoneMatched:false,colorMatched:metalScore(fallback)===2};
  }
  const fallback=imgs.find(x=>x.variant==='Default')||imgs[0];
  return{variant:fallback.variant,grayscale:false,stoneMatched:false,colorMatched:metalScore(fallback)===2};
}
function chooseVariant(p){return chooseImageMatch(p).variant}
function applyAutoImageMatch(item){
  if(!item)return;
  const match=chooseImageMatch(item);item.imageVariant=match.variant;item.imageGrayscale=!!match.grayscale;item.imageAutoMatched=true;
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
function addProductToDocument(p,{fromSearch=false}={}){
  if(!p)return false;const lot=String(p.lotNo);
  if(state.items.some(x=>x.lotNo===lot)){status('#addMessage',`LOTNO ${lot} 已在 ${documentLabels().short}。`,'error');return false}
  const rate=Number($('#salesRate').value)||0,usdUnitPrice=Math.ceil((Number(p.price)||0)*rate),match=chooseImageMatch(p);const item={...productSnapshot(p),id:Date.now()+Math.random(),seq:state.items.length+1,qty:1,usdUnitPrice,currencyPrices:{},quote14kCurrencyPrices:{},unitPrice:0,imageVariant:match.variant,imageGrayscale:!!match.grayscale,imageAutoMatched:true};state.items.push(item);item.unitPrice=convertedFromUsd(effectiveUsdPrice(item));
  status('#addMessage',`已加入 ${p.artNo} / LOTNO ${lot}`,'ok');renderItems();if(fromSearch)renderStockSearch();return true;
}
function addLot(raw){
  const lot=normalizeLotInput(raw);
  if(!lot){status('#addMessage','請輸入 LOTNO。','error');refocusLotInput();return false}
  let p=state.documentType==='quotation'?state.stockCatalog.get(lot):state.products.get(lot);
  if(p&&state.documentType==='quotation'){const row=state.stockRowByLot.get(lot),itemStatus=inventoryStatusFromVector(inventoryVectorFromRow(row||{}));if(!['AVAILABLE','CONSIGNED','SOLD_ON_HAND','SOLD_DELIVERED'].includes(itemStatus))p=null}
  if(!p){const message=state.documentType==='quotation'?`找不到狀態有效的 LOTNO ${lot}。`:`找不到可售 LOTNO ${lot}。可到「配套搜尋」查看已售／寄賣狀態。`;status('#addMessage',message,'error');refocusLotInput(true);return false}
  if(!addProductToDocument(p))return false;$('#lotInput').value='';setTimeout(()=>{$('#invoiceItems').scrollTo({top:0,behavior:'smooth'});refocusLotInput()},50);return true;
}
$('#addLotBtn').onclick=()=>addLot($('#lotInput').value);
$('#lotInput').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();addLot(e.target.value)}};
$('#lotInput').oninput=e=>{e.target.value=e.target.value.replace(/[，,。\.\-–—_]/g,' ')};
$('#lotInput').onfocus=e=>setTimeout(()=>e.target.select(),50);
function getImg(item){
  if(item?.customImage?.file&&item.customImage.url)return {...item.customImage,grayscale:false};
  const arr=state.imageFiles.get(item.artNo)||[];
  const selected=arr.find(x=>x.variant===item.imageVariant)||arr[0];
  return selected?{...selected,grayscale:!!item?.imageGrayscale}:selected;
}
function placeholder(t){return `data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="200" height="200"><rect width="100%" height="100%" fill="#f1f5f9"/><text x="50%" y="50%" text-anchor="middle" dominant-baseline="middle" font-family="Arial" font-size="18" fill="#64748b">${t}</text></svg>`)}`}
function revokeCustomImage(item){
  const url=item?.customImage?.url;
  if(url&&String(url).startsWith('blob:')){try{URL.revokeObjectURL(url)}catch{}}
  if(item)item.customImage=null;
}
function releaseCustomImages(items=state.items){for(const item of items||[])revokeCustomImage(item)}
async function prepareItemImageFile(file,item,source='upload'){
  if(!file||!item)return;
  if(!String(file.type||'').startsWith('image/')&&!/\.(jpe?g|png|webp|heic|heif)$/i.test(file.name||''))throw new Error('請選擇圖片檔案。');
  let output=file;
  try{
    const sourceUrl=URL.createObjectURL(file);
    try{
      const img=await new Promise((resolve,reject)=>{const i=new Image();i.onload=()=>resolve(i);i.onerror=reject;i.src=sourceUrl});
      const iw=img.naturalWidth||img.width,ih=img.naturalHeight||img.height,maxSide=1400;
      const scale=Math.min(1,maxSide/Math.max(iw,ih));
      const w=Math.max(1,Math.round(iw*scale)),h=Math.max(1,Math.round(ih*scale));
      const canvas=document.createElement('canvas');canvas.width=w;canvas.height=h;
      const ctx=canvas.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,w,h);ctx.drawImage(img,0,0,w,h);
      const blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/jpeg',.86));
      if(blob){
        const safeArt=String(item.artNo||'item').replace(/[^0-9A-Za-z._-]+/g,'_');
        const safeLot=String(item.lotNo||'').replace(/[^0-9A-Za-z._-]+/g,'_');
        output=new File([blob],`${safeArt}_${safeLot}_${source==='camera'?'camera':'upload'}.jpg`,{type:'image/jpeg',lastModified:Date.now()});
      }
    }finally{URL.revokeObjectURL(sourceUrl)}
  }catch(err){
    console.warn('圖片壓縮失敗，改用原圖。',err);
  }
  revokeCustomImage(item);
  item.customImage={file:output,url:URL.createObjectURL(output),fileName:output.name||file.name||'image',source};
  renderItems();schedulePreview();
}

function renderItems(){
  const box=$('#invoiceItems');box.innerHTML='';
  if(!state.items.length){state.editingItemId=null;box.className='invoice-items empty-state';box.textContent='尚未加入貨品。';updateTotals();updateBackToTopButton();return}
  box.className='invoice-items';
  displayItems().forEach(item=>{
    const node=$('#itemTemplate').content.firstElementChild.cloneNode(true);node.dataset.itemId=String(item.id);
    $('.item-seq',node).textContent=item.seq;$('.item-artno',node).textContent=item.artNo;$('.item-lot',node).textContent=`LOTNO ${item.lotNo}`;
    const nonEmptyDescriptions=effectiveDescriptions(item).map(x=>norm(x)).filter(Boolean);
    $('.item-desc',node).textContent=nonEmptyDescriptions.slice(0,2).join('\n');
    $('.item-full-desc',node).textContent=nonEmptyDescriptions.join('\n');
    const usd=baseUsdPrice(item),code=currencyCode(),note=$('.item-price-note',node);
    if(quote14KMode()){const p18=parse18KDesc1((item.descriptions||[])[0]),m=goldMetrics();if(p18&&m){const w14=roundUpWeight05(p18.weight*K14_WEIGHT_FACTOR),diff=goldDifferenceUsd(item),u14=effectiveUsdPrice(item);note.textContent=code==='USD'?`14K：${p18.weight.toFixed(2)}g → ${w14.toFixed(2)}g · ${fmt(usd,'USD')} − ${fmt(diff,'USD')} → ${fmt(u14,'USD')}`:(fxPricingReady()?`14K：${p18.weight.toFixed(2)}g → ${w14.toFixed(2)}g · ${fmt(u14,'USD')} × ${currentFxRate()} → ${fmt(item.unitPrice,code)}`:`14K：${w14.toFixed(2)}g · 等待 FX Rate`);note.classList.add('quote14-price-note');const badge=document.createElement('span');badge.className='quote14-badge';badge.textContent='14K';$('.item-artno',node).after(badge)}else note.textContent='14K 參考報價：等待 Kitco Ask／DESC1 無法辨識 18K 金重';}
    else note.textContent=code==='USD'?`${item.price}u × ${Number($('#salesRate').value)||0} → ${fmt(usd,'USD')}`:(fxPricingReady()?`${item.price}u × ${Number($('#salesRate').value)||0} → ${fmt(usd,'USD')} × ${currentFxRate()} → ${fmt(item.unitPrice,code)}`:`${item.price}u × ${Number($('#salesRate').value)||0} → ${fmt(usd,'USD')} · 等待 FX Rate`);
    const thumbImg=getImg(item),thumb=$('.item-thumb',node);thumb.src=thumbImg?.url||placeholder(item.artNo);thumb.classList.toggle('grayscale-image',!!thumbImg?.grayscale);
    const controls=$('.item-controls',node),toggle=$('.item-edit-toggle',node);
    toggle.onclick=()=>{
      const isOpen=String(state.editingItemId??'')===String(item.id);
      setEditingItem(box,isOpen?null:item.id);
    };
    const sel=$('.variant-select',node),arr=state.imageFiles.get(item.artNo)||[],custom=item.customImage;
    if(custom){
      const o=document.createElement('option');o.value='__custom__';o.textContent=custom.source==='camera'?'圖片：即時拍照':'圖片：上傳';o.selected=true;sel.appendChild(o);
    }
    if(arr.length){
      arr.forEach(x=>{const o=document.createElement('option');o.value=x.variant;o.textContent='圖片：'+x.variant;o.selected=!custom&&!item.imageGrayscale&&x.variant===item.imageVariant;sel.appendChild(o)});
      const mono=document.createElement('option');mono.value='__grayscale__';mono.textContent='圖片：黑白';mono.selected=!custom&&!!item.imageGrayscale;sel.appendChild(mono);
      sel.disabled=false;
    }else if(!custom){sel.innerHTML='<option>沒有圖片</option>';sel.disabled=true}
    sel.onchange=e=>{
      const value=e.target.value;
      if(value==='__custom__')return;
      if(custom)revokeCustomImage(item);
      if(value==='__grayscale__'){
        if(!item.imageVariant||!arr.some(x=>x.variant===item.imageVariant))item.imageVariant=chooseImageMatch(item).variant;
        item.imageGrayscale=true;item.imageAutoMatched=false;
      }else{
        item.imageVariant=value;item.imageGrayscale=false;item.imageAutoMatched=false;
      }
      renderItems();schedulePreview();
    };
    const uploadBtn=$('.upload-image-btn',node),cameraBtn=$('.camera-image-btn',node),uploadInput=$('.upload-image-input',node),cameraInput=$('.camera-image-input',node),restoreBtn=$('.restore-db-image',node),customNote=$('.custom-image-note',node);
    uploadBtn.onclick=()=>uploadInput.click();
    cameraBtn.onclick=()=>cameraInput.click();
    uploadInput.onchange=async e=>{const f=e.target.files?.[0];if(!f)return;uploadBtn.disabled=true;uploadBtn.textContent='處理中…';try{await prepareItemImageFile(f,item,'upload')}catch(err){alert(err.message||err)}finally{uploadBtn.disabled=false;uploadBtn.textContent='上傳圖片';e.target.value=''}};
    cameraInput.onchange=async e=>{const f=e.target.files?.[0];if(!f)return;cameraBtn.disabled=true;cameraBtn.textContent='處理中…';try{await prepareItemImageFile(f,item,'camera')}catch(err){alert(err.message||err)}finally{cameraBtn.disabled=false;cameraBtn.textContent='📷 即時拍照';e.target.value=''}};
    if(custom){
      restoreBtn.classList.remove('hidden');restoreBtn.textContent=arr.length?'使用原資料庫圖片':'移除現場圖片';
      customNote.classList.remove('hidden');customNote.textContent=custom.source==='camera'?'目前使用：即時拍照圖片':'目前使用：上傳圖片';
      restoreBtn.onclick=()=>{revokeCustomImage(item);renderItems();schedulePreview()};
    }
    $('.qty-input',node).value=item.qty;$('.price-input',node).value=item.unitPrice;
    $('.qty-input',node).onchange=e=>{item.qty=Math.max(1,Number(e.target.value)||1);updateTotals()};
    $('.price-input',node).onchange=e=>{const code=currencyCode(),raw=Math.max(0,Number(e.target.value)||0),value=code==='EUR'?Math.round(raw):(code==='USD'&&quote14KMode()?Math.ceil(raw-1e-9):roundUnitPrice(raw,code)),overrides=activePriceOverrides(item);overrides[code]=value;item.unitPrice=value;if(code==='USD'&&!quote14KMode())item.usdUnitPrice=value;e.target.value=value;updateTotals()};
    $('.delete-item',node).onclick=()=>{if(confirm(`刪除 ${item.artNo}？`)){revokeCustomImage(item);if(String(state.editingItemId??'')===String(item.id))state.editingItemId=null;state.items=state.items.filter(x=>x.id!==item.id);normalizeItemSequence();renderItems()}};
    box.appendChild(node)
  });
  setEditingItem(box,state.editingItemId,{syncSort:false});
  installItemSorting(box);
  updateTotals();updateBackToTopButton()
}
function setEditingItem(box,itemId,{syncSort=true}={}){
  if(!box)return;
  const wanted=itemId==null?null:String(itemId);
  const exists=wanted&&[...box.querySelectorAll('.invoice-item')].some(node=>node.dataset.itemId===wanted);
  state.editingItemId=exists?wanted:null;
  box.querySelectorAll('.invoice-item').forEach(node=>{
    const open=!!state.editingItemId&&node.dataset.itemId===String(state.editingItemId);
    node.classList.toggle('editing',open);
    const controls=$('.item-controls',node);if(controls)controls.classList.toggle('open',open);
    const toggle=$('.item-edit-toggle',node);if(toggle)toggle.textContent=open?'完成':'編輯';
  });
  if(syncSort)syncSortAvailability(box);
}
function syncSortAvailability(box=$('#invoiceItems')){
  if(!box)return;
  const editing=!!box.querySelector('.invoice-item.editing');
  box.classList.toggle('sorting-locked',editing);
  box.querySelectorAll('.drag-handle').forEach(handle=>{
    handle.disabled=editing;
    handle.setAttribute('aria-disabled',String(editing));
    handle.title=editing?'請先按「完成」收起編輯，再拖曳排序':'長按 ≡ 約 0.3 秒，再上下拖曳排序';
  });
  if(state.sortable&&typeof state.sortable.option==='function')state.sortable.option('disabled',editing);
}
function installItemSorting(box){
  if(state.sortable){try{state.sortable.destroy()}catch{}state.sortable=null}
  if(!box||!window.Sortable){
    if(box)box.querySelectorAll('.drag-handle').forEach(handle=>{handle.disabled=true;handle.setAttribute('aria-disabled','true');handle.title='拖曳排序元件未能載入'});
    return;
  }
  state.sortable=Sortable.create(box,{
    animation:160,
    handle:'.drag-handle',
    draggable:'.invoice-item',
    forceFallback:true,
    fallbackOnBody:true,
    // iPhone guard rails: require an intentional long press and real movement
    // before drag starts, then swap only after crossing roughly half a row.
    delay:300,
    delayOnTouchOnly:true,
    fallbackTolerance:10,
    touchStartThreshold:10,
    swapThreshold:0.5,
    invertSwap:false,
    scroll:true,
    scrollSensitivity:42,
    scrollSpeed:7,
    chosenClass:'sort-chosen',
    ghostClass:'sort-ghost',
    dragClass:'sort-drag',
    fallbackClass:'sort-fallback',
    onStart:()=>box.classList.add('sorting-active'),
    onEnd:()=>{
      box.classList.remove('sorting-active');
      const byId=new Map(state.items.map(item=>[String(item.id),item]));
      const visible=[...box.querySelectorAll('.invoice-item')].map(node=>byId.get(node.dataset.itemId)).filter(Boolean);
      if(visible.length!==state.items.length){renderItems();return}
      // The working list is newest/highest Item No. at the top, while formal output is 1 → N.
      // Reverse the visible order back into formal order, then rebuild sequence numbers.
      state.items=[...visible].reverse();
      state.items.forEach((item,i)=>item.seq=i+1);
      box.querySelectorAll('.invoice-item').forEach(node=>{
        const item=byId.get(node.dataset.itemId),seq=$('.item-seq',node);
        if(item&&seq)seq.textContent=item.seq;
      });
      schedulePreview();
      syncSortAvailability(box);
    }
  });
  syncSortAvailability(box);
}
$('#scrollLatestBtn').onclick=()=>$('#invoiceItems').scrollTo({top:0,behavior:'smooth'});$('#clearInvoiceBtn').onclick=()=>{if(confirm('清空目前 Invoice？')){releaseCustomImages();state.items=[];renderItems()}};
function reprice(){const r=Number($('#salesRate').value)||0;state.items.forEach(x=>{x.usdUnitPrice=Math.ceil(x.price*r);x.currencyPrices={};x.quote14kCurrencyPrices={}});syncEffectivePrices();renderItems()}$('#salesRate').onchange=reprice;$('#currency').onchange=handleCurrencyChange;$('#refreshFxBtn').onclick=()=>fetchReferenceFxRate();let fxInputTimer=null;$('#fxRate').oninput=e=>{clearTimeout(fxInputTimer);fxInputTimer=setTimeout(()=>{const n=Number(e.target.value);if(Number.isFinite(n)&&n>0){state.fx={rate:n,date:'',source:'manual',fetching:false};syncEffectivePrices({clearCurrentOverride:true});renderItems();renderCustomerSummary();setFxStatus('使用手動 FX Rate · EUR Unit Price 四捨五入至整數。','warn')}else{syncEffectivePrices({clearCurrentOverride:true});renderItems();setFxStatus('請輸入大於 0 的 FX Rate。','error')}},160)};$('#discountAmount').oninput=updateTotals;['invoiceNo','invoiceDate','shipmentMethod','customerCode','customerName','customerAddress','customerTerms','remark'].forEach(id=>$('#'+id)?.addEventListener('input',schedulePreview));
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
  return ({USD:'US DOLLARS',EUR:'EUROS',GBP:'BRITISH POUNDS',CNY:'CHINESE YUAN',JPY:'JAPANESE YEN',HKD:'HONG KONG DOLLARS'})[String(code||'').toUpperCase()]||String(code||'').toUpperCase();
}
function renderPreview(){
  const t=totals();
  const rows=formalItems().map((x,i)=>{const selected=getImg(x),img=selected?.url||placeholder('No Image'),gray=selected?.grayscale?' grayscale-image':'';return `<tr><td>${i+1}</td><td>Lot.No. : ${esc(x.lotNo)}<br>${esc(x.artNo)}</td><td>${effectiveDescriptions(x).filter(Boolean).map(esc).join('<br>')}</td><td class="preview-picture"><img class="${gray.trim()}" src="${esc(img)}" alt="${esc(x.artNo)}"></td><td class="qty-cell">${x.qty}</td><td class="unit-cell">${esc(x.unit)}</td><td class="num">${fmt(x.unitPrice)}</td><td class="num">${fmt(x.qty*x.unitPrice)}</td></tr>`}).join('');
  $('#invoiceDocument').innerHTML=`<div class="letterhead"><h2>UNIVERSE GEMS &amp; JEWELLERY CO.</h2><p>UNIT 11-12, 10/F., FU HANG INDUSTRIAL BUILDING, NO. 1 HOK YUEN STREET EAST,<br>HUNG HOM, KOWLOON, HONG KONG · TEL : (852) 2363 5409 · FAX : (852) 2765 0343</p></div><div class="doc-title">${documentLabels().title}</div><div class="doc-grid screen-preview"><div class="doc-meta">No. : <strong>${esc($('#invoiceNo').value)}</strong><br>${documentLabels().date} : ${esc(englishInvoiceDate($('#invoiceDate').value))}<br>Shipment Method : ${esc($('#shipmentMethod').value)}<br>Currency : ${esc($('#currency').value)}<br><br>Customer : <strong>${esc($('#customerName').value)}</strong><br>${esc($('#customerAddress').value).replace(/\n/g,'<br>')}</div><div class="doc-meta print-only print-banker"><strong>Vendor's Banker</strong><br>The Hong Kong &amp; Shanghai Banking Corporation Ltd.<br>Address : 41 Ma Tau Wai Road,Hung Hom,Kowloon,Hong Kong<br>A/C # : 012-593570-001<br>A/C Name : Universe Gems &amp; Jewellery Co.</div></div><table class="doc-table"><thead><tr><th>No.</th><th>Article No.</th><th>Description</th><th>Picture</th><th class="qty-head">Quantity</th><th class="unit-head">Unit</th><th class="num">Unit Price</th><th class="num amount-head"><span>Amount</span><small>F.O.B. Value</small></th></tr></thead><tbody>${rows}</tbody></table><div class="doc-footer"><div class="doc-totals"><div><span>Total Quantity :</span><strong>${t.qty}</strong></div><div><span>Sub Total:</span><strong>${fmt(t.sub)}</strong></div><div><span>Discount:</span><strong>${discountDisplay(t.discount)}</strong></div><div class="total"><span>Total : (${esc(currencyCode())})</span><strong>${fmt(t.total)}</strong></div></div><p class="remark-preview"><strong>Remark :</strong><br>${esc($('#remark').value).replace(/\n/g,'<br>')}</p></div>`;
}


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
function grayscaleCanvas(ctx,width,height){
  const imageData=ctx.getImageData(0,0,width,height),d=imageData.data;
  for(let i=0;i<d.length;i+=4){const y=Math.round(.299*d[i]+.587*d[i+1]+.114*d[i+2]);d[i]=d[i+1]=d[i+2]=y}
  ctx.putImageData(imageData,0,0);
}
async function imageFileToJpegDataUrl(file,maxSide=620,quality=.82,grayscale=false){
  if(!file)return null;
  const source=URL.createObjectURL(file);
  try{
    const img=await new Promise((resolve,reject)=>{const i=new Image();i.onload=()=>resolve(i);i.onerror=reject;i.src=source});
    const scale=Math.min(1,maxSide/Math.max(img.naturalWidth||img.width,img.naturalHeight||img.height));
    const w=Math.max(1,Math.round((img.naturalWidth||img.width)*scale));
    const h=Math.max(1,Math.round((img.naturalHeight||img.height)*scale));
    const canvas=document.createElement('canvas');canvas.width=w;canvas.height=h;
    const ctx=canvas.getContext('2d');ctx.fillStyle='#ffffff';ctx.fillRect(0,0,w,h);ctx.drawImage(img,0,0,w,h);if(grayscale)grayscaleCanvas(ctx,w,h);
    return canvas.toDataURL('image/jpeg',quality);
  }finally{URL.revokeObjectURL(source)}
}
async function imageFileToJpegAsset(file,maxSide=700,quality=.84,grayscale=false){
  if(!file)return null;
  const source=URL.createObjectURL(file);
  try{
    const img=await new Promise((resolve,reject)=>{const i=new Image();i.onload=()=>resolve(i);i.onerror=reject;i.src=source});
    const scale=Math.min(1,maxSide/Math.max(img.naturalWidth||img.width,img.naturalHeight||img.height));
    const width=Math.max(1,Math.round((img.naturalWidth||img.width)*scale));
    const height=Math.max(1,Math.round((img.naturalHeight||img.height)*scale));
    const canvas=document.createElement('canvas');canvas.width=width;canvas.height=height;
    const ctx=canvas.getContext('2d');ctx.fillStyle='#ffffff';ctx.fillRect(0,0,width,height);ctx.drawImage(img,0,0,width,height);if(grayscale)grayscaleCanvas(ctx,width,height);
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

  const docLabels=documentLabels();
  // Convert the imported Invoice template title / document number when selected.
  if(state.documentType==='consignment'||state.documentType==='quotation'){
    ws.eachRow(row=>row.eachCell(cell=>{
      if(typeof cell.value==='string'){
        cell.value=cell.value
          .replace(/Sales Invoice/gi,state.documentType==='quotation'?'Quotation':'Sales Consign')
          .replace(/Invoice No\.?/gi,state.documentType==='quotation'?'Quotation No.':'Consign No.');
      }
    }));
  }
  // All document types use the same concise date label in preview and exports.
  ws.eachRow(row=>row.eachCell(cell=>{
    if(typeof cell.value==='string'){
      cell.value=cell.value.replace(/(?:Invoice Date|Consign Date|Consignment Date|Quotation Date)/gi,'Date');
    }
  }));

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
  const inv=norm($('#invoiceNo').value)||formatDocumentNo();
  const invoiceDateText=norm($('#invoiceDate').value);
  const invoiceDateEnglish=englishInvoiceDate(invoiceDateText||today());
  setMapped('Invoice No.',inv);
  setMapped('Invoice Date',invoiceDateEnglish);
  {
    const dateAddress=getMap('Invoice Date');
    if(dateAddress){const dateCell=ws.getCell(dateAddress.split(':')[0]);dateCell.numFmt='@';dateCell.value=invoiceDateEnglish;}
  }
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
  const baseContentRows=Math.min(4,originalContentRows);
  const separatorSourceRow=firstItemRow+baseContentRows;
  const columnCount=Math.max(9,ws.columnCount||9);

  // A:I widths matched to the user-approved iPhone Excel widths from INV260003.
  // Keep true 100% scale and 1.0 cm side margins; do not globally shrink text/images.
  const templateColumnWidths={A:9.2890625,B:13.49609375,C:23.94921875,D:4.7890625,E:4.7890625,F:9.43359375,G:6.53125,H:9.72265625,I:11.90234375};
  for(const [col,width] of Object.entries(templateColumnWidths))ws.getColumn(col).width=width;

  const captureCell=(cell)=>({
    value:cell.value,style:cloneStyle(cell.style),numFmt:cell.numFmt,
    alignment:cloneStyle(cell.alignment),border:cloneStyle(cell.border),
    fill:cloneStyle(cell.fill),font:cloneStyle(cell.font),protection:cloneStyle(cell.protection)
  });
  const contentStyle=[];
  for(let c=1;c<=columnCount;c++)contentStyle.push(captureCell(ws.getRow(firstItemRow).getCell(c)));
  const separatorStyle=[];
  for(let c=1;c<=columnCount;c++)separatorStyle.push(captureCell(ws.getRow(separatorSourceRow).getCell(c)));
  // Locked to the approved Invoice Template print geometry.
  const contentHeight=10.2;
  const separatorHeight=10.2;

  const footerRows=[];
  for(let r=footerBaseRow;r<=originalFooterEnd;r++){
    const row=[];
    for(let c=1;c<=columnCount;c++)row.push(captureCell(ws.getRow(r).getCell(c)));
    footerRows.push({height:ws.getRow(r).height,row});
  }

  // Preserve all drawings already embedded in the imported template, including the company letterhead.
  for(let r=firstItemRow;r<=originalFooterEnd;r++){
    try{ws.unMergeCells(`D${r}:E${r+3}`)}catch{}
  }

  // Each item uses at least 4 content rows. Extra DESC rows extend only the text area.
  // One additional 10.2 pt separator row follows every item; the image remains fixed to the first 4 rows.
  const itemPlans=formalItems().map(item=>{
    const lines=[articleDescriptionFor(item),...effectiveDescriptions(item)].map(norm).filter(Boolean);
    const contentRows=Math.max(4,lines.length);
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

  function buildFixedRowPagePlan(plans,maxItemRows=54){
    const pages=[];
    if(!plans.length)return {pages};
    let start=0,usedRows=0;
    for(let i=0;i<plans.length;i++){
      const rows=plans[i].totalRows;
      if(usedRows>0&&usedRows+rows>maxItemRows){
        pages.push({start,end:i-1,usedRows});
        start=i;usedRows=0;
      }
      usedRows+=rows;
    }
    pages.push({start,end:plans.length-1,usedRows});
    return {pages};
  }

  let rowCursor=firstItemRow;
  let missingImages=0;
  // Pagination rule (Template row model):
  // - Header and Footer are outside the Item-row count.
  // - A page without the Footer may use up to 54 Item rows.
  // - The Footer reserves 16 rows. Therefore, when it shares the final page,
  //   that page may use up to 38 Item rows (54 - 16).
  // - If the final Item page uses more than 38 rows, the Footer starts on a
  //   new page and the previous page is allowed to use the full 54 Item rows.
  // - An Item (including its separator row) is never split across pages.
  const fullItemRowsPerPage=54;
  const footerReservedRows=16;
  const finalItemRowsWithFooter=fullItemRowsPerPage-footerReservedRows; // 38
  const fixedRowPlan=buildFixedRowPagePlan(itemPlans,fullItemRowsPerPage);
  const itemPages=fixedRowPlan.pages||[];
  const pageStartIndexes=new Set(itemPages.slice(1).map(p=>p.start));
  const lastItemPage=itemPages.length?itemPages[itemPages.length-1]:null;
  const footerNeedsOwnPage=!!lastItemPage&&lastItemPage.usedRows>finalItemRowsWithFooter;

  for(let i=0;i<itemPlans.length;i++){
    const {item,lines,contentRows,totalRows}=itemPlans[i];
    const start=rowCursor,contentEnd=start+contentRows-1,separatorRow=contentEnd+1;
    if(pageStartIndexes.has(i)){
      try{ws.getRow(Math.max(firstItemRow,start-1)).addPageBreak()}catch{}
    }
    for(let r=start;r<=contentEnd;r++)applyRowStyle(r,contentStyle,contentHeight);
    applyRowStyle(separatorRow,separatorStyle,separatorHeight);

    ws.getCell(`${noCol}${start}`).value=i+1;
    ws.getCell(`${noCol}${start}`).alignment={...cloneStyle(ws.getCell(`${noCol}${start}`).alignment),horizontal:'center',vertical:'middle'};
    ws.getCell(`${lotCol}${start}`).value=`Lot.No. : ${item.lotNo}`;
    ws.getCell(`${artCol}${start+1}`).value=item.artNo;
    ws.getCell(`${lotCol}${start}`).font={...cloneStyle(ws.getCell(`${lotCol}${start}`).font),bold:false};
    ws.getCell(`${artCol}${start+1}`).font={...cloneStyle(ws.getCell(`${artCol}${start+1}`).font),bold:false};

    for(let r=0;r<contentRows;r++){
      const cell=ws.getCell(`${descCol}${start+r}`);
      cell.value=lines[r]||'';
      cell.alignment={...cloneStyle(cell.alignment),vertical:'middle',wrapText:false};
    }

    try{ws.mergeCells(`${imageStartCol}${start}:${imageEndCol}${start+3}`)}catch{}
    ws.getCell(`${imageStartCol}${start}`).value=null;
    ws.getCell(`${imageStartCol}${start}`).alignment={horizontal:'center',vertical:'middle'};

    ws.getCell(`${qtyCol}${start}`).value=item.qty;
    ws.getCell(`${qtyCol}${start}`).numFmt='0';
    ws.getCell(`${unitCol}${start}`).value=item.unit;
    ws.getCell(`${unitPriceCol}${start}`).value=item.unitPrice;
    ws.getCell(`${unitPriceCol}${start}`).numFmt=currencyExcelFormat();
    ws.getCell(`${amountCol}${start}`).value=item.qty*item.unitPrice;
    ws.getCell(`${amountCol}${start}`).numFmt=currencyExcelFormat();

    const selected=getImg(item);
    if(selected?.file){
      try{
        const asset=await imageFileToJpegAsset(selected.file,620,.84,!!selected.grayscale);
        const imageId=wb.addImage({base64:asset.base64,extension:'jpeg'});
        const imageStartColNo=excelColNumber(imageStartCol),imageEndColNo=excelColNumber(imageEndCol);
        const imageEndRow=start+3;
        let boxW=0,boxH=0;
        for(let c=imageStartColNo;c<=imageEndColNo;c++)boxW+=excelColPixels(ws,c);
        for(let r=start;r<=imageEndRow;r++)boxH+=excelRowPixels(ws,r);
        const pad=1,maxW=Math.max(20,boxW-pad*2),maxH=Math.max(20,boxH-pad*2);
        const scale=Math.min(maxW/asset.width,maxH/asset.height);
        const width=Math.max(20,Math.round(asset.width*scale));
        const height=Math.max(20,Math.round(asset.height*scale));
        const xOffset=(boxW-width)/2,yOffset=(boxH-height)/2;
        ws.addImage(imageId,{tl:{col:imageAnchorCol(ws,imageStartColNo,imageEndColNo,xOffset),row:imageAnchorRow(ws,start,imageEndRow,yOffset)},ext:{width,height},editAs:'oneCell'});
      }catch{missingImages++}
    }else missingImages++;

    rowCursor+=totalRows;
    setExcelExportStatus(`正在依 Template Map 建立 Excel… ${i+1}/${state.items.length}`);
  }


  // When the final Item page exceeds the 38-row Item allowance available
  // beside the 16-row Footer, force the Footer onto its own page. The final
  // Item page can then use the full 54-row Item capacity instead of leaving
  // the Footer's 16-row reservation blank.
  if(footerNeedsOwnPage&&totalItemRows>0){
    try{ws.getRow(Math.max(firstItemRow,footerStart-1)).addPageBreak()}catch{}
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
  ws.getCell(subAddr).value=t.sub;ws.getCell(subAddr).numFmt=currencyExcelFormat();
  ws.getCell(discountAddr).value=t.discount;ws.getCell(discountAddr).numFmt=currencyExcelFormat(currencyCode(),true);
  ws.getCell(totalAddr).value=t.total;ws.getCell(totalAddr).numFmt=currencyExcelFormat();

  // Keep the visible Total currency code synchronized with the selected USD / EUR.
  // This updates the template label (for example, Total : (USD)) without touching Total Amount.
  const totalCurrencyLabel=`Total : (${currencyCode()})`;
  for(let r=footerStart;r<=requiredEnd;r++)for(let c=1;c<=columnCount;c++){
    const cell=ws.getRow(r).getCell(c),text=norm(cell.value);
    if(/^total\s*:\s*(?:\(\s*[A-Z]{3}\s*\))?\s*$/i.test(text))cell.value=totalCurrencyLabel;
  }

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


  // Uniform alignment requested for the complete Invoice sheet.
  for(let r=1;r<=requiredEnd;r++)for(let c=1;c<=columnCount;c++){
    const cell=ws.getRow(r).getCell(c);
    cell.alignment={...cloneStyle(cell.alignment),vertical:'middle',wrapText:false};
  }
  ws.pageSetup=ws.pageSetup||{};
  ws.pageSetup.paperSize=9;
  ws.pageSetup.orientation='portrait';
  // Use the approved Template's true 100% scale. Do not combine this with
  // Fit-to-Width because Excel may silently rescale and change page breaks.
  ws.pageSetup.fitToPage=false;
  ws.pageSetup.fitToWidth=undefined;
  ws.pageSetup.fitToHeight=undefined;
  ws.pageSetup.scale=100;
  ws.pageSetup.printArea=`A1:I${requiredEnd}`;
  ws.pageSetup.printTitlesRow=`1:${Math.max(1,firstItemRow-1)}`;
  // Print geometry: keep true 100% scale, existing row heights,
  // 1.0 cm margins, and a 54-row physical Item-page capacity.
  // ExcelJS stores margins in inches.
  const cmToIn=1/2.54;
  ws.pageSetup.margins={
    left:1.0*cmToIn,right:1.0*cmToIn,
    top:1.0*cmToIn,bottom:1.0*cmToIn,
    header:0.8*cmToIn,footer:0.8*cmToIn
  };
  ws.headerFooter=ws.headerFooter||{};ws.headerFooter.oddFooter='&RPage &P of &N';

  if(mapWs)wb.removeWorksheet(mapWs.id);
  ws.name=state.documentType==='consignment'?'Consignment':state.documentType==='quotation'?'Quotation':'Invoice';
  const buffer=await wb.xlsx.writeBuffer();
  downloadBlob(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),`${inv}.xlsx`);
  setExcelExportStatus(`已依 Template Map 輸出 ${state.invoiceTemplateName}${missingImages?`；${missingImages} 款沒有圖片`:''}。`,'ok');
}
async function exportInvoiceExcel(){
  if(!fxPricingReady())return alert(`請先取得或輸入 USD → ${currencyCode()} 的 FX Rate。`);
  if(!quote14KReady())return alert('14K 參考報價尚未取得 Kitco Ask，請先線上更新或手動輸入。');
  if(!state.items.length){alert('Invoice 沒有貨品。');return}
  if(typeof ExcelJS==='undefined'){setExcelExportStatus('Excel 輸出程式未載入，請連接網絡後重新開啟。','error');return}
  const btn=$('#exportExcelBtn');btn.disabled=true;setExcelExportStatus('正在建立 Excel Invoice…');
  try{
    if(state.invoiceTemplateBuffer){await exportInvoiceFromTemplate();return}
    const wb=new ExcelJS.Workbook();
    wb.creator='Universe Invoice PWA';wb.created=new Date();
    const ws=wb.addWorksheet(documentLabels().title,{pageSetup:{paperSize:9,orientation:'portrait',fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.25,right:.25,top:.35,bottom:.35,header:.15,footer:.15}}});
    ws.views=[{showGridLines:false}];
    ws.columns=[
      {key:'no',width:5},{key:'article',width:17},{key:'description',width:34},{key:'image',width:15},
      {key:'qty',width:9},{key:'unit',width:8},{key:'unitPrice',width:14},{key:'amount',width:14}
    ];
    const merge=(range,value,size=10,bold=false,align='left')=>{ws.mergeCells(range);const c=ws.getCell(range.split(':')[0]);c.value=value;c.font={name:'Arial',size,bold};c.alignment={vertical:'middle',horizontal:align,wrapText:true};return c};
    merge('A1:H1','UNIVERSE GEMS & JEWELLERY CO.',17,true,'center');ws.getRow(1).height=24;
    merge('A2:H2','UNIT 11-12, 10/F., FU HANG INDUSTRIAL BUILDING,',9,false,'center');
    merge('A3:H3','NO. 1 HOK YUEN STREET EAST, HUNG HOM, KOWLOON, HONG KONG',9,false,'center');
    merge('A4:H4','TEL : (852) 2363 5409     FAX : (852) 2765 0343',9,false,'center');
    ws.getRow(5).height=7;
    merge('A6:D6',documentLabels().title,16,true,'left');
    merge('E6:H6',`No. : ${norm($('#invoiceNo').value)}`,11,true,'right');
    merge('A7:D7',`${documentLabels().date} : ${englishInvoiceDate($('#invoiceDate').value)}`,10);
    merge('E7:H7',`Currency : ${norm($('#currency').value)}`,10,false,'right');
    merge('A8:D8',`Shipment Method : ${norm($('#shipmentMethod').value)}`,10);
    merge('E8:H8',`Customer Code : ${norm($('#customerCode').value)}`,10,false,'right');
    merge('A9:D11',`Customer : ${norm($('#customerName').value)}\n${norm($('#customerAddress').value)}`,10,true);
    merge('E9:H11',"Vendor's Banker\nThe Hong Kong & Shanghai Banking Corporation Ltd.\nAddress : 41 Ma Tau Wai Road, Hung Hom, Kowloon, Hong Kong\nA/C # : 012-593570-001\nA/C Name : Universe Gems & Jewellery Co.",9,false,'left');
    [9,10,11].forEach(r=>ws.getRow(r).height=20);
    const headerRow=13;
    const headers=['No.','Article No.','Description','Picture','Quantity','Unit','Unit Price','Amount'];
    headers.forEach((h,i)=>{const c=ws.getCell(headerRow,i+1);c.value=h;c.font={name:'Arial',size:10,bold:true};c.alignment={horizontal:'center',vertical:'middle',wrapText:true};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE5E7EB'}};applyThinBorder(c)});
    ws.getRow(headerRow).height=24;
    ws.mergeCells(`A${headerRow+1}:G${headerRow+1}`);const fob=ws.getCell(headerRow+1,8);fob.value='F.O.B. Value';fob.font={name:'Arial',size:10,bold:true};fob.alignment={vertical:'middle',horizontal:'right'};applyThinBorder(fob);ws.getRow(headerRow+1).height=21;
    let row=headerRow+2;
    let missingImages=0;
    const exportItems=formalItems();
    for(let i=0;i<exportItems.length;i++){
      const item=exportItems[i],start=row,end=row+4;
      for(let r=start;r<=end;r++)ws.getRow(r).height=10.5;
      ws.mergeCells(`A${start}:A${end}`);ws.getCell(`A${start}`).value=i+1;
      ws.getCell(`A${start}`).alignment={horizontal:'center',vertical:'middle'};
      ws.mergeCells(`B${start}:B${end}`);ws.getCell(`B${start}`).value=`Lot.No. : ${item.lotNo}\n${item.artNo}`;ws.getCell(`B${start}`).font={name:'Arial',size:10,bold:false};ws.getCell(`B${start}`).alignment={vertical:'top',wrapText:true};
      ws.mergeCells(`C${start}:C${end}`);ws.getCell(`C${start}`).value=[articleDescriptionFor(item),...effectiveDescriptions(item)].filter(Boolean).join('\n');ws.getCell(`C${start}`).alignment={vertical:'top',wrapText:true};ws.getCell(`C${start}`).font={name:'Arial',size:10};
      ws.mergeCells(`D${start}:D${end}`);
      ws.mergeCells(`E${start}:E${end}`);ws.getCell(`E${start}`).value=item.qty;ws.getCell(`E${start}`).alignment={horizontal:'center',vertical:'middle'};
      ws.mergeCells(`F${start}:F${end}`);ws.getCell(`F${start}`).value=item.unit;ws.getCell(`F${start}`).alignment={horizontal:'center',vertical:'middle'};
      ws.mergeCells(`G${start}:G${end}`);ws.getCell(`G${start}`).value=item.unitPrice;ws.getCell(`G${start}`).numFmt=currencyExcelFormat();ws.getCell(`G${start}`).alignment={horizontal:'right',vertical:'middle'};
      ws.mergeCells(`H${start}:H${end}`);ws.getCell(`H${start}`).value={formula:`E${start}*G${start}`,result:item.qty*item.unitPrice};ws.getCell(`H${start}`).numFmt=currencyExcelFormat();ws.getCell(`H${start}`).alignment={horizontal:'right',vertical:'middle'};
      for(let r=start;r<=end;r++)for(let c=1;c<=8;c++)applyThinBorder(ws.getCell(r,c));
      const selected=getImg(item);
      if(selected?.file){
        try{
          const dataUrl=await imageFileToJpegDataUrl(selected.file,620,.82,!!selected.grayscale);
          const imageId=wb.addImage({base64:dataUrl,extension:'jpeg'});
          ws.addImage(imageId,{tl:{col:3.01,row:start-1+.01},br:{col:3.99,row:start+3.99},editAs:'oneCell'});
        }catch{missingImages++}
      }else missingImages++;
      row=end+1;
      setExcelExportStatus(`正在建立 Excel Invoice… ${i+1}/${state.items.length}`);
    }
    const t=totals();
    ws.mergeCells(`A${row}:F${row}`);ws.getCell(`A${row}`).value='Total Quantity';ws.getCell(`A${row}`).font={bold:true};ws.getCell(`G${row}`).value=t.qty;ws.getCell(`G${row}`).font={bold:true};ws.getCell(`G${row}`).alignment={horizontal:'right'};
    row++;
    ws.mergeCells(`A${row}:F${row}`);ws.getCell(`A${row}`).value='Sub Total';ws.getCell(`A${row}`).font={bold:true};ws.getCell(`G${row}`).value=t.sub;ws.getCell(`G${row}`).numFmt=currencyExcelFormat();ws.getCell(`G${row}`).font={bold:true};ws.getCell(`G${row}`).alignment={horizontal:'right'};
    row++;
    ws.mergeCells(`A${row}:F${row}`);ws.getCell(`A${row}`).value='Discount Amount';ws.getCell(`G${row}`).value=t.discount;ws.getCell(`G${row}`).numFmt=currencyExcelFormat(currencyCode(),true);ws.getCell(`G${row}`).alignment={horizontal:'right'};
    row++;
    ws.mergeCells(`A${row}:F${row}`);ws.getCell(`A${row}`).value=`Total : (${currencyCode()})`;ws.getCell(`A${row}`).font={bold:true,size:12};ws.getCell(`G${row}`).value=t.total;ws.getCell(`G${row}`).numFmt=currencyExcelFormat();ws.getCell(`G${row}`).font={bold:true,size:12};ws.getCell(`G${row}`).alignment={horizontal:'right'};
    row+=2;
    ws.mergeCells(`A${row}:H${row+2}`);const remarkCell=ws.getCell(`A${row}`);remarkCell.value=`Remark :\n${norm($('#remark').value)}`;remarkCell.alignment={vertical:'top',wrapText:true};remarkCell.font={name:'Arial',size:10};row+=2;
    row+=2;merge(`A${row}:D${row}`,'Vender Signature : ______________________',10);merge(`E${row}:H${row}`,'Accept By : ______________________',10,false,'right');
    ws.headerFooter.oddFooter='&RPage &P of &N';
    ws.pageSetup.printArea=`A1:H${row}`;
    ws.autoFilter={from:{row:headerRow,column:1},to:{row:headerRow,column:8}};
    const buffer=await wb.xlsx.writeBuffer();
    const inv=norm($('#invoiceNo').value)||formatDocumentNo();
    downloadBlob(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),`${inv}.xlsx`);
    setExcelExportStatus(`Excel Invoice 已輸出${missingImages?`；${missingImages} 款沒有嵌入圖片`:''}。`,'ok');
  }catch(err){console.error(err);setExcelExportStatus('Excel 輸出失敗：'+(err.message||err),'error')}
  finally{btn.disabled=false}
}

$('#exportExcelBtn').onclick=exportInvoiceExcel;$('#exportPdfBtn').onclick=()=>{if(!quote14KReady())return alert('14K 參考報價尚未取得 Kitco Ask，請先線上更新或手動輸入。');renderPreview();setTimeout(()=>window.print(),80)};
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
function confirmStoredDocument(){
  const type=state.documentType,documentNo=norm($('#invoiceNo').value)||formatDocumentNo(),recall=state.recall&&state.recall.type===type&&state.recall.documentNo===documentNo?state.recall:null;
  if(!state.items.length&&!recall)return alert(`${documentLabels().short} 沒有貨品。`);
  if(!state.stockAllRows.length)return alert('請先匯入 jmsdata。');
  if(state.items.length&&!fxPricingReady())return alert(`請先取得或輸入 USD → ${currencyCode()} 的 FX Rate。`);
  if(state.items.length&&!quote14KReady())return alert('14K 參考報價尚未取得 Kitco Ask，請先線上更新或手動輸入。');
  if(type==='quotation'){if(!state.items.length)return alert('Quotation 沒有貨品。');releaseCustomImages();state.items=[];state.quote.karat='18K';$$('input[name="quoteKarat"]').forEach(x=>x.checked=x.value==='18K');updateGoldQuoteUI();advanceDocumentSequence(documentNo,type);renderItems();status('#addMessage',`Quotation ${documentNo} 已 Confirm（庫存及 jmsdata 沒有改動）；下一張為 ${$('#invoiceNo').value}。`,'ok');return;}
  const keys=storeKeysForType(type),targetStatus=type==='consignment'?'CONSIGNED':'SOLD_ON_HAND';
  if(!recall&&findLatestDocument(type,documentNo))return alert(`${documentNo} 已存在。請使用 Recall 重新開啟原文件，不要重複 Confirm 同一文件號碼。`);
  const revision=recall?recall.baseRevision+1:0,oldItems=recall?new Map(recall.items.map(x=>[String(docValue(x,'LOTNO')),x])):new Map(),newItems=formalItems(),newLots=new Set(newItems.map(x=>String(x.lotNo))),rollback=new Map(),documentStatus=newItems.length?'CONFIRMED':'CANCELLED';
  try{
    if(recall){for(const [lot,oldRow] of oldItems){if(newLots.has(lot))continue;const p=state.stockCatalog.get(lot)||{lotNo:lot,artNo:normArt(docValue(oldRow,'ARTNO'))},before=inventoryVectorFromRow(state.stockRowByLot.get(lot)||{}),original=vectorFromRecord(oldRow,'ORIGINAL'),currentStatus=inventoryStatusFromVector(before);if(currentStatus!==targetStatus)throw new Error(`${p.artNo} / LOTNO ${lot} 目前狀態是 ${historyStatusLabel(currentStatus)}，已不是這張文件原本的 ${historyStatusLabel(targetStatus)}；如貨品已 Delivered，請先在配套搜尋使用「取消已交貨」。`);if(!rollback.has(lot))rollback.set(lot,before);setInventoryVectorForLot(lot,original);appendTransaction({type,documentNo,revision,item:p,before,after:original})}}
    const itemRows=[];
    for(const item of newItems){const lot=String(item.lotNo),row=state.stockRowByLot.get(lot);if(!row)throw new Error(`jmsdata 找不到 LOTNO ${lot}`);const before=inventoryVectorFromRow(row),oldRow=oldItems.get(lot),isExisting=!!oldRow;let original=isExisting?vectorFromRecord(oldRow,'ORIGINAL'):(item._originalVector||before),after;
      const currentStatus=inventoryStatusFromVector(before);if(isExisting){if(currentStatus===targetStatus)after=inventoryVectorForStatus(targetStatus,before.m);else if(type==='invoice'&&currentStatus==='SOLD_DELIVERED')after=before;else throw new Error(`${item.artNo} / LOTNO ${lot} 的庫存狀態已在這張文件之後改為 ${historyStatusLabel(currentStatus)}，不能由舊文件 Recall 覆蓋。`)}else after=inventoryVectorForStatus(targetStatus,before.m);
      if(!isExisting&&inventoryStatusFromVector(before)!=='AVAILABLE')throw new Error(`${item.artNo} / LOTNO ${lot} 目前不是 Available，不能新增到 ${documentLabels(type).short}。`);
      if(!rollback.has(lot))rollback.set(lot,before);setInventoryVectorForLot(lot,after);appendTransaction({type,documentNo,revision,item,before,after});itemRows.push(itemRowForDocument(item,documentNo,revision,original,after));
    }
    const previousCreated=recall?docValue(recall.header,'CREATED_AT'):'';state.documentStore[keys.headers].push(headerRowForCurrentDocument(type,documentNo,revision,previousCreated,documentStatus));state.documentStore[keys.items].push(...itemRows);saveDocumentStore();rebuildInventoryMaps();const exported=exportJmsdata(),wasRecall=!!recall;
    state.recall=null;setRecallLock(false);$('#recallActive')?.classList.add('hidden');clearCurrentDocument();if(wasRecall)setDefaultInvoiceNo(true);else advanceDocumentSequence(documentNo,type);renderRecallResults();const action=documentStatus==='CANCELLED'?'已取消整張文件':revision?'已儲存修訂':'已 Confirm';status('#addMessage',`${documentLabels(type).short} ${documentNo} ${action}（Revision ${revision}），${exported?'已輸出 jmsdata.xls':'jmsdata 匯出失敗'}；目前文件號碼為 ${$('#invoiceNo').value}。`,exported?'ok':'warn');status('#stockStatus',`目前 Available Stock：${state.products.size} 件。`,'ok');
  }catch(err){console.error(err);for(const [lot,vector] of rollback){try{setInventoryVectorForLot(lot,vector)}catch{}}alert('Confirm 失敗：'+(err.message||err));rebuildInventoryMaps()}
}
function exportCurrentStockAfterConfirm(){return confirmStoredDocument()}
function exportRemaining(){exportCurrentStockAfterConfirm()}
$('#confirmInvoiceBtn').onclick=()=>{const l=documentLabels(),isCancel=!!state.recall&&!state.items.length&&state.documentType!=='quotation',message=isCancel?`目前已刪除全部貨品。確認將 ${state.recall.documentNo} 標記為 Cancelled，並把原有貨品恢復到加入文件前的庫存狀態？`:`${l.confirm}？`;if(confirm(message))exportCurrentStockAfterConfirm()};

function articleCore(value){const s=normArt(value),m=s.match(/(\d{3,})/);return m?m[1]:s.replace(/^[A-Z]+[-\s]*/,'').replace(/\.[A-Z0-9]+$/,'')}
function articleType(value){const m=normArt(value).match(/^([A-Z]+)/);return m?m[1]:'OTHER'}
function validStoneBreakdownCodes(){return [...activeStoneAliases().keys()].map(x=>norm(x).toUpperCase().replace(/\s+/g,'')).filter(Boolean).sort((a,b)=>b.length-a.length)}
function stoneCodesFromDescriptionLine(line){
  const raw=String(line||'').toUpperCase(),firstDash=raw.indexOf('-');if(firstDash<0)return[];
  const tail=raw.slice(firstDash+1),nextDash=tail.indexOf('-'),stoneBlock=(nextDash>=0?tail.slice(0,nextDash):tail).replace(/\s+/g,'');
  if(!stoneBlock)return[];
  const known=validStoneBreakdownCodes(),out=[];
  for(const part of stoneBlock.split(/[+\/，,]/).map(x=>x.replace(/[^A-Z0-9]/g,'')).filter(Boolean)){
    const code=known.find(c=>part.startsWith(c));
    if(code&&code!=='CDM'&&!out.includes(code))out.push(code);
  }
  return out;
}
function stoneCodesForProduct(p){const out=[];for(const line of (p.descriptions||[])){for(const code of stoneCodesFromDescriptionLine(line)){if(!out.includes(code))out.push(code)}}return out}
const STOCK_TYPE_ORDER=['RG','ER','PT','BR','NL','BL','BG'];
function stockTypeRank(type){const t=norm(type).toUpperCase(),i=STOCK_TYPE_ORDER.indexOf(t);return i>=0?i:STOCK_TYPE_ORDER.length}
function compareStockTypes(a,b){const ra=stockTypeRank(a),rb=stockTypeRank(b);return ra-rb||String(a).localeCompare(String(b))}
function stoneOrderList(){const out=[];for(const code of activeStoneAliases().keys()){const c=norm(code).toUpperCase();if(c&&c!=='CDM'&&!out.includes(c))out.push(c)}return out}
function stoneOrderRank(code){const c=norm(code).toUpperCase(),list=stoneOrderList(),i=list.indexOf(c);return i>=0?i:list.length+100}
function colorGroupOrderList(){const out=[];for(const code of stoneOrderList()){const group=stoneGroupForCode(code);if(group&&!out.includes(group))out.push(group)}return out}
function colorGroupRank(group){const g=norm(group).toUpperCase(),list=colorGroupOrderList(),i=list.indexOf(g);return i>=0?i:list.length+100}
function stockSortStoneCodes(p,preferredStones=[]){
  const codes=stoneCodesForProduct(p),preferred=(Array.isArray(preferredStones)?preferredStones:[]).map(x=>norm(x).toUpperCase());
  const matched=preferred.length?codes.filter(code=>preferred.includes(norm(code).toUpperCase())):codes;
  const source=matched.length?matched:codes;
  return [...source].sort((a,b)=>colorGroupRank(stoneGroupForCode(a))-colorGroupRank(stoneGroupForCode(b))||stoneOrderRank(a)-stoneOrderRank(b)||String(a).localeCompare(String(b)));
}
function primaryStoneCodeForProduct(p,preferredStones=[]){return stockSortStoneCodes(p,preferredStones)[0]||''}
function stockColorGroupForProduct(p,preferredStones=[]){
  const preferred=Array.isArray(preferredStones)?preferredStones:[];
  if(!preferred.length&&isMultiColorProduct(p))return'MULTI';
  return stoneGroupForCode(primaryStoneCodeForProduct(p,preferred));
}
function numericArticleCore(value){const n=Number(articleCore(value));return Number.isFinite(n)?n:-1}
function compareColorSortedStock(a,b,preferredStones=[]){
  const coreDiff=numericArticleCore(b.artNo)-numericArticleCore(a.artNo);if(coreDiff)return coreDiff;
  const ga=stockColorGroupForProduct(a,preferredStones),gb=stockColorGroupForProduct(b,preferredStones),groupDiff=colorGroupRank(ga)-colorGroupRank(gb);if(groupDiff)return groupDiff;
  if(ga!==gb){if(!ga)return 1;if(!gb)return-1;const groupAlpha=ga.localeCompare(gb);if(groupAlpha)return groupAlpha}
  const sa=primaryStoneCodeForProduct(a,preferredStones),sb=primaryStoneCodeForProduct(b,preferredStones),stoneDiff=stoneOrderRank(sa)-stoneOrderRank(sb);if(stoneDiff)return stoneDiff;
  if(sa!==sb){if(!sa)return 1;if(!sb)return-1;const stoneAlpha=sa.localeCompare(sb);if(stoneAlpha)return stoneAlpha}
  const ta=articleType(a.artNo),tb=articleType(b.artNo),typeDiff=compareStockTypes(ta,tb);if(typeDiff)return typeDiff;
  return normArt(a.artNo).localeCompare(normArt(b.artNo))||String(a.lotNo).localeCompare(String(b.lotNo));
}
function compareWildcardStock(a,b){return compareColorSortedStock(a,b,[])}
function inventorySearchRecords(){const map=new Map();for(const p of state.stockCatalog.values())map.set(String(p.lotNo),{...productSnapshot(p),status:'AVAILABLE'});for(const p of state.products.values())map.set(String(p.lotNo),{...productSnapshot(p),status:'AVAILABLE'});for(const h of state.inventoryHistory.values())map.set(String(h.lotNo),{...h,status:h.status||'AVAILABLE'});return [...map.values()]}
function filterButtonHTML(value,label,active){return `<button type="button" class="filter-chip${active?' active':''}" data-value="${esc(value)}" aria-pressed="${active?'true':'false'}">${esc(label)}</button>`}
function stoneFilterLabel(code){const c=norm(code).toUpperCase();return c==='SKY BTO'||c==='SKY BT'?'SKY':c}
function selectedStockFilters(kind){const key=kind==='type'?'types':kind==='stone'?'stones':'statuses',v=state.stockSearch[key];return Array.isArray(v)?v:[]}
function toggleStockFilter(kind,value){
  const key=kind==='type'?'types':kind==='stone'?'stones':'statuses';
  if(value==='ALL'){state.stockSearch[key]=[];renderStockSearch();return}
  const selected=new Set(selectedStockFilters(kind));
  if(selected.has(value))selected.delete(value);else selected.add(value);
  state.stockSearch[key]=[...selected];renderStockSearch();
}
function stockFilterSummaryText(){const t=selectedStockFilters('type'),s=selectedStockFilters('stone');return `款式：${t.length?t.join(', '):'全部'} · 石頭：${s.length?s.map(stoneFilterLabel).join(', '):'全部'}`}
function runStockSearch(){const raw=norm($('#stockSearchInput').value).toUpperCase(),core=raw==='*'?'*':articleCore(raw);state.stockSearch.query=core||raw;state.stockSearch.types=[];state.stockSearch.stones=[];state.stockSearch.statuses=[];state.stockSearch.filtersOpen=false;renderStockSearch()}
function renderStockSearch(){
  const resultsEl=$('#stockSearchResults');if(!resultsEl)return;const q=norm(state.stockSearch.query||$('#stockSearchInput')?.value).toUpperCase(),filterControls=$('#stockFilterControls'),filterPanel=$('#stockFilterPanel');if(!q){resultsEl.innerHTML='';$('#stockSearchMessage').textContent='請輸入款號搜尋。';$('#stockTypeFilters').classList.add('hidden');$('#stockStoneFilters').classList.add('hidden');$('#stockSearchSummary').classList.add('hidden');filterControls?.classList.add('hidden');filterPanel?.classList.add('hidden');return}
  const wildcard=q==='*',core=wildcard?'*':articleCore(q),all=wildcard?inventorySearchRecords():inventorySearchRecords().filter(x=>articleCore(x.artNo)===core||normArt(x.artNo).includes(normArt(q)));const types=[...new Set(all.map(x=>articleType(x.artNo)))].sort(compareStockTypes),stones=[...new Set(all.flatMap(stoneCodesForProduct))].sort((a,b)=>stoneOrderRank(a)-stoneOrderRank(b)||a.localeCompare(b));
  const selectedTypes=selectedStockFilters('type'),selectedStones=selectedStockFilters('stone'),selectedStatuses=selectedStockFilters('status');
  const typeBox=$('#stockTypeFilters'),stoneBox=$('#stockStoneFilters');typeBox.dataset.kind='type';stoneBox.dataset.kind='stone';typeBox.innerHTML=`<div class="filter-chips">${filterButtonHTML('ALL','全部',selectedTypes.length===0)}${types.map(x=>filterButtonHTML(x,x,selectedTypes.includes(x))).join('')}</div>`;stoneBox.innerHTML=`<div class="filter-chips">${filterButtonHTML('ALL','全部',selectedStones.length===0)}${stones.map(x=>filterButtonHTML(x,stoneFilterLabel(x),selectedStones.includes(x))).join('')}</div>`;typeBox.classList.toggle('hidden',!types.length);stoneBox.classList.toggle('hidden',!stones.length);typeBox.querySelectorAll('.filter-chip').forEach(b=>b.onclick=()=>toggleStockFilter('type',b.dataset.value));stoneBox.querySelectorAll('.filter-chip').forEach(b=>b.onclick=()=>toggleStockFilter('stone',b.dataset.value));
  const hasFilters=types.length||stones.length;if(filterControls){filterControls.classList.toggle('hidden',!hasFilters);$('#stockFilterSelectionSummary').textContent=stockFilterSummaryText();$('#stockFilterToggle').textContent=state.stockSearch.filtersOpen?'收起篩選':'展開篩選'}if(filterPanel)filterPanel.classList.toggle('hidden',!hasFilters||!state.stockSearch.filtersOpen);
  const shown=all.filter(x=>(selectedTypes.length===0||selectedTypes.includes(articleType(x.artNo)))&&(selectedStones.length===0||stoneCodesForProduct(x).some(code=>selectedStones.includes(code)))&&(selectedStatuses.length===0||selectedStatuses.includes(x.status))).sort((a,b)=>compareColorSortedStock(a,b,selectedStones));
  const counts={AVAILABLE:0,CONSIGNED:0,SOLD_ON_HAND:0,SOLD_DELIVERED:0};for(const x of all)counts[x.status]=(counts[x.status]||0)+1;$('#stockSearchMessage').textContent=wildcard?`全部庫存：找到 ${all.length} 件，目前顯示 ${shown.length} 件。`:`搜尋 ${core}：找到 ${all.length} 件，目前顯示 ${shown.length} 件。`;const summary=$('#stockSearchSummary');summary.innerHTML=`<button type="button" class="stock-summary-chip${selectedStatuses.includes('AVAILABLE')?' active':''}" data-status="AVAILABLE" aria-pressed="${selectedStatuses.includes('AVAILABLE')?'true':'false'}">Avail ${counts.AVAILABLE||0}</button><button type="button" class="stock-summary-chip${selectedStatuses.includes('CONSIGNED')?' active':''}" data-status="CONSIGNED" aria-pressed="${selectedStatuses.includes('CONSIGNED')?'true':'false'}">Consign ${counts.CONSIGNED||0}</button><button type="button" class="stock-summary-chip${selectedStatuses.includes('SOLD_ON_HAND')?' active':''}" data-status="SOLD_ON_HAND" aria-pressed="${selectedStatuses.includes('SOLD_ON_HAND')?'true':'false'}">Sold-OH ${counts.SOLD_ON_HAND||0}</button><button type="button" class="stock-summary-chip${selectedStatuses.includes('SOLD_DELIVERED')?' active':''}" data-status="SOLD_DELIVERED" aria-pressed="${selectedStatuses.includes('SOLD_DELIVERED')?'true':'false'}">Deliv ${counts.SOLD_DELIVERED||0}</button>`;summary.querySelectorAll('[data-status]').forEach(b=>b.onclick=()=>toggleStockFilter('status',b.dataset.status));summary.classList.remove('hidden');
  if(!shown.length){resultsEl.innerHTML='<div class="notice">沒有符合目前款式／石頭／狀態篩選的貨品。</div>';return}
  resultsEl.innerHTML='';for(const x of shown){const card=document.createElement('div');card.className='stock-result';const imageList=state.imageFiles.get(x.artNo)||[],imageMatch=chooseImageMatch(x),matchedImage=imageList.find(img=>img.variant===imageMatch.variant)||imageList.find(img=>img.variant.toUpperCase()===String(imageMatch.variant).toUpperCase())||imageList.find(img=>img.variant==='Default')||imageList[0],img=matchedImage?.url||placeholder(x.artNo),grayClass=imageMatch.grayscale?'grayscale-image':'',inDoc=state.items.some(i=>String(i.lotNo)===String(x.lotNo)),canAdd=x.status==='AVAILABLE'||state.documentType==='quotation'&&['CONSIGNED','SOLD_ON_HAND','SOLD_DELIVERED'].includes(x.status);card.innerHTML=`<img class="${grayClass}" src="${esc(img)}" alt="${esc(x.artNo)}"><div><h4>${esc(x.artNo)}</h4><div class="lot">LOTNO ${esc(x.lotNo)}</div><div class="desc">${esc((x.descriptions||[]).join('\n'))}</div></div><div class="stock-result-actions"><span class="stock-status ${historyStatusClass(x.status)}">${esc(historyStatusLabel(x.status))}</span>${inDoc?'<span class="stock-current-doc">已在目前文件</span>':''}</div>`;const actions=$('.stock-result-actions',card);if(canAdd&&!inDoc){const b=document.createElement('button');b.type='button';b.textContent=`加入目前 ${documentLabels().short}`;b.onclick=()=>{if(addProductToDocument(x,{fromSearch:true})){b.disabled=true;b.textContent='已加入';renderStockSearch()}};actions.appendChild(b)}if(x.status==='SOLD_ON_HAND'){const d=document.createElement('button');d.type='button';d.className='ghost';d.textContent='標記已交貨';d.onclick=()=>{if(confirm(`${x.artNo} / LOTNO ${x.lotNo}\n標記為 Sold - Delivered？`))markInventoryDelivered(x.lotNo)};actions.appendChild(d)}if(x.status==='SOLD_DELIVERED'){const u=document.createElement('button');u.type='button';u.className='ghost';u.textContent='取消已交貨';u.onclick=()=>{if(confirm(`${x.artNo} / LOTNO ${x.lotNo}\n已確認貨品由客人退回，並取消 Delivered？\n系統會恢復到交貨前狀態。`))undoInventoryDelivered(x.lotNo)};actions.appendChild(u)}resultsEl.appendChild(card)}
}
function activePanelId(){return $('.tab.active')?.dataset.tab||''}
function updateBackToTopButton(){const btn=$('#backToTopBtn');if(!btn)return;const panel=activePanelId();let visible=false,title='回到頂部';if(panel==='invoice'){visible=($('#invoiceItems')?.scrollTop||0)>180;title=`回到 ${documentLabels().short} 貨品列表頂部`}else if(panel==='stockSearch'){const section=$('#stockSearch'),hasResults=($('#stockSearchResults')?.children.length||0)>0;visible=!!section&&hasResults&&window.scrollY>section.offsetTop+260;title='回到配套搜尋頂部'}btn.title=title;btn.setAttribute('aria-label',title);btn.classList.toggle('hidden',!visible)}
function scrollCurrentPanelToTop(){const panel=activePanelId();if(panel==='invoice')$('#invoiceItems')?.scrollTo({top:0,behavior:'smooth'});else if(panel==='stockSearch'){const section=$('#stockSearch');if(section)window.scrollTo({top:Math.max(0,section.offsetTop-72),behavior:'smooth'})}setTimeout(updateBackToTopButton,350)}
$('#backToTopBtn').onclick=scrollCurrentPanelToTop;$('#invoiceItems').addEventListener('scroll',updateBackToTopButton,{passive:true});window.addEventListener('scroll',updateBackToTopButton,{passive:true});window.addEventListener('resize',updateBackToTopButton,{passive:true});
$('#stockSearchBtn').onclick=runStockSearch;$('#stockSearchInput').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();runStockSearch()}};$('#stockSearchInput').oninput=e=>{e.target.value=e.target.value.toUpperCase()};$('#stockFilterToggle').onclick=()=>{state.stockSearch.filtersOpen=!state.stockSearch.filtersOpen;renderStockSearch()};
$('#exportJmsdataBtn').onclick=exportJmsdata;$('#recallSearchBtn').onclick=renderRecallResults;$('#recallSearchInput').oninput=renderRecallResults;$('#recallSearchInput').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();renderRecallResults()}};$('#recallTypeFilter').onchange=renderRecallResults;$('#cancelRecallBtn').onclick=cancelRecall;
$$('input[name="quoteKarat"]').forEach(r=>r.onchange=e=>setQuoteKarat(e.target.value));$('#refreshGoldBtn').onclick=()=>fetchKitcoAsk();let goldInputTimer=null;$('#kitcoAskInput').oninput=e=>{clearTimeout(goldInputTimer);goldInputTimer=setTimeout(()=>{const n=Number(e.target.value);if(n>0)setKitcoAsk(n,{source:'manual',message:`手動 Kitco Ask：USD ${n.toFixed(2)} / oz`});else{state.quote.kitcoAsk=0;syncEffectivePrices({clearCurrentOverride:true});renderItems();updateGoldQuoteUI('請輸入大於 0 的 Kitco Ask。','error')}},180)};

updateFxPanel();updateDocumentTypeUI();updateGoldQuoteUI();renderCustomerSummary();renderItems();renderRecallResults();schedulePreview();updateBackToTopButton();if('serviceWorker'in navigator)navigator.serviceWorker.register('./sw.js').catch(()=>{});
