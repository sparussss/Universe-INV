const $=(s,r=document)=>r.querySelector(s),$$=(s,r=document)=>[...r.querySelectorAll(s)];
const APP_VERSION='0.14.31';
const state={products:new Map(),stockCatalog:new Map(),customers:new Map(),imageFiles:new Map(),imageFilesByName:new Map(),imageOverrides:new Map(),imageOverrideDirty:0,imageOverrideDirtyLots:new Set(),items:[],stockRows:[],stockAllRows:[],stockHeaders:[],stockRowByLot:new Map(),stockWorkbook:null,stockSheetName:'jmsdata',stockFileName:'jmsdata.xlsx',stockIntegrityIssues:[],stockDuplicateLots:[],stoneAliases:new Map(),stoneVariantAliases:new Map(),stoneGroups:new Map(),stoneEnglishNames:new Map(),diamondStoneCodes:new Set(),stoneMappingName:'',stoneDiagnostics:{duplicates:[],multiAlias:[],missingGroup:[],missingType:[],missingQuotation:[],prefixOverlaps:[]},articleMap:new Map(),articleMappingName:'',invoiceTemplateBuffer:null,invoiceTemplateName:'',documentType:'invoice',packageName:'',exhibitionName:'',sortable:null,scanner:null,scannerBusy:false,scannerRunning:false,scannerZoom:{min:1,max:1,step:1,current:1},fx:{rate:1,date:'',source:'usd',fetching:false},quote:{karat:'18K',currentLondonPm:0,currentLondonPmDate:'',source:'',historicalPm:{},goldPrices:new Map(),goldDates:[],goldDataName:'',goldDataRows:0},inventoryHistory:new Map(),documentStore:{invoiceHeaders:[],invoiceItems:[],consignmentHeaders:[],consignmentItems:[],quotationHeaders:[],quotationItems:[],transactions:[]},recall:null,deliveryReturns:new Set(),exhibitionSession:'',draft:{baseline:'',timer:null,prompted:false,restoring:false},stockSearch:{query:'*',types:[],stones:[],statuses:[],imageIssuesOnly:false,filtersOpen:false},editingItemId:null,stockImageEditLot:null,packageFiles:[],customPackageImages:new Map(),packageImageDirtyFiles:new Set(),dataMeta:{},importConflicts:[],health:{},imageIndexProgress:{done:0,total:0},diagnosticLot:'',recordsLoaded:false,recordsFileName:'jmsdata.xlsx / Universe Records',recordsFilePath:'',recordCounters:{invoice:1,consignment:1,quotation:1},recordsDirty:false};
const EXHIBITION_SESSION_KEY='universeExhibitionSession_v1',EXHIBITION_NAME_KEY='universeExhibitionName_v1',IMAGE_OVERRIDE_LOCAL_KEY='universeImageOverrides_v1';try{state.exhibitionSession=localStorage.getItem(EXHIBITION_SESSION_KEY)||'';state.exhibitionName=localStorage.getItem(EXHIBITION_NAME_KEY)||''}catch{}
function formalItems(){return [...state.items].sort((a,b)=>(Number(a.seq)||0)-(Number(b.seq)||0))}
function displayItems(){return formalItems().reverse()}
function normalizeItemSequence(){state.items=formalItems();state.items.forEach((item,i)=>item.seq=i+1)}
const norm=v=>String(v??'').trim(),normCode=v=>String(v??'').replace(/\s+/g,'').toUpperCase(),normArt=v=>norm(v).toUpperCase();
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const today=()=>{const d=new Date(),y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,'0'),day=String(d.getDate()).padStart(2,'0');return `${y}-${m}-${day}`};$('#invoiceDate').value=today();
function englishInvoiceDate(iso){const m=String(iso||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)return String(iso||'');const months=['January','February','March','April','May','June','July','August','September','October','November','December'];return `${Number(m[3])} ${months[Number(m[2])-1]}, ${m[1]}`;}
function discountDisplay(value){const n=Math.max(0,Number(value)||0);return n>0?`(${fmt(n)})`:fmt(0)}
function invoiceYear(){const raw=$('#invoiceDate')?.value,m=String(raw||'').match(/^(\d{4})-/);return m?m[1].slice(-2):String(new Date().getFullYear()).slice(-2)}
function documentPrefix(type=state.documentType){return type==='consignment'?'CON':type==='quotation'?'QUO':'INV'}
function documentSequenceKey(type=state.documentType,yy=invoiceYear()){const key=type==='consignment'?'Consign':type==='quotation'?'Quotation':'Invoice',scope=state.exhibitionSession?`_${state.exhibitionSession}`:'';return `universe${key}Seq_${yy}${scope}`}
function getNextDocumentSequence(type=state.documentType,yy=invoiceYear()){
  const saved=Number(state.recordCounters?.[type]||1);
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
  state.recordCounters[type]=next;state.recordsDirty=true;
  $('#invoiceNo').value=formatDocumentNo(next,yy,type);
}
function advanceInvoiceSequence(confirmedNo){advanceDocumentSequence(confirmedNo,'invoice')}
function saveExhibitionName(value=state.exhibitionName){
  const previous=state.exhibitionName;state.exhibitionName=norm(value);if(state.recordsLoaded&&previous!==state.exhibitionName)state.recordsDirty=true;if(state.dataMeta?.records)state.dataMeta.records.exhibitionName=state.exhibitionName;try{if(state.exhibitionName)localStorage.setItem(EXHIBITION_NAME_KEY,state.exhibitionName);else localStorage.removeItem(EXHIBITION_NAME_KEY)}catch{}
  const input=$('#exhibitionName');if(input&&document.activeElement!==input)input.value=state.exhibitionName;
  const current=$('#currentExhibitionStatus');if(current)current.textContent=state.exhibitionName?`目前展覽：${state.exhibitionName}`:'目前展覽：尚未命名';updateDataVersionPanel();
}
function syncExhibitionNameFromInput(){saveExhibitionName($('#exhibitionName')?.value||'')}
setDefaultInvoiceNo();
let previewTimer=null;
function schedulePreview(){clearTimeout(previewTimer);previewTimer=setTimeout(()=>{renderPreview();const el=$('#previewUpdatedAt');if(el)el.textContent='最後更新：'+new Date().toLocaleTimeString('zh-HK',{hour:'2-digit',minute:'2-digit',second:'2-digit'});},180)}
function status(id,msg,type=''){const el=$(id);el.textContent=msg;el.className='notice'+(type?' '+type:'')}
function documentLabels(type=state.documentType){if(type==='consignment')return{type:'consignment',title:'Sales Consign',short:'Consignment',date:'Date',no:'Consign No.',items:'Consignment 貨品',confirm:'Confirm Consignment 並更新 jmsdata.xlsx',export:'匯出 Excel Consignment'};if(type==='quotation')return{type:'quotation',title:'Quotation',short:'Quotation',date:'Date',no:'Quotation No.',items:'Quotation 貨品',confirm:'Confirm Quotation 並更新 jmsdata.xlsx（不扣庫存）',export:'匯出 Excel Quotation'};return{type:'invoice',title:'Sales Invoice',short:'Invoice',date:'Date',no:'Invoice No.',items:'Invoice 貨品',confirm:'Confirm Invoice 並更新 jmsdata.xlsx',export:'匯出 Excel Invoice'}}
function updateDocumentTypeUI(){
  const l=documentLabels();
  $('#documentDataHeading').textContent=`客人 ${l.short} 資料`;
  $('#documentNoLabel').textContent=l.no;
  $('#documentDateLabel').textContent=l.date;
  $('#documentItemsHeading').textContent=l.items;
  $('#confirmInvoiceBtn').textContent=l.confirm;
  const bulk=$('#markAllDeliveredBtn');if(bulk)bulk.classList.toggle('hidden',state.documentType!=='invoice');const pdfWrap=$('#pdfInvoiceImportWrap');if(pdfWrap)pdfWrap.classList.toggle('hidden',state.documentType!=='invoice');
  $('#exportExcelBtn').textContent=l.export;
  $('#invoiceNo').placeholder=`${documentPrefix()}YY0001`;
  setDefaultInvoiceNo(true);
  if(state.documentType!=='quotation'&&state.quote.karat!=='18K')state.quote.karat='18K';
  $$('input[name="quoteKarat"]').forEach(x=>x.checked=x.value===state.quote.karat);
  updateGoldQuoteUI();syncEffectivePrices();renderItems();syncPreviewAddonControls();renderPreview();
}
$$('input[name="documentType"]').forEach(r=>r.addEventListener('change',e=>{if(state.recall){e.preventDefault();e.target.checked=false;const current=$(`input[name="documentType"][value="${state.recall.type}"]`);if(current)current.checked=true;return alert('Recall 修改期間不能轉換文件類型；請先完成或取消 Recall。')}state.documentType=e.target.value;updateDocumentTypeUI()}));
function setImportCollapsed(key,collapsed=true){const card=document.querySelector(`[data-import-card="${key}"]`);if(!card)return;card.classList.toggle('collapsed',collapsed);const btn=card.querySelector('.import-toggle');if(btn){btn.textContent=collapsed?'展開':'收合';btn.setAttribute('aria-expanded',String(!collapsed))}}
$$('.import-toggle').forEach(btn=>btn.addEventListener('click',()=>{const card=btn.closest('.import-card');setImportCollapsed(card?.dataset.importCard,!card.classList.contains('collapsed'))}));
function field(row,names){const keys=Object.keys(row);for(const n of names){const k=keys.find(x=>x.trim().toUpperCase()===n);if(k)return row[k]}return''}

function fieldKey(row,names){const keys=Object.keys(row||{});for(const n of names){const k=keys.find(x=>x.trim().toUpperCase()===String(n).toUpperCase());if(k)return k}return''}
function setField(row,names,value){let key=fieldKey(row,names);if(!key){key=names[0];if(!state.stockHeaders.includes(key))state.stockHeaders.push(key)}row[key]=value}
const DOCUMENT_STORE_KEY='universeDocumentStore_v1';
const UNIVERSE_RECORDS_SHEET='Universe Records',UNIVERSE_RECORDS_SCHEMA_VERSION=1,UNIVERSE_RECORDS_CHUNK_SIZE=30000;
const DOCUMENT_SHEETS={invoiceHeaders:'Invoice Header',invoiceItems:'Invoice Items',consignmentHeaders:'Consignment Header',consignmentItems:'Consignment Items',quotationHeaders:'Quotation Header',quotationItems:'Quotation Items',transactions:'Transaction History'};
const HEADER_COLUMNS=['DOCUMENT_NO','EXHIBITION_SESSION','EXHIBITION_NAME','REVISION','DOCUMENT_STATUS','DOCUMENT_DATE','CUSTOMER_CODE','CUSTOMER','ADDRESS','SALES_RATE','CURRENCY','FX_RATE','SHIPMENT_METHOD','TERMS','DISCOUNT','REMARK','QUOTE_KARAT','CURRENT_LONDON_PM','CURRENT_LONDON_PM_DATE','GOLD_DATA_NAME','QUOTE_HISTORICAL_PM_JSON','CREATED_AT','UPDATED_AT'];
const ITEM_COLUMNS=['DOCUMENT_NO','EXHIBITION_SESSION','REVISION','ITEM_SEQ','LOTNO','ARTNO','ARTICLE','DESC1','DESC2','DESC3','DESC4','DESC5','DESC6','QTY','UNIT','PRICE','USD_UNIT_PRICE','UNIT_PRICE','CURRENCY','MANUAL_PRICE_OVERRIDE','IMAGE_VARIANT','IMAGE_GRAYSCALE','IMAGE_AUTO_MATCHED','CUSTOM_IMAGE_FILE','COMPLETION_DATE','QUOTE_HISTORICAL_PM','QUOTE_HISTORICAL_PM_DATE','QUOTE_HISTORICAL_COMPANY_GOLD','QUOTE_CURRENT_COMPANY_GOLD','QUOTE_CALCULATED_USD_UNIT_PRICE','ORIGINAL_STATUS','ORIGINAL_I','ORIGINAL_J','ORIGINAL_K','ORIGINAL_L','ORIGINAL_M','AFTER_STATUS','AFTER_I','AFTER_J','AFTER_K','AFTER_L','AFTER_M'];
const TRANSACTION_COLUMNS=['TIMESTAMP','DOCUMENT_TYPE','DOCUMENT_NO','EXHIBITION_SESSION','REVISION','LOTNO','ARTNO','FROM_STATUS','TO_STATUS','FROM_I','FROM_J','FROM_K','FROM_L','FROM_M','TO_I','TO_J','TO_K','TO_L','TO_M'];
const IMAGE_OVERRIDE_SHEET='Image Overrides';
const IMAGE_OVERRIDE_COLUMNS=['LOTNO','ARTNO','IMAGE_FILE','IMAGE_VARIANT','GRAYSCALE','MODE','UPDATED_AT'];
function imageOverrideRow(raw){
  const lot=norm(field(raw,['LOTNO'])),art=normArt(field(raw,['ARTNO'])),fileName=norm(field(raw,['IMAGE_FILE','IMAGE FILE','FILE_NAME','FILENAME']));
  if(!lot||!fileName)return null;
  return{LOTNO:lot,ARTNO:art,IMAGE_FILE:fileName,IMAGE_VARIANT:norm(field(raw,['IMAGE_VARIANT','IMAGE VARIANT','VARIANT'])),GRAYSCALE:numberValue(field(raw,['GRAYSCALE']))?1:0,MODE:'MANUAL',UPDATED_AT:norm(field(raw,['UPDATED_AT','UPDATED AT']))||new Date().toISOString()};
}
function saveImageOverridesLocal(){try{localStorage.setItem(IMAGE_OVERRIDE_LOCAL_KEY,JSON.stringify([...state.imageOverrides.values()]))}catch(err){console.warn('Image Overrides 無法寫入本機儲存。',err)}}
function loadImageOverridesLocal(){try{const rows=JSON.parse(localStorage.getItem(IMAGE_OVERRIDE_LOCAL_KEY)||'[]'),map=new Map();for(const raw of Array.isArray(rows)?rows:[]){const row=imageOverrideRow(raw);if(row)map.set(String(row.LOTNO),row)}state.imageOverrides=map}catch{state.imageOverrides=new Map()}}
function importImageOverridesFromWorkbook(wb){
  const incoming=new Map();for(const raw of sheetRows(wb,IMAGE_OVERRIDE_SHEET)){const row=imageOverrideRow(raw);if(row)incoming.set(String(row.LOTNO),row)}
  const conflicts=[];for(const [lot,row] of incoming){const prev=state.imageOverrides.get(lot);if(prev&&stableRowJSON(prev)!==stableRowJSON(row))conflicts.push({lot,local:prev,incoming:row})}
  let merge=true;if(conflicts.length)merge=confirm(`匯入 jmsdata.xlsx 時發現 ${conflicts.length} 項 Image Override 與本機記錄不同。\n\n按「確定」：保留本機選圖並補入 jmsdata 內其他記錄。\n按「取消」：保留全部本機 Image Override，不匯入衝突資料。`);
  if(merge){for(const [lot,row] of incoming)if(!state.imageOverrides.has(lot))state.imageOverrides.set(lot,row)}
  const dirty=new Set();for(const [lot,row] of state.imageOverrides){const source=incoming.get(lot);if(!source||stableRowJSON(source)!==stableRowJSON(row))dirty.add(lot)}state.imageOverrideDirtyLots=dirty;state.imageOverrideDirty=dirty.size;saveImageOverridesLocal();updateImageOverrideStatus();return{conflicts:conflicts.length,dirty:dirty.size};
}
function imageOverrideForLot(lot){return state.imageOverrides.get(String(lot||''))||null}
function findImageByFileName(fileName,artNo=''){
  const exact=state.imageFilesByName.get(String(fileName||'').toUpperCase());if(exact)return exact;
  const arr=state.imageFiles.get(normArt(artNo))||[];return arr.find(x=>String(x.fileName||'').toUpperCase()===String(fileName||'').toUpperCase())||null;
}
function resolvedImageOverride(product){
  const row=imageOverrideForLot(product?.lotNo);if(!row)return null;
  const image=findImageByFileName(row.IMAGE_FILE,row.ARTNO||product?.artNo);if(!image)return null;
  return{row,image,variant:row.IMAGE_VARIANT||image.variant,grayscale:!!numberValue(row.GRAYSCALE),fileName:image.fileName,manualOverride:true};
}
function updateImageOverrideStatus(){
  const el=$('#imageOverrideStatus'),count=state.imageOverrides.size,dirty=state.imageOverrideDirty,newImages=state.packageImageDirtyFiles?.size||0;
  if(!el)return;
  if(dirty||newImages){
    const parts=[];if(dirty)parts.push(`${dirty} 項選圖記錄尚未寫入最新 jmsdata.xlsx`);if(newImages)parts.push(`${newImages} 張新增圖片只在目前 PWA 暫存`);
    el.textContent=parts.join('；')+'。';el.className='notice warn';
  }else if(count){el.textContent=`已載入 ${count} 項 LOTNO 手動選圖記錄。`;el.className='notice ok'}
  else{el.textContent='尚未有 LOTNO 手動選圖記錄。';el.className='notice'}
}
function confirmDiscardPendingImageOverrides(){const pending=state.imageOverrideDirty+(state.packageImageDirtyFiles?.size||0);return !pending||confirm(`尚有 ${pending} 項圖片修改只在目前 PWA 暫存。

手動選圖記錄可隨最新 jmsdata.xlsx 保存；新上傳／拍攝的圖片檔本身不會嵌入 jmsdata.xlsx，重新匯入／更換 Folder 後可能無法帶走。

仍然繼續？`)}
function saveImageOverride(product,image,grayscale=false){
  if(!product||!image)return;
  const key=String(product.lotNo),previous=state.imageOverrides.get(key),row={LOTNO:key,ARTNO:normArt(product.artNo),IMAGE_FILE:image.fileName,IMAGE_VARIANT:image.variant||'',GRAYSCALE:grayscale?1:0,MODE:'MANUAL',UPDATED_AT:new Date().toISOString()};
  state.imageOverrides.set(key,row);
  if(previous&&String(previous.IMAGE_FILE||'').toUpperCase()!==String(row.IMAGE_FILE||'').toUpperCase()){
    const oldKey=String(previous.IMAGE_FILE||'').toUpperCase(),stillUsed=[...state.imageOverrides.values()].some(entry=>String(entry.IMAGE_FILE||'').toUpperCase()===oldKey);
    if(!stillUsed&&state.customPackageImages.has(oldKey)){state.customPackageImages.delete(oldKey);state.packageImageDirtyFiles.delete(oldKey)}
  }
  if(image.manualSource&&image.file){const imageKey=String(image.fileName||'').toUpperCase();state.customPackageImages.set(imageKey,image.file);state.packageImageDirtyFiles.add(imageKey)}
  if(!previous||previous.ARTNO!==row.ARTNO||previous.IMAGE_FILE!==row.IMAGE_FILE||numberValue(previous.GRAYSCALE)!==numberValue(row.GRAYSCALE)){state.imageOverrideDirtyLots.add(key);state.imageOverrideDirty=state.imageOverrideDirtyLots.size}
  applyImageOverrideToCurrentItem(product);saveImageOverridesLocal();updateImageOverrideStatus();invalidateStockSearchImages(key);renderStockSearch();renderItems();runHealthCheck();schedulePreview();
}
function clearImageOverride(product){
  const key=String(product?.lotNo||''),previous=state.imageOverrides.get(key);if(!previous)return;
  state.imageOverrides.delete(key);const imageKey=String(previous.IMAGE_FILE||'').toUpperCase(),stillUsed=[...state.imageOverrides.values()].some(row=>String(row.IMAGE_FILE||'').toUpperCase()===imageKey);
  if(!stillUsed&&state.customPackageImages.has(imageKey)){state.customPackageImages.delete(imageKey);state.packageImageDirtyFiles.delete(imageKey)}
  state.imageOverrideDirtyLots.add(key);state.imageOverrideDirty=state.imageOverrideDirtyLots.size;applyImageOverrideToCurrentItem(product);saveImageOverridesLocal();updateImageOverrideStatus();invalidateStockSearchImages(key);renderStockSearch();renderItems();runHealthCheck();schedulePreview();
}
function applyImageOverrideToCurrentItem(product){
  const item=state.items.find(x=>String(x.lotNo)===String(product?.lotNo));if(!item)return;
  const match=chooseImageMatch(product);item.imageVariant=match.variant;item.imageGrayscale=!!match.grayscale;item.imageAutoMatched=!match.manualOverride;item.imageOverrideFile=match.fileName||'';
}
function emptyDocumentStore(){return{invoiceHeaders:[],invoiceItems:[],consignmentHeaders:[],consignmentItems:[],quotationHeaders:[],quotationItems:[],transactions:[]}}
function resetFormalRecordContext(){state.documentStore=emptyDocumentStore();state.recordCounters={invoice:1,consignment:1,quotation:1};state.recordsLoaded=false;state.recordsDirty=false;state.recall=null;state.deliveryReturns=new Set();state.exhibitionSession='';state.recordsFileName='jmsdata.xlsx / Universe Records';state.recordsFilePath='';delete state.dataMeta.records;setRecallLock(false);$('#recallActive')?.classList.add('hidden');renderRecallResults();try{localStorage.removeItem(DOCUMENT_STORE_KEY);localStorage.removeItem(EXHIBITION_SESSION_KEY)}catch{}}
function normalizeDocumentStore(raw){const base=emptyDocumentStore();for(const key of Object.keys(base))base[key]=Array.isArray(raw?.[key])?raw[key]:[];return base}
function loadDocumentStore(){state.documentStore=emptyDocumentStore();try{localStorage.removeItem(DOCUMENT_STORE_KEY)}catch{}}
function saveDocumentStore(){}
function normalizeRecordCounters(raw){const out={invoice:1,consignment:1,quotation:1};for(const type of Object.keys(out)){const n=Number(raw?.[type]);if(Number.isInteger(n)&&n>0)out[type]=n}return out}
function newExhibitionId(){const d=new Date(),p=n=>String(n).padStart(2,'0');return `EXH-${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}-${Math.random().toString(36).slice(2,7).toUpperCase()}`}
function officialDocumentCount(){return state.documentStore.invoiceHeaders.length+state.documentStore.consignmentHeaders.length+state.documentStore.quotationHeaders.length}
function universeRecordsPayload(){if(!state.exhibitionSession)state.exhibitionSession=newExhibitionId();if(!state.exhibitionName)saveExhibitionName(state.packageName||'Exhibition');return{app:'Universe Invoice',recordType:'Universe Records',schemaVersion:UNIVERSE_RECORDS_SCHEMA_VERSION,appVersion:APP_VERSION,exhibitionId:state.exhibitionSession,exhibitionName:state.exhibitionName||'',updatedAt:new Date().toISOString(),nextSequence:normalizeRecordCounters(state.recordCounters),documentStore:normalizeDocumentStore(state.documentStore)}}
function universeRecordsJson(){return JSON.stringify(universeRecordsPayload())}
function applyUniverseRecords(raw,{fallbackName='',fileName='jmsdata.xlsx',filePath=''}={}){if(!raw||raw.app!=='Universe Invoice'||raw.recordType!=='Universe Records'||Number(raw.schemaVersion)!==UNIVERSE_RECORDS_SCHEMA_VERSION)throw new Error('jmsdata.xlsx 的 Universe Records 工作表格式不正確或版本不支援。');state.documentStore=normalizeDocumentStore(raw.documentStore);state.recordCounters=normalizeRecordCounters(raw.nextSequence);state.exhibitionSession=norm(raw.exhibitionId)||newExhibitionId();state.recordsFileName=`${fileName||'jmsdata.xlsx'} / ${UNIVERSE_RECORDS_SHEET}`;state.recordsFilePath=filePath||'';state.recordsLoaded=true;state.recordsDirty=!norm(raw.exhibitionId)||!norm(raw.exhibitionName);try{localStorage.removeItem(DOCUMENT_STORE_KEY);localStorage.setItem(EXHIBITION_SESSION_KEY,state.exhibitionSession)}catch{}saveExhibitionName(norm(raw.exhibitionName)||norm(fallbackName)||state.packageName||'Exhibition');syncDocumentSequencesFromStore();renderRecallResults();updateDataVersionPanel();runHealthCheck();return{documents:officialDocumentCount(),exhibitionId:state.exhibitionSession,exhibitionName:state.exhibitionName,newTemplate:!norm(raw.exhibitionId)}}
function emptyUniverseRecordsForWorkbook(fallbackName='',fileName='jmsdata.xlsx'){state.documentStore=emptyDocumentStore();state.recordCounters={invoice:1,consignment:1,quotation:1};state.recall=null;state.deliveryReturns=new Set();state.exhibitionSession=newExhibitionId();state.recordsFileName=`${fileName||'jmsdata.xlsx'} / ${UNIVERSE_RECORDS_SHEET}`;state.recordsFilePath='';state.recordsLoaded=true;state.recordsDirty=true;state.draft.prompted=false;clearDocumentDraft();clearCurrentDocument({keepCustomer:false});try{localStorage.removeItem(DOCUMENT_STORE_KEY);localStorage.setItem(EXHIBITION_SESSION_KEY,state.exhibitionSession)}catch{}saveExhibitionName(norm(fallbackName)||'Exhibition');syncDocumentSequencesFromStore();renderRecallResults();updateDataVersionPanel();return{documents:0,exhibitionId:state.exhibitionSession,exhibitionName:state.exhibitionName,newTemplate:true}}
function universeRecordsWorksheet(){const json=universeRecordsJson(),rows=[['Universe Records','Embedded in jmsdata.xlsx'],['Schema Version',UNIVERSE_RECORDS_SCHEMA_VERSION],['Format','JSON chunks'],['Chunk','Data']];for(let i=0,n=1;i<json.length;i+=UNIVERSE_RECORDS_CHUNK_SIZE,n++)rows.push([n,json.slice(i,i+UNIVERSE_RECORDS_CHUNK_SIZE)]);const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=[{wch:12},{wch:42}];return ws}
function universeRecordsRawFromWorkbook(wb){const ws=wb?.Sheets?.[UNIVERSE_RECORDS_SHEET];if(!ws)return null;const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});let start=rows.findIndex(r=>norm(r?.[0]).toUpperCase()==='CHUNK'&&norm(r?.[1]).toUpperCase()==='DATA');if(start<0)return null;const chunks=rows.slice(start+1).filter(r=>Number(r?.[0])>0&&r?.[1]!==undefined&&r?.[1]!==null&&String(r[1]).length).sort((a,b)=>Number(a[0])-Number(b[0])).map(r=>String(r[1]));if(!chunks.length)return null;try{return JSON.parse(chunks.join(''))}catch(err){throw new Error('jmsdata.xlsx 的 Universe Records 工作表內容損壞：'+(err.message||err))}}
function importUniverseRecordsFromWorkbook(wb,{fallbackName='',fileName='jmsdata.xlsx',lastModified=0}={}){const raw=universeRecordsRawFromWorkbook(wb);const info=raw?applyUniverseRecords(raw,{fallbackName,fileName}):emptyUniverseRecordsForWorkbook(fallbackName,fileName);state.dataMeta.records={name:`${fileName||'jmsdata.xlsx'} / ${UNIVERSE_RECORDS_SHEET}`,lastModified:Number(lastModified)||0,count:info.documents,exhibitionId:info.exhibitionId,exhibitionName:info.exhibitionName,newTemplate:info.newTemplate,embedded:!!raw};updateDataVersionPanel();return info}
function setWorkbookSheetHidden(wb,name,hidden=1){if(!wb?.SheetNames?.includes(name))return;wb.Workbook=wb.Workbook||{};const existing=Array.isArray(wb.Workbook.Sheets)?wb.Workbook.Sheets:[];const byName=new Map(existing.map((x,i)=>[x?.name||wb.SheetNames[i],x||{}]));wb.Workbook.Sheets=wb.SheetNames.map(n=>{const meta=byName.get(n)||{name:n};meta.name=n;if(n===name)meta.Hidden=hidden;else if(meta.Hidden===undefined)meta.Hidden=0;return meta})}
function stableRowJSON(row){const out={};for(const k of Object.keys(row||{}).sort())out[k]=row[k];return JSON.stringify(out)}
function documentRowKey(storeKey,row){const no=norm(docValue(row,'DOCUMENT_NO')).toUpperCase(),session=recordSession(row),revision=revisionNumber(row);if(storeKey.endsWith('Headers'))return `${session}|${no}|${revision}`;if(storeKey.endsWith('Items'))return `${session}|${no}|${revision}|${numberValue(docValue(row,'ITEM_SEQ'))}`;return `${norm(docValue(row,'TIMESTAMP'))}|${session}|${no}|${revision}|${norm(docValue(row,'LOTNO'))}|${norm(docValue(row,'FROM_STATUS'))}|${norm(docValue(row,'TO_STATUS'))}`}
function mergeDocumentStores(localStore,incomingStore){
  const merged=emptyDocumentStore(),conflicts=[];
  for(const key of Object.keys(merged)){
    const map=new Map();for(const row of localStore?.[key]||[])map.set(documentRowKey(key,row),row);
    for(const row of incomingStore?.[key]||[]){const id=documentRowKey(key,row),prev=map.get(id);if(!prev){map.set(id,row);continue}if(stableRowJSON(prev)===stableRowJSON(row))continue;conflicts.push({kind:key,key:id,local:prev,incoming:row});if(key.endsWith('Headers')){const a=String(docValue(prev,'UPDATED_AT')||''),b=String(docValue(row,'UPDATED_AT')||'');if(b>a)map.set(id,row)}}
    merged[key]=[...map.values()];
  }
  return{store:normalizeDocumentStore(merged),conflicts};
}
function sheetRows(wb,name){const ws=wb?.Sheets?.[name];return ws?XLSX.utils.sheet_to_json(ws,{defval:''}):[]}
function syncDocumentSequencesFromStore(){const yy=invoiceYear();for(const type of ['invoice','consignment','quotation']){const prefix=documentPrefix(type);let next=getNextDocumentSequence(type,yy);for(const h of latestDocumentHeaders(type,{session:state.exhibitionSession})){const no=norm(docValue(h,'DOCUMENT_NO')).toUpperCase(),m=no.match(new RegExp(`^${prefix}(\\d{2})(\\d{4})$`));if(m&&m[1]===yy)next=Math.max(next,Number(m[2])+1)}state.recordCounters[type]=next}if(!state.recall&&!state.items.length)setDefaultInvoiceNo(true)}
function importDocumentStoreFromWorkbook(wb){
  const incoming=emptyDocumentStore();for(const [key,name] of Object.entries(DOCUMENT_SHEETS))incoming[key]=sheetRows(wb,name);
  const {store,conflicts}=mergeDocumentStores(state.documentStore,incoming);state.importConflicts=conflicts;
  let useMerged=true;if(conflicts.length)useMerged=confirm(`匯入 jmsdata.xlsx 時發現 ${conflicts.length} 項本機文件記錄與 jmsdata 同一 Revision 內容不同。\n\n按「確定」：合併資料，Header 以較新 UPDATED_AT 為準，其他衝突保留本機版本。\n按「取消」：只更新庫存資料，不匯入 jmsdata 內的文件記錄。`);
  if(useMerged)state.documentStore=store;
  if(!state.exhibitionSession){const headers=[...state.documentStore.invoiceHeaders,...state.documentStore.consignmentHeaders,...state.documentStore.quotationHeaders].filter(x=>recordSession(x)).sort((a,b)=>String(docValue(b,'UPDATED_AT')||docValue(b,'DOCUMENT_DATE')).localeCompare(String(docValue(a,'UPDATED_AT')||docValue(a,'DOCUMENT_DATE')))),latest=headers[0];if(latest){const session=recordSession(latest),name=norm(docValue(latest,'EXHIBITION_NAME'));if(session){state.exhibitionSession=session;try{localStorage.setItem(EXHIBITION_SESSION_KEY,session)}catch{}}if(name)saveExhibitionName(name)}}
  saveDocumentStore();syncDocumentSequencesFromStore();renderRecallResults();return{conflicts:conflicts.length,merged:useMerged};
}
function storeKeysForType(type){if(type==='consignment')return{headers:'consignmentHeaders',items:'consignmentItems'};if(type==='quotation')return{headers:'quotationHeaders',items:'quotationItems'};return{headers:'invoiceHeaders',items:'invoiceItems'}}
function docValue(row,name){return field(row,[name])}
function revisionNumber(row){return Number(docValue(row,'REVISION'))||0}
function recordSession(row){return norm(docValue(row,'EXHIBITION_SESSION'))}
function latestDocumentHeaders(type,{session=null}={}){const keys=storeKeysForType(type),latest=new Map();for(const row of state.documentStore[keys.headers]){const no=norm(docValue(row,'DOCUMENT_NO')).toUpperCase(),rowSession=recordSession(row);if(!no||session!==null&&rowSession!==norm(session))continue;const key=`${rowSession}|${no}`,prev=latest.get(key);if(!prev||revisionNumber(row)>revisionNumber(prev)||revisionNumber(row)===revisionNumber(prev)&&String(docValue(row,'UPDATED_AT'))>String(docValue(prev,'UPDATED_AT')))latest.set(key,row)}return[...latest.values()].sort((a,b)=>String(docValue(b,'UPDATED_AT')||docValue(b,'DOCUMENT_DATE')).localeCompare(String(docValue(a,'UPDATED_AT')||docValue(a,'DOCUMENT_DATE'))))}
function findLatestDocument(type,documentNo,session=state.exhibitionSession){const no=norm(documentNo).toUpperCase(),wantedSession=norm(session),header=latestDocumentHeaders(type,{session:wantedSession}).find(x=>norm(docValue(x,'DOCUMENT_NO')).toUpperCase()===no);if(!header)return null;const revision=revisionNumber(header),keys=storeKeysForType(type),items=state.documentStore[keys.items].filter(x=>norm(docValue(x,'DOCUMENT_NO')).toUpperCase()===no&&recordSession(x)===wantedSession&&revisionNumber(x)===revision).sort((a,b)=>(Number(docValue(a,'ITEM_SEQ'))||0)-(Number(docValue(b,'ITEM_SEQ'))||0));return{header,items,revision,session:wantedSession}}
function numberValue(v){const n=Number(v);return Number.isFinite(n)?n:0}
function inventoryVectorFromRow(row){const i=numberValue(field(row,['POQTY','DELIVERED'])),j=numberValue(field(row,['FQTY','AVAILABLE'])),k=numberValue(field(row,['RQTY','SOLD ON HAND','SOLD_ON_HAND'])),l=numberValue(field(row,['COQTY','CONSIGNMENT']));let m=numberValue(field(row,['TQTY','BALANCE']));if(!m)m=i+j+k+l||1;return{i,j,k,l,m}}
function vectorFromRecord(row,prefix){const p=String(prefix||'').toUpperCase();return{i:numberValue(docValue(row,`${p}_I`)),j:numberValue(docValue(row,`${p}_J`)),k:numberValue(docValue(row,`${p}_K`)),l:numberValue(docValue(row,`${p}_L`)),m:numberValue(docValue(row,`${p}_M`))}}
function inventoryStatusFromVector(v){const positives=[['SOLD_DELIVERED',v.i],['AVAILABLE',v.j],['SOLD_ON_HAND',v.k],['CONSIGNED',v.l]].filter(([,n])=>numberValue(n)>0);return positives.length===1?positives[0][0]:positives.length?'CONFLICT':'UNKNOWN'}
function inventoryVectorForStatus(status,total=1){const m=Math.max(1,numberValue(total)||1),v={i:0,j:0,k:0,l:0,m};if(status==='SOLD_DELIVERED')v.i=m;else if(status==='SOLD_ON_HAND')v.k=m;else if(status==='CONSIGNED')v.l=m;else v.j=m;return v}
function currentInventoryStatusForLot(lot){const row=state.stockRowByLot.get(String(lot||''));return row?inventoryStatusFromVector(inventoryVectorFromRow(row)):'UNKNOWN'}
function initialDeliveredChoiceForLot(lot,fallbackStatus=''){const current=currentInventoryStatusForLot(lot);if(current==='SOLD_DELIVERED')return true;if(current==='SOLD_ON_HAND')return false;return norm(fallbackStatus).toUpperCase()==='SOLD_DELIVERED'}
function setInvoiceItemDelivered(item,next,{confirmReturn=true}={}){
  if(state.documentType!=='invoice'||!item)return false;
  next=!!next;const lot=String(item.lotNo||''),current=currentInventoryStatusForLot(lot);
  if(!next&&current==='SOLD_DELIVERED'&&confirmReturn){if(!confirm(`${item.artNo} / LOTNO ${lot}
目前 jmsdata 是 Sold - Delivered。

確認貨品已退回／改為「未交貨」？
真正更新會在 Confirm Invoice 時才寫入 jmsdata.xlsx。`))return false}
  item.delivered=next;
  if(current==='SOLD_DELIVERED'&&!next)state.deliveryReturns.add(lot);else state.deliveryReturns.delete(lot);
  renderItems();scheduleDraftSave();return true;
}
function markAllInvoiceDelivered(){if(state.documentType!=='invoice')return;if(!state.items.length)return alert('Invoice 尚未有貨品。');for(const item of state.items){item.delivered=true;state.deliveryReturns.delete(String(item.lotNo))}renderItems();scheduleDraftSave();status('#addMessage',`Invoice 內 ${state.items.length} 款已全部標記為「已交貨」；jmsdata 會在 Confirm 時才更新。`,'ok')}
function deliveredRecallItemsNeedingReturn(){if(state.documentType!=='invoice'||!state.recall)return[];return state.items.filter(item=>currentInventoryStatusForLot(item.lotNo)==='SOLD_DELIVERED'&&!state.deliveryReturns.has(String(item.lotNo)))}
function vectorsEqual(a,b){return['i','j','k','l','m'].every(k=>numberValue(a?.[k])===numberValue(b?.[k]))}
function vectorFields(prefix,v){const p=String(prefix).toUpperCase();return{[`${p}_STATUS`]:inventoryStatusFromVector(v),[`${p}_I`]:numberValue(v.i),[`${p}_J`]:numberValue(v.j),[`${p}_K`]:numberValue(v.k),[`${p}_L`]:numberValue(v.l),[`${p}_M`]:numberValue(v.m)}}
function writeInventoryVector(row,v){setField(row,['POQTY','DELIVERED'],numberValue(v.i));setField(row,['FQTY','AVAILABLE'],numberValue(v.j));setField(row,['RQTY','SOLD ON HAND'],numberValue(v.k));setField(row,['COQTY','CONSIGNMENT'],numberValue(v.l));setField(row,['TQTY','BALANCE'],numberValue(v.m))}
function rebuildInventoryMaps(){const available=new Map(),history=new Map(),availableRows=[],previous=state.inventoryHistory instanceof Map?state.inventoryHistory:new Map();for(const [lot,p] of state.stockCatalog){const row=state.stockRowByLot.get(String(lot)),status=inventoryStatusFromVector(inventoryVectorFromRow(row||{})),old=previous.get(String(lot))||{};if(status==='AVAILABLE'){available.set(String(lot),p);if(row)availableRows.push(row)}else history.set(String(lot),{...old,...productSnapshot(p),status,updatedAt:Date.now(),docs:Array.isArray(old.docs)?old.docs:[]})}state.products=available;state.stockRows=availableRows;state.inventoryHistory=history;if(state.dataMeta?.stock)state.dataMeta.stock.available=available.size;saveInventoryHistory();updateDataVersionPanel();invalidateStockSearchData();renderStockSearch();updateTotals()}
function setInventoryVectorForLot(lot,vector){const key=String(lot),row=state.stockRowByLot.get(key);if(!row)throw new Error(`jmsdata 找不到 LOTNO ${key}`);writeInventoryVector(row,vector);return vector}
function appendTransaction({type,documentNo,revision,item,before,after,session=state.recall?.session??state.exhibitionSession}){if(vectorsEqual(before,after))return;state.documentStore.transactions.push({TIMESTAMP:new Date().toISOString(),DOCUMENT_TYPE:type==='consignment'?'Consignment':'Invoice',DOCUMENT_NO:documentNo,EXHIBITION_SESSION:session,REVISION:revision,LOTNO:item.lotNo,ARTNO:item.artNo,FROM_STATUS:inventoryStatusFromVector(before),TO_STATUS:inventoryStatusFromVector(after),FROM_I:before.i,FROM_J:before.j,FROM_K:before.k,FROM_L:before.l,FROM_M:before.m,TO_I:after.i,TO_J:after.j,TO_K:after.k,TO_L:after.l,TO_M:after.m})}
function headerRowForCurrentDocument(type,documentNo,revision,createdAt='',documentStatus='CONFIRMED'){const now=new Date().toISOString(),isQuote=type==='quotation',session=state.recall?.session??state.exhibitionSession,exhibitionName=state.recall?norm(docValue(state.recall.header,'EXHIBITION_NAME')):(state.exhibitionName||norm($('#exhibitionName')?.value)||state.packageName);return{DOCUMENT_NO:documentNo,EXHIBITION_SESSION:session,EXHIBITION_NAME:exhibitionName,REVISION:revision,DOCUMENT_STATUS:documentStatus,DOCUMENT_DATE:norm($('#invoiceDate').value),CUSTOMER_CODE:norm($('#customerCode').value),CUSTOMER:norm($('#customerName').value),ADDRESS:norm($('#customerAddress').value),SALES_RATE:numberValue($('#salesRate').value),CURRENCY:currencyCode(),FX_RATE:currencyCode()==='USD'?1:currentFxRate(),SHIPMENT_METHOD:norm($('#shipmentMethod').value),TERMS:norm($('#customerTerms').value),DISCOUNT:numberValue($('#discountAmount').value),REMARK:norm($('#remark').value),QUOTE_KARAT:isQuote?state.quote.karat:'',CURRENT_LONDON_PM:isQuote?numberValue(state.quote.currentLondonPm):'',CURRENT_LONDON_PM_DATE:isQuote?normalizeStockDate(state.quote.currentLondonPmDate):'',GOLD_DATA_NAME:isQuote?norm(state.quote.goldDataName):'',QUOTE_HISTORICAL_PM_JSON:isQuote?JSON.stringify(state.quote.historicalPm||{}):'',CREATED_AT:createdAt||now,UPDATED_AT:now}}
function itemRowForDocument(item,documentNo,revision,original,after){const desc=[...(item.descriptions||[])],isQuote=state.documentType==='quotation',q=isQuote?quotationPriceDetails(item):null,session=state.recall?.session??state.exhibitionSession;return{DOCUMENT_NO:documentNo,EXHIBITION_SESSION:session,REVISION:revision,ITEM_SEQ:Number(item.seq)||1,LOTNO:item.lotNo,ARTNO:item.artNo,ARTICLE:item.article||'',DESC1:desc[0]||'',DESC2:desc[1]||'',DESC3:desc[2]||'',DESC4:desc[3]||'',DESC5:desc[4]||'',DESC6:desc[5]||'',QTY:Number(item.qty)||1,UNIT:item.unit||'PC',PRICE:Number(item.price)||0,USD_UNIT_PRICE:Number(item.usdUnitPrice)||0,UNIT_PRICE:Number(item.unitPrice)||0,CURRENCY:currencyCode(),MANUAL_PRICE_OVERRIDE:isManualPriceOverride(item)?1:0,IMAGE_VARIANT:item.imageVariant||'',IMAGE_GRAYSCALE:item.imageGrayscale?1:0,IMAGE_AUTO_MATCHED:item.imageAutoMatched?1:0,CUSTOM_IMAGE_FILE:item.customImage?.fileName||'',COMPLETION_DATE:normalizeStockDate(item.completionDate),QUOTE_HISTORICAL_PM:isQuote&&q?.ready?q.historicalPm:'',QUOTE_HISTORICAL_PM_DATE:isQuote&&q?.ready?q.historicalSourceDate:'',QUOTE_HISTORICAL_COMPANY_GOLD:isQuote&&q?.ready?q.historical.base:'',QUOTE_CURRENT_COMPANY_GOLD:isQuote&&q?.ready?q.current.base:'',QUOTE_CALCULATED_USD_UNIT_PRICE:isQuote&&q?.ready?q.finalPrice:'',...vectorFields('ORIGINAL',original),...vectorFields('AFTER',after)}}
function worksheetFromRows(rows,headers){return rows.length?XLSX.utils.json_to_sheet(rows,{header:headers}):XLSX.utils.aoa_to_sheet([headers])}
function preserveWorksheetLayout(target,source){if(!target||!source)return target;for(const prop of ['!cols','!rows','!merges','!autofilter','!freeze','!margins'])if(source[prop])target[prop]=structuredClone(source[prop]);for(const address of Object.keys(target)){if(address[0]==='!')continue;const old=source[address],cell=target[address];if(!old||!cell)continue;if(old.s)cell.s=structuredClone(old.s);if(old.z)cell.z=old.z}return target}
function buildJmsdataWorkbook(){if(!state.stockAllRows.length)throw new Error('尚未匯入 jmsdata。');if(!state.recordsLoaded)emptyUniverseRecordsForWorkbook(state.packageName||state.exhibitionName||'Exhibition',state.stockFileName||'jmsdata.xlsx');const wb=XLSX.utils.book_new(),stockName=(state.stockSheetName||'jmsdata').slice(0,31),originalStock=state.stockWorkbook?.Sheets?.[state.stockSheetName],stockWs=preserveWorksheetLayout(XLSX.utils.json_to_sheet(state.stockAllRows,{header:state.stockHeaders}),originalStock);if(state.stockWorkbook?.Props)wb.Props=structuredClone(state.stockWorkbook.Props);if(state.stockWorkbook?.Custprops)wb.Custprops=structuredClone(state.stockWorkbook.Custprops);XLSX.utils.book_append_sheet(wb,stockWs,stockName);XLSX.utils.book_append_sheet(wb,universeRecordsWorksheet(),UNIVERSE_RECORDS_SHEET);const managed=new Set([state.stockSheetName,UNIVERSE_RECORDS_SHEET,...Object.values(DOCUMENT_SHEETS),IMAGE_OVERRIDE_SHEET]);if(state.stockWorkbook){for(const name of state.stockWorkbook.SheetNames||[]){if(managed.has(name)||name===state.stockSheetName)continue;XLSX.utils.book_append_sheet(wb,state.stockWorkbook.Sheets[name],name.slice(0,31))}}XLSX.utils.book_append_sheet(wb,worksheetFromRows([...state.imageOverrides.values()],IMAGE_OVERRIDE_COLUMNS),IMAGE_OVERRIDE_SHEET);setWorkbookSheetHidden(wb,UNIVERSE_RECORDS_SHEET,1);return wb}
function exportJmsdata(){try{const wb=buildJmsdataWorkbook();XLSX.writeFile(wb,'jmsdata.xlsx',{bookType:'xlsx'});state.recordsDirty=false;state.imageOverrideDirtyLots=new Set();state.imageOverrideDirty=0;if(state.dataMeta.records){state.dataMeta.records.count=officialDocumentCount();state.dataMeta.records.exhibitionId=state.exhibitionSession;state.dataMeta.records.exhibitionName=state.exhibitionName;state.dataMeta.records.embedded=true;state.dataMeta.records.name=`jmsdata.xlsx / ${UNIVERSE_RECORDS_SHEET}`}updateImageOverrideStatus();updateDataVersionPanel();const el=$('#jmsdataExportStatus');if(el){el.textContent=`已輸出最新 jmsdata.xlsx：Sheet 1 為庫存，Sheet 2「Universe Records」保存 ${officialDocumentCount()} 筆正式文件記錄${state.imageOverrides.size?`，另包括 ${state.imageOverrides.size} 項 LOTNO 圖片選擇`:''}。請在「檔案」App 儲存時確認取代舊檔。`;el.className='notice ok'}return true}catch(err){console.error(err);alert('匯出 jmsdata 失敗：'+(err.message||err));return false}}
function setRecallLock(locked){$$('input[name="documentType"]').forEach(r=>r.disabled=!!locked);const no=$('#invoiceNo');if(no)no.readOnly=!!locked;$('#cancelRecallBtn')?.classList.toggle('hidden',!locked)}
function clearCurrentDocument({keepCustomer=true}={}){releaseCustomImages();state.items=[];state.deliveryReturns=new Set();state.editingItemId=null;$('#discountAmount').value=0;$('#remark').value='';if(!keepCustomer){$('#customerCode').value='';$('#customerName').value='';$('#customerAddress').value='';$('#salesRate').value='';$('#customerTerms').value=''}resetPreviewAddonOptions();renderItems();renderCustomerSummary();schedulePreview()}
function restoreDocumentFields(header){$('#invoiceNo').value=norm(docValue(header,'DOCUMENT_NO'));$('#invoiceDate').value=norm(docValue(header,'DOCUMENT_DATE'))||today();$('#customerCode').value=norm(docValue(header,'CUSTOMER_CODE'));$('#customerName').value=norm(docValue(header,'CUSTOMER'));$('#customerAddress').value=norm(docValue(header,'ADDRESS'));$('#salesRate').value=numberValue(docValue(header,'SALES_RATE'));$('#currency').value=norm(docValue(header,'CURRENCY'))||'USD';$('#fxRate').value=numberValue(docValue(header,'FX_RATE'))||'';state.fx={rate:numberValue(docValue(header,'FX_RATE'))||1,date:'',source:'manual',fetching:false};$('#shipmentMethod').value=norm(docValue(header,'SHIPMENT_METHOD'));$('#customerTerms').value=norm(docValue(header,'TERMS'));$('#discountAmount').value=numberValue(docValue(header,'DISCOUNT'));$('#remark').value=norm(docValue(header,'REMARK'));if(state.documentType==='quotation'){state.quote.karat=['18K','14K','14K_SAME_WEIGHT'].includes(norm(docValue(header,'QUOTE_KARAT')))?norm(docValue(header,'QUOTE_KARAT')):'18K';state.quote.currentLondonPm=numberValue(docValue(header,'CURRENT_LONDON_PM'));state.quote.currentLondonPmDate=normalizeStockDate(docValue(header,'CURRENT_LONDON_PM_DATE'));state.quote.source='recall';const raw=norm(docValue(header,'QUOTE_HISTORICAL_PM_JSON'));if(raw)try{const x=JSON.parse(raw);if(x&&typeof x==='object')state.quote.historicalPm=x}catch{}$$('input[name="quoteKarat"]').forEach(x=>x.checked=x.value===state.quote.karat)}}
function recallItemFromRow(row){const lot=norm(docValue(row,'LOTNO')),p=state.stockCatalog.get(lot)||{lotNo:lot,artNo:normArt(docValue(row,'ARTNO')),price:numberValue(docValue(row,'PRICE')),unit:norm(docValue(row,'UNIT'))||'PC',article:norm(docValue(row,'ARTICLE')),descriptions:[1,2,3,4,5,6].map(n=>norm(docValue(row,`DESC${n}`))).filter(Boolean),completionDate:normalizeStockDate(docValue(row,'COMPLETION_DATE'))};const code=norm(docValue(row,'CURRENCY'))||currencyCode(),unitPrice=numberValue(docValue(row,'UNIT_PRICE')),item={...productSnapshot(p),id:Date.now()+Math.random(),seq:numberValue(docValue(row,'ITEM_SEQ'))||1,qty:numberValue(docValue(row,'QTY'))||1,usdUnitPrice:numberValue(docValue(row,'USD_UNIT_PRICE'))||Math.ceil((Number(p.price)||0)*(Number($('#salesRate').value)||0)),currencyPrices:{},quote18kCurrencyPrices:{},quote14kCurrencyPrices:{},quote14kSameWeightCurrencyPrices:{},manualPriceFlags:{},unitPrice,imageVariant:norm(docValue(row,'IMAGE_VARIANT'))||chooseImageMatch(p).variant,imageGrayscale:numberValue(docValue(row,'IMAGE_GRAYSCALE'))>0,imageAutoMatched:numberValue(docValue(row,'IMAGE_AUTO_MATCHED'))>0,delivered:state.documentType==='invoice'?initialDeliveredChoiceForLot(lot,docValue(row,'AFTER_STATUS')):false,_originalVector:vectorFromRecord(row,'ORIGINAL')};const qExpected=state.documentType==='quotation'?quotationPriceDetails(item):null,expected=convertedFromUsd(qExpected?.ready?qExpected.finalPrice:item.usdUnitPrice);if(numberValue(docValue(row,'MANUAL_PRICE_OVERRIDE'))||Math.abs(unitPrice-expected)>.001){const overrides=activePriceOverrides(item);overrides[code]=unitPrice;markManualPriceOverride(item,code)}return item}
function openRecallDocument(type,documentNo,session=state.exhibitionSession){if(!state.recordsLoaded)return alert('請先匯入最新的 jmsdata.xlsx；正式文件記錄保存在第二個「Universe Records」工作表。');if(!state.stockAllRows.length)return alert('請先匯入最新的 jmsdata.xlsx，才可 Recall 原文件。');const found=findLatestDocument(type,documentNo,session);if(!found)return alert('找不到這張文件。');releaseCustomImages();state.recall={type,documentNo:norm(docValue(found.header,'DOCUMENT_NO')),session:found.session,baseRevision:found.revision,header:found.header,items:found.items};state.documentType=type;const radio=$(`input[name="documentType"][value="${type}"]`);if(radio)radio.checked=true;updateDocumentTypeUI();restoreDocumentFields(found.header);if(type==='quotation'){state.quote.historicalPm=state.quote.historicalPm||{};for(const row of found.items){const d=normalizeStockDate(docValue(row,'COMPLETION_DATE')),pm=numberValue(docValue(row,'QUOTE_HISTORICAL_PM'));if(d&&pm>0)state.quote.historicalPm[d]=pm}}state.deliveryReturns=new Set();state.items=found.items.map(recallItemFromRow);normalizeItemSequence();setRecallLock(true);const banner=$('#recallActive');if(banner){banner.classList.remove('hidden');banner.innerHTML=`<strong>正在修改 ${esc(state.recall.documentNo)}</strong><span>Revision ${state.recall.baseRevision} → ${state.recall.baseRevision+1}</span>`}updateFxPanel();syncEffectivePrices();renderItems();renderCustomerSummary();renderPreview();status('#addMessage',`${state.recall.documentNo} 已 Recall，可刪除／新增貨品後重新 Confirm。`,'ok');window.scrollTo({top:$('#invoice').offsetTop-70,behavior:'smooth'})}
function cancelRecall(){if(!state.recall)return;const no=state.recall.documentNo;if(!confirm(`取消修改 ${no}？尚未 Confirm 的變更不會保存。`))return;state.recall=null;setRecallLock(false);$('#recallActive')?.classList.add('hidden');clearCurrentDocument();setDefaultInvoiceNo(true);clearDocumentDraft();markDraftBaseline();renderRecallResults();status('#addMessage',`已取消 Recall ${no}；庫存沒有變更。`,'ok')}
function renderRecallResults(){
  const box=$('#recallResults');if(!box)return;const q=norm($('#recallSearchInput')?.value).toUpperCase(),typeFilter=$('#recallTypeFilter')?.value||'all',types=typeFilter==='all'?['invoice','consignment','quotation']:[typeFilter],rows=[];
  for(const type of types){for(const h of latestDocumentHeaders(type,{session:state.exhibitionSession})){const no=norm(docValue(h,'DOCUMENT_NO')),customer=norm(docValue(h,'CUSTOMER')),code=norm(docValue(h,'CUSTOMER_CODE')),date=norm(docValue(h,'DOCUMENT_DATE')),exhibition=norm(docValue(h,'EXHIBITION_NAME')),hay=`${no} ${customer} ${code} ${date} ${exhibition}`.toUpperCase();if(q&&!hay.includes(q))continue;const session=recordSession(h),found=findLatestDocument(type,no,session);rows.push({type,h,session,count:found?.items.length||0,exhibition:exhibition||session||'未命名展覽'})}}
  rows.sort((a,b)=>String(docValue(b.h,'UPDATED_AT')||docValue(b.h,'DOCUMENT_DATE')).localeCompare(String(docValue(a.h,'UPDATED_AT')||docValue(a.h,'DOCUMENT_DATE'))));if(!rows.length){box.innerHTML='<div class="notice">找不到已保存的 Invoice／Consignment／Quotation。</div>';return}
  box.innerHTML='';const groups=new Map();for(const row of rows.slice(0,200)){const key=row.exhibition||'未命名展覽';if(!groups.has(key))groups.set(key,[]);groups.get(key).push(row)}
  for(const [name,groupRows] of groups){const section=document.createElement('section');section.className='recall-exhibition-group';const head=document.createElement('div');head.className='recall-exhibition-head';head.innerHTML=`<strong>${esc(name)}</strong><span>${groupRows.length} 張文件</span>`;section.appendChild(head);for(const row of groupRows){const no=norm(docValue(row.h,'DOCUMENT_NO')),documentStatus=norm(docValue(row.h,'DOCUMENT_STATUS')).toUpperCase(),statusText=documentStatus==='CANCELLED'?'Cancelled':'Confirmed',card=document.createElement('div');card.className='recall-result';card.innerHTML=`<div><strong>${esc(no)}</strong><span>${row.type==='consignment'?'Consignment':row.type==='quotation'?'Quotation':'Invoice'} · Rev ${revisionNumber(row.h)} · ${esc(docValue(row.h,'DOCUMENT_DATE'))} · ${statusText}</span><small>${esc(docValue(row.h,'CUSTOMER_CODE'))} · ${esc(docValue(row.h,'CUSTOMER'))} · ${row.count} 件</small></div><button type="button">重新開啟</button>`;$('button',card).onclick=()=>openRecallDocument(row.type,no,row.session);section.appendChild(card)}box.appendChild(section)}
}
const INVENTORY_HISTORY_KEY='universeInventoryHistory_v1';
const LATEST_LONDON_PM_CACHE_KEY='universeKitcoLatestLondonPm_v1';
const KITCO_HISTORY_CACHE_KEY='universeKitcoLondonPmHistory_v1';
const GOLDSILVER_HISTORY_PAGE='https://goldsilver.com/price-charts/historical-london-fix/';
const K14_WEIGHT_FACTOR=.83,K14_WEIGHT_STEP=.05,TROY_OZ_G=31.1034768,GOLD_PRODUCTION_FACTOR=1.12,GOLD_PRICE_MULTIPLIER=4;
function loadInventoryHistory(){try{const raw=JSON.parse(localStorage.getItem(INVENTORY_HISTORY_KEY)||'[]');state.inventoryHistory=new Map((Array.isArray(raw)?raw:[]).filter(x=>x&&x.lotNo).map(x=>[String(x.lotNo),x]))}catch{state.inventoryHistory=new Map()}}
function saveInventoryHistory(){try{localStorage.setItem(INVENTORY_HISTORY_KEY,JSON.stringify([...state.inventoryHistory.values()]))}catch{}}
function productSnapshot(p){return{lotNo:p.lotNo,artNo:p.artNo,price:Number(p.price)||0,unit:p.unit||'PC',article:p.article||'',descriptions:[...(p.descriptions||[])],desc2:p.desc2||'',completionDate:p.completionDate||''}}
function historyStatusLabel(status){return status==='CONSIGNED'?'Consigned':status==='SOLD_ON_HAND'?'Sold - On Hand':status==='SOLD_DELIVERED'?'Sold - Delivered':'Available'}
function historyStatusClass(status){return status==='CONSIGNED'?'consigned':status==='SOLD_ON_HAND'?'sold-on-hand':status==='SOLD_DELIVERED'?'sold-delivered':'available'}
function recordInventoryMovement(item,statusValue,docNo=''){
  const lot=String(item.lotNo),prev=state.inventoryHistory.get(lot)||productSnapshot(item),docs=Array.isArray(prev.docs)?prev.docs:[];
  docs.push({type:state.documentType,no:docNo,date:norm($('#invoiceDate')?.value),customerCode:norm($('#customerCode')?.value),customer:norm($('#customerName')?.value),at:Date.now()});
  state.inventoryHistory.set(lot,{...prev,...productSnapshot(item),status:statusValue,docs,updatedAt:Date.now()});saveInventoryHistory();
}
function quoteMode(){return state.documentType==='quotation'}
function quote14KReferenceMode(){return quoteMode()&&state.quote.karat==='14K'}
function quote14KSameWeightMode(){return quoteMode()&&state.quote.karat==='14K_SAME_WEIGHT'}
function quote14KMode(){return quote14KReferenceMode()||quote14KSameWeightMode()}
function quoteModeName(){return state.quote.karat==='14K_SAME_WEIGHT'?'14K 同金重':state.quote.karat==='14K'?'14K 參考':'18K'}
function roundUpWeight05(v){const n=Number(v)||0;return Math.round(Math.ceil((n-1e-9)/K14_WEIGHT_STEP)*K14_WEIGHT_STEP*100)/100}
function parse18KDesc1(v){const raw=norm(v),m=raw.match(/^(\d+(?:\.\d+)?)\s*([A-Z]+(?:\/[A-Z]+)*)750(.*)$/i);return m?{raw,weight:Number(m[1]),metal:m[2].toUpperCase(),suffix:m[3]||''}:null}
function estimate14KWeightFromItem(item){const p=parse18KDesc1((item.descriptions||[])[0]);return p?roundUpWeight05(p.weight*K14_WEIGHT_FACTOR):0}
function quotationCurrentWeight(p){return quote14KReferenceMode()?roundUpWeight05(p.weight*K14_WEIGHT_FACTOR):p.weight}
function effectiveDescriptions(item){const d=[...(item.descriptions||[])];if(!quote14KMode()||!d.length)return d;const p=parse18KDesc1(d[0]);if(!p)return d;const w=quotationCurrentWeight(p);d[0]=`${w.toFixed(2)}${p.metal}585${p.suffix} (${p.raw})`;return d}
function normalizeStockDate(value){
  if(value instanceof Date&&!Number.isNaN(value.getTime()))return `${value.getFullYear()}-${String(value.getMonth()+1).padStart(2,'0')}-${String(value.getDate()).padStart(2,'0')}`;
  const raw=norm(value);if(!raw)return'';
  const serial=Number(raw);if(Number.isFinite(serial)&&serial>20000&&serial<90000){const d=new Date(Date.UTC(1899,11,30)+Math.round(serial)*86400000);return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`}
  let m=raw.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);if(m)return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  m=raw.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);if(m)return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  const d=new Date(raw);return Number.isNaN(d.getTime())?'':`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function goldNumber(value){const n=Number(String(value??'').replace(/[$,\s]/g,''));return Number.isFinite(n)&&n>0?n:0}
function goldFileLatest(){const dates=state.quote.goldDates||[],date=dates.length?dates[dates.length-1]:'';return date?{date,pm:Number(state.quote.goldPrices.get(date))||0}:null}
function findGoldPriceOnOrBefore(value){const target=normalizeStockDate(value),dates=state.quote.goldDates||[];if(!target||!dates.length)return null;let lo=0,hi=dates.length-1,best=-1;while(lo<=hi){const mid=(lo+hi)>>1;if(dates[mid]<=target){best=mid;lo=mid+1}else hi=mid-1}if(best<0)return null;const date=dates[best],pm=Number(state.quote.goldPrices.get(date))||0;return pm>0?{date,pm}:null}
async function importGoldSilverFile(file){
  const wb=await readWB(file),sheetName=wb.SheetNames.find(name=>name.trim().toUpperCase()==='GOLDSILVER.COM')||wb.SheetNames[0],rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:''});
  const headerIndex=rows.findIndex(row=>row.some(value=>norm(value).toUpperCase()==='DATE')&&row.some(value=>norm(value).toUpperCase()==='GOLD PM'));
  if(headerIndex<0)throw new Error('GoldSilver.xlsx 找不到 Date / Gold PM 欄位');
  const header=rows[headerIndex].map(value=>norm(value).toUpperCase()),dateCol=header.indexOf('DATE'),pmCol=header.indexOf('GOLD PM'),prices=new Map();
  for(const row of rows.slice(headerIndex+1)){const date=normalizeStockDate(row[dateCol]),pm=goldNumber(row[pmCol]);if(date&&pm>0)prices.set(date,pm)}
  if(!prices.size)throw new Error('GoldSilver.xlsx 沒有有效 Gold PM 資料');
  const dates=[...prices.keys()].sort();state.quote.goldPrices=prices;state.quote.goldDates=dates;state.quote.goldDataName=file.name||'GoldSilver.xlsx';state.quote.goldDataRows=prices.size;
  const latest=goldFileLatest();state.quote.currentLondonPmDate=latest.date;state.quote.currentLondonPm=latest.pm;state.quote.source='file';state.quote.historicalPm={};state.dataMeta.gold={name:file.name||'GoldSilver.xlsx',lastModified:Number(file.lastModified)||0,count:prices.size,latestDate:latest.date,latestPm:latest.pm};saveLatestLondonPmCache();saveHistoricalGoldCache();
  syncEffectivePrices({clearCurrentOverride:true});updateGoldQuoteUI();updateDataVersionPanel();runHealthCheck();renderItems();schedulePreview();
  return{count:prices.size,latestDate:latest.date,latestPm:latest.pm,sheetName};
}
function companyGoldPrice(londonPm){const n=Number(londonPm)||0;return n>0?Math.ceil((n*1.01)/10-1e-9)*10:0}
function goldValuePerGram(companyPrice,purity){return companyPrice>0?companyPrice/TROY_OZ_G*purity*GOLD_PRODUCTION_FACTOR:0}
function goldMetrics(londonPm=state.quote.currentLondonPm){const pm=Number(londonPm)||0;if(!(pm>0))return null;const base=companyGoldPrice(pm),g18=goldValuePerGram(base,.75),g14=goldValuePerGram(base,.585);return{pm,base,g18,g14}}
function historicalLondonPmDetails(item){
  const date=normalizeStockDate(item?.completionDate),manual=Number(state.quote.historicalPm?.[date]);
  if(date&&Number.isFinite(manual)&&manual>0)return{pm:manual,sourceDate:date,source:'manual'};
  if(date&&date===normalizeStockDate(state.quote.currentLondonPmDate)&&Number(state.quote.currentLondonPm)>0)return{pm:Number(state.quote.currentLondonPm),sourceDate:date,source:state.quote.source||'current'};
  const fileMatch=findGoldPriceOnOrBefore(date);if(fileMatch)return{pm:fileMatch.pm,sourceDate:fileMatch.date,source:'file'};
  return{pm:0,sourceDate:'',source:''};
}
function historicalLondonPm(item){return historicalLondonPmDetails(item).pm}
function quotationPriceDetails(item){
  const p=parse18KDesc1((item?.descriptions||[])[0]),date=normalizeStockDate(item?.completionDate),latestDate=normalizeStockDate(state.quote.currentLondonPmDate),current=goldMetrics(),historicalDetail=historicalLondonPmDetails(item),historicalPm=historicalDetail.pm,historical=goldMetrics(historicalPm),rate=Number($('#salesRate')?.value)||0;
  const missing=[];if(!p)missing.push('DESC1 金重');if(!date)missing.push('LDATE');if(!latestDate)missing.push('GoldSilver.xlsx 最新交易日期');if(!current)missing.push('最新 London PM 金價');if(date&&!historical)missing.push(`${date} 或之前的 London PM`);
  if(missing.length)return{ready:false,missing,date,latestDate,p,current,historical,historicalPm,historicalSourceDate:historicalDetail.sourceDate,historicalSource:historicalDetail.source};
  const currentWeight=quotationCurrentWeight(p),currentPerGram=quote14KMode()?current.g14:current.g18;
  const historicalTotal=p.weight*historical.g18,currentTotal=currentWeight*currentPerGram,adjustment=(currentTotal-historicalTotal)*GOLD_PRICE_MULTIPLIER*rate,rawBase=Math.max(0,(Number(item.price)||0)*rate),rawPrice=Math.max(0,rawBase+adjustment),finalPrice=Math.ceil(rawPrice-1e-9);
  return{ready:true,date,latestDate,p,current,historical,historicalPm,historicalSourceDate:historicalDetail.sourceDate,historicalSource:historicalDetail.source,currentWeight,currentPerGram,historicalTotal,currentTotal,adjustment,rawBase,rawPrice,finalPrice};
}
function quotationGoldReady(){return !quoteMode()||!!normalizeStockDate(state.quote.currentLondonPmDate)&&!!goldMetrics()&&state.items.every(item=>quotationPriceDetails(item).ready)}
function quote14KReady(){return quotationGoldReady()}
function effectiveUsdPrice(item){if(!quoteMode())return baseUsdPrice(item);const q=quotationPriceDetails(item);return q.ready?q.finalPrice:baseUsdPrice(item)}
function activePriceModeKey(){return quoteMode()?(quote14KSameWeightMode()?'quote14same':quote14KReferenceMode()?'quote14':'quote18'):'regular'}
function manualPriceFlagKey(code=currencyCode()){return `${activePriceModeKey()}:${code}`}
function markManualPriceOverride(item,code=currencyCode()){item.manualPriceFlags=item.manualPriceFlags||{};item.manualPriceFlags[manualPriceFlagKey(code)]=true}
function clearManualPriceOverride(item,code=currencyCode()){if(item?.manualPriceFlags)delete item.manualPriceFlags[manualPriceFlagKey(code)]}
function isManualPriceOverride(item,code=currencyCode()){return !!item?.manualPriceFlags?.[manualPriceFlagKey(code)]}
function activePriceOverrides(item){item.currencyPrices=item.currencyPrices||{};item.quote18kCurrencyPrices=item.quote18kCurrencyPrices||{};item.quote14kCurrencyPrices=item.quote14kCurrencyPrices||{};item.quote14kSameWeightCurrencyPrices=item.quote14kSameWeightCurrencyPrices||{};item.manualPriceFlags=item.manualPriceFlags||{};return quoteMode()?(quote14KSameWeightMode()?item.quote14kSameWeightCurrencyPrices:quote14KReferenceMode()?item.quote14kCurrencyPrices:item.quote18kCurrencyPrices):item.currencyPrices}
const DOCUMENT_DRAFT_KEY='universeDocumentDraft_v2';
function draftItemSnapshot(item){return{...productSnapshot(item),id:item.id,seq:Number(item.seq)||1,qty:Number(item.qty)||1,usdUnitPrice:Number(item.usdUnitPrice)||0,unitPrice:Number(item.unitPrice)||0,currencyPrices:{...(item.currencyPrices||{})},quote18kCurrencyPrices:{...(item.quote18kCurrencyPrices||{})},quote14kCurrencyPrices:{...(item.quote14kCurrencyPrices||{})},quote14kSameWeightCurrencyPrices:{...(item.quote14kSameWeightCurrencyPrices||{})},manualPriceFlags:{...(item.manualPriceFlags||{})},imageVariant:item.imageVariant||'',imageGrayscale:!!item.imageGrayscale,imageAutoMatched:!!item.imageAutoMatched,imageOverrideFile:item.imageOverrideFile||'',delivered:!!item.delivered}}
function captureDocumentDraft(){return{version:4,savedAt:new Date().toISOString(),packageName:state.packageName||'',exhibitionSession:state.exhibitionSession||'',exhibitionName:state.exhibitionName||'',documentType:state.documentType,fields:{invoiceNo:norm($('#invoiceNo')?.value),invoiceDate:norm($('#invoiceDate')?.value),customerCode:norm($('#customerCode')?.value),customerName:norm($('#customerName')?.value),customerAddress:norm($('#customerAddress')?.value),salesRate:norm($('#salesRate')?.value),currency:currencyCode(),fxRate:norm($('#fxRate')?.value),shipmentMethod:norm($('#shipmentMethod')?.value),customerTerms:norm($('#customerTerms')?.value),discountAmount:norm($('#discountAmount')?.value),remark:norm($('#remark')?.value)},quote:{karat:state.quote.karat,currentLondonPm:Number(state.quote.currentLondonPm)||0,currentLondonPmDate:normalizeStockDate(state.quote.currentLondonPmDate),historicalPm:{...(state.quote.historicalPm||{})},source:state.quote.source||''},items:formalItems().map(draftItemSnapshot),deliveryReturns:[...state.deliveryReturns],recall:state.recall?{type:state.recall.type,documentNo:state.recall.documentNo,session:state.recall.session||'',baseRevision:state.recall.baseRevision}:null}}
function draftComparable(data){const copy=structuredClone(data);delete copy.savedAt;delete copy.packageName;return JSON.stringify(copy)}
function updateDraftStatus(text='',kind=''){const el=$('#draftStatus');if(!el)return;const wrap=el.closest('.workspace-tools');if(!text){el.textContent='';el.className='draft-status';wrap?.classList.add('hidden');return}el.textContent=text;el.className='draft-status'+(kind?' '+kind:'');wrap?.classList.remove('hidden')}
function markDraftBaseline(){if(state.draft.timer){clearTimeout(state.draft.timer);state.draft.timer=null}state.draft.baseline=draftComparable(captureDocumentDraft())}
function clearDocumentDraft(){try{localStorage.removeItem(DOCUMENT_DRAFT_KEY)}catch{}updateDraftStatus()}
function saveDocumentDraft(){if(state.draft.restoring)return;const data=captureDocumentDraft();if(draftComparable(data)===state.draft.baseline){clearDocumentDraft();return}try{localStorage.setItem(DOCUMENT_DRAFT_KEY,JSON.stringify(data));updateDraftStatus(`草稿已自動儲存 · ${new Date().toLocaleTimeString('zh-HK',{hour:'2-digit',minute:'2-digit'})}`,'ok')}catch(err){console.warn('草稿無法儲存。',err);updateDraftStatus('草稿自動儲存失敗','warn')}}
function scheduleDraftSave(){if(state.draft.restoring)return;if(state.draft.timer)clearTimeout(state.draft.timer);state.draft.timer=setTimeout(()=>{state.draft.timer=null;saveDocumentDraft()},280)}
function loadDocumentDraft(){try{const x=JSON.parse(localStorage.getItem(DOCUMENT_DRAFT_KEY)||'null');return x&&x.version>=2&&x.fields?x:null}catch{return null}}
function restoreDocumentDraft(data){if(!data)return false;state.draft.restoring=true;try{releaseCustomImages();if(data.exhibitionSession!==undefined){state.exhibitionSession=norm(data.exhibitionSession);try{if(state.exhibitionSession)localStorage.setItem(EXHIBITION_SESSION_KEY,state.exhibitionSession);else localStorage.removeItem(EXHIBITION_SESSION_KEY)}catch{}}if(data.exhibitionName!==undefined)saveExhibitionName(data.exhibitionName);state.documentType=['invoice','consignment','quotation'].includes(data.documentType)?data.documentType:'invoice';const radio=$(`input[name="documentType"][value="${state.documentType}"]`);if(radio)radio.checked=true;updateDocumentTypeUI();const f=data.fields||{};$('#invoiceNo').value=f.invoiceNo||formatDocumentNo();$('#invoiceDate').value=f.invoiceDate||today();$('#customerCode').value=f.customerCode||'';$('#customerName').value=f.customerName||'';$('#customerAddress').value=f.customerAddress||'';$('#salesRate').value=f.salesRate||'';$('#currency').value=f.currency||'USD';$('#fxRate').value=f.fxRate||'';state.fx={rate:currencyCode()==='USD'?1:numberValue(f.fxRate),date:'',source:f.fxRate?'manual':'pending',fetching:false};$('#shipmentMethod').value=f.shipmentMethod||'';$('#customerTerms').value=f.customerTerms||'';$('#discountAmount').value=f.discountAmount||0;$('#remark').value=f.remark||'';if(data.quote){state.quote.karat=['18K','14K','14K_SAME_WEIGHT'].includes(data.quote.karat)?data.quote.karat:'18K';state.quote.currentLondonPm=numberValue(data.quote.currentLondonPm);state.quote.currentLondonPmDate=normalizeStockDate(data.quote.currentLondonPmDate);state.quote.historicalPm=data.quote.historicalPm&&typeof data.quote.historicalPm==='object'?data.quote.historicalPm:{};state.quote.source=data.quote.source||'draft'}$$('input[name="quoteKarat"]').forEach(x=>x.checked=x.value===state.quote.karat);state.items=(data.items||[]).map((saved,i)=>{const p=state.stockCatalog.get(String(saved.lotNo))||saved,item={...productSnapshot(p),id:saved.id||Date.now()+Math.random(),seq:Number(saved.seq)||i+1,qty:Number(saved.qty)||1,usdUnitPrice:Number(saved.usdUnitPrice)||0,unitPrice:Number(saved.unitPrice)||0,currencyPrices:{...(saved.currencyPrices||{})},quote18kCurrencyPrices:{...(saved.quote18kCurrencyPrices||{})},quote14kCurrencyPrices:{...(saved.quote14kCurrencyPrices||{})},quote14kSameWeightCurrencyPrices:{...(saved.quote14kSameWeightCurrencyPrices||{})},manualPriceFlags:{...(saved.manualPriceFlags||{})},imageVariant:saved.imageVariant||'',imageGrayscale:!!saved.imageGrayscale,imageAutoMatched:!!saved.imageAutoMatched,imageOverrideFile:saved.imageOverrideFile||'',delivered:state.documentType==='invoice'?(saved.delivered!==undefined?!!saved.delivered:initialDeliveredChoiceForLot(saved.lotNo||p.lotNo,'')):false};return item});state.deliveryReturns=new Set((data.deliveryReturns||[]).map(String));normalizeItemSequence();state.recall=null;if(data.recall){const found=findLatestDocument(data.recall.type,data.recall.documentNo,data.recall.session||'');if(found){state.recall={type:data.recall.type,documentNo:data.recall.documentNo,session:found.session,baseRevision:Number(data.recall.baseRevision)||found.revision,header:found.header,items:found.items}}}setRecallLock(!!state.recall);const banner=$('#recallActive');if(banner){banner.classList.toggle('hidden',!state.recall);if(state.recall)banner.innerHTML=`<strong>正在修改 ${esc(state.recall.documentNo)}</strong><span>Revision ${state.recall.baseRevision} → ${state.recall.baseRevision+1} · 已恢復草稿</span>`}updateFxPanel();syncEffectivePrices();renderItems();renderCustomerSummary();updateGoldQuoteUI();renderPreview();updateDraftStatus(`已恢復 ${documentLabels().short} 草稿 · ${state.items.length} 款`,'ok');return true}finally{state.draft.restoring=false}}
function maybeOfferDraftRestore(){const data=loadDocumentDraft();if(!data){updateDraftStatus();return false}if(!state.stockAllRows.length){updateDraftStatus(`發現未完成 ${documentLabels(data.documentType).short} 草稿；匯入展覽資料包後可恢復。`,'warn');return false}if(state.draft.prompted)return false;state.draft.prompted=true;const no=data.fields?.invoiceNo||'',count=(data.items||[]).length,when=data.savedAt?new Date(data.savedAt).toLocaleString('zh-HK'):'';if(confirm(`發現未完成 ${documentLabels(data.documentType).short} 草稿${no?` ${no}`:''}，共 ${count} 款${when?`\n最後儲存：${when}`:''}。\n\n是否恢復？`)){restoreDocumentDraft(data);return true}clearDocumentDraft();markDraftBaseline();return false}
function syncDraftAfterDataImport(){const data=loadDocumentDraft();if(!data){markDraftBaseline();updateDraftStatus();return false}const current=norm(state.exhibitionSession),draftSession=norm(data.exhibitionSession);if(!current||!draftSession||current!==draftSession){clearDocumentDraft();state.draft.prompted=false;markDraftBaseline();updateDraftStatus('已清除上一個展覽的未完成草稿；目前展覽由 jmsdata.xlsx 決定。','ok');return false}return maybeOfferDraftRestore()}

function saveLatestLondonPmCache(){try{const date=normalizeStockDate(state.quote.currentLondonPmDate),value=Number(state.quote.currentLondonPm)||0;if(date||value>0)localStorage.setItem(LATEST_LONDON_PM_CACHE_KEY,JSON.stringify({date,value,updatedAt:Date.now()}));else localStorage.removeItem(LATEST_LONDON_PM_CACHE_KEY)}catch{}}
function loadLatestLondonPmCache(){try{const x=JSON.parse(localStorage.getItem(LATEST_LONDON_PM_CACHE_KEY)||'null');if(x&&typeof x==='object'){state.quote.currentLondonPmDate=normalizeStockDate(x.date);state.quote.currentLondonPm=Number(x.value)>0?Number(x.value):0;state.quote.source='cache';return x}}catch{}return null}
function saveHistoricalGoldCache(){try{localStorage.setItem(KITCO_HISTORY_CACHE_KEY,JSON.stringify(state.quote.historicalPm||{}))}catch{}}
function loadHistoricalGoldCache(){try{const x=JSON.parse(localStorage.getItem(KITCO_HISTORY_CACHE_KEY)||'{}');state.quote.historicalPm=x&&typeof x==='object'&&!Array.isArray(x)?x:{}}catch{state.quote.historicalPm={}}}
function requiredHistoricalDates(){const map=new Map();for(const item of state.items){const date=normalizeStockDate(item.completionDate);if(date)map.set(date,(map.get(date)||0)+1)}return [...map.entries()].sort((a,b)=>a[0].localeCompare(b[0]))}
function renderHistoricalGoldInputs(){
  const box=$('#historicalGoldInputs');if(!box)return;
  if(!quoteMode()){box.innerHTML='';return}
  const dates=requiredHistoricalDates(),missingDate=state.items.filter(x=>!normalizeStockDate(x.completionDate)).length;
  if(!state.items.length){box.innerHTML='<div class="notice">加入 Quotation 貨品後，系統會按 Column R（LDATE）從 GoldSilver.xlsx 自動配對完成日或之前最近交易日的 London PM。</div>';return}
  box.innerHTML='';
  for(const [date,count] of dates){
    const detail=historicalLondonPmDetails({completionDate:date}),value=detail.pm,company=companyGoldPrice(value),fallback=detail.sourceDate&&detail.sourceDate!==date,sourceText=detail.source==='manual'?'手動覆蓋':detail.source==='file'?(fallback?`使用 ${detail.sourceDate} 最近交易日`:`GoldSilver.xlsx ${detail.sourceDate}`):detail.sourceDate?'共用最新 London PM':'尚未配對',row=document.createElement('div');
    row.className='historical-gold-row';row.innerHTML=`<div class="historical-gold-date"><strong>${esc(date)}</strong><span>${count} 款 · ${esc(sourceText)}</span></div><label>London PM${detail.sourceDate?` (${esc(detail.sourceDate)})`:''}<input type="number" min="0" step="0.01" inputmode="decimal" data-gold-date="${esc(date)}" value="${value>0?value:''}" placeholder="例如 4062.20"></label><div class="historical-company-price"><span>公司金價</span><strong>${company?`USD ${company.toFixed(0)} / oz`:'—'}</strong></div>`;
    const input=$('input',row);input.onchange=e=>setHistoricalLondonPm(date,e.target.value);box.appendChild(row);
  }
  if(missingDate){const warn=document.createElement('div');warn.className='notice error';warn.textContent=`有 ${missingDate} 款貨品沒有有效 LDATE，未能完成金價調整。`;box.appendChild(warn)}
}
function missingGoldSummary(){if(!normalizeStockDate(state.quote.currentLondonPmDate))return'請匯入包含有效 Gold PM 日期的 GoldSilver.xlsx。';if(!goldMetrics())return'請輸入最新 London PM 金價。';if(!state.items.length)return state.quote.goldDataRows?`已載入 ${state.quote.goldDataName}；加入 Quotation 貨品後會自動配對完成日金價。`:'加入 Quotation 貨品後，請輸入各完成日的 London PM。';const missing=[];for(const item of state.items){const d=quotationPriceDetails(item);if(!d.ready)for(const x of d.missing)if(!missing.includes(x))missing.push(x)}return missing.length?`尚欠：${missing.join('、')}`:'最新及完成日金價已齊全，Quotation 已重新計價。'}
function setGoldQuoteExpanded(expanded){
  const details=$('#goldQuoteDetails'),button=$('#toggleGoldQuoteBtn'),action=$('#goldQuoteToggleAction');if(!details||!button)return;
  details.classList.toggle('hidden',!expanded);button.setAttribute('aria-expanded',expanded?'true':'false');button.classList.toggle('expanded',expanded);if(action)action.textContent=expanded?'收起':'展開';
}
function toggleGoldQuoteDetails(){const details=$('#goldQuoteDetails');if(details)setGoldQuoteExpanded(details.classList.contains('hidden'))}
function updateGoldQuoteUI(message='',type=''){
  const panel=$('#quotationKaratPanel'),gold=$('#goldQuotePanel');if(!panel)return;
  panel.classList.toggle('hidden',!quoteMode());gold?.classList.toggle('hidden',!quoteMode());
  const pmInput=$('#currentLondonPmInput');if(pmInput&&document.activeElement!==pmInput)pmInput.value=state.quote.currentLondonPm>0?String(state.quote.currentLondonPm):'';
  const m=goldMetrics(),fileLatest=goldFileLatest(),meta=$('#currentLondonPmMeta'),toggleSummary=$('#goldQuoteToggleSummary');$('#companyGoldBase').textContent=m?`USD ${m.base.toFixed(0)} / oz`:'—';$('#gold18PerGram').textContent=m?`USD ${m.g18.toFixed(3)}`:'—';$('#gold14PerGram').textContent=m?`USD ${m.g14.toFixed(3)}`:'—';if(meta)meta.textContent=fileLatest?`${state.quote.goldDataName} 最新：${fileLatest.date} · USD ${fileLatest.pm.toFixed(2)} / oz`:'GoldSilver.xlsx 未載入，請先更新並匯入資料包。';if(toggleSummary){const date=normalizeStockDate(state.quote.currentLondonPmDate),pm=Number(state.quote.currentLondonPm)||0;toggleSummary.textContent=date&&pm>0?`最新 ${date} · USD ${pm.toFixed(2)} / oz`:fileLatest?`最新 ${fileLatest.date} · USD ${fileLatest.pm.toFixed(2)} / oz`:'按此展開並輸入金價資料'}renderHistoricalGoldInputs();
  const el=$('#goldQuoteStatus');if(el){const fresh=goldFreshnessInfo(),baseMessage=message||missingGoldSummary(),freshSuffix=fresh.status==='stale'?` · ${fresh.message}`:'';el.textContent=baseMessage+freshSuffix;el.className='fx-status'+(type?' '+type:quotationGoldReady()&&state.items.length&&fresh.status!=='stale'?' ok':' warn')}updateGoldFreshnessUI()
}
function latestLondonPmStatus(){const date=normalizeStockDate(state.quote.currentLondonPmDate),n=Number(state.quote.currentLondonPm)||0;return date&&n>0?`最新可用 London PM：${date} · USD ${n.toFixed(2)} / oz`:date?'請輸入最新可用 London PM 金價。':'請匯入包含有效 Gold PM 日期的 GoldSilver.xlsx。'}
function setCurrentLondonPmValue(value){const n=Number(value);state.quote.currentLondonPm=Number.isFinite(n)&&n>0?n:0;state.quote.source='manual';saveLatestLondonPmCache();syncEffectivePrices({clearCurrentOverride:true});renderItems();updateGoldQuoteUI(latestLondonPmStatus(),state.quote.currentLondonPm>0?'ok':'warn');schedulePreview();return state.quote.currentLondonPm>0}
function setHistoricalLondonPm(date,value){const key=normalizeStockDate(date),n=Number(value);if(!key)return;if(Number.isFinite(n)&&n>0)state.quote.historicalPm[key]=n;else delete state.quote.historicalPm[key];saveHistoricalGoldCache();syncEffectivePrices({clearCurrentOverride:true});renderItems();updateGoldQuoteUI(n>0?`${key} London PM 已設為 USD ${n.toFixed(2)} / oz。`:`已清除 ${key} London PM。`,n>0?'ok':'warn');schedulePreview()}
function openGoldSilverLondonHistory(){window.open(GOLDSILVER_HISTORY_PAGE,'_blank','noopener,noreferrer');updateGoldQuoteUI('已開啟 GoldSilver Historical London Fix 頁面。資料包內 GoldSilver.xlsx 仍是目前自動配對來源。','warn')}
function setQuoteKarat(value){state.quote.karat=['18K','14K','14K_SAME_WEIGHT'].includes(value)?value:'18K';$$('input[name="quoteKarat"]').forEach(x=>x.checked=x.value===state.quote.karat);syncEffectivePrices({clearCurrentOverride:true});updateGoldQuoteUI();renderItems();schedulePreview()}
loadInventoryHistory();loadDocumentStore();loadImageOverridesLocal();loadLatestLondonPmCache();loadHistoricalGoldCache();if($('#exhibitionName'))$('#exhibitionName').value=state.exhibitionName||'';saveExhibitionName(state.exhibitionName);
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
    if(clearCurrentOverride){delete overrides[code];clearManualPriceOverride(item,code)}
    const manual=Number(overrides[code]);
    item.unitPrice=Number.isFinite(manual)&&manual>=0?manual:convertedFromUsd(effectiveUsdPrice(item));
  }
  scheduleDraftSave();
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
function updateTotals(){const t=totals();$('#totalQty').textContent=t.qty;$('#subtotal').textContent=fmt(t.sub);const discountEl=$('#discountDisplay');if(discountEl)discountEl.textContent=discountDisplay(t.discount);$('#grandTotal').textContent=fmt(t.total);$('#productCount').textContent=state.products.size;$('#customerCount').textContent=state.customers.size;$('#invoiceCount').textContent=state.items.length;$('#headerTotal').textContent=fmt(t.total);schedulePreview();scheduleDraftSave()}
function activatePanel(panelId,tabButton=null){
  const leaving=activePanelId();if(leaving==='stockSearch')stockSearchPerf.scrollY=window.scrollY;
  $$('.tab').forEach(x=>x.classList.remove('active'));$$('.tab-panel').forEach(x=>x.classList.remove('active'));
  const panel=$('#'+panelId);if(panel)panel.classList.add('active');if(tabButton)tabButton.classList.add('active');
  const importBtn=$('#dataImportBtn');if(importBtn)importBtn.classList.toggle('active',panelId==='setup');
  if(panelId==='invoice'){renderCustomerSummary();renderRecallResults()}if(panelId==='preview')renderPreview();if(panelId==='stockSearch'){ensureStockSearchRendered();requestAnimationFrame(()=>{if(stockSearchPerf.scrollY>0)window.scrollTo({top:stockSearchPerf.scrollY,behavior:'auto'})})}requestAnimationFrame(updateBackToTopButton)
}
$$('.tab').forEach(b=>b.onclick=()=>activatePanel(b.dataset.tab,b));
$('#dataImportBtn').onclick=()=>activatePanel('setup');
async function readWB(file){if(typeof XLSX==='undefined')throw new Error('Excel 程式未載入');return XLSX.read(await file.arrayBuffer(),{type:'array',cellStyles:true,cellDates:true})}
async function importStockFile(f,{exhibitionNameHint=''}={}){
  if(!/\.xlsx$/i.test(f?.name||''))throw new Error('庫存主檔已統一為 jmsdata.xlsx；請選擇 .xlsx 檔案。');
  const wb=await readWB(f),sheetName=wb.SheetNames[0],ws=wb.Sheets[sheetName],rows=XLSX.utils.sheet_to_json(ws,{defval:''});
  state.stockWorkbook=wb;state.stockSheetName=sheetName||'jmsdata';state.stockFileName=f.name||'jmsdata.xlsx';state.stockHeaders=Object.keys(rows[0]||{});state.stockAllRows=rows;state.stockRowByLot=new Map();state.stockIntegrityIssues=[];state.stockDuplicateLots=[];
  const fullMap=new Map(),seenLots=new Map();
  for(const r of rows){const lot=norm(field(r,['LOTNO'])),art=normArt(field(r,['ARTNO'])),price=Number(field(r,['PRICE']));if(!lot||!art||!Number.isFinite(price))continue;if(seenLots.has(lot))state.stockDuplicateLots.push({lotNo:lot,first:seenLots.get(lot),duplicate:r});else seenLots.set(lot,r);const desc=[];for(let i=1;i<=6;i++){const v=norm(field(r,[`DESC${i}`]));if(v)desc.push(v)}const p={lotNo:lot,artNo:art,price,unit:norm(field(r,['UNIT']))||'PC',article:norm(field(r,['ARTICLE']))||'',descriptions:desc,desc2:norm(field(r,['DESC2'])),completionDate:normalizeStockDate(field(r,['LDATE']))};fullMap.set(lot,p);state.stockRowByLot.set(lot,r);let vector=inventoryVectorFromRow(r),sum=vector.i+vector.j+vector.k+vector.l;if(sum!==vector.m||!['AVAILABLE','CONSIGNED','SOLD_ON_HAND','SOLD_DELIVERED'].includes(inventoryStatusFromVector(vector)))state.stockIntegrityIssues.push({lotNo:lot,vector});}
  if(!fullMap.size)throw new Error('找不到有效 LOTNO / ARTNO / PRICE');state.stockCatalog=fullMap;state.importConflicts=[];resetFormalRecordContext();const imageInfo=importImageOverridesFromWorkbook(wb);const recordsInfo=importUniverseRecordsFromWorkbook(wb,{fallbackName:norm(exhibitionNameHint),fileName:f.name||'jmsdata.xlsx',lastModified:f.lastModified});rebuildInventoryMaps();state.dataMeta.stock={name:f.name||'',lastModified:Number(f.lastModified)||0,count:fullMap.size,available:state.products.size,duplicates:state.stockDuplicateLots.length,integrity:state.stockIntegrityIssues.length,documentConflicts:0,imageOverrideConflicts:imageInfo?.conflicts||0,recordsEmbedded:!!state.dataMeta.records?.embedded};const exportBtn=$('#exportJmsdataBtn');if(exportBtn)exportBtn.disabled=false;renderRecallResults();updateDataVersionPanel();runHealthCheck();return fullMap.size;
}
async function importCustomerFile(f){
  const wb=await readWB(f),ws=wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});const map=new Map();let defaultRateCount=0;
  for(const r of rows){const code=normCode(r[0]),company=norm(r[1]);if(!code||!company||code.includes('CUSTOMER'))continue;const raw=r[11],num=Number(raw),usesDefault=(raw===''||!Number.isFinite(num)),rate=usesDefault?0.34:num;if(usesDefault)defaultRateCount++;map.set(code,{code,company,address:[r[2],r[3],r[4]].map(norm).filter(Boolean).join('\n'),rate,terms:norm(r[10]),usesDefaultRate:usesDefault})}
  if(!map.size)throw new Error('找不到有效客戶資料');state.customers=map;state.dataMeta.customer={name:f.name||'',lastModified:Number(f.lastModified)||0,count:map.size,defaultRateCount};updateDataVersionPanel();runHealthCheck();return map.size;
}
async function importStoneFile(f){
  const wb=await readWB(f);const sheetName=wb.SheetNames.find(n=>n.trim().toUpperCase()==='STONE LIST')||wb.SheetNames[0];const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:''});
  const headerIndex=rows.findIndex(r=>r.some(v=>String(v).trim().toUpperCase()==='BREAKDOWN')&&r.some(v=>String(v).trim().toUpperCase()==='QUOTATION'));
  if(headerIndex<0)throw new Error('找不到 Stone List 標題列');
  const header=rows[headerIndex].map(v=>String(v).trim().toUpperCase()),typeCol=header.indexOf('石類'),englishCol=header.indexOf('英文石名'),bCol=header.indexOf('BREAKDOWN'),qCol=header.indexOf('QUOTATION'),gCol=header.indexOf('GROUP');
  if(typeCol<0||englishCol<0||bCol<0||qCol<0||gCol<0)throw new Error('Stone List 必須包含 石類 / 英文石名 / BREAKDOWN / QUOTATION / GROUP 欄位');
  const aliases=new Map(),variantAliases=new Map(),groups=new Map(),englishNames=new Map(),diamondCodes=new Set(),seen=new Map(),diag={duplicates:[],multiAlias:[],missingGroup:[],missingType:[],missingEnglish:[],missingQuotation:[],prefixOverlaps:[]};
  for(let ri=headerIndex+1;ri<rows.length;ri++){
    const r=rows[ri],stoneType=norm(r[typeCol]).replace(/\s+/g,''),englishName=norm(r[englishCol]),breakdown=norm(r[bCol]),quotation=norm(r[qCol]),group=norm(r[gCol]);if(!breakdown)continue;
    const codes=breakdown.split(/[,，]/).map(norm).filter(Boolean).map(x=>x.toUpperCase()),quotes=quotation.split(/[,，]/).map(norm).filter(Boolean).map(x=>x.toUpperCase()),isDiamond=stoneType.includes('鑽');
    for(const code of codes){const prev=seen.get(code);if(prev){if(prev.stoneType!==stoneType||prev.group!==group)diag.duplicates.push({code,row:ri+1,firstRow:prev.row,firstType:prev.stoneType,type:stoneType,firstGroup:prev.group,group});else if(!diag.multiAlias.some(x=>x.code===code))diag.multiAlias.push({code,firstRow:prev.row,row:ri+1})}else seen.set(code,{row:ri+1,stoneType,group});if(!stoneType)diag.missingType.push({code,row:ri+1});if(!englishName)diag.missingEnglish.push({code,row:ri+1});if(!group)diag.missingGroup.push({code,row:ri+1});if(!quotation)diag.missingQuotation.push({code,row:ri+1});if(englishName&&!englishNames.has(code))englishNames.set(code,englishName);if(!aliases.has(code))aliases.set(code,code);if(quotes.length){const list=variantAliases.get(code)||[];for(const q of quotes)if(!list.includes(q))list.push(q);variantAliases.set(code,list);aliases.set(code,quotes[quotes.length-1])}if(group)groups.set(code,group);if(isDiamond)diamondCodes.add(code)}
  }
  if(!aliases.size)throw new Error('Stone List 沒有有效 BREAKDOWN 資料');const codes=[...aliases.keys()].sort((a,b)=>a.length-b.length||a.localeCompare(b));for(let i=0;i<codes.length;i++)for(let j=i+1;j<codes.length;j++){const a=codes[i],b=codes[j];if(b.startsWith(a)&&a!==b)diag.prefixOverlaps.push({short:a,long:b})}
  state.stoneAliases=aliases;state.stoneVariantAliases=variantAliases;state.stoneGroups=groups;state.stoneEnglishNames=englishNames;state.diamondStoneCodes=diamondCodes;state.stoneMappingName=f.name;state.stoneDiagnostics=diag;state.dataMeta.stone={name:f.name||'',lastModified:Number(f.lastModified)||0,count:aliases.size,diamondCount:diamondCodes.size,...Object.fromEntries(Object.entries(diag).map(([k,v])=>[k,v.length]))};resetStockStoneDerivedCache();renderStockSearch();updateDataVersionPanel();runHealthCheck();return aliases.size;
}
async function importArticleFile(f){const wb=await readWB(f),ws=wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});const map=new Map();for(const row of rows){const prefix=normCode(row[0]),description=norm(row[1]);if(!prefix||!description||prefix==='PREFIX')continue;map.set(prefix,description)}if(!map.size)throw new Error('找不到 Prefix / Article Description 對照');state.articleMap=map;state.articleMappingName=f.name;state.dataMeta.article={name:f.name||'',lastModified:Number(f.lastModified)||0,count:map.size};updateDataVersionPanel();return map.size}
async function importTemplateFile(f){const buf=await f.arrayBuffer();if(typeof ExcelJS==='undefined')throw new Error('Excel 範本程式未載入');const test=new ExcelJS.Workbook();await test.xlsx.load(buf.slice(0));if(!test.worksheets.length)throw new Error('範本沒有工作表');state.invoiceTemplateBuffer=buf;state.invoiceTemplateName=f.name;state.dataMeta.template={name:f.name||'',lastModified:Number(f.lastModified)||0};updateDataVersionPanel();return true}
function imageDuplicatePreference(p){const nl=articleType(p?.art)==='NL';if(nl)return p?.dup===1?0:p?.dup===0?1:2+(Number(p?.dup)||0);return p?.dup===0?0:1+(Number(p?.dup)||0)}
function releaseImageObjectUrl(image){if(image?.url&&String(image.url).startsWith('blob:')){try{URL.revokeObjectURL(image.url)}catch{}}if(image)image.url=''}
function releaseAllIndexedImageUrls(){for(const arr of state.imageFiles.values())for(const image of arr)releaseImageObjectUrl(image)}
function ensureImageObjectUrl(image){if(!image?.file)return'';if(!image.url)image.url=URL.createObjectURL(image.file);image.lastUsed=Date.now();return image.url}
let lazyImageObserver=null;
function getLazyImageObserver(){if(lazyImageObserver||typeof IntersectionObserver==='undefined')return lazyImageObserver;if(typeof IntersectionObserver==='undefined')return null;lazyImageObserver=new IntersectionObserver(entries=>{for(const entry of entries){if(!entry.isIntersecting)continue;const img=entry.target,image=img._imageRef;if(image){img.src=ensureImageObjectUrl(image);img._imageLoadedFile=image.fileName||''}lazyImageObserver.unobserve(img)}},{rootMargin:'500px 0px'});return lazyImageObserver}
function setLazyImageElement(img,image,fallbackSrc){if(!img)return;img.loading='lazy';img.decoding='async';img._imageRef=image||null;img.src=fallbackSrc||placeholder('No Image');const observer=getLazyImageObserver();if(image){if(observer)observer.observe(img);else img.src=ensureImageObjectUrl(image)}}
function releaseEditorCandidateUrls(product){if(!product)return;const selected=chooseImageMatch(product),protectedNames=new Set([selected?.fileName].filter(Boolean).map(x=>String(x).toUpperCase()));for(const item of state.items){const im=getImg(item);if(im?.fileName)protectedNames.add(String(im.fileName).toUpperCase())}for(const image of state.imageFiles.get(product.artNo)||[]){if(image.manualSource)continue;if(protectedNames.has(String(image.fileName||'').toUpperCase()))continue;releaseImageObjectUrl(image)}}
async function importImageFiles(files){
  const all=[...files].filter(f=>String(f.type).startsWith('image/')||/\.(jpe?g|png|webp)$/i.test(f.name)),map=new Map(),arts=[...new Set([...state.stockCatalog.values(),...state.products.values(),...state.inventoryHistory.values()].map(x=>normArt(x.artNo)).filter(Boolean))].sort((a,b)=>b.length-a.length);releaseAllIndexedImageUrls();state.imageIndexProgress={done:0,total:all.length};
  const progress=$('#imageIndexProgress');if(progress){progress.classList.remove('hidden');progress.textContent=all.length?`正在建立圖片索引：0 / ${all.length}`:'沒有圖片可建立索引。'}
  for(let i=0;i<all.length;i++){const f=all[i],p=parseImage(f,arts);if(p){const key=p.art+'|'+p.variant.toUpperCase(),existing=map.get(key);if(!existing||imageDuplicatePreference(p)<imageDuplicatePreference(existing))map.set(key,p)}if((i+1)%250===0||i===all.length-1){state.imageIndexProgress.done=i+1;if(progress)progress.textContent=`正在建立圖片索引：${i+1} / ${all.length}`;await new Promise(resolve=>setTimeout(resolve,0))}}
  state.imageFiles=new Map();state.imageFilesByName=new Map();for(const p of map.values()){const image={variant:p.variant,url:'',fileName:p.file.name,file:p.file,dup:p.dup||0,lastUsed:0};const arr=state.imageFiles.get(p.art)||[];arr.push(image);state.imageFiles.set(p.art,arr);state.imageFilesByName.set(String(p.file.name||'').toUpperCase(),image)}for(const arr of state.imageFiles.values())arr.sort((a,b)=>a.variant==='Default'?-1:b.variant==='Default'?1:a.variant.localeCompare(b.variant));for(const item of state.items)if(!item.customImage)applyAutoImageMatch(item);state.dataMeta.images={count:all.length,matched:state.imageFiles.size,indexed:map.size};invalidateStockSearchImages();if(progress){progress.textContent=`圖片索引完成：${all.length} 張 · ${state.imageFiles.size} 個款號`;progress.className='notice ok image-index-progress'}updateImageOverrideStatus();updateDataVersionPanel();runHealthCheck();return{images:all.length,matched:state.imageFiles.size}
}
$('#stockInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;if(!confirmDiscardPendingImageOverrides()){e.target.value='';return}try{const count=await importStockFile(f,{exhibitionNameHint:''}),issues=state.stockIntegrityIssues.length,dups=state.stockDuplicateLots.length,conflicts=state.importConflicts?.length||0;status('#stockStatus',`已匯入 ${f.name}：${count} 件貨品；Available ${state.products.size} 件；Records ${state.recordsLoaded?officialDocumentCount():'未載入'}；圖片選擇 ${state.imageOverrides.size} 項${dups?`；重複 LOTNO ${dups}`:''}${issues?`；${issues} 列 I:M 需要檢查`:''}${conflicts?`；文件記錄衝突 ${conflicts}`:''}。`,dups||issues||conflicts?'warn':'ok');setImportCollapsed('stock',true);updateTotals();syncDraftAfterDataImport()}catch(err){status('#stockStatus','匯入失敗：'+err.message,'error');setImportCollapsed('stock',false)}};
$('#customerInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const count=await importCustomerFile(f);status('#customerStatus',`已匯入 ${f.name}：${count} 位客戶。`,'ok');setImportCollapsed('customer',true);updateTotals()}catch(err){status('#customerStatus','匯入失敗：'+err.message,'error');setImportCollapsed('customer',false)}};
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

function activeStoneAliases(){return state.stoneAliases}
function activeStoneVariantAliases(){return state.stoneVariantAliases}
function activeStoneGroups(){return state.stoneGroups}
function activeStoneEnglishNames(){return state.stoneEnglishNames}
function stoneEnglishNameForCode(code){return norm(activeStoneEnglishNames().get(String(code||'').toUpperCase())).toUpperCase()}
function isDiamondStoneCode(code){return state.diamondStoneCodes.has(norm(code).toUpperCase())}
function stoneImageAliasesForCode(code){const c=String(code||'').toUpperCase(),out=[];if(c)out.push(c);const list=activeStoneVariantAliases().get(c);if(Array.isArray(list))for(const value of list){const v=String(value||'').toUpperCase();if(v&&!out.includes(v))out.push(v)}const single=activeStoneAliases().get(c),v=String(single||'').toUpperCase();if(v&&!out.includes(v))out.push(v);return out}
function stoneGroupForCode(code){return norm(activeStoneGroups().get(String(code||'').toUpperCase())).toUpperCase()}

function fileMetaDate(ts){if(!Number(ts))return'';try{return new Date(Number(ts)).toLocaleString('zh-HK',{year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit'})}catch{return''}}
function updateDataVersionPanel(){
  const el=$('#dataVersionPanel');if(!el)return;const m=state.dataMeta||{},rows=[];
  rows.push(['PWA',`v${APP_VERSION}`]);
  rows.push(['Exhibition',state.exhibitionName||'尚未命名']);
  rows.push(['Records',m.records?`jmsdata.xlsx Sheet 2 · ${m.records.count} 筆文件 · ${m.records.exhibitionId||'未建立 ID'}`:'未載入']);
  rows.push(['jmsdata',m.stock?`${m.stock.name} · ${m.stock.count} 件 · Available ${m.stock.available}${m.stock.lastModified?` · ${fileMetaDate(m.stock.lastModified)}`:''}`:'未載入']);
  rows.push(['Stone List',m.stone?`${m.stone.name} · ${m.stone.count} 個 BREAKDOWN · 鑽石代號 ${m.stone.diamondCount}`:'未載入']);
  rows.push(['GoldSilver',m.gold?`${m.gold.name} · 最新 ${m.gold.latestDate} · USD ${Number(m.gold.latestPm||0).toFixed(2)}`:'未載入']);
  rows.push(['Customer',m.customer?`${m.customer.name} · ${m.customer.count} 位${m.customer.defaultRateCount?` · ${m.customer.defaultRateCount} 位使用預設 0.34`:''}`:'未載入']);
  rows.push(['Pictures',m.images?`${m.images.count} 張 · ${m.images.matched} 個款號`:'未載入']);
  rows.push(['Template',m.template?m.template.name:'未載入']);
  if(m.article)rows.push(['Article Mapping',`${m.article.name} · ${m.article.count} 個`]);
  el.innerHTML=rows.map(([k,v])=>`<div><span>${esc(k)}</span><strong>${esc(v)}</strong></div>`).join('');
  const statusEl=$('#currentExhibitionStatus');if(statusEl)statusEl.textContent=state.exhibitionName?`目前展覽：${state.exhibitionName}`:'目前展覽：尚未命名';
}
function dateOnlyLocal(iso){const m=String(iso||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)return null;return new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),12,0,0,0)}
function businessDaysSince(iso){const start=dateOnlyLocal(iso);if(!start)return null;const end=dateOnlyLocal(today());if(!end)return null;if(start>end)return-1;let n=0,d=new Date(start);d.setDate(d.getDate()+1);while(d<=end){const day=d.getDay();if(day!==0&&day!==6)n++;d.setDate(d.getDate()+1)}return n}
function goldFreshnessInfo(){const latest=goldFileLatest();if(!latest)return{status:'missing',message:'GoldSilver.xlsx 未載入。',businessDays:null};const days=businessDaysSince(latest.date);if(days===null)return{status:'warn',message:`最新 London PM：${latest.date} · 無法判斷資料新鮮度。`,businessDays:null};if(days<0)return{status:'warn',message:`最新 London PM 日期 ${latest.date} 晚於今天，請檢查 GoldSilver.xlsx。`,businessDays:days};if(days>=2)return{status:'stale',message:`⚠ 最新 London PM：${latest.date} · USD ${latest.pm.toFixed(2)} / oz · 已相隔 ${days} 個工作日，請確認 GoldSilver.xlsx 是否需要更新。`,businessDays:days};return{status:'ok',message:`✓ 最新 London PM：${latest.date} · USD ${latest.pm.toFixed(2)} / oz · 日期正常。`,businessDays:days}}
function updateGoldFreshnessUI(){const info=goldFreshnessInfo(),el=$('#goldFreshnessStatus');if(el){el.textContent=info.message;el.className='notice '+(info.status==='ok'?'ok':info.status==='missing'?'':'warn')}return info}
function stonePartsFromDescriptionLine(line){const raw=String(line||'').toUpperCase(),firstDash=raw.indexOf('-');if(firstDash<0)return[];const tail=raw.slice(firstDash+1),nextDash=tail.indexOf('-'),stoneBlock=(nextDash>=0?tail.slice(0,nextDash):tail).replace(/\s+/g,'');if(!stoneBlock)return[];return stoneBlock.split(/[+\/，,]/).map(x=>x.trim()).filter(Boolean).map(part=>({part,code:longestStoneBreakdownPrefix(part)}))}
function allStoneCodesForProduct(p){const out=[];for(const line of p?.descriptions||[]){for(const x of stonePartsFromDescriptionLine(line)){if(x.code&&!out.includes(x.code))out.push(x.code)}}return out}
function documentStoneDescriptionEntries(items=formalItems()){const out=[],seen=new Set();for(const item of items||[]){for(const line of item?.descriptions||[]){for(const x of stonePartsFromDescriptionLine(line)){const code=norm(x.code).toUpperCase();if(!code||seen.has(code))continue;const name=stoneEnglishNameForCode(code);if(!name)continue;seen.add(code);out.push({code,name,text:`${code} = ${name}`})}}}return out.sort((a,b)=>a.code.localeCompare(b.code,'en',{sensitivity:'base'}))}
function documentStoneDescriptionPairs(items=formalItems()){const entries=documentStoneDescriptionEntries(items),rows=Math.ceil(entries.length/2),pairs=[];for(let i=0;i<rows;i++)pairs.push([entries[i]||null,entries[i+rows]||null]);return pairs}
function invoiceWeightSummary(items=formalItems()){let goldGrams=0,allStoneCarats=0,diamondCarats=0,semiPreciousCarats=0;for(const item of items||[]){const qty=Math.max(0,Number(item?.qty)||0),desc=item?.descriptions||[],goldMatch=norm(desc[0]).match(/^(\d+(?:\.\d+)?)/);if(goldMatch)goldGrams+=Number(goldMatch[1])*qty;for(const line of desc.slice(1)){let lineCarats=0;for(const m of String(line||'').matchAll(/(\d+(?:\.\d+)?)\s*ct\b/ig))lineCarats+=Number(m[1]);if(!lineCarats)continue;lineCarats*=qty;allStoneCarats+=lineCarats;const codes=stonePartsFromDescriptionLine(line).map(x=>norm(x.code).toUpperCase()).filter(Boolean);if(codes.length&&codes.every(isDiamondStoneCode))diamondCarats+=lineCarats;else semiPreciousCarats+=lineCarats}}return{goldGrams,allStoneCarats,allStoneGrams:allStoneCarats*.2,diamondCarats,diamondGrams:diamondCarats*.2,semiPreciousCarats,semiPreciousGrams:semiPreciousCarats*.2}}
function addCalendarMonthsIso(value,months){const m=String(value||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)return'';const y=Number(m[1]),month=Number(m[2])-1,day=Number(m[3]),index=month+Number(months||0),ty=y+Math.floor(index/12),tm=((index%12)+12)%12,last=new Date(ty,tm+1,0).getDate(),td=Math.min(day,last);return `${ty}-${String(tm+1).padStart(2,'0')}-${String(td).padStart(2,'0')}`}
function invoicePaymentTermLines(total=totals().total,invoiceDate=norm($('#invoiceDate')?.value)||today()){const amount=Number(total)||0,installment=Math.ceil(amount*.2),amounts=[installment,installment,installment,installment,Math.round((amount-installment*4)*100)/100],months=[3,4,5,6],days=[90,120,150,180],lines=[`PAID BY BANK TT 20% - ${fmt(amounts[0])}`];for(let i=0;i<4;i++)lines.push(`TERMS ${days[i]} DAYS 20% (${addCalendarMonthsIso(invoiceDate,months[i])}) - ${fmt(amounts[i+1])} WIRE BANK TRANSFER`);return lines}
function normalizeExcelAddonOptions(v={}){return{paymentTerm:!!v.paymentTerm,remark:!!v.remark,goldWeight:!!v.goldWeight,semiPrecious:!!v.semiPrecious,diamond:!!v.diamond,allStone:!!v.allStone,grossWeight:!!v.grossWeight,declaration:!!v.declaration,stoneDescription:!!v.stoneDescription}}
function excelAddonRequested(v={}){const o=normalizeExcelAddonOptions(v);return o.paymentTerm||o.remark||o.declaration||o.stoneDescription}
function excelAddonNeedsStoneList(v={}){const o=normalizeExcelAddonOptions(v);return o.stoneDescription||(o.remark&&(o.semiPrecious||o.diamond))}
function excelAddonOptionsFromUI(){const goldWeight=!!$('#excelOptGold')?.checked,semiPrecious=!!$('#excelOptSemi')?.checked,diamond=!!$('#excelOptDiamond')?.checked,allStone=!!$('#excelOptAllStone')?.checked,grossWeight=!!$('#excelOptGross')?.checked;return normalizeExcelAddonOptions({paymentTerm:$('#excelOptPayment')?.checked,remark:goldWeight||semiPrecious||diamond||allStone||grossWeight,goldWeight,semiPrecious,diamond,allStone,grossWeight,declaration:$('#excelOptDeclaration')?.checked,stoneDescription:$('#excelOptStoneDescription')?.checked})}
function syncPreviewAddonControls(){
  const panel=$('#previewAddonOptions');if(panel)panel.classList.toggle('hidden',state.documentType!=='invoice');
  const hint=$('#previewAddonHint');if(hint){const o=excelAddonOptionsFromUI(),warnings=[];if(excelAddonRequested(o)&&!state.invoiceTemplateBuffer)warnings.push('附加資料正式匯出需要 Invoice Master Template');if(excelAddonNeedsStoneList(o)&&!state.dataMeta?.stone)warnings.push('所選項目需要最新 Stone List');hint.textContent=warnings.join(' · ');hint.classList.toggle('warn',warnings.length>0)}
}
function resetPreviewAddonOptions(){for(const id of ['excelOptPayment','excelOptGold','excelOptSemi','excelOptDiamond','excelOptAllStone','excelOptGross','excelOptDeclaration','excelOptStoneDescription']){const el=$('#'+id);if(el)el.checked=false}syncPreviewAddonControls();schedulePreview()}
function setupPreviewAddonControls(){
  const ids=['excelOptPayment','excelOptGold','excelOptSemi','excelOptDiamond','excelOptAllStone','excelOptGross','excelOptDeclaration','excelOptStoneDescription'];
  for(const id of ids){const el=$('#'+id);if(!el)continue;el.addEventListener('change',()=>{syncPreviewAddonControls();schedulePreview()})}
  syncPreviewAddonControls();
}
function unknownStoneDiagnostics(){const out=[];if(!state.stockAllRows.length||!state.stoneAliases.size)return out;for(const row of state.stockAllRows){const lot=norm(field(row,['LOTNO'])),art=normArt(field(row,['ARTNO']));for(let i=2;i<=6;i++){const line=norm(field(row,[`DESC${i}`]));if(!line)continue;for(const x of stonePartsFromDescriptionLine(line)){if(!x.code)out.push({lotNo:lot,artNo:art,field:`DESC${i}`,line,token:x.part})}}}return out}
function imageCoverageDiagnostics(){if(!state.stockCatalog.size)return{missingProducts:0,missingArticles:[]};const missing=[];for(const p of state.stockCatalog.values())if(!(state.imageFiles.get(p.artNo)||[]).length)missing.push(p);return{missingProducts:missing.length,missingArticles:[...new Set(missing.map(x=>x.artNo))].sort()}}
function healthSummaryRows(){
  const unknown=unknownStoneDiagnostics(),images=imageCoverageDiagnostics(),gold=goldFreshnessInfo(),d=state.stoneDiagnostics||{},m=state.dataMeta||{},rows=[];
  if(m.package?.errors?.length)rows.push({level:'error',text:`資料包：${m.package.errors.join('；')}`});
  if(m.records){
    if(m.records.embedded)rows.push({level:'ok',text:`jmsdata Sheet 2「Universe Records」：${m.records.count} 筆文件 · ${state.exhibitionName||m.records.exhibitionName||'未命名展覽'}`});
    else if(m.records.newTemplate)rows.push({level:'ok',text:'新展覽：目前 jmsdata.xlsx 尚未有「Universe Records」工作表，屬正常；Recall 0 筆，文件流水號由 0001 開始，第一次 Confirm 後會自動建立隱藏的 Sheet 2。'});
  }
  if(!m.stock)rows.push({level:'error',text:'尚未載入 jmsdata。'});else{rows.push({level:state.stockDuplicateLots.length?'error':'ok',text:`jmsdata：${m.stock.count} 件 · 重複 LOTNO ${state.stockDuplicateLots.length}`});rows.push({level:state.stockIntegrityIssues.length?'error':'ok',text:`庫存 I:J:K:L / Balance：${state.stockIntegrityIssues.length?'有 '+state.stockIntegrityIssues.length+' 列需要檢查':'正常'}`})}
  if(!m.stone)rows.push({level:'error',text:'尚未載入 Stone List。'});else{rows.push({level:d.duplicates?.length?'warn':'ok',text:`Stone List BREAKDOWN 衝突：${d.duplicates?.length?d.duplicates.length+' 個同代號但石類／GROUP 不一致':'沒有衝突'}${d.multiAlias?.length?` · ${d.multiAlias.length} 個代號有多個 QUOTATION Alias（正常）`:''}`});rows.push({level:(d.missingGroup?.length||d.missingType?.length||d.missingEnglish?.length||d.missingQuotation?.length)?'warn':'ok',text:`Stone List 欄位：缺 GROUP ${d.missingGroup?.length||0} · 缺石類 ${d.missingType?.length||0} · 缺英文石名 ${d.missingEnglish?.length||0} · 缺 QUOTATION Alias ${d.missingQuotation?.length||0}`});rows.push({level:'info',text:`Stone List 鑽石 BREAKDOWN：${state.diamondStoneCodes.size} 個 · ${[...state.diamondStoneCodes].join(', ')||'沒有'}`});rows.push({level:'info',text:`Stone List 前綴重疊 ${d.prefixOverlaps?.length||0} 組（採用最長前綴配對，屬資訊提示）`});rows.push({level:unknown.length?'warn':'ok',text:`DESC2–DESC6 石種辨認：${unknown.length?unknown.length+' 個 token 未能辨認':'正常'}`})}
  if(m.customer)rows.push({level:m.customer.defaultRateCount?'warn':'ok',text:`Customer：${m.customer.count} 位 · ${m.customer.defaultRateCount||0} 位 Sales Rate 空白使用 0.34`});else rows.push({level:'warn',text:'Customer Excel 未載入。'});
  if(m.images)rows.push({level:images.missingProducts?'warn':'ok',text:`Pictures：${m.images.count} 張 · ${images.missingProducts?images.missingProducts+' 件貨品／'+images.missingArticles.length+' 個款號完全沒有圖片':'所有庫存款號都有圖片候選'}`});else rows.push({level:'warn',text:'Pictures 未載入。'});
  rows.push({level:gold.status==='ok'?'ok':gold.status==='missing'?'warn':'warn',text:gold.message});
  if(!m.template)rows.push({level:'warn',text:'Invoice Template 未載入；匯出會使用 PWA 標準格式。'});else rows.push({level:'ok',text:`Invoice Template：${m.template.name}`});
  if(state.importConflicts?.length)rows.push({level:'warn',text:`jmsdata 匯入：偵測到 ${state.importConflicts.length} 項文件記錄衝突，已按匯入時選擇處理。`});
  return{rows,unknown,images,gold};
}
function runHealthCheck(){
  const summary=healthSummaryRows();state.health=summary;const el=$('#healthCheckResults'),headline=$('#healthCheckStatus');if(!el||!headline){updateGoldFreshnessUI();return summary}
  const critical=summary.rows.filter(x=>x.level==='error').length,warnings=summary.rows.filter(x=>x.level==='warn').length;headline.textContent=critical?`資料包檢查：${critical} 項錯誤、${warnings} 項警告。請先處理錯誤。`:warnings?`資料包檢查：沒有阻擋錯誤，另有 ${warnings} 項警告可覆核。`:'資料包檢查：全部正常，可以開始展覽。';headline.className='notice '+(critical?'error':warnings?'warn':'ok');let html=summary.rows.map(x=>`<div class="health-row ${esc(x.level)}"><span>${x.level==='ok'?'✓':x.level==='error'?'✕':x.level==='warn'?'⚠':'○'}</span><div>${esc(x.text)}</div></div>`).join('');if(state.stockDuplicateLots.length)html+=`<details class="health-detail"><summary>重複 LOTNO 明細</summary>${state.stockDuplicateLots.slice(0,50).map(x=>`<div>${esc(x.lotNo)}</div>`).join('')}</details>`;if(state.stockIntegrityIssues.length)html+=`<details class="health-detail"><summary>I:J:K:L / Balance 異常 LOTNO</summary>${state.stockIntegrityIssues.slice(0,50).map(x=>`<div>${esc(x.lotNo)} · I ${x.vector.i} / J ${x.vector.j} / K ${x.vector.k} / L ${x.vector.l} / M ${x.vector.m}</div>`).join('')}</details>`;if(summary.unknown.length)html+=`<details class="health-detail"><summary>未辨認石種 token（${summary.unknown.length}）</summary>${summary.unknown.slice(0,80).map(x=>`<div>${esc(x.lotNo)} · ${esc(x.artNo)} · ${esc(x.field)} · <strong>${esc(x.token)}</strong> · ${esc(x.line)}</div>`).join('')}</details>`;if(summary.images.missingArticles.length)html+=`<details class="health-detail"><summary>完全沒有圖片的款號（${summary.images.missingArticles.length}）</summary>${summary.images.missingArticles.slice(0,80).map(x=>`<div>${esc(x)}</div>`).join('')}</details>`;el.innerHTML=html;
  updateGoldFreshnessUI();return summary;
}
function stoneConsistencyDetailsHTML(){const d=state.stoneDiagnostics||{},parts=[];if(!state.stoneAliases.size)return'Stone List 尚未載入。';if(d.duplicates?.length)parts.push(`<strong>BREAKDOWN 衝突（同代號但石類／GROUP 不一致）</strong><div>${d.duplicates.slice(0,30).map(x=>`${esc(x.code)}（Row ${x.firstRow} / ${x.row} · ${esc(x.firstType||'')} / ${esc(x.type||'')} · ${esc(x.firstGroup||'')} / ${esc(x.group||'')}）`).join('<br>')}</div>`);if(d.multiAlias?.length)parts.push(`<strong>多個 QUOTATION Alias（正常）</strong><div>${d.multiAlias.slice(0,60).map(x=>`${esc(x.code)}（例如 Row ${x.firstRow} / ${x.row}）`).join('<br>')}</div>`);if(d.missingGroup?.length)parts.push(`<strong>缺 GROUP</strong><div>${d.missingGroup.slice(0,30).map(x=>`${esc(x.code)} · Row ${x.row}`).join('<br>')}</div>`);if(d.missingType?.length)parts.push(`<strong>缺石類</strong><div>${d.missingType.slice(0,30).map(x=>`${esc(x.code)} · Row ${x.row}`).join('<br>')}</div>`);if(d.missingEnglish?.length)parts.push(`<strong>缺英文石名</strong><div>${d.missingEnglish.slice(0,30).map(x=>`${esc(x.code)} · Row ${x.row}`).join('<br>')}</div>`);if(d.missingQuotation?.length)parts.push(`<strong>缺 QUOTATION Alias</strong><div>${d.missingQuotation.slice(0,30).map(x=>`${esc(x.code)} · Row ${x.row}`).join('<br>')}</div>`);if(d.prefixOverlaps?.length)parts.push(`<strong>前綴重疊（正常，最長前綴優先）</strong><div>${d.prefixOverlaps.slice(0,60).map(x=>`${esc(x.short)} / ${esc(x.long)}`).join('<br>')}</div>`);return parts.length?parts.join('<hr>'):'Stone List 一致性沒有發現問題。'}
function showStoneConsistencyDetails(){const dialog=$('#stoneConsistencyDialog'),box=$('#stoneConsistencyContent');if(box)box.innerHTML=stoneConsistencyDetailsHTML();if(dialog?.showModal)dialog.showModal();else dialog?.setAttribute('open','')}
function backupLocalStorageSnapshot(){const out={};try{for(let i=0;i<localStorage.length;i++){const key=localStorage.key(i);if(key&&key.startsWith('universe'))out[key]=localStorage.getItem(key)}}catch{}return out}
function backupFileStamp(){const d=new Date(),p=n=>String(n).padStart(2,'0');return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`}
function exportPwaBackup(){saveDocumentDraft();saveInventoryHistory();saveImageOverridesLocal();syncExhibitionNameFromInput();const payload={app:'Universe Invoice',backupVersion:1,appVersion:APP_VERSION,exportedAt:new Date().toISOString(),exhibitionSession:state.exhibitionSession||'',exhibitionName:state.exhibitionName||'',localStorage:backupLocalStorageSnapshot(),dataMeta:state.dataMeta||{},notes:'Confirmed Invoice / Consignment / Quotation history is NOT stored here. Official records are embedded in Sheet 2 (Universe Records) of jmsdata.xlsx.'};downloadBlob(new Blob([JSON.stringify(payload,null,2)],{type:'application/json'}),`UniverseInvoice_PWA_Backup_${backupFileStamp()}.json`);status('#backupStatus','PWA 本機備份已匯出。正式 Confirm 文件歷史不在此備份內，請保存最新 jmsdata.xlsx（Sheet 2：Universe Records）。','ok')}
function compareAppVersions(a,b){const aa=String(a||'0').split('.').map(x=>Number(x)||0),bb=String(b||'0').split('.').map(x=>Number(x)||0),n=Math.max(aa.length,bb.length);for(let i=0;i<n;i++){const d=(aa[i]||0)-(bb[i]||0);if(d)return d}return 0}
async function restorePwaBackupFile(file){try{const raw=JSON.parse(await file.text());if(raw?.app!=='Universe Invoice'||Number(raw?.backupVersion)!==1||typeof raw.localStorage!=='object')throw new Error('不是有效或受支援的 Universe Invoice PWA 備份檔。');const newer=compareAppVersions(raw.appVersion,APP_VERSION)>0;const msg=`備份版本：v${raw.appVersion||'未知'}\n匯出時間：${raw.exportedAt?new Date(raw.exportedAt).toLocaleString('zh-HK'):'未知'}\nExhibition：${raw.exhibitionName||'未命名'}\n\n還原會覆蓋目前 PWA 本機草稿及設定，但不會改動你在「檔案」App 的原始資料包；正式 Confirm 文件歷史請以 jmsdata.xlsx 的第二個「Universe Records」工作表為準。${newer?'\n\n⚠ 備份來自較新的 PWA 版本，可能有相容性風險。':''}\n\n確定還原？`;if(!confirm(msg))return;for(let i=localStorage.length-1;i>=0;i--){const key=localStorage.key(i);if(key&&key.startsWith('universe'))localStorage.removeItem(key)}for(const [key,value] of Object.entries(raw.localStorage||{}))if(key.startsWith('universe')&&value!==null)localStorage.setItem(key,String(value));alert('PWA 備份已還原。頁面會重新載入；請再匯入最新展覽資料包以恢復庫存、Stone List、GoldSilver 及 Pictures。');location.reload()}catch(err){alert('還原 PWA 備份失敗：'+(err.message||err))}}
function clearAllPwaLocalData(){if(!confirm('這是高風險操作。\n\n「清空所有 PWA 本機資料」會刪除本機草稿、Image Overrides 暫存、Exhibition 暫存及設定。正式 Confirm 文件歷史以資料包內 jmsdata.xlsx 的第二個「Universe Records」工作表為準。\n\n不會刪除「檔案」App 內的 Excel／Pictures。\n\n確定繼續？'))return;if(!confirm('最後確認：真的要清空所有 PWA 本機資料？此操作不能由 PWA 自動復原。'))return;try{for(let i=localStorage.length-1;i>=0;i--){const key=localStorage.key(i);if(key&&key.startsWith('universe'))localStorage.removeItem(key)}}catch{}alert('所有 PWA 本機資料已清空。頁面會重新載入。');location.reload()}
function diagnosticProductByLot(lot){const key=normalizeLotInput(lot);return state.stockCatalog.get(key)||state.products.get(key)||state.inventoryHistory.get(key)||state.items.find(x=>String(x.lotNo)===key)||null}
function buildDiagnosticData(product){if(!product)return null;const row=state.stockRowByLot.get(String(product.lotNo)),vector=row?inventoryVectorFromRow(row):null,statusName=vector?inventoryStatusFromVector(vector):'UNKNOWN',allCodes=allStoneCodesForProduct(product),stones=allCodes.map(code=>({code,type:isDiamondStoneCode(code)?'Diamond':'Color Stone',group:stoneGroupForCode(code)||'',aliases:stoneImageAliasesForCode(code)})),unknown=[];for(const line of product.descriptions||[])for(const x of stonePartsFromDescriptionLine(line))if(!x.code)unknown.push({line,token:x.part});const imgs=state.imageFiles.get(product.artNo)||[],match=chooseImageMatch(product),override=imageOverrideForLot(product.lotNo),q=quotationPriceDetails(product),rate=Number($('#salesRate')?.value)||0;return{generatedAt:new Date().toISOString(),lotNo:product.lotNo,artNo:product.artNo,status:statusName,inventoryVector:vector,priceU:Number(product.price)||0,salesRate:rate,baseUsdUnitPrice:Math.ceil((Number(product.price)||0)*rate),descriptions:[...(product.descriptions||[])],completionDate:product.completionDate||'',metalToken:originalMetalToken(product)||'',stoneCodes:stones,unknownStoneTokens:unknown,multiColor:isMultiColorProduct(product),imageCandidates:imgs.map(x=>({fileName:x.fileName,variant:x.variant,dup:x.dup||0,metalToken:imageMetalToken(x.variant),multi:isMultiImageVariant(x.variant)})),selectedImage:{...match},imageOverride:override||null,quotation:state.documentType==='quotation'?{mode:quoteModeName(),ready:!!q?.ready,details:q}:null}}
function diagnosticText(data){if(!data)return'';const lines=[];lines.push(`Universe Invoice Diagnostic · v${APP_VERSION}`,`Generated: ${data.generatedAt}`,`LOTNO: ${data.lotNo}`,`ARTNO: ${data.artNo}`,`Status: ${data.status}`,`U Price: ${data.priceU}u`,`Sales Rate: ${data.salesRate}`,`Base USD Unit Price: ${data.baseUsdUnitPrice}`,`Completion Date: ${data.completionDate||'—'}`,`Metal: ${data.metalToken||'—'}`,'','Descriptions:',...data.descriptions.map((x,i)=>`DESC${i+1}: ${x}`),'','Stones:',...data.stoneCodes.map(x=>`${x.code} · ${x.type} · GROUP=${x.group||'—'} · Alias=${(x.aliases||[]).join(', ')||'—'}`),`MULTI: ${data.multiColor?'Yes':'No'}`);if(data.unknownStoneTokens.length)lines.push('','Unknown stone tokens:',...data.unknownStoneTokens.map(x=>`${x.token} · ${x.line}`));lines.push('','Image candidates:',...data.imageCandidates.map(x=>`${x.fileName} · variant=${x.variant} · dup=${x.dup} · metal=${x.metalToken||'—'} · multi=${x.multi?'Yes':'No'}`),'',`Selected image: ${data.selectedImage.fileName||data.selectedImage.variant||'No Image'} · grayscale=${data.selectedImage.grayscale?'Yes':'No'} · manual=${data.selectedImage.manualOverride?'Yes':'No'}`);if(data.imageOverride)lines.push(`Image Override: ${JSON.stringify(data.imageOverride)}`);if(data.quotation)lines.push('',`Quotation mode: ${data.quotation.mode}`,`Quotation ready: ${data.quotation.ready?'Yes':'No'}`,JSON.stringify(data.quotation.details,null,2));return lines.join('\n')}
function openDiagnosticReport(product){if(!product)return alert('找不到這件貨。');state.diagnosticLot=String(product.lotNo);const data=buildDiagnosticData(product),dialog=$('#diagnosticDialog'),box=$('#diagnosticContent');if(box)box.textContent=diagnosticText(data);if(dialog?.showModal)dialog.showModal();else dialog?.setAttribute('open','')}
function openDiagnosticByLot(){const p=diagnosticProductByLot($('#diagnosticLotInput')?.value||'');if(!p)return alert('找不到這個 LOTNO；請先匯入 jmsdata。');openDiagnosticReport(p)}
function exportDiagnosticReport(){const p=diagnosticProductByLot(state.diagnosticLot);if(!p)return alert('診斷貨品已不存在。');const data=buildDiagnosticData(p);downloadBlob(new Blob([diagnosticText(data)],{type:'text/plain;charset=utf-8'}),`UniverseInvoice_Diagnostic_${p.lotNo}_${backupFileStamp()}.txt`)}
$('#stoneMappingInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const count=await importStoneFile(f);status('#stoneMappingStatus',`已匯入 ${f.name}：${count} 個石種代碼對照。`,'ok');setImportCollapsed('stone',true);for(const item of state.items)if(!item.customImage)applyAutoImageMatch(item);renderItems()}catch(err){status('#stoneMappingStatus','匯入失敗：'+(err.message||err),'error');setImportCollapsed('stone',false)}};
$('#articleMappingInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const count=await importArticleFile(f);status('#articleMappingStatus',`已匯入 ${f.name}：${count} 個 Article 對照。`,'ok');setImportCollapsed('article',true);schedulePreview()}catch(err){status('#articleMappingStatus','匯入失敗：'+(err.message||err),'error');setImportCollapsed('article',false)}};
$('#invoiceTemplateInput').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{await importTemplateFile(f);status('#invoiceTemplateStatus',`已匯入 ${f.name}；匯出文件時會套用此範本。`,'ok');setImportCollapsed('template',true)}catch(err){state.invoiceTemplateBuffer=null;status('#invoiceTemplateStatus','匯入失敗：'+(err.message||err),'error');setImportCollapsed('template',false)}};

function imageArtAliasWithoutSuffixDot(art){
  const value=String(art||'').toUpperCase();
  return /\.[A-Z0-9]+$/.test(value)?value.replace(/\.([A-Z0-9]+)$/,'$1'):'';
}
function imageStemMatchesPrefix(stemUpper,prefix){return stemUpper===prefix||stemUpper.startsWith(prefix+' ')}
function parseImage(file,knownArts=null){
  const stem=file.name.replace(/\.[^.]+$/,'').trim(),stemUpper=stem.toUpperCase();
  const arts=knownArts||[...new Set([...state.stockCatalog.values(),...state.products.values(),...state.inventoryHistory.values()].map(x=>normArt(x.artNo)).filter(Boolean))].sort((a,b)=>b.length-a.length);
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
$('#imageFolderInput').onchange=async e=>{const result=await importImageFiles(e.target.files);for(const item of state.items)if(!item.customImage)applyAutoImageMatch(item);status('#imageStatus',`已選擇圖片 Folder：${result.images} 張圖片，配對 ${result.matched} 個款號。`,'ok');setImportCollapsed('images',true);renderItems();renderStockSearch()};
$('#exhibitionPackageInput').onchange=async e=>{
  const files=[...e.target.files];if(!files.length)return;
  if(!confirmDiscardPendingImageOverrides()){e.target.value='';return}
  const root=(files[0].webkitRelativePath||'').split('/')[0]||'Exhibition Package';state.packageName=root;state.packageFiles=files;state.customPackageImages=new Map();state.packageImageDirtyFiles=new Set();state.dataMeta.package={name:root,fileCount:files.length,loadedAt:new Date().toISOString(),errors:[]};updateDataVersionPanel();
  resetFormalRecordContext();
  const excels=files.filter(f=>/\.(xlsx?|xls)$/i.test(f.name));
  const byName=(re)=>excels.find(f=>re.test(f.name));
  let stock=byName(/^jmsdata(?:\s.*|\(\d+\))?\.xlsx$/i)||byName(/^jmsdata.*\.xlsx$/i);
  let customer=byName(/customer/i);
  let stone=byName(/stone\s*list.*shape.*cutting/i);
  let article=byName(/article\s*mapping/i);
  let template=byName(/invoice\s*(master\s*)?template/i)||byName(/template/i);
  let gold=byName(/^goldsilver(?:\s.*)?\.(xls|xlsx)$/i)||byName(/gold\s*silver/i)||byName(/historical\s*london/i);
  const imageFiles=files.filter(f=>String(f.type).startsWith('image/')||/\.(jpe?g|png|webp)$/i.test(f.name));
  const lines=[];let errors=[];
  status('#packageStatus',`正在讀取 ${root}…`);
  try{
    if(!stock)errors.push('找不到 jmsdata.xlsx');else{const n=await importStockFile(stock,{exhibitionNameHint:root});lines.push(`✓ ${stock.name} · ${n} 件貨品 · Available ${state.products.size} · Records ${officialDocumentCount()} 筆${state.dataMeta.records?.newTemplate?' · 新展覽空白記錄':''} · 圖片選擇 ${state.imageOverrides.size} 項 · ${new Date(stock.lastModified).toLocaleString('zh-HK')}`)}
    if(!customer)errors.push('找不到客戶 Excel');else{const n=await importCustomerFile(customer);lines.push(`✓ ${customer.name} · ${n} 位客戶`)}
    if(!stone)errors.push('找不到 Stone List & Shape & Cutting.xlsx');else{const n=await importStoneFile(stone);lines.push(`✓ ${stone.name} · ${n} 個石種代碼`)}
    if(!template)errors.push('找不到 Invoice Template.xlsx');else{await importTemplateFile(template);lines.push(`✓ ${template.name} · Template 已載入`)}
    if(article){const n=await importArticleFile(article);lines.push(`✓ ${article.name} · ${n} 個 Article 對照`)}else lines.push('○ Article Mapping 未提供（不顯示 Article）');
    if(gold){const g=await importGoldSilverFile(gold);lines.push(`✓ ${gold.name} · ${g.count} 個 Gold PM 交易日 · 最新 ${g.latestDate} USD ${g.latestPm.toFixed(2)} / oz`)}else{state.quote.goldPrices=new Map();state.quote.goldDates=[];state.quote.goldDataName='';state.quote.goldDataRows=0;delete state.dataMeta.gold;lines.push('○ GoldSilver.xlsx 未提供（Quotation 金價日期無法自動配對）')}
    if(!imageFiles.length){errors.push('找不到 Pictures 內的圖片');delete state.dataMeta.images}else{const r=await importImageFiles(imageFiles);for(const item of state.items)if(!item.customImage)applyAutoImageMatch(item);renderStockSearch();lines.push(`✓ Pictures · ${r.images} 張圖片 · 配對 ${r.matched} 個款號`)}
    $('#packageSummary').innerHTML=lines.map(x=>`<div>${esc(x)}</div>`).join('')+(errors.length?`<div class="package-errors">${errors.map(x=>'✕ '+esc(x)).join('<br>')}</div>`:'');
    if(errors.length)status('#packageStatus',`${root} 未完整載入，請補回缺少的資料。`,'error');else{status('#packageStatus',`${root} 已完成匯入，可以開始建立文件。`,'ok');document.querySelector('.advanced-imports').open=false}
    status('#stockStatus',stock?`已由資料包匯入 ${stock.name}：Available ${state.products.size} 件；圖片選擇 ${state.imageOverrides.size} 項。文件歷史來自 jmsdata.xlsx 的第二個「Universe Records」工作表。`:'尚未匯入倉存。',stock?'ok':'error');
    status('#customerStatus',customer?`已由資料包匯入 ${customer.name}：${state.customers.size} 位客戶。`:'尚未匯入客戶。',customer?'ok':'error');
    status('#imageStatus',imageFiles.length?`已由資料包匯入 Pictures：${imageFiles.length} 張圖片，配對 ${state.imageFiles.size} 個款號。`:'尚未匯入圖片。',imageFiles.length?'ok':'error');
    if(stone)status('#stoneMappingStatus',`已由資料包匯入 ${stone.name}。`,'ok');
    if(template)status('#invoiceTemplateStatus',`已由資料包匯入 ${template.name}。`,'ok');
    if(article)status('#articleMappingStatus',`已由資料包匯入 ${article.name}。`,'ok');
    state.dataMeta.package.errors=[...errors];updateDataVersionPanel();runHealthCheck();updateGoldQuoteUI(gold?`已由 ${gold.name} 載入最新及歷史 London PM。`:'GoldSilver.xlsx 未載入；請先更新檔案並重新匯入資料包。',gold?'ok':'warn');
    updateTotals();renderItems();syncDraftAfterDataImport();
  }catch(err){console.error(err);if(state.dataMeta.package)state.dataMeta.package.errors=[String(err.message||err)];updateDataVersionPanel();runHealthCheck();status('#packageStatus','資料包匯入失敗：'+(err.message||err),'error')}
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
const MULTI_COLOR_GROUP_THRESHOLD=3,MULTI_STONE_CODE_THRESHOLD=3,MULTI_MIN_COLOR_GROUPS=2;
function imageStoneCodesForProduct(p){
  return stoneCodesForProduct(p).filter(code=>!isDiamondStoneCode(code));
}
function normalDesiredStoneVariants(p){
  const out=[];for(const code of imageStoneCodesForProduct(p)){for(const variant of stoneImageAliasesForCode(code)){const v=String(variant||'').toUpperCase();if(v&&!out.includes(v))out.push(v)}}return out;
}
function desiredStoneVariants(p){return normalDesiredStoneVariants(p)}
function productColorGroups(p){
  const groups=[];for(const code of imageStoneCodesForProduct(p)){const group=stoneGroupForCode(code);if(group&&!groups.includes(group))groups.push(group)}return groups;
}
function isMultiColorProduct(p){
  const groups=productColorGroups(p),codes=imageStoneCodesForProduct(p);
  return groups.includes('MULTI')||groups.length>=MULTI_COLOR_GROUP_THRESHOLD||(codes.length>=MULTI_STONE_CODE_THRESHOLD&&groups.length>=MULTI_MIN_COLOR_GROUPS);
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
function isMultiImageVariant(variant){return imageStoneSignature(variant).split('+').some(x=>/\bMULTI\b/.test(x.trim()))}
function variantMatchesProductStoneCode(variant,code){return stoneImageAliasesForCode(code).some(alias=>variantContainsStone(variant,alias))}
function matchedProductStoneIndexes(variant,codes){const out=[];codes.forEach((code,index)=>{if(variantMatchesProductStoneCode(variant,code))out.push(index)});return out}
function chooseAutomaticImageMatch(p){
  const imgs=state.imageFiles.get(p?.artNo)||[];
  if(!imgs.length)return{variant:'Default',grayscale:false,stoneMatched:false,colorMatched:false};
  const codes=imageStoneCodesForProduct(p),metal=originalMetalToken(p),multiColor=isMultiColorProduct(p),preferNlDuplicate=articleType(p?.artNo)==='NL';
  const metalScore=img=>{const t=imageMetalToken(img.variant);return metal&&t===metal?2:!t?1:0};
  const nlDuplicateScore=img=>preferNlDuplicate&&Number(img?.dup)===1?1:0;
  const regularImgs=imgs.filter(x=>x.variant!=='Default'&&!isMultiImageVariant(x.variant));
  const scoreRegular=(img,index,requireAll=false)=>{
    const matchedIndexes=matchedProductStoneIndexes(img.variant,codes),matchedCount=matchedIndexes.length;if(!matchedCount||requireAll&&matchedCount!==codes.length)return null;
    const tokenCount=imageStoneTokenCount(img.variant),allMatched=matchedCount===codes.length;
    let stoneScore=matchedCount*1000+(allMatched?5000:0)-Math.abs(tokenCount-matchedCount)*80;
    stoneScore+=Math.max(0,40-(matchedIndexes[0]||0)*5);
    return{img,index,matchedCount,allMatched,stoneScore,metalScore:metalScore(img),nlDuplicateScore:nlDuplicateScore(img),tokenCount};
  };
  const sortScored=(a,b)=>b.stoneScore-a.stoneScore||b.metalScore-a.metalScore||b.nlDuplicateScore-a.nlDuplicateScore||a.tokenCount-b.tokenCount||a.img.variant.length-b.img.variant.length||a.index-b.index;
  const exactCombos=regularImgs.map((img,index)=>scoreRegular(img,index)).filter(x=>x&&x.allMatched&&codes.length>1).sort(sortScored);
  if(exactCombos.length){const best=exactCombos[0];return{variant:best.img.variant,grayscale:false,stoneMatched:true,colorMatched:best.metalScore===2}}
  if(multiColor){
    const multi=imgs.filter(x=>isMultiImageVariant(x.variant)).map((img,index)=>({img,index,metalScore:metalScore(img),nlDuplicateScore:nlDuplicateScore(img)})).sort((a,b)=>b.metalScore-a.metalScore||b.nlDuplicateScore-a.nlDuplicateScore||a.img.variant.length-b.img.variant.length||a.index-b.index)[0];
    if(multi)return{variant:multi.img.variant,grayscale:false,stoneMatched:true,colorMatched:multi.metalScore===2};
    // No MULTI image exists: use a real single-stone match before falling back to grayscale.
    // This deliberately does not use carat weight. Product stone order and exact aliases decide the result.
    const normal=regularImgs.map((img,index)=>scoreRegular(img,index)).filter(Boolean).sort(sortScored)[0];
    if(normal)return{variant:normal.img.variant,grayscale:false,stoneMatched:true,colorMatched:normal.metalScore===2};
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
function chooseImageMatch(p){
  const override=resolvedImageOverride(p);
  if(override)return{variant:override.variant,grayscale:override.grayscale,stoneMatched:true,colorMatched:true,manualOverride:true,fileName:override.fileName};
  return chooseAutomaticImageMatch(p);
}
function chooseVariant(p){return chooseImageMatch(p).variant}
function applyAutoImageMatch(item){
  if(!item)return;
  const match=chooseImageMatch(item);item.imageVariant=match.variant;item.imageGrayscale=!!match.grayscale;item.imageAutoMatched=!match.manualOverride;item.imageOverrideFile=match.fileName||'';
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

const PDFJS_WORKER_URL='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
function parsePdfMoney(v){return Number(String(v||'').replace(/[$,\s]/g,''))||0}
function parseInvoicePdfDate(raw){
  const text=norm(raw).replace(/,/g,' '),m=text.match(/\b(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})\b/);if(!m)return'';
  const months={january:1,february:2,march:3,april:4,may:5,june:6,july:7,august:8,september:9,october:10,november:11,december:12},mo=months[m[2].toLowerCase()];if(!mo)return'';
  return `${m[3]}-${String(mo).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
}
function groupPdfTextItems(items){
  const rows=[];for(const item of items||[]){const str=norm(item.str);if(!str)continue;const x=Number(item.transform?.[4]||0),y=Number(item.transform?.[5]||0);let row=rows.find(r=>Math.abs(r.y-y)<2.5);if(!row){row={y,parts:[]};rows.push(row)}row.parts.push({x,str})}
  return rows.sort((a,b)=>b.y-a.y).map(r=>r.parts.sort((a,b)=>a.x-b.x).map(p=>p.str).join(' ').replace(/\s+/g,' ').trim()).filter(Boolean);
}
function parseInvoicePdfLines(lines){
  const all=lines.join('\n'),invoice=(all.match(/\bNo\.?\s*:\s*(INV\d{6})\b/i)||[])[1]||'',dateRaw=(all.match(/(?:Invoice\s+Date|Date)\s*:\s*([^\n]+)/i)||[])[1]||'',date=parseInvoicePdfDate(dateRaw);
  let total=0;for(const line of lines){const m=line.match(/\bTotal\s*:\s*\(\s*[A-Z]{3}\s*\)\s*\$?([\d,]+(?:\.\d{2})?)/i);if(m)total=parsePdfMoney(m[1])}
  const items=[];let pending=null;
  const pushPending=()=>{if(pending&&pending.lotNo&&pending.qty&&Number.isFinite(pending.unitPrice))items.push(pending);pending=null};
  for(const line of lines){
    const lotMatch=line.match(/(?:^|\s)(\d+)\s+Lot\.?\s*No\.?\s*:\s*(\d+)/i)||line.match(/Lot\.?\s*No\.?\s*:\s*(\d+)/i);
    if(lotMatch){pushPending();pending={seq:Number(lotMatch.length>2?lotMatch[1]:items.length+1)||items.length+1,lotNo:String(lotMatch.length>2?lotMatch[2]:lotMatch[1]),qty:0,unitPrice:NaN};
      const tail=line.slice((lotMatch.index||0)+lotMatch[0].length),pm=tail.match(/\b(\d+(?:\.\d+)?)\s+(PC|PR|ST)\s+\$?([\d,]+(?:\.\d{2})?)\s+\$?([\d,]+(?:\.\d{2})?)/i);if(pm){pending.qty=Number(pm[1])||1;pending.unitPrice=parsePdfMoney(pm[3]);pushPending()}continue;
    }
    if(pending){const pm=line.match(/^\s*(\d+(?:\.\d+)?)\s+(PC|PR|ST)\s+\$?([\d,]+(?:\.\d{2})?)\s+\$?([\d,]+(?:\.\d{2})?)/i);if(pm){pending.qty=Number(pm[1])||1;pending.unitPrice=parsePdfMoney(pm[3]);pushPending()}}
  }
  pushPending();return{invoiceNo:invoice,date,total,items};
}
async function readInvoicePdf(file){
  if(!window.pdfjsLib)throw new Error('PDF 讀取程式未載入，請確認網絡後重新開啟 PWA。');
  pdfjsLib.GlobalWorkerOptions.workerSrc=PDFJS_WORKER_URL;const data=await file.arrayBuffer(),doc=await pdfjsLib.getDocument({data}).promise,lines=[];
  for(let n=1;n<=doc.numPages;n++){const page=await doc.getPage(n),content=await page.getTextContent();lines.push(...groupPdfTextItems(content.items))}
  return parseInvoicePdfLines(lines);
}
function importedPdfValidation(parsed){
  const errors=[];if(!parsed.invoiceNo)errors.push('讀不到 Invoice No.');if(!parsed.date)errors.push('讀不到 Date');if(!parsed.items.length)errors.push('讀不到 LOTNO / Qty / Unit Price');if(!(parsed.total>=0))errors.push('讀不到 Total');
  const seen=new Set();for(const r of parsed.items){if(seen.has(r.lotNo))errors.push(`LOTNO ${r.lotNo} 在 PDF 重複`);seen.add(r.lotNo);const p=state.products.get(r.lotNo);if(!p)errors.push(`LOTNO ${r.lotNo} 目前不是 Available／jmsdata 找不到可售貨品`);if(!(r.qty>0))errors.push(`LOTNO ${r.lotNo} Qty 無效`);if(!(r.unitPrice>=0))errors.push(`LOTNO ${r.lotNo} Unit Price 無效`)}
  if(findLatestDocument('invoice',parsed.invoiceNo))errors.push(`${parsed.invoiceNo} 已存在 Records，請用 Recall，不要重複匯入 PDF`);return errors;
}
async function importInvoicePdfFile(file){
  const box=$('#invoicePdfImportStatus'),btn=$('#importInvoicePdfBtn');if(!file)return;btn.disabled=true;status('#invoicePdfImportStatus','正在讀取 Invoice PDF…','');box?.classList.remove('hidden');
  try{
    if(state.documentType!=='invoice')throw new Error('「匯入 PDF」只適用於 Invoice。');if(!state.stockAllRows.length)throw new Error('請先匯入展覽資料包 / jmsdata.xlsx。');if(!norm($('#customerCode').value)||!norm($('#customerName').value))throw new Error('請先選擇客戶，再匯入 PDF。');
    const parsed=await readInvoicePdf(file),errors=importedPdfValidation(parsed);if(errors.length)throw new Error(errors.slice(0,12).join('\n')+(errors.length>12?`\n…另有 ${errors.length-12} 項`:''));
    if(state.items.length&&!confirm(`目前草稿已有 ${state.items.length} 款貨品。\n\n匯入 PDF 會用 ${parsed.invoiceNo} 的 ${parsed.items.length} 款貨品取代目前草稿，是否繼續？`))return;
    releaseCustomImages();state.items=[];state.recall=null;state.deliveryReturns=new Set();setRecallLock(false);$('#recallActive')?.classList.add('hidden');$('#invoiceNo').value=parsed.invoiceNo;$('#invoiceDate').value=parsed.date;
    const code=currencyCode();for(const r of parsed.items){const p=state.products.get(r.lotNo);addProductToDocument(p,{fromSearch:true});const item=state.items[state.items.length-1];item.qty=r.qty;const overrides=activePriceOverrides(item);overrides[code]=r.unitPrice;markManualPriceOverride(item,code);item.unitPrice=r.unitPrice;if(code==='USD')item.usdUnitPrice=r.unitPrice}
    state.items.forEach((x,i)=>x.seq=i+1);const sub=roundCurrency(state.items.reduce((a,x)=>a+x.qty*(Number(x.unitPrice)||0),0),code),diff=roundCurrency(sub-parsed.total,code);$('#discountAmount').value=diff>=0?diff:0;renderItems();updateTotals();schedulePreview();scheduleDraftSave();
    const rebuilt=totals().total,delta=Math.abs(rebuilt-parsed.total),message=`已匯入 ${parsed.invoiceNo} · ${parsed.date} · ${parsed.items.length} 款 / ${state.items.reduce((a,x)=>a+x.qty,0)} 件 · PDF Total ${fmt(parsed.total,code)} · PWA 重算 ${fmt(rebuilt,code)}${delta<0.01?' ✓ 一致':' ⚠ 不一致'}`;status('#invoicePdfImportStatus',message,delta<0.01?'ok':'warn');status('#addMessage',`PDF 已作為正式 Invoice 草稿輸入；Confirm 後會正常更新 jmsdata.xlsx / Records。`,'ok');
  }catch(err){console.error(err);status('#invoicePdfImportStatus','PDF 匯入失敗：'+(err.message||err),'error');alert('PDF 匯入失敗：\n'+(err.message||err))}finally{btn.disabled=false}
}
function addProductToDocument(p,{fromSearch=false}={}){
  if(!p)return false;const lot=String(p.lotNo);
  if(state.items.some(x=>x.lotNo===lot)){status('#addMessage',`LOTNO ${lot} 已在 ${documentLabels().short}。`,'error');return false}
  const rate=Number($('#salesRate').value)||0,usdUnitPrice=Math.ceil((Number(p.price)||0)*rate),match=chooseImageMatch(p);const item={...productSnapshot(p),id:Date.now()+Math.random(),seq:state.items.length+1,qty:1,usdUnitPrice,currencyPrices:{},quote18kCurrencyPrices:{},quote14kCurrencyPrices:{},quote14kSameWeightCurrencyPrices:{},unitPrice:0,imageVariant:match.variant,imageGrayscale:!!match.grayscale,imageAutoMatched:!match.manualOverride,imageOverrideFile:match.fileName||'',delivered:false};state.items.push(item);item.unitPrice=convertedFromUsd(effectiveUsdPrice(item));
  status('#addMessage',`已加入 ${p.artNo} / LOTNO ${lot}`,'ok');renderItems();if(!fromSearch)invalidateStockSearchView();return true;
}
function addLot(raw){
  const lot=normalizeLotInput(raw);
  if(!lot){status('#addMessage','請輸入 LOTNO。','error');refocusLotInput();return false}
  let p=state.documentType==='quotation'?state.stockCatalog.get(lot):state.products.get(lot);
  if(p&&state.documentType==='quotation'){const row=state.stockRowByLot.get(lot),itemStatus=inventoryStatusFromVector(inventoryVectorFromRow(row||{}));if(!['AVAILABLE','CONSIGNED','SOLD_ON_HAND','SOLD_DELIVERED'].includes(itemStatus))p=null}
  if(!p){const message=state.documentType==='quotation'?`找不到狀態有效的 LOTNO ${lot}。`:`找不到可售 LOTNO ${lot}。可到「款式搜尋」查看已售／寄賣狀態。`;status('#addMessage',message,'error');refocusLotInput(true);return false}
  if(!addProductToDocument(p))return false;$('#lotInput').value='';setTimeout(()=>{$('#invoiceItems').scrollTo({top:0,behavior:'smooth'});refocusLotInput()},50);return true;
}
$('#addLotBtn').onclick=()=>addLot($('#lotInput').value);
$('#lotInput').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();addLot(e.target.value)}};
$('#lotInput').oninput=e=>{e.target.value=e.target.value.replace(/[，,。\.\-–—_]/g,' ')};
$('#lotInput').onfocus=e=>setTimeout(()=>e.target.select(),50);
$('#importInvoicePdfBtn').onclick=()=>$('#invoicePdfInput').click();$('#invoicePdfInput').onchange=async e=>{const f=e.target.files?.[0];e.target.value='';if(f)await importInvoicePdfFile(f)};
function getImg(item){
  if(item?.customImage?.file&&item.customImage.url)return {...item.customImage,grayscale:false};
  const overrideRow=imageOverrideForLot(item?.lotNo),overrideFile=item?.imageOverrideFile||overrideRow?.IMAGE_FILE;
  if(overrideFile){const selected=findImageByFileName(overrideFile,item?.artNo);if(selected)return{...selected,grayscale:item?.imageOverrideFile?!!item?.imageGrayscale:!!numberValue(overrideRow?.GRAYSCALE)}}
  const arr=state.imageFiles.get(item.artNo)||[];
  const selected=arr.find(x=>x.variant===item.imageVariant)||arr[0];
  return selected?{...selected,grayscale:!!item?.imageGrayscale}:selected;
}
function imageDisplayForProduct(product){
  const match=chooseImageMatch(product),arr=state.imageFiles.get(product?.artNo)||[];
  const image=match.fileName?findImageByFileName(match.fileName,product?.artNo):arr.find(x=>x.variant===match.variant)||arr.find(x=>String(x.variant).toUpperCase()===String(match.variant).toUpperCase())||arr.find(x=>x.variant==='Default')||arr[0]||null;
  return{match,image,src:placeholder(product?.artNo||'No Image'),grayscale:!!match.grayscale,manual:!!match.manualOverride};
}
function imageNeedsAttention(product){
  if(resolvedImageOverride(product))return false;
  const arr=state.imageFiles.get(product?.artNo)||[];if(!arr.length)return true;
  return !!chooseAutomaticImageMatch(product).grayscale;
}
function currentDocumentImageStats(){
  let normal=0,fallback=0,missing=0;
  for(const item of state.items){const selected=getImg(item);if(!selected?.file)missing++;else if(imageNeedsAttention(item))fallback++;else normal++}
  return{total:state.items.length,normal,fallback,missing,issues:fallback+missing};
}
function renderDocumentImageWarning(){
  const el=$('#documentImageWarning');if(!el)return;const x=currentDocumentImageStats();
  if(!x.total){el.textContent='圖片：尚未加入貨品。';el.className='notice image-warning-summary';return}
  el.textContent=`${x.total} 款貨品 · 圖片正常 ${x.normal} · 待處理 ${x.issues}${x.issues?`（黑白 ${x.fallback} · 無圖 ${x.missing}）`:''}`;
  el.className='notice image-warning-summary '+(x.issues?'warn':'ok');
}
function confirmImageIssuesBeforeExport(){const x=currentDocumentImageStats();return !x.issues||confirm(`有 ${x.issues} 款圖片待處理（黑白 ${x.fallback}／無圖 ${x.missing}），仍然繼續匯出？`)}
function imageEditorProduct(){
  const lot=String(state.stockImageEditLot||'');
  return state.items.find(x=>String(x.lotNo)===lot)||state.stockCatalog.get(lot)||state.products.get(lot)||state.inventoryHistory.get(lot)||null;
}
function registerManualImage(product,file,source='upload'){
  const image={variant:`MANUAL ${product.lotNo}`,url:URL.createObjectURL(file),fileName:file.name,file,dup:0,manualSource:source};
  const arr=state.imageFiles.get(product.artNo)||[];
  const previousIndex=arr.findIndex(x=>String(x.fileName).toUpperCase()===String(image.fileName).toUpperCase());
  if(previousIndex>=0){const old=arr[previousIndex];if(old?.url?.startsWith('blob:'))try{URL.revokeObjectURL(old.url)}catch{}arr[previousIndex]=image}else arr.unshift(image);
  state.imageFiles.set(product.artNo,arr);state.imageFilesByName.set(String(image.fileName).toUpperCase(),image);
  return image;
}
async function compressManualImage(file,product,source='upload'){
  if(!file||!product)throw new Error('沒有選擇圖片。');
  if(!String(file.type||'').startsWith('image/')&&!/\.(jpe?g|png|webp|heic|heif)$/i.test(file.name||''))throw new Error('請選擇圖片檔案。');
  let blob=file;
  try{
    const sourceUrl=URL.createObjectURL(file);
    try{
      const img=await new Promise((resolve,reject)=>{const i=new Image();i.onload=()=>resolve(i);i.onerror=reject;i.src=sourceUrl});
      const iw=img.naturalWidth||img.width,ih=img.naturalHeight||img.height,maxSide=1800,scale=Math.min(1,maxSide/Math.max(iw,ih)),w=Math.max(1,Math.round(iw*scale)),h=Math.max(1,Math.round(ih*scale));
      const canvas=document.createElement('canvas');canvas.width=w;canvas.height=h;const ctx=canvas.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,w,h);ctx.drawImage(img,0,0,w,h);
      const converted=await new Promise(resolve=>canvas.toBlob(resolve,'image/jpeg',.88));if(converted)blob=converted;
    }finally{URL.revokeObjectURL(sourceUrl)}
  }catch(err){console.warn('圖片壓縮失敗，改用原圖。',err)}
  const safeArt=String(product.artNo||'ITEM').replace(/[^0-9A-Za-z._-]+/g,'_'),safeLot=String(product.lotNo||'').replace(/[^0-9A-Za-z._-]+/g,'_'),tag=source==='camera'?'CAMERA':'UPLOAD';
  return new File([blob],`${safeArt} MANUAL ${safeLot} ${tag}.JPG`,{type:'image/jpeg',lastModified:Date.now()});
}
function renderStockImageEditor(product){
  const dialog=$('#stockImageDialog'),grid=$('#stockImageChoices');if(!dialog||!grid||!product)return;
  $('#stockImageEditorTitle').textContent=`編輯圖片 · ${product.artNo}`;$('#stockImageEditorMeta').textContent=`LOTNO ${product.lotNo} · 最後選用圖片會寫入 jmsdata 的 ${IMAGE_OVERRIDE_SHEET} 工作表。`;
  grid.innerHTML='';const arr=state.imageFiles.get(product.artNo)||[],current=resolvedImageOverride(product);
  if(!arr.length)grid.innerHTML='<div class="notice warn">這個款號在 Pictures 內沒有可選圖片，可使用上傳圖片或即時拍照。</div>';
  for(const image of arr){
    const card=document.createElement('div');card.className='image-choice-card';if(current?.image?.fileName===image.fileName)card.classList.add('selected');
    const preview=document.createElement('img');preview.alt=image.fileName;setLazyImageElement(preview,image,placeholder(product.artNo));
    const name=document.createElement('div');name.className='image-choice-name';name.textContent=image.fileName;
    const buttons=document.createElement('div');buttons.className='image-choice-actions';
    const color=document.createElement('button');color.type='button';color.textContent='選用彩色';color.onclick=()=>{saveImageOverride(product,image,false);dialog.close()};
    const gray=document.createElement('button');gray.type='button';gray.className='ghost';gray.textContent='選用黑白';gray.onclick=()=>{saveImageOverride(product,image,true);dialog.close()};
    buttons.append(color,gray);card.append(preview,name,buttons);grid.appendChild(card);
  }
  const restore=$('#restoreAutoImageBtn');restore.disabled=!imageOverrideForLot(product.lotNo);restore.onclick=()=>{clearImageOverride(product);dialog.close()};
}
function openStockImageEditor(product){
  const dialog=$('#stockImageDialog');if(!dialog||!product)return;
  state.stockImageEditLot=String(product.lotNo);renderStockImageEditor(product);
  const statusEl=$('#stockImageSourceStatus');if(statusEl)statusEl.className='notice hidden';
  if(!dialog.open){if(typeof dialog.showModal==='function')dialog.showModal();else dialog.setAttribute('open','')}
}
async function handleImageEditorFile(file,source){
  const product=imageEditorProduct(),statusEl=$('#stockImageSourceStatus');if(!file||!product)return;
  if(statusEl){statusEl.textContent='正在處理圖片…';statusEl.className='notice'}
  try{const output=await compressManualImage(file,product,source),image=registerManualImage(product,output,source);renderStockImageEditor(product);if(statusEl){statusEl.textContent=`已加入 ${image.fileName}，請選擇彩色或黑白。`;statusEl.className='notice ok'}}catch(err){if(statusEl){statusEl.textContent=err.message||String(err);statusEl.className='notice error'}}
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
  if(!state.items.length){state.editingItemId=null;box.className='invoice-items empty-state';box.textContent='尚未加入貨品。';renderHistoricalGoldInputs();renderDocumentImageWarning();updateTotals();updateBackToTopButton();return}
  box.className='invoice-items';
  displayItems().forEach(item=>{
    const node=$('#itemTemplate').content.firstElementChild.cloneNode(true);node.dataset.itemId=String(item.id);
    $('.item-seq',node).textContent=item.seq;$('.item-artno',node).textContent=item.artNo;$('.item-lot',node).textContent=`LOTNO ${item.lotNo}`;
    const nonEmptyDescriptions=effectiveDescriptions(item).map(x=>norm(x)).filter(Boolean);
    $('.item-desc',node).textContent=nonEmptyDescriptions.slice(0,2).join('\n');
    $('.item-full-desc',node).textContent=nonEmptyDescriptions.join('\n');
    const usd=baseUsdPrice(item),code=currencyCode(),note=$('.item-price-note',node);
    if(quoteMode()){
      const q=quotationPriceDetails(item),mode=quoteModeName();
      if(q.ready){
        const adjustment=`${q.adjustment>=0?'+':'−'}${fmt(Math.abs(q.adjustment),'USD')}`,historyDate=q.historicalSourceDate&&q.historicalSourceDate!==q.date?`${q.date}→${q.historicalSourceDate}`:q.date;
        if(quote14KMode()){const weightText=quote14KSameWeightMode()?`${q.currentWeight.toFixed(2)}g（同金重）`:`${q.p.weight.toFixed(2)}g → ${q.currentWeight.toFixed(2)}g`;note.textContent=code==='USD'?`${mode}：${weightText} · 完成日18K ${historyDate} ${fmt(q.historicalTotal,'USD')} → 最新14K ${fmt(q.currentTotal,'USD')} · 調整 ${adjustment} → ${fmt(q.finalPrice,'USD')}`:(fxPricingReady()?`${mode}：USD ${q.finalPrice.toFixed(0)} × ${currentFxRate()} → ${fmt(item.unitPrice,code)}`:`${mode}：USD ${q.finalPrice.toFixed(0)} · 等待 FX Rate`)}
        else note.textContent=code==='USD'?`18K：完成日 ${historyDate} 公司金價 USD ${q.historical.base.toFixed(0)} → 最新 USD ${q.current.base.toFixed(0)} · 調整 ${adjustment} → ${fmt(q.finalPrice,'USD')}`:(fxPricingReady()?`18K：USD ${q.finalPrice.toFixed(0)} × ${currentFxRate()} → ${fmt(item.unitPrice,code)}`:`18K：USD ${q.finalPrice.toFixed(0)} · 等待 FX Rate`);
      }else note.textContent=`${mode} Quotation：等待 ${q.missing.join('／')}`;
      note.classList.add('quote14-price-note');if(quote14KMode()){const badge=document.createElement('span');badge.className='quote14-badge';badge.textContent=quote14KSameWeightMode()?'14K 同重':'14K';$('.item-artno',node).after(badge)}
    }else note.textContent=code==='USD'?`${item.price}u × ${Number($('#salesRate').value)||0} → ${fmt(usd,'USD')}`:(fxPricingReady()?`${item.price}u × ${Number($('#salesRate').value)||0} → ${fmt(usd,'USD')} × ${currentFxRate()} → ${fmt(item.unitPrice,code)}`:`${item.price}u × ${Number($('#salesRate').value)||0} → ${fmt(usd,'USD')} · 等待 FX Rate`);
    const manualFlag=$('.manual-price-flag',node);if(manualFlag)manualFlag.classList.toggle('hidden',!isManualPriceOverride(item));const thumbImg=getImg(item),thumb=$('.item-thumb',node);if(thumbImg?.file)setLazyImageElement(thumb,thumbImg,placeholder(item.artNo));else thumb.src=thumbImg?.url||placeholder(item.artNo);thumb.classList.toggle('grayscale-image',!!thumbImg?.grayscale);
    const controls=$('.item-controls',node),toggle=$('.item-edit-toggle',node);
    toggle.onclick=()=>{
      const isOpen=String(state.editingItemId??'')===String(item.id);
      setEditingItem(box,isOpen?null:item.id);
    };
    const deliveryBtn=$('.item-delivery-toggle',node);
    if(deliveryBtn){
      deliveryBtn.classList.toggle('hidden',state.documentType!=='invoice');
      if(state.documentType==='invoice'){deliveryBtn.textContent=item.delivered?'已交貨 ✓':'未交貨';deliveryBtn.classList.toggle('delivered',!!item.delivered);deliveryBtn.onclick=()=>setInvoiceItemDelivered(item,!item.delivered)}
    }
    const imageSelection=$('.item-image-selection',node),editImageBtn=$('.edit-item-image-btn',node),selectedImage=getImg(item);
    if(imageSelection){const source=imageOverrideForLot(item.lotNo)?'手動':item.imageAutoMatched?'自動':'文件選擇';imageSelection.textContent=`圖片：${selectedImage?.fileName||item.imageVariant||'沒有圖片'} · ${item.imageGrayscale?'黑白':'彩色'} · ${source}`}
    if(editImageBtn)editImageBtn.onclick=()=>openStockImageEditor(item);
    $('.qty-input',node).value=item.qty;$('.price-input',node).value=item.unitPrice;
    $('.qty-input',node).onchange=e=>{item.qty=Math.max(1,Number(e.target.value)||1);updateTotals();scheduleDraftSave()};
    $('.price-input',node).onchange=e=>{const code=currencyCode(),raw=Math.max(0,Number(e.target.value)||0),value=code==='EUR'?Math.round(raw):(code==='USD'&&quoteMode()?Math.ceil(raw-1e-9):roundUnitPrice(raw,code)),overrides=activePriceOverrides(item);overrides[code]=value;markManualPriceOverride(item,code);item.unitPrice=value;if(code==='USD'&&!quoteMode())item.usdUnitPrice=value;e.target.value=value;const flag=$('.manual-price-flag',node);if(flag)flag.classList.remove('hidden');updateTotals();scheduleDraftSave()};
    $('.delete-item',node).onclick=()=>{const lot=String(item.lotNo),current=currentInventoryStatusForLot(lot);if(state.documentType==='invoice'&&state.recall&&current==='SOLD_DELIVERED'&&!state.deliveryReturns.has(lot))return alert(`${item.artNo} / LOTNO ${lot} 目前已交貨。\n請先按右邊「已交貨 ✓」切換成「未交貨」，確認貨品已退回後再刪除。`);if(confirm(`刪除 ${item.artNo}？`)){revokeCustomImage(item);if(String(state.editingItemId??'')===String(item.id))state.editingItemId=null;state.items=state.items.filter(x=>x.id!==item.id);normalizeItemSequence();renderItems()}};
    box.appendChild(node)
  });
  setEditingItem(box,state.editingItemId,{syncSort:false});
  installItemSorting(box);
  renderHistoricalGoldInputs();renderDocumentImageWarning();updateTotals();updateBackToTopButton()
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
      schedulePreview();scheduleDraftSave();
      syncSortAvailability(box);
    }
  });
  syncSortAvailability(box);
}
$('#clearInvoiceBtn').onclick=()=>{const blocked=deliveredRecallItemsNeedingReturn();if(blocked.length)return alert(`目前有 ${blocked.length} 款貨品在 jmsdata 已是 Sold - Delivered。\n請先逐款按「已交貨 ✓」改成「未交貨」，確認貨品已退回後，才可清空／取消 Invoice。`);if(confirm(`清空目前 ${documentLabels().short} 貨品？`)){releaseCustomImages();state.items=[];resetPreviewAddonOptions();renderItems();if(state.recall)scheduleDraftSave();else{state.deliveryReturns=new Set();clearDocumentDraft();markDraftBaseline()}}};
$('#markAllDeliveredBtn').onclick=markAllInvoiceDelivered;
function reprice(){const r=Number($('#salesRate').value)||0;state.items.forEach(x=>{x.usdUnitPrice=Math.ceil(x.price*r);x.currencyPrices={};x.quote18kCurrencyPrices={};x.quote14kCurrencyPrices={};x.quote14kSameWeightCurrencyPrices={};x.manualPriceFlags={}});syncEffectivePrices();renderItems()}$('#salesRate').onchange=reprice;$('#currency').onchange=handleCurrencyChange;$('#refreshFxBtn').onclick=()=>fetchReferenceFxRate();let fxInputTimer=null;$('#fxRate').oninput=e=>{clearTimeout(fxInputTimer);fxInputTimer=setTimeout(()=>{const n=Number(e.target.value);if(Number.isFinite(n)&&n>0){state.fx={rate:n,date:'',source:'manual',fetching:false};syncEffectivePrices({clearCurrentOverride:true});renderItems();renderCustomerSummary();setFxStatus('使用手動 FX Rate · EUR Unit Price 四捨五入至整數。','warn')}else{syncEffectivePrices({clearCurrentOverride:true});renderItems();setFxStatus('請輸入大於 0 的 FX Rate。','error')}},160)};$('#discountAmount').oninput=updateTotals;['invoiceNo','invoiceDate','shipmentMethod','customerCode','customerName','customerAddress','customerTerms','remark'].forEach(id=>$('#'+id)?.addEventListener('input',()=>{schedulePreview();scheduleDraftSave()}));
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
function updatePreviewScale(){
  const host=$('#invoiceDocument');if(!host)return;
  const shell=host.querySelector('.preview-continuous-shell'),page=shell?.querySelector('.preview-continuous');if(!shell||!page)return;
  const available=Math.max(260,host.clientWidth||794),scale=Math.min(1,available/794);
  page.style.transform=`scale(${scale})`;
  shell.style.height=`${Math.ceil(page.scrollHeight*scale)}px`;
}
function previewAddonFooterHtml(t,items){
  const addons=state.documentType==='invoice'?excelAddonOptionsFromUI():normalizeExcelAddonOptions(),blocks=[];
  const totalAmount=state.documentType==='invoice'?`<div class="preview-total-amount"><span>Total Amount :</span><strong>${esc(currencyWords(currencyCode()))} ${esc(numberToWords(t.total))}</strong></div>`:'';
  if(addons.paymentTerm){
    const lines=invoicePaymentTermLines(t.total,norm($('#invoiceDate')?.value));
    blocks.push(`<section class="preview-addon-block preview-payment-term"><div class="preview-addon-title">Payment Term :</div>${lines.map(x=>`<div class="preview-addon-line">${esc(x)}</div>`).join('')}</section>`);
  }
  const manualRemark=norm($('#remark')?.value);
  if(addons.remark){
    const weight=invoiceWeightSummary(items),rows=[];
    if(addons.goldWeight)rows.push(['TOTAL GOLD WEIGHT :',weight.goldGrams,'g',false]);
    if(addons.semiPrecious)rows.push(['TOTAL SEMI-PRECIOUS STONES WEIGHT :',weight.semiPreciousGrams,`g (${weight.semiPreciousCarats.toFixed(2)} CARATS)`,false]);
    if(addons.diamond)rows.push(['TOTAL DIAMOND WEIGHT :',weight.diamondGrams,`g (${weight.diamondCarats.toFixed(2)} CARATS)`,false]);
    if(addons.allStone)rows.push(['TOTAL STONES WEIGHT :',weight.allStoneGrams,`g (${weight.allStoneCarats.toFixed(2)} CARATS)`,false]);
    if(addons.grossWeight)rows.push(['TOTAL GROSS WEIGHT :',weight.goldGrams+weight.allStoneGrams,'g',true]);
    blocks.push(`<section class="preview-addon-block"><div class="preview-addon-title">Remarks :</div><div class="preview-weight-grid">${rows.map(([label,value,unit,gross])=>`<div class="preview-weight-label">${esc(label)}</div><div class="preview-weight-value${gross?' gross':''}">${Number(value||0).toFixed(2)}</div><div class="preview-weight-unit">${esc(unit)}</div>`).join('')}</div>${manualRemark?`<div class="preview-addon-line preview-manual-remark">${esc(manualRemark).replace(/\n/g,'<br>')}</div>`:''}</section>`);
  }else{
    blocks.push(`<section class="preview-addon-block preview-standard-remark"><div class="preview-addon-title">Remark :</div><div class="preview-addon-line">${esc(manualRemark).replace(/\n/g,'<br>')||'&nbsp;'}</div></section>`);
  }
  if(addons.declaration)blocks.push(`<section class="preview-addon-block preview-declaration"><div class="preview-addon-line">WE, UNIVERSE GEMS &amp; JEWELLERY COMPANY, HEREBY CONFIRM THAT ALL DIAMONDS AND</div><div class="preview-addon-line">SEMI-PRECIOUS STONES ARE NATURAL.</div></section>`);
  if(addons.stoneDescription){
    const pairs=documentStoneDescriptionPairs(items);
    blocks.push(`<section class="preview-addon-block"><div class="preview-addon-title">Stone Decsription :</div><div class="preview-stone-grid">${pairs.map(([left,right])=>`<div>${esc(left?.text||'')}</div><div>${esc(right?.text||'')}</div>`).join('')}</div></section>`);
  }
  return `${totalAmount}${blocks.join('')}`;
}
function renderPreview(){
  const t=totals(),items=formalItems();
  const banker=`<div class="doc-meta preview-banker"><strong>Vendor's Banker</strong><br>The Hong Kong &amp; Shanghai Banking Corporation Ltd.<br>Address : 41 Ma Tau Wai Road, Hung Hom, Kowloon, Hong Kong<br>A/C # : 012-593570-001<br>A/C Name : Universe Gems &amp; Jewellery Co.</div>`;
  const rows=items.map((x,index)=>{
    const selected=getImg(x),img=selected?.file?ensureImageObjectUrl(selected):(selected?.url||placeholder('No Image')),gray=selected?.grayscale?' grayscale-image':'';
    const descriptions=[articleDescriptionFor(x),...effectiveDescriptions(x)].filter(Boolean).slice(0,4).map(esc).join('<br>');
    return `<tr class="preview-item-row"><td>${index+1}</td><td>Lot.No. : ${esc(x.lotNo)}<br>${esc(x.artNo)}</td><td>${descriptions}</td><td class="preview-picture"><img class="${gray.trim()}" src="${esc(img)}" alt="${esc(x.artNo)}"></td><td class="qty-cell">${x.qty}</td><td class="unit-cell">${esc(x.unit)}</td><td class="num">${fmt(x.unitPrice)}</td><td class="num">${fmt(x.qty*x.unitPrice)}</td></tr>`;
  }).join('');
  const addonFooter=previewAddonFooterHtml(t,items);
  const footer=`<div class="doc-footer preview-final-footer"><div class="doc-totals"><div><span>Total Quantity :</span><strong>${t.qty}</strong></div><div><span>Sub Total :</span><strong>${fmt(t.sub)}</strong></div><div><span>Discount :</span><strong>${discountDisplay(t.discount)}</strong></div><div class="total"><span>Total : (${esc(currencyCode())})</span><strong>${fmt(t.total)}</strong></div></div><div class="preview-addon-footer">${addonFooter}</div><div class="signature-row"><span>Vender Signature : ____________________</span><span>Accept By : ____________________</span></div></div>`;
  const html=`<div class="preview-continuous-shell"><section class="preview-continuous"><div class="letterhead"><img class="preview-company-logo" src="./company-logo.png" alt="Universe Gems &amp; Jewellery Company"></div><div class="doc-title">${documentLabels().title}</div><div class="doc-grid screen-preview"><div class="doc-meta">No. : <strong>${esc($('#invoiceNo').value)}</strong><br>${documentLabels().date} : ${esc(englishInvoiceDate($('#invoiceDate').value))}<br>Shipment Method : ${esc($('#shipmentMethod').value)}<br>Currency : ${esc($('#currency').value)}<br><br>Customer : <strong>${esc($('#customerName').value)}</strong><br>${esc($('#customerAddress').value).replace(/\n/g,'<br>')}</div>${banker}</div><table class="doc-table"><colgroup><col class="col-no"><col class="col-article"><col class="col-desc"><col class="col-picture"><col class="col-qty"><col class="col-unit"><col class="col-price"><col class="col-amount"></colgroup><thead><tr><th>No.</th><th>Article No.</th><th>Description</th><th>Picture</th><th class="qty-head">Quantity</th><th class="unit-head">Unit</th><th class="num">Unit Price</th><th class="num amount-head"><span>Amount</span><small>F.O.B. Value</small></th></tr></thead><tbody>${rows}</tbody></table>${footer}</section></div>`;
  $('#invoiceDocument').innerHTML=html;
  syncPreviewAddonControls();
  requestAnimationFrame(updatePreviewScale);
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
function excelColLetter(no){let n=Math.max(1,Number(no)||1),out='';while(n>0){n--;out=String.fromCharCode(65+(n%26))+out;n=Math.floor(n/26)}return out}
function excelColPixels(ws,colNo){const width=Number(ws.getColumn(colNo).width)||8.43;return Math.max(12,Math.round(width*7+5))}
function excelRowPixels(ws,rowNo){const points=Number(ws.getRow(rowNo).height)||15;return Math.max(8,points*96/72)}
function imageAnchorCol(ws,startColNo,endColNo,offsetPx){let col=startColNo-1,remain=Math.max(0,offsetPx);for(let c=startColNo;c<=endColNo;c++){const px=excelColPixels(ws,c);if(remain<=px)return col+remain/px;remain-=px;col+=1}return endColNo}
function imageAnchorRow(ws,startRow,endRow,offsetPx){let row=startRow-1,remain=Math.max(0,offsetPx);for(let r=startRow;r<=endRow;r++){const px=excelRowPixels(ws,r);if(remain<=px)return row+remain/px;remain-=px;row+=1}return endRow}
const EXCEL_PX_PER_CM=96/2.54;
const BL_EXCEL_IMAGE_WIDTH_CM=4;
function excelImageContainPlacement(ws,startColNo,endColNo,startRow,endRow,asset,padPx=1,targetWidthPx=0){
  let boxW=0,boxH=0;
  for(let c=startColNo;c<=endColNo;c++)boxW+=excelColPixels(ws,c);
  for(let r=startRow;r<=endRow;r++)boxH+=excelRowPixels(ws,r);
  const sourceW=Math.max(1,Number(asset?.width)||1),sourceH=Math.max(1,Number(asset?.height)||1);
  const pad=Math.max(0,Number(padPx)||0);
  const maxW=Math.max(1,boxW-pad*2),maxH=Math.max(1,boxH-pad*2);
  let scale=Math.min(maxW/sourceW,maxH/sourceH);
  // BL bracelet images are extremely wide and thin. Give them a requested
  // width of 4 cm, while still limiting the result to the available image box
  // and using one identical scale for width and height.
  const requestedWidth=Math.max(0,Number(targetWidthPx)||0);
  if(requestedWidth>0)scale=Math.min(requestedWidth/sourceW,maxW/sourceW,maxH/sourceH);
  const width=Math.max(1,sourceW*scale);
  const height=Math.max(1,sourceH*scale);
  const xOffset=Math.max(0,(boxW-width)/2),yOffset=Math.max(0,(boxH-height)/2);
  return {
    tl:{col:imageAnchorCol(ws,startColNo,endColNo,xOffset),row:imageAnchorRow(ws,startRow,endRow,yOffset)},
    ext:{width,height},
    editAs:'oneCell'
  };
}
function excelImagePlacementForItem(ws,startColNo,endColNo,startRow,endRow,asset,item,padPx=1){
  const targetWidth=articleType(item?.artNo)==='BL'?BL_EXCEL_IMAGE_WIDTH_CM*EXCEL_PX_PER_CM:0;
  return excelImageContainPlacement(ws,startColNo,endColNo,startRow,endRow,asset,padPx,targetWidth);
}
function rowRangeHeightPoints(ws,start,end){let total=0;for(let r=start;r<=end;r++)total+=Number(ws.getRow(r).height)||15;return total}
function copyTemplateRowStyle(ws,sourceRow,targetRow){
  const src=ws.getRow(sourceRow),dst=ws.getRow(targetRow);dst.height=src.height;
  for(let c=1;c<=12;c++){
    const s=src.getCell(c),d=dst.getCell(c);
    d.style=cloneStyle(s.style);d.numFmt=s.numFmt;d.alignment=cloneStyle(s.alignment);d.border=cloneStyle(s.border);d.fill=cloneStyle(s.fill);d.font=cloneStyle(s.font);d.protection=cloneStyle(s.protection);
  }
}
async function exportInvoiceFromTemplate(exportOptions={}){
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
  const addons=normalizeExcelAddonOptions(exportOptions);
  const addPayment=state.documentType==='invoice'&&addons.paymentTerm;
  const addWeights=state.documentType==='invoice'&&addons.remark;
  const addDeclaration=state.documentType==='invoice'&&addons.declaration;
  const addStoneDescription=state.documentType==='invoice'&&addons.stoneDescription;
  const stoneDescriptionPairs=addStoneDescription?documentStoneDescriptionPairs(formalItems()):[];
  const selectedWeightCount=addWeights?[addons.goldWeight,addons.semiPrecious,addons.diamond,addons.allStone,addons.grossWeight].filter(Boolean).length:0;
  const findOriginalFooterLabelRow=(text)=>{const needle=String(text||'').toLowerCase();for(let r=footerBaseRow;r<=originalFooterEnd;r++)for(let c=1;c<=Math.max(9,ws.columnCount||9);c++){if(norm(ws.getRow(r).getCell(c).value).toLowerCase().includes(needle))return r}return 0};
  const originalRemarkRow=findOriginalFooterLabelRow('remark')||footerBaseRow+9;
  const originalSignatureRow=findOriginalFooterLabelRow('vender signature')||findOriginalFooterLabelRow('accept by')||originalRemarkRow+3;
  const extraPaymentRows=addPayment?7:0;
  const stoneRows=addStoneDescription?Math.max(1,stoneDescriptionPairs.length):0;
  const manualRemarkRows=addWeights&&norm($('#remark')?.value)?1:0;
  let postRemarkRows=0;
  if(addWeights)postRemarkRows+=selectedWeightCount+manualRemarkRows;
  if(addDeclaration||addStoneDescription)postRemarkRows+=1;
  if(addDeclaration){postRemarkRows+=2;postRemarkRows+=1}
  if(addStoneDescription)postRemarkRows+=1+stoneRows+1;
  if(addWeights&&!addDeclaration&&!addStoneDescription)postRemarkRows+=1;
  const builtInRowsAfterRemark=Math.max(0,originalSignatureRow-originalRemarkRow-1);
  const extraRemarkRows=Math.max(0,postRemarkRows-builtInRowsAfterRemark);
  const originalContentRows=Math.max(1,footerBaseRow-firstItemRow-1);
  const baseContentRows=Math.min(4,originalContentRows);
  const separatorSourceRow=firstItemRow+baseContentRows;
  const columnCount=Math.max(9,ws.columnCount||9);

  // Keep the normal approved widths unless this document contains a BL item.
  // BL exports widen the Picture area to at least 4 cm and rebalance the compact
  // numeric columns, while the page remains fitted to one A4 page wide.
  const hasWideBlImage=formalItems().some(item=>articleType(item.artNo)==='BL');
  const templateColumnWidths=hasWideBlImage
    ? {A:5,B:13.49609375,C:23.94921875,D:10.3,E:10.3,F:6.5,G:5,H:8.5,I:11.2}
    : {A:9.2890625,B:13.49609375,C:23.94921875,D:4.7890625,E:4.7890625,F:9.43359375,G:6.53125,H:9.72265625,I:11.90234375};
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
  const requiredEnd=footerStart+footerRows.length+extraPaymentRows+extraRemarkRows-1;
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
  const footerReservedRows=16+extraPaymentRows+extraRemarkRows;
  const finalItemRowsWithFooter=fullItemRowsPerPage-footerReservedRows;
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
        ws.addImage(imageId,excelImagePlacementForItem(ws,imageStartColNo,imageEndColNo,start,imageEndRow,asset,item,1));
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

  {
    const blankSourceIndex=Math.max(0,Math.min(footerRows.length-1,originalSignatureRow-footerBaseRow-1));
    const blankCaptured=footerRows[blankSourceIndex];
    let targetOffset=0;
    for(let offset=0;offset<footerRows.length;offset++){
      const sourceRowNo=footerBaseRow+offset;
      if(sourceRowNo===originalRemarkRow&&extraPaymentRows>0){
        for(let k=0;k<extraPaymentRows;k++){
          const targetRow=footerStart+targetOffset++,row=ws.getRow(targetRow);row.height=blankCaptured?.height;
          for(let c=1;c<=columnCount;c++)applyCaptured(row.getCell(c),blankCaptured.row[c-1],false);
        }
      }
      if(sourceRowNo===originalSignatureRow&&extraRemarkRows>0){
        for(let k=0;k<extraRemarkRows;k++){
          const targetRow=footerStart+targetOffset++,row=ws.getRow(targetRow);row.height=blankCaptured?.height;
          for(let c=1;c<=columnCount;c++)applyCaptured(row.getCell(c),blankCaptured.row[c-1],false);
        }
      }
      const targetRow=footerStart+targetOffset++,captured=footerRows[offset],row=ws.getRow(targetRow);
      row.height=captured.height;
      for(let c=1;c<=columnCount;c++)applyCaptured(row.getCell(c),captured.row[c-1],true);
    }
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
  // Customer-specific footer text inherits the current Invoice Master Template style.
  // B on the original Remark row is used as the label style; C as the body style.
  const remarkTemplateOffset=Math.max(0,originalRemarkRow-footerBaseRow);
  const remarkLabelTemplateCaptured=footerRows[remarkTemplateOffset]?.row?.[Math.min(1,columnCount-1)]||null;
  const remarkTemplateCaptured=footerRows[remarkTemplateOffset]?.row?.[Math.min(2,columnCount-1)]||null;
  const remarkTemplateHeight=Number(footerRows[remarkTemplateOffset]?.height)||Number(ws.getRow(footerStart+remarkTemplateOffset).height)||15;
  const addonRowHeight=10.5;

  // Keep the existing approved sheet alignment behaviour, then re-apply Template styles
  // to any optional Invoice footer cells below.
  for(let r=1;r<=requiredEnd;r++)for(let c=1;c<=columnCount;c++){
    const cell=ws.getRow(r).getCell(c);
    cell.alignment={...cloneStyle(cell.alignment),vertical:'middle',wrapText:false};
  }
  const mergeStyledText=(rowNo,startCol,endCol,value,styleRef=remarkTemplateCaptured)=>{
    const a=excelColLetter(startCol),b=excelColLetter(Math.min(columnCount,endCol));
    try{ws.unMergeCells(`${a}${rowNo}:${b}${rowNo}`)}catch{}
    if(startCol<=Math.min(columnCount,endCol))try{ws.mergeCells(`${a}${rowNo}:${b}${rowNo}`)}catch{}
    const cell=ws.getCell(`${a}${rowNo}`);if(styleRef)applyCaptured(cell,styleRef,false);cell.value=value||'';
    ws.getRow(rowNo).height=addonRowHeight;
    return cell;
  };
  if(remarkLabel){
    const manualRemark=norm($('#remark').value),remarkRow=remarkLabel.r;

    if(addPayment){
      const paymentLabelRow=remarkRow-7,paymentLines=invoicePaymentTermLines(t.total,norm($('#invoiceDate').value));
      const paymentLabelCell=ws.getCell(`B${paymentLabelRow}`);if(remarkLabelTemplateCaptured)applyCaptured(paymentLabelCell,remarkLabelTemplateCaptured,false);paymentLabelCell.value='Payment Term :';ws.getRow(paymentLabelRow).height=addonRowHeight;
      for(let i=0;i<paymentLines.length;i++)mergeStyledText(paymentLabelRow+1+i,2,9,paymentLines[i]);
      mergeStyledText(paymentLabelRow+6,2,9,'');
    }

    let cursor=remarkRow+1;
    if(addWeights){
      const labelCell=ws.getCell(`B${remarkRow}`);if(remarkLabelTemplateCaptured)applyCaptured(labelCell,remarkLabelTemplateCaptured,false);labelCell.value='Remarks :';ws.getRow(remarkRow).height=addonRowHeight;
      const weight=invoiceWeightSummary(formalItems());
      const setWeightRow=(rowNo,label,value,unit)=>{
        mergeStyledText(rowNo,2,3,label);
        const valueCell=ws.getCell(`D${rowNo}`);if(remarkTemplateCaptured)applyCaptured(valueCell,remarkTemplateCaptured,false);
        valueCell.value=Number(value)||0;valueCell.numFmt='0.00';
        valueCell.alignment={...cloneStyle(valueCell.alignment),horizontal:'center',vertical:'middle',wrapText:false};
        mergeStyledText(rowNo,5,9,unit);
      };
      if(addons.goldWeight){setWeightRow(cursor++,'TOTAL GOLD WEIGHT :',weight.goldGrams,'g')}
      if(addons.semiPrecious){setWeightRow(cursor++,'TOTAL SEMI-PRECIOUS STONES WEIGHT :',weight.semiPreciousGrams,`g (${weight.semiPreciousCarats.toFixed(2)} CARATS)`)}
      if(addons.diamond){setWeightRow(cursor++,'TOTAL DIAMOND WEIGHT :',weight.diamondGrams,`g (${weight.diamondCarats.toFixed(2)} CARATS)`)}
      if(addons.allStone){setWeightRow(cursor++,'TOTAL STONES WEIGHT :',weight.allStoneGrams,`g (${weight.allStoneCarats.toFixed(2)} CARATS)`)}
      if(addons.grossWeight){const grossRow=cursor++;setWeightRow(grossRow,'TOTAL GROSS WEIGHT :',weight.goldGrams+weight.allStoneGrams,'g');const grossCell=ws.getCell(`D${grossRow}`);grossCell.border={...cloneStyle(grossCell.border),top:{style:'thin'},bottom:{style:'double'}}}
      if(manualRemark){mergeStyledText(cursor++,2,9,'');mergeStyledText(cursor++,2,9,manualRemark);}
    }else{
      try{ws.unMergeCells(`C${remarkRow}:H${remarkRow}`)}catch{}
      try{ws.mergeCells(`C${remarkRow}:H${remarkRow}`)}catch{}
      const remarkCell=ws.getCell(`C${remarkRow}`);if(remarkTemplateCaptured)applyCaptured(remarkCell,remarkTemplateCaptured,false);remarkCell.value=manualRemark;
    }

    if(addDeclaration||addStoneDescription)mergeStyledText(cursor++,2,9,'');
    if(addDeclaration){
      mergeStyledText(cursor++,2,9,'WE, UNIVERSE GEMS & JEWELLERY COMPANY, HEREBY CONFIRM THAT ALL DIAMONDS AND');
      mergeStyledText(cursor++,2,9,'SEMI-PRECIOUS STONES ARE NATURAL.');
      mergeStyledText(cursor++,2,9,'');
    }
    if(addStoneDescription){
      const stoneLabelCell=ws.getCell(`B${cursor}`);if(remarkLabelTemplateCaptured)applyCaptured(stoneLabelCell,remarkLabelTemplateCaptured,false);stoneLabelCell.value='Stone Decsription :';ws.getRow(cursor).height=addonRowHeight;cursor++;
      for(let i=0;i<stoneRows;i++){
        const [left,right]=stoneDescriptionPairs[i]||[];
        mergeStyledText(cursor,2,3,left?.text||'');
        mergeStyledText(cursor,4,8,right?.text||'');
        cursor++;
      }
      mergeStyledText(cursor++,2,9,'');
    }else if(addWeights&&!addDeclaration){
      mergeStyledText(cursor++,2,9,'');
    }
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
  downloadBlob(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),excelDocumentFileName());
  setExcelExportStatus(`已依 Template Map 輸出 ${state.invoiceTemplateName}${missingImages?`；${missingImages} 款沒有圖片`:''}。`,'ok');
}

function safeExportFilePart(v){return norm(v).replace(/[\\/:*?"<>|]/g,' ').replace(/\s+/g,' ').trim()||'Customer'}
function currentExportRevision(){if(state.recall&&state.recall.type===state.documentType&&state.recall.documentNo===norm($('#invoiceNo').value))return Math.max(0,Number(state.recall.baseRevision)||0)+1;const latest=findLatestDocument(state.documentType,norm($('#invoiceNo').value));return latest?Math.max(0,numberValue(docValue(latest.header,'REVISION'))):0}
function excelDocumentFileName(){const labels=documentLabels(),no=norm($('#invoiceNo').value)||formatDocumentNo(),customer=safeExportFilePart($('#customerName').value),qty=totals().qty,rev=currentExportRevision(),suffix=rev>0?`(${rev})`:'';return `${labels.short}(${no})_${customer}(${qty})${suffix}.xlsx`}
async function exportInvoiceExcel(){
  if(!fxPricingReady())return alert(`請先取得或輸入 USD → ${currencyCode()} 的 FX Rate。`);
  if(!quote14KReady())return alert('Quotation 金價資料未完成。請輸入最新 London PM，並補齊所有貨品完成日的 London PM。');
  if(!state.items.length){alert(`${documentLabels().short} 沒有貨品。`);return}
  let exportOptions=normalizeExcelAddonOptions();
  if(state.documentType==='invoice'){
    exportOptions=excelAddonOptionsFromUI();
    if(exportOptions.remark&&!exportOptions.goldWeight&&!exportOptions.semiPrecious&&!exportOptions.diamond&&!exportOptions.allStone&&!exportOptions.grossWeight)return alert('Remark 已選取，請至少選擇一項重量：Gold Weight、Semi-Precious、Diamond、All Stone 或 Gross Weight。');
    if(excelAddonRequested(exportOptions)&&!state.invoiceTemplateBuffer)return alert('已選擇 Invoice 附加資料，請先匯入目前展覽使用的 Invoice Master Template.xlsx。');
    if(excelAddonNeedsStoneList(exportOptions)&&!state.dataMeta?.stone)return alert('所選附加資料需要 Stone List，請先匯入目前展覽使用的最新 Stone List。');
  }
  if(!confirmImageIssuesBeforeExport())return;
  if(typeof ExcelJS==='undefined'){setExcelExportStatus('Excel 輸出程式未載入，請連接網絡後重新開啟。','error');return}
  const btn=$('#exportExcelBtn');btn.disabled=true;setExcelExportStatus('正在建立 Excel Invoice…');
  try{
    if(state.invoiceTemplateBuffer){await exportInvoiceFromTemplate(exportOptions);return}
    const wb=new ExcelJS.Workbook();
    wb.creator='Universe Invoice PWA';wb.created=new Date();
    const ws=wb.addWorksheet(documentLabels().title,{pageSetup:{paperSize:9,orientation:'portrait',fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.25,right:.25,top:.35,bottom:.35,header:.15,footer:.15}}});
    ws.views=[{showGridLines:false}];
    const exportItems=formalItems();
    const hasWideBlImage=exportItems.some(item=>articleType(item.artNo)==='BL');
    ws.columns=hasWideBlImage?[
      {key:'no',width:5},{key:'article',width:15},{key:'description',width:31.9},{key:'image',width:21.3},
      {key:'qty',width:7.5},{key:'unit',width:7.5},{key:'unitPrice',width:14},{key:'amount',width:14}
    ]:[
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
          const asset=await imageFileToJpegAsset(selected.file,620,.82,!!selected.grayscale);
          const imageId=wb.addImage({base64:asset.base64,extension:'jpeg'});
          ws.addImage(imageId,excelImagePlacementForItem(ws,4,4,start,start+3,asset,item,1));
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
    const manualRemark=norm($('#remark').value);
    ws.mergeCells(`A${row}:B${row}`);const remarkLabelCell=ws.getCell(`A${row}`);remarkLabelCell.value='Remark :';remarkLabelCell.font={name:'Arial',size:8,bold:true};remarkLabelCell.alignment={vertical:'middle'};
    ws.mergeCells(`C${row}:H${row}`);const remarkCell=ws.getCell(`C${row}`);remarkCell.value=manualRemark;remarkCell.alignment={vertical:'middle',wrapText:true};remarkCell.font={name:'Arial',size:8};if(String(remarkCell.value||'').includes('\n'))ws.getRow(row).height=22;
    row+=1;merge(`A${row}:D${row}`,'Vender Signature : ______________________',10);merge(`E${row}:H${row}`,'Accept By : ______________________',10,false,'right');
    ws.headerFooter.oddFooter='&RPage &P of &N';
    ws.pageSetup.printArea=`A1:H${row}`;
    ws.autoFilter={from:{row:headerRow,column:1},to:{row:headerRow,column:8}};
    const buffer=await wb.xlsx.writeBuffer();
    const inv=norm($('#invoiceNo').value)||formatDocumentNo();
    downloadBlob(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),excelDocumentFileName());
    setExcelExportStatus(`Excel Invoice 已輸出${missingImages?`；${missingImages} 款沒有嵌入圖片`:''}。`,'ok');
  }catch(err){console.error(err);setExcelExportStatus('Excel 輸出失敗：'+(err.message||err),'error')}
  finally{btn.disabled=false}
}

$('#exportExcelBtn').onclick=exportInvoiceExcel;
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
window.addEventListener('resize',()=>requestAnimationFrame(updatePreviewScale));
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
  if(!state.stockAllRows.length)return alert('請先匯入 jmsdata.xlsx。');
  if(state.items.length&&!fxPricingReady())return alert(`請先取得或輸入 USD → ${currencyCode()} 的 FX Rate。`);
  if(state.items.length&&!quote14KReady())return alert('Quotation 金價資料未完成。請輸入最新 London PM，並補齊所有貨品完成日的 London PM。');
  const keys=storeKeysForType(type),isQuote=type==='quotation',targetStatus=type==='consignment'?'CONSIGNED':'SOLD_ON_HAND';
  if(!recall&&findLatestDocument(type,documentNo))return alert(`${documentNo} 已存在。請使用 Recall 重新開啟原文件，不要重複 Confirm 同一文件號碼。`);
  const revision=recall?recall.baseRevision+1:0,oldItems=recall?new Map(recall.items.map(x=>[String(docValue(x,'LOTNO')),x])):new Map(),newItems=formalItems(),newLots=new Set(newItems.map(x=>String(x.lotNo))),rollback=new Map(),documentStatus=newItems.length?'CONFIRMED':'CANCELLED';
  try{
    if(recall&&!isQuote){
      for(const [lot,oldRow] of oldItems){
        if(newLots.has(lot))continue;
        const p=state.stockCatalog.get(lot)||{lotNo:lot,artNo:normArt(docValue(oldRow,'ARTNO'))},before=inventoryVectorFromRow(state.stockRowByLot.get(lot)||{}),original=vectorFromRecord(oldRow,'ORIGINAL'),currentStatus=inventoryStatusFromVector(before);
        if(type==='invoice'){
          if(!['SOLD_ON_HAND','SOLD_DELIVERED'].includes(currentStatus))throw new Error(`${p.artNo} / LOTNO ${lot} 目前狀態是 ${historyStatusLabel(currentStatus)}，已不是這張 Invoice 的銷售狀態，為避免覆蓋其他庫存變更不能自動回復。`);
          if(currentStatus==='SOLD_DELIVERED'&&!state.deliveryReturns.has(lot))throw new Error(`${p.artNo} / LOTNO ${lot} 目前已是 Sold - Delivered。請返回 Invoice，先將該貨品由「已交貨 ✓」改成「未交貨」，確認貨品已退回後再刪除／取消文件。`);
        }else if(currentStatus!==targetStatus)throw new Error(`${p.artNo} / LOTNO ${lot} 目前狀態是 ${historyStatusLabel(currentStatus)}，已不是這張文件原本的 ${historyStatusLabel(targetStatus)}，不能由舊文件 Recall 覆蓋。`);
        if(!rollback.has(lot))rollback.set(lot,before);setInventoryVectorForLot(lot,original);appendTransaction({type,documentNo,revision,item:p,before,after:original});
      }
    }
    const itemRows=[];
    for(const item of newItems){
      const lot=String(item.lotNo),row=state.stockRowByLot.get(lot);if(!row)throw new Error(`jmsdata 找不到 LOTNO ${lot}`);
      const before=inventoryVectorFromRow(row),oldRow=oldItems.get(lot),isExisting=!!oldRow;let original=isExisting?vectorFromRecord(oldRow,'ORIGINAL'):(item._originalVector||before),after=before;
      if(!isQuote){
        const currentStatus=inventoryStatusFromVector(before);
        if(type==='invoice'){
          const desiredStatus=item.delivered?'SOLD_DELIVERED':'SOLD_ON_HAND';
          if(isExisting){
            if(!['SOLD_ON_HAND','SOLD_DELIVERED'].includes(currentStatus))throw new Error(`${item.artNo} / LOTNO ${lot} 的庫存狀態已在這張文件之後改為 ${historyStatusLabel(currentStatus)}，不能由舊 Invoice Recall 覆蓋。`);
            if(currentStatus==='SOLD_DELIVERED'&&desiredStatus==='SOLD_ON_HAND'&&!state.deliveryReturns.has(lot))throw new Error(`${item.artNo} / LOTNO ${lot} 目前已是 Sold - Delivered。請先在本 Invoice 按「已交貨 ✓」改為「未交貨」，確認貨品已退回後再 Confirm。`);
            after=inventoryVectorForStatus(desiredStatus,before.m);
          }else{
            if(currentStatus!=='AVAILABLE')throw new Error(`${item.artNo} / LOTNO ${lot} 目前不是 Available，不能新增到 Invoice。`);
            after=inventoryVectorForStatus(desiredStatus,before.m);
          }
        }else{
          if(isExisting){if(currentStatus===targetStatus)after=inventoryVectorForStatus(targetStatus,before.m);else throw new Error(`${item.artNo} / LOTNO ${lot} 的庫存狀態已在這張文件之後改為 ${historyStatusLabel(currentStatus)}，不能由舊文件 Recall 覆蓋。`)}
          else{if(currentStatus!=='AVAILABLE')throw new Error(`${item.artNo} / LOTNO ${lot} 目前不是 Available，不能新增到 ${documentLabels(type).short}。`);after=inventoryVectorForStatus(targetStatus,before.m)}
        }
        if(!rollback.has(lot))rollback.set(lot,before);setInventoryVectorForLot(lot,after);appendTransaction({type,documentNo,revision,item,before,after});
      }
      itemRows.push(itemRowForDocument(item,documentNo,revision,original,after));
    }
    const previousCreated=recall?docValue(recall.header,'CREATED_AT'):'';state.documentStore[keys.headers].push(headerRowForCurrentDocument(type,documentNo,revision,previousCreated,documentStatus));state.documentStore[keys.items].push(...itemRows);state.recordsLoaded=true;state.recordsDirty=true;if(!isQuote)rebuildInventoryMaps();const wasRecall=!!recall;if(!wasRecall)advanceDocumentSequence(documentNo,type);if(state.dataMeta.records)state.dataMeta.records.count=officialDocumentCount();const stockExported=exportJmsdata();
    state.recall=null;state.deliveryReturns=new Set();setRecallLock(false);$('#recallActive')?.classList.add('hidden');clearDocumentDraft();clearCurrentDocument();if(isQuote){state.quote.karat='18K';$$('input[name="quoteKarat"]').forEach(x=>x.checked=x.value==='18K');updateGoldQuoteUI()}if(wasRecall)setDefaultInvoiceNo(true);markDraftBaseline();renderRecallResults();updateDataVersionPanel();const action=documentStatus==='CANCELLED'?'已取消整張文件':revision?'已儲存修訂':'已 Confirm',inventoryNote=isQuote?'庫存不變；':' ';status('#addMessage',`${documentLabels(type).short} ${documentNo} ${action}（Revision ${revision}），${inventoryNote}${stockExported?'已輸出最新 jmsdata.xlsx（庫存＋Records）':'jmsdata 匯出失敗'}；目前文件號碼為 ${$('#invoiceNo').value}。`,stockExported?'ok':'warn');if(!isQuote)status('#stockStatus',`目前 Available Stock：${state.products.size} 件。`,'ok');
  }catch(err){console.error(err);for(const [lot,vector] of rollback){try{setInventoryVectorForLot(lot,vector)}catch{}}alert('Confirm 失敗：'+(err.message||err));if(!isQuote)rebuildInventoryMaps()}
}

function documentNoIsValid(no=norm($('#invoiceNo')?.value),type=state.documentType){return new RegExp(`^${documentPrefix(type)}\\d{2}\\d{4}$`,'i').test(no)}
function confirmPreflightError(){
  const no=norm($('#invoiceNo')?.value),rate=Number($('#salesRate')?.value),hasCustomer=!!(norm($('#customerCode')?.value)&&norm($('#customerName')?.value));
  if(!hasCustomer)return'請先選擇／輸入有效客戶（Customer Code 及 Customer Name）。';
  if(!(Number.isFinite(rate)&&rate>0))return'Sales Rate 必須大於 0。';
  if(!documentNoIsValid(no))return`${documentLabels().short} No. 格式無效，應為 ${documentPrefix()}YY0001，例如 ${formatDocumentNo(1)}。`;
  if(!state.stockAllRows.length)return'請先匯入 jmsdata.xlsx。';
  const isCancel=!!state.recall&&!state.items.length;if(!state.items.length&&!isCancel)return`${documentLabels().short} 沒有貨品。`;
  if(state.documentType==='invoice'&&state.recall){const newLots=new Set(state.items.map(x=>String(x.lotNo))),blocked=state.recall.items.filter(row=>{const lot=String(docValue(row,'LOTNO'));return !newLots.has(lot)&&currentInventoryStatusForLot(lot)==='SOLD_DELIVERED'&&!state.deliveryReturns.has(lot)});if(blocked.length)return`有 ${blocked.length} 款已交貨貨品被刪除／取消，但尚未在本 Invoice 確認「未交貨」。請返回修改，先確認貨品已退回。`;}
  if(state.items.length&&!fxPricingReady())return`請先取得或輸入 USD → ${currencyCode()} 的 FX Rate。`;
  if(state.items.length&&!quote14KReady())return'Quotation 金價資料未完成。請更新 GoldSilver.xlsx 或輸入最新 London PM。';
  return'';
}
function openFinalConfirmDialog(){
  const error=confirmPreflightError();if(error)return alert(error);
  const t=totals(),img=currentDocumentImageStats(),manual=state.items.filter(x=>isManualPriceOverride(x)).length,isCancel=!!state.recall&&!state.items.length,dialog=$('#finalConfirmDialog'),box=$('#finalConfirmSummary');
  const rows=[['Document No.',norm($('#invoiceNo').value)],['Customer',`${norm($('#customerCode').value)} · ${norm($('#customerName').value)}`],['Sales Rate',norm($('#salesRate').value)],['Currency',currencyCode()],['Quantity',String(t.qty)],['Sub Total',fmt(t.sub)],['Discount',t.discount?`(${fmt(t.discount)})`:fmt(0)],['Total',fmt(t.total)],['圖片待處理',String(img.issues)],['人手改價',String(manual)]];if(state.documentType==='invoice'){const delivered=state.items.filter(x=>x.delivered).length;rows.push(['已交貨',String(delivered)],['未交貨',String(state.items.length-delivered)])}
  box.innerHTML=`<div class="final-confirm-grid">${rows.map(([k,v])=>`<div><span>${esc(k)}</span><strong>${esc(v)}</strong></div>`).join('')}</div>${img.issues?`<div class="notice warn">有 ${img.issues} 款圖片待處理（黑白 ${img.fallback}／無圖 ${img.missing}）。可繼續 Confirm，但請留意交客文件。</div>`:''}${manual?`<div class="notice warn">有 ${manual} 款 Unit Price 曾經人手修改。</div>`:''}${isCancel?'<div class="notice warn">目前已刪除全部貨品；Confirm 後會將這張 Recall 文件標記為 Cancelled。</div>':''}`;
  $('#finalConfirmBtn').textContent=isCancel?'確定取消文件':'確定 Confirm';
  if(dialog?.showModal)dialog.showModal();else dialog?.setAttribute('open','');
}

function exportCurrentStockAfterConfirm(){return confirmStoredDocument()}
function exportRemaining(){exportCurrentStockAfterConfirm()}
$('#confirmInvoiceBtn').onclick=openFinalConfirmDialog;$('#closeFinalConfirmBtn').onclick=()=>$('#finalConfirmDialog').close();$('#cancelFinalConfirmBtn').onclick=()=>$('#finalConfirmDialog').close();$('#finalConfirmBtn').onclick=()=>{$('#finalConfirmDialog').close();exportCurrentStockAfterConfirm()};

const STOCK_SEARCH_INITIAL_BATCH=42,STOCK_SEARCH_BATCH_SIZE=54;
const stockSearchPerf={dataRevision:0,imageRevision:0,recordsRevision:-1,records:[],renderKey:'',shown:[],renderedCount:0,sentinelObserver:null,imageMeta:new Map(),issueScanToken:0,scrollY:0,baseMessage:''};
const stockDerivedCache={breakdownCodes:null,stoneOrder:null,stoneOrderRank:null,colorGroups:null,colorGroupRank:null,stoneCodes:new Map(),signatures:new Map()};
function disconnectStockSearchSentinel(){if(stockSearchPerf.sentinelObserver){try{stockSearchPerf.sentinelObserver.disconnect()}catch{}stockSearchPerf.sentinelObserver=null}}
function invalidateStockSearchView(){stockSearchPerf.renderKey='';disconnectStockSearchSentinel();stockSearchPerf.issueScanToken++}
function invalidateStockSearchData(){stockSearchPerf.dataRevision++;stockSearchPerf.recordsRevision=-1;stockSearchPerf.records=[];invalidateStockSearchView()}
function invalidateStockSearchImages(lot=''){stockSearchPerf.imageRevision++;if(lot)stockSearchPerf.imageMeta.delete(String(lot));else stockSearchPerf.imageMeta.clear();invalidateStockSearchView()}
function resetStockStoneDerivedCache(){stockDerivedCache.breakdownCodes=null;stockDerivedCache.stoneOrder=null;stockDerivedCache.stoneOrderRank=null;stockDerivedCache.colorGroups=null;stockDerivedCache.colorGroupRank=null;stockDerivedCache.stoneCodes.clear();stockDerivedCache.signatures.clear();invalidateStockSearchData();invalidateStockSearchImages()}
function stockSearchViewKey(){const f=state.stockSearch||{},items=state.items.map(x=>String(x.lotNo)).sort().join(','),arr=v=>(Array.isArray(v)?v:[]).slice().sort().join(',');return [stockSearchPerf.dataRevision,stockSearchPerf.imageRevision,state.documentType,f.query||'*',arr(f.types),arr(f.stones),arr(f.statuses),f.imageIssuesOnly?1:0,f.filtersOpen?1:0,items].join('|')}
function ensureStockSearchRendered(){const resultsEl=$('#stockSearchResults');if(stockSearchPerf.renderKey===stockSearchViewKey()&&resultsEl&&resultsEl.childElementCount)return;renderStockSearch()}
function stockRecordCache(){if(stockSearchPerf.recordsRevision===stockSearchPerf.dataRevision)return stockSearchPerf.records;stockSearchPerf.records=inventorySearchRecordsUncached();stockSearchPerf.recordsRevision=stockSearchPerf.dataRevision;return stockSearchPerf.records}
function stockImageMeta(product){const lot=String(product?.lotNo||''),cached=stockSearchPerf.imageMeta.get(lot);if(cached&&cached.revision===stockSearchPerf.imageRevision)return cached;const display=imageDisplayForProduct(product),issue=!display.image||(!display.manual&&!!display.grayscale),meta={revision:stockSearchPerf.imageRevision,display,issue};stockSearchPerf.imageMeta.set(lot,meta);return meta}
function scheduleStockIssueCount(records){const token=++stockSearchPerf.issueScanToken,button=$('#stockSearchSummary [data-image-issues]');if(button)button.textContent='圖片待處理 …';let i=0,count=0;const step=deadline=>{if(token!==stockSearchPerf.issueScanToken)return;let n=0;while(i<records.length&&(n<36||deadline?.timeRemaining?.()>3)){if(stockImageMeta(records[i]).issue)count++;i++;n++}if(i<records.length){if(typeof requestIdleCallback==='function')requestIdleCallback(step,{timeout:90});else setTimeout(()=>step(null),0);return}const current=$('#stockSearchSummary [data-image-issues]');if(current)current.textContent=`圖片待處理 ${count}`};if(typeof requestIdleCallback==='function')requestIdleCallback(step,{timeout:90});else setTimeout(()=>step(null),0)}
function createStockResultCard(x){const card=document.createElement('div');card.className='stock-result';card.dataset.lot=String(x.lotNo);const meta=stockImageMeta(x),display=meta.display,grayClass=display.grayscale?'grayscale-image':'',inDoc=state.items.some(i=>String(i.lotNo)===String(x.lotNo)),canAdd=x.status==='AVAILABLE'||state.documentType==='quotation'&&['CONSIGNED','SOLD_ON_HAND','SOLD_DELIVERED'].includes(x.status);const imageBadge=display.manual?'<span class="stock-image-badge manual">手動圖片</span>':meta.issue?'<span class="stock-image-badge attention">圖片待處理</span>':'';const rawPrice=Number(x.price),uPrice=Number.isFinite(rawPrice)?new Intl.NumberFormat('en-US',{useGrouping:false,maximumFractionDigits:2}).format(rawPrice)+'u':'';card.innerHTML=`<div class="stock-result-image-wrap"><div class="stock-result-thumb-box"><img class="${grayClass}" src="${esc(display.src)}" alt="${esc(x.artNo)}">${imageBadge}</div></div><div><h4>${esc(x.artNo)}</h4><div class="lot">LOTNO ${esc(x.lotNo)}${uPrice?` · <span class="stock-u-price">${esc(uPrice)}</span>`:''}</div><div class="desc">${esc((x.descriptions||[]).join('\n'))}</div></div><div class="stock-result-actions"><span class="stock-status ${historyStatusClass(x.status)}">${esc(historyStatusLabel(x.status))}</span>${inDoc?'<span class="stock-current-doc">已在目前文件</span>':''}</div>`;const cardImg=$('img',card);if(cardImg&&display.image)setLazyImageElement(cardImg,display.image,display.src);const imageWrap=$('.stock-result-image-wrap',card),actions=$('.stock-result-actions',card);const edit=document.createElement('button');edit.type='button';edit.className='ghost stock-image-edit-btn';edit.textContent='編輯圖片';edit.onclick=()=>openStockImageEditor(x);imageWrap.appendChild(edit);const diagnostic=document.createElement('button');diagnostic.type='button';diagnostic.className='ghost';diagnostic.textContent='診斷';diagnostic.onclick=()=>openDiagnosticReport(x);actions.appendChild(diagnostic);if(canAdd&&!inDoc){const b=document.createElement('button');b.type='button';b.textContent=`加入目前 ${documentLabels().short}`;b.onclick=()=>{if(addProductToDocument(x,{fromSearch:true})){b.disabled=true;b.textContent='已加入';const tag=document.createElement('span');tag.className='stock-current-doc';tag.textContent='已在目前文件';actions.insertBefore(tag,diagnostic);stockSearchPerf.renderKey=stockSearchViewKey()}};actions.appendChild(b)}return card}
function updateStockBatchSentinel(){const resultsEl=$('#stockSearchResults');if(!resultsEl)return;resultsEl.querySelector('.stock-load-more')?.remove();disconnectStockSearchSentinel();if(stockSearchPerf.renderedCount>=stockSearchPerf.shown.length)return;const more=document.createElement('button');more.type='button';more.className='stock-load-more ghost';more.textContent=`繼續向下捲動載入更多 · ${stockSearchPerf.renderedCount} / ${stockSearchPerf.shown.length}`;more.onclick=()=>appendStockSearchBatch();resultsEl.appendChild(more);if(typeof IntersectionObserver!=='undefined'){stockSearchPerf.sentinelObserver=new IntersectionObserver(entries=>{if(entries.some(x=>x.isIntersecting))appendStockSearchBatch()},{rootMargin:'900px 0px'});stockSearchPerf.sentinelObserver.observe(more)}}
function appendStockSearchBatch(count=STOCK_SEARCH_BATCH_SIZE){const resultsEl=$('#stockSearchResults');if(!resultsEl)return;const start=stockSearchPerf.renderedCount,end=Math.min(stockSearchPerf.shown.length,start+Math.max(1,count));if(start>=end){updateStockBatchSentinel();return}resultsEl.querySelector('.stock-load-more')?.remove();const frag=document.createDocumentFragment();for(let i=start;i<end;i++)frag.appendChild(createStockResultCard(stockSearchPerf.shown[i]));resultsEl.appendChild(frag);stockSearchPerf.renderedCount=end;updateStockBatchSentinel();requestAnimationFrame(updateBackToTopButton)}
function articleCore(value){const s=normArt(value),m=s.match(/(\d{3,})/);return m?m[1]:s.replace(/^[A-Z]+[-\s]*/,'').replace(/\.[A-Z0-9]+$/,'')}
function articleType(value){const m=normArt(value).match(/^([A-Z]+)/);return m?m[1]:'OTHER'}
function validStoneBreakdownCodes(){if(stockDerivedCache.breakdownCodes)return stockDerivedCache.breakdownCodes;stockDerivedCache.breakdownCodes=[...activeStoneAliases().keys()].map(x=>norm(x).toUpperCase().replace(/\s+/g,'')).filter(Boolean).sort((a,b)=>b.length-a.length||a.localeCompare(b));return stockDerivedCache.breakdownCodes}
function longestStoneBreakdownPrefix(value){const part=String(value||'').toUpperCase().replace(/[^A-Z0-9]/g,'');if(!part)return'';return validStoneBreakdownCodes().find(code=>part.startsWith(code))||''}
function stoneCodesFromDescriptionLine(line){
  const raw=String(line||'').toUpperCase(),firstDash=raw.indexOf('-');if(firstDash<0)return[];
  const tail=raw.slice(firstDash+1),nextDash=tail.indexOf('-'),stoneBlock=(nextDash>=0?tail.slice(0,nextDash):tail).replace(/\s+/g,'');
  if(!stoneBlock)return[];
  const out=[];
  for(const part of stoneBlock.split(/[+\/，,]/).filter(Boolean)){
    // Stone List BREAKDOWN uses the leading code; the remainder is shape/cutting.
    // Longest-prefix matching is essential for pairs such as BO/BOX and GT/GTQ.
    const code=longestStoneBreakdownPrefix(part);
    if(code&&!isDiamondStoneCode(code)&&!out.includes(code))out.push(code);
  }
  return out;
}
function stoneCodesForProduct(p){const key=`${String(p?.lotNo||p?.artNo||'')}|${(p?.descriptions||[]).join('\u0001')}`,cached=stockDerivedCache.stoneCodes.get(key);if(cached)return cached;const out=[];for(const line of (p?.descriptions||[])){for(const code of stoneCodesFromDescriptionLine(line)){if(!out.includes(code))out.push(code)}}stockDerivedCache.stoneCodes.set(key,out);return out}
const STOCK_TYPE_ORDER=['RG','ER','PT','BR','NL','BL','BG'];
function stockTypeRank(type){const t=norm(type).toUpperCase(),i=STOCK_TYPE_ORDER.indexOf(t);return i>=0?i:STOCK_TYPE_ORDER.length}
function compareStockTypes(a,b){const ra=stockTypeRank(a),rb=stockTypeRank(b);return ra-rb||String(a).localeCompare(String(b))}
function stoneOrderList(){if(stockDerivedCache.stoneOrder)return stockDerivedCache.stoneOrder;const out=[];for(const code of activeStoneAliases().keys()){const c=norm(code).toUpperCase();if(c&&!isDiamondStoneCode(c)&&!out.includes(c))out.push(c)}stockDerivedCache.stoneOrder=out;stockDerivedCache.stoneOrderRank=new Map(out.map((x,i)=>[x,i]));return out}
function stoneOrderRank(code){const c=norm(code).toUpperCase(),list=stoneOrderList(),i=stockDerivedCache.stoneOrderRank?.get(c);return Number.isInteger(i)?i:list.length+100}
function colorGroupOrderList(){if(stockDerivedCache.colorGroups)return stockDerivedCache.colorGroups;const out=[];for(const code of stoneOrderList()){const group=stoneGroupForCode(code);if(group&&!out.includes(group))out.push(group)}stockDerivedCache.colorGroups=out;stockDerivedCache.colorGroupRank=new Map(out.map((x,i)=>[norm(x).toUpperCase(),i]));return out}
function colorGroupRank(group){const g=norm(group).toUpperCase(),list=colorGroupOrderList(),i=stockDerivedCache.colorGroupRank?.get(g);return Number.isInteger(i)?i:list.length+100}
function stockSortStoneCodes(p,preferredStones=[]){
  const codes=stoneCodesForProduct(p),preferred=(Array.isArray(preferredStones)?preferredStones:[]).map(x=>norm(x).toUpperCase());
  const matched=preferred.length?codes.filter(code=>preferred.includes(norm(code).toUpperCase())):codes;
  const source=matched.length?matched:codes;
  return [...source].sort((a,b)=>colorGroupRank(stoneGroupForCode(a))-colorGroupRank(stoneGroupForCode(b))||stoneOrderRank(a)-stoneOrderRank(b)||String(a).localeCompare(String(b)));
}
function primaryStoneCodeForProduct(p,preferredStones=[]){return stockSortStoneCodes(p,preferredStones)[0]||''}
function stockStoneSignatureCodes(p){const key=`${String(p?.lotNo||p?.artNo||'')}|${(p?.descriptions||[]).join('\u0001')}`,cached=stockDerivedCache.signatures.get(key);if(cached)return cached;const out=stockSortStoneCodes(p,[]);stockDerivedCache.signatures.set(key,out);return out}
function compareStockStoneSignatures(a,b){
  const aa=stockStoneSignatureCodes(a),bb=stockStoneSignatureCodes(b),n=Math.min(aa.length,bb.length);
  for(let i=0;i<n;i++){
    const ga=stoneGroupForCode(aa[i]),gb=stoneGroupForCode(bb[i]),groupDiff=colorGroupRank(ga)-colorGroupRank(gb);if(groupDiff)return groupDiff;
    const stoneDiff=stoneOrderRank(aa[i])-stoneOrderRank(bb[i]);if(stoneDiff)return stoneDiff;
    const alpha=String(aa[i]).localeCompare(String(bb[i]));if(alpha)return alpha;
  }
  return aa.length-bb.length;
}
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
  // Within the same colour group, keep the complete stone combination together
  // before splitting by RG / ER / PT etc. Example: PTQ stays together, then PAM+PTQ.
  const signatureDiff=compareStockStoneSignatures(a,b);if(signatureDiff)return signatureDiff;
  const ta=articleType(a.artNo),tb=articleType(b.artNo),typeDiff=compareStockTypes(ta,tb);if(typeDiff)return typeDiff;
  return normArt(a.artNo).localeCompare(normArt(b.artNo))||String(a.lotNo).localeCompare(String(b.lotNo));
}
function compareWildcardStock(a,b){return compareColorSortedStock(a,b,[])}
function inventorySearchRecordsUncached(){const map=new Map();for(const p of state.stockCatalog.values())map.set(String(p.lotNo),{...productSnapshot(p),status:'AVAILABLE'});for(const p of state.products.values())map.set(String(p.lotNo),{...productSnapshot(p),status:'AVAILABLE'});for(const h of state.inventoryHistory.values())map.set(String(h.lotNo),{...h,status:h.status||'AVAILABLE'});return [...map.values()]}
function inventorySearchRecords(){return stockRecordCache()}
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
function stockFilterSummaryText(){const t=selectedStockFilters('type'),s=selectedStockFilters('stone'),image=state.stockSearch.imageIssuesOnly?' · 圖片：待處理':'';return `款式：${t.length?t.join(', '):'全部'} · 石頭：${s.length?s.map(stoneFilterLabel).join(', '):'全部'}${image}`}
function runStockSearch(){const raw=norm($('#stockSearchInput').value).toUpperCase(),core=!raw||raw==='*'?'*':articleCore(raw);state.stockSearch.query=core||raw||'*';state.stockSearch.types=[];state.stockSearch.stones=[];state.stockSearch.statuses=[];state.stockSearch.imageIssuesOnly=false;state.stockSearch.filtersOpen=false;renderStockSearch()}
function renderStockSearch(){
  const resultsEl=$('#stockSearchResults');if(!resultsEl)return;
  disconnectStockSearchSentinel();stockSearchPerf.issueScanToken++;
  const q=norm(state.stockSearch.query||$('#stockSearchInput')?.value||'*').toUpperCase(),filterControls=$('#stockFilterControls'),filterPanel=$('#stockFilterPanel'),records=inventorySearchRecords();
  if(!records.length){resultsEl.innerHTML='';$('#stockSearchMessage').textContent='尚未匯入 jmsdata。';$('#stockTypeFilters').classList.add('hidden');$('#stockStoneFilters').classList.add('hidden');$('#stockSearchSummary').classList.add('hidden');filterControls?.classList.add('hidden');filterPanel?.classList.add('hidden');stockSearchPerf.renderKey=stockSearchViewKey();return}
  const wildcard=!q||q==='*',core=wildcard?'*':articleCore(q),all=wildcard?records:records.filter(x=>articleCore(x.artNo)===core||normArt(x.artNo).includes(normArt(q)));
  const types=[...new Set(all.map(x=>articleType(x.artNo)))].sort(compareStockTypes),stones=[...new Set(all.flatMap(stoneCodesForProduct))].sort((a,b)=>stoneOrderRank(a)-stoneOrderRank(b)||a.localeCompare(b));
  const selectedTypes=selectedStockFilters('type'),selectedStones=selectedStockFilters('stone'),selectedStatuses=selectedStockFilters('status');
  const typeBox=$('#stockTypeFilters'),stoneBox=$('#stockStoneFilters');typeBox.dataset.kind='type';stoneBox.dataset.kind='stone';
  typeBox.innerHTML=`<div class="filter-chips">${filterButtonHTML('ALL','全部',selectedTypes.length===0)}${types.map(x=>filterButtonHTML(x,x,selectedTypes.includes(x))).join('')}</div>`;
  stoneBox.innerHTML=`<div class="filter-chips">${filterButtonHTML('ALL','全部',selectedStones.length===0)}${stones.map(x=>filterButtonHTML(x,stoneFilterLabel(x),selectedStones.includes(x))).join('')}</div>`;
  typeBox.classList.toggle('hidden',!types.length);stoneBox.classList.toggle('hidden',!stones.length);
  typeBox.querySelectorAll('.filter-chip').forEach(b=>b.onclick=()=>toggleStockFilter('type',b.dataset.value));stoneBox.querySelectorAll('.filter-chip').forEach(b=>b.onclick=()=>toggleStockFilter('stone',b.dataset.value));
  const hasFilters=types.length||stones.length;if(filterControls){filterControls.classList.toggle('hidden',!hasFilters);$('#stockFilterSelectionSummary').textContent=stockFilterSummaryText();$('#stockFilterToggle').textContent=state.stockSearch.filtersOpen?'收起篩選':'展開篩選'}if(filterPanel)filterPanel.classList.toggle('hidden',!hasFilters||!state.stockSearch.filtersOpen);
  const baseShown=all.filter(x=>(selectedTypes.length===0||selectedTypes.includes(articleType(x.artNo)))&&(selectedStones.length===0||stoneCodesForProduct(x).some(code=>selectedStones.includes(code)))&&(selectedStatuses.length===0||selectedStatuses.includes(x.status)));
  const shown=(state.stockSearch.imageIssuesOnly?baseShown.filter(x=>stockImageMeta(x).issue):baseShown).sort((a,b)=>compareColorSortedStock(a,b,selectedStones));
  const counts={AVAILABLE:0,CONSIGNED:0,SOLD_ON_HAND:0,SOLD_DELIVERED:0};for(const x of all)counts[x.status]=(counts[x.status]||0)+1;
  stockSearchPerf.baseMessage=(wildcard?`全部庫存：找到 ${all.length} 件`:`搜尋 ${core}：找到 ${all.length} 件`)+`，目前顯示 ${shown.length} 件。`;$('#stockSearchMessage').textContent=stockSearchPerf.baseMessage;
  const summary=$('#stockSearchSummary');summary.innerHTML=`<button type="button" class="stock-summary-chip${selectedStatuses.includes('AVAILABLE')?' active':''}" data-status="AVAILABLE" aria-pressed="${selectedStatuses.includes('AVAILABLE')?'true':'false'}">Avail ${counts.AVAILABLE||0}</button><button type="button" class="stock-summary-chip${selectedStatuses.includes('CONSIGNED')?' active':''}" data-status="CONSIGNED" aria-pressed="${selectedStatuses.includes('CONSIGNED')?'true':'false'}">Consign ${counts.CONSIGNED||0}</button><button type="button" class="stock-summary-chip${selectedStatuses.includes('SOLD_ON_HAND')?' active':''}" data-status="SOLD_ON_HAND" aria-pressed="${selectedStatuses.includes('SOLD_ON_HAND')?'true':'false'}">Sold-OH ${counts.SOLD_ON_HAND||0}</button><button type="button" class="stock-summary-chip${selectedStatuses.includes('SOLD_DELIVERED')?' active':''}" data-status="SOLD_DELIVERED" aria-pressed="${selectedStatuses.includes('SOLD_DELIVERED')?'true':'false'}">Deliv ${counts.SOLD_DELIVERED||0}</button><button type="button" class="stock-summary-chip image-issue-chip${state.stockSearch.imageIssuesOnly?' active':''}" data-image-issues="1" aria-pressed="${state.stockSearch.imageIssuesOnly?'true':'false'}">圖片待處理 …</button>`;
  summary.querySelectorAll('[data-status]').forEach(b=>b.onclick=()=>toggleStockFilter('status',b.dataset.status));const issueBtn=summary.querySelector('[data-image-issues]');if(issueBtn)issueBtn.onclick=()=>{state.stockSearch.imageIssuesOnly=!state.stockSearch.imageIssuesOnly;renderStockSearch()};summary.classList.remove('hidden');
  stockSearchPerf.shown=shown;stockSearchPerf.renderedCount=0;stockSearchPerf.renderKey=stockSearchViewKey();resultsEl.innerHTML='';
  if(!shown.length){resultsEl.innerHTML=`<div class="notice">沒有符合目前${state.stockSearch.imageIssuesOnly?'圖片待處理／':''}款式／石頭／狀態篩選的貨品。</div>`;scheduleStockIssueCount(baseShown);return}
  appendStockSearchBatch(STOCK_SEARCH_INITIAL_BATCH);scheduleStockIssueCount(baseShown)
}
function activePanelId(){return $('.tab-panel.active')?.id||''}
function updateBackToTopButton(){const btn=$('#backToTopBtn');if(!btn)return;const panel=activePanelId();let visible=false,title='回到頂部';if(panel==='invoice'){visible=($('#invoiceItems')?.scrollTop||0)>180;title=`回到 ${documentLabels().short} 貨品列表頂部`}else if(panel==='stockSearch'){const section=$('#stockSearch'),hasResults=($('#stockSearchResults')?.children.length||0)>0;visible=!!section&&hasResults&&window.scrollY>section.offsetTop+260;title='回到款式搜尋頂部'}btn.title=title;btn.setAttribute('aria-label',title);btn.classList.toggle('hidden',!visible)}
function scrollCurrentPanelToTop(){const panel=activePanelId();if(panel==='invoice')$('#invoiceItems')?.scrollTo({top:0,behavior:'smooth'});else if(panel==='stockSearch'){const section=$('#stockSearch');if(section)window.scrollTo({top:Math.max(0,section.offsetTop-72),behavior:'smooth'})}setTimeout(updateBackToTopButton,350)}
$('#backToTopBtn').onclick=scrollCurrentPanelToTop;$('#invoiceItems').addEventListener('scroll',updateBackToTopButton,{passive:true});window.addEventListener('scroll',updateBackToTopButton,{passive:true});window.addEventListener('resize',updateBackToTopButton,{passive:true});
$('#stockSearchBtn').onclick=runStockSearch;$('#stockSearchInput').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();runStockSearch()}};$('#stockSearchInput').oninput=e=>{e.target.value=e.target.value.toUpperCase();if(!norm(e.target.value)){state.stockSearch.query='*';state.stockSearch.types=[];state.stockSearch.stones=[];state.stockSearch.statuses=[];state.stockSearch.imageIssuesOnly=false;renderStockSearch()}};$('#stockFilterToggle').onclick=()=>{state.stockSearch.filtersOpen=!state.stockSearch.filtersOpen;renderStockSearch()};
$('#exportJmsdataBtn').onclick=exportJmsdata;$('#closeStockImageDialogBtn').onclick=()=>$('#stockImageDialog').close();$('#stockImageDialog').addEventListener('close',()=>{const p=imageEditorProduct();if(p)releaseEditorCandidateUrls(p);state.stockImageEditLot=null;ensureStockSearchRendered()});$('#stockUploadImageBtn').onclick=()=>$('#stockUploadImageInput').click();$('#stockCameraImageBtn').onclick=()=>$('#stockCameraImageInput').click();$('#stockUploadImageInput').onchange=async e=>{const f=e.target.files?.[0];e.target.value='';if(f)await handleImageEditorFile(f,'upload')};$('#stockCameraImageInput').onchange=async e=>{const f=e.target.files?.[0];e.target.value='';if(f)await handleImageEditorFile(f,'camera')};window.addEventListener('pagehide',()=>{if(state.draft.timer){clearTimeout(state.draft.timer);state.draft.timer=null}saveDocumentDraft()});window.addEventListener('beforeunload',e=>{saveDocumentDraft();if(!state.imageOverrideDirty&&!(state.packageImageDirtyFiles?.size||0))return;e.preventDefault();e.returnValue=''});$('#recallSearchBtn').onclick=renderRecallResults;$('#recallSearchInput').oninput=renderRecallResults;$('#recallSearchInput').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();renderRecallResults()}};$('#recallTypeFilter').onchange=renderRecallResults;$('#cancelRecallBtn').onclick=cancelRecall;
$$('input[name="quoteKarat"]').forEach(r=>r.onchange=e=>setQuoteKarat(e.target.value));$('#toggleGoldQuoteBtn').onclick=toggleGoldQuoteDetails;$('#refreshGoldBtn').onclick=openGoldSilverLondonHistory;let goldInputTimer=null;$('#currentLondonPmInput').oninput=e=>{clearTimeout(goldInputTimer);goldInputTimer=setTimeout(()=>setCurrentLondonPmValue(e.target.value),180)};
$('#exhibitionName')?.addEventListener('change',syncExhibitionNameFromInput);$('#exhibitionName')?.addEventListener('blur',syncExhibitionNameFromInput);
$('#runHealthCheckBtn').onclick=runHealthCheck;$('#showStoneConsistencyBtn').onclick=showStoneConsistencyDetails;$('#closeStoneConsistencyBtn').onclick=()=>$('#stoneConsistencyDialog').close();
$('#diagnosticLotBtn').onclick=openDiagnosticByLot;$('#diagnosticLotInput').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();openDiagnosticByLot()}};$('#closeDiagnosticBtn').onclick=()=>$('#diagnosticDialog').close();$('#exportDiagnosticBtn').onclick=exportDiagnosticReport;
$('#exportPwaBackupBtn').onclick=exportPwaBackup;$('#importPwaBackupBtn').onclick=()=>$('#pwaBackupInput').click();$('#pwaBackupInput').onchange=async e=>{const f=e.target.files?.[0];e.target.value='';if(f)await restorePwaBackupFile(f)};$('#clearAllPwaDataBtn').onclick=clearAllPwaLocalData;

const appVersionEyebrow=$('#appVersionEyebrow');if(appVersionEyebrow)appVersionEyebrow.textContent=`Jewellery Exhibition PWA · v${APP_VERSION}`;
setupPreviewAddonControls();updateFxPanel();updateDocumentTypeUI();updateGoldQuoteUI();renderCustomerSummary();renderItems();renderRecallResults();updateImageOverrideStatus();updateDataVersionPanel();runHealthCheck();renderStockSearch();schedulePreview();updateBackToTopButton();markDraftBaseline();maybeOfferDraftRestore();if('serviceWorker'in navigator)navigator.serviceWorker.register('./sw.js',{updateViaCache:'none'}).then(reg=>reg.update().catch(()=>{})).catch(()=>{});
