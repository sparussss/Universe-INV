'use strict';

const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];

const STONE_MAP = {
  'SKY BTO': 'SKY BT', 'SKY BT': 'SKY BT', 'AMCT': 'AMCT', 'QAM': 'AM', 'LBT': 'LBT',
  'BTO': 'BT', 'YCT': 'CT', 'GPS': 'G.AM', 'GAM': 'G.AM', 'GPD': 'PD', 'RQZ': 'RQZ',
  'MG': 'MG', 'PTQ': 'PTR', 'AQ': 'AQ', 'PAM': 'P.AM', 'MCT': 'MCT', 'RGT': 'RGT',
  'TZ': 'TZ', 'IO': 'IO', 'GTQ': 'GTR', 'GT': 'GT', 'ALEX': 'ALEX', 'KU': 'KU',
  'LQZ': 'LQZ', 'SQZ': 'SQZ', 'BSA': 'BSA', 'PSA': 'PSA', 'GGT': 'GGT', 'OSA': 'OSA',
  'YSA': 'YSA', 'SSU': 'SSU', 'RRU': 'RRU', 'GEM': 'GEM', 'GSA': 'GSA', 'WSA': 'WSA',
  'ZSA': 'ZSA', 'ZSP': 'ZSP', 'DIA': 'DIA', 'AG': 'AG', 'AMZ': 'AMZ', 'BCH': 'BCH',
  'BO': 'BO', 'GMA': 'GMA', 'LAB': 'LAB', 'LAP': 'LAP', 'MOON': 'MOON', 'OPAL': 'OPAL',
  'WPL': 'WPL', 'RCH': 'RCH', 'TE': 'TE', 'TQ': 'TQ'
};
const STONE_ALIASES = Object.keys(STONE_MAP).sort((a, b) => b.length - a.length);

const state = {
  products: new Map(),
  customers: new Map(),
  stockHeader: [],
  stockRows: [],
  stockLotIndex: -1,
  stockFileName: '',
  imageIndex: new Map(),
  items: [],
  insertIndex: null,
  deletedItem: null,
  deletedTimer: null,
  scanner: null,
  scannerBusy: false,
  scannerStarted: false,
  currentZoom: 3,
  torchOn: false,
  lastScan: { value: '', time: 0 },
  feedbackTimer: null,
  audioContext: null,
  soldHistory: new Map()
};

function normalizeText(value) {
  return String(value ?? '').trim();
}
function normalizeCode(value) {
  return normalizeText(value).replace(/\s+/g, '').toUpperCase();
}
function normalizeArticle(value) {
  return normalizeText(value).replace(/\s+/g, ' ').toUpperCase();
}
function normalizeVariant(value) {
  return normalizeText(value).replace(/\s+/g, ' ').toUpperCase();
}
function safeFileName(value) {
  return normalizeText(value).replace(/[\\/:*?"<>|]+/g, '_') || 'Invoice';
}
function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}
function showStatus(selector, text, type = '') {
  const el = $(selector);
  el.textContent = text;
  el.className = `status${type ? ` ${type}` : ''}`;
}
function formatMoney(value) {
  const currency = $('#currency').value || 'USD';
  return new Intl.NumberFormat('en-US', { style: 'currency', currency, minimumFractionDigits: 2 }).format(Number(value) || 0);
}
function todayIso() {
  return new Date().toISOString().slice(0, 10);
}
$('#invoiceDate').value = todayIso();

function switchTab(name) {
  $$('.tab').forEach(button => button.classList.toggle('active', button.dataset.tab === name));
  $$('.tab-panel').forEach(panel => panel.classList.toggle('active', panel.id === name));
  if (name === 'invoice') {
    renderCustomerSummary();
    renderItems();
  }
}
$$('.tab').forEach(button => button.addEventListener('click', () => switchTab(button.dataset.tab)));
$('#goInvoiceBtn').addEventListener('click', () => switchTab('invoice'));

function readWorkbook(file) {
  if (!window.XLSX) throw new Error('Excel 程式未載入，請連接網絡後重新開啟。');
  return file.arrayBuffer().then(buffer => XLSX.read(buffer, { type: 'array', cellDates: false }));
}

function findHeaderIndex(headers, names) {
  const normalized = headers.map(h => normalizeText(h).toUpperCase().replace(/[.\s_]/g, ''));
  for (const name of names) {
    const target = name.toUpperCase().replace(/[.\s_]/g, '');
    const index = normalized.indexOf(target);
    if (index >= 0) return index;
  }
  return -1;
}

$('#stockInput').addEventListener('change', async event => {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    const workbook = await readWorkbook(file);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
    if (!rows.length) throw new Error('Excel 沒有資料。');
    const header = rows[0].map(normalizeText);
    const lotIndex = findHeaderIndex(header, ['LOTNO', 'LOT NO']);
    const artIndex = findHeaderIndex(header, ['ARTNO', 'ART NO']);
    const priceIndex = findHeaderIndex(header, ['PRICE']);
    const unitIndex = findHeaderIndex(header, ['UNIT']);
    const descIndexes = [1, 2, 3, 4, 5, 6].map(n => findHeaderIndex(header, [`DESC${n}`]));
    if ([lotIndex, artIndex, priceIndex].some(index => index < 0)) throw new Error('找不到 LOTNO、ARTNO 或 PRICE 欄位。');

    const products = new Map();
    const validRows = [];
    for (const row of rows.slice(1)) {
      const lotNo = normalizeCode(row[lotIndex]);
      const artNo = normalizeArticle(row[artIndex]);
      const price = Number(row[priceIndex]);
      if (!lotNo || !artNo || !Number.isFinite(price)) continue;
      const descriptions = descIndexes.map(index => index >= 0 ? normalizeText(row[index]) : '').filter(Boolean);
      const desc2 = descIndexes[1] >= 0 ? normalizeText(row[descIndexes[1]]) : '';
      const product = {
        lotNo,
        artNo,
        price,
        unit: unitIndex >= 0 ? (normalizeText(row[unitIndex]) || 'PC') : 'PC',
        descriptions,
        desc2,
        rawRow: [...row]
      };
      products.set(lotNo, product);
      validRows.push([...row]);
    }
    if (!products.size) throw new Error('沒有找到有效貨品資料。');
    state.products = products;
    state.stockHeader = header;
    state.stockRows = validRows;
    state.stockLotIndex = lotIndex;
    state.stockFileName = file.name;
    state.items = state.items.filter(item => products.has(item.lotNo));
    showStatus('#stockStatus', `已匯入 ${file.name}：${products.size} 件貨品。`, 'ok');
    rebuildImageIndexFromCurrentFolder();
    renderItems();
  } catch (error) {
    showStatus('#stockStatus', `匯入失敗：${error.message}`, 'error');
  }
});

$('#customerInput').addEventListener('change', async event => {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    const workbook = await readWorkbook(file);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
    const customers = new Map();
    for (const row of rows) {
      const code = normalizeCode(row[0]);
      const company = normalizeText(row[1]);
      if (!code || !company || code.includes('CUSTOMER')) continue;
      const rawRate = row[11];
      const parsedRate = Number(rawRate);
      const rate = rawRate === '' || rawRate === null || rawRate === undefined || !Number.isFinite(parsedRate) ? 0.34 : parsedRate;
      const address = [row[2], row[3], row[4]].map(normalizeText).filter(Boolean).join('\n');
      customers.set(code, {
        code,
        company,
        address,
        rate,
        terms: normalizeText(row[10]),
        contact: normalizeText(row[9]),
        phone: normalizeText(row[6]),
        email: normalizeText(row[8])
      });
    }
    if (!customers.size) throw new Error('沒有找到有效客戶資料。');
    state.customers = customers;
    showStatus('#customerStatus', `已匯入 ${file.name}：${customers.size} 位客戶。`, 'ok');
  } catch (error) {
    showStatus('#customerStatus', `匯入失敗：${error.message}`, 'error');
  }
});

function stripDuplicateSuffix(baseName) {
  const match = baseName.match(/\s*\((\d+)\)\s*$/);
  if (!match) return { clean: baseName.trim(), rank: 0 };
  return { clean: baseName.slice(0, match.index).trim(), rank: Number(match[1]) || 1 };
}

function parseImageFile(file, artNos) {
  const extensionMatch = file.name.match(/\.([^.]+)$/);
  if (!extensionMatch) return null;
  const extension = extensionMatch[1].toLowerCase();
  if (!['jpg', 'jpeg', 'png', 'webp'].includes(extension)) return null;
  const base = file.name.slice(0, -extensionMatch[0].length).replace(/\s+/g, ' ').trim();
  const { clean, rank } = stripDuplicateSuffix(base);
  const upper = clean.toUpperCase();
  const artNo = artNos.find(art => upper === art || upper.startsWith(`${art} `));
  if (!artNo) return null;
  const variant = clean.slice(artNo.length).trim() || 'Default';
  return { artNo, variant, variantKey: normalizeVariant(variant), duplicateRank: rank, file, extension };
}

function rebuildImageIndexFromCurrentFolder() {
  const input = $('#pictureFolderInput');
  const files = [...(input.files || [])];
  if (!files.length || !state.products.size) return;
  buildImageIndex(files);
}

function buildImageIndex(files) {
  for (const entries of state.imageIndex.values()) {
    for (const entry of entries) if (entry.url) URL.revokeObjectURL(entry.url);
  }
  state.imageIndex.clear();
  const artNos = [...new Set([...state.products.values()].map(product => product.artNo))].sort((a, b) => b.length - a.length);
  let matched = 0;
  for (const file of files) {
    const parsed = parseImageFile(file, artNos);
    if (!parsed) continue;
    parsed.url = URL.createObjectURL(file);
    const list = state.imageIndex.get(parsed.artNo) || [];
    list.push(parsed);
    state.imageIndex.set(parsed.artNo, list);
    matched += 1;
  }
  for (const [artNo, entries] of state.imageIndex) {
    entries.sort((a, b) => {
      if (a.variantKey === b.variantKey) return a.duplicateRank - b.duplicateRank;
      if (a.variantKey === 'DEFAULT') return -1;
      if (b.variantKey === 'DEFAULT') return 1;
      return a.variantKey.localeCompare(b.variantKey);
    });
    state.imageIndex.set(artNo, entries);
  }
  const total = files.filter(file => /\.(jpg|jpeg|png|webp)$/i.test(file.name)).length;
  const skipped = Math.max(0, total - matched);
  showStatus('#pictureStatus', `已選擇 Folder：${total} 張圖片；其中 ${matched} 張符合目前倉存，${skipped} 張略過。`, 'ok');
  for (const item of state.items) {
    if (!item.imageVariantManual) item.imageVariant = chooseAutomaticVariant(item);
  }
  renderItems();
}

$('#pictureFolderInput').addEventListener('change', event => {
  const files = [...(event.target.files || [])];
  if (!files.length) return;
  if (!state.products.size) {
    showStatus('#pictureStatus', '請先匯入倉存表，再選擇圖片 Folder。', 'error');
    return;
  }
  buildImageIndex(files);
});

function uniqueVariantsForArt(artNo) {
  const entries = state.imageIndex.get(artNo) || [];
  const best = new Map();
  for (const entry of entries) {
    const current = best.get(entry.variantKey);
    if (!current || entry.duplicateRank < current.duplicateRank) best.set(entry.variantKey, entry);
  }
  return [...best.values()].sort((a, b) => {
    if (a.variantKey === 'DEFAULT') return -1;
    if (b.variantKey === 'DEFAULT') return 1;
    return a.variant.localeCompare(b.variant);
  });
}

function extractStoneVariants(desc2) {
  const text = normalizeText(desc2).toUpperCase().replace(/SKY\s+BTO/g, 'SKY BTO');
  const found = [];
  for (const alias of STONE_ALIASES) {
    const escaped = alias.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(`(?:^|[-+\\s])${escaped}(?=[A-Z0-9.(]|$)`, 'g');
    if (regex.test(text)) found.push(STONE_MAP[alias]);
  }
  return [...new Set(found)];
}

function chooseAutomaticVariant(item) {
  const variants = uniqueVariantsForArt(item.artNo);
  if (!variants.length) return 'Default';
  const stones = extractStoneVariants(item.desc2);
  const candidates = [];
  if (stones.length) {
    candidates.push(stones.join('+'));
    candidates.push(...stones);
  }
  candidates.push('Default');
  const normalizedCandidates = candidates.map(normalizeVariant);
  for (const candidate of normalizedCandidates) {
    const exact = variants.find(entry => entry.variantKey === candidate);
    if (exact) return exact.variant;
  }
  for (const candidate of normalizedCandidates) {
    const partial = variants.find(entry => entry.variantKey.includes(candidate) || candidate.includes(entry.variantKey));
    if (partial) return partial.variant;
  }
  return variants[0].variant;
}

function getImageEntry(item) {
  const variants = uniqueVariantsForArt(item.artNo);
  if (!variants.length) return null;
  const key = normalizeVariant(item.imageVariant || chooseAutomaticVariant(item));
  return variants.find(entry => entry.variantKey === key) || variants[0];
}

function placeholderDataUrl(text) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="240" height="240"><rect width="100%" height="100%" fill="#f1f5f9"/><text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" font-family="Arial" font-size="18" fill="#64748b">${escapeHtml(text)}</text></svg>`;
  return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(svg)}`;
}

function searchCustomers(query) {
  const raw = normalizeText(query).toUpperCase();
  const code = normalizeCode(query);
  if (!raw) return [];
  return [...state.customers.values()].filter(customer => customer.code.includes(code) || customer.company.toUpperCase().includes(raw)).slice(0, 15);
}

function renderCustomerResults() {
  const container = $('#customerResults');
  const matches = searchCustomers($('#customerSearch').value);
  container.innerHTML = '';
  if (!matches.length) {
    container.innerHTML = '<div class="status">找不到客戶。</div>';
    return;
  }
  for (const customer of matches) {
    const button = document.createElement('button');
    button.className = 'customer-result';
    button.innerHTML = `<span><strong>${escapeHtml(customer.code)} · ${escapeHtml(customer.company)}</strong><small>${escapeHtml(customer.address).replace(/\n/g, ' · ')}</small></span><span>${customer.rate}</span>`;
    button.addEventListener('click', () => selectCustomer(customer));
    container.appendChild(button);
  }
}

function selectCustomer(customer) {
  $('#customerCode').value = customer.code;
  $('#customerName').value = customer.company;
  $('#customerAddress').value = customer.address;
  $('#salesRate').value = customer.rate;
  $('#terms').value = customer.terms || '';
  $('#customerSearch').value = '';
  $('#customerResults').innerHTML = '';
  repriceAllItems();
  renderCustomerSummary();
  saveDraft();
}

$('#customerSearchBtn').addEventListener('click', renderCustomerResults);
$('#customerSearch').addEventListener('input', event => {
  if (event.target.value.length >= 2) renderCustomerResults();
  else $('#customerResults').innerHTML = '';
});
$('#customerSearch').addEventListener('keydown', event => {
  if (event.key === 'Enter') {
    event.preventDefault();
    renderCustomerResults();
  }
});

function renderCustomerSummary() {
  const code = normalizeText($('#customerCode').value);
  const name = normalizeText($('#customerName').value);
  const rate = normalizeText($('#salesRate').value);
  if (!code && !name) {
    $('#customerSummary').textContent = '尚未選擇客戶。';
    return;
  }
  $('#customerSummary').innerHTML = `<strong>${escapeHtml(code)} · ${escapeHtml(name)}</strong><br><span>${escapeHtml($('#customerAddress').value).replace(/\n/g, '<br>')}</span><br><small>Sales Rate ${escapeHtml(rate)} · ${escapeHtml($('#currency').value)}</small>`;
}

function nextSequence() {
  return state.items.reduce((max, item) => Math.max(max, item.seq || 0), 0) + 1;
}

function addByLot(rawValue, source = 'manual') {
  const lotNo = normalizeCode(rawValue).replace(/[^0-9A-Z]/g, '');
  if (!lotNo) return { status: 'error', lotNo: '', text: '請輸入 LOTNO' };
  if (state.items.some(item => item.lotNo === lotNo)) return { status: 'duplicate', lotNo, text: '已加入' };
  if (state.soldHistory.has(lotNo)) return { status: 'duplicate', lotNo, text: `已售 ${state.soldHistory.get(lotNo).invoiceNo}` };
  const product = state.products.get(lotNo);
  if (!product) return { status: 'error', lotNo, text: '找不到' };
  const rate = Number($('#salesRate').value);
  if (!Number.isFinite(rate) || rate <= 0) return { status: 'error', lotNo, text: '未選客戶' };
  const item = {
    id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`,
    seq: nextSequence(),
    lotNo: product.lotNo,
    artNo: product.artNo,
    price: product.price,
    unit: product.unit,
    descriptions: [...product.descriptions],
    desc2: product.desc2,
    qty: 1,
    unitPrice: Math.ceil(product.price * rate),
    imageVariant: 'Default',
    imageVariantManual: false,
    addedAt: Date.now()
  };
  item.imageVariant = chooseAutomaticVariant(item);
  if (state.insertIndex === null) state.items.push(item);
  else {
    state.items.splice(state.insertIndex, 0, item);
    state.insertIndex = null;
  }
  $('#lotInput').value = '';
  renderItems(item.id);
  saveDraft();
  if (source !== 'scan') showGlobalFeedback(lotNo, '', 'success', 1000);
  return { status: 'success', lotNo, text: '' };
}

$('#addLotBtn').addEventListener('click', () => {
  const result = addByLot($('#lotInput').value);
  showAddResult(result);
});
$('#lotInput').addEventListener('keydown', event => {
  if (event.key === 'Enter') {
    event.preventDefault();
    const result = addByLot(event.target.value);
    showAddResult(result);
  }
});
function showAddResult(result) {
  if (result.status === 'success') showStatus('#addStatus', `已加入 LOTNO ${result.lotNo}。`, 'ok');
  else showStatus('#addStatus', `${result.lotNo ? `${result.lotNo}：` : ''}${result.text}`, 'error');
  if (result.status !== 'success') showGlobalFeedback(result.lotNo || '—', result.text, result.status, result.status === 'error' ? 1800 : 1200);
}

$('#manualModeBtn').addEventListener('click', () => {
  $('#manualPanel').classList.remove('hidden');
  $('#lotInput').focus();
});

$('#voiceBtn').addEventListener('click', () => {
  const Recognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!Recognition) {
    $('#manualPanel').classList.remove('hidden');
    $('#lotInput').focus();
    showStatus('#addStatus', '請使用 iPhone 鍵盤的咪高峰輸入 LOTNO。');
    return;
  }
  try {
    const recognition = new Recognition();
    recognition.lang = 'zh-HK';
    recognition.interimResults = false;
    recognition.maxAlternatives = 1;
    showStatus('#addStatus', '請讀出 LOTNO。');
    recognition.onresult = event => {
      const transcript = event.results[0][0].transcript;
      const digits = transcript.replace(/[^0-9]/g, '');
      $('#lotInput').value = digits;
      const result = addByLot(digits, 'voice');
      showAddResult(result);
    };
    recognition.onerror = () => showStatus('#addStatus', '語音未能辨識，請再試或手動輸入。', 'error');
    recognition.start();
  } catch {
    $('#lotInput').focus();
  }
});

function calcTotals() {
  const qty = state.items.reduce((sum, item) => sum + (Number(item.qty) || 0), 0);
  const subtotal = state.items.reduce((sum, item) => sum + (Number(item.qty) || 0) * (Number(item.unitPrice) || 0), 0);
  const discount = Math.max(0, Number($('#discountAmount').value) || 0);
  return { qty, subtotal, discount, total: Math.max(0, subtotal - discount) };
}

function updateSummary() {
  const totals = calcTotals();
  $('#totalQty').textContent = totals.qty;
  $('#subtotal').textContent = formatMoney(totals.subtotal);
  $('#discountDisplay').textContent = formatMoney(totals.discount);
  $('#grandTotal').textContent = formatMoney(totals.total);
  $('#itemCountLabel').textContent = `共 ${state.items.length} 件`;
  $('#scannerItemCount').textContent = state.items.length;
}

function renderItems(latestId = null) {
  const container = $('#invoiceItems');
  container.innerHTML = '';
  if (!state.items.length) {
    container.className = 'invoice-items empty';
    container.textContent = '尚未加入貨品。';
    updateSummary();
    return;
  }
  container.className = 'invoice-items';
  const displayItems = [...state.items].sort((a, b) => (b.seq || 0) - (a.seq || 0));
  for (const item of displayItems) {
    const template = $('#itemTemplate').content.firstElementChild.cloneNode(true);
    if (item.id === latestId) template.classList.add('latest');
    $('.item-number', template).textContent = `${item.seq}.`;
    $('.item-artno', template).textContent = item.artNo;
    $('.item-lot', template).textContent = `LOTNO ${item.lotNo}`;
    $('.item-description', template).textContent = item.descriptions.join('\n');
    const imageEntry = getImageEntry(item);
    $('.item-image', template).src = imageEntry?.url || placeholderDataUrl(item.artNo);
    const variantSelect = $('.variant-select', template);
    const variants = uniqueVariantsForArt(item.artNo);
    if (!variants.length) {
      variantSelect.innerHTML = '<option>沒有圖片</option>';
      variantSelect.disabled = true;
    } else {
      for (const entry of variants) {
        const option = document.createElement('option');
        option.value = entry.variant;
        option.textContent = entry.variant;
        option.selected = normalizeVariant(entry.variant) === normalizeVariant(item.imageVariant);
        variantSelect.appendChild(option);
      }
    }
    $('.qty-input', template).value = item.qty;
    $('.price-input', template).value = item.unitPrice;
    variantSelect.addEventListener('change', event => {
      item.imageVariant = event.target.value;
      item.imageVariantManual = true;
      renderItems();
      saveDraft();
    });
    $('.qty-input', template).addEventListener('change', event => {
      item.qty = Math.max(1, Number(event.target.value) || 1);
      updateSummary();
      saveDraft();
    });
    $('.price-input', template).addEventListener('change', event => {
      item.unitPrice = Math.max(0, Math.ceil(Number(event.target.value) || 0));
      updateSummary();
      saveDraft();
    });
    $('.insert-before', template).addEventListener('click', () => prepareInsert(item, false));
    $('.insert-after', template).addEventListener('click', () => prepareInsert(item, true));
    $('.delete-item', template).addEventListener('click', () => deleteItem(item));
    $('.image-button', template).addEventListener('click', () => {
      if (!imageEntry) return;
      window.open(imageEntry.url, '_blank');
    });
    container.appendChild(template);
  }
  updateSummary();
  if (latestId) $('#itemViewport').scrollTop = 0;
}

function prepareInsert(item, after) {
  const index = state.items.findIndex(candidate => candidate.id === item.id);
  state.insertIndex = Math.max(0, index + (after ? 1 : 0));
  $('#lotInput').focus();
  showStatus('#addStatus', `下一件貨會插入 ${item.seq} 號貨品${after ? '下方' : '上方'}。`);
}

function deleteItem(item) {
  const index = state.items.findIndex(candidate => candidate.id === item.id);
  if (index < 0) return;
  state.deletedItem = { item, index };
  state.items.splice(index, 1);
  renderItems();
  $('#undoText').textContent = `已刪除 ${item.artNo} / ${item.lotNo}`;
  $('#undoBar').classList.remove('hidden');
  clearTimeout(state.deletedTimer);
  state.deletedTimer = setTimeout(() => {
    state.deletedItem = null;
    $('#undoBar').classList.add('hidden');
  }, 6000);
  saveDraft();
}

$('#undoBtn').addEventListener('click', () => {
  if (!state.deletedItem) return;
  state.items.splice(state.deletedItem.index, 0, state.deletedItem.item);
  const restoredId = state.deletedItem.item.id;
  state.deletedItem = null;
  clearTimeout(state.deletedTimer);
  $('#undoBar').classList.add('hidden');
  renderItems(restoredId);
  saveDraft();
});

$('#scrollLatestBtn').addEventListener('click', () => { $('#itemViewport').scrollTop = 0; });
$('#clearDraftBtn').addEventListener('click', () => {
  if (!state.items.length || confirm('確定清空目前 Invoice 草稿？')) {
    state.items = [];
    renderItems();
    saveDraft();
  }
});

function repriceAllItems() {
  const rate = Number($('#salesRate').value);
  if (!Number.isFinite(rate) || rate <= 0) return;
  for (const item of state.items) item.unitPrice = Math.ceil(item.price * rate);
  renderItems();
}

$('#salesRate').addEventListener('change', () => { repriceAllItems(); renderCustomerSummary(); saveDraft(); });
$('#currency').addEventListener('change', () => { renderCustomerSummary(); updateSummary(); saveDraft(); });
$('#discountAmount').addEventListener('input', () => { updateSummary(); saveDraft(); });
['#customerCode', '#customerName', '#customerAddress', '#terms', '#shipmentMethod', '#invoiceNo', '#invoiceDate', '#remark'].forEach(selector => {
  $(selector).addEventListener('change', () => { renderCustomerSummary(); saveDraft(); });
});

function showGlobalFeedback(lotNo, text, type, duration) {
  const box = $('#globalFeedback');
  box.className = `global-feedback${type === 'duplicate' ? ' duplicate' : type === 'error' ? ' error' : ''}`;
  $('#globalFeedbackLot').textContent = lotNo;
  $('#globalFeedbackText').textContent = text;
  clearTimeout(state.feedbackTimer);
  state.feedbackTimer = setTimeout(() => box.classList.add('hidden'), duration);
}

function showScannerFeedback(lotNo, text, type, duration) {
  const box = $('#scanFeedback');
  box.className = `scan-feedback${type === 'duplicate' ? ' duplicate' : type === 'error' ? ' error' : ''}`;
  $('#scanFeedbackLot').textContent = lotNo;
  $('#scanFeedbackText').textContent = text;
  clearTimeout(state.feedbackTimer);
  state.feedbackTimer = setTimeout(() => box.classList.add('hidden'), duration);
}

function initAudio() {
  try {
    state.audioContext = state.audioContext || new (window.AudioContext || window.webkitAudioContext)();
    state.audioContext.resume?.();
  } catch {}
}
function successBeep() {
  try {
    if (!state.audioContext) return;
    const oscillator = state.audioContext.createOscillator();
    const gain = state.audioContext.createGain();
    oscillator.frequency.value = 880;
    gain.gain.setValueAtTime(0.0001, state.audioContext.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.18, state.audioContext.currentTime + 0.01);
    gain.gain.exponentialRampToValueAtTime(0.0001, state.audioContext.currentTime + 0.12);
    oscillator.connect(gain).connect(state.audioContext.destination);
    oscillator.start();
    oscillator.stop(state.audioContext.currentTime + 0.13);
  } catch {}
}

$('#scanBtn').addEventListener('click', () => {
  initAudio();
  startScanner();
});
$('#closeScannerBtn').addEventListener('click', stopScanner);
$('#restartCameraBtn').addEventListener('click', startScanner);
$$('.zoom-btn').forEach(button => button.addEventListener('click', () => setZoom(Number(button.dataset.zoom))));
$('#torchBtn').addEventListener('click', toggleTorch);

async function chooseRearCamera() {
  const cameras = await Html5Qrcode.getCameras();
  if (!cameras?.length) throw new Error('找不到相機。');
  return cameras.find(camera => /back|rear|environment/i.test(camera.label)) || cameras[cameras.length - 1];
}

async function startScanner() {
  if (state.scannerBusy) return;
  state.scannerBusy = true;
  $('#restartCameraBtn').classList.add('hidden');
  $('#scannerDebug').textContent = '';
  const dialog = $('#scannerDialog');
  if (!dialog.open) dialog.showModal();
  try {
    if (!window.Html5Qrcode) throw new Error('掃描程式未載入，請連接網絡後重新開啟。');
    if (state.scannerStarted && state.scanner) {
      try { await state.scanner.stop(); } catch {}
      try { await state.scanner.clear(); } catch {}
      state.scannerStarted = false;
    }
    $('#reader').innerHTML = '';
    state.scanner = new Html5Qrcode('reader', { verbose: false });
    const camera = await chooseRearCamera();
    await state.scanner.start(
      camera.id,
      {
        fps: 18,
        qrbox: (width, height) => ({ width: Math.floor(width * 0.86), height: Math.max(72, Math.floor(height * 0.16)) }),
        aspectRatio: 1.7778,
        disableFlip: true,
        formatsToSupport: [
          Html5QrcodeSupportedFormats.CODE_128,
          Html5QrcodeSupportedFormats.CODE_39,
          Html5QrcodeSupportedFormats.EAN_13,
          Html5QrcodeSupportedFormats.EAN_8,
          Html5QrcodeSupportedFormats.UPC_A,
          Html5QrcodeSupportedFormats.UPC_E,
          Html5QrcodeSupportedFormats.ITF
        ]
      },
      onScanSuccess,
      () => {}
    );
    state.scannerStarted = true;
    await setZoom(3, true);
  } catch (error) {
    $('#scannerDebug').textContent = `相機無法啟動：${error.message || error}`;
    $('#restartCameraBtn').classList.remove('hidden');
  } finally {
    state.scannerBusy = false;
  }
}

async function stopScanner() {
  if (state.scannerBusy) return;
  state.scannerBusy = true;
  try {
    if (state.scannerStarted && state.scanner) await state.scanner.stop();
  } catch {}
  try { await state.scanner?.clear(); } catch {}
  state.scannerStarted = false;
  state.scanner = null;
  state.scannerBusy = false;
  if ($('#scannerDialog').open) $('#scannerDialog').close();
  $('#reader').innerHTML = '';
}

async function setZoom(value, silent = false) {
  state.currentZoom = value;
  $$('.zoom-btn').forEach(button => button.classList.toggle('active', Number(button.dataset.zoom) === value));
  if (!state.scannerStarted || !state.scanner) return;
  try {
    const capabilities = state.scanner.getRunningTrackCapabilities?.() || {};
    const max = Number(capabilities.zoom?.max ?? capabilities.zoom ?? value);
    const min = Number(capabilities.zoom?.min ?? 1);
    const target = Math.max(min, Math.min(value, max || value));
    await state.scanner.applyVideoConstraints({ advanced: [{ zoom: target }] });
    if (!silent && target !== value) $('#scannerDebug').textContent = `裝置最高支援 ${target}×。`;
    else if (!silent) $('#scannerDebug').textContent = '';
  } catch {
    if (!silent) $('#scannerDebug').textContent = '此裝置未提供網頁相機倍率控制。';
  }
}

async function toggleTorch() {
  if (!state.scannerStarted || !state.scanner) return;
  try {
    state.torchOn = !state.torchOn;
    await state.scanner.applyVideoConstraints({ advanced: [{ torch: state.torchOn }] });
    $('#torchBtn').classList.toggle('active', state.torchOn);
  } catch {
    $('#scannerDebug').textContent = '此裝置未提供網頁手電筒控制。';
  }
}

function onScanSuccess(decodedText) {
  const lotNo = normalizeCode(decodedText).replace(/[^0-9A-Z]/g, '');
  const now = Date.now();
  if (!lotNo || (state.lastScan.value === lotNo && now - state.lastScan.time < 1800)) return;
  state.lastScan = { value: lotNo, time: now };
  const result = addByLot(lotNo, 'scan');
  if (result.status === 'success') {
    successBeep();
    showScannerFeedback(lotNo, '', 'success', 1000);
  } else if (result.status === 'duplicate') {
    showScannerFeedback(lotNo, result.text, 'duplicate', 1200);
  } else {
    showScannerFeedback(lotNo, result.text, 'error', 2000);
  }
}

function numberToWords(number) {
  number = Math.floor(Number(number) || 0);
  if (number === 0) return 'ZERO';
  const ones = ['', 'ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'SIX', 'SEVEN', 'EIGHT', 'NINE', 'TEN', 'ELEVEN', 'TWELVE', 'THIRTEEN', 'FOURTEEN', 'FIFTEEN', 'SIXTEEN', 'SEVENTEEN', 'EIGHTEEN', 'NINETEEN'];
  const tens = ['', '', 'TWENTY', 'THIRTY', 'FORTY', 'FIFTY', 'SIXTY', 'SEVENTY', 'EIGHTY', 'NINETY'];
  const chunk = n => {
    const parts = [];
    if (n >= 100) { parts.push(`${ones[Math.floor(n / 100)]} HUNDRED`); n %= 100; }
    if (n >= 20) { parts.push(tens[Math.floor(n / 10)]); n %= 10; }
    if (n > 0) parts.push(ones[n]);
    return parts.join(' ');
  };
  const scales = [[1_000_000_000, 'BILLION'], [1_000_000, 'MILLION'], [1_000, 'THOUSAND'], [1, '']];
  const output = [];
  for (const [value, label] of scales) {
    if (number >= value) {
      const part = Math.floor(number / value);
      number %= value;
      output.push(`${chunk(part)}${label ? ` ${label}` : ''}`);
    }
  }
  return output.join(' ');
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

async function exportInvoiceXlsx() {
  if (!window.ExcelJS) throw new Error('XLSX 輸出程式未載入，請連接網絡後重新開啟。');
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Universe Invoice PWA';
  workbook.created = new Date();
  const sheet = workbook.addWorksheet('Sales Invoice', {
    pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.25, right: 0.25, top: 0.25, bottom: 0.35, header: 0.1, footer: 0.1 } }
  });
  sheet.views = [{ showGridLines: false }];
  const widths = [5, 18, 16, 16, 17, 8, 8, 13, 13];
  widths.forEach((width, index) => { sheet.getColumn(index + 1).width = width; });
  const thin = { style: 'thin', color: { argb: 'FFBFC7D1' } };
  const headerFont = { name: 'Arial', size: 9, bold: true };
  const bodyFont = { name: 'Arial', size: 9 };
  const items = [...state.items];
  const pages = [];
  for (let i = 0; i < items.length; i += 8) pages.push(items.slice(i, i + 8));
  if (!pages.length) pages.push([]);
  const totals = calcTotals();
  let row = 1;
  for (let pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
    const pageItems = pages[pageIndex];
    sheet.mergeCells(row, 1, row, 9);
    const companyCell = sheet.getCell(row, 1);
    companyCell.value = 'UNIVERSE GEMS & JEWELLERY CO.';
    companyCell.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FF164478' } };
    companyCell.alignment = { horizontal: 'center' };
    row += 1;
    sheet.mergeCells(row, 1, row + 1, 9);
    const addressCell = sheet.getCell(row, 1);
    addressCell.value = 'UNIT 11-12, 10/F., FU HANG INDUSTRIAL BUILDING, NO. 1 HOK YUEN STREET EAST,\nHUNG HOM, KOWLOON, HONG KONG · TEL : (852) 2363 5409 · FAX : (852) 2765 0343';
    addressCell.font = { name: 'Arial', size: 8 };
    addressCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    row += 2;
    sheet.mergeCells(row, 1, row, 4);
    sheet.getCell(row, 1).value = 'Sales Invoice';
    sheet.getCell(row, 1).font = { name: 'Arial', size: 14, bold: true };
    sheet.mergeCells(row, 5, row, 9);
    sheet.getCell(row, 5).value = "Vender's Banker";
    sheet.getCell(row, 5).font = headerFont;
    row += 1;
    const leftLines = [
      `No. : ${normalizeText($('#invoiceNo').value)}`,
      `Invoice Date : ${normalizeText($('#invoiceDate').value)}`,
      `Shipment Method : ${normalizeText($('#shipmentMethod').value)}`,
      `Currency : ${normalizeText($('#currency').value)}`,
      `Customer : ${normalizeText($('#customerName').value)}`,
      normalizeText($('#customerAddress').value)
    ].filter(Boolean).join('\n');
    const rightLines = [
      'The Hong Kong & Shanghai Banking Corporation Ltd.',
      'Address : 41 Ma Tau Wai Road,Hung Hom,Kowloon,Hong Kong',
      'A/C # : 012-593570-001',
      'A/C Name : Universe Gems & Jewellery Co.'
    ].join('\n');
    sheet.mergeCells(row, 1, row + 5, 4);
    sheet.getCell(row, 1).value = leftLines;
    sheet.getCell(row, 1).font = bodyFont;
    sheet.getCell(row, 1).alignment = { vertical: 'top', wrapText: true };
    sheet.mergeCells(row, 5, row + 5, 9);
    sheet.getCell(row, 5).value = rightLines;
    sheet.getCell(row, 5).font = bodyFont;
    sheet.getCell(row, 5).alignment = { vertical: 'top', wrapText: true };
    row += 6;
    const tableHeaderRow = row;
    ['No.', 'Article No.', 'Description', '', 'Picture', 'Quantity', 'Unit', 'Unit Price', 'Amount'].forEach((value, index) => {
      const cell = sheet.getCell(row, index + 1);
      cell.value = value;
      cell.font = headerFont;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = { top: thin, bottom: thin };
    });
    sheet.mergeCells(row, 3, row, 4);
    row += 1;
    sheet.mergeCells(row, 1, row, 9);
    sheet.getCell(row, 1).value = 'F.O.B. Value';
    sheet.getCell(row, 1).font = headerFont;
    sheet.getCell(row, 1).border = { bottom: thin };
    row += 1;

    for (const item of pageItems) {
      const startRow = row;
      const lines = Math.max(3, item.descriptions.length + 1);
      const endRow = row + Math.max(3, Math.ceil(lines / 2));
      sheet.mergeCells(startRow, 1, endRow, 1);
      sheet.mergeCells(startRow, 2, endRow, 2);
      sheet.mergeCells(startRow, 3, endRow, 4);
      sheet.mergeCells(startRow, 5, endRow, 5);
      sheet.mergeCells(startRow, 6, endRow, 6);
      sheet.mergeCells(startRow, 7, endRow, 7);
      sheet.mergeCells(startRow, 8, endRow, 8);
      sheet.mergeCells(startRow, 9, endRow, 9);
      sheet.getCell(startRow, 1).value = item.seq;
      sheet.getCell(startRow, 2).value = `Lot.No. : ${item.lotNo}\n${item.artNo}`;
      sheet.getCell(startRow, 3).value = item.descriptions.join('\n');
      sheet.getCell(startRow, 6).value = item.qty;
      sheet.getCell(startRow, 7).value = item.unit;
      sheet.getCell(startRow, 8).value = item.unitPrice;
      sheet.getCell(startRow, 9).value = item.qty * item.unitPrice;
      for (let col = 1; col <= 9; col += 1) {
        const cell = sheet.getCell(startRow, col);
        cell.font = bodyFont;
        cell.alignment = { vertical: 'top', horizontal: col >= 6 ? 'center' : 'left', wrapText: true };
        cell.border = { bottom: thin };
      }
      sheet.getCell(startRow, 8).numFmt = '$#,##0.00';
      sheet.getCell(startRow, 9).numFmt = '$#,##0.00';
      const imageEntry = getImageEntry(item);
      if (imageEntry?.file) {
        try {
          const dataUrl = await fileToDataUrl(imageEntry.file);
          const extension = imageEntry.extension === 'jpg' ? 'jpeg' : imageEntry.extension;
          const imageId = workbook.addImage({ base64: dataUrl, extension });
          sheet.addImage(imageId, { tl: { col: 4.12, row: startRow - 0.88 }, ext: { width: 92, height: Math.max(72, (endRow - startRow + 1) * 20 - 6) }, editAs: 'oneCell' });
        } catch {}
      }
      for (let r = startRow; r <= endRow; r += 1) sheet.getRow(r).height = 20;
      row = endRow + 1;
    }

    if (pageIndex === pages.length - 1) {
      row += 1;
      sheet.mergeCells(row, 1, row, 5);
      sheet.getCell(row, 1).value = `Total Quantity : ${totals.qty}`;
      sheet.getCell(row, 1).font = headerFont;
      sheet.mergeCells(row, 6, row, 8);
      sheet.getCell(row, 6).value = 'Sub Total:';
      sheet.getCell(row, 9).value = totals.subtotal;
      sheet.getCell(row, 9).numFmt = '$#,##0.00';
      row += 1;
      if (totals.discount > 0) {
        sheet.mergeCells(row, 6, row, 8);
        sheet.getCell(row, 6).value = 'DISCOUNT AMOUNT';
        sheet.getCell(row, 9).value = -totals.discount;
        sheet.getCell(row, 9).numFmt = '($#,##0.00)';
        row += 1;
      }
      sheet.mergeCells(row, 6, row, 8);
      sheet.getCell(row, 6).value = `Total : (${normalizeText($('#currency').value)})`;
      sheet.getCell(row, 6).font = headerFont;
      sheet.getCell(row, 9).value = totals.total;
      sheet.getCell(row, 9).numFmt = '$#,##0.00';
      sheet.getCell(row, 9).font = headerFont;
      row += 2;
      sheet.mergeCells(row, 1, row + 1, 9);
      sheet.getCell(row, 1).value = `Total Amount : ${normalizeText($('#currency').value)} ${numberToWords(totals.total)}`;
      sheet.getCell(row, 1).alignment = { wrapText: true };
      sheet.getCell(row, 1).font = bodyFont;
      row += 2;
      sheet.mergeCells(row, 1, row, 9);
      sheet.getCell(row, 1).value = `Remark : ${normalizeText($('#remark').value)}`;
      row += 2;
      sheet.mergeCells(row, 1, row, 4);
      sheet.getCell(row, 1).value = 'Vender Signature : __________________';
      sheet.mergeCells(row, 6, row, 9);
      sheet.getCell(row, 6).value = 'Accept By : __________________';
      row += 2;
      sheet.mergeCells(row, 1, row, 5);
      sheet.getCell(row, 1).value = 'UNIVERSE GEMS & JEWELLERY CO.';
      sheet.getCell(row, 1).font = headerFont;
    }
    sheet.mergeCells(row + 1, 7, row + 1, 9);
    sheet.getCell(row + 1, 7).value = `Page ${pageIndex + 1} of ${pages.length}`;
    sheet.getCell(row + 1, 7).alignment = { horizontal: 'right' };
    sheet.getCell(row + 1, 7).font = { name: 'Arial', size: 8 };
    if (pageIndex < pages.length - 1) {
      sheet.getRow(row + 1).addPageBreak();
      row += 3;
    } else row += 2;
  }
  sheet.pageSetup.printArea = `A1:I${row}`;
  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${safeFileName($('#invoiceNo').value)}.xlsx`);
}

function exportRemainingStockXlsx() {
  if (!window.XLSX) throw new Error('Excel 程式未載入。');
  const soldLots = new Set(state.items.map(item => item.lotNo));
  const remainingRows = state.stockRows.filter(row => !soldLots.has(normalizeCode(row[state.stockLotIndex])));
  const sheet = XLSX.utils.aoa_to_sheet([state.stockHeader, ...remainingRows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, 'Remaining Stock');
  const invoiceNo = safeFileName($('#invoiceNo').value);
  XLSX.writeFile(workbook, `Remaining_Stock_${invoiceNo}_${remainingRows.length}pcs.xlsx`);
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 3000);
}

$('#confirmBtn').addEventListener('click', async () => {
  if (!state.items.length) {
    showGlobalFeedback('—', '沒有貨品', 'error', 1800);
    return;
  }
  if (!normalizeText($('#customerName').value)) {
    showGlobalFeedback('—', '未選客戶', 'error', 1800);
    return;
  }
  const button = $('#confirmBtn');
  button.disabled = true;
  button.textContent = '正在輸出…';
  try {
    await exportInvoiceXlsx();
    await new Promise(resolve => setTimeout(resolve, 500));
    exportRemainingStockXlsx();
    const invoiceNo = normalizeText($('#invoiceNo').value);
    for (const item of state.items) {
      state.soldHistory.set(item.lotNo, { invoiceNo, customerCode: normalizeText($('#customerCode').value), date: todayIso() });
      state.products.delete(item.lotNo);
    }
    state.stockRows = state.stockRows.filter(row => !state.items.some(item => item.lotNo === normalizeCode(row[state.stockLotIndex])));
    state.items = [];
    renderItems();
    saveDraft();
    showGlobalFeedback(invoiceNo, '已輸出', 'success', 1500);
  } catch (error) {
    showGlobalFeedback('—', error.message || '輸出失敗', 'error', 2200);
  } finally {
    button.disabled = false;
    button.textContent = 'Confirm Invoice 並輸出 Excel';
  }
});

function saveDraft() {
  try {
    localStorage.setItem('universe-invoice-v07-draft', JSON.stringify({
      customerCode: $('#customerCode').value,
      customerName: $('#customerName').value,
      customerAddress: $('#customerAddress').value,
      salesRate: $('#salesRate').value,
      currency: $('#currency').value,
      terms: $('#terms').value,
      shipmentMethod: $('#shipmentMethod').value,
      invoiceNo: $('#invoiceNo').value,
      invoiceDate: $('#invoiceDate').value,
      discountAmount: $('#discountAmount').value,
      remark: $('#remark').value,
      items: state.items
    }));
  } catch {}
}

function restoreDraft() {
  try {
    const draft = JSON.parse(localStorage.getItem('universe-invoice-v07-draft') || 'null');
    if (!draft) return;
    const fields = ['customerCode', 'customerName', 'customerAddress', 'salesRate', 'currency', 'terms', 'shipmentMethod', 'invoiceNo', 'invoiceDate', 'discountAmount', 'remark'];
    for (const field of fields) if (draft[field] !== undefined && $(`#${field}`)) $(`#${field}`).value = draft[field];
    if (Array.isArray(draft.items)) state.items = draft.items;
  } catch {}
}

restoreDraft();
renderCustomerSummary();
renderItems();
updateSummary();
if ('serviceWorker' in navigator) window.addEventListener('load', () => navigator.serviceWorker.register('./sw.js').catch(() => {}));
