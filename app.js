/* SchoolCollect · Local-first. Student and payment data never leave this browser. */
const DB_NAME = 'so-thu-hoc-sinh-local';
const DB_VERSION = 1;
const FEES = [
  { key: 'insurance', label: 'Bảo hiểm y tế (BHYT)', short: 'BHYT' },
  { key: 'mandatory', label: 'Bảo hiểm bắt buộc (BHTT)', short: 'BHTT' },
  { key: 'service', label: 'Dịch vụ khác', short: 'Dịch vụ' },
  { key: 'other', label: 'Chưa phân loại', short: 'Chưa rõ' }
];
let db;
let activeImport = null;
let toastTimer;

const $ = (s, root = document) => root.querySelector(s);
const $$ = (s, root = document) => [...root.querySelectorAll(s)];
const num = value => Math.round(Number(value) || 0);
const money = n => new Intl.NumberFormat('vi-VN').format(num(n)) + ' ₫';
const dateTime = value => value ? new Intl.DateTimeFormat('vi-VN', { dateStyle: 'short', timeStyle: 'short' }).format(new Date(value)) : '—';
const escapeHTML = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const slug = value => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/đ/g, 'd').replace(/[^a-z0-9]+/g, ' ').trim();

function openDatabase() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const d = req.result;
      if (!d.objectStoreNames.contains('students')) d.createObjectStore('students', { keyPath: 'code' });
      if (!d.objectStoreNames.contains('transactions')) d.createObjectStore('transactions', { keyPath: 'id' });
      if (!d.objectStoreNames.contains('history')) d.createObjectStore('history', { keyPath: 'id' });
      if (!d.objectStoreNames.contains('meta')) d.createObjectStore('meta', { keyPath: 'key' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
function request(store, method, value) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, method === 'clear' ? 'readwrite' : value === undefined ? 'readonly' : 'readwrite');
    const req = value === undefined ? tx.objectStore(store)[method]() : tx.objectStore(store)[method](value);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
function all(store) { return request(store, 'getAll'); }
async function putMany(store, entries) {
  if (!entries.length) return;
  await new Promise((resolve, reject) => {
    const tx = db.transaction(store, 'readwrite');
    const objectStore = tx.objectStore(store);
    entries.forEach(entry => objectStore.put(entry));
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}
async function clearAll() {
  await new Promise((resolve, reject) => {
    const tx = db.transaction(['students', 'transactions', 'history', 'meta'], 'readwrite');
    ['students', 'transactions', 'history', 'meta'].forEach(name => tx.objectStore(name).clear());
    tx.oncomplete = resolve; tx.onerror = () => reject(tx.error);
  });
}

function parseCsv(text, delimiter) {
  const rows = []; let row = []; let value = ''; let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) { if (ch === '"' && text[i + 1] === '"') { value += '"'; i++; } else if (ch === '"') quoted = false; else value += ch; }
    else if (ch === '"') quoted = true;
    else if (ch === delimiter) { row.push(value); value = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(value); rows.push(row); row = []; value = ''; }
    else value += ch;
  }
  if (value.length || row.length) { row.push(value); rows.push(row); }
  while (rows.length && rows[rows.length - 1].every(cell => !String(cell).trim())) rows.pop();
  return rows;
}
function bestDelimiter(text) {
  const first = text.replace(/^\uFEFF/, '').split(/\r?\n/, 1)[0] || '';
  return [',', ';', '\t'].map(d => [d, first.split(d).length]).sort((a, b) => b[1] - a[1])[0][0];
}
async function readZipEntry(bytes, entry) {
  let offset = entry.localOffset;
  if (bytes[offset] !== 0x50 || bytes[offset + 1] !== 0x4b) throw new Error('Tệp Excel không đúng định dạng ZIP/XLSX.');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const nameLength = view.getUint16(offset + 26, true), extraLength = view.getUint16(offset + 28, true);
  offset += 30 + nameLength + extraLength;
  const data = bytes.slice(offset, offset + entry.compressedSize);
  if (entry.method === 0) return data;
  if (entry.method !== 8 || !('DecompressionStream' in window)) throw new Error('Trình duyệt chưa hỗ trợ đọc XLSX ngoại tuyến. Hãy dùng Excel .xlsx hoặc xuất CSV.');
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
async function parseXlsx(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65558); i--) if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('Không tìm thấy cấu trúc XLSX trong tệp.');
  const count = view.getUint16(eocd + 10, true), cdOffset = view.getUint32(eocd + 16, true);
  const entries = new Map(); let p = cdOffset;
  for (let i = 0; i < count; i++) {
    if (view.getUint32(p, true) !== 0x02014b50) break;
    const method = view.getUint16(p + 10, true), compressedSize = view.getUint32(p + 20, true);
    const nameLength = view.getUint16(p + 28, true), extraLength = view.getUint16(p + 30, true), commentLength = view.getUint16(p + 32, true);
    const localOffset = view.getUint32(p + 42, true);
    const name = new TextDecoder().decode(bytes.slice(p + 46, p + 46 + nameLength));
    entries.set(name, { method, compressedSize, localOffset }); p += 46 + nameLength + extraLength + commentLength;
  }
  const getText = async name => entries.has(name) ? new TextDecoder().decode(await readZipEntry(bytes, entries.get(name))) : '';
  const workbook = new DOMParser().parseFromString(await getText('xl/workbook.xml'), 'application/xml');
  const relations = new DOMParser().parseFromString(await getText('xl/_rels/workbook.xml.rels'), 'application/xml');
  const sharedXml = await getText('xl/sharedStrings.xml');
  const sharedDoc = sharedXml ? new DOMParser().parseFromString(sharedXml, 'application/xml') : null;
  const shared = sharedDoc ? [...sharedDoc.querySelectorAll('si')].map(si => [...si.querySelectorAll('t')].map(t => t.textContent).join('')) : [];
  const headerWords=['ma hoc sinh','ma hs theo khoan nop','ma khach hang','ho va ten','ten khach hang','khoan nop','so tien','so hoa don','ngay giao dich','trang thai giao dich'];
  const candidates=[];
  for (const sheetNode of workbook.querySelectorAll('sheet')) {
    const relId=sheetNode.getAttribute('r:id');
    const rel=[...relations.querySelectorAll('Relationship')].find(node=>node.getAttribute('Id')===relId);
    const target=rel?.getAttribute('Target')||'worksheets/sheet1.xml';
    const sheetPath=target.startsWith('/')?target.slice(1):`xl/${target.replace(/^\.\//,'')}`;
    const sheetXml=await getText(sheetPath);if(!sheetXml)continue;
    const doc=new DOMParser().parseFromString(sheetXml,'application/xml');
    const rows=[...doc.querySelectorAll('sheetData row')].map(rowNode=>{
      const cells=[];
      for(const cell of rowNode.querySelectorAll(':scope > c')){
        const ref=cell.getAttribute('r')||'';const col=[...ref.matchAll(/[A-Z]+/g)][0]?.[0]||'A';
        let index=0;for(const char of col)index=index*26+char.charCodeAt(0)-64;index--;
        const type=cell.getAttribute('t');let val=cell.querySelector('v')?.textContent??'';
        if(type==='s')val=shared[Number(val)]??'';else if(type==='inlineStr')val=[...cell.querySelectorAll('is t')].map(t=>t.textContent).join('');
        cells[index]=val;
      }
      return cells.map(value=>value??'');
    });
    const headers=(rows[0]||[]).map(x=>slug(x));
    const score=headers.reduce((n,h)=>n+(headerWords.some(w=>h===w||h.includes(w))?1:0),0);
    const dataRows=rows.slice(1).filter(row=>row.some(value=>String(value??'').trim())).length;
    candidates.push({name:sheetNode.getAttribute('name')||'Trang tính',rows,score,dataRows});
  }
  const chosen=candidates.sort((a,b)=>b.score-a.score||b.dataRows-a.dataRows)[0];
  if(!chosen)throw new Error('Không tìm thấy trang tính có dữ liệu trong file Excel.');
  chosen.rows.sourceSheet=chosen.name;
  return chosen.rows;
}
async function readRows(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'xlsx') return parseXlsx(file);
  const buffer = await file.arrayBuffer(); let text;
  try { text = new TextDecoder('utf-8', { fatal: true }).decode(buffer); }
  catch { text = new TextDecoder('windows-1258').decode(buffer); }
  return parseCsv(text.replace(/^\uFEFF/, ''), bestDelimiter(text));
}
function guessColumn(headers, field) {
  const normal = headers.map(h => slug(h));
  const patterns = {
    code: ['ma hoc sinh', 'ma hs', 'student code', 'student id', 'ma dinh danh', 'ma so hs'],
    name: ['ho va ten', 'ten hoc sinh', 'ho ten', 'student name', 'ten'],
    className: ['lop', 'khoi lop', 'class', 'ma lop'],
    due: ['so tien phai thu', 'phai thu', 'hoc phi', 'muc thu', 'tien thu'],
    dueInsurance: ['bhyt', 'bao hiem y te', 'so tien bao hiem y te', 'bao hiem y te phai thu', 'insurance'],
    dueMandatory: ['bhtt', 'bao hiem bat buoc', 'bat buoc'],
    dueService: ['dich vu khac phai thu', 'so tien dich vu khac', 'dich vu khac', 'service due'],
    dueParking: ['gui xe phai thu', 'so tien gui xe', 'phi gui xe', 'gui xe'],
    dueWater: ['nuoc uong phai thu', 'so tien nuoc uong', 'phi nuoc uong', 'nuoc uong'],
    studentFeeAmount: ['so tien', 'so tien phai thu', 'fee amount'],
    studentFeeCategory: ['khoan nop', 'loai khoan nop', 'fee item'],
    paymentCode: ['ma hs theo khoan nop', 'ma khach hang', 'ma thanh toan', 'customer code'],
    amount: ['so tien giao dich', 'amount', 'credit', 'ghi co', 'so tien'],
    txnId: ['so hoa don', 'ma hoa don', 'ma giao dich', 'transaction id', 'ma tham chieu', 'reference', 'trace', 'so but toan'],
    date: ['ngay giao dich', 'thoi gian', 'transaction date', 'ngay hach toan', 'ngay'],
    content: ['noi dung chuyen khoan', 'ten khach hang', 'noi dung', 'dien giai', 'transaction content', 'description', 'chi tiet'],
    studentCode: ['ma khach hang', 'ma hs theo khoan nop', 'ma hoc sinh', 'ma hs', 'student code', 'student id', 'ma dinh danh'],
    feeCategory: ['khoan nop', 'khoan thu', 'loai khoan thu', 'danh muc thu', 'fee category', 'fee type'],
    bankStatus: ['trang thai giao dich', 'trang thai', 'status']
  }[field] || [];
  const exact = normal.findIndex(h => patterns.includes(h));
  if (exact >= 0) return String(exact);
  return String(normal.findIndex(h => h && patterns.some(p => h.includes(p) || p.includes(h))));
}
function fieldSelect(id, label, headers, field, required = false) {
  const guessed = guessColumn(headers, field);
  return `<div class="mapping-field"><label for="${id}">${label}${required ? ' <em>*</em>' : ''}</label><select id="${id}" data-field="${field}"><option value="-1">— Không chọn —</option>${headers.map((h, i) => `<option value="${i}" ${String(i) === guessed ? 'selected' : ''}>${escapeHTML(h || `Cột ${i + 1}`)}</option>`).join('')}</select></div>`;
}
function previewHtml(headers, rows) {
  return `<div class="preview-box"><strong>Xem trước 3 dòng đầu</strong><div class="table-wrap"><table><thead><tr>${headers.map(h => `<th>${escapeHTML(h)}</th>`).join('')}</tr></thead><tbody>${rows.slice(1, 4).map(row => `<tr>${headers.map((_, i) => `<td>${escapeHTML(row[i] || '')}</td>`).join('')}</tr>`).join('')}</tbody></table></div></div>`;
}
function openImportModal(kind, file, rows) {
  if (!rows.length || rows.length < 2) return toast('File không có dòng tiêu đề hoặc dữ liệu.', true);
  const headers = rows[0].map((h, i) => String(h || `Cột ${i + 1}`).trim());
  const longFormat=kind==='students'&&['code','paymentCode','studentFeeCategory','studentFeeAmount'].every(field=>Number(guessColumn(headers,field))>=0);
  activeImport = { kind, file, rows, headers, longFormat };
  $('#modalEyebrow').textContent = kind === 'students' ? 'DANH SÁCH HỌC SINH' : 'BÁO CÁO THU';
  $('#modalTitle').textContent = kind === 'students' ? 'Ghép cột danh sách học sinh' : 'Ghép cột báo cáo thu';
  const summary = `<div class="file-summary"><span class="file-badge">${file.name.toLowerCase().endsWith('.xlsx') ? 'XLSX' : 'CSV'}</span><div><strong>${escapeHTML(file.name)}</strong><small>${rows.length - 1} dòng dữ liệu · ${headers.length} cột${rows.sourceSheet?` · Trang tính: ${escapeHTML(rows.sourceSheet)}`:''}</small></div></div>`;
  const fields = kind === 'students'
    ? longFormat
      ? `<div class="mapping-grid">${fieldSelect('map-code','Mã học sinh',headers,'code',true)}${fieldSelect('map-payment-code','Mã HS theo khoản nộp',headers,'paymentCode',true)}${fieldSelect('map-name','Họ và tên',headers,'name',true)}${fieldSelect('map-class','Lớp',headers,'className',true)}${fieldSelect('map-fee-category','Khoản nộp (BHYT/BHTT)',headers,'studentFeeCategory',true)}${fieldSelect('map-fee-amount','Số tiền phải thu',headers,'studentFeeAmount',true)}</div><p class="mapping-intro">Mỗi học sinh có một dòng BHYT và một dòng BHTT. Web tự ghép hai dòng theo mã học sinh và giữ mã từng khoản để đối soát với cột “Mã khách hàng” của ngân hàng.</p>`
      : `<div class="mapping-grid">${fieldSelect('map-code','Mã học sinh',headers,'code',true)}${fieldSelect('map-name','Họ và tên',headers,'name',true)}${fieldSelect('map-class','Lớp',headers,'className')}${fieldSelect('map-due','Tổng phải thu (nếu có)',headers,'due')}${fieldSelect('map-insurance','Phải thu · BHYT',headers,'dueInsurance')}${fieldSelect('map-mandatory','Phải thu · BHTT',headers,'dueMandatory')}${fieldSelect('map-service','Phải thu · Dịch vụ khác',headers,'dueService')}${fieldSelect('map-parking','Phải thu · Gửi xe',headers,'dueParking')}${fieldSelect('map-water','Phải thu · Nước uống',headers,'dueWater')}</div><p class="mapping-intro">Nhập riêng số phải thu BHYT và BHTT. Web chỉ tính khoản đã thu khi số tiền khớp chính xác.</p>`
    : `<div class="mapping-grid">${fieldSelect('map-amount','Số tiền giao dịch',headers,'amount',true)}${fieldSelect('map-student','Mã khách hàng / mã khoản nộp',headers,'studentCode')}${fieldSelect('map-txn','Số hóa đơn / mã giao dịch',headers,'txnId')}${fieldSelect('map-date','Ngày giao dịch',headers,'date')}${fieldSelect('map-content','Tên khách hàng / nội dung',headers,'content')}${fieldSelect('map-category','Khoản thu (nếu có)',headers,'feeCategory')}${fieldSelect('map-bank-status','Trạng thái giao dịch',headers,'bankStatus')}</div><p class="mapping-intro">Mã khách hàng kết thúc bằng YT hoặc TT sẽ được ghép với đúng khoản BHYT hoặc BHTT. Chỉ giao dịch thành công, đúng mã khoản và đúng số tiền mới được tính đã thu.</p>`;
  $('#modalBody').innerHTML = `${summary}${fields}${previewHtml(headers, rows)}`;
  $('#modalConfirm').textContent = kind === 'students' ? 'Nhập danh sách' : 'Nhập & đối soát';
  $('#modalBackdrop').classList.add('open');
}
function closeModal() { $('#modalBackdrop').classList.remove('open'); activeImport = null; }
function getMap() { return Object.fromEntries($$('#modalBody select').map(s => [s.dataset.field, Number(s.value)])); }
function cell(row, map, field) { const index = map[field]; return index === undefined || index < 0 ? '' : String(row[index] ?? '').trim(); }
function parseAmount(value) {
  let s = String(value || '').replace(/[^\d,.-]/g, '').trim();
  if (!s) return 0;
  if (s.includes(',') && s.includes('.')) s = s.lastIndexOf(',') > s.lastIndexOf('.') ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (s.includes(',')) s = /^-?\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
  else if ((s.match(/\./g) || []).length > 1 || /^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  return Math.round(Number(s) || 0);
}
function normalizeDate(value) {
  if (!value) return '';
  if (/^\d+(\.\d+)?$/.test(value) && Number(value) > 15000 && Number(value) < 90000) return new Date(Date.UTC(1899, 11, 30) + Number(value) * 86400000).toISOString().slice(0, 10);
  const m = value.match(/^(\d{1,2})[/. -](\d{1,2})[/. -](\d{2,4})/);
  if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  const d = new Date(value); return Number.isNaN(d.getTime()) ? value : d.toISOString().slice(0, 10);
}
function isSuccessfulBankStatus(value) {
  const status=slug(value);
  if(!status)return true;
  if(/khong|that bai|huy|tu choi|failed|cancel|pending/.test(status))return false;
  return /thanh cong|success|successful|completed|hoan tat/.test(status);
}
function getFeeKey(value, content = '') {
  const raw = slug(value || '');
  if (/bhtt|bao hiem bat buoc|bat buoc|mandatory/.test(raw)) return 'mandatory';
  if (/bhyt|bao hiem y te|bao hiem|insurance|y te/.test(raw)) return 'insurance';
  if (/dich vu|gui xe|nuoc uong|ban tru|parking|service|an uong/.test(raw)) return 'service';
  const s = slug(content || '');
  if (/bhtt|bao hiem bat buoc|bat buoc|mandatory/.test(s)) return 'mandatory';
  if (/bhyt|bao hiem y te|bao hiem|insurance|y te/.test(s)) return 'insurance';
  if (/dich vu|gui xe|nuoc uong|ban tru|parking|service|an uong/.test(s)) return 'service';
  return 'other';
}
function feeCategoryFromPaymentCode(value) {
  const code=String(value||'').trim().toUpperCase().replace(/\s+/g,'');
  if(code.endsWith('BHTT')||code.endsWith('TT'))return 'mandatory';
  if(code.endsWith('BHYT')||code.endsWith('YT'))return 'insurance';
  return 'other';
}
function getFeeLabel(key) { return FEES.find(f => f.key === key)?.label || 'Chưa phân loại'; }
function serviceDetail(raw, content) {
  const rawText = String(raw || '').trim();
  if (rawText && !/^(dich vu khac|dich vu|bao hiem)$/i.test(slug(rawText))) return rawText;
  let detail = String(content || '').trim().replace(/^(thu|nop|chuyen khoan|thanh toan)\s+/i, '');
  detail = detail.replace(/\s*[-–—]\s*(ma\s*)?hs\s*[a-z0-9-]+.*$/i, '').replace(/\s*[-–—]\s*hs\d+.*$/i, '').trim();
  return detail ? detail.charAt(0).toLocaleUpperCase('vi-VN') + detail.slice(1) : 'Dịch vụ khác';
}
function studentFeesFromRow(row, map) {
  const hasServiceDetails=map.dueParking>=0||map.dueWater>=0;
  const fields=[['insurance',getFeeLabel('insurance'),'dueInsurance'],['mandatory',getFeeLabel('mandatory'),'dueMandatory']];
  if(hasServiceDetails){fields.push(['service','Gửi xe','dueParking'],['service','Nước uống','dueWater']);}
  else fields.push(['service','Dịch vụ khác','dueService']);
  const dueItems=fields.filter(([, ,field])=>map[field]>=0).map(([category,name,field])=>({id:`${category}:${slug(name)}`,category,name,amount:parseAmount(cell(row,map,field))})).filter(item=>item.amount>0);
  const hasBreakdown=fields.some(([, ,field])=>map[field]>=0);
  if(hasServiceDetails&&map.dueService>=0){
    const serviceTotal=parseAmount(cell(row,map,'dueService'));
    const detailedService=dueItems.filter(item=>item.category==='service').reduce((sum,item)=>sum+item.amount,0);
    if(serviceTotal>detailedService)dueItems.push({id:'service:other',category:'service',name:'Dịch vụ khác',amount:serviceTotal-detailedService});
  }
  const statedDue=parseAmount(cell(row,map,'due'));
  const detailedTotal=dueItems.reduce((sum,item)=>sum+item.amount,0);
  if(hasBreakdown&&statedDue>detailedTotal)dueItems.push({id:'other:unclassified',category:'other',name:'Chưa phân loại',amount:statedDue-detailedTotal});
  if(!hasBreakdown&&statedDue>0)dueItems.push({id:'other:unclassified',category:'other',name:'Chưa phân loại',amount:statedDue});
  const due=hasBreakdown?Math.max(statedDue,detailedTotal):statedDue;
  const dueByCategory=dueItems.reduce((out,item)=>(out[item.category]=(out[item.category]||0)+item.amount,out),{});
  return {due,dueItems,dueByCategory,hasFeeBreakdown:hasBreakdown};
}
function studentsFromFeeRows(dataRows,map,now) {
  const grouped=new Map(),paymentOwners=new Map();
  for(let i=0;i<dataRows.length;i++) {
    const row=dataRows[i],code=cell(row,map,'code'),name=cell(row,map,'name'),className=cell(row,map,'className');
    const paymentCode=cell(row,map,'paymentCode').toUpperCase().replace(/\s+/g,'');
    const rawCategory=cell(row,map,'studentFeeCategory'),category=getFeeKey(rawCategory),amount=parseAmount(cell(row,map,'studentFeeAmount'));
    if(!code||!name||!className||!paymentCode||!amount||!['insurance','mandatory'].includes(category))
      throw new Error(`Dòng ${i+2}: thiếu mã, tên, lớp, khoản BHYT/BHTT hoặc số tiền hợp lệ.`);
    if(feeCategoryFromPaymentCode(paymentCode)!==category)throw new Error(`Dòng ${i+2}: mã khoản nộp ${paymentCode} không khớp loại ${rawCategory}.`);
    let student=grouped.get(slug(code));
    if(!student){student={code,name,className,due:0,dueItems:[],dueByCategory:{insurance:0,mandatory:0,service:0,other:0},updatedAt:now,hasFeeBreakdown:true};grouped.set(slug(code),student);}
    if(student.name!==name||student.className!==className)throw new Error(`Mã học sinh ${code} có tên hoặc lớp không thống nhất trong file.`);
    const normalizedPaymentCode=slug(paymentCode);
    const owner=paymentOwners.get(normalizedPaymentCode);
    if(owner&&owner!==slug(code))throw new Error(`Mã HS theo khoản nộp ${paymentCode} bị gán cho nhiều học sinh.`);
    paymentOwners.set(normalizedPaymentCode,slug(code));
    if(student.dueItems.some(item=>slug(item.paymentCode)===normalizedPaymentCode))throw new Error(`Mã HS theo khoản nộp ${paymentCode} bị lặp. Hãy kiểm tra file trước khi nhập.`);
    student.dueItems.push({id:`payment:${normalizedPaymentCode}`,paymentCode,category,name:getFeeLabel(category),amount});
    student.due+=amount;student.dueByCategory[category]+=amount;
  }
  for(const student of grouped.values()){
    const counts=student.dueItems.reduce((out,item)=>(out[item.category]=(out[item.category]||0)+1,out),{});
    if(counts.insurance!==1||counts.mandatory!==1)throw new Error(`Học sinh ${student.code} cần đúng một dòng BHYT và một dòng BHTT trong file.`);
  }
  if(!grouped.size)throw new Error('Không tìm thấy dòng học sinh hợp lệ trong file.');
  return [...grouped.values()];
}
async function confirmImport() {
  if (!activeImport) return;
  const { kind, file, rows, longFormat } = activeImport; const map = getMap();
  if (kind === 'students' && (map.code < 0 || map.name < 0 || (longFormat && [map.paymentCode,map.studentFeeCategory,map.studentFeeAmount,map.className].some(x=>x<0)))) return toast('Hãy chọn đủ cột mã, tên, lớp, mã khoản nộp, loại khoản và số tiền.', true);
  if (kind === 'bank' && map.amount < 0) return toast('Hãy chọn cột số tiền giao dịch.', true);
  const dataRows = rows.slice(1).filter(row => row.some(v => String(v ?? '').trim()));
  const now = new Date().toISOString(); let summary;
  if (kind === 'students') {
    let items;
    try { items = longFormat ? studentsFromFeeRows(dataRows,map,now) : dataRows.map(row => {
      return { code:cell(row,map,'code'), name:cell(row,map,'name'), className:cell(row,map,'className'), ...studentFeesFromRow(row,map), updatedAt:now };
    }).filter(s => s.code && s.name); }
    catch(error) { return toast(error.message||'Danh sách học sinh chưa đúng định dạng.',true); }
    await putMany('students', items);
    const codes = new Map((await all('students')).map(s => [slug(s.code), s]));
    const knownTransactions = await all('transactions');
    const rematched = knownTransactions.map(t => {
      const student = t.reportedStudentCode ? codes.get(slug(t.reportedStudentCode)) : null;
      return student ? { ...t, studentCode:student.code, studentName:student.name, matched:true } : { ...t, studentCode:'', studentName:'', matched:false };
    });
    await putMany('transactions', rematched);
    const feeItems=items.reduce((sum,s)=>sum+studentDueItems(s).length,0);
    summary = { rows:dataRows.length, imported:items.length, detail:`${items.length.toLocaleString('vi-VN')} học sinh · ${feeItems.toLocaleString('vi-VN')} món phải thu được nhập/cập nhật` };
    $('#studentLastImport').textContent = `Gần nhất: ${file.name} · ${items.length.toLocaleString('vi-VN')} học sinh`;
  } else {
    const students = await all('students'); const byCode = new Map(students.map(s => [slug(s.code), s]));
    const byPaymentCode=new Map();students.forEach(s=>studentDueItems(s).forEach(item=>{if(item.paymentCode)byPaymentCode.set(slug(item.paymentCode),s);}));
    const items = dataRows.map((row, i) => {
      const amount = parseAmount(cell(row, map, 'amount')); if (!amount) return null;
      const code = cell(row, map, 'studentCode'); const student = code ? byCode.get(slug(code))||byPaymentCode.get(slug(code)) : null;
      const date = normalizeDate(cell(row, map, 'date')); const content = cell(row, map, 'content');
      const rawCategory = cell(row, map, 'feeCategory'); const codeCategory=feeCategoryFromPaymentCode(code); const feeCategory = codeCategory!=='other'?codeCategory:getFeeKey(rawCategory, content);
      const feeDetail = feeCategory === 'service' ? serviceDetail(rawCategory, content) : getFeeLabel(feeCategory);
      const bankStatus=cell(row,map,'bankStatus');
      const ref = cell(row, map, 'txnId'); const normalizedRef = slug(ref);
      const id = normalizedRef ? `ref:${normalizedRef}` : `row:${slug(date)}:${amount}:${slug(code || content)}:${i}`;
      return { id, ref, date, content, amount, feeCategory, feeDetail, reportedStudentCode:code, reportedPaymentCode:code, bankStatus, studentCode:student?.code || '', studentName:student?.name || '', sourceFile:file.name, importedAt:now, matched:!!student };
    }).filter(Boolean);
    const oldItems=await all('transactions'),oldById=new Map(oldItems.map(t=>[t.id,t]));
    const additions=[],updates=[];let duplicates=0;
    for(const t of items){const old=oldById.get(t.id);if(!old){additions.push(t);oldById.set(t.id,t);}else if(!isSuccessfulBankStatus(old.bankStatus)&&isSuccessfulBankStatus(t.bankStatus)){updates.push(t);oldById.set(t.id,t);}else duplicates++;}
    await putMany('transactions',[...additions,...updates]);
    summary = { rows:dataRows.length, imported:items.length, detail:`${additions.length} giao dịch mới · ${updates.length} giao dịch cập nhật thành công · ${duplicates} dòng trùng được bỏ qua` };
    $('#bankLastImport').textContent = `Gần nhất: ${file.name} · ${additions.length.toLocaleString('vi-VN')} giao dịch mới`;
  }
  await request('history', 'put', { id:crypto.randomUUID(), kind:kind === 'students' ? 'Danh sách học sinh' : 'Báo cáo thu', fileName:file.name, rows:summary.rows, imported:summary.imported, detail:summary.detail, at:now });
  closeModal(); await refresh(); toast(summary.detail);
}
function toast(message, error = false) {
  const el = $('#toast'); el.textContent = message; el.classList.toggle('error', error); el.classList.add('show');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => el.classList.remove('show'), 3500);
}
async function chooseFile(kind, file) {
  if (!file) return;
  try { const rows = await readRows(file); openImportModal(kind, file, rows); }
  catch (error) { console.error(error); toast(error.message || 'Không đọc được file này.', true); }
}
function setPage(page) {
  $$('.page').forEach(p => p.classList.toggle('active', p.id === `page-${page}`));
  $$('.nav-item[data-page]').forEach(b => b.classList.toggle('active', b.dataset.page === page));
  const labels = {
    dashboard:['Tổng quan','Theo dõi tiến độ thu theo thời gian thực trên thiết bị này'],
    students:['Học sinh','Danh sách và số phải thu chi tiết theo từng học sinh'],
    fees:['Khoản thu','Theo dõi riêng bảo hiểm, dịch vụ khác và từng nội dung dịch vụ'],
    qr:['Tạo mã QR','Tạo QR thanh toán theo từng món thu của từng học sinh'],
    imports:['Nhập dữ liệu','Cập nhật danh sách học sinh và báo cáo thu gần nhất'],
    history:['Tra cứu & báo cáo','Lịch sử các lần nhập dữ liệu trên thiết bị này'],
    settings:['Sao lưu & cài đặt','Bảo vệ và chuyển dữ liệu theo quy trình của trường']
  };
  $('#topTitle').textContent = labels[page]?.[0] || 'SchoolCollect';
  $('#topSubtitle').textContent = labels[page]?.[1] || '';
  $('#sidebar').classList.remove('open'); window.scrollTo({ top:0, behavior:'smooth' });
}
function transactionCategory(t) {
  const codeCategory=feeCategoryFromPaymentCode(t.reportedPaymentCode||t.reportedStudentCode);
  if(codeCategory!=='other')return codeCategory;
  return ['insurance','mandatory','service','other'].includes(t.feeCategory) ? t.feeCategory : getFeeKey('', t.content);
}
function studentDueItems(student) {
  if (Array.isArray(student.dueItems)) return student.dueItems.filter(x=>num(x.amount)>0).map(x=>({...x,amount:num(x.amount)}));
  if (student.dueByCategory && Object.keys(student.dueByCategory).length) {
    const items=[];
    for (const key of ['insurance','mandatory','service','other']) { const amount=num(student.dueByCategory[key]); if(amount) items.push({id:`legacy:${key}`,category:key,name:key==='other'?'Chưa phân loại':getFeeLabel(key),amount,legacy:true}); }
    const gap=Math.max(0,num(student.due)-items.reduce((sum,x)=>sum+x.amount,0));
    if(gap)items.push({id:'legacy:unclassified',category:'other',name:'Chưa phân loại',amount:gap});
    return items;
  }
  return num(student.due)>0?[{id:'legacy:unclassified',category:'other',name:'Chưa phân loại',amount:num(student.due)}]:[];
}
function studentDueByCategory(student) {
  return studentDueItems(student).reduce((out,item)=>(out[item.category]=(out[item.category]||0)+item.amount,out),{insurance:0,mandatory:0,service:0,other:0});
}
function reconcileTransactions(students, transactions) {
  const byCode=new Map(students.map(s=>[slug(s.code),s]));const byPaymentCode=new Map();
  students.forEach(student=>studentDueItems(student).forEach(item=>{if(item.paymentCode)byPaymentCode.set(slug(item.paymentCode),{student,item});}));
  const claimed=new Set();
  return [...transactions].sort((a,b)=>(a.date||a.importedAt||'').localeCompare(b.date||b.importedAt||'')).map(t=>{
    const reportedCode=t.reportedPaymentCode||t.reportedStudentCode||t.studentCode||'';
    const alias=byPaymentCode.get(slug(reportedCode));const student=byCode.get(slug(reportedCode))||alias?.student;
    if(!isSuccessfulBankStatus(t.bankStatus))return {...t,studentCode:student?.code||'',studentName:student?.name||'',matched:false,paymentStatus:'bank_not_successful',feeCategory:alias?.item.category||transactionCategory(t)};
    if(!student)return {...t,studentCode:'',studentName:'',matched:false,paymentStatus:reportedCode?'unmatched':'missing_code'};
    const studentItems=studentDueItems(student), category=alias?.item.category||transactionCategory(t), amount=num(t.amount);
    const candidates=alias?[alias.item]:category==='other'?[]:studentItems.filter(item=>!item.legacy&&item.category===category&&item.amount===amount);
    if(!candidates.length) {
      const categoryItems=studentItems.filter(item=>item.category===category);
      const status=categoryItems.some(item=>item.legacy)?'missing_category':categoryItems.length?'amount_mismatch':studentItems.length?'missing_category':'no_due';
      return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:status,feeCategory:category};
    }
    if(candidates.length===1&&candidates[0].amount!==amount)return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:'amount_mismatch',feeCategory:category};
    const reportedDetail=slug(t.feeDetail||'');
    const byName=reportedDetail?candidates.filter(item=>slug(item.name)===reportedDetail):[];
    const possible=byName.length?byName:candidates;
    if(possible.length!==1)return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:'ambiguous',feeCategory:category};
    if(possible[0].amount!==amount)return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:'amount_mismatch',feeCategory:category};
    const item=possible[0], key=`${student.code}|${item.id}`;
    if(claimed.has(key))return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:'duplicate',feeCategory:item.category};
    claimed.add(key);
    return {...t,studentCode:student.code,studentName:student.name,matched:true,paymentStatus:'valid',feeCategory:item.category,feeDetail:item.name,matchedDueItem:item.name,matchedDueItemId:item.id};
  });
}
function feeSummaries(students, transactions) {
  const summaries = Object.fromEntries(FEES.map(f => [f.key, { ...f, due:0, paid:0, remain:0, dueItems:0, paidItems:0, pct:0 }]));
  const paidByStudent = new Map(), paidDueItems=new Set();
  transactions.filter(t => t.paymentStatus==='valid' && t.studentCode).forEach(t => {
    const key = transactionCategory(t); const amount = num(t.amount);
    summaries[key].paid += amount;
    if (!paidByStudent.has(t.studentCode)) paidByStudent.set(t.studentCode, {});
    const map = paidByStudent.get(t.studentCode); map[key] = (map[key] || 0) + amount;
    paidDueItems.add(`${t.studentCode}|${t.matchedDueItemId||slug(t.matchedDueItem||'')}`);
  });
  students.forEach(student => {
    studentDueItems(student).forEach(item=>{const s=summaries[item.category]||summaries.other;s.due+=item.amount;s.dueItems++;if(paidDueItems.has(`${student.code}|${item.id||slug(item.name)}`))s.paidItems++;});
  });
  Object.values(summaries).forEach(s => { s.remain = Math.max(0, s.due - s.paid); s.pct = s.due ? Math.min(100, Math.round(s.paid / s.due * 100)) : 0; });
  return { summaries, paidByStudent };
}
function totals(students, transactions) {
  const { summaries, paidByStudent } = feeSummaries(students, transactions);
  const due = Object.values(summaries).reduce((sum,s)=>sum+s.due,0);
  const paid = transactions.filter(t=>t.paymentStatus==='valid').reduce((sum,t)=>sum+num(t.amount),0);
  const remain = Math.max(0,due-paid);
  const dueItems = Object.values(summaries).reduce((sum,s)=>sum+s.dueItems,0);
  const paidItems = Object.values(summaries).reduce((sum,s)=>sum+s.paidItems,0);
  return { due, paid, remain, dueItems, paidItems, unpaidItems:Math.max(0,dueItems-paidItems), paidByStudent, summaries, pct:due ? Math.min(100,Math.round(paid/due*100)) : 0 };
}
function renderStudents(students, transactions) {
  const { paidByStudent } = totals(students, transactions); const query = slug($('#studentSearch')?.value || '');
  const filtered = students.filter(s => !query || slug(`${s.code} ${s.name} ${s.className}`).includes(query));
  $('#studentCountLabel').textContent = `${students.length.toLocaleString('vi-VN')} học sinh`;
  $('#studentsTable').innerHTML = filtered.length ? filtered.map(s => {
    const due = studentDueByCategory(s); const paid = paidByStudent.get(s.code) || {};
    const insPaid = num(paid.insurance), mandatoryPaid=num(paid.mandatory), svcPaid = num(paid.service), otherPaid=num(paid.other);
    const totalDue=studentDueItems(s).reduce((sum,x)=>sum+x.amount,0);
    const totalRemain = Math.max(0,totalDue-(insPaid+mandatoryPaid+svcPaid+otherPaid));
    return `<tr><td><strong>${escapeHTML(s.code)}</strong></td><td>${escapeHTML(s.name)}</td><td>${escapeHTML(s.className || '—')}</td><td>${money(due.insurance)}</td><td>${money(insPaid)}</td><td class="remain-cell">${money(Math.max(0,due.insurance-insPaid))}</td><td>${money(due.mandatory)}</td><td>${money(mandatoryPaid)}</td><td class="remain-cell">${money(Math.max(0,due.mandatory-mandatoryPaid))}</td><td>${money(Math.max(0,due.other-otherPaid))}</td><td><strong>${money(totalRemain)}</strong></td></tr>`;
  }).join('') : `<tr><td colspan="11" class="empty-cell">${students.length ? 'Không tìm thấy học sinh phù hợp.' : 'Chưa có học sinh. Hãy tải file danh sách ban đầu.'}</td></tr>`;
}
function renderClasses(students, transactions) {
  const groups = new Map();
  students.forEach(s => { const key=s.className||'Chưa xếp lớp'; const g=groups.get(key)||{students:[],count:0,due:0,paid:0,dueItems:0,paidItems:0};g.students.push(s);g.count++;groups.set(key,g); });
  const rows = [...groups.entries()].sort((a,b)=>a[0].localeCompare(b[0],'vi')).slice(0,30);
  $('#classTable').innerHTML = rows.length ? rows.map(([name,g])=>{
    const codes=new Set(g.students.map(s=>s.code));const classTransactions=transactions.filter(t=>codes.has(t.studentCode));
    const m=totals(g.students,classTransactions);const pct=m.due?Math.min(100,Math.round(m.paid/m.due*100)):0;
    return `<tr><td><strong>${escapeHTML(name)}</strong></td><td>${g.count}</td><td>${m.dueItems}</td><td>${m.paidItems}</td><td>${money(m.remain)}</td><td><div class="class-progress"><span>${pct}%</span><span class="tiny-track"><i style="width:${pct}%"></i></span></div></td></tr>`;
  }).join('') : '<tr><td colspan="6" class="empty-cell">Chưa có dữ liệu. Nhập danh sách học sinh để bắt đầu.</td></tr>';
}
function renderTransactions(transactions, limit) {
  const sorted=[...transactions].sort((a,b)=>(b.date||b.importedAt||'').localeCompare(a.date||a.importedAt||''));
  return (limit?sorted.slice(0,limit):sorted).map(t=>{
    const key=transactionCategory(t);const badgeClass=key==='service'?'service':key==='other'?'unknown':'';
    const statuses={valid:'Khớp món thu',unmatched:'Không tìm thấy mã khoản',missing_code:'Thiếu mã HS',amount_mismatch:'Sai số tiền món',missing_category:'Sai loại khoản',no_due:'Không có món phải thu',duplicate:'Trùng món thu',ambiguous:'Món tiền chưa phân biệt được',bank_not_successful:'Giao dịch không thành công'};
    return `<tr><td>${escapeHTML(t.date||'—')}</td><td>${escapeHTML(t.ref||t.id.slice(0,18))}</td><td><span class="category-badge ${badgeClass}">${getFeeLabel(key)}</span></td><td title="${escapeHTML(t.content)}">${escapeHTML((t.content||t.feeDetail||'—').slice(0,60))}</td><td>${escapeHTML(t.studentName||t.reportedStudentCode||'—')}</td><td><strong>${money(t.amount)}</strong></td><td><span class="status-badge ${t.paymentStatus==='valid'?'':'unmatched'}">${statuses[t.paymentStatus]||'Chưa đối soát'}</span></td></tr>`;
  }).join('');
}
function renderFeeProgress(summaries) {
  const visible=FEES.map(f=>summaries[f.key]).filter(s=>s.dueItems||s.paid);
  $('#feeProgressList').innerHTML=visible.length?visible.map((s,i)=>`<div class="fee-progress-row"><span class="fee-index">${String(i+1).padStart(2,'0')}</span><div class="fee-main"><strong>${s.label}</strong><small>${s.paidItems} / ${s.dueItems} món thu đủ</small></div><div class="fee-track"><i style="width:${s.pct}%"></i></div><div class="fee-progress-meta"><strong>${s.pct}%</strong><small>${money(s.paid)}</small></div></div>`).join(''):'<div class="empty-inline">Nhập danh sách học sinh và số phải thu theo từng khoản để xem tiến độ.</div>';
}
function renderChart(transactions) {
  const daily=new Map();
  transactions.filter(t=>t.paymentStatus==='valid').forEach(t=>{const date=t.date||String(t.importedAt||'').slice(0,10)||'Chưa rõ ngày';daily.set(date,(daily.get(date)||0)+num(t.amount));});
  const points=[...daily.entries()].sort((a,b)=>a[0].localeCompare(b[0]));
  $('#chartTotal').textContent=money(transactions.filter(t=>t.paymentStatus==='valid').reduce((sum,t)=>sum+num(t.amount),0)).replace(' ₫','');
  if(!points.length){$('#collectionChart').innerHTML='<div class="chart-empty">Chưa có giao dịch để lập biểu đồ.</div>';return;}
  let running=0;const values=points.map(([,amount])=>(running+=amount));const max=Math.max(...values,1);const width=480,height=165,pad={l:34,r:9,t:14,b:27};
  const coords=values.map((v,i)=>({x:pad.l+(points.length===1?(width-pad.l-pad.r)/2:i*(width-pad.l-pad.r)/(points.length-1)),y:pad.t+(1-v/max)*(height-pad.t-pad.b)}));
  const path=coords.map((p,i)=>`${i?'L':'M'}${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');const area=`${path} L${coords.at(-1).x},${height-pad.b} L${coords[0].x},${height-pad.b} Z`;
  const grid=[0,1,2,3].map(i=>{const y=pad.t+i*(height-pad.t-pad.b)/3;const label=money(max*(1-i/3)).replace(' ₫','');return `<line x1="${pad.l}" y1="${y}" x2="${width-pad.r}" y2="${y}" stroke="#e7eeeb" stroke-dasharray="3 4"/><text x="0" y="${y+3}" fill="#899994" font-size="9">${escapeHTML(label)}</text>`}).join('');
  const dates=points.length<5?points.map((p,i)=>i):[0,Math.floor((points.length-1)/3),Math.floor(2*(points.length-1)/3),points.length-1];
  const labels=[...new Set(dates)].map(i=>`<text x="${coords[i].x}" y="${height-5}" text-anchor="middle" fill="#899994" font-size="9">${escapeHTML(points[i][0].slice(5)||points[i][0])}</text>`).join('');
  $('#collectionChart').innerHTML=`<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Biểu đồ lũy kế số tiền đã thu"><defs><linearGradient id="chartFill" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="#159987" stop-opacity=".2"/><stop offset="1" stop-color="#159987" stop-opacity="0"/></linearGradient></defs>${grid}<path d="${area}" fill="url(#chartFill)"/><path d="${path}" fill="none" stroke="#0c8b7a" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>${coords.map(p=>`<circle cx="${p.x}" cy="${p.y}" r="2.7" fill="#fff" stroke="#0c8b7a" stroke-width="2"/>`).join('')}${labels}</svg>`;
}
function renderFeeDetails(students, transactions, summaries) {
  const visible=['insurance','mandatory','service'].map(key=>summaries[key]).filter(s=>s.due||s.paid);
  $('#feeCategoryCards').innerHTML=visible.length?visible.map(s=>{
    const feeTx=transactions.filter(t=>transactionCategory(t)===s.key);const detailNames=[...new Set(feeTx.filter(t=>s.key==='service').map(t=>t.feeDetail||serviceDetail('',t.content)))].filter(Boolean).slice(0,6);
    const cls=s.key==='service'?'service-card':'';
    const description=s.key==='insurance'?'Bảo hiểm y tế theo mã khoản YT':s.key==='mandatory'?'Bảo hiểm bắt buộc theo mã khoản TT':'Gửi xe, nước uống và các dịch vụ khác';
    return `<article class="panel fee-detail-card ${cls}"><div class="fee-detail-title"><div><h2>${s.label}</h2><p>${description}</p></div><span class="category-badge ${s.key==='service'?'service':''}">${s.dueItems} món</span></div><div class="fee-total">${money(s.due)}</div><div class="fee-card-progress"><i style="width:${s.pct}%"></i></div><div class="fee-card-foot">Đã ghi nhận ${money(s.paid)} · Còn ${money(s.remain)} · ${s.pct}% giá trị</div><div class="fee-breakdown"><div><span>Phải thu</span><strong>${s.dueItems} món</strong></div><div><span>Đã thu đủ</span><strong>${s.paidItems} món</strong></div><div><span>Giao dịch</span><strong>${feeTx.length}</strong></div></div>${detailNames.length?`<div class="service-breakdown">${detailNames.map(n=>`<span class="service-chip">${escapeHTML(n)}</span>`).join('')}</div>`:''}</article>`;
  }).join(''):'<div class="panel empty-state">Chưa có số liệu theo từng khoản thu. Nhập danh sách học sinh và báo cáo thu để bắt đầu.</div>';
  const serviceTx=transactions.filter(t=>transactionCategory(t)==='service');const breakdown=new Map();
  serviceTx.forEach(t=>{const key=t.feeDetail||serviceDetail('',t.content);const b=breakdown.get(key)||{count:0,codes:new Set(),amount:0};b.count++;if(t.studentCode||t.reportedStudentCode)b.codes.add(t.studentCode||t.reportedStudentCode);if(t.paymentStatus==='valid')b.amount+=num(t.amount);breakdown.set(key,b);});
  $('#serviceTxnCount').textContent=`${serviceTx.length} giao dịch`;
  $('#serviceBreakdownTable').innerHTML=breakdown.size?[...breakdown.entries()].sort((a,b)=>b[1].amount-a[1].amount).map(([name,b])=>`<tr><td><strong>${escapeHTML(name)}</strong></td><td>${b.count}</td><td>${b.codes.size}</td><td><strong>${money(b.amount)}</strong></td></tr>`).join(''):'<tr><td colspan="4" class="empty-cell">Chưa có giao dịch dịch vụ.</td></tr>';
}
function qrText(value,maxLength=25){return String(value||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toUpperCase().replace(/[^A-Z0-9 _.-]/g,' ').replace(/\s+/g,' ').trim().slice(0,maxLength);}
function emvTag(id,value){const text=String(value);if(text.length>99)throw new Error(`Trường QR ${id} vượt độ dài cho phép.`);return id+String(text.length).padStart(2,'0')+text;}
function crc16ccitt(value){let crc=0xFFFF;for(let i=0;i<value.length;i++){crc^=value.charCodeAt(i)<<8;for(let bit=0;bit<8;bit++)crc=crc&0x8000?(crc<<1)^0x1021:crc<<1;crc&=0xFFFF;}return crc.toString(16).toUpperCase().padStart(4,'0');}
function buildVietQrPayload(config,amount,remark){
  const accountInfo=emvTag('00','A000000727')+emvTag('01',emvTag('00',config.bin)+emvTag('01',config.accountNumber))+emvTag('02','QRIBFTTA');
  const reference=emvTag('08',qrText(remark,25));
  let payload='000201010212'+emvTag('38',accountInfo)+'52040000'+'5303704'+emvTag('54',String(num(amount)))+'5802VN'+emvTag('59',qrText(config.accountName,25))+'6007DONGHA'+emvTag('62',reference)+'6304';
  payload+=crc16ccitt(payload);return payload;
}
function qrCandidates(students,transactions){
  const paid=new Set(transactions.filter(t=>t.paymentStatus==='valid').map(t=>`${t.studentCode}|${t.matchedDueItemId||''}`));
  return students.flatMap(student=>studentDueItems(student).filter(item=>item.category!=='other'&&!item.legacy&&item.amount>0).map(item=>({student,item,paid:paid.has(`${student.code}|${item.id}`)})));
}
function renderQrPage(students,transactions,config){
  if(config){$('#qrBankBin').value=config.bin||'';$('#qrAccountNumber').value=config.accountNumber||'';$('#qrAccountName').value=config.accountName||'';$('#qrConfigStatus').textContent='Đã lưu trên thiết bị này';}
  $('#downloadQrs').hidden=true;$('#qrSelectionCount').textContent='Chưa có QR được tạo';$('#qrPreviewGrid').innerHTML='<div class="panel qr-empty-state">Chọn điều kiện rồi bấm “Tạo QR”.</div>';
  const currentClass=$('#qrClassFilter').value||'all';const classes=[...new Set(students.map(s=>s.className).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'vi'));
  $('#qrClassFilter').innerHTML='<option value="all">Tất cả lớp</option>'+classes.map(name=>`<option value="${escapeHTML(name)}">${escapeHTML(name)}</option>`).join('');
  if(classes.includes(currentClass))$('#qrClassFilter').value=currentClass;
  const candidates=qrCandidates(students,transactions);$('#qrDueCount').textContent=`${candidates.filter(x=>!x.paid).length} món còn phải thu`;
  if(students.some(s=>!s.hasFeeBreakdown))$('#qrPreviewSummary').textContent='Có học sinh chưa có dữ liệu chi tiết từng khoản. Hãy nhập lại danh sách phải thu để tạo QR chính xác.';
}
async function saveQrConfig(){
  const config={key:'qrAccount',bin:$('#qrBankBin').value.trim(),accountNumber:$('#qrAccountNumber').value.trim(),accountName:qrText($('#qrAccountName').value,25)};
  if(!/^\d{6}$/.test(config.bin))return toast('Mã BIN cần có đúng 6 chữ số.',true);
  if(!/^\d{4,30}$/.test(config.accountNumber))return toast('Hãy nhập số tài khoản nhận gồm 4–30 chữ số.',true);
  if(!config.accountName)return toast('Hãy nhập tên chủ tài khoản.',true);
  await request('meta','put',config);$('#qrAccountName').value=config.accountName;$('#qrConfigStatus').textContent='Đã lưu trên thiết bị này';toast('Đã lưu tài khoản nhận cục bộ.');
}
function qrFilename(entry){return `${slug(entry.student.className||'lop')}_${slug(entry.student.name)}_${slug(entry.student.code)}_${slug(entry.item.name)}.png`;}
function renderQrCards(entries){
  if(!entries.length){$('#qrPreviewGrid').innerHTML='<div class="panel qr-empty-state">Không có món phù hợp với điều kiện đã chọn.</div>';return;}
  $('#qrPreviewGrid').innerHTML=entries.map((entry,index)=>`<article class="panel qr-result-card"><div class="qr-result-top"><span class="category-badge ${entry.item.category==='service'?'service':''}">${escapeHTML(entry.item.name)}</span><span class="qr-class-tag">${escapeHTML(entry.student.className||'Chưa xếp lớp')}</span></div><div class="qr-person"><strong>${escapeHTML(entry.student.name)}</strong><small>${escapeHTML(entry.student.code)}</small></div><img src="${entry.png}" alt="Mã QR ${escapeHTML(entry.student.code)} ${escapeHTML(entry.item.name)}"><div class="qr-amount">${money(entry.item.amount)}</div><div class="qr-result-footer"><span>${escapeHTML(entry.remark)}</span><a class="button button-outline button-small" href="${entry.png}" download="${escapeHTML(entry.filename)}">Lưu PNG</a></div></article>`).join('');
}
async function generateQrs(){
  const config=await request('meta','get','qrAccount');if(!config?.bin||!config?.accountNumber||!config?.accountName)return toast('Hãy lưu tài khoản nhận tiền của trường trước.',true);
  if(typeof QRCode==='undefined'||typeof JSZip==='undefined')return toast('Thiếu bộ tạo ảnh cục bộ. Tải lại trang sau khi kết nối mạng.',true);
  const [students,stored]=await Promise.all([all('students'),all('transactions')]);const transactions=reconcileTransactions(students,stored);
  const className=$('#qrClassFilter').value,fee=$('#qrFeeFilter').value,status=$('#qrStatusFilter').value;
  const entries=qrCandidates(students,transactions).filter(x=>(className==='all'||x.student.className===className)&&(fee==='all'||x.item.category===fee)&&(status==='all'||!x.paid));
  if(!entries.length)return toast('Không có món phải thu phù hợp để tạo QR.',true);
  const output=[];const unique=new Set();
  for(const entry of entries){
    const key=`${entry.student.code}|${entry.item.id}`;if(unique.has(key))continue;unique.add(key);
    const code=entry.item.category==='insurance'?'BHYT':entry.item.category==='mandatory'?'BHTT':qrText(entry.item.name,12).replace(/\s+/g,'');
    const remark=qrText(entry.item.paymentCode||`${entry.student.code} ${code}`,25);const payload=buildVietQrPayload(config,entry.item.amount,remark);
    const holder=document.createElement('div');new QRCode(holder,{text:payload,width:240,height:240,correctLevel:QRCode.CorrectLevel.M});
    const canvas=holder.querySelector('canvas');if(!canvas)throw new Error('Không tạo được ảnh QR trên trình duyệt này.');
    output.push({...entry,remark,filename:qrFilename(entry),png:canvas.toDataURL('image/png')});
  }
  renderQrCards(output);$('#qrSelectionCount').textContent=`${output.length} ảnh QR đã tạo`;
  $('#qrPreviewSummary').textContent=`${output.length} mã, mỗi mã gắn với đúng một học sinh và một món thu.`;
  const zip=new JSZip();for(const item of output)zip.folder(qrText(item.student.className||'Chua_xep_lop',30)||'Chua_xep_lop').file(item.filename,item.png.split(',')[1],{base64:true});
  $('#downloadQrs').onclick=async()=>{const blob=await zip.generateAsync({type:'blob'});download(`SchoolCollect_QR_${new Date().toISOString().slice(0,10)}.zip`,blob,'application/zip');};
  $('#downloadQrs').hidden=false;
}
async function refresh() {
  const [students,storedTransactions,history,qrConfig]=await Promise.all([all('students'),all('transactions'),all('history'),request('meta','get','qrAccount')]);
  const transactions=reconcileTransactions(students,storedTransactions);
  const priorById=new Map(storedTransactions.map(t=>[t.id,t]));
  if(transactions.some(t=>{const old=priorById.get(t.id);return !old||['paymentStatus','matched','studentCode','studentName','matchedDueItem','matchedDueItemId'].some(key=>t[key]!==old[key]);})) await putMany('transactions',transactions);
  const t=totals(students,transactions);const classes=new Set(students.map(s=>s.className).filter(Boolean));
  $('#statStudents').textContent=students.length.toLocaleString('vi-VN');$('#statClasses').textContent=students.length?`${classes.size} lớp`:'Chưa có danh sách';
  $('#statFeeItems').textContent=t.dueItems.toLocaleString('vi-VN');$('#statDueAmount').textContent=`${money(t.due)} phải thu`;
  $('#statPaidItems').textContent=t.paidItems.toLocaleString('vi-VN');$('#statPaidAmount').textContent=`${money(t.paid)} đã ghi nhận`;
  $('#statUnpaidItems').textContent=t.unpaidItems.toLocaleString('vi-VN');$('#statRemainAmount').textContent=`${money(t.remain)} còn lại`;
  $('#completionPercent').textContent=`${t.pct}%`;
  $('#completionRing').style.background=`conic-gradient(var(--teal) ${t.pct*3.6}deg,#dce8e5 0deg)`;
  const legacy=students.some(s=>!s.hasFeeBreakdown);
  const notice=students.length?`${students.length} học sinh · ${transactions.filter(x=>x.paymentStatus==='valid').length} món thu khớp chính xác / ${transactions.length} giao dịch. ${legacy?'Danh sách cũ chỉ có tổng phải thu; hãy nhập lại file chi tiết từng khoản.':'Dữ liệu đang lưu riêng trên máy này.'}`:'Chọn danh sách học sinh và báo cáo thu để bắt đầu theo dõi.';
  $('#dataNoticeText').textContent=notice;
  renderFeeProgress(t.summaries);renderChart(transactions);renderClasses(students,transactions);renderStudents(students,transactions);renderFeeDetails(students,transactions,t.summaries);
  renderQrPage(students,transactions,qrConfig);
  const transactionHtml=transactions.length?renderTransactions(transactions):'<tr><td colspan="7" class="empty-cell">Chưa có báo cáo thu.</td></tr>';
  $('#transactionsTable').innerHTML=transactionHtml;$('#importsTransactionsTable').innerHTML=transactionHtml;
  ['allTxnCount','importsAllTxnCount'].forEach(id=>{const el=$(`#${id}`);if(el)el.textContent=transactions.length;});
  ['matchedTxnCount','importsMatchedTxnCount'].forEach(id=>{const el=$(`#${id}`);if(el)el.textContent=transactions.filter(x=>x.paymentStatus==='valid').length;});
  ['unmatchedTxnCount','importsUnmatchedTxnCount'].forEach(id=>{const el=$(`#${id}`);if(el)el.textContent=transactions.filter(x=>x.paymentStatus!=='valid').length;});
  $('#recentTransactions').innerHTML=transactions.length?[...transactions].sort((a,b)=>(b.importedAt||'').localeCompare(a.importedAt||'')).slice(0,4).map(x=>`<div class="recent-row"><strong><span class="category-badge ${transactionCategory(x)==='service'?'service':''}">${getFeeLabel(transactionCategory(x))}</span> ${escapeHTML(x.studentName||x.content||'Giao dịch thu')}</strong><span>${escapeHTML(x.date||'—')}</span><b>${money(x.amount)}</b></div>`).join(''):'<div class="empty-inline">Chưa có giao dịch được nhập.</div>';
  $('#historyTable').innerHTML=history.length?history.sort((a,b)=>b.at.localeCompare(a.at)).map(h=>`<tr><td>${dateTime(h.at)}</td><td>${escapeHTML(h.kind)}</td><td>${escapeHTML(h.fileName)}</td><td>${h.rows}</td><td>${escapeHTML(h.detail)}</td></tr>`).join(''):'<tr><td colspan="5" class="empty-cell">Chưa có lịch sử nhập file.</td></tr>';
  $('#storageStatus').textContent='Kho trình duyệt đã sẵn sàng';
}
function download(name, content, type='application/json') {
  const url=URL.createObjectURL(new Blob([content],{type}));const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
async function backup() {
  const password=prompt('Đặt mật khẩu cho tệp sao lưu (ít nhất 8 ký tự):');if(password===null)return;
  if(password.length<8)return toast('Mật khẩu cần có ít nhất 8 ký tự.',true);
  try {
    const data=JSON.stringify({version:1,createdAt:new Date().toISOString(),students:await all('students'),transactions:await all('transactions'),history:await all('history'),meta:await all('meta')});
    const salt=crypto.getRandomValues(new Uint8Array(16)),iv=crypto.getRandomValues(new Uint8Array(12));
    const base=await crypto.subtle.importKey('raw',new TextEncoder().encode(password),'PBKDF2',false,['deriveKey']);
    const key=await crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:250000,hash:'SHA-256'},base,{name:'AES-GCM',length:256},false,['encrypt']);
    const cipher=await crypto.subtle.encrypt({name:'AES-GCM',iv},key,new TextEncoder().encode(data));
    download(`so-thu-sao-luu-${new Date().toISOString().slice(0,10)}.sctbackup`,JSON.stringify({format:'SCTBACKUP1',salt:[...salt],iv:[...iv],cipher:[...new Uint8Array(cipher)]}));
    toast('Đã tạo bản sao lưu mã hóa. Hãy cất mật khẩu riêng.');
  } catch(e) { console.error(e);toast('Không tạo được tệp sao lưu.',true); }
}
async function restore(file) {
  const password=prompt('Nhập mật khẩu của bản sao lưu:');if(password===null)return;
  try {
    const payload=JSON.parse(await file.text());if(payload.format!=='SCTBACKUP1')throw new Error('Tệp sao lưu không đúng định dạng.');
    const salt=new Uint8Array(payload.salt),iv=new Uint8Array(payload.iv),cipher=new Uint8Array(payload.cipher);
    const base=await crypto.subtle.importKey('raw',new TextEncoder().encode(password),'PBKDF2',false,['deriveKey']);
    const key=await crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:250000,hash:'SHA-256'},base,{name:'AES-GCM',length:256},false,['decrypt']);
    const raw=await crypto.subtle.decrypt({name:'AES-GCM',iv},key,cipher);const data=JSON.parse(new TextDecoder().decode(raw));
    if(!confirm('Khôi phục sẽ thay thế toàn bộ dữ liệu hiện có trên máy này. Tiếp tục?'))return;
    await clearAll();await Promise.all([putMany('students',data.students||[]),putMany('transactions',data.transactions||[]),putMany('history',data.history||[]),putMany('meta',data.meta||[])]);
    await refresh();toast('Đã khôi phục dữ liệu từ bản sao lưu.');
  } catch(e) { console.error(e);toast('Không mở được bản sao lưu. Kiểm tra đúng tệp và mật khẩu.',true); }
}
function exportStudents() {
  all('students').then(items=>{
    const headers=['Mã học sinh','Họ và tên','Lớp','Số tiền phải thu','Mã khoản BHYT','BHYT phải thu','Mã khoản BHTT','BHTT phải thu','Dịch vụ khác phải thu','Gửi xe phải thu','Nước uống phải thu'];
    const lines=items.map(s=>{const fees=studentDueItems(s);const amt=name=>fees.filter(x=>slug(x.name)===slug(name)).reduce((sum,x)=>sum+x.amount,0);const health=fees.find(x=>x.category==='insurance');const mandatory=fees.find(x=>x.category==='mandatory');const insurance=fees.filter(x=>x.category==='insurance').reduce((sum,x)=>sum+x.amount,0);const required=fees.filter(x=>x.category==='mandatory').reduce((sum,x)=>sum+x.amount,0);const service=fees.filter(x=>x.category==='service'&&!['gui xe','nuoc uong'].includes(slug(x.name))).reduce((sum,x)=>sum+x.amount,0);return [s.code,s.name,s.className,s.due,health?.paymentCode||'',insurance,mandatory?.paymentCode||'',required,service,amt('Gửi xe'),amt('Nước uống')].map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(',');});
    download('danh-sach-hoc-sinh.csv','\uFEFF'+headers.join(',')+'\r\n'+lines.join('\r\n'),'text/csv;charset=utf-8');
  });
}
function wire() {
  $$('.nav-item[data-page]').forEach(btn=>btn.addEventListener('click',()=>setPage(btn.dataset.page)));
  $$('[data-go]').forEach(btn=>btn.addEventListener('click',()=>setPage(btn.dataset.go)));
  $('#studentImportButton').onclick=$('#studentImportButton2').onclick=()=>$('#studentFileInput').click();
  $('#bankImportButton').onclick=()=>$('#bankFileInput').click();
  $('#studentFileInput').onchange=e=>{chooseFile('students',e.target.files[0]);e.target.value='';};
  $('#bankFileInput').onchange=e=>{chooseFile('bank',e.target.files[0]);e.target.value='';};
  $('#restoreFileInput').onchange=e=>{restore(e.target.files[0]);e.target.value='';};
  $('#modalClose').onclick=$('#modalCancel').onclick=closeModal;$('#modalConfirm').onclick=confirmImport;
  $('#modalBackdrop').addEventListener('click',e=>{if(e.target.id==='modalBackdrop')closeModal();});
  $('#studentSearch').addEventListener('input',async()=>renderStudents(await all('students'),await all('transactions')));
  $('#saveQrConfig').onclick=saveQrConfig;$('#generateQrs').onclick=()=>generateQrs().catch(e=>{console.error(e);toast(e.message||'Không tạo được mã QR.',true);});
  ['qrFeeFilter','qrClassFilter','qrStatusFilter'].forEach(id=>$(`#${id}`).addEventListener('change',()=>$('#downloadQrs').hidden=true));
  $('#exportStudents').onclick=exportStudents;$('#backupButton').onclick=$('#backupButtonTop').onclick=backup;$('#restoreButton').onclick=()=>$('#restoreFileInput').click();
  $('#clearDataButton').onclick=async()=>{if(confirm('Xóa toàn bộ dữ liệu học sinh, giao dịch và lịch sử trên trình duyệt này?')){await clearAll();await refresh();toast('Đã xóa dữ liệu trên máy này.');}};
  $('#menuToggle').onclick=()=>$('#sidebar').classList.toggle('open');
  if('serviceWorker' in navigator&&location.protocol.startsWith('http'))navigator.serviceWorker.register('sw.js').catch(err=>console.warn('Offline cache:',err));
}
document.addEventListener('DOMContentLoaded',async()=>{
  try { db=await openDatabase();wire();await refresh(); }
  catch(error){console.error(error);$('#storageStatus').textContent='Không mở được kho dữ liệu';toast('Trình duyệt không cho phép lưu dữ liệu cục bộ.',true);}
});
