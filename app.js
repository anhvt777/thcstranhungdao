/* Sổ thu học sinh · Local-first. This file never sends school data over the network. */
const DB_NAME = 'so-thu-hoc-sinh-local';
const DB_VERSION = 1;
let db;
let activeImport = null;
let toastTimer;

const $ = (s, root = document) => root.querySelector(s);
const $$ = (s, root = document) => [...root.querySelectorAll(s)];
const money = n => new Intl.NumberFormat('vi-VN').format(Math.round(Number(n) || 0)) + ' ₫';
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
  if (entry.method !== 8 || !('DecompressionStream' in window)) throw new Error('Trình duyệt chưa hỗ trợ đọc XLSX ngoại tuyến. Hãy dùng bản Excel .xlsx mới hoặc xuất CSV.');
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
  const workbookXml = await getText('xl/workbook.xml');
  const relationXml = await getText('xl/_rels/workbook.xml.rels');
  const workbook = new DOMParser().parseFromString(workbookXml, 'application/xml');
  const sheetNode = workbook.querySelector('sheet');
  if (!sheetNode) throw new Error('Không tìm thấy trang tính trong file Excel.');
  const relId = sheetNode.getAttribute('r:id');
  const relations = new DOMParser().parseFromString(relationXml, 'application/xml');
  const rel = [...relations.querySelectorAll('Relationship')].find(node => node.getAttribute('Id') === relId);
  const target = rel?.getAttribute('Target') || 'worksheets/sheet1.xml';
  const sheetPath = target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`;
  const sheetXml = await getText(sheetPath);
  if (!sheetXml) throw new Error('Không đọc được trang tính đầu tiên trong file Excel.');
  const sharedXml = await getText('xl/sharedStrings.xml');
  const sharedDoc = sharedXml ? new DOMParser().parseFromString(sharedXml, 'application/xml') : null;
  const shared = sharedDoc ? [...sharedDoc.querySelectorAll('si')].map(si => [...si.querySelectorAll('t')].map(t => t.textContent).join('')) : [];
  const doc = new DOMParser().parseFromString(sheetXml, 'application/xml');
  const rows = [...doc.querySelectorAll('sheetData row')].map(rowNode => {
    const cells = [];
    for (const cell of rowNode.querySelectorAll(':scope > c')) {
      const ref = cell.getAttribute('r') || '';
      const col = [...ref.matchAll(/[A-Z]+/g)][0]?.[0] || 'A';
      let index = 0; for (const char of col) index = index * 26 + char.charCodeAt(0) - 64; index--;
      const type = cell.getAttribute('t');
      let val = cell.querySelector('v')?.textContent ?? '';
      if (type === 's') val = shared[Number(val)] ?? '';
      else if (type === 'inlineStr') val = [...cell.querySelectorAll('is t')].map(t => t.textContent).join('');
      cells[index] = val;
    }
    return cells.map(value => value ?? '');
  });
  return rows;
}
async function readRows(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'xlsx') return parseXlsx(file);
  const buffer = await file.arrayBuffer();
  let text;
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
    due: ['phai thu', 'so tien phai thu', 'hoc phi', 'so tien', 'muc thu', 'tien thu'],
    amount: ['so tien', 'amount', 'so tien giao dich', 'credit', 'ghi co'],
    txnId: ['ma giao dich', 'transaction id', 'ma tham chieu', 'reference', 'trace', 'so but toan'],
    date: ['ngay giao dich', 'thoi gian', 'transaction date', 'ngay hach toan', 'ngay'],
    content: ['noi dung', 'dien giai', 'transaction content', 'description', 'chi tiet'],
    studentCode: ['ma hoc sinh', 'ma hs', 'student code', 'student id', 'ma dinh danh']
  }[field] || [];
  const exact = normal.findIndex(h => patterns.includes(h));
  if (exact >= 0) return String(exact);
  return String(normal.findIndex(h => patterns.some(p => h.includes(p) || p.includes(h))));
}
function fieldSelect(id, label, headers, field, required = false) {
  const guessed = guessColumn(headers, field);
  return `<div class="mapping-field"><label for="${id}">${label}${required ? ' <em>*</em>' : ''}</label><select id="${id}" data-field="${field}"><option value="-1">— Không chọn —</option>${headers.map((h, i) => `<option value="${i}" ${String(i) === guessed ? 'selected' : ''}>${escapeHTML(h || `Cột ${i + 1}`)}</option>`).join('')}</select></div>`;
}
function previewHtml(headers, rows) {
  return `<div class="preview-box"><strong>Xem trước 3 dòng đầu</strong><div class="table-wrap"><table><thead><tr>${headers.slice(0, 5).map(h => `<th>${escapeHTML(h)}</th>`).join('')}</tr></thead><tbody>${rows.slice(1, 4).map(row => `<tr>${headers.slice(0, 5).map((_, i) => `<td>${escapeHTML(row[i] || '')}</td>`).join('')}</tr>`).join('')}</tbody></table></div></div>`;
}
function openImportModal(kind, file, rows) {
  if (!rows.length || rows.length < 2) return toast('File không có dòng tiêu đề hoặc dữ liệu.', true);
  const headers = rows[0].map((h, i) => String(h || `Cột ${i + 1}`).trim());
  activeImport = { kind, file, rows, headers };
  $('#modalEyebrow').textContent = kind === 'students' ? 'DANH SÁCH HỌC SINH' : 'BÁO CÁO NGÂN HÀNG';
  $('#modalTitle').textContent = kind === 'students' ? 'Ghép cột danh sách học sinh' : 'Ghép cột báo cáo thu';
  const summary = `<div class="file-summary"><span class="file-badge">${file.name.toLowerCase().endsWith('.xlsx') ? 'XLSX' : 'CSV'}</span><div><strong>${escapeHTML(file.name)}</strong><small>${rows.length - 1} dòng dữ liệu · ${headers.length} cột</small></div></div>`;
  const fields = kind === 'students'
    ? `<div class="mapping-grid">${fieldSelect('map-code','Mã học sinh',headers,'code',true)}${fieldSelect('map-name','Họ và tên',headers,'name',true)}${fieldSelect('map-class','Lớp',headers,'className')}${fieldSelect('map-due','Số tiền phải thu',headers,'due')}</div>`
    : `<div class="mapping-grid">${fieldSelect('map-amount','Số tiền giao dịch',headers,'amount',true)}${fieldSelect('map-student','Mã học sinh trong giao dịch',headers,'studentCode')}${fieldSelect('map-txn','Mã giao dịch/tham chiếu',headers,'txnId')}${fieldSelect('map-date','Ngày giao dịch',headers,'date')}${fieldSelect('map-content','Nội dung chuyển khoản',headers,'content')}</div><p class="mapping-intro">Nếu báo cáo không có cột mã học sinh, giao dịch sẽ được ghi nhận là chưa khớp để rà soát thủ công.</p>`;
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
async function confirmImport() {
  if (!activeImport) return;
  const { kind, file, rows, headers } = activeImport; const map = getMap();
  if (kind === 'students' && (map.code < 0 || map.name < 0)) return toast('Hãy chọn cột mã học sinh và họ tên.', true);
  if (kind === 'bank' && map.amount < 0) return toast('Hãy chọn cột số tiền giao dịch.', true);
  const dataRows = rows.slice(1).filter(row => row.some(v => String(v ?? '').trim()));
  const now = new Date().toISOString(); let summary;
  if (kind === 'students') {
    const items = dataRows.map(row => ({ code: cell(row, map, 'code'), name: cell(row, map, 'name'), className: cell(row, map, 'className'), due: parseAmount(cell(row, map, 'due')), updatedAt: now }))
      .filter(s => s.code && s.name);
    await putMany('students', items);
    const codes = new Map((await all('students')).map(s => [slug(s.code), s]));
    const knownTransactions = await all('transactions');
    const rematched = knownTransactions.map(t => {
      const student = t.reportedStudentCode ? codes.get(slug(t.reportedStudentCode)) : null;
      return student ? { ...t, studentCode: student.code, studentName: student.name, matched: true } : t;
    });
    await putMany('transactions', rematched);
    summary = { rows: dataRows.length, imported: items.length, detail: `${items.length} học sinh được thêm/cập nhật` };
    $('#studentLastImport').textContent = `Gần nhất: ${file.name} · ${items.length} học sinh`;
  } else {
    const students = await all('students'); const byCode = new Map(students.map(s => [slug(s.code), s]));
    const items = dataRows.map((row, i) => {
      const amount = parseAmount(cell(row, map, 'amount')); if (!amount) return null;
      const code = cell(row, map, 'studentCode'); const student = code ? byCode.get(slug(code)) : null;
      const date = normalizeDate(cell(row, map, 'date'));
      const content = cell(row, map, 'content'); const ref = cell(row, map, 'txnId');
      const normalizedRef = slug(ref);
      const id = normalizedRef ? `ref:${normalizedRef}` : `row:${slug(date)}:${amount}:${slug(code || content)}:${i}`;
      return { id, ref, date, content, amount, reportedStudentCode: code, studentCode: student?.code || '', studentName: student?.name || '', sourceFile: file.name, importedAt: now, matched: !!student };
    }).filter(Boolean);
    const oldIds = new Set((await all('transactions')).map(t => t.id));
    const added = items.filter(t => { if (oldIds.has(t.id)) return false; oldIds.add(t.id); return true; });
    await putMany('transactions', added);
    summary = { rows: dataRows.length, imported: items.length, detail: `${added.length} giao dịch mới · ${items.length - added.length} giao dịch đã có được bỏ qua` };
    $('#bankLastImport').textContent = `Gần nhất: ${file.name} · ${added.length} giao dịch mới`;
  }
  await request('history', 'put', { id: crypto.randomUUID(), kind: kind === 'students' ? 'Danh sách học sinh' : 'Báo cáo ngân hàng', fileName: file.name, rows: summary.rows, imported: summary.imported, detail: summary.detail, at: now });
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
  const labels = { dashboard:'Tổng quan', students:'Danh sách học sinh', imports:'Nhập báo cáo thu', history:'Lịch sử nhập file', settings:'Sao lưu & cài đặt' };
  $('#topTitle').textContent = labels[page] || 'Tổng quan'; $('#sidebar').classList.remove('open'); window.scrollTo({ top:0, behavior:'smooth' });
}
function totals(students, transactions) {
  const paidByCode = new Map();
  transactions.filter(t => t.matched && t.studentCode).forEach(t => paidByCode.set(t.studentCode, (paidByCode.get(t.studentCode) || 0) + t.amount));
  const due = students.reduce((sum, s) => sum + (s.due || 0), 0);
  const paid = transactions.filter(t => t.matched).reduce((sum, t) => sum + t.amount, 0);
  return { due, paid, remain:Math.max(0,due-paid), paidByCode, pct:due ? Math.min(100,Math.round(paid/due*100)) : 0 };
}
function renderStudents(students, txn) {
  const { paidByCode } = totals(students, txn); const query = slug($('#studentSearch')?.value || '');
  const filtered = students.filter(s => !query || slug(`${s.code} ${s.name} ${s.className}`).includes(query));
  $('#studentCountLabel').textContent = `${students.length.toLocaleString('vi-VN')} học sinh`;
  $('#studentsTable').innerHTML = filtered.length ? filtered.slice(0, 500).map(s => {
    const paid = paidByCode.get(s.code) || 0; return `<tr><td><strong>${escapeHTML(s.code)}</strong></td><td>${escapeHTML(s.name)}</td><td>${escapeHTML(s.className || '—')}</td><td>${money(s.due)}</td><td>${money(paid)}</td><td>${money(Math.max(0,s.due-paid))}</td></tr>`;
  }).join('') : `<tr><td colspan="6" class="empty-cell">${students.length ? 'Không tìm thấy học sinh phù hợp.' : 'Chưa có học sinh. Hãy tải file danh sách ban đầu.'}</td></tr>`;
}
function renderClasses(students, txn) {
  const { paidByCode } = totals(students, txn); const groups = new Map();
  students.forEach(s => { const key = s.className || 'Chưa xếp lớp'; const g = groups.get(key) || { count:0,due:0,paid:0 }; g.count++; g.due += s.due || 0; g.paid += paidByCode.get(s.code) || 0; groups.set(key,g); });
  const rows = [...groups.entries()].sort((a,b)=>a[0].localeCompare(b[0],'vi')).slice(0,30);
  $('#classTable').innerHTML = rows.length ? rows.map(([name,g])=>{const pct=g.due?Math.min(100,Math.round(g.paid/g.due*100)):0;return `<tr><td><strong>${escapeHTML(name)}</strong></td><td>${g.count}</td><td>${money(g.due)}</td><td>${money(g.paid)}</td><td>${money(Math.max(0,g.due-g.paid))}</td><td><div class="class-progress"><span>${pct}%</span><span class="tiny-track"><i style="width:${pct}%"></i></span></div></td></tr>`}).join(''):`<tr><td colspan="6" class="empty-cell">Chưa có dữ liệu. Nhập danh sách học sinh để bắt đầu.</td></tr>`;
}
function renderTransactions(transactions, limit) {
  const sorted = [...transactions].sort((a,b)=>(b.date||b.importedAt).localeCompare(a.date||a.importedAt));
  const html = (limit ? sorted.slice(0,limit) : sorted.slice(0,500)).map(t=>`<tr><td>${escapeHTML(t.date||'—')}</td><td>${escapeHTML(t.ref||t.id.slice(0,18))}</td><td title="${escapeHTML(t.content)}">${escapeHTML((t.content||'—').slice(0,50))}</td><td>${escapeHTML(t.studentName||'—')}</td><td><strong>${money(t.amount)}</strong></td><td><span class="status-badge ${t.matched?'':'unmatched'}">${t.matched?'Đã khớp':'Chưa khớp'}</span></td></tr>`).join('');
  return html;
}
async function refresh() {
  const [students, transactions, history] = await Promise.all([all('students'),all('transactions'),all('history')]);
  const t = totals(students, transactions); const classes = new Set(students.map(s=>s.className).filter(Boolean));
  $('#statStudents').textContent = students.length.toLocaleString('vi-VN'); $('#statClasses').textContent = students.length ? `${classes.size} lớp · dữ liệu trên máy này` : 'Chưa có danh sách học sinh';
  $('#statDue').innerHTML = `${money(t.due).replace(' ₫','')} <small>₫</small>`; $('#statPaid').innerHTML = `${money(t.paid).replace(' ₫','')} <small>₫</small>`; $('#statRemain').innerHTML = `${money(t.remain).replace(' ₫','')} <small>₫</small>`;
  $('#statMatch').textContent = `${transactions.filter(x=>x.matched).length} giao dịch khớp`;
  $('#progressPercent').textContent = `${t.pct}%`; $('#progressFill').style.width = `${t.pct}%`;
  $('#legendPaid').textContent = money(t.paid); $('#legendDue').textContent = money(t.remain);
  $('#chartEmpty').style.display = students.length && transactions.length ? 'none' : 'flex';
  $('#step1').classList.toggle('muted', !students.length); $('#step2').classList.toggle('muted', !transactions.length);
  renderClasses(students,transactions); renderStudents(students,transactions);
  $('#transactionsTable').innerHTML = transactions.length ? renderTransactions(transactions) : '<tr><td colspan="6" class="empty-cell">Chưa có báo cáo ngân hàng.</td></tr>';
  $('#allTxnCount').textContent = transactions.length; $('#matchedTxnCount').textContent = transactions.filter(x=>x.matched).length; $('#unmatchedTxnCount').textContent = transactions.filter(x=>!x.matched).length;
  $('#recentTransactions').innerHTML = transactions.length ? [...transactions].sort((a,b)=>b.importedAt.localeCompare(a.importedAt)).slice(0,4).map(x=>`<div class="recent-row"><strong>${escapeHTML(x.studentName||x.content||'Giao dịch thu')}</strong><span>${escapeHTML(x.date||dateTime(x.importedAt))}</span><b>${money(x.amount)}</b></div>`).join('') : 'Chưa có giao dịch được nhập.';
  $('#historyTable').innerHTML = history.length ? history.sort((a,b)=>b.at.localeCompare(a.at)).map(h=>`<tr><td>${dateTime(h.at)}</td><td>${escapeHTML(h.kind)}</td><td>${escapeHTML(h.fileName)}</td><td>${h.rows}</td><td>${escapeHTML(h.detail)}</td></tr>`).join('') : '<tr><td colspan="5" class="empty-cell">Chưa có lịch sử nhập file.</td></tr>';
  $('#storageStatus').textContent = 'Kho trình duyệt đã sẵn sàng';
}
function download(name, content, type='application/json') {
  const url = URL.createObjectURL(new Blob([content],{type})); const a=document.createElement('a'); a.href=url; a.download=name; a.click(); setTimeout(()=>URL.revokeObjectURL(url),1000);
}
async function backup() {
  const password = prompt('Đặt mật khẩu cho tệp sao lưu (ít nhất 8 ký tự):');
  if (password === null) return;
  if (password.length < 8) return toast('Mật khẩu cần có ít nhất 8 ký tự.', true);
  try {
    const data = JSON.stringify({ version:1, createdAt:new Date().toISOString(), students:await all('students'), transactions:await all('transactions'), history:await all('history') });
    const salt=crypto.getRandomValues(new Uint8Array(16)), iv=crypto.getRandomValues(new Uint8Array(12));
    const base=await crypto.subtle.importKey('raw',new TextEncoder().encode(password),'PBKDF2',false,['deriveKey']);
    const key=await crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:250000,hash:'SHA-256'},base,{name:'AES-GCM',length:256},false,['encrypt']);
    const cipher=await crypto.subtle.encrypt({name:'AES-GCM',iv},key,new TextEncoder().encode(data));
    const payload={format:'SCTBACKUP1',salt:[...salt],iv:[...iv],cipher:[...new Uint8Array(cipher)]};
    download(`so-thu-sao-luu-${new Date().toISOString().slice(0,10)}.sctbackup`,JSON.stringify(payload));
    toast('Đã tạo bản sao lưu mã hóa. Hãy cất mật khẩu riêng.');
  } catch(e) { console.error(e); toast('Không tạo được tệp sao lưu.',true); }
}
async function restore(file) {
  const password=prompt('Nhập mật khẩu của bản sao lưu:'); if(password===null)return;
  try {
    const payload=JSON.parse(await file.text()); if(payload.format!=='SCTBACKUP1') throw new Error('Tệp sao lưu không đúng định dạng.');
    const salt=new Uint8Array(payload.salt),iv=new Uint8Array(payload.iv),cipher=new Uint8Array(payload.cipher);
    const base=await crypto.subtle.importKey('raw',new TextEncoder().encode(password),'PBKDF2',false,['deriveKey']);
    const key=await crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:250000,hash:'SHA-256'},base,{name:'AES-GCM',length:256},false,['decrypt']);
    const raw=await crypto.subtle.decrypt({name:'AES-GCM',iv},key,cipher); const data=JSON.parse(new TextDecoder().decode(raw));
    if(!confirm('Khôi phục sẽ thay thế toàn bộ dữ liệu hiện có trên máy này. Tiếp tục?'))return;
    await clearAll(); await Promise.all([putMany('students',data.students||[]),putMany('transactions',data.transactions||[]),putMany('history',data.history||[])]);
    await refresh(); toast('Đã khôi phục dữ liệu từ bản sao lưu.');
  } catch(e) { console.error(e); toast('Không mở được bản sao lưu. Kiểm tra đúng tệp và mật khẩu.',true); }
}
function exportStudents() { all('students').then(items=>download('danh-sach-hoc-sinh.csv','\uFEFFMã học sinh,Họ và tên,Lớp,Số tiền phải thu\r\n'+items.map(s=>[s.code,s.name,s.className,s.due].map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(',')).join('\r\n'),'text/csv;charset=utf-8')); }

function wire() {
  $('#todayLabel').textContent = new Intl.DateTimeFormat('vi-VN',{weekday:'short',day:'2-digit',month:'2-digit',year:'numeric'}).format(new Date());
  $$('.nav-item[data-page]').forEach(btn=>btn.addEventListener('click',()=>setPage(btn.dataset.page)));
  $$('[data-go]').forEach(btn=>btn.addEventListener('click',()=>setPage(btn.dataset.go)));
  $('#studentImportButton').onclick=$('#studentImportButton2').onclick=()=>$('#studentFileInput').click();
  $('#bankImportButton').onclick=()=>$('#bankFileInput').click();
  $('#studentFileInput').onchange=e=>{chooseFile('students',e.target.files[0]);e.target.value='';};
  $('#bankFileInput').onchange=e=>{chooseFile('bank',e.target.files[0]);e.target.value='';};
  $('#restoreFileInput').onchange=e=>{restore(e.target.files[0]);e.target.value='';};
  $('#modalClose').onclick=$('#modalCancel').onclick=closeModal; $('#modalConfirm').onclick=confirmImport;
  $('#modalBackdrop').addEventListener('click',e=>{if(e.target.id==='modalBackdrop')closeModal();});
  $('#studentSearch').addEventListener('input',async()=>renderStudents(await all('students'),await all('transactions')));
  $('#exportStudents').onclick=exportStudents; $('#backupButton').onclick=backup; $('#restoreButton').onclick=()=>$('#restoreFileInput').click();
  $('#clearDataButton').onclick=async()=>{if(confirm('Xóa toàn bộ dữ liệu học sinh, giao dịch và lịch sử trên trình duyệt này?')){await clearAll();await refresh();toast('Đã xóa dữ liệu trên máy này.');}};
  $('#menuToggle').onclick=()=>$('#sidebar').classList.toggle('open');
  $('.notice-close').onclick=()=>$('.notice-close').parentElement.remove();
  if('serviceWorker' in navigator && location.protocol.startsWith('http')) navigator.serviceWorker.register('sw.js').catch(err=>console.warn('Offline cache:',err));
}

document.addEventListener('DOMContentLoaded', async () => {
  try { db=await openDatabase(); wire(); await refresh(); }
  catch(error) { console.error(error); $('#storageStatus').textContent='Không mở được kho dữ liệu'; toast('Trình duyệt không cho phép lưu dữ liệu cục bộ.',true); }
});
