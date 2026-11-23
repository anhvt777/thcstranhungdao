/* THCS Trần Hưng Đạo: báo cáo còn phải thu và bản nháp cổng phụ huynh.
   Read-only: không ghi vào IndexedDB, không tải lên máy chủ. */
(function () {
  'use strict';

  function snapshot(students, transactions) {
    const valid = transactions.filter(t => t.paymentStatus === 'valid' && t.studentCode && t.matchedDueItemId);
    const paidKeys = new Set(valid.map(t => String(t.studentCode) + '|' + String(t.matchedDueItemId)));
    const rows = [];
    for (const student of students) {
      const items = studentDueItems(student);
      for (const item of items) {
        const amount = num(item.amount);
        const paid = paidKeys.has(String(student.code) + '|' + String(item.id)) ? amount : 0;
        rows.push({
          studentCode: String(student.code || ''),
          name: String(student.name || ''),
          className: String(student.className || ''),
          birthDate: String(student.birthDate || ''),
          itemId: String(item.id || ''),
          itemName: String(item.name || ''),
          category: String(item.category || ''),
          periodId: String(item.periodId || item.period || ''),
          periodType: String(item.periodType || ''),
          paymentCode: String(item.paymentCode || ''),
          qrAccountNumber: String(item.qrAccountNumber || ''),
          due: amount, paid, remaining: Math.max(0, amount - paid),
          legacy: Boolean(item.legacy)
        });
      }
    }
    return { rows, unmatched: transactions.filter(t => t.paymentStatus !== 'valid' && t.paymentStatus !== 'voided').length };
  }

  function csvCell(value) {
    const text = String(value == null ? '' : value);
    // Ngăn Excel thực thi nội dung tên / ghi chú dưới dạng công thức.
    const safe = /^[\s]*[=+@\-\t\r]/.test(text) ? "'" + text : text;
    return '"' + safe.replace(/"/g, '""') + '"';
  }

  function exportRemaining(rows) {
    const filter = document.getElementById('parentDraftClass').value;
    const outstanding = rows.filter(r => r.remaining > 0 && (!filter || filter === r.className))
      .sort((a,b) => a.className.localeCompare(b.className,'vi') || a.name.localeCompare(b.name,'vi'));
    const table = [
      ['Lớp','Mã học sinh','Họ và tên','Khoản thu','Loại khoản','Kỳ thu','Phải thu','Đã thu khớp','Còn thiếu','Mã thanh toán','Cần xác minh dữ liệu'],
      ...outstanding.map(r=>[r.className,r.studentCode,r.name,r.itemName,r.category,r.periodId,r.due,r.paid,r.remaining,r.paymentCode,r.legacy?'Khoản cũ chưa tách mã':''])
    ];
    download('THD_CAC_KHOAN_CON_THU_' + new Date().toISOString().slice(0,10) + '.csv',
      '\uFEFF' + table.map(row=>row.map(csvCell).join(',')).join('\r\n'), 'text/csv;charset=utf-8');
    return outstanding.length;
  }

  function exportParentDraft(students, rows) {
    const byStudent = new Map();
    for (const student of students) {
      if (student.status === 'inactive') continue;
      byStudent.set(String(student.code),{
        student_id: String(student.code || ''),
        name: String(student.name || ''),
        class_name: String(student.className || ''),
        birth_date: String(student.birthDate || ''),
        charges: []
      });
    }
    for (const row of rows) {
      const target = byStudent.get(row.studentCode);
      if (!target) continue;
      target.charges.push({
        charge_id: row.itemId,
        name: row.itemName,
        category: row.category,
        period: row.periodId,
        period_id: row.periodId,
        period_type: row.periodType,
        due: row.due,
        paid: row.paid,
        remaining: row.remaining,
        payment_reference: row.paymentCode || row.qrAccountNumber,
        qr_account_number: row.qrAccountNumber,
        qr_payload: null,
        verification_required: row.legacy
      });
    }
    const draft = {
      schema: 'school-parent-sync-draft/1',
      school_id: 'THCSTRANHUNGDAO',
      school_name: 'TRƯỜNG THCS TRẦN HƯNG ĐẠO',
      generated_at: new Date().toISOString(),
      status: 'OFFLINE_DRAFT_NOT_APPROVED',
      students: [...byStudent.values()]
    };
    download('THD_PARENT_DRAFT_' + new Date().toISOString().slice(0,10) + '.json',
      JSON.stringify(draft,null,2),'application/json;charset=utf-8');
    return draft.students.length;
  }

  async function prepare() {
    const [students,raw] = await Promise.all([all('students'),all('transactions')]);
    return { students, ...snapshot(students, reconcileTransactions(students,raw)) };
  }

  async function run(type) {
    const status = document.getElementById('parentDraftStatus');
    try {
      const data = await prepare();
      if (!data.students.length) {
        status.textContent = 'Chưa có danh sách học sinh. Hãy nhập dữ liệu trước khi xuất.';
        return;
      }
      const classes = [...new Set(data.students.map(s=>s.className).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'vi'));
      const select = document.getElementById('parentDraftClass');
      const selected = select.value;
      select.replaceChildren(new Option('Tất cả lớp',''),...classes.map(c=>new Option(c,c)));
      select.value = classes.includes(selected) ? selected : '';
      if (type === 'summary') {
        const visible = data.rows.filter(r=>!select.value||r.className===select.value);
        const sum = key => visible.reduce((total,r)=>total+r[key],0);
        status.textContent = visible.length + ' món · Phải thu: ' + money(sum('due')) +
          ' · Đã thu khớp: ' + money(sum('paid')) + ' · Còn thiếu: ' + money(sum('remaining')) +
          ' · ' + data.unmatched + ' giao dịch chưa khớp/chưa xác minh (không tính là đã thu).';
        return;
      }
      if (type === 'csv') {
        const count = exportRemaining(data.rows);
        status.textContent = 'Đã xuất ' + count + ' khoản còn phải thu. Có ' +
          data.unmatched + ' giao dịch chưa khớp/chưa xác minh; cần xử lý trước khi chốt số liệu.';
        return;
      }
      if (type === 'parent') {
        if (!confirm('Tệp chứa tên, lớp, ngày sinh và tình trạng thanh toán học sinh. Chỉ lưu ở nơi được phép và KHÔNG tải lên GitHub/công khai. Xuất bản nháp cục bộ?')) return;
        const count = exportParentDraft(data.students,data.rows);
        status.textContent = 'Đã xuất bản nháp ' + count +
          ' học sinh. Chưa đồng bộ hoặc công bố lên cổng phụ huynh; ' +
          data.unmatched + ' giao dịch còn cần xác minh.';
      }
    } catch (error) {
      console.error('THD draft export:',error);
      status.textContent = 'Không xuất được dữ liệu: ' + (error.message || 'Lỗi không xác định');
    }
  }

  document.addEventListener('DOMContentLoaded',function () {
    const status = document.getElementById('parentDraftStatus');
    if (!status) return;
    document.getElementById('parentDraftRefresh').addEventListener('click',()=>run('summary'));
    document.getElementById('parentDraftClass').addEventListener('change',()=>run('summary'));
    document.getElementById('parentDraftCsv').addEventListener('click',()=>run('csv'));
    document.getElementById('parentDraftExport').addEventListener('click',()=>run('parent'));
    // Tránh tranh chấp với quá trình mở IndexedDB ban đầu; chỉ khởi chạy khi có yêu cầu.
    status.textContent = 'Bấm "Tính số còn thu" để đọc dữ liệu đang lưu trên máy.';
  });
})();
