// Kiểm tra tĩnh các bất biến nền tảng trước khi tiếp nhận nâng cấp từ Trưng Vương.
// Chạy: node tests/source-invariants.mjs
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';
const app = readFileSync(new URL('../app.js', import.meta.url), 'utf8');
const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const readme = readFileSync(new URL('../README.md', import.meta.url), 'utf8');
assert.match(app, /const DB_NAME\s*=\s*['"]so-thu-hoc-sinh-local['"]/,
  'Không được đổi tên IndexedDB của Trần Hưng Đạo.');
assert.doesNotMatch(app, /DB_NAME\s*=\s*['"]so-thu-hoc-sinh-trungvuong-local['"]/,
  'Không được dùng dữ liệu IndexedDB của Trưng Vương.');
assert.match(html, /TRƯỜNG THCS TRẦN HƯNG ĐẠO/,
  'Giao diện cần giữ nhận diện Trần Hưng Đạo.');
assert.match(readme, /\.sctbackup/,
  'Phải duy trì quy trình sao lưu dữ liệu cục bộ.');
assert.match(html, /app\.js/,
  'HTML phải tải ứng dụng chính.');
console.log('PASS: Bất biến nguồn THCS Trần Hưng Đạo.');
