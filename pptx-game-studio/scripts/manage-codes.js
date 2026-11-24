#!/usr/bin/env node
/**
 * Cấp mã cho khách: lấy N mã chưa phát (status 'unused' & chưa có issued_to), đánh dấu 'issued' trong
 * file nguồn và xuất 1 file riêng chứa mã giao cho khách.
 *   node scripts/manage-codes.js -i output-codes/SLIDE30-....json -n 2 -c "Cô Lan - THCS A" --note "CK 20k Momo"
 *   node scripts/manage-codes.js -i output-codes/SLIDE30-....json --stats
 * Lưu ý: trạng thái 'issued' chỉ để bạn quản lý kho mã; CSDL vẫn coi mã là 'unused' cho tới khi khách nạp.
 */
const fs = require('fs');
const path = require('path');

function args() {
  const a = process.argv.slice(2), o = {};
  const alias = { i: 'input', n: 'count', c: 'customer', o: 'out' };
  for (let i = 0; i < a.length; i++) {
    const m = a[i].match(/^--?([a-z]+)(?:=(.*))?$/i);
    if (!m) continue;
    const k = alias[m[1]] || m[1];
    o[k] = m[2] != null ? m[2] : (a[i + 1] && !a[i + 1].startsWith('-') ? a[++i] : true);
  }
  return o;
}
const o = args();
if (!o.input || o.input === true) { console.error('Thiếu --input (-i) đường dẫn file .json hoặc .csv'); process.exit(1); }
// File .json và .csv cùng tên là một kho mã: luôn đọc .json (nếu có) và ghi lại cả hai để không cấp trùng.
const input = path.resolve(o.input);
const base = input.replace(/\.(json|csv)$/i, '');
const jsonFile = base + '.json', csvFile = base + '.csv';
const file = fs.existsSync(jsonFile) ? jsonFile : input;
const isCsv = /\.csv$/i.test(file);
const COLS = ['code', 'credits', 'status', 'expires_at', 'created_at', 'issued_to', 'issued_at', 'note'];

function parseCsv(txt) {
  const lines = txt.replace(/^﻿/, '').split(/\r?\n/).filter(Boolean);
  const head = lines.shift().split(',');
  return lines.map((ln) => {
    const cells = []; let cur = '', q = false;
    for (let i = 0; i < ln.length; i++) {
      const ch = ln[i];
      if (q) { if (ch === '"' && ln[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; }
      else if (ch === '"') q = true; else if (ch === ',') { cells.push(cur); cur = ''; } else cur += ch;
    }
    cells.push(cur);
    const r = {}; head.forEach((h, i) => { r[h] = cells[i] === '' || cells[i] == null ? null : cells[i]; });
    r.credits = parseInt(r.credits, 10);
    return r;
  });
}
const csvCell = (v) => (v == null ? '' : /[",\n]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : String(v));
const toCsv = (rows) => '﻿' + COLS.join(',') + '\n' + rows.map((r) => COLS.map((c) => csvCell(r[c])).join(',')).join('\n') + '\n';

const raw = fs.readFileSync(file, 'utf8');
const rows = isCsv ? parseCsv(raw) : JSON.parse(raw);
const free = rows.filter((r) => r.status === 'unused' && !r.issued_to);

if (o.stats) {
  const by = rows.reduce((m, r) => { const k = r.issued_to ? 'issued' : r.status; m[k] = (m[k] || 0) + 1; return m; }, {});
  console.log(`Tổng ${rows.length} mã:`, by);
  process.exit(0);
}
const n = Math.max(1, parseInt(o.count, 10) || 1);
const customer = o.customer && o.customer !== true ? String(o.customer) : '';
if (!customer) { console.error('Thiếu --customer (-c) tên khách hàng'); process.exit(1); }
if (free.length < n) { console.error(`Chỉ còn ${free.length} mã trống, không đủ ${n} mã.`); process.exit(1); }

const now = new Date().toISOString();
const picked = free.slice(0, n);
picked.forEach((r) => { r.status = 'issued'; r.issued_to = customer; r.issued_at = now; r.note = o.note && o.note !== true ? String(o.note) : r.note; });
if (fs.existsSync(jsonFile) || !isCsv) fs.writeFileSync(jsonFile, JSON.stringify(rows, null, 2));
if (fs.existsSync(csvFile) || isCsv) fs.writeFileSync(csvFile, toCsv(rows));

const safe = customer.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-zA-Z0-9]+/g, '-').replace(/^-|-$/g, '') || 'khach';
const out = o.out && o.out !== true ? path.resolve(o.out) : path.join(path.dirname(file), `cap-cho-${safe}-${now.slice(0, 10)}.txt`);
const txt = [`MÃ KÍCH HOẠT PPTX GAME STUDIO`, `Khách hàng: ${customer}`, `Ngày cấp: ${now.slice(0, 10)}`, o.note && o.note !== true ? `Ghi chú: ${o.note}` : '', '',
  ...picked.map((r, i) => `${i + 1}. ${r.code}  (+${r.credits} lượt xuất${r.expires_at ? `, hạn ${String(r.expires_at).slice(0, 10)}` : ''})`), '',
  'Cách dùng: mở trang PPTX Game Studio > "Nhập mã kích hoạt" > dán mã > "Nạp lượt".'].filter((l) => l !== null).join('\n');
fs.writeFileSync(out, txt + '\n');
console.log(`Đã cấp ${n} mã cho "${customer}". Còn ${free.length - n} mã trống.\nFile gửi khách: ${out}`);
picked.forEach((r) => console.log('  ' + r.code));
