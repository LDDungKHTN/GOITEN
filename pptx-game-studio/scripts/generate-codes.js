#!/usr/bin/env node
/**
 * Sinh hàng loạt mã kích hoạt ngẫu nhiên (không trùng, bỏ ký tự dễ nhầm 0 O 1 I L).
 *   node scripts/generate-codes.js --count 50 --credits 30 --prefix SLIDE30 --length 6 --expires 2027-06-30
 * Xuất 3 file vào ./output-codes/: .json, .csv, .sql (dán file .sql vào Supabase > SQL Editor để nạp mã).
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function args() {
  const a = process.argv.slice(2), o = {};
  const alias = { n: 'count', c: 'credits', p: 'prefix', l: 'length', e: 'expires' };
  for (let i = 0; i < a.length; i++) {
    const m = a[i].match(/^--?([a-z]+)(?:=(.*))?$/i);
    if (!m) continue;
    const k = alias[m[1]] || m[1];
    o[k] = m[2] != null ? m[2] : (a[i + 1] && !a[i + 1].startsWith('-') ? a[++i] : true);
  }
  return o;
}
const o = args();
if (o.help) {
  console.log('Dùng: node scripts/generate-codes.js [-n 50] [-c 30] [-p SLIDE30] [-l 6] [-e 2027-12-31]');
  process.exit(0);
}
const count = Math.max(1, parseInt(o.count, 10) || 50);
const credits = Math.max(1, parseInt(o.credits, 10) || 30);
const prefix = String(o.prefix || `SLIDE${credits}`).toUpperCase().replace(/[^A-Z0-9]/g, '');
const length = Math.min(16, Math.max(4, parseInt(o.length, 10) || 6));
let expires = null;
if (o.expires && o.expires !== true) {
  const d = new Date(o.expires);
  if (isNaN(d)) { console.error('Ngày hết hạn không hợp lệ:', o.expires); process.exit(1); }
  expires = d.toISOString();
}

const ALPHA = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
function rand(n) {
  let s = '';
  while (s.length < n) {
    const b = crypto.randomBytes(1)[0];
    if (b < 256 - (256 % ALPHA.length)) s += ALPHA[b % ALPHA.length];
  }
  return s;
}
const set = new Set();
while (set.size < count) set.add(`${prefix}-${rand(length)}`);
const now = new Date().toISOString();
const codes = [...set].map((code) => ({ code, credits, status: 'unused', expires_at: expires, created_at: now, issued_to: null, issued_at: null, note: null }));

const dir = path.join(process.cwd(), 'output-codes');
fs.mkdirSync(dir, { recursive: true });
const base = path.join(dir, `${prefix}-${now.slice(0, 19).replace(/[:T]/g, '-')}-${count}ma`);
fs.writeFileSync(base + '.json', JSON.stringify(codes, null, 2));
fs.writeFileSync(base + '.csv', '﻿code,credits,status,expires_at,created_at,issued_to,issued_at,note\n' +
  codes.map((c) => [c.code, c.credits, c.status, c.expires_at || '', c.created_at, '', '', ''].join(',')).join('\n') + '\n');
const q = (v) => (v == null ? 'NULL' : `'${String(v).replace(/'/g, "''")}'`);
fs.writeFileSync(base + '.sql', `-- ${count} mã gói ${credits} lượt (${prefix}) – tạo lúc ${now}\n` +
  'INSERT INTO activation_codes (code, credits, status, expires_at) VALUES\n' +
  codes.map((c) => `  (${q(c.code)}, ${c.credits}, 'unused', ${q(c.expires_at)})`).join(',\n') + '\nON CONFLICT (code) DO NOTHING;\n');
console.log(`Đã tạo ${count} mã (${credits} lượt/mã):\n  ${base}.json\n  ${base}.csv\n  ${base}.sql`);
console.log('Ví dụ:', codes.slice(0, 3).map((c) => c.code).join(', '));
