/**
 * PPTX Game Studio – API server (tùy chọn, dùng khi KHÔNG dùng Supabase trực tiếp).
 * Gọi lại đúng các stored function trong server/schema.sql nên logic giống hệt bản Supabase.
 *
 *   DATABASE_URL=postgres://user:pass@host:5432/db  PORT=8787  ALLOW_ORIGIN=https://lddungkhtn.github.io  npm start
 *
 * Sau đó trong config.js đặt: BACKEND: 'api', API_BASE: 'https://dia-chi-server'
 */
const express = require('express');
const cors = require('cors');
const { Pool } = require('pg');

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.PGSSL === 'false' ? false : { rejectUnauthorized: false }
});
const app = express();
app.use(express.json({ limit: '10kb' }));
app.use(cors({ origin: (process.env.ALLOW_ORIGIN || '*').split(',').map((s) => s.trim()) }));

const validUser = (id) => typeof id === 'string' && /^dev_[a-z0-9]{12,40}$/.test(id);

app.get('/api/health', (req, res) => res.json({ ok: true }));

app.get('/api/balance', async (req, res) => {
  const { userId } = req.query;
  if (!validUser(userId)) return res.status(400).json({ message: 'Mã thiết bị không hợp lệ.' });
  try {
    const { rows } = await pool.query('SELECT * FROM get_export_balance($1)', [userId]);
    res.json({ remainingExports: rows[0].remaining_exports });
  } catch (e) { console.error(e); res.status(500).json({ message: 'Lỗi máy chủ.' }); }
});

app.post('/api/redeem-code', async (req, res) => {
  const { code, userId } = req.body || {};
  if (!validUser(userId) || typeof code !== 'string' || code.length > 60) return res.status(400).json({ success: false, message: 'Dữ liệu không hợp lệ.' });
  try {
    const { rows } = await pool.query('SELECT * FROM redeem_activation_code($1, $2)', [code, userId]);
    const r = rows[0];
    res.status(r.ok ? 200 : 400).json({ success: r.ok, message: r.message, addedCredits: r.added_credits, remainingExports: r.remaining_exports });
  } catch (e) { console.error(e); res.status(500).json({ success: false, message: 'Lỗi máy chủ.' }); }
});

app.post('/api/consume-export', async (req, res) => {
  const { userId, selectedCount, estimatedSlides } = req.body || {};
  if (!validUser(userId)) return res.status(400).json({ allowed: false, message: 'Mã thiết bị không hợp lệ.' });
  try {
    const { rows } = await pool.query('SELECT * FROM consume_export_credit($1, $2, $3)', [userId, parseInt(selectedCount, 10) || 0, parseInt(estimatedSlides, 10) || 0]);
    const r = rows[0];
    res.status(r.allowed ? 200 : 402).json({ allowed: r.allowed, message: r.message, remainingExports: r.remaining_exports });
  } catch (e) { console.error(e); res.status(500).json({ allowed: false, message: 'Lỗi máy chủ.' }); }
});

const port = process.env.PORT || 8787;
app.listen(port, () => console.log(`PPTX Game Studio API đang chạy tại cổng ${port}`));
