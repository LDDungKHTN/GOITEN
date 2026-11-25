/*!
 * PPTX Game Studio – Lớp kết nối quản lý lượt xuất & mã kích hoạt.
 * Mọi phép kiểm tra mã / trừ lượt ở chế độ thật đều do CSDL xử lý nguyên tử (stored function).
 */
(function (root) {
  'use strict';
  const C = root.GS_CONFIG || {};

  function deviceId() {
    const K = 'gs_device_id';
    let id = null;
    try { id = localStorage.getItem(K); } catch (e) { /* trình duyệt chặn storage */ }
    if (!id || !/^dev_[a-z0-9]{12,40}$/.test(id)) {
      const a = new Uint8Array(12);
      (root.crypto || root.msCrypto).getRandomValues(a);
      id = 'dev_' + Array.from(a, (b) => (b % 36).toString(36)).join('');
      try { localStorage.setItem(K, id); } catch (e) { /* bỏ qua */ }
    }
    return id;
  }

  function mode() {
    const b = (C.BACKEND || 'auto').toLowerCase();
    if (b === 'auto') return C.SUPABASE_URL && C.SUPABASE_ANON_KEY ? 'supabase' : C.API_BASE ? 'api' : 'demo';
    return b;
  }

  // Chuẩn hóa mã: chấp nhận chữ thường, dấu cách, thiếu gạch nối (gs abcd efgh -> GS-ABCD-EFGH)
  const normCode = (c) => {
    const s = String(c || '').trim().toUpperCase().replace(/\s+/g, '');
    const raw = s.replace(/[^A-Z0-9]/g, '');
    return /^GS[A-Z0-9]{8}$/.test(raw) ? `GS-${raw.slice(2, 6)}-${raw.slice(6)}` : s;
  };

  /* ---------- Chế độ trên trình duyệt (không cần máy chủ) ----------
   * Mã kích hoạt được kiểm tra bằng "dấu vân tay" PBKDF2 trong js/code-hashes.js (không chứa mã thật).
   * Lượt xuất lưu trên trình duyệt; mỗi mã chỉ nạp được 1 lần trên mỗi trình duyệt. */
  const enc = (s) => new TextEncoder().encode(s);
  async function codeHash(code, cfg) {
    const key = await crypto.subtle.importKey('raw', enc(code), 'PBKDF2', false, ['deriveBits']);
    const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', salt: enc(cfg.salt), iterations: cfg.iter, hash: 'SHA-256' }, key, 256);
    return Array.from(new Uint8Array(bits)).slice(0, 16).map((b) => b.toString(16).padStart(2, '0')).join('');
  }
  const demo = {
    KEY: 'gs_local_v2',
    _get() {
      try { return JSON.parse(localStorage.getItem(this.KEY) || 'null'); } catch (e) { return null; }
    },
    _set(s) { try { localStorage.setItem(this.KEY, JSON.stringify(s)); } catch (e) { /* bỏ qua */ } },
    _state() {
      let s = this._get();
      if (!s) { s = { remaining: C.DEMO_START_CREDITS != null ? C.DEMO_START_CREDITS : 5, used: [] }; this._set(s); }
      return s;
    },
    async balance() { return { remainingExports: this._state().remaining }; },
    async redeem(code) {
      const s = this._state(); code = normCode(code);
      const cfg = root.GS_CODES;
      if (!code) return { success: false, message: 'Vui lòng nhập mã kích hoạt.', remainingExports: s.remaining };
      if (!cfg || !cfg.hashes || !cfg.hashes.length || !(root.crypto && crypto.subtle)) {
        return { success: false, message: 'Trang chưa sẵn sàng kiểm tra mã. Vui lòng tải lại trang.', remainingExports: s.remaining };
      }
      // chống dò mã: tối đa 10 lần sai liên tiếp mỗi 10 phút
      const now = Date.now();
      s.fails = (s.fails || []).filter((t) => now - t < 600000);
      if (s.fails.length >= 10) { this._set(s); return { success: false, message: 'Bạn nhập sai quá nhiều lần. Vui lòng thử lại sau 10 phút.', remainingExports: s.remaining }; }
      const h = await codeHash(code, cfg);
      if (!cfg.hashes.includes(h)) {
        s.fails.push(now); this._set(s);
        return { success: false, message: 'Mã kích hoạt không đúng. Kiểm tra lại từng ký tự.', remainingExports: s.remaining };
      }
      if (s.used.includes(h)) return { success: false, message: 'Mã này đã được sử dụng.', remainingExports: s.remaining };
      const credits = cfg.credits || 5;
      s.used.push(h); s.remaining += credits; s.fails = []; this._set(s);
      return { success: true, message: `Nạp thành công ${credits} lượt xuất!`, addedCredits: credits, remainingExports: s.remaining };
    },
    async consume() {
      const s = this._state();
      if (s.remaining < 1) return { allowed: false, message: 'Bạn đã hết lượt xuất PPTX.', remainingExports: 0 };
      s.remaining -= 1; this._set(s);
      return { allowed: true, message: 'Đã trừ 1 lượt xuất.', remainingExports: s.remaining };
    }
  };

  /* ---------- SUPABASE (gọi stored function qua REST RPC) ---------- */
  async function rpc(fn, body) {
    const url = C.SUPABASE_URL.replace(/\/+$/, '') + '/rest/v1/rpc/' + fn;
    const r = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', apikey: C.SUPABASE_ANON_KEY, Authorization: 'Bearer ' + C.SUPABASE_ANON_KEY },
      body: JSON.stringify(body)
    });
    const data = await r.json().catch(() => null);
    if (!r.ok) throw new Error((data && (data.message || data.hint)) || ('Lỗi máy chủ (' + r.status + ')'));
    return Array.isArray(data) ? data[0] : data;
  }
  const supa = {
    async balance(uid) { const d = await rpc('get_export_balance', { p_user_id: uid }); return { remainingExports: d ? d.remaining_exports : 0 }; },
    async redeem(code, uid) {
      const d = await rpc('redeem_activation_code', { p_code: normCode(code), p_user_id: uid });
      return { success: !!d.ok, message: d.message, addedCredits: d.added_credits || 0, remainingExports: d.remaining_exports };
    },
    async consume(uid, selectedCount, estimatedSlides) {
      const d = await rpc('consume_export_credit', { p_user_id: uid, p_selected_count: selectedCount, p_estimated_slides: estimatedSlides });
      return { allowed: !!d.allowed, message: d.message, remainingExports: d.remaining_exports };
    }
  };

  /* ---------- API Node.js riêng ---------- */
  async function api(path, body) {
    const r = await fetch(C.API_BASE.replace(/\/+$/, '') + path, {
      method: body ? 'POST' : 'GET', headers: { 'Content-Type': 'application/json' }, body: body ? JSON.stringify(body) : undefined
    });
    const data = await r.json().catch(() => ({}));
    if (!r.ok && r.status !== 400 && r.status !== 402) throw new Error(data.message || ('Lỗi máy chủ (' + r.status + ')'));
    return data;
  }
  const node = {
    balance: (uid) => api('/api/balance?userId=' + encodeURIComponent(uid)),
    redeem: (code, uid) => api('/api/redeem-code', { code: normCode(code), userId: uid }),
    consume: (uid, selectedCount, estimatedSlides) => api('/api/consume-export', { userId: uid, selectedCount, estimatedSlides })
  };

  const impl = { demo, supabase: supa, api: node };

  root.GSBackend = {
    mode,
    userId: deviceId,
    balance() { return impl[mode()].balance(deviceId()); },
    redeem(code) { return impl[mode()].redeem(code, deviceId()); },
    consume(selectedCount, estimatedSlides) { return impl[mode()].consume(deviceId(), selectedCount, estimatedSlides); }
  };
})(typeof self !== 'undefined' ? self : this);
