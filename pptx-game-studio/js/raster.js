/*!
 * PPTX Game Studio – Chế độ KHÓA NỘI DUNG.
 * Vẽ từng hình của slide thành ảnh PNG (chữ, khung, màu… được "in" vào ảnh) để file PowerPoint
 * xuất ra không còn chữ/hình chỉnh sửa được: người dùng không thể sửa câu hỏi hay sao chép bố cục
 * để tự làm lại, nhưng hiệu ứng, trigger, âm thanh vẫn chạy y nguyên.
 * Các hình tĩnh nằm dưới cùng được gộp chung thành 1 ảnh nền.
 */
(function (root) {
  'use strict';
  const K = 144; // điểm ảnh / inch (slide 1920 x 1080)
  const W = 13.333, H = 7.5;

  function rgba(hex, transp) {
    const n = parseInt(String(hex || '000000').replace('#', ''), 16);
    return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${1 - (transp || 0) / 100})`;
  }
  function family(face) {
    if (face === 'Segoe UI Black') return "'Segoe UI Black','Segoe UI',system-ui,sans-serif";
    if (face === 'Segoe UI Symbol') return "'Segoe UI Symbol','Segoe UI Emoji','Segoe UI',system-ui,sans-serif";
    return "'Segoe UI','Segoe UI Emoji',system-ui,-apple-system,Roboto,Arial,sans-serif";
  }
  const imgCache = new Map();
  function loadImg(src) {
    if (imgCache.has(src)) return imgCache.get(src);
    const p = new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; });
    imgCache.set(src, p);
    return p;
  }

  function path(g, sh, w, h) {
    g.beginPath();
    const s = sh.shape;
    if (sh.kind === 'text') { g.rect(0, 0, w, h); return; }
    if (s === 'ellipse') { g.ellipse(w / 2, h / 2, w / 2, h / 2, 0, 0, Math.PI * 2); return; }
    if (s === 'roundRect') {
      const r = Math.min((sh.radius || 0.1) * K, w / 2, h / 2);
      g.moveTo(r, 0); g.arcTo(w, 0, w, h, r); g.arcTo(w, h, 0, h, r); g.arcTo(0, h, 0, 0, r); g.arcTo(0, 0, w, 0, r); g.closePath(); return;
    }
    if (s === 'hexagon') {
      const d = 0.25 * Math.min(w, h);
      g.moveTo(d, 0); g.lineTo(w - d, 0); g.lineTo(w, h / 2); g.lineTo(w - d, h); g.lineTo(d, h); g.lineTo(0, h / 2); g.closePath(); return;
    }
    if (s === 'star5') {
      const P = [[50, 0], [61, 35], [98, 35], [68, 57], [79, 91], [50, 70], [21, 91], [32, 57], [2, 35], [39, 35]];
      P.forEach(([px, py], i) => (i ? g.lineTo : g.moveTo).call(g, (px / 100) * w, (py / 100) * h)); g.closePath(); return;
    }
    g.rect(0, 0, w, h);
  }

  function runsOf(sh) {
    const base = { color: sh.color || 'FFFFFF', size: sh.size || 18, bold: !!sh.bold, face: sh.face };
    if (sh.runs) return sh.runs.map((r) => Object.assign({}, base, r.color ? { color: r.color } : {}, r.size ? { size: r.size } : {},
      r.bold != null ? { bold: r.bold } : {}, r.face ? { face: r.face } : {}, { t: r.t || '' }));
    return [Object.assign({}, base, { t: sh.text || '' })];
  }
  const fontOf = (r) => `${r.face === 'Segoe UI Black' ? 900 : r.bold ? 700 : 400} ${(r.size / 72) * K}px ${family(r.face)}`;

  function drawText(g, sh, w, h) {
    const runs = runsOf(sh);
    if (!runs.some((r) => r.t)) return;
    const pad = (sh.pad || [6, 6, 3, 3]).map((p) => (p / 72) * K); // [l, r, b, t]
    const maxW = Math.max(10, w - pad[0] - pad[1]);
    // tách thành các "mẩu" (từ + khoảng trắng) giữ định dạng
    const lines = [[]];
    let lineW = 0;
    const push = (piece) => {
      g.font = fontOf(piece);
      const pw = g.measureText(piece.t).width;
      const cur = lines[lines.length - 1];
      if (lineW + pw > maxW && cur.length && piece.t.trim()) { lines.push([]); lineW = 0; if (!piece.t.trim()) return; }
      const tgt = lines[lines.length - 1];
      if (!tgt.length && !piece.t.trim()) return; // bỏ khoảng trắng đầu dòng
      tgt.push(Object.assign({}, piece, { w: pw }));
      lineW += pw;
    };
    runs.forEach((r) => {
      String(r.t).split('\n').forEach((para, pi) => {
        if (pi > 0) { lines.push([]); lineW = 0; }
        para.split(/(\s+)/).filter((x) => x !== '').forEach((tok) => {
          g.font = fontOf(r);
          if (g.measureText(tok).width > maxW && tok.trim()) { // từ quá dài: cắt theo ký tự
            let buf = '';
            Array.from(tok).forEach((ch) => { if (g.measureText(buf + ch).width > maxW && buf) { push(Object.assign({}, r, { t: buf })); buf = ''; } buf += ch; });
            if (buf) push(Object.assign({}, r, { t: buf }));
          } else push(Object.assign({}, r, { t: tok }));
        });
      });
    });
    while (lines.length > 1 && !lines[lines.length - 1].length) lines.pop();
    const lh = (ln) => Math.max(...(ln.length ? ln : [runs[runs.length - 1]]).map((p) => (p.size / 72) * K)) * 1.2;
    const total = lines.reduce((s, ln) => s + lh(ln), 0);
    let y = sh.valign === 'top' ? pad[3] : sh.valign === 'bottom' ? h - pad[2] - total : pad[3] + (h - pad[3] - pad[2] - total) / 2;
    g.textBaseline = 'alphabetic';
    lines.forEach((ln) => {
      const LH = lh(ln);
      // bỏ khoảng trắng cuối dòng khi tính căn lề
      let trimmed = ln.slice();
      while (trimmed.length && !trimmed[trimmed.length - 1].t.trim()) trimmed.pop();
      const lw = trimmed.reduce((s, p) => s + p.w, 0);
      const al = sh.align || 'center';
      let x = al === 'left' ? pad[0] : al === 'right' ? w - pad[1] - lw : pad[0] + (maxW - lw) / 2;
      const base = y + LH * 0.8;
      trimmed.forEach((p) => { g.font = fontOf(p); g.fillStyle = rgba(p.color); g.fillText(p.t, x, base); x += p.w; });
      y += LH;
    });
  }

  async function drawShape(g, sh, x, y) {
    const w = sh.w * K, h = sh.h * K;
    g.save();
    g.translate(x, y);
    if (sh.kind === 'image') {
      if (sh.data) {
        const im = await loadImg(sh.data);
        const ar = w / h, ir = im.naturalWidth / im.naturalHeight;
        let sw = im.naturalWidth, shh = im.naturalHeight, sx = 0, sy = 0;
        if (ir > ar) { sw = shh * ar; sx = (im.naturalWidth - sw) / 2; } else { shh = sw / ar; sy = (im.naturalHeight - shh) / 2; }
        g.drawImage(im, sx, sy, sw, shh, 0, 0, w, h);
      }
      g.restore(); return;
    }
    if (sh.kind !== 'text' && (sh.fill || sh.line)) {
      path(g, sh, w, h);
      if (sh.fill) {
        if (sh.shadow) { g.shadowColor = 'rgba(0,0,0,.38)'; g.shadowBlur = (10 / 72) * K; g.shadowOffsetY = (3 / 72) * K; }
        g.fillStyle = rgba(sh.fill, sh.alpha); g.fill();
        g.shadowColor = 'transparent'; g.shadowBlur = 0; g.shadowOffsetY = 0;
      }
      if (sh.line) {
        const lw = ((sh.lineW || 1) / 72) * K;
        g.lineWidth = lw; g.strokeStyle = rgba(sh.line);
        if (sh.dash) g.setLineDash([lw * 4, lw * 3]);
        g.stroke(); g.setLineDash([]);
      }
    }
    drawText(g, sh, w, h);
    g.restore();
  }

  function canvas(w, h) { const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(w)); c.height = Math.max(1, Math.round(h)); return c; }

  /** Vẽ 1 hình thành ảnh PNG trong suốt, giữ nguyên tâm hình (để hiệu ứng phóng to/thu nhỏ đúng tâm). */
  async function shapeImage(sh) {
    const pad = sh.shadow ? 0.16 : 0.04;
    const rad = ((sh.rot || 0) * Math.PI) / 180;
    const bw = Math.abs(sh.w * Math.cos(rad)) + Math.abs(sh.h * Math.sin(rad));
    const bh = Math.abs(sh.w * Math.sin(rad)) + Math.abs(sh.h * Math.cos(rad));
    const cw = bw + pad * 2, ch = bh + pad * 2;
    const c = canvas(cw * K, ch * K), g = c.getContext('2d');
    g.translate((cw * K) / 2, (ch * K) / 2);
    if (rad) g.rotate(rad);
    await drawShape(g, sh, (-sh.w * K) / 2, (-sh.h * K) / 2);
    const cx = sh.x + sh.w / 2, cy = sh.y + sh.h / 2;
    return { data: c.toDataURL('image/png'), x: cx - cw / 2, y: cy - ch / 2, w: cw, h: ch };
  }

  /** Vẽ cả slide: trả về ảnh nền (gồm các hình tĩnh phía dưới) + ảnh từng hình còn lại. */
  async function rasterizeSlide(sd) {
    if (document.fonts && document.fonts.ready) await document.fonts.ready;
    const active = new Set();
    sd.main.forEach((gr) => gr.effects.forEach((e) => active.add(e.target)));
    sd.triggers.forEach((t) => { active.add(t.on); t.effects.forEach((e) => active.add(e.target)); });
    sd.shapes.forEach((s) => { if (s.link) active.add(s.name); });
    const merged = new Set();
    for (const s of sd.shapes) { if (active.has(s.name)) break; merged.add(s.name); }

    const bg = canvas(W * K, H * K), g = bg.getContext('2d');
    g.fillStyle = rgba(sd.bg); g.fillRect(0, 0, bg.width, bg.height);
    for (const s of sd.shapes) {
      if (!merged.has(s.name)) continue;
      g.save();
      const cx = (s.x + s.w / 2) * K, cy = (s.y + s.h / 2) * K;
      g.translate(cx, cy); if (s.rot) g.rotate((s.rot * Math.PI) / 180);
      await drawShape(g, s, (-s.w * K) / 2, (-s.h * K) / 2);
      g.restore();
    }
    // dấu bản quyền mờ
    g.font = `600 ${0.13 * K}px ${family()}`; g.fillStyle = 'rgba(255,255,255,.28)'; g.textAlign = 'right';
    g.fillText('Tạo bởi PPTX Game Studio', W * K - 0.15 * K, H * K - 0.12 * K);
    const images = {};
    for (const s of sd.shapes) {
      if (merged.has(s.name)) continue;
      if (s.kind === 'image' && !s.data) continue;
      images[s.name] = await shapeImage(s);
    }
    return { bg: bg.toDataURL('image/jpeg', 0.92), merged, images };
  }

  root.GSRaster = { rasterizeSlide, shapeImage };
})(typeof self !== 'undefined' ? self : this);
