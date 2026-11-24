/*!
 * PPTX Game Studio – Scene builder
 * Mô tả mỗi trò chơi thành 1 "cảnh" (slide) gồm: các hình (shapes), chuỗi hiệu ứng bấm chuột (main)
 * và các trigger (bấm vào 1 hình cụ thể để chạy hiệu ứng). Cùng một mô tả được dùng cho:
 *   - Trình phát HTML (xem trước / trình chiếu online)  -> js/player.js
 *   - Bộ xuất PowerPoint (.pptx) có animation thật         -> js/pptx-export.js
 * Đơn vị: inch, slide 16:9 = 13.333 x 7.5 in. Cỡ chữ: point.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.GameScenes = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const W = 13.333, H = 7.5;
  const F = 'Segoe UI', FB = 'Segoe UI Black', FSYM = 'Segoe UI Symbol';

  const GAME_TYPES = {
    millionaire: { label: 'Ai là triệu phú', icon: '💰' },
    picture: { label: 'Nhìn hình đoán chữ', icon: '🖼️' },
    puzzle: { label: 'Mảnh ghép bí mật', icon: '🧩' },
    truefalse: { label: 'Đúng / Sai', icon: '⚖️' },
    tiles: { label: 'Lật ô số', icon: '🔢' }
  };

  const TRANSITIONS = ['fade', 'push', 'zoom', 'cover', 'wipe'];

  /* ---------------- tiện ích ---------------- */
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const nfc = (s) => String(s == null ? '' : s).normalize('NFC');
  const short = (s, n) => { s = nfc(s).replace(/\s+/g, ' ').trim(); return s.length > n ? s.slice(0, n - 1) + '…' : s; };

  function estLines(text, wIn, pt, factor) {
    const cw = (pt / 72) * factor;
    const perLine = Math.max(1, Math.floor(wIn / cw));
    return String(text).split('\n').reduce((n, para) => {
      const words = para.split(/\s+/).filter(Boolean);
      let lines = 1, cur = 0;
      for (const w of words) {
        const len = w.length;
        if (cur === 0) cur = len;
        else if (cur + 1 + len <= perLine) cur += 1 + len;
        else { lines++; cur = len; }
        while (cur > perLine) { lines++; cur -= perLine; }
      }
      return n + lines;
    }, 0);
  }
  /** Ước lượng cỡ chữ lớn nhất để text vừa khung w x h (inch). */
  function fit(text, w, h, max, min, factor) {
    min = min || 11; factor = factor || 0.58;
    text = nfc(text) || ' ';
    for (let pt = max; pt > min; pt -= 1) {
      if (estLines(text, w, pt, factor) * pt * 1.25 / 72 <= h) return pt;
    }
    return min;
  }
  const E = (target, type, delay, dur) => ({ target, type, delay: delay || 0, dur: dur || 500 });

  class Slide {
    constructor(o) {
      this.key = o.key; this.label = o.label || ''; this.bg = o.bg || '1E1B4B';
      this.transition = o.transition || 'fade';
      this.shapes = []; this.main = []; this.triggers = []; this.notes = o.notes || '';
      this.itemId = o.itemId || null;
    }
    add(o) {
      const s = Object.assign({ kind: 'shape', shape: 'rect', x: 0, y: 0, w: 1, h: 1 }, o);
      if (!s.name) s.name = 'o' + this.shapes.length;
      this.shapes.push(s);
      return s.name;
    }
    auto(effects) { if (effects.length) this.main.unshift({ auto: true, effects }); }
    click(effects) { if (effects.length) this.main.push({ auto: false, effects }); }
    on(target, effects) { if (effects.length) this.triggers.push({ on: target, effects }); }
  }

  /* ---------------- phần trang trí chung ---------------- */
  function deco(sl, c) {
    sl.add({ name: 'bgA', shape: 'ellipse', x: -1.7, y: -2.0, w: 5.4, h: 5.4, fill: c[0], alpha: 84 });
    sl.add({ name: 'bgB', shape: 'ellipse', x: 10.1, y: 4.5, w: 4.8, h: 4.8, fill: c[1], alpha: 87 });
    sl.add({ name: 'bgC', shape: 'ellipse', x: 8.4, y: -1.3, w: 2.5, h: 2.5, fill: c[2], alpha: 89 });
    sl.add({ name: 'bgD', shape: 'ellipse', x: 0.9, y: 6.55, w: 0.45, h: 0.45, fill: c[1], alpha: 55 });
    sl.add({ name: 'bgE', shape: 'ellipse', x: 12.85, y: 1.15, w: 0.26, h: 0.26, fill: c[0], alpha: 45 });
  }

  function header(sl, label, accent, accentText, ctx, timer) {
    sl.add({ name: 'hdr', shape: 'roundRect', radius: 0.27, x: 0.4, y: 0.3, w: 3.35, h: 0.55, fill: accent,
      text: label, size: fit(label, 3.05, 0.45, 17, 10), bold: true, color: accentText, face: F, shadow: true });
    sl.add({ name: 'cnt', kind: 'text', x: 9.6, y: 0.3, w: 2.45, h: 0.55, text: `Câu ${ctx.no} / ${ctx.total}`,
      size: 15, bold: true, color: 'FFFFFF', align: 'right' });
    sl.add({ name: 'home', shape: 'ellipse', x: 12.22, y: 0.26, w: 0.66, h: 0.63, fill: 'FFFFFF', alpha: 82,
      line: 'FFFFFF', lineW: 1.25, text: '⌂', face: FSYM, size: 20, color: 'FFFFFF', bold: true, link: 1, tip: 'Về danh sách trò chơi' });
    const t = parseInt(timer, 10) || 0;
    if (t > 0) {
      sl.add({ name: 'tbtn', shape: 'roundRect', radius: 0.27, x: 3.95, y: 0.3, w: 1.4, h: 0.55, fill: '000000', alpha: 55,
        line: accent, lineW: 1.75, text: `⏱ ${t}s`, size: 15, bold: true, color: 'FFFFFF' });
      sl.add({ name: 'ttrk', shape: 'roundRect', radius: 0.12, x: 5.55, y: 0.45, w: 3.85, h: 0.25, fill: '000000', alpha: 55 });
      sl.add({ name: 'tbar', shape: 'roundRect', radius: 0.12, x: 5.55, y: 0.45, w: 3.85, h: 0.25, fill: accent });
      sl.add({ name: 'tend', shape: 'roundRect', radius: 0.3, x: 5.75, y: 0.22, w: 3.45, h: 0.7, fill: 'DC2626',
        line: 'FFFFFF', lineW: 2, text: 'HẾT GIỜ!', size: 22, bold: true, face: FB, color: 'FFFFFF', shadow: true });
      const ms = t * 1000;
      sl.on('tbtn', [E('tbtn', 'pulse', 0, 300), E('tbar', 'wipeIn', 0, 350), E('tbar', 'timerOut', 400, ms),
        E('tend', 'zoomIn', ms + 400, 400), E('tend', 'pulse', ms + 850, 500), E('tend', 'zoomOut', ms + 2600, 400)]);
    }
  }

  function stamp(sl, name, text, fill, color, rot, y) {
    sl.add({ name, shape: 'roundRect', radius: 0.3, x: 4.17, y: y == null ? 2.2 : y, w: 5.0, h: 1.25, fill, line: 'FFFFFF',
      lineW: 3, text, size: 40, bold: true, face: FB, color, rot, shadow: true });
  }

  /* ---------------- 0. Slide mở đầu (menu) ---------------- */
  function buildTitle(project, entries) {
    const sl = new Slide({ key: 'title', label: 'Mở đầu', bg: '1E1B4B', transition: 'fade' });
    deco(sl, ['6366F1', 'EC4899', 'F59E0B']);
    const title = nfc(project.title) || 'TRÒ CHƠI BÀI GIẢNG';
    const sub = nfc(project.subtitle) || `${entries.length} trò chơi • Bấm vào tên trò chơi để bắt đầu`;
    sl.add({ name: 'ttl', kind: 'text', x: 0.8, y: 0.75, w: 11.73, h: 1.85, text: title.toUpperCase(),
      size: fit(title, 11.3, 1.7, 52, 26, 0.66), bold: true, face: FB, color: 'FFFFFF', shadow: true });
    sl.add({ name: 'bar', shape: 'roundRect', radius: 0.04, x: 5.17, y: 2.7, w: 3.0, h: 0.09, fill: 'F59E0B' });
    sl.add({ name: 'sub', kind: 'text', x: 0.8, y: 2.85, w: 11.73, h: 0.65, text: sub, size: fit(sub, 11.4, 0.6, 22, 12), color: 'C7D2FE' });

    const n = entries.length;
    const fx = [E('ttl', 'zoomIn', 0, 700), E('bar', 'wipeIn', 500, 500), E('sub', 'fadeIn', 700, 600)];
    if (n) {
      const cols = n <= 5 ? 1 : n <= 12 ? 2 : 3;
      const rows = Math.ceil(n / cols);
      const X = cols === 1 ? 2.2 : 0.9, AW = cols === 1 ? 8.93 : 11.53, Y = 3.75, AH = 3.4;
      const gap = 0.14, gx = 0.3;
      const bh = clamp((AH - (rows - 1) * gap) / rows, 0.34, 0.66);
      const bw = (AW - (cols - 1) * gx) / cols;
      entries.forEach((en, i) => {
        const c = Math.floor(i / rows), r = i % rows;
        const gt = GAME_TYPES[en.it.type] || { label: '', icon: '' };
        const label = `${i + 1}. ${gt.label}`;
        const extra = short(en.it.question || en.it.statement || en.it.title || en.it.keyword || '', cols === 1 ? 52 : 26);
        const size = clamp(Math.round(bh * 72 * 0.34), 10, 18);
        sl.add({ name: 'm' + i, shape: 'roundRect', radius: Math.min(0.18, bh / 2), x: X + c * (bw + gx), y: Y + r * (bh + gap), w: bw, h: bh,
          fill: 'FFFFFF', alpha: 87, line: 'A5B4FC', lineW: 1.25, align: 'left', pad: [14, 8, 2, 2], link: en.slideNo, tip: label,
          runs: [{ t: `${gt.icon}  ` }, { t: label, bold: true, color: 'FFFFFF' }, { t: extra ? `  —  ${extra}` : '', color: 'C7D2FE' }], size });
        fx.push(E('m' + i, 'floatIn', 900 + i * 110, 450));
      });
    }
    sl.auto(fx);
    sl.notes = 'Slide mở đầu: bấm vào tên trò chơi để chuyển nhanh tới trò đó. Ở mỗi trò chơi, nút ⌂ ở góc phải sẽ quay lại slide này.';
    return sl;
  }

  /* ---------------- 1. Ai là triệu phú ---------------- */
  function buildMillionaire(it, ctx) {
    const sl = new Slide({ key: 'millionaire', label: 'Ai là triệu phú', bg: '0B1640', transition: ctx.transition, itemId: it.id });
    deco(sl, ['3B82F6', 'F5C542', '8B5CF6']);
    const LINE = 'F5C542', BOX = '16286E';
    header(sl, 'AI LÀ TRIỆU PHÚ', LINE, '0B1640', ctx, it.timer);
    const q = nfc(it.question) || 'Nhập câu hỏi…';
    const answers = (it.answers || []).slice(0, 4).map(nfc);
    while (answers.length < 4) answers.push('');
    const correct = clamp(parseInt(it.correct, 10) || 0, 0, 3);

    sl.add({ name: 'q', shape: 'hexagon', x: 0.6, y: 1.2, w: 10.75, h: 2.05, fill: BOX, line: LINE, lineW: 2.5,
      text: q, size: fit(q, 9.3, 1.75, 30, 14), bold: true, color: 'FFFFFF', shadow: true });
    sl.add({ name: 'ff', shape: 'ellipse', x: 11.6, y: 1.3, w: 1.2, h: 0.88, fill: '1E3A8A', line: LINE, lineW: 2,
      text: '50:50', size: 16, bold: true, color: 'FDE68A', shadow: true });
    sl.add({ name: 'ffl', kind: 'text', x: 11.35, y: 2.22, w: 1.7, h: 0.35, text: 'Trợ giúp', size: 11, color: 'C7D2FE' });

    const pos = [[0.6, 3.55], [6.98, 3.55], [0.6, 4.8], [6.98, 4.8]], bw = 5.75, bh = 1.0, L = 'ABCD';
    const intro = [E('q', 'zoomIn', 0, 650), E('ff', 'zoomIn', 500, 400)];
    answers.forEach((a, i) => {
      const [x, y] = pos[i];
      const size = fit(L[i] + ': ' + a, bw - 1.1, bh - 0.2, 22, 11);
      const base = { shape: 'hexagon', x, y, w: bw, h: bh, lineW: 2, align: 'left', pad: [30, 22, 2, 2], size };
      sl.add(Object.assign({}, base, { name: 'a' + i, fill: BOX, line: LINE, shadow: true,
        runs: [{ t: L[i] + ':  ', color: 'F5C542', bold: true }, { t: a, color: 'FFFFFF' }] }));
      intro.push(E('a' + i, 'floatIn', 650 + i * 160, 450));
    });
    // lớp phủ kết quả cho từng đáp án
    answers.forEach((a, i) => {
      const [x, y] = pos[i];
      const size = fit(L[i] + ': ' + a, bw - 1.1, bh - 0.2, 22, 11);
      const ok = i === correct;
      sl.add({ name: (ok ? 'ok' : 'bad') + i, shape: 'hexagon', x, y, w: bw, h: bh, lineW: 2.5, align: 'left', pad: [30, 22, 2, 2], size,
        fill: ok ? '16A34A' : 'DC2626', line: ok ? 'BBF7D0' : 'FECACA', shadow: true,
        runs: [{ t: L[i] + ':  ', color: 'FFFFFF', bold: true }, { t: a, color: 'FFFFFF', bold: true }] });
    });
    const wrong = [0, 1, 2, 3].filter((i) => i !== correct);
    const drop = [wrong[(correct + 1) % 3], wrong[(correct + 2) % 3]];
    drop.forEach((i) => {
      const [x, y] = pos[i];
      sl.add({ name: 'dim' + i, shape: 'hexagon', x, y, w: bw, h: bh, fill: '0B1640', alpha: 18, line: '334155', lineW: 1.5 });
    });
    const ex = nfc(it.explanation);
    if (ex) sl.add({ name: 'ex', shape: 'roundRect', radius: 0.18, x: 0.6, y: 6.05, w: 12.13, h: 1.05, fill: 'FFF7D6', line: 'F5C542', lineW: 1.5,
      text: '💡 ' + ex, size: fit('💡 ' + ex, 11.6, 0.9, 20, 11), color: '1F2937', bold: false, align: 'left', pad: [16, 12, 2, 2] });
    stamp(sl, 'win', 'CHÍNH XÁC!', 'F5C542', '0B1640', -5, 1.6);
    stamp(sl, 'lose', 'CHƯA ĐÚNG!', 'DC2626', 'FFFFFF', 4, 1.6);

    sl.auto(intro);
    answers.forEach((a, i) => {
      if (i === correct) {
        sl.on('a' + i, [E('ok' + i, 'zoomIn', 0, 450), E('ok' + i, 'pulse', 500, 500), E('win', 'zoomIn', 0, 450),
          E('win', 'zoomOut', 1700, 400)].concat(ex ? [E('ex', 'floatIn', 600, 500)] : []));
      } else {
        sl.on('a' + i, [E('bad' + i, 'fadeIn', 0, 300), E('bad' + i, 'pulse', 300, 400), E('lose', 'zoomIn', 0, 400), E('lose', 'zoomOut', 1500, 350)]);
      }
    });
    sl.on('ff', [E('ff', 'pulse', 0, 300)].concat(drop.map((i, k) => E('dim' + i, 'fadeIn', 250 + k * 200, 500)), [E('ff', 'fadeOut', 900, 400)]));
    sl.click([E('ok' + correct, 'zoomIn', 0, 500), E('ok' + correct, 'pulse', 550, 500)].concat(ex ? [E('ex', 'floatIn', 400, 500)] : []));
    sl.notes = `Ai là triệu phú – Đáp án đúng: ${L[correct]}.\n• Bấm trực tiếp vào ô đáp án học sinh chọn: đúng → xanh + "CHÍNH XÁC!", sai → đỏ + "CHƯA ĐÚNG!".\n• Nút 50:50 loại bỏ 2 phương án sai.${parseInt(it.timer, 10) ? '\n• Bấm nút ⏱ để bắt đầu đếm giờ.' : ''}\n• Bấm ra ngoài (hoặc phím →) để công bố đáp án; bấm tiếp để sang trò chơi sau.`;
    return sl;
  }

  /* ---------------- 2. Nhìn hình đoán chữ ---------------- */
  function layoutLetters(answer, X, Y, AW, AH, maxS) {
    const words = nfc(answer).toUpperCase().trim().split(/\s+/).filter(Boolean).map((w) => Array.from(w));
    if (!words.length) return { s: 0, tiles: [] };
    for (let s = maxS; s >= 0.28; s -= 0.02) {
      const g = s * 0.14, wg = s * 0.6;
      const lines = []; let cur = [], cw = 0, bad = false;
      for (const w of words) {
        const ww = w.length * s + (w.length - 1) * g;
        if (ww > AW) { bad = true; break; }
        if (!cur.length) { cur = [w]; cw = ww; }
        else if (cw + wg + ww <= AW) { cur.push(w); cw += wg + ww; }
        else { lines.push({ words: cur, w: cw }); cur = [w]; cw = ww; }
      }
      if (bad) continue;
      if (cur.length) lines.push({ words: cur, w: cw });
      const lg = s * 0.28, th = lines.length * s + (lines.length - 1) * lg;
      if (th > AH && s > 0.3) continue;
      const tiles = [];
      let y = Y + (AH - th) / 2;
      lines.forEach((ln) => {
        let x = X + (AW - ln.w) / 2;
        ln.words.forEach((w) => { w.forEach((ch) => { tiles.push({ ch, x, y }); x += s + g; }); x += wg - g; });
        y += s + lg;
      });
      return { s, tiles };
    }
    return { s: 0, tiles: [] };
  }

  function buildPicture(it, ctx) {
    const sl = new Slide({ key: 'picture', label: 'Nhìn hình đoán chữ', bg: '0F3D3E', transition: ctx.transition, itemId: it.id });
    deco(sl, ['14B8A6', 'FDE047', '22D3EE']);
    const AC = '2DD4BF';
    header(sl, 'NHÌN HÌNH ĐOÁN CHỮ', AC, '083344', ctx, it.timer);
    sl.add({ name: 'frm', shape: 'roundRect', radius: 0.18, x: 0.6, y: 1.2, w: 6.2, h: 5.75, fill: 'FFFFFF', shadow: true });
    const img = it._img || it.image;
    if (img) sl.add({ name: 'img', kind: 'image', x: 0.75, y: 1.35, w: 5.9, h: 5.45, data: img });
    else sl.add({ name: 'img', kind: 'text', x: 0.75, y: 1.35, w: 5.9, h: 5.45, text: '🖼️\nChưa có ảnh', size: 28, color: '94A3B8' });
    const prompt = nfc(it.question) || 'Hình ảnh này gợi nhắc đến từ khóa nào?';
    sl.add({ name: 'pr', kind: 'text', x: 7.15, y: 1.2, w: 5.6, h: 1.0, text: prompt, size: fit(prompt, 5.4, 0.9, 26, 13), bold: true, color: 'FFFFFF', align: 'left' });
    const fx = [E('frm', 'zoomIn', 0, 600), E('img', 'fadeIn', 300, 700), E('pr', 'floatIn', 500, 500)];
    const hint = nfc(it.hint);
    if (hint) {
      sl.add({ name: 'hb', shape: 'roundRect', radius: 0.18, x: 7.15, y: 2.3, w: 5.6, h: 1.25, fill: 'FFFFFF', alpha: 86, line: AC, lineW: 1.5,
        dash: true, text: '💡  Bấm để xem gợi ý', size: 18, bold: true, color: 'CCFBF1' });
      sl.add({ name: 'hint', shape: 'roundRect', radius: 0.18, x: 7.15, y: 2.3, w: 5.6, h: 1.25, fill: 'FEF3C7', line: 'F59E0B', lineW: 1.5,
        text: 'Gợi ý: ' + hint, size: fit('Gợi ý: ' + hint, 5.3, 1.1, 20, 11), color: '1F2937', align: 'left', pad: [12, 10, 2, 2] });
      fx.push(E('hb', 'floatIn', 700, 500));
      sl.on('hb', [E('hint', 'zoomIn', 0, 450)]);
    }
    const { s, tiles } = layoutLetters(it.answer || '', 7.15, 3.8, 5.6, 2.45, 0.8);
    const letters = tiles.length;
    tiles.forEach((t, k) => {
      sl.add({ name: 'q' + k, shape: 'roundRect', radius: s * 0.18, x: t.x, y: t.y, w: s, h: s, fill: 'FFFFFF', alpha: 80, line: '5EEAD4', lineW: 1.25,
        text: '?', size: Math.round(s * 72 * 0.45), bold: true, color: 'CCFBF1' });
      fx.push(E('q' + k, 'zoomIn', 900 + k * 60, 300));
    });
    tiles.forEach((t, k) => {
      sl.add({ name: 'L' + k, shape: 'roundRect', radius: s * 0.18, x: t.x, y: t.y, w: s, h: s, fill: 'FDE047', line: 'F59E0B', lineW: 1.5,
        text: t.ch, size: Math.round(s * 72 * 0.52), bold: true, face: FB, color: '1F2937', shadow: true });
      sl.on('q' + k, [E('L' + k, 'zoomIn', 0, 350)]);
    });
    const info = letters ? `Từ khóa gồm ${letters} chữ cái • bấm vào ô "?" để mở từng chữ` : 'Chưa nhập từ khóa đáp án';
    sl.add({ name: 'inf', kind: 'text', x: 7.15, y: 6.35, w: 5.6, h: 0.55, text: info, size: 13, color: '99F6E4' });
    fx.push(E('inf', 'fadeIn', 1000, 500));
    sl.auto(fx);
    if (letters) sl.click(tiles.map((t, k) => E('L' + k, 'zoomIn', k * 110, 350)).concat([E('frm', 'pulse', letters * 110 + 200, 600)]));
    sl.notes = `Nhìn hình đoán chữ – Đáp án: ${nfc(it.answer)}.\n• Bấm ô "Bấm để xem gợi ý" để hiện gợi ý.\n• Bấm từng ô "?" để mở từng chữ cái.\n• Bấm ra ngoài (phím →) để mở toàn bộ đáp án.`;
    return sl;
  }

  /* ---------------- 3. Mảnh ghép bí mật ---------------- */
  const PIECE_COLORS = ['F43F5E', 'F59E0B', '10B981', '3B82F6', '8B5CF6', 'EC4899', '14B8A6', 'EAB308', '6366F1'];
  function gridFor(n) { return n <= 4 ? [2, 2] : n <= 6 ? [3, 2] : [3, 3]; }

  function popup(sl, k, title, color, question, answer) {
    const q = nfc(question) || '(Chưa nhập câu hỏi)';
    sl.add({ name: 'P' + k, shape: 'roundRect', radius: 0.22, x: 1.9, y: 1.45, w: 9.53, h: 4.55, fill: 'FFFFFF', line: color, lineW: 4, shadow: true,
      valign: 'top', pad: [26, 26, 8, 22],
      runs: [{ t: title + '\n', color, bold: true, size: 18, face: FB }, { t: q, color: '1F2937', bold: true, size: fit(q, 8.9, 2.3, 32, 14) }] });
    const a = nfc(answer);
    const at = a ? '✔  ' + a : '✔  Mở mảnh ghép';
    sl.add({ name: 'A' + k, shape: 'roundRect', radius: 0.2, x: 2.4, y: 4.6, w: 8.53, h: 1.1, fill: '16A34A', line: 'BBF7D0', lineW: 2, shadow: true,
      text: at, size: fit(at, 8.1, 0.95, 28, 12), bold: true, color: 'FFFFFF' });
  }

  function buildPuzzle(it, ctx) {
    const sl = new Slide({ key: 'puzzle', label: 'Mảnh ghép bí mật', bg: '2E1065', transition: ctx.transition, itemId: it.id });
    deco(sl, ['A855F7', 'F472B6', '60A5FA']);
    const AC = 'C084FC';
    header(sl, 'MẢNH GHÉP BÍ MẬT', AC, '2E1065', ctx, it.timer);
    const n = [4, 6, 9].includes(+it.pieceCount) ? +it.pieceCount : 4;
    const pieces = (it.pieces || []).slice(0, n);
    while (pieces.length < n) pieces.push({ question: '', answer: '' });
    const [cols, rows] = gridFor(n);
    const X = 0.6, Y = 1.2, AW = 7.6, AH = 5.7;
    sl.add({ name: 'frm', shape: 'roundRect', radius: 0.12, x: X - 0.1, y: Y - 0.1, w: AW + 0.2, h: AH + 0.2, fill: 'FFFFFF', shadow: true });
    const img = it._img || it.image;
    if (img) sl.add({ name: 'img', kind: 'image', x: X, y: Y, w: AW, h: AH, data: img });
    const tw = AW / cols, th = AH / rows;
    const fx = [E('frm', 'zoomIn', 0, 600)];
    pieces.forEach((p, k) => {
      const c = k % cols, r = Math.floor(k / cols);
      sl.add({ name: 'T' + k, shape: 'rect', x: X + c * tw, y: Y + r * th, w: tw, h: th, fill: PIECE_COLORS[k], line: 'FFFFFF', lineW: 3,
        text: String(k + 1), size: Math.round(Math.min(tw, th) * 72 * 0.42), bold: true, face: FB, color: 'FFFFFF' });
      fx.push(E('T' + k, 'zoomIn', 350 + k * 110, 400));
    });
    if (img) sl.add({ name: 'full', kind: 'image', x: X, y: Y, w: AW, h: AH, data: img });
    const kw = nfc(it.keyword) || '';
    const blanks = kw ? kw.toUpperCase().split(/\s+/).map((w) => Array.from(w).map(() => '_').join(' ')).join('   ') : '? ? ?';
    sl.add({ name: 'kwl', kind: 'text', x: 8.6, y: 1.2, w: 4.15, h: 0.5, text: 'TỪ KHÓA BÍ MẬT', size: 16, bold: true, color: AC });
    sl.add({ name: 'kwb', shape: 'roundRect', radius: 0.16, x: 8.6, y: 1.75, w: 4.15, h: 1.2, fill: 'FFFFFF', alpha: 86, line: AC, lineW: 1.5,
      text: blanks, size: fit(blanks, 3.9, 1.05, 30, 12, 0.5), bold: true, color: 'FFFFFF' });
    sl.add({ name: 'kw', shape: 'roundRect', radius: 0.16, x: 8.6, y: 1.75, w: 4.15, h: 1.2, fill: 'FDE047', line: 'F59E0B', lineW: 2.5,
      text: (kw || '???').toUpperCase(), size: fit(kw.toUpperCase() || '???', 3.9, 1.05, 34, 12, 0.66), bold: true, face: FB, color: '2E1065', shadow: true });
    const ins = `• Chọn một mảnh ghép (số 1 – ${n}) để mở câu hỏi.\n• Bấm vào khung câu hỏi để xem đáp án.\n• Bấm vào ô đáp án để lật mảnh ghép.\n• Đoán được từ khóa càng sớm, điểm càng cao!`;
    sl.add({ name: 'ins', kind: 'text', x: 8.6, y: 3.2, w: 4.15, h: 3.3, text: ins, size: 14, color: 'E9D5FF', align: 'left', valign: 'top' });
    fx.push(E('kwl', 'fadeIn', 400, 400), E('kwb', 'floatIn', 600, 500), E('ins', 'fadeIn', 900, 600));
    pieces.forEach((p, k) => popup(sl, k, `MẢNH GHÉP SỐ ${k + 1}`, PIECE_COLORS[k], p.question, p.answer));
    sl.auto(fx);
    pieces.forEach((p, k) => {
      sl.on('T' + k, [E('T' + k, 'pulse', 0, 250), E('P' + k, 'zoomIn', 200, 450)]);
      sl.on('P' + k, [E('A' + k, 'floatIn', 0, 450)]);
      sl.on('A' + k, [E('A' + k, 'zoomOut', 0, 300), E('P' + k, 'zoomOut', 150, 400), E('T' + k, 'zoomOut', 600, 550)]);
    });
    sl.click((img ? [E('full', 'fadeIn', 0, 900)] : []).concat([E('kw', 'zoomIn', 700, 550), E('kw', 'pulse', 1350, 600)]));
    sl.notes = `Mảnh ghép bí mật – Từ khóa: ${kw}.\n• Bấm vào số trên mảnh ghép → hiện câu hỏi; bấm vào khung câu hỏi → hiện đáp án; bấm vào đáp án → lật mảnh ghép.\n• Bấm ra ngoài (phím →) để mở toàn bộ hình và công bố từ khóa.`;
    return sl;
  }

  /* ---------------- 4. Đúng / Sai ---------------- */
  function buildTrueFalse(it, ctx) {
    const sl = new Slide({ key: 'truefalse', label: 'Đúng / Sai', bg: '0F172A', transition: ctx.transition, itemId: it.id });
    deco(sl, ['0EA5E9', '22C55E', 'F43F5E']);
    const AC = '38BDF8';
    header(sl, 'ĐÚNG HAY SAI?', AC, '0F172A', ctx, it.timer);
    const st = nfc(it.statement) || 'Nhập nhận định…';
    const isTrue = it.isTrue !== false && it.isTrue !== 'false';
    sl.add({ name: 's', shape: 'roundRect', radius: 0.25, x: 0.9, y: 1.3, w: 11.53, h: 2.2, fill: '1E293B', line: AC, lineW: 2.5, shadow: true,
      text: st, size: fit(st, 10.9, 1.9, 30, 14), bold: true, color: 'FFFFFF', pad: [20, 20, 4, 14] });
    sl.add({ name: 'lbl', shape: 'roundRect', radius: 0.2, x: 5.42, y: 1.08, w: 2.5, h: 0.44, fill: AC, text: 'NHẬN ĐỊNH', size: 13, bold: true, color: '0F172A' });
    const cards = { T: [1.5, 3.85], F: [7.13, 3.85] }, cw = 4.7, ch = 1.85;
    sl.add({ name: 'cT', shape: 'roundRect', radius: 0.3, x: cards.T[0], y: cards.T[1], w: cw, h: ch, fill: '16A34A', line: 'BBF7D0', lineW: 2.5,
      text: '✔  ĐÚNG', size: 40, bold: true, face: FB, color: 'FFFFFF', shadow: true });
    sl.add({ name: 'cF', shape: 'roundRect', radius: 0.3, x: cards.F[0], y: cards.F[1], w: cw, h: ch, fill: 'DC2626', line: 'FECACA', lineW: 2.5,
      text: '✘  SAI', size: 40, bold: true, face: FB, color: 'FFFFFF', shadow: true });
    const good = isTrue ? 'T' : 'F', badk = isTrue ? 'F' : 'T';
    sl.add({ name: 'ring', shape: 'roundRect', radius: 0.38, x: cards[good][0] - 0.16, y: cards[good][1] - 0.16, w: cw + 0.32, h: ch + 0.32, line: 'FDE047', lineW: 6 });
    sl.add({ name: 'dim', shape: 'roundRect', radius: 0.3, x: cards[badk][0], y: cards[badk][1], w: cw, h: ch, fill: '0F172A', alpha: 28 });
    const ex = nfc(it.explanation);
    if (ex) sl.add({ name: 'ex', shape: 'roundRect', radius: 0.18, x: 0.9, y: 6.0, w: 11.53, h: 1.1, fill: 'FFF7D6', line: 'F59E0B', lineW: 1.5,
      text: '💡 ' + ex, size: fit('💡 ' + ex, 11.0, 0.95, 20, 11), color: '1F2937', align: 'left', pad: [16, 12, 2, 2] });
    stamp(sl, 'win', 'CHÍNH XÁC!', 'FDE047', '1F2937', -6, 2.3);
    stamp(sl, 'lose', 'CHƯA ĐÚNG!', 'DC2626', 'FFFFFF', 5, 2.3);
    sl.auto([E('s', 'zoomIn', 0, 600), E('lbl', 'fadeIn', 400, 400), E('cT', 'floatIn', 600, 500), E('cF', 'floatIn', 780, 500)]);
    sl.on('c' + good, [E('c' + good, 'pulse', 0, 450), E('ring', 'zoomIn', 0, 450), E('win', 'zoomIn', 100, 450), E('win', 'zoomOut', 1800, 400)]
      .concat(ex ? [E('ex', 'floatIn', 700, 500)] : []));
    sl.on('c' + badk, [E('c' + badk, 'pulse', 0, 400), E('lose', 'zoomIn', 50, 400), E('lose', 'zoomOut', 1500, 350)]);
    sl.click([E('ring', 'zoomIn', 0, 450), E('dim', 'fadeIn', 200, 500)].concat(ex ? [E('ex', 'floatIn', 500, 500)] : []));
    sl.notes = `Đúng / Sai – Kết luận: ${isTrue ? 'ĐÚNG' : 'SAI'}.\n• Bấm vào thẻ học sinh chọn để chấm ngay.\n• Bấm ra ngoài (phím →) để công bố kết luận và giải thích.`;
    return sl;
  }

  /* ---------------- 5. Lật ô số ---------------- */
  const TILE_COLORS = ['F97316', 'EF4444', 'EAB308', '22C55E', '06B6D4', '8B5CF6', 'EC4899', '14B8A6'];
  function buildTiles(it, ctx) {
    const sl = new Slide({ key: 'tiles', label: 'Lật ô số', bg: '431407', transition: ctx.transition, itemId: it.id });
    deco(sl, ['F97316', 'FACC15', 'EF4444']);
    const AC = 'FB923C';
    header(sl, 'LẬT Ô SỐ', AC, '431407', ctx, it.timer);
    const cells = (it.cells || []).slice(0, 8);
    while (cells.length < 4) cells.push({ question: '', answer: '' });
    const n = cells.length;
    const title = nfc(it.title) || 'Chọn một ô số bất kỳ!';
    sl.add({ name: 'tt', kind: 'text', x: 0.6, y: 1.05, w: 12.13, h: 0.8, text: title, size: fit(title, 11.8, 0.75, 28, 14), bold: true, color: 'FFFFFF' });
    const perRow = n <= 4 ? [n] : [Math.ceil(n / 2), Math.floor(n / 2)];
    const AY = 2.05, AHt = 4.9, gap = 0.3;
    const rows = perRow.length;
    const maxCols = perRow[0];
    const th = Math.min((AHt - (rows - 1) * gap) / rows, 2.6);
    const tw = Math.min((12.13 - (maxCols - 1) * gap) / maxCols, th * 1.35);
    const totalH = rows * th + (rows - 1) * gap;
    const fx = [E('tt', 'fadeIn', 0, 500)];
    let k = 0;
    perRow.forEach((cnt, r) => {
      const rowW = cnt * tw + (cnt - 1) * gap;
      const x0 = (W - rowW) / 2, y = AY + (AHt - totalH) / 2 + r * (th + gap);
      for (let c = 0; c < cnt; c++, k++) {
        const x = x0 + c * (tw + gap);
        const a = nfc(cells[k].answer) || '—';
        sl.add({ name: 'D' + k, shape: 'roundRect', radius: 0.2, x, y, w: tw, h: th, fill: 'FFFFFF', alpha: 86, line: 'FDBA74', lineW: 1.5, dash: true,
          runs: [{ t: `Ô ${k + 1}\n`, color: 'FDBA74', size: 13, bold: true }, { t: a, color: 'FFFFFF', bold: true, size: fit(a, tw - 0.3, th - 0.6, 30, 10) }] });
        sl.add({ name: 'T' + k, shape: 'roundRect', radius: 0.2, x, y, w: tw, h: th, fill: TILE_COLORS[k % 8], line: 'FFFFFF', lineW: 2.5, shadow: true,
          text: String(k + 1), size: Math.round(Math.min(tw, th) * 72 * 0.46), bold: true, face: FB, color: 'FFFFFF' });
        fx.push(E('T' + k, 'zoomIn', 250 + k * 120, 420));
      }
    });
    cells.forEach((c, i) => popup(sl, i, `Ô SỐ ${i + 1}`, TILE_COLORS[i % 8], c.question, c.answer));
    sl.auto(fx);
    cells.forEach((c, i) => {
      sl.on('T' + i, [E('T' + i, 'pulse', 0, 250), E('P' + i, 'zoomIn', 200, 450)]);
      sl.on('P' + i, [E('A' + i, 'floatIn', 0, 450)]);
      sl.on('A' + i, [E('A' + i, 'zoomOut', 0, 300), E('P' + i, 'zoomOut', 150, 400), E('D' + i, 'fadeIn', 550, 10), E('T' + i, 'zoomOut', 550, 550)]);
    });
    sl.notes = `Lật ô số – ${n} ô.\n• Bấm vào ô số → hiện câu hỏi; bấm vào khung câu hỏi → hiện đáp án; bấm vào đáp án → ô được lật và ghi lại đáp án trên bảng.\n• Bảng cuối cùng chính là bảng tổng kết đáp án.`;
    return sl;
  }

  /* ---------------- slide kết thúc ---------------- */
  function buildEnd(project) {
    const sl = new Slide({ key: 'end', label: 'Kết thúc', bg: '1E1B4B', transition: 'zoom' });
    deco(sl, ['EC4899', '6366F1', '22C55E']);
    const conf = ['F43F5E', 'F59E0B', '10B981', '3B82F6', 'A855F7', 'FDE047', '22D3EE', 'FB923C'];
    const fx = [];
    for (let i = 0; i < 18; i++) {
      const x = 0.5 + ((i * 2.17) % 12.2), y = 0.4 + ((i * 1.31) % 6.4);
      const sz = 0.18 + (i % 4) * 0.08;
      sl.add({ name: 'c' + i, shape: i % 3 === 0 ? 'star5' : i % 3 === 1 ? 'ellipse' : 'rect', x, y, w: sz, h: sz, fill: conf[i % 8], rot: (i * 37) % 90 });
      fx.push(E('c' + i, 'zoomIn', 700 + i * 60, 400));
    }
    const msg = nfc(project.endMessage) || 'Cảm ơn các em đã tham gia!';
    sl.add({ name: 'big', kind: 'text', x: 0.8, y: 2.0, w: 11.73, h: 1.7, text: 'HOÀN THÀNH!', size: 66, bold: true, face: FB, color: 'FDE047', shadow: true });
    sl.add({ name: 'msg', kind: 'text', x: 0.8, y: 3.75, w: 11.73, h: 0.8, text: msg, size: fit(msg, 11.4, 0.75, 28, 14), color: 'FFFFFF' });
    sl.add({ name: 'back', shape: 'roundRect', radius: 0.28, x: 4.67, y: 5.0, w: 4.0, h: 0.62, fill: 'FFFFFF', alpha: 85, line: 'A5B4FC', lineW: 1.5,
      text: '⌂  Về danh sách trò chơi', size: 16, bold: true, color: 'FFFFFF', link: 1 });
    sl.auto([E('big', 'zoomIn', 0, 700), E('big', 'pulse', 800, 600), E('msg', 'floatIn', 500, 600), E('back', 'fadeIn', 1100, 500)].concat(fx));
    return sl;
  }

  const BUILDERS = { millionaire: buildMillionaire, picture: buildPicture, puzzle: buildPuzzle, truefalse: buildTrueFalse, tiles: buildTiles };

  /** Dựng toàn bộ bài trình chiếu từ project (chỉ các mục includedInExport). */
  function buildDeck(project) {
    const items = (project.items || []).filter((i) => i.includedInExport !== false && BUILDERS[i.type]);
    const entries = items.map((it, i) => ({ it, slideNo: i + 2 }));
    const slides = [buildTitle(project, entries)];
    items.forEach((it, i) => slides.push(BUILDERS[it.type](it, { no: i + 1, total: items.length, transition: TRANSITIONS[(i + 1) % TRANSITIONS.length] })));
    slides.push(buildEnd(project));
    return { title: nfc(project.title) || 'Trò chơi bài giảng', W, H, slides };
  }

  /* ---------------- dữ liệu mặc định ---------------- */
  let seq = 0;
  const uid = () => 'it' + Date.now().toString(36) + (seq++).toString(36) + Math.random().toString(36).slice(2, 6);
  function newItem(type) {
    const base = { id: uid(), type, includedInExport: true, timer: 0 };
    switch (type) {
      case 'millionaire': return Object.assign(base, { question: '', answers: ['', '', '', ''], correct: 0, explanation: '', timer: 30 });
      case 'picture': return Object.assign(base, { question: 'Hình ảnh này gợi nhắc đến từ khóa nào?', image: '', hint: '', answer: '', timer: 0 });
      case 'puzzle': return Object.assign(base, { image: '', pieceCount: 4, keyword: '', pieces: [0, 1, 2, 3].map(() => ({ question: '', answer: '' })) });
      case 'truefalse': return Object.assign(base, { statement: '', isTrue: true, explanation: '', timer: 15 });
      case 'tiles': return Object.assign(base, { title: 'Chọn một ô số bất kỳ!', cells: [0, 1, 2, 3].map(() => ({ question: '', answer: '' })) });
    }
    return base;
  }

  return { W, H, GAME_TYPES, buildDeck, newItem, fit, layoutLetters, _builders: BUILDERS };
});
