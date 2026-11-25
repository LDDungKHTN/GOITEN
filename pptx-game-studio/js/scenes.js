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
    tiles: { label: 'Lật ô số', icon: '🔢' },
    obstacle: { label: 'Vượt chướng ngại vật', icon: '🏃' },
    rescue: { label: 'Giải cứu thú cưng', icon: '🐶' },
    fruit: { label: 'Hứng hoa quả', icon: '🍎' },
    space: { label: 'Bảo vệ Trái Đất', icon: '🌍' }
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
  // Hiệu ứng: target = tên hình, type = loại, delay/dur (ms), snd = âm thanh phát cùng (tùy chọn)
  const E = (target, type, delay, dur, snd) => {
    const e = { target, type, delay: delay || 0, dur: dur || 500 };
    if (snd) e.snd = snd;
    return e;
  };
  // Hiệu ứng "wipe" theo hướng: 'left' (từ trái), 'right', 'up' (từ trên), 'down'
  const EW = (target, dir, delay, dur, snd) => Object.assign(E(target, 'wipeIn', delay, dur, snd), { dir });

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
      sl.on('tbtn', [E('tbtn', 'pulse', 0, 300), E('tbar', 'wipeIn', 0, 350), E('tbar', 'timerOut', 400, ms, 'tick' + t),
        E('tend', 'zoomIn', ms + 400, 400, 'timeup'), E('tend', 'pulse', ms + 850, 500), E('tend', 'zoomOut', ms + 2600, 400)]);
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
    const fx = [E('ttl', 'zoomIn', 0, 700, 'intro'), E('bar', 'wipeIn', 500, 500), E('sub', 'fadeIn', 700, 600)];
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
        const room = (cols === 1 ? 72 : cols === 2 ? 36 : 24) - label.length - 5;
        let extra = room >= 8 ? short(en.it.question || en.it.statement || en.it.title || en.it.keyword || '', room) : '';
        const size = Math.min(clamp(Math.round(bh * 72 * 0.34), 10, 18), fit(`${gt.icon}  ${label}  —  ${extra}`, bw - 0.5, bh - 0.08, 18, 9, 0.64));
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
    const intro = [E('q', 'zoomIn', 0, 650, 'reveal'), E('ff', 'zoomIn', 500, 400)];
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
        sl.on('a' + i, [E('ok' + i, 'zoomIn', 0, 450, 'correct'), E('ok' + i, 'pulse', 500, 500), E('win', 'zoomIn', 0, 450),
          E('win', 'zoomOut', 1700, 400)].concat(ex ? [E('ex', 'floatIn', 600, 500)] : []));
      } else {
        sl.on('a' + i, [E('bad' + i, 'fadeIn', 0, 300, 'wrong'), E('bad' + i, 'pulse', 300, 400), E('lose', 'zoomIn', 0, 400), E('lose', 'zoomOut', 1500, 350)]);
      }
    });
    sl.on('ff', [E('ff', 'pulse', 0, 300, 'whoosh')].concat(drop.map((i, k) => E('dim' + i, 'fadeIn', 250 + k * 200, 500)), [E('ff', 'fadeOut', 900, 400)]));
    sl.click([E('ok' + correct, 'zoomIn', 0, 500, 'correct'), E('ok' + correct, 'pulse', 550, 500)].concat(ex ? [E('ex', 'floatIn', 400, 500)] : []));
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
    const fx = [E('frm', 'zoomIn', 0, 600, 'reveal'), E('img', 'fadeIn', 300, 700), E('pr', 'floatIn', 500, 500)];
    const hint = nfc(it.hint);
    if (hint) {
      sl.add({ name: 'hb', shape: 'roundRect', radius: 0.18, x: 7.15, y: 2.3, w: 5.6, h: 1.25, fill: 'FFFFFF', alpha: 86, line: AC, lineW: 1.5,
        dash: true, text: '💡  Bấm để xem gợi ý', size: 18, bold: true, color: 'CCFBF1' });
      sl.add({ name: 'hint', shape: 'roundRect', radius: 0.18, x: 7.15, y: 2.3, w: 5.6, h: 1.25, fill: 'FEF3C7', line: 'F59E0B', lineW: 1.5,
        text: 'Gợi ý: ' + hint, size: fit('Gợi ý: ' + hint, 5.3, 1.1, 20, 11), color: '1F2937', align: 'left', pad: [12, 10, 2, 2] });
      fx.push(E('hb', 'floatIn', 700, 500));
      sl.on('hb', [E('hint', 'zoomIn', 0, 450, 'reveal')]);
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
      sl.on('q' + k, [E('L' + k, 'zoomIn', 0, 350, 'pop')]);
    });
    const info = letters ? `Từ khóa gồm ${letters} chữ cái • bấm vào ô "?" để mở từng chữ` : 'Chưa nhập từ khóa đáp án';
    sl.add({ name: 'inf', kind: 'text', x: 7.15, y: 6.35, w: 5.6, h: 0.55, text: info, size: 13, color: '99F6E4' });
    fx.push(E('inf', 'fadeIn', 1000, 500));
    sl.auto(fx);
    if (letters) sl.click(tiles.map((t, k) => E('L' + k, 'zoomIn', k * 110, 350, k === 0 ? 'reveal' : null)).concat([E('frm', 'pulse', letters * 110 + 200, 600, 'win')]));
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
    const fx = [E('frm', 'zoomIn', 0, 600, 'reveal')];
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
      sl.on('T' + k, [E('T' + k, 'pulse', 0, 250, 'pop'), E('P' + k, 'zoomIn', 200, 450)]);
      sl.on('P' + k, [E('A' + k, 'floatIn', 0, 450, 'reveal')]);
      sl.on('A' + k, [E('A' + k, 'zoomOut', 0, 300), E('P' + k, 'zoomOut', 150, 400), E('T' + k, 'zoomOut', 600, 550, 'whoosh')]);
    });
    sl.click((img ? [E('full', 'fadeIn', 0, 900, 'win')] : []).concat([E('kw', 'zoomIn', 700, 550, img ? null : 'win'), E('kw', 'pulse', 1350, 600)]));
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
    sl.auto([E('s', 'zoomIn', 0, 600, 'reveal'), E('lbl', 'fadeIn', 400, 400), E('cT', 'floatIn', 600, 500), E('cF', 'floatIn', 780, 500)]);
    sl.on('c' + good, [E('c' + good, 'pulse', 0, 450, 'correct'), E('ring', 'zoomIn', 0, 450), E('win', 'zoomIn', 100, 450), E('win', 'zoomOut', 1800, 400)]
      .concat(ex ? [E('ex', 'floatIn', 700, 500)] : []));
    sl.on('c' + badk, [E('c' + badk, 'pulse', 0, 400, 'wrong'), E('lose', 'zoomIn', 50, 400), E('lose', 'zoomOut', 1500, 350)]);
    sl.click([E('ring', 'zoomIn', 0, 450, 'correct'), E('dim', 'fadeIn', 200, 500)].concat(ex ? [E('ex', 'floatIn', 500, 500)] : []));
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
    const fx = [E('tt', 'fadeIn', 0, 500, 'reveal')];
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
      sl.on('T' + i, [E('T' + i, 'pulse', 0, 250, 'pop'), E('P' + i, 'zoomIn', 200, 450)]);
      sl.on('P' + i, [E('A' + i, 'floatIn', 0, 450, 'reveal')]);
      sl.on('A' + i, [E('A' + i, 'zoomOut', 0, 300), E('P' + i, 'zoomOut', 150, 400), E('D' + i, 'fadeIn', 550, 10), E('T' + i, 'zoomOut', 550, 550, 'whoosh')]);
    });
    sl.notes = `Lật ô số – ${n} ô.\n• Bấm vào ô số → hiện câu hỏi; bấm vào khung câu hỏi → hiện đáp án; bấm vào đáp án → ô được lật và ghi lại đáp án trên bảng.\n• Bảng cuối cùng chính là bảng tổng kết đáp án.`;
    return sl;
  }

  /* ---------------- 6. Vượt chướng ngại vật ---------------- */
  // chỉ dùng emoji có sẵn trên Windows 10 / PowerPoint 2016+
  const OBST_ICONS = ['🧱', '🔥', '🌊', '🌵', '⛰️', '⚡', '🐊', '🌋'];
  function buildObstacle(it, ctx) {
    const sl = new Slide({ key: 'obstacle', label: 'Vượt chướng ngại vật', bg: '14532D', transition: ctx.transition, itemId: it.id });
    deco(sl, ['22C55E', 'FACC15', '38BDF8']);
    const AC = '4ADE80';
    header(sl, 'VƯỢT CHƯỚNG NGẠI VẬT', AC, '14532D', ctx, it.timer);
    const steps = (it.steps || []).slice(0, 8);
    while (steps.length < 3) steps.push({ question: '', answer: '' });
    const n = steps.length;
    const title = nfc(it.title) || 'Vượt qua tất cả chướng ngại vật để về đích!';
    sl.add({ name: 'tt', kind: 'text', x: 0.6, y: 0.98, w: 12.13, h: 0.55, text: title, size: fit(title, 11.8, 0.5, 22, 12), bold: true, color: 'FFFFFF' });

    // Các điểm trên đường đua: xuất phát, n chướng ngại vật, đích
    const pts = [];
    if (n <= 4) {
      const y = 4.25;
      for (let i = 0; i <= n + 1; i++) pts.push({ x: 1.0 + (11.33 * i) / (n + 1), y });
    } else {
      // hàng 1: xuất phát + k1 chướng ngại (trái → phải); hàng 2: k2 chướng ngại + đích (phải → trái)
      const k1 = Math.ceil(n / 2), k2 = n - k1, y1 = 3.0, y2 = 5.75, xl = 1.0, xr = 12.2;
      for (let i = 0; i <= k1; i++) pts.push({ x: xl + ((xr - xl) * i) / k1, y: y1 });
      for (let i = 0; i <= k2; i++) pts.push({ x: xr - ((xr - xl) * i) / k2, y: y2 });
    }
    // Đường đua (nền) và vệt tiến trình (hiện dần theo hướng chạy)
    const RW = 0.42;
    const segs = [];
    for (let i = 1; i < pts.length; i++) {
      const a = pts[i - 1], b = pts[i];
      segs.push(a.y === b.y
        ? { x: Math.min(a.x, b.x) - RW / 2, y: a.y - RW / 2, w: Math.abs(b.x - a.x) + RW, h: RW, dir: b.x >= a.x ? 'left' : 'right' }
        : { x: a.x - RW / 2, y: Math.min(a.y, b.y) - RW / 2, w: RW, h: Math.abs(b.y - a.y) + RW, dir: b.y >= a.y ? 'up' : 'down' });
    }
    segs.forEach((s, i) => sl.add({ name: 'rd' + i, shape: 'roundRect', radius: RW / 2, x: s.x, y: s.y, w: s.w, h: s.h, fill: '0F172A', alpha: 35, line: '86EFAC', lineW: 1.25, dash: true }));
    segs.forEach((s, i) => sl.add({ name: 'pg' + i, shape: 'roundRect', radius: RW / 2, x: s.x, y: s.y, w: s.w, h: s.h, fill: 'FACC15' }));
    const S = 1.15, R = 0.82;
    const st = pts[0], fin = pts[n + 1];
    sl.add({ name: 'start', shape: 'roundRect', radius: 0.18, x: st.x - 0.62, y: st.y + 0.55, w: 1.24, h: 0.42, fill: 'FFFFFF', alpha: 85, text: 'XUẤT PHÁT', size: 11, bold: true, color: 'FFFFFF' });
    sl.add({ name: 'fin', shape: 'ellipse', x: fin.x - 0.55, y: fin.y - 0.55, w: 1.1, h: 1.1, fill: 'FFFFFF', line: 'FACC15', lineW: 3, text: '🏁', size: 30, shadow: true });
    sl.add({ name: 'finl', kind: 'text', x: fin.x - 0.8, y: fin.y + 0.58, w: 1.6, h: 0.4, text: 'ĐÍCH', size: 13, bold: true, color: 'FDE047' });
    const fx = [E('tt', 'fadeIn', 0, 500, 'reveal')];
    steps.forEach((s, k) => {
      const p = pts[k + 1];
      sl.add({ name: 'D' + k, shape: 'ellipse', x: p.x - 0.42, y: p.y - 0.42, w: 0.84, h: 0.84, fill: '16A34A', line: 'FFFFFF', lineW: 2.5, text: '✔', size: 26, bold: true, color: 'FFFFFF' });
      sl.add({ name: 'O' + k, shape: 'roundRect', radius: 0.22, x: p.x - S / 2, y: p.y - S / 2, w: S, h: S, fill: PIECE_COLORS[k % PIECE_COLORS.length], line: 'FFFFFF', lineW: 2.5, shadow: true,
        runs: [{ t: OBST_ICONS[k % OBST_ICONS.length] + '\n', size: 26 }, { t: String(k + 1), size: 20, bold: true, face: FB, color: 'FFFFFF' }], size: 20, pad: [2, 2, 2, 2] });
      fx.push(E('O' + k, 'zoomIn', 400 + k * 130, 420));
    });
    // nhân vật ở từng vị trí (chỉ hiện 1 vị trí tại một thời điểm)
    const runner = (name, p, dy) => sl.add({ name, shape: 'ellipse', x: p.x - R / 2, y: p.y - R / 2 - (dy || 0), w: R, h: R, fill: 'FDE047', line: 'F59E0B', lineW: 3, text: '🏃', size: 26, shadow: true });
    runner('R0', st, 0);
    steps.forEach((s, k) => runner('R' + (k + 1), pts[k + 1], 0.95));
    runner('RF', fin, 0.95);
    sl.add({ name: 'goal', shape: 'roundRect', radius: 0.3, x: 3.67, y: 2.75, w: 6.0, h: 1.5, fill: 'FDE047', line: 'FFFFFF', lineW: 4, text: '🏆  VỀ ĐÍCH!', size: 48, bold: true, face: FB, color: '14532D', shadow: true, rot: -4 });
    fx.push(E('R0', 'zoomIn', 400 + n * 130, 450, 'step'));

    // Hộp câu hỏi cho từng chướng ngại vật
    steps.forEach((s, k) => {
      const color = PIECE_COLORS[k % PIECE_COLORS.length];
      const q = nfc(s.question) || '(Chưa nhập câu hỏi)';
      sl.add({ name: 'P' + k, shape: 'roundRect', radius: 0.22, x: 1.9, y: 1.3, w: 9.53, h: 4.95, fill: 'FFFFFF', line: color, lineW: 4, shadow: true, valign: 'top', pad: [26, 26, 8, 20],
        runs: [{ t: `CHƯỚNG NGẠI VẬT ${k + 1}  ${OBST_ICONS[k % OBST_ICONS.length]}\n`, color, bold: true, size: 18, face: FB }, { t: q, color: '1F2937', bold: true, size: fit(q, 8.9, 1.9, 30, 13) }] });
      const a = nfc(s.answer), at = a ? 'Đáp án: ' + a : 'Đáp án: (chưa nhập)';
      sl.add({ name: 'A' + k, shape: 'roundRect', radius: 0.18, x: 2.4, y: 3.75, w: 8.53, h: 1.0, fill: 'ECFDF5', line: '16A34A', lineW: 2, text: at, size: fit(at, 8.1, 0.85, 26, 12), bold: true, color: '14532D' });
      sl.add({ name: 'OK' + k, shape: 'roundRect', radius: 0.2, x: 2.4, y: 5.0, w: 4.15, h: 0.9, fill: '16A34A', line: 'BBF7D0', lineW: 2, text: '✔  Vượt qua', size: 22, bold: true, color: 'FFFFFF', shadow: true });
      sl.add({ name: 'NO' + k, shape: 'roundRect', radius: 0.2, x: 6.78, y: 5.0, w: 4.15, h: 0.9, fill: 'DC2626', line: 'FECACA', lineW: 2, text: '✘  Chưa vượt qua', size: 22, bold: true, color: 'FFFFFF', shadow: true });
    });
    sl.auto(fx);
    const close = (k) => [E('A' + k, 'zoomOut', 0, 300), E('OK' + k, 'zoomOut', 0, 300), E('NO' + k, 'zoomOut', 0, 300), E('P' + k, 'zoomOut', 120, 400)];
    steps.forEach((s, k) => {
      const prev = 'R' + k;
      sl.on('O' + k, [E('O' + k, 'pulse', 0, 250, 'pop'), E('P' + k, 'zoomIn', 200, 450)]);
      sl.on('P' + k, [E('A' + k, 'floatIn', 0, 400, 'reveal'), E('OK' + k, 'zoomIn', 300, 350), E('NO' + k, 'zoomIn', 380, 350)]);
      sl.on('OK' + k, close(k).concat([E('O' + k, 'zoomOut', 500, 450, 'correct'), E('D' + k, 'zoomIn', 850, 350),
        EW('pg' + k, segs[k].dir, 600, 500), E(prev, 'fadeOut', 1000, 300), E('R' + (k + 1), 'zoomIn', 1150, 450, 'step')]));
      sl.on('NO' + k, close(k).concat([E('O' + k, 'pulse', 500, 500, 'wrong')]));
    });
    // Về đích: bấm vào lá cờ 🏁 sau khi vượt hết chướng ngại vật
    sl.on('fin', [EW('pg' + n, segs[n].dir, 0, 500), E('R' + n, 'fadeOut', 400, 300), E('RF', 'zoomIn', 550, 450, 'step'),
      E('goal', 'zoomIn', 1100, 600, 'win'), E('goal', 'pulse', 1800, 600), E('goal', 'zoomOut', 4200, 500)]);
    sl.notes = `Vượt chướng ngại vật – ${n} chướng ngại.\n• Bấm vào chướng ngại vật theo thứ tự 1 → ${n} để hiện câu hỏi; bấm vào khung câu hỏi để hiện đáp án.\n• Bấm "Vượt qua" → nhân vật tiến lên; bấm "Chưa vượt qua" → giữ nguyên.\n• Khi vượt hết, bấm vào lá cờ 🏁 để nhân vật về đích.`;
    return sl;
  }

  /* ============ Nhóm trò nhiều câu trắc nghiệm trên 1 slide (giải cứu, hứng quả, bảo vệ Trái Đất) ============
   * Mỗi câu là một "nhóm" hình. Bấm ra ngoài (phím →) để chuyển sang câu tiếp theo.
   * Bấm vào phương án: đúng / sai sẽ chạy hiệu ứng riêng của từng trò. */
  const LET = 'ABCD';
  function roundQs(it) {
    const qs = (it.questions || []).slice(0, 8).map((q) => {
      const ans = (q.answers || []).map(nfc);
      let idx = [0, 1, 2, 3].filter((i) => (ans[i] || '').trim());
      if (idx.length < 2) idx = [0, 1];
      let c = parseInt(q.correct, 10) || 0;
      if (!idx.includes(c)) c = idx[0];
      return { q: nfc(q.question) || '(Chưa nhập câu hỏi)', opts: idx.map((i) => ({ i, t: ans[i] || '' })), correct: c };
    });
    if (!qs.length) qs.push({ q: '(Chưa nhập câu hỏi)', opts: [{ i: 0, t: '' }, { i: 1, t: '' }], correct: 0 });
    return qs;
  }
  /**
   * Dựng phần câu hỏi chung.
   * cfg.qBox: vị trí khung câu hỏi; cfg.place(k, m, j) -> {x,y,w,h} vị trí phương án thứ j (m phương án) của câu k
   * cfg.optStyle(k, j, opt) -> thuộc tính hình phương án; cfg.onRight(k, q, j, name) / cfg.onWrong(...) -> mảng hiệu ứng thêm
   * cfg.intro: hiệu ứng mở đầu (khung cảnh); cfg.finale: hiệu ứng cuối cùng
   */
  function quizRounds(sl, qs, cfg) {
    const groups = [];
    qs.forEach((q, k) => {
      const g = [];
      const qb = cfg.qBox;
      const qt = `Câu ${k + 1}/${qs.length}: ${q.q}`;
      sl.add(Object.assign({ name: `Q${k}`, shape: 'roundRect', radius: 0.18, text: qt, size: fit(qt, qb.w - 0.4, qb.h - 0.15, 26, 12), bold: true,
        color: '1F2937', fill: 'FFFFFF', line: cfg.accent, lineW: 2.5, shadow: true, align: 'center', pad: [14, 14, 4, 4] }, qb));
      g.push(`Q${k}`);
      q.opts.forEach((o, j) => {
        const name = `o${k}_${j}`;
        const p = cfg.place(k, q.opts.length, j);
        const label = `${LET[o.i]}. ${o.t}`;
        sl.add(Object.assign({ name, text: label, size: fit(label, p.w - 0.3, p.h - 0.2, 22, 10), bold: true, color: 'FFFFFF', shadow: true }, p, cfg.optStyle(k, j, o)));
        g.push(name);
      });
      groups.push(g);
    });
    (cfg.extraShapes || (() => {}))();
    const intro = (cfg.intro || []).concat(groups[0].map((n, j) => E(n, j ? 'zoomIn' : 'floatIn', 700 + j * 120, 400, j ? null : 'reveal')));
    sl.auto(intro);
    for (let k = 1; k < qs.length; k++) {
      sl.click(groups[k - 1].map((n) => E(n, 'fadeOut', 0, 300)).concat(cfg.clearShared ? cfg.clearShared(k - 1) : [],
        groups[k].map((n, j) => E(n, j ? 'zoomIn' : 'floatIn', 350 + j * 110, 400, j ? null : 'reveal'))));
    }
    if (cfg.finale) sl.click(groups[qs.length - 1].map((n) => E(n, 'fadeOut', 0, 300)).concat(cfg.clearShared ? cfg.clearShared(qs.length - 1) : [], cfg.finale));
    qs.forEach((q, k) => q.opts.forEach((o, j) => {
      const name = `o${k}_${j}`;
      sl.on(name, o.i === q.correct ? cfg.onRight(k, q, j, name) : cfg.onWrong(k, q, j, name));
    }));
    // Bỏ các hình phụ dùng chung (dấu ✔/✘, tia laser, quả thối…) không có hiệu ứng nào tác động,
    // nếu không chúng sẽ hiện sẵn ngay từ đầu.
    if (cfg.lazy) {
      const used = new Set();
      sl.main.forEach((g) => g.effects.forEach((e) => used.add(e.target)));
      sl.triggers.forEach((t) => t.effects.forEach((e) => used.add(e.target)));
      sl.shapes = sl.shapes.filter((s) => !cfg.lazy.test(s.name) || used.has(s.name));
    }
  }
  const optRow = (x0, w0, y, h, gap) => (k, m, j) => { const w = (w0 - (m - 1) * gap) / m; return { x: x0 + j * (w + gap), y, w, h }; };

  /* ---------------- 7. Giải cứu thú cưng ---------------- */
  function buildRescue(it, ctx) {
    const sl = new Slide({ key: 'rescue', label: 'Giải cứu thú cưng', bg: '0C4A6E', transition: ctx.transition, itemId: it.id });
    deco(sl, ['38BDF8', 'FDE047', '22D3EE']);
    const AC = '38BDF8';
    header(sl, 'GIẢI CỨU THÚ CƯNG', AC, '0C4A6E', ctx, it.timer);
    const qs = roundQs(it), n = qs.length;
    const pet = it.pet === 'cat' ? '🐱' : '🐶';
    // Căn phòng kín
    const RX = 0.5, RY = 1.1, RW = 5.4, RH = 5.95, FLOOR = RY + RH - 0.75;
    sl.add({ name: 'room', shape: 'roundRect', radius: 0.15, x: RX, y: RY, w: RW, h: RH, fill: '475569', line: '94A3B8', lineW: 4, shadow: true });
    sl.add({ name: 'wnd', shape: 'roundRect', radius: 0.08, x: RX + 0.4, y: RY + 0.9, w: 1.1, h: 0.9, fill: '1E293B', line: 'CBD5E1', lineW: 2, text: '🌙', size: 20 });
    sl.add({ name: 'pipe', shape: 'rect', x: RX + 2.4, y: RY, w: 0.6, h: 0.35, fill: '64748B', line: '334155', lineW: 1 });
    sl.add({ name: 'flow', shape: 'rect', x: RX + 2.55, y: RY + 0.35, w: 0.3, h: FLOOR - RY - 0.35, fill: '38BDF8', alpha: 55 });
    sl.add({ name: 'floor', shape: 'rect', x: RX + 0.1, y: FLOOR, w: RW - 0.2, h: 0.62, fill: '334155' });
    // Nước dâng: mỗi câu trả lời sai dâng thêm 1 lớp
    const WH = (FLOOR - RY - 1.0) / n;
    for (let k = 0; k < n; k++) sl.add({ name: 'W' + k, shape: 'rect', x: RX + 0.1, y: FLOOR - (k + 1) * WH, w: RW - 0.2, h: WH, fill: '0EA5E9', alpha: 45 });
    // lớp "phòng khô" phủ lên nước khi giải cứu xong (nước rút hết)
    sl.add({ name: 'dry', shape: 'rect', x: RX + 0.1, y: FLOOR - n * WH, w: RW - 0.2, h: n * WH, fill: '475569' });
    sl.add({ name: 'pet', shape: 'ellipse', x: RX + RW / 2 - 0.8, y: FLOOR - 1.45, w: 1.6, h: 1.4, fill: 'FDE68A', line: 'F59E0B', lineW: 3, text: pet, size: 50, shadow: true });
    sl.add({ name: 'sad', kind: 'text', x: RX + RW / 2 + 0.6, y: FLOOR - 1.9, w: 0.8, h: 0.6, text: '💦', size: 24 });
    // Van thoát nước dưới sàn: mỗi câu đúng mở 1 van
    const VW = Math.min(0.5, (RW - 0.6) / n - 0.08);
    for (let k = 0; k < n; k++) {
      const x = RX + 0.3 + k * ((RW - 0.6) / n) + ((RW - 0.6) / n - VW) / 2;
      sl.add({ name: 'V' + k, shape: 'ellipse', x, y: FLOOR + 0.1, w: VW, h: VW * 0.8, fill: '1F2937', line: '94A3B8', lineW: 1.5, text: '⛔', size: Math.round(VW * 30) });
      sl.add({ name: 'VO' + k, shape: 'ellipse', x, y: FLOOR + 0.1, w: VW, h: VW * 0.8, fill: '22C55E', line: 'BBF7D0', lineW: 1.5, text: '✔', size: Math.round(VW * 30), bold: true, color: 'FFFFFF' });
    }
    const OC = ['EF4444', '3B82F6', '16A34A', 'D97706'];
    quizRounds(sl, qs, {
      accent: AC,
      qBox: { x: 6.25, y: 1.1, w: 6.6, h: 2.0 },
      place: (k, m, j) => ({ x: 6.25, y: 3.3 + j * 0.925, w: 6.6, h: 0.78 }),
      optStyle: (k, j, o) => ({ shape: 'roundRect', radius: 0.18, fill: OC[o.i], line: 'FFFFFF', lineW: 2, align: 'left', pad: [18, 10, 2, 2] }),
      extraShapes: () => {
        for (let j = 0; j < 4; j++) {
          const p = { x: 12.15, y: 3.3 + j * (3.7 / 4) + 0.12, w: 0.6, h: 0.6 };
          sl.add(Object.assign({ name: 'ok' + j, shape: 'ellipse', fill: '16A34A', line: 'FFFFFF', lineW: 2, text: '✔', size: 20, bold: true }, p));
          sl.add(Object.assign({ name: 'no' + j, shape: 'ellipse', fill: 'DC2626', line: 'FFFFFF', lineW: 2, text: '✘', size: 20, bold: true }, p));
        }
        sl.add({ name: 'goal', shape: 'roundRect', radius: 0.3, x: 3.17, y: 2.6, w: 7.0, h: 1.6, fill: 'FDE047', line: 'FFFFFF', lineW: 4, text: `${pet}  GIẢI CỨU THÀNH CÔNG!`, size: fit('🐶  GIẢI CỨU THÀNH CÔNG!', 6.4, 0.75, 36, 18, 0.72), bold: true, face: FB, color: '0C4A6E', shadow: true, rot: -3 });
      },
      lazy: /^(ok|no)\d$/,
      // dấu ✔/✘ cạnh phương án, tự tắt sau 1,6 giây
      intro: [E('room', 'zoomIn', 0, 600, 'reveal'), E('pet', 'zoomIn', 400, 500), EW('flow', 'up', 300, 700)],
      onRight: (k, q, j) => [E(rowBadge('ok', q, j), 'zoomIn', 0, 350, 'correct'), E(rowBadge('ok', q, j), 'zoomOut', 1600, 300), E('VO' + k, 'zoomIn', 300, 450, 'whoosh'), E('pet', 'pulse', 600, 500)],
      onWrong: (k, q, j) => [E(rowBadge('no', q, j), 'zoomIn', 0, 350, 'wrong'), E(rowBadge('no', q, j), 'zoomOut', 1600, 300), EW('W' + k, 'down', 250, 900),
        E('sad', 'zoomIn', 300, 300), E('sad', 'zoomOut', 1400, 300), E('pet', 'pulse', 500, 400)],
      finale: [E('flow', 'fadeOut', 0, 400), EW('dry', 'up', 200, 1200, 'whoosh'), E('goal', 'zoomIn', 1300, 600, 'win'), E('pet', 'pulse', 2000, 600)]
    });
    sl.notes = `Giải cứu thú cưng – ${n} câu.\n• Bấm vào phương án học sinh chọn: đúng → mở 1 van thoát nước; sai → nước dâng thêm 1 mức.\n• Bấm ra ngoài (phím →) để sang câu tiếp theo; sau câu cuối, bấm thêm 1 lần để giải cứu thú cưng.`;
    return sl;
  }
  // vị trí dấu ✔/✘ theo hàng: khi có m phương án, hàng j dùng huy hiệu j (đặt cạnh khung phương án)
  function rowBadge(kind, q, j) { return kind + j; }

  /* ---------------- 8. Hứng hoa quả ---------------- */
  const FRUITS = ['🍎', '🍊', '🍋', '🍉', '🍇', '🍓', '🍑', '🍐'];
  function buildFruit(it, ctx) {
    const sl = new Slide({ key: 'fruit', label: 'Hứng hoa quả', bg: '0F3D2E', transition: ctx.transition, itemId: it.id });
    deco(sl, ['22C55E', 'FACC15', 'F97316']);
    const AC = 'FACC15';
    header(sl, 'HỨNG HOA QUẢ', AC, '0F3D2E', ctx, it.timer);
    const qs = roundQs(it), n = qs.length;
    // cây
    sl.add({ name: 'trunk', shape: 'rect', x: 6.0, y: 3.9, w: 1.33, h: 2.6, fill: '7C4A1E', line: '5B3412', lineW: 2 });
    sl.add({ name: 'crown1', shape: 'ellipse', x: 0.4, y: 2.05, w: 12.53, h: 2.45, fill: '15803D', shadow: true });
    sl.add({ name: 'crown2', shape: 'ellipse', x: 1.3, y: 1.85, w: 10.7, h: 1.4, fill: '16A34A' });
    sl.add({ name: 'ground', shape: 'rect', x: 0, y: 6.5, w: W, h: 1.0, fill: '3F6212' });
    const BX = 4.9, BY = 5.15;
    sl.add({ name: 'animal', shape: 'ellipse', x: BX + 0.55, y: BY + 0.55, w: 1.45, h: 1.3, fill: 'FDE68A', line: 'B45309', lineW: 2.5, text: it.pet || '🐻', size: 44, shadow: true });
    sl.add({ name: 'basket', shape: 'roundRect', radius: 0.15, x: BX, y: BY, w: 2.55, h: 0.7, fill: 'B45309', line: '78350F', lineW: 2.5, text: '🧺 GIỎ', size: 16, bold: true, color: 'FFFFFF' });
    sl.add({ name: 'bin', shape: 'roundRect', radius: 0.12, x: 11.3, y: 5.55, w: 1.4, h: 1.2, fill: '475569', line: '1F2937', lineW: 2, text: '🗑️', size: 34 });
    sl.add({ name: 'scl', kind: 'text', x: 0.4, y: 6.55, w: 1.4, h: 0.5, text: 'Đã hứng:', size: 14, bold: true, color: 'FDE047', align: 'left' });
    for (let k = 0; k < n; k++) sl.add({ name: 'S' + k, shape: 'ellipse', x: 1.75 + k * 0.62, y: 6.55, w: 0.52, h: 0.52, fill: 'FFFFFF', alpha: 15, text: FRUITS[k % FRUITS.length], size: 18 });
    for (let k = 0; k < n; k++) sl.add({ name: 'SE' + k, shape: 'ellipse', x: 1.75 + k * 0.62, y: 6.55, w: 0.52, h: 0.52, fill: '000000', alpha: 65, line: 'FFFFFF', lineW: 1, dash: true });
    quizRounds(sl, qs, {
      accent: AC,
      qBox: { x: 0.5, y: 1.0, w: 12.33, h: 0.95 },
      place: (k, m, j) => fruitPlace(m === 2 ? j + 1 : j),
      optStyle: (k, j, o) => ({ shape: 'ellipse', fill: ['DC2626', 'EA580C', 'CA8A04', '7C3AED'][o.i], line: 'FEF3C7', lineW: 3, runs: [{ t: FRUITS[(k * 4 + j) % FRUITS.length] + ' ', size: 22 }, { t: `${LET[o.i]}. ${o.t}`, bold: true }], pad: [8, 8, 2, 2] }),
      extraShapes: () => {
        for (let j = 0; j < 4; j++) {
          // quả thối tại vị trí phương án j (bố cục 4 quả); khi ít quả hơn dùng vị trí tương ứng
          sl.add({ name: 'rot' + j, shape: 'ellipse', x: 0, y: 2.4, w: 1.55, h: 1.55, fill: '57534E', line: '292524', lineW: 2, text: '🤢', size: 40 });
        }
        sl.add({ name: 'catch', shape: 'ellipse', x: BX + 0.8, y: BY - 0.85, w: 0.95, h: 0.95, fill: 'FDE047', line: 'FFFFFF', lineW: 2, text: '🍎', size: 30, shadow: true });
        sl.add({ name: 'trash', shape: 'ellipse', x: 11.6, y: 4.75, w: 0.8, h: 0.8, fill: '57534E', line: '292524', lineW: 2, text: '🤢', size: 24 });
        sl.add({ name: 'goal', shape: 'roundRect', radius: 0.3, x: 3.17, y: 2.55, w: 7.0, h: 1.5, fill: 'FDE047', line: 'FFFFFF', lineW: 4, text: '🧺  THU HOẠCH THÀNH CÔNG!', size: fit('🧺  THU HOẠCH THÀNH CÔNG!', 6.4, 0.75, 36, 18, 0.72), bold: true, face: FB, color: '14532D', shadow: true, rot: -3 });
      },
      lazy: /^rot\d$/,
      intro: [E('crown1', 'zoomIn', 0, 600, 'reveal'), E('animal', 'floatIn', 400, 500)],
      onRight: (k, q, j, name) => [E(name, 'zoomOut', 0, 350, 'whoosh'), E('catch', 'dropIn', 300, 500, 'pop'), E('catch', 'zoomOut', 1300, 300), E('SE' + k, 'fadeOut', 700, 300, 'correct'), E('S' + k, 'pulse', 1000, 400), E('animal', 'pulse', 800, 400)],
      onWrong: (k, q, j, name) => [E(rotAt(q, j), 'zoomIn', 0, 350, 'wrong'), E(name, 'fadeOut', 100, 300), E(rotAt(q, j), 'zoomOut', 1000, 300), E('trash', 'dropIn', 1250, 450), E('trash', 'zoomOut', 2200, 300), E('bin', 'pulse', 1500, 400)],
      clearShared: () => [],
      finale: [E('goal', 'zoomIn', 300, 600, 'win'), E('animal', 'pulse', 1000, 600)]
    });
    // đặt quả thối đúng giữa ô phương án tương ứng
    sl.shapes.filter((s) => /^rot\d$/.test(s.name)).forEach((s) => { const p = fruitPlace(+s.name.slice(3)); s.x = p.x + (p.w - 1.55) / 2; });
    sl.notes = `Hứng hoa quả – ${n} câu.\n• Bấm vào quả mang đáp án học sinh chọn: đúng → quả rơi vào giỏ, thêm 1 quả vào hàng "Đã hứng"; sai → quả thối và bị vứt vào thùng rác.\n• Bấm ra ngoài (phím →) để sang câu tiếp theo; sau câu cuối bấm thêm 1 lần để tổng kết.`;
    return sl;
  }
  function rotAt(q, j) { return 'rot' + slot(q, j); }
  const fruitPlace = (j) => { const m = 4, w = 2.85, x0 = (W - (m * w + (m - 1) * 0.3)) / 2; return { x: x0 + j * (w + 0.3), y: 2.4, w, h: 1.55 }; };

  /* ---------------- 9. Bảo vệ Trái Đất ---------------- */
  function buildSpace(it, ctx) {
    const sl = new Slide({ key: 'space', label: 'Bảo vệ Trái Đất', bg: '0B1026', transition: ctx.transition, itemId: it.id });
    for (let i = 0; i < 26; i++) {
      const s = 0.04 + (i % 3) * 0.03;
      sl.add({ name: 'st' + i, shape: 'ellipse', x: (i * 3.71) % 13.1 + 0.1, y: (i * 2.39) % 7.2 + 0.15, w: s, h: s, fill: 'FFFFFF', alpha: 20 + (i % 4) * 15 });
    }
    const AC = 'A78BFA';
    header(sl, 'BẢO VỆ TRÁI ĐẤT', AC, '0B1026', ctx, it.timer);
    const qs = roundQs(it), n = qs.length;
    sl.add({ name: 'earth', shape: 'ellipse', x: 4.42, y: 5.45, w: 4.5, h: 4.5, fill: '1D4ED8', line: '60A5FA', lineW: 6, text: '🌍', size: 96, valign: 'top', pad: [0, 0, 0, 6], shadow: true });
    sl.add({ name: 'shield', shape: 'ellipse', x: 3.95, y: 5.0, w: 5.45, h: 5.45, line: '38BDF8', lineW: 3, dash: true });
    sl.add({ name: 'ship', kind: 'text', x: 6.17, y: 4.35, w: 1.0, h: 0.9, text: '🚀', size: 40 });
    const P4 = (j) => { const w = 2.85, x0 = (W - (4 * w + 3 * 0.3)) / 2; return { x: x0 + j * (w + 0.3), y: 2.25, w, h: 1.5 }; };
    const place = (k, m, j) => P4(m === 2 ? j + 1 : j);
    quizRounds(sl, qs, {
      accent: AC,
      qBox: { x: 0.5, y: 1.0, w: 12.33, h: 1.0 },
      place,
      optStyle: () => ({ shape: 'ellipse', fill: '78716C', line: 'D6D3D1', lineW: 3 }),
      extraShapes: () => {
        for (let j = 0; j < 4; j++) {
          const p = P4(j), cx = p.x + p.w / 2;
          sl.add({ name: 'beam' + j, shape: 'rect', x: cx - 0.06, y: p.y + p.h, w: 0.12, h: 4.4 - (p.y + p.h) + 0.6, fill: 'F43F5E', shadow: true });
          sl.add({ name: 'boom' + j, shape: 'star5', x: cx - 0.85, y: p.y - 0.1, w: 1.7, h: 1.7, fill: 'FDE047', line: 'F97316', lineW: 3 });
          sl.add({ name: 'alien' + j, shape: 'ellipse', x: cx - 0.55, y: p.y + 0.2, w: 1.1, h: 1.1, fill: '4ADE80', line: '14532D', lineW: 2, text: '👾', size: 32 });
        }
        for (let k = 0; k < n; k++) sl.add({ name: 'inv' + k, kind: 'text', x: 4.6 + (k % 4) * 1.0, y: 5.6 + Math.floor(k / 4) * 0.55, w: 0.9, h: 0.6, text: '👾', size: 26 });
        sl.add({ name: 'hit', shape: 'ellipse', x: 4.42, y: 5.45, w: 4.5, h: 4.5, fill: 'DC2626', alpha: 45 });
        sl.add({ name: 'goal', shape: 'roundRect', radius: 0.3, x: 3.17, y: 2.5, w: 7.0, h: 1.5, fill: 'FDE047', line: 'FFFFFF', lineW: 4, text: '🌍  TRÁI ĐẤT AN TOÀN!', size: fit('🌍  TRÁI ĐẤT AN TOÀN!', 6.4, 0.75, 36, 18, 0.72), bold: true, face: FB, color: '1E1B4B', shadow: true, rot: -3 });
      },
      lazy: /^(beam|boom|alien)\d$/,
      intro: [E('earth', 'floatIn', 0, 700, 'reveal'), E('shield', 'zoomIn', 400, 500), E('ship', 'zoomIn', 600, 400)],
      onRight: (k, q, j, name) => {
        const s = slot(q, j);
        return [E('ship', 'pulse', 0, 250), EW('beam' + s, 'down', 0, 250, 'whoosh'), E('beam' + s, 'fadeOut', 450, 250), E('boom' + s, 'zoomIn', 250, 250),
          E(name, 'zoomOut', 300, 300), E('alien' + s, 'zoomIn', 450, 300), E('boom' + s, 'zoomOut', 650, 250), E('alien' + s, 'zoomOut', 1000, 350, 'correct')];
      },
      onWrong: (k, q, j, name) => {
        const s = slot(q, j), c = slot(q, q.opts.findIndex((o) => o.i === q.correct));
        return [E('ship', 'pulse', 0, 250), EW('beam' + s, 'down', 0, 250, 'whoosh'), E('beam' + s, 'fadeOut', 450, 250), E('boom' + s, 'zoomIn', 250, 250),
          E(name, 'zoomOut', 300, 300), E('boom' + s, 'zoomOut', 650, 250), E('alien' + c, 'zoomIn', 800, 300, 'wrong'), E('alien' + c, 'zoomOut', 1400, 300),
          E('inv' + k, 'dropIn', 1450, 500), E('hit', 'fadeIn', 1750, 200), E('hit', 'fadeOut', 2050, 400), E('earth', 'pulse', 1750, 500)];
      },
      clearShared: () => [],
      finale: [E('shield', 'pulse', 0, 600), E('goal', 'zoomIn', 300, 600, 'win'), E('earth', 'pulse', 1000, 600)]
    });
    // Với câu ít phương án, hình chùm tia/nổ dùng vị trí theo bố cục 4 phương án gần nhất
    sl.notes = `Bảo vệ Trái Đất – ${n} câu.\n• Bấm vào thiên thạch mang đáp án học sinh chọn: tàu bắn tia laser; đúng → thiên thạch vỡ, sinh vật ngoài hành tinh bị tiêu diệt; sai → sinh vật trong thiên thạch đúng bay xuống tấn công Trái Đất.\n• Bấm ra ngoài (phím →) để sang câu tiếp theo; sau câu cuối bấm thêm 1 lần để tổng kết.`;
    return sl;
  }
  // Phương án luôn nằm ở 1 trong 4 ô cố định: câu 2 phương án (Đúng/Sai) dùng 2 ô giữa
  function slot(q, j) { return q.opts.length === 2 ? j + 1 : j; }

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
    sl.auto([E('big', 'zoomIn', 0, 700, 'win'), E('big', 'pulse', 800, 600), E('msg', 'floatIn', 500, 600), E('back', 'fadeIn', 1100, 500)].concat(fx));
    return sl;
  }

  const BUILDERS = { millionaire: buildMillionaire, picture: buildPicture, puzzle: buildPuzzle, truefalse: buildTrueFalse, tiles: buildTiles, obstacle: buildObstacle,
    rescue: buildRescue, fruit: buildFruit, space: buildSpace };

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
      case 'rescue': case 'fruit': case 'space':
        return Object.assign(base, { pet: type === 'rescue' ? 'dog' : type === 'fruit' ? '🐻' : undefined,
          questions: [0, 1, 2].map(() => ({ question: '', answers: ['', '', '', ''], correct: 0 })) });
      case 'obstacle': return Object.assign(base, { title: 'Vượt qua tất cả chướng ngại vật để về đích!', steps: [0, 1, 2, 3].map(() => ({ question: '', answer: '' })) });
    }
    return base;
  }

  return { W, H, GAME_TYPES, buildDeck, newItem, fit, layoutLetters, _builders: BUILDERS };
});
