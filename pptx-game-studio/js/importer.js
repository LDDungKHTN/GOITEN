/*!
 * PPTX Game Studio – Nhập câu hỏi từ file Word (.docx), ảnh (OCR) hoặc văn bản dán vào,
 * nhận dạng câu hỏi và chuyển thành các trò chơi.
 *
 * Định dạng nhận dạng được (linh hoạt):
 *   Câu 1: Nội dung câu hỏi?          (hoặc "1." / "1)" / "Bài 1:")
 *   A. ...   B. ...   C. ...   D. ... (mỗi phương án 1 dòng hoặc chung 1 dòng)
 *   Đáp án: B                         (hoặc đánh dấu * trước phương án đúng,
 *                                      hoặc trong Word: gạch chân / tô màu / tô đậm phương án đúng)
 *   Giải thích: ...
 *   Câu đúng/sai: "Đáp án: Đúng" / "Đáp án: Sai";  câu tự luận: "Đáp án: <nội dung>"
 *   Bảng đáp án cuối bài: "ĐÁP ÁN  1-B  2-C  3.A ..."
 */
(function (root) {
  'use strict';
  const L = 'ABCD';
  const nfc = (s) => String(s == null ? '' : s).normalize('NFC');

  /* ================= Đọc file Word (.docx) ================= */
  const unxml = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

  async function readDocx(file, JSZip) {
    const zip = await JSZip.loadAsync(file);
    const docFile = zip.file('word/document.xml');
    if (!docFile) throw new Error('File không phải định dạng Word .docx hợp lệ');
    const xml = await docFile.async('string');
    const relsXml = zip.file('word/_rels/document.xml.rels') ? await zip.file('word/_rels/document.xml.rels').async('string') : '';
    const rels = {};
    relsXml.replace(/<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g, (m, id, t) => { rels[id] = t; return m; });
    relsXml.replace(/<Relationship [^>]*Target="([^"]+)"[^>]*Id="([^"]+)"/g, (m, t, id) => { rels[id] = rels[id] || t; return m; });
    const images = [];
    const lines = [];
    const paras = xml.match(/<w:p[ >][\s\S]*?<\/w:p>|<w:p\/>/g) || [];
    for (const p of paras) {
      let text = '';
      const flags = [];
      const runs = p.match(/<w:r[ >][\s\S]*?<\/w:r>/g) || [];
      for (const r of runs) {
        const rpr = (r.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0];
        const b = /<w:b\/>|<w:b w:val="(1|true|on)"\/>/.test(rpr);
        const u = /<w:u w:val="(?!none)[^"]+"/.test(rpr);
        const col = (rpr.match(/<w:color w:val="([0-9A-Fa-f]{6})"/) || [])[1];
        const colored = !!col && !/^(000000|auto)$/i.test(col) && !isDarkGray(col);
        const hl = /<w:highlight w:val="(?!none)[^"]+"/.test(rpr) || /<w:shd [^>]*w:fill="(?!auto|FFFFFF)[0-9A-Fa-f]{6}"/.test(rpr);
        const f = { b, emph: u || colored || hl };
        r.replace(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<w:br\/>/g, (m, t) => {
          const s = t != null ? unxml(t) : m === '<w:tab/>' ? '\t' : '\n';
          text += s; for (let i = 0; i < s.length; i++) flags.push(f);
          return m;
        });
        const blip = r.match(/<a:blip [^>]*r:embed="([^"]+)"/);
        if (blip && rels[blip[1]]) {
          const target = rels[blip[1]].replace(/^\//, '');
          const zf = zip.file(target.startsWith('word/') ? target : 'word/' + target);
          if (zf && /\.(png|jpe?g|gif|bmp)$/i.test(target)) {
            const ext = target.split('.').pop().toLowerCase();
            const b64 = await zf.async('base64');
            images.push(`data:image/${ext === 'jpg' ? 'jpeg' : ext};base64,${b64}`);
            lines.push({ text: `[Ảnh ${images.length}]`, flags: [] });
          }
        }
      }
      text.split('\n').forEach((seg, i, arr) => {
        const off = arr.slice(0, i).reduce((s, x) => s + x.length + 1, 0);
        lines.push({ text: seg, flags: flags.slice(off, off + seg.length) });
      });
    }
    // Phương án mỗi dòng 1 cái (A., B., C., D. trên các dòng liên tiếp): so định dạng giữa các dòng
    const out = lines.map(markOptions);
    const single = /^\s*\(?([A-Da-d])\s*[.):]\s*\S/;
    for (let i = 0; i < lines.length; i++) {
      const m = lines[i].text.match(single);
      if (!m || m[1].toUpperCase() !== 'A') continue;
      const grp = [i];
      for (let j = i + 1; j < lines.length && grp.length < 4; j++) {
        const mj = lines[j].text.match(single);
        if (!mj || mj[1].toUpperCase() !== L[grp.length]) break;
        grp.push(j);
      }
      if (grp.length < 2 || grp.some((k) => /^\*/.test(out[k]) || optionRanges(lines[k].text.trim()).length > 1)) continue;
      const ratio = (k, key) => {
        const t = lines[k].text, f = lines[k].flags; let n = 0, c = 0;
        const start = (t.match(single)[0].length) - 1;
        for (let x = start; x < t.length; x++) { if (!t[x].trim()) continue; n++; if (f[x] && f[x][key]) c++; }
        return n ? c / n : 0;
      };
      let marked = grp.filter((k) => ratio(k, 'emph') > 0.6);
      if (marked.length !== 1) { const b = grp.filter((k) => ratio(k, 'b') > 0.6); marked = b.length === 1 ? b : []; }
      if (marked.length === 1) out[marked[0]] = '*' + out[marked[0]];
      i = grp[grp.length - 1];
    }
    return { text: out.join('\n').replace(/\n{3,}/g, '\n\n').trim(), images };
  }
  function isDarkGray(hex) { const n = parseInt(hex, 16); return ((n >> 16) & 255) < 70 && ((n >> 8) & 255) < 70 && (n & 255) < 70; }

  // Trong Word: phương án được gạch chân / tô màu / tô đậm (khác các phương án khác) => thêm dấu *
  const OPT_RE = /(^|[\s\t])\(?([A-Da-d])\s*[.):]\s*/g;
  function optionRanges(text) {
    const res = []; let m;
    OPT_RE.lastIndex = 0;
    const found = [];
    while ((m = OPT_RE.exec(text))) found.push({ label: m[2].toUpperCase(), start: m.index + m[1].length, textStart: m.index + m[0].length });
    // phải bắt đầu bằng A và theo thứ tự
    if (!found.length || found[0].start > 1 || found[0].label !== 'A') return [];
    let want = 0;
    for (const f of found) { if (f.label === L[want]) { res.push(f); want++; } }
    res.forEach((r, i) => { r.end = i + 1 < res.length ? res[i + 1].start : text.length; });
    return res;
  }
  function markOptions(line) {
    const lw = line.text.length - line.text.replace(/^\s+/, '').length;
    const text = line.text.slice(lw);
    const flags = line.flags.slice(lw);
    const ops = optionRanges(text);
    if (ops.length < 2 || !flags.length) return text;
    line = { flags };
    const ratio = (o, key) => {
      let n = 0, k = 0;
      for (let i = o.textStart; i < o.end; i++) { if (!text[i] || !text[i].trim()) continue; n++; if (line.flags[i] && line.flags[i][key]) k++; }
      return n ? k / n : 0;
    };
    let marked = ops.filter((o) => ratio(o, 'emph') > 0.6);
    if (marked.length !== 1) { const bold = ops.filter((o) => ratio(o, 'b') > 0.6); marked = bold.length === 1 ? bold : []; }
    if (marked.length !== 1) return text;
    const at = marked[0].start;
    return text.slice(0, at) + '*' + text.slice(at);
  }

  /* ================= OCR ảnh (Tesseract.js, tải khi cần) ================= */
  let tessP = null;
  function loadTesseract() {
    if (root.Tesseract) return Promise.resolve(root.Tesseract);
    if (!tessP) tessP = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
      s.onload = () => res(root.Tesseract); s.onerror = () => { tessP = null; rej(new Error('Không tải được bộ nhận dạng chữ (cần Internet)')); };
      document.head.appendChild(s);
    });
    return tessP;
  }
  async function ocrImage(file, onProgress) {
    const T = await loadTesseract();
    const worker = await T.createWorker('vie', 1, {
      logger: (m) => { if (onProgress && m.status) onProgress(m.status, m.progress || 0); }
    });
    try {
      const src = await prepForOcr(file);
      const { data } = await worker.recognize(src);
      return nfc(data.text).replace(/[ \t]+\n/g, '\n').trim();
    } finally { await worker.terminate(); }
  }
  // phóng to ảnh nhỏ, chuyển xám để nhận dạng tốt hơn
  function prepForOcr(file) {
    return new Promise((res) => {
      const url = URL.createObjectURL(file);
      const img = new Image();
      img.onload = () => {
        const k = Math.max(1, Math.min(3, 2000 / Math.max(img.naturalWidth, img.naturalHeight)));
        const c = document.createElement('canvas'); c.width = img.naturalWidth * k; c.height = img.naturalHeight * k;
        const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0, c.width, c.height);
        const d = g.getImageData(0, 0, c.width, c.height), a = d.data;
        for (let i = 0; i < a.length; i += 4) { const v = a[i] * 0.299 + a[i + 1] * 0.587 + a[i + 2] * 0.114; a[i] = a[i + 1] = a[i + 2] = v; }
        g.putImageData(d, 0, 0);
        URL.revokeObjectURL(url); res(c);
      };
      img.onerror = () => { URL.revokeObjectURL(url); res(file); };
      img.src = url;
    });
  }

  /* ================= Phân tích văn bản thành câu hỏi ================= */
  const Q_RE = /^(?:câu|cau|bài|bai|question|q)\s*(\d+)\s*[:.)\-–]?\s*(.*)$/i;
  const NUM_RE = /^(\d{1,3})\s*[.)]\s+(.+)$/;
  const ANS_RE = /^(?:\*\s*)?(?:đáp\s*án(?:\s*đúng)?|đa|đ\/a|answer|key|chọn)\s*[:.\-–=]?\s*(.*)$/i;
  const EXP_RE = /^(?:giải\s*thích|lời\s*giải|hướng\s*dẫn|gợi\s*ý|chú\s*thích)\s*[:.\-–]?\s*(.*)$/i;
  const KEY_HEAD = /^(?:bảng\s*)?đáp\s*án\s*[:.]?\s*$|^(?:answer\s*key|keys?)\s*[:.]?\s*$/i;
  const TF_RE = /^(đúng|sai|đ|s|true|false|t|f)\b\.?$/i;
  const isTrueWord = (s) => /^(đúng|đ|true|t)$/i.test(s.trim().replace(/\.$/, ''));

  function parseText(text, images) {
    images = images || [];
    const lines = nfc(text).split(/\r?\n/).map((s) => s.replace(/\s+$/, ''));
    const qs = [];
    let q = null, keyMode = false, pendingImg = null;
    const keys = {};
    const newQ = (num, body) => {
      q = { num, text: body || '', options: [], marked: -1, answer: '', explanation: '', image: pendingImg || '' };
      pendingImg = null; qs.push(q);
    };
    for (let raw of lines) {
      const t = raw.trim();
      if (!t) continue;
      const im = t.match(/^\[Ảnh (\d+)\]$/);
      if (im) { const d = images[+im[1] - 1]; if (d) { if (q && !q.options.length && !q.answer) q.image = q.image || d; else pendingImg = d; } continue; }
      if (KEY_HEAD.test(t)) { keyMode = true; continue; }
      if (keyMode) {
        let found = false;
        t.replace(/(\d{1,3})\s*[-.:)]\s*([A-Da-d]|đúng|sai|Đ|S)(?![\p{L}])/giu, (m, n, v) => { keys[+n] = v; found = true; return m; });
        if (found) continue;
        keyMode = false;
      }
      let m = t.match(Q_RE) || t.match(NUM_RE);
      if (m) { newQ(+m[1], m[2]); continue; }
      if (!q) { // chưa có "Câu 1": coi dòng có dấu ? là câu hỏi
        if (/\?\s*$/.test(t)) newQ(qs.length + 1, t);
        continue;
      }
      if ((m = t.match(ANS_RE))) { q.answer = m[1].trim(); continue; }
      if ((m = t.match(EXP_RE))) { q.explanation = (q.explanation ? q.explanation + ' ' : '') + m[1].trim(); continue; }
      // phương án: dòng bắt đầu bằng nhãn kế tiếp (A, rồi B…), có thể nhiều phương án trên 1 dòng
      const lead = t.match(/^\*?\s*\(?([A-Da-d])\s*[.):]\s*/);
      if (lead && L.indexOf(lead[1].toUpperCase()) === q.options.length && q.options.length < 4) {
        splitOptions(t, q.options.length).forEach((s) => { q.options[s.idx] = s.text; if (s.star) q.marked = s.idx; });
        continue;
      }
      if (q.options.length) q.options[q.options.length - 1] += ' ' + t;
      else q.text = (q.text ? q.text + ' ' : '') + t;
    }
    return qs.map((x) => finalize(x, keys)).filter((x) => x.text || x.options.length);
  }
  function splitOptions(t, startIdx) {
    const out = [];
    const re = /(^|\s)(\*?)\s*\(?([A-Da-d])\s*[.):]\s*/g;
    const hits = []; let m, want = startIdx || 0;
    while ((m = re.exec(t))) {
      const idx = L.indexOf(m[3].toUpperCase());
      if (idx !== want) continue; // chỉ nhận nhãn đúng thứ tự A → B → C → D
      hits.push({ idx, star: !!m[2], at: m.index + m[1].length, start: m.index + m[0].length });
      want++;
    }
    hits.forEach((h, i) => {
      let s = t.slice(h.start, i + 1 < hits.length ? hits[i + 1].at : t.length).trim();
      let star = h.star;
      if (/\*$/.test(s)) { star = true; s = s.replace(/\s*\*$/, ''); }
      out.push({ idx: h.idx, text: s, star });
    });
    return out;
  }
  function finalize(q, keys) {
    const r = { num: q.num, text: q.text.replace(/\s+/g, ' ').trim(), image: q.image, explanation: q.explanation.trim(), options: [], correct: -1, answer: '', kind: 'short' };
    const opts = q.options.filter((x) => x != null).map((s) => s.replace(/\s+/g, ' ').trim());
    let ans = (q.answer || '').trim();
    if (!ans && keys[q.num] != null) ans = String(keys[q.num]);
    if (opts.length >= 2) {
      r.kind = 'mcq'; r.options = opts.slice(0, 4);
      let c = q.marked;
      const lm = ans.match(/^\(?([A-Da-d])\b/);
      if (lm) c = L.indexOf(lm[1].toUpperCase());
      else if (ans) { const k = r.options.findIndex((o) => o.toLowerCase() === ans.toLowerCase()); if (k >= 0) c = k; }
      r.correct = c >= 0 && c < r.options.length ? c : 0;
      r.answerKnown = c >= 0;
      r.answer = r.options[r.correct] || '';
    } else if (TF_RE.test(ans) || /\((đúng|đ)\s*\/\s*(sai|s)\)\s*\??$/i.test(r.text)) {
      r.kind = 'tf'; r.isTrue = ans ? isTrueWord(ans) : true; r.answerKnown = !!ans;
      r.text = r.text.replace(/\s*\((đúng|đ)\s*\/\s*(sai|s)\)\s*\??$/i, '').trim();
    } else { r.kind = 'short'; r.answer = ans; r.answerKnown = !!ans; }
    return r;
  }

  /* ================= Chuyển câu hỏi thành trò chơi ================= */
  function chunk(arr, max, min) {
    if (!arr.length) return [];
    const n = Math.ceil(arr.length / max);
    const size = Math.max(min || 1, Math.ceil(arr.length / n));
    const out = [];
    for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
    return out;
  }
  function qa(q) {
    if (q.kind === 'mcq') return { question: q.text + '\n' + q.options.map((o, i) => `${L[i]}. ${o}`).join('   '), answer: `${L[q.correct]}. ${q.options[q.correct]}` };
    if (q.kind === 'tf') return { question: q.text + ' (Đúng hay Sai?)', answer: q.isTrue ? 'Đúng' : 'Sai' };
    return { question: q.text, answer: q.answer };
  }
  /**
   * @param qs  danh sách câu hỏi từ parseText
   * @param target 'auto' | 'millionaire' | 'truefalse' | 'tiles' | 'puzzle' | 'obstacle' | 'picture'
   */
  function toItems(qs, target, newItem) {
    const items = [], skipped = [];
    const mk = (type, data) => items.push(Object.assign(newItem(type), data));
    const group = (list, type) => {
      if (type === 'tiles') chunk(list, 8, 4).forEach((c, i) => mk('tiles', { title: `Lật ô số – phần ${i + 1}`, cells: c.map(qa) }));
      if (type === 'obstacle') chunk(list, 6, 3).forEach((c, i) => mk('obstacle', { title: `Vượt chướng ngại vật – chặng ${i + 1}`, steps: c.map(qa) }));
      if (type === 'puzzle') chunk(list, 9, 4).forEach((c) => {
        const pc = c.length <= 4 ? 4 : c.length <= 6 ? 6 : 9; // số mảnh vừa đủ chứa các câu
        const img = (c.find((q) => q.image) || {}).image || '';
        mk('puzzle', { pieceCount: pc, keyword: '', image: img, pieces: c.map(qa).concat(Array(pc - c.length).fill(0).map(() => ({ question: '', answer: '' }))) });
      });
    };
    const mill = (q) => mk('millionaire', { question: q.text, answers: [0, 1, 2, 3].map((i) => q.options[i] || ''), correct: q.correct, explanation: q.explanation });
    const tf = (q) => mk('truefalse', { statement: q.text, isTrue: q.isTrue !== false, explanation: q.explanation });
    const pic = (q) => mk('picture', { question: q.text || 'Hình ảnh này gợi nhắc đến từ khóa nào?', image: q.image, answer: q.answer || (q.kind === 'mcq' ? q.options[q.correct] : ''), hint: q.explanation });
    switch (target) {
      case 'millionaire': {
        const rest = [];
        qs.forEach((q) => (q.kind === 'mcq' ? mill(q) : q.kind === 'tf' ? tf(q) : rest.push(q)));
        group(rest, 'tiles'); break;
      }
      case 'truefalse':
        qs.forEach((q) => {
          if (q.kind === 'tf') tf(q);
          else if (q.kind === 'mcq') {
            const right = Math.random() < 0.5, wrongIdx = [0, 1, 2, 3].filter((i) => i !== q.correct && q.options[i]);
            const pickI = right || !wrongIdx.length ? q.correct : wrongIdx[Math.floor(Math.random() * wrongIdx.length)];
            mk('truefalse', { statement: `${q.text.replace(/[?:]\s*$/, '')}: ${q.options[pickI]}`, isTrue: pickI === q.correct, explanation: `Đáp án đúng: ${q.options[q.correct]}. ${q.explanation}`.trim() });
          } else if (q.answer) mk('truefalse', { statement: `${q.text.replace(/[?:]\s*$/, '')}: ${q.answer}`, isTrue: true, explanation: q.explanation });
          else skipped.push(q);
        });
        break;
      case 'picture':
        qs.forEach((q) => (q.image ? pic(q) : skipped.push(q))); break;
      case 'rescue': case 'fruit': case 'space': {
        // chỉ dùng câu trắc nghiệm và đúng/sai (đúng/sai thành 2 phương án)
        const usable = [];
        qs.forEach((q) => {
          if (q.kind === 'mcq') usable.push({ question: q.text, answers: [0, 1, 2, 3].map((i) => q.options[i] || ''), correct: q.correct });
          else if (q.kind === 'tf') usable.push({ question: q.text, answers: ['Đúng', 'Sai', '', ''], correct: q.isTrue === false ? 1 : 0 });
          else skipped.push(q);
        });
        const names = { rescue: 'Giải cứu thú cưng', fruit: 'Hứng hoa quả', space: 'Bảo vệ Trái Đất' };
        chunk(usable, 8, 1).forEach((c, i, all) => mk(target, { title: all.length > 1 ? `${names[target]} – lượt ${i + 1}` : names[target], questions: c }));
        break;
      }
      case 'tiles': case 'puzzle': case 'obstacle':
        group(qs, target); break;
      default: { // tự động
        const rest = [];
        qs.forEach((q) => {
          if (q.kind === 'mcq') mill(q);
          else if (q.kind === 'tf') tf(q);
          else if (q.image) pic(q);
          else rest.push(q);
        });
        group(rest, rest.length >= 3 && rest.length < 4 ? 'obstacle' : 'tiles');
      }
    }
    return { items, skipped };
  }

  root.GSImporter = { readDocx, ocrImage, parseText, toItems, loadTesseract };
})(typeof self !== 'undefined' ? self : this);
