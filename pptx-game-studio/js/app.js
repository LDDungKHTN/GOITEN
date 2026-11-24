/*! PPTX Game Studio – Giao diện soạn bài */
(function () {
  'use strict';
  const GS = window.GameScenes, B = window.GSBackend, C = window.GS_CONFIG || {};
  const $ = (s, r) => (r || document).querySelector(s);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const DRAFT_KEY = 'gs_draft_v1';
  const TIMERS = [0, 10, 15, 20, 30, 45, 60, 90, 120];
  const ASPECT = { picture: 5.9 / 5.45, puzzle: 7.6 / 5.7 };

  let project = null, selectedId = null, credits = null;

  /* ================= tiện ích ================= */
  function toast(msg, kind) {
    const t = $('#toast');
    t.textContent = msg; t.className = 'toast ' + (kind || ''); t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(() => { t.hidden = true; }, kind === 'bad' ? 5200 : 3200);
  }
  function busy(on, text) { $('#mBusy').hidden = !on; if (text) $('#busyText').textContent = text; }
  function confirmDlg(title, html, yes, no) {
    return new Promise((resolve) => {
      $('#cfTitle').textContent = title; $('#cfBody').innerHTML = html;
      $('#cfYes').textContent = yes || 'Đồng ý'; $('#cfNo').textContent = no || 'Quay lại';
      $('#cfNo').hidden = no === false;
      const m = $('#mConfirm'); m.hidden = false;
      const done = (v) => { m.hidden = true; $('#cfYes').onclick = $('#cfNo').onclick = null; resolve(v); };
      $('#cfYes').onclick = () => done(true); $('#cfNo').onclick = () => done(false);
      $('#cfYes').focus();
    });
  }
  function download(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 4000);
  }
  function slug(s) {
    return (String(s || 'tro-choi').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D')
      .replace(/[^a-zA-Z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 50) || 'tro-choi');
  }
  const today = () => new Date().toISOString().slice(0, 10);
  const getPath = (o, p) => p.split('.').reduce((a, k) => (a == null ? a : a[k]), o);
  function setPath(o, p, v) { const ks = p.split('.'); const last = ks.pop(); const t = ks.reduce((a, k) => a[k], o); t[last] = v; }
  const sel = () => project.items.find((i) => i.id === selectedId) || null;

  /* ================= ảnh ================= */
  function loadImg(src) {
    return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('Không đọc được ảnh')); i.src = src; });
  }
  async function toDataUrl(src, maxSide) {
    const img = await loadImg(src);
    const k = Math.min(1, (maxSide || 1600) / Math.max(img.naturalWidth, img.naturalHeight));
    const c = document.createElement('canvas'); c.width = Math.round(img.naturalWidth * k); c.height = Math.round(img.naturalHeight * k);
    const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0, c.width, c.height);
    return c.toDataURL('image/jpeg', 0.88);
  }
  async function cropTo(src, aspect) {
    const img = await loadImg(src);
    const iw = img.naturalWidth, ih = img.naturalHeight;
    let sw = iw, sh = iw / aspect;
    if (sh > ih) { sh = ih; sw = ih * aspect; }
    const sx = (iw - sw) / 2, sy = (ih - sh) / 2;
    const ow = Math.min(1600, Math.round(sw)), oh = Math.round(ow / aspect);
    const c = document.createElement('canvas'); c.width = ow; c.height = oh;
    const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, ow, oh); g.drawImage(img, sx, sy, sw, sh, 0, 0, ow, oh);
    return c.toDataURL('image/jpeg', 0.9);
  }
  async function prepItem(it) {
    const c = JSON.parse(JSON.stringify(it));
    if (c.image && ASPECT[c.type]) { try { c._img = await cropTo(c.image, ASPECT[c.type]); } catch (e) { c._img = ''; } }
    return c;
  }

  /* ================= lưu trữ ================= */
  function normalize(p) {
    const out = { title: String(p && p.title || ''), subtitle: String(p && p.subtitle || ''), items: [] };
    (p && Array.isArray(p.items) ? p.items : []).forEach((it) => {
      if (!it || !GS.GAME_TYPES[it.type]) return;
      const base = GS.newItem(it.type);
      const m = Object.assign(base, it);
      if (!m.id) m.id = base.id;
      out.items.push(m);
    });
    return out;
  }
  function saveLocal() {
    try { localStorage.setItem(DRAFT_KEY, JSON.stringify(project)); }
    catch (e) {
      if (!saveLocal.warned) { saveLocal.warned = true; toast('Bản nháp có nhiều ảnh nên không tự lưu được trên trình duyệt – hãy bấm "Lưu nháp" để lưu file.', 'bad'); }
    }
  }
  let saveT;
  function changed(o) {
    o = o || {};
    clearTimeout(saveT); saveT = setTimeout(saveLocal, 500);
    renderList();
    if (o.editor) renderEditor();
    schedulePreview(!!o.animate);
  }

  /* ================= lượt xuất ================= */
  function setCredits(n) {
    credits = typeof n === 'number' ? n : null;
    $('#creditNum').textContent = credits == null ? '?' : credits;
    const b = $('#creditBadge');
    b.classList.toggle('zero', credits === 0);
    b.classList.toggle('one', credits === 1);
  }
  async function refreshCredits() {
    try { const r = await B.balance(); setCredits(r.remainingExports); }
    catch (e) { setCredits(null); $('#modeNote').innerHTML += '<br><span style="color:#dc2626">Không kết nối được máy chủ: ' + esc(e.message) + '</span>'; }
  }
  function renderMode() {
    const m = B.mode();
    $('#modeNote').innerHTML = m === 'demo'
      ? `<b class="demo">Chế độ DEMO</b> – lượt xuất lưu trên trình duyệt này để dùng thử. Mã thử: <code>${esc(Object.keys(C.DEMO_CODES || {})[0] || '')}</code>`
      : `<b class="live">● Đang kết nối máy chủ</b> – lượt xuất và mã kích hoạt được quản lý an toàn trên máy chủ.`;
    $('#contactNote').textContent = C.CONTACT || '';
    $('#userId').textContent = B.userId().slice(4, 12).toUpperCase();
  }
  function openRedeem(msg) {
    $('#mRedeem').hidden = false;
    const m = $('#redeemMsg'); m.textContent = msg || ''; m.className = 'msg' + (msg ? ' bad' : '');
    $('#codeInput').value = ''; setTimeout(() => $('#codeInput').focus(), 50);
  }
  async function doRedeem() {
    const code = $('#codeInput').value.trim();
    const m = $('#redeemMsg');
    if (!code) { m.textContent = 'Vui lòng nhập mã kích hoạt.'; m.className = 'msg bad'; return; }
    const btn = $('#btnDoRedeem'); btn.disabled = true; m.textContent = 'Đang kiểm tra mã…'; m.className = 'msg';
    try {
      const r = await B.redeem(code);
      if (typeof r.remainingExports === 'number') setCredits(r.remainingExports);
      m.textContent = r.message || (r.success ? 'Nạp thành công!' : 'Mã không hợp lệ.');
      m.className = 'msg ' + (r.success ? 'ok' : 'bad');
      if (r.success) { toast(`Đã cộng ${r.addedCredits} lượt. Hiện có ${r.remainingExports} lượt.`, 'ok'); setTimeout(() => { $('#mRedeem').hidden = true; }, 1300); }
    } catch (e) { m.textContent = 'Lỗi: ' + e.message; m.className = 'msg bad'; }
    finally { btn.disabled = false; }
  }

  /* ================= danh sách ================= */
  function summary(it) {
    switch (it.type) {
      case 'millionaire': return it.question;
      case 'picture': return it.answer ? 'Đáp án: ' + it.answer : it.question;
      case 'puzzle': return it.keyword ? `Từ khóa: ${it.keyword} • ${it.pieceCount} mảnh` : `${it.pieceCount} mảnh ghép`;
      case 'truefalse': return it.statement;
      case 'tiles': return `${(it.cells || []).length} ô số • ${it.title || ''}`;
      case 'obstacle': return `${(it.steps || []).length} chướng ngại vật • ${it.title || ''}`;
    }
    return '';
  }
  function issues(it) {
    const w = [];
    const blank = (s) => !String(s || '').trim();
    switch (it.type) {
      case 'millionaire':
        if (blank(it.question)) w.push('chưa có câu hỏi');
        if ((it.answers || []).some(blank)) w.push('còn đáp án trống');
        break;
      case 'picture':
        if (!it.image) w.push('chưa có ảnh');
        if (blank(it.answer)) w.push('chưa có từ khóa đáp án');
        break;
      case 'puzzle':
        if (!it.image) w.push('chưa có ảnh bí mật');
        if (blank(it.keyword)) w.push('chưa có từ khóa');
        if ((it.pieces || []).slice(0, it.pieceCount).some((p) => blank(p.question))) w.push('còn mảnh ghép chưa có câu hỏi');
        break;
      case 'truefalse':
        if (blank(it.statement)) w.push('chưa có nhận định');
        break;
      case 'tiles':
        if ((it.cells || []).some((c) => blank(c.question))) w.push('còn ô số chưa có câu hỏi');
        break;
      case 'obstacle':
        if ((it.steps || []).some((c) => blank(c.question))) w.push('còn chướng ngại vật chưa có câu hỏi');
        break;
    }
    return w;
  }
  function renderList() {
    const ul = $('#itemList');
    const items = project.items;
    const n = items.filter((i) => i.includedInExport !== false).length;
    $('#selCount').innerHTML = `Đã chọn: <b>${n}/${items.length}</b> mục xuất PPTX`;
    if (!items.length) { ul.innerHTML = '<li class="empty">Chưa có trò chơi nào. Chọn dạng trò chơi bên dưới rồi bấm “Thêm”, hoặc bấm “✨ Mẫu thử”.</li>'; return; }
    ul.innerHTML = items.map((it, i) => {
      const gt = GS.GAME_TYPES[it.type];
      const w = issues(it);
      const on = it.includedInExport !== false;
      return `<li class="item${it.id === selectedId ? ' active' : ''}${on ? '' : ' off'}" data-id="${it.id}">
        <input type="checkbox" data-act="toggle" ${on ? 'checked' : ''} title="Chọn để xuất PPTX" aria-label="Chọn mục ${i + 1}">
        <span class="ico">${gt.icon}</span>
        <div class="txt"><div class="t1">${i + 1}. ${esc(gt.label)}</div><div class="t2">${esc(summary(it) || '(chưa nhập nội dung)')}</div>
        ${w.length ? `<div class="warn">⚠ ${esc(w.join(', '))}</div>` : ''}</div>
        <div class="ops">
          <button data-act="up" title="Lên">↑</button><button data-act="down" title="Xuống">↓</button>
          <button data-act="dup" title="Nhân bản">⧉</button><button data-act="del" title="Xóa">🗑</button>
        </div></li>`;
    }).join('');
  }
  async function onListClick(e) {
    const li = e.target.closest('.item'); if (!li) return;
    const id = li.dataset.id, idx = project.items.findIndex((i) => i.id === id), it = project.items[idx];
    const act = e.target.dataset.act;
    if (act === 'toggle') { it.includedInExport = e.target.checked; changed(); return; }
    if (act === 'up' && idx > 0) { project.items.splice(idx - 1, 0, project.items.splice(idx, 1)[0]); changed(); return; }
    if (act === 'down' && idx < project.items.length - 1) { project.items.splice(idx + 1, 0, project.items.splice(idx, 1)[0]); changed(); return; }
    if (act === 'dup') {
      const c = JSON.parse(JSON.stringify(it)); c.id = GS.newItem(it.type).id;
      project.items.splice(idx + 1, 0, c); selectedId = c.id; changed({ editor: true, animate: true }); return;
    }
    if (act === 'del') {
      const ok = await confirmDlg('Xóa trò chơi?', `Xóa mục <b>${idx + 1}. ${esc(GS.GAME_TYPES[it.type].label)}</b>? Thao tác này không hoàn tác được.`, 'Xóa');
      if (!ok) return;
      project.items.splice(idx, 1);
      if (selectedId === id) selectedId = (project.items[idx] || project.items[idx - 1] || {}).id || null;
      changed({ editor: true, animate: true }); return;
    }
    if (act) return;
    if (selectedId !== id) { selectedId = id; renderList(); renderEditor(); refreshPreview(true); }
  }

  /* ================= trình soạn ================= */
  const timerSel = (it) => `<label class="field"><span>⏱ Đồng hồ đếm giờ (bấm nút ⏱ trên slide để chạy)</span>
    <select data-path="timer" data-num>${TIMERS.map((t) => `<option value="${t}" ${+it.timer === t ? 'selected' : ''}>${t ? t + ' giây' : 'Tắt'}</option>`).join('')}</select></label>`;
  const txt = (label, path, val, o) => {
    o = o || {};
    return `<label class="field"><span>${label}</span>${o.multi
      ? `<textarea data-path="${path}" rows="${o.rows || 3}" maxlength="${o.max || 400}" placeholder="${esc(o.ph || '')}">${esc(val)}</textarea>`
      : `<input data-path="${path}" value="${esc(val)}" maxlength="${o.max || 200}" placeholder="${esc(o.ph || '')}">`}</label>`;
  };
  const imgBox = (it, label) => `<div class="field"><span>${label}</span><div class="img-box">
      <div class="img-thumb" style="${it.image ? `background-image:url('${it.image}')` : ''}">${it.image ? '' : 'Chưa có ảnh'}</div>
      <div class="img-tools">
        <div class="row"><button class="btn sm" data-img="pick">📁 Tải ảnh lên</button>${it.image ? '<button class="btn sm ghost danger none" data-img="clear">Xóa</button>' : ''}</div>
        <div class="row"><input type="url" id="imgUrl" placeholder="hoặc dán link ảnh (https://…)"><button class="btn sm ghost none" data-img="url">Lấy</button></div>
        <input type="file" id="imgFile" accept="image/*" hidden>
      </div></div></div><p class="hint">Ảnh sẽ được cắt vừa khung tự động. Nên dùng ảnh ngang, rõ nét.</p>`;

  function renderEditor() {
    const box = $('#editor');
    const it = sel();
    if (!it) { box.innerHTML = '<div class="empty">Chọn một trò chơi trong danh sách để soạn nội dung.</div>'; return; }
    const gt = GS.GAME_TYPES[it.type];
    const no = project.items.indexOf(it) + 1;
    let h = `<h2><span>${gt.icon}</span> ${no}. ${esc(gt.label)}</h2>`;
    switch (it.type) {
      case 'millionaire':
        h += txt('Câu hỏi', 'question', it.question, { multi: true, ph: 'Nhập câu hỏi…' });
        h += '<div class="field"><span>4 phương án – tích chọn đáp án đúng</span>';
        ['A', 'B', 'C', 'D'].forEach((L, i) => {
          h += `<div class="ans-row"><span class="lt">${L}</span><input data-path="answers.${i}" value="${esc(it.answers[i])}" maxlength="120" placeholder="Phương án ${L}">
            <label class="ok"><input type="radio" name="correct" data-path="correct" data-num value="${i}" ${+it.correct === i ? 'checked' : ''}>Đúng</label></div>`;
        });
        h += '</div>' + txt('Giải thích (hiện khi công bố đáp án)', 'explanation', it.explanation, { multi: true, rows: 2 }) + timerSel(it);
        break;
      case 'picture':
        h += txt('Câu hỏi / lời dẫn', 'question', it.question, { max: 140 });
        h += imgBox(it, 'Ảnh minh họa');
        h += txt('Gợi ý (ẩn, bấm để mở)', 'hint', it.hint, { multi: true, rows: 2, max: 200 });
        h += txt('Từ khóa đáp án', 'answer', it.answer, { max: 40, ph: 'VD: Cầu vồng' });
        h += '<p class="hint">Mỗi chữ cái thành 1 ô "?" – học sinh đoán, giáo viên bấm từng ô để mở. Nên dưới 20 chữ cái.</p>' + timerSel(it);
        break;
      case 'puzzle':
        h += imgBox(it, 'Ảnh bí mật (bị che bởi các mảnh ghép)');
        h += txt('Từ khóa bí mật', 'keyword', it.keyword, { max: 40, ph: 'VD: Mặt Trời' });
        h += `<div class="field"><span>Số mảnh ghép</span><div class="seg" data-seg="pieceCount">${[4, 6, 9].map((n) => `<button data-v="${n}" class="${+it.pieceCount === n ? 'on' : ''}">${n} mảnh</button>`).join('')}</div></div>`;
        h += '<div class="sub-list">';
        it.pieces.slice(0, it.pieceCount).forEach((p, k) => {
          h += `<div class="sub"><div class="sub-h">Mảnh ${k + 1}</div>${txt('Câu hỏi', `pieces.${k}.question`, p.question, { multi: true, rows: 2 })}${txt('Đáp án', `pieces.${k}.answer`, p.answer, { max: 120 })}</div>`;
        });
        h += '</div>' + timerSel(it);
        break;
      case 'truefalse':
        h += txt('Nhận định / phát biểu', 'statement', it.statement, { multi: true, ph: 'Nhập một nhận định để học sinh đánh giá Đúng hay Sai' });
        h += `<div class="field"><span>Kết luận</span><div class="seg" data-seg="isTrue"><button data-v="true" class="${it.isTrue !== false ? 'on' : ''}">✔ ĐÚNG</button><button data-v="false" class="${it.isTrue === false ? 'on' : ''}">✘ SAI</button></div></div>`;
        h += txt('Giải thích', 'explanation', it.explanation, { multi: true, rows: 2 }) + timerSel(it);
        break;
      case 'tiles': case 'obstacle': {
        const key = it.type === 'tiles' ? 'cells' : 'steps', min = it.type === 'tiles' ? 4 : 3;
        const unit = it.type === 'tiles' ? 'Ô số' : 'Chướng ngại vật';
        h += txt(it.type === 'tiles' ? 'Tiêu đề bảng' : 'Tiêu đề / lời dẫn', 'title', it.title, { max: 100 });
        h += '<div class="sub-list">';
        it[key].forEach((c, k) => {
          h += `<div class="sub"><div class="sub-h">${unit} ${k + 1}${it[key].length > min ? `<button class="link" data-cell-del="${k}">Xóa</button>` : ''}</div>
            ${txt('Câu hỏi', `${key}.${k}.question`, c.question, { multi: true, rows: 2 })}${txt('Đáp án', `${key}.${k}.answer`, c.answer, { max: 120 })}</div>`;
        });
        h += `</div>${it[key].length < 8 ? `<button class="btn sm" style="margin:10px 0" data-cell-add>＋ Thêm ${unit.toLowerCase()}</button>` : ''}<p class="hint">Từ ${min} đến 8 ${unit.toLowerCase()}.${it.type === 'obstacle' ? ' Nhân vật vượt lần lượt từng chướng ngại để về đích.' : ''}</p>` + timerSel(it);
        break;
      }
    }
    box.innerHTML = h;
  }

  function onEditorInput(e) {
    const t = e.target, it = sel(); if (!it || !t.dataset.path) return;
    if (t.type === 'radio' && !t.checked) return;
    let v = t.value;
    if ('num' in t.dataset) v = parseInt(v, 10) || 0;
    setPath(it, t.dataset.path, v);
    changed();
  }
  async function onEditorClick(e) {
    const it = sel(); if (!it) return;
    const seg = e.target.closest('[data-seg] button');
    if (seg) {
      const key = seg.parentElement.dataset.seg;
      let v = seg.dataset.v;
      if (key === 'pieceCount') { v = +v; while (it.pieces.length < v) it.pieces.push({ question: '', answer: '' }); }
      if (key === 'isTrue') v = v === 'true';
      it[key] = v; changed({ editor: true }); return;
    }
    const listKey = it.type === 'obstacle' ? 'steps' : 'cells';
    if ('cellAdd' in e.target.dataset) { if (it[listKey].length < 8) it[listKey].push({ question: '', answer: '' }); changed({ editor: true }); return; }
    if (e.target.dataset.cellDel != null) { it[listKey].splice(+e.target.dataset.cellDel, 1); changed({ editor: true }); return; }
    const im = e.target.dataset.img;
    if (im === 'pick') { $('#imgFile').click(); return; }
    if (im === 'clear') { it.image = ''; changed({ editor: true }); return; }
    if (im === 'url') {
      const url = $('#imgUrl').value.trim();
      if (!/^https?:\/\//i.test(url)) { toast('Link ảnh phải bắt đầu bằng http:// hoặc https://', 'bad'); return; }
      busy(true, 'Đang tải ảnh…');
      try {
        const r = await fetch(url, { mode: 'cors' });
        if (!r.ok) throw new Error('HTTP ' + r.status);
        const blob = await r.blob();
        if (!/^image\//.test(blob.type)) throw new Error('không phải file ảnh');
        const ou = URL.createObjectURL(blob);
        it.image = await toDataUrl(ou); URL.revokeObjectURL(ou);
        changed({ editor: true });
      } catch (err) {
        toast('Không lấy được ảnh từ link này (website chặn tải trực tiếp). Hãy lưu ảnh về máy rồi bấm "Tải ảnh lên".', 'bad');
      } finally { busy(false); }
    }
  }
  async function onImageFile(e) {
    const f = e.target.files && e.target.files[0]; const it = sel();
    if (!f || !it) return;
    if (!/^image\//.test(f.type)) { toast('Vui lòng chọn file ảnh.', 'bad'); return; }
    busy(true, 'Đang xử lý ảnh…');
    const ou = URL.createObjectURL(f);
    try { it.image = await toDataUrl(ou); changed({ editor: true }); }
    catch (err) { toast('Không đọc được ảnh này.', 'bad'); }
    finally { URL.revokeObjectURL(ou); busy(false); }
  }

  /* ================= xem trước ================= */
  let player, pvT;
  function previewDeck() {
    const items = project.items.filter((i) => i.includedInExport !== false || i.id === selectedId);
    return GS.buildDeck(Object.assign({}, project, { items }));
  }
  function schedulePreview(animate) { clearTimeout(pvT); pvT = setTimeout(() => refreshPreview(animate), animate ? 0 : 260); }
  function refreshPreview(animate) {
    const deck = previewDeck();
    let idx = player.deck ? player.index : 0;
    if (animate || !player.deck) { const k = deck.slides.findIndex((s) => s.itemId === selectedId); idx = k >= 0 ? k : 0; }
    idx = Math.min(idx, deck.slides.length - 1);
    player.load(deck, idx, { instant: !animate, transition: false });
    renderChips();
  }
  function renderChips() {
    const d = player.deck; if (!d) return;
    $('#pvChips').innerHTML = d.slides.map((s, i) => {
      const lb = s.key === 'title' ? 'Mở đầu' : s.key === 'end' ? 'Kết thúc' : `${i}. ${s.label}`;
      return `<button class="chip${i === player.index ? ' on' : ''}" data-i="${i}">${esc(lb)}</button>`;
    }).join('');
  }
  function onPvChange(info) {
    const s = player.deck.slides[info.index];
    $('#pvTitle').textContent = s.key === 'title' ? 'Slide mở đầu (menu)' : s.key === 'end' ? 'Slide kết thúc' : s.label;
    $('#pvInfo').textContent = `Slide ${info.index + 1}/${info.total} • Bước bấm ${Math.max(0, info.step - (s.main[0] && s.main[0].auto ? 1 : 0))}/${s.main.filter((g) => !g.auto).length} • ${s.triggers.length} trigger`;
    document.querySelectorAll('#pvChips .chip').forEach((c, i) => c.classList.toggle('on', i === info.index));
  }

  /* ================= trình chiếu ================= */
  let pres = null;
  function openPresenter() {
    const items = project.items.filter((i) => i.includedInExport !== false);
    if (!items.length) { toast('Hãy tích chọn ít nhất 1 trò chơi để trình chiếu.', 'bad'); return; }
    const deck = GS.buildDeck(Object.assign({}, project, { items }));
    const el = $('#presenter'); el.hidden = false;
    $('#prStage').innerHTML = '';
    pres = new window.GamePlayer($('#prStage'), {
      onChange: (i) => { $('#prInfo').textContent = `${i.index + 1} / ${i.total}`; },
      onEnd: closePresenter
    });
    pres.load(deck, 0, { transition: false });
    if (el.requestFullscreen) el.requestFullscreen().catch(() => { /* trình duyệt từ chối */ });
  }
  function closePresenter() {
    if (!pres) return;
    pres._clear(); pres = null; $('#presenter').hidden = true; $('#prStage').innerHTML = '';
    if (document.fullscreenElement) document.exitFullscreen().catch(() => { /* bỏ qua */ });
  }

  /* ================= xuất PPTX ================= */
  async function doExport() {
    const items = project.items.filter((i) => i.includedInExport !== false);
    if (!items.length) { toast('Bạn chưa chọn mục nào để xuất. Hãy tích chọn ít nhất 1 trò chơi (không bị trừ lượt).', 'bad'); return; }
    const warns = [];
    items.forEach((it) => { const w = issues(it); if (w.length) warns.push(`<li><b>${project.items.indexOf(it) + 1}. ${esc(GS.GAME_TYPES[it.type].label)}</b>: ${esc(w.join(', '))}</li>`); });
    if (warns.length && !(await confirmDlg('Một số mục chưa hoàn chỉnh', `<ul>${warns.join('')}</ul><p>Bạn vẫn muốn xuất file?</p>`, 'Vẫn xuất', 'Quay lại sửa'))) return;
    if (credits === 0) { openRedeem('Bạn đã hết lượt xuất PPTX. Vui lòng nhập mã kích hoạt để nạp thêm lượt.'); return; }
    const slides = items.length + 2;
    if (!(await confirmDlg('Xuất PowerPoint', `<p>Xuất <b>${items.length}</b> trò chơi thành <b>${slides} slide</b> có hiệu ứng tương tác.</p>
      <label class="opt-row"><input type="checkbox" id="optProtect" checked><span><b>🔒 Khóa nội dung (khuyên dùng)</b><br><span class="muted">Chữ và hình được chuyển thành ảnh, khóa không cho di chuyển/sửa — người nhận file không sửa được câu hỏi hay sao chép bố cục để tự làm lại. Hiệu ứng, trigger, âm thanh vẫn chạy bình thường.</span></span></label>
      <label class="opt-row"><input type="checkbox" id="optSound" checked><span><b>🔊 Kèm hiệu ứng âm thanh</b></span></label>
      <p>Thao tác này trừ <b>1 lượt xuất</b>${credits != null ? ` (hiện còn ${credits})` : ''}.</p>`, 'Xuất file'))) return;
    const protect = !!($('#optProtect') && $('#optProtect').checked);
    const sound = !!($('#optSound') && $('#optSound').checked);
    busy(true, 'Đang chuẩn bị hình ảnh…');
    try {
      const prepared = [];
      for (const it of items) prepared.push(await prepItem(it));
      const deck = GS.buildDeck(Object.assign({}, project, { items: prepared }));
      busy(true, protect ? 'Đang khóa nội dung và tạo PowerPoint…' : 'Đang tạo PowerPoint có hiệu ứng…');
      const blob = await window.PptxExporter.exportDeck(deck, { PptxGenJS: window.PptxGenJS, JSZip: window.JSZip, Sounds: window.GSSounds }, {
        protect, sound, rasterize: window.GSRaster && window.GSRaster.rasterizeSlide,
        onProgress: (i, n) => busy(true, `Đang khóa nội dung slide ${i + 1}/${n}…`)
      });
      busy(true, 'Đang xác nhận lượt xuất…');
      const r = await B.consume(items.length, deck.slides.length);
      if (typeof r.remainingExports === 'number') setCredits(r.remainingExports);
      if (!r.allowed) { busy(false); openRedeem(r.message || 'Bạn đã hết lượt xuất PPTX.'); return; }
      download(blob, `${slug(project.title)}-${today()}.pptx`);
      busy(false);
      toast(`Đã xuất file PowerPoint! Còn ${r.remainingExports} lượt.`, 'ok');
    } catch (e) {
      console.error(e); busy(false);
      toast('Lỗi khi xuất file: ' + e.message, 'bad');
    }
  }

  /* ================= nhập câu hỏi từ Word / ảnh ================= */
  const imp = { images: [], result: null };
  const KIND = { mcq: ['Trắc nghiệm', ''], tf: ['Đúng/Sai', 'tf'], short: ['Tự luận ngắn', 'short'] };
  function impStatus(html) { $('#impStatus').innerHTML = html || ''; }
  function openImport() { $('#mImport').hidden = false; impAnalyze(); }
  async function impFiles(files) {
    files = Array.from(files || []);
    if (!files.length) return;
    let added = '';
    for (const f of files) {
      const name = f.name || 'file';
      try {
        if (/\.docx$/i.test(name) || /wordprocessingml/.test(f.type)) {
          impStatus(`Đang đọc <b>${esc(name)}</b>…`);
          const r = await window.GSImporter.readDocx(f, window.JSZip);
          const off = imp.images.length;
          imp.images.push(...r.images);
          added += '\n' + r.text.replace(/\[Ảnh (\d+)\]/g, (m, n) => `[Ảnh ${+n + off}]`) + '\n';
        } else if (/^image\//.test(f.type)) {
          impStatus(`Đang nhận dạng chữ trong <b>${esc(name)}</b>… (lần đầu cần tải bộ nhận dạng tiếng Việt, vui lòng chờ)`);
          const text = await window.GSImporter.ocrImage(f, (st, p) => {
            const lb = /recogniz/i.test(st) ? 'Đang nhận dạng chữ' : /load|initializ/i.test(st) ? 'Đang tải bộ nhận dạng tiếng Việt' : 'Đang xử lý';
            impStatus(`${lb} – <b>${esc(name)}</b> ${Math.round(p * 100)}%<div class="bar"><i style="width:${Math.round(p * 100)}%"></i></div>`);
          });
          added += '\n' + text + '\n';
        } else if (/\.doc$/i.test(name)) {
          toast(`"${name}" là Word bản cũ (.doc). Hãy mở bằng Word và Lưu thành .docx rồi chọn lại.`, 'bad');
        } else toast(`Không hỗ trợ file "${name}". Hãy dùng .docx hoặc ảnh.`, 'bad');
      } catch (e) { console.error(e); toast(`Không đọc được "${name}": ${e.message}`, 'bad'); }
    }
    if (added.trim()) {
      const ta = $('#impText');
      ta.value = (ta.value.trim() ? ta.value.trim() + '\n\n' : '') + added.trim();
      impStatus(`✔ Đã đọc xong. Kiểm tra nội dung bên dưới (sửa nếu nhận dạng sai) rồi bấm “Thêm vào bài”.`);
    } else impStatus('');
    impAnalyze();
  }
  function impAnalyze() {
    const text = $('#impText').value;
    const box = $('#impPreview');
    if (!text.trim()) { box.innerHTML = ''; $('#impAdd').disabled = true; imp.result = null; return; }
    const qs = window.GSImporter.parseText(text, imp.images);
    const target = $('#impTarget').value;
    const r = window.GSImporter.toItems(qs, target, GS.newItem);
    imp.result = r;
    const cnt = (k) => qs.filter((q) => q.kind === k).length;
    const made = {};
    r.items.forEach((i) => { made[i.type] = (made[i.type] || 0) + 1; });
    const madeTxt = Object.keys(made).map((t) => `${made[t]} ${GS.GAME_TYPES[t].label}`).join(', ');
    box.innerHTML = `<div class="sum">Tìm thấy ${qs.length} câu: ${cnt('mcq')} trắc nghiệm, ${cnt('tf')} đúng/sai, ${cnt('short')} tự luận ngắn
      ${r.items.length ? `→ sẽ tạo <span style="color:var(--ac)">${esc(madeTxt)}</span>` : ''}</div>
      ${r.skipped.length ? `<div class="muted small">⚠ ${r.skipped.length} câu không phù hợp dạng đã chọn nên bị bỏ qua${target === 'picture' ? ' (dạng này cần câu hỏi có ảnh trong file Word)' : ''}.</div>` : ''}
      <ol>${qs.map((q) => `<li><span class="badge ${KIND[q.kind][1]}">${KIND[q.kind][0]}</span>${q.answerKnown ? '' : '<span class="badge warn">chưa rõ đáp án</span>'}${q.image ? '<span class="badge">có ảnh</span>' : ''}
        ${esc(q.text.slice(0, 110))}${q.text.length > 110 ? '…' : ''}
        <span class="muted"> → ${esc(q.kind === 'mcq' ? 'ABCD'[q.correct] + '. ' + q.options[q.correct] : q.kind === 'tf' ? (q.isTrue ? 'Đúng' : 'Sai') : q.answer || '?')}</span></li>`).join('')}</ol>
      ${!qs.length ? '<div class="muted">Chưa nhận ra câu hỏi nào. Hãy trình bày theo mẫu “Câu 1: …”, “A. … B. …”, “Đáp án: …”.</div>' : ''}`;
    $('#impAdd').disabled = !r.items.length;
    $('#impAdd').textContent = r.items.length ? `＋ Thêm ${r.items.length} trò chơi vào bài` : '＋ Thêm vào bài';
  }
  function impAdd() {
    if (!imp.result || !imp.result.items.length) return;
    const items = imp.result.items;
    project.items.push(...items);
    selectedId = items[0].id;
    $('#mImport').hidden = true;
    $('#impText').value = ''; imp.images = []; imp.result = null; impStatus(''); $('#impPreview').innerHTML = '';
    changed({ editor: true, animate: true });
    toast(`Đã thêm ${items.length} trò chơi. Kiểm tra lại các mục có dấu ⚠ nếu có.`, 'ok');
  }
  function bindImport() {
    $('#btnImport').onclick = openImport;
    $('#impFile').onchange = (e) => { const f = e.target.files; impFiles(f).finally(() => { e.target.value = ''; }); };
    const dz = $('#dropZone');
    ['dragenter', 'dragover'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove('over'); }));
    dz.addEventListener('drop', (e) => impFiles(e.dataTransfer.files));
    let t;
    $('#impText').addEventListener('input', () => { clearTimeout(t); t = setTimeout(impAnalyze, 300); });
    $('#impTarget').onchange = impAnalyze;
    $('#impAdd').onclick = impAdd;
  }

  /* ================= âm thanh ================= */
  function bindSound() {
    let m = false;
    try { m = localStorage.getItem('gs_muted') === '1'; } catch (e) { /* bỏ qua */ }
    const apply = () => { window.GSSounds.setMuted(m); $('#btnSound').textContent = m ? '🔇' : '🔊'; $('#btnSound').title = m ? 'Đang tắt âm thanh – bấm để bật' : 'Đang bật âm thanh – bấm để tắt'; };
    $('#btnSound').onclick = () => { m = !m; try { localStorage.setItem('gs_muted', m ? '1' : '0'); } catch (e) { /* bỏ qua */ } apply(); if (!m) window.GSSounds.play('pop'); };
    apply();
  }

  /* ================= khởi động ================= */
  async function loadSample() {
    project = normalize(await window.GSSamples.sampleProject());
    selectedId = project.items[0] && project.items[0].id;
    $('#pTitle').value = project.title; $('#pSub').value = project.subtitle;
    changed({ editor: true, animate: true });
  }
  function bind() {
    $('#itemList').addEventListener('click', onListClick);
    $('#editor').addEventListener('input', onEditorInput);
    $('#editor').addEventListener('change', (e) => { if (e.target.id === 'imgFile') onImageFile(e); else if (e.target.type === 'radio') onEditorInput(e); });
    $('#editor').addEventListener('click', onEditorClick);
    $('#pTitle').addEventListener('input', (e) => { project.title = e.target.value; changed(); });
    $('#pSub').addEventListener('input', (e) => { project.subtitle = e.target.value; changed(); });
    $('#selAll').onclick = () => { project.items.forEach((i) => { i.includedInExport = true; }); changed(); };
    $('#selNone').onclick = () => { project.items.forEach((i) => { i.includedInExport = false; }); changed(); };
    $('#btnAdd').onclick = () => {
      const it = GS.newItem($('#addType').value);
      project.items.push(it); selectedId = it.id; changed({ editor: true, animate: true });
      setTimeout(() => { const f = $('#editor textarea, #editor input'); if (f) f.focus(); $('#editor').scrollIntoView({ behavior: 'smooth', block: 'start' }); }, 50);
    };
    $('#btnSample').onclick = async () => {
      if (project.items.length && !(await confirmDlg('Tải bộ mẫu?', 'Bộ câu hỏi hiện tại sẽ được thay bằng bộ mẫu 5 trò chơi. Hãy “Lưu nháp” trước nếu muốn giữ lại.', 'Tải mẫu'))) return;
      await loadSample(); toast('Đã tải bộ mẫu 5 trò chơi.', 'ok');
    };
    $('#btnSave').onclick = () => {
      download(new Blob([JSON.stringify({ app: 'pptx-game-studio', version: 1, project }, null, 1)], { type: 'application/json' }), `${slug(project.title)}-nhap.json`);
    };
    $('#btnOpen').onclick = () => $('#fileDraft').click();
    $('#fileDraft').onchange = async (e) => {
      const f = e.target.files[0]; e.target.value = ''; if (!f) return;
      try {
        const data = JSON.parse(await f.text());
        const p = normalize(data.project || data);
        if (!p.items.length && !(data.project || data).items) throw new Error('File không đúng định dạng nháp');
        project = p; selectedId = p.items[0] && p.items[0].id;
        $('#pTitle').value = p.title; $('#pSub').value = p.subtitle;
        changed({ editor: true, animate: true }); toast('Đã mở bản nháp.', 'ok');
      } catch (err) { toast('Không mở được file nháp: ' + err.message, 'bad'); }
    };
    $('#btnExport').onclick = doExport;
    $('#btnRedeem').onclick = () => openRedeem();
    $('#btnDoRedeem').onclick = doRedeem;
    $('#codeInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') doRedeem(); });
    document.querySelectorAll('[data-close]').forEach((b) => { b.onclick = () => { b.closest('.modal').hidden = true; }; });
    $('#pvPrev').onclick = () => player.prev();
    $('#pvNext').onclick = () => player.next();
    $('#pvReplay').onclick = () => player.replay();
    $('#pvChips').onclick = (e) => { const b = e.target.closest('.chip'); if (b) player.show(+b.dataset.i); };
    $('#btnPresent').onclick = openPresenter;
    $('#prNext').onclick = () => pres && pres.next();
    $('#prPrev').onclick = () => pres && pres.prev();
    $('#prClose').onclick = closePresenter;
    document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement && pres) closePresenter(); });
    document.addEventListener('keydown', (e) => {
      if (pres) {
        if (['ArrowRight', 'PageDown', ' ', 'Enter'].includes(e.key)) { e.preventDefault(); pres.next(); }
        else if (['ArrowLeft', 'PageUp'].includes(e.key)) { e.preventDefault(); pres.prev(); }
        else if (e.key === 'Home') pres.show(0);
        else if (e.key === 'Escape') closePresenter();
        return;
      }
      if (e.key === 'Escape') document.querySelectorAll('.modal:not(.busy)').forEach((m) => { if (!m.hidden && m.id !== 'mConfirm') m.hidden = true; });
    });
  }

  async function init() {
    player = new window.GamePlayer($('#pvStage'), { onChange: onPvChange });
    bind(); bindImport(); bindSound(); renderMode(); refreshCredits();
    let saved = null;
    try { saved = JSON.parse(localStorage.getItem(DRAFT_KEY) || 'null'); } catch (e) { saved = null; }
    if (saved && Array.isArray(saved.items) && saved.items.length) {
      project = normalize(saved); selectedId = project.items[0].id;
      $('#pTitle').value = project.title; $('#pSub').value = project.subtitle;
      changed({ editor: true, animate: true });
    } else await loadSample();
  }
  init();
})();
