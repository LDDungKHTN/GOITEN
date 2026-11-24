/*!
 * PPTX Game Studio – Trình phát HTML: hiển thị đúng các "cảnh" của GameScenes với hiệu ứng
 * và trigger giống hệt khi trình chiếu file PowerPoint.
 */
(function (root) {
  'use strict';
  const W = 13.333, H = 7.5;
  const ENTR = new Set(['fadeIn', 'zoomIn', 'floatIn', 'wipeIn']);
  const EXIT = new Set(['fadeOut', 'zoomOut', 'timerOut']);
  const u = (inch) => `${(inch / W) * 100}cqw`;   // inch -> đơn vị theo bề rộng slide
  const pt = (p) => u(p / 72);

  function rgba(hex, transp) {
    hex = String(hex || '000000').replace('#', '');
    const n = parseInt(hex, 16);
    return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${1 - (transp || 0) / 100})`;
  }
  function clipFor(sh) {
    if (sh.shape === 'hexagon') {
      const d = (0.25 * Math.min(sh.w, sh.h)) / sh.w * 100;
      return `polygon(${d}% 0, ${100 - d}% 0, 100% 50%, ${100 - d}% 100%, ${d}% 100%, 0 50%)`;
    }
    if (sh.shape === 'star5') return 'polygon(50% 0,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%)';
    return '';
  }
  const KF = {
    fadeIn: [{ opacity: 0 }, { opacity: 1 }],
    zoomIn: [{ transform: 'scale(0)', opacity: 0 }, { transform: 'scale(1)', opacity: 1 }],
    floatIn: [{ transform: `translateY(${u(0.75)})`, opacity: 0 }, { transform: 'translateY(0)', opacity: 1 }],
    wipeIn: [{ clipPath: 'inset(0 100% 0 0)' }, { clipPath: 'inset(0 0 0 0)' }],
    fadeOut: [{ opacity: 1 }, { opacity: 0 }],
    zoomOut: [{ transform: 'scale(1)', opacity: 1 }, { transform: 'scale(0)', opacity: 0 }],
    timerOut: [{ clipPath: 'inset(0 0 0 0)' }, { clipPath: 'inset(0 100% 0 0)' }],
    pulse: [{ transform: 'scale(1)' }, { transform: 'scale(1.13)' }, { transform: 'scale(1)' }]
  };
  const EASE = { zoomIn: 'cubic-bezier(.2,1.25,.4,1)', floatIn: 'cubic-bezier(.15,.85,.3,1)', zoomOut: 'cubic-bezier(.5,0,.9,.4)', timerOut: 'linear', pulse: 'ease-in-out' };

  class Player {
    constructor(host, opts) {
      this.host = host; this.opts = opts || {};
      host.classList.add('gs-player');
      this.stage = document.createElement('div');
      this.stage.className = 'gs-stage';
      host.appendChild(this.stage);
      this.deck = null; this.index = 0; this.timers = []; this.anims = []; this.layer = null;
      this.stage.addEventListener('click', (e) => this._onClick(e));
    }
    load(deck, index, o) { this.deck = deck; this.show(index || 0, Object.assign({ transition: false }, o)); }
    get slide() { return this.deck && this.deck.slides[this.index]; }

    _clear() {
      this.timers.forEach(clearTimeout); this.timers = [];
      this.anims.forEach((a) => { try { a.cancel(); } catch (e) { /* bỏ qua */ } }); this.anims = [];
    }

    _build(sd) {
      const layer = document.createElement('div');
      layer.className = 'gs-slide';
      layer.style.background = rgba(sd.bg);
      const nodes = {};
      const trig = new Set(sd.triggers.map((t) => t.on));
      sd.shapes.forEach((sh) => {
        const o = document.createElement('div');
        o.className = 'gs-shp';
        o.dataset.name = sh.name;
        Object.assign(o.style, { left: (sh.x / W) * 100 + '%', top: (sh.y / H) * 100 + '%', width: (sh.w / W) * 100 + '%', height: (sh.h / H) * 100 + '%' });
        if (sh.rot) o.style.transform = `rotate(${sh.rot}deg)`;
        const a = document.createElement('div'); a.className = 'gs-anim';
        if (sh.shadow) a.style.filter = `drop-shadow(0 ${u(0.04)} ${u(0.08)} rgba(0,0,0,.4))`;
        o.appendChild(a);
        let body;
        if (sh.kind === 'image') {
          body = document.createElement('img'); body.className = 'gs-img'; body.src = sh.data; body.alt = ''; body.draggable = false;
        } else {
          body = document.createElement('div'); body.className = 'gs-body';
          const clip = clipFor(sh);
          let fillEl = body;
          const lw = sh.line ? u((sh.lineW || 1) / 72) : 0;
          if (clip) {
            body.style.clipPath = clip;
            if (sh.line) body.style.background = rgba(sh.line);
            fillEl = document.createElement('div'); fillEl.className = 'gs-fill';
            fillEl.style.clipPath = clip; if (sh.line) fillEl.style.inset = lw;
            body.appendChild(fillEl);
          } else if (sh.line) {
            body.style.border = `${lw} ${sh.dash ? 'dashed' : 'solid'} ${rgba(sh.line)}`;
          }
          if (sh.fill) fillEl.style.background = rgba(sh.fill, sh.alpha);
          if (sh.shape === 'ellipse') body.style.borderRadius = '50%';
          else if (sh.shape === 'roundRect') body.style.borderRadius = u(Math.min(sh.radius || 0.1, Math.min(sh.w, sh.h) / 2));
          const tx = document.createElement('div'); tx.className = 'gs-text';
          const pad = sh.pad || [6, 6, 3, 3];
          Object.assign(tx.style, {
            padding: `${pt(pad[3])} ${pt(pad[1])} ${pt(pad[2])} ${pt(pad[0])}`,
            justifyContent: sh.valign === 'top' ? 'flex-start' : sh.valign === 'bottom' ? 'flex-end' : 'center',
            textAlign: sh.align || 'center', color: rgba(sh.color || 'FFFFFF'), fontSize: pt(sh.size || 18),
            fontWeight: sh.face === 'Segoe UI Black' ? 900 : sh.bold ? 700 : 400, fontFamily: faceCss(sh.face)
          });
          const p = document.createElement('div');
          if (sh.runs) sh.runs.forEach((r) => {
            const s = document.createElement('span'); s.textContent = r.t || '';
            if (r.color) s.style.color = rgba(r.color);
            if (r.size) s.style.fontSize = pt(r.size);
            if (r.bold != null) s.style.fontWeight = r.bold ? 700 : 400;
            if (r.face) { s.style.fontFamily = faceCss(r.face); if (r.face === 'Segoe UI Black') s.style.fontWeight = 900; }
            p.appendChild(s);
          });
          else p.textContent = sh.text || '';
          tx.appendChild(p); fillEl.appendChild(tx);
          if (!sh.fill && !sh.text && !sh.runs) o.style.pointerEvents = 'none';
        }
        a.appendChild(body);
        if (trig.has(sh.name)) o.classList.add('gs-trig');
        if (sh.link) { o.classList.add('gs-link'); o.dataset.link = sh.link; if (sh.tip) o.title = sh.tip; }
        if (sh.kind === 'text' && !sh.link && !trig.has(sh.name)) o.style.pointerEvents = 'none';
        layer.appendChild(o);
        nodes[sh.name] = { o, a };
      });
      // ẩn các hình có hiệu ứng đầu tiên là "xuất hiện"
      const first = {};
      sd.main.forEach((g) => g.effects.forEach((e) => { if (!(e.target in first)) first[e.target] = e.type; }));
      sd.triggers.forEach((t) => t.effects.forEach((e) => { if (!(e.target in first)) first[e.target] = e.type; }));
      Object.keys(first).forEach((k) => { if (ENTR.has(first[k]) && nodes[k]) nodes[k].o.style.visibility = 'hidden'; });
      return { layer, nodes };
    }

    show(i, o) {
      o = o || {};
      if (!this.deck) return;
      i = Math.max(0, Math.min(this.deck.slides.length - 1, i));
      this._clear();
      const prev = this.layer;
      this.index = i;
      const sd = this.deck.slides[i];
      const { layer, nodes } = this._build(sd);
      this.layer = layer; this.nodes = nodes; this.mainPos = 0;
      this.stage.appendChild(layer);
      if (prev) {
        if (o.transition !== false) {
          const t = this._transition(sd.transition, layer, prev);
          t.onfinish = () => prev.remove();
          setTimeout(() => prev.remove(), 1200);
        } else prev.remove();
      }
      const auto = sd.main[0] && sd.main[0].auto;
      if (auto) { this._play(sd.main[0].effects, !!o.instant, o.transition !== false && prev ? 450 : 0); this.mainPos = 1; }
      if (o.revealAll) while (this.mainPos < sd.main.length) this._play(sd.main[this.mainPos++].effects, true);
      this._emit();
    }

    _transition(kind, layer, prev) {
      const d = { duration: 700, easing: 'cubic-bezier(.3,.7,.2,1)' };
      switch (kind) {
        case 'push':
          prev.animate([{ transform: 'translateY(0)' }, { transform: 'translateY(-100%)' }], d);
          return layer.animate([{ transform: 'translateY(100%)' }, { transform: 'translateY(0)' }], d);
        case 'zoom': return layer.animate([{ transform: 'scale(.4)', opacity: 0 }, { transform: 'scale(1)', opacity: 1 }], d);
        case 'cover': return layer.animate([{ transform: 'translateX(100%)' }, { transform: 'translateX(0)' }], d);
        case 'wipe': return layer.animate([{ clipPath: 'inset(0 100% 0 0)' }, { clipPath: 'inset(0 0 0 0)' }], d);
        default: return layer.animate([{ opacity: 0 }, { opacity: 1 }], d);
      }
    }

    _run(e, instant) {
      const n = this.nodes[e.target]; if (!n) return;
      const isIn = ENTR.has(e.type), isOut = EXIT.has(e.type);
      if (isIn) n.o.style.visibility = 'visible';
      if (instant) { if (isOut) n.o.style.visibility = 'hidden'; return; }
      if (!KF[e.type]) return;
      const a = n.a.animate(KF[e.type], { duration: Math.max(1, e.dur), easing: EASE[e.type] || 'ease-out' });
      this.anims.push(a);
      if (isOut) {
        n.o.style.visibility = 'visible';
        a.onfinish = () => { n.o.style.visibility = 'hidden'; };
      }
      a.addEventListener('finish', () => { const k = this.anims.indexOf(a); if (k >= 0) this.anims.splice(k, 1); });
    }

    _play(effects, instant, offset) {
      const list = effects.slice().sort((a, b) => a.delay - b.delay);
      if (instant) { list.forEach((e) => this._run(e, true)); return; }
      list.forEach((e) => {
        const d = e.delay + (offset || 0);
        if (d <= 0) this._run(e);
        else this.timers.push(setTimeout(() => this._run(e), d));
      });
    }

    /** Bước tiếp: chạy nhóm hiệu ứng bấm chuột tiếp theo, hết thì sang slide sau. */
    next() {
      const sd = this.slide; if (!sd) return;
      if (this.mainPos < sd.main.length) { this._play(sd.main[this.mainPos].effects); this.mainPos++; this._emit(); }
      else if (this.index < this.deck.slides.length - 1) this.show(this.index + 1);
      else if (this.opts.onEnd) this.opts.onEnd();
    }
    prev() { if (this.index > 0) this.show(this.index - 1, { instant: true, revealAll: false }); else this.show(0); }
    replay() { this.show(this.index, { transition: false }); }
    goTo(i) { this.show(i); }

    _onClick(ev) {
      if (!this.layer) return;
      const el = ev.target.closest('.gs-shp');
      if (el && el.style.visibility !== 'hidden') {
        const name = el.dataset.name;
        const t = this.slide.triggers.find((x) => x.on === name);
        if (t) { this._play(t.effects); return; }
        if (el.dataset.link) { this.show(parseInt(el.dataset.link, 10) - 1); return; }
      }
      if (this.opts.clickAdvance !== false) this.next();
    }
    _emit() {
      if (this.opts.onChange) this.opts.onChange({ index: this.index, total: this.deck.slides.length, step: this.mainPos, steps: this.slide.main.length });
    }
  }
  function faceCss(face) {
    if (face === 'Segoe UI Black') return "'Segoe UI Black','Segoe UI',system-ui,sans-serif";
    if (face === 'Segoe UI Symbol') return "'Segoe UI Symbol','Segoe UI',system-ui,sans-serif";
    return "'Segoe UI',system-ui,-apple-system,Roboto,'Helvetica Neue',Arial,sans-serif";
  }
  root.GamePlayer = Player;
})(typeof self !== 'undefined' ? self : this);
