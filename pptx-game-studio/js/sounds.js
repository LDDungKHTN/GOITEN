/*!
 * PPTX Game Studio – Hiệu ứng âm thanh tự tổng hợp (không cần file ngoài, không bản quyền).
 * Cùng một bộ tổng hợp được dùng để: phát trên web (Web Audio) và nhúng file .wav vào PowerPoint.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.GSSounds = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const NAMES = ['pop', 'correct', 'wrong', 'reveal', 'whoosh', 'step', 'timeup', 'win', 'intro'];

  function mk(sr, sec) { return new Float32Array(Math.ceil(sr * sec)); }
  // nốt nhạc có bao (attack nhanh, tắt dần), hỗ trợ trượt tần số
  function note(b, sr, t0, dur, f0, o) {
    o = o || {};
    const f1 = o.to || f0, vol = o.vol == null ? 0.5 : o.vol, type = o.type || 'sine';
    const s0 = Math.floor(t0 * sr), n = Math.floor(dur * sr);
    let ph = 0;
    for (let i = 0; i < n && s0 + i < b.length; i++) {
      const t = i / n;
      const f = f0 + (f1 - f0) * t;
      ph += (2 * Math.PI * f) / sr;
      let v;
      if (type === 'square') v = Math.sin(ph) > 0 ? 0.6 : -0.6;
      else if (type === 'triangle') v = (2 / Math.PI) * Math.asin(Math.sin(ph));
      else if (type === 'bell') v = Math.sin(ph) * 0.7 + Math.sin(ph * 2.01) * 0.2 + Math.sin(ph * 3.02) * 0.1;
      else v = Math.sin(ph);
      const att = Math.min(1, i / (sr * 0.004));
      const env = att * Math.pow(1 - t, o.decay || 2);
      b[s0 + i] += v * env * vol;
    }
  }
  function noise(b, sr, t0, dur, o) {
    o = o || {};
    const s0 = Math.floor(t0 * sr), n = Math.floor(dur * sr);
    let lp = 0, seed = 12345;
    for (let i = 0; i < n && s0 + i < b.length; i++) {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      const w = (seed / 0x7fffffff) * 2 - 1;
      const t = i / n;
      const cut = o.sweep ? 0.02 + 0.5 * Math.sin(Math.PI * t) : (o.cut || 0.3);
      lp += (w - lp) * cut;
      const env = Math.sin(Math.PI * Math.min(1, t * (o.shape || 1))) * (1 - t * 0.3);
      b[s0 + i] += lp * env * (o.vol || 0.5);
    }
  }
  function finish(b) {
    let m = 0;
    for (let i = 0; i < b.length; i++) m = Math.max(m, Math.abs(b[i]));
    const k = m > 0.95 ? 0.95 / m : 1;
    const fade = Math.min(200, b.length);
    for (let i = 0; i < b.length; i++) { b[i] *= k; if (i > b.length - fade) b[i] *= (b.length - i) / fade; }
    return b;
  }

  /** Tạo mẫu âm thanh (Float32 -1..1). name: tên hiệu ứng hoặc 'tick<số giây>' */
  function synth(name, sr) {
    sr = sr || 22050;
    let b;
    const tk = /^tick(\d+)$/.exec(name);
    if (tk) {
      const sec = Math.max(1, Math.min(300, +tk[1]));
      b = mk(sr, sec);
      for (let s = 0; s < sec; s++) {
        const last = sec - s <= 5;
        note(b, sr, s, 0.05, last ? 1400 : 1000, { type: 'square', vol: last ? 0.45 : 0.3, decay: 3 });
        if (!last) note(b, sr, s + 0.5, 0.04, 700, { type: 'square', vol: 0.18, decay: 3 });
      }
      return finish(b);
    }
    switch (name) {
      case 'pop': b = mk(sr, 0.16); note(b, sr, 0, 0.14, 520, { to: 1250, vol: 0.7, decay: 1.5 }); break;
      case 'correct': b = mk(sr, 0.62);
        note(b, sr, 0, 0.16, 880, { type: 'bell', vol: 0.55 }); note(b, sr, 0.12, 0.5, 1318.5, { type: 'bell', vol: 0.6, decay: 2.5 }); break;
      case 'wrong': b = mk(sr, 0.6);
        note(b, sr, 0, 0.26, 220, { to: 180, type: 'square', vol: 0.32, decay: 0.6 }); note(b, sr, 0.27, 0.32, 175, { to: 120, type: 'square', vol: 0.32, decay: 1 }); break;
      case 'reveal': b = mk(sr, 0.55);
        [784, 988, 1175, 1568].forEach((f, i) => note(b, sr, i * 0.07, 0.3, f, { type: 'bell', vol: 0.32, decay: 2.5 })); break;
      case 'whoosh': b = mk(sr, 0.4); noise(b, sr, 0, 0.38, { sweep: true, vol: 0.9 }); break;
      case 'step': b = mk(sr, 0.35);
        note(b, sr, 0, 0.14, 330, { to: 760, type: 'triangle', vol: 0.55, decay: 1 }); note(b, sr, 0.14, 0.18, 660, { to: 990, type: 'triangle', vol: 0.45, decay: 1.5 }); break;
      case 'timeup': b = mk(sr, 0.95);
        for (let i = 0; i < 3; i++) note(b, sr, i * 0.28, 0.22, 950, { type: 'square', vol: 0.32, decay: 0.8 }); break;
      case 'win': b = mk(sr, 1.6);
        [523.25, 659.25, 783.99].forEach((f, i) => note(b, sr, i * 0.12, 0.2, f, { type: 'triangle', vol: 0.45, decay: 1.2 }));
        [523.25, 659.25, 783.99, 1046.5].forEach((f) => note(b, sr, 0.38, 1.15, f, { type: 'bell', vol: 0.28, decay: 2 }));
        noise(b, sr, 0.38, 1.1, { cut: 0.6, vol: 0.12, shape: 1 }); break;
      case 'intro': b = mk(sr, 1.3);
        [392, 523.25, 659.25, 783.99, 1046.5].forEach((f, i) => note(b, sr, i * 0.08, 0.25, f, { type: 'bell', vol: 0.32, decay: 2 }));
        [523.25, 783.99, 1046.5].forEach((f) => note(b, sr, 0.45, 0.85, f, { type: 'triangle', vol: 0.22, decay: 1.6 })); break;
      default: b = mk(sr, 0.1);
    }
    return finish(b);
  }

  /** Mã hóa WAV PCM mono. bits 8 hoặc 16 */
  function wav(samples, sr, bits) {
    bits = bits || 16;
    const bps = bits / 8, n = samples.length;
    const buf = new ArrayBuffer(44 + n * bps), v = new DataView(buf);
    const str = (o, s) => { for (let i = 0; i < s.length; i++) v.setUint8(o + i, s.charCodeAt(i)); };
    str(0, 'RIFF'); v.setUint32(4, 36 + n * bps, true); str(8, 'WAVE'); str(12, 'fmt ');
    v.setUint32(16, 16, true); v.setUint16(20, 1, true); v.setUint16(22, 1, true); v.setUint32(24, sr, true);
    v.setUint32(28, sr * bps, true); v.setUint16(32, bps, true); v.setUint16(34, bits, true); str(36, 'data'); v.setUint32(40, n * bps, true);
    for (let i = 0; i < n; i++) {
      const s = Math.max(-1, Math.min(1, samples[i]));
      if (bits === 8) v.setUint8(44 + i, Math.round((s + 1) * 127.5));
      else v.setInt16(44 + i * 2, Math.round(s * 32767), true);
    }
    return new Uint8Array(buf);
  }
  /** File .wav dùng để nhúng vào PowerPoint */
  function wavFor(name) {
    if (/^tick\d+$/.test(name)) return wav(synth(name, 8000), 8000, 8);
    return wav(synth(name, 22050), 22050, 16);
  }

  /* ---------- phát trên trình duyệt ---------- */
  let ctx = null, muted = false;
  const cache = {};
  const live = new Set();
  function play(name) {
    if (muted || !name || typeof window === 'undefined') return;
    try {
      ctx = ctx || new (window.AudioContext || window.webkitAudioContext)();
      if (ctx.state === 'suspended') ctx.resume();
      let buf = cache[name];
      if (!buf) {
        const data = synth(name, ctx.sampleRate > 44100 ? 44100 : ctx.sampleRate);
        buf = ctx.createBuffer(1, data.length, ctx.sampleRate > 44100 ? 44100 : ctx.sampleRate);
        buf.copyToChannel ? buf.copyToChannel(data, 0) : buf.getChannelData(0).set(data);
        cache[name] = buf;
      }
      const src = ctx.createBufferSource(); src.buffer = buf;
      const g = ctx.createGain(); g.gain.value = 0.8;
      src.connect(g).connect(ctx.destination);
      src.start();
      live.add(src); src.onended = () => live.delete(src);
    } catch (e) { /* trình duyệt không hỗ trợ âm thanh */ }
  }
  function stopAll() { live.forEach((s) => { try { s.stop(); } catch (e) { /* bỏ qua */ } }); live.clear(); }
  function setMuted(m) { muted = !!m; if (muted) stopAll(); }

  return { NAMES, synth, wav, wavFor, play, stopAll, setMuted, isMuted: () => muted };
});
