/*!
 * PPTX Game Studio – Bộ xuất PowerPoint có hiệu ứng thật.
 * 1) Dùng PptxGenJS vẽ hình/chữ/ảnh cho từng slide (mỗi hình được đặt tên = objectName).
 * 2) Mở file vừa tạo bằng JSZip và chèn XML <p:transition> + <p:timing> vào từng slide:
 *    - Chuỗi hiệu ứng tự chạy khi vào slide và khi bấm chuột (main sequence).
 *    - Trigger: bấm vào 1 hình cụ thể -> chạy hiệu ứng (interactive sequence).
 *    - Liên kết nhảy slide (nút ⌂, menu) được gắn trực tiếp vào hình.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PptxExporter = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  const F = 'Segoe UI';

  const ENTR = { fadeIn: 1, zoomIn: 1, floatIn: 1, wipeIn: 1, dropIn: 1 };
  const EXIT = { fadeOut: 1, zoomOut: 1, timerOut: 1 };

  function addShape(slide, sh) {
    const base = { x: sh.x, y: sh.y, w: sh.w, h: sh.h, objectName: sh.name };
    if (sh.kind === 'image') {
      if (!sh.data) return;
      slide.addImage(Object.assign(base, { data: sh.data }));
      return;
    }
    const o = Object.assign(base, {
      fontFace: sh.face || F, fontSize: sh.size || 18, color: sh.color || 'FFFFFF', bold: !!sh.bold,
      align: sh.align || 'center', valign: sh.valign || 'middle', margin: sh.pad || [6, 6, 3, 3], lineSpacingMultiple: 1.0,
      paraSpaceBefore: 0, paraSpaceAfter: 0
    });
    if (sh.kind !== 'text') {
      o.shape = sh.shape || 'rect';
      if (sh.fill) o.fill = { color: sh.fill, transparency: sh.alpha || 0 };
      if (sh.line) o.line = { color: sh.line, width: sh.lineW || 1, dashType: sh.dash ? 'dash' : 'solid' };
      if (sh.radius && o.shape === 'roundRect') o.rectRadius = Math.min(sh.radius, Math.min(sh.w, sh.h) / 2);
      if (sh.shadow) o.shadow = { type: 'outer', blur: 8, offset: 3, angle: 90, color: '000000', opacity: 0.35 };
    }
    if (sh.rot) o.rotate = sh.rot;
    let text = sh.text || '';
    if (sh.runs) {
      text = [];
      sh.runs.forEach((r) => {
        if (!r.t) return;
        const parts = String(r.t).split('\n');
        parts.forEach((p, i) => {
          const last = i === parts.length - 1;
          if (p === '' && last) return;
          text.push({ text: p, options: { color: r.color || o.color, bold: r.bold != null ? r.bold : o.bold,
            fontSize: r.size || o.fontSize, fontFace: r.face || o.fontFace, breakLine: !last } });
        });
      });
      if (!text.length) text = '';
    }
    slide.addText(text, o);
  }

  /* ---------- XML hiệu ứng ---------- */
  // Mỗi hiệu ứng trên cùng 1 hình phải có grpId riêng (0,1,2…) và 1 dòng <p:bldP> tương ứng,
  // nếu trùng PowerPoint sẽ âm thầm bỏ hiệu ứng.
  function TimingBuilder() { this.id = 1; this.grp = new Map(); this.bld = []; this.snds = new Map(); }
  TimingBuilder.prototype.nid = function () { return ++this.id; };
  // âm thanh gắn vào hiệu ứng: <p:audio> trỏ tới file .wav trong ppt/media
  // (cấu trúc lấy đúng theo file PowerPoint tự lưu: <p:subTnLst> nằm sau <p:childTnLst> của hiệu ứng)
  TimingBuilder.prototype.audio = function (name, effectId) {
    if (!this.snds.has(name)) this.snds.set(name, `rIdGSnd${this.snds.size + 1}`);
    const rid = this.snds.get(name);
    return `<p:subTnLst><p:audio><p:cMediaNode><p:cTn display="0" masterRel="sameClick"><p:stCondLst><p:cond evt="begin" delay="0"><p:tn val="${effectId}"/></p:cond></p:stCondLst>` +
      '<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn>' +
      `<p:tgtEl><p:sndTgt r:embed="${rid}" name="${name}.wav"/></p:tgtEl></p:cMediaNode></p:audio></p:subTnLst>`;
  };
  const WIPE_SUB = { left: 8, right: 2, up: 1, down: 4 };

  TimingBuilder.prototype.effect = function (e, spid, nodeType, isPic) {
    const grpId = this.grp.get(spid) || 0;
    this.grp.set(spid, grpId + 1);
    if (!isPic) this.bld.push(`<p:bldP spid="${spid}" grpId="${grpId}" animBg="1"/>`);
    const tgt = `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`;
    const self = this, dur = Math.max(1, Math.round(e.dur));
    const setVis = (val, delay) => `<p:set><p:cBhvr><p:cTn id="${self.nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${val}"/></p:to></p:set>`;
    const fx = (trans, filter, d) => `<p:animEffect transition="${trans}" filter="${filter}"><p:cBhvr><p:cTn id="${self.nid()}" dur="${d}"/>${tgt}</p:cBhvr></p:animEffect>`;
    const val = (v) => (typeof v === 'number' ? `<p:fltVal val="${v}"/>` : `<p:strVal val="${v}"/>`);
    const prop = (attr, from, to, d) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${self.nid()}" dur="${d}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val>${val(from)}</p:val></p:tav><p:tav tm="100000"><p:val>${val(to)}</p:val></p:tav></p:tavLst></p:anim>`;

    let preset, cls, sub = 0, extra = '';
    const oid = this.nid();
    let inner = '';
    switch (e.type) {
      case 'fadeIn': preset = 10; cls = 'entr'; inner = setVis('visible', 0) + fx('in', 'fade', dur); break;
      case 'zoomIn': preset = 53; cls = 'entr'; sub = 16;
        inner = setVis('visible', 0) + prop('ppt_w', 0, '#ppt_w', dur) + prop('ppt_h', 0, '#ppt_h', dur) + fx('in', 'fade', dur); extra = ' decel="50000"'; break;
      case 'floatIn': preset = 42; cls = 'entr';
        inner = setVis('visible', 0) + fx('in', 'fade', dur) + prop('ppt_x', '#ppt_x', '#ppt_x', dur) + prop('ppt_y', '#ppt_y+.1', '#ppt_y', dur); extra = ' decel="100000"'; break;
      case 'dropIn': preset = 42; cls = 'entr'; // rơi từ trên xuống
        inner = setVis('visible', 0) + fx('in', 'fade', Math.round(dur / 2)) + prop('ppt_x', '#ppt_x', '#ppt_x', dur) + prop('ppt_y', '#ppt_y-.25', '#ppt_y', dur); extra = ' accel="60000"'; break;
      case 'wipeIn': { const d = WIPE_SUB[e.dir] ? e.dir : 'left'; preset = 22; cls = 'entr'; sub = WIPE_SUB[d]; inner = setVis('visible', 0) + fx('in', `wipe(${d})`, dur); break; }
      case 'fadeOut': preset = 10; cls = 'exit'; inner = fx('out', 'fade', dur) + setVis('hidden', dur - 1); break;
      case 'zoomOut': preset = 53; cls = 'exit'; sub = 32; extra = ' accel="50000"';
        inner = prop('ppt_w', '#ppt_w', 0, dur) + prop('ppt_h', '#ppt_h', 0, dur) + fx('out', 'fade', dur) + setVis('hidden', dur - 1); break;
      case 'timerOut': preset = 22; cls = 'exit'; sub = 2; inner = fx('out', 'wipe(right)', dur) + setVis('hidden', dur - 1); break;
      case 'pulse': preset = 6; cls = 'emph';
        inner = `<p:animScale><p:cBhvr><p:cTn id="${self.nid()}" dur="${Math.round(dur / 2)}" autoRev="1" fill="hold"/>${tgt}</p:cBhvr><p:by x="113000" y="113000"/></p:animScale>`; break;
      default: return '';
    }
    const snd = e.snd ? this.audio(e.snd, oid) : '';
    return `<p:par><p:cTn id="${oid}" presetID="${preset}" presetClass="${cls}" presetSubtype="${sub}"${extra} fill="hold" grpId="${grpId}" nodeType="${nodeType}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst>${snd}</p:cTn></p:par>`;
  };

  /** Các hiệu ứng trong 1 lần bấm, nhóm theo thời điểm bắt đầu (delay). */
  TimingBuilder.prototype.group = function (effects, firstType, ids, pics) {
    const list = effects.filter((e) => ids[e.target] != null);
    const byDelay = new Map();
    list.slice().sort((a, b) => a.delay - b.delay).forEach((e) => {
      const d = Math.max(0, Math.round(e.delay));
      if (!byDelay.has(d)) byDelay.set(d, []);
      byDelay.get(d).push(e);
    });
    let xml = '', first = true;
    for (const [d, arr] of byDelay) {
      xml += `<p:par><p:cTn id="${this.nid()}" fill="hold"><p:stCondLst><p:cond delay="${d}"/></p:stCondLst><p:childTnLst>`;
      arr.forEach((e, i) => {
        const nt = i === 0 ? (first ? firstType : 'afterEffect') : 'withEffect';
        xml += this.effect(e, ids[e.target], nt, pics.has(e.target));
      });
      xml += '</p:childTnLst></p:cTn></p:par>';
      first = false;
    }
    return xml;
  };

  function timingXml(sd, ids, pics) {
    const tb = new TimingBuilder();
    const main = sd.main.filter((g) => g.effects.some((e) => ids[e.target] != null));
    const trig = sd.triggers.filter((t) => ids[t.on] != null && t.effects.some((e) => ids[e.target] != null));
    if (!main.length && !trig.length) return { xml: '', snds: tb.snds };
    let body = '';
    if (main.length) {
      const mainId = tb.nid(); // = 2
      let inner = '';
      main.forEach((g, gi) => {
        const auto = g.auto && gi === 0;
        const cond = auto ? `<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="${mainId}"/></p:cond>` : '<p:cond delay="indefinite"/>';
        inner += `<p:par><p:cTn id="${tb.nid()}" fill="hold"><p:stCondLst>${cond}</p:stCondLst><p:childTnLst>${tb.group(g.effects, auto ? 'afterEffect' : 'clickEffect', ids, pics)}</p:childTnLst></p:cTn></p:par>`;
      });
      body += `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${mainId}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${inner}</p:childTnLst></p:cTn>` +
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>';
    }
    trig.forEach((t) => {
      const sp = ids[t.on];
      const clk = `<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl></p:cond>`;
      const seqId = tb.nid();
      const parId = tb.nid();
      const g = tb.group(t.effects, 'clickEffect', ids, pics);
      body += `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${seqId}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">` +
        `<p:stCondLst>${clk}</p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>` +
        `<p:par><p:cTn id="${parId}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${g}</p:childTnLst></p:cTn></p:par>` +
        `</p:childTnLst></p:cTn><p:nextCondLst>${clk}</p:nextCondLst></p:seq>`;
    });
    const bld = tb.bld.join('');
    return { xml: `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${body}</p:childTnLst></p:cTn></p:par></p:tnLst>${bld ? `<p:bldLst>${bld}</p:bldLst>` : ''}</p:timing>`, snds: tb.snds };
  }

  const PIC_LOCK = '<a:picLocks noGrp="1" noSelect="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noCrop="1"/>';
  const SP_LOCK = '<a:spLocks noGrp="1" noSelect="1" noRot="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noTextEdit="1"/>';

  /** Đánh dấu "bản cuối" (Mark as Final) để PowerPoint mở ở chế độ chỉ đọc */
  async function markFinal(zip) {
    const core = zip.file('docProps/core.xml');
    if (core) {
      let x = await core.async('string');
      if (!/contentStatus/.test(x)) x = x.replace('</cp:coreProperties>', '<cp:contentStatus>Final</cp:contentStatus></cp:coreProperties>');
      zip.file('docProps/core.xml', x);
    }
    zip.file('docProps/custom.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="_MarkAsFinal"><vt:bool>true</vt:bool></property>' +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="GameStudioProtected"><vt:lpwstr>PPTX Game Studio</vt:lpwstr></property></Properties>');
    let ct = await zip.file('[Content_Types].xml').async('string');
    if (!/custom\.xml/.test(ct)) ct = ct.replace('</Types>', '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>');
    zip.file('[Content_Types].xml', ct);
    let rels = await zip.file('_rels/.rels').async('string');
    if (!/custom-properties/.test(rels)) rels = rels.replace('</Relationships>', '<Relationship Id="rIdGSCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>');
    zip.file('_rels/.rels', rels);
  }

  const TRANS = {
    fade: '<p:fade/>', push: '<p:push dir="u"/>', zoom: '<p:zoom/>', cover: '<p:cover dir="l"/>', wipe: '<p:wipe dir="r"/>'
  };
  const transitionXml = (t) => `<p:transition spd="med">${TRANS[t] || TRANS.fade}</p:transition>`;

  /**
   * Xuất deck -> Blob/ArrayBuffer .pptx
   * @param deck kết quả GameScenes.buildDeck()
   * @param libs { PptxGenJS, JSZip }
   */
  async function exportDeck(deck, libs, opts) {
    opts = opts || {};
    const { PptxGenJS, JSZip, Sounds } = libs;
    const protect = !!(opts.protect && opts.rasterize);
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.title = deck.title; pptx.subject = 'Trò chơi bài giảng'; pptx.author = 'PPTX Game Studio'; pptx.company = 'PPTX Game Studio';
    pptx.theme = { headFontFace: F, bodyFontFace: F };
    for (let i = 0; i < deck.slides.length; i++) {
      const sd = deck.slides[i];
      const s = pptx.addSlide();
      if (protect) {
        // Chế độ khóa: mọi hình được "in" thành ảnh, phần nền tĩnh gộp thành 1 ảnh nền.
        if (opts.onProgress) opts.onProgress(i, deck.slides.length);
        const r = await opts.rasterize(sd);
        s.background = { data: r.bg };
        sd.shapes.forEach((sh) => {
          if (r.merged.has(sh.name)) return;
          const im = r.images[sh.name];
          if (im) s.addImage({ data: im.data, x: im.x, y: im.y, w: im.w, h: im.h, objectName: sh.name, altText: ' ' });
        });
      } else {
        s.background = { color: sd.bg };
        sd.shapes.forEach((sh) => addShape(s, sh));
      }
      if (sd.notes) s.addNotes(sd.notes);
    }
    const buf = await pptx.write({ outputType: 'arraybuffer' });
    const zip = await JSZip.loadAsync(buf);
    const media = new Set();
    for (let i = 0; i < deck.slides.length; i++) {
      const sd = deck.slides[i];
      const path = `ppt/slides/slide${i + 1}.xml`;
      const relPath = `ppt/slides/_rels/slide${i + 1}.xml.rels`;
      let xml = await zip.file(path).async('string');
      let rels = await zip.file(relPath).async('string');
      // Đổi mọi ảnh (<p:pic>) thành hình chữ nhật tô nền bằng ảnh: trông y hệt, nhưng PowerPoint xử lý
      // như shape nên trigger vẫn chạy khi ảnh vừa có hiệu ứng ở chuỗi chính vừa có trong trigger.
      xml = xml.replace(/<p:pic>([\s\S]*?)<\/p:pic>/g, (m, body) => {
        const cNvPr = (body.match(/<p:cNvPr[^>]*\/>|<p:cNvPr[\s\S]*?<\/p:cNvPr>/) || [''])[0];
        const blip = (body.match(/<a:blip r:embed="([^"]+)"/) || [])[1];
        const xfrm = (body.match(/<a:xfrm[\s\S]*?<\/a:xfrm>/) || [''])[0];
        if (!cNvPr || !blip || !xfrm) return m;
        return `<p:sp><p:nvSpPr>${cNvPr}<p:cNvSpPr>${protect ? SP_LOCK : ''}</p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
          `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${blip}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`;
      });
      const ids = {}, pics = new Set();
      xml.replace(/<p:cNvPr id="(\d+)" name="([^"]*)"/g, (m, id, name) => { ids[name] = +id; return m; });
      xml.replace(/<p:pic>\s*<p:nvPicPr>\s*<p:cNvPr id="\d+" name="([^"]*)"/g, (m, name) => { pics.add(name); return m; });
      // Liên kết nhảy slide gắn vào cả hình
      let ln = 0;
      sd.shapes.forEach((sh) => {
        if (!sh.link || ids[sh.name] == null || sh.link > deck.slides.length) return;
        const rid = `rIdGS${++ln}`;
        rels = rels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide${sh.link}.xml"/></Relationships>`);
        const tip = sh.tip ? ` tooltip="${sh.tip.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')}"` : '';
        const re = new RegExp(`(<p:cNvPr id="${ids[sh.name]}" name="${sh.name}"[^>]*?)(/?)>`);
        xml = xml.replace(re, (m, head, selfClose) => selfClose
          ? `${head}><a:hlinkClick r:id="${rid}" action="ppaction://hlinksldjump"${tip}/></p:cNvPr>`
          : `${head}><a:hlinkClick r:id="${rid}" action="ppaction://hlinksldjump"${tip}/>`);
      });
      const tm = timingXml(sd, ids, pics);
      // Âm thanh: thêm file .wav (mỗi loại 1 lần cho cả bài) và quan hệ tới slide
      if (Sounds && opts.sound !== false) {
        tm.snds.forEach((rid, name) => {
          const file = `gs-${name}.wav`;
          if (!media.has(file)) { zip.file('ppt/media/' + file, Sounds.wavFor(name)); media.add(file); }
          rels = rels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="../media/${file}"/></Relationships>`);
        });
      } else tm.xml = tm.xml.replace(/<p:subTnLst>.*?<\/p:subTnLst>/g, '');
      if (protect) xml = xml.replace(/<a:picLocks[^>]*\/>/g, PIC_LOCK).replace(/<p:cNvSpPr(?: txBox="1")?\/>/g, `<p:cNvSpPr>${SP_LOCK}</p:cNvSpPr>`).replace(/<p:cNvSpPr><\/p:cNvSpPr>/g, `<p:cNvSpPr>${SP_LOCK}</p:cNvSpPr>`);
      xml = xml.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + transitionXml(sd.transition) + tm.xml);
      zip.file(path, xml);
      zip.file(relPath, rels);
    }
    if (media.size) {
      let ct = await zip.file('[Content_Types].xml').async('string');
      if (!/Extension="wav"/i.test(ct)) ct = ct.replace('<Default ', '<Default Extension="wav" ContentType="audio/wav"/><Default ');
      zip.file('[Content_Types].xml', ct);
    }
    if (protect) await markFinal(zip);
    return zip.generateAsync({ type: opts.outputType || 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 },
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  }

  return { exportDeck, timingXml, ENTR, EXIT };
});
