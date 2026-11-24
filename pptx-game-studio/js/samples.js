/*! PPTX Game Studio – Bộ câu hỏi mẫu (ảnh minh họa vẽ bằng SVG, chuyển sang PNG khi tải). */
(function (root) {
  'use strict';
  const RAINBOW = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 590 545">
  <defs><linearGradient id="sky" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7dd3fc"/><stop offset="1" stop-color="#e0f2fe"/></linearGradient></defs>
  <rect width="590" height="545" fill="url(#sky)"/>
  <g fill="none" stroke-width="22">
    <circle cx="295" cy="470" r="300" stroke="#ef4444"/><circle cx="295" cy="470" r="278" stroke="#f97316"/>
    <circle cx="295" cy="470" r="256" stroke="#facc15"/><circle cx="295" cy="470" r="234" stroke="#22c55e"/>
    <circle cx="295" cy="470" r="212" stroke="#3b82f6"/><circle cx="295" cy="470" r="190" stroke="#4f46e5"/>
    <circle cx="295" cy="470" r="168" stroke="#a855f7"/>
  </g>
  <g fill="#fff"><circle cx="70" cy="300" r="42"/><circle cx="120" cy="285" r="52"/><circle cx="170" cy="305" r="38"/><rect x="40" y="300" width="160" height="40" rx="20"/>
  <circle cx="440" cy="250" r="40"/><circle cx="490" cy="232" r="52"/><circle cx="540" cy="255" r="36"/><rect x="410" y="250" width="160" height="40" rx="20"/></g>
  <circle cx="520" cy="80" r="44" fill="#fde047"/>
  <path d="M0 420 Q150 380 300 420 T590 410 V545 H0z" fill="#4ade80"/><path d="M0 470 Q160 440 320 470 T590 465 V545 H0z" fill="#16a34a"/>
  <g fill="#38bdf8" opacity=".8"><path d="M90 150 l-8 22 a9 9 0 1 0 16 0z"/><path d="M200 110 l-8 22 a9 9 0 1 0 16 0z"/><path d="M380 130 l-8 22 a9 9 0 1 0 16 0z"/></g>
</svg>`;

  const SUN = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 760 570">
  <defs><radialGradient id="g" cx=".5" cy=".55" r=".7"><stop offset="0" stop-color="#fff7cc"/><stop offset=".45" stop-color="#fdba74"/><stop offset="1" stop-color="#7c2d12"/></radialGradient>
  <radialGradient id="s" cx=".5" cy=".5" r=".5"><stop offset="0" stop-color="#fffbeb"/><stop offset=".55" stop-color="#fde047"/><stop offset="1" stop-color="#f59e0b"/></radialGradient></defs>
  <rect width="760" height="570" fill="url(#g)"/>
  <g stroke="#fde68a" stroke-width="14" stroke-linecap="round" opacity=".9">
    <line x1="380" y1="40" x2="380" y2="95"/><line x1="380" y1="395" x2="380" y2="440"/><line x1="180" y1="240" x2="235" y2="240"/><line x1="525" y1="240" x2="580" y2="240"/>
    <line x1="240" y1="100" x2="278" y2="138"/><line x1="482" y1="342" x2="520" y2="380"/><line x1="520" y1="100" x2="482" y2="138"/><line x1="278" y1="342" x2="240" y2="380"/>
  </g>
  <circle cx="380" cy="240" r="130" fill="url(#s)"/>
  <path d="M0 410 C120 380 220 430 380 405 S640 380 760 410 V570 H0z" fill="#0e7490"/>
  <path d="M0 450 C140 430 260 470 400 450 S650 430 760 455 V570 H0z" fill="#155e75"/>
  <g stroke="#fde68a" stroke-width="6" stroke-linecap="round" opacity=".7"><line x1="300" y1="430" x2="460" y2="430"/><line x1="330" y1="465" x2="430" y2="465"/><line x1="350" y1="500" x2="410" y2="500"/></g>
  <g fill="none" stroke="#1f2937" stroke-width="5" stroke-linecap="round"><path d="M120 150 q15 -14 30 0 q15 -14 30 0"/><path d="M600 120 q12 -11 24 0 q12 -11 24 0"/></g>
</svg>`;

  function svgToPng(svg, w, h) {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const c = document.createElement('canvas'); c.width = w; c.height = h;
        c.getContext('2d').drawImage(img, 0, 0, w, h);
        try { resolve(c.toDataURL('image/png')); } catch (e) { resolve(''); }
      };
      img.onerror = () => resolve('');
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    });
  }

  async function sampleProject() {
    const N = root.GameScenes.newItem;
    const [rainbow, sun] = await Promise.all([svgToPng(RAINBOW, 1180, 1090), svgToPng(SUN, 1520, 1140)]);
    const a = N('millionaire');
    Object.assign(a, { question: 'Hành tinh nào gần Mặt Trời nhất trong Hệ Mặt Trời?', answers: ['Sao Kim', 'Sao Thủy', 'Trái Đất', 'Sao Hỏa'], correct: 1,
      explanation: 'Sao Thủy cách Mặt Trời trung bình khoảng 58 triệu km – gần nhất trong 8 hành tinh.', timer: 30 });
    const b = N('picture');
    Object.assign(b, { question: 'Hiện tượng tự nhiên trong hình là gì?', image: rainbow, hint: 'Xuất hiện trên bầu trời khi vừa mưa vừa nắng, gồm 7 màu sắc.', answer: 'Cầu vồng', timer: 20 });
    const c = N('puzzle');
    Object.assign(c, { image: sun, pieceCount: 4, keyword: 'Mặt Trời', pieces: [
      { question: 'Trái Đất tự quay quanh trục một vòng mất khoảng bao lâu?', answer: 'Khoảng 24 giờ (1 ngày)' },
      { question: 'Hiện tượng ngày và đêm luân phiên nhau là do đâu?', answer: 'Trái Đất tự quay quanh trục' },
      { question: 'Cây xanh dùng năng lượng ánh sáng để thực hiện quá trình gì?', answer: 'Quang hợp' },
      { question: 'Một năm Trái Đất quay được mấy vòng quanh ngôi sao gần nhất?', answer: '1 vòng (khoảng 365 ngày)' }] });
    const d = N('truefalse');
    Object.assign(d, { statement: 'Âm thanh có thể truyền được trong chân không.', isTrue: false,
      explanation: 'Âm thanh cần môi trường vật chất (rắn, lỏng, khí) để truyền đi. Trong chân không, âm không truyền được.', timer: 15 });
    const e = N('tiles');
    Object.assign(e, { title: 'Khám phá về nước – chọn một ô số bất kỳ!', cells: [
      { question: 'Công thức hóa học của nước là gì?', answer: 'H₂O' },
      { question: 'Ở áp suất tiêu chuẩn, nước sôi ở bao nhiêu độ C?', answer: '100 °C' },
      { question: 'Nước đá bắt đầu nóng chảy ở nhiệt độ nào?', answer: '0 °C' },
      { question: 'Quá trình nước chuyển từ thể lỏng sang thể hơi gọi là gì?', answer: 'Sự bay hơi' },
      { question: 'Hơi nước gặp lạnh chuyển thành nước lỏng gọi là gì?', answer: 'Sự ngưng tụ' },
      { question: 'Nước tồn tại ở những thể nào?', answer: 'Rắn, lỏng, khí' }] });
    const f = N('obstacle');
    Object.assign(f, { title: 'Chinh phục Hệ Mặt Trời – vượt 5 chướng ngại để về đích!', timer: 20, steps: [
      { question: 'Hệ Mặt Trời có bao nhiêu hành tinh?', answer: '8 hành tinh' },
      { question: 'Hành tinh lớn nhất trong Hệ Mặt Trời là gì?', answer: 'Sao Mộc' },
      { question: 'Hành tinh nào được gọi là “Hành tinh Đỏ”?', answer: 'Sao Hỏa' },
      { question: 'Vệ tinh tự nhiên của Trái Đất là gì?', answer: 'Mặt Trăng' },
      { question: 'Hành tinh nào có vành đai rõ và đẹp nhất?', answer: 'Sao Thổ' }] });
    return { title: 'Ôn tập Khoa học tự nhiên', subtitle: '6 trò chơi tương tác • Bấm vào tên trò chơi để bắt đầu', items: [a, b, c, d, e, f] };
  }

  root.GSSamples = { sampleProject };
})(typeof self !== 'undefined' ? self : this);
