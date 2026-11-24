/*
 * CẤU HÌNH PPTX GAME STUDIO
 * --------------------------------------------------------------------------
 * BACKEND:
 *   'auto'     : có SUPABASE_URL + SUPABASE_ANON_KEY -> dùng Supabase, ngược lại chạy DEMO.
 *   'supabase' : gọi thẳng các hàm SQL trên Supabase (chạy online được trên GitHub Pages).
 *   'api'      : gọi server Node.js riêng (server/server.js) theo API_BASE.
 *   'demo'     : lượt xuất lưu trên trình duyệt – CHỈ để thử giao diện, không chống gian lận.
 *
 * SUPABASE_ANON_KEY là khóa "anon / publishable" (công khai được). TUYỆT ĐỐI không dán khóa
 * service_role vào đây.
 */
window.GS_CONFIG = {
  BACKEND: 'auto',
  SUPABASE_URL: '',        // ví dụ: 'https://abcdxyz.supabase.co'
  SUPABASE_ANON_KEY: '',   // khóa anon/publishable của project Supabase
  API_BASE: '',            // ví dụ: 'https://ten-server-cua-anh.onrender.com'
  DEMO_START_CREDITS: 2,   // số lượt dùng thử ở chế độ DEMO
  DEMO_CODES: { 'DEMO-15LUOT': 15, 'DEMO-30LUOT': 30 }, // mã thử chỉ dùng ở chế độ DEMO
  CONTACT: 'Liên hệ thầy/cô quản trị để mua mã kích hoạt.'
};
