/*
 * CẤU HÌNH PPTX GAME STUDIO
 * --------------------------------------------------------------------------
 * BACKEND:
 *   'auto'     : có SUPABASE_URL + SUPABASE_ANON_KEY -> dùng Supabase, ngược lại kiểm tra mã ngay trên trình duyệt.
 *   'supabase' : gọi thẳng các hàm SQL trên Supabase (chống dùng lại mã trên nhiều máy).
 *   'api'      : gọi server Node.js riêng (server/server.js) theo API_BASE.
 *   'demo'     : kiểm tra mã trên trình duyệt bằng danh sách "dấu vân tay" trong js/code-hashes.js
 *                (file này KHÔNG chứa mã thật; mã thật chỉ nằm trong file quản lý mã trên máy của thầy).
 *
 * SUPABASE_ANON_KEY là khóa "anon / publishable" (công khai được). TUYỆT ĐỐI không dán khóa service_role vào đây.
 */
window.GS_CONFIG = {
  BACKEND: 'auto',
  SUPABASE_URL: '',
  SUPABASE_ANON_KEY: '',
  API_BASE: '',
  DEMO_START_CREDITS: 5,   // số lượt xuất miễn phí khi dùng thử
  CREDITS_PER_CODE: 5,     // mỗi mã kích hoạt cộng 5 lượt
  // Thông tin mua mã (hiện trong hộp "Nhập mã kích hoạt")
  PAYMENT: {
    bank: 'BIDV',
    account: '8821672343',
    holder: 'LE DINH DUNG',
    prices: [['1 mã (5 lượt xuất)', '10.000đ'], ['2 mã (10 lượt xuất)', '18.000đ'], ['3 mã (15 lượt xuất)', '25.000đ']],
    note: 'Nội dung chuyển khoản: MUA MA + số lượng mã + tên Facebook/Zalo của bạn (VD: MUA MA 2 NGUYEN VAN A). Sau khi chuyển khoản, nhắn tin cho thầy kèm ảnh chụp giao dịch để nhận mã.'
  }
};
