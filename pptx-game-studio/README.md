# PPTX Game Studio

Trang web soạn **trò chơi bài giảng** và xuất **PowerPoint (.pptx) có hiệu ứng tương tác thật**:
Ai là triệu phú · Nhìn hình đoán chữ · Mảnh ghép bí mật · Đúng / Sai · Lật ô số.

**Chạy online:** https://lddungkhtn.github.io/GOITEN/pptx-game-studio/

## Điểm nổi bật

- **Ưu tiên hiệu ứng thay vì nhiều slide:** mỗi trò chơi gói gọn trong **1 slide**. Cả bộ chỉ gồm
  `số trò chơi + 2` slide (slide mở đầu dạng menu và slide kết thúc).
- **Trigger thật trong PowerPoint:** bấm vào ô đáp án để chấm đúng/sai, nút 50:50, đồng hồ đếm giờ ⏱,
  bấm ô "?" để mở từng chữ cái, bấm mảnh ghép / ô số → hiện câu hỏi → hiện đáp án → lật mảnh.
- **Menu có liên kết:** slide mở đầu có danh sách trò chơi bấm để nhảy tới; mỗi slide có nút ⌂ quay về menu.
- **Hiệu ứng mượt:** zoom, float-in, fade, wipe, pulse + chuyển slide (fade/push/zoom/cover/wipe). Chỉ dùng hiệu
  ứng chuẩn của PowerPoint 2010 trở lên nên chạy ổn định, file nhẹ.
- **Xem trước giống hệt PowerPoint:** trang web dùng chung một mô tả cảnh với bộ xuất PPTX, có thể bấm thử
  trigger ngay trên trang, và có chế độ **⛶ Trình chiếu** toàn màn hình trên trình duyệt.
- Ghi chú người dẫn (speaker notes) hướng dẫn cách chơi ở từng slide.

## Tính năng mới (bản 2)

- **📥 Nhập từ Word / ảnh:** đọc file .docx (nhận đáp án đúng qua "Đáp án: B", dấu `*`, hoặc phương án được
  gạch chân / tô màu / tô đậm trong Word; nhận cả ảnh trong Word và bảng đáp án cuối bài) hoặc ảnh chụp đề
  (nhận dạng chữ tiếng Việt bằng Tesseract.js, tải khi cần). Nội dung đọc được hiển thị để sửa trước khi chuyển
  thành trò chơi theo lựa chọn (tự động / từng dạng).
- **🔊 Âm thanh** cho mọi trò chơi (tự tổng hợp, không bản quyền) — phát trên web và nhúng vào PowerPoint.
- **🏃 Vượt chướng ngại vật:** 3–8 chướng ngại trên đường đua, trả lời đúng thì nhân vật tiến lên, bấm 🏁 để về đích.
- **🔒 Khóa nội dung khi xuất:** mọi chữ/hình được chuyển thành ảnh, khóa không cho chọn/di chuyển/sửa, file được
  đánh dấu "bản cuối" và có dấu bản quyền mờ — người nhận không sửa câu hỏi hay sao chép bố cục để tự làm lại được,
  trong khi hiệu ứng, trigger, âm thanh vẫn chạy.

## Trò chơi từ Đấu trường Tri thức (bản 3)

Ba trò lấy ý tưởng từ https://lddungkhtn.github.io/ chuyển thành dạng PowerPoint: mỗi trò là 1 slide chứa 1–8 câu
trắc nghiệm (2–4 phương án, câu Đúng/Sai dùng 2 phương án). Bấm vào phương án để chấm; bấm ra ngoài (phím →) sang câu
tiếp; sau câu cuối bấm thêm 1 lần để xem màn kết thúc.

- **🐶 Giải cứu thú cưng:** cún con / mèo con bị nhốt trong phòng, nước đổ từ trần. Đúng: mở thêm 1 van thoát nước.
  Sai: nước dâng thêm 1 mức. Kết thúc: nước rút, "Giải cứu thành công!".
- **🍎 Hứng hoa quả:** các quả mang đáp án treo trên cây, con vật cầm giỏ phía dưới. Đúng: quả rơi vào giỏ, thêm 1 quả
  vào hàng "Đã hứng". Sai: quả thối và bị vứt vào thùng rác.
- **🌍 Bảo vệ Trái Đất:** các thiên thạch mang đáp án lơ lửng trên Trái Đất. Đúng: tàu bắn laser, thiên thạch vỡ,
  sinh vật ngoài hành tinh bị tiêu diệt. Sai: sinh vật trong thiên thạch đúng lao xuống tấn công Trái Đất.

Nhập từ Word / ảnh cũng chuyển được câu trắc nghiệm và Đúng/Sai thẳng vào 3 trò này.

## Cấu trúc

```text
pptx-game-studio/
├── index.html            # Giao diện (không cần build)
├── config.js             # Cấu hình backend (DEMO / Supabase / API Node)
├── css/app.css
├── js/
│   ├── scenes.js         # Mô tả 5 dạng trò chơi thành "cảnh" (hình + hiệu ứng + trigger)
│   ├── player.js         # Trình phát HTML (xem trước / trình chiếu online)
│   ├── pptx-export.js    # Xuất PPTX bằng PptxGenJS + chèn XML animation/trigger/transition
│   ├── backend.js        # Lượt xuất & mã kích hoạt (demo / supabase / api)
│   ├── samples.js        # Bộ câu hỏi mẫu
│   └── app.js            # Logic giao diện soạn bài
├── vendor/               # pptxgenjs 3.12.0, jszip 3.10.1 (đóng gói sẵn, không phụ thuộc CDN)
├── server/
│   ├── schema.sql        # Bảng + stored function (PostgreSQL / Supabase)
│   └── server.js         # API Node.js (tùy chọn)
├── scripts/
│   ├── generate-codes.js # Sinh mã kích hoạt hàng loạt (.json/.csv/.sql)
│   └── manage-codes.js   # Cấp mã cho khách
└── HANDOVER_PLAN.md      # Kế hoạch bàn giao gốc
```

## Lượt xuất & mã kích hoạt

**Đang dùng Supabase** (project `pptx-game-studio`, vùng Singapore): mỗi mã chỉ nạp được đúng 1 lần trên toàn hệ thống,
lượt xuất lưu trên máy chủ. Thêm mã mới: trong `quan-ly-ma.html` bấm "Xuất SQL (Supabase)" rồi chạy file đó trong
Supabase → SQL Editor. Nếu bỏ trống `SUPABASE_URL` trong `config.js`, trang tự quay về chế độ kiểm tra mã trên trình
duyệt mô tả dưới đây.

### Chế độ không máy chủ (dự phòng)

- Người dùng mới được **2 lượt xuất miễn phí**. Mỗi mã kích hoạt cộng **5 lượt** (giá và số tài khoản cấu hình
  trong `config.js` → `PAYMENT`).
- Trang chỉ chứa `js/code-hashes.js` = "dấu vân tay" PBKDF2 của các mã, **không chứa mã thật**. Danh sách mã thật
  và trạng thái bán nằm trong file `quan-ly-ma.html` trên máy của thầy (không đưa lên GitHub).
- Tạo thêm mã: mở `quan-ly-ma.html` → "Tạo thêm mã" → "Tải file code-hashes.js" → chép đè vào `js/` → đẩy lên GitHub.
- Giới hạn của chế độ không máy chủ: lượt xuất lưu trên trình duyệt; một mã chỉ nạp được 1 lần trên mỗi trình duyệt,
  nhưng nếu người mua chia sẻ mã cho người khác thì người đó vẫn nạp được trên máy của họ. Muốn chặn hẳn
  (mỗi mã dùng đúng 1 lần trên toàn hệ thống) cần bật Supabase như hướng dẫn dưới đây (file quản lý có nút xuất SQL).

## Bật hệ thống mã kích hoạt online (Supabase – miễn phí)

1. Tạo tài khoản và project tại https://supabase.com.
2. Vào **SQL Editor → New query**, dán toàn bộ `server/schema.sql` → **Run**.
3. Vào **Project Settings → API**, chép **Project URL** và khóa **anon / publishable**.
4. Mở `config.js`, điền `SUPABASE_URL` và `SUPABASE_ANON_KEY` rồi đẩy lên GitHub.
   (Chỉ dùng khóa *anon*; **không bao giờ** dán khóa `service_role`.)
5. Trang tự chuyển sang chế độ máy chủ ("● Đang kết nối máy chủ").

Người dùng mới được 2 lượt dùng thử (đổi trong `schema.sql`, cột `remaining_exports DEFAULT 2`).

## Tạo và bán mã kích hoạt

```bash
# Tạo 50 mã, mỗi mã cộng 30 lượt, hết hạn cuối 2027
node scripts/generate-codes.js -n 50 -c 30 -p SLIDE30 -e 2027-12-31
# → output-codes/*.sql : dán vào Supabase SQL Editor và Run để nạp mã vào CSDL

# Cấp 2 mã cho khách, xuất file .txt để gửi
node scripts/manage-codes.js -i output-codes/SLIDE30-....json -n 2 -c "Cô Lan - THCS A" --note "CK 20k Momo"
node scripts/manage-codes.js -i output-codes/SLIDE30-....json --stats
```

Thư mục `output-codes/` đã được chặn trong `.gitignore` — **không đưa mã thật lên GitHub**.

## Server Node.js (tùy chọn, thay cho Supabase trực tiếp)

```bash
npm install
DATABASE_URL=postgres://... ALLOW_ORIGIN=https://lddungkhtn.github.io npm start
```
Trong `config.js` đặt `BACKEND: 'api'` và `API_BASE` là địa chỉ server.
Endpoint: `GET /api/balance`, `POST /api/redeem-code`, `POST /api/consume-export`.

## Nguyên tắc tính lượt

- Mỗi lần bấm xuất file hợp lệ = trừ 1 lượt (không phụ thuộc số slide / số trò chơi).
- Không trừ lượt khi chưa chọn mục nào, khi soạn bài, xem trước hay trình chiếu trên web.
- File được tạo xong mới trừ lượt; lỗi tạo file thì không trừ.
- Số dư lượt và danh sách mã chỉ nằm trong CSDL; trình duyệt chỉ giữ mã thiết bị ngẫu nhiên.
- Chế độ DEMO (khi chưa cấu hình Supabase) lưu lượt trên trình duyệt — chỉ để dùng thử.

## Giới hạn đã biết

- Việc tạo file diễn ra trên trình duyệt, nên người rành kỹ thuật vẫn có thể tự chạy mã JavaScript để
  tạo file mà không qua bước trừ lượt. Muốn chặn hoàn toàn cần chuyển bước tạo PPTX lên máy chủ.
- Mã thiết bị lưu trong trình duyệt: xóa dữ liệu trình duyệt sẽ thành "thiết bị mới" (được lượt dùng thử
  mới, nhưng lượt đã nạp gắn với mã thiết bị cũ).
- Font dùng Segoe UI (có sẵn trên Windows). Mở trên máy Mac, PowerPoint sẽ thay bằng font tương đương.
