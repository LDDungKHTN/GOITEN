# KẾ HOẠCH BÀN GIAO DỰ ÁN: PPTX GAME STUDIO & HỆ THỐNG MÃ KÍCH HOẠT

Tài liệu này tổng hợp toàn bộ kiến trúc, yêu cầu tính năng, thiết kế cơ sở dữ liệu và lộ trình triển khai chi tiết cho công cụ tạo trò chơi bài giảng PowerPoint (`.pptx`) tích hợp cơ chế nạp lượt xuất qua mã kích hoạt (Activation Code).

---

## 1. TỔNG QUAN DỰ ÁN

### 1.1. Mục tiêu
- **Dành cho Giáo viên:** Cung cấp ứng dụng web đơn giản, trực quan giúp soạn bộ câu hỏi và trò chơi bài giảng (dạng *Ai là triệu phú*, *Nhìn hình đoán chữ*, *Mảnh ghép*, *Đúng / Sai*, *Lật ô số*), xem trước sinh động và xuất thẳng ra file PowerPoint (`.pptx`) chuẩn tỷ lệ 16:9 bằng `PptxGenJS`.
- **Dành cho Chủ sở hữu:** Quản lý lượt xuất `.pptx` an toàn, chống gian lận lách lượt ở client, hỗ trợ mô hình kinh doanh bán lượt theo gói hoặc theo mã kích hoạt dùng 1 lần (Activation Code / Voucher).

### 1.2. Nguyên tắc vận hành cốt lõi
1. **Một lần bấm xuất file `.pptx` hợp lệ = trừ 1 lượt** (không tính lượt theo số lượng slide hay số lượng trò chơi trong bộ).
2. **Không trừ lượt khi chưa chọn mục nào** hoặc khi người dùng chỉ thực hiện thao tác soạn bài / xem trước (Preview).
3. **Bảo mật tuyệt đối lượt xuất & mã kích hoạt:** Không lưu số dư lượt xuất thật hoặc danh sách mã trong `localStorage` hay mã nguồn JavaScript client. Việc xác thực mã và trừ lượt xuất phải do Backend API / Database xử lý nguyên tử (atomic transaction).
4. **Hỗ trợ chế độ ngoại tuyến / Demo:** Cho phép bật/tắt `DEMO_MODE` để kiểm thử giao diện frontend trước khi nối Backend.

---

## 2. KIẾN TRÚC & CẤU TRÚC THƯ MỤC DỰ ÁN

```text
pptx-game-studio/
├── index.html               # Frontend duy nhất (HTML5, Tailwind CDN, PptxGenJS 3.12.0 bundle, Lucide Icons)
├── server/
│   ├── server.js            # Node.js + Express API Server (/api/redeem-code, /api/consume-export)
│   └── schema.sql           # File khởi tạo Database PostgreSQL / Supabase (Tables + Stored Functions)
├── scripts/
│   ├── generate-codes.js    # Script CLI sinh hàng loạt mã kích hoạt ngẫu nhiên (xuất JSON, CSV, SQL)
│   └── manage-codes.js      # Script CLI lọc mã trống, đánh dấu mã 'issued' & xuất danh sách cho khách
├── output-codes/            # Thư mục chứa các lô mã kích hoạt đã tạo (Chú ý: Không commit file mã thật)
├── HANDOVER_PLAN.md         # File kế hoạch bàn giao chi tiết (Tài liệu này)
└── package.json             # Cấu hình dự án & phụ thuộc Node.js (express, cors, pg)
```

---

## 3. THIẾT KẾ PHẠM VI TÍNH NĂNG FRONTEND (`index.html`)

### 3.1. Giao diện & Trải nghiệm Người dùng
- **Thanh Header Top:**
  - Tiêu đề ứng dụng + Huy hiệu người dùng.
  - Nút **Mẫu thử** (tải bộ câu hỏi mẫu 5 game).
  - Nút **Mở nháp** (Import JSON) & **Lưu nháp** (Export JSON).
  - Nút chính **Xuất PowerPoint (.pptx)**.
- **Bảng Quản lý Danh sách (Bên trái):**
  - Huy hiệu hiển thị **Số lượt xuất còn lại** (Tự đổi màu: Xanh = >1 lượt, Vàng = 1 lượt, Đỏ = 0 lượt).
  - Nút **Nhập mã kích hoạt** (Mở Modal nạp lượt).
  - Thanh công cụ chọn nhanh: **Chọn tất cả / Bỏ chọn tất cả** và bộ đếm `Đã chọn: X/Y mục xuất PPTX`.
  - Checkbox chọn/bỏ chọn từng mục (`includedInExport: true/false`).
  - Nút sắp xếp thứ tự (lên/xuống), nhân bản và xóa mục.
  - Dropdown **"Thêm trò chơi / Câu hỏi mới"** với 5 dạng game.
  - Form chỉnh sửa câu hỏi linh hoạt theo từng dạng game.
- **Bảng Xem trước Slide (Bên phải):**
  - Khung mô phỏng màn hình chuẩn Widescreen `16:9` (`aspect-ratio: 16/9`).
  - Thanh chuyển đổi Slide Step Tabs (ví dụ: *Slide 1: Câu hỏi* | *Slide 2: Đáp án*).
  - Render HTML/CSS động đúng theo dữ liệu đang chỉnh sửa.

### 3.2. Quy tắc sinh Slide PowerPoint Widescreen 16:9 (`PptxGenJS`)

| Dạng trò chơi | Dữ liệu đầu vào | Quy tắc Slide xuất ra PPTX |
| :--- | :--- | :--- |
| **1. Ai là triệu phú** | Câu hỏi, 4 đáp án (A,B,C,D), vị trí đáp án đúng, giải thích | **Slide 1:** Khung câu hỏi + 4 ô đáp án.<br>**Slide 2:** Tô màu xanh nổi bật ô đáp án đúng + hiển thị ghi chú giải thích ở đáy slide. |
| **2. Nhìn hình đoán chữ** | Ảnh minh họa (Base64/Link), gợi ý, từ/cụm từ đáp án | **Slide 1:** Ảnh bên trái + ô gợi ý bên phải.<br>**Slide 2:** Ảnh + khung đáp án nổi bật chữ lớn màu xanh ngọc. |
| **3. Mảnh ghép** | Ảnh nền bí mật, số mảnh (4 hoặc 6), gợi ý từng mảnh, từ khóa chính | **Chuỗi Slide lật mở:** Slide 1 che toàn bộ ảnh bằng các ô mảnh ghép -> Các slide tiếp theo ẩn dần từng ô che -> Slide cuối hiện trọn ảnh + banner từ khóa chính. |
| **4. Đúng / Sai** | Nhận định/Phát biểu, kết luận (ĐÚNG hay SAI), giải thích | **Slide 1:** Nhận định + 2 thẻ ĐÚNG/SAI chờ chọn.<br>**Slide 2:** Tô sáng thẻ kết luận đúng, làm mờ thẻ sai, kèm giải thích. |
| **5. Lật ô số** | Tiêu đề bảng, 4–8 ô số (câu hỏi + đáp án từng ô) | **Slide 1:** Bảng tổng quan các ô số.<br>**Cặp Slide từng ô:** Slide câu hỏi ô `k` -> Slide đáp án ô `k`.<br>**Slide cuối:** Bảng tổng kết đáp án tất cả các ô số. |

---

## 4. CHUYÊN BIỆT BACKEND & DATABASE

### 4.1. Cấu trúc Bảng CSDL (PostgreSQL / Supabase)

```sql
-- 1. Bảng người dùng / phiên thiết bị
CREATE TABLE IF NOT EXISTS users (
  id TEXT PRIMARY KEY,
  remaining_exports INTEGER NOT NULL DEFAULT 2 CHECK (remaining_exports >= 0),
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 2. Bảng mã kích hoạt
CREATE TABLE IF NOT EXISTS activation_codes (
  code TEXT PRIMARY KEY,
  credits INTEGER NOT NULL CHECK (credits > 0),
  status TEXT NOT NULL DEFAULT 'unused' CHECK (status IN ('unused', 'used', 'disabled')),
  used_by TEXT REFERENCES users(id),
  used_at TIMESTAMPTZ,
  expires_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 3. Bảng nhật ký xuất file PPTX
CREATE TABLE IF NOT EXISTS export_logs (
  id BIGSERIAL PRIMARY KEY,
  user_id TEXT NOT NULL REFERENCES users(id),
  selected_count INTEGER NOT NULL DEFAULT 0,
  estimated_slides INTEGER NOT NULL DEFAULT 0,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
```

### 4.2. Các Hàm Xử lý Nguyên tử (Stored Functions)

1. **`redeem_activation_code(p_code TEXT, p_user_id TEXT)`**
   - Khóa bản ghi mã kích hoạt (`FOR UPDATE`) để chống đua tranh (race condition).
   - Kiểm tra mã có tồn tại, `status = 'unused'`, và `expires_at` còn hạn.
   - Đổi `status` thành `'used'`, cập nhật `used_by` và `used_at`.
   - Cộng `credits` vào `remaining_exports` của `users`.
   - Trả về: `(ok BOOLEAN, message TEXT, added_credits INT, remaining_exports INT)`.

2. **`consume_export_credit(p_user_id TEXT, p_selected_count INT, p_estimated_slides INT)`**
   - Chỉ cập nhật khi `remaining_exports >= 1` (`remaining_exports = remaining_exports - 1`).
   - Ghi 1 dòng nhật ký vào `export_logs`.
   - Trả về: `(allowed BOOLEAN, message TEXT, remaining_exports INT)`.

### 4.3. Định nghĩa REST API Endpoints

- **`POST /api/redeem-code`**
  - Input: `{ code: "SLIDE30-7K9M2X", userId: "dev_xyz123" }`
  - Output thành công (200): `{ success: true, message: "...", addedCredits: 30, remainingExports: 32 }`
  - Output thất bại (400): `{ success: false, message: "Mã này đã được sử dụng.", remainingExports: 2 }`
- **`POST /api/consume-export`**
  - Input: `{ userId: "dev_xyz123", selectedCount: 4, estimatedSlides: 12 }`
  - Output được phép (200): `{ allowed: true, message: "...", remainingExports: 1 }`
  - Output hết lượt (402): `{ allowed: false, message: "Bạn đã hết lượt xuất PPTX.", remainingExports: 0 }`

---

## 5. BỘ CÔNG CỤ UTILITY SCRIPTS CLI (NODE.JS)

### 5.1. Script `scripts/generate-codes.js`
- **Công dụng:** Tự động tạo hàng loạt mã kích hoạt không trùng lặp, loại bỏ ký tự dễ nhầm lẫn (`0`, `O`, `1`, `I`).
- **Tham số CLI:**
  - `--count`, `-n`: Số lượng mã cần tạo (Mặc định: 50)
  - `--credits`, `-c`: Số lượt xuất mỗi mã cộng (Mặc định: 30)
  - `--prefix`, `-p`: Tiền tố gói (Mặc định: `SLIDE30`)
  - `--length`, `-l`: Độ dài chuỗi ngẫu nhiên (Mặc định: 6)
  - `--expires`, `-e`: Ngày hết hạn ISO (Mặc định: `null`)
- **Đầu ra:** Xuất đồng thời 3 file trong thư mục `./output-codes/`: `.json`, `.csv`, và `.sql` (`INSERT ... ON CONFLICT DO NOTHING`).

### 5.2. Script `scripts/manage-codes.js`
- **Công dụng:** Quản lý kho mã, lọc các mã chưa phát/chưa dùng, trích xuất đúng số lượng mã giao cho khách và cập nhật file nguồn.
- **Tham số CLI:**
  - `--input`, `-i`: Đường dẫn file mã nguồn (`.json` hoặc `.csv`)
  - `--count`, `-n`: Số lượng mã cần cấp (Mặc định: 1)
  - `--customer`, `-c`: Tên/Thông tin khách hàng nhận mã
  - `--note`: Ghi chú đơn hàng (VD: "CK 20k qua Momo")
  - `--out`, `-o`: Tên file xuất riêng cho khách
- **Đầu ra:** Đánh dấu `status = 'issued'` kèm `issued_to` và `issued_at` vào file nguồn, xuất 1 file chứa riêng mã cấp cho khách.

---

## 6. LỘ TRÌNH THỰC THI CHI TIẾT CHO CLAUDE CODE

Khi thực hiện dự án này, hãy tuân thủ lộ trình 4 bước sau:

### Bước 1: Kiểm tra & Hoàn thiện Frontend Tĩnh (`index.html`)
1. Đảm bảo file `index.html` chạy tốt trên trình duyệt mà không cần build tool.
2. Kiểm tra tính năng tích chọn/bỏ chọn câu hỏi, bộ đếm câu hỏi được chọn xuất PPTX.
3. Kiểm tra logic tạo file `.pptx` của cả 5 dạng game bằng `PptxGenJS`.
4. Đảm bảo UI modal nạp lượt và huy hiệu số lượt hoạt động mượt ở chế độ `CONFIG.DEMO_MODE = true`.

### Bước 2: Khởi tạo Database & Backend API Server
1. Tạo file `server/schema.sql` chứa toàn bộ câu lệnh khởi tạo bảng và 2 stored functions (`redeem_activation_code`, `consume_export_credit`).
2. Khởi tạo `package.json` với các phụ thuộc: `express`, `cors`, `pg`.
3. Viết `server/server.js` triển khai 2 endpoint `/api/redeem-code` và `/api/consume-export`.
4. Đổi `CONFIG.DEMO_MODE = false` trong `index.html` và trỏ `CONFIG.API_BASE` về API server.

### Bước 3: Triển khai Bộ Script CLI Quản lý Mã Kích Hoạt
1. Viết script `scripts/generate-codes.js` có hỗ trợ nhận tham số dòng lệnh.
2. Viết script `scripts/manage-codes.js` để lọc mã trống và xuất mã cho khách.
3. Tạo file `.gitignore` để chắn không cho commit các file mã trong `output-codes/` lên Git.

### Bước 4: Kiểm thử Tích hợp End-to-End (E2E)
1. Chạy script tạo 10 mã mẫu gói 15 lượt.
2. Thực hiện nạp mã thành công trên giao diện web -> Số lượt tăng tương ứng.
3. Thử nhập mã đã dùng -> Hệ thống báo lỗi "Mã này đã được sử dụng".
4. Bấm xuất file PPTX -> Hệ thống trừ 1 lượt. Bấm cho đến khi số lượt về 0 -> Hệ thống tự chặn và mở modal yêu cầu nạp lượt.

---

*Tài liệu kế hoạch bàn giao hoàn tất và sẵn sàng để triển khai ngay!*