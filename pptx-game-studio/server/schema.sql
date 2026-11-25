-- =====================================================================
-- PPTX Game Studio – CSDL quản lý lượt xuất & mã kích hoạt
-- Dùng cho PostgreSQL / Supabase. Chạy toàn bộ file này 1 lần trong
-- Supabase > SQL Editor > New query > Run.
-- Bảng được khóa bằng RLS (không ai đọc/ghi trực tiếp được), trình duyệt chỉ
-- được gọi 3 hàm bên dưới; mọi kiểm tra/trừ lượt xử lý nguyên tử trong CSDL.
-- =====================================================================

-- 1. Bảng người dùng / thiết bị
CREATE TABLE IF NOT EXISTS users (
  id TEXT PRIMARY KEY,
  remaining_exports INTEGER NOT NULL DEFAULT 5 CHECK (remaining_exports >= 0),
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
  note TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 3. Nhật ký xuất file
CREATE TABLE IF NOT EXISTS export_logs (
  id BIGSERIAL PRIMARY KEY,
  user_id TEXT NOT NULL REFERENCES users(id),
  selected_count INTEGER NOT NULL DEFAULT 0,
  estimated_slides INTEGER NOT NULL DEFAULT 0,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 4. Nhật ký nạp mã (kể cả nhập sai – để phát hiện dò mã)
CREATE TABLE IF NOT EXISTS redeem_attempts (
  id BIGSERIAL PRIMARY KEY,
  user_id TEXT NOT NULL,
  code TEXT NOT NULL,
  ok BOOLEAN NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS redeem_attempts_user_time ON redeem_attempts (user_id, created_at);

ALTER TABLE users ENABLE ROW LEVEL SECURITY;
ALTER TABLE activation_codes ENABLE ROW LEVEL SECURITY;
ALTER TABLE export_logs ENABLE ROW LEVEL SECURITY;
ALTER TABLE redeem_attempts ENABLE ROW LEVEL SECURITY;
REVOKE ALL ON users, activation_codes, export_logs, redeem_attempts FROM anon, authenticated;

-- Kiểm tra định dạng mã thiết bị
CREATE OR REPLACE FUNCTION _valid_user_id(p_user_id TEXT) RETURNS BOOLEAN
LANGUAGE sql IMMUTABLE AS $$ SELECT p_user_id ~ '^dev_[a-z0-9]{12,40}$' $$;

-- ---------------------------------------------------------------------
-- Lấy số lượt còn lại (tự tạo người dùng mới với lượt dùng thử mặc định)
-- ---------------------------------------------------------------------
CREATE OR REPLACE FUNCTION get_export_balance(p_user_id TEXT)
RETURNS TABLE (remaining_exports INT)
LANGUAGE plpgsql SECURITY DEFINER SET search_path = public AS $$
BEGIN
  IF NOT _valid_user_id(p_user_id) THEN
    RAISE EXCEPTION 'Mã thiết bị không hợp lệ';
  END IF;
  INSERT INTO users (id) VALUES (p_user_id) ON CONFLICT (id) DO NOTHING;
  RETURN QUERY SELECT u.remaining_exports FROM users u WHERE u.id = p_user_id;
END $$;

-- ---------------------------------------------------------------------
-- Nạp mã kích hoạt (khóa dòng FOR UPDATE chống dùng 1 mã 2 lần)
-- ---------------------------------------------------------------------
CREATE OR REPLACE FUNCTION redeem_activation_code(p_code TEXT, p_user_id TEXT)
RETURNS TABLE (ok BOOLEAN, message TEXT, added_credits INT, remaining_exports INT)
LANGUAGE plpgsql SECURITY DEFINER SET search_path = public AS $$
DECLARE
  v_code activation_codes%ROWTYPE;
  v_left INT;
  v_fails INT;
BEGIN
  IF NOT _valid_user_id(p_user_id) THEN
    RAISE EXCEPTION 'Mã thiết bị không hợp lệ';
  END IF;
  p_code := upper(regexp_replace(coalesce(p_code, ''), '\s', '', 'g'));
  INSERT INTO users (id) VALUES (p_user_id) ON CONFLICT (id) DO NOTHING;
  SELECT u.remaining_exports INTO v_left FROM users u WHERE u.id = p_user_id;

  -- chống dò mã: tối đa 10 lần nhập sai / 15 phút / thiết bị
  SELECT count(*) INTO v_fails FROM redeem_attempts
   WHERE user_id = p_user_id AND NOT redeem_attempts.ok AND created_at > NOW() - INTERVAL '15 minutes';
  IF v_fails >= 10 THEN
    RETURN QUERY SELECT FALSE, 'Bạn nhập sai quá nhiều lần. Vui lòng thử lại sau 15 phút.'::TEXT, 0, v_left; RETURN;
  END IF;

  SELECT * INTO v_code FROM activation_codes WHERE code = p_code FOR UPDATE;
  IF NOT FOUND THEN
    INSERT INTO redeem_attempts (user_id, code, ok) VALUES (p_user_id, left(p_code, 60), FALSE);
    RETURN QUERY SELECT FALSE, 'Mã kích hoạt không tồn tại.'::TEXT, 0, v_left; RETURN;
  END IF;
  IF v_code.status = 'used' THEN
    INSERT INTO redeem_attempts (user_id, code, ok) VALUES (p_user_id, p_code, FALSE);
    RETURN QUERY SELECT FALSE, 'Mã này đã được sử dụng.'::TEXT, 0, v_left; RETURN;
  END IF;
  IF v_code.status = 'disabled' THEN
    INSERT INTO redeem_attempts (user_id, code, ok) VALUES (p_user_id, p_code, FALSE);
    RETURN QUERY SELECT FALSE, 'Mã này đã bị vô hiệu hóa.'::TEXT, 0, v_left; RETURN;
  END IF;
  IF v_code.expires_at IS NOT NULL AND v_code.expires_at < NOW() THEN
    INSERT INTO redeem_attempts (user_id, code, ok) VALUES (p_user_id, p_code, FALSE);
    RETURN QUERY SELECT FALSE, 'Mã này đã hết hạn.'::TEXT, 0, v_left; RETURN;
  END IF;

  UPDATE activation_codes SET status = 'used', used_by = p_user_id, used_at = NOW() WHERE code = p_code;
  UPDATE users SET remaining_exports = users.remaining_exports + v_code.credits, updated_at = NOW()
   WHERE id = p_user_id RETURNING users.remaining_exports INTO v_left;
  INSERT INTO redeem_attempts (user_id, code, ok) VALUES (p_user_id, p_code, TRUE);
  RETURN QUERY SELECT TRUE, format('Nạp thành công %s lượt xuất!', v_code.credits)::TEXT, v_code.credits, v_left;
END $$;

-- ---------------------------------------------------------------------
-- Trừ 1 lượt khi xuất file (nguyên tử: chỉ trừ khi còn >= 1)
-- ---------------------------------------------------------------------
CREATE OR REPLACE FUNCTION consume_export_credit(p_user_id TEXT, p_selected_count INT, p_estimated_slides INT)
RETURNS TABLE (allowed BOOLEAN, message TEXT, remaining_exports INT)
LANGUAGE plpgsql SECURITY DEFINER SET search_path = public AS $$
DECLARE v_left INT;
BEGIN
  IF NOT _valid_user_id(p_user_id) THEN
    RAISE EXCEPTION 'Mã thiết bị không hợp lệ';
  END IF;
  IF coalesce(p_selected_count, 0) < 1 THEN
    SELECT u.remaining_exports INTO v_left FROM users u WHERE u.id = p_user_id;
    RETURN QUERY SELECT FALSE, 'Chưa chọn mục nào để xuất – không trừ lượt.'::TEXT, coalesce(v_left, 0); RETURN;
  END IF;
  INSERT INTO users (id) VALUES (p_user_id) ON CONFLICT (id) DO NOTHING;
  UPDATE users SET remaining_exports = users.remaining_exports - 1, updated_at = NOW()
   WHERE id = p_user_id AND users.remaining_exports >= 1
   RETURNING users.remaining_exports INTO v_left;
  IF NOT FOUND THEN
    RETURN QUERY SELECT FALSE, 'Bạn đã hết lượt xuất PPTX.'::TEXT, 0; RETURN;
  END IF;
  INSERT INTO export_logs (user_id, selected_count, estimated_slides) VALUES (p_user_id, p_selected_count, coalesce(p_estimated_slides, 0));
  RETURN QUERY SELECT TRUE, 'Đã trừ 1 lượt xuất.'::TEXT, v_left;
END $$;

-- Chỉ cho phép trình duyệt (anon) gọi 3 hàm này
REVOKE ALL ON FUNCTION get_export_balance(TEXT), redeem_activation_code(TEXT, TEXT), consume_export_credit(TEXT, INT, INT) FROM PUBLIC;
GRANT EXECUTE ON FUNCTION get_export_balance(TEXT), redeem_activation_code(TEXT, TEXT), consume_export_credit(TEXT, INT, INT) TO anon, authenticated;
