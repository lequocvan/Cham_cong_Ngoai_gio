# Kiểm tra dữ liệu trong các bảng _yyyymm thay thế sp_TuDongKiemTraDuLieu (event 11h00 hàng ngày) trong mysql
# 1. Kiểm tra macn -> don_vi
# 2. Kiểm tra quoctich -> quoc_gia
# 3. Kiểm tra kieukh -> loai_khach_hang
# 4. Kiểm tra loaitien -> loai_tien
# 5. Kiểm tra loaigd -> ma_loai_nghiep_vu_pcrt
# 6. Kiểm tra kenhct -> kenh_chuyen_tien
# 7. Kiểm tra loaigt (DWT, EFT, PTR dùng 'loaigt') -> loai_giay_to
# 8. Kiểm tra loaigto (Riêng CTR dùng 'loaigto') -> loai_giay_to
# 9. Kiểm tra loaihanghoa -> loai_hang_hoa
# 10. Kiểm tra loaitk -> loai_tai_khoan

# TRUNCATE danh_sach_bang_da_kiem_tra;
# TRUNCATE log_kiem_tra_du_lieu;

### 🔒 Lưu ý về logic bảo toàn:
# Khi bạn gõ đích danh bảng hoặc chọn quét `all`, hệ thống sẽ **không lưu** bản ghi vào bảng `danh_sach_bang_da_kiem_tra` sau khi chạy xong.
# Điều này giúp bảo toàn cơ chế quét tự động: các bảng đó vẫn được giữ nguyên trạng thái để lần sau bạn gõ Enter thì hệ thống vẫn nhận diện nó là bảng cần quét. 
# Lệnh `INSERT IGNORE` được sử dụng để tránh lỗi trùng lặp dữ liệu (`Duplicate entry`) nếu có xung đột chỉ mục bảng.

from datetime import datetime
from textwrap import dedent
import os
import sys
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, text

# ==============================================================================
# 1. ĐỊNH NGHĨA LỚP GIẢ LẬP (CHỈ DÙNG KHI CHẠY TERMINAL ĐỘC LẬP)
# ==============================================================================
class DummyDB:
    def __init__(self, engine):
        self.engine = engine

    def get_engine(self, bind=None):
        """Giả lập method get_engine của Flask-SQLAlchemy"""
        return self.engine


# ==============================================================================
# 2. HÀM XỬ LÝ DỮ LIỆU CHÍNH (HỖ TRỢ LỌC THEO NGÀY QUA CỘT THOIDIEM)
# ==============================================================================
def tu_dong_kiem_tra_du_lieu_v2(db):
    """Hàm xử lý đối chiếu danh mục, tự động lọc bảng và lưu vết thoidiem giao dịch lỗi"""
    engine = db.get_engine(bind="db_bc48")

    with engine.connect() as conn:
        # --- TÍNH NĂNG 1: HỎI DỌN DẸP DỮ LIỆU CŨ ---
        print("\n=== CẤU HÌNH KHỞI ĐỘNG ===")
        clear_choice = (
            input(
                "Bạn có muốn TRUNCATE (làm sạch) dữ liệu log cũ trong MySQL không? (y/n): "
            )
            .strip()
            .lower()
        )

        if clear_choice == "y":
            print("-> Đang làm sạch các bảng log...")
            try:
                conn.execute(text("TRUNCATE TABLE log_kiem_tra_du_lieu;"))
                conn.execute(text("TRUNCATE TABLE danh_sach_bang_da_kiem_tra;"))
                conn.commit()
                print("✓ Đã TRUNCATE thành công hai bảng log!")
            except Exception as e:
                print(f"⚠️ Không thể TRUNCATE ({e}), chuyển sang dùng DELETE...")
                conn.execute(text("DELETE FROM log_kiem_tra_du_lieu;"))
                conn.execute(text("DELETE FROM danh_sach_bang_da_kiem_tra;"))
                conn.commit()
                print("✓ Đã DELETE sạch dữ liệu hai bảng log!")

        # --- TÍNH NĂNG LỌC DỮ LIỆU THEO NGÀY ---
        print("-" * 50)
        print("CẤU HÌNH NGÀY KIỂM TRA:")
        print("[Ấn Enter]: Quét toàn bộ các ngày có trong bảng (Quét tất cả các bảng).")
        print("[Gõ ngày cụ thể]: Điền dạng yyyymmdd (Ví dụ: 20260520)")
        print("-" * 50)
        target_date = input("Nhập ngày muốn quét: ").strip()
        
        target_ym = ""  # Biến lưu chuỗi yyyymm để lọc tên bảng
        if target_date != "":
            if len(target_date) >= 6:
                target_ym = target_date[:6]
            print(f"-> Chế độ: Chỉ quét dữ liệu phát sinh ngày: {target_date}")
        else:
            print("-> Chế độ: Quét toàn bộ dữ liệu (không lọc ngày).")

        # --- QUÉT VÀ LỌC TÊN BẢNG PHÙ HỢP ---
        if target_ym != "":
            regex_pattern = f"^(ctr|dwt|eft|ptr)_{target_ym}$"
            print(f"-> Tự động cấu hình lọc bảng có tên kết thúc bằng: _{target_ym}")
        else:
            regex_pattern = "^(ctr|dwt|eft|ptr)_[0-9]{6}$"

        query_tables = text(
            dedent(
                """
            SELECT table_name 
            FROM information_schema.tables 
            WHERE table_schema = DATABASE() 
              AND (table_name REGEXP :pattern)
        """
            )
        )
        all_tables = [row[0] for row in conn.execute(query_tables, {"pattern": regex_pattern}).fetchall()]

        # Đọc danh sách các bảng đã từng check trước đó
        query_checked = text("SELECT ten_bang FROM danh_sach_bang_da_kiem_tra")
        try:
            checked_tables = {row[0] for row in conn.execute(query_checked).fetchall()}
        except Exception:
            checked_tables = set()

        tables_unpreprocessed = [t for t in all_tables if t not in checked_tables]

        # --- TÍNH NĂNG 2: LỰA CHỌN BẢNG CẦN QUÉT ---
        print(f"\nHệ thống tìm thấy tổng cộng {len(all_tables)} bảng báo cáo phù hợp tiêu chí tháng/năm.")
        print(f"Trong đó có {len(tables_unpreprocessed)} bảng mới chưa quét.")
        print("-" * 50)
        print("LỰA CHỌN CHẾ ĐỘ QUÉT BẢNG:")
        print("1. [Ấn Enter]: Chỉ quét các bảng mới chưa từng kiểm tra.")
        print("2. [Gõ 'all']: Quét lại toàn bộ tất cả các bảng phù hợp.")
        print("3. [Gõ tên bảng]: Điền đích danh các bảng, cách nhau bằng dấu phẩy.")
        print("-" * 50)

        user_input = input("Nhập lựa chọn của bạn: ").strip()

        if user_input == "":
            tables_to_check = tables_unpreprocessed
            mode_text = "Quét các bảng mới chưa xử lý"
        elif user_input.lower() == "all":
            tables_to_check = all_tables
            mode_text = "Quét lại toàn bộ tất cả các bảng"
        else:
            custom_tables = [t.strip() for t in user_input.split(",") if t.strip()]
            tables_to_check = [t for t in custom_tables if t in all_tables]

            invalid_tables = [t for t in custom_tables if t not in all_tables]
            if invalid_tables:
                print(f"⚠️ Cảnh báo: Các bảng sau không tồn tại hoặc không khớp tháng/năm: {invalid_tables}")

            mode_text = f"Quét danh sách bảng chỉ định ({len(tables_to_check)} bảng)"

        if not tables_to_check:
            print("\n❌ Không có bảng nào hợp lệ để tiến hành kiểm tra!")
            return

        print(f"\n--- Khởi động tiến trình: {mode_text} ---\n")

        # --- NẠP TOÀN BỘ 10 DANH MỤC HỆ THỐNG VÀO RAM ---
        print("Đang nạp danh mục hệ thống vào bộ nhớ RAM...")
        set_don_vi = set(pd.read_sql("SELECT MaNH8so_moi FROM don_vi", conn)["MaNH8so_moi"].astype(str).str.strip().str.lower())
        set_quoc_gia = set(pd.read_sql("SELECT ma_Alpha2 FROM quoc_gia", conn)["ma_Alpha2"].astype(str).str.strip().str.lower())
        set_lkh = set(pd.read_sql("SELECT ma_loai_khach_hang FROM loai_khach_hang", conn)["ma_loai_khach_hang"].astype(str).str.strip().str.lower())
        set_loai_tien = set(pd.read_sql("SELECT loaitien FROM loai_tien", conn)["loaitien"].astype(str).str.strip().str.lower())
        set_loaigd = set(pd.read_sql("SELECT ma_nghiep_vu FROM ma_loai_nghiep_vu_pcrt", conn)["ma_nghiep_vu"].astype(str).str.strip().str.lower())
        set_kenhct = set(pd.read_sql("SELECT ma_kenh_ct FROM kenh_chuyen_tien", conn)["ma_kenh_ct"].astype(str).str.strip().str.lower())
        set_giay_to = set(pd.read_sql("SELECT ma_loai_giay_to FROM loai_giay_to", conn)["ma_loai_giay_to"].astype(str).str.strip().str.lower())
        set_hang_hoa = set(pd.read_sql("SELECT ma_loai_hang_hoa FROM loai_hang_hoa", conn)["ma_loai_hang_hoa"].astype(str).str.strip().str.lower())
        set_tai_khoan = set(pd.read_sql("SELECT ma_loai_tai_khoan FROM loai_tai_khoan", conn)["ma_loai_tai_khoan"].astype(str).str.strip().str.lower())
        print("✓ Đã nạp xong tất cả danh mục.\n")

        # --- VÒNG LẶP KIỂM TRA CHO TỪNG BẢNG GIAO DỊCH ---
        for table_name in tables_to_check:
            # 1. TỐI ƯU LỆNH XÓA LOG CŨ: Xóa trực tiếp theo ngày dựa vào cột thoidiem mới thêm
            if target_date != "":
                print(f"-> Đang quét bảng: {table_name} (Chỉ lọc ngày {target_date})")
                sql_delete_by_date = text(
                    "DELETE FROM log_kiem_tra_du_lieu WHERE ten_bang = :t AND thoidiem LIKE :date_pattern"
                )
                conn.execute(sql_delete_by_date, {"t": table_name, "date_pattern": f"{target_date}%"})
            else:
                print(f"-> Đang quét bảng: {table_name}")
                conn.execute(
                    text("DELETE FROM log_kiem_tra_du_lieu WHERE ten_bang = :t"),
                    {"t": table_name}
                )
            conn.commit()
            
            # 2. ĐỌC DỮ LIỆU TỪ BẢNG GIAO DỊCH
            if target_date != "":
                sql_select_target = text(f"SELECT * FROM {table_name} WHERE thoidiem LIKE :date_pattern")
                df_target = pd.read_sql(sql_select_target, conn, params={"date_pattern": f"{target_date}%"})
            else:
                sql_select_target = text(f"SELECT * FROM {table_name}")
                df_target = pd.read_sql(sql_select_target, conn)

            # 3. KIỂM TRA DỮ LIỆU BẢNG RỖNG
            if df_target.empty:
                print(f"   => Không có dữ liệu phù hợp trong bảng {table_name}, bỏ qua.")
                if target_date == "":
                    conn.execute(
                        text("INSERT IGNORE INTO danh_sach_bang_da_kiem_tra (ten_bang) VALUES (:t)"),
                        {"t": table_name}
                    )
                    conn.commit()
                continue

            columns = df_target.columns.tolist()
            list_errors = []

            # Hàm kiểm tra bằng Pandas (Bổ sung thu thập row["thoidiem"])
            def check_column(col_name, master_set_lower):
                # Đảm bảo bảng có đầy đủ các cột định danh quan trọng
                if col_name in columns and "magd" in columns and "thoidiem" in columns:
                    v_str = df_target[col_name].astype(str).str.strip().str.lower()
                    ignored_values = {"", "nan", "none"}
                    
                    is_error_mask = (
                        df_target[col_name].notna() 
                        & (~v_str.isin(ignored_values)) 
                        & (~v_str.isin(master_set_lower))
                    )
                    
                    df_err = df_target[is_error_mask]
                    for _, row in df_err.iterrows():
                        # Đẩy thêm phần tử row["thoidiem"] vào vị trí index số 4
                        list_errors.append((table_name, col_name, str(row[col_name]), row["magd"], str(row["thoidiem"])))

            # --- THỰC THI 10 TIÊU CHÍ KIỂM TRA ---
            check_column("macn", set_don_vi)
            check_column("quoctich", set_quoc_gia)
            check_column("kieukh", set_lkh)
            check_column("loaitien", set_loai_tien)
            check_column("loaigd", set_loaigd)
            check_column("kenhct", set_kenhct)

            if table_name.lower().startswith("ctr"):
                check_column("loaigto", set_giay_to)
            else:
                check_column("loaigt", set_giay_to)

            check_column("loaihanghoa", set_hang_hoa)
            check_column("loaitk", set_tai_khoan)

            # --- GHI LOG LỖI VÀO CƠ SỞ DỮ LIỆU ---
            if list_errors:
                print(f"   => Tìm thấy {len(list_errors)} bản ghi lỗi.")
                # Thêm cột thoidiem vào cấu trúc INSERT INTO
                sql_insert = text(
                    "INSERT INTO log_kiem_tra_du_lieu (ten_bang, ten_cot, gia_tri_loi, ma_giao_dich, thoidiem) "
                    "VALUES (:ten_bang, :ten_cot, :gia_tri_loi, :ma_giao_dich, :thoidiem)"
                )
                conn.execute(
                    sql_insert,
                    [
                        {
                            "ten_bang": x[0],
                            "ten_cot": x[1],
                            "gia_tri_loi": x[2],
                            "ma_giao_dich": x[3],
                            "thoidiem": x[4], # Gán giá trị thời điểm lỗi phát sinh
                        }
                        for x in list_errors
                    ],
                )
                conn.commit()
            else:
                print("   => Đoạn dữ liệu quét hợp lệ (0 lỗi).")

            # --- CẬP NHẬT LOG BẢNG ĐÃ HOÀN THÀNH ---
            if target_date == "":
                conn.execute(
                    text("INSERT IGNORE INTO danh_sach_bang_da_kiem_tra (ten_bang) VALUES (:t)"),
                    {"t": table_name}
                )
                conn.commit()

        print("\n--- Tiến trình hoàn tất thành công ---")

# ==============================================================================
# 3. KÍCH HOẠT KHI CHẠY ĐỘC LẬP TỪ TERMINAL
# ==============================================================================
if __name__ == "__main__":
    print("Đang cấu hình môi trường chạy độc lập...")

    env_path = Path(__file__).resolve().parent / ".env"
    if env_path.exists():
        load_dotenv(dotenv_path=env_path)
        print("-> Đã nạp cấu hình từ file .env.")
    else:
        print("⚠️ Cảnh báo: Không tìm thấy file .env.")

    raw_uri = os.getenv("SQLALCHEMY_BINDS_BC48")
    if not raw_uri:
        raw_uri = os.getenv("DATABASE_URL")
        print("-> Không tìm thấy SQLALCHEMY_BINDS_BC48, dùng DATABASE_URL.")

    if not raw_uri:
        print("❌ LỖI: Không tìm thấy chuỗi kết nối trong file .env!", file=sys.stderr)
        sys.exit(1)

    DATABASE_URI = raw_uri.strip('"').strip("'")

    if DATABASE_URI.startswith("mysql://"):
        DATABASE_URI = DATABASE_URI.replace("mysql://", "mysql+pymysql://", 1)

    try:
        print("-> Đang thiết lập kết nối tới MySQL...")
        raw_engine = create_engine(DATABASE_URI, pool_recycle=3600, pool_size=5)

        with raw_engine.connect() as test_conn:
            test_conn.execute(text("SELECT 1"))
        print("✓ Kết nối cơ sở dữ liệu thành công!")

        db_standalone = DummyDB(raw_engine)
        tu_dong_kiem_tra_du_lieu_v2(db_standalone)

    except Exception as err:
        print(f"❌ Lỗi kết nối hoặc thực thi hệ thống: {err}", file=sys.stderr)
