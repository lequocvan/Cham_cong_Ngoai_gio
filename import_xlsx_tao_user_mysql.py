import pandas as pd
import numpy as np
from sqlalchemy import create_engine, text
from werkzeug.security import generate_password_hash
from datetime import datetime
from dotenv import load_dotenv
import os

# --- TỰ ĐỘNG LOAD CẤU HÌNH TỪ .env ---
load_dotenv()

# Sử dụng đúng tên biến DATABASE_URL từ file .env của bạn
DB_URI = os.getenv('DATABASE_URL')
FILE_PATH = 'thong_tin_nguoi_lao_dong_tao_user.xlsx'

# Khởi tạo engine kết nối database
if not DB_URI:
    raise ValueError("Không tìm thấy biến DATABASE_URL trong file .env!")

engine = create_engine(DB_URI)

def clean_date(val):
    if pd.isna(val) or val is None: return None
    if isinstance(val, datetime): return val
    try:
        res = pd.to_datetime(val)
        return res.to_pydatetime() if not pd.isna(res) else None
    except:
        return None

def clean_phone_or_gttt(val):
    """Xử lý chuẩn hóa số điện thoại hoặc GTTT, đảm bảo không mất số 0 ở đầu"""
    if pd.isna(val) or val is None: return None
    
    # Ép về chuỗi và loại bỏ phần thập phân (.0 nếu có)
    val_str = str(val).split('.')[0].strip()
    
    if not val_str or val_str.lower() == 'none':
        return None
        
    # Xử lý riêng cho Số điện thoại Việt Nam: 
    # Nếu bắt đầu bằng số khác 0 và có độ dài 9 chữ số (bị Excel cắt mất số 0), tự động thêm '0' vào đầu
    if len(val_str) == 9 and not val_str.startswith('0'):
        val_str = '0' + val_str
        
    return val_str

def import_data():
    try:
        # QUAN TRỌNG: Ép kiểu các cột mã, số điện thoại, GTTT thành chuỗi ngay khi đọc Excel để giữ nguyên số 0
        df = pd.read_excel(FILE_PATH, dtype={
            'ma_nhan_vien': str, 
            'so_dien_thoai': str, 
            'so_gttt': str,
            'ma_phong_ban': str
        })
        
        df.columns = [str(c).strip() for c in df.columns]
        df = df.replace({np.nan: None, pd.NaT: None})

        success_count = 0
        update_count = 0

        with engine.connect() as conn:
            with conn.begin(): # Mở transaction
                for _, row in df.iterrows():
                    ma_nv = str(row.get('ma_nhan_vien', '')).strip()
                    if not ma_nv or ma_nv.lower() == 'none': continue

                    # Đầy đủ tham số khớp hoàn toàn với cấu trúc bảng thong_tin_nguoi_lao_dong
                    params = {
                        'ma': ma_nv,
                        'ten': str(row.get('ho_ten', '')).strip(),
                        'ns': clean_date(row.get('ngay_sinh')),
                        'gt': row.get('gioi_tinh'),
                        'gttt': clean_phone_or_gttt(row.get('so_gttt')),
                        'dt': clean_phone_or_gttt(row.get('so_dien_thoai')),
                        'mail': row.get('mail_Agribank'),
                        'dc': row.get('dia_chi'),
                        'np_phep': clean_date(row.get('ngay_tinh_phep')),
                        'nv_agri': clean_date(row.get('ngay_vao_Agribank')),
                        'pb': row.get('ma_phong_ban'),
                        'dv': row.get('ma_hieu_2'),
                        'cv': row.get('chuc_vu'),
                        'tt': int(row.get('trang_thai')) if row.get('trang_thai') is not None else 1
                    }

                    # --- Thao tác với bảng thong_tin_nguoi_lao_dong ---
                    check_nv = conn.execute(
                        text("SELECT id FROM thong_tin_nguoi_lao_dong WHERE ma_nhan_vien = :ma"),
                        {'ma': ma_nv}
                    ).fetchone()

                    if not check_nv:
                        # INSERT (Thêm mới)
                        conn.execute(text("""
                            INSERT INTO thong_tin_nguoi_lao_dong 
                            (ma_nhan_vien, ho_ten, ngay_sinh, gioi_tinh, so_gttt, so_dien_thoai, mail_Agribank, 
                             dia_chi, ngay_tinh_phep, ngay_vao_Agribank, ma_phong_ban, ma_hieu_2, chuc_vu, trang_thai)
                            VALUES (:ma, :ten, :ns, :gt, :gttt, :dt, :mail, :dc, :np_phep, :nv_agri, :pb, :dv, :cv, :tt)
                        """), params)
                        success_count += 1
                    else:
                        # UPDATE (Cập nhật thông tin mới nhất từ Excel)
                        conn.execute(text("""
                            UPDATE thong_tin_nguoi_lao_dong SET 
                            ho_ten=:ten, ngay_sinh=:ns, gioi_tinh=:gt, so_gttt=:gttt, so_dien_thoai=:dt, mail_Agribank=:mail, 
                            dia_chi=:dc, ngay_tinh_phep=:np_phep, ngay_vao_Agribank=:nv_agri, 
                            ma_phong_ban=:pb, ma_hieu_2=:dv, chuc_vu=:cv, trang_thai=:tt
                            WHERE ma_nhan_vien = :ma
                        """), params)
                        update_count += 1

                    # --- Đồng bộ tài khoản sang bảng users ---
                    check_user = conn.execute(
                        text("SELECT ma_nhan_vien FROM users WHERE ma_nhan_vien = :ma"),
                        {'ma': ma_nv}
                    ).fetchone()

                    if not check_user:
                        # Mật khẩu mặc định là mã nhân viên, đã mã hóa an toàn bằng werkzeug hash
                        hashed_pw = generate_password_hash(ma_nv)
                        conn.execute(text("""
                            INSERT INTO users (ma_nhan_vien, password_hash, fullname, role, is_active)
                            VALUES (:ma, :pw, :ten, 'LAP_BANG', 1)
                        """), {'ma': ma_nv, 'pw': hashed_pw, 'ten': params['ten']})

        print(f"Hoàn thành! Thêm mới: {success_count} nhân viên, Cập nhật: {update_count} nhân viên.")

    except Exception as e:
        print(f"Lỗi xảy ra: {e}")

if __name__ == '__main__':
    import_data()
