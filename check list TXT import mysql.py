import os
import pymysql
import pandas as pd
from dotenv import load_dotenv
from urllib.parse import urlparse, unquote  # Thư viện bóc tách URI chính xác

# Thư viện giao diện chọn thư mục
import tkinter as tk
from tkinter import filedialog

# 1. Nạp biến môi trường từ file .env (Giữ nguyên không sửa đổi .env)
load_dotenv()

# Đọc chuỗi URI từ biến có sẵn của bạn
# Đảm bảo bạn đã có biến này trong file .env
db_uri = os.environ.get("SQLALCHEMY_BINDS_BC48")

def parse_mysql_uri_standard(uri):
    """
    Hàm bóc tách chuỗi kết nối chuẩn,
    Bảo vệ toàn vẹn mật khẩu có ký tự đặc biệt.
    """
    if not uri:
        raise ValueError("Không tìm thấy cấu hình SQLALCHEMY_BINDS_BC48 trong file .env")
    
    # Chuẩn hóa tiền tố
    if "://" in uri:
        schemeless_uri = uri.split("://", 1)[1]
    else:
        schemeless_uri = uri

    # Dùng urlparse để bóc tách cấu trúc URL
    parsed = urlparse(f"http://{schemeless_uri}")
    
    # Giải mã các ký tự đặc biệt trong user và password
    user = unquote(parsed.username) if parsed.username else ""
    password = unquote(parsed.password) if parsed.password else ""
    host = parsed.hostname
    port = parsed.port if parsed.port else 3306
    database = parsed.path.lstrip('/')

    return {
        'host': host,
        'port': port,
        'user': user,
        'password': password,
        'database': database,
        'charset': 'utf8mb4'
    }

def get_folder_path_from_ui():
    """Hàm mở cửa sổ giao diện để người dùng chọn thư mục"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    selected_folder = filedialog.askdirectory(title="Vui lòng chọn thư mục chứa các file cần kiểm tra")
    root.destroy()
    return selected_folder

def check_files_and_export_excel():
    # Bước 1: Chọn thư mục
    folder_path = get_folder_path_from_ui()
    if not folder_path:
        print("Bạn đã hủy chọn thư mục. Tiến trình dừng lại.")
        return
        
    print(f"Đang xử lý thư mục đã chọn: {folder_path}")

    # Lấy danh sách file thực tế, chuẩn hóa về viết thường toàn bộ để so sánh
    files_in_folder_raw = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    if not files_in_folder_raw:
        print("Thư mục được chọn trống hoặc không chứa file.")
        return

    # Bước 2: Kết nối DB và lấy danh sách file, chuẩn hóa về viết thường toàn bộ
    files_in_db_normalized = set()
    files_in_db_original = [] # Lưu danh sách gốc để phục vụ ghi log/debug nếu cần
    try:
        # Bóc tách cấu hình từ URI
        db_config = parse_mysql_uri_standard(db_uri)
        
        # Kết nối trực tiếp qua pymysql thuần
        connection = pymysql.connect(**db_config)
        with connection.cursor() as cursor:
            # Truy vấn tên file gốc
            query = "SELECT file_name FROM log_file_imports_latest"
            cursor.execute(query)
            
            # Đưa vào Set dưới dạng VIẾT THƯỜNG TOÀN BỘ để tối ưu tìm kiếm
            rows = cursor.fetchall()
            for row in rows:
                original_name = row[0]
                files_in_db_original.append(original_name)
                files_in_db_normalized.add(original_name.lower())
            
    except Exception as e:
        print(f"Lỗi khi kết nối hoặc truy vấn Cơ sở dữ liệu: {e}")
        return
    finally:
        if 'connection' in locals() and connection.open:
            connection.close()

    # Bước 3: So sánh dữ liệu (bằng cách chuẩn hóa cả hai về viết thường)
    matched_files = []
    unmatched_files = []

    for file_name_original in files_in_folder_raw:
        # Chuẩn hóa tên file thực tế về viết thường
        file_name_lower = file_name_original.lower()
        
        # So sánh với Set đã được chuẩn hóa
        if file_name_lower in files_in_db_normalized:
            matched_files.append({
                "Tên file thực tế": file_name_original,
                "Trạng thái": "Trùng khớp (Bất chấp hoa thường)"
            })
        else:
            unmatched_files.append({
                "Tên file thực tế": file_name_original,
                "Trạng thái": "Không trùng khớp"
            })

    # Bước 4: Ghi kết quả ra file Excel
    df_matched = pd.DataFrame(matched_files)
    df_unmatched = pd.DataFrame(unmatched_files)

    output_excel = "ket_qua_kiem_tra.xlsx"
    
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        if not df_matched.empty:
            df_matched.to_excel(writer, sheet_name="Trùng khớp", index=False)
        else:
            pd.DataFrame(columns=["Tên file thực tế", "Trạng thái"]).to_excel(writer, sheet_name="Trùng khớp", index=False)
            
        if not df_unmatched.empty:
            df_unmatched.to_excel(writer, sheet_name="Không trùng khớp", index=False)
        else:
            pd.DataFrame(columns=["Tên file thực tế", "Trạng thái"]).to_excel(writer, sheet_name="Không trùng khớp", index=False)

    print(f"\n Quá trình kiểm tra hoàn tất!")
    print(f"-> Kết quả đã được ghi vào file: {os.path.abspath(output_excel)}")
    print(f"   + Số file trùng khớp (bất chấp hoa thường): {len(matched_files)}")
    print(f"   + Số file không trùng khớp: {len(unmatched_files)}")

if __name__ == "__main__":
    check_files_and_export_excel()
