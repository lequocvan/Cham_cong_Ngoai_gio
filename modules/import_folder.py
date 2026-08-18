import os
from pathlib import Path
from sqlalchemy import create_engine
# Import hàm xử lý từ file bc48.py của anh; chạy script này dưới dạng một công cụ dòng lệnh (Terminal) thuần túy để nạp data vào MySQL mà không muốn phụ thuộc vào các cấu hình loằng ngoằng của Flask App
from bc48 import process_import_files

class DummyDB:
    def __init__(self, engine):
        self.engine = engine

    def get_engine(self, bind=None):
        """Giả lập method get_engine của Flask-SQLAlchemy"""
        return self.engine

def import_all_files_from_folder(folder_path):
    path = Path(folder_path)
    if not path.exists() or not path.is_dir():
        print(f"❌ Thư mục không tồn tại: {folder_path}")
        return
    
    # 1. Tự tạo kết nối đến MySQL 
    engine = create_engine(
        "mysql+pymysql://write:ktgsnb%401050@127.0.0.1:3306/bc48?charset=utf8mb4",
        connect_args={"local_infile": 1}
    )
    db_dummy = DummyDB(engine)
    
    # 2. Tìm tất cả file .TXT
    all_files = [f for f in path.iterdir() if f.is_file() and f.suffix.upper() == '.TXT']
    total_files = len(all_files)
    print(f"📂 Tìm thấy {total_files} file .TXT.")
    
    if total_files == 0:
        return

    # 3. Đọc dữ liệu raw bytes
    print("⏳ Đang đọc dữ liệu các file vào bộ nhớ...")
    files_data = []
    for file_path in all_files:
        with open(file_path, 'rb') as f:
            files_data.append((file_path.name, f.read()))
    
    print("🚀 Bắt đầu nạp đồng loạt vào MySQL...")
    
    # 4. Truyền db_dummy (đã có get_engine) vào hàm xử lý
    generator = process_import_files(db_dummy, files_data, current_user_name="Hệ thống (Import Folder)")
    
    for response_line in generator:
        clean_line = response_line.replace("data: ", "").strip()
        if clean_line:
            print(f"➡️ {clean_line}")

    print("\n✅ QUÁ TRÌNH IMPORT ĐỒNG LOẠT HOÀN TẤT!")

if __name__ == "__main__":
    TARGET_FOLDER = "/Users/lequocvan/Downloads/TXT GLD nap vao bc48 2025 den 20260514"
    import_all_files_from_folder(TARGET_FOLDER)
