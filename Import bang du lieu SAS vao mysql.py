import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, text, inspect

# --- TỰ ĐỘNG LOAD CẤU HÌNH TỪ .env ---
load_dotenv()

DB_URI = os.getenv("DATABASE_URL")
if not DB_URI:
    raise ValueError("Không tìm thấy biến DATABASE_URL trong file .env!")

engine = create_engine(DB_URI)

# --- CẤU HÌNH ---
FOLDER_PATH = "./data_files_sas"
SINGLE_TABLE_NAME = "master_data_sas"


def sanitize_column_name(col_name: str) -> str:
    """Loại bỏ ký tự đặc biệt trong tên cột để không bị lỗi SQL."""
    col_str = str(col_name).strip()
    col_clean = re.sub(r"\W+", "_", col_str).strip("_").lower()
    return col_clean if col_clean else "unnamed_col"


def delete_existing_file_records(db_engine, table_name: str, filename: str):
    """
    Xóa toàn bộ dữ liệu cũ thuộc về file này trong MySQL trước khi ghi dữ liệu mới.
    """
    inspector = inspect(db_engine)
    # Bỏ qua nếu bảng chưa tồn tại (lần chạy đầu tiên)
    if not inspector.has_table(table_name):
        return

    with db_engine.begin() as conn:
        result = conn.execute(
            text(f"DELETE FROM `{table_name}` WHERE `source_filename` = :fn"),
            {"fn": filename}
        )
        deleted_count = result.rowcount
        if deleted_count > 0:
            print(f"   -> Đã xóa {deleted_count} dòng dữ liệu cũ của file '{filename}'.")


def sync_table_schema(db_engine, table_name: str, df: pd.DataFrame):
    """Tự động thêm cột mới (ALTER TABLE) vào MySQL nếu file chứa cột chưa có."""
    inspector = inspect(db_engine)
    if not inspector.has_table(table_name):
        return

    existing_columns = [col['name'] for col in inspector.get_columns(table_name)]
    missing_columns = [col for col in df.columns if col not in existing_columns]
    
    if missing_columns:
        with db_engine.begin() as conn:
            for col in missing_columns:
                print(f"   -> Phát hiện cột mới '{col}'. Đang thêm vào bảng '{table_name}'...")
                conn.execute(text(f'ALTER TABLE `{table_name}` ADD COLUMN `{col}` TEXT NULL;'))


def import_all_files():
    folder = Path(FOLDER_PATH)
    files = list(folder.glob("*.csv")) + list(folder.glob("*.xlsx")) + list(folder.glob("*.xls"))

    if not files:
        print(f"Không tìm thấy file nào trong: {FOLDER_PATH}")
        return

    for file in files:
        try:
            print(f"Đang xử lý: {file.name}...")
            
            # 1. Đọc dữ liệu từ file
            if file.suffix.lower() == ".csv":
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)
                
            if df.empty:
                print(f"   -> File rỗng, bỏ qua.\n")
                continue

            # 2. Chuẩn hóa tên cột & Thêm thông tin metadata
            df.columns = [sanitize_column_name(c) for c in df.columns]
            df["source_filename"] = file.name
            df["imported_at"] = datetime.now()

            # 3. Đồng bộ cấu trúc bảng (thêm cột mới nếu cần)
            sync_table_schema(engine, SINGLE_TABLE_NAME, df)

            # 4. CÁCH 2: Xóa dữ liệu cũ của file này trong MySQL nếu đã tồn tại
            delete_existing_file_records(engine, SINGLE_TABLE_NAME, file.name)

            # 5. Ghi dữ liệu mới vào bảng master
            df.to_sql(
                name=SINGLE_TABLE_NAME,
                con=engine,
                if_exists="append",
                index=False,
                chunksize=1000
            )
            print(f"   -> Đã ghi {len(df)} dòng mới vào bảng '{SINGLE_TABLE_NAME}' thành công.\n")

        except Exception as e:
            print(f" Lỗi khi import file {file.name}: {e}\n")


if __name__ == "__main__":
    if not os.path.exists(FOLDER_PATH):
        os.makedirs(FOLDER_PATH)
        print(f"Đã tạo thư mục '{FOLDER_PATH}'. Hãy copy các file dữ liệu vào đây và chạy lại.")
    else:
        import_all_files()
