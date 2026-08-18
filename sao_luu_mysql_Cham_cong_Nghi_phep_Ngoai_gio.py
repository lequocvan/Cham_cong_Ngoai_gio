import os
import subprocess
import shutil
import datetime
import zipfile
from urllib.parse import urlparse # Dùng để tách DATABASE_URL
from dotenv import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.service_account import Credentials

load_dotenv()

# --- TRÍCH XUẤT THÔNG TIN TỪ DATABASE_URL ---
db_url = os.getenv("DATABASE_URL")
parsed_url = urlparse(db_url)

# parsed_url.netloc có dạng 'user:password@host:port'
DB_USER = parsed_url.username
DB_PASS = parsed_url.password
DB_HOST = parsed_url.hostname
DB_PORT = parsed_url.port or 3306
# parsed_url.path có dạng '/db_name', cần bỏ dấu '/' ở đầu
DB_NAME = parsed_url.path.lstrip('/')

# Các thông số khác
LAN_DESTINATION = os.getenv("LAN_DESTINATION")
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE")

def backup_mysql():
    now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    file_sql = f"backup_{DB_NAME}_{now}.sql"
    file_zip = f"{file_sql}.zip"

    try:
        # 1. Tạo bản sao 1: MySQL Dump
        print(f"--- Đang dump dữ liệu từ host: {DB_HOST} ---")
        
        dump_cmd = [
            "mysqldump",
            f"--host={DB_HOST}",
            f"--port={DB_PORT}",
            f"--user={DB_USER}",
            f"--password={DB_PASS}",
            DB_NAME
        ]
        
        with open(file_sql, "w", encoding="utf-8") as f:
            subprocess.run(dump_cmd, stdout=f, check=True)

        # Nén file
        print(f"--- Đang nén file: {file_zip} ---")
        with zipfile.ZipFile(file_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(file_sql)

        # 2. Tạo bản sao 2: Copy qua LAN
        print(f"--- Đang sao chép qua LAN ---")
        if os.path.exists(LAN_DESTINATION):
            shutil.copy2(file_zip, os.path.join(LAN_DESTINATION, file_zip))
        else:
            print(f"Cảnh báo: Không kết nối được LAN {LAN_DESTINATION}")

        # 3. Tạo bản sao 3: Đẩy lên Google Drive
        print("--- Đang tải lên Google Drive ---")
        upload_to_drive(file_zip)

        # Dọn dẹp
        if os.path.exists(file_sql): os.remove(file_sql)
        if os.path.exists(file_zip): os.remove(file_zip)
        print(">>> HOÀN TẤT QUY TRÌNH 3-2-1 <<<")

    except Exception as e:
        print(f"Lỗi: {e}")

def upload_to_drive(file_path):
    # (Hàm upload_to_drive giữ nguyên như các câu trả lời trước)
    try:
        creds = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, 
            scopes=['https://www.googleapis.com/auth/drive.file']
        )
        service = build('drive', 'v3', credentials=creds)
        file_metadata = {'name': os.path.basename(file_path), 'parents': [DRIVE_FOLDER_ID]}
        media = MediaFileUpload(file_path, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(f"Drive File ID: {file.get('id')}")
    except Exception as e:
        print(f"Lỗi Drive: {e}")

if __name__ == "__main__":
    backup_mysql()
