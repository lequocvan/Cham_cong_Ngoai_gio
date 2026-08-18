import os
import re
import sys
import uuid
import shutil
import tempfile
import traceback
import gc
import logging
from datetime import datetime
from urllib.parse import urlparse, unquote
import pandas as pd
import pymysql
from dotenv import load_dotenv
from sqlalchemy import create_engine, text

# ==============================================================================
# 1. CẤU HÌNH BIẾN MÔI TRƯỜNG & DANH SÁCH 17 CỘT CHUẨN (CẬP NHẬT THÁNG 05/2026)
# ==============================================================================
load_dotenv()
db_uri = os.environ.get("SQLALCHEMY_BINDS_BC48")

# Danh sách 17 cột nghiệp vụ chuẩn phân hệ Error CSV cập nhật mới nhất tính đến 05/2026
CSV_ERROR_COLUMNS = [
    "TRANG_THAI", "MA_NGANHANG", "TEN_NGANHANG", "NGAY_BAOCAO",
    "TEN_FILE", "DONG_TIEUDE", "LOAI_BAOCAO", "HINHTHUC_GUI",
    "SOLAN_GUI", "MA_LOI", "MA_GIAO_DICH", "MOTA_LOI",
    "DONG_LOI", "GIAODICH_LOI", "NGAY_GUI", "YEU_CAU", "GHI_CHU"
]

VALID_MODULES = {"CTR", "DWT", "EFT", "PTR"}
COLLATE_STANDARD = "utf8mb4_general_ci"

# Định nghĩa tên tệp nhật ký vật lý riêng cho phân hệ CSV ngoài Database
CSV_LOG_FILE = "csv_import_history.log"

# ==============================================================================
# 1B. CẤU HÌNH HỆ THỐNG GHI LOG RA FILE VẬT LÝ VÀ CONSOLE
# ==============================================================================
logger = logging.getLogger("AML_CSV_Batch")
logger.setLevel(logging.INFO)

# Định dạng dòng log vật lý chuẩn hóa: Thời gian - Cấp độ - Thông điệp
log_formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

# Handler ghi log ra file (Tự động append dữ liệu liên tục)
file_handler = logging.FileHandler(CSV_LOG_FILE, encoding='utf-8')
file_handler.setFormatter(log_formatter)
logger.addHandler(file_handler)

# Handler xuất log ra màn hình Console để tiện theo dõi trực tiếp
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(log_formatter)
logger.addHandler(console_handler)

# ==============================================================================
# 2. CÁC HÀM TRỢ GIÚP HỆ THỐNG VÀ LOGGING KÉP
# ==============================================================================
def get_bc48_engine(uri):
    """Khởi tạo SQLAlchemy Engine hỗ trợ LOAD DATA LOCAL INFILE tối ưu tốc độ"""
    if not uri:
        raise ValueError("Không tìm thấy cấu hình SQLALCHEMY_BINDS_BC48 trong file .env")
    return create_engine(uri, connect_args={"local_infile": True})

def log_to_system(conn, *, file_name, loai, trang_thai, ma_nv_import, ngay_bc=None, ghi_chu=None, so_dong=0):
    """
    HỆ THỐNG LOGGING KÉP ĐỒNG BỘ:
    1. Ghi vết vào bảng `logs_nap_csv` trong database để quản trị trên giao diện Web.
    2. Ghi vết chi tiết ra file vật lý riêng `csv_import_history.log`.
    """
    # Bước 1: Ghi vào DB
    try:
        sql = text("""
            INSERT INTO logs_nap_csv (file_name, loai_bc, ngay_baocao, trang_thai, user_import, ghi_chu, so_dong_du_lieu_csv) 
            VALUES (:fname, :loai, :ngay, :status, :ma_nv, :note, :so_dong)
        """)
        conn.execute(sql, {
            "fname": file_name, 
            "loai": loai,
            "ngay": ngay_bc, 
            "status": trang_thai, 
            "ma_nv": ma_nv_import, 
            "note": ghi_chu,
            "so_dong": so_dong
        })
    except Exception as db_err:
        logger.error(f"Thất bại khi ghi log vào Database cho file {file_name}: {str(db_err)}")

    # Bước 2: Ghi vào tệp tin vật lý riêng biệt ngoài DB
    log_msg = f"File: {file_name} | Phân hệ: {loai} | Ngày BC: {ngay_bc or 'N/A'} | Trạng thái: {trang_thai} | Số dòng: {so_dong} | Người thực hiện: {ma_nv_import} | Ghi chú: {ghi_chu}"
    
    if trang_thai in ['SUCCESS', 'EMPTY_REPORT']:
        logger.info(log_msg)
    elif trang_thai in ['QUARANTINED', 'INTEGRITY_ERROR']:
        logger.warning(log_msg)
    else:
        logger.error(log_msg)

def move_to_quarantine(file_path, filename, reason, folder_path):
    """Cách ly các file có lỗi cấu trúc vật lý hoặc không thể phân tích cứu vãn"""
    quarantine_dir = os.path.join(folder_path, 'quarantine')
    os.makedirs(quarantine_dir, exist_ok=True)
    dest_path = os.path.join(quarantine_dir, f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{filename}")
    try:
        shutil.copy(file_path, dest_path)
        logger.warning(f"-> ĐÃ TIẾN HÀNH CÁCH LY FILE: {filename} | Lý do: {reason}")
    except Exception as e:
        logger.error(f"Không thể sao chép tệp tin qua vùng cách ly: {str(e)}")

def verify_integrity(conn, table_name, file_name, expected_count):
    """Kiểm tra toàn vẹn đối chiếu chéo số dòng nạp thực tế trong DB"""
    result = conn.execute(
        text(f"SELECT COUNT(*) FROM `{table_name}` WHERE `ten_file_goc` = :fname"), 
        {"fname": file_name}
    ).scalar()
    return result == expected_count

# ==============================================================================
# 3. LUỒNG XỬ LÝ CHUNK VÀ STAGING (LOAD DATA LOCAL INFILE)
# ==============================================================================
def load_to_staging(conn, df, target_table):
    """Sử dụng cơ chế Staging trung gian nạp tốc độ cao và an toàn cấu trúc"""
    if df.empty:
        return

    unique_id = uuid.uuid4().hex[:8]
    staging_table = f"temp_staging_{target_table}_{unique_id}"
    conn.execute(text(f"CREATE TABLE `{staging_table}` LIKE `{target_table}`"))

    tf_path = None
    try:
        with tempfile.NamedTemporaryFile(mode='w+', delete=False, suffix='.csv', encoding='utf-8', newline='') as tf:
            df.to_csv(tf, index=False, header=False, encoding='utf-8', quoting=1, lineterminator='\r\n')
            tf_path = tf.name.replace('\\', '/')
            
        conn.execute(text(f"""
            LOAD DATA LOCAL INFILE '{tf_path}' 
            INTO TABLE `{staging_table}` 
            CHARACTER SET utf8mb4
            FIELDS TERMINATED BY ',' 
            ENCLOSED BY '"' 
            ESCAPED BY '"'
            LINES TERMINATED BY '\\r\\n'
        """))
        
        conn.execute(text(f"INSERT INTO `{target_table}` SELECT * FROM `{staging_table}`"))
        
    finally:
        try:
            conn.execute(text(f"DROP TABLE IF EXISTS `{staging_table}`"))
        except Exception:
            pass
        if tf_path and os.path.exists(tf_path):
            try:
                os.remove(tf_path)
            except Exception:
                pass

# ==============================================================================
# 4. HÀM XỬ LÝ LÕI FILE CSV (CORE LOGIC)
# ==============================================================================
def process_csv_core(engine, file_path, filename):
    """
    Đọc file bằng cơ chế Chunksize giúp tối ưu bộ nhớ.
    Tự động Auto-healing cấu trúc và gom nhóm theo LOAI_BAOCAO + Tháng của NGAY_BAOCAO.
    """
    pure_filename = os.path.basename(filename)
    col_defs = ", ".join([f"`{col}` TEXT" for col in CSV_ERROR_COLUMNS])
    full_col_defs = f"{col_defs}, `ten_file_goc` VARCHAR(150) DEFAULT NULL"

    try:
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            first_line = f.readline()
            if not first_line or first_line.strip() == "":
                return None, 0, "FILE_EMPTY", None
    except Exception as e:
        return None, 0, f"CANNOT_READ_PHYSICAL: {str(e)}", None

    # Tạo dữ liệu dự phòng từ tên file
    backup_loai = "UNKNOWN"
    match_loai = re.search(r'(ctr|dwt|eft|ptr)', pure_filename, re.IGNORECASE)
    if match_loai:
        backup_loai = match_loai.group(1).upper()

    backup_ngay = None
    match_date = re.search(r'\b(20\d{4,6})\b|_(\d{6,8})_', pure_filename)
    if match_date:
        raw_date = match_date.group(1) if match_date.group(1) else match_date.group(2)
        backup_ngay = f"{raw_date[:6]}01" if len(raw_date) == 6 else raw_date[:8]

    try:
        reader = pd.read_csv(file_path, dtype=str, skipinitialspace=True, keep_default_na=True, chunksize=50000)
    except Exception as e:
        return None, 0, f"PARSING_ERROR: {str(e)}", None

    total_rows = 0
    inserted_metadata = {}
    tables_cleared = set()

    for chunk in reader:
        if chunk.empty:
            continue
            
        chunk.columns = [col.strip().upper() for col in chunk.columns]
        chunk = chunk.dropna(how='all')
        if chunk.empty:
            continue

        # --- AUTO-HEALING CẤU TRÚC THÁNG 05/2026 ---
        chunk = chunk.reindex(columns=CSV_ERROR_COLUMNS)
        chunk = chunk.fillna("N/A")
        chunk['ten_file_goc'] = pure_filename

        # Ép kiểu về chuỗi an toàn trước khi xử lý cắt chuỗi
        chunk['NGAY_BAOCAO'] = chunk['NGAY_BAOCAO'].astype(str).apply(lambda x: re.sub(r'\D', '', x))

        # SỬA TẠI ĐÂY: Thêm .str vào trước upper() để xử lý chuẩn Pandas Series
        chunk['LOAI_BAOCAO'] = chunk['LOAI_BAOCAO'].astype(str).str.strip().str.upper()

        chunk = chunk[chunk['LOAI_BAOCAO'].isin(VALID_MODULES) & chunk['NGAY_BAOCAO'].str.match(r'^\d{8}$')]
        if chunk.empty:
            continue

        grouped = chunk.groupby(['LOAI_BAOCAO', 'NGAY_BAOCAO'])
        
        with engine.connect() as conn:
            for (loai, ngay_bc), group in grouped:
                loai_str = str(loai).strip().upper()
                ngay_str = str(ngay_bc).strip()
                month = ngay_str[:6]
                table_name = f"{loai_str.lower()}_{month}_error"

                conn.execute(text(f"""
                    CREATE TABLE IF NOT EXISTS `{table_name}` (
                        {full_col_defs},
                        INDEX idx_tfg (`ten_file_goc`(50)),
                        INDEX idx_mgd (`MA_GIAO_DICH`(30))
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE={COLLATE_STANDARD}
                """))
                
                try:
                    conn.execute(text(f"ALTER TABLE `{table_name}` ADD COLUMN `ten_file_goc` VARCHAR(150) DEFAULT NULL"))
                except Exception:
                    pass

                if table_name not in tables_cleared:
                    conn.execute(text(f"DELETE FROM `{table_name}` WHERE `ten_file_goc` = :fname"), {"fname": pure_filename})
                    tables_cleared.add(table_name)

                load_to_staging(conn, group, table_name)

                log_key = (table_name, ngay_str)
                if log_key not in inserted_metadata:
                    inserted_metadata[log_key] = 0
                inserted_metadata[log_key] += len(group)
                
            conn.commit()
        total_rows += len(chunk)

    if total_rows == 0 or not inserted_metadata:
        return {}, 0, None, (backup_loai, backup_ngay)

    formatted_meta = {}
    total_actual_rows = 0
    for (t_name, n_bc), count in inserted_metadata.items():
        formatted_meta[(t_name, n_bc)] = {
            "loai": t_name.split('_')[0].upper(),
            "ngay": n_bc,
            "so_dong": count
        }
        total_actual_rows += count

    return formatted_meta, total_actual_rows, None, None

# ==============================================================================
# 5. HÀM ĐIỀU PHỐI QUÉT THƯ MỤC CHẠY BATCH HÀNG LOẠT
# ==============================================================================
def process_directory_csv_import_batch(folder_path, current_user_name="System_Auto"):
    """Quét thư mục, phân loại nghiệp vụ, kiểm tra toàn vẹn và thực hiện log kép"""
    if not os.path.isdir(folder_path):
        logger.error(f"Đường dẫn thư mục quét dữ liệu không tồn tại: {folder_path}")
        return

    archive_path = os.path.join(folder_path, 'archive')
    os.makedirs(archive_path, exist_ok=True)
    
    csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]
    total_files = len(csv_files)
    
    if not csv_files:
        logger.info("Thư mục trống, không tìm thấy file .CSV báo cáo lỗi nào cần nạp.")
        return

    logger.info("=" * 70)
    logger.info(f"HỆ THỐNG AML PHÁT HIỆN: {total_files} TỆP TIN CSV CẦN NẠP BATCH.")
    logger.info("=" * 70)

    engine = get_bc48_engine(db_uri)
    results = {"success": [], "error": []}

    for idx, filename in enumerate(csv_files, 1):
        pure_filename = os.path.basename(filename)
        file_path = os.path.join(folder_path, pure_filename)
        logger.info(f"[{idx}/{total_files}] Bắt đầu phân tích tệp tin: {pure_filename}")
        
        is_success_flag = False
        
        try:
            logs_meta, total_count, err_type, backup_info = process_csv_core(engine, file_path, pure_filename)

            # Trường hợp 1: File hỏng cấu trúc vật lý hoặc file trống rỗng rơm
            if err_type:
                move_to_quarantine(file_path, pure_filename, err_type, folder_path)
                with engine.connect() as conn:
                    log_to_system(
                        conn=conn, file_name=pure_filename, loai="NONE", trang_thai='QUARANTINED',
                        ma_nv_import=current_user_name, ghi_chu=f"Lỗi cấu trúc hoặc file trống: {err_type}"
                    )
                    conn.commit()
                results["error"].append(f"{pure_filename}: Lỗi cấu trúc ({err_type})")
                if os.path.exists(file_path): os.remove(file_path)
                continue

            with engine.connect() as conn:
                # Trường hợp 2: Nội dung trống rỗng sau sàng lọc (Chỉ có dòng tiêu đề Header)
                if total_count == 0 and backup_info:
                    b_loai, b_ngay = backup_info
                    log_to_system(
                        conn=conn, file_name=pure_filename, loai=b_loai, trang_thai='EMPTY_REPORT',
                        ma_nv_import=current_user_name, ngay_bc=b_ngay, so_dong=0, ghi_chu="File chỉ chứa dòng tiêu đề, trống dữ liệu nghiệp vụ"
                    )
                    is_success_flag = True
                    results["success"].append(f"{pure_filename} (Empty)")
                
                # Trường hợp 3: Nạp thành công dữ liệu phân đoạn -> Đối chiếu chéo toàn vẹn
                else:
                    rows_per_table = {}
                    for (t_name, _), info in logs_meta.items():
                        rows_per_table[t_name] = rows_per_table.get(t_name, 0) + info["so_dong"]

                    all_tables_verified = True
                    for t_name, expected_count in rows_per_table.items():
                        if not verify_integrity(conn, t_name, pure_filename, expected_count):
                            all_tables_verified = False
                            break
                    
                    if all_tables_verified:
                        # ĐỘT PHÁ LOG ĐA DÒNG KÉP: Có bao nhiêu ngày/phân hệ báo cáo, ghi bấy nhiêu vết riêng biệt
                        for (t_name, ngay_bc_log), info in logs_meta.items():
                            log_to_system(
                                conn=conn, file_name=pure_filename, loai=info["loai"], trang_thai='SUCCESS',
                                ma_nv_import=current_user_name, ngay_bc=ngay_bc_log, so_dong=info["so_dong"],
                                ghi_chu=f"Đã tự động cấu trúc lại và nạp thành công vào bảng mục tiêu `{t_name}`"
                            )
                        is_success_flag = True
                        results["success"].append(pure_filename)
                    else:
                        err_msg = "Lỗi toàn vẹn dữ liệu: Số lượng dòng nạp thực tế không khớp dữ liệu phân tích"
                        move_to_quarantine(file_path, pure_filename, err_msg, folder_path)
                        log_to_system(
                            conn=conn, file_name=pure_filename, loai="MULTI", trang_thai='INTEGRITY_ERROR',
                            ma_nv_import=current_user_name, ghi_chu=f"{err_msg} (Quarantined)", so_dong=sum(rows_per_table.values())
                        )
                        results["error"].append(f"{pure_filename}: Lỗi kiểm toán toàn vẹn")
                        if os.path.exists(file_path): os.remove(file_path)
                
                conn.commit()

        except Exception as e:
            logger.error(f"Lỗi hệ thống nghiêm trọng tại tệp tin {pure_filename}: {str(e)}")
            logger.error(traceback.format_exc())
            try:
                loai_fail = "UNKNOWN"
                match_loai_fail = re.search(r'(ctr|dwt|eft|ptr)', pure_filename, re.IGNORECASE)
                if match_loai_fail: loai_fail = match_loai_fail.group(1).upper()

                with engine.connect() as conn:
                    log_to_system(
                        conn=conn, file_name=pure_filename, loai=loai_fail, trang_thai='FAILED',
                        ma_nv_import=current_user_name, ghi_chu=f"Lỗi Runtime: {str(e)[:240]}", so_dong=0
                    )
                    conn.commit()
            except Exception as log_err:
                logger.error(f"Không thể cập nhật trạng thái FAILED vào hệ thống log chéo: {str(log_err)}")
            
            results["error"].append(f"{pure_filename}: {str(e)}")

        # Lưu trữ tệp tin khi luồng đọc chính thức khép lại hoàn toàn
        if is_success_flag and os.path.exists(file_path):
            try:
                shutil.move(file_path, os.path.join(archive_path, pure_filename))
                logger.info(f"Chuyển kho lưu trữ thành công: {pure_filename} -> /archive/")
            except Exception as move_err:
                logger.error(f"Lỗi File Lock hệ điều hành, không thể di chuyển {pure_filename} sang mục lưu trữ: {str(move_err)}")

        gc.collect()

    logger.info("=" * 70)
    logger.info(f"HOÀN THÀNH TIẾN TRÌNH BATCH CSV. THÀNH CÔNG: {len(results['success'])} | THẤT BẠI: {len(results['error'])}")
    logger.info("=" * 70)
    return results

# ==============================================================================
# 6. KHỐI KHỞI CHẠY CHƯƠNG TRÌNH ĐỘC LẬP
# ==============================================================================
if __name__ == "__main__":
    if os.name == 'posix':  
        PATH_THU_MUC = os.path.expanduser("/Users/lequocvan/Downloads/Cham_cong_Ngoai_gio/uploads_manual")
    else:  
        PATH_THU_MUC = r"C:\CSV_Error_Batch_2026"

    if not os.path.exists(PATH_THU_MUC):
        os.makedirs(PATH_THU_MUC)
        logger.info(f"Đã khởi tạo thư mục đích tự động quét tệp tin CSV: {PATH_THU_MUC}")

    # Chạy nạp Batch hàng loạt
    process_directory_csv_import_batch(PATH_THU_MUC, current_user_name="System_Batch")
