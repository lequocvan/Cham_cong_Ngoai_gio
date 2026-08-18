import os
import re
import pymysql
import pymysql.cursors
from datetime import datetime
from urllib.parse import urlparse, unquote
from dotenv import load_dotenv

# ==============================================================================
# 1. CẤU HÌNH BIẾN MÔI TRƯỜNG & ĐỊNH DẠNG CỘT CHUẨN CỦA 4 PHÂN HỆ AML
# ==============================================================================
load_dotenv()
db_uri = os.environ.get("SQLALCHEMY_BINDS_BC48")

# Định nghĩa file log vật lý riêng cho các file bị từ chối nạp
REJECTED_LOG_FILE = "rejected_files.log"

FILE_STRUCTURES = {
    "CTR": [("macn", "TEXT"), ("thoidiem", "TEXT"), ("loaigd", "TEXT"), ("magd", "TEXT"), ("loaitien", "TEXT"), ("sotien", "TEXT"), ("quydoi", "TEXT"), ("lydomucdich", "TEXT"), ("kieukh", "TEXT"), ("tenkh", "TEXT"), ("ten_eng", "TEXT"), ("quoctich", "TEXT"), ("diachitt", "TEXT"), ("noioht", "TEXT"), ("ngaysinh", "TEXT"), ("loaigto", "TEXT"), ("sogt", "TEXT"), ("sothithuc", "TEXT"), ("so_dt", "TEXT"), ("lq_kieukh", "TEXT"), ("lq_tenkh", "TEXT"), ("lq_tenta", "TEXT"), ("lq_quoctich", "TEXT"), ("lq_diachitt", "TEXT"), ("lq_noioht", "TEXT"), ("lq_ngaysinh", "TEXT"), ("lq_loaigto", "TEXT"), ("lq_sogto", "TEXT"), ("lq_sothithuc", "TEXT"), ("lq_so_dt", "TEXT"), ("bd_manh", "TEXT"), ("bd_sotk", "TEXT"), ("bd_tentk", "TEXT"), ("bd_loaitientk", "TEXT"), ("bd_ngaymotk", "TEXT"), ("bq_loaitk", "TEXT"), ("bd_status_tk", "TEXT")],
    "DWT": [("macn", "TEXT"), ("thoidiem", "TEXT"), ("loaigd", "TEXT"), ("kenhct", "TEXT"), ("magd", "TEXT"), ("thamchieu", "TEXT"), ("loaitien", "TEXT"), ("sotien", "TEXT"), ("quydoi", "TEXT"), ("lydomucdich", "TEXT"), ("kieukh", "TEXT"), ("tenkh", "TEXT"), ("ten_ta", "TEXT"), ("quoctich", "TEXT"), ("diachitt", "TEXT"), ("noioht", "TEXT"), ("ngaysinh", "TEXT"), ("loaigt", "TEXT"), ("sogt", "TEXT"), ("so_thithuc", "TEXT"), ("sodt", "TEXT"), ("ma_nh", "TEXT"), ("sotk", "TEXT"), ("tentk", "TEXT"), ("loaitientk", "TEXT"), ("ngaymotk", "TEXT"), ("loaitk", "TEXT"), ("status_tk", "TEXT"), ("lq_manh", "TEXT"), ("lq_macn", "TEXT"), ("lq_sotk", "TEXT"), ("lq_tentk", "TEXT"), ("lq_loaigt", "TEXT"), ("lq_sogt", "TEXT"), ("lq_tenkh", "TEXT")],
    "EFT": [("macn", "TEXT"), ("loaigd", "TEXT"), ("kenhct", "TEXT"), ("thoidiem", "TEXT"), ("magd", "TEXT"), ("thamchieu", "TEXT"), ("loaitien", "TEXT"), ("sotien", "TEXT"), ("quydoi", "TEXT"), ("lydomucdich", "TEXT"), ("kieukh", "TEXT"), ("tenkh", "TEXT"), ("ten_ta", "TEXT"), ("quoctich", "TEXT"), ("diachitt", "TEXT"), ("noiott", "TEXT"), ("ngaysinh", "TEXT"), ("loaigt", "TEXT"), ("sogt", "TEXT"), ("so_thithuc", "TEXT"), ("sodt", "TEXT"), ("ma_nh", "TEXT"), ("sotk", "TEXT"), ("tentk", "TEXT"), ("loaitientk", "TEXT"), ("ngaymotk", "TEXT"), ("loaitk", "TEXT"), ("status_tk", "TEXT"), ("td_manh", "TEXT"), ("td_ma_sw", "TEXT"), ("td_ten", "TEXT"), ("td_diachi", "TEXT"), ("td_tinh", "TEXT"), ("td_quocgia", "TEXT"), ("doiung_matc", "TEXT"), ("doiung_tentc", "TEXT"), ("doiung_diachi", "TEXT"), ("doiung_tinh", "TEXT"), ("doiung_quocgia", "TEXT"), ("khdu_tenkh", "TEXT"), ("khdu_ngaysinh", "TEXT"), ("khdu_sogiayto", "TEXT"), ("khdu_diachi", "TEXT"), ("khdu_quocgia", "TEXT"), ("khdu_sotk", "TEXT"), ("khdu_tentk", "TEXT")],
    "PTR": [("macn", "TEXT"), ("magd", "TEXT"), ("thoidiem", "TEXT"), ("kyhieumb", "TEXT"), ("loaihanghoa", "TEXT"), ("soluong_donvi", "TEXT"), ("loaitien", "TEXT"), ("sotien", "TEXT"), ("quydoi", "TEXT"), ("lydomucdich", "TEXT"), ("diadiem", "TEXT"), ("kieukh", "TEXT"), ("tenkh", "TEXT"), ("ten_ta", "TEXT"), ("quoctich", "TEXT"), ("diachitt", "TEXT"), ("noioht", "TEXT"), ("ngaysinh", "TEXT"), ("loaigt", "TEXT"), ("sogt", "TEXT"), ("so_thithuc", "TEXT"), ("sodt", "TEXT"), ("bc_macn", "TEXT"), ("bc_tentc", "TEXT"), ("bc_tenta", "TEXT"), ("bc_quocgia", "TEXT"), ("bc_diachi", "TEXT"), ("bc_loaigiayto", "TEXT"), ("bc_sogt", "TEXT"), ("bc_sodt", "TEXT")]
}

DON_VI_CACHE = {}

# ==============================================================================
# 2. HÀM BỔ TRỢ HỆ THỐNG
# ==============================================================================
def log_to_file(filename, error_code, details):
    """ Ghi vết các file bị từ chối nạp ra file log riêng biệt """
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_message = f"[{now_str}] FILE: {filename} | ERROR: {error_code} | DETAILS: {details}\n"
    try:
        with open(REJECTED_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(log_message)
    except Exception as le:
        print(f"[!] Không thể ghi log ra file riêng: {le}")

def parse_mysql_uri_standard(uri):
    """ Hàm bóc tách chuỗi kết nối chuẩn sang Dict dùng cho PyMySQL """
    if not uri:
        raise ValueError("Không tìm thấy cấu hình SQLALCHEMY_BINDS_BC48 trong file .env")
    schemeless_uri = uri.split("://", 1)[1] if "://" in uri else uri
    parsed = urlparse(f"http://{schemeless_uri}")
    return {
        'host': parsed.hostname,
        'port': parsed.port if parsed.port else 3306,
        'user': unquote(parsed.username) if parsed.username else "",
        'password': unquote(parsed.password) if parsed.password else "",
        'database': parsed.path.lstrip('/'),
        'charset': 'utf8mb4',
        'cursorclass': pymysql.cursors.DictCursor
    }

def validate_filename(filename):
    """ Kiểm tra định dạng tên file nghiêm ngặt (Lỗi cấu trúc RE2) """
    fn_upper = filename.upper()
    if not fn_upper.endswith('.TXT'):
        return False, "RE2_loi_chi_tiet_6: File phải có đuôi .TXT"

    parts = fn_upper.replace('.TXT', '').split('_')
    if len(parts) != 5:
        return False, f"RE2_loi_chi_tiet_0: Cấu trúc tên file sai (nhận được {len(parts)} phần)."
    
    ma_nh, ngay_bc, loai_bc, hinh_thuc, stt = parts
    
    if ma_nh != "01204001":
        return False, f"RE2_loi_chi_tiet_1: Mã ngân hàng {ma_nh} không hợp lệ."
    if not re.match(r'^\d{8}$', ngay_bc):
        return False, "RE2_loi_chi_tiet_2: Ngày báo cáo sai định dạng yyyymmdd."
    try:
        datetime.strptime(ngay_bc, "%Y%m%d")
    except ValueError:
        return False, f"RE2_loi_chi_tiet_2: Ngày báo cáo '{ngay_bc}' không tồn tại trên lịch."
    
    if loai_bc not in ["CTR", "DWT", "EFT", "PTR"]:
        return False, f"RE2_loi_chi_tiet_3: Loại báo cáo {loai_bc} không hợp lệ."
    if hinh_thuc not in ["GLD", "GLA", "GBS"]:
        return False, f"RE2_loi_chi_tiet_4: Hình thức gửi {hinh_thuc} không hợp lệ."
    if not re.match(r'^\d{3}$', stt):
        return False, f"RE2_loi_chi_tiet_5: STT {stt} phải là 3 chữ số."
        
    return True, None

def get_related_branches(cursor, header_macn):
    """ Lấy cây đơn vị liên đới áp dụng cơ chế Cache """
    if not header_macn:
        return set()
    
    header_macn = header_macn.strip()
    if header_macn in DON_VI_CACHE:
        return DON_VI_CACHE[header_macn]
        
    branches = {header_macn}
    try:
        query_dv = """
            SELECT DISTINCT TRIM(`MaNH8so_moi`) as macn_con
            FROM `don_vi`
            WHERE `ma_hieu_1` = (
                SELECT `ma_hieu_1` FROM `don_vi` WHERE TRIM(`MaNH8so_moi`) = %s LIMIT 1
            ) 
            AND `MaNH8so_moi` IS NOT NULL AND `MaNH8so_moi` != '';
        """
        cursor.execute(query_dv, (header_macn,))
        for r in cursor.fetchall():
            if r['macn_con']:
                branches.add(r['macn_con'])
    except Exception as edv:
        print(f"[!] Cảnh báo kiểm tra liên đới cây đơn vị cho {header_macn}: {edv}")
        
    DON_VI_CACHE[header_macn] = branches
    return branches

def run_post_import_logic_check(cursor, table_name, target_month, filename):
    """ Kiểm tra đối chiếu logic ngày phát sinh giao dịch hậu kỳ """
    file_date = filename.split('_')[1] 
    
    query = f"""
        INSERT IGNORE INTO log_loi_logic_ngay_thang 
        (table_name, thoidiem_loi, ten_file_loi, error_message)
        SELECT DISTINCT %s, CAST(thoidiem AS CHAR), ten_file_goc, 
               CASE 
                   WHEN LEFT(TRIM(thoidiem), 6) <> %s THEN 'thoidiem khong thuoc thang bao cao (TXT dung ky bao cao)'
                   WHEN LEFT(TRIM(thoidiem), 8) <> %s THEN 'thoidiem khong khop voi ngay trong ten file (TXT khop ngay phat sinh)'
                   WHEN LEFT(ten_file_goc, 15) NOT LIKE CONCAT('%%', %s, '%%') THEN 'File khong thuoc thang bao cao'
                   ELSE 'Loi logic ngay thang khong xac dinh'
               END
        FROM `{table_name}`
        WHERE LEFT(TRIM(thoidiem), 6) <> %s 
           OR LEFT(TRIM(thoidiem), 8) <> %s
           OR ten_file_goc NOT LIKE CONCAT('%%', %s, '%%')
    """
    cursor.execute(query, (table_name, target_month, file_date, target_month, target_month, file_date, target_month))

# ==============================================================================
# 3. LOGIC XỬ LÝ CHÍNH CHO TỪNG TỆP TIN
# ==============================================================================
def process_single_file(connection, cursor, folder_path, filename, current_user_name):
    """ Đọc, kiểm định lỗi nghiêm ngặt và đẩy batch dữ liệu vào MySQL """
    is_valid, error_detail = validate_filename(filename)
    
    if not is_valid:
        sql_err = """
            INSERT INTO `log_import_errors` 
            (file_name, ma_loi_bc48, header_content, user_import, status) 
            VALUES (%s, 'RE2', %s, %s, 'ERROR_CRITICAL')
        """
        cursor.execute(sql_err, (filename, error_detail, current_user_name))
        log_to_file(filename, "RE2", error_detail)
        print(f"[-] Từ chối: {filename} -> Lỗi cấu trúc RE2")
        return

    parts = filename.upper().replace('.TXT', '').split('_')
    file_ma_nh, file_date_full, prefix, file_hinh_thuc, file_stt = parts
    table_name = f"{prefix.lower()}_{file_date_full[:6]}"

    current_structure = FILE_STRUCTURES.get(prefix)
    if not current_structure: 
        print(f"[-] Bỏ qua {filename}: Thiếu cấu hình ánh xạ phân hệ {prefix}")
        return

    # Quét chỉ mục cột quan trọng phục vụ bóc tách dữ liệu thô
    try:
        macn_col_idx = [idx for idx, col in enumerate(current_structure) if col[0].lower() == 'macn'][0]
    except IndexError:
        print(f"[-] Bỏ qua {filename}: Phân hệ {prefix} thiếu trường cấu trúc macn")
        return

    try:
        magd_col_idx = [idx for idx, col in enumerate(current_structure) if col[0].lower() == 'magd'][0]
    except IndexError:
        magd_col_idx = None

    # --- ĐỌC VÀ GIẢI MÃ TỆP TIN ---
    file_full_path = os.path.join(folder_path, filename)
    content = None
    with open(file_full_path, 'rb') as f:
        raw_content = f.read()

    for encoding in ['utf-8-sig', 'utf-16', 'utf-8', 'latin-1']:
        try:
            content = raw_content.decode(encoding)
            break
        except: 
            continue
    
    if content is None:
        log_to_file(filename, "DECODE_FAIL", "Không thể xác định bảng mã giải mã file.")
        print(f"[-] Bỏ qua {filename}: Lỗi giải mã dữ liệu nhị phân.")
        return

    all_lines_raw = [l for l in content.splitlines() if l.strip()]
    if not all_lines_raw: 
        sql_err = """
            INSERT INTO `log_import_errors` 
            (file_name, ma_loi_bc48, header_content, user_import, status) 
            VALUES (%s, 'RE3', 'File trống hoàn toàn', %s, 'ERROR_CRITICAL')
        """
        cursor.execute(sql_err, (filename, current_user_name))
        log_to_file(filename, "RE3", "File trống hoàn toàn không có dữ liệu.")
        print(f"[-] Từ chối: {filename} -> Lỗi RE3 (File trống)")
        return

    # --- KIỂM TRA DÒNG TIÊU ĐỀ (RE3 & RE6) ---
    header_line_raw = all_lines_raw[0]
    re3_error = None

    if header_line_raw.startswith(' ') or header_line_raw.endswith(' '):
        re3_error = "Dòng tiêu đề chứa khoảng trắng thừa biên."
    
    header_parts = header_line_raw.split('#')
    if re3_error is None and len(header_parts) != 6:
        re3_error = f"Cấu trúc tiêu đề sai số lượng thành phần ({len(header_parts)}/6)."
    
    if re3_error:
        sql_err = """
            INSERT INTO `log_import_errors` 
            (file_name, ma_loi_bc48, header_content, user_import, status) 
            VALUES (%s, 'RE3', %s, %s, 'ERROR_CRITICAL')
        """
        cursor.execute(sql_err, (filename, header_line_raw, current_user_name))
        log_to_file(filename, "RE3", re3_error)
        print(f"[-] Từ chối: {filename} -> Lỗi RE3 ({re3_error})")
        return

    header_line = header_line_raw.strip()
    header_macn = header_parts[0].strip().upper() 
    header_date = header_parts[1].strip()         
    header_loai_bc = header_parts[2].strip().upper()
    header_hinh_thuc = header_parts[3].strip().upper()
    header_stt = header_parts[4].strip()

    re6_details = []
    if header_macn != file_ma_nh: re6_details.append("Sai biệt mã NH")
    if header_loai_bc != prefix: re6_details.append("Mâu thuẫn loại BC")
    if header_hinh_thuc != file_hinh_thuc: re6_details.append("Sai lệch hình thức gửi")
    if header_stt != file_stt: re6_details.append("Lệch số thứ tự file")
        
    if re6_details:
        re6_msg = "Thông tin tiêu đề không khớp tên file: " + ", ".join(re6_details)
        sql_err = """
            INSERT INTO `log_import_errors` 
            (file_name, ma_loi_bc48, header_content, user_import, status) 
            VALUES (%s, 'RE6', %s, %s, 'ERROR_CRITICAL')
        """
        cursor.execute(sql_err, (filename, f"LỖI RE6: {re6_msg} | Header: {header_line}", current_user_name))
        log_to_file(filename, "RE6", re6_msg)
        print(f"[-] Từ chối: {filename} -> Lỗi khớp dữ liệu RE6")
        return 

    # --- KIỂM TOÁN CHÊNH LỆCH DÒNG VÀ NGÀY (RE4 / RF3) ---
    raw_qty = "".join(filter(str.isdigit, header_parts[5].strip()))
    declared_rows = int(raw_qty) if raw_qty else 0
    data_lines = all_lines_raw[1:]
    actual_rows = len(data_lines)

    final_import_status = "SUCCESS"
    error_code = None
    
    if declared_rows != actual_rows:
        error_code = "RE4" 
        final_import_status = "WARNING_RE4"
    elif header_date != file_date_full:
        error_code = "RF3" 
        final_import_status = "WARNING_RF3"
    
    if error_code:
        sql_err = """
            INSERT INTO `log_import_errors` 
            (file_name, ma_loi_bc48, header_content, declared_rows, actual_rows, user_import, status)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """
        cursor.execute(sql_err, (filename, error_code, header_line, declared_rows, actual_rows, current_user_name, final_import_status))

    # --- THU THẬP DỮ LIỆU THÔ, MÃ CHI NHÁNH & MAGD ĐỘNG (BỔ SUNG CHÍNH) ---
    unique_macns = set()
    unique_magds = set()
    cleaned_data_lines = []

    for line in data_lines:
        line_clean = line.replace('\r', '').replace('\n', '').strip()
        f_split = line_clean.split('#')
        f_cleaned = [item.strip() for item in f_split] 
        cleaned_data_lines.append(f_cleaned)
        
        # Gom cụm chi nhánh phục vụ chiến lược GLD
        if len(f_cleaned) > macn_col_idx and f_cleaned[macn_col_idx]:
            unique_macns.add(f_cleaned[macn_col_idx])
            
        # Gom cụm magd phục vụ chiến lược GLA
        if magd_col_idx is not None and len(f_cleaned) > magd_col_idx and f_cleaned[magd_col_idx]:
            unique_magds.add(f_cleaned[magd_col_idx])

    related_branches = get_related_branches(cursor, header_macn)
    unique_macns.update(related_branches)

    if not unique_macns and header_macn:
        unique_macns.add(header_macn)

    # --- THAO TÁC CƠ SỞ DỮ LIỆU TỐC ĐỘ CAO ---
    cols_sql = ", ".join([f"`{c[0]}` TEXT" for c in current_structure])
    
    # CHUẨN HÓA: Đồng bộ bảng động toàn vẹn về utf8mb4_general_ci theo phân rã dữ liệu ngân hàng
    create_sql = f"""
        CREATE TABLE IF NOT EXISTS `{table_name}` (
            {cols_sql},
            `ten_file_goc` VARCHAR(150) DEFAULT NULL,
            `hinh_thuc_file` VARCHAR(10) DEFAULT NULL,
            INDEX idx_macn (`macn`(20)),
            INDEX idx_td (`thoidiem`(10)),
            INDEX idx_file (`ten_file_goc`(50)),
            INDEX idx_ht (`hinh_thuc_file`(10))
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci
    """
    cursor.execute(create_sql)

    cursor.execute(f"SHOW COLUMNS FROM `{table_name}`")
    existing_cols = {row['Field'].lower() for row in cursor.fetchall()}
    
    if "ten_file_goc" not in existing_cols:
        cursor.execute(f"ALTER TABLE `{table_name}` ADD COLUMN `ten_file_goc` VARCHAR(150) DEFAULT NULL, ADD INDEX idx_file (`ten_file_goc`(50))")
    if "hinh_thuc_file" not in existing_cols:
        cursor.execute(f"ALTER TABLE `{table_name}` ADD COLUMN `hinh_thuc_file` VARCHAR(10) DEFAULT NULL, ADD INDEX idx_ht (`hinh_thuc_file`(10))")

    for col_name, _ in current_structure:
        if col_name.lower() not in existing_cols:
            cursor.execute(f"ALTER TABLE `{table_name}` ADD COLUMN `{col_name}` TEXT")

    # ==========================================================================
    # 3. CHỐNG TRÙNG LẶP DỮ LIỆU THÔNG MINH - RẼ NHÁNH GLA ĐỘNG (MỤC 4B)
    # ==========================================================================
    if file_hinh_thuc == "GBS":
        # Chiến lược GBS: Xóa toàn bộ dữ liệu cũ gắn liền với file gốc này
        delete_sql = f"DELETE FROM `{table_name}` WHERE `ten_file_goc` = %s"
        cursor.execute(delete_sql, (filename,))
        
    elif file_hinh_thuc == "GLA":
        # Chiến lược GLA: Xóa đích danh dựa theo danh sách mã giao dịch (magd) có trong file thô
        if unique_magds:
            format_magd_strings = ','.join(['%s'] * len(unique_magds))
            delete_sql = f"DELETE FROM `{table_name}` WHERE `magd` IN ({format_magd_strings})"
            cursor.execute(delete_sql, list(unique_magds))
        else:
            #print(f"[!] Cảnh báo file GLA {filename} không thu thập được magd hợp lệ. Fallback sang chiến lược GLD rộng.")
            print(f"[!] File GLA {filename} rỗng dữ liệu (Khai báo: {declared_rows} dòng, Thực tế: {actual_rows} dòng). Kích hoạt Fallback.")
            # Fallback an toàn nếu không bóc được magd từ nội dung file
            format_strings = ','.join(['%s'] * len(unique_macns))
            delete_sql = f"""
                DELETE FROM `{table_name}` 
                WHERE LEFT(REPLACE(REPLACE(TRIM(`thoidiem`), '\\r', ''), '\\n', ''), 8) = %s 
                  AND REPLACE(REPLACE(TRIM(`macn`), '\\r', ''), '\\n', '') IN ({format_strings})
            """
            cursor.execute(delete_sql, [file_date_full.strip()] + list(unique_macns))
            
    else:
        # Chiến lược GLD (Mặc định): Xóa diện rộng theo Ngày phát sinh + Cây đơn vị liên đới
        format_strings = ','.join(['%s'] * len(unique_macns))
        delete_sql = f"""
            DELETE FROM `{table_name}` 
            WHERE LEFT(REPLACE(REPLACE(TRIM(`thoidiem`), '\\r', ''), '\\n', ''), 8) = %s 
              AND REPLACE(REPLACE(TRIM(`macn`), '\\r', ''), '\\n', '') IN ({format_strings})
        """
        cursor.execute(delete_sql, [file_date_full.strip()] + list(unique_macns))

    # --- 4. THỰC THI BATCH INSERT DUNG LƯỢNG LỚN ---
    if actual_rows > 0:
        column_names_sql = ", ".join([f"`{c[0]}`" for c in current_structure])
        full_columns_sql = f"{column_names_sql}, `ten_file_goc`, `hinh_thuc_file`"
        value_placeholders = ", ".join(["%s"] * (len(current_structure) + 2))
        insert_query = f"INSERT INTO `{table_name}` ({full_columns_sql}) VALUES ({value_placeholders})"
        
        batch_size = 3000
        for i in range(0, actual_rows, batch_size):
            batch_lines = cleaned_data_lines[i:i + batch_size]
            batch_params = []
            
            for fields in batch_lines:
                row_data = []
                for j in range(len(current_structure)):
                    if j < len(fields):
                        row_data.append(str(fields[j]).replace('\r', '').replace('\n', '').strip())
                    else:
                        row_data.append("")
                row_data.append(filename)
                row_data.append(file_hinh_thuc)
                batch_params.append(tuple(row_data))
            
            cursor.executemany(insert_query, batch_params)

    # --- 5. KIỂM TRA KIỂM TOÁN NGHIỆP VỤ HẬU KỲ ---
    try:
        target_month = file_date_full[:6]
        run_post_import_logic_check(cursor, table_name, target_month, filename)
    except Exception as e:
        print(f"[!] Cảnh báo lỗi đối chiếu logic nghiệp vụ: {e}")

    # --- 6. GHI VẾT THÔNG TIN LỊCH SỬ NẠP TỆP TIN ---
    try:
        sql_log = """
            INSERT INTO `log_file_imports` 
            (file_name, header_raw, macn, thoidiem, loai_bc, hinh_thuc, stt, so_luong, user_import) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        cursor.execute(sql_log, (
            filename,          # file_name
            header_line,       # header_raw
            header_macn,       # macn
            file_date_full,    # thoidiem
            header_loai_bc,    # loai_bc
            header_hinh_thuc,  # hinh_thuc
            header_stt,        # stt
            actual_rows,       # so_luong
            current_user_name  # user_import
        ))
    except Exception as e:
        print(f"[!] Bỏ qua ghi log_file_imports nếu chưa tạo bảng: {e}")

    sql_latest = """
        INSERT INTO `log_file_imports_latest` (file_name, last_import_date, last_user, status) 
        VALUES (%s, NOW(), %s, %s)
        ON DUPLICATE KEY UPDATE last_import_date = NOW(), last_user = VALUES(last_user), status = VALUES(status)
    """
    cursor.execute(sql_latest, (filename, current_user_name, final_import_status))
    
    print(f"[+] Thành công: {filename} -> `{table_name}` (+{actual_rows} dòng)")

# ==============================================================================
# 4. HÀM ĐIỀU PHỐI QUÉT THƯ MỤC CHÍNH
# ==============================================================================
def process_directory_import_standard(folder_path, current_user_name="system"):
    if not os.path.isdir(folder_path):
        print(f"[x] Đường dẫn mục tiêu không hợp lệ: {folder_path}")
        return

    txt_files = [f for f in os.listdir(folder_path) if f.upper().endswith('.TXT')]
    total_files = len(txt_files)
    if not txt_files:
        print(f"[*] Thư mục trống rỗng, không tìm thấy file đuôi .TXT cần nạp.")
        return

    print(f"\n" + "="*70)
    print(f"[+] HỆ THỐNG AML PHÁT HIỆN: {total_files} TỆP TIN DỮ LIỆU.")
    print(f"[+] KHỞI TẠO ĐƯỜNG TRUYỀN DATABASE DUY NHẤT ĐỂ XỬ LÝ HÀNG LOẠT...")
    print("="*70)

    db_config = parse_mysql_uri_standard(db_uri)
    connection = pymysql.connect(**db_config)

    try:
        with connection.cursor() as cursor:
            for idx, filename in enumerate(txt_files, 1):
                print(f"[{idx}/{total_files}] Đang xử lý: {filename}")
                try:
                    process_single_file(connection, cursor, folder_path, filename, current_user_name)
                    connection.commit()
                except Exception as fe:
                    connection.rollback()
                    print(f"[x] Lỗi nghiêm trọng tại file {filename}: {str(fe)} -> Đã Rollback dữ liệu file này.")
    finally:
        connection.close()
        print(f"\n" + "="*70)
        print("[+] HOÀN THÀNH TIẾN TRÌNH NẠP. KẾT NỐI HỆ THỐNG ĐÃ ĐÓNG AN TOÀN.")
        print("="*70)

if __name__ == "__main__":
    if os.name == 'posix':  
        PATH_THU_MUC = os.path.expanduser("/Users/lequocvan/Downloads/Cham_cong_Ngoai_gio/uploads_manual")
    else:  
        PATH_THU_MUC = r"C:\18_TXT_GLA\CTR_GLA 2025"

    process_directory_import_standard(PATH_THU_MUC, current_user_name="system")
