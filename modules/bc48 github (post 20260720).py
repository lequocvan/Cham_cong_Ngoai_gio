from flask import current_app, Response, send_file, flash, redirect, url_for, request, make_response, stream_with_context
from flask_login import current_user
from pathlib import Path # Thư viện xử lý đường dẫn đa nền tảng
from datetime import datetime
from sqlalchemy import text
from dateutil.relativedelta import relativedelta
from textwrap import dedent
import os, io, csv, time, shutil, traceback
import re
import pandas as pd
import numpy as np
import tempfile
import json
import pymysql
import uuid
import gc # Thêm thư viện giải phóng bộ nhớ
import zipfile
import string
import secrets 
import pyzipper
import math
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
from email.message import EmailMessage
from decimal import Decimal
from io import BytesIO

# --- CẤU TRÚC CỘT file TXT ---
FILE_STRUCTURES = {
    ...
}

# Đến tháng 05/2026 file csv có 17 cột;
CSV_ERROR_COLUMNS = [
...]

# hàm get_bc48_engine(db) trả về một đối tượng SQLAlchemy Engine được cấu hình đa CSDL (Multiple Binds) trong Flask-SQLAlchemy
def get_bc48_engine(db):
    return db.get_engine(bind='db_bc48')


################################################################################################################################################
# Thực hiện nạp TXT do Core trả ra hàng ngày
################################################################################################################################################
def validate_filename(filename):
    """Trả về (True, None) nếu hợp lệ, (False, error_detail) nếu lỗi."""
    fn_upper = filename.upper()
    if not fn_upper.endswith('.TXT'):
        return False, "RE2_loi_chi_tiet_6: File phải có đuôi .TXT"

    # Tách tên file (loại bỏ đuôi .TXT trước khi split)
    parts = fn_upper.replace('.TXT', '').split('_')
    if len(parts) != 5:
        return False, f"RE2_loi_chi_tiet_0: Cấu trúc tên file sai (yêu cầu 5 phần (ko tính TXT), nhận được {len(parts)})."
    
    ma_nh, ngay_bc, loai_bc, hinh_thuc, stt = parts
    
    if ma_nh != "01204001":
        return False, f"RE2_loi_chi_tiet_1: Mã ngân hàng {ma_nh} không hợp lệ."
    if not re.match(r'^\d{8}$', ngay_bc):
        return False, "RE2_loi_chi_tiet_2: Ngày báo cáo sai định dạng yyyymmdd."
    try:
        # datetime.strptime sẽ ném ra ValueError nếu ngày không tồn tại (ví dụ: 20260230)
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

def run_post_import_logic_check(engine, table_name, target_month, filename):
    """
    """
    # Trích xuất ngày từ cột ten_file_goc để đối chiếu (ví dụ: 20260520)
    file_date = filename.split('_')[1] 
    
    query = text(f"""
        INSERT IGNORE INTO log_loi_logic_ngay_thang 
        (table_name, thoidiem_loi, ten_file_loi, error_message)
        SELECT :tbl, thoidiem, ten_file_goc, 
               CASE 
                   WHEN LEFT(thoidiem, 6) <> :m THEN 'thoidiem khong thuoc thang bao cao (TXT dung ky bao cao)'
                   WHEN LEFT(thoidiem, 8) <> :f_date THEN 'thoidiem khong khop voi ngay trong ten file (TXT khop ngay phat sinh)'
                   WHEN LEFT(ten_file_goc, 15) NOT LIKE CONCAT('%', :m, '%') THEN 'File khong thuoc thang bao cao'
                   ELSE 'Loi logic ngay thang khong xac dinh'
               END
        FROM `{table_name}`
        WHERE LEFT(thoidiem, 6) <> :m 
           OR LEFT(thoidiem, 8) <> :f_date
           OR ten_file_goc NOT LIKE CONCAT('%', :m, '%')
    """)
    
    with engine.begin() as conn:
        conn.execute(query, {
            "tbl": table_name, 
            "m": target_month,
            "f_date": file_date
        })

def process_import_files(db, files_data, current_user_name="Hệ thống"):
    print(f"\n>>> [DEBUG] BẮT ĐẦU TIẾN TRÌNH: {len(files_data)} file", flush=True)
    engine = get_bc48_engine(db)
    
    for filename, raw_content in files_data:
        # --- BƯỚC 1: Kiểm tra tên file (Lỗi RE2) ---
        is_valid, error_detail = validate_filename(filename)
        if not is_valid:
            # Ghi lỗi RE2 vào log_import_errors
            try:
                with engine.begin() as log_conn:
                    log_conn.execute(text("""
                        INSERT INTO `log_import_errors` 
                        (file_name, ma_loi_bc48, header_content, user_import, status) 
                        VALUES (:fn, 'RE2', :err, :user, 'Chờ xử lý')
                    """), {"fn": filename, "err": error_detail, "user": current_user_name})
            except Exception as le:
                print(f"Lỗi ghi log RE2: {le}")
            
            yield f"data: {json.dumps({'status': 'error', 'msg': error_detail})}\n\n"
            continue

        # --- NẾU TÊN FILE TXT HỢP LỆ, BÓC TÁCH THÔNG TIN TỪ TÊN FILE ---
        parts = filename.upper().replace('.TXT', '').split('_')
        file_ma_nh = parts[0]     # Mã ngân hàng từ tên file
        file_date_full = parts[1] # yyyymmdd từ tên file
        prefix = parts[2]         # CTR/DWT/EFT/PTR
        file_hinh_thuc = parts[3] # GLD/GLA/GBS từ tên file
        file_stt = parts[4]       # STT từ tên file
        
        table_name = f"{prefix.lower()}_{file_date_full[:6]}"

        current_structure = FILE_STRUCTURES.get(prefix)
        if not current_structure: 
            yield f"data: {json.dumps({'status': 'error', 'msg': f'Không tìm thấy cấu trúc cho tiền tố {prefix}'})}\n\n"
            continue

        # Xác định vị trí cột 'macn' trong body để phục vụ dọn dẹp dữ liệu trùng
        try:
            macn_col_idx = [idx for idx, col in enumerate(current_structure) if col[0].lower() == 'macn'][0]
        except IndexError:
            yield f"data: {json.dumps({'status': 'error', 'msg': f'Cấu trúc file {prefix} thiếu cột macn'})}\n\n"
            continue
        
        try:
            # 2. Giải mã nội dung file bằng cơ chế thử sai encoding
            content = None
            for encoding in ['utf-8-sig', 'utf-16', 'utf-8', 'latin-1']:
                try:
                    content = raw_content.decode(encoding)
                    break
                except: 
                    continue
            
            if content is None:
                yield f"data: {json.dumps({'status': 'error', 'msg': f'Không thể đọc file {filename}'})}\n\n"
                continue

            # Thu thập các dòng thực tế
            all_lines_raw = [l for l in content.splitlines() if l.strip()]
            if not all_lines_raw: 
                yield f"data: {json.dumps({'status': 'error', 'msg': f'File {filename} trống rỗng không có dữ liệu'})}\n\n"
                continue

            # --- BƯỚC 2: KIỂM TRA DÒNG TIÊU ĐỀ (Lỗi RE3) ---
            header_line_raw = all_lines_raw[0]
            re3_error = None

            # Kiểm tra khoảng trắng ở rìa dòng tiêu đề
            if header_line_raw.startswith(' ') or header_line_raw.endswith(' '):
                re3_error = "Dòng tiêu đề không được bắt đầu hoặc kết thúc bằng khoảng trắng."
            
            # Kiểm tra số lượng phần tử phân tách bởi dấu #
            header_parts = header_line_raw.split('#')
            if re3_error is None and len(header_parts) != 6:
                re3_error = f"Dòng tiêu đề phải có đúng 6 phần phân tách bằng dấu # (nhận được {len(header_parts)})."
            
            if re3_error:
                try:
                    with engine.begin() as log_conn:
                        log_conn.execute(text("""
                            INSERT INTO `log_import_errors` 
                            (file_name, ma_loi_bc48, header_content, user_import, status) 
                            VALUES (:fn, 'RE3', :head, :user, 'Chờ xử lý')
                        """), {"fn": filename, "head": header_line_raw, "user": current_user_name})
                except Exception as le: 
                    print(f"Lỗi ghi log RE3: {le}")
                yield f"data: {json.dumps({'status': 'error', 'msg': f'Lỗi RE3 tại file {filename}: {re3_error}'})}\n\n"
                continue

            # --- BƯỚC 2B: KIỂM TRA ĐỐI CHIẾU TIÊU ĐỀ VỚI TÊN FILE (Lỗi RE6) ---
            header_line = header_line_raw.strip()
            header_macn = header_parts[0].strip().upper() 
            header_date = header_parts[1].strip()         
            header_loai_bc = header_parts[2].strip().upper()
            header_hinh_thuc = header_parts[3].strip().upper()
            header_stt = header_parts[4].strip()

            re6_details = []
            if header_macn != file_ma_nh:
                re6_details.append(f"Mã NH lệch ({file_ma_nh} vs {header_macn})")
            if header_loai_bc != prefix:
                re6_details.append(f"Loại BC lệch ({prefix} vs {header_loai_bc})")
            if header_hinh_thuc != file_hinh_thuc:
                re6_details.append(f"Hình thức lệch ({file_hinh_thuc} vs {header_hinh_thuc})")
            if header_stt != file_stt:
                re6_details.append(f"STT lệch ({file_stt} vs {header_stt})")
                
            if re6_details:
                re6_msg = "Tên file và dòng tiêu đề không trùng khớp: " + ", ".join(re6_details)
                try:
                    with engine.begin() as log_conn:
                        log_conn.execute(text("""
                            INSERT INTO `log_import_errors` 
                            (file_name, ma_loi_bc48, header_content, user_import, status) 
                            VALUES (:fn, 'RE6', :head, :user, 'Chờ xử lý')
                        """), {"fn": filename, "head": f"LỖI RE6: {re6_msg} | Header gốc: {header_line}", "user": current_user_name})
                except Exception as le: 
                    print(f"Lỗi ghi log RE6: {le}")
                yield f"data: {json.dumps({'status': 'error', 'msg': f'Lỗi RE6 tại file {filename}: {re6_msg}'})}\n\n"
                continue 
            
            # Đọc số lượng dòng được khai báo từ cột thứ 6 của tiêu đề
            raw_qty = "".join(filter(str.isdigit, header_parts[5].strip()))
            declared_rows = int(raw_qty) if raw_qty else 0

            data_lines = all_lines_raw[1:]
            actual_rows = len(data_lines)

            # Kiểm tra lỗi RE4 (Sai lệch số dòng) và RF3 (Sai lệch ngày báo cáo)
            error_code = None
            if declared_rows != actual_rows:
                error_code = "RE4" 
            elif header_date != file_date_full:
                error_code = "RF3" 
            
            if error_code:
                try:
                    with engine.begin() as log_conn:
                        log_conn.execute(text("""
                            INSERT INTO `log_import_errors` 
                            (file_name, ma_loi_bc48, header_content, declared_rows, actual_rows, user_import)
                            VALUES (:fn, :code, :head, :decl, :act, :user)
                        """), {
                            "fn": filename, "code": error_code, "head": header_line, 
                            "decl": declared_rows, "act": actual_rows, "user": current_user_name
                        })
                except Exception as le: 
                    print(f"Lỗi ghi log RE4/RF3: {le}")

            # --- BƯỚC 3: XỬ LÝ CHỐNG LẶP THEO MÃ CHI NHÁNH (Diện rộng) ---
            unique_macns = set()
            cleaned_data_lines = []

            # Đọc, phân tách nội dung chi tiết bằng dấu # và làm sạch khoảng trắng
            for line in data_lines:
                f = line.strip().split('#')
                f_cleaned = [item.strip() for item in f] 
                cleaned_data_lines.append(f_cleaned)

                if len(f_cleaned) > macn_col_idx:
                    val = f_cleaned[macn_col_idx]
                    if val: 
                        unique_macns.add(val)

            # Truy vấn quét danh mục các chi nhánh liên đới từ bảng don_vi
            if header_macn:
                unique_macns.add(header_macn)
                try:
                    with engine.begin() as conn_dv:
                        query_dv = text("""
                            SELECT DISTINCT TRIM(`MaNH8so_moi`) as macn_con
                            FROM `don_vi`
                            WHERE `ma_hieu_1` = (
                                SELECT `ma_hieu_1` 
                                FROM `don_vi` 
                                WHERE TRIM(`MaNH8so_moi`) = :hm 
                                LIMIT 1
                            ) 
                            AND `MaNH8so_moi` IS NOT NULL 
                            AND `MaNH8so_moi` != '';
                        """)
                        res_dv = conn_dv.execute(query_dv, {"hm": header_macn})
                        for r in res_dv:
                            if r.macn_con:
                                unique_macns.add(r.macn_con)
                    print(f">>> [DEBUG DB] File {filename}: Quét được {len(unique_macns)} chi nhánh liên quan (Mẹ + Con).", flush=True)
                except Exception as edv:
                    print(f">>> [WARNING DB] Lỗi tra cứu liên đới chi nhánh: {edv}.", flush=True)

            if not unique_macns and header_macn:
                unique_macns.add(header_macn)

            # Phát tín hiệu kiểm tra ban đầu lên giao diện UI
            yield f"data: {json.dumps({'status': 'check', 'filename': filename, 'declared': declared_rows, 'actual': actual_rows, 'error_code': error_code})}\n\n"
            yield f"data: {json.dumps({'status': 'start', 'filename': filename, 'total': actual_rows})}\n\n"
            
            # --- BƯỚC 4: TIẾN TRÌNH THAO TÁC CƠ SỞ DỮ LIỆU ---
            
            # 4A: Đảm bảo bảng tháng tồn tại và tích hợp cả 2 cột ẩn hệ thống ở cuối bảng
            with engine.begin() as conn:
                cols_sql = ", ".join([f"`{c[0]}` TEXT" for c in current_structure])
                create_sql = f"""
                    CREATE TABLE IF NOT EXISTS `{table_name}` (
                        {cols_sql},
                        `ten_file_goc` VARCHAR(150) DEFAULT NULL,
                        `hinh_thuc_file` VARCHAR(10) DEFAULT NULL,  -- Cột hệ thống phân biệt hình thức dữ liệu
                        INDEX idx_macn (`macn`(20)),
                        INDEX idx_td (`thoidiem`(10)),
                        INDEX idx_file (`ten_file_goc`(50)),
                        INDEX idx_ht (`hinh_thuc_file`(10))  -- Index hỗ trợ tăng tốc độ lọc báo cáo
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci
                """
                conn.execute(text(create_sql))

                # Cơ chế Auto-Migration: Kiểm tra và tự động đồng bộ hóa các cột ẩn nếu chạy trên các bảng tháng cũ
                db_cols_res = conn.execute(text(f"SHOW COLUMNS FROM `{table_name}`"))
                existing_cols = {row[0].lower() for row in db_cols_res}
                
                # Kiểm tra & tự động thêm cột `ten_file_goc` nếu bảng cũ chưa có
                if "ten_file_goc" not in existing_cols:
                    conn.execute(text(f"ALTER TABLE `{table_name}` ADD COLUMN `ten_file_goc` VARCHAR(150) DEFAULT NULL"))
                    conn.execute(text(f"ALTER TABLE `{table_name}` ADD INDEX idx_file (`ten_file_goc`(50))"))
                    print(f">>> [Migration] Đã bổ sung cột hệ thống `ten_file_goc` và Index vào bảng {table_name}", flush=True)

                # Kiểm tra & tự động thêm cột `hinh_thuc_file` nếu bảng cũ chưa có
                if "hinh_thuc_file" not in existing_cols:
                    conn.execute(text(f"ALTER TABLE `{table_name}` ADD COLUMN `hinh_thuc_file` VARCHAR(10) DEFAULT NULL"))
                    conn.execute(text(f"ALTER TABLE `{table_name}` ADD INDEX idx_ht (`hinh_thuc_file`(10))"))
                    print(f">>> [Migration] Đã bổ sung cột hệ thống `hinh_thuc_file` và Index vào bảng {table_name}", flush=True)

                # Kiểm tra thêm các cột nghiệp vụ mới nếu cấu trúc FILE_STRUCTURES được thay đổi trong tương lai
                for col_name, col_type in current_structure:
                    if col_name.lower() not in existing_cols:
                        alter_sql = f"ALTER TABLE `{table_name}` ADD COLUMN `{col_name}` TEXT"
                        conn.execute(text(alter_sql))
                        print(f">>> [Migration] Đã bổ sung cột nghiệp vụ mới `{col_name}` vào bảng {table_name}", flush=True)

            # 🔥 4B: XỬ LÝ CHÍNH - RẼ NHÁNH XÓA TRÙNG CHO GLA THEO MAGD ĐỘNG
            # Chống trùng lặp dữ liệu khi người dùng nạp lại (re-import) file
            # GLD: gửi báo cáo lần đầu theo đúng quy định; GLA: gửi lại báo cáo đối vác giao dịch sai định dạng, lỗi được Cục PCRT thông báo và yêu cầu gửi lại; GBS: gửi bổ sung các giao dịch khi tự rà soát và phát hiện các giao dịch báo cáo còn thiếu hoặc báo cáo bổ sung theo đề nghị của CQNN.
            # Lệnh DELETE của GLA sử dụng đồng thời bộ 3 điều kiện chặn: Đúng ngày (thoidiem) + Đúng cụm chi nhánh (macn) + Đúng Mã giao dịch (magd)
            with engine.begin() as conn:
                # =========================================================================
                # 🚀 XÓA SẠCH VẾT LỊCH SỬ CŨ CỦA CHÍNH FILE NÀY TRONG ARCHIVE
                # =========================================================================
                clear_old_archive_sql = """
                    DELETE FROM `bc48_deleted_records_archive` 
                    WHERE `deleted_by_file` = :new_file
                """
                conn.execute(text(clear_old_archive_sql), {"new_file": filename})
                
                # --- [CHUẨN BỊ SQL ĐỘNG ĐỂ LƯU TRỮ DỮ LIỆU SẮP XÓA] ---
                # Tạo chuỗi JSON_OBJECT động dựa trên cấu trúc các cột của bảng hiện tại
                # Ví dụ: JSON_OBJECT('macn', `macn`, 'thoidiem', `thoidiem`, ...)
                json_fields = ", ".join([f"'{c[0].lower()}', `{c[0]}`" for c in current_structure])
                json_object_sql = f"JSON_OBJECT({json_fields})"
                
                # Định nghĩa các cột cố định trích xuất ra ngoài để tạo Index tìm kiếm cho nhanh
                # Kiểm tra xem cấu trúc hiện tại có cột magd không
                has_magd = any(c[0].lower() == 'magd' for c in current_structure)
                magd_select = "`magd`" if has_magd else "NULL"

                # Lệnh SQL tổng quát để backup dữ liệu
                archive_base_sql = f"""
                    INSERT INTO `bc48_deleted_records_archive` (
                        `loai_bc`, `ten_table_goc`, `ten_file_goc_cu`, `hinh_thuc_file_cu`,
                        `macn`, `thoidiem`, `magd`, `raw_record_json`,
                        `deleted_by_file`, `deleted_by_user`, `deleted_at`
                    )
                    SELECT 
                        :prefix, :table_name, `ten_file_goc`, `hinh_thuc_file`,
                        `macn`, `thoidiem`, {magd_select}, {json_object_sql},
                        :new_file, :user_import, NOW()
                    FROM `{table_name}`
                """

                # 🟢 LOGIC CHO GBS (Bổ sung): Chỉ dọn dẹp file trùng tên nếu nạp lại
                if file_hinh_thuc == "GBS":
                    # 1. Sao lưu dữ liệu của chính file này trước khi xóa
                    archive_query = text(archive_base_sql + " WHERE `ten_file_goc` = :fn")
                    conn.execute(archive_query, {
                        "prefix": prefix, "table_name": table_name, 
                        "new_file": filename, "user_import": current_user_name, "fn": filename
                    })
                    
                    # 2. Xóa dữ liệu cũ
                    # 🟢 LOGIC CHO GBS (Bổ sung): Giữ nguyên dữ liệu cũ, chỉ dọn dẹp file trùng tên nếu nạp lại
                    delete_sql = f"DELETE FROM `{table_name}` WHERE `ten_file_goc` = :fn"
                    result = conn.execute(text(delete_sql), {"fn": filename})
                    print(f">>> [DEBUG GBS] File {filename}: Đã dọn sạch {result.rowcount} dòng cũ của chính tệp này.", flush=True)

                elif file_hinh_thuc == "GLA":
                    # 🟡 LOGIC CHO GLA (Gửi lại giao dịch sửa lỗi): Bóc tách chính xác mảng các chuỗi magd từ cleaned_data_lines
                    
                    # 1. Tự động xác định vị trí cột 'magd' dựa theo cấu trúc file đang nạp
                    try:
                        magd_col_idx = [idx for idx, col in enumerate(current_structure) if col[0].lower() == 'magd'][0]
                        # Trích xuất danh sách magd thô từ các dòng dữ liệu, loại bỏ khoảng trắng và dòng trống
                        list_ma_gd = [
                            fields[magd_col_idx].strip()
                            for fields in cleaned_data_lines 
                            if len(fields) > magd_col_idx and fields[magd_col_idx] and fields[magd_col_idx].strip()
                        ]
                    except IndexError:
                        list_ma_gd = []
                        print(f">>> [WARNING] Cấu trúc {prefix} không tìm thấy cấu trúc hoặc vị trí cột 'magd' để đối chiếu.", flush=True)

                    if list_ma_gd:
                        # Sử dụng cơ chế chia nhỏ magd_list nếu số lượng giao dịch sửa đổi quá lớn (tránh nghẽn lock)
                        list_ma_gd_unique = list(set(list_ma_gd))
                        t_clean = file_date_full.strip()
                        # SQL Lưu trữ cho GLA
                        archive_sql = archive_base_sql + """
                            WHERE `thoidiem` LIKE :t_like
                              AND `macn` IN :m_list
                              AND `magd` IN :magd_list
                        """
                        
                        # Thực hiện xóa khu trú: Đúng ngày phát sinh + Thuộc cụm chi nhánh + Đúng Mã giao dịch lỗi
                        # 🌟 TỐI ƯU 1: Sử dụng Index trực tiếp, loại bỏ toàn bộ hàm bọc cột của MySQL
                        # Vì khi insert dữ liệu đã sạch, ta chỉ cần so sánh trực tiếp chuỗi bằng "=" và "IN"
                        delete_sql = f"""
                            DELETE FROM `{table_name}` 
                            WHERE `thoidiem` LIKE :t_like
                              AND `macn` IN :m_list
                              AND `magd` IN :magd_list
                        """
                        t_clean = file_date_full.strip()

                        # Xử lý theo từng Chunk để giảm tải cho bộ nhớ và tránh Lock bảng dài hạn
                        chunk_size = 1000
                        total_archived = 0
                        total_deleted = 0
                        
                        for chunk_idx in range(0, len(list_ma_gd_unique), chunk_size):
                            sub_magd_list = list_ma_gd_unique[chunk_idx:chunk_idx + chunk_size]
                            params = {
                                "prefix": prefix, "table_name": table_name,
                                "new_file": filename, "user_import": current_user_name,
                                "t_like": f"{t_clean}%", "m_list": list(unique_macns),
                                "magd_list": sub_magd_list
                            }
                            # Chạy lệnh Sao lưu cho chunk hiện tại
                            arc_res = conn.execute(text(archive_sql), params)
                            total_archived += arc_res.rowcount
                            
                            # Chạy lệnh Xóa cho chunk hiện tại
                            del_res = conn.execute(text(delete_sql), params)
                            total_deleted += del_res.rowcount
                        print(f">>> [DEBUG GLA] File {filename}: Đã sao lưu {total_archived} dòng và dọn sạch {total_deleted} dòng lỗi cũ.", flush=True)
                    else:
                        # Fallback nếu file GLA trống mã giao dịch: Xóa theo tên file
                        archive_query = text(archive_base_sql + " WHERE `ten_file_goc` = :fn")
                        conn.execute(archive_query, {
                            "prefix": prefix, "table_name": table_name, 
                            "new_file": filename, "user_import": current_user_name, "fn": filename
                        })
                        conn.execute(text(f"DELETE FROM `{table_name}` WHERE `ten_file_goc` = :fn"), {"fn": filename})

                elif file_hinh_thuc == "GLD":
                    # 🔴 LOGIC CHO GLD (Lần đầu) - TỐI ƯU ĐỒNG BỘ
                    t_clean = file_date_full.strip()
                    
                    # 1. Sao lưu diện rộng trước khi xóa
                    archive_sql = archive_base_sql + """
                        WHERE `thoidiem` LIKE :t_like
                          AND `macn` IN :m_list
                    """
                    arc_res = conn.execute(text(archive_sql), {
                        "prefix": prefix, "table_name": table_name,
                        "new_file": filename, "user_import": current_user_name,
                        "t_like": f"{t_clean}%", "m_list": list(unique_macns)
                    })
                    
                    # 2. Thực hiện xóa diện rộng
                    delete_sql = f"""
                        DELETE FROM `{table_name}` 
                        WHERE `thoidiem` LIKE :t_like
                          AND `macn` IN :m_list
                    """
                    del_res = conn.execute(text(delete_sql), {
                        "t_like": f"{t_clean}%", "m_list": list(unique_macns)
                    })
                    print(f">>> [DEBUG GLD] File {filename}: Đã sao lưu {arc_res.rowcount} dòng và dọn sạch diện rộng {del_res.rowcount} dòng cũ ngày {t_clean}.", flush=True)

                # Thực hiện thiết kế Giải pháp 1: Lệnh INSERT gọi tên cột tường minh để cô lập cột ẩn hệ thống
                # --- TIẾN HÀNH BATCH INSERT DỮ LIỆU MỚI TƯỜNG MINH ---
                if actual_rows > 0:
                    # 1. Khởi tạo chuỗi danh sách các cột nghiệp vụ động: `macn`, `thoidiem`, ...
                    column_names_sql = ", ".join([f"`{c[0]}`" for c in current_structure])
                    # 2. Đính kèm đích danh 2 cột ẩn hệ thống vào cuối danh sách cột SQL
                    full_columns_sql = f"{column_names_sql}, `ten_file_goc`, `hinh_thuc_file`"
                    
                    # 3. Tạo chuỗi placeholders giá trị động: :v0, :v1, ...
                    value_placeholders = ", ".join([f":v{i}" for i in range(len(current_structure))])
                    # 4. Đính kèm 2 placeholders hệ thống `:v_file` và `:v_ht` vào cuối chuỗi giá trị
                    full_values_sql = f"{value_placeholders}, :v_file, :v_ht"
                    
                    # Khởi tạo câu lệnh SQL INSERT hoàn chỉnh
                    insert_query = text(f"INSERT INTO `{table_name}` ({full_columns_sql}) VALUES ({full_values_sql})")
                
                    # Chia nhỏ dữ liệu thành từng Batch để tối ưu hóa bộ nhớ và tốc độ ghi của InnoDB
                    batch_size = 3000
                    for i in range(0, actual_rows, batch_size):
                        batch_lines = cleaned_data_lines[i:i + batch_size]
                        batch_params = []
                        for fields in batch_lines:
                            row_dict = {}
                            # Đọc dữ liệu chi tiết theo đúng độ dài cấu trúc của FILE_STRUCTURES hiện hành
                            for j in range(len(current_structure)):
                                if j < len(fields):
                                    val_clean = str(fields[j]).replace('\r', '').replace('\n', '').strip()
                                    row_dict[f"v{j}"] = val_clean
                                else:
                                    row_dict[f"v{j}"] = ""
                                    
                            # Gán chính xác giá trị cho hai tham số hệ thống phục vụ truy vết tương lai
                            row_dict["v_file"] = filename
                            row_dict["v_ht"] = file_hinh_thuc # Nhận giá trị "GLD", "GLA" hoặc "GBS"
                            batch_params.append(row_dict)
                    
                        conn.execute(insert_query, batch_params)
                        yield f"data: {json.dumps({'status': 'progress', 'current': min(i + batch_size, actual_rows)})}\n\n"
                        
            # Kiểm tra logic giữa hậu tố tên bảng (yyyymm) và thoidiem, ten_file_goc
            try:
                target_month = file_date_full[:6]
                run_post_import_logic_check(engine, table_name, target_month, filename)
                print(f">>> [Check] Đã kiểm tra logic ngày tháng cho file {filename}", flush=True)
            except Exception as e:
                print(f">>> [Error] Lỗi khi kiểm tra logic sau import: {e}", flush=True)
            
            # --- BƯỚC 5: GHI LOG NHẬT KÝ VÀ TRẠNG THÁI SAU KHI NẠP THÀNH CÔNG ---
            try:
                header_data = {
                    "fn": filename,
                    "raw": header_line,
                    "macn": header_parts[0].strip() if len(header_parts) > 0 else None,
                    "thoidiem": file_date_full, 
                    "loai_bc": header_parts[2].strip() if len(header_parts) > 2 else None,
                    "hinh_thuc": header_parts[3].strip() if len(header_parts) > 3 else None,
                    "stt": header_parts[4].strip() if len(header_parts) > 4 else None,
                    "sl": actual_rows,
                    "user": current_user_name
                }
                    
                with engine.begin() as log_conn:
                    # Lớp log lịch sử: Lưu vết toàn bộ các lần import (Insert mới)
                    log_conn.execute(text("""
                        INSERT INTO `log_file_imports` 
                        (file_name, header_raw, macn, thoidiem, loai_bc, hinh_thuc, stt, so_luong, user_import, import_date) 
                        VALUES (:fn, :raw, :macn, :thoidiem, :loai_bc, :hinh_thuc, :stt, :sl, :user, NOW())
                    """), header_data)

                    # Lớp log dashboard: Chỉ lưu trạng thái mới nhất của từng tệp tin (Upsert)
                    log_conn.execute(text("""
                        INSERT INTO `log_file_imports_latest` 
                        (file_name, last_import_date, last_user, status) 
                        VALUES (:fn, NOW(), :user, 'SUCCESS')
                        ON DUPLICATE KEY UPDATE 
                            last_import_date = NOW(),
                            last_user = VALUES(last_user),
                            status = 'SUCCESS'
                    """), {"fn": filename, "user": current_user_name})
            except Exception as le:
                print(f"Lỗi ghi log lịch sử thành công: {le}")

            yield f"data: {json.dumps({'status': 'complete', 'filename': filename})}\n\n"

        except Exception as e:
            import traceback
            traceback.print_exc()
            yield f"data: {json.dumps({'status': 'error', 'msg': f'Lỗi hệ thống tại file {filename}: {str(e)}'})}\n\n"



#Gửi mail báo lỗi: cấu trúc file csv do Cục PCRT trả ra có sai lệch so với cấu trúc đã khai báo
def send_alert_email(smtp_config, filename, details, user_name):
    """
    Gửi cảnh báo cấu trúc file cho Admin thông qua SMTP sử dụng cấu hình từ .env.
    """
    try:
        msg = EmailMessage()
        msg['Subject'] = f"[CẢNH BÁO] Hệ thống phát hiện file CSV không chuẩn: {filename}"
        msg['From'] = smtp_config['email']
        msg['To'] = "pcrtagribank@gmail.com"  # Thay thế bằng email quản trị viên của bạn
        
        content = f"""
        Chào Admin,
        
        Hệ thống vừa thực hiện nạp dữ liệu từ file CSV và phát hiện cấu trúc không khớp hoàn toàn với chuẩn 17 cột quy định.
        
        - Tên file: {filename}
        - Chi tiết thay đổi: {details}
        - Người thực hiện: {user_name}
        - Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        
        Hệ thống đã tự động bù cột thiếu hoặc loại bỏ cột thừa để hoàn tất quá trình nạp. 
        Vui lòng kiểm tra lại file gốc nếu dữ liệu có sai khác.
        """
        msg.set_content(content)

        # Kết nối và gửi email
        with smtplib.SMTP(smtp_config['server'], int(smtp_config['port']), timeout=60) as s:
            if smtp_config.get('use_tls'):
                s.starttls()
            s.login(smtp_config['email'], smtp_config['password'])
            s.send_message(msg)
            
    except Exception as e:
        # Ghi lỗi vào log của Flask để Admin biết việc gửi cảnh báo bị lỗi
        current_app.logger.error(f"Gửi email cảnh báo thất bại cho file {filename}: {str(e)}")

    
################################################################################################################################################
# Sao kê theo ten_file_goc từ các bảng ctr_yyyymm; dwt_yyyymm; eft_yyyymm; ptr_yyyymm
################################################################################################################################################
def lay_danh_sach_bang_theo_khoang_thoi_gian(db, loai_bc, tu_ngay, den_ngay):
    """
    Truy vấn thẳng vào information_schema để lấy chính xác các bảng tồn tại
    nằm trong khoảng thời gian cần tra cứu. Đã sửa lỗi tương thích định dạng chuỗi ngày.
    """
    # CHUẨN HÓA: Loại bỏ dấu gạch ngang nếu có để đưa về một chuẩn duy nhất yyyymmdd trước khi parse
    tu_ngay_clean = tu_ngay.replace('-', '')
    den_ngay_clean = den_ngay.replace('-', '')
    
    try:
        start_dt = datetime.strptime(tu_ngay_clean, "%Y%m%d")
        end_dt = datetime.strptime(den_ngay_clean, "%Y%m%d")
    except Exception:
        # Dự phòng trường hợp định dạng truyền vào vẫn giữ nguyên yyyy-mm-dd
        try:
            start_dt = datetime.strptime(tu_ngay, "%Y-%m-%d")
            end_dt = datetime.strptime(den_ngay, "%Y-%m-%d")
        except Exception:
            return [] # Trả về rỗng nếu định dạng ngày hoàn toàn không hợp lệ
            
    prefix = loai_bc.lower()
    
    thang_nam_list = []
    current_year = start_dt.year
    current_month = start_dt.month
    
    while (current_year < end_dt.year) or (current_year == end_dt.year and current_month <= end_dt.month):
        thang_nam_list.append(f"{prefix}_{current_year}{current_month:02d}")
        if current_month == 12:
            current_month = 1
            current_year += 1
        else:
            current_month += 1

    if not thang_nam_list:
        return []

    engine = get_bc48_engine(db)
    with engine.connect() as conn:
        db_name = conn.execute(text("SELECT DATABASE()")).scalar()
        
        sql = text("""
            SELECT LOWER(table_name) 
            FROM information_schema.tables 
            WHERE table_schema = :db_name 
              AND table_name IN :tables
        """)
        
        res = conn.execute(sql, {"db_name": db_name, "tables": tuple(thang_nam_list)}).fetchall()
        tables_in_db = [r[0] for r in res]
        
    tables_in_db.sort(reverse=True)
    return tables_in_db


def count_saoke_txt(db, loai_bc, tu_ngay, den_ngay, ten_file_goc=''):
    """ Đếm tổng số dòng phục vụ phân trang """
    list_tables = lay_danh_sach_bang_theo_khoang_thoi_gian(db, loai_bc, tu_ngay, den_ngay)
    if not list_tables:
        return 0
        
    engine = get_bc48_engine(db)
    queries = []
    
    # Loại bỏ dấu gạch ngang của ngày để khớp với định dạng chuỗi TEXT (yyyymmdd hh:mm:ss) trong DB
    tu_ngay_clean = tu_ngay.replace('-', '')
    den_ngay_clean = den_ngay.replace('-', '')
    
    params = {
        "tu_ngay": f"{tu_ngay_clean} 00:00:00", 
        "den_ngay": f"{den_ngay_clean} 23:59:59"
    }
    
    where_file_clause = ""
    if ten_file_goc:
        where_file_clause = "AND ten_file_goc LIKE :ten_file"
        params["ten_file"] = f"%{ten_file_goc}%"
        
    for table in list_tables:
        queries.append(f"SELECT COUNT(*) as total FROM {table} WHERE thoidiem BETWEEN :tu_ngay AND :den_ngay {where_file_clause}")
        
    final_sql = f"SELECT SUM(total) FROM ( {' UNION ALL '.join(queries)} ) as k"
    
    with engine.connect() as conn:
        result = conn.execute(text(final_sql), params).scalar()
        return result or 0


def van_tin_saoke_txt(db, loai_bc, tu_ngay, den_ngay, page=1, per_page=50, ten_file_goc=''):
    """ Truy vấn dữ liệu sao kê chi tiết có phân trang cuốn chiếu từng bảng """
    list_tables = lay_danh_sach_bang_theo_khoang_thoi_gian(db, loai_bc, tu_ngay, den_ngay)
    if not list_tables:
        return []
        
    engine = get_bc48_engine(db)
    
    target_offset = (page - 1) * per_page  
    records_to_collect = per_page          
    final_results = []
    
    # Đồng bộ định dạng chuỗi TEXT loại bỏ gạch ngang cho hàm vấn tin dữ liệu
    tu_ngay_clean = tu_ngay.replace('-', '')
    den_ngay_clean = den_ngay.replace('-', '')
    
    params = {
        "tu_ngay": f"{tu_ngay_clean} 00:00:00", 
        "den_ngay": f"{den_ngay_clean} 23:59:59"
    }
    
    where_file_clause = ""
    if ten_file_goc:
        where_file_clause = "AND ten_file_goc LIKE :ten_file"
        params["ten_file"] = f"%{ten_file_goc}%"

    with engine.connect() as conn:
        for table in list_tables:
            if records_to_collect <= 0:
                break  
                
            # 1. Đếm số dòng thỏa mãn điều kiện tại bảng hiện tại
            count_sql = f"SELECT COUNT(*) FROM {table} WHERE thoidiem BETWEEN :tu_ngay AND :den_ngay {where_file_clause}"
            table_rows = conn.execute(text(count_sql), params).scalar() or 0
            
            if table_rows == 0:
                continue
                
            # 2. Kiểm tra vị trí offset phân trang
            if target_offset >= table_rows:
                target_offset -= table_rows
                continue
            
            # 3. SELECT ĐẦY ĐỦ CÁC TRƯỜNG THEO YÊU CẦU
            query_sql = f"""
                SELECT 
                    magd, macn, ten_file_goc, hinh_thuc_file, thoidiem, 
                    kieukh, sogt, tenkh, ngaysinh, quoctich, diachitt, 
                    loaitien, sotien, quydoi, lydomucdich
                FROM {table} 
                WHERE thoidiem BETWEEN :tu_ngay AND :den_ngay {where_file_clause}
                ORDER BY thoidiem DESC
                LIMIT :limit OFFSET :offset
            """
            
            local_params = params.copy()
            local_params["limit"] = records_to_collect
            local_params["offset"] = target_offset
            
            rows = conn.execute(text(query_sql), local_params).fetchall()
            
            for r in rows:
                row_dict = dict(r._mapping) if hasattr(r, '_mapping') else dict(r)
                final_results.append(row_dict)
            
            records_to_collect -= len(rows)
            target_offset = 0

    return final_results

################################################################################################################
################################################################################################################
def sanitize_and_save(df):
    """Làm sạch và ép kiểu dữ liệu trước khi nạp vào DB"""
    # Loại bỏ các dòng mà tất cả giá trị đều là NaN/trống
    df = df.dropna(how='all')
    if df.empty:
        return df
    
    # 1. Ép định dạng NGAY_BAOCAO về YYYYMMDD
    if 'NGAY_BAOCAO' in df.columns:
        df['NGAY_BAOCAO'] = pd.to_datetime(df['NGAY_BAOCAO'], errors='coerce').dt.strftime('%Y%m%d').fillna('00000000')
    
    # 2. Xử lý khoảng trắng thừa ở tất cả các cột string
    df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    # --- ĐỘT PHÁ TỐI ƯU CẤU TRÚC TẠI ĐÂY ---
    # reindex tự động: Thêm cột thiếu gán NaN, giữ lại cột đúng cấu trúc, loại bỏ cột thừa lạ, xếp chuẩn thứ tự 17 cột.
    df = df.reindex(columns=CSV_ERROR_COLUMNS)
    
    # 3. Lấp đầy tất cả ô trống hoặc cột mới bổ sung bằng "N/A"
    df = df.fillna("N/A")

    return df


def log_to_db(conn, *, file_name, loai, trang_thai, ma_nv_import, ngay_bc=None, ghi_chu=None, so_dong=0):
    """
    Hàm ghi log vào bảng logs_nap_csv.
    Bắt buộc truyền tham số dạng keyword (tên_tham_số=giá_trị) để chống lệch cột trong DB.
    """
    sql = text("""
        INSERT INTO logs_nap_csv (file_name, loai_bc, ngay_baocao, trang_thai, user_import, ghi_chu, so_dong_du_lieu_csv) 
        VALUES (:fname, :loai, :ngay, :status, :ma_nv, :note, :so_dong)
    """)
    conn.execute(sql, {
        "fname": file_name, 
        "loai": loai,
        "ngay": ngay_bc, # Định dạng YYYYMMDD hoặc YYYY-MM-DD
        "status": trang_thai, 
        "ma_nv": ma_nv_import, 
        "note": ghi_chu,
        "so_dong": so_dong
    })
    
def load_to_staging(conn, df, target_table, file_name):
    """
    Nạp dữ liệu vào bảng tạm thông qua LOAD DATA LOCAL INFILE và chuyển sang bảng chính.
    Đảm bảo sửa triệt để lỗi hụt dòng do ký tự xuống dòng nội bộ (Embedded Newlines).
    """
    if df.empty:
        return

    # Khởi tạo định danh độc nhất cho bảng tạm vật lý
    unique_id = uuid.uuid4().hex[:8]
    staging_table = f"temp_staging_{target_table}_{unique_id}"
    
    # 1. Tạo bảng tạm dựa trên cấu trúc bảng chính
    conn.execute(text(f"CREATE TABLE `{staging_table}` LIKE `{target_table}`"))

    tf_path = None
    try:
        # 2. Lưu df vào file tạm để chuẩn bị LOAD DATA
        with tempfile.NamedTemporaryFile(mode='w+', delete=False, suffix='.csv', encoding='utf-8', newline='') as tf:
            df.to_csv(
                tf, 
                index=False, 
                header=False, 
                encoding='utf-8', 
                quoting=1,               # QUOTE_ALL: Bọc tất cả các trường dữ liệu trong dấu ""
                lineterminator='\r\n'    # Ký tự ngắt dòng vật lý chuẩn
            )
            tf_path = tf.name.replace('\\', '/')
            
        # 3. Load cực nhanh dữ liệu file tạm vào bảng Staging tạm thời
        conn.execute(text(f"""
            LOAD DATA LOCAL INFILE '{tf_path}' 
            INTO TABLE `{staging_table}` 
            CHARACTER SET utf8mb4
            FIELDS TERMINATED BY ',' 
            ENCLOSED BY '"' 
            ESCAPED BY '"'
            LINES TERMINATED BY '\\r\\n'
        """))
        
        # 4. Chuyển toàn bộ dữ liệu từ staging sang bảng đích chính thức
        conn.execute(text(f"""
            INSERT INTO `{target_table}` 
            SELECT * FROM `{staging_table}`;
        """))
        
    finally:
        # --- KHỐI GIẢI PHÓNG TÀI NGUYÊN TUYỆT ĐỐI ---
        # Luôn giải phóng bảng Staging vật lý trong MySQL
        try:
            conn.execute(text(f"DROP TABLE IF EXISTS `{staging_table}`"))
        except Exception as drop_err:
            current_app.logger.error(f"Không thể xóa bảng tạm {staging_table}: {str(drop_err)}")
            
        # Giải phóng file .csv tạm thời trên ổ cứng Server
        if tf_path and os.path.exists(tf_path):
            try:
                os.remove(tf_path)
            except Exception:
                pass
    
def validate_csv_columns(df, required_columns):
    """Kiểm tra sự tồn tại của các cột bắt buộc"""
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        return False, f"Thiếu các cột: {', '.join(missing_cols)}"
    return True, None

def validate_csv_physical_structure(file_storage):
    """
    Kiểm tra cấu trúc header an toàn, tương thích mọi phiên bản Python,
    Hỗ trợ thông minh cả luồng bytes (Web/Flask) và luồng str (Thư mục Server)
    Khắc phục triệt để lỗi chặt đôi byte ký tự tiếng Việt (UTF-8).
    """
    original_pos = file_storage.tell() if hasattr(file_storage, 'tell') else 0
    file_storage.seek(0)
    
    try:
        first_line = file_storage.readline() # Đọc dòng đầu tiên (Header)
        if not first_line:
            return False, [], "File trống, không chứa dữ liệu"

        # KIỂM TRA ĐỘNG KIỂU DỮ LIỆU ĐỂ XỬ LÝ DECODE AN TOÀN
        if isinstance(first_line, bytes):
            # Nếu là bytes (Nạp từ giao diện Web), tiến hành giải mã UTF-8
            header_text = first_line.decode('utf-8-sig', errors='ignore').strip()
        else:
            # Nếu đã là str (Nạp từ thư mục Server), chỉ cần loại bỏ khoảng trắng dư thừa
            header_text = first_line.strip()

        # Tách các cột theo dấu phẩy và chuyển thành chữ in hoa
        header = [col.strip().upper() for col in header_text.split(',')]
        
        if not header or header == ['']:
            return False, [], "File không chứa dòng tiêu đề hợp lệ"
            
        # Tìm danh sách các cột chuẩn 2026 đang bị thiếu trong file cũ này
        missing_cols = [col for col in CSV_ERROR_COLUMNS if col not in header]
                
        return True, missing_cols, None
        
    except Exception as e:
        return False, [], f"Lỗi đọc header: {str(e)}"
        
    finally:
        if hasattr(file_storage, 'seek'):
            file_storage.seek(original_pos)

# file CSV có chứa dữ liệu của nhiều tháng hoặc nhiều loại báo cáo: "có bao nhiêu NGAY_BAOCAO thì thêm tương ứng bấy nhiêu dòng vào logs_nap_csv"
def verify_integrity(conn, table_name, file_name, expected_count):
    """
    Kiểm tra sự toàn vẹn: Đếm số dòng nạp vào có khớp với số dòng thực tế trong file hay không.
    Kiểm tra sự toàn vẹn: Đếm số dòng dựa trên cột hệ thống ten_file_goc
    """
    result = conn.execute(
        text(f"SELECT COUNT(*) FROM `{table_name}` WHERE `ten_file_goc` = :fname"), 
        {"fname": file_name}
    ).scalar()
    
    if result != expected_count:
        current_app.logger.error(f"CẢNH BÁO: Sai lệch dữ liệu! Table {table_name}, File {file_name}. Expected: {expected_count}, Actual: {result}")
        return False
    return True

def move_to_quarantine(file_storage, file_name, reason):
    """Di chuyển file rác vào thư mục cách ly để Admin kiểm tra lại"""
    base_dir = current_app.config.get('BASE_DIR', os.getcwd())
    quarantine_dir = os.path.join(base_dir, 'quarantine')
    os.makedirs(quarantine_dir, exist_ok=True)

    # Tạo đường dẫn lưu file
    dest_path = os.path.join(quarantine_dir, f"{datetime.now().strftime('%Y%m%d')}_{file_name}")

    # Dùng .save() của Flask FileStorage để lưu file vào ổ cứng
    try:
        # Quan trọng: Reset con trỏ file về đầu trước khi .save()
        file_storage.seek(0)
        file_storage.save(dest_path)
        current_app.logger.warning(f"File {file_name} đã bị cách ly tại {dest_path}. Lý do: {reason}")
    except Exception as e:
        current_app.logger.error(f"Lỗi khi di chuyển file vào quarantine: {str(e)}")

    return dest_path

def process_csv_error_files(db, files):
    """
    Hàm xử lý nạp CSV qua web:
    - Xử lý file lớn bằng chunking (100k dòng/lần).
    - Tự động bù cột thiếu (tối đa 3 cột).
    - Cách ly file lỗi cấu trúc vật lý.
    - Phát hiện file rỗng ruột.
    - Tách log thành từng dòng riêng biệt trong logs_nap_csv cho MỖI NGÀY BÁO CÁO của từng loại báo cáo thành công.
    """
    engine = get_bc48_engine(db)

    smtp_config = current_app.config.get('SMTP_CONFIG')
    error_messages = [] # Danh sách chứa lỗi để tổng hợp gửi 1 email duy nhất
    
    results = {"success": [], "error": []}
    ma_nhan_vien = getattr(current_user, 'ma_nhan_vien', 'Hệ thống') if current_user.is_authenticated else "Hệ thống"

    # 17 cột nghiệp vụ tĩnh từ file CSV
    col_defs = ", ".join([f"`{col}` TEXT" for col in CSV_ERROR_COLUMNS])
    # Chỉ thêm đúng 01 cột hệ thống ten_file_goc ở cuối bảng
    full_col_defs = f"{col_defs}, `ten_file_goc` VARCHAR(150) DEFAULT NULL"
    
    for file in files:
        if not file.filename.lower().endswith('.csv'): continue
        
        try:
            # 1. KIỂM TRA CẤU TRÚC VẬT LÝ
            is_valid, missing_cols, struct_error = validate_csv_physical_structure(file)
            if not is_valid:
                quarantine_path = move_to_quarantine(file, file.filename, err_msg)
                with engine.connect() as conn:
                    log_to_db(
                        conn=conn,
                        file_name=file.filename,
                        loai="NONE",
                        trang_thai='QUARANTINED',
                        ma_nv_import=ma_nhan_vien,
                        ngay_bc=None,
                        ghi_chu=f"Lỗi cấu trúc: {struct_error}",
                        so_dong=0
                    )
                    conn.commit()
                results["error"].append(f"{file.filename}: {struct_error} (Quarantined)")
                continue

            # 2. XỬ LÝ DỮ LIỆU BẰNG CHUNKSIZE (An toàn cho file 500MB)
            # Reset file pointer về đầu để read_csv đọc từ đầu
            file.seek(0)
            reader = pd.read_csv(file, dtype=str, skipinitialspace=True, chunksize=100000)
            
            rows_per_table = {}          # Lưu số dòng thực tế nạp vào từng bảng (theo tháng) để kiểm tra toàn vẹn
            logs_metadata = {}           # Lưu số dòng chi tiết theo cặp (table_name, ngay_baocao) phục vụ ghi log đa dòng
            tables_cleared = set()
            has_data = False
            
            for chunk in reader:
                if chunk.empty: continue
                
                # Bù cột thiếu trong mỗi chunk (nếu có)
                for col in missing_cols: chunk[col] = "N/A"

                # Loại bỏ các dòng mà tất cả các cột nghiệp vụ đều rỗng/NA
                chunk = chunk.dropna(how='all')
                # Hoặc loại bỏ nếu các cột chính như LOAI_BAOCAO, NGAY_BAOCAO bị null
                chunk = chunk[chunk['LOAI_BAOCAO'].notna() & chunk['NGAY_BAOCAO'].notna()]

                # Gán duy nhất cột hệ thống ten_file_goc vào DataFrame
                chunk['ten_file_goc'] = file.filename
                
                has_data = True
                
                # Gom nhóm nâng cao theo LOAI_BAOCAO và NGAY_BAOCAO cụ thể để tính log chi tiết đến từng ngày
                grouped = chunk.groupby(['LOAI_BAOCAO', 'NGAY_BAOCAO'])
                                
                with engine.connect() as conn:
                    for (loai, ngay_bc_val), group in grouped:
                        loai_str = str(loai).strip().upper()
                        ngay_str = str(ngay_bc_val).strip()
                        month = ngay_str[:6] # Lấy ra chuỗi yyyymm từ cột NGAY_BAOCAO
                        
                        if loai_str not in FILE_STRUCTURES or not re.match(r'^\d{8}$', ngay_str):
                            continue
                        
                        table_name = f"{loai_str.lower()}_{month}_error"
                        
                        # Khởi tạo bảng với cấu trúc mở rộng (gồm 17 cột nghiệp vụ + 1 cột ten_file_goc)
                        conn.execute(text(f"""
                            CREATE TABLE IF NOT EXISTS `{table_name}` (
                                {full_col_defs},
                                INDEX idx_tfg (`ten_file_goc`(50)),
                                INDEX idx_mgd (`MA_GIAO_DICH`(30))
                            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci
                        """))
                        # BẪY TỰ ĐỘNG BÙ CỘT HỆ THỐNG: Nếu bảng cũ chưa có cột 'ten_file_goc' thì tự add thêm
                        try:
                            conn.execute(text(f"ALTER TABLE `{table_name}` ADD COLUMN `ten_file_goc` VARCHAR(150) DEFAULT NULL"))
                            current_app.logger.info(f"Đã tự động bổ sung cột 'ten_file_goc' vào bảng cũ: {table_name}")
                        except Exception:
                            # Nếu bảng đã có sẵn cột này rồi, MySQL sẽ báo lỗi và Python sẽ bỏ qua an toàn
                            pass

                        # Cơ chế xóa dữ liệu cũ chống trùng lặp (Idempotency) theo ten_file_goc theo từng file trong phiên chạy
                        if table_name not in tables_cleared:
                            conn.execute(text(f"DELETE FROM `{table_name}` WHERE `ten_file_goc` = :fname"), {"fname": file.filename})
                            tables_cleared.add(table_name)

                        # Thu thập thông tin số dòng chi tiết cho từng Ngày báo cáo cụ thể
                        log_key = (table_name, ngay_str)
                        logs_metadata[log_key] = logs_metadata.get(log_key, 0) + len(group)
                        
                        # Loại bỏ cột phụ 'yyyymm' (nếu lỡ tạo) trước khi nạp qua Staging bằng LOAD DATA
                        cols_to_drop = [c for c in ['yyyymm'] if c in group.columns]
                        load_to_staging(conn, group.drop(columns=cols_to_drop), table_name, file.filename)

                        # Cập nhật số dòng đã nạp cho bảng tương ứng
                        rows_per_table[table_name] = rows_per_table.get(table_name, 0) + len(group)
                    conn.commit()

            # 3. KIỂM TRA FILE RỖNG RUỘT QUA GIAO DIỆN
            if not has_data:
                # KHỞI TẠO TRƯỚC ĐỂ TRÁNH LỖI REFERENCED BEFORE ASSIGNMENT
                loai_report = "UNKNOWN"
                
                # Bóc tách thông tin từ tên file phục vụ log file rỗng ruột
                match_loai = re.search(r'(ctr|dwt|eft|ptr)', file.filename, re.IGNORECASE)
                if match_loai:
                    loai_report = match_loai.group(1).upper()
                
                match_date = re.search(r'\b(20\d{4,6})\b|_(\d{6,8})_', file.filename)
                ngay_empty = None
                
                if match_date:
                    raw_date = match_date.group(1) if match_date.group(1) else match_date.group(2)
                    if len(raw_date) >= 6: ngay_empty = raw_date[:8] if len(raw_date) >= 8 else f"{raw_date[:6]}01"

                with engine.connect() as conn:
                    log_to_db(
                        conn=conn,
                        file_name=file.filename,
                        loai=loai_report, # <-- Đã an toàn, không bao giờ lo thiếu biến
                        trang_thai='EMPTY_REPORT',
                        ma_nv_import=ma_nhan_vien,
                        ngay_bc=ngay_empty,
                        ghi_chu="CSV chỉ có header, rỗng ruột",
                        so_dong=0
                    )
                    conn.commit()
                results["success"].append(f"{file.filename} (Empty)")
                continue

            # 4. KIỂM TRA TOÀN VỆN VÀ GHI LOG THÀNH CÔNG
            with engine.connect() as conn:
                any_integrity_error = False

                # Bước kiểm định chất lượng toàn bộ dữ liệu trên cấp độ bảng vật lý trước khi ghi log
                for table_name, count_expected in rows_per_table.items():
                    if not verify_integrity(conn, table_name, file.filename, count_expected):
                        any_integrity_error = True
                        current_app.logger.error(f"Lỗi toàn vẹn dữ liệu tại bảng: {table_name} - File: {file.filename}")

                if not any_integrity_error:
                    # ĐỘT PHÁ QUAN TRỌNG: Ghi thành từng dòng riêng biệt cho MỖI NGÀY BÁO CÁO có trong dữ liệu
                    for (table_name, ngay_bc_log), count_actual in logs_metadata.items():
                        loai_bc_log = table_name.split('_')[0].upper()
                        
                        log_to_db(
                            conn=conn,
                            file_name=file.filename,
                            loai=loai_bc_log,
                            trang_thai='SUCCESS',
                            ma_nv_import=ma_nhan_vien,
                            ngay_bc=ngay_bc_log,
                            ghi_chu=f"Đã nạp thành công dữ liệu ngày {ngay_bc_log} vào bảng `{table_name}`",
                            so_dong=count_actual
                        )
                    conn.commit()
                    results["success"].append(file.filename)
                else:
                    # Ghi log thất bại nếu phát hiện sai lệch dòng tại bất kỳ bảng nào
                    # --- TỰ ĐỘNG CÁCH LY KHI SAI LỆCH SỐ DÒNG ---
                    err_msg = "Lỗi toàn vẹn: Số dòng trong bảng dữ liệu không khớp dữ liệu phân tích"
                    quarantine_path = move_to_quarantine(file, file.filename, err_msg)

                    # Ghi log thất bại chi tiết vào database để theo dõi qua giao diện
                    log_to_db(
                        conn=conn,
                        file_name=file.filename,
                        loai="MULTI",
                        trang_thai='INTEGRITY_ERROR',
                        ma_nv_import=ma_nhan_vien,
                        ngay_bc=None,
                        ghi_chu=f"{err_msg} (Quarantined)",
                        so_dong=sum(rows_per_table.values())
                    )
                    conn.commit()
                    results["error"].append(f"{file.filename}: Lỗi kiểm tra toàn vẹn dữ liệu ở một trong các bảng")

                    # Rất quan trọng: Bỏ qua file này để chạy tiếp file khác trong danh sách, không làm sập ứng dụng
                    continue
            
            # 5. Gửi cảnh báo nếu có bù cột
            if missing_cols and smtp_config:
                error_messages.append(f"File {file.filename}: Cảnh báo cấu trúc (Bù {len(missing_cols)} cột)")
                ##final_note = f"Bù: {len(missing_cols)} cột"
                ##send_alert_email(smtp_config, file.filename, f"Cảnh báo cấu trúc: {final_note}", ma_nhan_vien)
            
        except Exception as e:
            current_app.logger.error(f"Lỗi nạp file {file.filename}: {traceback.format_exc()}")
            results["error"].append(f"{file.filename}: {str(e)}")
            try:
                with engine.connect() as conn:
                    log_to_db(
                        conn=conn,
                        file_name=file.filename,
                        loai="UNKNOWN",
                        trang_thai='ERROR',
                        ma_nv_import=ma_nhan_vien,
                        ngay_bc=None,
                        ghi_chu=str(e)[:250],
                        so_dong=0
                    )
                    conn.commit()
            except: 
                pass

    # Gửi email tổng hợp sau khi xong tất cả các file (ngoài vòng lặp for)
    if error_messages and smtp_config:
        content = "\n".join(error_messages)
        send_alert_email(smtp_config, "Tổng hợp lỗi nạp CSV", content, ma_nhan_vien)
        
    return results

def process_csv_core(engine, file_obj, filename, ma_nhan_vien):
    """
    Hàm lõi phân tích và nạp dữ liệu từ luồng file (Thư mục Server):
    - ĐỒNG BỘ NGHHIÊM NGẶT 100% logic lọc dòng trống và dòng lỗi giống luồng Web.
    - Sửa triệt để lỗi lệch dòng Expected và Actual do lệch pha khâu bù dữ liệu thô.
    """
    # 17 cột nghiệp vụ tĩnh gốc từ file CSV
    col_defs = ", ".join([f"`{col}` TEXT" for col in CSV_ERROR_COLUMNS])
    full_col_defs = f"{col_defs}, `ten_file_goc` VARCHAR(150) DEFAULT NULL"

    # 1. Kiểm tra cấu trúc header vật lý
    is_valid, missing_cols, struct_error = validate_csv_physical_structure(file_obj)
    if not is_valid:
        raise Exception(f"Lỗi cấu trúc: {struct_error}")

    file_obj.seek(0)

    # 2. Đọc file theo chunk (Để keep_default_na=True để dropna hoạt động chính xác giống luồng Web)
    reader = pd.read_csv(file_obj, dtype=str, skipinitialspace=True, keep_default_na=True, chunksize=100000)
    total_rows = 0
    inserted_logs_metadata = {}
    tables_cleared = set()

    pure_filename = os.path.basename(filename)

    # Trích xuất Loại BC và Ngày BC dự phòng từ tên file (Chỉ dùng làm cứu cánh khi file RỖNG HOÀN TOÀN)
    backup_loai_from_name = "UNKNOWN"
    match_loai = re.search(r'(ctr|dwt|eft|ptr)', pure_filename, re.IGNORECASE)
    if match_loai:
        backup_loai_from_name = match_loai.group(1).upper()
        
    backup_ngay_from_name = "00000000"
    match_date = re.search(r'\b(20\d{4,6})\b|_(\d{6,8})_', pure_filename)
    if match_date:
        raw_date = match_date.group(1) if match_date.group(1) else match_date.group(2)
        if len(raw_date) == 6:
            backup_ngay_from_name = f"{raw_date}01"
        elif len(raw_date) >= 8:
            backup_ngay_from_name = raw_date[:8]
    
    for chunk in reader:
        if chunk.empty: 
            continue
        
        # Bù các cột cấu trúc bị thiếu (nếu có)
        for col in missing_cols: 
            chunk[col] = "N/A"

        # ĐỒNG BỘ CHUẨN LUỒNG WEB: 
        # Bước 1: Loại bỏ dòng mà TẤT CẢ các cột nghiệp vụ đều là NaN/Trống
        chunk = chunk.dropna(how='all')
        
        # Bước 2: Loại bỏ thẳng tay các dòng khuyết cột chỉ số chính (LOAI_BAOCAO hoặc NGAY_BAOCAO)
        chunk = chunk[chunk['LOAI_BAOCAO'].notna() & chunk['NGAY_BAOCAO'].notna()]
        
        if chunk.empty: 
            continue

        # Ép DataFrame tuân thủ chuẩn vị trí 17 cột nghiêm ngặt
        chunk = chunk.reindex(columns=CSV_ERROR_COLUMNS)
        chunk = chunk.fillna("N/A")

        # Gán cột định danh hệ thống
        chunk['ten_file_goc'] = pure_filename

        # Gom nhóm đẩy nạp dữ liệu theo giá trị thực tế trong file
        grouped = chunk.groupby(['LOAI_BAOCAO', 'NGAY_BAOCAO'])
        
        with engine.connect() as conn:
            for (loai, ngay_bc), group in grouped:
                loai_str = str(loai).strip().upper()
                ngay_str = str(ngay_bc).strip()
                
                # BỘ LỌC CHẶN NGHIÊM NGẶT CỦA WEB: Không thuộc 4 loại báo cáo hoặc sai định dạng ngày -> BỎ QUA
                if loai_str not in ["CTR", "DWT", "EFT", "PTR"] or not re.match(r'^\d{8}$', ngay_str):
                    continue
                
                month = ngay_str[:6]
                table_name = f"{loai_str.lower()}_{month}_error"

                # Khởi tạo bảng vật lý
                conn.execute(text(f"""
                    CREATE TABLE IF NOT EXISTS `{table_name}` (
                        {full_col_defs},
                        INDEX idx_tfg (`ten_file_goc`(50)),
                        INDEX idx_mgd (`MA_GIAO_DICH`(30))
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci
                """))
                
                try:
                    conn.execute(text(f"ALTER TABLE `{table_name}` ADD COLUMN `ten_file_goc` VARCHAR(150) DEFAULT NULL"))
                except Exception:
                    pass
                
                # Khử trùng dữ liệu theo nguyên tắc Idempotency
                if table_name not in tables_cleared:
                    conn.execute(text(f"DELETE FROM `{table_name}` WHERE `ten_file_goc` = :fname"), {"fname": pure_filename})
                    tables_cleared.add(table_name)
                    
                cols_to_drop = [c for c in ['yyyymm'] if c in group.columns]
                final_insert_group = group.drop(columns=cols_to_drop)

                # NẠP DỮ LIỆU QUA STAGING
                load_to_staging(conn, final_insert_group, table_name, pure_filename)

                # Thu thập thông tin đếm dòng khớp cấu trúc log đa dòng ngoài giao diện
                log_key = (table_name, ngay_str)
                if log_key not in inserted_logs_metadata:
                    inserted_logs_metadata[log_key] = 0
                inserted_logs_metadata[log_key] += len(final_insert_group)
                
            conn.commit()
        
        total_rows += len(chunk)

    # --- TRƯỜNG HỢP A: FILE CHỈ CÓ HEADER HOẶC RỖNG SAU KHI LỌC ---
    if total_rows == 0 or not inserted_logs_metadata:
        loai_report = backup_loai_from_name
        ngay_bc = None if backup_ngay_from_name == "00000000" else backup_ngay_from_name
        return {}, 0, loai_report, ngay_bc
        
    # --- TRƯỜNG HỢP B: CÓ DỮ LIỆU THÀNH CÔNG ---
    # Ép cấu trúc trả về dạng phức hợp để không còn bị lỗi 'int' object is not subscriptable
    formatted_meta = {}
    total_actual_rows = 0
    for (t_name, n_bc), count in inserted_logs_metadata.items():
        formatted_meta[(t_name, n_bc)] = {
            "loai": t_name.split('_')[0].upper(),
            "ngay": n_bc,
            "so_dong": count
        }
        total_actual_rows += count

    return formatted_meta, total_actual_rows, None, None

# Tính năng "Nạp từ thư mục Server" dành cho các file csv 500MB
def import_csv_from_server_logic(db, folder_path):
    """
    Tính năng "Nạp tự động từ thư mục Server" dành cho các file CSV dung lượng lớn (500MB):
    - Đã sửa lỗi File Lock bằng cách đưa cấu trúc dịch chuyển file ra ngoài luồng đọc tệp.
    - Hỗ trợ toàn diện cấu trúc dữ liệu: Đơn ngày, đa ngày, đa tháng, đa loại báo cáo.
    - Tối ưu hóa I/O Buffer giúp xử lý file dung lượng lớn nhanh như môi trường Web.
    - Chuyển sang mở dạng Binary Mode ('rb') kết hợp Buffering 1MB giúp bảo toàn nguyên vẹn 100% số dòng.
    """
    archive_path = os.path.join(folder_path, 'archive')
    if not os.path.exists(archive_path): 
        os.makedirs(archive_path)
    
    results = {"success": [], "error": []}
    engine = get_bc48_engine(db)
    
    files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]
    
    for filename in files:
        # Đảm bảo lấy chính xác tên file thuần túy, loại bỏ mọi đường dẫn nếu có
        pure_filename = os.path.basename(filename)
        file_path = os.path.join(folder_path, pure_filename)
        ##file_path = os.path.join(folder_path, filename)
        is_success = False  # Biến cờ đánh dấu trạng thái xử lý thành công
        
        try:
            # GIẢI PHÁP VÀNG: Mở dạng 'rb' (Binary) để bảo toàn dữ liệu giống luồng Web,
            # kết hợp bộ đệm buffering=1024*1024 (1MB Cache) để xử lý file 500MB+ cực nhanh.
            with open(file_path, 'rb', buffering=1024*1024) as f:
                # Đưa con trỏ về đầu luồng tệp tin
                f.seek(0)

                # 1. Tận dụng hàm core nạp CSV - Nhận cấu trúc dữ liệu mới phân tách theo ngày báo cáo cụ thể
                # 1. TRUYỀN pure_filename VÀO HÀM CORE (Giống hoàn toàn luồng Web truyền file.filename)
                logs_meta, total_count, backup_loai, backup_ngay = process_csv_core(
                    engine, f, pure_filename, "System_Auto"
                )

                # 2. Log DB khi thành công
                with engine.connect() as conn:
                    # Trường hợp 1: File rỗng ruột (Chỉ có dòng tiêu đề)
                    if total_count == 0:
                        log_to_db(
                            conn=conn, 
                            file_name=pure_filename, 
                            loai=backup_loai, 
                            trang_thai='EMPTY_REPORT', 
                            ma_nv_import="System_Auto", 
                            ngay_bc=backup_ngay, 
                            so_dong=0, 
                            ghi_chu="CSV chỉ có header, rỗng ruột"
                        )
                    # Trường hợp 2: Có dữ liệu (Có thể đơn ngày hoặc đa ngày, đa tháng, đa loại)
                    else:
                        # Gom nhóm lại logs_meta theo table_name để kiểm tra toàn vẹn dữ liệu ở cấp độ bảng vật lý trước khi log
                        rows_per_table = {}
                        for (t_name, _), info in logs_meta.items():
                            rows_per_table[t_name] = rows_per_table.get(t_name, 0) + info["so_dong"]

                        all_tables_verified = True
                        for t_name, expected_count in rows_per_table.items():
                            # TRUYỀN pure_filename vào hàm xác thực toàn vẹn
                            if not verify_integrity(conn, t_name, pure_filename, expected_count):
                                all_tables_verified = False
                                break
                        
                        if all_tables_verified:
                            # Duyệt qua từng NGÀY BÁO CÁO cụ thể đã được ghi nhận thành công để tạo dòng log độc lập
                            for (t_name, ngay_bc_log), info in logs_meta.items():
                                log_to_db(
                                    conn=conn, 
                                    file_name=pure_filename, 
                                    loai=info["loai"], 
                                    trang_thai='SUCCESS', 
                                    ma_nv_import="System_Auto", 
                                    ngay_bc=ngay_bc_log, 
                                    so_dong=info["so_dong"], 
                                    ghi_chu=f"Hệ thống nạp thành công dữ liệu ngày {ngay_bc_log} vào bảng `{t_name}`" 
                                )
                        else:
                            raise Exception("Lỗi kiểm tra toàn vẹn: Số dòng trong bảng dữ liệu không khớp dữ liệu phân tích")
                    
                    conn.commit()
            
            # Cập nhật trạng thái thành công sau khi luồng 'with open' đã đóng hoàn toàn tệp
            is_success = True
            results["success"].append(pure_filename)
                
        except Exception as e:
            # IN LỖI RA CONSOLE ĐỂ BẠN ĐỌC
            print(f"--- LỖI XỬ LÝ FILE {pure_filename} ---")
            traceback.print_exc()

            # Bẫy tự động ghi log FAILED vào Database
            try:
                # 1. KHỞI TẠO GIÁ TRỊ MẶC ĐỊNH NGAY TẠI ĐÂY
                loai_report = "UNKNOWN"
                
                # Cố gắng bóc tách loại báo cáo từ tên file
                match_loai_fail = re.search(r'(ctr|dwt|eft|ptr)', pure_filename, re.IGNORECASE)
                if match_loai_fail:
                    loai_report = match_loai_fail.group(1).upper()

                # Thực hiện ghi nhận vết lỗi vào bảng log để quản trị theo dõi
                with engine.connect() as conn:
                    log_to_db(
                        conn=conn, 
                        file_name=pure_filename, 
                        loai=loai_report, 
                        trang_thai='FAILED', 
                        ma_nv_import="System_Auto", 
                        ngay_bc=None, 
                        so_dong=0, 
                        ghi_chu=f"Lỗi: {str(e)[:250]}"
                    )
                    conn.commit()
            except Exception as log_err:
                print(f"Không thể ghi log FAILED vào DB cho file {pure_filename}: {str(log_err)}")
            
            results["error"].append(f"{pure_filename}: {str(e)}")
        
        # SỬA ĐỔI QUAN TRỌNG: Lệnh di chuyển tệp nằm ngoài cấu trúc 'with open' của file_path
        if is_success:
            try:
                shutil.move(file_path, os.path.join(archive_path, pure_filename))
            except Exception as move_err:
                print(f"Không thể dịch chuyển file {pure_filename} sang thư mục archive: {str(move_err)}")

        # Cưỡng bức dọn dẹp RAM rác sau khi kết thúc một tệp tin lớn
        gc.collect()
            
    return results

####################################################################################
# Truy vấn lịch sử nạp dữ liệu TXT vào bc48; Nhật ký nạp TXT vào bc48
####################################################################################
def get_file_logs_query(start_date, end_date, loai_bc, page=1, per_page=50):
    # Tính toán offset
    offset = (page - 1) * per_page

    # Chuyển đổi date input (YYYY-MM-DD) sang định dạng YYYYMMDD để khớp với CHAR(8)
    # Ví dụ: '2026-05-06' -> '20260506'
    start_str = start_date.replace('-', '') if start_date and start_date.strip() else None
    end_str = end_date.replace('-', '') if end_date and end_date.strip() else None

    query = """
        SELECT l.file_name, l.macn, l.thoidiem, l.loai_bc, l.hinh_thuc, l.stt, 
               l.so_luong, l.user_import, l.import_date, l.header_raw, latest.status
        FROM log_file_imports l
        LEFT JOIN log_file_imports_latest latest ON l.file_name = latest.file_name
        WHERE 1=1
    """
    params = {}

    # Lọc thời gian: so sánh chuỗi YYYYMMDD
    if start_str:
        query += " AND l.thoidiem >= :start"
        params['start'] = start_str
    if end_str:
        query += " AND l.thoidiem <= :end"
        params['end'] = end_str
        
    # Lọc loại báo cáo
    if loai_bc and loai_bc.strip():
        query += " AND l.loai_bc = :loai_bc"
        params['loai_bc'] = loai_bc.strip()
        
    # Sắp xếp mới nhất lên đầu (dựa trên thoidiem và thời gian nạp file) và phân trang
    query += " ORDER BY l.thoidiem DESC, l.import_date DESC LIMIT :limit OFFSET :offset"
    params['limit'] = per_page
    params['offset'] = offset
    
    return query, params

def get_file_logs_stats(start_date, end_date, loai_bc):
    # Logic tách riêng để đếm dữ liệu
    start_str = start_date.replace('-', '') if start_date and start_date.strip() else None
    end_str = end_date.replace('-', '') if end_date and end_date.strip() else None
    
    query = """
        SELECT COUNT(DISTINCT file_name) as unique_files, COUNT(*) as total_records 
        FROM log_file_imports 
        WHERE 1=1
    """
    params = {}
    if start_str:
        query += " AND thoidiem >= :start"
        params['start'] = start_str
    if end_str:
        query += " AND thoidiem <= :end"
        params['end'] = end_str
    if loai_bc and loai_bc.strip():
        query += " AND loai_bc = :loai_bc"
        params['loai_bc'] = loai_bc.strip()
        
    return query, params

####################################################################################
# Truy vấn lịch sử nạp dữ liệu CSV vào bc48; Nhật ký nạp CSV
####################################################################################
def count_csv_logs(engine, start_date, end_date, loai_bc, ma_nhan_vien, status):
    # Xử lý date
    s_date = start_date.replace('-', '') if start_date and start_date.strip() else None
    e_date = end_date.replace('-', '') if end_date and end_date.strip() else None
    
    query = "SELECT COUNT(*) FROM logs_nap_csv WHERE 1=1"
    params = {}
    
    if s_date: query += " AND ngay_baocao >= :start"; params['start'] = s_date
    if e_date: query += " AND ngay_baocao <= :end"; params['end'] = e_date
    if loai_bc and loai_bc.strip(): query += " AND loai_bc = :loai_bc"; params['loai_bc'] = loai_bc.strip()
    
    # Dùng ma_nhan_vien
    if ma_nhan_vien and ma_nhan_vien.strip(): 
        query += " AND user_import = :ma_nhan_vien"
        params['ma_nhan_vien'] = ma_nhan_vien.strip()
    if status and status.strip(): 
        query += " AND trang_thai = :status"
        params['status'] = status.strip()
        
    with engine.connect() as conn:
        return conn.execute(text(query), params).scalar()

def get_csv_logs_data(engine, start_date, end_date, loai_bc, ma_nhan_vien, status, page=1, per_page=50, export=False):
    # Xử lý date Xử lý input date từ YYYY-MM-DD sang YYYYMMDD
    s_date = start_date.replace('-', '') if start_date and start_date.strip() else None
    e_date = end_date.replace('-', '') if end_date and end_date.strip() else None
    
    query = "SELECT * FROM logs_nap_csv WHERE 1=1"
    params = {}
    
    if s_date: query += " AND ngay_baocao >= :start"; params['start'] = s_date
    if e_date: query += " AND ngay_baocao <= :end"; params['end'] = e_date
    if loai_bc and loai_bc.strip(): query += " AND loai_bc = :loai_bc"; params['loai_bc'] = loai_bc.strip()
    
    # Dùng ma_nhan_vien
    if ma_nhan_vien and ma_nhan_vien.strip(): 
        query += " AND user_import = :ma_nhan_vien"
        params['ma_nhan_vien'] = ma_nhan_vien.strip()
    if status and status.strip(): 
        query += " AND trang_thai = :status"
        params['status'] = status.strip()

    # Chỉ áp dụng LIMIT/OFFSET khi không phải là export
    if not export:
        query += " ORDER BY ngay_baocao DESC, id DESC LIMIT :limit OFFSET :offset"
        params['limit'] = per_page + 1 # +1 để kiểm tra trang sau
        params['offset'] = (page - 1) * per_page
    else:
        # Nếu export, lấy hết không giới hạn
        query += " ORDER BY ngay_baocao DESC, id DESC"
    
    with engine.connect() as conn:
        result = conn.execute(text(query), params).fetchall()
        return [dict(row._mapping) for row in result]

def count_unique_files(engine, start_date, end_date, loai_bc, ma_nhan_vien, status):
    # Sử dụng logic xử lý điều kiện giống hệt hàm count_csv_logs
    s_date = start_date.replace('-', '') if start_date and start_date.strip() else None
    e_date = end_date.replace('-', '') if end_date and end_date.strip() else None
    
    query = "SELECT COUNT(DISTINCT file_name) FROM logs_nap_csv WHERE 1=1"
    params = {}
    
    if s_date: query += " AND ngay_baocao >= :start"; params['start'] = s_date
    if e_date: query += " AND ngay_baocao <= :end"; params['end'] = e_date
    if loai_bc and loai_bc.strip(): query += " AND loai_bc = :loai_bc"; params['loai_bc'] = loai_bc.strip()
    if ma_nhan_vien and ma_nhan_vien.strip(): 
        query += " AND user_import = :ma_nhan_vien"; params['ma_nhan_vien'] = ma_nhan_vien.strip()
    if status and status.strip(): 
        query += " AND trang_thai = :status"; params['status'] = status.strip()
        
    with engine.connect() as conn:
        return conn.execute(text(query), params).scalar()

####################################################################################
# Kết quả của CALL sp_tong_hop_sodong_bang_error;
####################################################################################
def get_tong_hop_sodong_pcrt(engine):
    query = "SELECT * FROM bang_thong_ke_loi_daily ORDER BY ngay_baocao DESC;"
    try:
        with engine.connect() as conn:
            df = pd.read_sql_query(query, conn)
            # THÊM DÒNG NÀY ĐỂ XEM TÊN CỘT TRONG LOG
            # print("DANH SÁCH CỘT:", df.columns.tolist()) 
            
            if 'thoi_gian_cap_nhat' in df.columns:
                df['thoi_gian_cap_nhat'] = df['thoi_gian_cap_nhat'].dt.strftime('%d/%m/%Y %H:%M:%S')
            return df.to_dict(orient='records')
    except Exception as e:
        print(f"Lỗi: {str(e)}")
        return []

def execute_procedure_daily_raw(engine):
    """
    Chỉ thực thi chạy Procedure kích hoạt tính toán/tổng hợp dưới Database 
    """
    query = "CALL sp_tong_hop_sodong_bang_error();"
    try:
        with engine.connect() as conn:
            # Thực thi và commit (nếu DB yêu cầu commit rõ ràng)
            conn.execute(text(query))
            return True, "Thành công"
    except Exception as e:
        print(f"Lỗi thực thi procedure: {str(e)}")
        return False, str(e)

# Chạy các lệnh
#TRUNCATE TABLE danh_sach_bang_da_kiem_tra;
#TRUNCATE TABLE log_kiem_tra_du_lieu;
#TRUNCATE TABLE danh_sach_bang_error_da_quet;
#TRUNCATE TABLE log_loi_xu_ly_bang;
#TRUNCATE TABLE bang_tonghop_ktra_gd_loi;
#TRUNCATE TABLE chi_tiet_loi_master;
#TRUNCATE TABLE log_chi_tiet_loi_phan_tach;
#TRUNCATE TABLE ket_qua_do_tim_loi;
#TRUNCATE TABLE giao_dich_error_khong_tim_thay;
#CALL sp_xuly_error_trung_tam();
#CALL sp_doi_soat_khac_phuc_loi();
#CALL sp_TuDongKiemTraDuLieu();
def execute_full_refresh_procedure(engine):
    try:
        with engine.begin() as conn:
            # 1. Tắt kiểm tra khóa ngoại bằng text()
            conn.execute(text("SET FOREIGN_KEY_CHECKS = 0"))
            
            # 2. Thực hiện TRUNCATE, trừ bảng "log_loi_logic_LoaiKH_LoaiGT"
            tables = [
                "danh_sach_bang_da_kiem_tra", "log_kiem_tra_du_lieu", 
                "danh_sach_bang_error_da_quet", "log_loi_xu_ly_bang",
                "bang_tonghop_ktra_gd_loi", "chi_tiet_loi_master",
                "log_chi_tiet_loi_phan_tach", "ket_qua_do_tim_loi",
                "giao_dich_error_khong_tim_thay"
            ]
            for table in tables:
                conn.execute(text(f"TRUNCATE TABLE {table}"))
            
            # 3. Gọi các procedure bằng text()
            conn.execute(text("CALL sp_xuly_error_trung_tam()"))
            conn.execute(text("CALL sp_doi_soat_khac_phuc_loi()"))
            conn.execute(text("CALL sp_TuDongKiemTraDuLieu()"))
            
            # 4. Bật lại kiểm tra khóa ngoại bằng text()
            conn.execute(text("SET FOREIGN_KEY_CHECKS = 1"))
            
        return True, "Thành công"
    except Exception as e:
        print(f"DEBUG ERROR: {str(e)}") 
        return False, str(e)

def get_tong_hop_loi_theo_thang(engine):
    """
    Lấy dữ liệu tổng hợp số lượng lỗi theo tháng từ bảng bang_tonghop_ktra_gd_loi.
    """
    # Gom nhóm theo tháng và cộng dồn số lượng lỗi, đồng thời gom các loại báo cáo lại thành chuỗi
    query = dedent("""
        SELECT 
            thang_nam,
            GROUP_CONCAT(DISTINCT loai_bc SEPARATOR ', ') as cac_loai_bc,
            SUM(so_luong_loi) as tong_dong_loi
        FROM bang_tonghop_ktra_gd_loi
        GROUP BY thang_nam
        ORDER BY thang_nam DESC
    """)
    
    try:
        with engine.connect() as conn:
            df = pd.read_sql_query(query, conn)
            return df.to_dict(orient='records')
    except Exception as e:
        print(f"Lỗi khi truy vấn bảng tháng bang_tonghop_ktra_gd_loi: {str(e)}")
        return []

def get_chi_tiet_loi_theo_ngay(engine):
    """
    [BỔ SUNG MỚI] Lấy dữ liệu chi tiết lỗi theo từng Ngày Báo Cáo (sắp xếp mới nhất lên đầu).
    """
    query = dedent("""
        SELECT 
            ngay_baocao,
            thang_nam,
            loai_bc,
            trang_thai,
            SUM(so_luong_loi) as so_luong_loi
        FROM bang_tonghop_ktra_gd_loi
        GROUP BY ngay_baocao, thang_nam, loai_bc, trang_thai
        ORDER BY ngay_baocao DESC, loai_bc ASC
    """)
    try:
        with engine.connect() as conn:
            df = pd.read_sql_query(query, conn)
            
            if not df.empty:
                # Chuyển đổi cột ngay_baocao (yyyy-mm-dd) sang định dạng dd/mm/yyyy để hiển thị ở giao diện
                df['ngay_baocao_fmt'] = pd.to_datetime(df['ngay_baocao']).dt.strftime('%d/%m/%Y')
            else:
                df['ngay_baocao_fmt'] = ''
                
            return df.to_dict(orient='records')
    except Exception as e:
        print(f"Lỗi khi truy vấn chi tiết ngày bang_tonghop_ktra_gd_loi: {str(e)}")
        return []

def get_filter_file_goc():
    # Lấy từ khóa người dùng gõ từ Select2 truyền lên (tham số mặc định của Select2 là 'q')
    search_term = request.args.get('q', '').strip()
    
    # Nếu người dùng chưa gõ hoặc gõ ít hơn 2 ký tự thì trả về mảng rỗng 
    # Việc này giúp chặn việc quét hàng chục triệu dòng vô ích khi vừa click vào ô
    if len(search_term) < 2:
        return jsonify({"results": []})
        
    # Câu lệnh SQL tối ưu: Chỉ lấy đích danh cột tên file, lọc theo LIKE và CHẶN LIMIT 20 dòng nhanh nhất
    query = """
        SELECT DISTINCT ten_file_goc 
        FROM chi_tiet_loi_master 
        WHERE ten_file_goc LIKE :search 
        ORDER BY ten_file_goc ASC 
        LIMIT 20
    """
    
    try:
        with db.engine.connect() as conn:
            result = conn.execute(text(query), {"search": f"%{search_term}%"}).mappings().all()
            
        # Cấu trúc kết quả đầu ra bắt buộc phải có cặp [id] và [text] để Select2 hiểu được
        output = [{"id": row["ten_file_goc"], "text": row["ten_file_goc"]} for row in result]
        return jsonify({"results": output})
        
    except Exception as e:
        # Ghi log lỗi nếu có và trả về mảng rỗng để giao diện không bị crash
        print(f"Lỗi truy vấn tên file gốc: {str(e)}")
        return jsonify({"results": []})


def get_master_filters_options(engine):
    """
    Lấy danh sách Loại BC và Tháng/Năm duy nhất trực tiếp từ DB 
    để làm dữ liệu mồi cho các ô Select box trên giao diện.
    """
    query_loai_bc = "SELECT DISTINCT loai_bc FROM chi_tiet_loi_master WHERE loai_bc IS NOT NULL ORDER BY loai_bc;"
    query_thang_nam = "SELECT DISTINCT thang_nam FROM chi_tiet_loi_master WHERE thang_nam IS NOT NULL ORDER BY thang_nam DESC;"
    
    # Khởi tạo sẵn cấu trúc, ten_file_goc để trống để Select2 Ajax tự lo khi gõ tìm kiếm
    options = {"loai_bc": [], "thang_nam": [], "ten_file_goc": []}
    try:
        with engine.connect() as conn:
            res_bc = conn.execute(text(query_loai_bc)).fetchall()
            res_thang = conn.execute(text(query_thang_nam)).fetchall()
            
            # CHỈ ĐỌC DỮ LIỆU CỦA 2 BIẾN TRÊN, KHÔNG GỌI query_ten_file Ở ĐÂY NỮA
            options["loai_bc"] = [r[0] for r in res_bc]
            options["thang_nam"] = [r[0] for r in res_thang]
    except Exception as e:
        print(f"Lỗi lấy options bộ lọc: {str(e)}")
    return options

def get_filter_file_goc_logic(engine, search_term):
    if len(search_term) < 2:
        return {"results": []}
        
    # Thêm điều kiện IS NOT NULL và loại bỏ chuỗi rỗng để sạch dữ liệu gợi ý
    query = """
        SELECT DISTINCT ten_file_goc 
        FROM chi_tiet_loi_master 
        WHERE ten_file_goc IS NOT NULL 
          AND ten_file_goc != ''
          AND ten_file_goc LIKE :search 
        ORDER BY ten_file_goc ASC 
        LIMIT 20
    """
    try:
        with engine.connect() as conn:
            result = conn.execute(text(query), {"search": f"%{search_term}%"}).mappings().all()
            
        # Kiểm tra điều kiện bảo vệ dòng dữ liệu hợp lệ
        output = [
            {"id": row["ten_file_goc"], "text": row["ten_file_goc"]} 
            for row in result 
            if row["ten_file_goc"]
        ]
        return {"results": output}
    except Exception as e:
        print(f"Lỗi truy vấn tên file gốc: {str(e)}")
        return {"results": []}

def get_chi_tiet_loi_master_serverside(engine, params):
    """
    Hàm xử lý Server-side Phân trang, Tìm kiếm, Lọc động cho DataTables
    """
    draw = int(params.get('draw', 1))
    start = int(params.get('start', 0))
    length = int(params.get('length', 10))
    search_value = params.get('search[value]', '')
    
    # Các bộ lọc nâng cao từ giao diện
    f_loai_bc = params.get('f_loai_bc', '')
    f_thang_nam = params.get('f_thang_nam', '')
    f_ngay_bc = params.get('f_ngay_bc', '')
    f_ten_file = params.get('f_ten_file', '')

    # 1. Đo đếm tổng số dòng ban đầu khi chưa lọc
    total_records_query = "SELECT COUNT(*) FROM chi_tiet_loi_master"
    
    # 2. Xây dựng mệnh đề WHERE cho bộ lọc nâng cao + ô tìm kiếm chung
    where_clauses = ["1=1"]
    sql_params = {}
    
    if f_loai_bc:
        where_clauses.append("loai_bc = :f_loai_bc")
        sql_params['f_loai_bc'] = f_loai_bc
    if f_thang_nam:
        where_clauses.append("thang_nam = :f_thang_nam")
        sql_params['f_thang_nam'] = f_thang_nam
    if f_ngay_bc:
        where_clauses.append("NGAY_BAOCAO = :f_ngay_bc")
        sql_params['f_ngay_bc'] = f_ngay_bc
    if f_ten_file: # BỔ SUNG
        where_clauses.append("ten_file_goc = :f_ten_file")
        sql_params['f_ten_file'] = f_ten_file
        
    if search_value:
        where_clauses.append("(ma_giao_dich LIKE :search OR ma_loi LIKE :search OR mota_loi LIKE :search OR ten_file_goc LIKE :search)")
        sql_params['search'] = f"%{search_value}%"
        
    where_str = " AND ".join(where_clauses)

    # 3. Đo đếm số dòng sau khi đã áp bộ lọc
    filtered_records_query = f"SELECT COUNT(*) FROM chi_tiet_loi_master WHERE {where_str}"

    # 4. Câu lệnh lấy dữ liệu trang hiện tại kèm LIMIT, OFFSET
    # Mặc định sắp xếp id DESC (mới nhất lên đầu)
    # Sử dụng 4 ký tự % để bảo toàn định dạng khi đi qua f-string và Pandas
    data_query = f"""
        SELECT 
            id, ma_giao_dich, loai_bc, thang_nam, trang_thai, ma_loi, mota_loi,
            DATE_FORMAT(NGAY_BAOCAO, '%Y-%m-%d') as NGAY_BAOCAO,
            HINHTHUC_GUI, SOLAN_GUI, ten_file_goc
        FROM chi_tiet_loi_master
        WHERE {where_str}
        ORDER BY NGAY_BAOCAO DESC, id DESC
        LIMIT :limit OFFSET :offset
    """
    sql_params['limit'] = length
    sql_params['offset'] = start

    try:
        with engine.connect() as conn:
            records_total = conn.execute(text(total_records_query)).scalar()
            records_filtered = conn.execute(text(filtered_records_query), sql_params).scalar()
            
            # Đọc dữ liệu phân trang bằng Pandas
            df = pd.read_sql_query(text(data_query), conn, params=sql_params)
            data_list = df.to_dict(orient='records')
            
            return {
                "draw": draw,
                "recordsTotal": records_total,
                "recordsFiltered": records_filtered,
                "data": data_list
            }
    except Exception as e:
        print(f"Lỗi xử lý Server-side DataTables: {str(e)}")
        return {"draw": draw, "recordsTotal": 0, "recordsFiltered": 0, "data": []}

# Tab 4: Từ ket_qua_do_tim_loi thống kê loaibc, hinhthuc, namthang, slgfile, sodong_loi, tong_quydoi_loi lấy chi tiết bảng dashboard_tke_loi_chi_tiet kết quả của sp_thong_ke_bang_loi
def get_dashboard_tke_loi_chi_tiet_serverside(engine, params):
    """
    Hàm xử lý Server-side Phân trang và Tìm kiếm cho bảng dashboard_tke_loi_chi_tiet (Tab 4)
    """
    draw = int(params.get('draw', 1))
    start = int(params.get('start', 0))
    length = int(params.get('length', 10))
    search_value = params.get('search[value]', '')
    
    # 1. Đếm tổng đại cục số bản ghi ban đầu
    total_records_query = "SELECT COUNT(*) FROM dashboard_tke_loi_chi_tiet"
    
    # 2. Xây dựng điều kiện tìm kiếm (ô Tìm kiếm chung của DataTables)
    where_clauses = ["1=1"]
    sql_params = {}
    
    if search_value:
        where_clauses.append("""
            (loai_bc LIKE :search 
             OR bang_goc_tim_thay LIKE :search 
             OR ma_loi_nhan LIKE :search 
             OR mota_loi_nhan LIKE :search)
        """)
        sql_params['search'] = f"%{search_value}%"
        
    where_str = " AND ".join(where_clauses)
    
    # 3. Đếm số lượng dòng sau khi áp bộ lọc tìm kiếm
    filtered_records_query = f"SELECT COUNT(*) FROM dashboard_tke_loi_chi_tiet WHERE {where_str}"
    
    # 4. Truy vấn tập dữ liệu phân trang thực tế
    data_query = f"""
        SELECT 
            loai_bc, bang_goc_tim_thay, nam_thang, ma_loi_nhan, mota_loi_nhan, loai_tien,
            so_luong_file, tong_so_dong_loi, tong_sotien_loi, tong_quydoi_loi,
            DATE_FORMAT(ngay_cap_nhat, '%Y-%m-%d %H:%M:%S') as ngay_cap_nhat
        FROM dashboard_tke_loi_chi_tiet
        WHERE {where_str}
        ORDER BY nam_thang DESC, loai_bc ASC, ma_loi_nhan ASC
        LIMIT :limit OFFSET :offset
    """
    sql_params['limit'] = length
    sql_params['offset'] = start
    
    try:
        with engine.connect() as conn:
            records_total = conn.execute(text(total_records_query)).scalar()
            records_filtered = conn.execute(text(filtered_records_query), sql_params).scalar()
            
            # Thực thi lấy dữ liệu trang hiện tại thông qua thư viện Pandas
            df = pd.read_sql_query(text(data_query), conn, params=sql_params)
            data_list = df.to_dict(orient='records')
            
            return {
                "draw": draw,
                "recordsTotal": records_total,
                "recordsFiltered": records_filtered,
                "data": data_list
            }
    except Exception as e:
        print(f"Lỗi xử lý Server-side Tab Thống kê mã lỗi: {str(e)}")
        return {"draw": draw, "recordsTotal": 0, "recordsFiltered": 0, "data": []}

def export_dashboard_tke_loi_chi_tiet_csv(engine):
    """
    Hàm truy vấn toàn bộ dữ liệu từ bảng dashboard_tke_loi_chi_tiet không phân trang,
    chuyển đổi cấu trúc và ép kiểu chuẩn hóa để xuất ra file CSV.
    """
    # Câu lệnh SQL quét toàn bộ bảng dữ liệu gốc của Tab 4
    query = """
        SELECT 
            loai_bc, 
            bang_goc_tim_thay, 
            nam_thang, 
            ma_loi_nhan, 
            mota_loi_nhan, 
            loai_tien,
            so_luong_file, 
            tong_so_dong_loi, 
            tong_sotien_loi, 
            tong_quydoi_loi,
            DATE_FORMAT(ngay_cap_nhat, '%Y-%m-%d %H:%M:%S') as ngay_cap_nhat
        FROM dashboard_tke_loi_chi_tiet
        ORDER BY nam_thang DESC, loai_bc ASC, ma_loi_nhan ASC
    """
    
    with engine.connect() as conn:
        # Đọc trọn vẹn data đổ vào DataFrame
        df = pd.read_sql_query(text(query), conn)
        
    # Chuẩn hóa tên cột dịch sang tiếng Việt tương ứng với giao diện bảng trên HTML của anh
    rename_mapping = {
        'loai_bc': 'Loại BC',
        'bang_goc_tim_thay': 'Bảng Gốc',
        'nam_thang': 'Năm Tháng',
        'ma_loi_nhan': 'Mã Lỗi Nhận',
        'mota_loi_nhan': 'Mô Tả Lỗi Nhận',
        'loai_tien': 'Loại Tiền',
        'so_luong_file': 'Số Lượng File',
        'tong_so_dong_loi': 'Tổng Dòng Lỗi',
        'tong_sotien_loi': 'Tổng Số Tiền Lỗi',
        'tong_quydoi_loi': 'Tổng Quy Đổi Lỗi',
        'ngay_cap_nhat': 'Ngày Cập Nhật'
    }
    df.rename(columns=rename_mapping, inplace=True)
    
    # Xử lý ghi dữ liệu ra Stream bộ nhớ RAM
    output = io.BytesIO()
    # Ghi mã Byte Order Mark (BOM) để Microsoft Excel hiển thị đúng font Tiếng Việt
    output.write(b'\xef\xbb\xbf')
    
    # Xuất dữ liệu từ DataFrame vào file CSV nhị phân
    df.to_csv(output, index=False, encoding='utf-8', mode='ab')
    output.seek(0)
    
    return output.getvalue()



####################################################################################
####################################################################################
def get_ket_qua_do_tim_data(engine, loai_bc=None, thang_nam=None, ma_hieu_1=None, ma_hieu_2=None, ma_giao_dich=None, mota_loi=None, page=1, per_page=50, tu_ngay=None, den_ngay=None, ten_file_goc=None, ten_file_error=None, export=False):
    params = {}
    # Truy vấn lấy toàn bộ các cột bao gồm cả ten_ma_hieu_1, ten_ma_hieu_2
    query = "SELECT * FROM ket_qua_do_tim_loi WHERE 1=1"

    if loai_bc and loai_bc != "ALL":
        query += " AND loai_bc = :loai_bc"
        params['loai_bc'] = loai_bc
    
    # SỬA TẠI ĐÂY: Ép kiểu thang_nam sang int để khớp với cột INT dưới DB
    if thang_nam:
        val_thang = str(thang_nam).strip()
        if val_thang.isdigit() and len(val_thang) == 6:
            # Ép kiểu hẳn sang int trong câu nối chuỗi để MySQL nhận diện Index INT trực tiếp
            query += f" AND thang_nam = {int(val_thang)}"
            
    if ma_hieu_1:
        query += " AND ma_hieu_1 LIKE :ma_hieu_1"
        params['ma_hieu_1'] = f"%{ma_hieu_1}%"
    if ma_hieu_2:
        query += " AND ma_hieu_2 LIKE :ma_hieu_2"
        params['ma_hieu_2'] = f"%{ma_hieu_2}%"
    if ma_giao_dich:
        query += " AND ma_giao_dich LIKE :ma_giao_dich"
        params['ma_giao_dich'] = f"%{ma_giao_dich}%"
    if mota_loi:
        query += " AND mota_loi LIKE :mota_loi"
        params['mota_loi'] = f"%{mota_loi}%"
    if tu_ngay and tu_ngay.strip() != "":
        query += " AND ngay_baocao >= :tu_ngay"
        params['tu_ngay'] = tu_ngay.strip()
    if den_ngay and den_ngay.strip() != "":
        query += " AND ngay_baocao <= :den_ngay"
        params['den_ngay'] = den_ngay.strip()
    # BỔ SUNG MỚI
    if ten_file_goc and ten_file_goc.strip() != "":
        query += " AND ten_file_goc LIKE :ten_file_goc"
        params['ten_file_goc'] = f"%{ten_file_goc.strip()}%"
    if ten_file_error and ten_file_error.strip() != "":
        query += " AND ten_file_error LIKE :ten_file_error"
        params['ten_file_error'] = f"%{ten_file_error.strip()}%"

    query += " ORDER BY id DESC"

    if not export:
        # Logic lấy dư 1 dòng để kiểm tra has_next
        offset = (page - 1) * per_page
        query += f" LIMIT {per_page + 1} OFFSET {offset}"

    with engine.connect() as conn:
        result = conn.execute(text(query), params)
        # Sử dụng _mapping để chuyển đổi Row sang Dict cho Jinja2 dễ xử lý
        return [dict(row._mapping) for row in result]

def count_ket_qua_do_tim(engine, loai_bc=None, thang_nam=None, ma_hieu_1=None, ma_hieu_2=None, ma_giao_dich=None, mota_loi=None, tu_ngay=None, den_ngay=None, ten_file_goc=None, ten_file_error=None):
    params = {}
    query = "SELECT COUNT(*) FROM ket_qua_do_tim_loi WHERE 1=1"
    
    if loai_bc and loai_bc != "ALL":
        query += " AND loai_bc = :loai_bc"
        params['loai_bc'] = loai_bc
    # SỬA TẠI ĐÂY: Đồng bộ ép kiểu int cho hàm đếm tổng số trang
    if thang_nam:
        val_thang = str(thang_nam).strip()
        if val_thang.isdigit() and len(val_thang) == 6:
            query += f" AND thang_nam = {int(val_thang)}"
            
    if ma_hieu_1:
        query += " AND ma_hieu_1 LIKE :ma_hieu_1"
        params['ma_hieu_1'] = f"%{ma_hieu_1}%"
    if ma_hieu_2:
        query += " AND ma_hieu_2 LIKE :ma_hieu_2"
        params['ma_hieu_2'] = f"%{ma_hieu_2}%"
    if ma_giao_dich:
        query += " AND ma_giao_dich LIKE :ma_giao_dich"
        params['ma_giao_dich'] = f"%{ma_giao_dich}%"
    if mota_loi:
        query += " AND mota_loi LIKE :mota_loi"
        params['mota_loi'] = f"%{mota_loi}%"
    if tu_ngay and tu_ngay.strip() != "":
        query += " AND ngay_baocao >= :tu_ngay"
        params['tu_ngay'] = tu_ngay.strip()
    if den_ngay and den_ngay.strip() != "":
        query += " AND ngay_baocao <= :den_ngay"
        params['den_ngay'] = den_ngay.strip()
    # BỔ SUNG MỚI
    if ten_file_goc and ten_file_goc.strip() != "":
        query += " AND ten_file_goc LIKE :ten_file_goc"
        params['ten_file_goc'] = f"%{ten_file_goc.strip()}%"
    if ten_file_error and ten_file_error.strip() != "":
        query += " AND ten_file_error LIKE :ten_file_error"
        params['ten_file_error'] = f"%{ten_file_error.strip()}%"
        
    with engine.connect() as conn:
        return conn.execute(text(query), params).scalar()

# Tao Procedure sp_xuly_error_trung_tam_by_date lay so luong va chi tiet trang thai (KIEM TRA GIAO DICH LOI) tu tat ca cac bang _error v12 (ket_qua_do_tim_loi), tu ngay den ngay
def execute_central_procedure_by_date(engine, tu_ngay, den_ngay):
    """
    Gọi Stored Procedure đối chiếu lỗi theo khoảng ngày chỉ định
    tu_ngay, den_ngay có định dạng chuỗi: 'YYYY-MM-DD'
    """
    query = "CALL sp_xuly_error_trung_tam_by_date(:tu_ngay, :den_ngay)"
    with engine.connect() as conn:
        # Commit tự động được kích hoạt tùy thuộc cấu hình sqlalchemy kết nối MySQL
        conn.execute(text(query), {"tu_ngay": tu_ngay, "den_ngay": den_ngay})
        conn.commit() # Đảm bảo dữ liệu được ghi nhận thay đổi ngay lập tức

####################################################################################
# Khai báo mail chi nhánh tiếp nhận csv lỗi để xử lý cập nhật
####################################################################################
def get_danh_sach_mail(engine):
    """Lấy danh sách khai báo mail của tất cả chi nhánh"""
    query = "SELECT * FROM danh_sach_mail_chi_nhanh ORDER BY ma_hieu_1"
    with engine.connect() as conn:
        result = conn.execute(text(query))
        return [dict(row._mapping) for row in result]

def save_mail_config(engine, data):
    """Lưu mới hoặc Cập nhật thông tin dựa trên ma_hieu_1"""
    query = """
        INSERT INTO danh_sach_mail_chi_nhanh (ma_hieu_1, ten_chi_nhanh, email_nhan, email_cc, trang_thai)
        VALUES (:ma_hieu_1, :ten_chi_nhanh, :email_nhan, :email_cc, :trang_thai)
        ON DUPLICATE KEY UPDATE 
            ten_chi_nhanh = :ten_chi_nhanh, 
            email_nhan = :email_nhan, 
            email_cc = :email_cc, 
            trang_thai = :trang_thai
    """
    with engine.connect() as conn:
        conn.execute(text(query), data)
        conn.commit()


####################################################################################
# Nút lệnh chỉ bóc tách mô tả lỗi của CSV process_and_save_split_errors; process_extract_errors_only
####################################################################################
#0: Mới, 1: Đã gửi mail, 2: CN đã phản hồi, 3: Đã khớp/Xong
def process_and_save_split_errors(conn, df_original):
    """
    Tách ma_loi và mota_loi, lưu vào DB và gọi Stored Procedure để map dữ liệu nghiệp vụ.
    """
    if df_original.empty:
        return pd.DataFrame()

    df = df_original.copy()
    
    # 1. Chuẩn hóa loại báo cáo (Phải khớp với ENUM trong DB)
    if 'loai_bc' in df.columns:
        df['loai_bc'] = df['loai_bc'].fillna('').astype(str).str.upper().str.strip()
    
    # 2. Tách chuỗi và cân bằng danh sách (Vectorization)
    df['ma_loi_list'] = df['ma_loi'].fillna('').astype(str).str.split(',').apply(lambda x: [i.strip() for i in x if i.strip()])
    df['mota_loi_list'] = df['mota_loi'].fillna('').astype(str).str.split(',').apply(lambda x: [i.strip() for i in x if i.strip()])

    len_ma = df['ma_loi_list'].str.len()
    len_mota = df['mota_loi_list'].str.len()
    max_len = np.maximum(len_ma, len_mota)

    # Dùng list comprehension nhanh hơn cho việc padding
    df['ma_loi_list'] = [l + ['N/A'] * (m - len(l)) for l, m in zip(df['ma_loi_list'], max_len)]
    df['mota_loi_list'] = [l + ['Chưa có mô tả'] * (m - len(l)) for l, m in zip(df['mota_loi_list'], max_len)]
    
    # 3. Giải nén (Explode)
    df_split = df.explode(['ma_loi_list', 'mota_loi_list'])
    
    # 4. Chuẩn bị dữ liệu cho insert
    # Đảm bảo các cột có kiểu dữ liệu phù hợp
    insert_data = pd.DataFrame({
        'id_goc': df_split['id'].astype(int),
        'loai_bc': df_split['loai_bc'],
        'ma_giao_dich': df_split['ma_giao_dich'].astype(str),
        'macn': df_split['macn'].astype(str),
        'ma_hieu_1': df_split['ma_hieu_1'].astype(str),
        'ten_ma_hieu_1': df_split['ten_ma_hieu_1'].astype(str),
        'ma_hieu_2': df_split['ma_hieu_2'].astype(str),
        'ten_ma_hieu_2': df_split['ten_ma_hieu_2'].astype(str),
        'ma_loi_don_le': df_split['ma_loi_list'].astype(str),
        'mota_loi_don_le': df_split['mota_loi_list'].astype(str),
        'thang_nam': df_split['thang_nam'].astype(int),
        'ngay_baocao': pd.to_datetime(df_split['ngay_baocao']).dt.date,
        'trang_thai_xuly': 0 
    })

    # 5. Ghi bulk data vào DB
    # Dùng try-except để bắt lỗi kết nối
    try:
        insert_data.to_sql(
            'log_chi_tiet_loi_phan_tach', 
            con=conn, 
            if_exists='append', 
            index=False, 
            chunksize=5000, # Giảm chunksize xuống một chút để ổn định hơn cho transaction
            method='multi'
        )
    except Exception as e:
        print(f"Lỗi khi ghi dữ liệu thô vào log_chi_tiet_loi_phan_tach: {e}")
        raise e # Dừng tiến trình nếu insert thất bại

    # 6. Gọi Stored Procedure để map 4 cột nghiệp vụ (ds_...)
    try:
        # Dùng text() để thực thi câu lệnh SQL
        conn.execute(text("CALL proc_cap_nhat_ds_loi()"))
        # Nếu phiên bản SQLAlchemy bạn dùng là 1.4 trở lên, 
        # lệnh commit trên đối tượng conn có thể cần thiết nếu nó không tự động commit
        if hasattr(conn, 'commit'):
            conn.commit()
    except Exception as e:
        print(f"Lưu ý: Không thể cập nhật tự động các cột nghiệp vụ (ds_...): {str(e)}")

    # Trả về kết quả phục vụ các logic hiển thị sau đó
    return df_split

# nút lệnh "Chỉ bóc tách mota_loi của CSV" trên site: /boctach-mota-loi
#Hàm điều hướng xử lý chính process_extract_errors_only, bóc tách cột mota_loi của CSV do Cục PCRT trả ra
#0: Mới, 1: Đã gửi mail, 2: CN đã phản hồi, 3: Đã khớp/Xong
def process_extract_errors_only(engine, loai_bc, raw_thang_nam):
    """
    Logic: Bóc tách dữ liệu lỗi từ bảng 'ket_qua_do_tim_loi'.
    Chuyển sang LEFT JOIN để bóc tách 100% dữ liệu lỗi gốc, không bị nuốt dòng.
    Quy trình xử lý bóc tách gói gọn trong 1 Transaction duy nhất đảm bảo an toàn dữ liệu.
    """
    # ---------------------------------------------------------
    # BƯỚC 1: XÂY DỰNG ĐIỀU KIỆN LỌC DỮ LIỆU ĐỘNG TỪ BẢNG GỐC
    # ---------------------------------------------------------
    conditions = ["1=1"]
    params = {}

    if loai_bc and loai_bc != "ALL":
        conditions.append("e.loai_bc = :loai_bc")
        params["loai_bc"] = loai_bc

    thang_nam_int = None
    if raw_thang_nam and raw_thang_nam != "ALL":
        try:
            # Loại bỏ khoảng trắng thừa
            raw_thang_nam = str(raw_thang_nam).strip()

            # TỰ ĐỘNG PHÂN TÍCH ĐỊNH DẠNG: YYYYMM (ví dụ: 202601) hoặc YYYY-MM (ví dụ: 2026-01)
            if len(raw_thang_nam) == 6 and raw_thang_nam.isdigit():
                dt_obj = datetime.strptime(raw_thang_nam, "%Y%m")
            else:
                dt_obj = datetime.strptime(raw_thang_nam, "%Y-%m")
                
            thang_nam_int = int(dt_obj.strftime("%Y%m"))
            conditions.append("e.thang_nam = :thang_nam")
            params["thang_nam"] = thang_nam_int
        except Exception as ex:
            return f"Định dạng Tháng/Năm '{raw_thang_nam}' không hợp lệ (Cần dạng YYYYMM hoặc YYYY-MM): {str(ex)}"

    where_clause = " WHERE " + " AND ".join(conditions)
    
    query = f"""
        SELECT e.*, m.ten_chi_nhanh, m.email_nhan, m.email_cc
        FROM ket_qua_do_tim_loi e
        LEFT JOIN danh_sach_mail_chi_nhanh m ON e.ma_hieu_1 = m.ma_hieu_1
        {where_clause}
    """
    
    try:
        # Sử dụng begin() để mở một Transaction duy nhất xuyên suốt quy trình
        with engine.begin() as conn:
            
            result = conn.execute(text(query), params)
            df_all = pd.DataFrame([dict(row._mapping) for row in result])

            if df_all.empty:
                return "Không tìm thấy dữ liệu lỗi thô nào phù hợp với điều kiện đã chọn."

            list_id_goc = df_all['id'].unique().tolist()

            # ---------------------------------------------------------
            # BƯỚC 2: XÓA SẠCH DỮ LIỆU CŨ PHÁT SINH TỪ CÁC ID GỐC NÀY
            # ---------------------------------------------------------
            delete_query = text("DELETE FROM log_chi_tiet_loi_phan_tach WHERE id_goc IN :ids")
            conn.execute(delete_query, {"ids": list_id_goc})

            # ---------------------------------------------------------
            # BƯỚC 3: TRUYỀN KẾT NỐI VÀO ĐỂ BÓC TÁCH VÀ LƯU DỮ LIỆU MỚI
            # ---------------------------------------------------------
            # Đã tối ưu gán trực tiếp trang_thai_xuly = 0 vào trong hàm này
            df_processed = process_and_save_split_errors(conn, df_all)
            
            if df_processed.empty:
                raise Exception("Quá trình phân tách lỗi không sinh ra dữ liệu mới.")

            # ---------------------------------------------------------
            # BƯỚC 4: ĐÃ LƯỢC BỎ CÂU LỆNH UPDATE THỪA THÃI GÂY HOÃN TIẾN TRÌNH (TABLE LOCK)
            # ---------------------------------------------------------

        # Hết khối transaction: Toàn bộ quá trình DELETE -> INSERT được COMMIT thành công.

        # Đóng gói kết quả gửi ra giao diện HTML ngoài khối transaction
        results_summary = []
        for ma_cn in df_all['ma_hieu_1'].unique():
            df_cn = df_all[df_all['ma_hieu_1'] == ma_cn]
            
            email_dest = "Chưa cấu hình mail"
            if not df_cn['email_nhan'].empty and pd.notna(df_cn['email_nhan'].iloc[0]):
                email_dest = df_cn['email_nhan'].iloc[0]
            
            status_text = "Đã làm sạch & tách lại (Tất cả)"
            if loai_bc != "ALL" and raw_thang_nam != "ALL":
                status_text = f"Đã làm sạch & tách lại {loai_bc} ({thang_nam_int})"
            elif loai_bc != "ALL":
                status_text = f"Đã làm sạch & tách lại {loai_bc}"
            elif raw_thang_nam != "ALL":
                status_text = f"Đã làm sạch & tách lại kỳ {thang_nam_int}"

            results_summary.append({
                "ma_cn": ma_cn if ma_cn else "N/A",
                "email": email_dest,
                "status": status_text
            })
            
        return results_summary

    except Exception as e:
        return f"Lỗi hệ thống khi thực hiện bóc tách mota_loi của CSV: {str(e)}"




####################################################################################
# Theo dõi kết quả khắc phục lỗi ket_qua_do_tim_loi
####################################################################################
def get_chi_tiet_loi_new(engine, filters, page=1, per_page=30):
    """
    Hàm phân trang dữ liệu lấy từ bảng mới kết hợp LEFT JOIN theo dõi khắc phục lỗi.
    Đã bổ sung bộ lọc tương đối cho Mã Hiệu 1 và Mã Hiệu 2.
    Đã bổ sung bộ lọc chính xác tuyệt đối cho Tên file gốc, Tên file error.
    """
    offset = (page - 1) * per_page
    where_clauses = ["1=1"]
    params = {"limit": per_page, "offset": offset}
    
    # 1. Lọc theo Khoảng ngày báo cáo
    if filters.get('tu_ngay'):
        where_clauses.append("k.ngay_baocao >= :tu_ngay")
        params['tu_ngay'] = filters['tu_ngay']
    if filters.get('den_ngay'):
        where_clauses.append("k.ngay_baocao <= :den_ngay")
        params['den_ngay'] = filters['den_ngay']
        
    # 2. Lọc theo Loại báo cáo (CTR, DWT, EFT, PTR)
    if filters.get('loai_bc') and filters['loai_bc'] != 'ALL':
        where_clauses.append("k.loai_bc = :loai_bc")
        params['loai_bc'] = filters['loai_bc']
        
    # 3. Lọc theo Mã chi nhánh 8 số
    if filters.get('macn'):
        where_clauses.append("TRIM(k.macn) = :macn")
        params['macn'] = filters['macn'].strip()
        
    # 4. Lọc theo Mô tả lỗi
    if filters.get('mota_loi_don_le') and filters['mota_loi_don_le'] != 'ALL':
        where_clauses.append("k.mota_loi = :mota_loi_don_le")
        params['mota_loi_don_le'] = filters['mota_loi_don_le']

    # =========================================================================
    # BỔ SUNG: 5. Lọc theo Mã Hiệu 1 và Mã Hiệu 2 (Tìm kiếm tương đối kèm ép chuỗi)
    # =========================================================================
    if filters.get('ma_hieu_1'):
        where_clauses.append("k.ma_hieu_1 LIKE :ma_hieu_1")
        params['ma_hieu_1'] = f"%{filters['ma_hieu_1'].strip()}%"
        
    if filters.get('ma_hieu_2'):
        where_clauses.append("k.ma_hieu_2 LIKE :ma_hieu_2")
        params['ma_hieu_2'] = f"%{filters['ma_hieu_2'].strip()}%"

    # =========================================================================
    # THAY ĐỔI: Lọc CHÍNH XÁC TỪNG KÝ TỰ cho Tên file gốc và Tên file error
    # =========================================================================
    if filters.get('ten_file_goc'):
        where_clauses.append("TRIM(k.ten_file_goc) = :ten_file_goc")
        params['ten_file_goc'] = filters['ten_file_goc'].strip()

    if filters.get('ten_file_error'):
        where_clauses.append("TRIM(k.ten_file_error) = :ten_file_error")
        params['ten_file_error'] = filters['ten_file_error'].strip()
    # =========================================================================

    # 6. Lọc theo Trạng thái khắc phục lỗi từ màn hình
    if filters.get('trang_thai_khac_phuc') and filters['trang_thai_khac_phuc'] != 'ALL':
        if filters['trang_thai_khac_phuc'] == '0':
            where_clauses.append("(t.trang_thai_khac_phuc IS NULL OR t.trang_thai_khac_phuc = 'CHUA_SUA')")
        elif filters['trang_thai_khac_phuc'] == '1':
            where_clauses.append("t.trang_thai_khac_phuc = 'DANG_SUA'") 
        elif filters['trang_thai_khac_phuc'] == '2':
            where_clauses.append("t.trang_thai_khac_phuc = 'DA_SUA_CHO_QUET'")
        elif filters['trang_thai_khac_phuc'] == '3':
            where_clauses.append("t.trang_thai_khac_phuc = 'DA_HOAN_THANH'")
        
    where_str = " AND ".join(where_clauses)
    
    # Câu lệnh SQL đếm tổng số dòng phục vụ phân trang
    count_sql = text(f"""
        SELECT COUNT(*) 
        FROM ket_qua_do_tim_loi k
        LEFT JOIN theo_doi_khac_phuc_loi t 
            ON k.ma_giao_dich = t.ma_giao_dich 
            AND k.bang_goc_tim_thay = t.bang_goc_tim_thay 
            AND k.ngay_baocao = t.ngay_baocao
        WHERE {where_str}
    """)
    
    # Câu lệnh SQL lấy chi tiết dữ liệu (Đã sửa k.mota_loi thành alias mota_loi_don_le để đồng bộ)
    data_sql = text(f"""
        SELECT 
            k.id, 
            k.ngay_baocao, 
            k.loai_bc, 
            k.ma_giao_dich, 
            k.macn, 
            k.ma_hieu_1, 
            k.ten_ma_hieu_1, 
            k.ma_hieu_2, 
            k.ten_ma_hieu_2, 
            k.mota_loi as mota_loi_don_le, 
            k.sotien, 
            k.loaitien, 
            k.ma_loi_f_ao, 
            k.ds_ma_nghiep_vu, 
            k.ds_ten_nghiep_vu, 
            k.ds_ten_cot_sql, 
            k.ds_ma_quy_dinh,
            k.kieukh, 
            k.tenkh, 
            k.loaigt, 
            k.sogt, 
            k.ten_file_goc, 
            k.ten_file_error,
            COALESCE(t.trang_thai_khac_phuc, '0') as trang_thai_khac_phuc
        FROM ket_qua_do_tim_loi k
        LEFT JOIN theo_doi_khac_phuc_loi t 
            ON k.ma_giao_dich = t.ma_giao_dich 
            AND k.bang_goc_tim_thay = t.bang_goc_tim_thay 
            AND k.ngay_baocao = t.ngay_baocao
        WHERE {where_str}
        ORDER BY k.ngay_baocao DESC, k.id DESC
        LIMIT :limit OFFSET :offset
    """)
    
    with engine.connect() as conn:
        count_params = {k: v for k, v in params.items() if k not in ['limit', 'offset']}
        total_items = conn.execute(count_sql, count_params).scalar() or 0
        result_set = conn.execute(data_sql, params).mappings().fetchall()
        
    total_pages = math.ceil(total_items / per_page) if total_items > 0 else 1
    
    safe_items = []
    for r in result_set:
        row_dict = {k.lower(): v for k, v in dict(r).items()}
        safe_items.append(row_dict)
        
    return {
        'items': safe_items,
        'page': page,
        'pages': total_pages,
        'total': total_items,
        'has_prev': page > 1,
        'has_next': page < total_pages,
        'prev_num': page - 1,
        'next_num': page + 1
    }

def export_sao_ke_duy_nhat_new(engine, filters):
    """
    Xuất Excel dữ liệu lỗi duy nhất theo khoảng ngày và điều kiện lọc từ bảng ket_qua_do_tim_loi
    """    
    where_clauses = ["1=1"]
    params = {}
    
    if filters.get('tu_ngay'):
        where_clauses.append("ngay_baocao >= :tu_ngay")
        params['tu_ngay'] = filters['tu_ngay']
    if filters.get('den_ngay'):
        where_clauses.append("ngay_baocao <= :den_ngay")
        params['den_ngay'] = filters['den_ngay']
    if filters.get('macn'):
        where_clauses.append("TRIM(macn) = :macn")
        params['macn'] = filters['macn'].strip()
    if filters.get('loai_bc') and filters['loai_bc'] != 'ALL':
        where_clauses.append("loai_bc = :loai_bc")
        params['loai_bc'] = filters['loai_bc']

    if filters.get('ma_hieu_1'):
        where_clauses.append("ma_hieu_1 LIKE :ma_hieu_1")
        params['ma_hieu_1'] = f"%{filters['ma_hieu_1'].strip()}%"
    if filters.get('ma_hieu_2'):
        where_clauses.append("ma_hieu_2 LIKE :ma_hieu_2")
        params['ma_hieu_2'] = f"%{filters['ma_hieu_2'].strip()}%"

    # Đồng bộ lọc chính xác tuyệt đối
    if filters.get('ten_file_goc'):
        where_clauses.append("TRIM(ten_file_goc) = :ten_file_goc")
        params['ten_file_goc'] = filters['ten_file_goc'].strip()
    if filters.get('ten_file_error'):
        where_clauses.append("TRIM(ten_file_error) = :ten_file_error")
        params['ten_file_error'] = filters['ten_file_error'].strip()
        
    where_str = " AND ".join(where_clauses)
    
    query = f"""
        SELECT ma_giao_dich, loai_bc, thang_nam, bang_goc_tim_thay, ma_loi_f_ao, mota_loi, ngay_baocao, macn, ma_hieu_2, sotien
        FROM ket_qua_do_tim_loi 
        WHERE {where_str}
        GROUP BY ma_giao_dich, bang_goc_tim_thay, ngay_baocao
    """
    with engine.connect() as conn:
        df = pd.read_sql_query(text(query), conn, params=params)
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='SaoKeLoiDuyNhat', index=False)
    output.seek(0)
    return output

# Tách riêng theo macn từ bảng ket_qua_do_tim_loi, bổ sung 3 cột: noi_dung, trang_thai_ghi_nhan, ngay_chinh_sua
def xu_ly_tach_theo_macn(engine, filters):
    where_clauses = ["1=1"]
    params = {}
    
    # Tái sử dụng bộ lọc tương tự hàm export_sao_ke_duy_nhat_new
    if filters.get('tu_ngay'):
        where_clauses.append("ngay_baocao >= :tu_ngay")
        params['tu_ngay'] = filters['tu_ngay']
    if filters.get('den_ngay'):
        where_clauses.append("ngay_baocao <= :den_ngay")
        params['den_ngay'] = filters['den_ngay']
    if filters.get('macn'):
        where_clauses.append("TRIM(macn) = :macn")
        params['macn'] = filters['macn'].strip()
    if filters.get('loai_bc') and filters['loai_bc'] != 'ALL':
        where_clauses.append("loai_bc = :loai_bc")
        params['loai_bc'] = filters['loai_bc']

    if filters.get('ma_hieu_1'):
        where_clauses.append("ma_hieu_1 LIKE :ma_hieu_1")
        params['ma_hieu_1'] = f"%{filters['ma_hieu_1'].strip()}%"
    if filters.get('ma_hieu_2'):
        where_clauses.append("ma_hieu_2 LIKE :ma_hieu_2")
        params['ma_hieu_2'] = f"%{filters['ma_hieu_2'].strip()}%"

    # Đồng bộ lọc chính xác tuyệt đối
    if filters.get('ten_file_goc'):
        where_clauses.append("TRIM(ten_file_goc) = :ten_file_goc")
        params['ten_file_goc'] = filters['ten_file_goc'].strip()
    if filters.get('ten_file_error'):
        where_clauses.append("TRIM(ten_file_error) = :ten_file_error")
        params['ten_file_error'] = filters['ten_file_error'].strip()
        
    where_str = " AND ".join(where_clauses)
    
    query = f"""
        SELECT * FROM ket_qua_do_tim_loi 
        WHERE {where_str}
        ORDER BY macn ASC, ngay_baocao DESC
    """
    with engine.connect() as conn:
        df = pd.read_sql_query(text(query), conn, params=params)
    return df

# Nút lệnh "4. Tách riêng theo macn (ket_qua_do_tim_loi_2)" trên bc48_ket_qua_gui_mail.html, gọi Stored Procedure sp_select_by_loai_bc với danh sách tham số
def xu_ly_tach_theo_macn_loi2(engine, filters):
    # 1. Thu thập danh sách loại báo cáo
    loai_bc_input = filters.get('loai_bc', 'ALL')
    if loai_bc_input and loai_bc_input != 'ALL':
        list_loai_bc = [loai_bc_input.lower()]
    else:
        list_loai_bc = ['ctr', 'dwt', 'eft', 'ptr']
        
    tu_ngay = filters.get('tu_ngay') if filters.get('tu_ngay') else None
    den_ngay = filters.get('den_ngay') if filters.get('den_ngay') else None
    
    all_dfs = []
    
    # 2. Thực thi gọi trong mysql sp_select_by_loai_bc
    with engine.connect() as conn:
        for bc in list_loai_bc:
            try:
                query = text("CALL sp_select_by_loai_bc(:loai_bc, :tu_ngay, :den_ngay)")
                df_chunk = pd.read_sql_query(
                    query, 
                    conn, 
                    params={'loai_bc': bc, 'tu_ngay': tu_ngay, 'den_ngay': den_ngay}
                )
                if not df_chunk.empty:
                    all_dfs.append(df_chunk)
            except Exception as e:
                print(f"Lỗi khi thực thi SP cho loại {bc}: {str(e)}")
                continue

    if not all_dfs:
        return pd.DataFrame()
        
    df = pd.concat(all_dfs, ignore_index=True)
    
    # =========================================================================
    # CRITICAL FIX: HÀM KHỬ DẤU TIẾNG VIỆT & CHUẨN HÓA TÊN CỘT TUYỆT ĐỐI
    # =========================================================================
    def clean_column_name(col_name):
        col_str = str(col_name).lower().strip()
        # Thay thế các ký tự tiếng Việt phổ biến
        col_str = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', col_str)
        col_str = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', col_str)
        col_str = re.sub(r'[ìíịỉĩ]', 'i', col_str)
        col_str = re.sub(r'[òóọỏõôồốộổỗơờớợởỡ]', 'o', col_str)
        col_str = re.sub(r'[ùúụủũưừứựửữ]', 'u', col_str)
        col_str = re.sub(r'[ỳýỵỷỹ]', 'y', col_str)
        col_str = re.sub(r'[đ]', 'd', col_str)
        # Loại bỏ khoảng trắng và ký tự đặc biệt
        col_str = re.sub(r'[^a-z0-9]', '', col_str)
        return col_str

    # Áp dụng chuẩn hóa: "Mã CN" hoặc "Mã CN" -> "macn"
    df.columns = [clean_column_name(col) for col in df.columns]
    
    # 3. Thực hiện lọc dữ liệu trên bộ DataFrame đã chuẩn hóa tên cột
    if filters.get('macn') and 'macn' in df.columns:
        df = df[df['macn'].astype(str).str.strip() == filters['macn'].strip()]
        
    if filters.get('mota_loi_don_le') and filters['mota_loi_don_le'] != 'ALL' and 'motaloi' in df.columns:
        df = df[df['motaloi'] == filters['mota_loi_don_le']]

    if filters.get('ma_hieu_1') and 'mahieu1' in df.columns:
        df = df[df['mahieu1'].astype(str).str.contains(filters['ma_hieu_1'].strip(), case=False, na=False)]

    if filters.get('ten_file_goc') and 'tenfilegoc' in df.columns:
        df = df[df['tenfilegoc'].astype(str).str.strip() == filters['ten_file_goc'].strip()]
        
    return df


def tao_mat_khau_ngau_nhien(length=8):
    """Sinh mật khẩu ngẫu nhiên an toàn chữ và số cho chi nhánh"""
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(length))

def tao_mat_khau_co_dinh():
    return f"123456"

def chi_xuat_csv_ket_qua_do_tim_loi(engine, loai_bc, tu_ngay, den_ngay, loai_mat_khau="fixed", nguoi_thuc_hien="Hệ thống", filters_data=None):
    """
    Xuất file nén ZIP chứa CSV lỗi lưu tại thư mục Downloads của máy vận hành.
    Đã sửa lỗi pyzipper và chuyển đổi truy vấn sang bảng mới ket_qua_do_tim_loi.
    """
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    export_dir = os.path.join(downloads_path, "BaoCao_Loi_CSV_Mahoa", datetime.now().strftime("%Y%m%d"))
    if not os.path.exists(export_dir):
        os.makedirs(export_dir)

    # ĐỔI: Sử dụng bảng mới k (ket_qua_do_tim_loi) và mặc định điều kiện 1=1
    query_conditions = ["1=1"]
    params = {}

    if tu_ngay and str(tu_ngay).strip():
        query_conditions.append("k.ngay_baocao >= :tu_ngay")
        params["tu_ngay"] = str(tu_ngay).strip()
    if den_ngay and str(den_ngay).strip():
        query_conditions.append("k.ngay_baocao <= :den_ngay")
        params["den_ngay"] = str(den_ngay).strip()

    if loai_bc and str(loai_bc).strip().upper() != 'ALL':
        loai_bc_upper = str(loai_bc).strip().upper()
        query_conditions.append("k.loai_bc = :loai_bc")
        params['loai_bc'] = loai_bc_upper
        fn_loai_bc = loai_bc_upper
    else:
        fn_loai_bc = "ALL"

    if filters_data is not None:
        macn_val = str(filters_data.get('macn', '')).strip()
        if macn_val:
            query_conditions.append("TRIM(k.macn) = :macn")
            params['macn'] = macn_val

        # Đồng bộ bộ lọc Mã Hiệu 1 và Mã Hiệu 2 dạng tìm kiếm tương đối LIKE từ giao diện
        ma_hieu_1_val = str(filters_data.get('ma_hieu_1', '')).strip()
        if ma_hieu_1_val:
            query_conditions.append("k.ma_hieu_1 LIKE :ma_hieu_1")
            params['ma_hieu_1'] = f"%{ma_hieu_1_val}%"

        ma_hieu_2_val = str(filters_data.get('ma_hieu_2', '')).strip()
        if ma_hieu_2_val:
            query_conditions.append("k.ma_hieu_2 LIKE :ma_hieu_2")
            params['ma_hieu_2'] = f"%{ma_hieu_2_val}%"

        # Nếu có lọc theo mô tả lỗi ngoài giao diện
        mota_loi_val = filters_data.get('mota_loi_don_le')
        if mota_loi_val and mota_loi_val != 'ALL':
            query_conditions.append("k.mota_loi = :mota_loi_don_le")
            params['mota_loi_don_le'] = mota_loi_val

    where_clause = " WHERE " + " AND ".join(query_conditions)
    
    # ĐỔI: Câu lệnh SQL lấy từ bảng mới ket_qua_do_tim_loi k
    query_loi = f"""
        SELECT 
            k.id, k.ngay_baocao, k.loai_bc, k.ma_giao_dich, k.macn, 
            k.ma_hieu_1, k.ten_ma_hieu_1, k.ma_hieu_2, k.ten_ma_hieu_2, 
            k.mota_loi, k.sotien, k.loaitien, k.ma_loi_f_ao, 
            k.ds_ma_nghiep_vu, k.ds_ten_nghiep_vu, k.ds_ten_cot_sql, k.ds_ma_quy_dinh, 
            k.tenkh, k.loaigt, k.sogt, k.ten_file_goc, k.ten_file_error
        FROM ket_qua_do_tim_loi k 
        {where_clause}
    """
    
    try:
        with engine.connect() as conn:
            res_loi = conn.execute(text(query_loi), params)
            df_loi = pd.DataFrame([dict(row._mapping) for row in res_loi])
            
            if df_loi.empty:
                return "Không tìm thấy dữ liệu sai sót nào thỏa mãn các tiêu chí lọc đang chọn."

            df_loi['macn'] = df_loi['macn'].astype(str).str.strip().str.upper()
            
            df_mail_system = lay_danh_sach_mail_toan_he_thong(conn)
            if df_mail_system.empty:
                return "Cảnh báo hệ thống: Danh mục thông tin hòm thư đơn vị trống."
                
            df_mail_system['macn'] = df_mail_system['macn'].astype(str).str.strip().str.upper()
            df_processed = pd.merge(df_loi, df_mail_system, on='macn', how='left')
            df_processed['email_nhan'] = df_processed['email_nhan'].fillna('ChuaCoEmail@agribank.com.vn')

            danh_sach_file_info = []
            today_str = datetime.now().strftime("%Y%m%d")
            
            tu_ngay_fn = str(tu_ngay).replace('-', '') if tu_ngay else "START"
            den_ngay_fn = str(den_ngay).replace('-', '') if den_ngay else "END"
            time_range_str = f"{tu_ngay_fn}_TO_{den_ngay_fn}"

            with engine.begin() as transaction_conn:
                for ma_cn in df_processed['macn'].unique():
                    df_cn = df_processed[df_processed['macn'] == ma_cn].copy()
                    email_dest = df_cn['email_nhan'].iloc[0]
                    ma_cn_chuandanh = str(ma_cn).strip().zfill(8)
                    
                    loai_bc_thuc_te = str(df_cn['loai_bc'].iloc[0]).strip().upper() if 'loai_bc' in df_cn.columns else fn_loai_bc

                    csv_name = f"{loai_bc_thuc_te}_{time_range_str}_error_{ma_cn_chuandanh}_{today_str}.csv"
                    zip_name = f"{loai_bc_thuc_te}_{time_range_str}_error_{ma_cn_chuandanh}_{today_str}.zip"
                    zip_path = os.path.join(export_dir, zip_name)
                    
                    exclude_cols = ['email_nhan', 'ten_chi_nhanh', 'email_cc']
                    final_cols = [c for c in df_cn.columns if c not in exclude_cols]
                    
                    # Tạo CSV trong bộ nhớ
                    csv_buffer = io.StringIO()
                    df_cn[final_cols].to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                    csv_data_bytes = csv_buffer.getvalue().encode('utf-8-sig')

                    if loai_mat_khau == 'random':
                        password_str = tao_mat_khau_ngau_nhien(8)
                    else:
                        password_str = '123456'

                    # =========================================================================
                    # SỬA LỖI PYZIPPER: Thay thế ENCRYPTION_AES bằng WZ_AES
                    # =========================================================================
                    with pyzipper.AESZipFile(zip_path, 'w', compression=pyzipper.ZIP_DEFLATED, encryption=pyzipper.WZ_AES) as zf:
                        zf.setpassword(password_str.encode('utf-8'))
                        zf.writestr(csv_name, csv_data_bytes)
                    # =========================================================================
                    
                    query_log_pass = text("""
                        INSERT INTO log_mat_khau_xuat_file (ten_file, manh8so_moi, loai_bao_cao, mat_khau, nguoi_xuat)
                        VALUES (:ten_file, :manh8so_moi, :loai_bc, :mat_khau, :nguoi_xuat)
                    """)
                    transaction_conn.execute(query_log_pass, {
                        "ten_file": zip_name, "manh8so_moi": ma_cn_chuandanh, 
                        "loai_bc": loai_bc_thuc_te, "mat_khau": password_str, "nguoi_xuat": nguoi_thuc_hien
                    })

                    danh_sach_file_info.append({
                        "ma_cn": ma_cn_chuandanh, "email": email_dest, "zip_path": zip_path,
                        "zip_name": zip_name, "password": password_str, "tu_ngay": tu_ngay, "den_ngay": den_ngay
                    })
                    
            return danh_sach_file_info
    except Exception as e:
        return f"Lỗi hệ thống trong quy trình xuất mã hóa tệp tin: {str(e)}"
def chi_xuat_csv_bangtonghop_loi(engine, loai_bc, tu_ngay, den_ngay, loai_mat_khau="fixed", nguoi_thuc_hien="Hệ thống", filters_data=None):
    """
    Xuất file nén ZIP chứa CSV lỗi lưu tại thư mục Downloads của máy vận hành
    """
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    export_dir = os.path.join(downloads_path, "BaoCao_Loi_CSV_Mahoa", datetime.now().strftime("%Y%m%d"))
    if not os.path.exists(export_dir):
        os.makedirs(export_dir)

    query_conditions = ["trang_thai_xuly = 0"]
    params = {}

    if tu_ngay and str(tu_ngay).strip():
        query_conditions.append("ngay_baocao >= :tu_ngay")
        params["tu_ngay"] = str(tu_ngay).strip()
    if den_ngay and str(den_ngay).strip():
        query_conditions.append("ngay_baocao <= :den_ngay")
        params["den_ngay"] = str(den_ngay).strip()

    if loai_bc and str(loai_bc).strip().upper() != 'ALL':
        loai_bc_upper = str(loai_bc).strip().upper()
        query_conditions.append("loai_bc = :loai_bc")
        params['loai_bc'] = loai_bc_upper
        fn_loai_bc = loai_bc_upper
    else:
        fn_loai_bc = "ALL"

    if filters_data is not None:
        macn_val = str(filters_data.get('macn', '')).strip()
        if macn_val:
            query_conditions.append("UPPER(macn) = :macn")
            params['macn'] = macn_val.upper()

        ma_hieu_1_val = str(filters_data.get('ma_hieu_1', '')).strip()
        if ma_hieu_1_val:
            query_conditions.append("UPPER(ma_hieu_1) = :ma_hieu_1")
            params['ma_hieu_1'] = ma_hieu_1_val.upper()

        ma_hieu_2_val = str(filters_data.get('ma_hieu_2', '')).strip()
        if ma_hieu_2_val:
            query_conditions.append("UPPER(ma_hieu_2) = :ma_hieu_2")
            params['ma_hieu_2'] = ma_hieu_2_val.upper()

    where_clause = " WHERE " + " AND ".join(query_conditions) if query_conditions else ""
    query_loi = f"SELECT * FROM bang_tong_hop_loi {where_clause}"
    
    try:
        with engine.connect() as conn:
            res_loi = conn.execute(text(query_loi), params)
            df_loi = pd.DataFrame([dict(row._mapping) for row in res_loi])
            
            if df_loi.empty:
                return "Không tìm thấy dữ liệu sai sót nào thỏa mãn các tiêu chí lọc đang chọn."

            df_loi['macn'] = df_loi['macn'].astype(str).str.strip().str.upper()
            
            df_mail_system = lay_danh_sach_mail_toan_he_thong(conn)
            if df_mail_system.empty:
                return "Cảnh báo hệ thống: Danh mục thông tin hòm thư đơn vị trống."
                
            df_mail_system['macn'] = df_mail_system['macn'].astype(str).str.strip().str.upper()
            df_processed = pd.merge(df_loi, df_mail_system, on='macn', how='left')
            df_processed['email_nhan'] = df_processed['email_nhan'].fillna('ChuaCoEmail@agribank.com.vn')

            danh_sach_file_info = []
            today_str = datetime.now().strftime("%Y%m%d")
            
            tu_ngay_fn = str(tu_ngay).replace('-', '') if tu_ngay else "START"
            den_ngay_fn = str(den_ngay).replace('-', '') if den_ngay else "END"
            time_range_str = f"{tu_ngay_fn}_TO_{den_ngay_fn}"

            with engine.begin() as transaction_conn:
                for ma_cn in df_processed['macn'].unique():
                    df_cn = df_processed[df_processed['macn'] == ma_cn].copy()
                    email_dest = df_cn['email_nhan'].iloc[0]
                    ma_cn_chuandanh = str(ma_cn).strip().zfill(8)
                    
                    loai_bc_thuc_te = str(df_cn['loai_bc'].iloc[0]).strip().upper() if 'loai_bc' in df_cn.columns else fn_loai_bc

                    csv_name = f"{loai_bc_thuc_te}_{time_range_str}_error_{ma_cn_chuandanh}_{today_str}.csv"
                    zip_name = f"{loai_bc_thuc_te}_{time_range_str}_error_{ma_cn_chuandanh}_{today_str}.zip"
                    zip_path = os.path.join(export_dir, zip_name)
                    
                    exclude_cols = ['email_nhan', 'ten_chi_nhanh', 'email_cc']
                    final_cols = [c for c in df_cn.columns if c not in exclude_cols]
                    
                    # Tạo CSV trong bộ nhớ
                    csv_buffer = io.StringIO()
                    df_cn[final_cols].to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                    csv_data_bytes = csv_buffer.getvalue().encode('utf-8-sig')

                    if loai_mat_khau == 'random':
                        password_str = tao_mat_khau_ngau_nhien(8)
                    else:
                        password_str = '123456'

                    # Nén trực tiếp từ memory vào file ZIP mã hóa
                    with pyzipper.AESZipFile(zip_path, 'w', compression=pyzipper.ZIP_DEFLATED, encryption=pyzipper.ENCRYPTION_MODE_PKWARE) as zf:
                        zf.setpassword(password_str.encode('utf-8'))
                        zf.writestr(csv_name, csv_data_bytes)
                    
                    query_log_pass = text("""
                        INSERT INTO log_mat_khau_xuat_file (ten_file, manh8so_moi, loai_bao_cao, mat_khau, nguoi_xuat)
                        VALUES (:ten_file, :manh8so_moi, :loai_bc, :mat_khau, :nguoi_xuat)
                    """)
                    transaction_conn.execute(query_log_pass, {
                        "ten_file": zip_name, "manh8so_moi": ma_cn_chuandanh, 
                        "loai_bc": loai_bc_thuc_te, "mat_khau": password_str, "nguoi_xuat": nguoi_thuc_hien
                    })

                    danh_sach_file_info.append({
                        "ma_cn": ma_cn_chuandanh, "email": email_dest, "zip_path": zip_path,
                        "zip_name": zip_name, "password": password_str, "tu_ngay": tu_ngay, "den_ngay": den_ngay
                    })
                    
            return danh_sach_file_info
    except Exception as e:
        return f"Lỗi hệ thống trong quy trình xuất mã hóa tệp tin: {str(e)}"

def export_sao_ke_duy_nhat(db, filters):
    """
    Xuất dữ liệu duy nhất theo ma_giao_dich từ bảng mới ket_qua_do_tim_loi.
    Đã đồng bộ hóa logic bộ lọc (LIKE cho Mã hiệu, trạng thái khắc phục kèm bảng theo dõi).
    """
    # Làm sạch bộ lọc
    valid_filters = {k: v.strip() for k, v in filters.items() if v and str(v).strip() != "" and str(v) != 'ALL'}
    
    where_clauses = ["1=1"]
    bind_params = {}
    
    # 1. Bộ lọc cơ bản bằng dấu "="
    if valid_filters.get('loai_bc'):
        where_clauses.append("k.loai_bc = :loai_bc")
        bind_params['loai_bc'] = valid_filters['loai_bc']
        
    if valid_filters.get('macn'):
        where_clauses.append("TRIM(k.macn) = :macn")
        bind_params['macn'] = valid_filters['macn']
        
    if valid_filters.get('mota_loi_don_le'):
        where_clauses.append("k.mota_loi = :mota_loi_don_le")
        bind_params['mota_loi_don_le'] = valid_filters['mota_loi_don_le']

    # 2. Bộ lọc tìm kiếm tương đối (LIKE) cho Mã hiệu 1 và Mã hiệu 2 đồng bộ với giao diện
    if valid_filters.get('ma_hieu_1'):
        where_clauses.append("k.ma_hieu_1 LIKE :ma_hieu_1")
        bind_params['ma_hieu_1'] = f"%{valid_filters['ma_hieu_1']}%"
        
    if valid_filters.get('ma_hieu_2'):
        where_clauses.append("k.ma_hieu_2 LIKE :ma_hieu_2")
        bind_params['ma_hieu_2'] = f"%{valid_filters['ma_hieu_2']}%"
        
    # 3. Áp dụng khoảng thời gian báo cáo
    if valid_filters.get('tu_ngay'):
        where_clauses.append("k.ngay_baocao >= :tu_ngay")
        bind_params['tu_ngay'] = valid_filters['tu_ngay']
    if valid_filters.get('den_ngay'):
        where_clauses.append("k.ngay_baocao <= :den_ngay")
        bind_params['den_ngay'] = valid_filters['den_ngay']

    # 4. Áp dụng trạng thái khắc phục lỗi từ bảng liên kết (theo_doi_khac_phuc_loi)
    if valid_filters.get('trang_thai_khac_phuc'):
        tt = valid_filters['trang_thai_khac_phuc']
        if tt == '0':
            where_clauses.append("(t.trang_thai_khac_phuc IS NULL OR t.trang_thai_khac_phuc = 'CHUA_SUA')")
        elif tt == '1':
            where_clauses.append("t.trang_thai_khac_phuc = 'DANG_SUA'")
        elif tt == '2':
            where_clauses.append("t.trang_thai_khac_phuc = 'DA_SUA_CHO_QUET'")
        elif tt == '3':
            where_clauses.append("t.trang_thai_khac_phuc = 'DA_HOAN_THANH'")
        
    # Tạo chuỗi WHERE hoàn chỉnh
    where_sql = "WHERE " + " AND ".join(where_clauses)
        
    # 5. Câu SQL chuẩn hóa sử dụng CTE và LEFT JOIN bảng theo dõi
    # Phân nhóm PARTITION BY ma_giao_dich, sắp xếp id DESC để lấy bản ghi mới nhất độc nhất
    raw_sql = f"""
        WITH RankedErrors AS (
            SELECT 
                k.id, k.ngay_baocao, k.loai_bc, k.ma_giao_dich, k.macn, 
                k.ma_hieu_1, k.ten_ma_hieu_1, k.ma_hieu_2, k.ten_ma_hieu_2, 
                k.mota_loi as mota_loi_don_le, k.sotien, k.loaitien, k.ma_loi_f_ao, 
                k.ds_ma_nghiep_vu, k.ds_ten_nghiep_vu, k.ds_ten_cot_sql, k.ds_ma_quy_dinh, 
                k.tenkh, k.loaigt, k.sogt, k.ten_file_goc, k.ten_file_error,
                COALESCE(t.trang_thai_khac_phuc, '0') as trang_thai_khac_phuc,
                ROW_NUMBER() OVER (
                    PARTITION BY k.ma_giao_dich 
                    ORDER BY k.id DESC
                ) as rn
            FROM ket_qua_do_tim_loi k
            LEFT JOIN theo_doi_khac_phuc_loi t 
                ON k.ma_giao_dich = t.ma_giao_dich 
                AND k.bang_goc_tim_thay = t.bang_goc_tim_thay 
                AND k.ngay_baocao = t.ngay_baocao
            {where_sql}
        )
        SELECT * FROM RankedErrors 
        WHERE rn = 1
        ORDER BY ngay_baocao DESC, id DESC;
    """
    
    engine = get_bc48_engine(db) 
    
    with engine.connect() as conn:
        result = conn.execute(text(dedent(raw_sql)), bind_params)
        columns = result.keys()
        df = pd.DataFrame([dict(row._mapping) for row in result], columns=columns)
    
    # Loại bỏ cột bổ trợ phân vùng 'rn' trước khi viết vào Excel
    if 'rn' in df.columns:
        df = df.drop(columns=['rn'])
        
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='SaoKeDuyNhat')
    output.seek(0)
    
    return output

def process_send_mail_errors(engine, smtp_config, loai_bc, tu_ngay, den_ngay, loai_mat_khau='fixed', nguoi_thuc_hien="Hệ thống"):
    """
    Quy trình phân phối tự động: Quét lỗi -> Sinh ZIP bảo mật -> Đẩy SMTP Server nội bộ
    Tối ưu: Hoàn toàn xử lý file CSV trên RAM, hạn chế tối đa ghi đĩa cứng.
    """
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    export_dir = os.path.join(downloads_path, "BaoCao_Loi_CSV_Mahoa", datetime.now().strftime("%Y%m%d"))
    
    if not os.path.exists(export_dir):
        os.makedirs(export_dir)

    query_conditions = ["trang_thai_xuly = 0", "ngay_baocao >= :tu_ngay", "ngay_baocao <= :den_ngay"]
    params = {"tu_ngay": tu_ngay, "den_ngay": den_ngay}

    if loai_bc and loai_bc != 'ALL':
        loai_bc_upper = str(loai_bc).strip().upper()
        query_conditions.append("loai_bc = :loai_bc")
        params['loai_bc'] = loai_bc_upper
        fn_loai_bc = loai_bc_upper
    else:
        fn_loai_bc = "ALL"

    where_clause = " WHERE " + " AND ".join(query_conditions)
    query_loi = f"SELECT * FROM bang_tong_hop_loi {where_clause}"
    
    try:
        with engine.connect() as conn:
            res_loi = conn.execute(text(query_loi), params)
            df_loi = pd.DataFrame([dict(row._mapping) for row in res_loi])
            if df_loi.empty:
                return f"Không tồn tại bản ghi lỗi chưa xử lý nào trong khoảng thời gian phân vùng đã chọn."

            df_loi['macn'] = df_loi['macn'].astype(str).str.strip().str.upper()

            df_mail_system = lay_danh_sach_mail_toan_he_thong(conn)
            if df_mail_system.empty:
                return "Lỗi: Danh mục tiếp nhận hòm thư hệ thống không khả dụng."

            df_mail_system['macn'] = df_mail_system['macn'].astype(str).str.strip().str.upper()
            df_all = pd.merge(df_loi, df_mail_system, on='macn', how='left')
            df_all['email_nhan'] = df_all['email_nhan'].fillna('ChuaCoEmail@agribank.com.vn')

            # Khởi tạo SMTP một lần duy nhất cho toàn bộ loop
            server = smtplib.SMTP(smtp_config['server'], int(smtp_config['port']), timeout=240)
            if smtp_config.get('use_tls', True):
                server.starttls()
            if smtp_config.get('email') and smtp_config.get('password'):
                server.login(smtp_config['email'], smtp_config['password'])

            results_summary = []
            today_str = datetime.now().strftime("%Y%m%d")
            time_range_str = f"{tu_ngay.replace('-','')}_TO_{den_ngay.replace('-','')}"

            with engine.begin() as transaction_conn:
                for ma_don_vi in df_all['macn'].unique():
                    df_cn = df_all[df_all['macn'] == ma_don_vi].copy()
                    email_dest = df_cn['email_nhan'].iloc[0]
                    email_cc = df_cn['email_cc'].iloc[0] if 'email_cc' in df_cn.columns and pd.notna(df_cn['email_cc'].iloc[0]) else None

                    ma_cn_chuandanh = str(ma_don_vi).strip().zfill(8)
                    loai_bc_thuc_te = str(df_cn['loai_bc'].iloc[0]).strip().upper() if 'loai_bc' in df_cn.columns else fn_loai_bc

                    base_name = f"{loai_bc_thuc_te}_{time_range_str}_error_{ma_cn_chuandanh}_{today_str}"
                    csv_name = f"{base_name}.csv"
                    zip_name = f"{base_name}.zip"
                    zip_path = os.path.join(export_dir, zip_name)

                    exclude_cols = ['email_nhan', 'ten_chi_nhanh', 'email_cc']
                    # Đảm bảo loại bỏ luôn id_chi_tiet khỏi nội dung file gửi chi nhánh nếu cần thiết
                    final_cols = [c for c in df_cn.columns if c not in exclude_cols and c != 'id_chi_tiet']
                    
                    # Xuất CSV sang Buffer RAM
                    csv_buffer = io.StringIO()
                    df_cn[final_cols].to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                    csv_data_bytes = csv_buffer.getvalue().encode('utf-8-sig')

                    if loai_mat_khau == 'random':
                        secret_password = tao_mat_khau_ngau_nhien(8)
                    else:
                        secret_password = '123456'
                    
                    try:
                        # Đóng gói trực tiếp bằng Bytes thô, loại bỏ hoàn toàn việc để lại file CSV thô trên ổ đĩa
                        with pyzipper.AESZipFile(zip_path, 'w', compression=pyzipper.ZIP_DEFLATED, encryption=pyzipper.ENCRYPTION_MODE_PKWARE) as zf:
                            zf.setpassword(secret_password.encode('utf-8'))
                            zf.writestr(csv_name, csv_data_bytes)
                    except Exception as zip_err:
                        results_summary.append({"ma_cn": ma_cn_chuandanh, "email": email_dest, "status": f"Thất bại (Lỗi nén ZIP: {str(zip_err)})"})
                        continue

                    # Gọi hàm gửi Mail nội bộ
                    success = send_via_opened_server(server, smtp_config['email'], email_dest, email_cc, ma_cn_chuandanh, zip_path, loai_bc_thuc_te)
                    
                    if success:
                        # Cập nhật trạng thái xử lý lỗi theo lô (Chunking 1000 dòng một lần)
                        # Lưu ý: Cần kiểm tra kĩ bảng của anh dùng cột 'id' hay 'id_chi_tiet' làm khóa chính.
                        id_col = 'id_chi_tiet' if 'id_chi_tiet' in df_cn.columns else 'id'
                        list_id_chi_tiet = df_cn[id_col].unique().tolist()
                        chunk_size = 1000
                        for i in range(0, len(list_id_chi_tiet), chunk_size):
                            sub_ids = list_id_chi_tiet[i:i+chunk_size]
                            transaction_conn.execute(
                                text(f"UPDATE bang_tong_hop_loi SET trang_thai_xuly = 1 WHERE {id_col} IN :ids"),
                                {"ids": sub_ids}
                            )
                        
                        # Ghi nhận nhật ký mật khẩu
                        transaction_conn.execute(text("""
                            INSERT INTO log_mat_khau_xuat_file (ten_file, manh8so_moi, loai_bao_cao, mat_khau, nguoi_xuat)
                            VALUES (:ten_file, :manh8so_moi, :loai_bc, :mat_khau, :nguoi_xuat)
                        """), {
                            "ten_file": zip_name, "manh8so_moi": ma_cn_chuandanh, "loai_bc": loai_bc_thuc_te,
                            "mat_khau": secret_password, "nguoi_xuat": nguoi_thuc_hien
                        })
                        
                        log_status = "Thành công" if email_dest != 'ChuaCoEmail@agribank.com.vn' else "Thành công (Chưa cấu hình Mail)"
                        transaction_conn.execute(text("""
                            INSERT INTO log_gui_mail_bc48 (manh8so_moi, ngay_gui, email_nhan, ten_file_dinh_kem, trang_thai, nguoi_thuc_hien)
                            VALUES (:manh8so_moi, NOW(), :email_nhan, :ten_file, :trang_thai, :nguoi)
                        """), {
                            "manh8so_moi": ma_cn_chuandanh, "email_nhan": email_dest, "ten_file": zip_name,
                            "trang_thai": log_status, "nguoi": nguoi_thuc_hien
                        })

                        results_summary.append({"ma_cn": ma_cn_chuandanh, "email": email_dest, "status": log_status})
                    else:
                        err_msg = "Thất bại (Mã CN thiếu cấu hình Email tiếp nhận)" if email_dest == 'ChuaCoEmail@agribank.com.vn' else "Thất bại (Lỗi cổng SMTP Server)"
                        results_summary.append({"ma_cn": ma_cn_chuandanh, "email": email_dest, "status": err_msg})

            try:
                server.quit()
            except:
                pass
                
        return results_summary
    except Exception as e:
        return f"Lỗi hệ thống trong tiến trình phân phối thư điện tử: {str(e)}"

def send_via_opened_server(server, from_email, to_email, cc_email, ma_cn, file_path, loai_bc):
    """
    Thực hiện đính kèm tệp tin và gửi qua kết nối SMTP đang mở.
    """
    try:
        if not to_email or to_email == 'ChuaCoEmail@agribank.com.vn':
            return False

        msg = MIMEMultipart()
        msg['From'] = from_email

        recipients = []
        str_to = str(to_email).strip()
        list_to = [x.strip() for x in str_to.split(',') if x.strip()]
        if list_to:
            recipients.extend(list_to)
            msg['To'] = ", ".join(list_to)
        else:
            return False
            
        str_cc = str(cc_email).strip() if cc_email else ""
        if str_cc and str_cc.lower() != 'none':
            list_cc = [x.strip() for x in str_cc.split(',') if x.strip()]
            if list_cc:
                recipients.extend(list_cc)
                msg['Cc'] = ", ".join(list_cc)
            
        msg['Subject'] = f"[{loai_bc}] ... - ĐƠN VỊ {ma_cn} - {datetime.now().strftime('%d/%m/%Y')}"
        
        body = (f"Kính gửi Đơn vị {ma_cn},\n\n"
                f"... {loai_bc} do ....\n"
                f"...\n\n"
                f"Trân trọng.")
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        if os.path.exists(file_path):
            with open(file_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                
                ten_file_dinh_kem = os.path.basename(file_path)
                part.add_header("Content-Disposition", f"attachment; filename={ten_file_dinh_kem}")
                msg.attach(part)
        else:
            return False
                    
        server.sendmail(from_email, recipients, msg.as_string())
        return True
    except Exception as e:
        print(f"Lỗi SMTP Engine tại đơn vị {ma_cn}: {e}")
        return False

def lay_danh_sach_mail_toan_he_thong(conn):
    """
    Kết hợp lấy hòm thư mặc định và hòm thư cấu hình đặc biệt, tối ưu hiệu năng DataFrame.
    """
    q_dv = """
        SELECT TRIM(MaNH8so_moi) as macn, ten_ma_hieu_2 as ten_chi_nhanh, TRIM(mail) as email_nhan, '' as email_cc
        FROM don_vi 
        WHERE MaNH8so_moi IS NOT NULL AND MaNH8so_moi != ''
    """
    res_dv = conn.execute(text(q_dv)).fetchall()
    df_dv = pd.DataFrame([dict(row._mapping) for row in res_dv]) if res_dv else pd.DataFrame(columns=['macn', 'ten_chi_nhanh', 'email_nhan', 'email_cc'])

    q_spec = """
        SELECT TRIM(ma_hieu_1) as macn, ten_chi_nhanh, TRIM(email_nhan) as email_nhan, TRIM(email_cc) as email_cc
        FROM danh_sach_mail_chi_nhanh 
        WHERE trang_thai = 1
    """
    res_spec = conn.execute(text(q_spec)).fetchall()
    df_spec = pd.DataFrame([dict(row._mapping) for row in res_spec]) if res_spec else pd.DataFrame()

    if df_dv.empty and df_spec.empty:
        return pd.DataFrame(columns=['macn', 'ten_chi_nhanh', 'email_nhan', 'email_cc'])

    df_dv_grouped = df_dv.groupby('macn').agg({
        'ten_chi_nhanh': 'first',
        'email_nhan': lambda x: ','.join(sorted(list(set([str(i).strip() for i in x if i])))),
        'email_cc': 'first'
    }).reset_index()

    if df_spec.empty:
        return df_dv_grouped

    df_spec_grouped = df_spec.groupby('macn').agg({
        'ten_chi_nhanh': 'first',
        'email_nhan': lambda x: ','.join(sorted(list(set([str(i).strip() for i in x if i])))),
        'email_cc': lambda x: ','.join(sorted(list(set([str(i).strip() for i in x if i]))))
    }).reset_index()

    # Ưu tiên lấy cấu hình bảng đặc biệt nếu trùng mã chi nhánh
    df_dv_filtered = df_dv_grouped[~df_dv_grouped['macn'].isin(df_spec_grouped['macn'])]
    df_final = pd.concat([df_spec_grouped, df_dv_filtered], ignore_index=True)
    return df_final

def lay_danh_sach_mail_tu_don_vi(conn):
    """
    Truy vấn danh mục đơn vị đang hoạt động, chuẩn hóa mã ngân hàng 8 số mới
    và gom nhóm email để tránh nhân bản dòng lỗi khi thực hiện so khớp dữ liệu.
    """
    query = """
        SELECT TRIM(MaNH8so_moi) as macn, ten_ma_hieu_2 as ten_chi_nhanh, TRIM(mail) as email_nhan
        FROM don_vi 
        WHERE MaNH8so_moi IS NOT NULL 
          AND MaNH8so_moi != ''
          AND mail IS NOT NULL 
          AND mail != ''
    """
    res = conn.execute(text(query)).fetchall()
    if not res:
        return pd.DataFrame(columns=['macn', 'ten_chi_nhanh', 'email_nhan'])
        
    df_dv = pd.DataFrame([dict(row._mapping) for row in res])
    
    # Đảm bảo gom nhóm sạch nếu 1 mã có nhiều dòng hoặc chứa chuỗi mail phân tách dấu phẩy
    df_grouped = df_dv.groupby('macn').agg({
        'ten_chi_nhanh': 'first',
        'email_nhan': lambda x: ','.join(sorted(list(set([str(i).strip() for i in x if i]))))
    }).reset_index()
    
    return df_grouped


def gui_mail_dinh_kem_zip(mail_config, target_email, cc_email, file_path, file_name, loai_bc, ma_cn):
    """
    Hàm lõi thực hiện kết nối SMTP Server nội bộ để gửi mail đính kèm file ZIP lỗi.
    """
    try:
        # 1. Khởi tạo đối tượng Message
        msg = MIMEMultipart()
        msg['From'] = mail_config.get('SENDER_EMAIL')
        
        # Xử lý chuỗi danh sách email To và Cc
        to_list = [email.strip() for email in target_email.split(',') if email.strip()]
        cc_list = [email.strip() for email in cc_email.split(',') if email.strip()] if cc_email else []
        
        msg['To'] = ', '.join(to_list)
        if cc_list:
            msg['Cc'] = ', '.join(cc_list)
            
        # Tiêu đề thư chuẩn quy định
        msg['Subject'] = f"[{loai_bc}] Thông báo sơ xuất, sai sót dữ liệu điện tử do NHNN trả lỗi - Đơn vị {ma_cn}"

        # 2. Nội dung Email (Format HTML/Text lịch sự)
        body = f"""
        <p>Kính gửi: <b>{ma_cn}</b>,</p>
        <p>... <b>{loai_bc}</b> của đơn vị.</p>
        <p>Chi tiết dữ liệu được đính kèm trong tệp tin nén mã hóa dưới đây.</p>
        <p><i>* ...</i></p>
        <br>
        <p>Trân trọng cảm ơn./.</p>
        <p>---------------------------------------------</p>
        <p><b>HỆ THỐNG AML AUTOMATION BÁO CÁO</b></p>
        """
        msg.attach(MIMEText(body, 'html', 'utf-8'))

        # 3. Đính kèm tệp tin ZIP vật lý
        if os.path.exists(file_path):
            with open(file_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {file_name}",
                )
                msg.attach(part)
        else:
            return f"Lỗi: Không tìm thấy file tại đường dẫn {file_path}"

        # 4. Thiết lập kết nối và gửi qua SMTP Server của hệ thống
        # Gom toàn bộ người nhận thực tế bao gồm To và Cc vào một danh sách gửi nhận diện
        all_recipients = to_list + cc_list

        with smtplib.SMTP(mail_config['SMTP_SERVER'], mail_config['SMTP_PORT']) as server:
            if mail_config.get('USE_TLS'):
                server.starttls()
            if mail_config.get('SMTP_USERNAME') and mail_config.get('SMTP_PASSWORD'):
                server.login(mail_config['SMTP_USERNAME'], mail_config['SMTP_PASSWORD'])
                
            server.sendmail(mail_config['SENDER_EMAIL'], all_recipients, msg.as_string())
            
        return "SUCCESS"
    except Exception as e:
        return f"Lỗi kết nối SMTP hoặc lỗi cấu trúc Mail: {str(e)}"    




def get_chi_tiet_loi(db_session, BangTongHopLoi_model, filters, page=1, per_page=50):
    query = db_session.query(BangTongHopLoi_model)
    
    # 1. Lọc các trường cơ bản (sử dụng list comprehension để làm sạch filters)
    # Loại bỏ các key không có giá trị trước khi lặp
    valid_filters = {k: v.strip() for k, v in filters.items() if v and str(v).strip() != "" and str(v) != 'ALL'}
    
    for key in ['loai_bc', 'macn', 'ma_hieu_1', 'ma_hieu_2', 'trang_thai_xuly', 'mota_loi_don_le']:
        if key in valid_filters:
            query = query.filter(getattr(BangTongHopLoi_model, key) == valid_filters[key])
            
    # 2. Lọc ngày tháng với kiểm tra định dạng
    tu = valid_filters.get('tu_ngay')
    den = valid_filters.get('den_ngay')
    
    # Đảm bảo chuyển đổi string ngày tháng sang đối tượng date nếu cần thiết 
    # hoặc so sánh trực tiếp nếu cột trong DB là String/Date chuẩn
    try:
        if tu:
            query = query.filter(BangTongHopLoi_model.ngay_baocao >= tu)
        if den:
            query = query.filter(BangTongHopLoi_model.ngay_baocao <= den)
    except Exception as e:
        # Nếu có lỗi về định dạng ngày tháng, in log để debug nhưng không làm treo hệ thống
        print(f"Lỗi lọc ngày tháng: {e}")
        
    # 3. Trả về kết quả phân trang
    # Đảm bảo query không bị trống trước khi thực hiện paginate
    return query.order_by(BangTongHopLoi_model.ngay_baocao.desc()).paginate(
        page=page, per_page=per_page, error_out=False
    )






# Chạy Stored Procedure proc_cap_nhat_tong_hop_loi trong mysql
def run_proc_tong_hop_loi(db):
    """Logic thực thi Procedure"""
    engine = get_bc48_engine(db)
    # Kết nối tới DB và chạy Procedure
    with engine.connect() as conn:
        # Sử dụng text() để bọc câu lệnh SQL
        conn.execute(text("CALL proc_cap_nhat_tong_hop_loi()"))
        conn.commit() # Đảm bảo commit thay đổi





















def lay_danh_sach_log_loi(engine, ma_giao_dich=None, trang_thai=None, filter_loai_bc=None, 
                          ma_don_vi=None, ngay_baocao=None, ma_loi_f_ao=None, page=1, per_page=20):
    """
    Bổ sung 3 tiêu chí cốt lõi bao gồm: Mã đơn vị (macn/ma_hieu_1), Ngày báo cáo (ngay_baocao) và Mã lỗi F ảo (ma_loi_f_ao).
    Sử dụng engine chung để lấy dữ liệu bảng log_chi_tiet_loi_phan_tach và đếm tổng số dòng.
    Hỗ trợ per_page=None để lấy toàn bộ dữ liệu không giới hạn (phục vụ xuất file Excel).
    """
    sql_where = []
    params = {}

    if ma_don_vi:
        sql_where.append("(macn = :ma_don_vi OR ma_hieu_1 = :ma_don_vi)")
        params['ma_don_vi'] = ma_don_vi.strip()

    if ngay_baocao:
        sql_where.append("ngay_baocao = :ngay_baocao")
        params['ngay_baocao'] = ngay_baocao

    if ma_loi_f_ao:
        sql_where.append("ma_loi_f_ao LIKE :ma_loi_f_ao")
        params['ma_loi_f_ao'] = f"%{ma_loi_f_ao.strip()}%"
    
    # Xây dựng điều kiện lọc động
    if ma_giao_dich:
        sql_where.append("ma_giao_dich LIKE :ma_giao_dich")
        params['ma_giao_dich'] = f"%{ma_giao_dich}%"
        
    if trang_thai is not None and trang_thai != "":
        sql_where.append("trang_thai_xuly = :trang_thai_xuly")
        params['trang_thai_xuly'] = int(trang_thai)

    # Đảm bảo loại bỏ khoảng trắng và chỉ append khi biến có giá trị thực tế
    if filter_loai_bc and filter_loai_bc.strip() != "":
        sql_where.append("loai_bc = :filter_loai_bc")
        params['filter_loai_bc'] = filter_loai_bc.strip()
        
    where_clause = " WHERE " + " AND ".join(sql_where) if sql_where else ""
    
    with engine.connect() as conn:
        # 1. Câu lệnh đếm tổng số dòng (Đồng bộ chuẩn điều kiện lọc)
        sql_count = text(f"SELECT COUNT(*) as total FROM log_chi_tiet_loi_phan_tach {where_clause}")
        result_count = conn.execute(sql_count, params).mappings().fetchone()
        total_rows = result_count['total'] if result_count else 0
        
        # 2. Phân trang dữ liệu
        if per_page is not None:
            offset = (page - 1) * per_page
            params['per_page'] = per_page
            params['offset'] = offset
            limit_clause = "LIMIT :per_page OFFSET :offset"
        else:
            limit_clause = ""

        sql_data = text(f"""
            SELECT id_chi_tiet, id_goc, loai_bc, ma_giao_dich, macn, 
                   ma_hieu_1, ten_ma_hieu_1, ma_hieu_2, ten_ma_hieu_2,
                   ma_loi_don_le, mota_loi_don_le, thang_nam, ngay_baocao,
                   trang_thai_xuly, ngay_phat_hien_lai, file_phan_hoi_tu_cn, ma_loi_f_ao,
                   ds_ma_nghiep_vu, ds_ten_nghiep_vu, ds_ten_cot_sql, ds_ma_quy_dinh
            FROM log_chi_tiet_loi_phan_tach
            {where_clause}
            ORDER BY id_chi_tiet DESC
            {limit_clause}
        """)
        
        result_data = conn.execute(sql_data, params).mappings().fetchall()
        data = [dict(row) for row in result_data]
        
    return data, total_rows




def cap_nhat_trang_thai_loi(engine, id_chi_tiet, trang_thai_moi, file_phan_hoi=None):
    """
    Cập nhật tiến độ xử lý và cập nhật tên tệp phản hồi từ chi nhánh
    """
    params = {
        'trang_thai': int(trang_thai_moi),
        'id_chi_tiet': id_chi_tiet,
        'file_phan_hoi': file_phan_hoi
    }
    
    if file_phan_hoi:
        sql = text("""
            UPDATE log_chi_tiet_loi_phan_tach 
            SET trang_thai_xuly = :trang_thai, file_phan_hoi_tu_cn = :file_phan_hoi 
            WHERE id_chi_tiet = :id_chi_tiet
        """)
    else:
        sql = text("""
            UPDATE log_chi_tiet_loi_phan_tach 
            SET trang_thai_xuly = :trang_thai 
            WHERE id_chi_tiet = :id_chi_tiet
        """)
        
    # Dùng commit() bên dưới context manager của connection để xác nhận ghi xuống DB
    with engine.begin() as conn:
        conn.execute(sql, params)
        
    return True


def xuat_bao_cao_chi_tiet_loi_phan_tach(data):
    # Chuyển đổi dữ liệu từ danh sách dict sang DataFrame
    export_data = []
    for row in data:
        ngay_bc = row.get('ngay_baocao')
        ngay_bc_str = ngay_bc.strftime('%d/%m/%Y') if ngay_bc else ''
        
        export_data.append({
            'ID': row.get('id_chi_tiet', ''),
            'Loại BC': row.get('loai_bc', ''),
            'Mã Giao Dịch': row.get('ma_giao_dich', ''),
            'Mã CN': row.get('macn', ''),
            'Mã Hiệu 1': row.get('ma_hieu_1', ''),
            'Tên Mã Hiệu 1': row.get('ten_ma_hieu_1', ''),
            'Mã Hiệu 2': row.get('ma_hieu_2', ''),
            'Tên Mã Hiệu 2': row.get('ten_ma_hieu_2', ''),
            'Mã Lỗi Đơn': row.get('ma_loi_don_le', ''),
            'Mô Tả Lỗi': row.get('mota_loi_don_le', ''),
            'Mã Lỗi F Ảo': row.get('ma_loi_f_ao', ''),
            'Ngày Báo Cáo': ngay_bc_str,
            # 4 Cột nghiệp vụ mới
            'Mã Nghiệp Vụ': row.get('ds_ma_nghiep_vu', ''),
            'Tên Nghiệp Vụ': row.get('ds_ten_nghiep_vu', ''),
            'Tên Cột SQL': row.get('ds_ten_cot_sql', ''),
            'Mã Quy Định': row.get('ds_ma_quy_dinh', ''),
            'Trạng Thái': row.get('trang_thai_xuly', '')
        })
    
    df = pd.DataFrame(export_data)
    
    # Ghi vào bộ nhớ đệm
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Log_Boc_Tach_Loi')
    output.seek(0)
    
    file_name = f"Log_BocTach_MotaLoi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return output, file_name


####################################################################################
# Dashboard để theo dõi trạng thái xử lý lỗi
####################################################################################
def get_dashboard_stats(engine):
    """Lấy dữ liệu thống kê tổng hợp và chi tiết cho Dashboard"""
    query_stats = text("""
        SELECT 
            ma_hieu_1, 
            ten_ma_hieu_1,
            COUNT(*) as tong_loi,
            SUM(CASE WHEN trang_thai_xuly = 1 THEN 1 ELSE 0 END) as dang_cho_sua,
            SUM(CASE WHEN trang_thai_xuly = 3 THEN 1 ELSE 0 END) as da_khac_phuc,
            MAX(COALESCE(ngay_phat_hien_lai, ngay_tao_log)) as ngay_gan_nhat,
            ROUND(SUM(CASE WHEN trang_thai_xuly = 3 THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) as ty_le
        FROM log_chi_tiet_loi_phan_tach
        GROUP BY ma_hieu_1, ten_ma_hieu_1
        ORDER BY ty_le ASC
    """)
    
    query_total = text("""
        SELECT 
            COUNT(*) as total,
            SUM(CASE WHEN trang_thai_xuly = 1 THEN 1 ELSE 0 END) as pending,
            SUM(CASE WHEN trang_thai_xuly = 3 THEN 1 ELSE 0 END) as resolved
        FROM log_chi_tiet_loi_phan_tach
    """)

    # Lấy toàn bộ chi tiết để hiển thị ở bảng dưới cùng (cho DataTables xử lý lọc/tìm kiếm)
    query_detail = text("""
        SELECT ma_hieu_1, ten_ma_hieu_1, ma_giao_dich, ma_loi_don_le, mota_loi_don_le, 
               ngay_tao_log, ngay_phat_hien_lai, trang_thai_xuly
        FROM log_chi_tiet_loi_phan_tach
        ORDER BY ngay_tao_log DESC
    """)
    
    with engine.connect() as conn:
        stats = conn.execute(query_stats).fetchall()
        totals = conn.execute(query_total).fetchone()
        details = conn.execute(query_detail).fetchall()
        
    return stats, totals, details

################################################################################################################
# Xem bảng cau_hinh_file_nghiep_vu trên db bc48
################################################################################################################
def view_cau_hinh_nghiep_vu(db): # THÊM tham số db vào đây
    """
    Hàm thực hiện truy vấn JOIN dữ liệu bảng cấu hình file nghiệp vụ (Chỉ đọc).
    """
    # Lấy engine của db_bc48
    engine = get_bc48_engine(db)

    try:
        # Sử dụng connect() thay vì begin() cho mục đích chỉ đọc
        with engine.connect() as conn:
            sql = text("""
                SELECT 
                    ch.loai_file, 
                    ch.ma_nghiep_vu, 
                    mlnv.ten_nghiep_vu, 
                    ch.ma_truong, 
                    ch.ten_cot_sql, 
                    ch.ma_quy_dinh, 
                    bbtc.dien_giai, 
                    ch.ghi_chu
                FROM cau_hinh_file_nghiep_vu ch
                JOIN ma_loai_nghiep_vu_pcrt mlnv ON ch.ma_nghiep_vu = mlnv.ma_nghiep_vu
                JOIN batbuoc_tuychon bbtc ON ch.ma_quy_dinh = bbtc.ma_quy_dinh
            """)
    
            # Thực thi truy vấn
            result = conn.execute(sql)
            
            # Fetch toàn bộ dữ liệu vào bộ nhớ trước khi đóng kết nối
            return result.fetchall() 
            
    except Exception as e:
        traceback.print_exc()
        print(f"Lỗi tại bc48.py: {str(e)}")
        raise e

################################################################################################################
################################################################################################################
def get_log_loi_logic_data(db, page=1, per_page=50, tu_ngay=None, den_ngay=None):
    """
    Lấy dữ liệu phân trang từ bảng log_loi_logic_LoaiKH_LoaiGT có bộ lọc khoảng thời gian giao dịch gốc
    """
    offset = (page - 1) * per_page
    
    # Chuẩn hóa chuỗi ngày truyền vào tương thích với định dạng chuỗi lưu trữ trong DB (YYYYMMDD)
    # Ví dụ: '2026-05-20' -> '20260520 00:00:00'
    str_tu_ngay = f"{tu_ngay.replace('-', '')} 00:00:00" if tu_ngay else "19700101 00:00:00"
    str_den_ngay = f"{den_ngay.replace('-', '')} 23:59:59" if den_ngay else "20991231 23:59:59"
    
    try:
        engine = get_bc48_engine(db)
        with engine.connect() as conn:
            # 1. Đếm tổng số bản ghi thỏa mãn điều kiện khoảng ngày lọc
            count_sql = text("""
                SELECT COUNT(*) FROM log_loi_logic_LoaiKH_LoaiGT 
                WHERE thoidiem_gd BETWEEN :tu_ngay AND :den_ngay
            """)
            total_records = conn.execute(count_sql, {"tu_ngay": str_tu_ngay, "den_ngay": str_den_ngay}).scalar() or 0

            # 2. Lấy chi tiết bản ghi phân trang dựa trên khoảng ngày lọc
            sql = text("""
                SELECT ten_bang_goc, thoi_diem_kiem_tra, thoidiem_gd, ma_giao_dich, 
                       kieukh_loi, loaigto_loi, sogt_loi, ngaysinh_loi, ghi_chu
                FROM log_loi_logic_LoaiKH_LoaiGT
                WHERE thoidiem_gd BETWEEN :tu_ngay AND :den_ngay
                ORDER BY thoi_diem_kiem_tra DESC, id DESC
                LIMIT :limit OFFSET :offset
            """)
            result = conn.execute(sql, {
                "tu_ngay": str_tu_ngay, 
                "den_ngay": str_den_ngay, 
                "limit": per_page, 
                "offset": offset
            })
            
            # 3. Ép kiểu dữ liệu sang Dictionary List phục vụ hiển thị Jinja2
            data = []
            for row in result:
                if hasattr(row, '_mapping'):
                    data.append(dict(row._mapping))
                else:
                    data.append(dict(row))
            
            return data, total_records
    except Exception as e:
        print(f"Lỗi khi lấy dữ liệu logic KH-GT theo khoảng ngày: {str(e)}")
        raise e


def run_sp_kiem_tra_logic_LoaiKH_LoaiGT(db, tu_ngay, den_ngay):
    """
    Chủ động gọi Procedure rà soát lỗi Khách hàng - Giấy tờ đồng thời truyền 2 tham số ngày
    """
    try:
        engine = get_bc48_engine(db)
        with engine.connect() as connection:
            conn = connection.connection
            cursor = conn.cursor()
            
            try:
                # Thực hiện gọi Procedure và truyền 2 tham số (tu_ngay, den_ngay) dạng chuỗi 'YYYY-MM-DD'
                cursor.execute("CALL sp_KiemTraLogic_LoaiKH_LoaiGT(%s, %s)", (tu_ngay, den_ngay))
                
                # Giải phóng sạch toàn bộ kết quả ẩn tránh treo nghẽn cursor kết nối
                if hasattr(cursor, 'stored_results'):
                    for r in cursor.stored_results():
                        r.fetchall()
                else:
                    while cursor.nextset():
                        cursor.fetchall()
                        
                conn.commit()
                return True
            except Exception as e:
                print(f"Lỗi thực thi SQL SP KH-GT với cặp tham số ({tu_ngay} -> {den_ngay}): {str(e)}")
                conn.rollback()
                return False
            finally:
                if cursor:
                    cursor.close()
    except Exception as e:
        print(f"Error executing sp_KiemTraLogic_LoaiKH_LoaiGT: {str(e)}")
        return False

################################################################################################################
# KQ tự dò tìm lỗi logic: LoaiTien != VND và SoTien = QuyDoi ; bảng log_loi_logic_TyGia lưu kết quả chạy sp_KiemTraLogic_TyGia
# duyệt qua toàn bộ các bảng nghiệp vụ có đuôi dạng _yyyymm (bắt đầu bằng ctr_, dwt_, eft_, ptr_), 
# kiểm tra lỗi logic: loaitien khác 'VND' nhưng sotien lại bằng quydoi (ngoại tệ thì số tiền giao dịch và số tiền quy đổi ra VND không thể bằng nhau)
# TRUNCATE log_loi_logic_TyGia;
# CALL sp_KiemTraLogic_TyGia;
################################################################################################################
def get_log_loi_logic_ty_gia_data(db, page, per_page, thang_nam=None, ngay_baocao=None):
    """
    Lấy dữ liệu phân trang từ bảng log_loi_logic_TyGia sử dụng get_bc48_engine,
    hỗ trợ lọc động theo thang_nam (đuôi tên bảng) và ngay_baocao (ngày của thoi_diem_kiem_tra).
    """
    offset = (page - 1) * per_page
    
    # Xây dựng mệnh đề WHERE động dựa trên bộ lọc đầu vào
    where_clauses = []
    query_params = []
    
    if thang_nam:
        # Tên bảng gốc kết thúc bằng chuỗi yyyymm (Ví dụ: ctr_202605)
        where_clauses.append("ten_bang_goc LIKE %s")
        query_params.append(f"%_{thang_nam}")
        
    if ngay_baocao:
        # Lọc theo ngày quét hệ thống
        where_clauses.append("DATE(thoi_diem_kiem_tra) = %s")
        query_params.append(ngay_baocao)
        
    where_stmt = ""
    if where_clauses:
        where_stmt = "WHERE " + " AND ".join(where_clauses)

    try:
        engine = get_bc48_engine(db)
        with engine.connect() as connection:
            conn = connection.connection
            
            # Khởi tạo cursor dạng Dictionary tương thích với Driver đang chạy
            try:
                cursor = conn.cursor(dictionary=True)
            except (TypeError, ValueError):
                try:
                    cursor = conn.cursor(pymysql.cursors.DictCursor)
                except Exception:
                    cursor = conn.cursor()
            
            try:
                # 1. Tính tổng số bản ghi lỗi có áp dụng bộ lọc
                count_query = f"SELECT COUNT(*) AS total FROM `log_loi_logic_TyGia` {where_stmt}"
                cursor.execute(count_query, tuple(query_params) if query_params else None)
                result = cursor.fetchone()
                
                if isinstance(result, dict):
                    total_records = result.get('total', 0)
                elif isinstance(result, tuple) or isinstance(result, list):
                    total_records = result[0]
                else:
                    total_records = 0

                # 2. Lấy dữ liệu phân trang có áp dụng bộ lọc
                query = f"""
                    SELECT id, ten_bang_goc, thoi_diem_kiem_tra, ma_giao_dich, loaitien, sotien, quydoi, ghi_chu
                    FROM `log_loi_logic_TyGia`
                    {where_stmt}
                    ORDER BY thoi_diem_kiem_tra DESC, id DESC
                    LIMIT %s OFFSET %s
                """
                
                # Gom tham số lọc cùng với tham số LIMIT, OFFSET
                exec_params = query_params + [per_page, offset]
                cursor.execute(query, tuple(exec_params))
                raw_data = cursor.fetchall()
                
                # 3. Chuẩn hóa dữ liệu đầu ra thành danh sách Dict để an toàn cho HTML
                data = []
                for row in raw_data:
                    if isinstance(row, dict):
                        data.append(row)
                    else:
                        data.append({
                            'id': row[0],
                            'ten_bang_goc': row[1],
                            'thoi_diem_kiem_tra': row[2],
                            'ma_giao_dich': row[3],
                            'loaitien': row[4],
                            'sotien': row[5],
                            'quydoi': row[6],
                            'ghi_chu': row[7]
                        })
                        
                return data, total_records
                
            finally:
                if cursor:
                    cursor.close()
    except Exception as e:
        print(f"Lỗi khi lấy dữ liệu logic tỷ giá: {str(e)}")
        raise e

def run_sp_kiem_tra_logic_ty_gia(db, thang_nam=None, ngay_baocao=None):
    """
    Thực thi stored procedure sp_KiemTraLogic_TyGia kèm tham số, dọn sạch resultset ẩn.
    Trả về Tuple: (Success_Status, Error_Message_If_Any)
    """
    try:
        engine = get_bc48_engine(db)
        with engine.connect() as connection:
            conn = connection.connection
            cursor = conn.cursor()
            
            try:
                # Thực thi Procedure truyền vào 2 tham số động
                cursor.execute("CALL sp_KiemTraLogic_TyGia(%s, %s)", (thang_nam, ngay_baocao))
                
                # Giải phóng sạch toàn bộ kết quả trả về ẩn từ lệnh SELECT tổng kết ở cuối SP
                if hasattr(cursor, 'stored_results'):
                    for r in cursor.stored_results():
                        r.fetchall()
                else:
                    while cursor.nextset():
                        cursor.fetchall()
                        
                conn.commit()
                return True, None
            except Exception as e:
                error_msg = f"Error trong quá trình thực thi SQL TyGia: {str(e)}"
                print(error_msg)
                conn.rollback()
                return False, str(e)
            finally:
                if cursor:
                    cursor.close()
    except Exception as e:
        error_msg = f"Error executing sp_KiemTraLogic_TyGia: {str(e)}"
        print(error_msg)
        return False, str(e)

################################################################################################################
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
################################################################################################################
def get_log_kiem_tra_du_lieu_data(db, page, per_page, thang_nam=None, ngay_baocao=None):
    """Lấy dữ liệu phân trang từ bảng log_kiem_tra_du_lieu, lọc động theo tháng hoặc ngày phát sinh lỗi"""
    offset = (page - 1) * per_page
    where_clauses = []
    query_params = []
    
    if thang_nam:
        # Tên bảng kết thúc bằng chuỗi yyyymm (Ví dụ: ctr_202605)
        where_clauses.append("ten_bang LIKE %s")
        query_params.append(f"%_{thang_nam}")
        
    if ngay_baocao:
        # Lọc theo chuỗi ngày bắt đầu của cột thoidiem (Ví dụ: '20260520%')
        where_clauses.append("thoidiem LIKE %s")
        query_params.append(f"{ngay_baocao}%")
        
    where_stmt = "WHERE " + " AND ".join(where_clauses) if where_clauses else ""

    try:
        engine = get_bc48_engine(db) # Hàm lấy Engine sẵn có của hệ thống anh
        with engine.connect() as connection:
            conn = connection.connection
            
            # Khởi tạo Cursor Dictionary an toàn theo Driver
            try:
                cursor = conn.cursor(dictionary=True)
            except (TypeError, ValueError):
                try:
                    cursor = conn.cursor(pymysql.cursors.DictCursor)
                except Exception:
                    cursor = conn.cursor()
            
            try:
                # 1. Đếm tổng số bản ghi lỗi đáp ứng điều kiện lọc
                count_query = f"SELECT COUNT(*) AS total FROM `log_kiem_tra_du_lieu` {where_stmt}"
                cursor.execute(count_query, tuple(query_params) if query_params else None)
                result = cursor.fetchone()
                total_records = result.get('total', 0) if isinstance(result, dict) else (result[0] if result else 0)

                # 2. Lấy dữ liệu phân trang thực tế
                query = f"""
                    SELECT id, ten_bang, ten_cot, gia_tri_loi, ma_giao_dich, thoidiem, thoi_diem_kiem_tra
                    FROM `log_kiem_tra_du_lieu`
                    {where_stmt}
                    ORDER BY thoi_diem_kiem_tra DESC, id DESC
                    LIMIT %s OFFSET %s
                """
                exec_params = query_params + [per_page, offset]
                cursor.execute(query, tuple(exec_params))
                raw_data = cursor.fetchall()
                
                # 3. Chuẩn hóa dữ liệu đầu ra an toàn cho Jinja2 Template
                data = []
                for row in raw_data:
                    if isinstance(row, dict):
                        data.append(row)
                    else:
                        data.append({
                            'id': row[0],
                            'ten_bang': row[1],
                            'ten_cot': row[2],
                            'gia_tri_loi': row[3],
                            'ma_giao_dich': row[4],
                            'thoidiem': row[5],
                            'thoi_diem_kiem_tra': row[6]
                        })
                return data, total_records
                
            finally:
                if cursor:
                    cursor.close()
    except Exception as e:
        print(f"Lỗi lấy dữ liệu log_kiem_tra_du_lieu: {str(e)}")
        raise e

def process_run_check_danh_muc(db, target_date="", target_thang=None):
    """
    Chuyển đổi logic từ terminal chạy ngầm hoàn toàn bằng bộ nhớ RAM xử lý dữ liệu Pandas 
    và đồng bộ ghi nhận vào bảng log_kiem_tra_du_lieu.
    - Nếu có target_date (8 ký tự YYYYMMDD): Chỉ quét dữ liệu ngày đó trong tháng đó.
    - Nếu không có target_date nhưng có target_thang (6 ký tự YYYYMM): Chỉ quét các bảng đuôi _YYYYMM.
    - Nếu trống cả hai: Quét toàn bộ lịch sử.
    """
    try:
        engine = get_bc48_engine(db)
        with engine.connect() as conn:
            # 1. XÁC ĐỊNH REGEX ĐỂ LỌC TÊN BẢNG
            # Ưu tiên lấy tháng từ target_date trước, nếu không có thì lấy từ target_thang
            current_ym = target_date[:6] if (target_date and len(target_date) >= 6) else (target_thang if target_thang else "")
            
            if current_ym != "":
                # Chỉ lọc ra các bảng của tháng được chỉ định (Ví dụ: ctr_202604, dwt_202604...)
                regex_pattern = f"^(ctr|dwt|eft|ptr)_{current_ym}$"
            else:
                # Quét toàn bộ lịch sử tất cả các tháng
                regex_pattern = "^(ctr|dwt|eft|ptr)_[0-9]{6}$"

            query_tables = text("""
                SELECT table_name 
                FROM information_schema.tables 
                WHERE table_schema = DATABASE() AND (table_name REGEXP :pattern)
            """)
            all_tables = [row[0] for row in conn.execute(query_tables, {"pattern": regex_pattern}).fetchall()]

            # Nạp danh mục hệ thống trực tiếp vào RAM
            set_don_vi = set(pd.read_sql("SELECT MaNH8so_moi FROM don_vi", conn)["MaNH8so_moi"].astype(str).str.strip().str.lower())
            set_quoc_gia = set(pd.read_sql("SELECT ma_Alpha2 FROM quoc_gia", conn)["ma_Alpha2"].astype(str).str.strip().str.lower())
            set_lkh = set(pd.read_sql("SELECT ma_loai_khach_hang FROM loai_khach_hang", conn)["ma_loai_khach_hang"].astype(str).str.strip().str.lower())
            set_loai_tien = set(pd.read_sql("SELECT loaitien FROM loai_tien", conn)["loaitien"].astype(str).str.strip().str.lower())
            set_loaigd = set(pd.read_sql("SELECT ma_nghiep_vu FROM ma_loai_nghiep_vu_pcrt", conn)["ma_nghiep_vu"].astype(str).str.strip().str.lower())
            set_kenhct = set(pd.read_sql("SELECT ma_kenh_ct FROM kenh_chuyen_tien", conn)["ma_kenh_ct"].astype(str).str.strip().str.lower())
            set_giay_to = set(pd.read_sql("SELECT ma_loai_giay_to FROM loai_giay_to", conn)["ma_loai_giay_to"].astype(str).str.strip().str.lower())
            set_hang_hoa = set(pd.read_sql("SELECT ma_loai_hang_hoa FROM loai_hang_hoa", conn)["ma_loai_hang_hoa"].astype(str).str.strip().str.lower())
            set_tai_khoan = set(pd.read_sql("SELECT ma_loai_tai_khoan FROM loai_tai_khoan", conn)["ma_loai_tai_khoan"].astype(str).str.strip().str.lower())

            for table_name in all_tables:
                # 2. Dọn dẹp log cũ theo phạm vi thông minh (Sử dụng index cột thoidiem mới)
                if target_date != "":
                    # Nếu chạy theo ngày: Chỉ xóa log của ngày đó trong bảng hiện tại
                    sql_delete = text("DELETE FROM log_kiem_tra_du_lieu WHERE ten_bang = :t AND thoidiem LIKE :date_pattern")
                    conn.execute(sql_delete, {"t": table_name, "date_pattern": f"{target_date}%"})
                else:
                    # Nếu chạy theo Tháng hoặc Toàn bộ: Xóa sạch log cũ của riêng bảng này để nạp mới hoàn toàn
                    sql_delete = text("DELETE FROM log_kiem_tra_du_lieu WHERE ten_bang = :t")
                    conn.execute(sql_delete, {"t": table_name})
                conn.commit()
                
                # 3. Đọc dữ liệu thô từ bảng giao dịch
                if target_date != "":
                    sql_select = text(f"SELECT * FROM {table_name} WHERE thoidiem LIKE :date_pattern")
                    df_target = pd.read_sql(sql_select, conn, params={"date_pattern": f"{target_date}%"})
                else:
                    sql_select = text(f"SELECT * FROM {table_name}")
                    df_target = pd.read_sql(sql_select, conn)

                if df_target.empty:
                    continue

                columns = df_target.columns.tolist()
                list_errors = []

                # Hàm rà soát danh mục nội bộ bằng Pandas
                def check_column(col_name, master_set_lower):
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
                            list_errors.append((table_name, col_name, str(row[col_name]), row["magd"], str(row["thoidiem"])))

                # Tiến hành rà soát 10 tiêu chí danh mục cốt lõi
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

                # 4. Đẩy khối lượng lỗi phát hiện ngược lại vào Database bảng log_kiem_tra_du_lieu
                if list_errors:
                    sql_insert = text(
                        "INSERT INTO log_kiem_tra_du_lieu (ten_bang, ten_cot, gia_tri_loi, ma_giao_dich, thoidiem) "
                        "VALUES (:ten_bang, :ten_cot, :gia_tri_loi, :ma_giao_dich, :thoidiem)"
                    )
                    conn.execute(
                        sql_insert,
                        [{"ten_bang": x[0], "ten_cot": x[1], "gia_tri_loi": x[2], "ma_giao_dich": x[3], "thoidiem": x[4]} for x in list_errors]
                    )
                    conn.commit()

            # =================================================================
            # 5. ĐỒNG BỘ CẬP NHẬT TRẠNG THÁI CHO BẢNG danh_sach_bang_da_kiem_tra
            # =================================================================
            if all_tables:
                # Nếu bảng đã tồn tại thì UPDATE lại thời gian kiểm tra mới nhất, nếu chưa thì INSERT mới
                sql_upsert_status = text("""
                    INSERT INTO danh_sach_bang_da_kiem_tra (ten_bang, ngay_kiem_tra_cuoi) 
                    VALUES (:t, NOW())
                    ON DUPLICATE KEY UPDATE ngay_kiem_tra_cuoi = NOW()
                """)
                conn.execute(sql_upsert_status, [{"t": table_name} for table_name in all_tables])
                conn.commit()

        return True, None
    except Exception as e:
        print(f"Error executing process_run_check_danh_muc: {str(e)}")
        return False, str(e)

def get_dashboard_summary_danh_muc(db, thang_nam=None, ngay_baocao=None):
    """
    Truy vấn gom nhóm (GROUP BY) lấy tổng số lượng lỗi phát hiện phân bổ theo từng tên bảng gốc.
    Được tối ưu chạy trực tiếp bằng Cursor thô cho tốc độ nhanh nhất phục vụ Dashboard.
    """
    where_clauses = []
    query_params = []
    
    if thang_nam:
        where_clauses.append("ten_bang LIKE %s")
        query_params.append(f"%_{thang_nam}")
        
    if ngay_baocao:
        where_clauses.append("thoidiem LIKE %s")
        query_params.append(f"{ngay_baocao}%")
        
    where_stmt = "WHERE " + " AND ".join(where_clauses) if where_clauses else ""
    
    query = f"""
        SELECT ten_bang, COUNT(*) AS so_luong 
        FROM `log_kiem_tra_du_lieu` 
        {where_stmt}
        GROUP BY ten_bang 
        ORDER BY so_luong DESC
    """
    
    summary_data = []
    try:
        engine = get_bc48_engine(db)
        with engine.connect() as connection:
            conn = connection.connection
            try:
                cursor = conn.cursor(dictionary=True)
            except (TypeError, ValueError):
                try:
                    cursor = conn.cursor(pymysql.cursors.DictCursor)
                except Exception:
                    cursor = conn.cursor()
                    
            try:
                cursor.execute(query, tuple(query_params) if query_params else None)
                rows = cursor.fetchall()
                
                for row in rows:
                    if isinstance(row, dict):
                        summary_data.append(row)
                    else:
                        summary_data.append({
                            'ten_bang': row[0],
                            'so_luong': row[1]
                        })
            finally:
                if cursor:
                    cursor.close()
    except Exception as e:
        print(f"Lỗi đồng bộ dữ liệu dashboard danh mục: {str(e)}")
        # Trả về mảng rỗng an toàn để giao diện không bị sập (crash)
        return []
        
    return summary_data


def get_all_log_errors_for_excel(db, thang_nam=None, ngay_baocao=None):
    """
    Truy vấn lấy TOÀN BỘ danh sách dòng lỗi (không phân trang LIMIT/OFFSET) 
    đáp ứng điều kiện lọc để phục vụ công tác kết xuất file Excel của quản trị viên.
    """
    where_clauses = []
    query_params = []
    
    if thang_nam:
        where_clauses.append("ten_bang LIKE %s")
        query_params.append(f"%_{thang_nam}")
        
    if ngay_baocao:
        where_clauses.append("thoidiem LIKE %s")
        query_params.append(f"{ngay_baocao}%")
        
    where_stmt = "WHERE " + " AND ".join(where_clauses) if where_clauses else ""
    
    query = f"""
        SELECT id, ten_bang, ten_cot, gia_tri_loi, ma_giao_dich, thoidiem, thoi_diem_kiem_tra
        FROM `log_kiem_tra_du_lieu`
        {where_stmt}
        ORDER BY ten_bang ASC, thoidiem DESC
    """
    
    all_data = []
    try:
        engine = get_bc48_engine(db)
        with engine.connect() as connection:
            conn = connection.connection
            try:
                cursor = conn.cursor(dictionary=True)
            except (TypeError, ValueError):
                try:
                    cursor = conn.cursor(pymysql.cursors.DictCursor)
                except Exception:
                    cursor = conn.cursor()
                    
            try:
                cursor.execute(query, tuple(query_params) if query_params else None)
                rows = cursor.fetchall()
                
                for row in rows:
                    if isinstance(row, dict):
                        all_data.append(row)
                    else:
                        all_data.append({
                            'id': row[0],
                            'ten_bang': row[1],
                            'ten_cot': row[2],
                            'gia_tri_loi': row[3],
                            'ma_giao_dich': row[4],
                            'thoidiem': row[5],
                            'thoi_diem_kiem_tra': row[6]
                        })
            finally:
                if cursor:
                    cursor.close()
    except Exception as e:
        print(f"Lỗi lấy dữ liệu kết xuất Excel danh mục: {str(e)}")
        raise e
        
    return all_data

################################################################################################################
# Tao Procedure thong ke hinh thuc GLD, GBS, GLA tu tat ca bang _yyyymm
# Thống kê theo hình thức file gửi (GLD, GBS, GLA) từ tất cả các bảng ctr_yyyymm; dwt_yyyymm; eft_yyyymm; ptr_yyyymm
# Trong khoảng thời gian có bao nhiêu dòng của loại báo cáo nào được GLA?
################################################################################################################
def get_dashboard_hinh_thuc_data(db):
    try:
        engine = get_bc48_engine(db)
        
        with engine.connect() as connection:
            # 1. Truy vấn tạo Ma Trận Tổng Hợp theo Ngày, Tháng, Loại
            query_matrix = connection.execute(text('''
                SELECT 
                    ngay_bao_cao AS NgayBaoCao, thang_bao_cao AS Thang, loai_bao_cao AS LoaiBaoCao,
                    SUM(CASE WHEN hinh_thuc_file = 'GLD' THEN so_luong ELSE 0 END) AS LanDau,
                    SUM(CASE WHEN hinh_thuc_file = 'GLA' THEN so_luong ELSE 0 END) AS GuiLai,
                    SUM(CASE WHEN hinh_thuc_file = 'GBS' THEN so_luong ELSE 0 END) AS BoSung,
                    SUM(CASE WHEN hinh_thuc_file NOT IN ('GLD', 'GLA', 'GBS') THEN so_luong ELSE 0 END) AS ChuaPhanLoai,
                    SUM(so_luong) AS TongGiaoDich,
                    MAX(thoi_diem_du_lieu) AS CapNhatCuoi
                FROM dashboard_thong_ke_hinh_thuc
                GROUP BY ngay_bao_cao, thang_bao_cao, loai_bao_cao
                ORDER BY ngay_bao_cao DESC, loai_bao_cao ASC
            ''')).fetchall()

            matrix_data = [
                {
                    'NgayBaoCao': row[0], 'Thang': row[1], 'Loại Báo Cáo': row[2],
                    'Lan dau (GLD)': int(row[3]), 'Gui lai (GLA)': int(row[4]), 'Bo sung (GBS)': int(row[5]), 'Chua phan loai': int(row[6]),
                    'Tong Giao Dich': int(row[7]), 'CapNhatCuoi': row[8] if row[8] else "N/A"
                }
                for row in query_matrix
            ]

            # 2. Truy vấn dữ liệu Tỷ lệ đóng góp theo Ngày
            query_ratio = connection.execute(text('''
                SELECT 
                    t.ngay_bao_cao AS NgayBaoCao, t.thang_bao_cao AS Thang, t.loai_bao_cao AS LoaiBaoCao, t.hinh_thuc_file AS HinhThucGui,
                    SUM(t.so_luong) AS TongSoGiaoDich, ROUND((SUM(t.so_luong) / m.TongLoaiTheoNgay) * 100, 2) AS TyLe
                FROM dashboard_thong_ke_hinh_thuc t
                INNER JOIN (
                    SELECT ngay_bao_cao, loai_bao_cao, SUM(so_luong) AS TongLoaiTheoNgay
                    FROM dashboard_thong_ke_hinh_thuc 
                    GROUP BY ngay_bao_cao, loai_bao_cao
                ) m ON t.ngay_bao_cao = m.ngay_bao_cao AND t.loai_bao_cao = m.loai_bao_cao
                GROUP BY t.ngay_bao_cao, t.thang_bao_cao, t.loai_bao_cao, t.hinh_thuc_file, m.TongLoaiTheoNgay
                ORDER BY t.ngay_bao_cao DESC, t.loai_bao_cao ASC, TongSoGiaoDich DESC
            ''')).fetchall()
            
            ratio_data = [
                {
                    'NgayBaoCao': row[0], 'Thang': row[1], 'Loại Báo Cáo': row[2],
                    'Hinh Thuc Gui': row[3], 'Tong So Giao Dich': int(row[4]),
                    'Ty Le (%)': float(row[5]) if row[5] is not None else 0.0
                }
                for row in query_ratio
            ]

            # 3. LẤY DANH SÁCH HÌNH THỨC FILE ĐỂ LÀM BỘ LỌC (Distinct)
            query_hinh_thuc = connection.execute(text('''
                SELECT DISTINCT hinh_thuc_file FROM dashboard_thong_ke_hinh_thuc 
                WHERE hinh_thuc_file IS NOT NULL AND hinh_thuc_file != ''
                ORDER BY hinh_thuc_file ASC
            ''')).fetchall()
            hinh_thuc_list = [row[0] for row in query_hinh_thuc]

            # 4. Truy vấn Chi tiết sao kê
            query_details = connection.execute(text('''
                SELECT ngay_bao_cao, thang_bao_cao, loai_bao_cao, hinh_thuc_file, so_luong, thoi_diem_du_lieu 
                FROM dashboard_thong_ke_hinh_thuc 
                ORDER BY ngay_bao_cao DESC
            ''')).fetchall()
            
            details_data = [{
                'Ngay': row[0], 'Thang': row[1], 'Loai': row[2], 
                'HinhThuc': row[3], 'SoLuong': int(row[4]) if row[4] is not None else 0,
                'ThoiDiem': row[5]
            } for row in query_details]
        
        return matrix_data, ratio_data, hinh_thuc_list, details_data

    except Exception as e:
        print(f"Lỗi hệ thống tại bc48.py: {str(e)}")
        traceback.print_exc()
        return [], [], [], []

def get_dashboard_hinh_thuc_details(db, start, length, search_value):
    try:
        engine = get_bc48_engine(db)
        with engine.connect() as connection:
            # Tạo điều kiện lọc nếu có search
            where_clause = ""
            if search_value:
                where_clause = f"WHERE ngay_bao_cao LIKE '%{search_value}%' OR hinh_thuc_file LIKE '%{search_value}%'"

            # Đếm tổng số dòng (phục vụ phân trang)
            count_query = text(f"SELECT COUNT(*) FROM dashboard_thong_ke_hinh_thuc {where_clause}")
            total_rows = connection.execute(count_query).scalar()

            # Lấy dữ liệu phân trang
            data_query = text(f'''
                SELECT ngay_bao_cao, thang_bao_cao, loai_bao_cao, hinh_thuc_file, so_luong, thoi_diem_du_lieu
                FROM dashboard_thong_ke_hinh_thuc
                {where_clause}
                ORDER BY ngay_bao_cao DESC
                LIMIT {length} OFFSET {start}
            ''')
            result = connection.execute(data_query).fetchall()
            
            data = [{
                'Ngay': row[0], 'Thang': row[1], 'Loai': row[2],
                'HinhThuc': row[3], 'SoLuong': int(row[4]), 'ThoiDiem': str(row[5])
            } for row in result]
            
            return data, total_rows
    except Exception as e:
        print(f"Lỗi: {e}")
        return [], 0

def get_sao_ke_gla(db, search_params, start, length):
    try:
        engine = get_bc48_engine(db)
        
        # 1. Chuẩn hóa tham số (định dạng YYYY-MM-DD HH:MM:SS)
        tu_ngay = search_params.get("tu_ngay") #or datetime.now().strftime('%Y-%m-%d')
        den_ngay = search_params.get("den_ngay") #or datetime.now().strftime('%Y-%m-%d')
        hinh_thuc = search_params.get("hinh_thuc") or "GLA" # Mặc định là GLA theo tên bảng
        loai_bc = search_params.get("loai_bc")

        # 2. Xây dựng điều kiện lọc
        where_clauses = []
        params = {
            "length": length,
            "start": start
        }

        # Chỉ thêm điều kiện thời gian nếu cả hai giá trị đều được chọn
        if tu_ngay and den_ngay:
            where_clauses.append("thoidiem >= :tu AND thoidiem <= :den")
            params["tu"] = f"{tu_ngay} 00:00:00"
            params["den"] = f"{den_ngay} 23:59:59"
        
        # Thêm các điều kiện khác nếu có
        if hinh_thuc:
            where_clauses.append("hinh_thuc_file = :ht")
            params["ht"] = hinh_thuc
        if loai_bc:
            where_clauses.append("loai_baocao = :loai")
            params["loai"] = loai_bc

        # Nếu không có điều kiện nào, dùng 1=1 để lấy toàn bộ
        where_sql = " AND ".join(where_clauses) if where_clauses else "1=1"

        with engine.connect() as conn:
            # 3. Đếm tổng số bản ghi
            count_sql = text(f"SELECT COUNT(*) FROM sao_ke_GLA_4loaibc WHERE {where_sql}")
            total = conn.execute(count_sql, params).scalar()

            if total == 0:
                return [], 0

            # 4. Lấy dữ liệu phân trang
            query_sql = text(f"""
                SELECT * FROM sao_ke_GLA_4loaibc 
                WHERE {where_sql} 
                ORDER BY thoidiem DESC 
                LIMIT :length OFFSET :start
            """)
            
            result = conn.execute(query_sql, params)
            data = [dict(row._mapping) for row in result.fetchall()]
            
            return data, total

    except Exception as e:
        print(f"Lỗi hệ thống tại bc48.py: {str(e)}")
        return [], 0

# Tab "SLgFile, SLgDong, HinhThuc, quydoi 4 loai_bc" lấy thông tin toàn bộ cột trong bảng dashboard_tke_slgfile_sodong_hinh_thuc
def get_dashboard_tke_slgfile(db, start, length, search_value):
    try:
        engine = get_bc48_engine(db)
        with engine.connect() as connection:
            # Xây dựng điều kiện lọc tìm kiếm nâng cao (loai_bang hoặc hinh_thuc_file)
            where_clause = ""
            params = {"length": length, "start": start}
            
            if search_value:
                where_clause = "WHERE loai_bang LIKE :search OR hinh_thuc_file LIKE :search OR nam_thang LIKE :search"
                params["search"] = f"%{search_value}%"

            # 1. Đếm tổng số dòng
            count_query = text(f"SELECT COUNT(*) FROM dashboard_tke_slgfile_sodong_hinh_thuc {where_clause}")
            total_rows = connection.execute(count_query, params).scalar()

            if total_rows == 0:
                return [], 0

            # 2. Lấy dữ liệu phân trang
            data_query = text(f'''
                SELECT 
                    loai_bang, ten_bang, nam_thang, hinh_thuc_file, loai_tien, 
                    so_luong_file, tong_so_dong, tong_sotien, tong_quydoi, ngay_cap_nhat
                FROM dashboard_tke_slgfile_sodong_hinh_thuc
                {where_clause}
                ORDER BY nam_thang DESC, loai_bang ASC
                LIMIT :length OFFSET :start
            ''')
            result = connection.execute(data_query, params).fetchall()
            
            data = [{
                'loai_bang': row[0],
                'ten_bang': row[1],
                'nam_thang': row[2],
                'hinh_thuc_file': row[3],
                'loai_tien': row[4],
                'so_luong_file': int(row[5]),
                'tong_so_dong': int(row[6]),
                'tong_sotien': float(row[7]) if row[7] else 0.0,
                'tong_quydoi': float(row[8]) if row[8] else 0.0,
                'ngay_cap_nhat': str(row[9])
            } for row in result]
            
            return data, total_rows
    except Exception as e:
        print(f"Lỗi tại get_dashboard_tke_slgfile: {e}")
        traceback.print_exc()
        return [], 0


################################################################################################################
# Thống kê từng ngày số lượng dòng KTraGDLoi, KTraFileThanhCong của từng loại báo cáo (CTR; DWT; EFT; PTR); bảng bang_thong_ke_loi_daily kết quả của sp_tong_hop_sodong_bang_error
################################################################################################################
def get_thong_ke_trang_thai_data(db, start, length, search_value, tu_ngay, den_ngay):
    try:
        engine = get_bc48_engine(db)
        with engine.connect() as connection:
            ##1. Định nghĩa điều kiện lọc rõ ràng bằng cách thêm alias 'main.' vào trước ngay_baocao
            where_clauses = ["1=1"]
            params = {"start": start, "length": length}
            
            if tu_ngay:
                where_clauses.append("main.ngay_baocao >= :tu_ngay")
                params['tu_ngay'] = tu_ngay
            if den_ngay:
                where_clauses.append("main.ngay_baocao <= :den_ngay")
                params['den_ngay'] = den_ngay
            if search_value:
                where_clauses.append("main.ngay_baocao LIKE :search")
                params['search'] = f"%{search_value}%"
            
            where_sql = " AND ".join(where_clauses)

            # 2. SQL JOIN hai thế giới lại với nhau; chuẩn hóa cấu trúc cột bảng bang_thong_ke_loi_daily có cột loai_baocao; bảng logs_nap_csv có cột loai_bc: main dùng loai_baocao, log_file dùng loai_bc
            query_str = '''
                SELECT * FROM (
                    SELECT 
                        main.ngay_baocao,
                        -- Dữ liệu thực tế quét từ ruột các bảng _error trong DB
                        SUM(CASE WHEN main.trang_thai = 'KIỂM TRA GIAO DỊCH LỖI' AND main.loai_baocao = 'CTR' THEN main.tong_so_dong ELSE 0 END) AS loi_ctr,
                        SUM(CASE WHEN main.trang_thai = 'KIỂM TRA GIAO DỊCH LỖI' AND main.loai_baocao = 'DWT' THEN main.tong_so_dong ELSE 0 END) AS loi_dwt,
                        SUM(CASE WHEN main.trang_thai = 'KIỂM TRA GIAO DỊCH LỖI' AND main.loai_baocao = 'EFT' THEN main.tong_so_dong ELSE 0 END) AS loi_eft,
                        SUM(CASE WHEN main.trang_thai = 'KIỂM TRA GIAO DỊCH LỖI' AND main.loai_baocao = 'PTR' THEN main.tong_so_dong ELSE 0 END) AS loi_ptr,
                        
                        SUM(CASE WHEN main.trang_thai = 'KIỂM TRA FILE THÀNH CÔNG' AND main.loai_baocao = 'CTR' THEN main.tong_so_dong ELSE 0 END) AS tc_ctr,
                        SUM(CASE WHEN main.trang_thai = 'KIỂM TRA FILE THÀNH CÔNG' AND main.loai_baocao = 'DWT' THEN main.tong_so_dong ELSE 0 END) AS tc_dwt,
                        SUM(CASE WHEN main.trang_thai = 'KIỂM TRA FILE THÀNH CÔNG' AND main.loai_baocao = 'EFT' THEN main.tong_so_dong ELSE 0 END) AS tc_eft,
                        SUM(CASE WHEN main.trang_thai = 'KIỂM TRA FILE THÀNH CÔNG' AND main.loai_baocao = 'PTR' THEN main.tong_so_dong ELSE 0 END) AS tc_ptr,
                        
                        -- Dữ liệu đếm từ File khi Python nạp (Lấy từ Subquery log_file bên dưới)
                        MAX(IFNULL(log_file.file_ctr, 0)) AS file_ctr,
                        MAX(IFNULL(log_file.file_dwt, 0)) AS file_dwt,
                        MAX(IFNULL(log_file.file_eft, 0)) AS file_eft,
                        MAX(IFNULL(log_file.file_ptr, 0)) AS file_ptr,
                        
                        -- Tổng số dòng hợp nhất 2 nguồn dữ liệu
                        SUM(main.tong_so_dong) AS tong_db,
                        MAX(IFNULL(log_file.tong_file, 0)) AS tong_file
                        
                    FROM bang_thong_ke_loi_daily main
                    LEFT JOIN (
                        -- Subquery độc lập gom số liệu từ bảng logs_nap_csv theo ngày
                        SELECT 
                            ngay_baocao,
                            SUM(CASE WHEN loai_bc = 'CTR' THEN so_dong_du_lieu_csv ELSE 0 END) AS file_ctr,
                            SUM(CASE WHEN loai_bc = 'DWT' THEN so_dong_du_lieu_csv ELSE 0 END) AS file_dwt,
                            SUM(CASE WHEN loai_bc = 'EFT' THEN so_dong_du_lieu_csv ELSE 0 END) AS file_eft,
                            SUM(CASE WHEN loai_bc = 'PTR' THEN so_dong_du_lieu_csv ELSE 0 END) AS file_ptr,
                            SUM(so_dong_du_lieu_csv) AS tong_file
                        FROM logs_nap_csv
                        GROUP BY ngay_baocao
                    ) log_file ON main.ngay_baocao = log_file.ngay_baocao
                    WHERE {where_cond}
                    GROUP BY main.ngay_baocao
                ) AS sub
                ORDER BY ngay_baocao DESC
                LIMIT :length OFFSET :start
            '''.format(where_cond=where_sql)

            # 3. Thực thi lấy dữ liệu
            result = connection.execute(text(query_str), params).fetchall()
            
            # 4. Đếm tổng số lượng ngày báo cáo phục vụ phân trang
            count_query = text(f"SELECT COUNT(DISTINCT main.ngay_baocao) FROM bang_thong_ke_loi_daily main WHERE {where_sql}")
            total = connection.execute(count_query, params).scalar()
            
            data = []
            for row in result:
                d = dict(row._mapping)
                for key in d:
                    if key != 'ngay_baocao' and d[key] is None:
                        d[key] = 0
                data.append(d)
                
            return data, total
            
    except Exception as e:
        print(f"Lỗi tại get_thong_ke_trang_thai_data: {str(e)}")
        traceback.print_exc()
        return [], 0

def export_thong_ke_trang_thai_to_excel(db, tu_ngay, den_ngay):
    data, _ = get_thong_ke_trang_thai_data_for_export(db, tu_ngay, den_ngay)
    df = pd.DataFrame(data)
    
    # Đổi tên cột cho đẹp (để làm Header)
    df.columns = ['Ngày', 'Loi_CTR', 'Loi_DWT', 'Loi_EFT', 'Loi_PTR', 'CTR', 'DWT', 'EFT', 'PTR', 'Tổng']
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='SoDong_LOI_tung_ngay_4loaibc', startrow=1) # Bắt đầu từ dòng 1 để chừa chỗ cho Header nhóm
        
        workbook = writer.book
        worksheet = writer.sheets['SoDong_LOI_tung_ngay_4loaibc']
        
        # Định dạng Header nhóm
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#343a40', 'font_color': 'white'})
        
        # Ghi các tiêu đề nhóm (Header tầng 1)
        worksheet.merge_range('B1:E1', 'KIỂM TRA GIAO DỊCH LỖI', header_format)
        worksheet.merge_range('F1:I1', 'KIỂM TRA FILE THÀNH CÔNG', header_format)
        
        # Định dạng các ô dữ liệu
        cell_format = workbook.add_format({'align': 'center', 'border': 1})
        for i, col in enumerate(df.columns):
            worksheet.set_column(i, i, 12, cell_format)

    output.seek(0)
    return output
# File Excel xuất ra dữ liệu dạng bảng trong db chỉ cột hàng
def export_thong_ke_trang_thai_to_excel_v01(db, tu_ngay, den_ngay):
    # 1. Gọi hàm lấy dữ liệu thô (tương tự như hàm bạn đã có)
    data, _ = get_thong_ke_trang_thai_data_for_export(db, tu_ngay, den_ngay)
    
    # 2. Xử lý thành DataFrame
    df = pd.DataFrame(data)
    
    # 3. Tạo file Excel và trả về object bytes (để app.py gửi đi)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='ThongKe')
    
    output.seek(0)
    return output

def get_thong_ke_trang_thai_data_for_export(db, tu_ngay, den_ngay):
    # Đảm bảo giá trị là None nếu là chuỗi rỗng
    tu_ngay = tu_ngay if tu_ngay and tu_ngay.strip() != "" else None
    den_ngay = den_ngay if den_ngay and den_ngay.strip() != "" else None
    
    engine = get_bc48_engine(db)
    with engine.connect() as connection:
        query_str = '''
            SELECT ngay_baocao,
                   SUM(CASE WHEN trang_thai = 'KIỂM TRA GIAO DỊCH LỖI' AND loai_baocao = 'CTR' THEN tong_so_dong ELSE 0 END) AS loi_ctr,
                   SUM(CASE WHEN trang_thai = 'KIỂM TRA GIAO DỊCH LỖI' AND loai_baocao = 'DWT' THEN tong_so_dong ELSE 0 END) AS loi_dwt,
                   SUM(CASE WHEN trang_thai = 'KIỂM TRA GIAO DỊCH LỖI' AND loai_baocao = 'EFT' THEN tong_so_dong ELSE 0 END) AS loi_eft,
                   SUM(CASE WHEN trang_thai = 'KIỂM TRA GIAO DỊCH LỖI' AND loai_baocao = 'PTR' THEN tong_so_dong ELSE 0 END) AS loi_ptr,
                   SUM(CASE WHEN trang_thai = 'KIỂM TRA FILE THÀNH CÔNG' AND loai_baocao = 'CTR' THEN tong_so_dong ELSE 0 END) AS tc_ctr,
                   SUM(CASE WHEN trang_thai = 'KIỂM TRA FILE THÀNH CÔNG' AND loai_baocao = 'DWT' THEN tong_so_dong ELSE 0 END) AS tc_dwt,
                   SUM(CASE WHEN trang_thai = 'KIỂM TRA FILE THÀNH CÔNG' AND loai_baocao = 'EFT' THEN tong_so_dong ELSE 0 END) AS tc_eft,
                   SUM(CASE WHEN trang_thai = 'KIỂM TRA FILE THÀNH CÔNG' AND loai_baocao = 'PTR' THEN tong_so_dong ELSE 0 END) AS tc_ptr,
                   SUM(tong_so_dong) AS tong_tat_ca
            FROM bang_thong_ke_loi_daily
            WHERE 1=1
        '''
        params = {}
        if tu_ngay and len(tu_ngay) == 8: # Đảm bảo là định dạng YYYYMMDD
            query_str += " AND ngay_baocao >= :tu_ngay"
            params['tu_ngay'] = tu_ngay
        if den_ngay and len(den_ngay) == 8:
            query_str += " AND ngay_baocao <= :den_ngay"
            params['den_ngay'] = den_ngay
            
        query_str += " GROUP BY ngay_baocao ORDER BY ngay_baocao DESC"
        
        result = connection.execute(text(query_str), params).fetchall()
        
        # Format dữ liệu: Chuyển đổi thành dict và ép kiểu số nguyên
        data = []
        for row in result:
            d = dict(row._mapping)
            # Ép kiểu tất cả các cột trừ 'ngay_baocao' thành int
            for key in d:
                if key != 'ngay_baocao':
                    d[key] = int(d[key] or 0)
            data.append(d)
            
        return data, None


################################################################################################################
# 
################################################################################################################
def du_lieu_bc48_thoidiem(db, loai_bc, tu_ngay, den_ngay):
    """Truy vấn dữ liệu từ các bảng tháng (ví dụ: ctr_202604)"""
    engine = get_bc48_engine(db)
    t_date = tu_ngay.replace('-', '')
    d_date = den_ngay.replace('-', '')
    
    start_month = int(t_date[:6])
    end_month = int(d_date[:6])
    
    all_data = []
    with engine.connect() as conn:
        # Lấy danh sách tất cả các bảng
        result = conn.execute(text("SHOW TABLES"))
        all_tables = [row[0] for row in result]
        
        target_tables = []
        for t in all_tables:
            if t.startswith(f"{loai_bc.lower()}_") and not t.endswith("_error"):
                try:
                    table_parts = t.split('_')
                    if len(table_parts) > 1:
                        table_month = int(table_parts[1])
                        if start_month <= table_month <= end_month:
                            target_tables.append(t)
                except: continue
        
        for table in target_tables:
            query = text(f"SELECT * FROM `{table}` WHERE `thoidiem` BETWEEN :t AND :d ORDER BY `thoidiem` DESC")
            res = conn.execute(query, {"t": t_date, "d": d_date})
            all_data.extend([dict(row._mapping) for row in res])
            
    return all_data

def thuc_thi_tong_hop_bc48(db, start_month, end_month):
    """
    Gọi Stored Procedure sp_TongHopToanBoBaoCao trên database BC48
    """
    engine = get_bc48_engine(db)
    try:
        # Sử dụng connection trực tiếp từ engine của db_bc48
        with engine.begin() as conn:
            sql = text("CALL sp_TongHopToanBoBaoCao(:start, :end)")
            result = conn.execute(sql, {"start": start_month, "end": end_month})
            
            # Lấy thông báo từ Procedure (nếu có SELECT ở cuối Procedure)
            message = "Tổng hợp thành công"
            row = result.fetchone()
            if row:
                message = row[0]
                
            return {"success": True, "msg": message}
    except Exception as e:
        print(f"Lỗi Procedure BC48: {str(e)}")
        return {"success": False, "msg": f"Lỗi thực thi: {str(e)}"}

def lay_lich_su_tong_hop(db):
    """Lấy dữ liệu từ bảng bao_cao_tong_hop_history từ database BC48"""
    engine = get_bc48_engine(db)
    try:
        with engine.connect() as conn:
            sql = text("SELECT * FROM bao_cao_tong_hop_history ORDER BY ngay_cap_nhat DESC")
            result = conn.execute(sql)

            # Chuyển đổi kết quả sang list các dict và ép kiểu số tiền
            data = []
            for row in result:
                row_dict = dict(row._mapping)
                # Ép kiểu Decimal cho các cột số tiền để chắc chắn không bị làm tròn
                row_dict['TongSoTien'] = Decimal(row_dict['TongSoTien']) if row_dict['TongSoTien'] is not None else Decimal('0')
                row_dict['TongQuyDoi'] = Decimal(row_dict['TongQuyDoi']) if row_dict['TongQuyDoi'] is not None else Decimal('0')
                row_dict['SoDong'] = int(row_dict['SoDong']) if row_dict['SoDong'] is not None else 0
                data.append(row_dict)
            return data
    except Exception as e:
        print(f"Lỗi lấy lịch sử BC48: {str(e)}")
        return []

def count_van_tin_bc48(db, loai_bc, tu_ngay, den_ngay):
    """Đếm tổng số bản ghi khớp với điều kiện thời gian trên tất cả các bảng tháng."""
    engine = get_bc48_engine(db)
    t_date = tu_ngay.replace('-', '') 
    d_date = den_ngay.replace('-', '')
    
    start_month = int(t_date[:6])
    end_month = int(d_date[:6])
    
    total_count = 0
    with engine.connect() as conn:
        result = conn.execute(text("SHOW TABLES"))
        all_tables = [row[0] for row in result]
        
        for t in all_tables:
            if t.startswith(f"{loai_bc.lower()}_") and not t.endswith("_error"):
                try:
                    table_month = int(t.split('_')[1])
                    if start_month <= table_month <= end_month:
                        # Chỉ SELECT COUNT(*), cực kỳ nhanh
                        count_query = text(f"SELECT COUNT(*) FROM `{t}` WHERE `thoidiem` BETWEEN :t AND :d")
                        table_count = conn.execute(count_query, {"t": t_date, "d": d_date}).scalar()
                        total_count += (table_count or 0)
                except: continue
                
    return total_count

def van_tin_bc48(db, loai_bc, tu_ngay, den_ngay, page=1, per_page=50):
    """Truy vấn dữ liệu từ các bảng tháng dựa trên loại và thời gian"""
    # Vấn tin từ ngày đến ngày của bảng CTR/DWT/EFT/PTR và Xuất Excel CSV (Full)
    engine = get_bc48_engine(db)
    t_date = tu_ngay.replace('-', '') 
    d_date = den_ngay.replace('-', '')
    
    start_month = int(t_date[:6])
    end_month = int(d_date[:6])
    
    offset = (page - 1) * per_page
    all_data = []
    
    with engine.connect() as conn:
        # 1. Lấy danh sách bảng và sắp xếp (mới nhất trước)
        result = conn.execute(text("SHOW TABLES"))
        all_tables = [row[0] for row in result]
        
        target_tables = []
        for t in all_tables:
            if t.startswith(f"{loai_bc.lower()}_") and not t.endswith("_error"):
                try:
                    table_month = int(t.split('_')[1])
                    if start_month <= table_month <= end_month:
                        target_tables.append(t)
                except: continue
        
        # Luôn lấy bảng mới nhất trước để đảm bảo tính thời gian
        target_tables.sort(reverse=True)

        # 2. Logic phân trang qua các bảng
        current_skip = offset
        
        for table in target_tables:
            # Nếu đã lấy đủ dữ liệu cho trang này thì dừng
            if len(all_data) >= per_page:
                break
                
            # Đếm số dòng trong bảng này để xem có cần nhảy qua (skip) không
            count_query = text(f"SELECT COUNT(*) FROM `{table}` WHERE `thoidiem` BETWEEN :t AND :d")
            row_count = conn.execute(count_query, {"t": t_date, "d": d_date}).scalar()
            
            if current_skip >= row_count:
                # Nếu trang này nằm hoàn toàn trước offset, bỏ qua bảng này
                current_skip -= row_count
                continue
            else:
                # Lấy dữ liệu từ bảng này với offset đã điều chỉnh
                limit_needed = per_page - len(all_data)
                # Thay vì chỉ dùng BETWEEN, hãy thêm điều kiện bắt dữ liệu: rỗng (NULL hoặc ''); không đúng chuẩn (ví dụ: chuỗi rác, định dạng khác)
                # dữ liệu trong cột thoidiem không tuân thủ định dạng chuẩn (YYYYMMDD hoặc YYYYMMDD HH:MM:SS)
                data_query = text(f"""
                    SELECT * FROM `{table}` 
                    WHERE (thoidiem BETWEEN :t AND :d) 
                       OR thoidiem IS NULL 
                       OR thoidiem = ''
                    ORDER BY `thoidiem` DESC 
                    LIMIT :limit OFFSET :offset
                """)
                res = conn.execute(data_query, {
                    "t": t_date, 
                    "d": d_date, 
                    "limit": limit_needed, 
                    "offset": current_skip
                })
                all_data.extend([dict(row._mapping) for row in res])
                
                # Sau khi đã nhảy qua phần offset, các bảng sau không cần skip nữa
                current_skip = 0
            
    return all_data



def generate_csv_stream_bc48(db, loai_bc, tu_ngay, den_ngay):
    """Thay thế hoàn toàn cho hàm xuat_excel cũ"""
    engine = get_bc48_engine(db)
    t_date = tu_ngay.replace('-', '')
    d_date = den_ngay.replace('-', '')
    
    start_month = int(t_date[:6])
    end_month = int(d_date[:6])

    with engine.connect() as conn:
        result = conn.execute(text("SHOW TABLES"))
        target_tables = [t[0] for t in result if t[0].startswith(f"{loai_bc.lower()}_") 
                         and not t[0].endswith("_error") 
                         and start_month <= int(t[0].split('_')[1]) <= end_month]

        output = io.StringIO()
        writer = csv.writer(output, quoting=csv.QUOTE_MINIMAL)

        # Ghi BOM để Excel không lỗi font tiếng Việt
        yield "\ufeff" 
        
        # Header
        writer.writerow(["Thoi diem", "Ma GD", "Loai tien", "So tien", "Quy doi", "Ten KH", "Ly do"])
        yield output.getvalue()
        output.seek(0)
        output.truncate(0)

        for table in target_tables:
            # Dùng stream_results để đọc từng dòng từ DB
            query = text(f"SELECT thoidiem, magd, loaitien, sotien, quydoi, tenkh, lydomucdich FROM `{table}` WHERE `thoidiem` BETWEEN :t AND :d")
            proxy = conn.execution_options(stream_results=True).execute(query, {"t": t_date, "d": d_date})
            
            for row in proxy:
                writer.writerow(list(row))
                yield output.getvalue()
                output.seek(0)
                output.truncate(0)

# Nút lệnh "Thực hiện Tổng hợp" trên site /bc48_pcrt, khi nhấn nút sẽ chạy sp_TongHopBaoCaoTheoDonVi trên db
def thuc_thi_tong_hop_don_vi_bc48(db, start_month, end_month):
    """
    Gọi Stored Procedure sp_TongHopBaoCaoTheoDonVi trên database BC48
    để tổng hợp dữ liệu phân rã theo tất cả các đơn vị.
    """
    engine = get_bc48_engine(db)
    try:
        # Chuyển đổi sang kiểu int nếu tham số truyền vào là string
        start_m = int(str(start_month).replace('-', ''))
        end_m = int(str(end_month).replace('-', ''))
        
        with engine.begin() as conn:
            # Procedure thực tế chỉ nhận 2 tham số đầu vào
            sql = text("CALL sp_TongHopBaoCaoTheoDonVi(:start, :end)")
            result = conn.execute(sql, {
                "start": start_m, 
                "end": end_m
            })
            
            # Mặc định thông báo nếu không lấy được result từ Procedure
            message = "Tổng hợp theo đơn vị hoàn tất"
            
            # Lấy thông báo "Tổng hợp thành công X bảng..." từ câu lệnh SELECT cuối cùng trong Procedure
            row = result.fetchone()
            if row:
                message = row[0]
                
            return {"success": True, "msg": message}
    except Exception as e:
        print(f"Lỗi Procedure Tổng hợp đơn vị BC48: {str(e)}")
        return {"success": False, "msg": f"Lỗi thực thi tổng hợp đơn vị: {str(e)}"}

def export_history_bc48_logic(db, mode):
    # CHÚ Ý: Bắt đầu bằng khối try
    try:
        # Lấy engine riêng của BC48
        engine = get_bc48_engine(db)
        
        if mode == 'toan_hang':
            sql = "SELECT LoaiBaoCao, TenBang, loaitien, SoDong, TongSoTien, TongQuyDoi, ngay_cap_nhat FROM bao_cao_tong_hop_history ORDER BY ngay_cap_nhat DESC"
            filename = f"bc48_TH_Toan_Hang_{datetime.now().strftime('%Y%m%d')}.xlsx"
        else:
            sql = "SELECT * FROM bao_cao_tong_hop_don_vi_history ORDER BY ngay_cap_nhat DESC"
            filename = f"bc48_TH_Don_Vi_{datetime.now().strftime('%Y%m%d')}.xlsx"

        # Thực hiện đọc dữ liệu
        with engine.connect() as conn:
            df = pd.read_sql(text(sql), conn)
        
        if df.empty:
            return None, "Không có dữ liệu"

        # Ghi ra file Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Data')
        
        output.seek(0)
        return output, filename

    except Exception as e:
        # Log lỗi để kiểm tra trong console
        print(f"Lỗi logic xuất Excel BC48: {traceback.format_exc()}")
        # Trả về None và nội dung lỗi để app.py xử lý tiếp
        return None, str(e)

def thuc_thi_tong_hop_don_vi_riengle(db, start_month, end_month, manh8so_moi):
    """
    Gọi Stored Procedure sp_TongHopBaoCaoTheoDonVi_RiengLe trên database BC48
    để tổng hợp dữ liệu cho DUY NHẤT một đơn vị cụ thể.
    """
    engine = get_bc48_engine(db)
    try:
        # 1. Chuẩn hóa tham số thời gian (YYYY-MM hoặc YYYYMM -> YYYYMM int)
        start_m = int(str(start_month).replace('-', ''))
        end_m = int(str(end_month).replace('-', ''))
        
        # 2. Đảm bảo mã ngân hàng không có khoảng trắng thừa
        target_ma = str(manh8so_moi).strip()
        
        with engine.begin() as conn:
            # 3. Procedure này nhận 3 tham số: start, end, và mã đơn vị
            sql = text("CALL sp_TongHopBaoCaoTheoDonVi_RiengLe(:start, :end, :ma)")
            result = conn.execute(sql, {
                "start": start_m, 
                "end": end_m,
                "ma": target_ma
            })
            
            # Mặc định thông báo
            message = f"Tổng hợp cho đơn vị {target_ma} hoàn tất"
            
            # 4. Lấy thông báo từ câu lệnh SELECT cuối cùng trong Procedure
            # (SELECT CONCAT('Thành công! Đã tổng hợp...', p_manh8so) AS Message)
            row = result.fetchone()
            if row:
                message = row[0]
                
            return {"success": True, "msg": message}

    except Exception as e:
        print(f"Lỗi Procedure sp_TongHopBaoCaoTheoDonVi_RiengLe Tổng hợp đơn vị riêng lẻ BC48: {str(e)}")
        return {"success": False, "msg": f"Lỗi thực thi: {str(e)}"}


################################################################################################################
################################################################################################################
def get_thong_tin_khach_hang(db, sogt='', tenkh='', data_month='', page=1, per_page=50):
    """
    Hàm tìm kiếm, phân trang và lấy dữ liệu phục vụ Stream CSV.
    Sử dụng raw SQL để đạt tốc độ tối đa.
    """
    engine = db.get_engine(bind='db_bc48')
    
    conditions = []
    params = {}
    
    # Tối ưu hóa Index
    if sogt:
        conditions.append("sogt = :sogt")
        params['sogt'] = sogt
    if tenkh:
        # Nếu dữ liệu cực lớn, khuyên dùng f"{tenkh}%" thay vì f"%{tenkh}%" để ăn index
        conditions.append("tenkh LIKE :tenkh")
        params['tenkh'] = f"%{tenkh}%"
    if data_month:
        conditions.append("data_month = :data_month")
        params['data_month'] = data_month
        
    where_clause = "WHERE " + " AND ".join(conditions) if conditions else ""
    
    # 1. Câu lệnh đếm tổng số dòng
    count_sql = text(f"SELECT COUNT(*) FROM thongtinkhachhang {where_clause}")
    
    # 2. Câu lệnh lấy dữ liệu phân đoạn
    offset = (page - 1) * per_page
    data_sql = text(f"""
        SELECT id, kieukh, sogt, tenkh, ngaysinh, quoctich, diachitt, sodt, data_source, data_month, created_at 
        FROM thongtinkhachhang 
        {where_clause}
        ORDER BY data_month DESC, id DESC
        LIMIT :limit OFFSET :offset
    """)
    
    params['limit'] = per_page
    params['offset'] = offset
    
    with engine.connect() as conn:
        # Chỉ đếm tổng số dòng ở trang đầu tiên hoặc khi cần thiết để tăng tốc
        total = conn.execute(count_sql, {k: v for k, v in params.items() if k not in ['limit', 'offset']}).scalar()
        
        # Lấy dữ liệu đoạn (Chunk) hiện tại
        rows = conn.execute(data_sql, params).mappings().all()
        data_list = [dict(row) for row in rows]
        
    total_pages = math.ceil(total / per_page) if total > 0 else 1
    
    return {
        'data': data_list,
        'total': total,
        'total_pages': total_pages
    }

def generate_ttkh_csv_stream(db, sogt='', tenkh='', data_month=''):
    """
    Hàm Generator xử lý luồng dữ liệu CSV cho hàng triệu dòng (Tầng Business Logic)
    """
    # 1. Xuất ký tự BOM (Byte Order Mark) để Microsoft Excel hiểu được UTF-8 (không lỗi font tiếng Việt)
    yield '\ufeff'
    
    # Sử dụng StringIO làm bộ đệm tạm cho thư viện csv
    output = io.StringIO()
    writer = csv.writer(output, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    
    # 2. Định nghĩa và ghi dòng tiêu đề (Header)
    headers = ["STT", "Tháng dữ liệu", "Kiểu KH", "Số giấy tờ", "Tên khách hàng", "Ngày sinh", "Quốc tịch", "Địa chỉ thường trú", "Số điện thoại", "Nguồn dữ liệu", "Ngày tạo"]
    writer.writerow(headers)
    yield output.getvalue()
    output.seek(0)
    output.truncate(0)
    
    # 3. Tiến hành bốc dữ liệu theo từng cụm (Chunk) 10.000 dòng để giải phóng RAM liên tục
    current_page = 1
    chunk_size = 10000
    stt_counter = 1
    
    while True:
        # Gọi lại hàm lấy dữ liệu phân trang đã viết trước đó
        result = get_thong_tin_khach_hang(
            db=db, sogt=sogt, tenkh=tenkh, data_month=data_month, page=current_page, per_page=chunk_size
        )
        data_chunk = result['data']
        
        # Nếu không còn dữ liệu ở cụm tiếp theo -> Ngắt vòng lặp, kết thúc luồng
        if not data_chunk:
            break
            
        # Duyệt qua từng dòng trong cụm dữ liệu vừa lấy
        for item in data_chunk:
            ngaysinh_str = item['ngaysinh'].strftime('%d/%m/%Y') if item['ngaysinh'] else ''
            created_str = item['created_at'].strftime('%d/%m/%Y %H:%M:%S') if item['created_at'] else ''
            
            # Định dạng mảng dữ liệu thô (Mẹo \t chống mất số 0 trong Excel)
            row_data = [
                stt_counter,
                f"\t{item['data_month']}", 
                item['kieukh'] or '',
                f"\t{item['sogt'] or ''}",  
                item['tenkh'].upper() if item['tenkh'] else '',
                ngaysinh_str,
                item['quoctich'] or '',
                item['diachitt'] or '',
                f"\t{item['sodt'] or ''}",  
                item['data_source'] or '',
                created_str
            ]
            
            writer.writerow(row_data)
            stt_counter += 1
            
        # Trả về chuỗi dữ liệu CSV của cụm hiện tại cho app.py đẩy về trình duyệt
        yield output.getvalue()
        
        # Giải phóng và xóa bộ nhớ đệm tạm thời StringIO để chuẩn bị nhận cụm mới
        output.seek(0)
        output.truncate(0)
        
        # Tăng trang lên cụm tiếp theo
        current_page += 1

