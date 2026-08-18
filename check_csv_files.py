import pandas as pd
import os
import glob
import shutil
from datetime import datetime

#script Python tập trung vào "Dry Run" (Kiểm tra mô phỏng). Script này sẽ thực hiện kiểm tra cấu trúc, định dạng ngày tháng và tính toàn vẹn dữ liệu
#Phân loại file: Bạn có thể chia file thành 2 nhóm: valid_files (nạp ngay) và invalid_files (cần gửi lại hoặc sửa thủ công)

CSV_ERROR_COLUMNS = [
    "TRANG_THAI", "MA_NGANHANG", "TEN_NGANHANG", "NGAY_BAOCAO",
    "TEN_FILE", "DONG_TIEUDE", "LOAI_BAOCAO", "HINHTHUC_GUI",
    "SOLAN_GUI", "MA_LOI", "MA_GIAO_DICH", "MOTA_LOI",
    "DONG_LOI", "GIAODICH_LOI", "NGAY_GUI", "YEU_CAU", "GHI_CHU"
]

def check_single_file(file_path):
    issues = []
    try:
        # Đọc header để kiểm tra cấu trúc
        df = pd.read_csv(file_path, nrows=1, dtype=str)
        
        # 1. Kiểm tra thiếu cột
        missing = [c for c in CSV_ERROR_COLUMNS if c not in df.columns]
        if missing:
            issues.append(f"Thiếu cột: {', '.join(missing)}")
            
        # 2. Kiểm tra file rỗng
        full_df = pd.read_csv(file_path, dtype=str)
        if full_df.empty:
            issues.append("File trống (chỉ có header)")
        
        # 3. Kiểm tra định dạng NGAY_BAOCAO
        if 'NGAY_BAOCAO' in full_df.columns and full_df['NGAY_BAOCAO'].isnull().all():
            issues.append("Cột NGAY_BAOCAO bị trống toàn bộ")
    
    except Exception as e:
        issues.append(f"Lỗi đọc file: {str(e)}")
        
    return issues

def run_batch_check(source_folder):
    valid_dir = os.path.join(source_folder, "valid_files")
    invalid_dir = os.path.join(source_folder, "invalid_files")
    os.makedirs(valid_dir, exist_ok=True)
    os.makedirs(invalid_dir, exist_ok=True)
    
    log_path = os.path.join(source_folder, f"check_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
    all_files = glob.glob(os.path.join(source_folder, "*.csv"))
    
    # Biến để theo dõi kết quả nạp cho phần in ra màn hình
    summary = {"valid": 0, "invalid": 0}
    
    print(f"Đang xử lý {len(all_files)} file...")
    
    with open(log_path, "w", encoding="utf-8") as log_file:
        for file_path in all_files:
            filename = os.path.basename(file_path)
            issues = check_single_file(file_path)
            
            if issues:
                shutil.copy2(file_path, os.path.join(invalid_dir, filename))
                log_file.write(f"[INVALID] {filename}: {', '.join(issues)}\n")
                summary["invalid"] += 1
            else:
                shutil.copy2(file_path, os.path.join(valid_dir, filename))
                log_file.write(f"[VALID] {filename}\n")
                summary["valid"] += 1
                
    print(f"Hoàn tất! Kiểm tra chi tiết tại: {log_path}")
    print(f"Tổng kết: {summary['valid']} file hợp lệ, {summary['invalid']} file lỗi.")

if __name__ == "__main__":
    folder = "/Users/lequocvan/Downloads/nap_dl_pcrt_2025/csv_2025"
    run_batch_check(folder)
