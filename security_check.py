import subprocess
import sys
import os

def run_security_scan():
    print("\n" + "="*80)
    print(" [BẢO MẬT] Đang quét hệ thống bằng pip-audit (Chờ tối đa 10s)...")
    print("="*80)

    # Xác định đường dẫn file requirements.txt
    req_path = os.path.join(os.path.dirname(__file__), "requirements.txt")
        
    # Kiểm tra file requirements.txt trước khi quét
    if not os.path.exists(req_path):
        print(f"\n [!] Bỏ qua quét: Không tìm thấy file cấu hình tại: {req_path}")
        print("     Hệ thống sẽ tiếp tục khởi động ứng dụng Flask...")
        print("="*80 + "\n")
        return

    try:
        # Tăng timeout lên 10s để giảm thiểu lỗi mạng trên Catalina
        # Bạn có thể thêm tham số "--local" nếu chỉ muốn quét cục bộ (nếu pip-audit hỗ trợ)
        ##cmd = [sys.executable, "-m", "pip_audit", "-r", req_path]
        # Thêm tham số --index-url để bỏ qua private repo bị lỗi kết nối
        cmd = [
            sys.executable, "-m", "pip_audit", 
            "-r", req_path,
            "--index-url", "https://pypi.org/simple"
        ]
        
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=10  # Tăng từ 5s lên 10s để ổn định hơn
        )
        
        if result.returncode == 0:
            print("\n [✓] Tuyệt vời! Không phát hiện lỗ hổng bảo mật nào trong các thư viện thực tế.")
        else:
            print("\n [⚠] CẢNH BÁO BẢO MẬT: Phát hiện một số thư viện chưa khớp bản vá!")
            if result.stdout:
                print(result.stdout.strip())
            if result.stderr:
                print("\nChi tiết lỗi kỹ thuật:")
                print(result.stderr.strip())
            
    except subprocess.TimeoutExpired:
        # Xử lý khi kết nối máy chủ quá lâu do sự cố SSL/mạng trên Catalina
        print("\n [!] Bỏ qua quét bảo mật: Không thể kết nối máy chủ (Timeout).")
        print("     Có thể do chứng chỉ SSL trên Catalina đã cũ hoặc mạng bị nghẽn.")
        print("     Hệ thống sẽ tiếp tục khởi động ứng dụng Flask...")
    except FileNotFoundError:
        print("\n [!] Không tìm thấy công cụ 'pip-audit'. Vui lòng chạy: pip install pip-audit")
    except Exception as e:
        print(f"\n [!] Lỗi khi quét: {str(e)}")
        
    print("="*80 + "\n")

if __name__ == "__main__":
    run_security_scan()
