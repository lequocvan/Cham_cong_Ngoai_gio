from app import app, db, User, bcrypt
from sqlalchemy import text

def sync_employees_to_users():
    with app.app_context():
        print("--- ĐANG BẮT ĐẦU ĐỒNG BỘ TÀI KHOẢN NHÂN VIÊN ---")
        
        # 1. Lấy danh sách nhân viên từ bảng gốc
        try:
            result = db.session.execute(text("SELECT ma_nhan_vien, ho_ten FROM thong_tin_nguoi_lao_dong"))
            employees = result.fetchall()
        except Exception as e:
            print(f"Lỗi khi đọc bảng nhân viên: {e}")
            return

        count_added = 0
        admin_id = "200739853"

        for emp in employees:
            ma_nv = emp.ma_nhan_vien
            ten_nv = emp.ho_ten

            # 2. Kiểm tra xem user này đã tồn tại trong bảng users chưa
            existing_user = User.query.filter_by(ma_nhan_vien=ma_nv).first()
            
            if not existing_user:
                # Kiểm tra điều kiện Admin
                if ma_nv == admin_id:
                    role_name = 'ADMIN'
                    admin_status = True
                else:
                    role_name = 'LAP_BANG'
                    admin_status = False
                
                # Tạo đối tượng User mới khớp với Model trong app.py
                new_user = User(
                    ma_nhan_vien=ma_nv,
                    fullname=ten_nv,     # Lấy ho_ten làm fullname
                    role=role_name,      # Gán role tương ứng
                    is_active=True,
                    is_admin=admin_status
                )
                
                # Sử dụng hàm set_password có sẵn trong model để băm mật khẩu
                # Mật khẩu mặc định là mã nhân viên
                new_user.set_password(ma_nv)
                
                db.session.add(new_user)
                count_added += 1
                print(f"Đã tạo: {ma_nv} - {ten_nv} (Admin: {admin_status})")

        # 3. Lưu vào database
        try:
            db.session.commit()
            print(f"--- HOÀN THÀNH ---")
            print(f"Tổng số tài khoản mới đã tạo: {count_added}")
        except Exception as e:
            db.session.rollback()
            print(f"Lỗi khi lưu vào Database: {e}")

if __name__ == "__main__":
    sync_employees_to_users()
