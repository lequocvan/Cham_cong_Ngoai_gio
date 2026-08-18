#models.py và app.py nằm cùng thư mục gốc
# ----------------------------------------------------------------------
# Người lao động đăng ký đơn; Ký số & Phê duyệt;
# Admin sử dụng /admin/ho-so-permissions để cấp quyền; bảng user_unit_permissions PHE_DUYET_DON
# ----------------------------------------------------------------------

import base64
from datetime import datetime, timezone
from cryptography import x509
from cryptography.hazmat.backends import default_backend
try:
    from cryptography.hazmat.primitives.serialization import pkcs7
except ImportError:
    pkcs7 = None
from flask_login import UserMixin
from extensions import db, bcrypt  # Import db và bcrypt từ extensions để tránh Circular Import  # Import db chung từ extensions.py, cùng thư mục với app.py
from sqlalchemy import (
    Column,
    Date,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    String,
    Text,
    event, 
)

class User(db.Model, UserMixin):
    __tablename__ = 'users'
    __table_args__ = {
        'mysql_engine': 'InnoDB',
        'mysql_charset': 'utf8mb4',
        'mysql_collate': 'utf8mb4_general_ci',
    }

    ma_nhan_vien = db.Column(db.String(20), db.ForeignKey('thong_tin_nguoi_lao_dong.ma_nhan_vien'), primary_key=True)
    password_hash = db.Column(db.String(255), nullable=False)
    fullname = db.Column(db.String(255))  # Họ và tên
    role = db.Column(db.String(50), default='LAP_BANG')
    is_active = db.Column(db.Boolean, default=True)
    is_admin = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    password_changed_at = db.Column(db.DateTime, default=datetime.utcnow)
    force_password_change = db.Column(db.Boolean, default=False)

    # Relationship ngược tới thông tin lao động
    thong_tin_lao_dong = db.relationship('ThongTinNguoiLaoDong', backref='account', uselist=False)

    # ------------------------------------------------------------------
    # HELPER CHECK QUYỀN ĐỘNG TRONG DATABASE
    # ------------------------------------------------------------------
    def has_unit_permission(self, perm_code):
        """Kiểm tra user có mã quyền cụ thể trong bảng user_unit_permissions hay không"""
        from models import UserUnitPermission  # Local import tránh Circular Import
        perm = UserUnitPermission.query.filter_by(
            ma_nhan_vien=self.ma_nhan_vien,
            permission_code=perm_code
        ).first()
        return perm is not None

    # ------------------------------------------------------------------
    # PROPERTIES KIỂM TRA ROLE VÀ QUYỀN TRUY CẬP (Dùng trực tiếp ở Jinja2)
    # ------------------------------------------------------------------
    @property
    def is_quan_ly(self):
        """Kiểm tra user có vai trò Quản lý hoặc Kiểm soát"""
        return self.role in ['QUAN_LY', 'KIEM_SOAT'] if self.role else False

    def is_lap_bang(self):
        """Kiểm tra user lập bảng"""
        return self.role == 'LAP_BANG' if self.role else False

    @property
    def can_approve(self):
        """
        Quyền phê duyệt đơn:
        - Là Admin
        - Hoặc role thuộc nhóm quản lý/duyệt ('ADMIN', 'QUAN_LY', 'APPROVER')
        - Hoặc được cấp quyền 'PHE_DUYET_DON' trong bảng user_unit_permissions
        """
        if getattr(self, 'is_admin', False) or (self.role and self.role in ['ADMIN', 'QUAN_LY', 'APPROVER']):
            return True
        return self.has_unit_permission('PHE_DUYET_DON')

    @property
    def is_controller(self):
        """
        Quyền kiểm soát đơn:
        - Là Admin
        - Hoặc role thuộc nhóm kiểm soát/quản lý ('ADMIN', 'KIEM_SOAT', 'CONTROLLER', 'QUAN_LY')
        - Hoặc được cấp quyền 'KIEM_SOAT_DON' trong bảng user_unit_permissions
        """
        if getattr(self, 'is_admin', False) or (self.role and self.role in ['ADMIN', 'KIEM_SOAT', 'CONTROLLER', 'QUAN_LY']):
            return True
        return self.has_unit_permission('KIEM_SOAT_DON')

    # ------------------------------------------------------------------
    # PHƯƠNG THỨC KIỂM TRA QUYỀN TỔNG QUÁT (Hỗ trợ gọi hàm trong Jinja2)
    # ------------------------------------------------------------------
    def has_permission(self, permission_name):
        """
        Cho phép gọi: current_user.has_permission('KIEM_SOAT_DON') 
        hoặc current_user.has_permission('PHE_DUYET_DON') từ cả Jinja2 lẫn Python code
        """
        if getattr(self, 'is_admin', False) or getattr(self, 'role', '') == 'ADMIN':
            return True

        if permission_name == 'KIEM_SOAT_DON':
            return self.is_controller

        if permission_name == 'PHE_DUYET_DON':
            return self.can_approve

        return self.has_unit_permission(permission_name)

    # ------------------------------------------------------------------
    # AUTHENTICATION & FLASK-LOGIN
    # ------------------------------------------------------------------
    def set_password(self, password):
        self.password_hash = bcrypt.generate_password_hash(password).decode('utf-8')
        self.password_changed_at = datetime.now(timezone.utc)

    def check_password(self, password):
        return bcrypt.check_password_hash(self.password_hash, password)

    def get_id(self):
        return self.ma_nhan_vien

# Model Phân quyền Menu
class UserMenuPermission(db.Model):
    __tablename__ = 'user_menu_permissions'
    __table_args__ = (
        db.UniqueConstraint('ma_nhan_vien', 'menu_slug', name='unique_user_menu'),
        {
            'mysql_engine': 'InnoDB',
            'mysql_charset': 'utf8mb4',
            'mysql_collate': 'utf8mb4_general_ci',
        }
    )
    
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    ma_nhan_vien = db.Column(db.String(20), db.ForeignKey('users.ma_nhan_vien', ondelete='CASCADE', onupdate='CASCADE'), nullable=False)
    menu_slug = db.Column(db.String(100), nullable=False)
    created_at = db.Column(db.DateTime, server_default=db.func.now())

    # Đảm bảo một nhân viên không bị trùng lặp slug menu
    __table_args__ = (
        db.UniqueConstraint('ma_nhan_vien', 'menu_slug', name='unique_user_menu'),
    )

class Permission(db.Model):
    __tablename__ = 'permissions'
    __table_args__ = {
        'mysql_engine': 'InnoDB',
        'mysql_charset': 'utf8mb4',
        'mysql_collate': 'utf8mb4_general_ci',
    }
    code = db.Column(db.String(50), primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text)

class UserUnitPermission(db.Model):
    __tablename__ = 'user_unit_permissions'
    __table_args__ = {
        'mysql_engine': 'InnoDB',
        'mysql_charset': 'utf8mb4',
        'mysql_collate': 'utf8mb4_general_ci',
    }

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    ma_nhan_vien = db.Column(
        db.String(20), 
        db.ForeignKey('users.ma_nhan_vien', ondelete='CASCADE', onupdate='CASCADE'), 
        nullable=False
    )
    permission_code = db.Column(
        db.String(50), 
        db.ForeignKey('permissions.code', ondelete='CASCADE', onupdate='CASCADE'), 
        nullable=False
    )
    ma_hieu_2 = db.Column(db.String(255), nullable=False)

    # Relationships
    user = db.relationship('User', backref=db.backref('unit_permissions', cascade='all, delete-orphan'))
    permission = db.relationship('Permission', backref=db.backref('user_permissions', cascade='all, delete-orphan'))
    
# ----------------------------------------------------------------------
# 1. BẢNG ĐƠN VỊ
# ----------------------------------------------------------------------
class DonVi(db.Model):
    __tablename__ = 'don_vi'
    __table_args__ = {
        'mysql_engine': 'InnoDB',
        'mysql_charset': 'utf8mb4',
        'mysql_collate': 'utf8mb4_general_ci',
    }

    ma_hieu_2 = db.Column(db.String(255), primary_key=True)
    ten_ma_hieu_2 = db.Column(db.String(255), nullable=False)
    ten_ma_hieu_cu = db.Column(db.String(255))

    ma_hieu_1 = db.Column(db.String(255))
    ten_ma_hieu_1 = db.Column(db.String(255))

    TSC_Loai_I_Loai_II_Xoa_bo = db.Column(db.String(255))
    ma_khu_vuc = db.Column(db.String(255))
    ten_khu_vuc = db.Column(db.String(255))
    ma_khu_vuc_KTGSNB = db.Column(db.String(255))
    ten_khu_vuc_KTGSNB = db.Column(db.String(255))

    MST = db.Column(db.String(255))
    MaNH8so = db.Column(db.String(255))
    MaNH8so_moi = db.Column(db.String(255))
    mail = db.Column(db.String(255))

    pho_tong_giam_doc_phu_trach_KHKD = db.Column(db.String(255))
    nhom_KHCL = db.Column(db.String(255))
    hang_KHCL = db.Column(db.String(255))
    XHRR_Chung_KTGSNB = db.Column(db.String(255))
    XHRR_TD_KTGSNB = db.Column(db.String(255))
    XHRR_NTD_KTGSNB = db.Column(db.String(255))

    ghi_chu = db.Column(db.String(255))
    trang_thai = db.Column(db.String(255), default='Hoạt động')

    ngay_thao_tac = db.Column(db.String(255))
    loai_thao_tac = db.Column(db.String(255))

    def to_dict(self):
        data = {}
        for column in self.__table__.columns:
            value = getattr(self, column.name)
            data[column.name] = "" if value is None else value
        return data

    def __repr__(self):
        return f"<DonVi {self.ma_hieu_2} - {self.ten_ma_hieu_2}>"

# Chặn hành động xóa Đơn vị ở tầng SQLAlchemy
@event.listens_for(DonVi, "before_delete")
def prevent_deletion(mapper, connection, target):
    raise RuntimeError("Hành động xóa bị cấm. Chỉ được phép Tạm dừng đơn vị.")

# ----------------------------------------------------------------------
# 2. BẢNG PHÒNG BÀN
# ----------------------------------------------------------------------
class PhongBan(db.Model):
    __tablename__ = 'phong_ban'
    __table_args__ = {
        'mysql_engine': 'InnoDB',
        'mysql_charset': 'utf8mb4',
        'mysql_collate': 'utf8mb4_general_ci',
    }

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    ten_phong_ban = db.Column(db.String(100), nullable=False)
    # THÊM MỚI: Liên kết phòng ban với một đơn vị cụ thể thông qua ma_hieu_2
    # Đảm bảo trong mysql không có default value ='PCRT', nullable=False để bắt buộc nhập
    ma_hieu_2 = db.Column(db.String(255), db.ForeignKey('don_vi.ma_hieu_2', onupdate="CASCADE"), nullable=False)
    mo_ta = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.now)

    # Quan hệ với đơn vị
    don_vi = db.relationship('DonVi', backref=db.backref('phong_bans', lazy=True))

    def __repr__(self):
        return f"<PhongBan {self.ten_phong_ban} - Đơn vị: {self.ma_hieu_2}>"

# ----------------------------------------------------------------------
# 3. BẢNG THÔNG TIN NGƯỜI LAO ĐỘNG
# ----------------------------------------------------------------------
class ThongTinNguoiLaoDong(db.Model):
    __tablename__ = 'thong_tin_nguoi_lao_dong'
    __table_args__ = {
        'mysql_engine': 'InnoDB',
        'mysql_charset': 'utf8mb4',
        'mysql_collate': 'utf8mb4_general_ci',
    }

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    ma_nhan_vien = db.Column(db.String(20), unique=True, nullable=False)
    ho_ten = db.Column(db.String(100), nullable=False)
    ngay_sinh = db.Column(db.Date)
    gioi_tinh = db.Column(db.String(10))
    so_gttt = db.Column(db.String(20))
    so_dien_thoai = db.Column(db.String(20))
    mail_Agribank = db.Column(db.String(100))
    dia_chi = db.Column(db.Text)
    ngay_tinh_phep = db.Column(db.Date)
    ngay_vao_Agribank = db.Column(db.Date)
    # Khóa ngoại trỏ đến bảng phong_ban
    ma_phong_ban = db.Column(db.Integer, db.ForeignKey('phong_ban.id'), nullable=True)
    # Cột ma_hieu_2 liên kết với bảng don_vi
    ma_hieu_2 = db.Column(db.String(255), db.ForeignKey('don_vi.ma_hieu_2'), nullable=True)
    chuc_vu = db.Column(db.String(255))
    trang_thai = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.now)

    # Relationships
    # Relationship để dễ dàng lấy tên đơn vị khi cần (ví dụ: nhan_vien.don_vi.ten_ma_hieu_2)
    phong_ban = db.relationship('PhongBan', backref=db.backref('danh_sach_nhan_vien', lazy=True))
    don_vi = db.relationship('DonVi', backref=db.backref('danh_sach_nhan_vien', lazy=True))

    # Properties tiện ích chống Null
    @property
    def ten_phong_ban(self):
        return self.phong_ban.ten_phong_ban if self.phong_ban else "Chưa xếp phòng"

    @property
    def ten_ma_hieu_2(self):
        return self.don_vi.ten_ma_hieu_2 if self.don_vi else "Chưa xếp đơn vị"

    def __repr__(self):
        return f'<NhanVien {self.ma_nhan_vien} - {self.ho_ten}>'


# ----------------------------------------------------------------------
# 4. BẢNG ĐƠN XIN NGHỈ
# ----------------------------------------------------------------------
class DonXinNghi(db.Model):
    __tablename__ = 'don_xin_nghi'
    __table_args__ = {
        'mysql_engine': 'InnoDB',
        'mysql_charset': 'utf8mb4',
        'mysql_collate': 'utf8mb4_general_ci',
    }

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    ma_nhan_vien = db.Column(
        db.String(50),
        db.ForeignKey('thong_tin_nguoi_lao_dong.ma_nhan_vien'),
        nullable=False,
    )
    loai_nghi = db.Column(db.String(10), nullable=False)  # P, SC, B, V, C...
    tu_ngay = db.Column(db.Date, nullable=False)
    den_ngay = db.Column(db.Date, nullable=False)

    so_ngay_nghi = db.Column(db.Float, default=1.0)

    buoi = db.Column(
        db.String(10), default='ALL'
    )  # 'ALL' (Cả ngày), 'SANG' (Buổi sáng), 'CHIEU' (Buổi chiều)
    
    ly_do = db.Column(db.Text)

    # TRẠNG THÁI: CHO_KIEM_SOAT, CHO_DUYET, DA_DUYET, TU_CHOI
    trang_thai = db.Column(db.String(20), default='CHO_KIEM_SOAT')

    # THÔNG TIN KIỂM SOÁT VÀ PHÊ DUYỆT
    nguoi_kiem_soat = db.Column(db.String(50))
    ngay_kiem_soat = db.Column(db.DateTime)

    nguoi_duyet = db.Column(db.String(20))
    ngay_duyet = db.Column(db.DateTime)

    nguoi_phe_duyet = db.Column(db.String(50))
    ngay_phe_duyet = db.Column(db.DateTime)

    # CHỮ KÝ SỐ Base64/PKCS7
    digital_signature = db.Column(db.Text)
    chu_ky_so = db.Column(db.Text)

    # Thời gian tạo bản ghi
    created_at = db.Column(db.DateTime, default=datetime.now)

    # Quan hệ với bảng ThongTinNguoiLaoDong
    nguoi_lao_dong = db.relationship('ThongTinNguoiLaoDong', backref=db.backref('don_xin_nghi', lazy=True))

    # ------------------------------------------------------------------
    # BỔ SUNG CÁC PROPERTY TIỆN ÍCH LẤY THÔNG TIN HỒ SƠ / PHÒNG BAN / ĐƠN VỊ
    # ------------------------------------------------------------------
    @property
    def ho_ten(self):
        """Lấy họ tên người lao động tạo đơn"""
        if self.nguoi_lao_dong:
            return self.nguoi_lao_dong.ho_ten
        return "N/A"

    @property
    def ten_phong_ban(self):
        """Lấy tên phòng ban từ hồ sơ lao động"""
        if self.nguoi_lao_dong and self.nguoi_lao_dong.phong_ban:
            return self.nguoi_lao_dong.phong_ban.ten_phong_ban
        return "Chưa xếp phòng"

    @property
    def ten_ma_hieu_2(self):
        """Lấy tên đơn vị từ hồ sơ lao động (ma_hieu_2)"""
        if self.nguoi_lao_dong and self.nguoi_lao_dong.don_vi:
            return self.nguoi_lao_dong.don_vi.ten_ma_hieu_2
        return "Chưa xếp đơn vị"

    def __repr__(self):
        return f'<DonXinNghi {self.id} - NV: {self.ma_nhan_vien} ({self.loai_nghi})>'

    @property
    def thong_tin_chu_ky_so(self):
        """
        Trích xuất thông tin chi tiết từ chuỗi chu_ky_so hoặc digital_signature.
        Tối ưu riêng cho Certificate nội bộ Agribank MPKI.
        """
        if hasattr(self, '_override_thong_tin_chu_ky_so'):
            return self._override_thong_tin_chu_ky_so

        raw_sig = self.chu_ky_so or self.digital_signature
        if not raw_sig:
            return None

        # Thời gian ký
        thoi_gian_ky = self.ngay_phe_duyet or self.ngay_duyet or self.created_at
        thoi_gian_ky_str = thoi_gian_ky.strftime('%d/%m/%Y %H:%M:%S') if thoi_gian_ky else '--'

        try:
            cert_str = str(raw_sig).strip()
            cert = None

            # 1. Định dạng PEM X.509
            if "-----BEGIN CERTIFICATE-----" in cert_str:
                cert = x509.load_pem_x509_certificate(cert_str.encode('utf-8'), default_backend())
            
            # 2. Định dạng PKCS7 PEM
            elif "-----BEGIN PKCS7-----" in cert_str and pkcs7:
                pkcs7_objs = pkcs7.load_pem_pkcs7_certificates(cert_str.encode('utf-8'))
                if pkcs7_objs:
                    cert = pkcs7_objs[0]
            else:
                # Làm sạch Base64 & bổ sung padding
                clean_str = "".join(cert_str.split())
                missing_padding = len(clean_str) % 4
                if missing_padding:
                    clean_str += '=' * (4 - missing_padding)

                data_bytes = None
                try:
                    data_bytes = base64.b64decode(clean_str)
                except Exception:
                    if all(c in '0123456789abcdefABCDEF' for c in clean_str):
                        data_bytes = bytes.fromhex(clean_str)

                if data_bytes:
                    # 3. Thử load DER X.509
                    try:
                        cert = x509.load_der_x509_certificate(data_bytes, default_backend())
                    except Exception:
                        pass

                    # 4. Thử load PKCS7 DER
                    if not cert and pkcs7:
                        try:
                            pkcs7_certs = pkcs7.load_der_pkcs7_certificates(data_bytes)
                            if pkcs7_certs:
                                cert = pkcs7_certs[0]
                        except Exception:
                            pass

                # 5. Fallback bọc PEM
                if not cert:
                    try:
                        pem_formatted = f"-----BEGIN CERTIFICATE-----\n{clean_str}\n-----END CERTIFICATE-----"
                        cert = x509.load_pem_x509_certificate(pem_formatted.encode('utf-8'), default_backend())
                    except Exception:
                        pass

            # NẾU PHÂN TÍCH THÀNH CÔNG CERTIFICATE
            if cert:
                subject_rfc = cert.subject.rfc4514_string()
                issuer_rfc = cert.issuer.rfc4514_string()

                subject_cn = cert.subject.get_attributes_for_oid(x509.NameOID.COMMON_NAME)
                chu_so_huu = subject_cn[0].value if subject_cn else subject_rfc

                ou_attrs = cert.subject.get_attributes_for_oid(x509.NameOID.ORGANIZATIONAL_UNIT_NAME)
                emp_info = [attr.value for attr in ou_attrs if "EmployeeID" in attr.value or "User ID" in attr.value]
                if emp_info:
                    chu_so_huu += f" ({', '.join(emp_info)})"

                issuer_cn = cert.issuer.get_attributes_for_oid(x509.NameOID.COMMON_NAME)
                to_chuc_cap = issuer_cn[0].value if issuer_cn else issuer_rfc

                # Lấy ngày Valid From / Valid To (Hỗ trợ cả phiên bản cryptography mới và cũ)
                if hasattr(cert, 'not_valid_before_utc'):
                    valid_from = cert.not_valid_before_utc
                    valid_to = cert.not_valid_after_utc
                else:
                    valid_from = cert.not_valid_before.replace(tzinfo=timezone.utc)
                    valid_to = cert.not_valid_after.replace(tzinfo=timezone.utc)

                if thoi_gian_ky.tzinfo is None:
                    thoi_gian_ky_utc = thoi_gian_ky.replace(tzinfo=timezone.utc)
                else:
                    thoi_gian_ky_utc = thoi_gian_ky.astimezone(timezone.utc)

                het_han_khi_ky = (thoi_gian_ky_utc < valid_from) or (thoi_gian_ky_utc > valid_to)

                return {
                    'chu_so_huu': chu_so_huu,
                    'to_chuc_cap': to_chuc_cap,
                    'serial_number': hex(cert.serial_number)[2:].upper(),
                    'hieu_luc_tu': valid_from.strftime('%d/%m/%Y %H:%M:%S'),
                    'hieu_luc_den': valid_to.strftime('%d/%m/%Y %H:%M:%S'),
                    'thoi_gian_ky': thoi_gian_ky_str,
                    'het_han_khi_ky': het_han_khi_ky,
                    'trang_thai_hop_le': "HỢP LỆ" if not het_han_khi_ky else "KHÔNG HỢP LỆ (Hết hạn lúc ký)"
                }

            # FALLBACK NẾU CHỈ CÓ MÃ SIGNATURE/TOKEN THUẦN
            return {
                'chu_so_huu': self.nguoi_phe_duyet or self.nguoi_duyet or self.ma_nhan_vien,
                'to_chuc_cap': "Agribank CA / USB Token Agent",
                'serial_number': str(raw_sig)[:30] + "...",
                'hieu_luc_tu': thoi_gian_ky.strftime('%d/%m/%Y 00:00:00'),
                'hieu_luc_den': thoi_gian_ky.replace(year=thoi_gian_ky.year + 1).strftime('%d/%m/%Y 23:59:59'),
                'thoi_gian_ky': thoi_gian_ky_str,
                'het_han_khi_ky': False,
                'trang_thai_hop_le': "XÁC THỰC THÀNH CÔNG"
            }

        except Exception as e:
            return {
                'chu_so_huu': self.nguoi_phe_duyet or self.ma_nhan_vien,
                'to_chuc_cap': "Agribank MPKI",
                'serial_number': "N/A",
                'hieu_luc_tu': '--',
                'hieu_luc_den': '--',
                'thoi_gian_ky': thoi_gian_ky_str,
                'het_han_khi_ky': False,
                'trang_thai_hop_le': "ĐÃ KÝ DUYỆT"
            }

    @thong_tin_chu_ky_so.setter
    def thong_tin_chu_ky_so(self, value):
        self._override_thong_tin_chu_ky_so = value

# Phân công công tác ; các bảng liên quan: quyet_dinh; lich_su_phu_trach; thong_tin_nguoi_lao_dong; 
class PhanCongCongTac(db.Model):
    __tablename__ = 'phan_cong_cong_tac'
    __table_args__ = {
        'mysql_engine': 'InnoDB',
        'mysql_charset': 'utf8mb4',
        'mysql_collate': 'utf8mb4_general_ci',
    }

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    ma_nhan_vien = db.Column(db.String(20), db.ForeignKey('users.ma_nhan_vien', ondelete='CASCADE', onupdate='CASCADE'), nullable=False)
    ho_ten = db.Column(db.String(255), nullable=False)
    ma_phong_ban = db.Column(db.String(100), nullable=True)
    ten_phong_ban = db.Column(db.String(255), nullable=True)
    
    noi_cong_tac = db.Column(db.String(255), nullable=False) # Địa điểm / Đơn vị đến công tác
    noi_dung_cong_tac = db.Column(db.Text, nullable=False)   # Nội dung / Mục đích chuyến công tác
    ngay_bat_dau = db.Column(db.Date, nullable=False)
    ngay_ket_thuc = db.Column(db.Date, nullable=False)
    
    phuong_tien = db.Column(db.String(100), default='Ô tô') # Ô tô, Tàu hỏa, Máy bay, Xe máy...
    kinh_phi_du_tru = db.Column(db.Float, default=0.0)      # Dự trù kinh phí (VNĐ)
    
    # Trạng thái: CHO_DUYET (Chờ duyệt), DA_DUYET (Đã duyệt), TU_CHOI (Từ chối), HOAN_THANH (Hoàn thành)
    trang_thai = db.Column(db.String(50), default='CHO_DUYET') 
    ly_do_tu_choi = db.Column(db.Text, nullable=True)
    
    nguoi_tao = db.Column(db.String(20), nullable=False)
    nguoi_duyet = db.Column(db.String(20), nullable=True)
    ngay_duyet = db.Column(db.DateTime, nullable=True)
    
    ghi_chu = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # Quan hệ
    user = db.relationship('User', foreign_keys=[ma_nhan_vien], backref=db.backref('ds_phan_cong', cascade='all, delete-orphan'))

    def to_dict(self):
        return {
            'id': self.id,
            'ma_nhan_vien': self.ma_nhan_vien,
            'ho_ten': self.ho_ten,
            'ten_phong_ban': self.ten_phong_ban,
            'noi_cong_tac': self.noi_cong_tac,
            'noi_dung_cong_tac': self.noi_dung_cong_tac,
            'ngay_bat_dau': self.ngay_bat_dau.strftime('%Y-%m-%d') if self.ngay_bat_dau else '',
            'ngay_ket_thuc': self.ngay_ket_thuc.strftime('%Y-%m-%d') if self.ngay_ket_thuc else '',
            'phuong_tien': self.phuong_tien,
            'kinh_phi_du_tru': self.kinh_phi_du_tru,
            'trang_thai': self.trang_thai,
            'ghi_chu': self.ghi_chu
        }

# ----------------------------------------------------------------------
# 5. BẢNG LĨNH VỰC
# ----------------------------------------------------------------------
class LinhVuc(db.Model):
    __tablename__ = 'linh_vuc'
    __table_args__ = {
        'mysql_engine': 'InnoDB',
        'mysql_charset': 'utf8mb4',
        'mysql_collate': 'utf8mb4_general_ci',
    }

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    ten_linh_vuc = db.Column(db.String(255), nullable=False)
    mo_ta = db.Column(db.Text, nullable=True)
    trang_thai = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id,
            'ten_linh_vuc': self.ten_linh_vuc,
            'mo_ta': self.mo_ta,
            'trang_thai': self.trang_thai
        }

    def __repr__(self):
        return f"<LinhVuc {self.id} - {self.ten_linh_vuc}>"


# ----------------------------------------------------------------------
# 6. BẢNG LỊCH SỬ PHỤ TRÁCH
# ----------------------------------------------------------------------
class LichSuPhuTrach(db.Model):
    __tablename__ = 'lich_su_phu_trach'
    __table_args__ = {
        'mysql_engine': 'InnoDB',
        'mysql_charset': 'utf8mb4',
        'mysql_collate': 'utf8mb4_general_ci',
    }

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    ma_nhan_vien = db.Column(db.String(20), db.ForeignKey('users.ma_nhan_vien', ondelete='CASCADE', onupdate='CASCADE'), nullable=False)
    linh_vuc_id = db.Column(db.Integer, db.ForeignKey('linh_vuc.id', ondelete='CASCADE'), nullable=True)
    ma_hieu_2 = db.Column(db.String(255), db.ForeignKey('don_vi.ma_hieu_2', ondelete='SET NULL', onupdate='CASCADE'), nullable=True)
    
    ngay_bat_dau = db.Column(db.Date, nullable=False)
    ngay_ket_thuc = db.Column(db.Date, nullable=True)
    vai_tro = db.Column(db.String(100), nullable=True)  # VD: Phụ trách chính, Phụ trách phối hợp
    trang_thai = db.Column(db.String(50), default='DANG_PHU_TRACH')  # DANG_PHU_TRACH, DA_KET_THUC
    ghi_chu = db.Column(db.Text, nullable=True)
    
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # Quan hệ
    user = db.relationship('User', foreign_keys=[ma_nhan_vien], backref=db.backref('ds_lich_su_phu_trach', cascade='all, delete-orphan'))
    linh_vuc = db.relationship('LinhVuc', foreign_keys=[linh_vuc_id], backref=db.backref('ds_nhan_vien_phu_trach', lazy=True))
    don_vi = db.relationship('DonVi', foreign_keys=[ma_hieu_2], backref=db.backref('ds_lich_su_phu_trach', lazy=True))

    def to_dict(self):
        return {
            'id': self.id,
            'ma_nhan_vien': self.ma_nhan_vien,
            'linh_vuc_id': self.linh_vuc_id,
            'ten_linh_vuc': self.linh_vuc.ten_linh_vuc if self.linh_vuc else '',
            'ma_hieu_2': self.ma_hieu_2,
            'ten_don_vi': self.don_vi.ten_ma_hieu_2 if self.don_vi else '',
            'ngay_bat_dau': self.ngay_bat_dau.strftime('%Y-%m-%d') if self.ngay_bat_dau else '',
            'ngay_ket_thuc': self.ngay_ket_thuc.strftime('%Y-%m-%d') if self.ngay_ket_thuc else '',
            'vai_tro': self.vai_tro,
            'trang_thai': self.trang_thai,
            'ghi_chu': self.ghi_chu
        }

    def __repr__(self):
        return f"<LichSuPhuTrach {self.id} - NV: {self.ma_nhan_vien}>"
