# extensions.py
# app.py và forum.py đều import db từ extensions.py
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from flask_login import LoginManager
from functools import wraps
from flask import abort
from flask_login import current_user

db = SQLAlchemy()
bcrypt = Bcrypt()
login_manager = LoginManager()

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # Kiểm tra user đã đăng nhập VÀ có quyền admin (is_admin)
        if not (current_user.is_authenticated and getattr(current_user, 'is_admin', False)):
            abort(403)
        return f(*args, **kwargs)
    return decorated_function
