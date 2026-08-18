# modules/forum.py
from functools import wraps
from flask import Blueprint, render_template, request, redirect, url_for, flash, abort
from flask_login import login_required, current_user
from datetime import datetime
import pytz

# Import db từ extensions, KHÔNG import từ app; extensions.py cùng thư mục với app.py
from extensions import db
from modules.vector_service import index_topic_to_chroma, search_similar_topics, remove_topic_from_chroma

# ----------------------------------------------------------------------
# 1. KHỞI TẠO BLUEPRINT
# ----------------------------------------------------------------------
forum_bp = Blueprint('forum', __name__, url_prefix='/forum')

HANOI_TZ = pytz.timezone('Asia/Ho_Chi_Minh')
def get_hanoi_now():
    return datetime.now(HANOI_TZ)

def get_forum_engine():
    return db.get_engine(bind='db_forum')

# ----------------------------------------------------------------------
# 2. DATABASE MODELS
# ----------------------------------------------------------------------
class ForumCategory(db.Model):
    __bind_key__ = 'db_forum'
    __tablename__ = 'forum_categories'
    __table_args__ = {'mysql_collate': 'utf8mb4_general_ci'}

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(150), nullable=False)
    description = db.Column(db.String(255), nullable=True)
    icon = db.Column(db.String(50), default='bi-chat-left-text')
    created_at = db.Column(db.DateTime, default=get_hanoi_now)

    topics = db.relationship('ForumTopic', backref='category', lazy=True, cascade="all, delete-orphan")


class ForumTopic(db.Model):
    __bind_key__ = 'db_forum'
    __tablename__ = 'forum_topics'
    __table_args__ = {'mysql_collate': 'utf8mb4_general_ci'}

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(255), nullable=False)
    content = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=get_hanoi_now)
    updated_at = db.Column(db.DateTime, default=get_hanoi_now, onupdate=get_hanoi_now)

    is_pinned = db.Column(db.Boolean, default=False)
    is_closed = db.Column(db.Boolean, default=False)
    # Bổ sung cột status trùng khớp với DB MySQL enum('pending','approved','rejected')
    status = db.Column(db.Enum('pending', 'approved', 'rejected'), default='approved')
    views_count = db.Column(db.Integer, default=0)

    category_id = db.Column(db.Integer, db.ForeignKey('forum_categories.id'), nullable=False)
    user_id = db.Column(db.Integer, nullable=False)
    author_name = db.Column(db.String(100), nullable=True)

    posts = db.relationship('ForumPost', backref='topic', lazy=True, cascade="all, delete-orphan")


class ForumPost(db.Model):
    __bind_key__ = 'db_forum'
    __tablename__ = 'forum_posts'
    __table_args__ = {'mysql_collate': 'utf8mb4_general_ci'}

    id = db.Column(db.Integer, primary_key=True)
    content = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=get_hanoi_now)
    updated_at = db.Column(db.DateTime, default=get_hanoi_now, onupdate=get_hanoi_now)

    topic_id = db.Column(db.Integer, db.ForeignKey('forum_topics.id'), nullable=False)
    user_id = db.Column(db.Integer, nullable=False)
    author_name = db.Column(db.String(100), nullable=True)


class ForumModerator(db.Model):
    __bind_key__ = 'db_forum'
    __tablename__ = 'forum_moderators'
    __table_args__ = {'mysql_collate': 'utf8mb4_general_ci'}

    id = db.Column(db.Integer, primary_key=True)
    category_id = db.Column(db.Integer, db.ForeignKey('forum_categories.id'), nullable=False)
    user_id = db.Column(db.Integer, nullable=False)
    created_at = db.Column(db.DateTime, default=get_hanoi_now)

# ----------------------------------------------------------------------
# HELPER FUNCTIONS & DECORATORS
# ----------------------------------------------------------------------
def is_forum_admin(user):
    if not user or not user.is_authenticated:
        return False
    return getattr(user, 'is_admin', False) or getattr(user, 'role', '') == 'admin'

def is_category_mod(user, category_id):
    if is_forum_admin(user):
        return True
    if not user or not user.is_authenticated or not category_id:
        return False
    return ForumModerator.query.filter_by(category_id=category_id, user_id=user.id).first() is not None

def mod_or_admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not current_user.is_authenticated:
            abort(401)
            
        if is_forum_admin(current_user):
            return f(*args, **kwargs)
            
        topic_id = kwargs.get('topic_id')
        category_id = kwargs.get('category_id')
        
        if topic_id:
            topic = ForumTopic.query.get_or_404(topic_id)
            category_id = topic.category_id
            
        if not is_category_mod(current_user, category_id):
            abort(403)
            
        return f(*args, **kwargs)
    return decorated_function

# ----------------------------------------------------------------------
# 3. ROUTES / CONTROLLERS
# ----------------------------------------------------------------------
@forum_bp.route('/')
@forum_bp.route('')
@login_required
def index():
    db.create_all(bind_key='db_forum')

    if ForumCategory.query.count() == 0:
        sample_cats = [
            ForumCategory(name="Thông Báo & Tin Tức Nội Bộ", description="Các thông báo chính thức từ ban quản lý", icon="bi-megaphone-fill"),
            ForumCategory(name="Nghiệp Vụ BC48 & PCRT", description="Trao đổi, giải đáp thắc mắc về báo cáo BC48 và chống rửa tiền", icon="bi-shield-check"),
            ForumCategory(name="Hỗ Trợ Kỹ Thuật & Phần Mềm", description="Hỏi đáp về hạ tầng CNTT, lỗi phần mềm, máy in", icon="bi-laptop"),
            ForumCategory(name="Thảo Luận Chung", description="Giao lưu, chia sẻ kinh nghiệm công việc hàng ngày", icon="bi-chat-dots-fill")
        ]
        db.session.add_all(sample_cats)
        db.session.commit()

    categories = ForumCategory.query.all()
    
    # Chỉ hiển thị bài viết có trạng thái 'approved'
    recent_topics = ForumTopic.query.filter_by(status='approved').order_by(
        ForumTopic.is_pinned.desc(), 
        ForumTopic.created_at.desc()
    ).limit(15).all()

    return render_template('forum/index.html', categories=categories, recent_topics=recent_topics)


@forum_bp.route('/topic/create', methods=['POST'])
@login_required
def create_topic():
    title = request.form.get('title', '').strip()
    category_id = request.form.get('category_id')
    content = request.form.get('content', '').strip()

    if not title or not category_id or not content:
        flash("Vui lòng điền đầy đủ thông tin bài viết!", "danger")
        return redirect(url_for('forum.index'))

    user_display_name = getattr(current_user, 'full_name', getattr(current_user, 'username', 'Thành viên'))

    new_topic = ForumTopic(
        title=title,
        content=content,
        category_id=category_id,
        user_id=current_user.id,
        author_name=user_display_name,
        status='approved'  # Mặc định duyệt bài (hoặc đổi thành 'pending' nếu cần duyệt)
    )
    db.session.add(new_topic)
    db.session.commit()

    # Đánh chỉ mục bài viết vào ChromaDB
    try:
        index_topic_to_chroma(
            topic_id=new_topic.id,
            title=new_topic.title,
            content=new_topic.content,
            category_id=new_topic.category_id,
            author_name=new_topic.author_name,
            status=new_topic.status
        )
    except Exception as e:
        print(f"[ChromaDB Error] Không thể đánh chỉ mục bài viết ID {new_topic.id}: {e}")

    flash("Đăng bài viết mới thành công!", "success")
    return redirect(url_for('forum.index'))


@forum_bp.route('/search', methods=['GET'])
@login_required
def search_topics():
    query = request.args.get('q', '').strip()
    if not query:
        return redirect(url_for('forum.index'))

    matched_ids = search_similar_topics(query, top_k=15)
    
    topics = []
    if matched_ids:
        # Lọc các bài viết vừa khớp ngữ nghĩa vừa có trạng thái approved
        fetched_topics = ForumTopic.query.filter(
            ForumTopic.id.in_(matched_ids),
            ForumTopic.status == 'approved'
        ).all()
        
        topic_dict = {t.id: t for t in fetched_topics}
        topics = [topic_dict[tid] for tid in matched_ids if tid in topic_dict]

    return render_template('forum/search_results.html', topics=topics, query=query)


@forum_bp.route('/sync-vectors', methods=['GET'])
@login_required
def sync_existing_topics():
    """Route Admin dùng để đồng bộ dữ liệu cũ từ MySQL vào ChromaDB"""
    if not is_forum_admin(current_user):
        abort(403)
        
    approved_topics = ForumTopic.query.filter_by(status='approved').all()
    count = 0
    for t in approved_topics:
        index_topic_to_chroma(
            topic_id=t.id,
            title=t.title,
            content=t.content,
            category_id=t.category_id,
            author_name=t.author_name,
            status=t.status
        )
        count += 1
        
    flash(f"Đã đồng bộ thành công {count} bài viết sang ChromaDB!", "success")
    return redirect(url_for('forum.index'))
