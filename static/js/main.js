// Tự động ẩn Flash messages sau 3 giây
document.addEventListener("DOMContentLoaded", function() {
    setTimeout(function() {
        const alerts = document.querySelectorAll('#flash-container .alert');
        alerts.forEach(alert => {
            if (alert) {
                alert.style.transition = "all 0.6s ease";
                alert.style.opacity = "0";
                alert.style.transform = "translateX(50px)";
                setTimeout(() => { if (alert.parentNode) { alert.remove(); } }, 600);
            }
        });
    }, 3000); 
});

// Tải thông báo qua API
function loadNotifications() {
    $.get('/api/notifications', function(data) {
        if (data.unread_count > 0) {
            $('#noti-count').text(data.unread_count > 9 ? '9+' : data.unread_count).show();
        } else {
            $('#noti-count').hide();
        }

        let html = '';
        if (!data.notifications || data.notifications.length === 0) {
            html = '<div class="p-4 text-center text-muted small">Không có thông báo mới nào</div>';
        } else {
            data.notifications.forEach(n => {
                let unreadStyle = n.is_read ? '' : 'noti-unread';
                html += `
                <a class="noti-item ${unreadStyle}" href="${n.link}">
                    <div class="noti-icon"><i class="fas fa-info-circle"></i></div>
                    <div class="noti-content">
                        <div class="noti-time">${n.time}</div>
                        <div class="noti-title">${n.tieu_de}</div>
                        <div class="noti-desc">${n.noi_dung}</div>
                    </div>
                </a>`;
            });
        }
        $('#noti-list').html(html);
    }).fail(function() {
        console.log("Không thể tải thông báo.");
    });
}

// Đánh dấu tất cả thông báo là đã đọc
function markAllRead() {
    $.post('/api/notifications/read-all', function() {
        loadNotifications();
    });
}