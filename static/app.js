// ── Action dropdown / split button ──
function toggleDropdown(btn) {
  const dd = btn.closest('.action-dropdown, .split-btn');
  const isOpen = dd.classList.contains('open');
  document.querySelectorAll('.action-dropdown.open, .split-btn.open').forEach(d => d.classList.remove('open'));
  if (!isOpen) dd.classList.add('open');
}
document.addEventListener('click', (e) => {
  if (!e.target.closest('.action-dropdown, .split-btn')) {
    document.querySelectorAll('.action-dropdown.open, .split-btn.open').forEach(d => d.classList.remove('open'));
  }
});

// ── Toast auto-dismiss ──
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.toast').forEach(t => {
    setTimeout(() => t.style.opacity = '0', 3000);
    setTimeout(() => t.remove(), 3500);
  });
});

// ── Confirm delete ──
function confirmDelete(formId, message) {
  if (confirm(message || 'Bạn có chắc muốn xóa không?')) {
    document.getElementById(formId).submit();
  }
}

// ── Show/hide loading overlay ──
function showLoading(text) {
  let el = document.getElementById('loading-overlay');
  if (!el) {
    el = document.createElement('div');
    el.id = 'loading-overlay';
    el.className = 'loading-overlay';
    el.innerHTML = `<div class="spinner"></div><div class="loading-text">${text || 'Đang xử lý...'}</div>`;
    document.body.appendChild(el);
  }
}
function hideLoading() {
  const el = document.getElementById('loading-overlay');
  if (el) el.remove();
}

// ── Toast helper (JS-triggered) ──
function showToast(message, type = 'success') {
  let container = document.querySelector('.toast-container');
  if (!container) {
    container = document.createElement('div');
    container.className = 'toast-container';
    document.body.appendChild(container);
  }
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => toast.style.opacity = '0', 3000);
  setTimeout(() => toast.remove(), 3500);
}
