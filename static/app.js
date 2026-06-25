// ── Số tiền → chữ tiếng Việt ──
function numToViWords(n) {
  if (!n || n === 0) return '';
  const units = ['', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];

  function readGroup(n, needZeroHundred) {
    let s = '';
    const h = Math.floor(n / 100);
    const t = Math.floor((n % 100) / 10);
    const u = n % 10;
    if (h > 0)              s += units[h] + ' trăm';
    else if (needZeroHundred) s += 'không trăm';
    if (t === 0) {
      if (u > 0) s += (s ? ' lẻ ' : '') + units[u];
    } else if (t === 1) {
      s += (s ? ' ' : '') + 'mười';
      if (u === 5) s += ' lăm';
      else if (u > 0) s += ' ' + units[u];
    } else {
      s += (s ? ' ' : '') + units[t] + ' mươi';
      if (u === 1) s += ' mốt';
      else if (u === 5) s += ' lăm';
      else if (u > 0) s += ' ' + units[u];
    }
    return s.trim();
  }

  const ty  = Math.floor(n / 1000000000);
  const tr  = Math.floor((n % 1000000000) / 1000000);
  const ng  = Math.floor((n % 1000000) / 1000);
  const don = n % 1000;
  const parts = [];
  if (ty  > 0) parts.push(readGroup(ty,  false) + ' tỷ');
  if (tr  > 0) parts.push(readGroup(tr,  parts.length > 0) + ' triệu');
  if (ng  > 0) parts.push(readGroup(ng,  parts.length > 0) + ' nghìn');
  if (don > 0) parts.push(readGroup(don, parts.length > 0));
  if (!parts.length) return 'Không đồng chẵn';
  const raw = parts.join(' ');
  return raw.charAt(0).toUpperCase() + raw.slice(1) + ' đồng chẵn';
}

// ── Money input formatter (VN: dấu chấm ngàn) ──
function fmtMoneyInput(el) {
  const pos   = el.selectionStart;
  const before = el.value.length;
  const raw   = el.value.replace(/\D/g, '');
  const fmt   = raw ? raw.replace(/\B(?=(\d{3})+(?!\d))/g, '.') : '';
  el.value = fmt;
  // giữ cursor position
  const after = fmt.length;
  el.setSelectionRange(pos + (after - before), pos + (after - before));
}

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
