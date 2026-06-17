// ════════════════════════════════════════════════════════════
//   COLLEGE RESULT ANALYTICS — Shared JS Utilities
// ════════════════════════════════════════════════════════════

const API = {
  base: '/api',

  async post(endpoint, data) {
    const res = await fetch(this.base + endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'include',
      body: JSON.stringify(data)
    });
    return res.json();
  },

  async get(endpoint) {
    const res = await fetch(this.base + endpoint, {
      credentials: 'include'
    });
    return res.json();
  },

  async patch(endpoint, data) {
    const res = await fetch(this.base + endpoint, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'include',
      body: JSON.stringify(data)
    });
    return res.json();
  }
};

// Toast notification
function showToast(message, type = 'success') {
  let container = document.getElementById('toast-container');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    container.className = 'toast-container';
    document.body.appendChild(container);
  }

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  const icon = type === 'success' ? '✅' : '❌';
  toast.innerHTML = `<span>${icon}</span><span>${message}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(30px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

// Check authentication
async function checkAuth(requireAdmin = false) {
  try {
    const data = await API.get('/auth/me');
    if (!data.success) {
      window.location.href = '/login';
      return null;
    }
    if (requireAdmin && data.user.role !== 'Admin') {
      window.location.href = '/dashboard';
      return null;
    }
    return data.user;
  } catch (e) {
    window.location.href = '/login';
    return null;
  }
}

// Logout
async function logout() {
  try {
    await API.post('/auth/logout', {});
  } catch (e) {}
  window.location.href = '/login';
}

// Generate floating particles
function createParticles(count = 20) {
  const bg = document.querySelector('.analytics-bg');
  if (!bg) return;
  for (let i = 0; i < count; i++) {
    const p = document.createElement('div');
    p.className = 'particle';
    p.style.left = Math.random() * 100 + '%';
    p.style.animationDuration = (Math.random() * 15 + 10) + 's';
    p.style.animationDelay = (Math.random() * 10) + 's';
    p.style.setProperty('--drift', (Math.random() * 100 - 50) + 'px');
    bg.appendChild(p);
  }
}

// Format date
function formatDate(dateStr) {
  const d = new Date(dateStr);
  return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
}

// Escape HTML to prevent XSS
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// Role badge HTML
function roleBadge(role) {
  const cls = {
    'Student': 'badge-student',
    'Faculty': 'badge-faculty',
    'Researcher': 'badge-researcher',
    'Admin': 'badge-admin'
  }[role] || 'badge-student';
  return `<span class="badge-role ${cls}">${escapeHtml(role)}</span>`;
}
