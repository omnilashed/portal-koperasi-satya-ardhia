// Global App State
const API_BASE = ''; // Relative URL for Express API
let currentUser = null;
let rawPengajuanList = [];
let filteredPengajuanList = [];
let currentPage = 1;
const pageSize = 10;

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) lucide.createIcons();
  checkAuthSession();
});

// Toast System
function showToast(message, type = 'info') {
  const container = document.getElementById('toast-container');
  if (!container) return;
  const toast = document.createElement('div');
  
  const bgMap = {
    success: 'bg-emerald-600 text-white',
    error: 'bg-rose-600 text-white',
    warning: 'bg-amber-500 text-white',
    info: 'bg-slate-800 text-white'
  };

  const iconMap = {
    success: 'check-circle',
    error: 'alert-triangle',
    warning: 'alert-circle',
    info: 'info'
  };

  toast.className = (bgMap[type] || bgMap.info) + ' px-4 py-3 rounded-xl shadow-lg text-xs sm:text-sm font-medium flex items-center gap-2 transform transition-all duration-300 translate-y-2 opacity-0 pointer-events-auto max-w-sm';
  toast.innerHTML = '<i data-lucide=' + (iconMap[type] || 'info') + ' class=w-4 h-4 flex-shrink-0></i><span class=flex-1>' + message + '</span>';
  
  container.appendChild(toast);
  if (window.lucide) lucide.createIcons();

  setTimeout(() => toast.classList.remove('translate-y-2', 'opacity-0'), 10);
  setTimeout(() => {
    toast.classList.add('opacity-0', 'translate-x-full');
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

// Session Management
function checkAuthSession() {
  const saved = localStorage.getItem('koperasi_user');
  if (saved) {
    try {
      currentUser = JSON.parse(saved);
      renderDashboard();
    } catch (e) {
      localStorage.removeItem('koperasi_user');
      renderAuth();
    }
  } else {
    renderAuth();
  }
}

function renderAuth() {
  document.getElementById('auth-container').classList.remove('hidden');
  document.getElementById('main-container').classList.add('hidden');
  switchAuthView('login');
}

function renderDashboard() {
  document.getElementById('auth-container').classList.add('hidden');
  document.getElementById('main-container').classList.remove('hidden');

  document.getElementById('nav-user-name').textContent = currentUser.namaLengkap || currentUser.username;
  document.getElementById('nav-user-role').textContent = currentUser.role;
  document.getElementById('nav-user-nik').textContent = currentUser.nik || '-';

  const tabUsers = document.getElementById('tab-btn-users');
  const tabForm = document.getElementById('tab-btn-form');

  if (currentUser.role === 'Admin') {
    tabUsers.classList.remove('hidden');
    tabForm.classList.remove('hidden');
  } else if (currentUser.role === 'Financial') {
    tabUsers.classList.add('hidden');
    tabForm.classList.remove('hidden');
  } else if (currentUser.role === 'Pengurus') {
    tabUsers.classList.add('hidden');
    tabForm.classList.add('hidden');
  } else {
    tabUsers.classList.add('hidden');
    tabForm.classList.remove('hidden');
  }

  if (currentUser.role === 'User') {
    const formNama = document.getElementById('form-nama');
    const formNik = document.getElementById('form-nik');
    if (formNama) formNama.value = currentUser.namaLengkap || '';
    if (formNik) formNik.value = currentUser.nik || '';
  }

  switchDashboardTab('tracking');
  fetchPengajuanData();
  if (window.lucide) lucide.createIcons();
}

function switchAuthView(viewName) {
  ['login', 'register', 'forgot', 'reset'].forEach(v => {
    const el = document.getElementById('view-' + v);
    if (el) el.classList.add('hidden');
  });
  const target = document.getElementById('view-' + viewName);
  if (target) target.classList.remove('hidden');
  if (window.lucide) lucide.createIcons();
}

function switchDashboardTab(tabName) {
  ['tracking', 'form', 'users'].forEach(t => {
    const section = document.getElementById('tab-' + t);
    const btn = document.getElementById('tab-btn-' + t);
    if (section) section.classList.add('hidden');
    if (btn) {
      btn.classList.remove('border-brand-500', 'text-brand-600');
      btn.classList.add('border-transparent', 'text-slate-500');
    }
  });

  const activeSection = document.getElementById('tab-' + tabName);
  const activeBtn = document.getElementById('tab-btn-' + tabName);
  if (activeSection) activeSection.classList.remove('hidden');
  if (activeBtn) {
    activeBtn.classList.remove('border-transparent', 'text-slate-500');
    activeBtn.classList.add('border-brand-500', 'text-brand-600');
  }

  if (tabName === 'users' && currentUser && currentUser.role === 'Admin') {
    fetchUsersData();
  }
  if (window.lucide) lucide.createIcons();
}

// Authentication Actions
async function handleLogin(e) {
  e.preventDefault();
  const btn = document.getElementById('btn-login');
  const originalHtml = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = '<span>Memvalidasi...</span>';

  const username = document.getElementById('login-username').value;
  const password = document.getElementById('login-password').value;

  try {
    const res = await fetch(API_BASE + '/api/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    });
    const data = await res.json();

    if (data.success) {
      currentUser = data.user;
      localStorage.setItem('koperasi_user', JSON.stringify(currentUser));
      showToast('Selamat datang, ' + currentUser.namaLengkap + '!', 'success');
      renderDashboard();
    } else {
      showToast(data.message || 'Login gagal.', 'error');
    }
  } catch (err) {
    showToast('Koneksi server gagal: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalHtml;
    if (window.lucide) lucide.createIcons();
  }
}

async function handleRegister(e) {
  e.preventDefault();
  const btn = document.getElementById('btn-register');
  btn.disabled = true;
  btn.innerHTML = '<span>Memproses...</span>';

  const payload = {
    username: document.getElementById('reg-username').value,
    password: document.getElementById('reg-password').value,
    namaLengkap: document.getElementById('reg-nama').value,
    nik: document.getElementById('reg-nik').value,
    email: document.getElementById('reg-email').value,
    role: 'User'
  };

  try {
    const res = await fetch(API_BASE + '/api/register', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();

    if (data.success) {
      showToast(data.message, 'success');
      document.getElementById('form-register').reset();
      switchAuthView('login');
    } else {
      showToast(data.message || 'Registrasi gagal.', 'error');
    }
  } catch (err) {
    showToast('Koneksi error: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<span>Kirim Pendaftaran</span>';
  }
}

async function handleForgotPassword(e) {
  e.preventDefault();
  const btn = document.getElementById('btn-forgot');
  btn.disabled = true;
  btn.innerHTML = '<span>Mengirim OTP...</span>';

  const email = document.getElementById('forgot-email').value;

  try {
    const res = await fetch(API_BASE + '/api/auth/forgot-password', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email })
    });
    const data = await res.json();

    if (data.success) {
      showToast(data.message, 'success');
      if (data.devOtp) console.log('[DEV OTP]:', data.devOtp);
      document.getElementById('reset-email').value = email;
      switchAuthView('reset');
    } else {
      showToast(data.message || 'Gagal mengirim OTP.', 'error');
    }
  } catch (err) {
    showToast('Koneksi error: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<span>Kirim Kode OTP</span>';
  }
}

async function handleResetPassword(e) {
  e.preventDefault();
  const btn = document.getElementById('btn-reset');
  btn.disabled = true;
  btn.innerHTML = '<span>Menyimpan...</span>';

  const payload = {
    email: document.getElementById('reset-email').value,
    otp: document.getElementById('reset-otp').value,
    newPassword: document.getElementById('reset-new-password').value
  };

  try {
    const res = await fetch(API_BASE + '/api/auth/reset-password', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();

    if (data.success) {
      showToast(data.message, 'success');
      document.getElementById('form-reset').reset();
      switchAuthView('login');
    } else {
      showToast(data.message || 'Reset password gagal.', 'error');
    }
  } catch (err) {
    showToast('Koneksi error: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<span>Simpan Password Baru</span>';
  }
}

function handleLogout() {
  localStorage.removeItem('koperasi_user');
  currentUser = null;
  showToast('Sesi Anda telah berakhir.', 'info');
  renderAuth();
}

// Form Pengajuan Logic
function toggleCustomTipe(val) {
  const customInput = document.getElementById('form-tipe-custom');
  if (val === 'CUSTOM') {
    customInput.classList.remove('hidden');
    customInput.required = true;
  } else {
    customInput.classList.add('hidden');
    customInput.required = false;
  }
}

async function handlePengajuanSubmit(e) {
  e.preventDefault();
  const btn = document.getElementById('btn-submit-pengajuan');
  const originalHtml = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = '<span>Menyimpan Pengajuan...</span>';

  const selectTipe = document.getElementById('form-tipe').value;
  const tipeFormulir = selectTipe === 'CUSTOM' ? document.getElementById('form-tipe-custom').value : selectTipe;

  const payload = {
    tipeFormulir,
    namaKaryawan: document.getElementById('form-nama').value,
    nikUnitKerja: document.getElementById('form-nik').value,
    noRekening: document.getElementById('form-rekening').value,
    atasNamaRekening: document.getElementById('form-atas-nama').value,
    totalDitransfer: document.getElementById('form-nominal').value,
    cicilan: document.getElementById('form-cicilan').value,
    keterangan: document.getElementById('form-keterangan').value
  };

  try {
    const res = await fetch(API_BASE + '/api/pengajuan', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();

    if (data.success) {
      showToast(data.message, 'success');
      document.getElementById('form-pengajuan').reset();
      switchDashboardTab('tracking');
      fetchPengajuanData();
    } else {
      showToast(data.message || 'Gagal menyimpan pengajuan.', 'error');
    }
  } catch (err) {
    showToast('Koneksi error: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalHtml;
    if (window.lucide) lucide.createIcons();
  }
}

// Data Fetching: Pengajuan
async function fetchPengajuanData() {
  if (!currentUser) return;
  
  const tbody = document.getElementById('table-pengajuan-body');
  tbody.innerHTML = '<tr><td colspan=9 class=py-8 text-center text-slate-400>Memuat data pengajuan...</td></tr>';

  let url = API_BASE + '/api/pengajuan?role=' + encodeURIComponent(currentUser.role);
  if (currentUser.role === 'User') {
    url += '&nik=' + encodeURIComponent(currentUser.nik) + '&nama=' + encodeURIComponent(currentUser.namaLengkap);
  }

  try {
    const res = await fetch(url);
    const json = await res.json();

    if (json.success) {
      rawPengajuanList = json.data || [];
      calculateKpiMetrics(rawPengajuanList);
      applyPengajuanFilters();
    } else {
      showToast(json.message || 'Gagal mengambil data.', 'error');
      tbody.innerHTML = '<tr><td colspan=9 class=py-8 text-center text-rose-500>Error memuat data</td></tr>';
    }
  } catch (err) {
    tbody.innerHTML = '<tr><td colspan=9 class=py-8 text-center text-rose-500>Koneksi backend terputus</td></tr>';
  }
}

function calculateKpiMetrics(list) {
  document.getElementById('kpi-total').textContent = list.length;
  document.getElementById('kpi-created').textContent = list.filter(x => (x.status || '').toLowerCase() === 'created').length;
  document.getElementById('kpi-process').textContent = list.filter(x => (x.status || '').toLowerCase() === 'on process').length;
  document.getElementById('kpi-done').textContent = list.filter(x => (x.status || '').toLowerCase() === 'done').length;
}

function applyPengajuanFilters() {
  const keyword = (document.getElementById('filter-search').value || '').toLowerCase();
  const status = document.getElementById('filter-status').value;
  const tipe = document.getElementById('filter-tipe').value;

  filteredPengajuanList = rawPengajuanList.filter(item => {
    const matchSearch = 
      (item.namaKaryawan || '').toLowerCase().includes(keyword) ||
      (item.nikUnitKerja || '').toLowerCase().includes(keyword) ||
      (item.noRekening || '').toLowerCase().includes(keyword) ||
      (item.keterangan || '').toLowerCase().includes(keyword);

    const matchStatus = (status === 'ALL') || (item.status === status);
    const matchTipe = (tipe === 'ALL') || (item.tipeFormulir === tipe);

    return matchSearch && matchStatus && matchTipe;
  });

  currentPage = 1;
  renderPengajuanTable();
}

function renderPengajuanTable() {
  const tbody = document.getElementById('table-pengajuan-body');
  const total = filteredPengajuanList.length;

  if (total === 0) {
    tbody.innerHTML = '<tr><td colspan=9 class=py-8 text-center text-slate-400>Tidak ada data pengajuan yang sesuai.</td></tr>';
    document.getElementById('pagination-info').textContent = 'Menampilkan 0 data';
    document.getElementById('pagination-controls').innerHTML = '';
    return;
  }

  const totalPages = Math.ceil(total / pageSize);
  const startIdx = (currentPage - 1) * pageSize;
  const pageItems = filteredPengajuanList.slice(startIdx, startIdx + pageSize);

  const statusBadge = (status) => {
    switch (status) {
      case 'Created':
        return '<span class=px-2.5 py-1 text-xs font-semibold rounded-full bg-amber-50 text-amber-700 border border-amber-200>Created</span>';
      case 'On Process':
        return '<span class=px-2.5 py-1 text-xs font-semibold rounded-full bg-blue-50 text-blue-700 border border-blue-200>On Process</span>';
      case 'Done':
        return '<span class=px-2.5 py-1 text-xs font-semibold rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200>Done</span>';
      case 'Rejected':
        return '<span class=px-2.5 py-1 text-xs font-semibold rounded-full bg-rose-50 text-rose-700 border border-rose-200>Rejected</span>';
      default:
        return '<span class=px-2.5 py-1 text-xs font-semibold rounded-full bg-slate-100 text-slate-700>' + status + '</span>';
    }
  };

  const formatRupiah = (val) => {
    const num = parseFloat(String(val).replace(/[^0-9.-]+/g, '')) || 0;
    return 'Rp ' + num.toLocaleString('id-ID');
  };

  tbody.innerHTML = pageItems.map(item => {
    let actionCol = '';

    if (currentUser.role === 'Pengurus') {
      if (item.status === 'Created') {
        actionCol = '<div class=flex items-center justify-center gap-1.5>' +
          '<button onclick=updatePengajuanStatus(' + item.rowIndex + ', \'On Process\') class=px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-semibold shadow-sm transition>Approve</button>' +
          '<button onclick=updatePengajuanStatus(' + item.rowIndex + ', \'Rejected\') class=px-2.5 py-1 bg-rose-600 hover:bg-rose-700 text-white rounded-lg text-xs font-semibold shadow-sm transition>Reject</button>' +
          '</div>';
      } else {
        actionCol = '<span class=text-xs text-slate-400 italic>Selesai Evaluasi</span>';
      }
    } else if (currentUser.role === 'Financial' || currentUser.role === 'Admin') {
      actionCol = '<div class=flex items-center justify-center>' +
        '<select onchange=updatePengajuanStatus(' + item.rowIndex + ', this.value) class=text-xs bg-slate-50 border border-slate-200 rounded-lg px-2 py-1 focus:ring-1 focus:ring-brand-500>' +
        '<option value=Created' + (item.status === 'Created' ? ' selected' : '') + '>Created</option>' +
        '<option value=On Process' + (item.status === 'On Process' ? ' selected' : '') + '>On Process</option>' +
        '<option value=Done' + (item.status === 'Done' ? ' selected' : '') + '>Done</option>' +
        '<option value=Rejected' + (item.status === 'Rejected' ? ' selected' : '') + '>Rejected</option>' +
        '</select></div>';
    } else {
      actionCol = '<span class=text-xs text-slate-400>Readonly</span>';
    }

    return '<tr class=hover:bg-slate-50/80 transition>' +
      '<td class=py-3 px-4 text-xs font-medium text-slate-500 whitespace-nowrap>' + item.timestamp + '</td>' +
      '<td class=py-3 px-4 font-semibold text-slate-800>' + item.tipeFormulir + '</td>' +
      '<td class=py-3 px-4>' +
        '<div class=font-medium text-slate-900>' + item.namaKaryawan + '</div>' +
        '<div class=text-[11px] text-slate-400>' + item.nikUnitKerja + '</div>' +
      '</td>' +
      '<td class=py-3 px-4>' +
        '<div class=font-mono text-xs text-slate-800>' + item.noRekening + '</div>' +
        '<div class=text-[11px] text-slate-400>a.n ' + item.atasNamaRekening + '</div>' +
      '</td>' +
      '<td class=py-3 px-4 font-bold text-slate-900 whitespace-nowrap>' + formatRupiah(item.totalDitransfer) + '</td>' +
      '<td class=py-3 px-4 text-slate-600>' + item.cicilan + '</td>' +
      '<td class=py-3 px-4 max-w-[200px] truncate text-slate-500 title=' + item.keterangan + '>' + item.keterangan + '</td>' +
      '<td class=py-3 px-4 text-center whitespace-nowrap>' + statusBadge(item.status) + '</td>' +
      '<td class=py-3 px-4 text-center whitespace-nowrap>' + actionCol + '</td>' +
    '</tr>';
  }).join('');

  document.getElementById('pagination-info').textContent = 'Menampilkan ' + (startIdx + 1) + '-' + Math.min(startIdx + pageSize, total) + ' dari ' + total + ' data';
  
  let pageHtml = '';
  for (let i = 1; i <= totalPages; i++) {
    const activeClass = i === currentPage ? 'bg-brand-500 text-white' : 'bg-slate-100 hover:bg-slate-200 text-slate-600';
    pageHtml += '<button onclick=goToPage(' + i + ') class=w-7 h-7 rounded-lg text-xs font-semibold transition ' + activeClass + '>' + i + '</button>';
  }
  document.getElementById('pagination-controls').innerHTML = pageHtml;
}

function goToPage(p) {
  currentPage = p;
  renderPengajuanTable();
}

async function updatePengajuanStatus(rowIndex, newStatus) {
  if (!confirm('Konfirmasi ubah status pengajuan baris ini menjadi ' + newStatus + '?')) {
    fetchPengajuanData();
    return;
  }

  try {
    const res = await fetch(API_BASE + '/api/pengajuan/status', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        rowIndex,
        newStatus,
        userRole: currentUser.role
      })
    });
    const data = await res.json();

    if (data.success) {
      showToast(data.message, 'success');
      fetchPengajuanData();
    } else {
      showToast(data.message || 'Gagal mengubah status.', 'error');
      fetchPengajuanData();
    }
  } catch (err) {
    showToast('Koneksi error: ' + err.message, 'error');
    fetchPengajuanData();
  }
}

// User Management (Admin)
async function fetchUsersData() {
  const tbody = document.getElementById('table-users-body');
  tbody.innerHTML = '<tr><td colspan=7 class=py-6 text-center text-slate-400>Memuat data pengguna...</td></tr>';

  try {
    const res = await fetch(API_BASE + '/api/users?role=Admin');
    const json = await res.json();

    if (json.success) {
      const users = json.data || [];
      if (users.length === 0) {
        tbody.innerHTML = '<tr><td colspan=7 class=py-6 text-center text-slate-400>Belum ada akun terdaftar.</td></tr>';
        return;
      }

      const statusBadge = (st) => {
        switch (st) {
          case 'Active':
            return '<span class=px-2.5 py-0.5 text-xs font-semibold rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200>Active</span>';
          case 'Pending':
            return '<span class=px-2.5 py-0.5 text-xs font-semibold rounded-full bg-amber-50 text-amber-700 border border-amber-200>Pending</span>';
          case 'Deactive':
            return '<span class=px-2.5 py-0.5 text-xs font-semibold rounded-full bg-slate-100 text-slate-600 border border-slate-200>Deactive</span>';
          case 'Rejected':
            return '<span class=px-2.5 py-0.5 text-xs font-semibold rounded-full bg-rose-50 text-rose-700 border border-rose-200>Rejected</span>';
          default:
            return '<span class=px-2.5 py-0.5 text-xs font-semibold rounded-full bg-slate-100 text-slate-700>' + st + '</span>';
        }
      };

      tbody.innerHTML = users.map(u => {
        let actionBtn = '';
        if (u.accountStatus !== 'Active') {
          actionBtn += '<button onclick=updateUserStatus(\'' + u.username + '\', \'Active\') class=px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-semibold shadow-sm transition mr-1>Set Active</button>';
        }
        if (u.accountStatus === 'Pending') {
          actionBtn += '<button onclick=updateUserStatus(\'' + u.username + '\', \'Rejected\') class=px-2.5 py-1 bg-rose-600 hover:bg-rose-700 text-white rounded-lg text-xs font-semibold shadow-sm transition mr-1>Reject</button>';
        }
        if (u.accountStatus === 'Active') {
          actionBtn += '<button onclick=updateUserStatus(\'' + u.username + '\', \'Deactive\') class=px-2.5 py-1 bg-slate-600 hover:bg-slate-700 text-white rounded-lg text-xs font-semibold shadow-sm transition>Deactivate</button>';
        }

        return '<tr class=hover:bg-slate-50/80 transition>' +
          '<td class=py-3 px-4 font-mono font-bold text-slate-900>' + u.username + '</td>' +
          '<td class=py-3 px-4 font-medium text-slate-800>' + u.namaLengkap + '</td>' +
          '<td class=py-3 px-4><span class=px-2 py-0.5 text-[11px] font-semibold rounded bg-slate-100 text-slate-700>' + u.role + '</span></td>' +
          '<td class=py-3 px-4 text-slate-600>' + (u.nik || '-') + '</td>' +
          '<td class=py-3 px-4 text-slate-600>' + (u.email || '-') + '</td>' +
          '<td class=py-3 px-4 text-center>' + statusBadge(u.accountStatus) + '</td>' +
          '<td class=py-3 px-4 text-center><div class=flex items-center justify-center>' + actionBtn + '</div></td>' +
        '</tr>';
      }).join('');
    } else {
      tbody.innerHTML = '<tr><td colspan=7 class=py-6 text-center text-rose-500>' + json.message + '</td></tr>';
    }
  } catch (err) {
    tbody.innerHTML = '<tr><td colspan=7 class=py-6 text-center text-rose-500>Error koneksi user API</td></tr>';
  }
}

async function updateUserStatus(username, status) {
  if (!confirm('Ubah status akun @' + username + ' menjadi ' + status + '?')) return;

  try {
    const res = await fetch(API_BASE + '/api/users/status', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, status })
    });
    const data = await res.json();

    if (data.success) {
      showToast(data.message, 'success');
      fetchUsersData();
    } else {
      showToast(data.message || 'Gagal mengubah status pengguna.', 'error');
    }
  } catch (err) {
    showToast('Koneksi error: ' + err.message, 'error');
  }
}