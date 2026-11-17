/**
 * ============================================================================
 * SiJaAnti PRO - Client-Side Engine (app.js)
 * Standar: /gas-pro-api, /gas-scale-turbo, /gas-pwa-github, /ui-ux-pro-max
 * ============================================================================
 * Fitur Utama:
 * 1. PWA Service Worker & Install to Home Screen Handler
 * 2. Pure REST API Fetch (Anti-CORS text/plain POST, In-flight Deduplication)
 * 3. SWR Caching & Optimistic UI (Zero-latency instant interaction)
 * 4. Zero-Latency Clash Detection Engine (<3ms)
 * 5. Interactive Matrix Plotting & Drag/Click Slot Assignment
 * 6. Teacher Availability Grid & Master Data CRUD
 * 7. SheetJS Excel Export, Word Export, Browser Print & Drive Archive
 * 8. Explicit Global Window Handlers
 * ============================================================================
 */

// ============================================================================
// 1. SAFE STORAGE WRAPPER (Anti-error di iframe/private mode)
// ============================================================================
const SafeStorage = {
  _mem: {},
  get: function(key) {
    try {
      if (typeof window !== 'undefined' && 'localStorage' in window && window.localStorage) {
        return window.localStorage.getItem(key);
      }
    } catch (e) {}
    return SafeStorage._mem[key] || null;
  },
  set: function(key, val) {
    try {
      if (typeof window !== 'undefined' && 'localStorage' in window && window.localStorage) {
        window.localStorage.setItem(key, val);
        return;
      }
    } catch (e) {}
    SafeStorage._mem[key] = val;
  },
  remove: function(key) {
    try {
      if (typeof window !== 'undefined' && 'localStorage' in window && window.localStorage) {
        window.localStorage.removeItem(key);
        return;
      }
    } catch (e) {}
    delete SafeStorage._mem[key];
  }
};

// ============================================================================
// 2. STATE STORE GLOBAL
// ============================================================================
const AppState = {
  isLoggedIn: false,
  authToken: '',
  currentUser: {
    role: 'admin',
    username: '19791104 200501 1 003',
    name: 'Bambang Sugiantoro, S.Pd.SD',
    nip: '19791104 200501 1 003',
    title: 'Tim Kurikulum (Admin)',
    menuAccess: ['dashboard', 'beban', 'matrix', 'availability', 'timeconfig', 'master', 'school', 'users', 'export']
  },
  activeView: 'dashboard',
  matrixMode: 'class', // 'class' | 'single-class' | 'teacher'
  selectedDay: 'Senin',
  selectedClassId: 'K07', // 4A
  selectedTeacherId: 'G02', // Ahmad Fauzi
  activeMasterTab: 'guru',
  selectedMasterIds: new Set(),
  users: [],
  beban: [],
  timeConfig: {
    days: {},
    classLimits: []
  },
  schoolLogoBase64: null,
  master: {
    guru: [],
    kelas: [],
    mapel: [],
    ruang: [],
    waktu: [],
    sekolah: []
  },
  jadwal: [],
  requests: [],
  auditTrail: [],
  clashes: [],
  activeSlotEdit: null,
  jtmChartInstance: null,
  loginSelectedRole: 'admin',
  deferredInstallPrompt: null,
  _pendingImport: null,
  _generatedScheduleTemp: null
};

// ============================================================================
// 3. PURE REST API CLIENT (/gas-pro-api & /gas-scale-turbo)
// ============================================================================
const _inflight = {};
let _writeEpoch = 0;

/**
 * HTTP GET ke Google Apps Script REST API
 */
async function apiGet(action, params = {}) {
  const query = new URLSearchParams({ action, ...params }).toString();
  const cacheKey = `GET_${query}`;

  if (_inflight[cacheKey]) return _inflight[cacheKey];

  const promise = (async () => {
    try {
      const res = await fetch(`${APP_CONFIG.GAS_URL}?${query}`, {
        method: 'GET',
        mode: 'cors'
      });
      const json = await res.json();
      return json;
    } catch (err) {
      console.warn('API GET Error:', err);
      return { success: false, message: err.message };
    } finally {
      delete _inflight[cacheKey];
    }
  })();

  _inflight[cacheKey] = promise;
  return promise;
}

/**
 * HTTP POST ke Google Apps Script REST API
 * KRUSIAL CORS: Wajib Content-Type text/plain;charset=utf-8 agar browser tidak trigger preflight OPTIONS!
 */
async function apiPost(action, data = {}) {
  _writeEpoch++; // Batalkan caching in-flight saat operasi tulis

  const token = AppState.authToken || SafeStorage.get(APP_CONFIG.STORAGE_KEY_AUTH + '_TOKEN') || '';
  const payload = {
    action: action,
    data: data,
    token: token
  };

  try {
    const res = await fetch(APP_CONFIG.GAS_URL, {
      method: 'POST',
      mode: 'cors',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8' // WAJIB text/plain untuk GAS CORS
      },
      body: JSON.stringify(payload)
    });
    const json = await res.json();
    return json;
  } catch (err) {
    console.error('API POST Error:', err);
    return { success: false, message: 'Gagal menghubungi server REST API: ' + err.message };
  }
}

// ============================================================================
// 4. PWA REGISTRATION & INSTALL TO HOME SCREEN (/gas-pwa-github)
// ============================================================================
function initPWA() {
  // 1. Daftarkan Service Worker
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./sw.js')
        .then(reg => console.log('✅ PWA Service Worker aktif:', reg.scope))
        .catch(err => console.warn('⚠️ PWA Service Worker gagal didaftarkan:', err));
    });
  }

  // 2. Tangkap event 'beforeinstallprompt' untuk Tombol Install PWA di Header / Sidebar
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    AppState.deferredInstallPrompt = e;

    const btnHeader = document.getElementById('btn-install-pwa-header');
    const btnSidebar = document.getElementById('btn-install-pwa-sidebar');
    const banner = document.getElementById('pwa-install-banner');

    if (btnHeader) btnHeader.style.display = 'inline-flex';
    if (btnSidebar) btnSidebar.style.display = 'flex';
    if (banner) banner.style.display = 'flex';
  });

  window.addEventListener('appinstalled', () => {
    console.log('🎉 SiJaAnti PRO PWA berhasil dipasang ke Layar Utama!');
    AppState.deferredInstallPrompt = null;
    hideInstallUI();
    showToast('Aplikasi berhasil dipasang di perangkat Anda!', 'success');
  });
}

async function triggerPWAInstall() {
  if (!AppState.deferredInstallPrompt) {
    showToast('Buka menu browser Anda lalu pilih "Tambahkan ke Layar Utama" (Add to Home Screen)', 'info');
    return;
  }

  AppState.deferredInstallPrompt.prompt();
  const { outcome } = await AppState.deferredInstallPrompt.userChoice;
  console.log('PWA Install choice:', outcome);

  if (outcome === 'accepted') {
    hideInstallUI();
  }
  AppState.deferredInstallPrompt = null;
}

function hideInstallUI() {
  const btnHeader = document.getElementById('btn-install-pwa-header');
  const btnSidebar = document.getElementById('btn-install-pwa-sidebar');
  const banner = document.getElementById('pwa-install-banner');
  if (btnHeader) btnHeader.style.display = 'none';
  if (btnSidebar) btnSidebar.style.display = 'none';
  if (banner) banner.style.display = 'none';
}

// ============================================================================
// 5. BOOT APPLICATION & SWR SYNC
// ============================================================================
document.addEventListener('DOMContentLoaded', bootApplication);

function bootApplication() {
  registerGlobalHandlers();
  initPWA();
  loadFromLocalStorage();
  checkAuthSession();
  loadSchoolIdentity(); // Sinkronkan identitas sekolah & logo ke seluruh elemen tampilan awal
  populateLoginTeacherDropdown();
  calculateClashes();

  if (AppState.isLoggedIn) {
    showAppLayout();
    setupRoleUI();
    renderCurrentView();
  } else {
    showLoginLayout();
  }

  // Tarik data terbaru dari backend REST API di background (SWR)
  fetchServerData();
}

function checkAuthSession() {
  const sessionRaw = SafeStorage.get(APP_CONFIG.STORAGE_KEY_AUTH);
  const token = SafeStorage.get(APP_CONFIG.STORAGE_KEY_AUTH + '_TOKEN');
  if (sessionRaw) {
    try {
      const sess = JSON.parse(sessionRaw);
      if (sess && sess.isLoggedIn && sess.user) {
        AppState.isLoggedIn = true;
        AppState.currentUser = sess.user;
        AppState.authToken = token || '';
        return;
      }
    } catch (e) {}
  }
  AppState.isLoggedIn = false;
}

function showLoginLayout() {
  const loginEl = document.getElementById('view-login');
  const appLayoutEl = document.getElementById('app-layout');
  if (loginEl) {
    loginEl.style.display = 'flex';
    loginEl.classList.remove('hidden');
  }
  if (appLayoutEl) {
    appLayoutEl.style.display = 'none';
    appLayoutEl.classList.add('hidden');
  }
}

function showAppLayout() {
  const loginEl = document.getElementById('view-login');
  const appLayoutEl = document.getElementById('app-layout');
  if (loginEl) {
    loginEl.style.display = 'none';
    loginEl.classList.add('hidden');
  }
  if (appLayoutEl) {
    appLayoutEl.style.display = 'block';
    appLayoutEl.classList.remove('hidden');
  }
}

function selectLoginRole(roleKey) {
  AppState.loginSelectedRole = roleKey;
  
  ['admin', 'kepsek', 'guru'].forEach(r => {
    const tab = document.getElementById(`login-tab-${r}`);
    if (tab) {
      if (r === roleKey) {
        tab.className = 'flex-1 py-2 rounded-xl bg-white text-primary shadow-sm transition-all text-center font-bold';
      } else {
        tab.className = 'flex-1 py-2 rounded-xl text-slate-600 hover:text-slate-900 transition-all text-center font-medium';
      }
    }
  });

  const teacherWrap = document.getElementById('login-teacher-wrap');
  const usernameInput = document.getElementById('login-username');
  const usernameLabel = document.getElementById('login-username-label');

  if (roleKey === 'guru') {
    if (teacherWrap) teacherWrap.classList.remove('hidden');
    if (usernameLabel) usernameLabel.textContent = 'NIP Guru:';
    const teacherSel = document.getElementById('login-teacher-select');
    if (teacherSel && teacherSel.value) updateLoginTeacherCredentials(teacherSel.value);
  } else if (roleKey === 'kepsek') {
    if (teacherWrap) teacherWrap.classList.add('hidden');
    if (usernameLabel) usernameLabel.textContent = 'NIP Kepala Sekolah:';
    if (usernameInput) usernameInput.value = '19740512 199803 2 004';
  } else {
    if (teacherWrap) teacherWrap.classList.add('hidden');
    if (usernameLabel) usernameLabel.textContent = 'NIP / Username Kurikulum:';
    if (usernameInput) usernameInput.value = '19791104 200501 1 003';
  }
}

function populateLoginTeacherDropdown() {
  const sel = document.getElementById('login-teacher-select');
  if (!sel) return;
  const teachers = AppState.master.guru || [];
  sel.innerHTML = teachers.map(g => `
    <option value="${g.id}" ${g.id === 'G02' ? 'selected' : ''}>${g.nama} (${g.mapel_utama})</option>
  `).join('');
  if (sel.value) updateLoginTeacherCredentials(sel.value);
}

function updateLoginTeacherCredentials(teacherId) {
  const g = (AppState.master.guru || []).find(it => it.id === teacherId);
  const usernameInput = document.getElementById('login-username');
  if (g && usernameInput) usernameInput.value = g.nip || g.id;
}

async function handleLoginSubmit(e) {
  if (e && e.preventDefault) e.preventDefault();
  const role = AppState.loginSelectedRole;
  const username = document.getElementById('login-username')?.value || '';
  const teacherId = document.getElementById('login-teacher-select')?.value || 'G02';

  const btn = document.getElementById('btn-login-submit');
  if (btn) btn.classList.add('opacity-75');

  showToast('Memverifikasi sesi dengan REST API...', 'info');

  // Panggil endpoint authLogin backend
  const res = await apiPost('authLogin', { role, username, teacherId });
  if (btn) btn.classList.remove('opacity-75');

  if (res && res.success && res.user) {
    executeLogin(res.user, res.token);
  } else {
    // Fallback offline login untuk kenyamanan pengujian
    fallbackLocalLogin(role, teacherId);
  }
}

function quickLogin(roleKey) {
  selectLoginRole(roleKey);
  fallbackLocalLogin(roleKey, 'G02');
}

function fallbackLocalLogin(roleKey, teacherId = 'G02') {
  let userObj = {};
  if (roleKey === 'admin') {
    userObj = { role: 'admin', username: '19791104 200501 1 003', name: 'Bambang Sugiantoro, S.Pd.SD', nip: '19791104 200501 1 003', title: 'Tim Kurikulum (Admin)' };
  } else if (roleKey === 'kepsek') {
    userObj = { role: 'kepsek', username: '19740512 199803 2 004', name: 'Dra. Sri Mulyani, M.Pd', nip: '19740512 199803 2 004', title: 'Kepala Sekolah' };
  } else {
    const g = (AppState.master.guru || []).find(it => it.id === teacherId) || { id: 'G02', nama: 'Ahmad Fauzi', gelar: 'S.Pd.SD', nip: '19820315 200604 1 008', mapel_utama: 'Guru Kelas 4' };
    userObj = { role: 'guru', teacherId: g.id, username: g.nip, name: g.nama + (g.gelar ? ', ' + g.gelar : ''), nip: g.nip, title: g.mapel_utama || 'Guru Kelas 4A' };
    AppState.selectedTeacherId = g.id;
  }
  executeLogin(userObj, 'OFFLINE_TOKEN_' + Date.now());
}

function executeLogin(userObj, token) {
  AppState.isLoggedIn = true;
  AppState.currentUser = userObj;
  AppState.authToken = token || '';

  SafeStorage.set(APP_CONFIG.STORAGE_KEY_AUTH, JSON.stringify({ isLoggedIn: true, user: userObj, timestamp: Date.now() }));
  if (token) SafeStorage.set(APP_CONFIG.STORAGE_KEY_AUTH + '_TOKEN', token);

  showAppLayout();
  setupRoleUI();

  if (userObj.role === 'guru') {
    navigateView('matrix');
  } else {
    navigateView('dashboard');
  }

  showToast(`Selamat datang, ${userObj.name}`, 'success');
}

function logoutUser() {
  AppState.isLoggedIn = false;
  AppState.authToken = '';
  SafeStorage.remove(APP_CONFIG.STORAGE_KEY_AUTH);
  SafeStorage.remove(APP_CONFIG.STORAGE_KEY_AUTH + '_TOKEN');
  showLoginLayout();
  showToast('Anda telah keluar dari sesi SiJaAnti.', 'info');
}

// ============================================================================
// 6. SPA ROUTING ENGINE (Murni Anti-Reload)
// ============================================================================
function navigateView(viewName) {
  if (!AppState.isLoggedIn) {
    showLoginLayout();
    return;
  }

  AppState.activeView = viewName;
  showAppLayout();

  const views = ['dashboard', 'beban', 'matrix', 'availability', 'timeconfig', 'master', 'school', 'users', 'export'];
  views.forEach(v => {
    const el = document.getElementById(`view-${v}`);
    if (el) {
      el.style.display = 'none';
      el.classList.add('hidden');
    }
    const navBtn = document.getElementById(`nav-btn-${v}`);
    if (navBtn) {
      navBtn.classList.remove('bg-white/20', 'text-white', 'shadow-inner', 'font-bold');
      navBtn.classList.add('text-white/90', 'font-semibold');
    }
  });

  const targetView = document.getElementById(`view-${viewName}`);
  if (targetView) {
    targetView.style.display = 'block';
    targetView.classList.remove('hidden');
  }

  const activeNavBtn = document.getElementById(`nav-btn-${viewName}`);
  if (activeNavBtn) {
    activeNavBtn.classList.remove('text-white/90', 'font-semibold');
    activeNavBtn.classList.add('bg-white/20', 'text-white', 'shadow-inner', 'font-bold');
  }

  closeMobileSidebar();

  try {
    renderCurrentView();
  } catch (err) {
    console.error('Render view error:', err);
  }
}

function renderCurrentView() {
  switch (AppState.activeView) {
    case 'dashboard': renderDashboardView(); break;
    case 'beban': renderBebanView(); break;
    case 'matrix': renderMatrixView(); break;
    case 'availability': renderAvailabilityView(); break;
    case 'timeconfig': renderTimeConfigView(); break;
    case 'master': renderMasterView(); break;
    case 'school': renderSchoolView(); break;
    case 'users': renderUsersView(); break;
    case 'export': renderExportView(); break;
  }
}

function switchRole(roleKey) {
  quickLogin(roleKey);
}

function setupRoleUI() {
  const nameEl = document.getElementById('current-user-name');
  const roleEl = document.getElementById('current-user-role');
  const avatarEl = document.getElementById('user-avatar-badge');
  const roleSel = document.getElementById('role-select');
  
  const user = AppState.currentUser || {};
  if (nameEl) nameEl.textContent = user.name || 'Pengguna';
  if (roleEl) roleEl.textContent = user.title || 'Kurikulum';
  if (roleSel && user.role) roleSel.value = user.role;
  if (avatarEl) {
    const rawName = user.name || 'Guru';
    avatarEl.textContent = rawName.split(' ').filter(Boolean).map(n => n[0]).slice(0, 2).join('') || 'U';
  }

  const approveBtn = document.getElementById('btn-approve-kbm');
  if (approveBtn) approveBtn.style.display = user.role === 'kepsek' ? 'flex' : 'none';

  // REVISI 2: Filter Hak Akses Menu Navigasi Berdasarkan currentUser.menuAccess
  const menuAccess = user.menuAccess || '*';
  const allNavViews = ['dashboard', 'beban', 'matrix', 'availability', 'timeconfig', 'master', 'export', 'school', 'users'];
  allNavViews.forEach(v => {
    const btn = document.getElementById(`nav-btn-${v}`);
    if (btn) {
      let allowed = false;
      if (menuAccess === '*' || user.role === 'admin') {
        allowed = true;
      } else if (Array.isArray(menuAccess)) {
        allowed = menuAccess.includes(v);
      } else if (typeof menuAccess === 'string') {
        allowed = menuAccess.split(',').map(s => s.trim()).includes(v);
      }
      btn.style.display = allowed ? 'flex' : 'none';
    }
  });
}

// ============================================================================
// 7. ZERO-LATENCY CLASH DETECTION ENGINE (<3ms)
// ============================================================================
function calculateClashes() {
  const clashes = [];
  const jadwal = AppState.jadwal || [];
  const requests = AppState.requests || [];
  const masterGuru = AppState.master.guru || [];
  const masterKelas = AppState.master.kelas || [];
  const masterRuang = AppState.master.ruang || [];

  const guruMap = {};
  masterGuru.forEach(g => { guruMap[g.id] = g.nama + (g.gelar ? ', ' + g.gelar : ''); });

  const kelasMap = {};
  masterKelas.forEach(k => { kelasMap[k.id] = k.nama_rombel || k.kode_kelas; });

  const ruangMap = {};
  masterRuang.forEach(r => { ruangMap[r.id] = r.nama; });

  // 1. Bentrok Guru: Mengajar di 2 kelas berbeda pada Hari & JP sama
  const teacherSlotMap = {};
  jadwal.forEach(item => {
    if (!item.guru_id || item.guru_id === '-' || !item.hari || !item.jp) return;
    const key = `${item.hari}_${item.jp}_${item.guru_id}`;
    if (!teacherSlotMap[key]) teacherSlotMap[key] = [];
    teacherSlotMap[key].push(item);
  });

  Object.keys(teacherSlotMap).forEach(key => {
    const items = teacherSlotMap[key];
    if (items.length > 1) {
      const first = items[0];
      const kelasNames = items.map(it => kelasMap[it.kelas_id] || it.kelas_id).join(' dan ');
      const guruName = guruMap[first.guru_id] || first.guru_id;

      items.forEach(it => {
        clashes.push({
          type: 'GURU_CLASH',
          jadwalId: it.id,
          hari: it.hari,
          jp: it.jp,
          guruId: it.guru_id,
          kelasId: it.kelas_id,
          ruangId: it.ruang_id,
          title: `Bentrok Guru: ${guruName}`,
          message: `${guruName} dijadwalkan mengajar sekaligus di ${kelasNames} pada ${it.hari} JP ${it.jp}.`
        });
      });
    }
  });

  // 2. Bentrok Ketersediaan Guru (Request Berhalangan)
  requests.forEach(req => {
    if (req.status === 'Disetujui' || !req.status) {
      const conflictingJadwal = jadwal.filter(it => 
        it.guru_id === req.guru_id && 
        it.hari === req.hari && 
        parseInt(it.jp) === parseInt(req.jp)
      );

      conflictingJadwal.forEach(it => {
        const guruName = guruMap[it.guru_id] || it.guru_id;
        const kelasName = kelasMap[it.kelas_id] || it.kelas_id;
        clashes.push({
          type: 'AVAILABILITY_CLASH',
          jadwalId: it.id,
          hari: it.hari,
          jp: it.jp,
          guruId: it.guru_id,
          kelasId: it.kelas_id,
          title: `Bentrok Ketersediaan: ${guruName}`,
          message: `${guruName} berhalangan mengajar di ${kelasName} pada ${it.hari} JP ${it.jp} (${req.alasan || 'Permohonan Guru'}).`
        });
      });
    }
  });

  // 3. Bentrok Ruangan Khusus
  const roomSlotMap = {};
  jadwal.forEach(item => {
    if (!item.ruang_id || !item.hari || !item.jp) return;
    if (['R_13', 'R_14', 'R_15'].includes(item.ruang_id)) {
      const key = `${item.hari}_${item.jp}_${item.ruang_id}`;
      if (!roomSlotMap[key]) roomSlotMap[key] = [];
      roomSlotMap[key].push(item);
    }
  });

  Object.keys(roomSlotMap).forEach(key => {
    const items = roomSlotMap[key];
    if (items.length > 1) {
      const first = items[0];
      const rName = ruangMap[first.ruang_id] || first.ruang_id;
      const kNames = items.map(it => kelasMap[it.kelas_id] || it.kelas_id).join(' & ');
      items.forEach(it => {
        clashes.push({
          type: 'ROOM_CLASH',
          jadwalId: it.id,
          hari: it.hari,
          jp: it.jp,
          ruangId: it.ruang_id,
          kelasId: it.kelas_id,
          title: `Bentrok Ruangan: ${rName}`,
          message: `${rName} digunakan bersamaan oleh ${kNames} pada ${it.hari} JP ${it.jp}.`
        });
      });
    }
  });

  AppState.clashes = clashes;
  updateClashBadgesUI();
  return clashes;
}

function updateClashBadgesUI() {
  const count = AppState.clashes.length;
  
  const headerClash = document.getElementById('header-clash-indicator');
  const headerCount = document.getElementById('header-clash-count');
  if (headerClash && headerCount) {
    if (count > 0) {
      headerClash.classList.remove('hidden');
      headerClash.classList.add('flex');
      headerCount.textContent = `${count} Bentrok`;
    } else {
      headerClash.classList.add('hidden');
      headerClash.classList.remove('flex');
    }
  }

  const sidebarBadge = document.getElementById('sidebar-clash-badge');
  if (sidebarBadge) {
    if (count > 0) {
      sidebarBadge.classList.remove('hidden');
      sidebarBadge.textContent = count;
    } else {
      sidebarBadge.classList.add('hidden');
    }
  }

  const dashBanner = document.getElementById('dashboard-clash-banner');
  const dashCount = document.getElementById('clash-banner-count');
  const kpiCount = document.getElementById('kpi-clash-count');
  const kpiStatus = document.getElementById('kpi-clash-status');
  const kpiIconWrap = document.getElementById('kpi-clash-icon-wrap');

  if (dashBanner && dashCount) {
    if (count > 0) {
      dashBanner.classList.remove('hidden');
      dashCount.textContent = `${count} Konflik`;
      if (kpiCount) { kpiCount.textContent = count; kpiCount.className = 'text-3xl font-black text-rose-600'; }
      if (kpiStatus) { kpiStatus.textContent = 'Ada Bentrok!'; kpiStatus.className = 'text-xs text-rose-700 font-bold'; }
      if (kpiIconWrap) kpiIconWrap.className = 'w-11 h-11 rounded-xl bg-rose-50 text-rose-600 flex items-center justify-center';
    } else {
      dashBanner.classList.add('hidden');
      if (kpiCount) { kpiCount.textContent = '0'; kpiCount.className = 'text-3xl font-black text-emerald-600'; }
      if (kpiStatus) { kpiStatus.textContent = 'Jadwal Aman'; kpiStatus.className = 'text-xs text-emerald-700 font-semibold'; }
      if (kpiIconWrap) kpiIconWrap.className = 'w-11 h-11 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center';
    }
  }

  const matrixAlert = document.getElementById('matrix-clash-alert-box');
  const matrixAlertText = document.getElementById('matrix-clash-alert-text');
  if (matrixAlert && matrixAlertText) {
    if (count > 0) {
      matrixAlert.classList.remove('hidden');
      matrixAlertText.textContent = `Terdeteksi ${count} bentrok jadwal mengajar! Segera tinjau & perbaiki.`;
    } else {
      matrixAlert.classList.add('hidden');
    }
  }
}

// ============================================================================
// 8. RENDER VIEW: EXECUTIVE DASHBOARD
// ============================================================================
function renderDashboardView() {
  const masterGuru = AppState.master.guru || [];
  const masterKelas = AppState.master.kelas || [];
  const jadwal = AppState.jadwal || [];

  const kpiGuru = document.getElementById('kpi-total-guru');
  const kpiRombel = document.getElementById('kpi-total-rombel');
  const kpiPlotPercent = document.getElementById('kpi-plotting-percent');
  const kpiPlotJP = document.getElementById('kpi-plotting-jp');
  const kpiProgBar = document.getElementById('kpi-progress-bar');

  if (kpiGuru) kpiGuru.textContent = masterGuru.length || 24;
  if (kpiRombel) kpiRombel.textContent = masterKelas.length || 12;

  const totalSlotsCapacity = (masterKelas.length || 12) * 28;
  const plottedSlotsCount = jadwal.length;
  const percent = Math.min(100, Math.round((plottedSlotsCount / (totalSlotsCapacity || 1)) * 100));

  if (kpiPlotPercent) kpiPlotPercent.textContent = `${percent}%`;
  if (kpiPlotJP) kpiPlotJP.textContent = `${plottedSlotsCount} / ${totalSlotsCapacity} JP`;
  if (kpiProgBar) kpiProgBar.style.width = `${percent}%`;

  renderJTMChart();
  renderRombelReadiness();
  renderAuditTrail();
}

function renderJTMChart() {
  const container = document.getElementById('chart-parent-container');
  const canvas = document.getElementById('jtmChart');
  if (!container) return;

  const masterGuru = AppState.master.guru || [];
  const jadwal = AppState.jadwal || [];

  const teacherJtmMap = {};
  masterGuru.forEach(g => { teacherJtmMap[g.id] = 0; });
  jadwal.forEach(it => {
    if (it.guru_id && teacherJtmMap[it.guru_id] !== undefined) {
      teacherJtmMap[it.guru_id] += 1;
    }
  });

  const sampleTeachers = masterGuru.length > 0 ? masterGuru.slice(1, 13) : [];
  const labels = sampleTeachers.map(g => (g && g.nama ? g.nama.split(' ')[0] : 'Guru'));
  const dataJTM = sampleTeachers.map(g => (g && g.id && teacherJtmMap[g.id] ? teacherJtmMap[g.id] : 0));

  if (typeof Chart !== 'undefined' && canvas) {
    try {
      if (AppState.jtmChartInstance) AppState.jtmChartInstance.destroy();
      const ctx = canvas.getContext('2d');
      AppState.jtmChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [{
            label: 'JTM Terplot (JP/Minggu)',
            data: dataJTM,
            backgroundColor: dataJTM.map(val => val >= 24 ? '#10B981' : val >= 20 ? '#2563EB' : '#F59E0B'),
            borderRadius: 6,
            barPercentage: 0.6
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            y: { beginAtZero: true, suggestedMax: 28, ticks: { stepSize: 4 } },
            x: { grid: { display: false } }
          }
        }
      });
      return;
    } catch (e) {
      console.warn('ChartJS error fallback:', e);
    }
  }

  // Fallback CSS Bars
  let html = `<div class="w-full h-full flex items-end justify-between gap-2 pt-6 pb-2 px-2">`;
  sampleTeachers.forEach((g, idx) => {
    const jtm = dataJTM[idx] || 0;
    const heightPercent = Math.min(100, Math.round((jtm / 28) * 100));
    const barColor = jtm >= 24 ? 'bg-emerald-500' : jtm >= 20 ? 'bg-indigo-600' : 'bg-amber-500';
    html += `
      <div class="flex-1 flex flex-col items-center gap-1 group">
        <span class="text-[10px] font-bold text-slate-700">${jtm}JP</span>
        <div class="w-full bg-slate-100 rounded-t-lg h-44 flex items-end">
          <div class="w-full ${barColor} rounded-t-lg transition-all duration-300" style="height: ${heightPercent}%"></div>
        </div>
        <span class="text-[10px] font-semibold text-slate-600 truncate max-w-[45px]">${labels[idx]}</span>
      </div>
    `;
  });
  html += `</div>`;
  container.innerHTML = html;
}

function renderRombelReadiness() {
  const container = document.getElementById('rombel-readiness-list');
  if (!container) return;

  const masterKelas = AppState.master.kelas || [];
  const jadwal = AppState.jadwal || [];

  let html = '';
  masterKelas.forEach(k => {
    const classJadwalCount = jadwal.filter(it => it.kelas_id === k.id).length;
    const targetSlots = 28;
    const progress = Math.min(100, Math.round((classJadwalCount / targetSlots) * 100));

    html += `
      <div class="p-2.5 rounded-xl bg-slate-50 border border-border-subtle flex items-center justify-between gap-3 text-xs">
        <div class="flex items-center gap-2">
          <div class="w-7 h-7 rounded-lg bg-blue-100 text-blue-800 font-extrabold flex items-center justify-center text-[11px]">
            ${k.kode_kelas}
          </div>
          <div>
            <div class="font-bold text-text-primary">${k.nama_rombel}</div>
            <div class="text-[10px] text-text-secondary">${classJadwalCount} dari ${targetSlots} JP terisi</div>
          </div>
        </div>
        <div class="flex items-center gap-2">
          <span class="font-bold text-text-primary">${progress}%</span>
          <span class="material-symbols-outlined text-[16px] ${progress >= 100 ? 'text-emerald-500' : 'text-amber-500'}">
            ${progress >= 100 ? 'check_circle' : 'pending'}
          </span>
        </div>
      </div>
    `;
  });
  container.innerHTML = html;
}

function renderAuditTrail() {
  const tbody = document.getElementById('audit-trail-tbody');
  if (!tbody) return;

  const logs = AppState.auditTrail && AppState.auditTrail.length > 0 ? AppState.auditTrail : [
    { timestamp: '15 Juli 2024 10:15', user: 'Tim Kurikulum', aksi: 'Plotting Jadwal', keterangan: 'Menyusun jadwal KBM Kelas 4A s.d 6B.' },
    { timestamp: '15 Juli 2024 09:30', user: 'Joko Widodo, S.Pd.I', aksi: 'Request Ketersediaan', keterangan: 'Mengajukan libur KBM Kamis jam 5-6 (sekolah mitra).' }
  ];

  tbody.innerHTML = logs.slice(0, 5).map(l => `
    <tr class="hover:bg-slate-50 transition-colors">
      <td class="py-2.5 px-3 font-medium text-text-secondary whitespace-nowrap">${l.timestamp}</td>
      <td class="py-2.5 px-3 font-bold text-text-primary whitespace-nowrap">${l.user}</td>
      <td class="py-2.5 px-3 whitespace-nowrap">
        <span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-blue-100 text-blue-800">${l.aksi}</span>
      </td>
      <td class="py-2.5 px-3 text-text-secondary">${l.keterangan}</td>
    </tr>
  `).join('');
}

// ============================================================================
// 9. RENDER VIEW: SCHEDULING MATRIX
// ============================================================================
function setMatrixMode(mode) {
  AppState.matrixMode = mode;
  
  ['class', 'single', 'teacher'].forEach(m => {
    const btn = document.getElementById(`btn-mode-${m}`);
    if (btn) {
      if ((mode === 'single-class' && m === 'single') || mode === m) {
        btn.className = 'px-3 py-1.5 rounded-lg bg-white text-primary shadow-sm transition-all font-bold';
      } else {
        btn.className = 'px-3 py-1.5 rounded-lg text-text-secondary hover:text-text-primary transition-all font-semibold';
      }
    }
  });

  const daysWrap = document.getElementById('filter-days-container');
  const classWrap = document.getElementById('class-select-wrap');
  const teacherWrap = document.getElementById('teacher-select-wrap');

  if (mode === 'class') {
    if (daysWrap) daysWrap.style.display = 'flex';
    if (classWrap) classWrap.style.display = 'none';
    if (teacherWrap) teacherWrap.style.display = 'none';
  } else if (mode === 'single-class') {
    if (daysWrap) daysWrap.style.display = 'none';
    if (classWrap) classWrap.style.display = 'block';
    if (teacherWrap) teacherWrap.style.display = 'none';
  } else {
    if (daysWrap) daysWrap.style.display = 'none';
    if (classWrap) classWrap.style.display = 'none';
    if (teacherWrap) teacherWrap.style.display = 'block';
  }

  renderMatrixTable();
}

function setSelectedDay(dayName) {
  AppState.selectedDay = dayName;
  document.querySelectorAll('.matrix-day-btn').forEach(btn => {
    if (btn.getAttribute('data-day') === dayName) {
      btn.className = 'matrix-day-btn px-3 py-1.5 rounded-lg text-xs font-bold bg-primary text-white transition-colors';
    } else {
      btn.className = 'matrix-day-btn px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-100 text-text-secondary hover:bg-slate-200 transition-colors';
    }
  });
  renderMatrixTable();
}

function renderMatrixView() {
  populateMatrixSelectors();
  renderMatrixTable();
}

function populateMatrixSelectors() {
  const classSel = document.getElementById('matrix-class-select');
  if (classSel && classSel.options.length <= 1) {
    classSel.innerHTML = (AppState.master.kelas || []).map(k => `
      <option value="${k.id}" ${k.id === AppState.selectedClassId ? 'selected' : ''}>${k.nama_rombel}</option>
    `).join('');
  }

  const teacherSel = document.getElementById('matrix-teacher-select');
  if (teacherSel && teacherSel.options.length <= 1) {
    teacherSel.innerHTML = (AppState.master.guru || []).map(g => `
      <option value="${g.id}" ${g.id === AppState.selectedTeacherId ? 'selected' : ''}>${g.nama} (${g.mapel_utama})</option>
    `).join('');
  }
}

function renderMatrixTable() {
  const thead = document.getElementById('matrix-thead');
  const tbody = document.getElementById('matrix-tbody');
  if (!thead || !tbody) return;

  const mode = AppState.matrixMode;
  const jadwal = AppState.jadwal || [];
  const masterKelas = AppState.master.kelas || [];
  const masterWaktu = AppState.master.waktu || [];
  const masterMapel = AppState.master.mapel || [];
  const masterGuru = AppState.master.guru || [];

  const mapelMap = {}; masterMapel.forEach(m => { mapelMap[m.id] = m; });
  const guruMap = {}; masterGuru.forEach(g => { guruMap[g.id] = g; });
  const clashJadwalIds = new Set(AppState.clashes.map(c => c.jadwalId));

  if (mode === 'class') {
    const dayWaktu = masterWaktu.filter(w => w.hari === AppState.selectedDay);
    
    let thHtml = `<tr><th class="p-3 sticky-time-col border-r border-border-subtle w-36">Waktu & JP</th>`;
    masterKelas.forEach(k => {
      const hasClash = AppState.clashes.some(c => c.hari === AppState.selectedDay && c.kelasId === k.id);
      thHtml += `
        <th class="p-3 text-center border-r border-border-subtle min-w-[140px]">
          <div class="flex items-center justify-center gap-1">
            <span>${k.nama_rombel}</span>
            ${hasClash ? '<span class="text-rose-500 font-bold" title="Ada bentrok di kelas ini">⚠️</span>' : ''}
          </div>
        </th>
      `;
    });
    thHtml += `</tr>`;
    thead.innerHTML = thHtml;

    let tbodyHtml = '';
    dayWaktu.forEach(w => {
      tbodyHtml += `<tr>`;
      tbodyHtml += `
        <td class="p-2.5 sticky-time-col border-r border-border-subtle bg-slate-50 font-semibold text-text-primary">
          <div class="text-xs font-bold">${w.nama_slot}</div>
          <div class="text-[10px] text-text-secondary">${w.jam_mulai} - ${w.jam_selesai}</div>
        </td>
      `;

      if (w.is_break) {
        tbodyHtml += `
          <td colspan="${masterKelas.length}" class="p-3 text-center slot-striped border-r border-border-subtle text-text-muted font-bold tracking-wider uppercase text-[11px]">
            <div class="flex items-center justify-center gap-1.5 text-slate-500">
              <span class="material-symbols-outlined text-[16px]">lock</span>
              <span>${w.nama_slot} (${w.jam_mulai} - ${w.jam_selesai})</span>
            </div>
          </td>
        `;
      } else {
        masterKelas.forEach(k => {
          const item = jadwal.find(it => it.hari === AppState.selectedDay && parseInt(it.jp) === parseInt(w.jp) && it.kelas_id === k.id);
          const hasClash = item && clashJadwalIds.has(item.id);
          tbodyHtml += `
            <td class="p-1.5 border-r border-border-subtle align-top">
              ${renderMatrixCell(item, AppState.selectedDay, w.jp, k.id, hasClash, mapelMap, guruMap)}
            </td>
          `;
        });
      }
      tbodyHtml += `</tr>`;
    });
    tbody.innerHTML = tbodyHtml;

  } else if (mode === 'single-class') {
    const selectedClass = masterKelas.find(k => k.id === AppState.selectedClassId) || masterKelas[0];
    const days = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];

    let thHtml = `<tr><th class="p-3 sticky-time-col border-r border-border-subtle w-32">Jam Pelajaran</th>`;
    days.forEach(d => { thHtml += `<th class="p-3 text-center border-r border-border-subtle min-w-[150px]">${d}</th>`; });
    thHtml += `</tr>`;
    thead.innerHTML = thHtml;

    let tbodyHtml = '';
    for (let jp = 1; jp <= 8; jp++) {
      tbodyHtml += `
        <tr>
          <td class="p-2.5 sticky-time-col border-r border-border-subtle bg-slate-50 font-bold text-text-primary text-center">
            <div class="text-xs">JP ${jp}</div>
          </td>
      `;
      days.forEach(d => {
        const item = jadwal.find(it => it.hari === d && parseInt(it.jp) === jp && it.kelas_id === selectedClass.id);
        const hasClash = item && clashJadwalIds.has(item.id);
        tbodyHtml += `
          <td class="p-1.5 border-r border-border-subtle align-top">
            ${renderMatrixCell(item, d, jp, selectedClass.id, hasClash, mapelMap, guruMap)}
          </td>
        `;
      });
      tbodyHtml += `</tr>`;
    }
    tbody.innerHTML = tbodyHtml;

  } else {
    // Mode Guru
    const teacherId = AppState.selectedTeacherId;
    const days = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];

    let thHtml = `<tr><th class="p-3 sticky-time-col border-r border-border-subtle w-32">Jam Pelajaran</th>`;
    days.forEach(d => { thHtml += `<th class="p-3 text-center border-r border-border-subtle min-w-[150px]">${d}</th>`; });
    thHtml += `</tr>`;
    thead.innerHTML = thHtml;

    let tbodyHtml = '';
    for (let jp = 1; jp <= 8; jp++) {
      tbodyHtml += `
        <tr>
          <td class="p-2.5 sticky-time-col border-r border-border-subtle bg-slate-50 font-bold text-text-primary text-center">
            <div class="text-xs">JP ${jp}</div>
          </td>
      `;
      days.forEach(d => {
        const item = jadwal.find(it => it.hari === d && parseInt(it.jp) === jp && it.guru_id === teacherId);
        const hasClash = item && clashJadwalIds.has(item.id);
        tbodyHtml += `
          <td class="p-1.5 border-r border-border-subtle align-top">
            ${renderMatrixCell(item, d, jp, item ? item.kelas_id : '', hasClash, mapelMap, guruMap)}
          </td>
        `;
      });
      tbodyHtml += `</tr>`;
    }
    tbody.innerHTML = tbodyHtml;
  }
}

function renderMatrixCell(item, hari, jp, kelasId, hasClash, mapelMap, guruMap) {
  if (!item || !item.mapel_id) {
    return `
      <div onclick="openPlotModal('${hari}', ${jp}, '${kelasId}')" class="h-20 border border-dashed border-slate-300 rounded-xl p-2 bg-white hover:bg-blue-50/50 hover:border-primary/50 transition-all cursor-pointer flex flex-col items-center justify-center text-text-muted group">
        <span class="material-symbols-outlined text-[18px] group-hover:text-primary transition-colors">add</span>
        <span class="text-[10px] font-medium group-hover:text-primary transition-colors">Isi Slot</span>
      </div>
    `;
  }

  const mapel = mapelMap[item.mapel_id] || { nama: item.mapel_id, kode: item.mapel_id };
  const guru = guruMap[item.guru_id] || { nama: item.guru_id || 'Belum Ditentukan' };

  if (hasClash) {
    return `
      <div onclick="openPlotModal('${hari}', ${jp}, '${kelasId}', '${item.id}')" class="h-20 bg-rose-50 border-2 border-rose-400 rounded-xl p-2 flex flex-col justify-between cursor-pointer shadow-sm hover:shadow-md transition-all relative overflow-hidden group">
        <div class="flex items-start justify-between gap-1">
          <span class="text-[11px] font-extrabold text-rose-800 leading-tight truncate">${mapel.nama}</span>
          <span class="material-symbols-outlined text-[16px] text-rose-600 animate-bounce flex-shrink-0">warning</span>
        </div>
        <div class="text-[10px] text-rose-700 font-semibold truncate">${guru.nama}</div>
        <div class="flex items-center justify-between text-[9px] font-bold text-rose-900 border-t border-rose-200/60 pt-1">
          <span>BENTROK!</span>
          <span class="underline">Perbaiki &rarr;</span>
        </div>
      </div>
    `;
  }

  return `
    <div onclick="openPlotModal('${hari}', ${jp}, '${kelasId}', '${item.id}')" class="h-20 bg-emerald-50/70 border border-emerald-300 rounded-xl p-2 flex flex-col justify-between cursor-pointer hover:border-emerald-500 hover:shadow-sm transition-all group">
      <div class="flex items-start justify-between gap-1">
        <span class="text-[11px] font-bold text-emerald-950 leading-tight truncate">${mapel.nama}</span>
        <span class="text-[9px] font-extrabold text-emerald-700 px-1 py-0.2 bg-emerald-200/70 rounded">${mapel.kode}</span>
      </div>
      <div class="text-[10px] text-emerald-800 font-medium truncate">${guru.nama}</div>
      <div class="flex items-center justify-between text-[9px] text-emerald-700/80 border-t border-emerald-200/60 pt-1">
        <span>${item.ruang_id || 'Kelas'}</span>
        <span class="opacity-0 group-hover:opacity-100 transition-opacity font-bold text-emerald-800">Ubah &rarr;</span>
      </div>
    </div>
  `;
}

// ============================================================================
// 10. MODAL SLOT PLOTTING HANDLERS
// ============================================================================
function openPlotModal(hari, jp, kelasId, existingJadwalId) {
  AppState.activeSlotEdit = { hari, jp, kelasId, id: existingJadwalId };
  
  const modal = document.getElementById('modal-plot-slot');
  const subtitle = document.getElementById('modal-slot-subtitle');
  const deleteBtn = document.getElementById('btn-delete-slot');

  const kelas = (AppState.master.kelas || []).find(k => k.id === kelasId);
  const kName = kelas ? kelas.nama_rombel : kelasId;
  if (subtitle) subtitle.textContent = `${hari}, JP ${jp} (${kName})`;

  const mapelSel = document.getElementById('modal-select-mapel');
  const guruSel = document.getElementById('modal-select-guru');
  const ruangSel = document.getElementById('modal-select-ruang');

  if (mapelSel) {
    mapelSel.innerHTML = (AppState.master.mapel || []).map(m => `
      <option value="${m.id}">${m.nama} (${m.kode})</option>
    `).join('');
  }

  if (guruSel) {
    guruSel.innerHTML = (AppState.master.guru || []).map(g => `
      <option value="${g.id}">${g.nama} - ${g.mapel_utama}</option>
    `).join('');
  }

  if (ruangSel) {
    ruangSel.innerHTML = (AppState.master.ruang || []).map(r => `
      <option value="${r.id}">${r.nama}</option>
    `).join('');
  }

  if (existingJadwalId) {
    const existing = AppState.jadwal.find(it => it.id === existingJadwalId);
    if (existing) {
      if (mapelSel) mapelSel.value = existing.mapel_id;
      if (guruSel) guruSel.value = existing.guru_id;
      if (ruangSel) ruangSel.value = existing.ruang_id || 'R_01';
    }
    if (deleteBtn) deleteBtn.style.display = 'block';
  } else {
    if (deleteBtn) deleteBtn.style.display = 'none';
  }

  checkModalSlotClash();
  if (modal) {
    modal.style.display = 'flex';
    modal.classList.remove('hidden');
  }
}

function closePlotModal() {
  const modal = document.getElementById('modal-plot-slot');
  if (modal) {
    modal.style.display = 'none';
    modal.classList.add('hidden');
  }
  AppState.activeSlotEdit = null;
}

function checkModalSlotClash() {
  if (!AppState.activeSlotEdit) return;
  const { hari, jp, kelasId, id } = AppState.activeSlotEdit;
  const guruId = document.getElementById('modal-select-guru')?.value;

  const warningBox = document.getElementById('modal-clash-warning');
  const warningText = document.getElementById('modal-clash-text');

  const otherClass = AppState.jadwal.find(it => 
    it.id !== id && 
    it.hari === hari && 
    parseInt(it.jp) === parseInt(jp) && 
    it.guru_id === guruId
  );

  const blockedReq = AppState.requests.find(r => 
    r.guru_id === guruId && 
    r.hari === hari && 
    parseInt(r.jp) === parseInt(jp)
  );

  if (otherClass) {
    const k = (AppState.master.kelas || []).find(k => k.id === otherClass.kelas_id);
    const kName = k ? k.nama_rombel : otherClass.kelas_id;
    if (warningBox && warningText) {
      warningBox.classList.remove('hidden');
      warningText.textContent = `Peringatan: Guru ini sudah dijadwalkan di ${kName} pada jam yang sama!`;
    }
  } else if (blockedReq) {
    if (warningBox && warningText) {
      warningBox.classList.remove('hidden');
      warningText.textContent = `Peringatan: Guru berhalangan mengajar (${blockedReq.alasan || 'Permohonan'}).`;
    }
  } else {
    if (warningBox) warningBox.classList.add('hidden');
  }
}

function confirmSlotPlotting() {
  if (!AppState.activeSlotEdit) return;
  const { hari, jp, kelasId, id } = AppState.activeSlotEdit;
  const mapelId = document.getElementById('modal-select-mapel').value;
  const guruId = document.getElementById('modal-select-guru').value;
  const ruangId = document.getElementById('modal-select-ruang').value;

  const newId = id || `J_${kelasId}_${hari}_${jp}`;
  const newItem = {
    id: newId,
    hari: hari,
    jp: jp,
    kelas_id: kelasId,
    guru_id: guruId,
    mapel_id: mapelId,
    ruang_id: ruangId,
    updated_at: new Date().toISOString()
  };

  const idx = AppState.jadwal.findIndex(it => it.id === newId || (it.hari === hari && parseInt(it.jp) === parseInt(jp) && it.kelas_id === kelasId));
  if (idx >= 0) {
    AppState.jadwal[idx] = newItem;
  } else {
    AppState.jadwal.push(newItem);
  }

  saveToLocalStorage();
  calculateClashes();
  closePlotModal();
  renderMatrixTable();
  showToast('Slot diperbarui secara instan (Optimistic UI)', 'success');
}

function deleteCurrentSlotAssignment() {
  if (!AppState.activeSlotEdit) return;
  const { hari, jp, kelasId, id } = AppState.activeSlotEdit;

  AppState.jadwal = AppState.jadwal.filter(it => 
    it.id !== id && !(it.hari === hari && parseInt(it.jp) === parseInt(jp) && it.kelas_id === kelasId)
  );

  saveToLocalStorage();
  calculateClashes();
  closePlotModal();
  renderMatrixTable();
  showToast('Slot jadwal telah dikosongkan', 'info');
}

function openClashDetailModal() {
  const modal = document.getElementById('modal-clash-details');
  const countSub = document.getElementById('clash-modal-count-sub');
  const list = document.getElementById('clash-details-list');

  if (countSub) countSub.textContent = `${AppState.clashes.length} konflik terdeteksi`;

  if (list) {
    if (AppState.clashes.length === 0) {
      list.innerHTML = `<div class="p-4 text-center text-slate-500 font-semibold">Tidak ada konflik jadwal. Semua jadwal aman!</div>`;
    } else {
      list.innerHTML = AppState.clashes.map(c => `
        <div class="p-3.5 rounded-xl bg-rose-50 border border-rose-200 text-xs text-rose-900 space-y-1">
          <div class="flex items-center justify-between font-bold">
            <span class="flex items-center gap-1.5 text-rose-700">
              <span class="material-symbols-outlined text-[16px]">error</span>
              <span>${c.title}</span>
            </span>
            <span class="text-[10px] px-2 py-0.5 rounded bg-rose-200 text-rose-800">${c.hari} JP ${c.jp}</span>
          </div>
          <p class="text-text-secondary leading-snug">${c.message}</p>
        </div>
      `).join('');
    }
  }

  if (modal) {
    modal.style.display = 'flex';
    modal.classList.remove('hidden');
  }
}

function closeClashDetailModal() {
  const modal = document.getElementById('modal-clash-details');
  if (modal) {
    modal.style.display = 'none';
    modal.classList.add('hidden');
  }
}

// ============================================================================
// 11. TEACHER AVAILABILITY GRID
// ============================================================================
function renderAvailabilityView() {
  const teacherSel = document.getElementById('avail-teacher-select');
  if (teacherSel && teacherSel.options.length <= 1) {
    teacherSel.innerHTML = (AppState.master.guru || []).map(g => `
      <option value="${g.id}">${g.nama} (${g.mapel_utama})</option>
    `).join('');
  }
  loadTeacherAvailabilityMatrix();
}

function loadTeacherAvailabilityMatrix() {
  const tbody = document.getElementById('avail-matrix-tbody');
  const teacherSel = document.getElementById('avail-teacher-select');
  if (!tbody || !teacherSel) return;

  const currentTeacherId = teacherSel.value || 'G14';
  const days = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const teacherRequests = AppState.requests.filter(r => r.guru_id === currentTeacherId);

  let html = '';
  for (let jp = 1; jp <= 8; jp++) {
    html += `<tr>`;
    html += `<td class="p-2.5 bg-slate-50 font-bold text-center border-r border-border-subtle">JP ${jp}</td>`;
    
    days.forEach(d => {
      const isBlocked = teacherRequests.some(r => r.hari === d && parseInt(r.jp) === jp);
      html += `
        <td class="p-2 text-center border-r border-border-subtle">
          <button type="button" onclick="toggleTeacherSlotAvailability('${currentTeacherId}', '${d}', ${jp})" class="w-full py-2.5 px-1 rounded-xl text-xs font-bold transition-all ${isBlocked ? 'bg-rose-100 text-rose-800 border-2 border-rose-300 hover:bg-rose-200' : 'bg-emerald-50 text-emerald-800 border border-emerald-300 hover:bg-emerald-100'}">
            ${isBlocked ? '✕ Berhalangan' : '✓ Tersedia'}
          </button>
        </td>
      `;
    });
    html += `</tr>`;
  }
  tbody.innerHTML = html;
}

function toggleTeacherSlotAvailability(teacherId, hari, jp) {
  const existingIdx = AppState.requests.findIndex(r => r.guru_id === teacherId && r.hari === hari && parseInt(r.jp) === parseInt(jp));

  if (existingIdx >= 0) {
    AppState.requests.splice(existingIdx, 1);
  } else {
    AppState.requests.push({
      id: `REQ_${teacherId}_${hari}_${jp}`,
      guru_id: teacherId,
      hari: hari,
      jp: jp,
      alasan: 'Sekolah Mitra / Kepentingan Dinas',
      status: 'Disetujui',
      timestamp: new Date().toISOString()
    });
  }

  saveToLocalStorage();
  calculateClashes();
  loadTeacherAvailabilityMatrix();
  showToast('Ketersediaan diperbarui instan', 'info');
}

function clearCurrentTeacherBlocks() {
  const teacherSel = document.getElementById('avail-teacher-select');
  if (!teacherSel) return;
  const currentTeacherId = teacherSel.value;

  AppState.requests = AppState.requests.filter(r => r.guru_id !== currentTeacherId);
  saveToLocalStorage();
  calculateClashes();
  loadTeacherAvailabilityMatrix();
  showToast('Semua slot ketersediaan dibuka', 'success');
}

function applyMitraTemplate() {
  const teacherSel = document.getElementById('avail-teacher-select');
  if (!teacherSel) return;
  const currentTeacherId = teacherSel.value;

  const templateSlots = [
    { hari: 'Kamis', jp: 5 }, { hari: 'Kamis', jp: 6 }, { hari: 'Kamis', jp: 7 },
    { hari: 'Sabtu', jp: 1 }, { hari: 'Sabtu', jp: 2 }
  ];

  templateSlots.forEach(s => {
    if (!AppState.requests.some(r => r.guru_id === currentTeacherId && r.hari === s.hari && r.jp === s.jp)) {
      AppState.requests.push({
        id: `REQ_${currentTeacherId}_${s.hari}_${s.jp}`,
        guru_id: currentTeacherId,
        hari: s.hari,
        jp: s.jp,
        alasan: 'Penugasan Sekolah Mitra',
        status: 'Disetujui',
        timestamp: new Date().toISOString()
      });
    }
  });

  saveToLocalStorage();
  calculateClashes();
  loadTeacherAvailabilityMatrix();
  showToast('Template Mitra berhasil diterapkan', 'success');
}

// ============================================================================
// 12. MASTER DATA CRUD & IMPORT / EXPORT (/gas-pro-api)
// ============================================================================
function setMasterTab(tabName) {
  AppState.activeMasterTab = tabName;
  AppState.selectedMasterIds.clear();
  updateBulkDeleteButton();

  document.querySelectorAll('.master-tab-btn').forEach(btn => {
    btn.className = 'master-tab-btn px-4 py-2.5 rounded-t-xl text-xs font-bold text-text-secondary hover:text-text-primary flex items-center gap-2';
  });
  const activeBtn = document.getElementById(`tab-master-${tabName}`);
  if (activeBtn) {
    activeBtn.className = 'master-tab-btn px-4 py-2.5 rounded-t-xl text-xs font-bold bg-white text-primary border-t-2 border-primary shadow-sm flex items-center gap-2';
  }
  renderMasterView();
}

function renderMasterView() {
  const thead = document.getElementById('master-thead');
  const tbody = document.getElementById('master-tbody');
  if (!thead || !tbody) return;

  const tab = AppState.activeMasterTab;
  const items = AppState.master[tab] || [];
  const allChecked = items.length > 0 && items.every(it => AppState.selectedMasterIds.has(it.id));

  if (tab === 'guru') {
    thead.innerHTML = `
      <tr>
        <th class="p-3 w-10 text-center">
          <input type="checkbox" id="master-check-all" onchange="toggleSelectAllMaster(this.checked)" class="rounded text-primary focus:ring-primary cursor-pointer" ${allChecked ? 'checked' : ''}/>
        </th>
        <th class="p-3">Nama Guru & NIP</th>
        <th class="p-3">Mapel Utama</th>
        <th class="p-3">Status</th>
        <th class="p-3 text-center">Target JTM</th>
        <th class="p-3 text-center">JTM Terplot</th>
        <th class="p-3">Status Beban</th>
        <th class="p-3 text-right">Aksi</th>
      </tr>
    `;

    const teacherJtmMap = {};
    AppState.master.guru.forEach(g => { teacherJtmMap[g.id] = 0; });
    AppState.jadwal.forEach(it => {
      if (it.guru_id && teacherJtmMap[it.guru_id] !== undefined) {
        teacherJtmMap[it.guru_id] += 1;
      }
    });

    tbody.innerHTML = (AppState.master.guru || []).map(g => {
      const terplot = teacherJtmMap[g.id] || 0;
      const target = g.target_jtm || 24;
      const isComplete = terplot >= target;
      const isChecked = AppState.selectedMasterIds.has(g.id);

      return `
        <tr class="hover:bg-slate-50 transition-colors ${isChecked ? 'bg-blue-50/50' : ''}">
          <td class="p-3 text-center">
            <input type="checkbox" value="${g.id}" onchange="toggleMasterItemSelection('${g.id}', this.checked)" class="master-row-check rounded text-primary focus:ring-primary cursor-pointer" ${isChecked ? 'checked' : ''}/>
          </td>
          <td class="p-3 font-semibold text-text-primary">
            <div class="font-bold text-sm text-text-primary">${g.nama}${g.gelar ? ', ' + g.gelar : ''}</div>
            <div class="text-[11px] text-text-secondary">NIP. ${g.nip || '-'}</div>
          </td>
          <td class="p-3 text-text-primary font-medium">${g.mapel_utama}</td>
          <td class="p-3">
            <span class="px-2 py-0.5 rounded text-[10px] font-bold ${g.status_kepegawaian === 'PNS' ? 'bg-blue-100 text-blue-800' : g.status_kepegawaian === 'P3K' ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-200 text-slate-700'}">
              ${g.status_kepegawaian || 'PNS'}
            </span>
          </td>
          <td class="p-3 text-center font-bold text-text-primary">${target} JP</td>
          <td class="p-3 text-center font-bold ${isComplete ? 'text-emerald-600' : 'text-amber-600'}">${terplot} JP</td>
          <td class="p-3">
            <div class="w-24 bg-slate-100 rounded-full h-2 overflow-hidden mb-1">
              <div class="h-2 rounded-full ${isComplete ? 'bg-emerald-500' : 'bg-amber-500'}" style="width: ${Math.min(100, (terplot/target)*100)}%"></div>
            </div>
            <span class="text-[10px] font-semibold ${isComplete ? 'text-emerald-700' : 'text-amber-700'}">
              ${isComplete ? 'Tercapai' : `${target - terplot} JP Tersisa`}
            </span>
          </td>
          <td class="p-3 text-right">
            <button onclick="editMasterItem('guru', '${g.id}')" class="p-1 rounded text-text-secondary hover:text-primary hover:bg-slate-100" title="Edit Guru">
              <span class="material-symbols-outlined text-[18px]">edit</span>
            </button>
          </td>
        </tr>
      `;
    }).join('');

  } else if (tab === 'kelas') {
    thead.innerHTML = `
      <tr>
        <th class="p-3 w-10 text-center">
          <input type="checkbox" id="master-check-all" onchange="toggleSelectAllMaster(this.checked)" class="rounded text-primary focus:ring-primary cursor-pointer" ${allChecked ? 'checked' : ''}/>
        </th>
        <th class="p-3">Kode Rombel</th>
        <th class="p-3">Nama Rombongan Belajar</th>
        <th class="p-3">Tingkat</th>
        <th class="p-3">Wali Kelas</th>
        <th class="p-3 text-center">Jumlah Siswa</th>
      </tr>
    `;
    const guruMap = {};
    AppState.master.guru.forEach(g => { guruMap[g.id] = g.nama; });

    tbody.innerHTML = (AppState.master.kelas || []).map(k => {
      const isChecked = AppState.selectedMasterIds.has(k.id);
      return `
        <tr class="hover:bg-slate-50 transition-colors ${isChecked ? 'bg-blue-50/50' : ''}">
          <td class="p-3 text-center">
            <input type="checkbox" value="${k.id}" onchange="toggleMasterItemSelection('${k.id}', this.checked)" class="master-row-check rounded text-primary focus:ring-primary cursor-pointer" ${isChecked ? 'checked' : ''}/>
          </td>
          <td class="p-3 font-bold text-primary">${k.kode_kelas}</td>
          <td class="p-3 font-bold text-text-primary">${k.nama_rombel}</td>
          <td class="p-3 font-semibold text-text-secondary">Kelas ${k.tingkat}</td>
          <td class="p-3 font-medium text-text-primary">${guruMap[k.wali_kelas_id] || '-'}</td>
          <td class="p-3 text-center font-bold text-text-primary">${k.total_siswa || 28} Anak</td>
        </tr>
      `;
    }).join('');

  } else if (tab === 'mapel') {
    thead.innerHTML = `
      <tr>
        <th class="p-3 w-10 text-center">
          <input type="checkbox" id="master-check-all" onchange="toggleSelectAllMaster(this.checked)" class="rounded text-primary focus:ring-primary cursor-pointer" ${allChecked ? 'checked' : ''}/>
        </th>
        <th class="p-3">Kode</th>
        <th class="p-3">Nama Mata Pelajaran</th>
        <th class="p-3">Kategori</th>
        <th class="p-3 text-center">Alokasi JP / Minggu</th>
      </tr>
    `;
    tbody.innerHTML = (AppState.master.mapel || []).map(m => {
      const isChecked = AppState.selectedMasterIds.has(m.id);
      return `
        <tr class="hover:bg-slate-50 transition-colors ${isChecked ? 'bg-blue-50/50' : ''}">
          <td class="p-3 text-center">
            <input type="checkbox" value="${m.id}" onchange="toggleMasterItemSelection('${m.id}', this.checked)" class="master-row-check rounded text-primary focus:ring-primary cursor-pointer" ${isChecked ? 'checked' : ''}/>
          </td>
          <td class="p-3 font-bold text-primary">${m.kode}</td>
          <td class="p-3 font-bold text-text-primary">${m.nama}</td>
          <td class="p-3">
            <span class="px-2 py-0.5 rounded text-[10px] font-bold ${m.kategori === 'Wajib' ? 'bg-indigo-100 text-indigo-800' : 'bg-amber-100 text-amber-800'}">
              ${m.kategori}
            </span>
          </td>
          <td class="p-3 text-center font-bold text-text-primary">${m.alokasi_jp_minggu} JP</td>
        </tr>
      `;
    }).join('');

  } else {
    thead.innerHTML = `
      <tr>
        <th class="p-3 w-10 text-center">
          <input type="checkbox" id="master-check-all" onchange="toggleSelectAllMaster(this.checked)" class="rounded text-primary focus:ring-primary cursor-pointer" ${allChecked ? 'checked' : ''}/>
        </th>
        <th class="p-3">Kode Ruang</th>
        <th class="p-3">Nama Fasilitas & Ruangan</th>
        <th class="p-3">Jenis</th>
        <th class="p-3 text-center">Kapasitas</th>
        <th class="p-3">Fasilitas Penunjang</th>
      </tr>
    `;
    tbody.innerHTML = (AppState.master.ruang || []).map(r => {
      const isChecked = AppState.selectedMasterIds.has(r.id);
      return `
        <tr class="hover:bg-slate-50 transition-colors ${isChecked ? 'bg-blue-50/50' : ''}">
          <td class="p-3 text-center">
            <input type="checkbox" value="${r.id}" onchange="toggleMasterItemSelection('${r.id}', this.checked)" class="master-row-check rounded text-primary focus:ring-primary cursor-pointer" ${isChecked ? 'checked' : ''}/>
          </td>
          <td class="p-3 font-bold text-primary">${r.kode}</td>
          <td class="p-3 font-bold text-text-primary">${r.nama}</td>
          <td class="p-3 font-medium text-text-secondary">${r.jenis}</td>
          <td class="p-3 text-center font-bold text-text-primary">${r.kapasitas} Orang</td>
          <td class="p-3 text-text-secondary">${r.fasilitas || '-'}</td>
        </tr>
      `;
    }).join('');
  }

  updateBulkDeleteButton();
}

function toggleSelectAllMaster(checked) {
  const tab = AppState.activeMasterTab;
  const items = AppState.master[tab] || [];
  if (checked) {
    items.forEach(it => AppState.selectedMasterIds.add(it.id));
  } else {
    AppState.selectedMasterIds.clear();
  }
  renderMasterView();
}

function toggleMasterItemSelection(id, checked) {
  if (checked) {
    AppState.selectedMasterIds.add(id);
  } else {
    AppState.selectedMasterIds.delete(id);
  }
  updateBulkDeleteButton();
}

function updateBulkDeleteButton() {
  const btn = document.getElementById('btn-bulk-delete-master');
  const txt = document.getElementById('bulk-delete-count-text');
  const count = AppState.selectedMasterIds.size;
  if (!btn) return;

  if (count > 0) {
    btn.style.display = 'inline-flex';
    if (txt) txt.textContent = `Hapus (${count})`;
  } else {
    btn.style.display = 'none';
  }
}

async function deleteSelectedMasterItems() {
  const count = AppState.selectedMasterIds.size;
  if (count === 0) return;
  const tab = AppState.activeMasterTab;

  if (!confirm(`Apakah Anda yakin ingin menghapus ${count} data ${tab} terpilih?`)) {
    return;
  }

  const idsToDelete = Array.from(AppState.selectedMasterIds);
  AppState.master[tab] = (AppState.master[tab] || []).filter(it => !idsToDelete.includes(it.id));
  AppState.selectedMasterIds.clear();

  saveToLocalStorage();
  renderMasterView();
  showToast(`${count} data ${tab} berhasil dihapus dari memori`, 'success');

  // Sinkronkan ke server
  const res = await apiPost('batchDeleteMaster', { tab: tab, ids: idsToDelete });
  if (res && res.success) {
    showToast(`Sinkronisasi hapus ke database Sheets berhasil`, 'success');
  }
}

// ----------------------------------------------------------------------------
// REVISI 3: TEMPLATE UNDUH & IMPOR FILE EXCEL / CSV
// ----------------------------------------------------------------------------
function downloadMasterTemplate() {
  if (typeof XLSX === 'undefined') {
    return showToast('Library Excel (SheetJS) belum siap. Silakan muat ulang.', 'warning');
  }

  const tab = AppState.activeMasterTab;
  let headers = [];
  let sampleData = [];
  let filename = `Template_Master_${tab.toUpperCase()}_SDN_Pilangrejo1.xlsx`;

  if (tab === 'guru') {
    headers = ['nama', 'gelar', 'nip', 'mapel_utama', 'target_jtm', 'status_kepegawaian'];
    sampleData = [
      { nama: 'Budi Santoso', gelar: 'S.Pd.SD', nip: '198501012010011005', mapel_utama: 'Guru Kelas 4', target_jtm: 24, status_kepegawaian: 'PNS' },
      { nama: 'Siti Aminah', gelar: 'S.Pd.I', nip: '199002022015022003', mapel_utama: 'Pendidikan Agama Islam', target_jtm: 24, status_kepegawaian: 'P3K' }
    ];
  } else if (tab === 'kelas') {
    headers = ['kode_kelas', 'nama_rombel', 'tingkat', 'wali_kelas_id', 'total_siswa'];
    sampleData = [
      { kode_kelas: '4A', nama_rombel: 'Kelas 4A (Garuda)', tingkat: 4, wali_kelas_id: 'G02', total_siswa: 30 },
      { kode_kelas: '4B', nama_rombel: 'Kelas 4B (Rajawali)', tingkat: 4, wali_kelas_id: 'G11', total_siswa: 29 }
    ];
  } else if (tab === 'mapel') {
    headers = ['kode', 'nama', 'kategori', 'alokasi_jp_minggu'];
    sampleData = [
      { kode: 'MTK', nama: 'Matematika', kategori: 'Wajib', alokasi_jp_minggu: 5 },
      { kode: 'IPAS', nama: 'Ilmu Pengetahuan Alam & Sosial', kategori: 'Wajib', alokasi_jp_minggu: 5 }
    ];
  } else {
    headers = ['kode', 'nama', 'jenis', 'kapasitas', 'fasilitas'];
    sampleData = [
      { kode: 'RK-4A', nama: 'Ruang Kelas 4A', jenis: 'Kelas', kapasitas: 32, fasilitas: 'Papan Tulis, Proyektor' },
      { kode: 'LAB-KOMP', nama: 'Laboratorium Komputer', jenis: 'Laboratorium', kapasitas: 36, fasilitas: '32 Unit PC, AC' }
    ];
  }

  const ws = XLSX.utils.json_to_sheet(sampleData, { header: headers });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `Master_${tab}`);
  XLSX.writeFile(wb, filename);

  showToast(`Template ${tab.toUpperCase()} berhasil diunduh (.xlsx)`, 'success');
}

function handleMasterFileSelect(e) {
  const file = e.target.files && e.target.files[0];
  if (!file) return;

  if (typeof XLSX === 'undefined') {
    return showToast('Library SheetJS belum siap, silakan refresh halaman.', 'error');
  }

  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

      if (!rows || rows.length === 0) {
        showToast('Berkas tidak berisi data valid.', 'warning');
        return;
      }

      openImportPreviewModal(rows);
    } catch (err) {
      showToast('Gagal membaca berkas Excel/CSV: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
  e.target.value = '';
}

function openImportPreviewModal(rows) {
  AppState._pendingImport = rows;
  const modal = document.getElementById('modal-import-preview');
  const thead = document.getElementById('import-preview-thead');
  const tbody = document.getElementById('import-preview-tbody');
  const summary = document.getElementById('import-summary-text');
  const tab = AppState.activeMasterTab;

  if (summary) {
    summary.textContent = `Terdeteksi ${rows.length} baris data siap diimpor ke Master ${tab.toUpperCase()}.`;
  }

  if (rows.length > 0 && thead && tbody) {
    const keys = Object.keys(rows[0]);
    thead.innerHTML = `<tr>${keys.map(k => `<th class="p-2 border border-slate-200">${k}</th>`).join('')}</tr>`;
    tbody.innerHTML = rows.slice(0, 10).map(r => `
      <tr>${keys.map(k => `<td class="p-2 border border-slate-200">${r[k] !== undefined ? r[k] : ''}</td>`).join('')}</tr>
    `).join('') + (rows.length > 10 ? `<tr><td colspan="${keys.length}" class="p-2 text-center text-text-secondary italic">...dan ${rows.length - 10} baris lainnya</td></tr>` : '');
  }

  if (modal) {
    modal.style.display = 'flex';
    modal.classList.remove('hidden');
  }
}

function closeImportPreviewModal() {
  const modal = document.getElementById('modal-import-preview');
  if (modal) {
    modal.style.display = 'none';
    modal.classList.add('hidden');
  }
  AppState._pendingImport = null;
}

async function confirmImportData() {
  const rows = AppState._pendingImport;
  if (!rows || rows.length === 0) return;

  const tab = AppState.activeMasterTab;
  let counter = (AppState.master[tab] || []).length;

  const formattedItems = rows.map(r => {
    counter++;
    const prefix = tab === 'guru' ? 'G' : tab === 'kelas' ? 'K' : tab === 'mapel' ? 'M' : 'R_';
    const id = r.id || `${prefix}${counter.toString().padStart(2, '0')}`;
    return { ...r, id };
  });

  AppState.master[tab] = (AppState.master[tab] || []).concat(formattedItems);
  saveToLocalStorage();
  closeImportPreviewModal();
  renderMasterView();
  showToast(`Berhasil mengimpor ${formattedItems.length} data ke Master ${tab.toUpperCase()}!`, 'success');

  // Kirim batch ke REST API Backend
  const res = await apiPost('batchImportMaster', { tab: tab, items: formattedItems });
  if (res && res.success) {
    showToast(`Sinkronisasi ${formattedItems.length} data ke Google Sheets berhasil!`, 'success');
  }
}

function openAddMasterModal() {
  const modal = document.getElementById('modal-master-crud');
  const title = document.getElementById('modal-master-title');
  const container = document.getElementById('master-form-fields');
  const tab = AppState.activeMasterTab;

  if (title) title.textContent = `Tambah Data ${tab.toUpperCase()}`;

  if (container) {
    if (tab === 'guru') {
      container.innerHTML = `
        <div>
          <label class="block text-xs font-bold text-text-primary mb-1">Nama Guru Lengkap:</label>
          <input type="text" id="crud-guru-nama" required placeholder="misal: Siti Nurhaliza" class="w-full text-xs border rounded-lg p-2.5 bg-slate-50"/>
        </div>
        <div>
          <label class="block text-xs font-bold text-text-primary mb-1">Gelar Akademik:</label>
          <input type="text" id="crud-guru-gelar" placeholder="misal: S.Pd.SD" class="w-full text-xs border rounded-lg p-2.5 bg-slate-50"/>
        </div>
        <div>
          <label class="block text-xs font-bold text-text-primary mb-1">NIP (Nomor Induk Pegawai):</label>
          <input type="text" id="crud-guru-nip" placeholder="19xxxxxxxxxxxxxx" class="w-full text-xs border rounded-lg p-2.5 bg-slate-50"/>
        </div>
        <div>
          <label class="block text-xs font-bold text-text-primary mb-1">Mata Pelajaran Utama:</label>
          <input type="text" id="crud-guru-mapel" required placeholder="misal: Guru Kelas 3" class="w-full text-xs border rounded-lg p-2.5 bg-slate-50"/>
        </div>
        <div class="grid grid-cols-2 gap-3">
          <div>
            <label class="block text-xs font-bold text-text-primary mb-1">Target JTM:</label>
            <input type="number" id="crud-guru-target" value="24" class="w-full text-xs border rounded-lg p-2.5 bg-slate-50"/>
          </div>
          <div>
            <label class="block text-xs font-bold text-text-primary mb-1">Status:</label>
            <select id="crud-guru-status" class="w-full text-xs border rounded-lg p-2.5 bg-slate-50">
              <option value="PNS">PNS</option>
              <option value="P3K">P3K</option>
              <option value="Honorer">Honorer</option>
            </select>
          </div>
        </div>
      `;
    } else {
      container.innerHTML = `<div class="p-3 text-xs text-text-secondary">Formulir penambahan data aktif untuk ${tab}.</div>`;
    }
  }

  if (modal) {
    modal.style.display = 'flex';
    modal.classList.remove('hidden');
  }
}

function closeMasterModal() {
  const modal = document.getElementById('modal-master-crud');
  if (modal) {
    modal.style.display = 'none';
    modal.classList.add('hidden');
  }
}

function editMasterItem(type, id) {
  openAddMasterModal();
  if (type === 'guru') {
    const g = (AppState.master.guru || []).find(it => it.id === id);
    if (g) {
      const modalTitle = document.getElementById('modal-master-title');
      if (modalTitle) modalTitle.textContent = `Edit Data Guru (${g.nama})`;
      const namaIn = document.getElementById('crud-guru-nama');
      const gelarIn = document.getElementById('crud-guru-gelar');
      const nipIn = document.getElementById('crud-guru-nip');
      const mapelIn = document.getElementById('crud-guru-mapel');
      const targetIn = document.getElementById('crud-guru-target');
      const statusIn = document.getElementById('crud-guru-status');
      if (namaIn) namaIn.value = g.nama || '';
      if (gelarIn) gelarIn.value = g.gelar || '';
      if (nipIn) nipIn.value = g.nip || '';
      if (mapelIn) mapelIn.value = g.mapel_utama || '';
      if (targetIn) targetIn.value = g.target_jtm || 24;
      if (statusIn) statusIn.value = g.status_kepegawaian || 'PNS';
    }
  }
}

async function handleMasterSubmit(e) {
  if (e && e.preventDefault) e.preventDefault();
  const tab = AppState.activeMasterTab;

  if (tab === 'guru') {
    const nama = document.getElementById('crud-guru-nama')?.value || '';
    const gelar = document.getElementById('crud-guru-gelar')?.value || '';
    const nip = document.getElementById('crud-guru-nip')?.value || '';
    const mapel = document.getElementById('crud-guru-mapel')?.value || '';
    const target = parseInt(document.getElementById('crud-guru-target')?.value) || 24;
    const status = document.getElementById('crud-guru-status')?.value || 'PNS';

    const newGuru = {
      id: `G${(AppState.master.guru.length + 1).toString().padStart(2, '0')}`,
      nip, nama, gelar, mapel_utama: mapel, target_jtm: target, status_kepegawaian: status
    };

    AppState.master.guru.push(newGuru);
    saveToLocalStorage();
    closeMasterModal();
    renderMasterView();
    showToast('Data guru disimpan ke memori & antrean server', 'success');

    // Sinkronisasi ke REST API
    apiPost('saveMaster', { tab: 'guru', item: newGuru });
  }
}

// ============================================================================
// REVISI 1: PENGATURAN IDENTITAS SEKOLAH & LOGO RESMI
// ============================================================================
function renderSchoolView() {
  loadSchoolIdentity();
}

function loadSchoolIdentity() {
  const sekolah = AppState.master.sekolah || [];
  const cfg = {};
  if (Array.isArray(sekolah)) {
    sekolah.forEach(row => {
      if (row.kunci !== undefined) cfg[row.kunci] = row.nilai;
      else if (row.key !== undefined) cfg[row.key] = row.value;
      else if (row.Kunci !== undefined) cfg[row.Kunci] = row.Nilai;
    });
  } else if (typeof sekolah === 'object' && sekolah !== null) {
    Object.assign(cfg, sekolah);
  }

  const setVal = (id, val) => {
    const el = document.getElementById(id);
    if (el && val !== undefined && val !== null) el.value = val;
  };

  const setText = (id, val) => {
    const el = document.getElementById(id);
    if (el && val !== undefined && val !== null) el.textContent = val;
  };

  const namaSekolah = cfg.nama_sekolah || 'Sekolah Dasar Negeri Pilangrejo 1';
  const npsn = cfg.npsn || '20507812';
  const nss = cfg.nss || '101054002001';
  const akreditasi = cfg.akreditasi || 'A (Amat Baik)';
  const alamat = cfg.alamat || 'Jl. Raya Pilangrejo No. 12, Pilangkenceng, Kab. Madiun 63154';
  const email = cfg.email || 'sdn.pilangrejo1@madiunkab.go.id';
  const website = cfg.website || 'sdn-pilangrejo1.sch.id';
  const kepsek = cfg.kepala_sekolah || 'Dra. Sri Mulyani, M.Pd';
  const nipKepsek = cfg.nip_kepala_sekolah || '19740512 199803 2 004';
  const tahunAjaran = cfg.tahun_ajaran || '2024/2025';
  const semester = cfg.semester || 'Ganjil';
  const logoSekolah = cfg.logo_sekolah || AppState.schoolLogoBase64 || '';

  // 1. Form Inputs di Tab Identitas Sekolah
  setVal('sch-nama', namaSekolah);
  setVal('sch-npsn', npsn);
  setVal('sch-nss', nss);
  setVal('sch-akreditasi', akreditasi);
  setVal('sch-alamat', alamat);
  setVal('sch-email', email);
  setVal('sch-website', website);
  setVal('sch-kepsek', kepsek);
  setVal('sch-nip-kepsek', nipKepsek);
  setVal('sch-tahun-ajaran', tahunAjaran);
  setVal('sch-semester', semester);

  // 2. Tampilan Global UI: Subtitle Sidebar, Header, Login, Dashboard Badge
  setText('sidebar-school-subtitle', namaSekolah);
  setText('login-school-name-subtitle', `${namaSekolah} • Kab. Madiun`);
  setText('header-school-name', namaSekolah);
  setText('header-ta-text', `TP ${tahunAjaran}`);
  setText('dash-ta-badge', `Tahun Ajaran ${tahunAjaran}`);

  // 3. Kop Surat & Pengesahan Dokumen Cetak / Pratinjau SK
  setText('print-school-name', namaSekolah.toUpperCase());
  setText('print-school-alamat', `Alamat: ${alamat}`);
  setText('print-school-meta', `Email: ${email} | NPSN: ${npsn} | Akreditasi: ${akreditasi}`);
  setText('print-ta-sub', `SEMESTER ${String(semester).toUpperCase()} - TAHUN PELAJARAN ${tahunAjaran}`);
  setText('print-kepsek-ttd', kepsek.replace(/(Dr\.|Dra\.|Drs\.|H\.|Hj\.|,.*)/gi, '').trim() || kepsek);
  setText('print-kepsek-nama', kepsek);
  setText('print-kepsek-nip', `NIP. ${nipKepsek}`);

  // 4. Sinkronkan Logo di seluruh elemen tampilan
  updateSchoolLogoUI(logoSekolah);
}

function updateSchoolLogoUI(logoSrc) {
  if (logoSrc) AppState.schoolLogoBase64 = logoSrc;
  const src = logoSrc || AppState.schoolLogoBase64;

  // Pratinjau Form Pengaturan
  const imgEl = document.getElementById('school-logo-preview');
  const placeholderEl = document.getElementById('school-logo-placeholder');
  if (imgEl && placeholderEl) {
    if (src) {
      imgEl.src = src;
      imgEl.style.display = 'block';
      placeholderEl.style.display = 'none';
    } else {
      imgEl.style.display = 'none';
      placeholderEl.style.display = 'flex';
    }
  }

  // Logo Sidebar
  const sideImg = document.getElementById('sidebar-logo-img');
  const sideFallback = document.getElementById('sidebar-logo-fallback');
  if (sideImg && sideFallback) {
    if (src) {
      sideImg.src = src;
      sideImg.style.display = 'block';
      sideFallback.style.display = 'none';
    } else {
      sideImg.style.display = 'none';
      sideFallback.style.display = 'inline-block';
    }
  }

  // Logo Halaman Login
  const loginImg = document.getElementById('login-school-logo-img');
  const loginFallback = document.getElementById('login-school-logo-fallback');
  if (loginImg && loginFallback) {
    if (src) {
      loginImg.src = src;
      loginImg.style.display = 'block';
      loginFallback.style.display = 'none';
    } else {
      loginImg.style.display = 'none';
      loginFallback.style.display = 'inline-block';
    }
  }

  // Logo Kop Surat Dokumen Cetak
  const printImg = document.getElementById('print-school-logo-img');
  const printFallback = document.getElementById('print-logo-fallback');
  if (printImg && printFallback) {
    if (src) {
      printImg.src = src;
      printImg.style.display = 'block';
      printFallback.style.display = 'none';
    } else {
      printImg.style.display = 'none';
      printFallback.style.display = 'inline-block';
    }
  }
}

function handleSchoolLogoUpload(e) {
  const file = e.target.files && e.target.files[0];
  if (!file) return;

  if (file.size > 2 * 1024 * 1024) {
    showToast('Ukuran logo maksimal 2MB!', 'warning');
    return;
  }

  const reader = new FileReader();
  reader.onload = function(evt) {
    AppState.schoolLogoBase64 = evt.target.result;
    updateSchoolLogoUI(AppState.schoolLogoBase64);
    showToast('Pratinjau logo berhasil dimuat. Klik Simpan & Sinkronkan.', 'info');
  };
  reader.readAsDataURL(file);
}

async function saveSchoolIdentityForm() {
  const getVal = id => document.getElementById(id)?.value?.trim() || '';

  const config = {
    nama_sekolah: getVal('sch-nama'),
    npsn: getVal('sch-npsn'),
    nss: getVal('sch-nss'),
    akreditasi: getVal('sch-akreditasi'),
    alamat: getVal('sch-alamat'),
    email: getVal('sch-email'),
    website: getVal('sch-website'),
    kepala_sekolah: getVal('sch-kepsek'),
    nip_kepala_sekolah: getVal('sch-nip-kepsek'),
    tahun_ajaran: getVal('sch-tahun-ajaran'),
    semester: getVal('sch-semester'),
    logo_sekolah: AppState.schoolLogoBase64 || ''
  };

  // Optimistic UI update seketika di seluruh halaman
  AppState.master.sekolah = Object.keys(config).map(k => ({ kunci: k, nilai: config[k] }));
  saveToLocalStorage();
  loadSchoolIdentity();

  showToast('Menyimpan identitas sekolah & menyinkronkan ke database...', 'info');
  const btn = document.getElementById('btn-save-school');
  if (btn) btn.classList.add('opacity-75');

  const res = await apiPost('saveSchoolIdentity', {
    config: config,
    logoBase64: AppState.schoolLogoBase64 || '',
    ...config
  });
  if (btn) btn.classList.remove('opacity-75');

  if (res && res.success) {
    if (res.logo_url) {
      AppState.schoolLogoBase64 = res.logo_url;
      config.logo_sekolah = res.logo_url;
    }
    if (res.sekolah && Array.isArray(res.sekolah)) {
      AppState.master.sekolah = res.sekolah;
    }
    saveToLocalStorage();
    loadSchoolIdentity();
    showToast('Identitas sekolah berhasil disimpan & disinkronkan ke Google Sheets!', 'success');
  } else {
    showToast('Identitas sekolah disimpan di penyimpanan lokal browser', 'info');
  }
}

// ============================================================================
// REVISI 2: PENGATURAN USER & HAK AKSES MENU (RBAC)
// ============================================================================
function renderUsersView() {
  const tbody = document.getElementById('users-table-tbody');
  if (!tbody) return;

  const users = AppState.users || [];
  if (users.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="p-4 text-center text-slate-500">Belum ada akun pengguna. Klik Tambah Pengguna.</td></tr>`;
    return;
  }

  const roleLabels = {
    admin: { label: 'Kurikulum (Admin)', color: 'bg-indigo-100 text-indigo-800' },
    kepsek: { label: 'Kepala Sekolah', color: 'bg-emerald-100 text-emerald-800' },
    guru: { label: 'Guru Pengajar', color: 'bg-blue-100 text-blue-800' },
    operator: { label: 'Operator', color: 'bg-slate-100 text-slate-800' }
  };

  tbody.innerHTML = users.map(u => {
    const roleInfo = roleLabels[u.role] || { label: u.role, color: 'bg-slate-100 text-slate-800' };
    const perms = Array.isArray(u.menu_access) ? u.menu_access : (u.menu_access ? String(u.menu_access).split(',') : ['*']);
    const isAll = perms.includes('*');

    return `
      <tr class="hover:bg-slate-50 transition-colors">
        <td class="p-3 font-mono font-bold text-text-primary">${u.username}</td>
        <td class="p-3 font-bold text-text-primary">${u.nama || '-'}</td>
        <td class="p-3">
          <span class="px-2 py-0.5 rounded text-[10px] font-bold ${roleInfo.color}">${roleInfo.label}</span>
        </td>
        <td class="p-3">
          <div class="flex flex-wrap gap-1">
            ${isAll 
              ? `<span class="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold text-[10px]">Semua Menu (Full Access)</span>` 
              : perms.map(p => `<span class="px-1.5 py-0.5 rounded bg-slate-100 text-slate-700 text-[10px] font-semibold">${p}</span>`).join('')}
          </div>
        </td>
        <td class="p-3 text-center">
          <span class="px-2 py-0.5 rounded text-[10px] font-bold ${u.status === 'Aktif' || !u.status ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'}">
            ${u.status || 'Aktif'}
          </span>
        </td>
        <td class="p-3 text-right">
          <div class="flex items-center justify-end gap-1">
            <button type="button" onclick="openAddUserModal('${u.username}')" class="p-1 rounded text-text-secondary hover:text-primary hover:bg-slate-100" title="Edit Pengguna">
              <span class="material-symbols-outlined text-[18px]">edit</span>
            </button>
            <button type="button" onclick="deleteUser('${u.username}')" class="p-1 rounded text-text-secondary hover:text-rose-600 hover:bg-slate-100" title="Hapus Pengguna">
              <span class="material-symbols-outlined text-[18px]">delete</span>
            </button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function openAddUserModal(username = null) {
  const modal = document.getElementById('modal-user-crud');
  const title = document.getElementById('modal-user-title');
  const userIn = document.getElementById('crud-user-username');
  const namaIn = document.getElementById('crud-user-nama');
  const passIn = document.getElementById('crud-user-password');
  const roleIn = document.getElementById('crud-user-role');

  if (username) {
    const u = (AppState.users || []).find(it => it.username === username);
    if (u) {
      if (title) title.textContent = `Edit Pengguna (${u.username})`;
      if (userIn) { userIn.value = u.username; userIn.disabled = true; }
      if (namaIn) namaIn.value = u.nama || '';
      if (passIn) { passIn.value = ''; passIn.placeholder = 'Kosongkan jika tidak ubah sandi'; }
      if (roleIn) roleIn.value = u.role || 'guru';
      
      const perms = Array.isArray(u.menu_access) ? u.menu_access : (u.menu_access ? String(u.menu_access).split(',') : ['*']);
      document.querySelectorAll('.user-menu-chk').forEach(chk => {
        chk.checked = perms.includes('*') || perms.includes(chk.value);
      });
    }
  } else {
    if (title) title.textContent = 'Tambah Pengguna Baru';
    if (userIn) { userIn.value = ''; userIn.disabled = false; }
    if (namaIn) namaIn.value = '';
    if (passIn) { passIn.value = ''; passIn.placeholder = 'Masukkan kata sandi baru...'; }
    if (roleIn) roleIn.value = 'admin';
    autoSelectMenuCheckboxesByRole('admin');
  }

  if (modal) {
    modal.style.display = 'flex';
    modal.classList.remove('hidden');
  }
}

function closeUserModal() {
  const modal = document.getElementById('modal-user-crud');
  if (modal) {
    modal.style.display = 'none';
    modal.classList.add('hidden');
  }
}

function autoSelectMenuCheckboxesByRole(role) {
  document.querySelectorAll('.user-menu-chk').forEach(chk => {
    if (role === 'admin') chk.checked = true;
    else if (role === 'kepsek') chk.checked = ['dashboard', 'matrix', 'export', 'school'].includes(chk.value);
    else if (role === 'guru') chk.checked = ['matrix', 'availability'].includes(chk.value);
    else chk.checked = ['dashboard', 'master'].includes(chk.value);
  });
}

async function saveUserSubmit() {
  const username = document.getElementById('crud-user-username')?.value?.trim();
  const nama = document.getElementById('crud-user-nama')?.value?.trim();
  const password = document.getElementById('crud-user-password')?.value?.trim();
  const role = document.getElementById('crud-user-role')?.value || 'guru';

  if (!username) return showToast('Username/NIP tidak boleh kosong!', 'warning');
  if (!nama) return showToast('Nama pengguna tidak boleh kosong!', 'warning');

  const selectedMenus = [];
  document.querySelectorAll('.user-menu-chk:checked').forEach(chk => {
    selectedMenus.push(chk.value);
  });

  const existingIdx = (AppState.users || []).findIndex(it => it.username === username);
  if (existingIdx >= 0) {
    const existing = AppState.users[existingIdx];
    existing.nama = nama;
    existing.role = role;
    existing.menu_access = selectedMenus;
    if (password) existing.password_hash = password;
  } else {
    AppState.users.push({
      username: username,
      nama: nama,
      password_hash: password || '123456',
      role: role,
      menu_access: selectedMenus,
      status: 'Aktif'
    });
  }

  saveToLocalStorage();
  closeUserModal();
  renderUsersView();
  showToast(`Pengguna ${username} berhasil disimpan!`, 'success');

  const res = await apiPost('saveUsers', { users: AppState.users });
  if (res && res.success) {
    showToast('Sinkronisasi user ke Google Sheets berhasil!', 'success');
  }
}

async function deleteUser(username) {
  if (username === AppState.currentUser.username) {
    return showToast('Anda tidak dapat menghapus akun Anda sendiri!', 'error');
  }
  if (!confirm(`Hapus pengguna ${username}?`)) return;

  AppState.users = (AppState.users || []).filter(u => u.username !== username);
  saveToLocalStorage();
  renderUsersView();
  showToast(`Pengguna ${username} dihapus`, 'info');

  apiPost('saveUsers', { users: AppState.users });
}

// ============================================================================
// REVISI 2: PENGATURAN WAKTU, DURASI JP, & ISTIRAHAT DINAMIS
// ============================================================================
function parseTimeToMinutes(timeStr) {
  if (!timeStr) return 7 * 60;
  const parts = String(timeStr).split(':').map(Number);
  return (parts[0] || 0) * 60 + (parts[1] || 0);
}

function formatMinutesToTime(totalMins) {
  const h = Math.floor(totalMins / 60) % 24;
  const m = totalMins % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

function generateDayTimeSlots(dayConfig, dayName) {
  const jamMulai = dayConfig.jam_mulai || '07:00';
  const menitPerJp = parseInt(dayConfig.menit_per_jp) || 45;
  const jumlahJp = parseInt(dayConfig.jumlah_jp) || 8;
  const ist1Menit = parseInt(dayConfig.istirahat_1_menit) || 0;
  const ist1Setelah = parseInt(dayConfig.istirahat_1_setelah_jp) || 0;
  const ist2Menit = parseInt(dayConfig.istirahat_2_menit) || 0;
  const ist2Setelah = parseInt(dayConfig.istirahat_2_setelah_jp) || 0;

  let currentMin = parseTimeToMinutes(jamMulai);
  const slots = [];

  for (let jp = 1; jp <= jumlahJp; jp++) {
    const startStr = formatMinutesToTime(currentMin);
    const endStr = formatMinutesToTime(currentMin + menitPerJp);
    currentMin += menitPerJp;

    slots.push({
      hari: dayName,
      jp: jp,
      jam_mulai: startStr,
      jam_selesai: endStr,
      is_break: false,
      nama_slot: `KBM JP ${jp}`
    });

    // Istirahat 1
    if (ist1Menit > 0 && ist1Setelah === jp) {
      const bStart = formatMinutesToTime(currentMin);
      const bEnd = formatMinutesToTime(currentMin + ist1Menit);
      currentMin += ist1Menit;
      slots.push({
        hari: dayName,
        jp: 91,
        jam_mulai: bStart,
        jam_selesai: bEnd,
        is_break: true,
        nama_slot: 'Istirahat 1'
      });
    }

    // Istirahat 2 (Sholat Dzuhur)
    if (ist2Menit > 0 && ist2Setelah === jp) {
      const bStart = formatMinutesToTime(currentMin);
      const bEnd = formatMinutesToTime(currentMin + ist2Menit);
      currentMin += ist2Menit;
      slots.push({
        hari: dayName,
        jp: 92,
        jam_mulai: bStart,
        jam_selesai: bEnd,
        is_break: true,
        nama_slot: 'Istirahat 2 (Dzuhur)'
      });
    }
  }

  return {
    slots: slots,
    jamSelesai: formatMinutesToTime(currentMin)
  };
}

function getDayConfigOrDefault(d) {
  const daily = (AppState.timeConfig && AppState.timeConfig.dailyConfig) || {};
  if (daily[d]) return daily[d];

  // Default cerdas sesuai kebutuhan sekolah
  const isJumat = d === 'Jumat';
  const isSabtu = d === 'Sabtu';

  return {
    jam_mulai: '07:00',
    menit_per_jp: isJumat || isSabtu ? 40 : 45,
    jumlah_jp: isJumat ? 6 : (isSabtu ? 6 : 8),
    istirahat_1_menit: 15,
    istirahat_1_setelah_jp: isJumat ? 4 : 3,
    istirahat_2_menit: isJumat ? 0 : 30,
    istirahat_2_setelah_jp: isJumat ? 0 : 6
  };
}

function renderTimeConfigView() {
  const daysContainer = document.getElementById('time-days-config-container');
  const classLimitsTbody = document.getElementById('class-time-limits-tbody');

  const days = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];

  if (daysContainer) {
    daysContainer.innerHTML = days.map(d => {
      const cfg = getDayConfigOrDefault(d);
      const calc = generateDayTimeSlots(cfg, d);

      const menitOptions = [30, 35, 40, 45, 50].map(m => `
        <option value="${m}" ${parseInt(cfg.menit_per_jp) === m ? 'selected' : ''}>${m} Menit</option>
      `).join('');

      return `
        <div class="p-4 bg-slate-50 border border-border-subtle rounded-2xl space-y-3 transition-all hover:border-primary/40 hover:shadow-sm">
          <div class="flex items-center justify-between border-b border-border-subtle pb-2">
            <div class="flex items-center gap-2">
              <span class="w-3 h-3 rounded-full ${d === 'Jumat' ? 'bg-emerald-500' : 'bg-primary'}"></span>
              <span class="font-black text-sm text-text-primary tracking-wide">${d}</span>
            </div>
            <span id="time-day-badge-${d}" class="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-blue-100 text-blue-900 border border-blue-200">
              ${cfg.jam_mulai} – ${calc.jamSelesai} (${cfg.jumlah_jp} JP @ ${cfg.menit_per_jp}m)
            </span>
          </div>

          <!-- Baris Utama: Jam Mulai, Menit/JP, Jumlah JP, Jam Selesai -->
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
            <div>
              <label class="text-[10px] font-bold text-slate-600 block mb-1">Jam Mulai KBM:</label>
              <input type="time" id="time-start-${d}" value="${cfg.jam_mulai}" onchange="recalcDaySchedule('${d}')" class="w-full text-xs border rounded-lg p-2 bg-white font-bold text-slate-900 border-slate-300">
            </div>

            <div>
              <label class="text-[10px] font-bold text-slate-600 block mb-1">Menit per JP:</label>
              <select id="time-menit-${d}" onchange="recalcDaySchedule('${d}')" class="w-full text-xs border rounded-lg p-2 bg-white font-bold text-slate-900 border-slate-300 cursor-pointer">
                ${menitOptions}
              </select>
            </div>

            <div>
              <label class="text-[10px] font-bold text-slate-600 block mb-1">Jumlah JP:</label>
              <input type="number" id="time-jp-${d}" value="${cfg.jumlah_jp}" min="1" max="10" oninput="recalcDaySchedule('${d}')" class="w-full text-xs border rounded-lg p-2 bg-white font-bold text-center text-slate-900 border-slate-300">
            </div>

            <div>
              <label class="text-[10px] font-bold text-slate-600 block mb-1">Jam Selesai (Auto):</label>
              <input type="time" id="time-end-${d}" value="${calc.jamSelesai}" readonly class="w-full text-xs border rounded-lg p-2 bg-slate-100 font-black text-center text-primary border-slate-300 cursor-not-allowed">
            </div>
          </div>

          <!-- Baris Istirahat: Istirahat I & Istirahat II -->
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2 border-t border-slate-200/60 text-xs">
            <!-- Istirahat 1 -->
            <div class="bg-white p-2.5 rounded-xl border border-slate-200 flex items-center justify-between gap-2">
              <span class="font-bold text-slate-700 text-[11px] whitespace-nowrap">Istirahat I:</span>
              <div class="flex items-center gap-1.5">
                <input type="number" id="time-ist1-dur-${d}" value="${cfg.istirahat_1_menit}" min="0" max="60" oninput="recalcDaySchedule('${d}')" class="w-14 border rounded p-1 text-center font-bold bg-slate-50 border-slate-300">
                <span class="text-[10px] text-slate-500 font-semibold">m, stlh</span>
                <select id="time-ist1-after-${d}" onchange="recalcDaySchedule('${d}')" class="border rounded p-1 text-xs font-bold bg-slate-50 border-slate-300 cursor-pointer">
                  <option value="0" ${cfg.istirahat_1_setelah_jp == 0 ? 'selected' : ''}>-</option>
                  <option value="1" ${cfg.istirahat_1_setelah_jp == 1 ? 'selected' : ''}>JP 1</option>
                  <option value="2" ${cfg.istirahat_1_setelah_jp == 2 ? 'selected' : ''}>JP 2</option>
                  <option value="3" ${cfg.istirahat_1_setelah_jp == 3 ? 'selected' : ''}>JP 3</option>
                  <option value="4" ${cfg.istirahat_1_setelah_jp == 4 ? 'selected' : ''}>JP 4</option>
                  <option value="5" ${cfg.istirahat_1_setelah_jp == 5 ? 'selected' : ''}>JP 5</option>
                </select>
              </div>
            </div>

            <!-- Istirahat 2 -->
            <div class="bg-white p-2.5 rounded-xl border border-slate-200 flex items-center justify-between gap-2">
              <span class="font-bold text-slate-700 text-[11px] whitespace-nowrap">Istirahat II:</span>
              <div class="flex items-center gap-1.5">
                <input type="number" id="time-ist2-dur-${d}" value="${cfg.istirahat_2_menit}" min="0" max="60" oninput="recalcDaySchedule('${d}')" class="w-14 border rounded p-1 text-center font-bold bg-slate-50 border-slate-300">
                <span class="text-[10px] text-slate-500 font-semibold">m, stlh</span>
                <select id="time-ist2-after-${d}" onchange="recalcDaySchedule('${d}')" class="border rounded p-1 text-xs font-bold bg-slate-50 border-slate-300 cursor-pointer">
                  <option value="0" ${cfg.istirahat_2_setelah_jp == 0 ? 'selected' : ''}>-</option>
                  <option value="4" ${cfg.istirahat_2_setelah_jp == 4 ? 'selected' : ''}>JP 4</option>
                  <option value="5" ${cfg.istirahat_2_setelah_jp == 5 ? 'selected' : ''}>JP 5</option>
                  <option value="6" ${cfg.istirahat_2_setelah_jp == 6 ? 'selected' : ''}>JP 6</option>
                  <option value="7" ${cfg.istirahat_2_setelah_jp == 7 ? 'selected' : ''}>JP 7</option>
                </select>
              </div>
            </div>
          </div>

          <!-- Pratinjau Rangkaian Slot Waktu Harian -->
          <div id="time-preview-timeline-${d}" class="text-[10px] text-slate-600 bg-white/80 p-2 rounded-lg border border-slate-200 flex flex-wrap gap-1 items-center">
            ${calc.slots.map(s => `
              <span class="px-1.5 py-0.5 rounded ${s.is_break ? 'bg-amber-100 text-amber-900 font-bold' : 'bg-slate-100 text-slate-700 font-medium'}">
                ${s.is_break ? s.nama_slot : `JP ${s.jp}`}: ${s.jam_mulai}–${s.jam_selesai}
              </span>
            `).join('')}
          </div>
        </div>
      `;
    }).join('');
  }

  if (classLimitsTbody) {
    const classes = AppState.master.kelas || [];
    const classLimits = AppState.timeConfig.classLimits || [];
    const limitMap = {};
    classLimits.forEach(l => { limitMap[l.kelas_id] = l; });

    classLimitsTbody.innerHTML = classes.map(k => {
      const saved = limitMap[k.id] || {};
      const maxJp = saved.max_jp_hari || (k.tingkat <= 2 ? 6 : 7);
      const targetJp = saved.total_jp_minggu || (k.tingkat <= 2 ? 30 : 36);

      return `
        <tr class="hover:bg-slate-50">
          <td class="p-2.5 font-bold text-primary">${k.kode_kelas}</td>
          <td class="p-2.5 font-semibold text-text-secondary">Kelas ${k.tingkat}</td>
          <td class="p-2.5">
            <input type="number" id="limit-max-${k.id}" value="${maxJp}" min="4" max="9" class="w-16 border rounded p-1 text-center font-bold bg-slate-50 border-slate-300">
          </td>
          <td class="p-2.5">
            <input type="number" id="limit-target-${k.id}" value="${targetJp}" min="20" max="45" class="w-16 border rounded p-1 text-center font-bold bg-slate-50 border-slate-300">
          </td>
        </tr>
      `;
    }).join('');
  }
}

function recalcDaySchedule(d) {
  const currentCfg = {
    jam_mulai: document.getElementById(`time-start-${d}`)?.value || '07:00',
    menit_per_jp: parseInt(document.getElementById(`time-menit-${d}`)?.value) || 45,
    jumlah_jp: parseInt(document.getElementById(`time-jp-${d}`)?.value) || 8,
    istirahat_1_menit: parseInt(document.getElementById(`time-ist1-dur-${d}`)?.value) || 0,
    istirahat_1_setelah_jp: parseInt(document.getElementById(`time-ist1-after-${d}`)?.value) || 0,
    istirahat_2_menit: parseInt(document.getElementById(`time-ist2-dur-${d}`)?.value) || 0,
    istirahat_2_setelah_jp: parseInt(document.getElementById(`time-ist2-after-${d}`)?.value) || 0
  };

  const calc = generateDayTimeSlots(currentCfg, d);

  const endInput = document.getElementById(`time-end-${d}`);
  if (endInput) endInput.value = calc.jamSelesai;

  const badge = document.getElementById(`time-day-badge-${d}`);
  if (badge) {
    badge.textContent = `${currentCfg.jam_mulai} – ${calc.jamSelesai} (${currentCfg.jumlah_jp} JP @ ${currentCfg.menit_per_jp}m)`;
  }

  const timeline = document.getElementById(`time-preview-timeline-${d}`);
  if (timeline) {
    timeline.innerHTML = calc.slots.map(s => `
      <span class="px-1.5 py-0.5 rounded ${s.is_break ? 'bg-amber-100 text-amber-900 font-bold' : 'bg-slate-100 text-slate-700 font-medium'}">
        ${s.is_break ? s.nama_slot : `JP ${s.jp}`}: ${s.jam_mulai}–${s.jam_selesai}
      </span>
    `).join('');
  }
}

async function saveTimeConfiguration() {
  const days = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const dailyConfig = {};
  const allGeneratedSlots = [];

  days.forEach(d => {
    const cfg = {
      jam_mulai: document.getElementById(`time-start-${d}`)?.value || '07:00',
      menit_per_jp: parseInt(document.getElementById(`time-menit-${d}`)?.value) || 45,
      jumlah_jp: parseInt(document.getElementById(`time-jp-${d}`)?.value) || 8,
      istirahat_1_menit: parseInt(document.getElementById(`time-ist1-dur-${d}`)?.value) || 0,
      istirahat_1_setelah_jp: parseInt(document.getElementById(`time-ist1-after-${d}`)?.value) || 0,
      istirahat_2_menit: parseInt(document.getElementById(`time-ist2-dur-${d}`)?.value) || 0,
      istirahat_2_setelah_jp: parseInt(document.getElementById(`time-ist2-after-${d}`)?.value) || 0
    };

    const calc = generateDayTimeSlots(cfg, d);
    cfg.jam_selesai = calc.jamSelesai;
    dailyConfig[d] = cfg;
    allGeneratedSlots.push(...calc.slots);
  });

  const classLimits = (AppState.master.kelas || []).map(k => ({
    kelas_id: k.id,
    tingkat: k.tingkat,
    max_jp_hari: parseInt(document.getElementById(`limit-max-${k.id}`)?.value) || 7,
    total_jp_minggu: parseInt(document.getElementById(`limit-target-${k.id}`)?.value) || 36
  }));

  // Update AppState secara instan
  AppState.timeConfig = { dailyConfig, classLimits };
  AppState.master.waktu = allGeneratedSlots;
  saveToLocalStorage();
  showToast('Pengaturan waktu dan batas JP disimpan ke memori', 'info');

  // Perbarui tampilan pratinjau dokumen cetak jika sedang aktif
  renderPrintableDocument();

  const res = await apiPost('saveTimeConfig', { dailyConfig, classLimits, slots: allGeneratedSlots, waktu: allGeneratedSlots });
  if (res && res.success) {
    showToast('Konfigurasi waktu KBM berhasil disinkronkan ke Google Sheets!', 'success');
  }
}

// ============================================================================
// REVISI 5: BEBAN MENGAJAR & CONSTRAINT-SATISFACTION AUTO-GENERATOR
// ============================================================================
function renderBebanView() {
  const tbody = document.getElementById('teaching-assignments-tbody');
  const countBadge = document.getElementById('assignment-count-badge');
  const statAssignments = document.getElementById('stat-total-assignments');
  const statJpDemand = document.getElementById('stat-total-jp-demand');

  const assignments = AppState.beban || [];
  const guruMap = {}; (AppState.master.guru || []).forEach(g => { guruMap[g.id] = g.nama; });
  const mapelMap = {}; (AppState.master.mapel || []).forEach(m => { mapelMap[m.id] = m.nama; });
  const kelasMap = {}; (AppState.master.kelas || []).forEach(k => { kelasMap[k.id] = k.kode_kelas; });

  let totalJp = 0;
  assignments.forEach(a => {
    const kList = Array.isArray(a.kelas_ids) ? a.kelas_ids : String(a.kelas_ids || '').split(',').filter(Boolean);
    totalJp += (parseInt(a.jp_per_minggu) || 0) * kList.length;
  });

  if (countBadge) countBadge.textContent = `${assignments.length} Alokasi Tugas`;
  if (statAssignments) statAssignments.textContent = assignments.length;
  if (statJpDemand) statJpDemand.textContent = `${totalJp} JP`;

  if (!tbody) return;
  if (assignments.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="p-6 text-center text-slate-500">Belum ada alokasi beban mengajar. Klik <strong>Tambah Beban Mengajar</strong> untuk mulai.</td></tr>`;
    return;
  }

  tbody.innerHTML = assignments.map(a => {
    const kList = Array.isArray(a.kelas_ids) ? a.kelas_ids : String(a.kelas_ids || '').split(',').filter(Boolean);
    const kelasPills = kList.map(kid => `<span class="px-2 py-0.5 rounded bg-blue-100 text-blue-800 font-bold text-[10px] mr-1">${kelasMap[kid] || kid}</span>`).join('');
    const totalRowJp = (parseInt(a.jp_per_minggu) || 0) * kList.length;

    return `
      <tr class="hover:bg-slate-50 transition-colors">
        <td class="p-3 font-bold text-text-primary">${guruMap[a.guru_id] || a.guru_id}</td>
        <td class="p-3 font-medium text-text-primary">${mapelMap[a.mapel_id] || a.mapel_id}</td>
        <td class="p-3">${kelasPills}</td>
        <td class="p-3 text-center font-bold text-text-primary">${a.jp_per_minggu} JP</td>
        <td class="p-3 text-center font-extrabold text-primary">${totalRowJp} JP</td>
        <td class="p-3 text-text-secondary text-[11px]">${a.keterangan || '-'}</td>
        <td class="p-3 text-right">
          <button type="button" onclick="deleteAssignmentItem('${a.id}')" class="p-1 rounded text-text-secondary hover:text-rose-600 hover:bg-slate-100 cursor-pointer" title="Hapus Alokasi">
            <span class="material-symbols-outlined text-[18px]">delete</span>
          </button>
        </td>
      </tr>
    `;
  }).join('');
}

function openAddAssignmentModal(id = null) {
  const modal = document.getElementById('modal-assignment-crud');
  const guruSel = document.getElementById('assign-select-guru');
  const mapelSel = document.getElementById('assign-select-mapel');
  const kelasContainer = document.getElementById('assign-kelas-checkboxes');

  if (guruSel) {
    guruSel.innerHTML = (AppState.master.guru || []).map(g => `<option value="${g.id}">${g.nama} (${g.mapel_utama})</option>`).join('');
  }
  if (mapelSel) {
    mapelSel.innerHTML = (AppState.master.mapel || []).map(m => `<option value="${m.id}">${m.nama} (${m.kode})</option>`).join('');
  }
  if (kelasContainer) {
    kelasContainer.innerHTML = (AppState.master.kelas || []).map(k => `
      <label class="flex items-center gap-1.5 p-1 rounded hover:bg-white cursor-pointer">
        <input type="checkbox" value="${k.id}" class="assign-kelas-chk" onchange="updateAssignmentTotalPreview()">
        <span class="font-bold text-text-primary">${k.kode_kelas}</span>
      </label>
    `).join('');
  }

  updateAssignmentTotalPreview();

  if (modal) {
    modal.style.display = 'flex';
    modal.classList.remove('hidden');
  }
}

function closeAssignmentModal() {
  const modal = document.getElementById('modal-assignment-crud');
  if (modal) {
    modal.style.display = 'none';
    modal.classList.add('hidden');
  }
}

function updateAssignmentTotalPreview() {
  const chks = document.querySelectorAll('.assign-kelas-chk:checked');
  const jp = parseInt(document.getElementById('assign-jp-per-class')?.value) || 4;
  const preview = document.getElementById('assign-total-preview');
  if (preview) {
    preview.textContent = `${chks.length * jp} JP / Minggu (${chks.length} Kelas)`;
  }
}

async function saveAssignmentItem() {
  const guruId = document.getElementById('assign-select-guru')?.value;
  const mapelId = document.getElementById('assign-select-mapel')?.value;
  const jp = parseInt(document.getElementById('assign-jp-per-class')?.value) || 4;
  const ket = document.getElementById('assign-keterangan')?.value?.trim() || '';

  const selectedClasses = [];
  document.querySelectorAll('.assign-kelas-chk:checked').forEach(chk => {
    selectedClasses.push(chk.value);
  });

  if (selectedClasses.length === 0) {
    return showToast('Pilih minimal satu rombel / kelas!', 'warning');
  }

  const newAssignment = {
    id: `BM_${Date.now()}`,
    guru_id: guruId,
    mapel_id: mapelId,
    kelas_ids: selectedClasses,
    jp_per_minggu: jp,
    keterangan: ket
  };

  AppState.beban.push(newAssignment);
  saveToLocalStorage();
  closeAssignmentModal();
  renderBebanView();
  showToast('Alokasi beban mengajar berhasil ditambahkan!', 'success');

  apiPost('saveTeachingAssignments', { assignments: AppState.beban });
}

async function deleteAssignmentItem(id) {
  if (!confirm('Hapus alokasi beban mengajar ini?')) return;
  AppState.beban = (AppState.beban || []).filter(a => a.id !== id);
  saveToLocalStorage();
  renderBebanView();
  showToast('Alokasi dihapus', 'info');

  apiPost('saveTeachingAssignments', { assignments: AppState.beban });
}

// ----------------------------------------------------------------------------
// ALGORITMA CONSTRAINT-SATISFACTION AUTO-SCHEDULE GENERATOR
// ----------------------------------------------------------------------------
function runAutoScheduleGenerator() {
  const assignments = AppState.beban || [];
  if (assignments.length === 0) {
    return showToast('Tambahkan beban mengajar guru terlebih dahulu sebelum generate!', 'warning');
  }

  showToast('Menjalankan algoritma auto-generate jadwal anti bentrok...', 'info');

  const days = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const masterWaktu = AppState.master.waktu || [];
  const requests = AppState.requests || [];

  // Hitung slot KBM valid per hari
  const validSlotsByDay = {};
  days.forEach(d => {
    validSlotsByDay[d] = masterWaktu
      .filter(w => w.hari === d && !w.is_break && parseInt(w.jp) > 0)
      .map(w => parseInt(w.jp));
    if (validSlotsByDay[d].length === 0) {
      const defaultJp = d === 'Jumat' ? 5 : d === 'Sabtu' ? 6 : 7;
      validSlotsByDay[d] = Array.from({ length: defaultJp }, (_, i) => i + 1);
    }
  });

  // Ketersediaan Guru Map (Blocked slots)
  const teacherBlockedMap = {};
  requests.forEach(r => {
    if (r.status === 'Disetujui' || !r.status) {
      const key = `${r.guru_id}_${r.hari}_${r.jp}`;
      teacherBlockedMap[key] = true;
    }
  });

  // Buat daftar Task: pecah setiap assignment menjadi blok KBM (prioritaskan 2 JP)
  const tasks = [];
  assignments.forEach(a => {
    const kList = Array.isArray(a.kelas_ids) ? a.kelas_ids : String(a.kelas_ids || '').split(',').filter(Boolean);
    kList.forEach(kId => {
      let remainingJp = parseInt(a.jp_per_minggu) || 2;
      while (remainingJp > 0) {
        const blockSize = (remainingJp >= 2 && remainingJp !== 3) ? 2 : (remainingJp === 3 ? 3 : 1);
        tasks.push({
          assignmentId: a.id,
          guruId: a.guru_id,
          mapelId: a.mapel_id,
          kelasId: kId,
          blockSize: blockSize
        });
        remainingJp -= blockSize;
      }
    });
  });

  // Urutkan task: guru dengan blok terbanyak atau paling terkendala diletakkan duluan
  tasks.sort((a, b) => b.blockSize - a.blockSize);

  // State Pelacakan Slot
  const occupiedTeacher = {}; // `${hari}_${jp}_${guruId}` -> true
  const occupiedClass = {};   // `${hari}_${jp}_${kelasId}` -> true
  const classDayJpCount = {}; // `${hari}_${kelasId}` -> count
  const generatedSchedule = [];
  const unplacedTasks = [];

  tasks.forEach(task => {
    let placed = false;

    // Cari hari dan JP terbaik yang memuat blockSize berturutan
    for (const day of days) {
      if (placed) break;
      const slots = validSlotsByDay[day];
      const classDayKey = `${day}_${task.kelasId}`;
      const currentDayJp = classDayJpCount[classDayKey] || 0;

      // Batas JP per hari (maks 7 atau 8)
      if (currentDayJp + task.blockSize > 8) continue;

      for (let i = 0; i <= slots.length - task.blockSize; i++) {
        const blockJps = slots.slice(i, i + task.blockSize);

        // Periksa apakah seluruh JP dalam blok ini aman
        const isBlockValid = blockJps.every(jp => {
          const tKey = `${day}_${jp}_${task.guruId}`;
          const cKey = `${day}_${jp}_${task.kelasId}`;
          const isBlocked = teacherBlockedMap[tKey];
          return !isBlocked && !occupiedTeacher[tKey] && !occupiedClass[cKey];
        });

        if (isBlockValid) {
          // Tempatkan KBM
          blockJps.forEach(jp => {
            const tKey = `${day}_${jp}_${task.guruId}`;
            const cKey = `${day}_${jp}_${task.kelasId}`;
            occupiedTeacher[tKey] = true;
            occupiedClass[cKey] = true;

            generatedSchedule.push({
              id: `GEN_${task.kelasId}_${day}_${jp}`,
              hari: day,
              jp: jp,
              kelas_id: task.kelasId,
              guru_id: task.guruId,
              mapel_id: task.mapelId,
              ruang_id: `R_${task.kelasId}`
            });
          });

          classDayJpCount[classDayKey] = currentDayJp + task.blockSize;
          placed = true;
          break;
        }
      }
    }

    if (!placed) {
      unplacedTasks.push(task);
    }
  });

  AppState._generatedScheduleTemp = generatedSchedule;

  // Tampilkan Modal Hasil Generate & Rekomendasi
  showGenerateResultModal(tasks.length, generatedSchedule.length, unplacedTasks);
}

function showGenerateResultModal(totalTasks, placedSlots, unplacedTasks) {
  const modal = document.getElementById('modal-generate-result');
  const percentEl = document.getElementById('gen-result-percent');
  const slotsEl = document.getElementById('gen-result-slots');
  const unplacedContainer = document.getElementById('gen-unplaced-container');
  const unplacedList = document.getElementById('gen-unplaced-list');
  const headerEl = document.getElementById('gen-modal-header');

  const totalDemandJp = placedSlots + unplacedTasks.reduce((sum, t) => sum + t.blockSize, 0);
  const successPercent = totalDemandJp > 0 ? Math.round((placedSlots / totalDemandJp) * 100) : 100;

  if (percentEl) percentEl.textContent = `${successPercent}%`;
  if (slotsEl) slotsEl.textContent = `${placedSlots} / ${totalDemandJp} JP`;

  const guruMap = {}; (AppState.master.guru || []).forEach(g => { guruMap[g.id] = g.nama; });
  const mapelMap = {}; (AppState.master.mapel || []).forEach(m => { mapelMap[m.id] = m.nama; });
  const kelasMap = {}; (AppState.master.kelas || []).forEach(k => { kelasMap[k.id] = k.kode_kelas; });

  if (unplacedTasks.length > 0) {
    if (headerEl) headerEl.className = 'bg-amber-600 px-6 py-4 text-white flex items-center justify-between';
    if (unplacedContainer) unplacedContainer.style.display = 'block';
    if (unplacedList) {
      unplacedList.innerHTML = unplacedTasks.map(t => `
        <div class="p-2.5 rounded-lg bg-rose-50 border border-rose-200 text-rose-900 text-xs">
          <strong>${guruMap[t.guruId] || t.guruId}</strong> - ${mapelMap[t.mapelId] || t.mapelId} (${kelasMap[t.kelasId] || t.kelasId}) @ ${t.blockSize} JP
          <div class="text-[11px] text-text-secondary mt-0.5">Rekomendasi: Tambah kesediaan hari guru pengampu atau sesuaikan batas JP harian rombel.</div>
        </div>
      `).join('');
    }
  } else {
    if (headerEl) headerEl.className = 'bg-emerald-600 px-6 py-4 text-white flex items-center justify-between';
    if (unplacedContainer) unplacedContainer.style.display = 'none';
  }

  if (modal) {
    modal.style.display = 'flex';
    modal.classList.remove('hidden');
  }
}

function closeGenerateResultModal() {
  const modal = document.getElementById('modal-generate-result');
  if (modal) {
    modal.style.display = 'none';
    modal.classList.add('hidden');
  }
}

function applyGeneratedScheduleToMatrix() {
  if (!AppState._generatedScheduleTemp || AppState._generatedScheduleTemp.length === 0) {
    return showToast('Belum ada jadwal yang di-generate!', 'warning');
  }

  AppState.jadwal = AppState._generatedScheduleTemp;
  saveToLocalStorage();
  calculateClashes();
  closeGenerateResultModal();
  navigateView('matrix');
  renderMatrixTable();
  showToast('Jadwal hasil generator otomatis berhasil diterapkan ke Matriks KBM!', 'success');
}

// ============================================================================
// 13. PRINT PREVIEW & DOCUMENT EXPORT
// ============================================================================
function renderExportView() {
  populateExportClassDropdown();
  renderPrintableDocument();
}

function populateExportClassDropdown() {
  const sel = document.getElementById('export-class-select');
  if (!sel) return;
  const classes = AppState.master.kelas || [];
  const currVal = sel.value || 'ALL';
  
  sel.innerHTML = `<option value="ALL">Semua Kelas</option>` +
    classes.map(k => `<option value="${k.id}" ${k.id === currVal ? 'selected' : ''}>${k.nama_rombel || k.kode_kelas}</option>`).join('');
}

function renderPrintableDocument() {
  const container = document.getElementById('printable-table-container');
  if (!container) return;

  const masterKelas = AppState.master.kelas || [];
  const masterMapel = AppState.master.mapel || [];
  const masterGuru = AppState.master.guru || [];
  const jadwal = AppState.jadwal || [];

  const mapelMap = {}; masterMapel.forEach(m => { mapelMap[m.id] = m; });
  const guruMap = {}; masterGuru.forEach(g => { guruMap[g.id] = g; });

  // Buat mapping kode huruf guru: (A), (B), (C), dst. sesuai standar jadwal kedinasan
  const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
  const teacherCodeMap = {};
  masterGuru.forEach((g, idx) => {
    teacherCodeMap[g.id] = g.kode || letters[idx % letters.length] || `G${idx + 1}`;
  });

  const filterClass = document.getElementById('export-class-select')?.value || 'ALL';
  const targetClasses = (filterClass === 'ALL') ? masterKelas : masterKelas.filter(k => k.id === filterClass);

  if (targetClasses.length === 0) {
    container.innerHTML = `<div class="p-6 text-center text-slate-500 font-semibold">Tidak ada data kelas yang dapat ditampilkan.</div>`;
    return;
  }

  // Helper waktu dinamis yang terintegrasi dengan Menu Waktu dan Jam Belajar
  const getSlotTime = (dayName, jpNum) => {
    const slot = (AppState.master.waktu || []).find(w => w.hari === dayName && parseInt(w.jp) === jpNum && !w.is_break);
    if (slot && slot.jam_mulai && slot.jam_selesai) {
      return `${slot.jam_mulai.replace(':', '.')} – ${slot.jam_selesai.replace(':', '.')}`;
    }
    const cfg = getDayConfigOrDefault(dayName);
    const calc = generateDayTimeSlots(cfg, dayName);
    const genSlot = calc.slots.find(s => parseInt(s.jp) === jpNum && !s.is_break);
    if (genSlot) return `${genSlot.jam_mulai.replace(':', '.')} – ${genSlot.jam_selesai.replace(':', '.')}`;
    return '';
  };

  const getBreakTime = (dayName, breakIdx) => {
    const breaks = (AppState.master.waktu || []).filter(w => w.hari === dayName && w.is_break);
    const b = breaks[breakIdx - 1];
    if (b && b.jam_mulai && b.jam_selesai) {
      return `${b.jam_mulai.replace(':', '.')} – ${b.jam_selesai.replace(':', '.')}`;
    }
    const cfg = getDayConfigOrDefault(dayName);
    const calc = generateDayTimeSlots(cfg, dayName);
    const genBreaks = calc.slots.filter(s => s.is_break);
    const gb = genBreaks[breakIdx - 1];
    if (gb) return `${gb.jam_mulai.replace(':', '.')} – ${gb.jam_selesai.replace(':', '.')}`;
    return '';
  };

  // Helper untuk mendapatkan teks isi slot ({KODE_GURU}) {MAPEL}
  const getSlot = (kId, day, jp) => {
    const item = jadwal.find(it => it.hari === day && parseInt(it.jp) === jp && it.kelas_id === kId);
    if (!item) return '';
    const tCode = teacherCodeMap[item.guru_id] || '';
    const m = mapelMap[item.mapel_id];
    const mName = m ? (m.nama || m.kode) : item.mapel_id;
    return tCode ? `(${tCode}) ${mName}` : mName;
  };

  let html = '';

  targetClasses.forEach((k, kIdx) => {
    const className = k.nama_rombel || k.kode_kelas;
    html += `
      <div class="class-schedule-block mb-8 page-break-inside-avoid">
        <div class="text-sm font-black text-slate-900 tracking-wide uppercase mb-1.5 flex items-center justify-between">
          <span>KELAS ${className}</span>
          <span class="text-xs text-slate-500 font-bold no-print">Rombel ${k.kode_kelas}</span>
        </div>

        <table class="table-schedule-preview">
          <thead>
            <tr>
              <th style="width: 45px;">JAM</th>
              <th style="width: 105px;">WAKTU<br>PELAJARAN</th>
              <th style="width: 110px;">SENIN</th>
              <th style="width: 110px;">SELASA</th>
              <th style="width: 110px;">RABU</th>
              <th style="width: 110px;">KAMIS</th>
              <th style="width: 105px;">WAKTU<br>PELAJARAN</th>
              <th style="width: 110px;">JUM'AT</th>
              <th style="width: 110px;">SABTU</th>
            </tr>
          </thead>
          <tbody>
            <!-- Row 1: JAM I -->
            <tr>
              <td class="text-center font-bold">I</td>
              <td class="text-center font-medium">${getSlotTime('Senin', 1) || '07.00 – 07.45'}</td>
              <td class="text-center">${getSlot(k.id, 'Senin', 1)}</td>
              <td class="text-center">${getSlot(k.id, 'Selasa', 1)}</td>
              <td class="text-center">${getSlot(k.id, 'Rabu', 1)}</td>
              <td class="text-center">${getSlot(k.id, 'Kamis', 1)}</td>
              <td class="text-center font-medium">${getSlotTime('Jumat', 1) || '07.00 – 07.40'}</td>
              <td class="text-center">${getSlot(k.id, 'Jumat', 1)}</td>
              <td class="text-center">${getSlot(k.id, 'Sabtu', 1) || 'Istighosah'}</td>
            </tr>

            <!-- Row 2: JAM II -->
            <tr>
              <td class="text-center font-bold">II</td>
              <td class="text-center font-medium">${getSlotTime('Senin', 2) || '07.45 – 08.30'}</td>
              <td class="text-center">${getSlot(k.id, 'Senin', 2)}</td>
              <td class="text-center">${getSlot(k.id, 'Selasa', 2)}</td>
              <td class="text-center">${getSlot(k.id, 'Rabu', 2)}</td>
              <td class="text-center">${getSlot(k.id, 'Kamis', 2)}</td>
              <td class="text-center font-medium">${getSlotTime('Jumat', 2) || '07.40 – 08.20'}</td>
              <td class="text-center">${getSlot(k.id, 'Jumat', 2)}</td>
              <td class="text-center">${getSlot(k.id, 'Sabtu', 2)}</td>
            </tr>

            <!-- Row 3: JAM III -->
            <tr>
              <td class="text-center font-bold">III</td>
              <td class="text-center font-medium">${getSlotTime('Senin', 3) || '08.30 – 09.15'}</td>
              <td class="text-center">${getSlot(k.id, 'Senin', 3)}</td>
              <td class="text-center">${getSlot(k.id, 'Selasa', 3)}</td>
              <td class="text-center">${getSlot(k.id, 'Rabu', 3)}</td>
              <td class="text-center">${getSlot(k.id, 'Kamis', 3)}</td>
              <td class="text-center font-medium">${getSlotTime('Jumat', 3) || '08.20 – 09.00'}</td>
              <td class="text-center">${getSlot(k.id, 'Jumat', 3)}</td>
              <td class="text-center">${getSlot(k.id, 'Sabtu', 3)}</td>
            </tr>

            <!-- Row 4: Istirahat Senin-Kamis & Sabtu | JP 4 Jum'at -->
            <tr>
              <td class="text-center font-bold"></td>
              <td class="text-center font-medium">${getBreakTime('Senin', 1) || '09.15 – 09.30'}</td>
              <td colspan="4" class="cell-break">I s t i r a h a t</td>
              <td class="text-center font-medium">${getSlotTime('Jumat', 4) || '09.00 – 09.40'}</td>
              <td class="text-center">${getSlot(k.id, 'Jumat', 4)}</td>
              <td class="cell-break">Istirahat</td>
            </tr>

            <!-- Row 5: JAM IV Senin-Kamis & Sabtu | Istirahat Jum'at -->
            <tr>
              <td class="text-center font-bold">IV</td>
              <td class="text-center font-medium">${getSlotTime('Senin', 4) || '09.30 – 10.15'}</td>
              <td class="text-center">${getSlot(k.id, 'Senin', 4)}</td>
              <td class="text-center">${getSlot(k.id, 'Selasa', 4)}</td>
              <td class="text-center">${getSlot(k.id, 'Rabu', 4)}</td>
              <td class="text-center">${getSlot(k.id, 'Kamis', 4)}</td>
              <td class="text-center font-medium">${getBreakTime('Jumat', 1) || '09.40 – 09.55'}</td>
              <td class="cell-break">Istirahat</td>
              <td class="text-center">${getSlot(k.id, 'Sabtu', 4)}</td>
            </tr>

            <!-- Row 6: JAM V -->
            <tr>
              <td class="text-center font-bold">V</td>
              <td class="text-center font-medium">${getSlotTime('Senin', 5) || '10.15 – 11.00'}</td>
              <td class="text-center">${getSlot(k.id, 'Senin', 5)}</td>
              <td class="text-center">${getSlot(k.id, 'Selasa', 5)}</td>
              <td class="text-center">${getSlot(k.id, 'Rabu', 5)}</td>
              <td class="text-center">${getSlot(k.id, 'Kamis', 5)}</td>
              <td class="text-center font-medium">${getSlotTime('Jumat', 5) || '09.55 – 10.35'}</td>
              <td class="text-center">${getSlot(k.id, 'Jumat', 5)}</td>
              <td class="text-center">${getSlot(k.id, 'Sabtu', 5)}</td>
            </tr>

            <!-- Row 7: JAM VI -->
            <tr>
              <td class="text-center font-bold">VI</td>
              <td class="text-center font-medium">${getSlotTime('Senin', 6) || '11.00 – 11.45'}</td>
              <td class="text-center">${getSlot(k.id, 'Senin', 6)}</td>
              <td class="text-center">${getSlot(k.id, 'Selasa', 6)}</td>
              <td class="text-center">${getSlot(k.id, 'Rabu', 6)}</td>
              <td class="text-center">${getSlot(k.id, 'Kamis', 6)}</td>
              <td class="text-center font-medium">${getSlotTime('Jumat', 6) || '10.35 – 11.15'}</td>
              <td class="text-center">${getSlot(k.id, 'Jumat', 6)}</td>
              <td class="text-center">${getSlot(k.id, 'Sabtu', 6)}</td>
            </tr>

            <!-- Row 8: Jama'ah Sholat Dzuhur -->
            <tr>
              <td class="text-center font-bold"></td>
              <td class="text-center font-medium">${getBreakTime('Senin', 2) || '11.45 – 12.15'}</td>
              <td colspan="4" class="cell-dzuhur">Jama'ah Sholat Dzuhur</td>
              <td class="text-center font-medium"></td>
              <td class="text-center"></td>
              <td class="cell-dzuhur">Sholat Dzuhur</td>
            </tr>

            <!-- Row 9: JAM VII -->
            <tr>
              <td class="text-center font-bold">VII</td>
              <td class="text-center font-medium">${getSlotTime('Senin', 7) || '12.15 – 13.00'}</td>
              <td class="text-center">${getSlot(k.id, 'Senin', 7)}</td>
              <td class="text-center">${getSlot(k.id, 'Selasa', 7)}</td>
              <td class="text-center">${getSlot(k.id, 'Rabu', 7)}</td>
              <td class="text-center">${getSlot(k.id, 'Kamis', 7)}</td>
              <td class="text-center font-medium"></td>
              <td class="text-center"></td>
              <td class="text-center"></td>
            </tr>

            <!-- Row 10: JAM VIII -->
            <tr>
              <td class="text-center font-bold">VIII</td>
              <td class="text-center font-medium">${getSlotTime('Senin', 8) || '13.00 – 13.45'}</td>
              <td class="text-center">${getSlot(k.id, 'Senin', 8)}</td>
              <td class="text-center">${getSlot(k.id, 'Selasa', 8)}</td>
              <td class="text-center">${getSlot(k.id, 'Rabu', 8)}</td>
              <td class="text-center">${getSlot(k.id, 'Kamis', 8)}</td>
              <td class="text-center font-medium"></td>
              <td class="text-center"></td>
              <td class="text-center"></td>
            </tr>
          </tbody>
        </table>
      </div>
    `;
  });

  // Tambahkan Keterangan Kode Guru Pengampu di bawah tabel
  html += `
    <div class="mt-6 pt-4 border-t-2 border-slate-900 text-[11px] text-slate-800 page-break-inside-avoid">
      <div class="font-black text-xs uppercase mb-2">Keterangan Kode Guru Pengampu:</div>
      <div class="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-x-4 gap-y-1">
        ${masterGuru.map(g => `
          <div>
            <span class="font-extrabold text-blue-900">(${teacherCodeMap[g.id]})</span>
            <span class="font-medium">${g.nama}${g.gelar ? ', ' + g.gelar : ''}</span>
            <span class="text-slate-500 text-[10px]">(${g.mapel_utama || '-'})</span>
          </div>
        `).join('')}
      </div>
    </div>
  `;

  container.innerHTML = html;
}

function exportToExcel() {
  try {
    if (typeof XLSX === 'undefined') return showToast('Library SheetJS sedang dimuat...', 'warning');
    const container = document.getElementById('printable-table-container');
    const tables = container ? container.querySelectorAll('table') : [];
    if (!tables || tables.length === 0) return showToast('Tabel pratinjau belum siap', 'error');

    const wb = XLSX.utils.book_new();

    tables.forEach((tbl, idx) => {
      const parentBlock = tbl.closest('.class-schedule-block');
      const titleEl = parentBlock ? parentBlock.querySelector('.text-sm') : null;
      let sheetName = titleEl ? titleEl.textContent.trim().replace(/[\\/?*[\]:]/g, '') : `Kelas_${idx + 1}`;
      sheetName = sheetName.replace(/^KELAS\s*/i, 'Kelas ').substring(0, 31);
      
      const ws = XLSX.utils.table_to_sheet(tbl);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    XLSX.writeFile(wb, "Jadwal_Pelajaran_Resmi_Format_Sekolah.xlsx");
    showToast('File Excel (.xlsx) berhasil diunduh sesuai format jadwal!', 'success');
  } catch (err) {
    showToast('Gagal ekspor Excel: ' + err.message, 'error');
  }
}

function exportToWord() {
  try {
    const docContent = document.getElementById('printable-document')?.innerHTML;
    if (!docContent) return showToast('Dokumen belum siap', 'error');

    const header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>"+
      "<head><meta charset='utf-8'><title>Jadwal Pelajaran Resmi</title>"+
      "<style>body{font-family: Arial, sans-serif;} table{border-collapse:collapse;width:100%;margin-bottom:20px;} td,th{border:1px solid #000;padding:5px;font-size:9pt;text-align:center;} .cell-break{background-color:#c0c0c0;font-weight:bold;letter-spacing:2px;} .cell-dzuhur{background-color:#808080;color:#fff;font-weight:bold;} </style>"+
      "</head><body>";
    const footer = "</body></html>";
    const sourceHTML = header + docContent + footer;

    const source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
    const fileDownload = document.createElement("a");
    document.body.appendChild(fileDownload);
    fileDownload.href = source;
    fileDownload.download = 'Jadwal_Pelajaran_Resmi.doc';
    fileDownload.click();
    document.body.removeChild(fileDownload);
    showToast('Dokumen Word (.doc) berhasil diunduh', 'success');
  } catch (err) {
    showToast('Gagal ekspor Word: ' + err.message, 'error');
  }
}

async function savePdfToDriveFromBrowser() {
  showToast('Memproses arsip ke Google Drive...', 'info');
  const docContent = document.getElementById('printable-document')?.innerHTML || '';

  const res = await apiPost('archiveDrive', { htmlContent: docContent });
  if (res && res.success) {
    showToast('Berhasil diarsipkan di Google Drive!', 'success');
  } else {
    showToast('Gagal arsip Drive: ' + (res ? res.message : 'Error jaringan'), 'error');
  }
}

// ============================================================================
// 14. APPROVAL & SINKRONISASI SERVER
// ============================================================================
function openApprovalModal() {
  const modal = document.getElementById('modal-approval');
  const clashStat = document.getElementById('approval-clash-status');
  if (clashStat) {
    if (AppState.clashes.length > 0) {
      clashStat.textContent = `${AppState.clashes.length} Bentrok (Perlu diselesaikan sebelum pengesahan)`;
      clashStat.className = 'font-bold text-rose-600';
    } else {
      clashStat.textContent = '0 Bentrok (Sistem Aman Terverifikasi)';
      clashStat.className = 'font-bold text-emerald-600';
    }
  }
  if (modal) {
    modal.style.display = 'flex';
    modal.classList.remove('hidden');
  }
}

function closeApprovalModal() {
  const modal = document.getElementById('modal-approval');
  if (modal) {
    modal.style.display = 'none';
    modal.classList.add('hidden');
  }
}

async function submitScheduleApproval(status) {
  closeApprovalModal();
  showToast('Menyimpan pengesahan jadwal ke Google Sheets...', 'info');

  const pill = document.getElementById('dashboard-approval-pill');
  if (pill) {
    pill.textContent = 'Telah Disahkan Kepala Sekolah';
    pill.className = 'px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800';
  }

  const res = await apiPost('approveSchedule', { status: status });
  if (res && res.success) {
    showToast('Jadwal Pelajaran resmi disahkan oleh Kepala Sekolah!', 'success');
  } else {
    showToast('Pengesahan tersimpan di browser (offline)', 'warning');
  }
}

async function saveMatrixChangesToServer() {
  showToast('Menyimpan batch jadwal ke Google Sheets via REST API...', 'info');
  const icon = document.getElementById('sync-icon');
  if (icon) icon.classList.add('animate-spin');

  const res = await apiPost('saveSchedule', AppState.jadwal);
  if (icon) icon.classList.remove('animate-spin');

  if (res && res.success) {
    showToast(res.message || 'Database Google Sheets berhasil disinkronkan!', 'success');
  } else {
    showToast('Tersimpan di penyimpanan lokal (' + (res ? res.message : 'Offline') + ')', 'warning');
  }
}

async function saveAvailabilityToServer() {
  showToast('Menyimpan data ketersediaan ke server...', 'info');
  const res = await apiPost('saveAvailability', AppState.requests);
  if (res && res.success) {
    showToast('Ketersediaan waktu guru berhasil disimpan ke server!', 'success');
  } else {
    showToast('Ketersediaan guru tersimpan di memori lokal', 'info');
  }
}

function syncDataToServer() {
  saveMatrixChangesToServer();
}

async function fetchServerData() {
  const connPill = document.getElementById('connection-status-text');
  const res = await apiGet('getData');

  if (res && res.success && res.data) {
    const data = res.data;
    if (data.guru && data.guru.length > 0) AppState.master.guru = data.guru;
    if (data.kelas && data.kelas.length > 0) AppState.master.kelas = data.kelas;
    if (data.mapel && data.mapel.length > 0) AppState.master.mapel = data.mapel;
    if (data.ruang && data.ruang.length > 0) AppState.master.ruang = data.ruang;
    if (data.waktu && data.waktu.length > 0) AppState.master.waktu = data.waktu;
    if (data.jadwal && data.jadwal.length > 0) AppState.jadwal = data.jadwal;
    if (data.request && data.request.length > 0) AppState.requests = data.request;
    if (data.log && data.log.length > 0) AppState.auditTrail = data.log;
    if (data.user && data.user.length > 0) AppState.users = data.user;
    if (data.beban && data.beban.length > 0) AppState.beban = data.beban;
    if (data.waktu_kelas && data.waktu_kelas.length > 0) AppState.timeConfig.classLimits = data.waktu_kelas;
    if (data.sekolah && data.sekolah.length > 0) {
      AppState.master.sekolah = data.sekolah;
      loadSchoolIdentity();
    }

    saveToLocalStorage();
    populateLoginTeacherDropdown();
    calculateClashes();

    if (AppState.isLoggedIn) renderCurrentView();
    if (connPill) connPill.textContent = 'REST API Online';
  } else {
    if (connPill) connPill.textContent = 'Offline (Cache Aktif)';
  }
}

// ============================================================================
// 15. STORAGE LOCAL PERSISTENCE
// ============================================================================
function saveToLocalStorage() {
  SafeStorage.set(APP_CONFIG.STORAGE_KEY_DATA, JSON.stringify({
    master: AppState.master,
    jadwal: AppState.jadwal,
    requests: AppState.requests,
    auditTrail: AppState.auditTrail,
    users: AppState.users,
    beban: AppState.beban,
    timeConfig: AppState.timeConfig,
    schoolLogoBase64: AppState.schoolLogoBase64,
    savedAt: Date.now()
  }));
}

function loadFromLocalStorage() {
  const raw = SafeStorage.get(APP_CONFIG.STORAGE_KEY_DATA);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (parsed.master && parsed.master.guru && parsed.master.guru.length > 0) {
        AppState.master = parsed.master;
        AppState.jadwal = parsed.jadwal || [];
        AppState.requests = parsed.requests || [];
        AppState.auditTrail = parsed.auditTrail || [];
        AppState.users = parsed.users || [];
        AppState.beban = parsed.beban || [];
        AppState.timeConfig = parsed.timeConfig || { days: {}, classLimits: [] };
        AppState.schoolLogoBase64 = parsed.schoolLogoBase64 || null;
        if (AppState.users.length === 0) seedDefaultUsersAndBeban();
        return;
      }
    } catch (e) {}
  }
  seedDefaultLocalState();
}

function seedDefaultUsersAndBeban() {
  AppState.users = [
    {
      username: '19791104 200501 1 003',
      nama: 'Bambang Sugiantoro, S.Pd.SD',
      password_hash: '123456',
      role: 'admin',
      menu_access: ['dashboard', 'beban', 'matrix', 'availability', 'timeconfig', 'master', 'export', 'school', 'users'],
      status: 'Aktif'
    },
    {
      username: '19740512 199803 2 004',
      nama: 'Dra. Sri Mulyani, M.Pd',
      password_hash: '123456',
      role: 'kepsek',
      menu_access: ['dashboard', 'matrix', 'export', 'school'],
      status: 'Aktif'
    },
    {
      username: '19820315 200604 1 008',
      nama: 'Ahmad Fauzi, S.Pd.SD',
      password_hash: '123456',
      role: 'guru',
      menu_access: ['matrix', 'availability'],
      status: 'Aktif'
    }
  ];

  AppState.beban = [
    {
      id: 'BM_01',
      guru_id: 'G02', // Ahmad Fauzi
      mapel_id: 'M05', // Matematika
      kelas_ids: ['K07', 'K08'], // 4A, 4B
      jp_per_minggu: 4,
      keterangan: 'Matematika Fase B'
    },
    {
      id: 'BM_02',
      guru_id: 'G02', // Ahmad Fauzi
      mapel_id: 'M06', // IPAS
      kelas_ids: ['K09', 'K10'], // 5A, 5B
      jp_per_minggu: 2,
      keterangan: 'IPAS Terapan Fase C'
    },
    {
      id: 'BM_03',
      guru_id: 'G14', // Joko Widodo
      mapel_id: 'M01', // PAI
      kelas_ids: ['K01', 'K02', 'K07', 'K08'],
      jp_per_minggu: 3,
      keterangan: 'PAI & Budi Pekerti'
    },
    {
      id: 'BM_04',
      guru_id: 'G16', // Budi Santoso
      mapel_id: 'M07', // PJOK
      kelas_ids: ['K07', 'K08', 'K09', 'K10'],
      jp_per_minggu: 3,
      keterangan: 'PJOK Lapangan'
    }
  ];

  AppState.timeConfig = {
    dailyConfig: {
      Senin: { jam_mulai: '07:00', menit_per_jp: 45, jumlah_jp: 8, istirahat_1_menit: 15, istirahat_1_setelah_jp: 3, istirahat_2_menit: 30, istirahat_2_setelah_jp: 6 },
      Selasa: { jam_mulai: '07:00', menit_per_jp: 45, jumlah_jp: 8, istirahat_1_menit: 15, istirahat_1_setelah_jp: 3, istirahat_2_menit: 30, istirahat_2_setelah_jp: 6 },
      Rabu: { jam_mulai: '07:00', menit_per_jp: 45, jumlah_jp: 8, istirahat_1_menit: 15, istirahat_1_setelah_jp: 3, istirahat_2_menit: 30, istirahat_2_setelah_jp: 6 },
      Kamis: { jam_mulai: '07:00', menit_per_jp: 45, jumlah_jp: 8, istirahat_1_menit: 15, istirahat_1_setelah_jp: 3, istirahat_2_menit: 30, istirahat_2_setelah_jp: 6 },
      Jumat: { jam_mulai: '07:00', menit_per_jp: 40, jumlah_jp: 6, istirahat_1_menit: 15, istirahat_1_setelah_jp: 4, istirahat_2_menit: 0, istirahat_2_setelah_jp: 0 },
      Sabtu: { jam_mulai: '07:00', menit_per_jp: 40, jumlah_jp: 6, istirahat_1_menit: 15, istirahat_1_setelah_jp: 3, istirahat_2_menit: 30, istirahat_2_setelah_jp: 6 }
    },
    days: {},
    classLimits: (AppState.master.kelas || []).map(k => ({
      kelas_id: k.id,
      tingkat: k.tingkat,
      max_jp_hari: k.tingkat <= 2 ? 6 : 7,
      total_jp_minggu: k.tingkat <= 2 ? 30 : 36
    }))
  };
}

function seedDefaultLocalState() {
  AppState.master.waktu = [
    { hari: 'Senin', jp: 0, jam_mulai: '07:00', jam_selesai: '07:35', is_break: true, nama_slot: 'Upacara Bendera' },
    { hari: 'Senin', jp: 1, jam_mulai: '07:35', jam_selesai: '08:10', is_break: false, nama_slot: 'KBM JP 1' },
    { hari: 'Senin', jp: 2, jam_mulai: '08:10', jam_selesai: '08:45', is_break: false, nama_slot: 'KBM JP 2' },
    { hari: 'Senin', jp: 3, jam_mulai: '08:45', jam_selesai: '09:20', is_break: false, nama_slot: 'KBM JP 3' },
    { hari: 'Senin', jp: 99, jam_mulai: '09:20', jam_selesai: '09:40', is_break: true, nama_slot: 'Istirahat 1' },
    { hari: 'Senin', jp: 4, jam_mulai: '09:40', jam_selesai: '10:15', is_break: false, nama_slot: 'KBM JP 4' },
    { hari: 'Senin', jp: 5, jam_mulai: '10:15', jam_selesai: '10:50', is_break: false, nama_slot: 'KBM JP 5' },
    { hari: 'Senin', jp: 6, jam_mulai: '10:50', jam_selesai: '11:25', is_break: false, nama_slot: 'KBM JP 6' },
    { hari: 'Senin', jp: 7, jam_mulai: '11:25', jam_selesai: '12:00', is_break: false, nama_slot: 'KBM JP 7' },

    { hari: 'Selasa', jp: 1, jam_mulai: '07:00', jam_selesai: '07:35', is_break: false, nama_slot: 'KBM JP 1' },
    { hari: 'Selasa', jp: 2, jam_mulai: '07:35', jam_selesai: '08:10', is_break: false, nama_slot: 'KBM JP 2' },
    { hari: 'Selasa', jp: 3, jam_mulai: '08:10', jam_selesai: '08:45', is_break: false, nama_slot: 'KBM JP 3' },
    { hari: 'Selasa', jp: 4, jam_mulai: '08:45', jam_selesai: '09:20', is_break: false, nama_slot: 'KBM JP 4' },
    { hari: 'Selasa', jp: 99, jam_mulai: '09:20', jam_selesai: '09:40', is_break: true, nama_slot: 'Istirahat 1' },
    { hari: 'Selasa', jp: 5, jam_mulai: '09:40', jam_selesai: '10:15', is_break: false, nama_slot: 'KBM JP 5' },
    { hari: 'Selasa', jp: 6, jam_mulai: '10:15', jam_selesai: '10:50', is_break: false, nama_slot: 'KBM JP 6' },
    { hari: 'Selasa', jp: 7, jam_mulai: '10:50', jam_selesai: '11:25', is_break: false, nama_slot: 'KBM JP 7' },
    { hari: 'Selasa', jp: 8, jam_mulai: '11:25', jam_selesai: '12:00', is_break: false, nama_slot: 'KBM JP 8' },

    { hari: 'Rabu', jp: 1, jam_mulai: '07:00', jam_selesai: '07:35', is_break: false, nama_slot: 'KBM JP 1' },
    { hari: 'Rabu', jp: 2, jam_mulai: '07:35', jam_selesai: '08:10', is_break: false, nama_slot: 'KBM JP 2' },
    { hari: 'Rabu', jp: 3, jam_mulai: '08:10', jam_selesai: '08:45', is_break: false, nama_slot: 'KBM JP 3' },
    { hari: 'Rabu', jp: 99, jam_mulai: '08:45', jam_selesai: '09:05', is_break: true, nama_slot: 'Istirahat 1' },
    { hari: 'Rabu', jp: 4, jam_mulai: '09:05', jam_selesai: '09:40', is_break: false, nama_slot: 'KBM JP 4' },
    { hari: 'Rabu', jp: 5, jam_mulai: '09:40', jam_selesai: '10:15', is_break: false, nama_slot: 'KBM JP 5' },
    { hari: 'Rabu', jp: 6, jam_mulai: '10:15', jam_selesai: '10:50', is_break: false, nama_slot: 'KBM JP 6' },
    { hari: 'Rabu', jp: 7, jam_mulai: '10:50', jam_selesai: '11:25', is_break: false, nama_slot: 'KBM JP 7' },

    { hari: 'Kamis', jp: 1, jam_mulai: '07:00', jam_selesai: '07:35', is_break: false, nama_slot: 'KBM JP 1' },
    { hari: 'Kamis', jp: 2, jam_mulai: '07:35', jam_selesai: '08:10', is_break: false, nama_slot: 'KBM JP 2' },
    { hari: 'Kamis', jp: 3, jam_mulai: '08:10', jam_selesai: '08:45', is_break: false, nama_slot: 'KBM JP 3' },
    { hari: 'Kamis', jp: 99, jam_mulai: '08:45', jam_selesai: '09:05', is_break: true, nama_slot: 'Istirahat 1' },
    { hari: 'Kamis', jp: 4, jam_mulai: '09:05', jam_selesai: '09:40', is_break: false, nama_slot: 'KBM JP 4' },
    { hari: 'Kamis', jp: 5, jam_mulai: '09:40', jam_selesai: '10:15', is_break: false, nama_slot: 'KBM JP 5' },
    { hari: 'Kamis', jp: 6, jam_mulai: '10:15', jam_selesai: '10:50', is_break: false, nama_slot: 'KBM JP 6' },
    { hari: 'Kamis', jp: 7, jam_mulai: '10:50', jam_selesai: '11:25', is_break: false, nama_slot: 'KBM JP 7' },

    { hari: 'Jumat', jp: 1, jam_mulai: '07:00', jam_selesai: '07:35', is_break: false, nama_slot: 'KBM JP 1' },
    { hari: 'Jumat', jp: 2, jam_mulai: '07:35', jam_selesai: '08:10', is_break: false, nama_slot: 'KBM JP 2' },
    { hari: 'Jumat', jp: 3, jam_mulai: '08:10', jam_selesai: '08:45', is_break: false, nama_slot: 'KBM JP 3' },
    { hari: 'Jumat', jp: 99, jam_mulai: '08:45', jam_selesai: '09:05', is_break: true, nama_slot: 'Istirahat 1' },
    { hari: 'Jumat', jp: 4, jam_mulai: '09:05', jam_selesai: '09:40', is_break: false, nama_slot: 'KBM JP 4' },
    { hari: 'Jumat', jp: 5, jam_mulai: '09:40', jam_selesai: '10:15', is_break: false, nama_slot: 'KBM JP 5' },

    { hari: 'Sabtu', jp: 1, jam_mulai: '07:00', jam_selesai: '07:35', is_break: false, nama_slot: 'KBM JP 1' },
    { hari: 'Sabtu', jp: 2, jam_mulai: '07:35', jam_selesai: '08:10', is_break: false, nama_slot: 'KBM JP 2' },
    { hari: 'Sabtu', jp: 3, jam_mulai: '08:10', jam_selesai: '08:45', is_break: false, nama_slot: 'KBM JP 3' },
    { hari: 'Sabtu', jp: 99, jam_mulai: '08:45', jam_selesai: '09:05', is_break: true, nama_slot: 'Istirahat 1' },
    { hari: 'Sabtu', jp: 4, jam_mulai: '09:05', jam_selesai: '09:40', is_break: false, nama_slot: 'KBM JP 4' },
    { hari: 'Sabtu', jp: 5, jam_mulai: '09:40', jam_selesai: '10:15', is_break: false, nama_slot: 'KBM JP 5' },
    { hari: 'Sabtu', jp: 6, jam_mulai: '10:15', jam_selesai: '10:50', is_break: false, nama_slot: 'KBM JP 6' }
  ];

  AppState.master.guru = [
    { id: 'G01', nip: '19740512 199803 2 004', nama: 'Dra. Sri Mulyani', gelar: 'M.Pd', mapel_utama: 'Kepala Sekolah', target_jtm: 0, status_kepegawaian: 'PNS' },
    { id: 'G02', nip: '19820315 200604 1 008', nama: 'Ahmad Fauzi', gelar: 'S.Pd.SD', mapel_utama: 'Guru Kelas 4', target_jtm: 24, status_kepegawaian: 'PNS' },
    { id: 'G03', nip: '19850720 200902 2 005', nama: 'Siti Rahmawati', gelar: 'S.Pd.SD', mapel_utama: 'Guru Kelas 5', target_jtm: 24, status_kepegawaian: 'PNS' },
    { id: 'G04', nip: '19791104 200501 1 003', nama: 'Bambang Sugiantoro', gelar: 'S.Pd.SD', mapel_utama: 'Guru Kelas 6', target_jtm: 24, status_kepegawaian: 'PNS' },
    { id: 'G05', nip: '19880110 201101 2 012', nama: 'Dewi Lestari', gelar: 'S.Pd.SD', mapel_utama: 'Guru Kelas 1', target_jtm: 24, status_kepegawaian: 'PNS' },
    { id: 'G06', nip: '19900418 201402 2 007', nama: 'Nur Hidayah', gelar: 'S.Pd.SD', mapel_utama: 'Guru Kelas 2', target_jtm: 24, status_kepegawaian: 'P3K' },
    { id: 'G07', nip: '19860925 201001 1 009', nama: 'Eko Prasetyo', gelar: 'S.Pd.SD', mapel_utama: 'Guru Kelas 3', target_jtm: 24, status_kepegawaian: 'PNS' },
    { id: 'G08', nip: '19920814 201903 2 015', nama: 'Rina Anggraini', gelar: 'S.Pd.SD', mapel_utama: 'Guru Kelas 1', target_jtm: 24, status_kepegawaian: 'P3K' },
    { id: 'G09', nip: '19910522 201801 1 006', nama: 'Wahyu Hidayat', gelar: 'S.Pd.SD', mapel_utama: 'Guru Kelas 2', target_jtm: 24, status_kepegawaian: 'P3K' },
    { id: 'G10', nip: '19871203 201101 2 010', nama: 'Tri Astuti', gelar: 'S.Pd.SD', mapel_utama: 'Guru Kelas 3', target_jtm: 24, status_kepegawaian: 'PNS' },
    { id: 'G11', nip: '19930219 202012 1 004', nama: 'Dimas Pratama', gelar: 'S.Pd.SD', mapel_utama: 'Guru Kelas 4', target_jtm: 24, status_kepegawaian: 'P3K' },
    { id: 'G12', nip: '19890630 201503 2 008', nama: 'Endang Susilowati', gelar: 'S.Pd.SD', mapel_utama: 'Guru Kelas 5', target_jtm: 24, status_kepegawaian: 'PNS' },
    { id: 'G13', nip: '19841015 200801 1 007', nama: 'Agus Setiawan', gelar: 'S.Pd.SD', mapel_utama: 'Guru Kelas 6', target_jtm: 24, status_kepegawaian: 'PNS' },
    { id: 'G14', nip: '19810302 200604 1 011', nama: 'Joko Widodo', gelar: 'S.Pd.I', mapel_utama: 'Pendidikan Agama Islam', target_jtm: 24, status_kepegawaian: 'PNS' },
    { id: 'G15', nip: '19941128 202203 2 009', nama: 'Lestari Handayani', gelar: 'S.Pd.I', mapel_utama: 'Pendidikan Agama Islam', target_jtm: 24, status_kepegawaian: 'P3K' },
    { id: 'G16', nip: '19830817 200801 1 014', nama: 'Budi Santoso', gelar: 'S.Pd.Kor', mapel_utama: 'PJOK', target_jtm: 24, status_kepegawaian: 'PNS' },
    { id: 'G17', nip: '19950312 202302 1 005', nama: 'Rizky Kurniawan', gelar: 'S.Pd', mapel_utama: 'PJOK', target_jtm: 24, status_kepegawaian: 'Honorer' },
    { id: 'G18', nip: '19930704 202012 2 011', nama: 'Maya Indah', gelar: 'S.Pd', mapel_utama: 'Bahasa Inggris', target_jtm: 20, status_kepegawaian: 'Honorer' },
    { id: 'G19', nip: '19900919 201701 2 006', nama: 'Dian Kusuma', gelar: 'S.Pd', mapel_utama: 'Seni Rupa', target_jtm: 18, status_kepegawaian: 'Honorer' },
    { id: 'G20', nip: '19870108 201202 1 003', nama: 'Hadi Sucipto', gelar: 'S.Pd.B.Jawa', mapel_utama: 'Bahasa Jawa', target_jtm: 24, status_kepegawaian: 'PNS' },
    { id: 'G21', nip: '19960224 202401 2 002', nama: 'Anisa Fitriani', gelar: 'S.Pd', mapel_utama: 'Agama Kristen', target_jtm: 12, status_kepegawaian: 'Honorer' },
    { id: 'G22', nip: '19911011 201902 1 008', nama: 'Fajar Ramadhan', gelar: 'S.Kom', mapel_utama: 'Informatika', target_jtm: 18, status_kepegawaian: 'Honorer' },
    { id: 'G23', nip: '19860416 201101 2 009', nama: 'Ratna Juwita', gelar: 'S.Pd', mapel_utama: 'Pendamping Khusus', target_jtm: 18, status_kepegawaian: 'P3K' },
    { id: 'G24', nip: '19970829 202403 1 001', nama: 'Bayu Aji', gelar: 'S.Pd', mapel_utama: 'Pengayaan MTK', target_jtm: 16, status_kepegawaian: 'Honorer' }
  ];

  AppState.master.kelas = [
    { id: 'K01', kode_kelas: '1A', nama_rombel: 'Kelas 1A (Mawar)', tingkat: 1, wali_kelas_id: 'G05', total_siswa: 28 },
    { id: 'K02', kode_kelas: '1B', nama_rombel: 'Kelas 1B (Melati)', tingkat: 1, wali_kelas_id: 'G08', total_siswa: 27 },
    { id: 'K03', kode_kelas: '2A', nama_rombel: 'Kelas 2A (Dahlia)', tingkat: 2, wali_kelas_id: 'G06', total_siswa: 29 },
    { id: 'K04', kode_kelas: '2B', nama_rombel: 'Kelas 2B (Cempaka)', tingkat: 2, wali_kelas_id: 'G09', total_siswa: 28 },
    { id: 'K05', kode_kelas: '3A', nama_rombel: 'Kelas 3A (Kenanga)', tingkat: 3, wali_kelas_id: 'G07', total_siswa: 30 },
    { id: 'K06', kode_kelas: '3B', nama_rombel: 'Kelas 3B (Anggrek)', tingkat: 3, wali_kelas_id: 'G10', total_siswa: 29 },
    { id: 'K07', kode_kelas: '4A', nama_rombel: 'Kelas 4A (Garuda)', tingkat: 4, wali_kelas_id: 'G02', total_siswa: 31 },
    { id: 'K08', kode_kelas: '4B', nama_rombel: 'Kelas 4B (Rajawali)', tingkat: 4, wali_kelas_id: 'G11', total_siswa: 30 },
    { id: 'K09', kode_kelas: '5A', nama_rombel: 'Kelas 5A (Merak)', tingkat: 5, wali_kelas_id: 'G03', total_siswa: 32 },
    { id: 'K10', kode_kelas: '5B', nama_rombel: 'Kelas 5B (Kasuari)', tingkat: 5, wali_kelas_id: 'G12', total_siswa: 31 },
    { id: 'K11', kode_kelas: '6A', nama_rombel: 'Kelas 6A (Elang)', tingkat: 6, wali_kelas_id: 'G04', total_siswa: 32 },
    { id: 'K12', kode_kelas: '6B', nama_rombel: 'Kelas 6B (Alap-alap)', tingkat: 6, wali_kelas_id: 'G13', total_siswa: 31 }
  ];

  AppState.master.mapel = [
    { id: 'M01', kode: 'PAI', nama: 'Pendidikan Agama Islam', kategori: 'Wajib', alokasi_jp_minggu: 3 },
    { id: 'M02', kode: 'PAK', nama: 'Pendidikan Agama Kristen', kategori: 'Wajib', alokasi_jp_minggu: 3 },
    { id: 'M03', kode: 'PPKN', nama: 'Pendidikan Pancasila', kategori: 'Wajib', alokasi_jp_minggu: 4 },
    { id: 'M04', kode: 'BIND', nama: 'Bahasa Indonesia', kategori: 'Wajib', alokasi_jp_minggu: 6 },
    { id: 'M05', kode: 'MTK', nama: 'Matematika', kategori: 'Wajib', alokasi_jp_minggu: 5 },
    { id: 'M06', kode: 'IPAS', nama: 'Ilmu Pengetahuan Alam & Sosial', kategori: 'Wajib', alokasi_jp_minggu: 5 },
    { id: 'M07', kode: 'PJOK', nama: 'PJOK', kategori: 'Wajib', alokasi_jp_minggu: 3 },
    { id: 'M08', kode: 'SRUP', nama: 'Seni Rupa & Prakarya', kategori: 'Pilihan', alokasi_jp_minggu: 3 },
    { id: 'M09', kode: 'BING', nama: 'Bahasa Inggris', kategori: 'Pilihan', alokasi_jp_minggu: 2 },
    { id: 'M10', kode: 'BJAW', nama: 'Bahasa Jawa (Mulok)', kategori: 'Mulok', alokasi_jp_minggu: 2 },
    { id: 'M11', kode: 'KMPT', nama: 'Informatika', kategori: 'Pilihan', alokasi_jp_minggu: 2 }
  ];

  AppState.master.ruang = [
    { id: 'R_01', kode: 'RK-1A', nama: 'Ruang Kelas 1A', jenis: 'Kelas', kapasitas: 32, fasilitas: 'Papan Tulis, Smart TV, AC' },
    { id: 'R_07', kode: 'RK-4A', nama: 'Ruang Kelas 4A', jenis: 'Kelas', kapasitas: 32, fasilitas: 'Papan Tulis, Proyektor' },
    { id: 'R_09', kode: 'RK-5A', nama: 'Ruang Kelas 5A', jenis: 'Kelas', kapasitas: 32, fasilitas: 'Papan Tulis, Proyektor' },
    { id: 'R_11', kode: 'RK-6A', nama: 'Ruang Kelas 6A', jenis: 'Kelas', kapasitas: 32, fasilitas: 'Papan Tulis, Smart TV' },
    { id: 'R_13', kode: 'LAB-KOMP', nama: 'Lab Komputer', jenis: 'Laboratorium', kapasitas: 36, fasilitas: '32 Unit PC, AC' },
    { id: 'R_14', kode: 'LAP-OLR', nama: 'Lapangan Olahraga', jenis: 'Terbuka', kapasitas: 100, fasilitas: 'Gawang, Ring Basket' },
    { id: 'R_15', kode: 'PERPUS', nama: 'Perpustakaan', jenis: 'Khusus', kapasitas: 40, fasilitas: 'Meja Baca, AC' }
  ];

  AppState.requests = [
    { id: 'REQ_01', guru_id: 'G14', hari: 'Kamis', jp: 5, alasan: 'SDN Pilangkenceng 2', status: 'Disetujui' },
    { id: 'REQ_02', guru_id: 'G14', hari: 'Kamis', jp: 6, alasan: 'SDN Pilangkenceng 2', status: 'Disetujui' }
  ];

  AppState.jadwal = [
    { id: 'J_4A_SEN_1', hari: 'Senin', jp: 1, kelas_id: 'K07', guru_id: 'G02', mapel_id: 'M04', ruang_id: 'R_07' },
    { id: 'J_4A_SEN_2', hari: 'Senin', jp: 2, kelas_id: 'K07', guru_id: 'G02', mapel_id: 'M04', ruang_id: 'R_07' },
    { id: 'J_4A_SEN_3', hari: 'Senin', jp: 3, kelas_id: 'K07', guru_id: 'G02', mapel_id: 'M05', ruang_id: 'R_07' },
    { id: 'J_4A_SEN_4', hari: 'Senin', jp: 4, kelas_id: 'K07', guru_id: 'G14', mapel_id: 'M01', ruang_id: 'R_07' },
    { id: 'J_4A_SEN_5', hari: 'Senin', jp: 5, kelas_id: 'K07', guru_id: 'G14', mapel_id: 'M01', ruang_id: 'R_07' },
    { id: 'J_4A_SEN_6', hari: 'Senin', jp: 6, kelas_id: 'K07', guru_id: 'G20', mapel_id: 'M10', ruang_id: 'R_07' },
    { id: 'J_4A_SEN_7', hari: 'Senin', jp: 7, kelas_id: 'K07', guru_id: 'G20', mapel_id: 'M10', ruang_id: 'R_07' },

    { id: 'J_5A_SEN_1', hari: 'Senin', jp: 1, kelas_id: 'K09', guru_id: 'G03', mapel_id: 'M05', ruang_id: 'R_09' },
    { id: 'J_5A_SEN_2', hari: 'Senin', jp: 2, kelas_id: 'K09', guru_id: 'G03', mapel_id: 'M05', ruang_id: 'R_09' },
    { id: 'J_5A_SEN_3', hari: 'Senin', jp: 3, kelas_id: 'K09', guru_id: 'G03', mapel_id: 'M04', ruang_id: 'R_09' },
    { id: 'J_5A_SEN_4', hari: 'Senin', jp: 4, kelas_id: 'K09', guru_id: 'G16', mapel_id: 'M07', ruang_id: 'R_14' },
    { id: 'J_5A_SEN_5', hari: 'Senin', jp: 5, kelas_id: 'K09', guru_id: 'G16', mapel_id: 'M07', ruang_id: 'R_14' },
    { id: 'J_5A_SEN_6', hari: 'Senin', jp: 6, kelas_id: 'K09', guru_id: 'G16', mapel_id: 'M07', ruang_id: 'R_14' },
    { id: 'J_5A_SEN_7', hari: 'Senin', jp: 7, kelas_id: 'K09', guru_id: 'G03', mapel_id: 'M03', ruang_id: 'R_09' },

    { id: 'J_6A_SEN_1', hari: 'Senin', jp: 1, kelas_id: 'K11', guru_id: 'G04', mapel_id: 'M06', ruang_id: 'R_11' },
    { id: 'J_6A_SEN_2', hari: 'Senin', jp: 2, kelas_id: 'K11', guru_id: 'G04', mapel_id: 'M06', ruang_id: 'R_11' },
    { id: 'J_6A_SEN_3', hari: 'Senin', jp: 3, kelas_id: 'K11', guru_id: 'G18', mapel_id: 'M09', ruang_id: 'R_11' },
    { id: 'J_6A_SEN_4', hari: 'Senin', jp: 4, kelas_id: 'K11', guru_id: 'G18', mapel_id: 'M09', ruang_id: 'R_11' },
    { id: 'J_6A_SEN_5', hari: 'Senin', jp: 5, kelas_id: 'K11', guru_id: 'G04', mapel_id: 'M05', ruang_id: 'R_11' },
    { id: 'J_6A_SEN_6', hari: 'Senin', jp: 6, kelas_id: 'K11', guru_id: 'G04', mapel_id: 'M05', ruang_id: 'R_11' },
    { id: 'J_6A_SEN_7', hari: 'Senin', jp: 7, kelas_id: 'K11', guru_id: 'G04', mapel_id: 'M03', ruang_id: 'R_11' }
  ];

  saveToLocalStorage();
}

// ============================================================================
// 16. TOAST, DRAWER & GLOBAL HELPERS
// ============================================================================
function showToast(message, type = 'info') {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  const bgColors = {
    success: 'bg-emerald-600 text-white',
    error: 'bg-rose-600 text-white',
    warning: 'bg-amber-600 text-white',
    info: 'bg-slate-800 text-white'
  };
  const icons = {
    success: 'check_circle',
    error: 'error',
    warning: 'warning',
    info: 'info'
  };

  toast.className = `flex items-center gap-2.5 px-4 py-2.5 rounded-xl shadow-xl text-xs font-bold pointer-events-auto transition-all animate-toast ${bgColors[type] || bgColors.info}`;
  toast.innerHTML = `
    <span class="material-symbols-outlined text-[18px]">${icons[type] || 'info'}</span>
    <span>${message}</span>
  `;

  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(-10px)';
    setTimeout(() => { toast.remove(); }, 250);
  }, 3200);
}

function toggleMobileSidebar() {
  const sidebar = document.getElementById('sidebar');
  const backdrop = document.getElementById('mobile-backdrop');
  if (!sidebar) return;

  const isClosed = sidebar.classList.contains('-translate-x-full');
  if (isClosed) {
    sidebar.classList.remove('-translate-x-full');
    if (backdrop) {
      backdrop.style.display = 'block';
      backdrop.classList.remove('hidden');
    }
  } else {
    sidebar.classList.add('-translate-x-full');
    if (backdrop) {
      backdrop.style.display = 'none';
      backdrop.classList.add('hidden');
    }
  }
}

function closeMobileSidebar() {
  const sidebar = document.getElementById('sidebar');
  const backdrop = document.getElementById('mobile-backdrop');
  if (sidebar) sidebar.classList.add('-translate-x-full');
  if (backdrop) {
    backdrop.style.display = 'none';
    backdrop.classList.add('hidden');
  }
}

function handleGlobalSearch(keyword) {
  if (!keyword || keyword.trim() === '') return;
  const kw = keyword.toLowerCase().trim();

  const matchedClass = (AppState.master.kelas || []).find(k => k.nama_rombel.toLowerCase().includes(kw) || k.kode_kelas.toLowerCase().includes(kw));
  if (matchedClass) {
    AppState.selectedClassId = matchedClass.id;
    navigateView('matrix');
    setMatrixMode('single-class');
    showToast(`Menampilkan jadwal ${matchedClass.nama_rombel}`, 'info');
    return;
  }

  const matchedTeacher = (AppState.master.guru || []).find(g => g.nama.toLowerCase().includes(kw));
  if (matchedTeacher) {
    AppState.selectedTeacherId = matchedTeacher.id;
    navigateView('matrix');
    setMatrixMode('teacher');
    showToast(`Menampilkan jadwal ${matchedTeacher.nama}`, 'info');
    return;
  }
}

// ============================================================================
// 17. EKSPLISIT WINDOW BINDING
// ============================================================================
function registerGlobalHandlers() {
  window.navigateView = navigateView;
  window.selectLoginRole = selectLoginRole;
  window.updateLoginTeacherCredentials = updateLoginTeacherCredentials;
  window.handleLoginSubmit = handleLoginSubmit;
  window.quickLogin = quickLogin;
  window.logoutUser = logoutUser;
  window.switchRole = switchRole;
  window.toggleMobileSidebar = toggleMobileSidebar;
  window.closeMobileSidebar = closeMobileSidebar;
  window.setMatrixMode = setMatrixMode;
  window.setSelectedDay = setSelectedDay;
  window.renderMatrixTable = renderMatrixTable;
  window.openPlotModal = openPlotModal;
  window.closePlotModal = closePlotModal;
  window.checkModalSlotClash = checkModalSlotClash;
  window.confirmSlotPlotting = confirmSlotPlotting;
  window.deleteCurrentSlotAssignment = deleteCurrentSlotAssignment;
  window.openClashDetailModal = openClashDetailModal;
  window.closeClashDetailModal = closeClashDetailModal;
  window.loadTeacherAvailabilityMatrix = loadTeacherAvailabilityMatrix;
  window.toggleTeacherSlotAvailability = toggleTeacherSlotAvailability;
  window.clearCurrentTeacherBlocks = clearCurrentTeacherBlocks;
  window.applyMitraTemplate = applyMitraTemplate;
  window.saveAvailabilityToServer = saveAvailabilityToServer;
  window.setMasterTab = setMasterTab;
  window.openAddMasterModal = openAddMasterModal;
  window.closeMasterModal = closeMasterModal;
  window.editMasterItem = editMasterItem;
  window.handleMasterSubmit = handleMasterSubmit;
  window.renderPrintableDocument = renderPrintableDocument;
  window.exportToExcel = exportToExcel;
  window.exportToWord = exportToWord;
  window.savePdfToDriveFromBrowser = savePdfToDriveFromBrowser;
  window.openApprovalModal = openApprovalModal;
  window.closeApprovalModal = closeApprovalModal;
  window.submitScheduleApproval = submitScheduleApproval;
  window.saveMatrixChangesToServer = saveMatrixChangesToServer;
  window.syncDataToServer = syncDataToServer;
  window.handleGlobalSearch = handleGlobalSearch;
  window.showToast = showToast;
  window.triggerPWAInstall = triggerPWAInstall;

  // Revisi 1: Identitas Sekolah & Logo
  window.renderSchoolView = renderSchoolView;
  window.loadSchoolIdentity = loadSchoolIdentity;
  window.handleSchoolLogoUpload = handleSchoolLogoUpload;
  window.saveSchoolIdentityForm = saveSchoolIdentityForm;

  // Revisi 2: Pengaturan User (RBAC)
  window.renderUsersView = renderUsersView;
  window.openAddUserModal = openAddUserModal;
  window.closeUserModal = closeUserModal;
  window.autoSelectMenuCheckboxesByRole = autoSelectMenuCheckboxesByRole;
  window.saveUserSubmit = saveUserSubmit;
  window.deleteUser = deleteUser;

  // Revisi 3: Import, Template & Bulk Delete Master
  window.toggleSelectAllMaster = toggleSelectAllMaster;
  window.toggleMasterItemSelection = toggleMasterItemSelection;
  window.deleteSelectedMasterItems = deleteSelectedMasterItems;
  window.downloadMasterTemplate = downloadMasterTemplate;
  window.handleMasterFileSelect = handleMasterFileSelect;
  window.openImportPreviewModal = openImportPreviewModal;
  window.closeImportPreviewModal = closeImportPreviewModal;
  window.confirmImportData = confirmImportData;

  // Revisi 4: Pengaturan Waktu & JP
  window.renderTimeConfigView = renderTimeConfigView;
  window.recalcDaySchedule = recalcDaySchedule;
  window.saveTimeConfiguration = saveTimeConfiguration;

  // Revisi 5: Beban Mengajar & Auto-Generator
  window.renderBebanView = renderBebanView;
  window.openAddAssignmentModal = openAddAssignmentModal;
  window.closeAssignmentModal = closeAssignmentModal;
  window.updateAssignmentTotalPreview = updateAssignmentTotalPreview;
  window.saveAssignmentItem = saveAssignmentItem;
  window.deleteAssignmentItem = deleteAssignmentItem;
  window.runAutoScheduleGenerator = runAutoScheduleGenerator;
  window.closeGenerateResultModal = closeGenerateResultModal;
  window.applyGeneratedScheduleToMatrix = applyGeneratedScheduleToMatrix;
}

registerGlobalHandlers();
