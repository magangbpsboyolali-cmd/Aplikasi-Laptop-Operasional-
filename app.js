// ==========================================
// APLIKASI LAPTOP OPERASIONAL - MAIN APPLICATION
// BPS Kabupaten Boyolali
// ==========================================

// ==========================================
// GLOBAL STATE
// ==========================================
const AppState = {
    laptops: [],
    pegawai: [],
    peminjaman: [],
    pengembalian: [],
    keperluan: [],
    kodeAkses: [],
    authenticated: false,
    syncing: false,
    loaded: false,
    refreshTimer: null
};

const MAX_BORROW_DAYS = 5;

// ==========================================
// INITIALIZATION
// ==========================================
document.addEventListener('DOMContentLoaded', function () {
    setCurrentDate();
    initForms();
    showLoginPage();

    // Hash navigation on load
    var hash = window.location.hash.replace('#', '');
    if (hash && hash !== 'dashboard') showPage(hash);

    // Proactive load
    showLoading('Menyinkronkan data...');
    loadData();
});

function setCurrentDate() {
    const el = document.getElementById('currentDate');
    if (el) {
        const now = new Date();
        const options = { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' };
        el.textContent = now.toLocaleDateString('id-ID', options);
    }
    // Set default date for forms
    const today = new Date().toISOString().split('T')[0];
    const pinjamTgl = document.getElementById('pinjamTglPinjam');
    const pinjamTglKembali = document.getElementById('pinjamTglKembali');
    const kembaliTgl = document.getElementById('kembaliTglRealisasi');
    if (pinjamTgl) pinjamTgl.value = today;
    if (pinjamTglKembali) pinjamTglKembali.value = today;
    if (kembaliTgl) kembaliTgl.value = today;

    applyPeminjamanDateLimit();
}

function formatDateInputValue(dateObj) {
    if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) return '';
    return dateObj.toISOString().split('T')[0];
}

function parseDateValue(dateStr) {
    if (!dateStr) return null;
    var raw = String(dateStr).trim();
    if (!raw) return null;

    // Support dd/MM/yyyy format.
    var slashMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (slashMatch) {
        var day = slashMatch[1].padStart(2, '0');
        var month = slashMatch[2].padStart(2, '0');
        var year = slashMatch[3];
        var fromSlash = new Date(year + '-' + month + '-' + day + 'T00:00:00');
        return isNaN(fromSlash.getTime()) ? null : fromSlash;
    }

    var normalized = raw.split('T')[0];
    var parsed = new Date(normalized + 'T00:00:00');
    return isNaN(parsed.getTime()) ? null : parsed;
}

function addDaysToDate(dateStr, days) {
    const d = parseDateValue(dateStr);
    if (!d || isNaN(d.getTime())) return null;
    d.setDate(d.getDate() + days);
    return d;
}

function getBorrowDurationDays(startDateStr, endDateStr) {
    const start = parseDateValue(startDateStr);
    const end = parseDateValue(endDateStr);
    if (!start || !end) return NaN;

    const msPerDay = 24 * 60 * 60 * 1000;

    // FIX: pakai floor biar tidak kelebihan hari
    return Math.floor((end - start) / msPerDay) + 1;
}


function normalizeDate(d) {
    if (!d) return null;
    const newDate = new Date(d);
    newDate.setHours(0, 0, 0, 0);
    return newDate;
}

function applyPeminjamanDateLimit() {
    const tglPinjamEl = document.getElementById('pinjamTglPinjam');
    const tglKembaliEl = document.getElementById('pinjamTglKembali');
    if (!tglPinjamEl || !tglKembaliEl || !tglPinjamEl.value) return;

    const startDate = tglPinjamEl.value;
    const maxReturnDateDate = addDaysToDate(startDate, MAX_BORROW_DAYS);
    const maxReturnDate = formatDateInputValue(maxReturnDateDate);
    if (!maxReturnDate) return;

    tglKembaliEl.min = startDate;
    tglKembaliEl.max = maxReturnDate;

    if (!tglKembaliEl.value || tglKembaliEl.value < startDate || tglKembaliEl.value > maxReturnDate) {
        tglKembaliEl.value = maxReturnDate;
    }
}

function getHistoryStatus(peminjamanRow, tglRealisasi) {
    var tglPinjam = normalizeDate(parseDateValue(peminjamanRow.TGL_PINJAM));
    var rencanaKembali = normalizeDate(parseDateValue(peminjamanRow.TGL_KEMBALI_RENCANA));

    // ===============================
    // SUDAH DIKEMBALIKAN
    // ===============================
    if (tglRealisasi && tglRealisasi !== '-') {
        var realisasiDate = normalizeDate(parseDateValue(tglRealisasi));

        if (tglPinjam && realisasiDate) {
            var borrowDays = getBorrowDurationDays(
                peminjamanRow.TGL_PINJAM,
                tglRealisasi
            );

            // 🔴 MERAH → lebih dari 5 hari
            if (!isNaN(borrowDays) && borrowDays >= MAX_BORROW_DAYS) {
                return {
                    label: 'Dikembalikan lewat 5 hari',
                    className: 'history-status-red'
                };
            }
        }

        // 🟢 HIJAU → tepat waktu (≤ 5 hari)
        return {
            label: 'Dikembalikan tepat waktu',
            className: 'history-status-green'
        };
    }

    // ===============================
    // BELUM DIKEMBALIKAN
    // ===============================
    var today = normalizeDate(new Date());

    var batasMax = normalizeDate(
        addDaysToDate(peminjamanRow.TGL_PINJAM, MAX_BORROW_DAYS)
    );

    // 🔴 MERAH → lewat 5 hari dari pinjam
    if (batasMax && today > batasMax) {
        return {
            label: 'Belum dikembalikan > 5 hari',
            className: 'history-status-red'
        };
    }

    // 🟡 KUNING → lewat rencana tapi belum 5 hari
    if (rencanaKembali && today > rencanaKembali) {
        return {
            label: 'Lewat tanggal rencana',
            className: 'history-status-yellow'
        };
    }

    // 🟡 KUNING → masih masa peminjaman
    return {
        label: 'Masih masa peminjaman',
        className: 'history-status-yellow'
    };
}

// ==========================================
// NAVIGATION
// ==========================================
function showPage(pageName) {
    if (pageName !== 'login' && !AppState.authenticated) {
        showLoginPage();
        return;
    }

    // Hide all pages
    document.querySelectorAll('.page').forEach(function (p) {
        p.classList.remove('active');
    });
    // Show target page
    var targetPage = document.getElementById(pageName + 'Page');
    if (targetPage) {
        targetPage.classList.add('active');
    }
    // Update hash
    window.location.hash = pageName;
    // Scroll to top
    window.scrollTo(0, 0);
    // Load page data
    if (pageName === 'dashboard') {
        renderDashboard();
    } else if (pageName === 'peminjaman') {
        if (!AppState.keperluan || AppState.keperluan.length === 0) {
            return loadData(true).then(function () {
                populatePeminjamanForm();
            }).catch(function () {
                populatePeminjamanForm();
            });
        }
        populatePeminjamanForm();
    } else if (pageName === 'pengembalian') {
        populatePengembalianForm();
    }
}

// ==========================================
// FORMS INIT
// ==========================================
function initForms() {
    const formPinjam = document.getElementById('formPeminjaman');
    if (formPinjam) {
        formPinjam.addEventListener('submit', handlePeminjaman);
    }

    const formKembali = document.getElementById('formPengembalian');
    if (formKembali) {
        formKembali.addEventListener('submit', handlePengembalian);
    }

    const loginForm = document.getElementById('loginForm');
    if (loginForm) {
        loginForm.addEventListener('submit', handleLogin);
    }

    const tglPinjamEl = document.getElementById('pinjamTglPinjam');
    if (tglPinjamEl) {
        tglPinjamEl.addEventListener('change', applyPeminjamanDateLimit);
    }

    // Auto-fill tim on pegawai select
    const pinjamNama = document.getElementById('pinjamNama');
    if (pinjamNama) {
        pinjamNama.addEventListener('change', function () {
            const nip = String(this.value);
            const pegawai = AppState.pegawai.find(function (p) { return String(p.NIP) === nip; });
            const timInput = document.getElementById('pinjamTim');
            if (pegawai && timInput) {
                timInput.value = getPegawaiTim(pegawai);
            } else if (timInput) {
                timInput.value = '';
            }
        });
    }
}

function showLoginPage() {
    document.querySelectorAll('.page').forEach(function (p) {
        p.classList.remove('active');
    });
    var loginPage = document.getElementById('loginPage');
    if (loginPage) {
        loginPage.classList.add('active');
    }
    AppState.authenticated = false;
    updateHeaderAuthState();
    window.location.hash = 'login';
}

function logout() {
    AppState.authenticated = false;
    showToast('Logout berhasil.', 'success');
    updateHeaderAuthState();
    showLoginPage();
}

function updateHeaderAuthState() {
    var logoutBtn = document.getElementById('logoutButton');
    if (!logoutBtn) return;
    logoutBtn.style.display = AppState.authenticated ? '' : 'none';
}

function handleLogin(e) {
    e.preventDefault();
    var input = document.getElementById('loginCode');
    if (!input) return;

    var kode = String(input.value || '').trim();
    if (!kode) {
        showToast('Masukkan kode akses terlebih dahulu.', 'error');
        return;
    }

    if (!AppState.kodeAkses || AppState.kodeAkses.length === 0) {
        showLoading('Memuat data akses...');
        Promise.resolve(loadData(true)).finally(function () {
            hideLoading();
            if (isValidAccessCode(kode)) {
                AppState.authenticated = true;
                showToast('Login berhasil.', 'success');
                showPage('dashboard');
            } else {
                showToast('Kode akses tidak valid. Periksa kembali di spreadsheet.', 'error');
            }
        });
        return;
    }

    if (isValidAccessCode(kode)) {
        AppState.authenticated = true;
        updateHeaderAuthState();
        showToast('Login berhasil.', 'success');
        showPage('dashboard');
    } else {
        showToast('Kode akses tidak valid. Periksa kembali di spreadsheet.', 'error');
    }
}

function isValidAccessCode(kode) {
    if (!kode || !AppState.kodeAkses) return false;
    var enteredCode = String(kode).trim();
    return AppState.kodeAkses.some(function (row) {
        var value = row.KODE || row['KODE AKSES'] || row['Kode Akses'] || row['kode'] || row['kode akses'] || row.CODE || row['Code'];
        return String(value || '').trim() === enteredCode;
    });
}

function isPegawaiActive(pegawai) {
    if (!pegawai) return false;
    var status = String(
        pegawai.AKTIF ||
        pegawai.STATUS ||
        pegawai['AKTIF'] ||
        pegawai['STATUS'] ||
        pegawai['STATUS PEGAWAI'] ||
        pegawai['STATUS_AKTIF'] ||
        pegawai['IS_ACTIVE'] ||
        pegawai['IS AKTIF'] ||
        ''
    ).trim().toLowerCase();

    if (!status) return false;
    return ['aktif', 'active', 'ya', 'yes', 'true', '1'].includes(status);
}

function getKeperluanText(row) {
    if (!row || typeof row !== 'object') return '';
    for (var key in row) {
        if (!Object.prototype.hasOwnProperty.call(row, key)) continue;
        var value = String(row[key] || '').trim();
        if (value) return value;
    }
    return '';
}

function fetchKeperluanSheet() {
    return getSheetData('Keperluan').then(function (result) {
        if (result && result.success && Array.isArray(result.data)) {
            AppState.keperluan = result.data;
            saveToLocalStorage();
        }
    }).catch(function (err) {
        console.warn('Gagal mengambil sheet Keperluan:', err);
    });
}

// ==========================================
// DATA LOADING (from Spreadsheet)
// ==========================================
function loadData(silent) {
    const url = CONFIG.APPS_SCRIPT_URL;
    if (!url) {
        loadFromLocalStorage();
        updateSyncUI('offline');
        renderDashboard();
        return;
    }

    if (AppState.syncing) return;
    AppState.syncing = true;

    if (!silent) showLoading('Mengambil data dari spreadsheet...');
    updateSyncUI('syncing');

    return fetch(url + '?action=getAllData&t=' + Date.now())
        .then(function (res) {
            if (!res.ok) throw new Error('Network error');
            return res.json();
        })
        .then(function (result) {
            if (result.success && result.data) {
                AppState.laptops = result.data.data_laptop || [];
                AppState.pegawai = result.data.data_pegawai || [];
                // Patch: inject LAST_ACTION_TIME jika belum ada (agar sorting selalu benar)
                AppState.peminjaman = (result.data.data_peminjaman || []).map(function(row) {
                                    var timestampValue = row.LAST_ACTION_TIME || row.Timestamp;
                    var lastActionTime = 0;

                    if (timestampValue) {
                        if (typeof timestampValue === 'string' && timestampValue.trim()) {
                            lastActionTime = parseTimestamp(timestampValue);
                        } else if (!isNaN(Number(timestampValue))) {
                            lastActionTime = Number(timestampValue);
                        }
                    }

                    if (!lastActionTime && row.ID && /^PEM-\d+$/.test(row.ID)) {
                        lastActionTime = parseInt(row.ID.replace('PEM-', ''));
                    }
                    if (!lastActionTime && row.TGL_PINJAM) {
                        lastActionTime = new Date(row.TGL_PINJAM).getTime();
                    }
                    if (!lastActionTime) {
                        lastActionTime = 0;
                    }

                    row.LAST_ACTION_TIME = lastActionTime;
                    return row;
                });
                AppState.pengembalian = (result.data.data_pengembalian || []).map(function(row) {
                    var ts = 0;
                    if (row.Timestamp) {
                        ts = parseTimestamp(row.Timestamp);
                    }
                    if (!ts && row.ID && /^KEM-\d+$/.test(row.ID)) {
                        ts = parseInt(row.ID.replace('KEM-', ''));
                    }
                    if (!ts && row.TGL_REALISASI_PENGEMBALIAN) {
                        ts = new Date(row.TGL_REALISASI_PENGEMBALIAN).getTime();
                    }
                    row.LAST_ACTION_TIME = ts || 0;
                    return row;
                });
                AppState.kodeAkses = result.data.kode_akses || [];
                AppState.keperluan = result.data.keperluan || [];
                if (!AppState.keperluan.length) {
                    return fetchKeperluanSheet().then(function () {
                        saveToLocalStorage();
                    });
                }

                // Save to localStorage as backup
                saveToLocalStorage();

                AppState.loaded = true;
                updateSyncUI('connected');
            } else {
                throw new Error('Invalid response');
            }
        })
        .catch(function (err) {
            loadFromLocalStorage();
            updateSyncUI('offline');
            if (!silent) showToast('Gagal memuat dari server, menggunakan data lokal', 'warning');
        })
        .finally(function () {
            AppState.syncing = false;
            hideLoading();
            renderDashboard();
            startAutoRefresh();
        });
}

function refreshData(silent) {
    if (!silent) showLoading('Menyegarkan data...');
    return loadData(silent);
}

function startAutoRefresh() {
    if (AppState.refreshTimer) clearInterval(AppState.refreshTimer);
    if (!CONFIG.APPS_SCRIPT_URL) return;

    AppState.refreshTimer = setInterval(function () {
        if (!AppState.syncing) {
            loadData(true);
        }
    }, CONFIG.AUTO_REFRESH_MS);
}

function saveToLocalStorage() {
    try {
        localStorage.setItem('iv_laptops', JSON.stringify(AppState.laptops));
        localStorage.setItem('iv_pegawai', JSON.stringify(AppState.pegawai));
        localStorage.setItem('iv_peminjaman', JSON.stringify(AppState.peminjaman));
        localStorage.setItem('iv_pengembalian', JSON.stringify(AppState.pengembalian));
        localStorage.setItem('iv_kode_akses', JSON.stringify(AppState.kodeAkses));
        localStorage.setItem('iv_keperluan', JSON.stringify(AppState.keperluan));
    } catch (e) { }
}

function loadFromLocalStorage() {
    try {
        AppState.laptops = JSON.parse(localStorage.getItem('iv_laptops') || '[]');
        AppState.pegawai = JSON.parse(localStorage.getItem('iv_pegawai') || '[]');
        AppState.peminjaman = JSON.parse(localStorage.getItem('iv_peminjaman') || '[]');
        AppState.pengembalian = JSON.parse(localStorage.getItem('iv_pengembalian') || '[]');
        AppState.kodeAkses = JSON.parse(localStorage.getItem('iv_kode_akses') || '[]');
        AppState.keperluan = JSON.parse(localStorage.getItem('iv_keperluan') || '[]');
        AppState.loaded = true;
    } catch (e) {
    }
}

// ==========================================
// SYNC UI
// ==========================================
function updateSyncUI(status) {
    const dot = document.getElementById('syncDot');
    const label = document.getElementById('syncLabel');
    if (!dot || !label) return;

    dot.className = 'sync-dot';

    if (status === 'connected') {
        dot.classList.add('connected');
        label.textContent = 'Terhubung';
    } else if (status === 'syncing') {
        dot.classList.add('syncing');
        label.textContent = 'Sinkronisasi...';
    } else {
        label.textContent = 'Offline';
    }
}

// ==========================================
// DASHBOARD RENDER
// ==========================================
function renderDashboard() {
    renderStats();
    renderRiwayat();
}

function renderStats() {
    const laptops = AppState.laptops;

    const dipinjam = laptops.filter(function (l) {
        return normalizeLaptopStatus(l.STATUS) === 'dipinjam';
    }).length;
    const rusak = laptops.filter(function (l) {
        return normalizeLaptopStatus(l.STATUS) === 'rusak';
    }).length;
    const tersedia = laptops.filter(function (l) {
        return normalizeLaptopStatus(l.STATUS) === 'tersedia';
    }).length;

    animateNumber('statTersedia', tersedia);
    animateNumber('statDipinjam', dipinjam);
    animateNumber('statRusak', rusak);
}

function animateNumber(id, target) {
    const el = document.getElementById(id);
    if (!el) return;
    const start = parseInt(el.textContent) || 0;
    if (start === target) { el.textContent = target; return; }
    const duration = 400;
    const stepTime = 20;
    const steps = duration / stepTime;
    const increment = (target - start) / steps;
    let current = start;
    let step = 0;
    const timer = setInterval(function () {
        step++;
        current += increment;
        el.textContent = Math.round(current);
        if (step >= steps) {
            el.textContent = target;
            clearInterval(timer);
        }
    }, stepTime);
}

// ==========================================
// RIWAYAT TABLE (Dashboard)
// ==========================================
const HISTORY_PAGE_SIZE = 5;
let historyCurrentPage = 1;

function renderRiwayat(filteredData) {
    const body = document.getElementById('riwayatBody');
    const empty = document.getElementById('riwayatEmpty');
    const table = document.getElementById('riwayatTable');

    // Remove old load more button if exists
    const oldBtn = document.getElementById('btnLoadMoreHistory');
    if (oldBtn && oldBtn.parentNode && oldBtn.parentNode.classList.contains('table-footer-actions')) {
        oldBtn.parentNode.remove();
    }

    // Remove old pagination container if exists to re-render fresh
    const oldPag = document.getElementById('historyPagination');
    if (oldPag) oldPag.remove();

    if (!body) return;

    // Combine peminjaman + pengembalian info
    let peminjaman = filteredData || AppState.peminjaman;
    // Filter agar hanya data unik berdasarkan ID
    const seen = new Set();
    peminjaman = peminjaman.filter(function(item) {
        if (!item.ID) return true;
        if (seen.has(item.ID)) return false;
        seen.add(item.ID);
        return true;
    });
    // Urutkan berdasarkan LAST_ACTION_TIME / Timestamp terbaru descending
    peminjaman = [...peminjaman].sort(function(a, b) {
        var tglA = Number(a.LAST_ACTION_TIME) || parseTimestamp(a.Timestamp) || new Date(a.TGL_PINJAM).getTime() || 0;
        var tglB = Number(b.LAST_ACTION_TIME) || parseTimestamp(b.Timestamp) || new Date(b.TGL_PINJAM).getTime() || 0;
        return tglB - tglA;
    });

    // Reset page if new filter applied (ad-hoc checking if filteredData is passed and page is out of bounds)
    if (filteredData) {
        const maxPage = Math.ceil(filteredData.length / HISTORY_PAGE_SIZE) || 1;
        if (historyCurrentPage > maxPage) historyCurrentPage = 1;
    }

    body.innerHTML = '';

    if (peminjaman.length === 0) {
        if (table) table.style.display = 'none';
        if (empty) empty.style.display = 'block';
        return;
    }

    if (table) table.style.display = '';
    if (empty) empty.style.display = 'none';

    // Pagination logic
    const totalPages = Math.ceil(peminjaman.length / HISTORY_PAGE_SIZE);

    // Ensure current page is valid
    if (historyCurrentPage > totalPages) historyCurrentPage = totalPages;
    if (historyCurrentPage < 1) historyCurrentPage = 1;

    const startIndex = (historyCurrentPage - 1) * HISTORY_PAGE_SIZE;
    const endIndex = startIndex + HISTORY_PAGE_SIZE;
    const visibleData = peminjaman.slice(startIndex, endIndex);

    visibleData.forEach(function (p, idx) {
        // Find matching pengembalian
        var kembali = AppState.pengembalian.find(function (k) {
            return k.PEMINJAMAN_ID === p.ID;
        });

        var tglRealisasi = kembali ? (kembali.TGL_REALISASI_PENGEMBALIAN || '-') : '-';
        var statusRiwayat = getHistoryStatus(p, tglRealisasi);

        // Laptop info
        var laptop = AppState.laptops.find(function (l) { return l.ID === p.LAPTOP_ID; });
        var laptopName = laptop ? ('[' + laptop.ID + '] ' + laptop.TYPE) : (p.LAPTOP_ID || '-');

        // Status indicator logic:
        // - Jika sudah realisasi: tampil hijau/merah sesuai status
        // - Jika belum realisasi + dalam rencana: kosong
        // - Jika belum realisasi + lewat rencana: kuning (reminder)
        // - Jika belum realisasi + lewat 5 hari: merah
        var statusIndicator = '';
        if (tglRealisasi && tglRealisasi !== '-') {
            // Sudah dikembalikan - tampil indicator
            statusIndicator = '<span class="history-status-dot ' + statusRiwayat.className + '" title="' + escapeHtml(statusRiwayat.label) + '"></span>';
        } else {
            // Belum dikembalikan - cek kondisi untuk reminder/alert
            var today = normalizeDate(new Date());
            var rencanaKembali = normalizeDate(parseDateValue(p.TGL_KEMBALI_RENCANA));
            var batasMax = normalizeDate(addDaysToDate(p.TGL_PINJAM, MAX_BORROW_DAYS));
            
            // Kuning: reminder jika sudah lewat rencana tapi belum 5 hari
            if (rencanaKembali && today > rencanaKembali && (!batasMax || today <= batasMax)) {
                statusIndicator = '<span class="history-status-dot history-status-yellow" title="Reminder: Sudah lewat tanggal rencana kembali"></span>';
            }
            // Merah: kritis jika sudah lewat 5 hari
            else if (batasMax && today > batasMax) {
                statusIndicator = '<span class="history-status-dot history-status-red" title="Alert: Sudah lewat 5 hari dari peminjaman"></span>';
            }
        }

        var tr = document.createElement('tr');
        tr.style.cursor = 'pointer';
        tr.setAttribute('data-peminjaman-id', p.ID);
        tr.setAttribute('title', 'Klik untuk melihat detail');
        tr.onclick = function() {
            showDetailPeminjamanModal(p.ID);
        };
        tr.innerHTML =
            '<td>' + (startIndex + idx + 1) + '</td>' +
            '<td><strong>' + escapeHtml(laptopName) + '</strong></td>' +
            '<td>' + escapeHtml(p.NAMA_PEMINJAM || '-') + '</td>' +
            '<td>' + formatDate(p.TGL_PINJAM) + '</td>' +
            '<td>' + formatDate(p.TGL_KEMBALI_RENCANA) + '</td>' +
            '<td>' + formatDate(tglRealisasi) + '</td>' +
            '<td>' + statusIndicator + '</td>';
        body.appendChild(tr);
    });

    // Render Pagination Controls
    if (totalPages > 1) {
        const pagContainer = document.createElement('div');
        pagContainer.id = 'historyPagination';
        pagContainer.className = 'pagination-controls';
        pagContainer.style.display = 'flex';
        pagContainer.style.justifyContent = 'center';
        pagContainer.style.alignItems = 'center';
        pagContainer.style.gap = '1rem';
        pagContainer.style.marginTop = '1rem';

        // Prev Button
        const btnPrev = document.createElement('button');
        btnPrev.className = 'btn btn-outline btn-sm';
        btnPrev.innerHTML = '<i class="fas fa-chevron-left"></i>';
        btnPrev.disabled = historyCurrentPage === 1;
        btnPrev.onclick = prevHistoryPage;

        // Page Info
        const pageInfo = document.createElement('span');
        pageInfo.className = 'pagination-info';
        pageInfo.style.fontSize = '0.9rem';
        pageInfo.style.fontWeight = '500';
        pageInfo.textContent = 'Halaman ' + historyCurrentPage + ' dari ' + totalPages;

        // Next Button
        const btnNext = document.createElement('button');
        btnNext.className = 'btn btn-outline btn-sm';
        btnNext.innerHTML = '<i class="fas fa-chevron-right"></i>';
        btnNext.disabled = historyCurrentPage === totalPages;
        btnNext.onclick = nextHistoryPage;

        pagContainer.appendChild(btnPrev);
        pagContainer.appendChild(pageInfo);
        pagContainer.appendChild(btnNext);

        table.parentNode.insertAdjacentElement('afterend', pagContainer);
    }
}

function prevHistoryPage() {
    if (historyCurrentPage > 1) {
        historyCurrentPage--;
        renderRiwayat();
    }
}

function nextHistoryPage() {
    historyCurrentPage++;
    renderRiwayat();
}

// ==========================================
// FILTER SYSTEM (Dropdown-based)
// ==========================================
function onFilterTypeChange() {
    var type = document.getElementById('filterType').value;
    var area = document.getElementById('filterInputArea');
    if (!area) return;

    area.innerHTML = '';

    if (type === 'semua') {
        renderRiwayat();
        return;
    }

    if (type === 'laptop') {
        var input = document.createElement('input');
        input.type = 'text';
        input.id = 'filterValue';
        input.placeholder = 'Ketik nama/merk laptop...';
        input.oninput = applyFilter;
        area.appendChild(input);
    }

    if (type === 'nama') {
        var input = document.createElement('input');
        input.type = 'text';
        input.id = 'filterValue';
        input.placeholder = 'Ketik nama peminjam...';
        input.oninput = applyFilter;
        area.appendChild(input);
    }

    if (type === 'tanggal') {
        // Preset Dropdown
        var presetSelect = document.createElement('select');
        presetSelect.id = 'filterPreset';
        presetSelect.innerHTML =
            '<option value="custom">Pilih Rentang...</option>' +
            '<option value="1w">1 Minggu Terakhir</option>' +
            '<option value="2w">2 Minggu Terakhir</option>' +
            '<option value="1m">1 Bulan Terakhir</option>' +
            '<option value="4m">4 Bulan Terakhir</option>';
        presetSelect.onchange = function () {
            var val = this.value;
            if (val === 'custom') return;

            var end = new Date();
            var start = new Date();
            if (val === '1w') start.setDate(end.getDate() - 7);
            if (val === '2w') start.setDate(end.getDate() - 14);
            if (val === '1m') start.setMonth(end.getMonth() - 1);
            if (val === '4m') start.setMonth(end.getMonth() - 4);

            // Format YYYY-MM-DD (safe for local time if we use split on ISO string after adjusting timezone offset, 
            // but for simplicity here standard ISO split is ok or use simple formatter)
            // Using simple offset fix for timezone
            var toDateInputValue = function (date) {
                var local = new Date(date);
                local.setMinutes(date.getMinutes() - date.getTimezoneOffset());
                return local.toJSON().slice(0, 10);
            };

            document.getElementById('filterDari').value = toDateInputValue(start);
            document.getElementById('filterSampai').value = toDateInputValue(end);
            applyFilter();
        };
        area.appendChild(presetSelect);

        var labelDari = document.createElement('label');
        labelDari.textContent = 'Dari:';
        labelDari.style.marginLeft = '10px';
        area.appendChild(labelDari);

        var inputDari = document.createElement('input');
        inputDari.type = 'date';
        inputDari.id = 'filterDari';
        inputDari.onchange = function () {
            document.getElementById('filterPreset').value = 'custom';
            applyFilter();
        };
        area.appendChild(inputDari);

        var labelSampai = document.createElement('label');
        labelSampai.textContent = 'Sampai:';
        area.appendChild(labelSampai);

        var inputSampai = document.createElement('input');
        inputSampai.type = 'date';
        inputSampai.id = 'filterSampai';
        inputSampai.onchange = function () {
            document.getElementById('filterPreset').value = 'custom';
            applyFilter();
        };
        area.appendChild(inputSampai);

        var btnReset = document.createElement('button');
        btnReset.className = 'btn btn-outline btn-sm';
        btnReset.innerHTML = '<i class="fas fa-times"></i> Reset';
        btnReset.onclick = function () {
            document.getElementById('filterType').value = 'semua';
            onFilterTypeChange();
        };
        area.appendChild(btnReset);
    }

    // Show all data initially when switching filter type
    renderRiwayat();
}

function applyFilter() {
    var type = document.getElementById('filterType').value;
    var data = AppState.peminjaman;

    if (type === 'laptop') {
        var keyword = ((document.getElementById('filterValue') || {}).value || '').toLowerCase();
        if (keyword) {
            data = data.filter(function (p) {
                var laptop = AppState.laptops.find(function (l) { return l.ID === p.LAPTOP_ID; });
                var label = laptop ? (laptop.ID + ' ' + laptop.MERK + ' ' + laptop.TYPE).toLowerCase() : (p.LAPTOP_ID || '').toLowerCase();
                return label.indexOf(keyword) !== -1;
            });
        }
    }

    if (type === 'nama') {
        var keyword = ((document.getElementById('filterValue') || {}).value || '').toLowerCase();
        if (keyword) {
            data = data.filter(function (p) {
                return (p.NAMA_PEMINJAM || '').toLowerCase().indexOf(keyword) !== -1;
            });
        }
    }

    if (type === 'tanggal') {
        var dari = (document.getElementById('filterDari') || {}).value || '';
        var sampai = (document.getElementById('filterSampai') || {}).value || '';
        if (dari) {
            data = data.filter(function (p) { return (p.TGL_PINJAM || '') >= dari; });
        }
        if (sampai) {
            data = data.filter(function (p) { return (p.TGL_PINJAM || '') <= sampai; });
        }
    }

    renderRiwayat(data);
}

// ==========================================
// PEMINJAMAN FORM
// ==========================================
function populatePeminjamanForm() {
    // Populate laptop dropdown (only Tersedia)
    const laptopSelect = document.getElementById('pinjamLaptop');
    if (laptopSelect) {
        const currentVal = laptopSelect.value;
        laptopSelect.innerHTML = '<option value="">-- Pilih Laptop --</option>';
        AppState.laptops.forEach(function (l) {
            var status = (l.STATUS || '').toLowerCase();
            if (status === 'tersedia' || status === 'rusak ringan') {
                const opt = document.createElement('option');
                opt.value = l.ID;
                opt.textContent = l.ID + ' - ' + l.MERK + ' ' + l.TYPE + (l.NUP ? ' (NUP: ' + l.NUP + ')' : '');
                laptopSelect.appendChild(opt);
            }
        });
        if (currentVal) laptopSelect.value = currentVal;
    }

    // Populate pegawai dropdown
    const namaSelect = document.getElementById('pinjamNama');
    if (namaSelect) {
        const currentVal = namaSelect.value;
        namaSelect.innerHTML = '<option value="">-- Pilih Pegawai --</option>';
        AppState.pegawai.forEach(function (p) {
            if (!isPegawaiActive(p)) return;

            const opt = document.createElement('option');
            opt.value = String(p.NIP);
            opt.textContent = p.NAMA + ' (NIP: ' + p.NIP + ')';
            namaSelect.appendChild(opt);
        });
        if (currentVal) namaSelect.value = currentVal;
    }

    // Populate keperluan dropdown
    const keperluanSelect = document.getElementById('pinjamKeperluan');
    if (keperluanSelect) {
        const currentVal = keperluanSelect.value;
        keperluanSelect.innerHTML = '<option value="">-- Pilih Keperluan --</option>';
        AppState.keperluan.forEach(function (row) {
            var label = getKeperluanText(row);
            if (!label) return;
            var opt = document.createElement('option');
            opt.value = label;
            opt.textContent = label;
            keperluanSelect.appendChild(opt);
        });
        if (currentVal) keperluanSelect.value = currentVal;
    }

    // Set tanggal default
    const today = new Date().toISOString().split('T')[0];
    const tglPinjam = document.getElementById('pinjamTglPinjam');
    if (tglPinjam && !tglPinjam.value) tglPinjam.value = today;

    applyPeminjamanDateLimit();
}


let pendingPeminjamanData = null;
let peminjamanInProgress = false;
    // Cegah double submit
    const btn = document.querySelector('#formPeminjaman button[type="submit"]');
    if (btn) btn.disabled = true;
    setTimeout(function() { if (btn) btn.disabled = false; }, 5000); // fallback re-enable

function handlePeminjaman(e) {
    e.preventDefault();

    const laptopId = document.getElementById('pinjamLaptop').value;
    const nip = document.getElementById('pinjamNama').value;
    const tim = document.getElementById('pinjamTim').value;
    const keperluan = document.getElementById('pinjamKeperluan').value;
    const deskripsi = document.getElementById('pinjamDeskripsi').value;
    const tglPinjam = document.getElementById('pinjamTglPinjam').value;
    const tglKembali = document.getElementById('pinjamTglKembali').value;

    if (!laptopId || !nip || !keperluan || !tglPinjam || !tglKembali) {
        showToast('Mohon lengkapi semua field yang wajib!', 'error');
        return;
    }

    // Validate dates
    const parsedTglPinjam = parseDateValue(tglPinjam);
    const parsedTglKembali = parseDateValue(tglKembali);
    if (!parsedTglPinjam || !parsedTglKembali) {
        showToast('Format tanggal tidak valid. Gunakan tanggal yang disediakan sistem.', 'error');
        return;
    }

    if (parsedTglKembali < parsedTglPinjam) {
        showToast('Tanggal kembali tidak boleh sebelum tanggal pinjam!', 'error');
        return;
    }

    const durasiPinjamHari = getBorrowDurationDays(tglPinjam, tglKembali);
    if (isNaN(durasiPinjamHari) || durasiPinjamHari > MAX_BORROW_DAYS) {
        showToast('Maksimal peminjaman adalah 5 hari dari tanggal pinjam.', 'error');
        return;
    }

    // Find Names for Confirmation
    const pegawai = AppState.pegawai.find(function (p) { return String(p.NIP) === String(nip); });
    const namaPeminjam = pegawai ? pegawai.NAMA : nip;

    const laptop = AppState.laptops.find(l => l.ID == laptopId);
    const laptopName = laptop ? (laptop.MERK + ' ' + laptop.TYPE) : laptopId;

    // Store data for confirmation
    pendingPeminjamanData = {
        laptopId,
        nip,
        namaPeminjam,
        tim,
        keperluan,
        deskripsi,
        tglPinjam,
        tglKembali
    };

    // Show Confirmation Modal
    const message = `
        <div style="text-align: left; background: var(--bg-subtle); padding: 1rem; border-radius: var(--radius-md); font-size: 0.9rem;">
            <div style="margin-bottom: 0.5rem;"><strong>Peminjam:</strong> ${escapeHtml(namaPeminjam)}</div>
            <div style="margin-bottom: 0.5rem;"><strong>Laptop:</strong> ${escapeHtml(laptopName)}</div>
            <div style="margin-bottom: 0.5rem;"><strong>Keperluan:</strong> ${escapeHtml(keperluan)}</div>
            <div style="margin-bottom: 0.5rem;"><strong>Tanggal:</strong> ${formatDate(tglPinjam)} s.d ${formatDate(tglKembali)}</div>
        </div>
        <p style="margin-top: 1rem;">Apakah data di atas sudah benar?</p>
    `;

    showConfirmModal('Konfirmasi Peminjaman', message, processPeminjaman);
}

function processPeminjaman() {
    if (!pendingPeminjamanData || peminjamanInProgress) return;
    peminjamanInProgress = true;

    closeConfirmModal();
    showLoading('Menyimpan peminjaman...');

    // Generate ID
    const now = Date.now();
    const id = 'PEM-' + now;

    const row = {
        ID: id,
        LAPTOP_ID: pendingPeminjamanData.laptopId,
        NAMA_PEMINJAM: pendingPeminjamanData.namaPeminjam,
        NIP: pendingPeminjamanData.nip,
        TIM: pendingPeminjamanData.tim,
        DIVISI: pendingPeminjamanData.tim,
        KEPERLUAN: pendingPeminjamanData.keperluan,
        DESKRIPSI_KEPERLUAN: pendingPeminjamanData.deskripsi,
        TGL_PINJAM: pendingPeminjamanData.tglPinjam,
        TGL_KEMBALI_RENCANA: pendingPeminjamanData.tglKembali,
        STATUS: 'Aktif',
        Timestamp: formatTimestamp(new Date())
    };

    // 1. Append to Data Peminjaman
    appendToSheet(CONFIG.SHEETS.DATA_PEMINJAMAN, row)
        .then(function () {
            // 2. Update laptop status to Dipinjam
            return updateSheetRow(CONFIG.SHEETS.DATA_LAPTOP, 'ID', pendingPeminjamanData.laptopId, { STATUS: 'Dipinjam' });
        })
        .then(function () {
            // Update local state
            AppState.peminjaman.push(row);

            const laptop = AppState.laptops.find(function (l) { return l.ID === pendingPeminjamanData.laptopId; });
            if (laptop) laptop.STATUS = 'Dipinjam';

            saveToLocalStorage();
            hideLoading();
            showToast('Peminjaman berhasil disimpan!', 'success');

            // Cleanup
            document.getElementById('formPeminjaman').reset();
            const today = new Date().toISOString().split('T')[0];
            const tglEl = document.getElementById('pinjamTglPinjam');
            if (tglEl) tglEl.value = today;

            populatePeminjamanForm();
            showPage('dashboard');
            renderDashboard();
            peminjamanInProgress = false;
            const btn = document.querySelector('#formPeminjaman button[type="submit"]');
            if (btn) btn.disabled = false;
        })
        .catch(function (error) {
            hideLoading();
            showToast('Gagal menyimpan peminjaman: ' + error.message, 'error');
            peminjamanInProgress = false;
            const btn = document.querySelector('#formPeminjaman button[type="submit"]');
            if (btn) btn.disabled = false;
        });
}

// Confirmation Modal Helpers
function showConfirmModal(title, htmlContent, onConfirm) {
    const modal = document.getElementById('confirmModal');
    if (!modal) return;

    const titleEl = modal.querySelector('.modal-title');
    const textEl = document.getElementById('confirmModalText');
    const btn = document.getElementById('confirmModalBtn');

    titleEl.textContent = title;
    textEl.innerHTML = htmlContent;

    // Clone button to remove old listeners
    const newBtn = btn.cloneNode(true);
    btn.parentNode.replaceChild(newBtn, btn);

    newBtn.onclick = onConfirm;

    modal.classList.add('show');
}

function closeConfirmModal() {
    const modal = document.getElementById('confirmModal');
    if (modal) modal.classList.remove('show');
}


// ==========================================
// PENGEMBALIAN FORM
// ==========================================
function populatePengembalianForm() {
    // Populate pegawai dropdown
    const select = document.getElementById('kembaliPegawai');
    if (select) {
        select.innerHTML = '<option value="">-- Pilih Pegawai --</option>';

        // Only show pegawai who have active loans
        const activeBorrowers = new Set();
        AppState.peminjaman.forEach(function (p) {
            if ((p.STATUS || '').toLowerCase() === 'aktif') {
                activeBorrowers.add(p.NAMA_PEMINJAM);
            }
        });

        AppState.pegawai.forEach(function (p) {
            if (activeBorrowers.has(p.NAMA)) {
                const opt = document.createElement('option');
                opt.value = p.NAMA;
                var tim = getPegawaiTim(p);
                opt.textContent = p.NAMA + (tim ? (' - ' + tim) : '');
                select.appendChild(opt);
            }
        });

        // Also add names from peminjaman that might not be in pegawai
        activeBorrowers.forEach(function (name) {
            const exists = AppState.pegawai.some(function (p) { return p.NAMA === name; });
            if (!exists) {
                const opt = document.createElement('option');
                opt.value = name;
                opt.textContent = name;
                select.appendChild(opt);
            }
        });
    }

    // Hide step 2 and 3
    hideElement('kembaliListCard');
    hideElement('kembaliFormCard');

    // Reset date
    const today = new Date().toISOString().split('T')[0];
    const tgl = document.getElementById('kembaliTglRealisasi');
    if (tgl) tgl.value = today;
}

function onSelectPegawaiKembali() {
    const nama = document.getElementById('kembaliPegawai').value;
    const listCard = document.getElementById('kembaliListCard');
    const tbody = document.getElementById('kembaliListBody');
    const empty = document.getElementById('kembaliEmpty');

    hideElement('kembaliFormCard');

    if (!nama) {
        hideElement('kembaliListCard');
        return;
    }

    showElement('kembaliListCard');

    // Filter active peminjaman for this person
    const activeLoans = AppState.peminjaman.filter(function (p) {
        return (p.STATUS || '').toLowerCase() === 'aktif' && p.NAMA_PEMINJAM === nama;
    });

    tbody.innerHTML = '';

    if (activeLoans.length === 0) {
        const table = listCard.querySelector('table');
        if (table) table.style.display = 'none';
        if (empty) empty.style.display = 'block';
        return;
    }

    const table = listCard.querySelector('table');
    if (table) table.style.display = '';
    if (empty) empty.style.display = 'none';

    activeLoans.forEach(function (loan) {
        const laptop = AppState.laptops.find(function (l) { return l.ID === loan.LAPTOP_ID; });
        const merkType = laptop ? (laptop.MERK + ' ' + laptop.TYPE) : '-';

        const tr = document.createElement('tr');
        tr.innerHTML =
            '<td><strong>' + escapeHtml(loan.LAPTOP_ID || '-') + '</strong></td>' +
            '<td>' + escapeHtml(merkType) + '</td>' +
            '<td>' + formatDate(loan.TGL_PINJAM) + '</td>' +
            '<td>' + formatDate(loan.TGL_KEMBALI_RENCANA) + '</td>' +
            '<td>' + escapeHtml(loan.KEPERLUAN || '-') + '</td>' +
            '<td><button class="btn btn-sm btn-success" onclick="selectLoanForReturn(\'' + loan.ID + '\')"><i class="fas fa-check"></i> Pilih</button></td>';
        tbody.appendChild(tr);
    });
}

function selectLoanForReturn(peminjamanId) {
    const loan = AppState.peminjaman.find(function (p) { return p.ID === peminjamanId; });
    if (!loan) return;

    const laptop = AppState.laptops.find(function (l) { return l.ID === loan.LAPTOP_ID; });

    // Fill detail grid
    const grid = document.getElementById('returnDetailGrid');
    if (grid) {
        grid.innerHTML =
            detailItem('Kode Laptop', loan.LAPTOP_ID || '-') +
            detailItem('Merk/Type', laptop ? (laptop.MERK + ' ' + laptop.TYPE) : '-') +
            detailItem('Nama Peminjam', loan.NAMA_PEMINJAM || '-') +
            detailItem('Tim', loan.TIM || loan.DIVISI || '-') +
            detailItem('Keperluan', loan.KEPERLUAN || '-') +
            detailItem('Tgl Pinjam', formatDate(loan.TGL_PINJAM)) +
            detailItem('Perkiraan Kembali', formatDate(loan.TGL_KEMBALI_RENCANA));
    }

    // Set hidden ID
    const hiddenId = document.getElementById('kembaliPeminjamanId');
    if (hiddenId) hiddenId.value = peminjamanId;

    // Set today as realisasi date
    const today = new Date().toISOString().split('T')[0];
    const tgl = document.getElementById('kembaliTglRealisasi');
    if (tgl) tgl.value = today;

    showElement('kembaliFormCard');

    // Scroll to form
    document.getElementById('kembaliFormCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function detailItem(label, value) {
    return '<div class="return-detail-item">' +
        '<span class="rd-label">' + label + '</span>' +
        '<span class="rd-value">' + escapeHtml(value) + '</span>' +
        '</div>';
}

function cancelReturn() {
    hideElement('kembaliFormCard');
}

// ==========================================
// PENGEMBALIAN FORM HANDLERS
// ==========================================
// Photo handling functions removed as requested


let pendingPengembalianData = null;

function handlePengembalian(e) {
    e.preventDefault();

    // Validate inputs
    const tglRealisasi = document.getElementById('kembaliTglRealisasi').value;
    const kondisi = document.getElementById('kembaliKondisi').value;
    const catatan = document.getElementById('kembaliCatatan').value;
    const peminjamanId = document.getElementById('kembaliPeminjamanId').value;

    if (!peminjamanId || !tglRealisasi || !kondisi) {
        showToast('Mohon lengkapi semua field!', 'error');
        return;
    }

    const peminjaman = AppState.peminjaman.find(function (p) { return String(p.ID) === String(peminjamanId); });

    if (!peminjaman) {
        showToast('Data peminjaman tidak ditemukan!', 'error');
        return;
    }

    const laptop = AppState.laptops.find(l => l.ID === peminjaman.LAPTOP_ID);
    const laptopName = laptop ? (laptop.MERK + ' ' + laptop.TYPE) : peminjaman.LAPTOP_ID;

    // Store for confirmation
    pendingPengembalianData = {
        peminjamanId,
        tglRealisasi,
        kondisi,
        catatan,
        peminjaman,
        laptopName
    };

    // Show Modal
    const message = `
        <div style="text-align: left; background: var(--bg-subtle); padding: 1rem; border-radius: var(--radius-md); font-size: 0.9rem;">
            <div style="margin-bottom: 0.5rem;"><strong>Peminjam:</strong> ${escapeHtml(peminjaman.NAMA_PEMINJAM)}</div>
            <div style="margin-bottom: 0.5rem;"><strong>Laptop:</strong> ${escapeHtml(laptopName)}</div>
            <div style="margin-bottom: 0.5rem;"><strong>Tgl Realisasi:</strong> ${formatDate(tglRealisasi)}</div>
            <div style="margin-bottom: 0.5rem;"><strong>Kondisi:</strong> ${escapeHtml(kondisi)}</div>
            <div style="margin-bottom: 0.5rem;"><strong>Catatan:</strong> ${escapeHtml(catatan || '-')}</div>
        </div>
        <p style="margin-top: 1rem;">Pastikan laptop sudah dicek. Proses pengembalian sekarang?</p>
    `;

    showConfirmModal('Konfirmasi Pengembalian', message, processPengembalian);
}

function processPengembalian() {
    const { peminjamanId, tglRealisasi, kondisi, catatan, peminjaman } = pendingPengembalianData;
    closeConfirmModal();

    showLoading('Menyinkronkan data ke spreadsheet...');

    // Create Data Pengembalian Row
    const now = Date.now();
    const returnRow = {
        ID: 'KEM-' + now,
        PEMINJAMAN_ID: peminjamanId,
        LAPTOP_ID: peminjaman.LAPTOP_ID,
        NAMA_PEMINJAM: peminjaman.NAMA_PEMINJAM,
        TGL_PINJAM: peminjaman.TGL_PINJAM,
        TGL_KEMBALI_RENCANA: peminjaman.TGL_KEMBALI_RENCANA,
        TGL_REALISASI_PENGEMBALIAN: tglRealisasi,

        // DUAL KEYS for compatibility
        KONDISI: kondisi,
        KONDISI_PENGEMBALIAN: kondisi,
        CATATAN: catatan,
        CATATAN_PENGEMBALIAN: catatan,

        STATUS: 'Selesai',
        Timestamp: formatTimestamp(new Date())
    };

    // Determine new laptop status
    let newLaptopStatus = 'Tersedia';
    if (kondisi === 'Rusak Ringan') newLaptopStatus = 'Rusak Ringan';
    if (kondisi === 'Rusak Berat') newLaptopStatus = 'Rusak Berat';

    // 1. Append to Data Pengembalian
    appendToSheet('Data Pengembalian', returnRow)
        .then(function () {
            // 2. Update Data Peminjaman Status
            return updateSheetRow('Data Peminjaman', 'ID', peminjamanId, {
                STATUS: 'Selesai',
                TGL_REALISASI_PENGEMBALIAN: tglRealisasi,
                KONDISI_PENGEMBALIAN: kondisi
            }).catch(e => {
                console.warn('Loan record update failed (might be deleted):', e);
                return { success: true };
            });
        })
        .then(function () {
            // 3. Update Data Laptop Status
            return updateSheetRow('Data Laptop', 'ID', peminjaman.LAPTOP_ID, { STATUS: newLaptopStatus });
        })
        .then(function () {
            // SUCCESS
            AppState.pengembalian.push(returnRow);
            peminjaman.STATUS = 'Selesai';
            peminjaman.TGL_REALISASI_PENGEMBALIAN = tglRealisasi;

            const laptop = AppState.laptops.find(function (l) { return l.ID === peminjaman.LAPTOP_ID; });
            if (laptop) laptop.STATUS = newLaptopStatus;

            saveToLocalStorage();
            hideLoading();
            showToast('Pengembalian berhasil diproses!', 'success');

            // UI Update
            showPage('dashboard');
            renderDashboard();
            refreshData();

            // Reset
            pendingPengembalianData = null;
        })
        .catch(function (err) {
            hideLoading();
            showToast('Sinkronisasi gagal, data disimpan secara lokal.', 'warning');

            // Fallback Logic (so Riwayat still works)
            AppState.pengembalian.push(returnRow);
            peminjaman.STATUS = 'Selesai';
            const laptopIdx = AppState.laptops.findIndex(l => l.ID === peminjaman.LAPTOP_ID);
            if (laptopIdx !== -1) AppState.laptops[laptopIdx].STATUS = newLaptopStatus;

            saveToLocalStorage();
            showPage('dashboard');
            renderDashboard();
            pendingPengembalianData = null;
        });
}

// ==========================================
// SPREADSHEET API HELPERS
// ==========================================
function appendToSheet(sheetName, rowObj) {
    var url = CONFIG.APPS_SCRIPT_URL;
    if (!url) return Promise.reject('No URL');

    // Use POST with text/plain to avoid CORS preflight, handled by e.postData.contents in GAS
    var payload = JSON.stringify({
        action: 'appendRow',
        sheet: sheetName,
        row: rowObj
    });

    return fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'text/plain;charset=utf-8',
        },
        body: payload
    }).then(function (res) { return res.json(); })
        .then(function (data) {
            if (!data.success) throw new Error(data.error || 'Append failed');
            return data;
        });
}

function updateSheetRow(sheetName, matchCol, matchVal, updates) {
    var url = CONFIG.APPS_SCRIPT_URL;
    if (!url) return Promise.reject('No URL');

    var payload = JSON.stringify({
        action: 'updateRow',
        sheet: sheetName,
        matchCol: matchCol,
        matchVal: matchVal,
        row: updates
    });

    return fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'text/plain;charset=utf-8',
        },
        body: payload
    }).then(function (res) { return res.json(); })
        .then(function (data) {
            if (!data.success) throw new Error(data.error || 'Update failed');
            return data;
        });
}

// ==========================================
// UI HELPERS
// ==========================================
function showToast(message, type) {
    type = type || 'success';
    const toast = document.getElementById('toast');
    const icon = document.getElementById('toastIcon');
    const msg = document.getElementById('toastMsg');
    if (!toast) return;

    // Remove old classes
    toast.className = 'toast';

    const icons = {
        success: 'fas fa-check-circle',
        error: 'fas fa-times-circle',
        warning: 'fas fa-exclamation-circle'
    };

    if (icon) icon.className = icons[type] || icons.success;
    if (msg) msg.textContent = message;

    toast.classList.add('toast-' + type, 'show');

    setTimeout(function () {
        toast.classList.remove('show');
    }, 3500);
}

function showLoading(text) {
    const overlay = document.getElementById('loadingOverlay');
    const txt = document.getElementById('loadingText');
    if (txt) txt.textContent = text || 'Memuat...';
    if (overlay) overlay.classList.add('show');
}

function hideLoading() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) overlay.classList.remove('show');
}

function showElement(id) {
    const el = document.getElementById(id);
    if (el) el.style.display = '';
}

function hideElement(id) {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
}

function formatDate(dateStr) {
    if (!dateStr || dateStr === '-') return '-';
    try {
        const str = String(dateStr);
        // Handle yyyy-MM-dd format
        if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
            const parts = str.split('-');
            const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
            return parseInt(parts[2]) + ' ' + months[parseInt(parts[1]) - 1] + ' ' + parts[0];
        }
        // Try Date parse
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return dateStr;
        return d.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
    } catch (e) {
        return dateStr;
    }
}

function formatTimestamp(date) {
    if (!date) return '';
    var d = new Date(date);
    if (isNaN(d.getTime())) return '';
    var local = new Date(d.getTime() - d.getTimezoneOffset() * 60000);
    return local.toISOString().slice(0, 19);
}

function parseTimestamp(value) {
    if (!value) return 0;
    var str = String(value || '').trim();
    if (!str) return 0;

    // Normalize common timestamp variants, including one-digit minutes/seconds
    var m = str.match(/^([0-9]{4}-[0-9]{2}-[0-9]{2})[ T]([0-9]{1,2}):([0-9]{1,2})(?::([0-9]{1,2}))?$/);
    if (m) {
        var datePart = m[1];
        var hour = m[2].padStart(2, '0');
        var minute = m[3].padStart(2, '0');
        var second = (m[4] || '00').padStart(2, '0');
        str = datePart + 'T' + hour + ':' + minute + ':' + second;
    }

    var ts = Date.parse(str);
    if (!isNaN(ts)) return ts;

    ts = Date.parse(str.replace(' ', 'T'));
    return isNaN(ts) ? 0 : ts;
}

function escapeHtml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function getPegawaiTim(pegawai) {
    if (!pegawai) return '';
    return pegawai['TIM'] ||
        pegawai['Tim'] ||
        pegawai['TIM DIVISI'] ||
        pegawai['TIM (DIVISI)'] ||
        pegawai['TIM(DIVISI)'] ||
        pegawai['DIVISI'] ||
        pegawai['Divisi'] ||
        pegawai.DIVISI ||
        '';
}

function normalizeLaptopStatus(status) {
    var s = String(status || '').trim().toLowerCase();

    if (s === 'dipinjam') return 'dipinjam';

    // ❌ HANYA rusak berat yang dianggap rusak
    if (s === 'rusak berat') return 'rusak';

    // ✅ rusak ringan tetap dianggap tersedia
    return 'tersedia';
}

function showStatusLaptopModal(statusType) {
    var modal = document.getElementById('statusLaptopModal');
    var title = document.getElementById('statusLaptopModalTitle');
    var body = document.getElementById('statusLaptopModalBody');
    var empty = document.getElementById('statusLaptopModalEmpty');
    if (!modal || !title || !body || !empty) return;

    var statusLabel = {
        tersedia: 'Tersedia',
        dipinjam: 'Dipinjam',
        rusak: 'Rusak'
    };
    title.textContent = 'Riwayat Laptop - ' + (statusLabel[statusType] || 'Semua');

    var filtered = AppState.laptops.filter(function (laptop) {
        return normalizeLaptopStatus(laptop.STATUS) === statusType;
    });

    body.innerHTML = '';
    if (!filtered.length) {
        empty.style.display = 'block';
    } else {
        empty.style.display = 'none';
        filtered.forEach(function (laptop, idx) {
            var tr = document.createElement('tr');
            tr.style.cursor = 'pointer';
            tr.title = 'Klik untuk lihat riwayat peminjam laptop ini';
            tr.onclick = function () {
                showLaptopBorrowHistoryModal(laptop.ID || '');
            };
            tr.innerHTML =
                '<td>' + (idx + 1) + '</td>' +
                '<td><strong>' + escapeHtml(laptop.ID || '-') + '</strong></td>' +
                '<td>' + escapeHtml(laptop.MERK || '-') + '</td>' +
                '<td>' + escapeHtml(laptop.TYPE || '-') + '</td>' +
                '<td>' + escapeHtml(laptop.NUP || '-') + '</td>' +
                '<td>' + escapeHtml(laptop.STATUS || '-') + '</td>';
            body.appendChild(tr);
        });
    }

    modal.style.display = 'flex';
    modal.classList.add('show');
}

function closeStatusLaptopModal() {
    var modal = document.getElementById('statusLaptopModal');
    if (!modal) return;
    modal.classList.remove('show');
    modal.style.display = 'none';
}

function showLaptopBorrowHistoryModal(laptopId) {
    var modal = document.getElementById('laptopBorrowHistoryModal');
    var title = document.getElementById('laptopBorrowHistoryTitle');
    var body = document.getElementById('laptopBorrowHistoryBody');
    var empty = document.getElementById('laptopBorrowHistoryEmpty');
    if (!modal || !title || !body || !empty || !laptopId) return;

    var laptop = AppState.laptops.find(function (l) { return l.ID === laptopId; });
    var laptopName = laptop ? ((laptop.MERK || '-') + ' ' + (laptop.TYPE || '-')) : laptopId;
    title.textContent = 'Riwayat Peminjam Laptop - ' + laptopId + ' (' + laptopName + ')';

    var rows = AppState.peminjaman
        .filter(function (p) { return String(p.LAPTOP_ID || '') === String(laptopId); })
        .sort(function (a, b) {
            var ta = Number(a.LAST_ACTION_TIME) || parseTimestamp(a.Timestamp) || new Date(a.TGL_PINJAM).getTime() || 0;
            var tb = Number(b.LAST_ACTION_TIME) || parseTimestamp(b.Timestamp) || new Date(b.TGL_PINJAM).getTime() || 0;
            return tb - ta;
        });

    body.innerHTML = '';
    if (!rows.length) {
        empty.style.display = 'block';
    } else {
        empty.style.display = 'none';
        rows.forEach(function (row, idx) {
            var kembali = AppState.pengembalian.find(function (k) {
                return String(k.PEMINJAMAN_ID || '') === String(row.ID || '');
            });
            var tglRealisasi = kembali ? (kembali.TGL_REALISASI_PENGEMBALIAN || '-') : (row.TGL_REALISASI_PENGEMBALIAN || '-');
            var kondisi = kembali ? (kembali.KONDISI_PENGEMBALIAN || kembali.KONDISI || '-') : '-';
            var catatan = kembali ? (kembali.CATATAN_PENGEMBALIAN || kembali.CATATAN || '-') : '-';

            // Apply condition styling
            var kondisiClass = 'condition-baik';
            if (kondisi === 'Rusak Ringan') {
                kondisiClass = 'condition-rusak-ringan';
            } else if (kondisi === 'Rusak Berat') {
                kondisiClass = 'condition-rusak-berat';
            }
            var kondisiHtml = kondisi === '-' ? '-' : '<span class="' + kondisiClass + '">' + escapeHtml(kondisi) + '</span>';

            var tr = document.createElement('tr');
            tr.innerHTML =
                '<td>' + (idx + 1) + '</td>' +
                '<td>' + escapeHtml(row.NAMA_PEMINJAM || '-') + '</td>' +
                '<td>' + formatDate(row.TGL_PINJAM) + '</td>' +
                '<td>' + formatDate(row.TGL_KEMBALI_RENCANA) + '</td>' +
                '<td>' + formatDate(tglRealisasi) + '</td>' +
                '<td>' + kondisiHtml + '</td>' +
                '<td>' + escapeHtml(catatan) + '</td>' +
                '<td>' + escapeHtml(row.STATUS || '-') + '</td>';
            body.appendChild(tr);
        });
    }

    closeStatusLaptopModal();
    modal.style.display = 'flex';
    modal.classList.add('show');
}

function closeLaptopBorrowHistoryModal() {
    var modal = document.getElementById('laptopBorrowHistoryModal');
    if (!modal) return;
    modal.classList.remove('show');
    modal.style.display = 'none';
}

function showDetailPeminjamanModal(peminjamanId) {
    // Find peminjaman data
    var peminjaman = AppState.peminjaman.find(function(p) { return p.ID === peminjamanId; });
    if (!peminjaman) {
        alert('Data peminjaman tidak ditemukan');
        return;
    }

    // Find laptop data
    var laptop = AppState.laptops.find(function(l) { return l.ID === peminjaman.LAPTOP_ID; });
    
    // Find pengembalian data
    var kembali = AppState.pengembalian.find(function(k) { return k.PEMINJAMAN_ID === peminjamanId; });

    // Populate data peminjam
    document.getElementById('detailPeminjamanNama').textContent = peminjaman.NAMA_PEMINJAM || '-';
    document.getElementById('detailPeminjamanNip').textContent = peminjaman.NIP || '-';
    document.getElementById('detailPeminjamanDivisi').textContent = peminjaman.DIVISI || '-';

    // Populate data laptop
    document.getElementById('detailPeminjamanLaptopId').textContent = peminjaman.LAPTOP_ID || '-';
    document.getElementById('detailPeminjamanMerk').textContent = laptop ? (laptop.MERK || '-') : '-';
    document.getElementById('detailPeminjamanType').textContent = laptop ? (laptop.TYPE || '-') : '-';

    // Populate data peminjaman
    document.getElementById('detailPeminjamanTglPinjam').textContent = formatDate(peminjaman.TGL_PINJAM);
    document.getElementById('detailPeminjamanTglRencana').textContent = formatDate(peminjaman.TGL_KEMBALI_RENCANA);

    // Populate keperluan
    document.getElementById('detailPeminjamanKeperluan').textContent = peminjaman.KEPERLUAN || '-';
    var deskripsiEl = document.getElementById('detailPeminjamanDeskripsi');
    if (peminjaman.DESKRIPSI_KEPERLUAN) {
        deskripsiEl.textContent = peminjaman.DESKRIPSI_KEPERLUAN;
        document.getElementById('detailPeminjamanDeskripsiRow').style.display = 'block';
    } else {
        document.getElementById('detailPeminjamanDeskripsiRow').style.display = 'none';
    }

    // Show/hide pengembalian section
    if (kembali) {
        document.getElementById('detailPeminjamanTglRealisasiRow').style.display = 'block';
        document.getElementById('detailPeminjamanTglRealisasi').textContent = formatDate(kembali.TGL_REALISASI_PENGEMBALIAN);
        
        // Show pengembalian section
        document.getElementById('detailPeminjamanPengembalianSection').style.display = 'block';
        
        // Kondisi with color coding
        var kondisi = kembali.KONDISI_PENGEMBALIAN || '-';
        var kondisiEl = document.getElementById('detailPeminjamanKondisi');
        kondisiEl.textContent = kondisi;
        
        // Add class based on condition
        kondisiEl.className = 'detail-value';
        if (kondisi === 'Baik') {
            kondisiEl.style.color = '#10b981';
            kondisiEl.style.fontWeight = '600';
        } else if (kondisi === 'Rusak Ringan' || kondisi === 'rusak-ringan') {
            kondisiEl.style.color = '#f59e0b';
            kondisiEl.style.fontWeight = '600';
        } else if (kondisi === 'Rusak Berat' || kondisi === 'rusak-berat') {
            kondisiEl.style.color = '#ef4444';
            kondisiEl.style.fontWeight = '600';
        }
        
        // Catatan
        document.getElementById('detailPeminjamanCatatan').textContent = kembali.CATATAN_PENGEMBALIAN || '-';
    } else {
        document.getElementById('detailPeminjamanTglRealisasiRow').style.display = 'none';
        document.getElementById('detailPeminjamanPengembalianSection').style.display = 'none';
    }

    // Show modal
    var modal = document.getElementById('detailPeminjamanModal');
    if (!modal) return;
    modal.style.display = 'flex';
    modal.classList.add('show');
}

function closeDetailPeminjamanModal() {
    var modal = document.getElementById('detailPeminjamanModal');
    if (!modal) return;
    modal.classList.remove('show');
    modal.style.display = 'none';
}
