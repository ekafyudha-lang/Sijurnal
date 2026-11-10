// Logika bisnis Si Jurnal. Berasal dari Code.gs (Google Apps Script); layanan Google ditiru oleh runtime.ts di atas Supabase.
export function createApp(G) {
const { SpreadsheetApp, PropertiesService, CacheService, LockService, DriveApp, Utilities, Session, Logger } = G;

/**
 * ===================================================================
 * SI JURNAL - BACKEND (Google Apps Script)
 * Arsitektur: GAS + Google Sheets sebagai database
 * ===================================================================
 *
 * SHEET YANG DIGUNAKAN (dibuat otomatis via setupDatabase()):
 *  - Users   : username | password_hash | role
 *  - COA     : account_code | account_name | category | normal_balance | status | opening_debit | opening_credit
 *              account_code berjenjang: "1" (ASET), "1.01" (Kas), "1.01.01" (sub-akun), dst.
 *              category otomatis mengikuti digit root kode (1=Aset,2=Kewajiban,3=Ekuitas,4=Pendapatan,5=Beban).
 *              Node yang punya sub-akun = header (rollup); node tanpa sub-akun = akun yang bisa dipakai di Jurnal.
 *  - Journal : date | ref_no | account_code | account_name | description | debit | credit | timestamp
 *  - Assets  : asset_id | asset_name | purchase_date | cost | salvage_value | useful_life_years | accumulated_dep | last_dep_date
 *  - Members / Savings / MemberTrans : modul koperasi (lihat bagian KOPERASI DAGANG di bawah)
 *
 * SESI LOGIN disimpan di CacheService (token -> username), berlaku 6 jam.
 *
 * CATATAN PENTING:
 *  - Script ini HARUS container-bound (dibuat lewat Extensions > Apps Script
 *    dari dalam sebuah Google Sheet), bukan standalone script. Kalau dibuat
 *    standalone lalu di-attach manual, getActiveSpreadsheet() bisa null dan
 *    semua fungsi simpan/baca akan gagal.
 *  - Setiap kali kode ini diubah, WAJIB Deploy > Manage deployments > Edit
 *    (pensil) > Version: New version > Deploy, supaya Web App yang sudah
 *    live memakai kode terbaru.
 *  - Untuk debug: buka Apps Script Editor > menu "Executions" (ikon jam) di
 *    sidebar kiri untuk melihat log & error setiap kali fungsi backend
 *    dipanggil dari halaman web.
 */

// ---------- KONFIGURASI ----------
var SESSION_DURATION_SEC = 6 * 60 * 60; // 6 jam

var SHEET_USERS = 'Users';
var SHEET_COA = 'COA';
var SHEET_JOURNAL = 'Journal';
var SHEET_ASSETS = 'Assets';

// Pemetaan kategori COA <-> digit root kode berjenjang.
var CATEGORY_TO_ROOT_CODE = { 'Aset': '1', 'Kewajiban': '2', 'Ekuitas': '3', 'Pendapatan': '4', 'Beban': '5' };
var ROOT_CODE_TO_CATEGORY = { '1': 'Aset', '2': 'Kewajiban', '3': 'Ekuitas', '4': 'Pendapatan', '5': 'Beban' };

// Nama akun khusus yang dipakai otomatis oleh sistem (posting penyusutan).
// Dicari berdasarkan NAMA (bukan kode tetap), supaya tidak rapuh terhadap perubahan/migrasi kode akun.
var NAMA_AKUN_BEBAN_PENYUSUTAN = 'Beban Penyusutan';
var NAMA_AKUN_AKUM_PENYUSUTAN = 'Akumulasi Penyusutan';


// ===================================================================
// UTIL
// ===================================================================

/**
 * Baris kosong berikutnya di sheet, sambil memastikan kapasitas baris cukup.
 * Sheet Google baru hanya berisi 1000 baris; menulis lewat getRange(...).setValues() melewati batas itu
 * akan error ("out of bounds"), jadi baris ditambah otomatis dalam blok besar sebelum menulis.
 */
function nextRow_(sheet, n) {
  var start = sheet.getLastRow() + 1;
  var need = start + (n || 1) - 1 + 200;          // sisakan cadangan 200 baris
  var max = sheet.getMaxRows();
  if (need > max) sheet.insertRowsAfter(max, Math.max(need - max, 2000));
  return start;
}

/** Ambil spreadsheet aktif. Melempar error jelas kalau script tidak ter-bind. */
function getSS_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      'Script Apps Script ini tidak terhubung ke Google Sheet manapun. ' +
      'Pastikan Anda membuka project ini lewat menu Extensions > Apps Script ' +
      'DARI DALAM sebuah Google Sheet (container-bound script), bukan dari script.google.com langsung.'
    );
  }
  return ss;
}

/** Ambil sheet by name. Melempar error jelas kalau sheet belum ada. */
function getSheetOrThrow_(name) {
  var sheet = getSS_().getSheetByName(name);
  if (!sheet) {
    throw new Error(
      'Sheet "' + name + '" belum ada. Jalankan fungsi setupDatabase() ' +
      'sekali dari Apps Script Editor terlebih dahulu.'
    );
  }
  return sheet;
}

function getOrCreateSheet_(name, headers) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row.join('') === '') continue; // skip baris kosong
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      obj[headers[c]] = row[c];
    }
    obj._row = i + 1; // nomor baris asli di sheet (1-based, termasuk header)
    out.push(obj);
  }
  return out;
}

function hashPassword_(password) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return digest.map(function (byte) {
    var v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function fmtDate_(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd');
}

function inRange_(dateStr, startDate, endDate) {
  if (!startDate && !endDate) return true;
  var d = new Date(dateStr);
  if (startDate && d < new Date(startDate)) return false;
  if (endDate && d > new Date(endDate + 'T23:59:59')) return false;
  return true;
}

/** Validasi token sesi. Mengembalikan username, atau null kalau tidak valid/expired. */

// ==================== ROLE & HAK AKSES ====================
var PERM_CATALOG_ = [
  { key: 'dashboard', label: 'Dashboard', group: 'Umum' },
  { key: 'members', label: 'Data Anggota', group: 'Koperasi' },
  { key: 'savings', label: 'Simpanan & Transaksi Anggota', group: 'Koperasi' },
  { key: 'shu', label: 'SHU', group: 'Koperasi' },
  { key: 'coa', label: 'Akun (COA) & Pengaturan', group: 'Keuangan' },
  { key: 'journal', label: 'Jurnal Kas/Umum & Bukti', group: 'Keuangan' },
  { key: 'reports', label: 'Laporan Keuangan', group: 'Keuangan' },
  { key: 'assets', label: 'Aset Tetap', group: 'Keuangan' },
  { key: 'items', label: 'Master Barang / Stok', group: 'Operasional' },
  { key: 'purchases_tbs', label: 'Pembelian TBS', group: 'Operasional' },
  { key: 'sales_tbs', label: 'Penjualan TBS', group: 'Operasional' },
  { key: 'purchases_barang', label: 'Pembelian Non TBS (Barang)', group: 'Operasional' },
  { key: 'sales_barang', label: 'Penjualan Non TBS (Barang)', group: 'Operasional' },
  { key: 'suppliers', label: 'Data Pemasok / Petani', group: 'Pemasok' },
  { key: 'garden_progress', label: 'Progres Kebun Pemasok', group: 'Pemasok' }
];
var DEFAULT_ROLES_ = {
  'Keuangan': { permissions: ['dashboard', 'members', 'savings', 'shu', 'coa', 'journal', 'reports', 'assets'] },
  'Pembelian TBS': { permissions: ['purchases_tbs'] },
  'Pembelian Non TBS': { permissions: ['purchases_barang'] }
};
var RESERVED_ROLES_ = ['administrator', 'anggota'];

function allPermKeys_() { return PERM_CATALOG_.map(function (p) { return p.key; }); }

function getRolesConfig_() {
  var raw = PropertiesService.getDocumentProperties().getProperty('roles_config');
  if (!raw) return JSON.parse(JSON.stringify(DEFAULT_ROLES_));
  try { return JSON.parse(raw) || {}; } catch (e) { return JSON.parse(JSON.stringify(DEFAULT_ROLES_)); }
}

function permsForRole_(role) {
  role = String(role || '');
  if (role === 'Administrator') return allPermKeys_();
  if (role === 'Anggota') return [];
  var cfg = getRolesConfig_();
  var valid = allPermKeys_();
  var list = (cfg[role] && cfg[role].permissions) || [];
  return list.filter(function (k) { return valid.indexOf(k) >= 0; });
}

// Unit yang boleh dipakai role: null = semua unit (Administrator / role tanpa pembatasan), atau array nama unit.
function unitsForRole_(role) {
  role = String(role || '');
  if (role === 'Administrator') return null;
  var cfg = getRolesConfig_();
  var u = cfg[role] && cfg[role].units;
  return Array.isArray(u) ? u : null;
}
function sessionUnits_(token) {
  return unitsForRole_(CacheService.getScriptCache().get('role_' + token));
}
// null jika unit boleh dipakai sesi ini; selain itu respons ditolak. Role terbatas wajib memilih unit miliknya.
function checkUnitAccess_(token, unit) {
  var allowed = sessionUnits_(token);
  if (allowed === null) return null;
  unit = String(unit || '');
  if (allowed.indexOf(unit) >= 0) return null;
  return { success: false, forbidden: true, message: unit ? 'Anda tidak memiliki akses ke unit "' + unit + '".' : 'Pilih unit koperasi. Akun Anda dibatasi pada unit: ' + allowed.join(', ') + '.' };
}
// Saring baris berdasarkan kolom unit sesuai batas unit sesi.
function filterByUnits_(token, rows) {
  var allowed = sessionUnits_(token);
  if (allowed === null) return rows;
  return rows.filter(function (r) { return allowed.indexOf(String(r.unit || '')) >= 0; });
}
function lookupTradeUnit_(sheetName, idField, id) {
  var row = sheetToObjects_(getTradeSheet_(sheetName)).filter(function (r) { return String(r[idField]) === String(id); })[0];
  return row ? String(row.unit || '') : '';
}

function sessionPerms_(token) {
  var role = CacheService.getScriptCache().get('role_' + token);
  return permsForRole_(role);
}

// Aksi edit/hapus hanya untuk Administrator (kembalikan null jika boleh).
function checkAdminAction_(token) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  if (CacheService.getScriptCache().get('role_' + token) !== 'Administrator') {
    return { success: false, forbidden: true, message: 'Hanya Administrator yang boleh mengedit atau menghapus data.' };
  }
  return null;
}

function forbiddenResponse_() {
  return { success: false, forbidden: true, message: 'Anda tidak memiliki hak akses untuk fitur ini. Hubungi Administrator.' };
}

// Kembalikan null jika boleh; kalau tidak, objek respons (sesi habis / dilarang).
// anyOf = salah satu hak yang cukup.
function checkPerm_(token, anyOf) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  var mine = sessionPerms_(token);
  for (var i = 0; i < anyOf.length; i++) if (mine.indexOf(anyOf[i]) >= 0) return null;
  return forbiddenResponse_();
}

// Hak pembelian/penjualan menurut jenis (TBS | Barang). prefix: 'purchases' | 'sales'
function checkTradePerm_(token, prefix, kind) {
  return checkPerm_(token, [prefix + '_' + (kindOf_(kind) === 'Barang' ? 'barang' : 'tbs')]);
}

function lookupTradeKind_(sheetName, idField, id) {
  var row = sheetToObjects_(getTradeSheet_(sheetName)).filter(function (r) { return String(r[idField]) === String(id); })[0];
  return row ? kindOf_(row.kind) : 'TBS';
}

// Jenis (TBS/Barang) yang boleh dilihat pada daftar, berdasarkan hak.
function allowedKinds_(token, prefixes) {
  var mine = sessionPerms_(token), out = [];
  if (prefixes.some(function (p) { return mine.indexOf(p + '_tbs') >= 0; })) out.push('TBS');
  if (prefixes.some(function (p) { return mine.indexOf(p + '_barang') >= 0; })) out.push('Barang');
  return out;
}

function requireSession_(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var username = cache.get('sess_' + token);
  if (!username) return null;
  // Sesi berrole "Anggota" TIDAK boleh memakai fungsi staf/administrasi apa pun.
  if (cache.get('role_' + token) === 'Anggota') return null;
  return username;
}

/** Sesi apa pun (staf maupun anggota). Hanya untuk fungsi yang aman bagi semua role, mis. ubah password sendiri. */
function requireAnySession_(token) {
  if (!token) return null;
  return CacheService.getScriptCache().get('sess_' + token);
}

/**
 * Sesi khusus anggota. Mengembalikan { username, member_id } atau null.
 * member_id SELALU diambil dari sesi di server (bukan dari input browser), dan dicek ulang ke sheet Users
 * setiap panggilan, jadi akun yang dihapus langsung tidak bisa dipakai lagi.
 */
function requireMemberSession_(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var username = cache.get('sess_' + token);
  if (!username || cache.get('role_' + token) !== 'Anggota') return null;
  var memberId = cache.get('mem_' + token);
  if (!memberId) return null;
  var sheet = getSS_().getSheetByName(SHEET_USERS);
  if (!sheet) return null;
  var ok = sheetToObjects_(sheet).some(function (u) {
    return String(u.username).toLowerCase() === String(username).toLowerCase() &&
           u.role === 'Anggota' && String(u.member_id) === String(memberId);
  });
  return ok ? { username: username, member_id: memberId } : null;
}

/** Hanya role Administrator. Mengembalikan username atau null. */
function requireAdmin_(token) {
  var username = requireSession_(token);
  if (!username) return null;
  return CacheService.getScriptCache().get('role_' + token) === 'Administrator' ? username : null;
}

function sessionExpiredResponse_() {
  return { success: false, sessionExpired: true, message: 'Sesi Anda telah berakhir. Silakan login kembali.' };
}

/** Jalankan fn di dalam lock, supaya dua simpan bersamaan tidak saling menimpa baris. */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  var gotLock = lock.tryLock(10000); // tunggu maksimal 10 detik
  if (!gotLock) {
    throw new Error('Server sedang sibuk memproses permintaan lain, coba lagi sebentar.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ===================================================================
// DIAGNOSTIK (opsional, dipanggil manual dari Apps Script Editor
// untuk memastikan koneksi ke spreadsheet & sheet sudah benar)
// ===================================================================
function runDiagnostics() {
  var report = [];
  try {
    var ss = getSS_();
    report.push('OK: Terhubung ke spreadsheet "' + ss.getName() + '"');
  } catch (err) {
    report.push('GAGAL: ' + err.message);
    Logger.log(report.join('\n'));
    return report.join('\n');
  }
  [SHEET_USERS, SHEET_COA, SHEET_JOURNAL, SHEET_ASSETS, SHEET_MEMBERS, SHEET_SAVINGS, SHEET_MEMBER_TRANS, SHEET_ITEMS, SHEET_PURCHASES, SHEET_SALES].forEach(function (name) {
    var sheet = getSS_().getSheetByName(name);
    if (sheet) {
      report.push('OK: Sheet "' + name + '" ada, ' + (sheet.getLastRow() - 1) + ' baris data.');
    } else {
      report.push('BELUM ADA: Sheet "' + name + '" — jalankan setupDatabase().');
    }
  });
  var msg = report.join('\n');
  Logger.log(msg);
  return msg;
}

/**
 * DEBUG: Tes baca sheet COA langsung, TANPA lewat token/sesi/web app.
 * Jalankan manual: pilih "debugGetCOARaw" di dropdown toolbar > klik Jalankan (▶),
 * lalu lihat hasilnya di panel "Log eksekusi" atau console di bawah editor.
 */
function debugGetCOARaw() {
  try {
    var ss = getSS_();
    Logger.log('Spreadsheet aktif: "' + ss.getName() + '" (ID: ' + ss.getId() + ')');

    var sheet = ss.getSheetByName(SHEET_COA);
    if (!sheet) {
      Logger.log('GAGAL: Sheet "COA" tidak ditemukan di spreadsheet ini.');
      return;
    }
    Logger.log('Sheet COA ditemukan. Last row: ' + sheet.getLastRow() + ', Last col: ' + sheet.getLastColumn());

    var raw = sheet.getDataRange().getValues();
    Logger.log('Data mentah (getValues), jumlah baris termasuk header: ' + raw.length);
    Logger.log(JSON.stringify(raw));

    var data = sheetToObjects_(sheet);
    Logger.log('Hasil sheetToObjects_: ' + data.length + ' baris');
    Logger.log(JSON.stringify(data));
  } catch (err) {
    Logger.log('ERROR: ' + err.message + ' | stack: ' + err.stack);
  }
}

// ===================================================================
// SETUP DATABASE
// ===================================================================
function setupDatabase() {
  try {
    var usersSheet = getOrCreateSheet_(SHEET_USERS, ['username', 'password_hash', 'role', 'member_id']);
    ensureUsersColumns_(usersSheet);
    if (usersSheet.getLastRow() < 2) {
      usersSheet.appendRow(['admin', hashPassword_('Admin@2026'), 'Administrator']);
    }

    // COA berjenjang: account_code = "1", "1.01", "1.01.01", dst.
    // category & normal_balance disimpan per baris; opening_debit/opening_credit = saldo awal per akun.
    var coaSheet = getOrCreateSheet_(SHEET_COA, ['account_code', 'account_name', 'category', 'normal_balance', 'status', 'opening_debit', 'opening_credit']);
    // PENTING: paksa kolom account_code (kolom A) selalu jadi TEKS, supaya nilai
    // seperti "1.01.01" tidak otomatis diubah Google Sheets jadi tanggal/angka.
    coaSheet.getRange(2, 1, 998, 1).setNumberFormat('@');
    ensureCoaColumns_(coaSheet);
    if (coaSheet.getLastRow() < 2) {
      var defaultCOA = [
        ['1', 'ASET', 'Aset', 'Debit', 'Aktif', 0, 0],
        ['1.01', 'Kas', 'Aset', 'Debit', 'Aktif', 0, 0],
        ['1.02', 'Bank', 'Aset', 'Debit', 'Aktif', 0, 0],
        ['1.03', 'Piutang Usaha', 'Aset', 'Debit', 'Aktif', 0, 0],
        ['1.04', 'Piutang Anggota', 'Aset', 'Debit', 'Aktif', 0, 0],
        ['1.05', 'Persediaan Barang Dagang', 'Aset', 'Debit', 'Aktif', 0, 0],
        ['1.06', 'Peralatan', 'Aset', 'Debit', 'Aktif', 0, 0],
        ['1.07', NAMA_AKUN_AKUM_PENYUSUTAN + ' Aset Tetap', 'Aset', 'Kredit', 'Aktif', 0, 0],
        ['2', 'KEWAJIBAN', 'Kewajiban', 'Kredit', 'Aktif', 0, 0],
        ['2.01', 'Utang Usaha', 'Kewajiban', 'Kredit', 'Aktif', 0, 0],
        ['2.02', 'Simpanan Sukarela Anggota', 'Kewajiban', 'Kredit', 'Aktif', 0, 0],
        ['3', 'EKUITAS', 'Ekuitas', 'Kredit', 'Aktif', 0, 0],
        ['3.01', 'Simpanan Pokok', 'Ekuitas', 'Kredit', 'Aktif', 0, 0],
        ['3.02', 'Simpanan Wajib', 'Ekuitas', 'Kredit', 'Aktif', 0, 0],
        ['3.03', 'Dana Cadangan', 'Ekuitas', 'Kredit', 'Aktif', 0, 0],
        ['4', 'PENDAPATAN', 'Pendapatan', 'Kredit', 'Aktif', 0, 0],
        ['4.01', 'Penjualan Barang Dagang', 'Pendapatan', 'Kredit', 'Aktif', 0, 0],
        ['5', 'BEBAN', 'Beban', 'Debit', 'Aktif', 0, 0],
        ['5.01', 'Harga Pokok Penjualan', 'Beban', 'Debit', 'Aktif', 0, 0],
        ['5.02', 'Beban Gaji', 'Beban', 'Debit', 'Aktif', 0, 0],
        ['5.03', 'Beban Sewa', 'Beban', 'Debit', 'Aktif', 0, 0],
        ['5.04', NAMA_AKUN_BEBAN_PENYUSUTAN, 'Beban', 'Debit', 'Aktif', 0, 0]
      ];
      defaultCOA.forEach(function (row) { coaSheet.appendRow(row); });
    }

    getOrCreateSheet_(SHEET_JOURNAL, ['date', 'ref_no', 'account_code', 'account_name', 'description', 'debit', 'credit', 'timestamp', 'unit']);
    var assetsSheet = getOrCreateSheet_(SHEET_ASSETS, ['asset_id', 'asset_name', 'purchase_date', 'cost', 'salvage_value', 'useful_life_years', 'accumulated_dep', 'last_dep_date']);
    assetsSheet.getRange(2, 1, 998, 1).setNumberFormat('@'); // asset_id juga selalu teks
    ensureCoopSheets_(); // sheet koperasi: Members, Savings, MemberTrans

    Logger.log('setupDatabase: OK, spreadsheet = ' + getSS_().getName());
    return { success: true, message: 'Database berhasil disiapkan. Login default: admin / Admin@2026' };
  } catch (err) {
    Logger.log('setupDatabase ERROR: ' + err.message);
    return { success: false, message: 'Gagal setup database: ' + err.message };
  }
}

// ===================================================================
// AUTH
// ===================================================================
function loginUser(username, password) {
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) return { success: false, message: 'Database belum disiapkan. Hubungi pengelola aplikasi.' };

    var users = sheetToObjects_(sheet);
    var hash = hashPassword_(password);
    var found = users.filter(function (u) {
      return String(u.username).toLowerCase() === String(username).toLowerCase() && u.password_hash === hash;
    })[0];

    if (!found) return { success: false, message: 'Username atau password salah.' };

    var role = String(found.role || '');
    var memberId = String(found.member_id || '');
    var memberName = '';
    if (role === 'Anggota') {
      var m = memberId ? sheetToObjects_(getCoopSheet_(SHEET_MEMBERS)).filter(function (r) { return String(r.member_id) === memberId; })[0] : null;
      if (!m) return { success: false, message: 'Akun ini belum terhubung ke data anggota. Hubungi pengurus koperasi.' };
      memberName = String(m.name || '');
    }

    var token = Utilities.getUuid();
    var cache = CacheService.getScriptCache();
    cache.put('sess_' + token, found.username, SESSION_DURATION_SEC);
    cache.put('role_' + token, role, SESSION_DURATION_SEC);
    if (role === 'Anggota') cache.put('mem_' + token, memberId, SESSION_DURATION_SEC);

    Logger.log('loginUser: sukses untuk "' + found.username + '" (' + role + ')');
    return { success: true, token: token, user: { username: found.username, role: role, member_name: memberName, perms: permsForRole_(role), units: unitsForRole_(role) } };
  } catch (err) {
    Logger.log('loginUser ERROR: ' + err.message);
    return { success: false, message: 'Error login: ' + err.message };
  }
}

// ===================================================================
// CHART OF ACCOUNTS (COA) - berjenjang: "1", "1.01", "1.01.01", dst.
// Node yang punya sub-akun = header (nilainya rollup dari anaknya).
// Node tanpa sub-akun = akun yang bisa dipilih di Jurnal.
// ===================================================================

/** Migrasi: pastikan sheet COA punya kolom opening_debit/opening_credit (utk instalasi lama). */
function ensureCoaColumns_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 5);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('opening_debit') === -1) sheet.getRange(1, 6).setValue('opening_debit');
  if (headers.indexOf('opening_credit') === -1) sheet.getRange(1, 7).setValue('opening_credit');
}

/** Natural-sort comparator untuk kode berjenjang seperti "1.01.02" vs "1.1". */
function coaCodeCompare_(a, b) {
  var pa = String(a).split('.').map(Number);
  var pb = String(b).split('.').map(Number);
  var len = Math.max(pa.length, pb.length);
  for (var i = 0; i < len; i++) {
    var na = pa[i] === undefined ? -1 : pa[i];
    var nb = pb[i] === undefined ? -1 : pb[i];
    if (na !== nb) return na - nb;
  }
  return 0;
}

/** Kode child berikutnya di bawah parentCode pada hirarki COA. */
function nextCoaChildCode_(codes, parentCode) {
  var siblingNumbers = [];
  codes.forEach(function (code) {
    var isDirectChild = parentCode
      ? (code.indexOf(parentCode + '.') === 0 && code.substring(parentCode.length + 1).indexOf('.') === -1)
      : (code.indexOf('.') === -1);
    if (isDirectChild) {
      var lastSeg = code.substring(code.lastIndexOf('.') + 1);
      siblingNumbers.push(Number(lastSeg) || 0);
    }
  });
  var nextNum = siblingNumbers.length ? Math.max.apply(null, siblingNumbers) + 1 : (parentCode ? 1 : codes.filter(function (c) { return c.indexOf('.') === -1; }).length + 1);
  var lastSegStr = parentCode ? (nextNum < 10 ? '0' + nextNum : String(nextNum)) : String(nextNum);
  return parentCode ? (parentCode + '.' + lastSegStr) : lastSegStr;
}

/** Kategori otomatis dari digit root kode ("1.01.01" -> root "1" -> "Aset"). */
function deriveCoaCategory_(code) {
  var root = String(code).split('.')[0];
  return ROOT_CODE_TO_CATEGORY[root] || 'Lainnya';
}

function defaultNormalBalance_(category) {
  return (category === 'Aset' || category === 'Beban') ? 'Debit' : 'Kredit';
}

/** Saldo awal akun (opening_debit/opening_credit) dikonversi ke satu angka sesuai saldo normalnya. */
function openingSeed_(coaRow, isDebitNormal) {
  var od = Number(coaRow.opening_debit) || 0;
  var oc = Number(coaRow.opening_credit) || 0;
  return isDebitNormal ? (od - oc) : (oc - od);
}

function getCOAList(token) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  try {
    var sheet = getSheetOrThrow_(SHEET_COA);
    ensureCoaColumns_(sheet);
    var rows = sheetToObjects_(sheet).map(function (r) {
      return {
        account_code: String(r.account_code), account_name: r.account_name,
        category: r.category, normal_balance: r.normal_balance, status: r.status,
        own_debit: Number(r.opening_debit) || 0,
        own_credit: Number(r.opening_credit) || 0
      };
    });
    rows.sort(function (a, b) { return coaCodeCompare_(a.account_code, b.account_code); });

    var childrenCount = {};
    rows.forEach(function (r) { childrenCount[r.account_code] = 0; });
    rows.forEach(function (r) {
      var parent = r.account_code.indexOf('.') === -1 ? null : r.account_code.substring(0, r.account_code.lastIndexOf('.'));
      if (parent && childrenCount.hasOwnProperty(parent)) childrenCount[parent]++;
    });

    function rollup(code, field) {
      var node = rows.filter(function (r) { return r.account_code === code; })[0];
      if (!node) return 0;
      if (childrenCount[code] === 0) return field === 'debit' ? node.own_debit : node.own_credit;
      var total = 0;
      rows.forEach(function (r) {
        if (r.account_code === code || r.account_code.indexOf(code + '.') !== 0) return;
        var directParent = r.account_code.substring(0, r.account_code.lastIndexOf('.'));
        if (directParent === code) total += rollup(r.account_code, field);
      });
      return total;
    }

    var data = rows.map(function (r) {
      var isLeaf = childrenCount[r.account_code] === 0;
      return {
        account_code: r.account_code,
        account_name: r.account_name,
        category: r.category,
        normal_balance: r.normal_balance,
        status: r.status,
        level: r.account_code.split('.').length - 1,
        is_leaf: isLeaf,
        opening_debit: isLeaf ? r.own_debit : rollup(r.account_code, 'debit'),
        opening_credit: isLeaf ? r.own_credit : rollup(r.account_code, 'credit')
      };
    });

    var totalDebit = 0, totalCredit = 0;
    data.forEach(function (r) { if (r.level === 0) { totalDebit += r.opening_debit; totalCredit += r.opening_credit; } });

    return { success: true, data: data, totalDebit: totalDebit, totalCredit: totalCredit };
  } catch (err) {
    Logger.log('getCOAList ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** Tambah akun/sub-akun baru. payload = { parent_code, name, normal_balance }. */
function saveAccount(token, payload) {
  var deny_ = checkPerm_(token, ['coa']); if (deny_) return deny_;
  try {
    if (!payload.name) return { success: false, message: 'Nama akun wajib diisi.' };
    return withLock_(function () {
      var sheet = getSheetOrThrow_(SHEET_COA);
      ensureCoaColumns_(sheet);
      var codes = sheetToObjects_(sheet).map(function (r) { return String(r.account_code); });
      var parentCode = payload.parent_code ? String(payload.parent_code) : '';
      var newCode = nextCoaChildCode_(codes, parentCode);
      var category = deriveCoaCategory_(newCode);
      var normalBalance = payload.normal_balance || defaultNormalBalance_(category);

      var newRow = nextRow_(sheet);
      sheet.getRange(newRow, 1).setNumberFormat('@');
      sheet.getRange(newRow, 1, 1, 7).setValues([[newCode, payload.name, category, normalBalance, 'Aktif', 0, 0]]);
      SpreadsheetApp.flush();
      Logger.log('saveAccount: tambah ' + newCode + ' (' + payload.name + ')');
      return { success: true, message: 'Akun "' + payload.name + '" berhasil ditambahkan sebagai "' + newCode + '".', code: newCode };
    });
  } catch (err) {
    Logger.log('saveAccount ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** Ubah nama/saldo normal akun yang sudah ada (kode & kategori tidak berubah). */
function renameAccount(token, code, newName, newNormalBalance) {
  var deny_ = checkPerm_(token, ['coa']); if (deny_) return deny_;
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    if (!newName) return { success: false, message: 'Nama akun wajib diisi.' };
    return withLock_(function () {
      var sheet = getSheetOrThrow_(SHEET_COA);
      var rows = sheetToObjects_(sheet);
      var existing = rows.filter(function (r) { return String(r.account_code) === String(code); })[0];
      if (!existing) return { success: false, message: 'Akun tidak ditemukan.' };
      sheet.getRange(existing._row, 2).setValue(newName);
      if (newNormalBalance) sheet.getRange(existing._row, 4).setValue(newNormalBalance);
      Logger.log('renameAccount: ' + code + ' -> ' + newName);
      return { success: true, message: 'Akun berhasil diperbarui.' };
    });
  } catch (err) {
    Logger.log('renameAccount ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function updateSaldoAwalBulk(token, updates) {
  var deny_ = checkPerm_(token, ['coa']); if (deny_) return deny_;
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    return withLock_(function () {
      var sheet = getSheetOrThrow_(SHEET_COA);
      ensureCoaColumns_(sheet);
      var rows = sheetToObjects_(sheet);
      var rowByCode = {};
      rows.forEach(function (r) { rowByCode[String(r.account_code)] = r._row; });

      updates.forEach(function (u) {
        var rowNum = rowByCode[String(u.code)];
        if (!rowNum) return;
        var d = Number(u.opening_debit) || 0;
        var c = Number(u.opening_credit) || 0;
        sheet.getRange(rowNum, 6, 1, 2).setValues([[d, c]]);
      });
      SpreadsheetApp.flush();
      Logger.log('updateSaldoAwalBulk: ' + updates.length + ' akun diperbarui');
      return { success: true, message: 'Saldo awal berhasil disimpan.' };
    });
  } catch (err) {
    Logger.log('updateSaldoAwalBulk ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** Hapus akun. Ditolak kalau masih punya sub-akun, atau sudah dipakai di transaksi Jurnal. */
function deleteAccount(token, code) {
  var deny_ = checkPerm_(token, ['coa']); if (deny_) return deny_;
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    return withLock_(function () {
      var sheet = getSheetOrThrow_(SHEET_COA);
      var rows = sheetToObjects_(sheet);
      var hasChild = rows.some(function (r) { return String(r.account_code).indexOf(String(code) + '.') === 0; });
      if (hasChild) return { success: false, message: 'Tidak bisa menghapus: akun ini masih punya sub-akun. Hapus sub-akunnya terlebih dahulu.' };
      var existing = rows.filter(function (r) { return String(r.account_code) === String(code); })[0];
      if (!existing) return { success: false, message: 'Akun tidak ditemukan.' };

      var journalRows = sheetToObjects_(getSheetOrThrow_(SHEET_JOURNAL));
      var used = journalRows.some(function (r) { return String(r.account_code) === String(code); });
      if (used) return { success: false, message: 'Akun ini sudah dipakai di transaksi Jurnal, tidak bisa dihapus. Nonaktifkan lewat Edit kalau perlu.' };

      sheet.deleteRow(existing._row);
      Logger.log('deleteAccount: hapus ' + code);
      return { success: true, message: 'Akun berhasil dihapus.' };
    });
  } catch (err) {
    Logger.log('deleteAccount ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/**
 * Import massal akun COA dari file CSV yang diupload user.
 * rows = [{ code, name, normal_balance, opening_debit, opening_credit }, ...]
 * - code WAJIB format kode berjenjang yang valid (mis. "1", "1.01", "1.01.01"),
 *   dan sebaiknya menyertakan akun induk (header)-nya juga supaya pohon COA lengkap.
 * - Kalau code sudah ada di COA -> baris tsb DIPERBARUI (nama/saldo normal/saldo awal).
 * - Kalau code belum ada -> ditambahkan sebagai akun baru.
 * - category otomatis diturunkan dari digit root code, tidak perlu diisi di CSV.
 */
function importCoaCSV(token, rows) {
  var deny_ = checkPerm_(token, ['coa']); if (deny_) return deny_;
  try {
    if (!rows || !rows.length) return { success: false, message: 'Tidak ada data untuk diimpor.' };
    return withLock_(function () {
      var sheet = getSheetOrThrow_(SHEET_COA);
      ensureCoaColumns_(sheet);
      var existingRows = sheetToObjects_(sheet);
      var rowByCode = {};
      existingRows.forEach(function (r) { rowByCode[String(r.account_code)] = r._row; });

      var codeRegex = /^[1-5](\.\d+)*$/;
      var added = 0, updated = 0, skipped = [];

      rows.forEach(function (item) {
        var code = String(item.code || '').trim();
        var name = String(item.name || '').trim();
        if (!code || !name) { skipped.push(code || '(baris kosong)'); return; }
        if (!codeRegex.test(code)) { skipped.push(code + ' (format kode tidak valid)'); return; }

        var category = deriveCoaCategory_(code);
        var normalBalance = (item.normal_balance === 'Debit' || item.normal_balance === 'Kredit')
          ? item.normal_balance : defaultNormalBalance_(category);
        var openingDebit = Number(item.opening_debit) || 0;
        var openingCredit = Number(item.opening_credit) || 0;

        if (rowByCode[code]) {
          var rowNum = rowByCode[code];
          sheet.getRange(rowNum, 2, 1, 6).setValues([[name, category, normalBalance, 'Aktif', openingDebit, openingCredit]]);
          updated++;
        } else {
          var newRow = nextRow_(sheet);
          sheet.getRange(newRow, 1).setNumberFormat('@');
          sheet.getRange(newRow, 1, 1, 7).setValues([[code, name, category, normalBalance, 'Aktif', openingDebit, openingCredit]]);
          rowByCode[code] = newRow;
          added++;
        }
      });
      SpreadsheetApp.flush();
      var msg = 'Impor selesai: ' + added + ' akun baru ditambahkan, ' + updated + ' akun diperbarui.';
      if (skipped.length) msg += ' Dilewati (' + skipped.length + '): ' + skipped.slice(0, 10).join(', ') + (skipped.length > 10 ? ', ...' : '');
      Logger.log('importCoaCSV: ' + msg);
      return { success: true, message: msg, added: added, updated: updated, skipped: skipped };
    });
  } catch (err) {
    Logger.log('importCoaCSV ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/**
 * MIGRASI SATU KALI: ubah COA lama (format flat "1-1001", kategori bebas)
 * menjadi hirarki berjenjang ("1.01", dst), dan REMAP semua kode akun lama
 * di sheet Journal ke kode baru supaya data transaksi historis tetap valid.
 * Jalankan manual sekali dari Apps Script Editor: pilih "migrateCoaToHierarchical"
 * di dropdown toolbar > klik Jalankan. AMAN dijalankan hanya SEKALI -
 * kalau COA sudah berjenjang (semua kode sudah mengandung format "angka.angka"
 * atau cuma "1"-"5"), fungsi ini akan berhenti tanpa mengubah apa pun.
 */
function migrateCoaToHierarchical() {
  var coaSheet = getSheetOrThrow_(SHEET_COA);
  var oldRows = sheetToObjects_(coaSheet);

  var alreadyHierarchical = oldRows.length > 0 && oldRows.every(function (r) {
    return /^[1-5](\.\d+)*$/.test(String(r.account_code));
  });
  if (alreadyHierarchical) {
    var skipMsg = 'COA sudah dalam format berjenjang. Tidak ada yang perlu dimigrasi.';
    Logger.log('migrateCoaToHierarchical: ' + skipMsg);
    return skipMsg;
  }

  // Kelompokkan akun lama per kategori, urutan sesuai baris asli.
  var byCategory = {};
  oldRows.forEach(function (r) {
    var cat = r.category || 'Lainnya';
    if (!byCategory[cat]) byCategory[cat] = [];
    byCategory[cat].push(r);
  });

  var oldToNew = {};
  var newRows = [];
  ['Aset', 'Kewajiban', 'Ekuitas', 'Pendapatan', 'Beban'].forEach(function (cat) {
    var rootCode = CATEGORY_TO_ROOT_CODE[cat];
    newRows.push([rootCode, cat === 'Aset' ? 'ASET' : cat === 'Kewajiban' ? 'KEWAJIBAN' : cat === 'Ekuitas' ? 'EKUITAS' : cat === 'Pendapatan' ? 'PENDAPATAN' : 'BEBAN', cat, defaultNormalBalance_(cat), 'Aktif', 0, 0]);
    (byCategory[cat] || []).forEach(function (r, idx) {
      var newCode = rootCode + '.' + (idx + 1 < 10 ? '0' + (idx + 1) : (idx + 1));
      oldToNew[String(r.account_code)] = newCode;
      newRows.push([newCode, r.account_name, cat, r.normal_balance || defaultNormalBalance_(cat), r.status || 'Aktif', Number(r.opening_debit) || 0, Number(r.opening_credit) || 0]);
    });
  });
  // Kategori lain di luar 5 kategori standar, taruh sebagai root sendiri berurutan mulai dari "6".
  var extraRootNum = 6;
  Object.keys(byCategory).forEach(function (cat) {
    if (['Aset', 'Kewajiban', 'Ekuitas', 'Pendapatan', 'Beban'].indexOf(cat) !== -1) return;
    var rootCode = String(extraRootNum++);
    newRows.push([rootCode, cat.toUpperCase(), cat, 'Debit', 'Aktif', 0, 0]);
    byCategory[cat].forEach(function (r, idx) {
      var newCode = rootCode + '.' + (idx + 1 < 10 ? '0' + (idx + 1) : (idx + 1));
      oldToNew[String(r.account_code)] = newCode;
      newRows.push([newCode, r.account_name, cat, r.normal_balance || 'Debit', r.status || 'Aktif', Number(r.opening_debit) || 0, Number(r.opening_credit) || 0]);
    });
  });

  // Tulis ulang sheet COA dari awal dengan struktur baru.
  var lastRow = coaSheet.getLastRow();
  if (lastRow > 1) coaSheet.getRange(2, 1, lastRow - 1, coaSheet.getLastColumn()).clearContent();
  ensureCoaColumns_(coaSheet);
  coaSheet.getRange(2, 1, 998, 1).setNumberFormat('@');
  if (newRows.length) coaSheet.getRange(2, 1, newRows.length, 7).setValues(newRows);
  SpreadsheetApp.flush();

  // Remap kode akun lama -> baru di semua baris Journal, supaya transaksi historis tetap valid.
  var journalSheet = getSheetOrThrow_(SHEET_JOURNAL);
  var journalValues = journalSheet.getDataRange().getValues();
  var remapped = 0;
  for (var i = 1; i < journalValues.length; i++) {
    var oldCode = String(journalValues[i][2]);
    if (oldToNew.hasOwnProperty(oldCode)) {
      journalValues[i][2] = oldToNew[oldCode];
      remapped++;
    }
  }
  if (journalValues.length > 1) {
    journalSheet.getRange(1, 1, journalValues.length, journalValues[0].length).setValues(journalValues);
    journalSheet.getRange(2, 3, Math.max(journalValues.length - 1, 0), 1).setNumberFormat('@');
  }
  SpreadsheetApp.flush();

  var summary = 'Migrasi selesai: ' + oldRows.length + ' akun dipindah ke format berjenjang, ' + remapped + ' baris transaksi Jurnal di-remap ke kode baru. Peta kode lama->baru:\n' +
    Object.keys(oldToNew).map(function (k) { return k + ' -> ' + oldToNew[k]; }).join('\n');
  Logger.log('migrateCoaToHierarchical: ' + summary);
  return summary;
}


// ===================================================================
// JURNAL UMUM
// ===================================================================
function getJournals(token, startDate, endDate) {
  var deny_ = checkPerm_(token, ['journal', 'reports']); if (deny_) return deny_;
  try {
    var sheet = getSheetOrThrow_(SHEET_JOURNAL);
    var rows = sheetToObjects_(sheet)
      .filter(function (r) { return inRange_(fmtDate_(r.date), startDate, endDate); })
      .map(function (r) {
        return {
          date: fmtDate_(r.date), ref_no: r.ref_no, account_code: r.account_code,
          account_name: r.account_name, description: r.description,
          debit: Number(r.debit) || 0, credit: Number(r.credit) || 0, unit: String(r.unit || '')
        };
      })
      .sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
    rows = filterByUnits_(token, rows);
    return { success: true, data: rows };
  } catch (err) {
    Logger.log('getJournals ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function saveJournal(token, payload) {
  var deny_ = checkPerm_(token, ['journal']); if (deny_) return deny_;
  try {
    if (!payload.date || !payload.ref_no) {
      return { success: false, message: 'Tanggal dan No. Referensi wajib diisi.' };
    }
    if (!payload.entries || payload.entries.length < 2) {
      return { success: false, message: 'Jurnal minimal harus punya 2 baris transaksi.' };
    }
    var totalDebit = 0, totalCredit = 0;
    payload.entries.forEach(function (e) { totalDebit += Number(e.debit) || 0; totalCredit += Number(e.credit) || 0; });

    if (totalDebit !== totalCredit) {
      return { success: false, message: 'Jurnal tidak balance! Total Debit (' + totalDebit + ') harus sama dengan Total Kredit (' + totalCredit + ').' };
    }
    if (totalDebit === 0) {
      return { success: false, message: 'Nilai transaksi tidak boleh 0.' };
    }

    var pc_ = periodClosedCheck_(payload.date); if (pc_) return pc_;
    var uc = checkUnit_(payload.unit);
    if (uc.error) return { success: false, message: uc.error };
    var ua_ = checkUnitAccess_(token, uc.unit); if (ua_) return ua_;
    return withLock_(function () {
      var sheet = getSheetOrThrow_(SHEET_JOURNAL);
      ensureJournalUnitCol_(sheet);
      var now = new Date();
      payload.entries.forEach(function (e) {
        sheet.appendRow([payload.date, payload.ref_no, e.account_code, e.account_name, payload.description, Number(e.debit) || 0, Number(e.credit) || 0, now, uc.unit]);
      });
      SpreadsheetApp.flush();
      Logger.log('saveJournal: ' + payload.ref_no + ' (' + payload.entries.length + ' baris)');
      return { success: true, message: 'Jurnal berhasil disimpan.' };
    });
  } catch (err) {
    Logger.log('saveJournal ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deleteJournalEntry(token, refNo) {
  var deny_ = checkPerm_(token, ['journal']); if (deny_) return deny_;
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    return withLock_(function () {
      var sheet = getSheetOrThrow_(SHEET_JOURNAL);
      var rows = sheetToObjects_(sheet).filter(function (r) { return r.ref_no === refNo; });
      if (rows.length === 0) return { success: false, message: 'No. Ref tidak ditemukan.' };
      // hapus dari baris terbawah ke atas supaya index tidak bergeser
      rows.sort(function (a, b) { return b._row - a._row; });
      rows.forEach(function (r) { sheet.deleteRow(r._row); });
      purgeEvidence_(refNo);
      Logger.log('deleteJournalEntry: ' + refNo + ' (' + rows.length + ' baris dihapus)');
      return { success: true, message: 'Jurnal ' + refNo + ' berhasil dihapus.' };
    });
  } catch (err) {
    Logger.log('deleteJournalEntry ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ===================================================================
// BUKU BESAR
// ===================================================================
function getLedger(token, accountCode, startDate, endDate) {
  var deny_ = checkPerm_(token, ['reports', 'journal']); if (deny_) return deny_;
  try {
    var coaRows = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
    var account = coaRows.filter(function (r) { return r.account_code === accountCode; })[0];
    if (!account) return { success: false, message: 'Akun tidak ditemukan.' };

    var journalRows = sheetToObjects_(getSheetOrThrow_(SHEET_JOURNAL))
      .filter(function (r) { return r.account_code === accountCode; })
      .map(function (r) { return { date: fmtDate_(r.date), ref_no: r.ref_no, description: r.description, debit: Number(r.debit) || 0, credit: Number(r.credit) || 0 }; })
      .sort(function (a, b) { return new Date(a.date) - new Date(b.date); });

    var isDebitNormal = account.normal_balance === 'Debit';
    var openingBalance = openingSeed_(account, isDebitNormal);
    var entries = [];

    journalRows.forEach(function (r) {
      var beforeStart = startDate && new Date(r.date) < new Date(startDate);
      var mutation = isDebitNormal ? (r.debit - r.credit) : (r.credit - r.debit);
      if (beforeStart) {
        openingBalance += mutation;
      }
    });

    var runningBalance = openingBalance;
    journalRows.forEach(function (r) {
      if (!inRange_(r.date, startDate, endDate)) return;
      var mutation = isDebitNormal ? (r.debit - r.credit) : (r.credit - r.debit);
      runningBalance += mutation;
      entries.push({ date: r.date, ref_no: r.ref_no, description: r.description, debit: r.debit, credit: r.credit, balance: runningBalance });
    });

    return {
      success: true,
      account: { account_code: account.account_code, account_name: account.account_name, normal_balance: account.normal_balance },
      openingBalance: openingBalance,
      entries: entries
    };
  } catch (err) {
    Logger.log('getLedger ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ===================================================================
// PENGATURAN PERUSAHAAN (mis. jumlah lembar saham beredar utk EPS)
// Disimpan di Document Properties (bukan sheet), karena ini nilai
// tunggal yang jarang berubah, bukan data transaksional.
// ===================================================================
function getCompanySettings(token) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  try {
    var props = PropertiesService.getDocumentProperties();
    var shares = Number(props.getProperty('shares_outstanding')) || 0;
    return { success: true, sharesOutstanding: shares };
  } catch (err) {
    Logger.log('getCompanySettings ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function saveCompanySettings(token, payload) {
  var deny_ = checkPerm_(token, ['coa']); if (deny_) return deny_;
  try {
    var shares = Number(payload.sharesOutstanding) || 0;
    if (shares < 0) return { success: false, message: 'Jumlah lembar saham tidak boleh negatif.' };
    PropertiesService.getDocumentProperties().setProperty('shares_outstanding', String(shares));
    Logger.log('saveCompanySettings: shares_outstanding = ' + shares);
    return { success: true, message: 'Pengaturan berhasil disimpan.', sharesOutstanding: shares };
  } catch (err) {
    Logger.log('saveCompanySettings ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ===================================================================
// PROFIL KOPERASI (kop cetakan). Boleh dibaca semua role, diubah hanya Administrator.
// ===================================================================
var DEFAULT_COOP_PROFILE = {
  name: 'Koperasi Produsen Mitra Maju Sejahtera',
  tagline: 'Bersama Petani, Menuju Sawit Sejahtera',
  address: '',
  contact: '',
  npwp: '',
  nib: '',
  owner: '',
  bank: '',
  units: ['Unit TBS', 'Unit Sarana Produksi'],
  ppn_rate: 12,
  ppn_dpp_lain: true,
  pph22_rate: 0.25,
  pph23_rate: 2,
  ppn_on: false,
  pph22_on: false,
  pph23_on: false
};
var LOGO_MIMES_ = { 'image/png': 'png', 'image/jpeg': 'jpg' };
var LOGO_MAX_BYTES_ = 1024 * 1024;

function getCoopProfile(token) {
  if (!requireAnySession_(token)) return sessionExpiredResponse_();
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty('coop_profile');
    var saved = raw ? JSON.parse(raw) : {};
    var p = {};
    Object.keys(DEFAULT_COOP_PROFILE).forEach(function (k) { p[k] = saved[k] !== undefined ? saved[k] : DEFAULT_COOP_PROFILE[k]; });
    // Data lama (sebelum aturan PMK 131/2024): tarif 11% tanpa DPP Nilai Lain -> tarif resmi 12% dengan DPP 11/12 (efektif tetap 11%).
    if (saved.ppn_dpp_lain === undefined && Number(saved.ppn_rate) === 11) p.ppn_rate = 12;
    p.logo = '';
    if (saved.logo_file_id) {
      try {
        var blob = DriveApp.getFileById(saved.logo_file_id).getBlob();
        p.logo = 'data:' + (saved.logo_mime || 'image/png') + ';base64,' + Utilities.base64Encode(blob.getBytes());
      } catch (e) { /* file hilang: tanpa logo */ }
    }
    return { success: true, profile: p };
  } catch (err) {
    Logger.log('getCoopProfile ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}
// Tarif pajak dalam persen, 0-100, maksimal 4 desimal.
// Daftar unit koperasi: array / teks per baris -> array rapi (unik, maks 20, tiap unit maks 40 karakter).
function cleanUnits_(v) {
  var arr = Array.isArray(v) ? v : String(v || '').split(/\r?\n|,/);
  var out = [];
  arr.forEach(function (x) {
    x = String(x || '').trim().slice(0, 40);
    if (x && out.map(function (y) { return y.toLowerCase(); }).indexOf(x.toLowerCase()) < 0 && out.length < 20) out.push(x);
  });
  return out;
}
function getUnitsList_() {
  var raw = PropertiesService.getDocumentProperties().getProperty('coop_profile');
  var saved = raw ? JSON.parse(raw) : {};
  return Array.isArray(saved.units) ? saved.units : DEFAULT_COOP_PROFILE.units;
}
// Kembalikan {unit} bila valid atau {error}. Unit kosong diizinkan.
function checkUnit_(u) {
  u = String(u || '').trim();
  if (!u) return { unit: '' };
  var list = getUnitsList_();
  var hit = list.filter(function (x) { return x.toLowerCase() === u.toLowerCase(); })[0];
  if (!hit) return { error: 'Unit "' + u + '" tidak ada di daftar unit. Atur di menu Profil Usaha.' };
  return { unit: hit };
}
function taxRate_(v, dflt) {
  var n = parseFloat(String(v === undefined || v === null || v === '' ? dflt : v).replace(',', '.'));
  if (isNaN(n) || n < 0) n = 0;
  if (n > 100) n = 100;
  return Math.round(n * 10000) / 10000;
}
function saveCoopProfile(token, payload) {
  if (!requireAdmin_(token)) return { success: false, message: 'Hanya Administrator yang dapat mengubah profil usaha.' };
  try {
    var name = String((payload && payload.name) || '').trim();
    if (!name) return { success: false, message: 'Nama usaha wajib diisi.' };
    return withLock_(function () {
      var props = PropertiesService.getDocumentProperties();
      var raw = props.getProperty('coop_profile');
      var old = raw ? JSON.parse(raw) : {};
      var p = {
        name: name.slice(0, 120),
        tagline: String(payload.tagline || '').trim().slice(0, 160),
        address: String(payload.address || '').trim().slice(0, 240),
        contact: String(payload.contact || '').trim().slice(0, 160),
        npwp: String(payload.npwp || '').trim().slice(0, 40),
        nib: String(payload.nib || '').trim().slice(0, 60),
        owner: String(payload.owner || '').trim().slice(0, 100),
        bank: String(payload.bank || '').trim().slice(0, 160),
        units: cleanUnits_(payload.units),
        ppn_rate: taxRate_(payload.ppn_rate, 12),
        ppn_dpp_lain: payload.ppn_dpp_lain === undefined ? true : !!payload.ppn_dpp_lain,
        pph22_rate: taxRate_(payload.pph22_rate, 0.25),
        pph23_rate: taxRate_(payload.pph23_rate, 2),
        ppn_on: !!payload.ppn_on, pph22_on: !!payload.pph22_on, pph23_on: !!payload.pph23_on,
        logo_file_id: old.logo_file_id || '',
        logo_mime: old.logo_mime || ''
      };
      var trashOld = function () { if (p.logo_file_id) { try { DriveApp.getFileById(p.logo_file_id).setTrashed(true); } catch (e) { /* sudah hilang */ } } };
      if (payload.logo_data) {
        var mime = String(payload.logo_mime || '');
        if (!LOGO_MIMES_[mime]) return { success: false, message: 'Logo harus berformat PNG atau JPG.' };
        var bytes = Utilities.base64Decode(String(payload.logo_data));
        if (bytes.length > LOGO_MAX_BYTES_) return { success: false, message: 'Ukuran logo maksimal 1 MB.' };
        var file = getEvidenceFolder_().createFile(Utilities.newBlob(bytes, mime, 'logo-usaha.' + LOGO_MIMES_[mime]));
        trashOld();
        p.logo_file_id = file.getId();
        p.logo_mime = mime;
      } else if (payload.remove_logo) {
        trashOld();
        p.logo_file_id = ''; p.logo_mime = '';
      }
      props.setProperty('coop_profile', JSON.stringify(p));
      return getCoopProfile(token).success ? { success: true, message: 'Profil usaha disimpan.', profile: getCoopProfile(token).profile } : { success: true, message: 'Profil usaha disimpan.' };
    });
  } catch (err) {
    Logger.log('saveCoopProfile ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ===================================================================
// RINGKASAN DASHBOARD
// ===================================================================
function getDashboardSummary(token) {
  var deny_ = checkPerm_(token, ['dashboard']); if (deny_) return deny_;
  try {
    var tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
    var coaRows = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
    var journalRows = sheetToObjects_(getSheetOrThrow_(SHEET_JOURNAL)).map(function (r) {
      return { date: fmtDate_(r.date), ref_no: r.ref_no, description: r.description, account_code: r.account_code, debit: Number(r.debit) || 0, credit: Number(r.credit) || 0 };
    });

    var coaByCode = {};
    coaRows.forEach(function (a) { coaByCode[a.account_code] = a; });

    function accBalance(codes, isDebitNormal) {
      var total = codes.reduce(function (s, code) {
        var acc = coaByCode[code];
        return s + (acc ? openingSeed_(acc, isDebitNormal) : 0);
      }, 0);
      journalRows.forEach(function (r) {
        if (codes.indexOf(r.account_code) === -1) return;
        total += isDebitNormal ? (r.debit - r.credit) : (r.credit - r.debit);
      });
      return total;
    }

    var cashCodes = coaRows.filter(function (a) {
      var n = String(a.account_name || '').toLowerCase();
      return a.category === 'Aset' && (n.indexOf('kas') !== -1 || n.indexOf('bank') !== -1);
    }).map(function (a) { return a.account_code; });

    var piutangCodes = coaRows.filter(function (a) {
      return a.category === 'Aset' && String(a.account_name || '').toLowerCase().indexOf('piutang') !== -1;
    }).map(function (a) { return a.account_code; });

    var hutangCodes = coaRows.filter(function (a) {
      return a.category === 'Kewajiban' && String(a.account_name || '').toLowerCase().indexOf('utang') !== -1;
    }).map(function (a) { return a.account_code; });

    var cashBalance = accBalance(cashCodes, true);
    var piutangBalance = accBalance(piutangCodes, true);
    var hutangBalance = accBalance(hutangCodes, false);

    function monthRange(offsetMonths) {
      var now = new Date();
      var d = new Date(now.getFullYear(), now.getMonth() + offsetMonths, 1);
      var start = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      var endD = new Date(d.getFullYear(), d.getMonth() + 1, 0);
      var end = Utilities.formatDate(endD, tz, 'yyyy-MM-dd');
      var label = Utilities.formatDate(d, tz, 'MMM');
      return { start: start, end: end, label: label };
    }

    function sumCategoryInRange(category, start, end) {
      var codes = coaRows.filter(function (a) { return a.category === category; }).map(function (a) { return a.account_code; });
      var total = 0;
      journalRows.forEach(function (r) {
        if (codes.indexOf(r.account_code) === -1) return;
        if (!inRange_(r.date, start, end)) return;
        total += category === 'Pendapatan' ? (r.credit - r.debit) : (r.debit - r.credit);
      });
      return total;
    }

    function pctChange(cur, prev) {
      if (!prev) return null;
      return ((cur - prev) / Math.abs(prev)) * 100;
    }

    var curMonth = monthRange(0);
    var prevMonth = monthRange(-1);
    var revenueThis = sumCategoryInRange('Pendapatan', curMonth.start, curMonth.end);
    var revenueLast = sumCategoryInRange('Pendapatan', prevMonth.start, prevMonth.end);
    var expenseThis = sumCategoryInRange('Beban', curMonth.start, curMonth.end);
    var expenseLast = sumCategoryInRange('Beban', prevMonth.start, prevMonth.end);
    var profitThis = revenueThis - expenseThis;
    var profitLast = revenueLast - expenseLast;

    var months = [];
    for (var i = 5; i >= 0; i--) months.push(monthRange(-i));
    var monthLabels = months.map(function (m) { return m.label; });
    var revenueSeries = months.map(function (m) { return sumCategoryInRange('Pendapatan', m.start, m.end); });
    var cashInSeries = [], cashOutSeries = [];
    months.forEach(function (m) {
      var cin = 0, cout = 0;
      journalRows.forEach(function (r) {
        if (cashCodes.indexOf(r.account_code) === -1) return;
        if (!inRange_(r.date, m.start, m.end)) return;
        cin += r.debit; cout += r.credit;
      });
      cashInSeries.push(cin);
      cashOutSeries.push(cout);
    });

    var grouped = {}, order = [];
    journalRows.forEach(function (r) {
      if (!grouped[r.ref_no]) { grouped[r.ref_no] = { ref_no: r.ref_no, date: r.date, description: r.description, debit: 0, credit: 0 }; order.push(r.ref_no); }
      grouped[r.ref_no].debit += r.debit;
      grouped[r.ref_no].credit += r.credit;
    });
    var recentJournal = order.map(function (ref) { return grouped[ref]; })
      .sort(function (a, b) { return new Date(b.date) - new Date(a.date); })
      .slice(0, 8);

    return {
      success: true,
      cashBalance: cashBalance,
      piutangBalance: piutangBalance,
      hutangBalance: hutangBalance,
      revenueThis: revenueThis, revenueChangePct: pctChange(revenueThis, revenueLast),
      expenseThis: expenseThis, expenseChangePct: pctChange(expenseThis, expenseLast),
      profitThis: profitThis, profitChangePct: pctChange(profitThis, profitLast),
      monthLabels: monthLabels,
      revenueSeries: revenueSeries,
      cashInSeries: cashInSeries,
      cashOutSeries: cashOutSeries,
      recentJournal: recentJournal,
      memberCount: coopSummary_().memberCount,
      totalSavings: coopSummary_().totalSavings
    };
  } catch (err) {
    Logger.log('getDashboardSummary ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ===================================================================
// LAPORAN KEUANGAN
// ===================================================================
function getFinancialStatements(token, startDate, endDate) {
  var deny_ = checkPerm_(token, ['reports']); if (deny_) return deny_;
  try {
    var allCoaRows = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
    var parentCodes = {};
    allCoaRows.forEach(function (r) {
      var code = String(r.account_code);
      if (code.indexOf('.') !== -1) parentCodes[code.substring(0, code.lastIndexOf('.'))] = true;
    });
    // Hanya akun leaf (bisa dipakai di Jurnal) yang dihitung sebagai baris laporan; akun header/rollup dilewati.
    var coaRows = allCoaRows.filter(function (r) { return !parentCodes[String(r.account_code)]; });

    var journalRows = sheetToObjects_(getSheetOrThrow_(SHEET_JOURNAL))
      .map(function (r) { return { account_code: r.account_code, date: fmtDate_(r.date), debit: Number(r.debit) || 0, credit: Number(r.credit) || 0 }; });

    function netBalanceFor(acc, isDebitNormal, filterRange) {
      var total = filterRange ? 0 : openingSeed_(acc, isDebitNormal);
      journalRows.forEach(function (r) {
        if (r.account_code !== acc.account_code) return;
        if (filterRange && !inRange_(r.date, startDate, endDate)) return;
        total += isDebitNormal ? (r.debit - r.credit) : (r.credit - r.debit);
      });
      return total;
    }

    var revenue = [], expense = [], assets = [], liabilities = [], equity = [];
    var totalRevenue = 0, totalExpense = 0, totalAssets = 0, totalLiabilities = 0, totalEquity = 0;

    coaRows.forEach(function (acc) {
      var isDebitNormal = acc.normal_balance === 'Debit';
      if (acc.category === 'Pendapatan') {
        var v = netBalanceFor(acc, isDebitNormal, true);
        revenue.push({ account_code: acc.account_code, account_name: acc.account_name, netBalance: v });
        totalRevenue += v;
      } else if (acc.category === 'Beban') {
        var v2 = netBalanceFor(acc, isDebitNormal, true);
        expense.push({ account_code: acc.account_code, account_name: acc.account_name, netBalance: v2 });
        totalExpense += v2;
      } else if (acc.category === 'Aset') {
        var v3 = netBalanceFor(acc, isDebitNormal, false); // neraca: kumulatif s.d. tanggal akhir
        assets.push({ account_code: acc.account_code, account_name: acc.account_name, netBalance: v3 });
        totalAssets += v3;
      } else if (acc.category === 'Kewajiban') {
        var v4 = netBalanceFor(acc, isDebitNormal, false);
        liabilities.push({ account_code: acc.account_code, account_name: acc.account_name, netBalance: v4 });
        totalLiabilities += v4;
      } else if (acc.category === 'Ekuitas') {
        var v5 = netBalanceFor(acc, isDebitNormal, false);
        equity.push({ account_code: acc.account_code, account_name: acc.account_name, netBalance: v5 });
        totalEquity += v5;
      }
    });

    var netProfit = totalRevenue - totalExpense;
    equity.push({ account_code: '-', account_name: 'SHU Tahun Berjalan', netBalance: netProfit });
    var totalEquityAndProfit = totalLiabilities + totalEquity + netProfit;

    return {
      success: true,
      incomeStatement: { revenue: revenue, expense: expense, totalRevenue: totalRevenue, totalExpense: totalExpense, netProfit: netProfit },
      balanceSheet: { assets: assets, liabilities: liabilities, equity: equity, totalAssets: totalAssets, totalEquityAndProfit: totalEquityAndProfit }
    };
  } catch (err) {
    Logger.log('getFinancialStatements ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ===================================================================
// LAPORAN PERBANDINGAN (bulanan / triwulan / semester / tahunan)
// statement: 'labarugi' | 'neraca' | 'ekuitas'
// ===================================================================
function comparePeriods_(periodType, year) {
  var y = Number(year) || new Date().getFullYear();
  var out = [];
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function lastDay(yy, m) { return new Date(yy, m, 0).getDate(); }
  function mk(label, yy, m1, m2) {
    return { label: label, start: yy + '-' + pad(m1) + '-01', end: yy + '-' + pad(m2) + '-' + pad(lastDay(yy, m2)) };
  }
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  var i;
  if (periodType === 'bulan') {
    for (i = 1; i <= 12; i++) out.push(mk(MON[i - 1] + ' ' + y, y, i, i));
  } else if (periodType === 'triwulan') {
    for (i = 1; i <= 4; i++) out.push(mk('TW ' + i + ' ' + y, y, (i - 1) * 3 + 1, i * 3));
  } else if (periodType === 'semester') {
    for (i = 1; i <= 2; i++) out.push(mk('Smt ' + i + ' ' + y, y, (i - 1) * 6 + 1, i * 6));
  } else {
    for (i = y - 3; i <= y; i++) out.push(mk(String(i), i, 1, 12));
  }
  return out;
}

function getComparativeReport(token, statement, periodType, year) {
  var deny_ = checkPerm_(token, ['reports']); if (deny_) return deny_;
  try {
    var periods = comparePeriods_(periodType, year);
    var allCoa = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
    var parents = {};
    allCoa.forEach(function (r) {
      var c = String(r.account_code);
      if (c.indexOf('.') !== -1) parents[c.substring(0, c.lastIndexOf('.'))] = true;
    });
    var coa = allCoa.filter(function (r) { return !parents[String(r.account_code)]; });
    var byAcc = {};
    sheetToObjects_(getSheetOrThrow_(SHEET_JOURNAL)).forEach(function (r) {
      var k = String(r.account_code);
      (byAcc[k] = byAcc[k] || []).push({ date: fmtDate_(r.date), debit: Number(r.debit) || 0, credit: Number(r.credit) || 0 });
    });
    // saldo bersih akun: seed (opsional) + mutasi dengan from <= tanggal <= to (string yyyy-MM-dd)
    function bal(acc, from, to, withSeed) {
      var deb = acc.normal_balance === 'Debit';
      var t = withSeed ? openingSeed_(acc, deb) : 0;
      (byAcc[String(acc.account_code)] || []).forEach(function (r) {
        if (from && r.date < from) return;
        if (to && r.date > to) return;
        t += deb ? (r.debit - r.credit) : (r.credit - r.debit);
      });
      return t;
    }
    function ofCat(cat) { return coa.filter(function (a) { return a.category === cat; }); }
    function rowsFor(cat, fn) {
      return ofCat(cat).map(function (a) {
        return { code: String(a.account_code), name: a.account_name, values: periods.map(function (p) { return fn(a, p); }) };
      });
    }
    function sum(rows) {
      return periods.map(function (p, i) { return rows.reduce(function (s, r) { return s + (r.values[i] || 0); }, 0); });
    }
    function add(a, b) { return a.map(function (v, i) { return v + b[i]; }); }
    function sub(a, b) { return a.map(function (v, i) { return v - b[i]; }); }
    var sections = [], grand = [];

    if (statement === 'labarugi') {
      var rev = rowsFor('Pendapatan', function (a, p) { return bal(a, p.start, p.end, false); });
      var exp = rowsFor('Beban', function (a, p) { return bal(a, p.start, p.end, false); });
      var tr = sum(rev), te = sum(exp);
      sections.push({ title: 'Pendapatan', rows: rev, totalLabel: 'Total Pendapatan', total: tr });
      sections.push({ title: 'Beban', rows: exp, totalLabel: 'Total Beban', total: te });
      grand.push({ label: 'SHU Bersih', values: sub(tr, te), kind: 'profit' });
    } else if (statement === 'neraca') {
      var asset = rowsFor('Aset', function (a, p) { return bal(a, null, p.end, true); });
      var liab = rowsFor('Kewajiban', function (a, p) { return bal(a, null, p.end, true); });
      var eq = rowsFor('Ekuitas', function (a, p) { return bal(a, null, p.end, true); });
      var cumProfit = periods.map(function (p) {
        var r = 0, e = 0;
        ofCat('Pendapatan').forEach(function (a) { r += bal(a, null, p.end, false); });
        ofCat('Beban').forEach(function (a) { e += bal(a, null, p.end, false); });
        return r - e;
      });
      eq.push({ code: '-', name: 'SHU Berjalan (kumulatif)', values: cumProfit });
      var ta = sum(asset), tl = sum(liab), tq = sum(eq);
      sections.push({ title: 'Aset', rows: asset, totalLabel: 'Total Aset', total: ta });
      sections.push({ title: 'Kewajiban', rows: liab, totalLabel: 'Total Kewajiban', total: tl });
      sections.push({ title: 'Ekuitas', rows: eq, totalLabel: 'Total Ekuitas', total: tq });
      grand.push({ label: 'Total Kewajiban + Ekuitas', values: add(tl, tq), kind: 'profit' });
      var shares = Number(PropertiesService.getDocumentProperties().getProperty('shares_outstanding')) || 0;
      if (shares > 0) grand.push({ label: 'Nilai Buku per Lembar Saham (' + shares + ' lembar)', values: tq.map(function (v) { return v / shares; }), kind: 'pershare' });
    } else {
      var op = rowsFor('Ekuitas', function (a, p) {
        var d = new Date(p.start); d.setDate(d.getDate() - 1);
        return bal(a, null, fmtDate_(d), true);
      });
      var mv = rowsFor('Ekuitas', function (a, p) { return bal(a, p.start, p.end, false); });
      var shu = periods.map(function (p) {
        var r = 0, e = 0;
        ofCat('Pendapatan').forEach(function (a) { r += bal(a, p.start, p.end, false); });
        ofCat('Beban').forEach(function (a) { e += bal(a, p.start, p.end, false); });
        return r - e;
      });
      var prior = periods.map(function (p) {
        var d0 = new Date(p.start); d0.setDate(d0.getDate() - 1);
        var lim = fmtDate_(d0), r0 = 0, e0 = 0;
        ofCat('Pendapatan').forEach(function (a) { r0 += bal(a, null, lim, false); });
        ofCat('Beban').forEach(function (a) { e0 += bal(a, null, lim, false); });
        return r0 - e0;
      });
      op.push({ code: '-', name: 'SHU Periode Sebelumnya (belum dibagi/ditutup)', values: prior });
      var top = sum(op), tmv = sum(mv);
      sections.push({ title: 'Saldo Awal', rows: op, totalLabel: 'Total Saldo Awal', total: top });
      var mvRows = mv.concat([{ code: '-', name: 'SHU Periode Berjalan', values: shu }]);
      var tmv2 = add(tmv, shu);
      sections.push({ title: 'Mutasi Periode', rows: mvRows, totalLabel: 'Total Mutasi', total: tmv2 });
      var tclose = add(top, tmv2);
      grand.push({ label: 'Total Saldo Akhir Ekuitas', values: tclose, kind: 'profit' });
      var shares2 = Number(PropertiesService.getDocumentProperties().getProperty('shares_outstanding')) || 0;
      if (shares2 > 0) grand.push({ label: 'Nilai Buku per Lembar Saham (' + shares2 + ' lembar)', values: tclose.map(function (v) { return v / shares2; }), kind: 'pershare' });
    }
    return { success: true, statement: statement, periodType: periodType, periods: periods, sections: sections, grand: grand };
  } catch (err) {
    Logger.log('getComparativeReport ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ===================================================================
// LAPORAN PERUBAHAN MODAL
// ===================================================================
function getEquityChanges(token, startDate, endDate) {
  var deny_ = checkPerm_(token, ['reports']); if (deny_) return deny_;
  try {
    var allCoaRows = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
    var parentCodes = {};
    allCoaRows.forEach(function (r) {
      var code = String(r.account_code);
      if (code.indexOf('.') !== -1) parentCodes[code.substring(0, code.lastIndexOf('.'))] = true;
    });
    var coaRows = allCoaRows.filter(function (r) { return !parentCodes[String(r.account_code)]; });

    var journalRows = sheetToObjects_(getSheetOrThrow_(SHEET_JOURNAL)).map(function (r) {
      return { account_code: r.account_code, date: fmtDate_(r.date), debit: Number(r.debit) || 0, credit: Number(r.credit) || 0 };
    });

    function balanceBefore(acc, dateLimit, isDebitNormal) {
      var total = openingSeed_(acc, isDebitNormal);
      journalRows.forEach(function (r) {
        if (r.account_code !== acc.account_code) return;
        if (dateLimit && new Date(r.date) >= new Date(dateLimit)) return;
        total += isDebitNormal ? (r.debit - r.credit) : (r.credit - r.debit);
      });
      return total;
    }

    var equityAccounts = coaRows.filter(function (a) { return a.category === 'Ekuitas'; });

    // "dayAfterEnd" dipakai supaya saldo akhir mencakup transaksi PADA tanggal akhir periode juga
    var dayAfterEnd = null;
    if (endDate) {
      var d = new Date(endDate);
      d.setDate(d.getDate() + 1);
      dayAfterEnd = fmtDate_(d);
    }

    var items = equityAccounts.map(function (acc) {
      var isDebitNormal = acc.normal_balance === 'Debit';
      var opening = balanceBefore(acc, startDate, isDebitNormal);
      var closing = balanceBefore(acc, dayAfterEnd, isDebitNormal);
      return {
        account_code: acc.account_code,
        account_name: acc.account_name,
        opening: opening,
        movement: closing - opening,
        closing: closing
      };
    });

    var finRes = getFinancialStatements(token, startDate, endDate);
    var netProfit = (finRes && finRes.success) ? finRes.incomeStatement.netProfit : 0;

    var totalOpening = items.reduce(function (s, i) { return s + i.opening; }, 0);
    var totalClosingExplicit = items.reduce(function (s, i) { return s + i.closing; }, 0);
    // Laba berjalan belum dijurnal ke akun modal manapun, jadi ditambahkan terpisah di sini
    var totalClosing = totalClosingExplicit + netProfit;

    return {
      success: true,
      items: items,
      netProfit: netProfit,
      totalOpening: totalOpening,
      totalClosing: totalClosing
    };
  } catch (err) {
    Logger.log('getEquityChanges ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ===================================================================
// LAPORAN ARUS KAS (metode langsung, estimasi otomatis)
// ===================================================================
function getCashFlow(token, startDate, endDate) {
  var deny_ = checkPerm_(token, ['reports']); if (deny_) return deny_;
  try {
    var coaRows = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
    var coaByCode = {};
    coaRows.forEach(function (a) { coaByCode[a.account_code] = a; });

    // Deteksi akun Kas/Bank dari nama
    var cashCodes = coaRows.filter(function (a) {
      var name = String(a.account_name || '').toLowerCase();
      return name.indexOf('kas') !== -1 || name.indexOf('bank') !== -1;
    }).map(function (a) { return a.account_code; });

    var journalRows = sheetToObjects_(getSheetOrThrow_(SHEET_JOURNAL)).map(function (r) {
      return { date: fmtDate_(r.date), ref_no: r.ref_no, account_code: r.account_code, debit: Number(r.debit) || 0, credit: Number(r.credit) || 0 };
    });

    var cashOpeningSeed = cashCodes.reduce(function (s, code) {
      var acc = coaByCode[code];
      return s + (acc ? openingSeed_(acc, true) : 0);
    }, 0);

    function cashBalanceBefore(dateLimit) {
      var total = cashOpeningSeed;
      journalRows.forEach(function (r) {
        if (cashCodes.indexOf(r.account_code) === -1) return;
        if (dateLimit && new Date(r.date) >= new Date(dateLimit)) return;
        total += r.debit - r.credit;
      });
      return total;
    }
    var openingCash = cashBalanceBefore(startDate);

    // Kelompokkan baris jurnal dalam rentang tanggal berdasarkan No. Ref
    var groups = {};
    journalRows.forEach(function (r) {
      if (!inRange_(r.date, startDate, endDate)) return;
      if (!groups[r.ref_no]) groups[r.ref_no] = [];
      groups[r.ref_no].push(r);
    });

    var operating = 0, investing = 0, financing = 0;

    Object.keys(groups).forEach(function (refNo) {
      var lines = groups[refNo];
      var cashLines = lines.filter(function (r) { return cashCodes.indexOf(r.account_code) !== -1; });
      if (cashLines.length === 0) return; // transaksi ini tidak menyentuh kas/bank
      var nonCashLines = lines.filter(function (r) { return cashCodes.indexOf(r.account_code) === -1; });

      var netCash = 0;
      cashLines.forEach(function (r) { netCash += r.debit - r.credit; });
      if (netCash === 0) return; // mis. transfer kas<->bank, tidak ada dampak arus kas neto

      // Klasifikasi berdasarkan kategori lawan akun (kalau seragam)
      var categories = {};
      nonCashLines.forEach(function (r) {
        var acc = coaByCode[r.account_code];
        categories[acc ? acc.category : 'Lainnya'] = true;
      });
      var catKeys = Object.keys(categories);
      var bucket = 'operating';
      if (catKeys.length === 1) {
        if (catKeys[0] === 'Aset') bucket = 'investing';
        else if (catKeys[0] === 'Kewajiban' || catKeys[0] === 'Ekuitas') bucket = 'financing';
      }
      if (bucket === 'operating') operating += netCash;
      else if (bucket === 'investing') investing += netCash;
      else financing += netCash;
    });

    var netChange = operating + investing + financing;
    var closingCash = openingCash + netChange;

    return {
      success: true,
      openingCash: openingCash,
      operating: operating,
      investing: investing,
      financing: financing,
      netChange: netChange,
      closingCash: closingCash
    };
  } catch (err) {
    Logger.log('getCashFlow ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ===================================================================
// ASET TETAP & PENYUSUTAN
// ===================================================================
function getAssets(token) {
  var deny_ = checkPerm_(token, ['assets']); if (deny_) return deny_;
  try {
    var sheet = getSheetOrThrow_(SHEET_ASSETS);
    var data = sheetToObjects_(sheet).map(function (r) {
      var cost = Number(r.cost) || 0;
      var accDep = Number(r.accumulated_dep) || 0;
      return {
        asset_id: r.asset_id, asset_name: r.asset_name, purchase_date: fmtDate_(r.purchase_date),
        cost: cost, salvage_value: Number(r.salvage_value) || 0, useful_life_years: Number(r.useful_life_years) || 0,
        accumulated_dep: accDep, net_book_value: cost - accDep
      };
    });
    Logger.log('getAssets: ' + data.length + ' baris ditemukan di sheet Assets');
    return { success: true, data: data };
  } catch (err) {
    Logger.log('getAssets ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function saveAsset(token, payload) {
  var deny_ = checkPerm_(token, ['assets']); if (deny_) return deny_;
  try {
    if (!payload || !payload.asset_name || !payload.cost) {
      return { success: false, message: 'Nama aset dan harga perolehan wajib diisi.' };
    }
    if (Number(payload.cost) <= 0) {
      return { success: false, message: 'Harga perolehan harus lebih besar dari 0.' };
    }

    return withLock_(function () {
      var sheet = getSheetOrThrow_(SHEET_ASSETS);
      var assetId = Utilities.getUuid();
      var newRow = [
        assetId, payload.asset_name, payload.purchase_date, Number(payload.cost),
        Number(payload.salvage_value) || 0, Number(payload.useful_life_years) || 1, 0, ''
      ];
      sheet.appendRow(newRow);
      SpreadsheetApp.flush(); // paksa tulis langsung, jangan tunggu batch

      var lastRow = sheet.getLastRow();
      Logger.log('saveAsset: OK, asset_id=' + assetId + ', ditulis di baris ' + lastRow + ' pada sheet "' + sheet.getName() + '" milik spreadsheet "' + getSS_().getName() + '"');

      return { success: true, message: 'Aset tetap berhasil ditambahkan.', asset_id: assetId };
    });
  } catch (err) {
    Logger.log('saveAsset ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function postDepreciationJournal(token, assetId) {
  var deny_ = checkPerm_(token, ['assets']); if (deny_) return deny_;
  var username = requireSession_(token);
  if (!username) return sessionExpiredResponse_();
  try {
    return withLock_(function () {
      var sheet = getSheetOrThrow_(SHEET_ASSETS);
      var rows = sheetToObjects_(sheet);
      var asset = rows.filter(function (r) { return r.asset_id === assetId; })[0];
      if (!asset) return { success: false, message: 'Aset tidak ditemukan.' };

      var cost = Number(asset.cost) || 0;
      var salvage = Number(asset.salvage_value) || 0;
      var lifeYears = Number(asset.useful_life_years) || 1;
      var accDep = Number(asset.accumulated_dep) || 0;
      var depreciableBase = cost - salvage;
      var monthlyDep = depreciableBase / (lifeYears * 12);

      if (accDep >= depreciableBase) {
        return { success: false, message: 'Aset ini sudah disusutkan penuh.' };
      }
      if (accDep + monthlyDep > depreciableBase) {
        monthlyDep = depreciableBase - accDep; // cap di sisa yang bisa disusutkan
      }

      var today = fmtDate_(new Date());
      var refNo = 'DEP-' + today + '-' + asset.asset_id.substring(0, 4);

      var coaRows = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
      var akunBeban = coaRows.filter(function (r) { return String(r.account_name || '').indexOf(NAMA_AKUN_BEBAN_PENYUSUTAN) !== -1; })[0];
      var akunAkum = coaRows.filter(function (r) { return String(r.account_name || '').indexOf(NAMA_AKUN_AKUM_PENYUSUTAN) !== -1; })[0];
      if (!akunBeban || !akunAkum) {
        return { success: false, message: 'Akun "Beban Penyusutan" / "Akumulasi Penyusutan" belum ada di COA. Jalankan setupDatabase() dari Apps Script Editor, atau pastikan nama akunnya mengandung kata tersebut.' };
      }

      var jSheet = getSheetOrThrow_(SHEET_JOURNAL);
      var now = new Date();
      var desc = 'Penyusutan bulanan: ' + asset.asset_name;
      jSheet.appendRow([today, refNo, akunBeban.account_code, akunBeban.account_name, desc, monthlyDep, 0, now]);
      jSheet.appendRow([today, refNo, akunAkum.account_code, akunAkum.account_name, desc, 0, monthlyDep, now]);

      sheet.getRange(asset._row, 7, 1, 2).setValues([[accDep + monthlyDep, today]]); // kolom 7=accumulated_dep, 8=last_dep_date
      SpreadsheetApp.flush();

      Logger.log('postDepreciationJournal: ' + refNo + ' untuk aset ' + asset.asset_name);
      return { success: true, message: 'Penyusutan sebesar ' + monthlyDep.toFixed(0) + ' berhasil diposting ke jurnal (' + refNo + ').' };
    });
  } catch (err) {
    Logger.log('postDepreciationJournal ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}


// ===================================================================
// KOPERASI DAGANG: ANGGOTA, SIMPANAN, VOLUME TRANSAKSI, SHU
// ===================================================================
// SHEET BARU (dibuat otomatis, aman untuk instalasi lama):
//  - Members     : member_id | member_no | name | join_date | phone | address | status | exit_date
//  - Savings     : tx_id | date | member_id | type | direction | amount | ref_no | note | timestamp
//                  type = Pokok | Wajib | Sukarela ; direction = Setor | Tarik
//                  Setiap baris Savings otomatis membuat 2 baris Journal dengan ref_no yang sama.
//  - MemberTrans : tx_id | date | member_id | kind | amount | note | timestamp
//                  Catatan volume transaksi anggota (Pembelian/Penjualan), HANYA untuk dasar hitung
//                  Jasa Transaksi SHU. Tidak membuat jurnal (jurnal penjualan tetap di Jurnal Umum).
// Pengaturan persentase SHU disimpan di Document Properties (kunci "shu_settings").

var SHEET_MEMBERS = 'Members';
var SHEET_SAVINGS = 'Savings';
var SHEET_MEMBER_TRANS = 'MemberTrans';

// Nama akun COA yang dipakai otomatis untuk jurnal simpanan (dicari berdasarkan NAMA, seperti akun penyusutan).
var SAVING_TYPE_ACCOUNT_NAME = { 'Pokok': 'Simpanan Pokok', 'Wajib': 'Simpanan Wajib', 'Sukarela': 'Simpanan Sukarela' };

var SHU_DEFAULTS = { cadangan: 40, jasa_simpanan: 25, jasa_transaksi: 25, pengurus: 5, pendidikan: 5 };
var SHU_LABELS = {
  cadangan: 'Dana Cadangan',
  jasa_simpanan: 'Jasa Simpanan Anggota',
  jasa_transaksi: 'Jasa Transaksi Anggota',
  pengurus: 'Dana Pengurus & Pengawas',
  pendidikan: 'Dana Pendidikan & Sosial'
};

function ensureCoopSheets_() {
  var ss = getSS_();
  var s;
  if (!ss.getSheetByName(SHEET_MEMBERS)) {
    s = getOrCreateSheet_(SHEET_MEMBERS, ['member_id', 'member_no', 'name', 'join_date', 'phone', 'address', 'status', 'exit_date']);
    s.getRange(2, 1, 998, 2).setNumberFormat('@'); // member_id & member_no selalu teks
    s.getRange(2, 5, 998, 1).setNumberFormat('@'); // phone teks supaya angka 0 di depan tidak hilang
  }
  if (!ss.getSheetByName(SHEET_SAVINGS)) {
    s = getOrCreateSheet_(SHEET_SAVINGS, ['tx_id', 'date', 'member_id', 'type', 'direction', 'amount', 'ref_no', 'note', 'timestamp']);
    s.getRange(2, 1, 998, 1).setNumberFormat('@');
    s.getRange(2, 3, 998, 1).setNumberFormat('@');
    s.getRange(2, 7, 998, 1).setNumberFormat('@');
  }
  if (!ss.getSheetByName(SHEET_MEMBER_TRANS)) {
    s = getOrCreateSheet_(SHEET_MEMBER_TRANS, ['tx_id', 'date', 'member_id', 'kind', 'amount', 'note', 'timestamp']);
    s.getRange(2, 1, 998, 1).setNumberFormat('@');
    s.getRange(2, 3, 998, 1).setNumberFormat('@');
  }
}

function getCoopSheet_(name) {
  ensureCoopSheets_();
  return getSheetOrThrow_(name);
}

/** Akun leaf saja (yang bisa dipakai di Jurnal): akun yang tidak punya sub-akun. */
function leafCoaRows_(rows) {
  var parents = {};
  rows.forEach(function (r) {
    var c = String(r.account_code);
    if (c.indexOf('.') !== -1) parents[c.substring(0, c.lastIndexOf('.'))] = true;
  });
  return rows.filter(function (r) { return !parents[String(r.account_code)]; });
}

/** Cari akun leaf berdasarkan nama: cocok persis dulu, kalau tidak ada baru yang mengandung kata tsb. */
function findLeafCoaByName_(rows, name) {
  var leaves = leafCoaRows_(rows);
  var n = String(name).toLowerCase();
  var exact = leaves.filter(function (r) { return String(r.account_name).toLowerCase() === n; })[0];
  if (exact) return exact;
  return leaves.filter(function (r) { return String(r.account_name).toLowerCase().indexOf(n) !== -1; })[0];
}

/** Tulis baris jurnal dengan kolom ref_no & account_code dipaksa teks (supaya "1.01" tidak jadi angka 1,01). */
// Pastikan sheet Journal punya kolom 'unit' (kolom ke-9); sheet lama dimigrasi otomatis.
function ensureJournalUnitCol_(sheet) {
  if (String(sheet.getRange(1, 9).getValue()) !== 'unit') sheet.getRange(1, 9).setValue('unit');
}
function appendJournalRows_(rows, unit) {
  var sheet = getSheetOrThrow_(SHEET_JOURNAL);
  ensureJournalUnitCol_(sheet);
  var u = String(unit || '');
  rows = rows.map(function (r) { return r.slice(0, 8).concat([u]); });
  var start = nextRow_(sheet, rows.length);
  sheet.getRange(start, 2, rows.length, 2).setNumberFormat('@');
  sheet.getRange(start, 1, rows.length, 9).setValues(rows);
}

/** Saldo simpanan per anggota s.d. endDate (yyyy-MM-dd, boleh null = semua). Hasil: { member_id: {Pokok, Wajib, Sukarela} } */
function computeSavingBalances_(endDate) {
  var bal = {};
  sheetToObjects_(getCoopSheet_(SHEET_SAVINGS)).forEach(function (r) {
    if (endDate && !inRange_(fmtDate_(r.date), null, endDate)) return;
    var id = String(r.member_id);
    if (!bal[id]) bal[id] = { Pokok: 0, Wajib: 0, Sukarela: 0 };
    if (bal[id][r.type] === undefined) return;
    var amt = Number(r.amount) || 0;
    bal[id][r.type] += (r.direction === 'Setor') ? amt : -amt;
  });
  return bal;
}

function nextMemberNo_(rows) {
  var max = 0;
  rows.forEach(function (r) {
    var n = parseInt(String(r.member_no).replace(/\D/g, ''), 10);
    if (!isNaN(n) && n > max) max = n;
  });
  var s = String(max + 1);
  while (s.length < 4) s = '0' + s;
  return 'A-' + s;
}

// ---------- ANGGOTA ----------
function getMembers(token) {
  var deny_ = checkPerm_(token, ['members', 'savings', 'shu']); if (deny_) return deny_;
  try {
    var bal = computeSavingBalances_(null);
    var accounts = memberAccountMap_();
    var data = sheetToObjects_(getCoopSheet_(SHEET_MEMBERS)).map(function (r) {
      var b = bal[String(r.member_id)] || { Pokok: 0, Wajib: 0, Sukarela: 0 };
      return {
        account_username: accounts[String(r.member_id)] || '',
        member_id: String(r.member_id), member_no: String(r.member_no), name: String(r.name || ''),
        join_date: fmtDate_(r.join_date), phone: String(r.phone || ''), address: String(r.address || ''),
        status: r.status || 'Aktif', exit_date: fmtDate_(r.exit_date),
        pokok: b.Pokok, wajib: b.Wajib, sukarela: b.Sukarela, total: b.Pokok + b.Wajib + b.Sukarela
      };
    });
    data.sort(function (a, b) { return a.member_no < b.member_no ? -1 : (a.member_no > b.member_no ? 1 : 0); });
    return { success: true, data: data };
  } catch (err) {
    Logger.log('getMembers ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** Tambah anggota baru (tanpa member_id) atau ubah anggota (dengan member_id). */
function saveMember(token, p) {
  if (p && p.member_id) { var adm_ = checkAdminAction_(token); if (adm_) return adm_; }
  var deny_ = checkPerm_(token, ['members']); if (deny_) return deny_;
  try {
    if (!p || !String(p.name || '').trim()) return { success: false, message: 'Nama anggota wajib diisi.' };
    return withLock_(function () {
      var sheet = getCoopSheet_(SHEET_MEMBERS);
      var rows = sheetToObjects_(sheet);
      var name = String(p.name).trim();
      var status = p.status === 'Keluar' ? 'Keluar' : 'Aktif';
      var joinDate = p.join_date || fmtDate_(new Date());

      if (p.member_id) {
        var ex = rows.filter(function (r) { return String(r.member_id) === String(p.member_id); })[0];
        if (!ex) return { success: false, message: 'Anggota tidak ditemukan.' };
        var exitDate = status === 'Keluar' ? (p.exit_date || fmtDate_(ex.exit_date) || fmtDate_(new Date())) : '';
        sheet.getRange(ex._row, 3, 1, 6).setValues([[name, joinDate, p.phone || '', p.address || '', status, exitDate]]);
        SpreadsheetApp.flush();
        return { success: true, message: 'Data anggota "' + name + '" berhasil diperbarui.' };
      }

      var no = String(p.member_no || '').trim();
      if (no) {
        var dup = rows.some(function (r) { return String(r.member_no).toLowerCase() === no.toLowerCase(); });
        if (dup) return { success: false, message: 'No. Anggota "' + no + '" sudah dipakai.' };
      } else {
        no = nextMemberNo_(rows);
      }
      var id = Utilities.getUuid();
      var newRow = nextRow_(sheet);
      sheet.getRange(newRow, 1, 1, 8).setValues([[id, no, name, joinDate, p.phone || '', p.address || '', status, status === 'Keluar' ? (p.exit_date || fmtDate_(new Date())) : '']]);
      SpreadsheetApp.flush();
      Logger.log('saveMember: ' + no + ' ' + name);
      return { success: true, message: 'Anggota "' + name + '" berhasil ditambahkan dengan No. Anggota ' + no + '.', member_id: id };
    });
  } catch (err) {
    Logger.log('saveMember ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deleteMember(token, memberId) {
  var deny_ = checkPerm_(token, ['members']); if (deny_) return deny_;
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    return withLock_(function () {
      var sheet = getCoopSheet_(SHEET_MEMBERS);
      var ex = sheetToObjects_(sheet).filter(function (r) { return String(r.member_id) === String(memberId); })[0];
      if (!ex) return { success: false, message: 'Anggota tidak ditemukan.' };
      var hasSaving = sheetToObjects_(getCoopSheet_(SHEET_SAVINGS)).some(function (r) { return String(r.member_id) === String(memberId); });
      var hasTrans = sheetToObjects_(getCoopSheet_(SHEET_MEMBER_TRANS)).some(function (r) { return String(r.member_id) === String(memberId); });
      if (hasSaving || hasTrans) {
        return { success: false, message: 'Anggota ini sudah punya transaksi simpanan/volume transaksi, tidak bisa dihapus. Ubah statusnya menjadi "Keluar" saja.' };
      }
      sheet.deleteRow(ex._row);
      deleteUsersByMember_(memberId);
      return { success: true, message: 'Anggota berhasil dihapus.' };
    });
  } catch (err) {
    Logger.log('deleteMember ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ---------- SIMPANAN ----------
function getSavings(token, memberId, startDate, endDate) {
  var deny_ = checkPerm_(token, ['savings']); if (deny_) return deny_;
  try {
    var names = {};
    sheetToObjects_(getCoopSheet_(SHEET_MEMBERS)).forEach(function (m) {
      names[String(m.member_id)] = { no: String(m.member_no), name: String(m.name || '') };
    });
    var data = sheetToObjects_(getCoopSheet_(SHEET_SAVINGS))
      .filter(function (r) {
        if (memberId && String(r.member_id) !== String(memberId)) return false;
        return inRange_(fmtDate_(r.date), startDate, endDate);
      })
      .map(function (r) {
        var m = names[String(r.member_id)] || { no: '-', name: '(anggota dihapus)' };
        return {
          tx_id: String(r.tx_id), date: fmtDate_(r.date), member_id: String(r.member_id),
          member_no: m.no, member_name: m.name, type: r.type, direction: r.direction,
          amount: Number(r.amount) || 0, ref_no: String(r.ref_no), note: String(r.note || '')
        };
      })
      .sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
    return { success: true, data: data };
  } catch (err) {
    Logger.log('getSavings ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/**
 * Catat setoran/penarikan simpanan dan otomatis buat jurnalnya.
 * p = { date, member_id, type: Pokok|Wajib|Sukarela, direction: Setor|Tarik, amount, cash_account_code, note }
 *  Setor : Debit Kas/Bank        | Kredit Simpanan X
 *  Tarik : Debit Simpanan X      | Kredit Kas/Bank
 */
function saveSaving(token, p) {
  var deny_ = checkPerm_(token, ['savings']); if (deny_) return deny_;
  try {
    var amount = Number(p && p.amount) || 0;
    if (!p || !p.date || !p.member_id || !SAVING_TYPE_ACCOUNT_NAME[p.type] ||
        (p.direction !== 'Setor' && p.direction !== 'Tarik') || !p.cash_account_code) {
      return { success: false, message: 'Data simpanan belum lengkap (tanggal, anggota, jenis, setor/tarik, akun kas/bank).' };
    }
    var pcs_ = periodClosedCheck_(p.date); if (pcs_) return pcs_;
    if (amount <= 0) return { success: false, message: 'Jumlah harus lebih besar dari 0.' };

    return withLock_(function () {
      var m = sheetToObjects_(getCoopSheet_(SHEET_MEMBERS)).filter(function (r) { return String(r.member_id) === String(p.member_id); })[0];
      if (!m) return { success: false, message: 'Anggota tidak ditemukan.' };
      var status = m.status || 'Aktif';

      if (p.direction === 'Setor' && status !== 'Aktif') {
        return { success: false, message: 'Anggota berstatus Keluar, tidak bisa menyetor simpanan.' };
      }
      if (p.direction === 'Tarik') {
        if ((p.type === 'Pokok' || p.type === 'Wajib') && status !== 'Keluar') {
          return { success: false, message: 'Simpanan Pokok/Wajib hanya boleh ditarik saat anggota berstatus Keluar. Ubah status anggota dulu.' };
        }
        var current = (computeSavingBalances_(null)[String(p.member_id)] || {})[p.type] || 0;
        if (amount > current) {
          return { success: false, message: 'Saldo Simpanan ' + p.type + ' (' + current + ') tidak cukup untuk penarikan ' + amount + '.' };
        }
      }

      var coaRows = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
      var cash = leafCoaRows_(coaRows).filter(function (r) { return String(r.account_code) === String(p.cash_account_code) && r.category === 'Aset'; })[0];
      if (!cash) return { success: false, message: 'Akun Kas/Bank tidak valid.' };
      var savingAcc = findLeafCoaByName_(coaRows, SAVING_TYPE_ACCOUNT_NAME[p.type]);
      if (!savingAcc) {
        return { success: false, message: 'Akun "' + SAVING_TYPE_ACCOUNT_NAME[p.type] + '" belum ada di COA. Buka halaman Kode Akun lalu klik "Tambah Akun Koperasi".' };
      }

      var txId = Utilities.getUuid();
      var refNo = 'SMP-' + String(p.date).replace(/-/g, '') + '-' + txId.substring(0, 4).toUpperCase();
      var now = new Date();

      var sSheet = getCoopSheet_(SHEET_SAVINGS);
      var sRow = nextRow_(sSheet);
      sSheet.getRange(sRow, 1, 1, 9).setValues([[txId, p.date, String(p.member_id), p.type, p.direction, amount, refNo, p.note || '', now]]);

      var desc = 'Simpanan ' + p.type + ' (' + p.direction + ') - ' + m.name + (p.note ? ' - ' + p.note : '');
      var debitAcc = p.direction === 'Setor' ? cash : savingAcc;
      var creditAcc = p.direction === 'Setor' ? savingAcc : cash;
      appendJournalRows_([
        [p.date, refNo, debitAcc.account_code, debitAcc.account_name, desc, amount, 0, now],
        [p.date, refNo, creditAcc.account_code, creditAcc.account_name, desc, 0, amount, now]
      ]);
      SpreadsheetApp.flush();
      Logger.log('saveSaving: ' + refNo + ' ' + p.type + ' ' + p.direction + ' ' + amount);
      return { success: true, message: 'Simpanan tercatat dan dijurnalkan otomatis (' + refNo + ').', ref_no: refNo };
    });
  } catch (err) {
    Logger.log('saveSaving ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** Hapus transaksi simpanan beserta jurnal otomatisnya (ref_no sama). */
function deleteSaving(token, txId) {
  var deny_ = checkPerm_(token, ['savings']); if (deny_) return deny_;
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    return withLock_(function () {
      var sSheet = getCoopSheet_(SHEET_SAVINGS);
      var ex = sheetToObjects_(sSheet).filter(function (r) { return String(r.tx_id) === String(txId); })[0];
      if (!ex) return { success: false, message: 'Transaksi simpanan tidak ditemukan.' };
      var pcd_ = periodClosedCheck_(ex.date); if (pcd_) return pcd_;

      // Jangan biarkan saldo simpanan menjadi negatif akibat menghapus setoran.
      if (ex.direction === 'Setor') {
        var bal = (computeSavingBalances_(null)[String(ex.member_id)] || {})[ex.type] || 0;
        if (bal - (Number(ex.amount) || 0) < 0) {
          return { success: false, message: 'Tidak bisa dihapus: saldo Simpanan ' + ex.type + ' anggota akan menjadi negatif (sudah ada penarikan setelahnya). Hapus penarikannya dulu.' };
        }
      }

      var refNo = String(ex.ref_no);
      var jSheet = getSheetOrThrow_(SHEET_JOURNAL);
      var jRows = sheetToObjects_(jSheet).filter(function (r) { return String(r.ref_no) === refNo; });
      jRows.sort(function (a, b) { return b._row - a._row; });
      jRows.forEach(function (r) { jSheet.deleteRow(r._row); });
      sSheet.deleteRow(ex._row);
      purgeEvidence_(refNo);
      SpreadsheetApp.flush();
      Logger.log('deleteSaving: ' + refNo + ' (' + jRows.length + ' baris jurnal dihapus)');
      return { success: true, message: 'Transaksi simpanan ' + refNo + ' dan jurnalnya berhasil dihapus.' };
    });
  } catch (err) {
    Logger.log('deleteSaving ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ---------- VOLUME TRANSAKSI ANGGOTA (dasar Jasa Transaksi SHU) ----------
function getMemberTrans(token, year) {
  var deny_ = checkPerm_(token, ['savings']); if (deny_) return deny_;
  try {
    var start = year ? year + '-01-01' : null;
    var end = year ? year + '-12-31' : null;
    var names = {};
    sheetToObjects_(getCoopSheet_(SHEET_MEMBERS)).forEach(function (m) {
      names[String(m.member_id)] = { no: String(m.member_no), name: String(m.name || '') };
    });
    var data = sheetToObjects_(getCoopSheet_(SHEET_MEMBER_TRANS))
      .filter(function (r) { return inRange_(fmtDate_(r.date), start, end); })
      .map(function (r) {
        var m = names[String(r.member_id)] || { no: '-', name: '(anggota dihapus)' };
        return {
          tx_id: String(r.tx_id), date: fmtDate_(r.date), member_id: String(r.member_id),
          member_no: m.no, member_name: m.name, kind: r.kind, amount: Number(r.amount) || 0, note: String(r.note || '')
        };
      })
      .sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
    return { success: true, data: data };
  } catch (err) {
    Logger.log('getMemberTrans ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function saveMemberTrans(token, p) {
  var deny_ = checkPerm_(token, ['savings']); if (deny_) return deny_;
  try {
    var amount = Number(p && p.amount) || 0;
    if (!p || !p.date || !p.member_id || (p.kind !== 'Pembelian' && p.kind !== 'Penjualan')) {
      return { success: false, message: 'Tanggal, anggota, dan jenis transaksi wajib diisi.' };
    }
    if (amount <= 0) return { success: false, message: 'Nilai transaksi harus lebih besar dari 0.' };
    return withLock_(function () {
      var m = sheetToObjects_(getCoopSheet_(SHEET_MEMBERS)).filter(function (r) { return String(r.member_id) === String(p.member_id); })[0];
      if (!m) return { success: false, message: 'Anggota tidak ditemukan.' };
      var sheet = getCoopSheet_(SHEET_MEMBER_TRANS);
      var row = nextRow_(sheet);
      sheet.getRange(row, 1, 1, 7).setValues([[Utilities.getUuid(), p.date, String(p.member_id), p.kind, amount, p.note || '', new Date()]]);
      SpreadsheetApp.flush();
      return { success: true, message: 'Volume transaksi anggota tercatat.' };
    });
  } catch (err) {
    Logger.log('saveMemberTrans ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deleteMemberTrans(token, txId) {
  var deny_ = checkPerm_(token, ['savings']); if (deny_) return deny_;
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    return withLock_(function () {
      var sheet = getCoopSheet_(SHEET_MEMBER_TRANS);
      var ex = sheetToObjects_(sheet).filter(function (r) { return String(r.tx_id) === String(txId); })[0];
      if (!ex) return { success: false, message: 'Data tidak ditemukan.' };
      sheet.deleteRow(ex._row);
      return { success: true, message: 'Data volume transaksi dihapus.' };
    });
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ---------- SHU ----------
function getShuSettings(token) {
  var deny_ = checkPerm_(token, ['shu']); if (deny_) return deny_;
  try {
    return { success: true, settings: readShuSettings_(), labels: SHU_LABELS };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function readShuSettings_() {
  var out = {};
  var raw = PropertiesService.getDocumentProperties().getProperty('shu_settings');
  var saved = {};
  try { saved = raw ? JSON.parse(raw) : {}; } catch (e) { saved = {}; }
  Object.keys(SHU_DEFAULTS).forEach(function (k) {
    out[k] = saved[k] !== undefined ? Number(saved[k]) || 0 : SHU_DEFAULTS[k];
  });
  return out;
}

function saveShuSettings(token, settings) {
  var deny_ = checkPerm_(token, ['shu']); if (deny_) return deny_;
  try {
    var clean = {}, total = 0;
    Object.keys(SHU_DEFAULTS).forEach(function (k) {
      var v = Number(settings && settings[k]);
      if (isNaN(v) || v < 0) v = 0;
      clean[k] = v;
      total += v;
    });
    if (Math.abs(total - 100) > 0.001) {
      return { success: false, message: 'Total persentase harus tepat 100% (sekarang ' + total + '%).' };
    }
    PropertiesService.getDocumentProperties().setProperty('shu_settings', JSON.stringify(clean));
    return { success: true, message: 'Pengaturan pembagian SHU berhasil disimpan.', settings: clean };
  } catch (err) {
    Logger.log('saveShuSettings ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/**
 * Hitung SHU tahun buku "year".
 *  - SHU = Pendapatan - Beban selama tahun tsb (dari Jurnal). Kalau rugi, SHU dibagikan = 0.
 *  - Jasa Simpanan  : dibagi sesuai proporsi total simpanan (Pokok+Wajib+Sukarela) tiap anggota per 31 Des.
 *  - Jasa Transaksi : dibagi sesuai proporsi volume transaksi anggota selama tahun tsb (sheet MemberTrans).
 *  - Anggota yang sudah keluar sebelum 31 Des tidak ikut dibagi.
 */
function calculateShu(token, year) {
  var deny_ = checkPerm_(token, ['shu']); if (deny_) return deny_;
  try {
    year = parseInt(year, 10);
    if (!year || year < 2000 || year > 2100) return { success: false, message: 'Tahun buku tidak valid.' };
    var start = year + '-01-01', end = year + '-12-31';

    var fin = getFinancialStatements(token, start, end);
    if (!fin || !fin.success) return { success: false, message: 'Gagal menghitung laba tahun buku: ' + (fin && fin.message) };
    var netProfit = fin.incomeStatement.netProfit;
    var distributable = netProfit > 0 ? netProfit : 0;

    var settings = readShuSettings_();
    var components = Object.keys(SHU_DEFAULTS).map(function (k) {
      return { key: k, label: SHU_LABELS[k], percent: settings[k], amount: distributable * settings[k] / 100 };
    });
    var amt = {};
    components.forEach(function (c) { amt[c.key] = c.amount; });

    var bal = computeSavingBalances_(end);
    var volume = {};
    sheetToObjects_(getCoopSheet_(SHEET_MEMBER_TRANS)).forEach(function (r) {
      if (!inRange_(fmtDate_(r.date), start, end)) return;
      var id = String(r.member_id);
      volume[id] = (volume[id] || 0) + (Number(r.amount) || 0);
    });

    var members = sheetToObjects_(getCoopSheet_(SHEET_MEMBERS)).filter(function (m) {
      var joined = fmtDate_(m.join_date);
      if (joined && new Date(joined) > new Date(end)) return false;
      if ((m.status || 'Aktif') === 'Keluar') {
        var exit = fmtDate_(m.exit_date);
        if (!exit || new Date(exit) <= new Date(end)) return false;
      }
      return true;
    }).map(function (m) {
      var b = bal[String(m.member_id)] || { Pokok: 0, Wajib: 0, Sukarela: 0 };
      return {
        member_id: String(m.member_id), member_no: String(m.member_no), name: String(m.name || ''),
        simpanan: Math.max(b.Pokok + b.Wajib + b.Sukarela, 0), volume: volume[String(m.member_id)] || 0
      };
    });
    members.sort(function (a, b) { return a.member_no < b.member_no ? -1 : (a.member_no > b.member_no ? 1 : 0); });

    var totalSimpanan = 0, totalVolume = 0;
    members.forEach(function (m) { totalSimpanan += m.simpanan; totalVolume += m.volume; });

    var totals = { simpanan: totalSimpanan, volume: totalVolume, jasa_simpanan: 0, jasa_transaksi: 0, total: 0 };
    members.forEach(function (m) {
      m.jasa_simpanan = totalSimpanan > 0 ? amt.jasa_simpanan * m.simpanan / totalSimpanan : 0;
      m.jasa_transaksi = totalVolume > 0 ? amt.jasa_transaksi * m.volume / totalVolume : 0;
      m.total = m.jasa_simpanan + m.jasa_transaksi;
      totals.jasa_simpanan += m.jasa_simpanan;
      totals.jasa_transaksi += m.jasa_transaksi;
      totals.total += m.total;
    });

    var warnings = [];
    if (netProfit <= 0) warnings.push('Hasil usaha tahun ' + year + ' nol atau rugi, sehingga tidak ada SHU yang dibagikan.');
    if (distributable > 0 && totalSimpanan <= 0) warnings.push('Belum ada saldo simpanan anggota, Jasa Simpanan tidak bisa dibagi.');
    if (distributable > 0 && totalVolume <= 0) warnings.push('Belum ada volume transaksi anggota tahun ' + year + ', Jasa Transaksi tidak bisa dibagi. Isi di bagian "Volume Transaksi Anggota".');
    if (!members.length) warnings.push('Belum ada anggota aktif pada tahun buku ini.');

    return {
      success: true, year: year, netProfit: netProfit, distributable: distributable,
      components: components, members: members, totals: totals, warnings: warnings
    };
  } catch (err) {
    Logger.log('calculateShu ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ---------- COA KOPERASI ----------
/** Tambahkan akun-akun standar koperasi yang belum ada di COA (aman dijalankan berulang, tidak menimpa apa pun). */
function setupCoopCoa(token) {
  var deny_ = checkPerm_(token, ['coa']); if (deny_) return deny_;
  try {
    return withLock_(function () {
      var sheet = getSheetOrThrow_(SHEET_COA);
      ensureCoaColumns_(sheet);
      var rows = sheetToObjects_(sheet);
      var codes = rows.map(function (r) { return String(r.account_code); });
      var names = rows.map(function (r) { return String(r.account_name).toLowerCase(); });

      var roots = [['1', 'ASET', 'Aset'], ['2', 'KEWAJIBAN', 'Kewajiban'], ['3', 'EKUITAS', 'Ekuitas'], ['4', 'PENDAPATAN', 'Pendapatan'], ['5', 'BEBAN', 'Beban']];
      var wanted = [
        ['1', 'Piutang Anggota', 'Debit'],
        ['1', 'Persediaan Barang Dagang', 'Debit'],
        ['2', 'Simpanan Sukarela Anggota', 'Kredit'],
        ['3', 'Simpanan Pokok', 'Kredit'],
        ['3', 'Simpanan Wajib', 'Kredit'],
        ['3', 'Dana Cadangan', 'Kredit'],
        ['1', 'Piutang Usaha', 'Debit'],
        ['2', 'Utang Usaha', 'Kredit'],
        ['4', 'Penjualan Barang Dagang', 'Kredit'],
        ['1', 'Piutang Penjualan', 'Debit'],
        ['4', 'Penjualan TBS', 'Kredit'],
        ['5', 'Harga Pokok Penjualan', 'Debit'],
        ['2', 'Utang PPN', 'Kredit'],
        ['1', 'PPh Dibayar Dimuka', 'Debit']
      ];
      var added = [];

      function addRow(code, name, category, normal) {
        var r = nextRow_(sheet);
        sheet.getRange(r, 1).setNumberFormat('@');
        sheet.getRange(r, 1, 1, 7).setValues([[code, name, category, normal, 'Aktif', 0, 0]]);
        codes.push(code);
        names.push(name.toLowerCase());
      }

      roots.forEach(function (root) {
        if (codes.indexOf(root[0]) === -1) addRow(root[0], root[1], root[2], defaultNormalBalance_(root[2]));
      });
      wanted.forEach(function (w) {
        if (names.indexOf(w[1].toLowerCase()) !== -1) return;
        var code = nextCoaChildCode_(codes, w[0]);
        addRow(code, w[1], deriveCoaCategory_(code), w[2]);
        added.push(code + ' ' + w[1]);
      });
      SpreadsheetApp.flush();
      Logger.log('setupCoopCoa: ' + added.length + ' akun ditambahkan');
      return {
        success: true,
        message: added.length
          ? added.length + ' akun koperasi ditambahkan: ' + added.join(', ') + '.'
          : 'Semua akun standar koperasi sudah ada. Tidak ada yang ditambahkan.'
      };
    });
  } catch (err) {
    Logger.log('setupCoopCoa ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** Ringkasan untuk dashboard: jumlah anggota aktif & total simpanan seluruh anggota. */
function coopSummary_() {
  try {
    var bal = computeSavingBalances_(null);
    var count = 0, total = 0;
    sheetToObjects_(getCoopSheet_(SHEET_MEMBERS)).forEach(function (m) {
      if ((m.status || 'Aktif') === 'Aktif') count++;
      var b = bal[String(m.member_id)];
      if (b) total += b.Pokok + b.Wajib + b.Sukarela;
    });
    return { memberCount: count, totalSavings: total };
  } catch (e) {
    return { memberCount: 0, totalSavings: 0 };
  }
}

// ===================================================================
// AKUN ANGGOTA (role "Anggota") - akses simpanan milik sendiri saja
// ===================================================================
// Sheet Users kini punya kolom ke-4: member_id (hanya terisi untuk role "Anggota").
// Anggota:
//  - hanya bisa memanggil getMyPortal() & changeMyPassword();
//  - SEMUA fungsi staf lain menolak sesi Anggota lewat requireSession_();
//  - member_id yang dipakai selalu dari sesi server, tidak pernah dari input browser.
// Administrator membuat/reset/hapus akun anggota lewat saveMemberAccount() / deleteMemberAccount().

function ensureUsersColumns_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 4);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('member_id') === -1) sheet.getRange(1, 4).setValue('member_id');
  sheet.getRange(2, 4, 998, 1).setNumberFormat('@');
}

/** { member_id: username } untuk semua akun role Anggota. */
function memberAccountMap_() {
  var map = {};
  var sheet = getSS_().getSheetByName(SHEET_USERS);
  if (!sheet) return map;
  sheetToObjects_(sheet).forEach(function (u) {
    if (u.role === 'Anggota' && u.member_id) map[String(u.member_id)] = String(u.username);
  });
  return map;
}

function deleteUsersByMember_(memberId) {
  var sheet = getSS_().getSheetByName(SHEET_USERS);
  if (!sheet) return 0;
  var rows = sheetToObjects_(sheet).filter(function (u) { return u.role === 'Anggota' && String(u.member_id) === String(memberId); });
  rows.sort(function (a, b) { return b._row - a._row; });
  rows.forEach(function (u) { sheet.deleteRow(u._row); });
  return rows.length;
}

function notAdminResponse_() {
  return { success: false, message: 'Fitur ini hanya untuk Administrator. Jika aplikasi baru diperbarui, login ulang dulu.' };
}

/**
 * Buat akun login anggota, atau reset password kalau anggota itu sudah punya akun. Hanya Administrator.
 * p = { member_id, username (hanya utk akun baru), password }
 */
function saveMemberAccount(token, p) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  if (!requireAdmin_(token)) return notAdminResponse_();
  try {
    if (!p || !p.member_id) return { success: false, message: 'Anggota belum dipilih.' };
    var password = String(p.password || '');
    if (password.length < 6) return { success: false, message: 'Password minimal 6 karakter.' };

    return withLock_(function () {
      var m = sheetToObjects_(getCoopSheet_(SHEET_MEMBERS)).filter(function (r) { return String(r.member_id) === String(p.member_id); })[0];
      if (!m) return { success: false, message: 'Anggota tidak ditemukan.' };

      var uSheet = getSheetOrThrow_(SHEET_USERS);
      ensureUsersColumns_(uSheet);
      var users = sheetToObjects_(uSheet);

      var existing = users.filter(function (u) { return u.role === 'Anggota' && String(u.member_id) === String(p.member_id); })[0];
      if (existing) {
        uSheet.getRange(existing._row, 2).setNumberFormat('@');
        uSheet.getRange(existing._row, 2).setValue(hashPassword_(password));
        SpreadsheetApp.flush();
        return { success: true, message: 'Password akun "' + existing.username + '" berhasil direset.' };
      }

      var username = String(p.username || '').trim();
      if (!/^[A-Za-z0-9._-]{3,30}$/.test(username)) {
        return { success: false, message: 'Username 3-30 karakter: huruf, angka, titik, strip, atau garis bawah.' };
      }
      var dup = users.some(function (u) { return String(u.username).toLowerCase() === username.toLowerCase(); });
      if (dup) return { success: false, message: 'Username "' + username + '" sudah dipakai.' };

      var row = nextRow_(uSheet);
      uSheet.getRange(row, 1, 1, 4).setNumberFormat('@');
      uSheet.getRange(row, 1, 1, 4).setValues([[username, hashPassword_(password), 'Anggota', String(p.member_id)]]);
      SpreadsheetApp.flush();
      Logger.log('saveMemberAccount: akun ' + username + ' untuk anggota ' + m.member_no);
      return { success: true, message: 'Akun "' + username + '" untuk ' + m.name + ' berhasil dibuat.' };
    });
  } catch (err) {
    Logger.log('saveMemberAccount ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deleteMemberAccount(token, memberId) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  if (!requireAdmin_(token)) return notAdminResponse_();
  try {
    return withLock_(function () {
      var n = deleteUsersByMember_(memberId);
      SpreadsheetApp.flush();
      return n ? { success: true, message: 'Akun login anggota dihapus.' } : { success: false, message: 'Anggota ini belum punya akun login.' };
    });
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/** Portal anggota: profil, saldo simpanan, dan riwayat transaksi MILIK SENDIRI saja. */
function getMyPortal(token, startDate, endDate) {
  var sess = requireMemberSession_(token);
  if (!sess) return sessionExpiredResponse_();
  try {
    var id = String(sess.member_id);
    var m = sheetToObjects_(getCoopSheet_(SHEET_MEMBERS)).filter(function (r) { return String(r.member_id) === id; })[0];
    if (!m) return { success: false, message: 'Data anggota tidak ditemukan. Hubungi pengurus koperasi.' };

    var b = computeSavingBalances_(null)[id] || { Pokok: 0, Wajib: 0, Sukarela: 0 };
    var data = sheetToObjects_(getCoopSheet_(SHEET_SAVINGS))
      .filter(function (r) { return String(r.member_id) === id && inRange_(fmtDate_(r.date), startDate, endDate); })
      .map(function (r) {
        return { date: fmtDate_(r.date), ref_no: String(r.ref_no), type: r.type, direction: r.direction, amount: Number(r.amount) || 0, note: String(r.note || '') };
      })
      .sort(function (a, c) { return new Date(c.date) - new Date(a.date); });

    return {
      success: true,
      profile: { member_no: String(m.member_no), name: String(m.name || ''), address: String(m.address || ''), join_date: fmtDate_(m.join_date), status: m.status || 'Aktif' },
      balances: { pokok: b.Pokok, wajib: b.Wajib, sukarela: b.Sukarela, total: b.Pokok + b.Wajib + b.Sukarela },
      data: data
    };
  } catch (err) {
    Logger.log('getMyPortal ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** Ubah password sendiri (staf maupun anggota). */
function changeMyPassword(token, oldPassword, newPassword) {
  var username = requireAnySession_(token);
  if (!username) return sessionExpiredResponse_();
  try {
    newPassword = String(newPassword || '');
    if (newPassword.length < 6) return { success: false, message: 'Password baru minimal 6 karakter.' };
    return withLock_(function () {
      var uSheet = getSheetOrThrow_(SHEET_USERS);
      var u = sheetToObjects_(uSheet).filter(function (r) { return String(r.username).toLowerCase() === String(username).toLowerCase(); })[0];
      if (!u) return { success: false, message: 'Akun tidak ditemukan.' };
      if (u.password_hash !== hashPassword_(String(oldPassword || ''))) return { success: false, message: 'Password lama salah.' };
      uSheet.getRange(u._row, 2).setNumberFormat('@');
      uSheet.getRange(u._row, 2).setValue(hashPassword_(newPassword));
      SpreadsheetApp.flush();
      return { success: true, message: 'Password berhasil diubah.' };
    });
  } catch (err) {
    Logger.log('changeMyPassword ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ===================================================================
// BARANG & STOK, PEMBELIAN, PENJUALAN, PELUNASAN
// ===================================================================
// SHEET BARU (dibuat otomatis):
//  - Items         : item_id | code | name | unit | sell_price | stock | avg_cost | status
//  - Purchases     : purchase_id | date | ref_no | supplier | total | paid | cash_account_code | note | timestamp
//  - PurchaseLines : line_id | purchase_id | item_id | item_name | qty | unit_cost | subtotal
//  - Sales         : sale_id | date | ref_no | customer | total | paid | cost_total | cash_account_code | note | timestamp
//  - SaleLines     : line_id | sale_id | item_id | item_name | qty | unit_price | unit_cost | subtotal
//  - TradePayments : pay_id | date | kind | ref_id | ref_no | amount | cash_account_code | journal_ref | timestamp
//
// STOK memakai harga pokok RATA-RATA BERGERAK (moving average):
//  - Pembelian menambah stok dan menghitung ulang avg_cost.
//  - Penjualan mengurangi stok pada avg_cost saat itu; nilainya dicatat sebagai HPP.
//
// JURNAL OTOMATIS (akun dicari berdasarkan NAMA, lihat tradeAccounts_):
//  Pembelian : Debit Persediaan | Kredit Kas/Bank (yang dibayar sekarang) + Kredit Utang Usaha (sisanya)
//  Penjualan : Debit Kas/Bank (yang diterima) + Debit Piutang Usaha (sisanya) | Kredit Penjualan Barang Dagang
//              Debit Harga Pokok Penjualan | Kredit Persediaan
//  Pelunasan : Piutang -> Debit Kas/Bank, Kredit Piutang Usaha ; Utang -> Debit Utang Usaha, Kredit Kas/Bank
// Jumlah "dibayar sekarang" lebih kecil dari total = sisanya menjadi piutang/utang (kredit / uang muka).

var SHEET_ITEMS = 'Items';
var SHEET_PURCHASES = 'Purchases';
var SHEET_PURCHASE_LINES = 'PurchaseLines';
var SHEET_SALES = 'Sales';
var SHEET_SALE_LINES = 'SaleLines';
var SHEET_TRADE_PAYMENTS = 'TradePayments';

// text = nomor kolom (1-based) yang dipaksa berformat teks (id, ref, kode akun)
var TRADE_SHEETS = {
  'Items': { headers: ['item_id', 'code', 'name', 'unit', 'sell_price', 'stock', 'avg_cost', 'status', 'item_type'], text: [1, 2] },
  'Purchases': { headers: ['purchase_id', 'date', 'ref_no', 'supplier', 'total', 'paid', 'cash_account_code', 'note', 'timestamp', 'vehicle_no', 'driver', 'kind', 'unit', 'supplier_id'], text: [1, 3, 7, 14] },
  'PurchaseLines': { headers: ['line_id', 'purchase_id', 'item_id', 'item_name', 'qty', 'unit_cost', 'subtotal', 'gross', 'tare', 'cut_pct', 'cut_kg'], text: [1, 2, 3] },
  'Sales': { headers: ['sale_id', 'date', 'ref_no', 'customer', 'total', 'paid', 'cost_total', 'cash_account_code', 'note', 'timestamp', 'vehicle_no', 'driver', 'receivable_account_code', 'kind', 'source_purchase_id', 'unit', 'tax_ppn', 'tax_pph22', 'tax_pph23', 'tax_ref'], text: [1, 3, 8, 13, 15, 20] },
  'SaleLines': { headers: ['line_id', 'sale_id', 'item_id', 'item_name', 'qty', 'unit_price', 'unit_cost', 'subtotal', 'gross', 'tare', 'cut_pct', 'cut_kg'], text: [1, 2, 3] },
  'Suppliers': { headers: ['supplier_id', 'code', 'name', 'group', 'phone', 'address', 'village', 'member_no', 'status', 'join_date', 'note', 'timestamp', 'unit'], text: [1, 2, 5, 8] },
  'Gardens': { headers: ['garden_id', 'supplier_id', 'name', 'location', 'area_ha', 'planting_year', 'trees', 'variety', 'land_status', 'lat', 'lng', 'status', 'note', 'timestamp', 'plantings'], text: [1, 2] },
  'GardenProgress': { headers: ['progress_id', 'garden_id', 'supplier_id', 'date', 'stage', 'progress_pct', 'activity', 'production_kg', 'note', 'by', 'timestamp'], text: [1, 2, 3] },
  'TradePayments': { headers: ['pay_id', 'date', 'kind', 'ref_id', 'ref_no', 'amount', 'cash_account_code', 'journal_ref', 'timestamp'], text: [1, 4, 5, 7, 8] }
};

var TRADE_MIGRATED_ = false;
function ensureTradeSheets_() {
  var ss = getSS_();
  Object.keys(TRADE_SHEETS).forEach(function (name) {
    var ex = ss.getSheetByName(name);
    if (ex) {
      // migrasi: tambah kolom baru (mis. kolom timbangan TBS) pada sheet lama
      if (!TRADE_MIGRATED_) {
        var want = TRADE_SHEETS[name].headers;
        var have = ex.getRange(1, 1, 1, Math.max(ex.getLastColumn(), 1)).getValues()[0].map(String);
        for (var c = have.length; c < want.length; c++) {
          if (have.indexOf(want[c]) === -1) ex.getRange(1, c + 1).setValue(want[c]);
        }
      }
      return;
    }
    var cfg = TRADE_SHEETS[name];
    var s = getOrCreateSheet_(name, cfg.headers);
    cfg.text.forEach(function (c) { s.getRange(2, c, 998, 1).setNumberFormat('@'); });
  });
  TRADE_MIGRATED_ = true;
}

// Hitung berat bayar dari timbangan TBS: (bruto - tara) x (1 - potongan%). Tanpa bruto, pakai qty apa adanya.
function weighLine_(L, k) {
  var gross = Number(L.gross) || 0, tare = Number(L.tare) || 0, cut = Number(L.cut_pct) || 0;
  if (gross > 0) {
    if (tare < 0 || tare >= gross) throw new Error('Tara pada baris ' + (k + 1) + ' harus lebih kecil dari berat bruto.');
    if (cut < 0 || cut >= 100) throw new Error('Potongan pada baris ' + (k + 1) + ' harus antara 0 dan 100%.');
    var cutKg = Number(L.cut_kg) || 0;
    if (cutKg < 0 || cutKg >= gross - tare) throw new Error('Potongan (kg) pada baris ' + (k + 1) + ' harus lebih kecil dari berat.');
    return { qty: q4_((gross - tare) * (1 - cut / 100) - cutKg), gross: gross, tare: tare, cut: cut, cutKg: cutKg };
  }
  return { qty: Number(L.qty) || 0, gross: '', tare: '', cut: '', cutKg: '' };
}

function getTradeSheet_(name) {
  ensureTradeSheets_();
  return getSheetOrThrow_(name);
}

function appendTradeRows_(name, rows) {
  if (!rows.length) return;
  var sheet = getTradeSheet_(name);
  var start = nextRow_(sheet, rows.length);
  TRADE_SHEETS[name].text.forEach(function (c) { sheet.getRange(start, c, rows.length, 1).setNumberFormat('@'); });
  sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
}

function deleteRowsDesc_(sheet, objs) {
  objs.slice().sort(function (a, b) { return b._row - a._row; }).forEach(function (o) { sheet.deleteRow(o._row); });
}

// Jenis transaksi/barang: 'TBS' (default, data lama) atau 'Barang' (barang dagang umum)
function kindOf_(v) { return String(v || '') === 'Barang' ? 'Barang' : 'TBS'; }
function q4_(n) { return Math.round(n * 10000) / 10000; }

var TRADE_ACCOUNT_NAMES = {
  persediaan: 'Persediaan Barang Dagang',
  hpp: 'Harga Pokok Penjualan',
  penjualan: 'Penjualan Barang Dagang',
  piutang: 'Piutang Usaha',
  utang: 'Utang Usaha',
  ppn: 'Utang PPN',
  pph: 'PPh Dibayar Dimuka'
};

/** Cari akun dagang di COA. Mengembalikan { acc: {kunci: baris COA}, missing(keys) -> pesan error atau null }. */
function tradeAccounts_(coaRows, kind) {
  var acc = {};
  Object.keys(TRADE_ACCOUNT_NAMES).forEach(function (k) { acc[k] = findLeafCoaByName_(coaRows, TRADE_ACCOUNT_NAMES[k]); });
  // Cadangan nama akun: COA lama memakai "Hpp TBS" / "Hpp Pupuk" dan "Persediaan ..."
  if (!acc.hpp) {
    var hppTry = kind === 'Barang' ? ['Hpp Pupuk', 'Hpp Barang', 'Hpp'] : ['Hpp TBS', 'Hpp'];
    for (var i = 0; i < hppTry.length && !acc.hpp; i++) acc.hpp = findLeafCoaByName_(coaRows, hppTry[i]);
  }
  if (!acc.ppn) {
    ['PPN Keluaran', 'Hutang PPN', 'PPN'].forEach(function (n) { if (!acc.ppn) acc.ppn = findLeafCoaByName_(coaRows, n); });
  }
  if (!acc.pph) {
    ['Pajak Dibayar Dimuka', 'Uang Muka Pajak', 'PPh 22', 'PPh'].forEach(function (n) { if (!acc.pph) acc.pph = findLeafCoaByName_(coaRows, n); });
  }
  if (!acc.persediaan) {
    var psTry = kind === 'Barang' ? ['Persediaan Pupuk', 'Persediaan Barang', 'Persediaan'] : ['Persediaan TBS', 'Persediaan'];
    for (var j = 0; j < psTry.length && !acc.persediaan; j++) acc.persediaan = findLeafCoaByName_(coaRows, psTry[j]);
  }
  return {
    acc: acc,
    missing: function (keys) {
      var miss = keys.filter(function (k) { return !acc[k]; }).map(function (k) { return '"' + TRADE_ACCOUNT_NAMES[k] + '"'; });
      return miss.length ? 'Akun ' + miss.join(', ') + ' belum ada di COA. Buka halaman Kode Akun lalu klik "Tambah Akun Koperasi".' : null;
    }
  };
}

function findCashAccount_(coaRows, code) {
  return leafCoaRows_(coaRows).filter(function (r) { return String(r.account_code) === String(code) && r.category === 'Aset'; })[0];
}

function tradeRef_(prefix, dateStr, id) {
  return prefix + '-' + String(dateStr).replace(/-/g, '') + '-' + id.substring(0, 4).toUpperCase();
}

function jrow_(date, ref, acc, desc, debit, credit, now) {
  return [date, ref, acc.account_code, acc.account_name, desc, debit || 0, credit || 0, now];
}

// ---------- BARANG ----------
function getItems(token) {
  var deny_ = checkPerm_(token, ['items', 'purchases_tbs', 'sales_tbs', 'purchases_barang', 'sales_barang']); if (deny_) return deny_;
  var itemsAll_ = sessionPerms_(token).indexOf('items') >= 0;
  var okKinds_ = itemsAll_ ? ['TBS', 'Barang'] : allowedKinds_(token, ['purchases', 'sales']);
  try {
    var totalValue = 0;
    var data = sheetToObjects_(getTradeSheet_(SHEET_ITEMS)).map(function (r) {
      var stock = Number(r.stock) || 0, avg = Number(r.avg_cost) || 0;
      totalValue += Math.round(stock * avg);
      return {
        item_id: String(r.item_id), code: String(r.code), name: String(r.name || ''), unit: String(r.unit || ''),
        sell_price: Number(r.sell_price) || 0, stock: stock, avg_cost: avg, value: Math.round(stock * avg),
        status: r.status || 'Aktif', item_type: kindOf_(r.item_type)
      };
    });
    data = data.filter(function (d) { return okKinds_.indexOf(d.item_type) >= 0; });
    totalValue = data.reduce(function (a, d) { return a + d.value; }, 0);
    data.sort(function (a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
    return { success: true, data: data, totalValue: totalValue };
  } catch (err) {
    Logger.log('getItems ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/**
 * Tambah / ubah barang. p = { item_id?, name, unit, sell_price, status, initial_stock?, initial_cost? }.
 * initial_stock/initial_cost hanya dipakai saat barang BARU (stok awal, TANPA jurnal: samakan dengan Saldo Awal
 * akun Persediaan di halaman Kode Akun). Setelah itu stok hanya berubah lewat Pembelian/Penjualan.
 */
function saveItem(token, p) {
  if (p && p.item_id) { var adm_ = checkAdminAction_(token); if (adm_) return adm_; }
  var deny_ = checkPerm_(token, ['items']); if (deny_) return deny_;
  try {
    if (!p || !String(p.name || '').trim()) return { success: false, message: 'Nama barang wajib diisi.' };
    var name = String(p.name).trim();
    var sell = Number(p.sell_price) || 0;
    if (sell < 0) return { success: false, message: 'Harga jual tidak boleh negatif.' };
    return withLock_(function () {
      var sheet = getTradeSheet_(SHEET_ITEMS);
      var rows = sheetToObjects_(sheet);
      var dup = rows.some(function (r) { return String(r.name).toLowerCase() === name.toLowerCase() && String(r.item_id) !== String(p.item_id || ''); });
      if (dup) return { success: false, message: 'Barang dengan nama "' + name + '" sudah ada.' };
      var status = p.status === 'Nonaktif' ? 'Nonaktif' : 'Aktif';

      if (p.item_id) {
        var ex = rows.filter(function (r) { return String(r.item_id) === String(p.item_id); })[0];
        if (!ex) return { success: false, message: 'Barang tidak ditemukan.' };
        sheet.getRange(ex._row, 3, 1, 3).setValues([[name, p.unit || '', sell]]);
        sheet.getRange(ex._row, 8).setValue(status);
        SpreadsheetApp.flush();
        return { success: true, message: 'Barang "' + name + '" diperbarui.' };
      }

      var stock = Number(p.initial_stock) || 0, cost = Number(p.initial_cost) || 0;
      if (stock < 0 || cost < 0) return { success: false, message: 'Stok awal dan harga pokok awal tidak boleh negatif.' };
      if (stock > 0 && cost <= 0) return { success: false, message: 'Isi harga pokok awal per satuan untuk stok awal.' };
      var max = 0;
      rows.forEach(function (r) { var n = parseInt(String(r.code).replace(/\D/g, ''), 10); if (!isNaN(n) && n > max) max = n; });
      var s = String(max + 1); while (s.length < 4) s = '0' + s;
      var code = 'BRG-' + s;
      appendTradeRows_(SHEET_ITEMS, [[Utilities.getUuid(), code, name, p.unit || '', sell, q4_(stock), q4_(cost), status, kindOf_(p.item_type)]]);
      SpreadsheetApp.flush();
      return { success: true, message: 'Barang "' + name + '" ditambahkan dengan kode ' + code + '.' };
    });
  } catch (err) {
    Logger.log('saveItem ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deleteItem(token, itemId) {
  var deny_ = checkPerm_(token, ['items']); if (deny_) return deny_;
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    return withLock_(function () {
      var sheet = getTradeSheet_(SHEET_ITEMS);
      var ex = sheetToObjects_(sheet).filter(function (r) { return String(r.item_id) === String(itemId); })[0];
      if (!ex) return { success: false, message: 'Barang tidak ditemukan.' };
      var used = sheetToObjects_(getTradeSheet_(SHEET_PURCHASE_LINES)).some(function (r) { return String(r.item_id) === String(itemId); }) ||
                 sheetToObjects_(getTradeSheet_(SHEET_SALE_LINES)).some(function (r) { return String(r.item_id) === String(itemId); });
      if (used) return { success: false, message: 'Barang ini sudah dipakai di transaksi, tidak bisa dihapus. Ubah statusnya menjadi Nonaktif.' };
      if ((Number(ex.stock) || 0) > 0) return { success: false, message: 'Barang ini masih punya stok, tidak bisa dihapus.' };
      sheet.deleteRow(ex._row);
      return { success: true, message: 'Barang dihapus.' };
    });
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ---------- PEMBELIAN ----------
// Total tagihan penjualan setelah pajak invoice: harga + PPN - PPh yang dipotong pembeli.
function saleNet_(r) {
  return (Number(r.total) || 0) + (Number(r.tax_ppn) || 0) - (Number(r.tax_pph22) || 0) - (Number(r.tax_pph23) || 0);
}
function tradeStatus_(total, paid) {
  var out = total - paid;
  return out <= 0 ? 'Lunas' : (paid > 0 ? 'Sebagian' : 'Belum Lunas');
}

function getPurchases(token, startDate, endDate, kind) {
  var deny_ = checkPerm_(token, ['purchases_tbs', 'purchases_barang', 'sales_tbs']); if (deny_) return deny_;
  var mine_ = sessionPerms_(token), okKinds_ = [];
  if (mine_.indexOf('purchases_tbs') >= 0 || mine_.indexOf('sales_tbs') >= 0) okKinds_.push('TBS');
  if (mine_.indexOf('purchases_barang') >= 0) okKinds_.push('Barang');
  if (kind && okKinds_.indexOf(kindOf_(kind)) < 0) return forbiddenResponse_();
  try {
    var lines = {};
    sheetToObjects_(getTradeSheet_(SHEET_PURCHASE_LINES)).forEach(function (l) {
      var k = String(l.purchase_id);
      (lines[k] = lines[k] || []).push({ item_name: String(l.item_name), qty: Number(l.qty) || 0, unit_cost: Number(l.unit_cost) || 0, subtotal: Number(l.subtotal) || 0, gross: Number(l.gross) || 0, tare: Number(l.tare) || 0, cut_pct: Number(l.cut_pct) || 0, cut_kg: Number(l.cut_kg) || 0 });
    });
    var data = sheetToObjects_(getTradeSheet_(SHEET_PURCHASES))
      .filter(function (r) { return inRange_(fmtDate_(r.date), startDate, endDate) && (!kind || kindOf_(r.kind) === kind) && okKinds_.indexOf(kindOf_(r.kind)) >= 0; })
      .map(function (r) {
        var total = Number(r.total) || 0, paid = Number(r.paid) || 0;
        return {
          purchase_id: String(r.purchase_id), date: fmtDate_(r.date), ref_no: String(r.ref_no), supplier: String(r.supplier || ''), supplier_id: String(r.supplier_id || ''),
          total: total, paid: paid, outstanding: total - paid, status: tradeStatus_(total, paid), note: String(r.note || ''),
          vehicle_no: String(r.vehicle_no || ''), driver: String(r.driver || ''), kind: kindOf_(r.kind), unit: String(r.unit || ''),
          lines: lines[String(r.purchase_id)] || []
        };
      })
      .sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
    data = filterByUnits_(token, data);
    return { success: true, data: data };
  } catch (err) {
    Logger.log('getPurchases ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/**
 * p = { date, supplier, note, paid, cash_account_code, lines: [{ item_id, qty, unit_cost }] }
 * paid = jumlah dibayar sekarang (0..total). Sisanya menjadi Utang Usaha.
 */
function savePurchase(token, p) {
  var deny_ = checkTradePerm_(token, 'purchases', p && p.kind); if (deny_) return deny_;
  try {
    if (!p || !p.date) return { success: false, message: 'Tanggal wajib diisi.' };
    var inLines = p.lines || [];
    if (!inLines.length) return { success: false, message: 'Tambahkan minimal satu barang.' };
    var kind = kindOf_(p.kind);
    var pc_ = periodClosedCheck_(p.date); if (pc_) return pc_;
    var uc = checkUnit_(p.unit);
    if (uc.error) return { success: false, message: uc.error };
    var ua_ = checkUnitAccess_(token, uc.unit); if (ua_) return ua_;

    return withLock_(function () {
      var itemSheet = getTradeSheet_(SHEET_ITEMS);
      var byId = {};
      sheetToObjects_(itemSheet).forEach(function (i) {
        byId[String(i.item_id)] = { row: i._row, name: String(i.name), type: kindOf_(i.item_type), stock: Number(i.stock) || 0, avg: Number(i.avg_cost) || 0, touched: false };
      });

      var purchaseId = Utilities.getUuid(), lineRows = [], total = 0;
      for (var k = 0; k < inLines.length; k++) {
        var L = inLines[k], it = byId[String(L.item_id)];
        var w = weighLine_(L, k), qty = w.qty, cost = Number(L.unit_cost) || 0;
        if (!it) return { success: false, message: 'Barang pada baris ' + (k + 1) + ' tidak ditemukan.' };
        if (it.type !== kind) return { success: false, message: 'Barang "' + it.name + '" berjenis ' + it.type + ', tidak bisa dipakai pada transaksi ' + kind + '.' };
        if (qty <= 0) return { success: false, message: 'Jumlah barang pada baris ' + (k + 1) + ' harus lebih dari 0.' };
        if (cost <= 0) return { success: false, message: 'Harga beli pada baris ' + (k + 1) + ' harus lebih dari 0.' };
        var sub = Math.round(qty * cost);
        var newStock = it.stock + qty;
        it.avg = it.stock > 0 ? (it.stock * it.avg + qty * cost) / newStock : cost;
        it.stock = newStock;
        it.touched = true;
        total += sub;
        lineRows.push([Utilities.getUuid(), purchaseId, String(L.item_id), it.name, qty, cost, sub, w.gross, w.tare, w.cut, w.cutKg]);
      }

      var paid = (p.paid === undefined || p.paid === null || p.paid === '') ? total : Number(p.paid);
      if (isNaN(paid) || paid < 0 || paid > total) return { success: false, message: 'Jumlah dibayar harus antara 0 dan total (' + total + ').' };

      var coaRows = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
      var t = tradeAccounts_(coaRows, kind);
      var need = ['persediaan'].concat(total - paid > 0 ? ['utang'] : []);
      var miss = t.missing(need);
      if (miss) return { success: false, message: miss };
      var cash = null;
      if (paid > 0) {
        cash = findCashAccount_(coaRows, p.cash_account_code);
        if (!cash) return { success: false, message: 'Pilih akun Kas/Bank untuk pembayaran.' };
      }

      var refNo = tradeRef_('PBL', p.date, purchaseId), now = new Date();
      var supplier = String(p.supplier || '').trim(), supplierId = '';
      if (p.supplier_id) {
        var sp_ = sheetToObjects_(getTradeSheet_(SHEET_SUPPLIERS)).filter(function (r) { return String(r.supplier_id) === String(p.supplier_id); })[0];
        if (!sp_) return { success: false, message: 'Pemasok yang dipilih tidak ditemukan.' };
        if (!inSupplierScope_(supplierScope_(token), sp_.unit)) return forbiddenResponse_();
        supplierId = String(sp_.supplier_id); supplier = String(sp_.name || '');
      }
      var desc = 'Pembelian ' + (kind === 'TBS' ? 'TBS' : 'barang') + (supplier ? ' dari ' + supplier : '');
      var j = [jrow_(p.date, refNo, t.acc.persediaan, desc, total, 0, now)];
      if (paid > 0) j.push(jrow_(p.date, refNo, cash, desc, 0, paid, now));
      if (total - paid > 0) j.push(jrow_(p.date, refNo, t.acc.utang, desc, 0, total - paid, now));

      appendTradeRows_(SHEET_PURCHASES, [[purchaseId, p.date, refNo, supplier, total, paid, cash ? String(cash.account_code) : '', p.note || '', now, String(p.vehicle_no || '').trim(), String(p.driver || '').trim(), kind, uc.unit, supplierId]]);
      appendTradeRows_(SHEET_PURCHASE_LINES, lineRows);
      appendJournalRows_(j, uc.unit);
      Object.keys(byId).forEach(function (id) {
        var it = byId[id];
        if (it.touched) itemSheet.getRange(it.row, 6, 1, 2).setValues([[q4_(it.stock), q4_(it.avg)]]);
      });
      SpreadsheetApp.flush();
      Logger.log('savePurchase: ' + refNo + ' total ' + total);
      return { success: true, message: 'Pembelian tercatat dan dijurnalkan (' + refNo + ').', ref_no: refNo };
    });
  } catch (err) {
    Logger.log('savePurchase ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deletePurchase(token, purchaseId) {
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  var deny0_ = checkPerm_(token, ['purchases_tbs', 'purchases_barang']); if (deny0_) return deny0_;
  var deny_ = checkTradePerm_(token, 'purchases', lookupTradeKind_(SHEET_PURCHASES, 'purchase_id', purchaseId)); if (deny_) return deny_;
  try {
    return withLock_(function () {
      var pSheet = getTradeSheet_(SHEET_PURCHASES);
      var head = sheetToObjects_(pSheet).filter(function (r) { return String(r.purchase_id) === String(purchaseId); })[0];
      if (!head) return { success: false, message: 'Pembelian tidak ditemukan.' };
      var pcd_ = periodClosedCheck_(head.date); if (pcd_) return pcd_;
      var hasPay = sheetToObjects_(getTradeSheet_(SHEET_TRADE_PAYMENTS)).some(function (r) { return r.kind === 'Utang' && String(r.ref_id) === String(purchaseId); });
      if (hasPay) return { success: false, message: 'Pembelian ini sudah punya pelunasan utang, tidak bisa dihapus.' };

      var lSheet = getTradeSheet_(SHEET_PURCHASE_LINES);
      var lines = sheetToObjects_(lSheet).filter(function (r) { return String(r.purchase_id) === String(purchaseId); });
      var itemSheet = getTradeSheet_(SHEET_ITEMS);
      var byId = {};
      sheetToObjects_(itemSheet).forEach(function (i) { byId[String(i.item_id)] = { row: i._row, name: String(i.name), stock: Number(i.stock) || 0, avg: Number(i.avg_cost) || 0 }; });

      // stok yang dibeli harus masih ada (belum terjual), kalau tidak pembelian tidak bisa dibatalkan
      for (var k = 0; k < lines.length; k++) {
        var it = byId[String(lines[k].item_id)];
        var qty = Number(lines[k].qty) || 0, cost = Number(lines[k].unit_cost) || 0;
        if (!it) continue;
        if (it.stock + 0.00001 < qty) return { success: false, message: 'Stok "' + it.name + '" (' + it.stock + ') lebih kecil dari jumlah yang dibeli (' + qty + '), sebagian sudah terjual. Pembelian tidak bisa dihapus.' };
        var remain = it.stock - qty;
        if (remain > 0) {
          var a = (it.stock * it.avg - qty * cost) / remain;
          if (a > 0) it.avg = a;
        }
        it.stock = remain;
      }
      Object.keys(byId).forEach(function (id) { itemSheet.getRange(byId[id].row, 6, 1, 2).setValues([[q4_(byId[id].stock), q4_(byId[id].avg)]]); });

      var jSheet = getSheetOrThrow_(SHEET_JOURNAL);
      deleteRowsDesc_(jSheet, sheetToObjects_(jSheet).filter(function (r) { return String(r.ref_no) === String(head.ref_no); }));
      purgeEvidence_(String(head.ref_no));
      deleteRowsDesc_(lSheet, lines);
      pSheet.deleteRow(head._row);
      SpreadsheetApp.flush();
      return { success: true, message: 'Pembelian ' + head.ref_no + ' dihapus, stok dan jurnalnya dikembalikan.' };
    });
  } catch (err) {
    Logger.log('deletePurchase ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ---------- PENJUALAN ----------
function getSales(token, startDate, endDate, kind) {
  var deny_ = checkPerm_(token, ['sales_tbs', 'sales_barang']); if (deny_) return deny_;
  var okKinds_ = allowedKinds_(token, ['sales']);
  if (kind && okKinds_.indexOf(kindOf_(kind)) < 0) return forbiddenResponse_();
  try {
    var lines = {};
    sheetToObjects_(getTradeSheet_(SHEET_SALE_LINES)).forEach(function (l) {
      var k = String(l.sale_id);
      (lines[k] = lines[k] || []).push({ item_name: String(l.item_name), qty: Number(l.qty) || 0, unit_price: Number(l.unit_price) || 0, subtotal: Number(l.subtotal) || 0, gross: Number(l.gross) || 0, tare: Number(l.tare) || 0, cut_pct: Number(l.cut_pct) || 0, cut_kg: Number(l.cut_kg) || 0 });
    });
    var data = sheetToObjects_(getTradeSheet_(SHEET_SALES))
      .filter(function (r) { return inRange_(fmtDate_(r.date), startDate, endDate) && (!kind || kindOf_(r.kind) === kind) && okKinds_.indexOf(kindOf_(r.kind)) >= 0; })
      .map(function (r) {
        var total = Number(r.total) || 0, paid = Number(r.paid) || 0, cost = Number(r.cost_total) || 0;
        var net = saleNet_(r);
        return {
          sale_id: String(r.sale_id), date: fmtDate_(r.date), ref_no: String(r.ref_no), customer: String(r.customer || ''),
          total: total, paid: paid, outstanding: Math.max(0, net - paid), status: tradeStatus_(net, paid), cost_total: cost,
          tax_ppn: Number(r.tax_ppn) || 0, tax_pph22: Number(r.tax_pph22) || 0, tax_pph23: Number(r.tax_pph23) || 0, tax_ref: String(r.tax_ref || ''), net_total: net,
          gross_profit: total - cost, note: String(r.note || ''), vehicle_no: String(r.vehicle_no || ''), driver: String(r.driver || ''), receivable_account_code: String(r.receivable_account_code || ''), kind: kindOf_(r.kind), unit: String(r.unit || ''), source_purchase_id: String(r.source_purchase_id || ''), lines: lines[String(r.sale_id)] || []
        };
      })
      .sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
    data = filterByUnits_(token, data);
    return { success: true, data: data };
  } catch (err) {
    Logger.log('getSales ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/**
 * p = { date, customer, note, paid, cash_account_code, lines: [{ item_id, qty, unit_price }] }
 * paid = jumlah diterima sekarang (0..total). Sisanya menjadi Piutang Usaha. Stok dikurangi & HPP dijurnalkan otomatis.
 */
function saveSale(token, p) {
  var deny_ = checkTradePerm_(token, 'sales', p && p.kind); if (deny_) return deny_;
  try {
    if (!p || !p.date) return { success: false, message: 'Tanggal wajib diisi.' };
    var inLines = p.lines || [];
    if (!inLines.length) return { success: false, message: 'Tambahkan minimal satu barang.' };
    var kind = kindOf_(p.kind);
    var pc_ = periodClosedCheck_(p.date); if (pc_) return pc_;
    var uc = checkUnit_(p.unit);
    if (uc.error) return { success: false, message: uc.error };
    var ua_ = checkUnitAccess_(token, uc.unit); if (ua_) return ua_;

    return withLock_(function () {
      var itemSheet = getTradeSheet_(SHEET_ITEMS);
      var byId = {};
      sheetToObjects_(itemSheet).forEach(function (i) {
        byId[String(i.item_id)] = { row: i._row, name: String(i.name), type: kindOf_(i.item_type), stock: Number(i.stock) || 0, avg: Number(i.avg_cost) || 0, touched: false };
      });

      var srcPurchase = '';
      if (p.source_purchase_id) {
        var sp = sheetToObjects_(getTradeSheet_(SHEET_PURCHASES)).filter(function (r) { return String(r.purchase_id) === String(p.source_purchase_id); })[0];
        if (!sp || kindOf_(sp.kind) !== 'TBS' || kind !== 'TBS') return { success: false, message: 'Pembelian sumber tidak valid (hanya untuk TBS).' };
        srcPurchase = String(p.source_purchase_id);
      }
      var saleId = Utilities.getUuid(), lineRows = [], total = 0, costTotal = 0;
      for (var k = 0; k < inLines.length; k++) {
        var L = inLines[k], it = byId[String(L.item_id)];
        var w = weighLine_(L, k), qty = w.qty, price = Number(L.unit_price) || 0;
        if (!it) return { success: false, message: 'Barang pada baris ' + (k + 1) + ' tidak ditemukan.' };
        if (it.type !== kind) return { success: false, message: 'Barang "' + it.name + '" berjenis ' + it.type + ', tidak bisa dipakai pada transaksi ' + kind + '.' };
        if (qty <= 0) return { success: false, message: 'Jumlah barang pada baris ' + (k + 1) + ' harus lebih dari 0.' };
        if (price <= 0) return { success: false, message: 'Harga jual pada baris ' + (k + 1) + ' harus lebih dari 0.' };
        if (it.stock + 0.00001 < qty) return { success: false, message: 'Stok "' + it.name + '" tidak cukup (tersisa ' + it.stock + ', diminta ' + qty + ').' };
        var sub = Math.round(qty * price);
        var lineCost = Math.round(qty * it.avg);
        lineRows.push([Utilities.getUuid(), saleId, String(L.item_id), it.name, qty, price, it.avg, sub, w.gross, w.tare, w.cut, w.cutKg]);
        it.stock = q4_(it.stock - qty);
        it.touched = true;
        total += sub;
        costTotal += lineCost;
      }

      var paid = (p.paid === undefined || p.paid === null || p.paid === '') ? total : Number(p.paid);
      if (isNaN(paid) || paid < 0 || paid > total) return { success: false, message: 'Jumlah diterima harus antara 0 dan total (' + total + ').' };

      var coaRows = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
      var t = tradeAccounts_(coaRows, kind);
      // akun piutang: boleh dipilih (mis. Piutang Usaha / Piutang Anggota); default Piutang Usaha
      var recv = null;
      if (total - paid > 0) {
        if (p.receivable_account_code) {
          recv = leafCoaRows_(coaRows).filter(function (r) { return String(r.account_code) === String(p.receivable_account_code) && r.category === 'Aset' && /piutang/i.test(String(r.account_name)); })[0];
          if (!recv) return { success: false, message: 'Akun piutang yang dipilih tidak valid (harus akun Piutang).' };
        } else recv = t.acc.piutang;
      }
      // akun pendapatan: boleh dipilih (sub akun Pendapatan); default Penjualan Barang Dagang
      var rev = null;
      if (p.revenue_account_code) {
        rev = leafCoaRows_(coaRows).filter(function (r) { return String(r.account_code) === String(p.revenue_account_code) && r.category === 'Pendapatan'; })[0];
        if (!rev) return { success: false, message: 'Akun pendapatan yang dipilih tidak valid (harus akun Pendapatan).' };
      } else rev = t.acc.penjualan;
      var need = (rev ? [] : ['penjualan']).concat(total - paid > 0 && !recv ? ['piutang'] : []).concat(costTotal > 0 ? ['hpp', 'persediaan'] : []);
      var miss = t.missing(need);
      if (miss) return { success: false, message: miss };
      var cash = null;
      if (paid > 0) {
        cash = findCashAccount_(coaRows, p.cash_account_code);
        if (!cash) return { success: false, message: 'Pilih akun Kas/Bank untuk penerimaan.' };
      }

      var refNo = tradeRef_('PJL', p.date, saleId), now = new Date();
      var customer = String(p.customer || '').trim();
      var desc = 'Penjualan ' + (kind === 'TBS' ? 'TBS' : 'barang') + (customer ? ' kepada ' + customer : '');
      var j = [];
      if (paid > 0) j.push(jrow_(p.date, refNo, cash, desc, paid, 0, now));
      if (total - paid > 0) j.push(jrow_(p.date, refNo, recv, desc, total - paid, 0, now));
      j.push(jrow_(p.date, refNo, rev, desc, 0, total, now));
      if (costTotal > 0) {
        j.push(jrow_(p.date, refNo, t.acc.hpp, 'HPP ' + desc, costTotal, 0, now));
        j.push(jrow_(p.date, refNo, t.acc.persediaan, 'HPP ' + desc, 0, costTotal, now));
      }

      appendTradeRows_(SHEET_SALES, [[saleId, p.date, refNo, customer, total, paid, costTotal, cash ? String(cash.account_code) : '', p.note || '', now, String(p.vehicle_no || '').trim(), String(p.driver || '').trim(), recv ? String(recv.account_code) : '', kind, srcPurchase, uc.unit]]);
      appendTradeRows_(SHEET_SALE_LINES, lineRows);
      appendJournalRows_(j, uc.unit);
      Object.keys(byId).forEach(function (id) {
        var it = byId[id];
        if (it.touched) itemSheet.getRange(it.row, 6).setValue(it.stock);
      });
      SpreadsheetApp.flush();
      Logger.log('saveSale: ' + refNo + ' total ' + total + ' hpp ' + costTotal);
      return { success: true, message: 'Penjualan tercatat dan dijurnalkan (' + refNo + ').', ref_no: refNo };
    });
  } catch (err) {
    Logger.log('saveSale ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deleteSale(token, saleId) {
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  var deny0_ = checkPerm_(token, ['sales_tbs', 'sales_barang']); if (deny0_) return deny0_;
  var deny_ = checkTradePerm_(token, 'sales', lookupTradeKind_(SHEET_SALES, 'sale_id', saleId)); if (deny_) return deny_;
  try {
    return withLock_(function () {
      var sSheet = getTradeSheet_(SHEET_SALES);
      var head = sheetToObjects_(sSheet).filter(function (r) { return String(r.sale_id) === String(saleId); })[0];
      if (!head) return { success: false, message: 'Penjualan tidak ditemukan.' };
      var pcd_ = periodClosedCheck_(head.date); if (pcd_) return pcd_;
      var hasPay = sheetToObjects_(getTradeSheet_(SHEET_TRADE_PAYMENTS)).some(function (r) { return r.kind === 'Piutang' && String(r.ref_id) === String(saleId); });
      if (hasPay) return { success: false, message: 'Penjualan ini sudah punya pelunasan piutang, tidak bisa dihapus.' };

      var lSheet = getTradeSheet_(SHEET_SALE_LINES);
      var lines = sheetToObjects_(lSheet).filter(function (r) { return String(r.sale_id) === String(saleId); });
      var itemSheet = getTradeSheet_(SHEET_ITEMS);
      var byId = {};
      sheetToObjects_(itemSheet).forEach(function (i) { byId[String(i.item_id)] = { row: i._row, stock: Number(i.stock) || 0, avg: Number(i.avg_cost) || 0 }; });

      // stok dikembalikan; avg_cost dihitung ulang dengan harga pokok saat barang dijual
      lines.forEach(function (l) {
        var it = byId[String(l.item_id)];
        if (!it) return;
        var qty = Number(l.qty) || 0, cost = Number(l.unit_cost) || 0;
        var newStock = it.stock + qty;
        it.avg = it.stock > 0 ? (it.stock * it.avg + qty * cost) / newStock : cost;
        it.stock = newStock;
      });
      Object.keys(byId).forEach(function (id) { itemSheet.getRange(byId[id].row, 6, 1, 2).setValues([[q4_(byId[id].stock), q4_(byId[id].avg)]]); });

      var jSheet = getSheetOrThrow_(SHEET_JOURNAL);
      var taxRef_ = String(head.tax_ref || '');
      deleteRowsDesc_(jSheet, sheetToObjects_(jSheet).filter(function (r) { return String(r.ref_no) === String(head.ref_no) || (taxRef_ && String(r.ref_no) === taxRef_); }));
      purgeEvidence_(String(head.ref_no));
      deleteRowsDesc_(lSheet, lines);
      sSheet.deleteRow(head._row);
      SpreadsheetApp.flush();
      return { success: true, message: 'Penjualan ' + head.ref_no + ' dihapus, stok dan jurnalnya dikembalikan.' };
    });
  } catch (err) {
    Logger.log('deleteSale ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ---------- PELUNASAN PIUTANG / UTANG ----------
/** p = { kind: 'Piutang' | 'Utang', ref_id, date, amount, cash_account_code } */
function payTrade(token, p) {
  var deny0_ = checkPerm_(token, ['purchases_tbs', 'purchases_barang', 'sales_tbs', 'sales_barang']); if (deny0_) return deny0_;
  if (p && p.ref_id && (p.kind === 'Piutang' || p.kind === 'Utang')) {
    var isS_ = p.kind === 'Piutang';
    var deny_ = checkTradePerm_(token, isS_ ? 'sales' : 'purchases', lookupTradeKind_(isS_ ? SHEET_SALES : SHEET_PURCHASES, isS_ ? 'sale_id' : 'purchase_id', p.ref_id));
    if (deny_) return deny_;
    var uaPay_ = checkUnitAccess_(token, lookupTradeUnit_(isS_ ? SHEET_SALES : SHEET_PURCHASES, isS_ ? 'sale_id' : 'purchase_id', p.ref_id)); if (uaPay_) return uaPay_;
  }
  try {
    if (!p || (p.kind !== 'Piutang' && p.kind !== 'Utang') || !p.ref_id || !p.date) return { success: false, message: 'Data pelunasan belum lengkap.' };
    var pcp_ = periodClosedCheck_(p.date); if (pcp_) return pcp_;
    var amount = Number(p.amount) || 0;
    if (amount <= 0) return { success: false, message: 'Jumlah pembayaran harus lebih dari 0.' };

    return withLock_(function () {
      var isSale = p.kind === 'Piutang';
      var sheet = getTradeSheet_(isSale ? SHEET_SALES : SHEET_PURCHASES);
      var idField = isSale ? 'sale_id' : 'purchase_id';
      var head = sheetToObjects_(sheet).filter(function (r) { return String(r[idField]) === String(p.ref_id); })[0];
      if (!head) return { success: false, message: 'Transaksi tidak ditemukan.' };
      var total = isSale ? saleNet_(head) : (Number(head.total) || 0), paid = Number(head.paid) || 0;
      var outstanding = total - paid;
      if (outstanding <= 0) return { success: false, message: 'Transaksi ini sudah lunas.' };
      if (amount > outstanding) return { success: false, message: 'Jumlah melebihi sisa ' + (isSale ? 'piutang' : 'utang') + ' (' + outstanding + ').' };

      var coaRows = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
      var t = tradeAccounts_(coaRows);
      var recvAcc = null;
      if (isSale && head.receivable_account_code) {
        recvAcc = leafCoaRows_(coaRows).filter(function (r) { return String(r.account_code) === String(head.receivable_account_code); })[0];
      }
      var miss = (isSale && recvAcc) ? null : t.missing([isSale ? 'piutang' : 'utang']);
      if (miss) return { success: false, message: miss };
      var cash = findCashAccount_(coaRows, p.cash_account_code);
      if (!cash) return { success: false, message: 'Pilih akun Kas/Bank.' };

      var payId = Utilities.getUuid(), now = new Date();
      var refNo = tradeRef_('PLN', p.date, payId);
      var who = isSale ? String(head.customer || '') : String(head.supplier || '');
      var desc = (isSale ? 'Penerimaan piutang ' : 'Pembayaran utang ') + head.ref_no + (who ? ' - ' + who : '');
      appendJournalRows_(isSale
        ? [jrow_(p.date, refNo, cash, desc, amount, 0, now), jrow_(p.date, refNo, recvAcc || t.acc.piutang, desc, 0, amount, now)]
        : [jrow_(p.date, refNo, t.acc.utang, desc, amount, 0, now), jrow_(p.date, refNo, cash, desc, 0, amount, now)], String(head.unit || ''));
      appendTradeRows_(SHEET_TRADE_PAYMENTS, [[payId, p.date, p.kind, String(p.ref_id), String(head.ref_no), amount, String(cash.account_code), refNo, now]]);
      sheet.getRange(head._row, 6).setValue(paid + amount);
      SpreadsheetApp.flush();
      Logger.log('payTrade: ' + refNo + ' ' + p.kind + ' ' + amount);
      return { success: true, message: 'Pembayaran tercatat dan dijurnalkan (' + refNo + '). Sisa: ' + (outstanding - amount) + '.' };
    });
  } catch (err) {
    Logger.log('payTrade ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}


// ============================================================
// IMPOR ANGGOTA (upload Excel/CSV)
// ============================================================
function importMembers(token, rows) {
  var deny_ = checkPerm_(token, ['members']); if (deny_) return deny_;
  try {
    if (!rows || !rows.length) return { success: false, message: 'Tidak ada data untuk diimpor.' };
    if (rows.length > 2000) return { success: false, message: 'Maksimal 2000 baris per impor. Pecah file menjadi beberapa bagian.' };
    return withLock_(function () {
      var sheet = getCoopSheet_(SHEET_MEMBERS);
      var existing = sheetToObjects_(sheet);
      var used = {};
      existing.forEach(function (r) { used[String(r.member_no).toLowerCase()] = true; });
      var max = 0;
      existing.forEach(function (r) { var n = parseInt(String(r.member_no).replace(/\D/g, ''), 10); if (!isNaN(n) && n > max) max = n; });

      var out = [], skipped = [], today = fmtDate_(new Date());
      rows.forEach(function (p, i) {
        var line = Number(p.row) || (i + 2);
        var name = String(p.name || '').trim();
        if (!name) { skipped.push({ row: line, reason: 'Nama kosong' }); return; }
        var join = String(p.join_date || '').trim() || today;
        if (!/^\d{4}-\d{2}-\d{2}$/.test(join) || isNaN(new Date(join).getTime())) { skipped.push({ row: line, reason: 'Tanggal masuk tidak valid (' + join + ')' }); return; }
        var status = String(p.status || '').trim().toLowerCase() === 'keluar' ? 'Keluar' : 'Aktif';
        var no = String(p.member_no || '').trim();
        if (no) {
          if (used[no.toLowerCase()]) { skipped.push({ row: line, reason: 'No. Anggota ' + no + ' sudah dipakai' }); return; }
        } else {
          do { max++; var s4 = String(max); while (s4.length < 4) s4 = '0' + s4; no = 'A-' + s4; } while (used[no.toLowerCase()]);
        }
        used[no.toLowerCase()] = true;
        out.push([Utilities.getUuid(), no, name, join, String(p.phone || '').trim(), String(p.address || '').trim(), status, status === 'Keluar' ? today : '']);
      });

      if (out.length) {
        var start = nextRow_(sheet, out.length);
        sheet.getRange(start, 2, out.length, 1).setNumberFormat('@');   // no. anggota
        sheet.getRange(start, 5, out.length, 1).setNumberFormat('@');   // telepon (jaga angka 0 di depan)
        sheet.getRange(start, 1, out.length, 8).setValues(out);
        SpreadsheetApp.flush();
      }
      Logger.log('importMembers: ' + out.length + ' masuk, ' + skipped.length + ' dilewati');
      return {
        success: true, imported: out.length, skipped: skipped,
        message: out.length + ' anggota berhasil diimpor' + (skipped.length ? ', ' + skipped.length + ' baris dilewati.' : '.')
      };
    });
  } catch (err) {
    Logger.log('importMembers ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}


// ============================================================
// BUKTI TRANSAKSI (upload JPG/PNG/PDF per No. Ref jurnal)
// File disimpan di folder Google Drive pemilik skrip; akses lewat aplikasi
// (cek sesi), bukan tautan Drive publik.
// ============================================================
var SHEET_EVIDENCE = 'Evidence';
var EVIDENCE_HEADERS = ['att_id', 'ref_no', 'file_id', 'file_name', 'mime', 'size', 'uploaded_by', 'timestamp'];
var EVIDENCE_MIMES = { 'image/jpeg': 'jpg', 'image/png': 'png', 'application/pdf': 'pdf' };
var EVIDENCE_MAX_BYTES = 4 * 1024 * 1024;
var EVIDENCE_MAX_PER_REF = 5;

function getEvidenceSheet_() {
  var ss = getSS_();
  var sh = ss.getSheetByName(SHEET_EVIDENCE);
  if (!sh) {
    sh = getOrCreateSheet_(SHEET_EVIDENCE, EVIDENCE_HEADERS);
    [1, 2, 3].forEach(function (c) { sh.getRange(2, c, 998, 1).setNumberFormat('@'); });
  }
  return sh;
}

function getEvidenceFolder_() {
  var props = PropertiesService.getDocumentProperties();
  var id = props.getProperty('evidence_folder_id');
  if (id) {
    try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* buat baru */ }
  }
  var folder = DriveApp.createFolder('Si Jurnal - Bukti Transaksi');
  props.setProperty('evidence_folder_id', folder.getId());
  return folder;
}

function purgeEvidence_(refNo) {
  try {
    var sh = getSS_().getSheetByName(SHEET_EVIDENCE);
    if (!sh) return;
    var rows = sheetToObjects_(sh).filter(function (r) { return String(r.ref_no) === String(refNo); });
    rows.sort(function (a, b) { return b._row - a._row; });
    rows.forEach(function (r) {
      try { DriveApp.getFileById(String(r.file_id)).setTrashed(true); } catch (e) { /* file sudah hilang */ }
      sh.deleteRow(r._row);
    });
  } catch (err) {
    Logger.log('purgeEvidence_ ERROR: ' + err.message);
  }
}

function uploadEvidence(token, p) {
  var deny_ = checkPerm_(token, ['journal']); if (deny_) return deny_;
  var user = requireSession_(token);
  if (!user) return sessionExpiredResponse_();
  try {
    if (!p || !p.ref_no || !p.data) return { success: false, message: 'Data bukti belum lengkap.' };
    var mime = String(p.mime || '');
    if (!EVIDENCE_MIMES[mime]) return { success: false, message: 'Format file harus JPG, PNG, atau PDF.' };
    return withLock_(function () {
      var refNo = String(p.ref_no);
      var exists = sheetToObjects_(getSheetOrThrow_(SHEET_JOURNAL)).some(function (r) { return String(r.ref_no) === refNo; });
      if (!exists) return { success: false, message: 'No. Ref jurnal "' + refNo + '" tidak ditemukan.' };
      var sh = getEvidenceSheet_();
      var count = sheetToObjects_(sh).filter(function (r) { return String(r.ref_no) === refNo; }).length;
      if (count >= EVIDENCE_MAX_PER_REF) return { success: false, message: 'Maksimal ' + EVIDENCE_MAX_PER_REF + ' bukti per jurnal.' };
      var bytes = Utilities.base64Decode(String(p.data));
      if (bytes.length > EVIDENCE_MAX_BYTES) return { success: false, message: 'Ukuran file maksimal 4 MB.' };
      var base = String(p.file_name || 'bukti').replace(/\.[^.]+$/, '').replace(/[^\w\- ]+/g, '_').slice(0, 60) || 'bukti';
      var name = refNo.replace(/[^\w\-]+/g, '_') + '_' + base + '.' + EVIDENCE_MIMES[mime];
      var file = getEvidenceFolder_().createFile(Utilities.newBlob(bytes, mime, name));
      var attId = Utilities.getUuid();
      var start = nextRow_(sh);
      [1, 2, 3].forEach(function (c) { sh.getRange(start, c).setNumberFormat('@'); });
      sh.getRange(start, 1, 1, 8).setValues([[attId, refNo, file.getId(), name, mime, bytes.length, user, new Date()]]);
      SpreadsheetApp.flush();
      Logger.log('uploadEvidence: ' + refNo + ' ' + name + ' ' + bytes.length + ' byte');
      return { success: true, message: 'Bukti berhasil diunggah.', att_id: attId };
    });
  } catch (err) {
    Logger.log('uploadEvidence ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function getEvidenceIndex(token) {
  var deny_ = checkPerm_(token, ['journal', 'reports']); if (deny_) return deny_;
  try {
    var sh = getSS_().getSheetByName(SHEET_EVIDENCE);
    var map = {};
    if (sh) sheetToObjects_(sh).forEach(function (r) { var k = String(r.ref_no); map[k] = (map[k] || 0) + 1; });
    return { success: true, data: map };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getEvidenceList(token, refNo) {
  var deny_ = checkPerm_(token, ['journal', 'reports']); if (deny_) return deny_;
  try {
    var sh = getSS_().getSheetByName(SHEET_EVIDENCE);
    var list = sh ? sheetToObjects_(sh).filter(function (r) { return String(r.ref_no) === String(refNo); }).map(function (r) {
      return { att_id: String(r.att_id), file_name: String(r.file_name), mime: String(r.mime), size: Number(r.size) || 0, uploaded_by: String(r.uploaded_by), timestamp: fmtDate_(r.timestamp) };
    }) : [];
    return { success: true, data: list };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getEvidenceFile(token, attId) {
  var deny_ = checkPerm_(token, ['journal', 'reports']); if (deny_) return deny_;
  try {
    var sh = getSS_().getSheetByName(SHEET_EVIDENCE);
    var r = sh ? sheetToObjects_(sh).filter(function (x) { return String(x.att_id) === String(attId); })[0] : null;
    if (!r) return { success: false, message: 'Bukti tidak ditemukan.' };
    var blob = DriveApp.getFileById(String(r.file_id)).getBlob();
    return { success: true, name: String(r.file_name), mime: String(r.mime), data: Utilities.base64Encode(blob.getBytes()) };
  } catch (err) {
    Logger.log('getEvidenceFile ERROR: ' + err.message);
    return { success: false, message: 'Gagal membuka file: ' + err.message };
  }
}

function deleteEvidence(token, attId) {
  var deny_ = checkPerm_(token, ['journal']); if (deny_) return deny_;
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    return withLock_(function () {
      var sh = getSS_().getSheetByName(SHEET_EVIDENCE);
      var r = sh ? sheetToObjects_(sh).filter(function (x) { return String(x.att_id) === String(attId); })[0] : null;
      if (!r) return { success: false, message: 'Bukti tidak ditemukan.' };
      try { DriveApp.getFileById(String(r.file_id)).setTrashed(true); } catch (e) { /* sudah hilang */ }
      sh.deleteRow(r._row);
      return { success: true, message: 'Bukti dihapus.' };
    });
  } catch (err) {
    return { success: false, message: err.message };
  }
}


// ==================== MANAJEMEN ROLE & PENGGUNA (Administrator) ====================
function getMyPerms(token) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  var role = CacheService.getScriptCache().get('role_' + token);
  return { success: true, role: role, perms: permsForRole_(role), units: unitsForRole_(role) };
}

function getRoles(token) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  if (!requireAdmin_(token)) return notAdminResponse_();
  var cfg = getRolesConfig_();
  var users = sheetToObjects_(getSheetOrThrow_(SHEET_USERS));
  var data = Object.keys(cfg).map(function (name) {
    return {
      name: name, permissions: permsForRole_(name), units: unitsForRole_(name),
      user_count: users.filter(function (u) { return String(u.role) === name; }).length
    };
  }).sort(function (a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
  return { success: true, data: data, catalog: PERM_CATALOG_, units_list: getUnitsList_() };
}

// p = { name, permissions:[...] } - buat baru atau perbarui hak sebuah role.
function saveRole(token, p) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  if (!requireAdmin_(token)) return notAdminResponse_();
  try {
    var name = String((p && p.name) || '').trim();
    if (name.length < 2 || name.length > 30) return { success: false, message: 'Nama role 2-30 karakter.' };
    if (RESERVED_ROLES_.indexOf(name.toLowerCase()) >= 0) return { success: false, message: 'Nama "' + name + '" dipakai sistem, pilih nama lain.' };
    var valid = allPermKeys_();
    var perms = ((p && p.permissions) || []).filter(function (k, i, a) { return valid.indexOf(k) >= 0 && a.indexOf(k) === i; });
    // Batas unit untuk pembelian, penjualan, dan kas. all_units=false -> hanya unit terpilih.
    var units = null;
    if (p && p.all_units === false) {
      var list = getUnitsList_();
      units = ((p.units) || []).filter(function (u, i, a) { return list.indexOf(u) >= 0 && a.indexOf(u) === i; });
      if (!units.length) return { success: false, message: 'Pilih minimal satu unit, atau pilih "Semua unit".' };
    }
    return withLock_(function () {
      var cfg = getRolesConfig_();
      // nama role tidak peka huruf besar/kecil
      Object.keys(cfg).forEach(function (n) { if (n !== name && n.toLowerCase() === name.toLowerCase()) { cfg[name] = cfg[n]; delete cfg[n]; } });
      var isNew = !cfg[name];
      cfg[name] = units ? { permissions: perms, units: units } : { permissions: perms };
      PropertiesService.getDocumentProperties().setProperty('roles_config', JSON.stringify(cfg));
      return { success: true, message: 'Role "' + name + '" ' + (isNew ? 'dibuat' : 'diperbarui') + '.' };
    });
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteRole(token, name) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  if (!requireAdmin_(token)) return notAdminResponse_();
  try {
    name = String(name || '');
    return withLock_(function () {
      var cfg = getRolesConfig_();
      if (!cfg[name]) return { success: false, message: 'Role tidak ditemukan.' };
      var used = sheetToObjects_(getSheetOrThrow_(SHEET_USERS)).filter(function (u) { return String(u.role) === name; }).length;
      if (used) return { success: false, message: 'Role "' + name + '" masih dipakai ' + used + ' pengguna. Pindahkan/hapus penggunanya dulu.' };
      delete cfg[name];
      PropertiesService.getDocumentProperties().setProperty('roles_config', JSON.stringify(cfg));
      return { success: true, message: 'Role "' + name + '" dihapus.' };
    });
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getStaffUsers(token) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  if (!requireAdmin_(token)) return notAdminResponse_();
  var data = sheetToObjects_(getSheetOrThrow_(SHEET_USERS))
    .filter(function (u) { return String(u.role) !== 'Anggota'; })
    .map(function (u) { return { username: String(u.username), role: String(u.role) }; });
  return { success: true, data: data };
}

// p = { username, role, password } - password wajib untuk akun baru, opsional (reset) untuk akun lama.
function saveStaffUser(token, p) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  if (!requireAdmin_(token)) return notAdminResponse_();
  try {
    var username = String((p && p.username) || '').trim();
    var role = String((p && p.role) || '');
    var password = String((p && p.password) || '');
    if (!/^[A-Za-z0-9._-]{3,30}$/.test(username)) return { success: false, message: 'Username 3-30 karakter: huruf, angka, titik, strip, atau garis bawah.' };
    if (!getRolesConfig_()[role]) return { success: false, message: 'Pilih role yang valid.' };
    return withLock_(function () {
      var uSheet = getSheetOrThrow_(SHEET_USERS);
      ensureUsersColumns_(uSheet);
      var users = sheetToObjects_(uSheet);
      var ex = users.filter(function (u) { return String(u.username).toLowerCase() === username.toLowerCase(); })[0];
      if (ex) {
        if (String(ex.role) === 'Administrator' || String(ex.role) === 'Anggota') return { success: false, message: 'Akun "' + ex.username + '" (' + ex.role + ') tidak bisa diubah dari menu ini.' };
        if (password && password.length < 6) return { success: false, message: 'Password minimal 6 karakter.' };
        uSheet.getRange(ex._row, 3).setValue(role);
        if (password) {
          uSheet.getRange(ex._row, 2).setNumberFormat('@');
          uSheet.getRange(ex._row, 2).setValue(hashPassword_(password));
        }
        SpreadsheetApp.flush();
        return { success: true, message: 'Akun "' + ex.username + '" diperbarui' + (password ? ' (password direset)' : '') + '. Perubahan role berlaku setelah pengguna login ulang.' };
      }
      if (password.length < 6) return { success: false, message: 'Password minimal 6 karakter.' };
      var row = nextRow_(uSheet);
      uSheet.getRange(row, 1, 1, 4).setNumberFormat('@');
      uSheet.getRange(row, 1, 1, 4).setValues([[username, hashPassword_(password), role, '']]);
      SpreadsheetApp.flush();
      return { success: true, message: 'Akun "' + username + '" (' + role + ') berhasil dibuat.' };
    });
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteStaffUser(token, username) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  if (!requireAdmin_(token)) return notAdminResponse_();
  try {
    username = String(username || '');
    if (username.toLowerCase() === String(requireSession_(token)).toLowerCase()) return { success: false, message: 'Anda tidak bisa menghapus akun yang sedang dipakai.' };
    return withLock_(function () {
      var uSheet = getSheetOrThrow_(SHEET_USERS);
      var ex = sheetToObjects_(uSheet).filter(function (u) { return String(u.username).toLowerCase() === username.toLowerCase(); })[0];
      if (!ex) return { success: false, message: 'Akun tidak ditemukan.' };
      if (String(ex.role) === 'Administrator' || String(ex.role) === 'Anggota') return { success: false, message: 'Akun ' + ex.role + ' tidak bisa dihapus dari menu ini.' };
      uSheet.deleteRow(ex._row);
      SpreadsheetApp.flush();
      return { success: true, message: 'Akun "' + ex.username + '" dihapus.' };
    });
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==================== TUTUP BUKU TAHUNAN ====================
var SHEET_CLOSING_LOG = 'ClosingLog';

function getClosedDate_() {
  return PropertiesService.getDocumentProperties().getProperty('last_closed_date') || '';
}

// null jika tanggal masih di periode terbuka; selain itu respons penolakan.
function periodClosedCheck_(date) {
  var closed = getClosedDate_();
  var d = fmtDate_(date);
  if (closed && d && d <= closed) {
    return { success: false, message: 'Periode sampai ' + closed + ' sudah ditutup buku. Transaksi bertanggal ' + d + ' tidak bisa dibuat atau dihapus.' };
  }
  return null;
}

function getArchiveFolder_() {
  var props = PropertiesService.getDocumentProperties();
  var id = props.getProperty('archive_folder_id');
  if (id) {
    try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* buat baru */ }
  }
  var folder = DriveApp.createFolder('Si Jurnal - Arsip Tutup Buku');
  props.setProperty('archive_folder_id', folder.getId());
  return folder;
}

function getClosingLogSheet_() {
  var ss = getSS_();
  var sh = ss.getSheetByName(SHEET_CLOSING_LOG);
  if (!sh) sh = getOrCreateSheet_(SHEET_CLOSING_LOG, ['closed_at', 'cutoff', 'backup_url', 'backup_id', 'journal_removed', 'trade_removed', 'profit', 'by', 'status']);
  return sh;
}

function isLeafCode_(code, allCodes) {
  var prefix = String(code) + '.';
  return !allCodes.some(function (c) { return String(c).indexOf(prefix) === 0; });
}

// Hitung rencana tutup buku tanpa mengubah data apa pun.
function computeClosingPlan_(cutoff, equityCode) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(cutoff || ''))) return { error: 'Tanggal tutup buku tidak valid.' };
  var today = fmtDate_(new Date());
  if (cutoff >= today) return { error: 'Tanggal tutup buku harus sebelum hari ini (' + today + ').' };
  var last = getClosedDate_();
  if (last && cutoff <= last) return { error: 'Periode sampai ' + last + ' sudah ditutup. Pilih tanggal setelahnya.' };

  var coaSheet = getSheetOrThrow_(SHEET_COA);
  var coaRows = sheetToObjects_(coaSheet);
  var codes = coaRows.map(function (r) { return String(r.account_code); });
  var eq = coaRows.filter(function (r) { return String(r.account_code) === String(equityCode); })[0];
  if (!eq || eq.category !== 'Ekuitas' || !isLeafCode_(eq.account_code, codes)) {
    return { error: 'Pilih akun Ekuitas (akun detail) sebagai penampung SHU/laba tahun berjalan.' };
  }

  // --- Jurnal
  var jSheet = getSheetOrThrow_(SHEET_JOURNAL);
  var jVals = jSheet.getDataRange().getValues();
  var jHead = jVals[0].map(String);
  var iDate = jHead.indexOf('date'), iAcc = jHead.indexOf('account_code'), iDeb = jHead.indexOf('debit'), iCred = jHead.indexOf('credit');
  if (iDate < 0 || iAcc < 0 || iDeb < 0 || iCred < 0) return { error: 'Struktur sheet Journal tidak dikenali.' };
  var move = {}, removed = 0, keptDC = 0, totalBeforeDC = 0, jDebit = 0, jCredit = 0;
  var knownCodes = {};
  codes.forEach(function (c) { knownCodes[c] = true; });
  for (var i = 1; i < jVals.length; i++) {
    var row = jVals[i];
    var dc = (Number(row[iDeb]) || 0) - (Number(row[iCred]) || 0);
    totalBeforeDC += dc;
    if (fmtDate_(row[iDate]) <= cutoff) {
      var code = String(row[iAcc]);
      if (!knownCodes[code]) return { error: 'Jurnal memakai kode akun "' + code + '" yang tidak ada di COA. Perbaiki dulu sebelum tutup buku.' };
      move[code] = (move[code] || 0) + dc;
      removed++;
      jDebit += Number(row[iDeb]) || 0; jCredit += Number(row[iCred]) || 0;
    } else {
      keptDC += dc;
    }
  }

  // --- Saldo awal baru per akun; akun Pendapatan/Beban ditutup ke akun ekuitas
  var openBefore = 0, openAfter = 0, profit = 0;
  var newOpen = coaRows.map(function (r) {
    var od = Number(r.opening_debit) || 0, oc = Number(r.opening_credit) || 0;
    openBefore += od - oc;
    var net = (od - oc) + (move[String(r.account_code)] || 0);
    if (r.category === 'Pendapatan' || r.category === 'Beban') {
      profit += -net;            // kredit - debit
      net = 0;
    }
    return net;
  });
  newOpen = newOpen.map(function (net, idx) {
    if (String(coaRows[idx].account_code) === String(eq.account_code)) net -= profit;   // laba menambah saldo kredit ekuitas
    openAfter += net;
    return net;
  });
  // Pemeriksaan keseimbangan: total (debit-kredit) seluruh buku sebelum == sesudah
  var before = openBefore + totalBeforeDC, after = openAfter + keptDC;
  if (Math.abs(before - after) > 0.5) return { error: 'Pemeriksaan keseimbangan gagal (selisih ' + (before - after) + '). Tutup buku dibatalkan, tidak ada data yang diubah.' };

  // --- Transaksi jual-beli yang sudah lunas dan bertanggal <= cutoff
  var pRows = sheetToObjects_(getTradeSheet_(SHEET_PURCHASES)), sRows = sheetToObjects_(getTradeSheet_(SHEET_SALES));
  var settled = function (r) { return fmtDate_(r.date) <= cutoff && saleNet_(r) - (Number(r.paid) || 0) <= 0.5; };
  var pDel = {}, sDel = {}, pCount = 0, sCount = 0, openP = 0, openS = 0;
  pRows.forEach(function (r) { if (settled(r)) { pDel[String(r.purchase_id)] = true; pCount++; } else if (fmtDate_(r.date) <= cutoff) openP++; });
  sRows.forEach(function (r) { if (settled(r)) { sDel[String(r.sale_id)] = true; sCount++; } else if (fmtDate_(r.date) <= cutoff) openS++; });

  return {
    cutoff: cutoff, equity: eq, coaRows: coaRows, newOpen: newOpen, profit: profit,
    journal: { removed: removed, kept: jVals.length - 1 - removed, debit: jDebit, credit: jCredit, iDate: iDate },
    trade: { purchases: pCount, sales: sCount, pDel: pDel, sDel: sDel, openPurchases: openP, openSales: openS }
  };
}

function planSummary_(plan) {
  return {
    cutoff: plan.cutoff, equity_name: plan.equity.account_name, profit: plan.profit,
    journal_removed: plan.journal.removed, journal_kept: plan.journal.kept, journal_debit: plan.journal.debit, journal_credit: plan.journal.credit,
    purchases_removed: plan.trade.purchases, sales_removed: plan.trade.sales,
    purchases_open: plan.trade.openPurchases, sales_open: plan.trade.openSales,
    accounts: plan.coaRows.length
  };
}

function getClosingInfo(token) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  if (!requireAdmin_(token)) return notAdminResponse_();
  try {
    var coaRows = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
    var codes = coaRows.map(function (r) { return String(r.account_code); });
    var equity = coaRows.filter(function (r) { return r.category === 'Ekuitas' && isLeafCode_(r.account_code, codes); })
      .map(function (r) { return { code: String(r.account_code), name: String(r.account_name) }; });
    var logSheet = getSS_().getSheetByName(SHEET_CLOSING_LOG);
    var log = logSheet ? sheetToObjects_(logSheet).map(function (r) {
      return { closed_at: fmtDate_(r.closed_at), cutoff: fmtDate_(r.cutoff), backup_url: String(r.backup_url || ''), journal_removed: Number(r.journal_removed) || 0,
        trade_removed: Number(r.trade_removed) || 0, profit: Number(r.profit) || 0, by: String(r.by || ''), status: String(r.status || '') };
    }).reverse() : [];
    return { success: true, last_closed: getClosedDate_(), equity_accounts: equity, log: log };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function previewClosing(token, p) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  if (!requireAdmin_(token)) return notAdminResponse_();
  try {
    var plan = computeClosingPlan_(p && p.cutoff, p && p.equity_code);
    if (plan.error) return { success: false, message: plan.error };
    return { success: true, summary: planSummary_(plan) };
  } catch (err) {
    Logger.log('previewClosing ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// Tulis ulang sheet hanya dengan baris yang dipertahankan. Jika gagal di tengah, data asli ditulis kembali.
function rewriteKeepRows_(sheet, keepFn) {
  return sheet.__filterRows(keepFn);
}

function runClosing(token, p) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  if (!requireAdmin_(token)) return notAdminResponse_();
  try {
    if (!p || String(p.confirm || '').trim().toUpperCase() !== 'TUTUP BUKU') return { success: false, message: 'Ketik "TUTUP BUKU" pada kolom konfirmasi.' };
    return withLock_(function () {
      var plan = computeClosingPlan_(p.cutoff, p.equity_code);
      if (plan.error) return { success: false, message: plan.error };
      var user = requireSession_(token);
      var tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
      var ss = getSS_();

      // 1) CADANGAN: potret seluruh tabel (sebelum diubah) disimpan sebagai file JSON di Storage Supabase.
      //    File diunggah lebih dulu; bila unggah gagal, seluruh perubahan database dibatalkan.
      var name = 'Arsip Si Jurnal s.d. ' + plan.cutoff + ' (dibuat ' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmm') + ')';
      var backup = SpreadsheetApp.__createBackup(name);
      var backupUrl = backup.url;

      // 2) Catat log, lalu terapkan perubahan
      var logSheet = getClosingLogSheet_();
      var logRow = nextRow_(logSheet);
      logSheet.getRange(logRow, 1, 1, 9).setValues([[new Date(), plan.cutoff, backupUrl, backup.id, 0, 0, plan.profit, user, 'berjalan']]);

      var coaSheet = getSheetOrThrow_(SHEET_COA);
      var oldOpen = coaSheet.getRange(2, 6, plan.coaRows.length, 2).getValues();
      var tradeRemoved = 0, jRemoved = 0;
      try {
        coaSheet.getRange(2, 6, plan.coaRows.length, 2).setValues(plan.newOpen.map(function (n) { return [n > 0 ? n : 0, n < 0 ? -n : 0]; }));
        var jSheet = getSheetOrThrow_(SHEET_JOURNAL);
        ensureJournalUnitCol_(jSheet);
        jRemoved = rewriteKeepRows_(jSheet, function (row) { return fmtDate_(row[plan.journal.iDate]) > plan.cutoff; });

        var pSheet = getTradeSheet_(SHEET_PURCHASES), sSheet = getTradeSheet_(SHEET_SALES);
        var pDel = plan.trade.pDel, sDel = plan.trade.sDel;
        var ph = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues()[0].map(String), sh = sSheet.getRange(1, 1, 1, sSheet.getLastColumn()).getValues()[0].map(String);
        var pi = ph.indexOf('purchase_id'), si = sh.indexOf('sale_id');
        tradeRemoved += rewriteKeepRows_(pSheet, function (r) { return !pDel[String(r[pi])]; });
        tradeRemoved += rewriteKeepRows_(sSheet, function (r) { return !sDel[String(r[si])]; });
        tradeRemoved += rewriteKeepRows_(getTradeSheet_(SHEET_PURCHASE_LINES), function (r) { return !pDel[String(r[1])]; });
        tradeRemoved += rewriteKeepRows_(getTradeSheet_(SHEET_SALE_LINES), function (r) { return !sDel[String(r[1])]; });
        tradeRemoved += rewriteKeepRows_(getTradeSheet_(SHEET_TRADE_PAYMENTS), function (r) { return !(pDel[String(r[3])] || sDel[String(r[3])]); });

        PropertiesService.getDocumentProperties().setProperty('last_closed_date', plan.cutoff);
        SpreadsheetApp.flush();
      } catch (err) {
        try { coaSheet.getRange(2, 6, plan.coaRows.length, 2).setValues(oldOpen); } catch (e2) { /* sudah dicoba */ }
        logSheet.getRange(logRow, 9).setValue('GAGAL: ' + err.message);
        return { success: false, message: 'Tutup buku gagal di tengah proses: ' + err.message + '. Saldo awal sudah dikembalikan; jika data terlihat janggal, hubungi pengelola aplikasi.' };
      }
      logSheet.getRange(logRow, 5, 1, 2).setValues([[jRemoved, tradeRemoved]]);
      logSheet.getRange(logRow, 9).setValue('selesai');
      Logger.log('runClosing: s.d. ' + plan.cutoff + ' jurnal ' + jRemoved + ' trade ' + tradeRemoved);
      return { success: true, message: 'Tutup buku s.d. ' + plan.cutoff + ' selesai. Cadangan tersimpan di Storage Supabase.', backup_url: backupUrl, summary: planSummary_(plan) };
    });
  } catch (err) {
    Logger.log('runClosing ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}


// ============================================================
// PAJAK INVOICE (PPN / PPh 22 / PPh 23) -> JURNAL
// ============================================================
// Saat invoice dibuat dengan pajak, dicatat jurnal pajak per penjualan (sekali saja):
//   Debit  Piutang (PPN - PPh)          <- tagihan bertambah PPN, berkurang PPh yang dipotong pembeli
//   Debit  PPh Dibayar Dimuka           <- PPh 22 / PPh 23 yang dipotong pembeli (kredit pajak)
//   Kredit Utang PPN                    <- PPN keluaran
// Sisa piutang penjualan = harga + PPN - PPh - dibayar.
var TAX_KEYS_ = ['ppn', 'pph22', 'pph23'];

function coopProfileRaw_() {
  var raw = PropertiesService.getDocumentProperties().getProperty('coop_profile');
  var saved = raw ? JSON.parse(raw) : {};
  var p = {};
  Object.keys(DEFAULT_COOP_PROFILE).forEach(function (k) { p[k] = saved[k] !== undefined ? saved[k] : DEFAULT_COOP_PROFILE[k]; });
  if (saved.ppn_dpp_lain === undefined && Number(saved.ppn_rate) === 11) p.ppn_rate = 12;
  return p;
}

/** Hitung pajak dari harga jual (DPP). Harus sama dengan perhitungan di invoice (Index.html: invoiceFoot). */
function calcInvoiceTax_(profile, total, keys) {
  var out = { ppn: 0, pph22: 0, pph23: 0 };
  keys.forEach(function (k) {
    var rate = Number(profile[k + '_rate']) || 0;
    if (rate <= 0) return;
    var base = total;
    if (k === 'ppn' && profile.ppn_dpp_lain !== false) base = Math.round(total * 11 / 12);
    out[k] = Math.round(base * rate / 100);
  });
  return out;
}

/** p = { sale_ids: [..], taxes: ['ppn','pph22','pph23'], date? }  Jurnal pajak per penjualan; penjualan yang sudah pernah dijurnal dilewati. */
function postInvoiceTax(token, p) {
  var deny0_ = checkPerm_(token, ['sales_tbs', 'sales_barang']); if (deny0_) return deny0_;
  try {
    if (!p || !p.sale_ids || !p.sale_ids.length) return { success: false, message: 'Penjualan belum dipilih.' };
    var keys = (p.taxes || []).filter(function (k) { return TAX_KEYS_.indexOf(k) >= 0; });
    var date = p.date && /^\d{4}-\d{2}-\d{2}$/.test(p.date) ? p.date : fmtDate_(new Date());
    var pcp_ = periodClosedCheck_(date); if (pcp_) return pcp_;
    return withLock_(function () {
      var sheet = getTradeSheet_(SHEET_SALES);
      var all = sheetToObjects_(sheet), byId = {};
      all.forEach(function (r) { byId[String(r.sale_id)] = r; });
      var heads = [];
      for (var i = 0; i < p.sale_ids.length; i++) {
        var h = byId[String(p.sale_ids[i])];
        if (!h) return { success: false, message: 'Penjualan tidak ditemukan.' };
        var d1 = checkTradePerm_(token, 'sales', h.kind); if (d1) return d1;
        var d2 = checkUnitAccess_(token, String(h.unit || '')); if (d2) return d2;
        heads.push(h);
      }
      var result = [];
      var todo = heads.filter(function (h) { return !String(h.tax_ref || ''); });
      if (keys.length && todo.length) {
        var profile = coopProfileRaw_();
        keys.forEach(function (k) { if (!(Number(profile[k + '_rate']) > 0)) throw new Error('Tarif ' + k.toUpperCase() + ' belum diisi di Profil Usaha.'); });
        var coaRows = sheetToObjects_(getSheetOrThrow_(SHEET_COA));
        var t = tradeAccounts_(coaRows);
        var need = [];
        if (keys.indexOf('ppn') >= 0) need.push('ppn');
        if (keys.indexOf('pph22') >= 0 || keys.indexOf('pph23') >= 0) need.push('pph');
        var miss = t.missing(need);
        if (miss) return { success: false, message: miss };
        var now = new Date();
        todo.forEach(function (h) {
          var total = Number(h.total) || 0;
          var tx = calcInvoiceTax_(profile, total, keys);
          var recv = null;
          if (h.receivable_account_code) recv = leafCoaRows_(coaRows).filter(function (r) { return String(r.account_code) === String(h.receivable_account_code); })[0];
          if (!recv) recv = t.acc.piutang;
          if (!recv) throw new Error(t.missing(['piutang']));
          var refNo = tradeRef_('PJK', date, String(h.sale_id));
          var desc = 'Pajak invoice ' + String(h.ref_no).replace(/^PJL/, 'INV') + (h.customer ? ' - ' + h.customer : '');
          var j = [];
          var pphTotal = tx.pph22 + tx.pph23;
          var recvNet = tx.ppn - pphTotal;
          if (recvNet > 0) j.push(jrow_(date, refNo, recv, desc, recvNet, 0, now));
          if (tx.pph22 > 0) j.push(jrow_(date, refNo, t.acc.pph, desc + ' (PPh 22 dipotong)', tx.pph22, 0, now));
          if (tx.pph23 > 0) j.push(jrow_(date, refNo, t.acc.pph, desc + ' (PPh 23 dipotong)', tx.pph23, 0, now));
          if (recvNet < 0) j.push(jrow_(date, refNo, recv, desc, 0, -recvNet, now));
          if (tx.ppn > 0) j.push(jrow_(date, refNo, t.acc.ppn, desc + ' (PPN keluaran)', 0, tx.ppn, now));
          if (!j.length) return;
          appendJournalRows_(j, String(h.unit || ''));
          sheet.getRange(h._row, 17, 1, 4).setValues([[tx.ppn, tx.pph22, tx.pph23, refNo]]);
          sheet.getRange(h._row, 20).setNumberFormat('@');
          h.tax_ppn = tx.ppn; h.tax_pph22 = tx.pph22; h.tax_pph23 = tx.pph23; h.tax_ref = refNo;
        });
        SpreadsheetApp.flush();
      }
      heads.forEach(function (h) {
        result.push({ sale_id: String(h.sale_id), tax_ppn: Number(h.tax_ppn) || 0, tax_pph22: Number(h.tax_pph22) || 0, tax_pph23: Number(h.tax_pph23) || 0, tax_ref: String(h.tax_ref || '') });
      });
      var posted = todo.filter(function (h) { return String(h.tax_ref || ''); }).length;
      return { success: true, data: result, posted: posted, message: posted ? 'Pajak ' + posted + ' penjualan dijurnalkan.' : 'Tidak ada jurnal pajak baru.' };
    });
  } catch (err) {
    Logger.log('postInvoiceTax ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ============================================================
// RIWAYAT & PEMBATALAN PELUNASAN PIUTANG / UTANG
// ============================================================
function tradeSheetFor_(kind) { return kind === 'Piutang' ? SHEET_SALES : SHEET_PURCHASES; }

/** Daftar pelunasan satu transaksi. kind: 'Piutang' | 'Utang'. */
function getTradePayments(token, kind, refId) {
  var isS = kind === 'Piutang';
  var deny0_ = checkPerm_(token, isS ? ['sales_tbs', 'sales_barang'] : ['purchases_tbs', 'purchases_barang', 'sales_tbs']); if (deny0_) return deny0_;
  try {
    var data = sheetToObjects_(getTradeSheet_(SHEET_TRADE_PAYMENTS))
      .filter(function (r) { return r.kind === kind && String(r.ref_id) === String(refId); })
      .map(function (r) { return { pay_id: String(r.pay_id), date: fmtDate_(r.date), amount: Number(r.amount) || 0, cash_account_code: String(r.cash_account_code || ''), journal_ref: String(r.journal_ref || '') }; })
      .sort(function (a, b) { return a.date < b.date ? -1 : 1; });
    return { success: true, data: data };
  } catch (err) {
    Logger.log('getTradePayments ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** Batalkan satu pelunasan (hanya Administrator): jurnalnya dihapus, jumlah dibayar dikurangi, catatan pelunasan dihapus. */
function cancelTradePayment(token, payId) {
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    return withLock_(function () {
      var pSheet = getTradeSheet_(SHEET_TRADE_PAYMENTS);
      var pay = sheetToObjects_(pSheet).filter(function (r) { return String(r.pay_id) === String(payId); })[0];
      if (!pay) return { success: false, message: 'Pelunasan tidak ditemukan.' };
      var pcp_ = periodClosedCheck_(pay.date); if (pcp_) return pcp_;
      var isSale = pay.kind === 'Piutang';
      var tSheet = getTradeSheet_(tradeSheetFor_(pay.kind));
      var idField = isSale ? 'sale_id' : 'purchase_id';
      var head = sheetToObjects_(tSheet).filter(function (r) { return String(r[idField]) === String(pay.ref_id); })[0];
      if (head) {
        var d1 = checkTradePerm_(token, isSale ? 'sales' : 'purchases', head.kind); if (d1) return d1;
        var amount = Number(pay.amount) || 0, paid = Number(head.paid) || 0;
        tSheet.getRange(head._row, 6).setValue(Math.max(0, paid - amount));
      }
      var jSheet = getSheetOrThrow_(SHEET_JOURNAL);
      var jref = String(pay.journal_ref || '');
      var jrows = jref ? sheetToObjects_(jSheet).filter(function (r) { return String(r.ref_no) === jref; }) : [];
      deleteRowsDesc_(jSheet, jrows);
      purgeEvidence_(jref);
      pSheet.deleteRow(pay._row);
      SpreadsheetApp.flush();
      Logger.log('cancelTradePayment: ' + jref + ' ' + pay.kind + ' ' + pay.amount);
      return { success: true, message: 'Pelunasan ' + (jref || '') + ' dibatalkan (' + jrows.length + ' baris jurnal dihapus). Transaksi kembali memiliki sisa ' + (isSale ? 'piutang' : 'utang') + '.' };
    });
  } catch (err) {
    Logger.log('cancelTradePayment ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ============================================================
// DATA PEMASOK / PETANI + PROGRES KEBUN
// ============================================================
//  Suppliers      : data pemasok/petani (kode PMS-001, kelompok tani, kontak, alamat, no. anggota bila ada)
//  Gardens        : kebun milik pemasok (luas, tahun tanam, populasi, bibit, status lahan, koordinat)
//  GardenProgress : catatan progres kebun per tanggal (tahap, persen, kegiatan, produksi)
// Tautan ke Pembelian TBS memakai NAMA pemasok (kolom "supplier" pembelian), tanpa mengubah data pembelian.
var SHEET_SUPPLIERS = 'Suppliers', SHEET_GARDENS = 'Gardens', SHEET_GARDEN_PROGRESS = 'GardenProgress';
var GARDEN_STAGES_ = ['Pembukaan Lahan', 'Penanaman', 'Perawatan TBM', 'Panen Perdana', 'TM Produktif', 'Replanting', 'Lainnya'];

// Cakupan unit data pemasok: null = Administrator (semua data, termasuk yang tanpa unit / global);
// selain itu daftar unit yang boleh dilihat. Role non-admin tanpa batas unit hanya melihat pemasok yang punya unit, bukan data tanpa unit.
function supplierScope_(token) {
  if (CacheService.getScriptCache().get('role_' + token) === 'Administrator') return null;
  var u = sessionUnits_(token);
  return u === null ? getUnitsList_() : u;
}
function inSupplierScope_(scope, unit) { return scope === null || scope.indexOf(String(unit || '')) >= 0; }
function sessionUsername_(token) { return CacheService.getScriptCache().get('sess_' + token) || ''; }
function supplierPerm_(token, both) { return checkPerm_(token, both ? ['suppliers', 'garden_progress'] : ['suppliers']); }
function numOrBlank_(v) { var n = parseFloat(String(v).replace(',', '.')); return isNaN(n) ? '' : n; }

function getSuppliers(token) {
  var deny_ = supplierPerm_(token, true); if (deny_) return deny_;
  try {
    var scope = supplierScope_(token);
    var supAll = sheetToObjects_(getTradeSheet_(SHEET_SUPPLIERS)), supUnit = {};
    supAll.forEach(function (r) { supUnit[String(r.supplier_id)] = String(r.unit || ''); });
    var gardens = sheetToObjects_(getTradeSheet_(SHEET_GARDENS)).filter(function (g) { return inSupplierScope_(scope, supUnit[String(g.supplier_id)]); }), gBy = {};
    gardens.forEach(function (g) {
      var k = String(g.supplier_id), o = gBy[k] = gBy[k] || { n: 0, area: 0 };
      o.n++; o.area += Number(g.area_ha) || 0;
    });
    // Total pembelian TBS per nama pemasok
    var buy = {};
    var qtyBy = {};
    // Pembelian lama (tanpa supplier_id) dicocokkan lewat nama hanya bila namanya milik tepat satu pemasok.
    var nameCount = {}, nameToId = {};
    supAll.forEach(function (r) { var k = String(r.name || '').trim().toLowerCase(); nameCount[k] = (nameCount[k] || 0) + 1; nameToId[k] = String(r.supplier_id); });
    sheetToObjects_(getTradeSheet_(SHEET_PURCHASE_LINES)).forEach(function (l) { var k = String(l.purchase_id); qtyBy[k] = (qtyBy[k] || 0) + (Number(l.qty) || 0); });
    sheetToObjects_(getTradeSheet_(SHEET_PURCHASES)).forEach(function (p) {
      if (kindOf_(p.kind) !== 'TBS' || !inSupplierScope_(scope, p.unit)) return;
      var nm = String(p.supplier || '').trim().toLowerCase();
      var sid = String(p.supplier_id || '') || (nameCount[nm] === 1 ? nameToId[nm] : '');
      if (!sid) return;
      var o = buy[sid] = buy[sid] || { kg: 0, total: 0, last: '' };
      o.kg += qtyBy[String(p.purchase_id)] || 0; o.total += Number(p.total) || 0;
      var d = fmtDate_(p.date); if (d > o.last) o.last = d;
    });
    var data = supAll.filter(function (r) { return inSupplierScope_(scope, r.unit); }).map(function (r) {
      var g = gBy[String(r.supplier_id)] || { n: 0, area: 0 }, b = buy[String(r.supplier_id)] || { kg: 0, total: 0, last: '' };
      return {
        supplier_id: String(r.supplier_id), code: String(r.code || ''), name: String(r.name || ''), group: String(r.group || ''),
        phone: String(r.phone || ''), address: String(r.address || ''), village: String(r.village || ''), member_no: String(r.member_no || ''),
        status: r.status || 'Aktif', join_date: fmtDate_(r.join_date), note: String(r.note || ''), unit: String(r.unit || ''),
        gardens: g.n, area_ha: Math.round(g.area * 10000) / 10000, buy_kg: Math.round(b.kg * 100) / 100, buy_total: b.total, last_buy: b.last
      };
    }).sort(function (a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
    return { success: true, data: data };
  } catch (err) {
    Logger.log('getSuppliers ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** p = { supplier_id?, name, group, phone, address, village, member_no, status, join_date, note } */
function saveSupplier(token, p) {
  if (p && p.supplier_id) { var adm_ = checkAdminAction_(token); if (adm_) return adm_; }
  var deny_ = supplierPerm_(token, false); if (deny_) return deny_;
  try {
    var name = String((p && p.name) || '').trim();
    if (!name) return { success: false, message: 'Nama pemasok/petani wajib diisi.' };
    if (name.length > 80) return { success: false, message: 'Nama terlalu panjang (maks 80 karakter).' };
    var join = String(p.join_date || '').trim();
    if (join && (!/^\d{4}-\d{2}-\d{2}$/.test(join) || isNaN(new Date(join).getTime()))) return { success: false, message: 'Tanggal bergabung tidak valid.' };
    var status = p.status === 'Nonaktif' ? 'Nonaktif' : 'Aktif';
    var f = function (v, n) { return String(v || '').trim().substring(0, n); };
    var uc = checkUnit_(p.unit); if (uc.error) return { success: false, message: uc.error };
    var scope = supplierScope_(token);
    if (scope !== null && !uc.unit) return { success: false, message: 'Pilih unit koperasi pemasok ini.' };
    if (!inSupplierScope_(scope, uc.unit)) return { success: false, forbidden: true, message: 'Anda tidak memiliki akses ke unit "' + uc.unit + '".' };
    return withLock_(function () {
      var sheet = getTradeSheet_(SHEET_SUPPLIERS);
      var rows = sheetToObjects_(sheet);
      // Nama boleh kembar, asalkan unit dan desa/dusunnya berbeda (kode PMS-xxx yang membedakan di sistem).
      var vil = f(p.village, 60).toLowerCase();
      var dup = rows.some(function (r) {
        return String(r.name).trim().toLowerCase() === name.toLowerCase() && String(r.supplier_id) !== String(p.supplier_id || '') &&
          String(r.unit || '') === uc.unit && String(r.village || '').trim().toLowerCase() === vil;
      });
      if (dup) return { success: false, message: 'Pemasok "' + name + '" dengan unit dan desa yang sama sudah ada. Isi desa/dusun yang berbeda, atau tambahkan keterangan pada nama.' };
      if (p.supplier_id) {
        var ex = rows.filter(function (r) { return String(r.supplier_id) === String(p.supplier_id); })[0];
        if (!ex) return { success: false, message: 'Pemasok tidak ditemukan.' };
        if (String(ex.name).trim().toLowerCase() !== name.toLowerCase()) linkLegacyPurchases_(rows, ex);
        sheet.getRange(ex._row, 3, 1, 9).setValues([[name, f(p.group, 60), f(p.phone, 30), f(p.address, 200), f(p.village, 60), f(p.member_no, 30), status, join, f(p.note, 300)]]);
        sheet.getRange(ex._row, 13).setValue(uc.unit);
        SpreadsheetApp.flush();
        return { success: true, message: 'Data pemasok "' + name + '" diperbarui.' };
      }
      // Nama baru yang sama dengan pemasok lama: tautkan dulu pembelian lama (hanya bernama) ke pemasok yang sudah ada, supaya tidak tercampur.
      var same = rows.filter(function (r) { return String(r.name).trim().toLowerCase() === name.toLowerCase(); });
      if (same.length === 1) linkLegacyPurchases_(rows, same[0]);
      var max = 0;
      rows.forEach(function (r) { var n = parseInt(String(r.code).replace(/\D/g, ''), 10); if (!isNaN(n) && n > max) max = n; });
      var code = 'PMS-' + ('000' + (max + 1)).slice(-3);
      appendTradeRows_(SHEET_SUPPLIERS, [[Utilities.getUuid(), code, name, f(p.group, 60), f(p.phone, 30), f(p.address, 200), f(p.village, 60), f(p.member_no, 30), status, join, f(p.note, 300), new Date(), uc.unit]]);
      SpreadsheetApp.flush();
      return { success: true, message: 'Pemasok "' + name + '" ditambahkan dengan kode ' + code + '.' };
    });
  } catch (err) {
    Logger.log('saveSupplier ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deleteSupplier(token, supplierId) {
  var deny_ = supplierPerm_(token, false); if (deny_) return deny_;
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    return withLock_(function () {
      var sheet = getTradeSheet_(SHEET_SUPPLIERS);
      var ex = sheetToObjects_(sheet).filter(function (r) { return String(r.supplier_id) === String(supplierId); })[0];
      if (!ex) return { success: false, message: 'Pemasok tidak ditemukan.' };
      var n = sheetToObjects_(getTradeSheet_(SHEET_GARDENS)).filter(function (g) { return String(g.supplier_id) === String(supplierId); }).length;
      if (n) return { success: false, message: 'Pemasok ini masih punya ' + n + ' kebun. Hapus kebunnya dulu (atau ubah statusnya menjadi Nonaktif).' };
      sheet.deleteRow(ex._row);
      SpreadsheetApp.flush();
      return { success: true, message: 'Pemasok "' + ex.name + '" dihapus.' };
    });
  } catch (err) {
    Logger.log('deleteSupplier ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** Isi supplier_id pada pembelian lama (hanya bernama) yang namanya sama dengan pemasok `sup`, bila namanya hanya milik pemasok itu. */
function linkLegacyPurchases_(allSuppliers, sup) {
  var nm = String(sup.name || '').trim().toLowerCase();
  var owners = allSuppliers.filter(function (r) { return String(r.name || '').trim().toLowerCase() === nm; });
  if (owners.length !== 1) return;
  var sheet = getTradeSheet_(SHEET_PURCHASES);
  sheetToObjects_(sheet).forEach(function (r) {
    if (!String(r.supplier_id || '') && String(r.supplier || '').trim().toLowerCase() === nm) {
      sheet.getRange(r._row, 14).setNumberFormat('@').setValue(String(sup.supplier_id));
    }
  });
}

// ---------- KEBUN ----------
/** Daftar tahun tanam kebun. Data lama (satu tahun tanam di kolom lama) dibaca sebagai satu baris. */
function plantingsOf_(g) {
  try {
    var arr = JSON.parse(String(g.plantings || ''));
    if (Array.isArray(arr)) return arr.map(function (x) { return { year: x.year === '' ? '' : Number(x.year), area_ha: Number(x.area_ha) || 0, trees: Number(x.trees) || 0 }; });
  } catch (e) { /* data lama */ }
  var y = g.planting_year === '' ? '' : Number(g.planting_year) || '', a = Number(g.area_ha) || 0, t = Number(g.trees) || 0;
  return (y === '' && !a && !t) ? [] : [{ year: y, area_ha: a, trees: t }];
}

function getGardens(token) {
  var deny_ = supplierPerm_(token, true); if (deny_) return deny_;
  try {
    var scope = supplierScope_(token);
    var sup = {};
    sheetToObjects_(getTradeSheet_(SHEET_SUPPLIERS)).forEach(function (r) { sup[String(r.supplier_id)] = r; });
    var last = {}, cnt = {}, prod = {}, monthly = {};
    sheetToObjects_(getTradeSheet_(SHEET_GARDEN_PROGRESS)).forEach(function (r) {
      var k = String(r.garden_id), d = fmtDate_(r.date);
      cnt[k] = (cnt[k] || 0) + 1; prod[k] = (prod[k] || 0) + (Number(r.production_kg) || 0);
      if (Number(r.production_kg) > 0) { var mo = d.substring(0, 7); var mm = monthly[k] = monthly[k] || {}; mm[mo] = Math.round(((mm[mo] || 0) + Number(r.production_kg)) * 100) / 100; }
      if (!last[k] || d > last[k].date || (d === last[k].date && String(r.timestamp) > String(last[k].ts))) {
        last[k] = { date: d, stage: String(r.stage || ''), pct: Number(r.progress_pct) || 0, activity: String(r.activity || ''), ts: String(r.timestamp || '') };
      }
    });
    var data = sheetToObjects_(getTradeSheet_(SHEET_GARDENS)).filter(function (g) { var sp0 = sup[String(g.supplier_id)]; return inSupplierScope_(scope, sp0 ? sp0.unit : ''); }).map(function (g) {
      var k = String(g.garden_id), sp = sup[String(g.supplier_id)] || {}, l = last[k] || null;
      return {
        garden_id: k, supplier_id: String(g.supplier_id), supplier_name: String(sp.name || '(pemasok dihapus)'), supplier_code: String(sp.code || ''), group: String(sp.group || ''), unit: String(sp.unit || ''),
        name: String(g.name || ''), location: String(g.location || ''), area_ha: Number(g.area_ha) || 0, planting_year: g.planting_year === '' ? '' : Number(g.planting_year) || '', plantings: plantingsOf_(g),
        trees: Number(g.trees) || 0, variety: String(g.variety || ''), land_status: String(g.land_status || ''), lat: g.lat === '' ? '' : Number(g.lat), lng: g.lng === '' ? '' : Number(g.lng),
        status: g.status || 'Aktif', note: String(g.note || ''),
        last: l ? { date: l.date, stage: l.stage, pct: l.pct, activity: l.activity } : null, progress_count: cnt[k] || 0, production_kg: Math.round((prod[k] || 0) * 100) / 100, monthly: monthly[k] || {}
      };
    }).sort(function (a, b) { return (a.supplier_name + a.name).toLowerCase() < (b.supplier_name + b.name).toLowerCase() ? -1 : 1; });
    return { success: true, data: data, stages: GARDEN_STAGES_, yield_curve: getYieldCurve_() };
  } catch (err) {
    Logger.log('getGardens ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** p = { garden_id?, supplier_id, name, location, area_ha, planting_year, trees, variety, land_status, lat, lng, status, note } */
function saveGarden(token, p) {
  if (p && p.garden_id) { var adm_ = checkAdminAction_(token); if (adm_) return adm_; }
  var deny_ = supplierPerm_(token, true); if (deny_) return deny_;
  try {
    var name = String((p && p.name) || '').trim();
    if (!name) return { success: false, message: 'Nama kebun wajib diisi (mis. "Blok A" atau nama lokasinya).' };
    // Satu kebun boleh punya beberapa tahun tanam: plantings = [{year, area_ha, trees}] (tahun kosong = belum diketahui, maksimal satu).
    var thisYear = new Date().getFullYear();
    var plantings = p.plantings;
    if (!Array.isArray(plantings)) plantings = [{ year: p.planting_year, area_ha: p.area_ha, trees: p.trees }];
    var plClean = [], seenYears = {};
    for (var pi = 0; pi < plantings.length; pi++) {
      var pl = plantings[pi] || {};
      var yr = pl.year === '' || pl.year === undefined || pl.year === null ? '' : Math.round(Number(pl.year));
      var ar = Number(pl.area_ha) || 0, tr = Math.round(Number(pl.trees) || 0);
      if (yr === '' && ar === 0 && tr === 0) continue;   // baris kosong
      if (yr !== '' && (isNaN(yr) || yr < 1950 || yr > thisYear + 1)) return { success: false, message: 'Tahun tanam tidak valid: ' + pl.year + '.' };
      if (ar < 0 || tr < 0) return { success: false, message: 'Luas dan jumlah pohon tidak boleh negatif.' };
      if (seenYears[yr]) return { success: false, message: yr === '' ? 'Hanya satu baris boleh tanpa tahun tanam.' : 'Tahun tanam ' + yr + ' diisi dua kali. Gabungkan jadi satu baris.' };
      seenYears[yr] = true;
      plClean.push({ year: yr, area_ha: Math.round(ar * 10000) / 10000, trees: tr });
    }
    plClean.sort(function (a, b) { return (a.year === '' ? 9999 : a.year) - (b.year === '' ? 9999 : b.year); });
    var area = Math.round(plClean.reduce(function (t, x) { return t + x.area_ha; }, 0) * 10000) / 10000;
    var trees = plClean.reduce(function (t, x) { return t + x.trees; }, 0);
    var years = plClean.filter(function (x) { return x.year !== ''; }).map(function (x) { return x.year; });
    var year = years.length ? Math.min.apply(null, years) : '';
    var lat = numOrBlank_(p.lat), lng = numOrBlank_(p.lng);
    if ((lat !== '' && (lat < -90 || lat > 90)) || (lng !== '' && (lng < -180 || lng > 180))) return { success: false, message: 'Koordinat tidak valid.' };
    var f = function (v, n) { return String(v || '').trim().substring(0, n); };
    var status = p.status === 'Nonaktif' ? 'Nonaktif' : 'Aktif';
    return withLock_(function () {
      var sup = sheetToObjects_(getTradeSheet_(SHEET_SUPPLIERS)).filter(function (r) { return String(r.supplier_id) === String(p.supplier_id); })[0];
      if (!sup) return { success: false, message: 'Pilih pemasok/petani pemilik kebun.' };
      if (!inSupplierScope_(supplierScope_(token), sup.unit)) return forbiddenResponse_();
      var sheet = getTradeSheet_(SHEET_GARDENS);
      var rows = sheetToObjects_(sheet);
      var dup = rows.some(function (r) { return String(r.supplier_id) === String(p.supplier_id) && String(r.name).trim().toLowerCase() === name.toLowerCase() && String(r.garden_id) !== String(p.garden_id || ''); });
      if (dup) return { success: false, message: 'Pemasok ini sudah punya kebun bernama "' + name + '".' };
      var vals = [name, f(p.location, 200), area, year, trees, f(p.variety, 60), f(p.land_status, 60), lat, lng, status, f(p.note, 300)];
      if (p.garden_id) {
        var ex = rows.filter(function (r) { return String(r.garden_id) === String(p.garden_id); })[0];
        if (!ex) return { success: false, message: 'Kebun tidak ditemukan.' };
        sheet.getRange(ex._row, 2, 1, 1).setValue(String(p.supplier_id));
        sheet.getRange(ex._row, 3, 1, 11).setValues([vals]);
        sheet.getRange(ex._row, 15).setValue(JSON.stringify(plClean));
        SpreadsheetApp.flush();
        return { success: true, message: 'Kebun "' + name + '" diperbarui.' };
      }
      appendTradeRows_(SHEET_GARDENS, [[Utilities.getUuid(), String(p.supplier_id)].concat(vals, [new Date(), JSON.stringify(plClean)])]);
      SpreadsheetApp.flush();
      return { success: true, message: 'Kebun "' + name + '" milik ' + sup.name + ' ditambahkan.' };
    });
  } catch (err) {
    Logger.log('saveGarden ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deleteGarden(token, gardenId) {
  var deny_ = supplierPerm_(token, true); if (deny_) return deny_;
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    return withLock_(function () {
      var sheet = getTradeSheet_(SHEET_GARDENS);
      var ex = sheetToObjects_(sheet).filter(function (r) { return String(r.garden_id) === String(gardenId); })[0];
      if (!ex) return { success: false, message: 'Kebun tidak ditemukan.' };
      var pSheet = getTradeSheet_(SHEET_GARDEN_PROGRESS);
      var prog = sheetToObjects_(pSheet).filter(function (r) { return String(r.garden_id) === String(gardenId); });
      deleteRowsDesc_(pSheet, prog);
      sheet.deleteRow(ex._row);
      SpreadsheetApp.flush();
      return { success: true, message: 'Kebun "' + ex.name + '" dan ' + prog.length + ' catatan progresnya dihapus.' };
    });
  } catch (err) {
    Logger.log('deleteGarden ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ---------- PROGRES KEBUN ----------
function gardenInScope_(token, gardenId) {
  var g = sheetToObjects_(getTradeSheet_(SHEET_GARDENS)).filter(function (r) { return String(r.garden_id) === String(gardenId); })[0];
  if (!g) return true;   // tidak ada: biarkan fungsi pemanggil melapor "tidak ditemukan"
  var sp = sheetToObjects_(getTradeSheet_(SHEET_SUPPLIERS)).filter(function (r) { return String(r.supplier_id) === String(g.supplier_id); })[0];
  return inSupplierScope_(supplierScope_(token), sp ? sp.unit : '');
}
function getGardenProgress(token, gardenId) {
  var deny_ = supplierPerm_(token, true); if (deny_) return deny_;
  if (!gardenInScope_(token, gardenId)) return forbiddenResponse_();
  try {
    var data = sheetToObjects_(getTradeSheet_(SHEET_GARDEN_PROGRESS))
      .filter(function (r) { return String(r.garden_id) === String(gardenId); })
      .map(function (r) {
        return { progress_id: String(r.progress_id), date: fmtDate_(r.date), stage: String(r.stage || ''), progress_pct: Number(r.progress_pct) || 0, activity: String(r.activity || ''), production_kg: Number(r.production_kg) || 0, note: String(r.note || ''), by: String(r.by || ''), _ts: String(r.timestamp || '') };
      })
      .sort(function (a, b) { return a.date === b.date ? (a._ts < b._ts ? 1 : -1) : (a.date < b.date ? 1 : -1); });
    return { success: true, data: data };
  } catch (err) {
    Logger.log('getGardenProgress ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** p = { garden_id, date, stage, progress_pct (0-100), activity, production_kg, note } */
function saveGardenProgress(token, p) {
  var deny_ = checkPerm_(token, ['garden_progress']); if (deny_) return deny_;
  try {
    if (!p || !p.garden_id) return { success: false, message: 'Kebun belum dipilih.' };
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(p.date || '')) || isNaN(new Date(p.date).getTime())) return { success: false, message: 'Tanggal tidak valid.' };
    var stage = String(p.stage || '');
    if (GARDEN_STAGES_.indexOf(stage) < 0) return { success: false, message: 'Pilih tahap progres.' };
    var pct = Number(p.progress_pct);
    if (isNaN(pct) || pct < 0 || pct > 100) return { success: false, message: 'Persentase progres harus antara 0 dan 100.' };
    var prod = Number(p.production_kg) || 0;
    if (prod < 0) return { success: false, message: 'Produksi tidak boleh negatif.' };
    if (!gardenInScope_(token, p.garden_id)) return forbiddenResponse_();
    return withLock_(function () {
      var g = sheetToObjects_(getTradeSheet_(SHEET_GARDENS)).filter(function (r) { return String(r.garden_id) === String(p.garden_id); })[0];
      if (!g) return { success: false, message: 'Kebun tidak ditemukan.' };
      appendTradeRows_(SHEET_GARDEN_PROGRESS, [[Utilities.getUuid(), String(g.garden_id), String(g.supplier_id), p.date, stage, pct, String(p.activity || '').trim().substring(0, 300), prod, String(p.note || '').trim().substring(0, 300), sessionUsername_(token), new Date()]]);
      SpreadsheetApp.flush();
      return { success: true, message: 'Progres kebun "' + g.name + '" dicatat (' + stage + ', ' + pct + '%).' };
    });
  } catch (err) {
    Logger.log('saveGardenProgress ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function deleteGardenProgress(token, progressId) {
  var deny_ = checkPerm_(token, ['garden_progress']); if (deny_) return deny_;
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    return withLock_(function () {
      var sheet = getTradeSheet_(SHEET_GARDEN_PROGRESS);
      var ex = sheetToObjects_(sheet).filter(function (r) { return String(r.progress_id) === String(progressId); })[0];
      if (!ex) return { success: false, message: 'Catatan progres tidak ditemukan.' };
      sheet.deleteRow(ex._row);
      SpreadsheetApp.flush();
      return { success: true, message: 'Catatan progres dihapus.' };
    });
  } catch (err) {
    Logger.log('deleteGardenProgress ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

/** Nama pemasok aktif untuk isian cepat di form Pembelian TBS (boleh dibaca pemegang hak Pembelian TBS). */
function getSupplierNames(token) {
  var deny_ = checkPerm_(token, ['purchases_tbs', 'suppliers', 'garden_progress']); if (deny_) return deny_;
  try {
    var scope = supplierScope_(token);
    var names = sheetToObjects_(getTradeSheet_(SHEET_SUPPLIERS))
      .filter(function (r) { return (r.status || 'Aktif') === 'Aktif' && inSupplierScope_(scope, r.unit); })
      .map(function (r) { return String(r.name || ''); })
      .filter(Boolean).sort(function (a, b) { return a.toLowerCase() < b.toLowerCase() ? -1 : 1; });
    return { success: true, data: names };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/** Pilihan pemasok aktif untuk form Pembelian TBS (kode, nama, unit, desa), sesuai cakupan unit akun. */
function getSupplierOptions(token) {
  var deny_ = checkPerm_(token, ['purchases_tbs', 'suppliers', 'garden_progress']); if (deny_) return deny_;
  try {
    var scope = supplierScope_(token);
    var data = sheetToObjects_(getTradeSheet_(SHEET_SUPPLIERS))
      .filter(function (r) { return (r.status || 'Aktif') === 'Aktif' && inSupplierScope_(scope, r.unit); })
      .map(function (r) { return { supplier_id: String(r.supplier_id), code: String(r.code || ''), name: String(r.name || ''), unit: String(r.unit || ''), village: String(r.village || ''), group: String(r.group || '') }; })
      .sort(function (a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
    return { success: true, data: data };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ============================================================
// KURVA HASIL SAWIT (target & maksimum produksi menurut umur tanaman)
// ============================================================
// Satuan: ton TBS per hektar per tahun. Angka bawaan adalah acuan umum kebun sawit rakyat/Tenera dan HARUS disesuaikan
// dengan kondisi kebun setempat (tanah, bibit, perawatan, curah hujan). Target bawaan = 80% dari maksimum.
// Umur = tahun berjalan - tahun tanam. from/to = rentang umur (to kosong = seterusnya).
var DEFAULT_YIELD_CURVE_ = [
  { from: 0, to: 2, max: 0, target: 0 },
  { from: 3, to: 3, max: 10, target: 8 },
  { from: 4, to: 4, max: 17, target: 13.6 },
  { from: 5, to: 5, max: 22, target: 17.6 },
  { from: 6, to: 6, max: 25, target: 20 },
  { from: 7, to: 7, max: 27, target: 21.6 },
  { from: 8, to: 8, max: 28, target: 22.4 },
  { from: 9, to: 12, max: 30, target: 24 },
  { from: 13, to: 15, max: 28, target: 22.4 },
  { from: 16, to: 18, max: 26, target: 20.8 },
  { from: 19, to: 21, max: 24, target: 19.2 },
  { from: 22, to: 25, max: 21, target: 16.8 },
  { from: 26, to: null, max: 18, target: 14.4 }
];
function getYieldCurve_() {
  var raw = PropertiesService.getDocumentProperties().getProperty('yield_curve');
  if (raw) { try { var c = JSON.parse(raw); if (Array.isArray(c) && c.length) return c; } catch (e) { /* pakai bawaan */ } }
  return JSON.parse(JSON.stringify(DEFAULT_YIELD_CURVE_));
}
function getYieldSettings(token) {
  var deny_ = supplierPerm_(token, true); if (deny_) return deny_;
  return { success: true, data: getYieldCurve_(), defaults: DEFAULT_YIELD_CURVE_ };
}
/** curve = [{from, to|null, max, target}] ; hanya Administrator. Kirim kosong untuk mengembalikan ke bawaan. */
function saveYieldCurve(token, curve) {
  var adm_ = checkAdminAction_(token); if (adm_) return adm_;
  try {
    if (!curve || !curve.length) {
      PropertiesService.getDocumentProperties().deleteProperty('yield_curve');
      return { success: true, message: 'Kurva hasil dikembalikan ke angka bawaan.' };
    }
    var out = [], prevTo = -1;
    for (var i = 0; i < curve.length; i++) {
      var c = curve[i] || {};
      var from = Math.round(Number(c.from)), to = (c.to === null || c.to === '' || c.to === undefined) ? null : Math.round(Number(c.to));
      var max = Number(c.max), target = Number(c.target);
      if (isNaN(from) || from < 0 || (to !== null && (isNaN(to) || to < from))) return { success: false, message: 'Rentang umur baris ' + (i + 1) + ' tidak valid.' };
      if (from !== prevTo + 1) return { success: false, message: 'Rentang umur harus berurutan tanpa celah atau tumpang tindih (baris ' + (i + 1) + ' seharusnya mulai dari umur ' + (prevTo + 1) + ').' };
      if (isNaN(max) || isNaN(target) || max < 0 || target < 0 || max > 80 || target > 80) return { success: false, message: 'Hasil baris ' + (i + 1) + ' harus antara 0 dan 80 ton/ha/tahun.' };
      if (target > max) return { success: false, message: 'Target baris ' + (i + 1) + ' tidak boleh melebihi maksimum.' };
      out.push({ from: from, to: to, max: Math.round(max * 100) / 100, target: Math.round(target * 100) / 100 });
      if (to === null) { if (i !== curve.length - 1) return { success: false, message: 'Baris "seterusnya" harus yang terakhir.' }; prevTo = 1e9; } else prevTo = to;
    }
    if (out[0].from !== 0) return { success: false, message: 'Baris pertama harus mulai dari umur 0.' };
    if (out[out.length - 1].to !== null) return { success: false, message: 'Baris terakhir harus berlaku "seterusnya" (kosongkan umur akhir).' };
    PropertiesService.getDocumentProperties().setProperty('yield_curve', JSON.stringify(out));
    return { success: true, message: 'Kurva hasil disimpan.' };
  } catch (err) {
    Logger.log('saveYieldCurve ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}
// ============================================================
// IMPOR MASSAL DARI EXCEL (menu Impor / Ekspor Excel)
// ============================================================
/** items = [{ row, data }] ; kind: journal | saving | item | supplier | garden. Tiap data disimpan lewat fungsi simpan biasa (hak akses ikut diperiksa). */
function bulkSave(token, kind, items) {
  if (!requireSession_(token)) return sessionExpiredResponse_();
  var fns = { journal: saveJournal, saving: saveSaving, item: saveItem, supplier: saveSupplier, garden: saveGarden };
  var fn = fns[String(kind)];
  if (!fn) return { success: false, message: 'Jenis impor tidak dikenal.' };
  if (!items || !items.length) return { success: false, message: 'Tidak ada data untuk diimpor.' };
  if (items.length > 200) return { success: false, message: 'Maksimal 200 data per kiriman.' };
  var results = items.map(function (it) {
    var res;
    try { res = fn(token, (it && it.data) || {}); } catch (err) { res = { success: false, message: err.message }; }
    return { row: it && it.row, ok: !!(res && res.success), message: (res && res.message) || '' };
  });
  return { success: true, results: results };
}

/** Impor akun COA dari Excel. Kode yang sudah ada DILEWATI (tidak diubah). rows = [{ row, code, name, normal_balance, opening_debit, opening_credit }] */
function importCoaExcel(token, rows) {
  var deny_ = checkPerm_(token, ['coa']); if (deny_) return deny_;
  try {
    if (!rows || !rows.length) return { success: false, message: 'Tidak ada data untuk diimpor.' };
    if (rows.length > 2000) return { success: false, message: 'Maksimal 2000 baris per impor.' };
    return withLock_(function () {
      var sheet = getSheetOrThrow_(SHEET_COA);
      var have = {};
      sheetToObjects_(sheet).forEach(function (r) { have[String(r.account_code)] = true; });
      var out = [], skipped = [];
      rows.forEach(function (it, i) {
        var line = Number(it.row) || (i + 2);
        var code = String(it.code || '').trim(), name = String(it.name || '').trim();
        if (!code || !name) { skipped.push({ row: line, reason: 'Kode atau nama kosong' }); return; }
        if (!/^[1-5](\.\d+)*$/.test(code)) { skipped.push({ row: line, reason: 'Format kode tidak valid' }); return; }
        if (have[code]) { skipped.push({ row: line, reason: 'Kode ' + code + ' sudah ada' }); return; }
        var od = Number(it.opening_debit) || 0, oc = Number(it.opening_credit) || 0;
        if (od < 0 || oc < 0) { skipped.push({ row: line, reason: 'Saldo awal tidak boleh negatif' }); return; }
        var category = deriveCoaCategory_(code);
        var nb = (it.normal_balance === 'Debit' || it.normal_balance === 'Kredit') ? it.normal_balance : defaultNormalBalance_(category);
        have[code] = true;
        out.push([code, name, category, nb, 'Aktif', od, oc]);
      });
      if (out.length) {
        var start = nextRow_(sheet, out.length);
        sheet.getRange(start, 1, out.length, 7).setValues(out);
      }
      return { success: true, imported: out.length, skipped: skipped, message: out.length + ' akun diimpor' + (skipped.length ? ', ' + skipped.length + ' dilewati.' : '.') };
    });
  } catch (err) {
    Logger.log('importCoaExcel ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

return { allPermKeys_, allowedKinds_, appendJournalRows_, appendTradeRows_, bulkSave, calcInvoiceTax_, calculateShu, cancelTradePayment, changeMyPassword, checkAdminAction_, checkPerm_, checkTradePerm_, checkUnitAccess_, checkUnit_, cleanUnits_, coaCodeCompare_, comparePeriods_, computeClosingPlan_, computeSavingBalances_, coopProfileRaw_, coopSummary_, debugGetCOARaw, defaultNormalBalance_, deleteAccount, deleteEvidence, deleteGarden, deleteGardenProgress, deleteItem, deleteJournalEntry, deleteMember, deleteMemberAccount, deleteMemberTrans, deletePurchase, deleteRole, deleteRowsDesc_, deleteSale, deleteSaving, deleteStaffUser, deleteSupplier, deleteUsersByMember_, deriveCoaCategory_, ensureCoaColumns_, ensureCoopSheets_, ensureJournalUnitCol_, ensureTradeSheets_, ensureUsersColumns_, filterByUnits_, findCashAccount_, findLeafCoaByName_, fmtDate_, forbiddenResponse_, gardenInScope_, getArchiveFolder_, getAssets, getCOAList, getCashFlow, getClosedDate_, getClosingInfo, getClosingLogSheet_, getCompanySettings, getComparativeReport, getCoopProfile, getCoopSheet_, getDashboardSummary, getEquityChanges, getEvidenceFile, getEvidenceFolder_, getEvidenceIndex, getEvidenceList, getEvidenceSheet_, getFinancialStatements, getGardenProgress, getGardens, getItems, getJournals, getLedger, getMemberTrans, getMembers, getMyPerms, getMyPortal, getOrCreateSheet_, getPurchases, getRoles, getRolesConfig_, getSS_, getSales, getSavings, getSheetOrThrow_, getShuSettings, getStaffUsers, getSupplierNames, getSupplierOptions, getSuppliers, getTradePayments, getTradeSheet_, getUnitsList_, getYieldCurve_, getYieldSettings, hashPassword_, importCoaCSV, importCoaExcel, importMembers, inRange_, inSupplierScope_, isLeafCode_, jrow_, kindOf_, leafCoaRows_, linkLegacyPurchases_, loginUser, lookupTradeKind_, lookupTradeUnit_, memberAccountMap_, migrateCoaToHierarchical, nextCoaChildCode_, nextMemberNo_, nextRow_, notAdminResponse_, numOrBlank_, openingSeed_, payTrade, periodClosedCheck_, permsForRole_, planSummary_, plantingsOf_, postDepreciationJournal, postInvoiceTax, previewClosing, purgeEvidence_, q4_, readShuSettings_, renameAccount, requireAdmin_, requireAnySession_, requireMemberSession_, requireSession_, rewriteKeepRows_, runClosing, runDiagnostics, saleNet_, saveAccount, saveAsset, saveCompanySettings, saveCoopProfile, saveGarden, saveGardenProgress, saveItem, saveJournal, saveMember, saveMemberAccount, saveMemberTrans, savePurchase, saveRole, saveSale, saveSaving, saveShuSettings, saveStaffUser, saveSupplier, saveYieldCurve, sessionExpiredResponse_, sessionPerms_, sessionUnits_, sessionUsername_, setupCoopCoa, setupDatabase, sheetToObjects_, supplierPerm_, supplierScope_, taxRate_, tradeAccounts_, tradeRef_, tradeSheetFor_, tradeStatus_, unitsForRole_, updateSaldoAwalBulk, uploadEvidence, weighLine_, withLock_ };
}
