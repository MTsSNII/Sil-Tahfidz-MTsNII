/**
 * ============================================================================
 * SISTEM LAPORAN PERKEMBANGAN TAHFIDZ (SiL Tahfidz)
 * MTsS Nurul Islam Indonesia
 * Backend Engine: Google Apps Script + Google Sheets + Google Drive
 * ============================================================================
 */

const PROP_KEY_FOLDER_ID = 'SIL_TAHFIDZ_FOLDER_ID';
const PROP_KEY_MASTER_DB_ID = 'SIL_TAHFIDZ_MASTER_DB_ID';
const PROP_KEY_ACTIVE_TA = 'SIL_TAHFIDZ_ACTIVE_TA';

/**
 * Entry point Web App
 */
function doGet(e) {
  ensureSystemInitialized();
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('SiL Tahfidz - MTsS Nurul Islam Indonesia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fungsi Inisialisasi Database Otomatis (Dijalankan pertama kali & saat ganti Tahun Ajaran)
 */
function initDatabase(customTahunAjaran) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    let folderId = props.getProperty(PROP_KEY_FOLDER_ID);
    let folder;

    if (folderId) {
      try {
        folder = DriveApp.getFolderById(folderId);
      } catch (err) {
        folder = null;
      }
    }

    if (!folder) {
      const folders = DriveApp.getFoldersByName('SiL_Tahfidz_Database');
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder('SiL_Tahfidz_Database');
        // Lindungi folder dari akses publik
        folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      }
      folderId = folder.getId();
      props.setProperty(PROP_KEY_FOLDER_ID, folderId);
    }

    const defaultTA = customTahunAjaran || getDefaultTahunAjaran();
    let activeTA = props.getProperty(PROP_KEY_ACTIVE_TA);
    if (!activeTA || customTahunAjaran) {
      activeTA = defaultTA;
      props.setProperty(PROP_KEY_ACTIVE_TA, activeTA);
    }

    // Cek Master DB (Menyimpan Admin, Konfigurasi Global, Guru)
    let masterDbId = props.getProperty(PROP_KEY_MASTER_DB_ID);
    let masterSs = null;
    if (masterDbId) {
      try {
        masterSs = SpreadsheetApp.openById(masterDbId);
      } catch (e) {
        masterSs = null;
      }
    }

    if (!masterSs) {
      masterSs = SpreadsheetApp.create('SiL_Tahfidz_MASTER_CONFIG');
      const file = DriveApp.getFileById(masterSs.getId());
      file.moveTo(folder);
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      masterDbId = masterSs.getId();
      props.setProperty(PROP_KEY_MASTER_DB_ID, masterDbId);

      setupMasterSheets(masterSs, activeTA, folderId);
    }

    // Buat/Pastikan Database Akademik untuk Tahun Ajaran Aktif
    const taDbKey = 'SIL_DB_TA_' + activeTA.replace(/[^a-zA-Z0-9]/g, '_');
    let taDbId = props.getProperty(taDbKey);
    let taSs = null;
    if (taDbId) {
      try {
        taSs = SpreadsheetApp.openById(taDbId);
      } catch (e) {
        taSs = null;
      }
    }

    if (!taSs) {
      taSs = SpreadsheetApp.create('DB_SiL_Tahfidz_TA_' + activeTA.replace(/\//g, '-'));
      const file = DriveApp.getFileById(taSs.getId());
      file.moveTo(folder);
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      taDbId = taSs.getId();
      props.setProperty(taDbKey, taDbId);

      setupAcademicSheets(taSs);
    }

    // Update Konfigurasi Spreadsheet ID aktif
    updateConfigValue(masterSs, 'ACTIVE_TA', activeTA);
    updateConfigValue(masterSs, 'ACTIVE_TA_DB_ID', taDbId);
    updateConfigValue(masterSs, 'FOLDER_ID', folderId);

    return {
      success: true,
      message: 'Database berhasil diinisialisasi untuk Tahun Ajaran ' + activeTA,
      activeTA: activeTA
    };
  } catch (error) {
    return { success: false, message: 'Gagal inisialisasi database: ' + error.message };
  } finally {
    lock.releaseLock();
  }
}

function ensureSystemInitialized() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(PROP_KEY_MASTER_DB_ID)) {
    initDatabase();
  }
}

function getDefaultTahunAjaran() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  return month >= 7 ? `${year}/${year + 1}` : `${year - 1}/${year}`;
}

function setupMasterSheets(ss, activeTA, folderId) {
  const defaultSheet = ss.getSheets()[0];
  defaultSheet.setName('KONFIGURASI');
  defaultSheet.appendRow(['Key', 'Value']);
  defaultSheet.appendRow(['APP_NAME', 'SiL Tahfidz']);
  defaultSheet.appendRow(['APP_SUBTITLE', 'Sistem Laporan Perkembangan Tahfidz Siswa MTsS Nurul Islam Indonesia']);
  defaultSheet.appendRow(['LANGUAGE', 'ID']);
  defaultSheet.appendRow(['SCHOOL_LOGO', '']);
  defaultSheet.appendRow(['ACTIVE_TA', activeTA]);
  defaultSheet.appendRow(['FOLDER_ID', folderId]);
  defaultSheet.appendRow(['ACTIVE_TA_DB_ID', '']);
  formatHeader(defaultSheet);

  const adminSheet = ss.insertSheet('ADMIN');
  adminSheet.appendRow(['ID_Admin', 'Nama', 'Email', 'Password', 'Reset_Token', 'Token_Expiry', 'Created_At']);
  formatHeader(adminSheet);

  const guruSheet = ss.insertSheet('DATA_GURU');
  guruSheet.appendRow(['ID_Guru', 'Nama_Guru', 'Username', 'Password', 'Created_At']);
  formatHeader(guruSheet);
}

function setupAcademicSheets(ss) {
  const sheetsConfig = [
    { name: 'DATA_MADRASAH', headers: ['ID_Madrasah', 'Kelas', 'Semester', 'Tahun_Ajaran', 'Nama_Madrasah', 'Nama_Guru', 'Nama_Kepala', 'Tanggal_Rapor'] },
    { name: 'DATA_SISWA_VII', headers: ['ID_Siswa', 'No', 'Nama_Siswa', 'NIS_NISN', 'TTL', 'Alamat', 'Nama_Orangtua', 'Semester', 'Tahun_Ajaran'] },
    { name: 'DATA_SISWA_VIII', headers: ['ID_Siswa', 'No', 'Nama_Siswa', 'NIS_NISN', 'TTL', 'Alamat', 'Nama_Orangtua', 'Semester', 'Tahun_Ajaran'] },
    { name: 'DATA_SISWA_IX', headers: ['ID_Siswa', 'No', 'Nama_Siswa', 'NIS_NISN', 'TTL', 'Alamat', 'Nama_Orangtua', 'Semester', 'Tahun_Ajaran'] },
    { name: 'DAFTAR_SURAH', headers: ['ID_Surah', 'No', 'Nama_Surah', 'Kelas', 'Semester', 'Tahun_Ajaran'] },
    { name: 'PROGRES_VII', headers: ['ID_Progres', 'ID_Siswa', 'Semester', 'Tahun_Ajaran', 'Target_Surah', 'Jumlah_Tercapai', 'Persentase', 'Catatan_Guru', 'Detail_JSON'] },
    { name: 'PROGRES_VIII', headers: ['ID_Progres', 'ID_Siswa', 'Semester', 'Tahun_Ajaran', 'Target_Surah', 'Jumlah_Tercapai', 'Persentase', 'Catatan_Guru', 'Detail_JSON'] },
    { name: 'PROGRES_IX', headers: ['ID_Progres', 'ID_Siswa', 'Semester', 'Tahun_Ajaran', 'Target_Surah', 'Jumlah_Tercapai', 'Persentase', 'Catatan_Guru', 'Detail_JSON'] },
    { name: 'DATA_RAPOR', headers: ['ID_Rapor', 'ID_Siswa', 'Kelas', 'Semester', 'Tahun_Ajaran', 'Generated_At', 'Status'] }
  ];

  const firstSheet = ss.getSheets()[0];
  sheetsConfig.forEach((cfg, idx) => {
    const sh = idx === 0 ? firstSheet : ss.insertSheet(cfg.name);
    sh.setName(cfg.name);
    sh.appendRow(cfg.headers);
    formatHeader(sh);
  });
}

function formatHeader(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    const range = sheet.getRange(1, 1, 1, lastCol);
    range.setBackground('#2D4A3E').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function getMasterSpreadsheet() {
  ensureSystemInitialized();
  const id = PropertiesService.getScriptProperties().getProperty(PROP_KEY_MASTER_DB_ID);
  return SpreadsheetApp.openById(id);
}

function getAcademicSpreadsheet(tahunAjaran) {
  ensureSystemInitialized();
  const props = PropertiesService.getScriptProperties();
  const ta = tahunAjaran || props.getProperty(PROP_KEY_ACTIVE_TA);
  const taDbKey = 'SIL_DB_TA_' + ta.replace(/[^a-zA-Z0-9]/g, '_');
  let taDbId = props.getProperty(taDbKey);

  if (!taDbId) {
    initDatabase(ta);
    taDbId = props.getProperty(taDbKey);
  }
  return SpreadsheetApp.openById(taDbId);
}

function updateConfigValue(masterSs, key, value) {
  const sheet = masterSs.getSheetByName('KONFIGURASI');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function getAppConfig() {
  try {
    const ss = getMasterSpreadsheet();
    const sheet = ss.getSheetByName('KONFIGURASI');
    const data = sheet.getDataRange().getValues();
    const cfg = {};
    for (let i = 1; i < data.length; i++) {
      cfg[data[i][0]] = data[i][1];
    }
    const adminSheet = ss.getSheetByName('ADMIN');
    cfg.HAS_ADMIN = adminSheet.getLastRow() > 1;
    return { success: true, config: cfg };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * ============================================================================
 * MANAJEMEN SESI & AUTENTIKASI AMAN
 * ============================================================================
 */
function createSessionToken(role, userData) {
  const token = Utilities.getUuid();
  const cache = CacheService.getScriptCache();
  const sessionPayload = JSON.stringify({
    role: role,
    user: userData,
    createdAt: new Date().getTime()
  });
  // Sesi berlaku 6 jam (21600 detik)
  cache.put('SESSION_' + token, sessionPayload, 21600);
  return token;
}

function validateSession(token, requiredRole) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get('SESSION_' + token);
  if (!raw) return null;
  const session = JSON.parse(raw);
  if (requiredRole && session.role !== requiredRole) return null;
  return session;
}

function logoutSession(token) {
  if (token) {
    CacheService.getScriptCache().remove('SESSION_' + token);
  }
  return { success: true };
}

/**
 * ============================================================================
 * MODUL ADMIN: DAFTAR, LOGIN, LUPA PASSWORD, PENGATURAN
 * ============================================================================
 */
function registerAdmin(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const { nama, email, confirmEmail, password, confirmPassword } = payload;
    if (!nama || !email || !password) {
      return { success: false, message: 'Semua kolom wajib diisi.' };
    }
    if (email.trim().toLowerCase() !== confirmEmail.trim().toLowerCase()) {
      return { success: false, message: 'Konfirmasi Email tidak cocok.' };
    }
    if (password !== confirmPassword) {
      return { success: false, message: 'Konfirmasi Password tidak cocok.' };
    }
    if (password.length < 4) {
      return { success: false, message: 'Password minimal 4 karakter.' };
    }

    const ss = getMasterSpreadsheet();
    const sheet = ss.getSheetByName('ADMIN');
    const data = sheet.getDataRange().getValues();
    const cleanEmail = email.trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === cleanEmail) {
        return { success: false, message: 'Email sudah terdaftar sebagai Admin.' };
      }
    }

    const idAdmin = 'ADM-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    sheet.appendRow([idAdmin, nama.trim(), cleanEmail, password, '', '', new Date().toISOString()]);

    return { success: true, message: 'Pendaftaran Admin berhasil! Silakan masuk menggunakan email dan password Anda.' };
  } catch (e) {
    return { success: false, message: 'Error pendaftaran: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function loginAdmin(email, password) {
  try {
    if (!email || !password) {
      return { success: false, message: 'Username/Email dan Password wajib diisi.' };
    }
    const ss = getMasterSpreadsheet();
    const sheet = ss.getSheetByName('ADMIN');
    const data = sheet.getDataRange().getValues();
    const cleanEmail = email.trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === cleanEmail && String(data[i][3]) === String(password)) {
        const userObj = {
          id: data[i][0],
          nama: data[i][1],
          email: data[i][2]
        };
        const token = createSessionToken('ADMIN', userObj);
        return { success: true, token: token, user: userObj };
      }
    }
    return { success: false, message: 'Email atau Password Admin salah.' };
  } catch (e) {
    return { success: false, message: 'Gagal login: ' + e.message };
  }
}

function requestPasswordReset(email) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!email) return { success: false, message: 'Masukkan email Admin yang terdaftar.' };
    const ss = getMasterSpreadsheet();
    const sheet = ss.getSheetByName('ADMIN');
    const data = sheet.getDataRange().getValues();
    const cleanEmail = email.trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === cleanEmail) {
        const resetToken = Math.floor(100000 + Math.random() * 900000).toString();
        const expiry = new Date(new Date().getTime() + 30 * 60000).toISOString(); // 30 menit
        sheet.getRange(i + 1, 5).setValue(resetToken);
        sheet.getRange(i + 1, 6).setValue(expiry);

        const appUrl = ScriptApp.getService().getUrl();
        const subject = 'Reset Password Admin - SiL Tahfidz MTsS Nurul Islam';
        const body = `Assalamu'alaikum ${data[i][1]},\n\n` +
          `Kami menerima permintaan reset password untuk akun Admin SiL Tahfidz Anda.\n` +
          `Kode Token Reset Password Anda adalah: ${resetToken}\n\n` +
          `Tautan Aplikasi: ${appUrl}\n` +
          `Masukkan kode token di atas pada halaman Reset Password di aplikasi (berlaku 30 menit).\n\n` +
          `Wassalamu'alaikum Wr. Wb.\nTim IT MTsS Nurul Islam Indonesia`;

        MailApp.sendEmail(cleanEmail, subject, body);
        return { success: true, message: 'Tautan dan Kode Reset Password telah dikirim ke email ' + cleanEmail };
      }
    }
    return { success: false, message: 'Email tidak ditemukan dalam database Admin.' };
  } catch (e) {
    return { success: false, message: 'Gagal mengirim email reset: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function confirmPasswordReset(email, token, newPassword) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!email || !token || !newPassword) {
      return { success: false, message: 'Semua data reset wajib diisi.' };
    }
    const ss = getMasterSpreadsheet();
    const sheet = ss.getSheetByName('ADMIN');
    const data = sheet.getDataRange().getValues();
    const cleanEmail = email.trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === cleanEmail) {
        const savedToken = String(data[i][4]);
        const expiry = new Date(data[i][5]);
        if (savedToken !== String(token).trim()) {
          return { success: false, message: 'Kode Token Reset salah.' };
        }
        if (new Date() > expiry) {
          return { success: false, message: 'Kode Token sudah kadaluarsa. Silakan minta ulang.' };
        }
        sheet.getRange(i + 1, 4).setValue(newPassword);
        sheet.getRange(i + 1, 5).setValue('');
        sheet.getRange(i + 1, 6).setValue('');
        return { success: true, message: 'Password berhasil diperbarui! Silakan login.' };
      }
    }
    return { success: false, message: 'Akun tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function saveAdminSettings(token, settingsPayload) {
  const session = validateSession(token, 'ADMIN');
  if (!session) return { success: false, message: 'Sesi habis. Silakan login ulang.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getMasterSpreadsheet();
    if (settingsPayload.appName !== undefined) {
      updateConfigValue(ss, 'APP_NAME', settingsPayload.appName);
    }
    if (settingsPayload.language !== undefined) {
      updateConfigValue(ss, 'LANGUAGE', settingsPayload.language);
    }
    if (settingsPayload.schoolLogo !== undefined) {
      updateConfigValue(ss, 'SCHOOL_LOGO', settingsPayload.schoolLogo);
    }
    if (settingsPayload.tahunAjaran && settingsPayload.tahunAjaran !== settingsPayload.oldTahunAjaran) {
      // Ketentuan #20: Buat ulang database ketika tahun ajaran berubah
      initDatabase(settingsPayload.tahunAjaran);
    }

    // Update Email / Password Admin jika diisi
    if (settingsPayload.newEmail || settingsPayload.newPassword) {
      const adminSheet = ss.getSheetByName('ADMIN');
      const data = adminSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === session.user.id) {
          if (settingsPayload.newEmail) {
            adminSheet.getRange(i + 1, 3).setValue(settingsPayload.newEmail.trim().toLowerCase());
          }
          if (settingsPayload.newPassword) {
            adminSheet.getRange(i + 1, 4).setValue(settingsPayload.newPassword);
          }
          break;
        }
      }
    }

    return { success: true, message: 'Pengaturan berhasil disimpan!' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function resetAdminSettings(token) {
  const session = validateSession(token, 'ADMIN');
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getMasterSpreadsheet();
    updateConfigValue(ss, 'APP_NAME', 'SiL Tahfidz');
    updateConfigValue(ss, 'LANGUAGE', 'ID');
    updateConfigValue(ss, 'SCHOOL_LOGO', '');
    return { success: true, message: 'Pengaturan aplikasi telah direset ke default.' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * MODUL DATA GURU PENGUJI (ADMIN & LOGIN GURU)
 * ============================================================================
 */
function getGuruList(token) {
  const session = validateSession(token, 'ADMIN');
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const ss = getMasterSpreadsheet();
  const sheet = ss.getSheetByName('DATA_GURU');
  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    list.push({
      id: data[i][0],
      nama: data[i][1],
      username: data[i][2],
      hasPassword: Boolean(data[i][3]) // Jangan kirim plain password ke frontend tanpa kebutuhan
    });
  }
  return { success: true, data: list };
}

function saveGuru(token, payload) {
  const session = validateSession(token, 'ADMIN');
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const { id, nama, username, password } = payload;
    if (!nama || !username || (!id && !password)) {
      return { success: false, message: 'Nama, Username, dan Password Guru wajib diisi.' };
    }

    const ss = getMasterSpreadsheet();
    const sheet = ss.getSheetByName('DATA_GURU');
    const data = sheet.getDataRange().getValues();
    const cleanUser = username.trim();

    // Cek duplikasi username
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === cleanUser.toLowerCase() && data[i][0] !== id) {
        return { success: false, message: 'Username Guru sudah digunakan.' };
      }
    }

    if (id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === id) {
          sheet.getRange(i + 1, 2).setValue(nama.trim());
          sheet.getRange(i + 1, 3).setValue(cleanUser);
          if (password) sheet.getRange(i + 1, 4).setValue(password);
          return { success: true, message: 'Data Guru berhasil diperbarui.' };
        }
      }
    }

    const newId = 'GRU-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    sheet.appendRow([newId, nama.trim(), cleanUser, password, new Date().toISOString()]);
    return { success: true, message: 'Data Guru Penguji berhasil ditambahkan.' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteGuru(token, idGuru) {
  const session = validateSession(token, 'ADMIN');
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getMasterSpreadsheet();
    const sheet = ss.getSheetByName('DATA_GURU');
    const data = sheet.getDataRange().getValues();

    if (idGuru === 'ALL') {
      if (sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
      return { success: true, message: 'Seluruh data Guru Penguji telah direset.' };
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idGuru) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Data Guru berhasil dihapus.' };
      }
    }
    return { success: false, message: 'Data Guru tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function loginGuru(username, password, kelas, semester) {
  try {
    if (!username || !password) {
      return { success: false, message: 'Username dan Password wajib diisi.' };
    }
    const ss = getMasterSpreadsheet();
    const sheet = ss.getSheetByName('DATA_GURU');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === username.trim().toLowerCase() && String(data[i][3]) === String(password)) {
        const userObj = {
          id: data[i][0],
          nama: data[i][1],
          username: data[i][2],
          kelas: kelas,
          semester: semester
        };
        const token = createSessionToken('GURU', userObj);
        return { success: true, token: token, user: userObj };
      }
    }
    return { success: false, message: 'Username atau Password Guru Penguji salah!' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * ============================================================================
 * MODUL DATA MADRASAH
 * ============================================================================
 */
function getDataMadrasah(token, kelas, semester, tahunAjaran) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const ss = getAcademicSpreadsheet(tahunAjaran);
  const sheet = ss.getSheetByName('DATA_MADRASAH');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(kelas) && String(data[i][2]) === String(semester)) {
      return {
        success: true,
        data: {
          id: data[i][0],
          kelas: data[i][1],
          semester: data[i][2],
          tahunAjaran: data[i][3],
          namaMadrasah: data[i][4],
          namaGuru: data[i][5],
          namaKepala: data[i][6],
          tanggalRapor: data[i][7]
        }
      };
    }
  }
  return {
    success: true,
    data: {
      kelas: kelas,
      semester: semester,
      tahunAjaran: tahunAjaran,
      namaMadrasah: 'MTsS Nurul Islam Indonesia',
      namaGuru: session.user.nama || '',
      namaKepala: '',
      tanggalRapor: ''
    }
  };
}

function saveDataMadrasah(token, payload) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getAcademicSpreadsheet(payload.tahunAjaran);
    const sheet = ss.getSheetByName('DATA_MADRASAH');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(payload.kelas) && String(data[i][2]) === String(payload.semester)) {
        sheet.getRange(i + 1, 4, 1, 5).setValues([[
          payload.tahunAjaran,
          payload.namaMadrasah,
          payload.namaGuru,
          payload.namaKepala,
          payload.tanggalRapor
        ]]);
        return { success: true, message: 'Data Madrasah berhasil disimpan.' };
      }
    }

    const id = 'MDR-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    sheet.appendRow([
      id,
      payload.kelas,
      payload.semester,
      payload.tahunAjaran,
      payload.namaMadrasah,
      payload.namaGuru,
      payload.namaKepala,
      payload.tanggalRapor
    ]);
    return { success: true, message: 'Data Madrasah berhasil disimpan.' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function resetDataMadrasah(token, kelas, semester, tahunAjaran) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getAcademicSpreadsheet(tahunAjaran);
    const sheet = ss.getSheetByName('DATA_MADRASAH');
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]) === String(kelas) && String(data[i][2]) === String(semester)) {
        sheet.deleteRow(i + 1);
      }
    }
    return { success: true, message: 'Data Madrasah telah direset.' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * MODUL DATA SISWA (PER KELAS: DATA_SISWA_VII, VIII, IX)
 * ============================================================================
 */
function getSiswaSheetName(kelas) {
  const k = String(kelas).toUpperCase().trim();
  if (k === 'VII' || k === '7') return 'DATA_SISWA_VII';
  if (k === 'VIII' || k === '8') return 'DATA_SISWA_VIII';
  if (k === 'IX' || k === '9') return 'DATA_SISWA_IX';
  throw new Error('Kelas tidak valid: ' + kelas);
}

function getSiswaList(token, kelas, semester, tahunAjaran) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const ss = getAcademicSpreadsheet(tahunAjaran);
  const sheet = ss.getSheetByName(getSiswaSheetName(kelas));
  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][7]) === String(semester)) {
      list.push({
        id: data[i][0],
        no: data[i][1],
        nama: data[i][2],
        nisNisn: data[i][3],
        ttl: data[i][4],
        alamat: data[i][5],
        namaOrangtua: data[i][6],
        semester: data[i][7],
        tahunAjaran: data[i][8]
      });
    }
  }
  return { success: true, data: list };
}

function saveSiswa(token, kelas, semester, tahunAjaran, siswaObj) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!siswaObj.nama || !siswaObj.nama.trim()) {
      return { success: false, message: 'Nama siswa wajib diisi.' };
    }

    const ss = getAcademicSpreadsheet(tahunAjaran);
    const sheet = ss.getSheetByName(getSiswaSheetName(kelas));
    const data = sheet.getDataRange().getValues();

    if (siswaObj.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === siswaObj.id) {
          sheet.getRange(i + 1, 3, 1, 5).setValues([[
            siswaObj.nama.trim(),
            siswaObj.nisNisn || '',
            siswaObj.ttl || '',
            siswaObj.alamat || '',
            siswaObj.namaOrangtua || ''
          ]]);
          return { success: true, message: 'Data siswa berhasil diperbarui.' };
        }
      }
    }

    const countInSem = data.filter((r, idx) => idx > 0 && String(r[7]) === String(semester)).length;
    const newId = 'SIS-' + kelas + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    sheet.appendRow([
      newId,
      countInSem + 1,
      siswaObj.nama.trim(),
      siswaObj.nisNisn || '',
      siswaObj.ttl || '',
      siswaObj.alamat || '',
      siswaObj.namaOrangtua || '',
      semester,
      tahunAjaran
    ]);
    return { success: true, message: 'Siswa berhasil ditambahkan.' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function importSiswaBulk(token, kelas, semester, tahunAjaran, rowsArray) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!Array.isArray(rowsArray) || rowsArray.length === 0) {
      return { success: false, message: 'Tidak ada data valid dalam file Excel yang diunggah.' };
    }

    const ss = getAcademicSpreadsheet(tahunAjaran);
    const sheet = ss.getSheetByName(getSiswaSheetName(kelas));
    const data = sheet.getDataRange().getValues();
    let currentNo = data.filter((r, idx) => idx > 0 && String(r[7]) === String(semester)).length;

    const newRows = [];
    rowsArray.forEach(item => {
      if (item.nama && String(item.nama).trim() !== '') {
        currentNo++;
        newRows.push([
          'SIS-' + kelas + '-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
          currentNo,
          String(item.nama).trim(),
          String(item.nisNisn || '').trim(),
          String(item.ttl || '').trim(),
          String(item.alamat || '').trim(),
          String(item.namaOrangtua || '').trim(),
          semester,
          tahunAjaran
        ]);
      }
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 9).setValues(newRows);
    }

    return { success: true, message: `Berhasil mengimpor ${newRows.length} data siswa.` };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteSiswa(token, kelas, semester, tahunAjaran, idSiswa) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getAcademicSpreadsheet(tahunAjaran);
    const sheet = ss.getSheetByName(getSiswaSheetName(kelas));
    const data = sheet.getDataRange().getValues();

    if (idSiswa === 'ALL') {
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][7]) === String(semester)) {
          sheet.deleteRow(i + 1);
        }
      }
      return { success: true, message: 'Seluruh data siswa pada kelas & semester ini telah direset.' };
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idSiswa) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Data siswa berhasil dihapus.' };
      }
    }
    return { success: false, message: 'Siswa tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * MODUL DAFTAR SURAH
 * ============================================================================
 */
function getDaftarSurah(token, kelas, semester, tahunAjaran) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const ss = getAcademicSpreadsheet(tahunAjaran);
  const sheet = ss.getSheetByName('DAFTAR_SURAH');
  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]) === String(kelas) && String(data[i][4]) === String(semester)) {
      list.push({
        id: data[i][0],
        no: data[i][1],
        namaSurah: data[i][2],
        kelas: data[i][3],
        semester: data[i][4]
      });
    }
  }
  return { success: true, data: list };
}

function saveSurahSingle(token, kelas, semester, tahunAjaran, surahObj) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!surahObj.namaSurah || !surahObj.namaSurah.trim()) {
      return { success: false, message: 'Nama surah tidak boleh kosong.' };
    }
    const ss = getAcademicSpreadsheet(tahunAjaran);
    const sheet = ss.getSheetByName('DAFTAR_SURAH');
    const data = sheet.getDataRange().getValues();

    if (surahObj.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === surahObj.id) {
          sheet.getRange(i + 1, 3).setValue(surahObj.namaSurah.trim());
          return { success: true, message: 'Nama Surah berhasil diperbarui.' };
        }
      }
    }

    const existingCount = data.filter((r, idx) => idx > 0 && String(r[3]) === String(kelas) && String(r[4]) === String(semester)).length;
    const newId = 'SRH-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    sheet.appendRow([newId, existingCount + 1, surahObj.namaSurah.trim(), kelas, semester, tahunAjaran]);
    return { success: true, message: 'Surah berhasil ditambahkan.' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function importSurahBulk(token, kelas, semester, tahunAjaran, surahArray) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const ss = getAcademicSpreadsheet(tahunAjaran);
    const sheet = ss.getSheetByName('DAFTAR_SURAH');
    const data = sheet.getDataRange().getValues();
    let count = data.filter((r, idx) => idx > 0 && String(r[3]) === String(kelas) && String(r[4]) === String(semester)).length;

    const newRows = [];
    surahArray.forEach(item => {
      if (item.namaSurah && String(item.namaSurah).trim() !== '') {
        count++;
        newRows.push([
          'SRH-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
          count,
          String(item.namaSurah).trim(),
          kelas,
          semester,
          tahunAjaran
        ]);
      }
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
    }
    return { success: true, message: `Berhasil mengimpor ${newRows.length} surah.` };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteSurah(token, kelas, semester, tahunAjaran, idSurah) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getAcademicSpreadsheet(tahunAjaran);
    const sheet = ss.getSheetByName('DAFTAR_SURAH');
    const data = sheet.getDataRange().getValues();

    if (idSurah === 'ALL') {
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][3]) === String(kelas) && String(data[i][4]) === String(semester)) {
          sheet.deleteRow(i + 1);
        }
      }
      return { success: true, message: 'Daftar surah telah direset.' };
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idSurah) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Surah dihapus.' };
      }
    }
    return { success: false, message: 'Surah tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * MODUL PROGRES HAFALAN (PER KELAS: PROGRES_VII, VIII, IX)
 * ============================================================================
 */
function getProgresSheetName(kelas) {
  const k = String(kelas).toUpperCase().trim();
  if (k === 'VII' || k === '7') return 'PROGRES_VII';
  if (k === 'VIII' || k === '8') return 'PROGRES_VIII';
  if (k === 'IX' || k === '9') return 'PROGRES_IX';
  throw new Error('Kelas tidak valid');
}

function getProgresSiswa(token, kelas, semester, tahunAjaran, idSiswa) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const ss = getAcademicSpreadsheet(tahunAjaran);
  const sheet = ss.getSheetByName(getProgresSheetName(kelas));
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(idSiswa) && String(data[i][2]) === String(semester)) {
      let detail = [];
      try {
        detail = JSON.parse(data[i][8] || '[]');
      } catch (e) {
        detail = [];
      }
      return {
        success: true,
        data: {
          id: data[i][0],
          idSiswa: data[i][1],
          semester: data[i][2],
          tahunAjaran: data[i][3],
          targetSurah: data[i][4],
          jumlahTercapai: data[i][5],
          persentase: data[i][6],
          catatanGuru: data[i][7],
          detailSurah: detail
        }
      };
    }
  }

  return {
    success: true,
    data: {
      id: '',
      idSiswa: idSiswa,
      semester: semester,
      tahunAjaran: tahunAjaran,
      targetSurah: '',
      jumlahTercapai: 0,
      persentase: 0,
      catatanGuru: '',
      detailSurah: []
    }
  };
}

function saveProgresSiswa(token, kelas, semester, tahunAjaran, payload) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!payload.idSiswa) return { success: false, message: 'Pilih siswa terlebih dahulu.' };

    const ss = getAcademicSpreadsheet(tahunAjaran);
    const sheet = ss.getSheetByName(getProgresSheetName(kelas));
    const data = sheet.getDataRange().getValues();
    const detailJson = JSON.stringify(payload.detailSurah || []);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(payload.idSiswa) && String(data[i][2]) === String(semester)) {
        sheet.getRange(i + 1, 5, 1, 5).setValues([[
          payload.targetSurah,
          payload.jumlahTercapai,
          payload.persentase,
          payload.catatanGuru,
          detailJson
        ]]);
        return { success: true, message: 'Progres hafalan siswa berhasil disimpan!' };
      }
    }

    const newId = 'PRG-' + kelas + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    sheet.appendRow([
      newId,
      payload.idSiswa,
      semester,
      tahunAjaran,
      payload.targetSurah,
      payload.jumlahTercapai,
      payload.persentase,
      payload.catatanGuru,
      detailJson
    ]);
    return { success: true, message: 'Progres hafalan siswa berhasil disimpan!' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function resetProgresSiswa(token, kelas, semester, tahunAjaran, idSiswa) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getAcademicSpreadsheet(tahunAjaran);
    const sheet = ss.getSheetByName(getProgresSheetName(kelas));
    const data = sheet.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][2]) === String(semester)) {
        if (idSiswa === 'ALL' || String(data[i][1]) === String(idSiswa)) {
          sheet.deleteRow(i + 1);
        }
      }
    }
    return { success: true, message: 'Data progres hafalan berhasil direset.' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * MODUL DATA RAPOR (SATUAN & SATU KELAS SEKALIGUS)
 * ============================================================================
 */
function getFullRaporClassData(token, kelas, semester, tahunAjaran) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Sesi tidak valid.' };

  try {
    const madrasahRes = getDataMadrasah(token, kelas, semester, tahunAjaran);
    const siswaRes = getSiswaList(token, kelas, semester, tahunAjaran);
    const surahRes = getDaftarSurah(token, kelas, semester, tahunAjaran);

    const ss = getAcademicSpreadsheet(tahunAjaran);
    const progSheet = ss.getSheetByName(getProgresSheetName(kelas));
    const progData = progSheet.getDataRange().getValues();

    const progresMap = {};
    for (let i = 1; i < progData.length; i++) {
      if (String(progData[i][2]) === String(semester)) {
        let detail = [];
        try { detail = JSON.parse(progData[i][8] || '[]'); } catch (e) {}
        progresMap[progData[i][1]] = {
          targetSurah: progData[i][4],
          jumlahTercapai: progData[i][5],
          persentase: progData[i][6],
          catatanGuru: progData[i][7],
          detailSurah: detail
        };
      }
    }

    return {
      success: true,
      madrasah: madrasahRes.data,
      siswaList: siswaRes.data,
      daftarSurah: surahRes.data,
      progresMap: progresMap
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
