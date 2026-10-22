// ==========================================================
// Code.gs v2 — e-pelaporan IKP RSUD Merauke
// Struktur sheet SAMA dengan versi awal (database lama tetap terpakai).
// Baru: sesi token server-side, password di-hash, kontrol akses per peran
// di server, audit trail terisi, filter periode & tren bulanan nyata.
// ==========================================================
function getDbSpreadsheet() {
  // 1. Coba ambil dari container-bound spreadsheet (jika dibuat dari Extensions > Apps Script)
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}

  // 2. Coba ambil ID dari Script Properties
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('SPREADSHEET_ID');
  if (savedId) {
    try {
      const ss = SpreadsheetApp.openById(savedId);
      if (ss) return ss;
    } catch (e) {}
  }

  // 3. Cari file spreadsheet bernama 'Database e-IKP RSUD Merauke' di Google Drive
  try {
    const files = DriveApp.getFilesByName('Database e-IKP RSUD Merauke');
    if (files.hasNext()) {
      const file = files.next();
      props.setProperty('SPREADSHEET_ID', file.getId());
      return SpreadsheetApp.openById(file.getId());
    }
  } catch (e) {}

  // 4. Jika belum ada sama sekali, buat Google Sheet baru di Drive pengguna secara otomatis
  const newSs = SpreadsheetApp.create('Database e-IKP RSUD Merauke');
  props.setProperty('SPREADSHEET_ID', newSs.getId());
  return newSs;
}

function setSpreadsheetId(id) {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', String(id).trim());
  return 'Database e-IKP berhasil dihubungkan ke spreadsheet ID: ' + id;
}

const TZ = 'Asia/Jayapura';
// Peran pengguna. KetuaKomite = Ketua Komite Mutu (hak sama dengan Komite + penandatangan).
// Manajemen = hanya melihat (semua laporan, lembar cetak, jejak audit, ekspor), tanpa grading/RCA.
// Direktur  = melihat semua laporan + MEREVIEW hasil RCA & rencana perbaikan (corrective action) yang dilaporkan Komite
//             (evaluasi, pembelajaran, umpan balik perbaikan sistem). Direktur tidak melakukan grading/RCA.
const ROLES_ALL = ['Staf', 'Karu', 'Komite', 'KetuaKomite', 'Manajemen', 'Direktur', 'TU', 'Admin'];
const ROLES_REGISTER = ['Staf', 'Karu', 'Komite', 'KetuaKomite', 'Manajemen', 'Direktur', 'TU'];
const ROLES_KOMITE = ['Komite', 'KetuaKomite'];
// Peran yang langsung aktif saat mendaftar (tanpa konfirmasi Admin). Karu & peran lain tetap menunggu Admin.
const ROLES_AUTO_APPROVE = ['Staf'];
const ROLES_VIEW_ALL = ['Admin', 'Komite', 'KetuaKomite', 'Manajemen', 'Direktur'];
const JENIS = ['KPCS', 'KNC', 'KTC', 'KTD', 'Sentinel'];
// Istilah baru: KPC (Kondisi Potensial Cedera) -> KPCS (Kondisi Potensial Cedera Signifikan).
// Data lama bernilai 'KPC' tetap terbaca sebagai KPCS (lihat normJenis & migrasiKPCkeKPCS).
function normJenis(v) { v = String(v == null ? '' : v).trim(); return v === 'KPC' ? 'KPCS' : v; }
const GRADES = ['Biru', 'Hijau', 'Kuning', 'Merah'];
const BLN = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
const FORM_INTERNAL = [
  ["h_pasien", "I. DATA PASIEN", "heading", "", "false", 1],
  ["q_nama_pasien", "Nama Pasien / Inisial", "text", "", "true", 2],
  ["q_no_rm", "No. Rekam Medis (RM)", "text", "", "true", 3],
  ["q_ruangan_pasien", "Ruangan Pasien", "text", "", "false", 4],
  ["q_umur_tahun", "Umur (Tahun)", "number", "", "false", 5],
  ["q_umur_bulan", "Umur (Bulan)", "number", "", "false", 6],
  ["q_kelompok_umur", "Kelompok Umur", "radio", "0-1 bulan,> 1 bulan - 1 tahun,> 1 tahun - 5 tahun,> 5 tahun - 15 tahun,> 15 tahun - 30 tahun,> 30 tahun - 65 tahun,> 65 tahun", "true", 7],
  ["q_jenis_kelamin", "Jenis Kelamin", "radio", "Laki-laki,Perempuan", "true", 8],
  ["q_penanggung_biaya", "Penanggung Biaya Pasien", "radio", "Pribadi,Asuransi swasta,Pemerintah,Perusahaan,BPJS,Lain-lain", "false", 9],
  ["q_tgl_masuk_rs", "Tanggal & Jam Masuk RS", "datetime-local", "", "false", 10],
  
  ["h_kejadian", "II. RINCIAN KEJADIAN", "heading", "", "false", 11],
  ["q_tgl_insiden", "1. Tanggal dan Waktu Insiden", "datetime-local", "", "true", 12],
  ["q_insiden", "2. Insiden (judul singkat)", "text", "", "true", 13],
  ["q_kronologi", "3. Kronologis Insiden", "textarea", "", "true", 14],
  ["q_jenis_insiden", "4. Jenis Insiden", "radio", "Kejadian Sentinel (Sentinel Event),Kejadian Tidak Diharapkan/KTD (Adverse Event),Kejadian Tidak Cedera/KTC,Kejadian Nyaris Cedera/KNC (Near Miss)", "true", 15],
  ["q_pelapor_pertama", "5. Orang Pertama Yang Melaporkan Insiden", "radio", "Karyawan: Dokter/Perawat/Petugas lainnya,Pasien,Keluarga/Pendamping Pasien,Pengunjung,Lain-lain", "true", 16],
  ["q_pelapor_lain", "5a. Jika Lain-lain, sebutkan", "text", "", "false", 17],
  ["q_menyangkut", "6. Insiden Menyangkut Pasien", "radio", "Pasien Rawat Inap,Pasien IGD,Pasien Rawat Jalan,Lain-lain", "true", 18],
  ["q_menyangkut_lain", "6a. Jika Lain-lain, sebutkan", "text", "", "false", 19],
  ["q_lokasi", "7. Lokasi Kejadian (tempat pasien berada)", "text", "", "true", 20],
  ["q_spesialisasi", "8. Insiden Terjadi Pada Pasien (kasus penyakit / spesialisasi, termasuk subspesialisasinya)", "select", "Penyakit Dalam,Anak,Bedah,Obstetri Ginekologi,THT,Mata,Saraf,Anestesi,Kulit & Kelamin,Jantung,Paru,Jiwa,Lain-lain", "false", 21],
  ["q_spesialisasi_lain", "8a. Jika Lain-lain, sebutkan", "text", "", "false", 22],
  ["q_unit_penyebab", "9. Unit / Departemen terkait yang menyebabkan insiden", "text", "", "false", 23],
  ["q_akibat", "10. Akibat Insiden terhadap Pasien", "radio", "Kematian,Cedera Berat / cedera irreversible,Cedera Sedang / cedera reversible,Cedera Ringan,Tidak Ada Cedera", "true", 24],
  
  // TAB/SECTION BARU: III. TINDAKAN DAN TIM
  ["h_tindakan_tim", "III. TINDAKAN DAN TIM", "heading", "", "false", 25],
  ["q_tindakan", "1. Tindakan yang dilakukan segera setelah kejadian dan hasilnya", "textarea", "", "true", 26],
  ["h_tim", "2. Tindakan dilakukan oleh (tim terdiri dari)", "heading", "", "false", 27],
  ["q_tim_dokter", "Dokter", "text", "", "false", 28],
  ["q_tim_perawat", "Perawat", "text", "", "false", 29],
  ["q_tim_lain", "Petugas lainnya", "text", "", "false", 30],
  
  // LANJUTAN RINCIAN KEJADIAN / PERTANYAAN LAIN
  ["q_kejadian_sama", "3. Apakah kejadian yang sama pernah terjadi di unit kerja lain?", "radio", "Ya,Tidak", "false", 31],
  ["q_kejadian_sama_detail", "3a. Jika Ya: kapan, dan langkah / tindakan apa yang diambil unit tersebut untuk mencegah terulangnya kejadian yang sama?", "textarea", "", "false", 32]
];
const RCA_K = ['rca_id', 'report_id', 'risk_level', 'judul_rca', 'ringkasan_kronologi', 'masalah_utama', 'why1', 'why2', 'why3', 'why4', 'why5', 'root_cause', 'contributing_factors', 'corrective_action', 'preventive_action', 'pic', 'target_selesai', 'status', 'created_at', 'indikator_keberhasilan'];

// ---------- Status alur kerja ----------
// Staf lapor -> Karu grading -> Komite regrading -> (Kuning/Merah) Komite isi RCA + rencana perbaikan
// -> DILAPORKAN KE DIREKSI -> Direktur review (evaluasi, pembelajaran, umpan balik sistem)
//    -> Disetujui / Disetujui dengan Catatan : Selesai
//    -> Dikembalikan untuk Revisi            : Revisi RCA (kembali ke Komite, lalu dikirim ulang ke Direksi)
const ST_KARU = 'Menunggu Grading Karu';
const ST_KOMITE = 'Menunggu Regrading Komite';
const ST_RCA = 'Proses RCA';
const ST_DIREKSI = 'Menunggu Review Direksi';
const ST_REVISI = 'Revisi RCA';
const ST_SELESAI = 'Selesai';
const ROLES_DIREKSI = ['Direktur'];
const REVIEW_KEPUTUSAN = ['Disetujui', 'Disetujui dengan Catatan', 'Dikembalikan untuk Revisi'];
const REVIEW_HEADER = ['review_id', 'report_id', 'rca_id', 'timestamp', 'username', 'nama', 'keputusan', 'evaluasi', 'pembelajaran', 'rekomendasi_sistem'];

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Sistem Komite RS - RSUD Merauke')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // wajib agar bisa dibuka dari aplikasi Android (PWA)
}

// ---------- Utilitas ----------
function sh(n) { return getDbSpreadsheet().getSheetByName(n); }
function ok(o) { return Object.assign({ success: true }, o || {}); }
// Sheet review Direksi dibuat otomatis saat pertama dipakai (database lama tidak perlu diubah manual).
function reviewSheet(createIfMissing) {
  const ss = getDbSpreadsheet();
  let s = ss.getSheetByName('Review_Direksi');
  if (!s && createIfMissing) {
    s = ss.insertSheet('Review_Direksi');
    s.appendRow(REVIEW_HEADER);
    s.getRange(1, 1, 1, REVIEW_HEADER.length).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  }
  return s;
}
function reviewRows(reportId) {
  const s = reviewSheet(false);
  if (!s || s.getLastRow() < 2) return [];
  return s.getDataRange().getValues().slice(1).filter(r => r[1] === reportId).map(r => ({
    review_id: r[0], report_id: r[1], rca_id: r[2], timestamp: fmt(r[3]), username: r[4], nama: r[5],
    keputusan: r[6], evaluasi: r[7], pembelajaran: r[8], rekomendasi_sistem: r[9]
  }));
}
// Kolom indikator_keberhasilan (kolom 20) ditambahkan otomatis pada sheet RCA_IKP lama.
function rcaSheet() {
  const s = sh('RCA_IKP');
  if (!String(s.getRange(1, 20).getValue()).trim()) s.getRange(1, 20).setValue('indikator_keberhasilan').setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  return s;
}
function direkturEmails() {
  const out = {};
  try {
    const st = sh('App_Settings');
    if (st) st.getDataRange().getValues().forEach(r => { if (r[0] === 'direksi_email' && r[1]) String(r[1]).split(/[,;\s]+/).forEach(e => { if (e) out[e] = true; }); });
    sh('Users').getDataRange().getValues().slice(1).forEach(r => {
      if (String(r[3]).trim() === 'Direktur' && String(r[5]).trim().toLowerCase() === 'approved' && /@/.test(String(r[6] || ''))) out[String(r[6]).trim()] = true;
    });
  } catch (e) {}
  return Object.keys(out).join(',');
}
function fail(e) {
  return e.message === 'AUTH'
    ? { success: false, code: 'AUTH', message: 'Sesi berakhir. Silakan masuk kembali.' }
    : { success: false, message: e.message };
}
function locked(fn) {
  const l = LockService.getScriptLock();
  try { l.waitLock(10000); return fn(); } catch (e) { return fail(e); } finally { try { l.releaseLock(); } catch (x) {} }
}
function fmt(v, f) { return v instanceof Date ? Utilities.formatDate(v, TZ, f || 'dd/MM/yyyy HH:mm') : v; }
function T(v, n) { return String(v || '').trim().slice(0, n || 2000); }
function escHtml(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

// ---------- Password & sesi ----------
function hashPw(pw, salt) {
  salt = salt || Utilities.getUuid().slice(0, 12);
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + pw, Utilities.Charset.UTF_8);
  return 'sha256$' + salt + '$' + Utilities.base64Encode(d);
}
function checkPw(pw, stored) {
  stored = String(stored).trim();
  if (stored.indexOf('sha256$') === 0) return hashPw(pw, stored.split('$')[1]) === stored;
  return stored === pw; // password lama (teks biasa), akan di-hash saat login
}
function findUser(un) {
  un = String(un).trim().toLowerCase();
  const d = sh('Users').getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]).trim().toLowerCase() === un) {
      return { row: i + 1, username: String(d[i][0]).trim(), pw: d[i][1], nama: String(d[i][2]).trim(), role: String(d[i][3]).trim(), ruangan: String(d[i][4]).trim(), status: String(d[i][5]).trim(), email: String(d[i][6] || '').trim(), nip: String(d[i][7] || '').trim(), ttd: String(d[i][8] || '').trim() };
    }
  }
  return null;
}
function pub(u) { return { username: u.username, nama: u.nama, role: u.role, ruangan: u.ruangan, nip: u.nip }; }
// Peran & ruangan SELALU diambil dari sheet, bukan dari browser.
function auth(token, roles) {
  const un = token && CacheService.getScriptCache().get('s_' + token);
  if (!un) throw new Error('AUTH');
  const u = findUser(un);
  if (!u || u.status.toLowerCase() !== 'approved') throw new Error('AUTH');
  if (roles && roles.indexOf(u.role) < 0) throw new Error('Akses ditolak untuk peran Anda.');
  return u;
}
function audit(id, u, action, from, to, details) {
  sh('Audit_Trail').appendRow(['AUD-' + Utilities.getUuid().slice(0, 8), id, new Date(), u.username, u.nama, u.role, action, from || '', to || '', details || '']);
  // Setiap aktivitas yang tercatat di audit trail ikut dikirim ke Telegram (gagal kirim tidak mengganggu proses utama).
  try { notifAktivitas(id, u, action, from, to, details); } catch (e) { Logger.log('notifAktivitas gagal: ' + e.message); }
}

function apiAuthenticateUser(username, password) {
  try {
    const c = CacheService.getScriptCache(), key = 'f_' + String(username).trim().toLowerCase(), pw = String(password).trim();
    if (Number(c.get(key) || 0) >= 5) return { success: false, message: 'Terlalu banyak percobaan gagal. Coba lagi dalam 10 menit.' };
    const u = findUser(username);
    if (!u || !checkPw(pw, u.pw)) { c.put(key, String(Number(c.get(key) || 0) + 1), 600); return { success: false, message: 'Username atau password salah.' }; }
    if (u.status.toLowerCase() !== 'approved') return { success: false, message: 'Akun belum disetujui Administrator (status: ' + u.status + ').' };
    c.remove(key);
    if (String(u.pw).indexOf('sha256$') !== 0) sh('Users').getRange(u.row, 2).setValue(hashPw(pw));
    const token = Utilities.getUuid() + Utilities.getUuid();
    c.put('s_' + token, u.username, 21600);
    return { success: true, token: token, user: pub(u) };
  } catch (e) { return fail(e); }
}
function apiRestoreSession(token) {
  try { return { success: true, token: token, user: pub(auth(token)) }; } catch (e) { return { success: false }; }
}
function apiLogout(token) { CacheService.getScriptCache().remove('s_' + token); return { success: true }; }

function maskEmail(em) {
  em = String(em || '').trim();
  const at = em.indexOf('@');
  if (at < 1) return 'email terdaftar';
  const local = em.slice(0, at), domain = em.slice(at);
  return (local.length <= 2 ? local[0] + '*' : local.slice(0, 2) + '***') + domain;
}
function otpDigest(un, otp) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(un).toLowerCase() + '|' + String(otp), Utilities.Charset.UTF_8);
  return Utilities.base64Encode(d);
}

function apiRequestPasswordReset(username) {
  try {
    const c = CacheService.getScriptCache();
    const un = String(username || '').trim().toLowerCase();
    if (!un) throw new Error('Username wajib diisi.');
    const reqKey = 'otpr_' + un;
    if (Number(c.get(reqKey) || 0) >= 3) {
      return { success: false, message: 'Terlalu banyak permintaan OTP. Coba lagi dalam 10 menit.' };
    }
    const u = findUser(un);
    c.put(reqKey, String(Number(c.get(reqKey) || 0) + 1), 600);
    if (!u || u.status.toLowerCase() !== 'approved') {
      return ok({ message: 'Jika username terdaftar dan disetujui, kode OTP 6 digit akan dikirim ke email pendaftar. Kode berlaku 10 menit.' });
    }
    if (!u.email || u.email.indexOf('@') < 1) {
      return { success: false, message: 'Akun ini belum memiliki email terdaftar. Hubungi Administrator untuk mereset password.' };
    }
    const otp = String(Math.floor(100000 + Math.random() * 900000));
    c.put('otp_' + un, otpDigest(un, otp), 600);
    c.put('otpn_' + un, '0', 600);
    MailApp.sendEmail({
      to: u.email,
      subject: '[e-IKP RSUD Merauke] Kode OTP Reset Password',
      htmlBody: '<div style="font-family:Arial,sans-serif;padding:20px;border:1px solid #e2e8f0;border-radius:8px;max-width:520px">' +
        '<h2 style="color:#1e3a8a;margin:0 0 8px">Reset Password e-IKP</h2>' +
        '<p>Yth. <b>' + escHtml(u.nama) + '</b>,</p>' +
        '<p>Kode OTP untuk mengatur ulang password akun <b>' + escHtml(u.username) + '</b> adalah:</p>' +
        '<p style="font-size:28px;letter-spacing:8px;font-weight:700;color:#00236f;margin:16px 0">' + otp + '</p>' +
        '<p>Kode berlaku <b>10 menit</b>. Jangan bagikan kode ini kepada siapa pun.</p>' +
        '<hr><small style="color:#64748b">Pesan otomatis Sistem e-IKP RSUD Merauke. Abaikan jika Anda tidak meminta reset password.</small></div>'
    });
    return ok({ message: 'Kode OTP 6 digit telah dikirim ke ' + maskEmail(u.email) + '. Kode berlaku 10 menit. Periksa juga folder spam.' });
  } catch (e) { return fail(e); }
}

function apiResetPassword(username, otp, newPw) {
  return locked(() => {
    const c = CacheService.getScriptCache();
    const un = String(username || '').trim().toLowerCase();
    const code = String(otp || '').trim();
    const pw = String(newPw || '').trim();
    if (!un || !/^[0-9]{6}$/.test(code)) throw new Error('Kode OTP tidak valid.');
    if (pw.length < 8) throw new Error('Password baru minimal 8 karakter.');
    const stored = c.get('otp_' + un);
    if (!stored) throw new Error('Kode OTP kedaluwarsa atau belum diminta. Kirim ulang kode.');
    const tries = Number(c.get('otpn_' + un) || 0);
    if (tries >= 5) {
      c.remove('otp_' + un);
      throw new Error('Terlalu banyak percobaan OTP salah. Kirim ulang kode.');
    }
    if (stored !== otpDigest(un, code)) {
      c.put('otpn_' + un, String(tries + 1), 600);
      throw new Error('Kode OTP salah.');
    }
    const u = findUser(un);
    if (!u || u.status.toLowerCase() !== 'approved') throw new Error('Akun tidak dapat direset.');
    sh('Users').getRange(u.row, 2).setValue(hashPw(pw));
    c.remove('otp_' + un);
    c.remove('otpn_' + un);
    c.remove('otpr_' + un);
    audit('USER:' + u.username, u, 'Reset password OTP');
    return ok({ message: 'Password berhasil diatur ulang. Silakan masuk dengan password baru.' });
  });
}

function apiChangePassword(token, oldPw, newPw) {
  return locked(() => {
    const u = auth(token);
    if (!checkPw(String(oldPw).trim(), u.pw)) throw new Error('Password lama salah.');
    if (String(newPw).trim().length < 8) throw new Error('Password baru minimal 8 karakter.');
    sh('Users').getRange(u.row, 2).setValue(hashPw(String(newPw).trim()));
    audit('USER:' + u.username, u, 'Ganti password');
    return ok({ message: 'Password berhasil diganti.' });
  });
}
// ---------- Profil pengguna ----------
// Pengguna hanya dapat mengubah profilnya sendiri. Username, peran, & status TIDAK dapat diubah di sini.
// Ruangan hanya boleh diubah oleh Staf (ruangan Karu menentukan hak akses grading, jadi dikunci).
function apiGetMyProfile(token) {
  try {
    const u = auth(token);
    return ok({ profile: { username: u.username, nama: u.nama, role: u.role, ruangan: u.ruangan, email: u.email, nip: u.nip, can_edit_ruangan: u.role === 'Staf' } });
  } catch (e) { return fail(e); }
}
function apiUpdateMyProfile(token, d) {
  return locked(() => {
    const u = auth(token);
    d = d || {};
    const nama = T(d.nama, 80);
    if (!nama) throw new Error('Nama lengkap wajib diisi.');
    const email = T(d.email, 100);
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Format email tidak valid.');
    const nip = T(d.nip, 30).replace(/\s+/g, ' ');
    if (nip && !/^[0-9][0-9 ]{5,29}$/.test(nip)) throw new Error('NIP hanya berisi angka (boleh spasi), minimal 6 digit. Kosongkan jika bukan pegawai.');
    // Mengganti email (dipakai untuk OTP reset password) harus dikonfirmasi dengan kata sandi
    if (email.toLowerCase() !== String(u.email || '').toLowerCase() && !checkPw(String(d.password || '').trim(), u.pw)) {
      throw new Error('Untuk mengubah email, masukkan kata sandi Anda saat ini dengan benar.');
    }
    let ruangan = u.ruangan;
    if (u.role === 'Staf' && d.ruangan != null) {
      ruangan = T(d.ruangan, 60);
      if (!ruangan) throw new Error('Ruangan wajib diisi.');
    }
    const us = sh('Users');
    if (!String(us.getRange(1, 8).getValue()).trim()) us.getRange(1, 8).setValue('nip').setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
    us.getRange(u.row, 3).setValue(nama);
    us.getRange(u.row, 5).setValue(ruangan);
    us.getRange(u.row, 7).setValue(email);
    us.getRange(u.row, 8).setNumberFormat('@').setValue(nip);
    const det = [];
    if (nama !== u.nama) det.push('nama');
    if (ruangan !== u.ruangan) det.push('ruangan');
    if (email !== u.email) det.push('email');
    if (nip !== u.nip) det.push('nip');
    audit('USER:' + u.username, u, 'Ubah profil', '', '', det.join(', ') || 'tanpa perubahan');
    return ok({ message: 'Profil berhasil diperbarui.', user: pub(findUser(u.username)) });
  });
}
function apiRegisterUser(d) {
  return locked(() => {
    const un = String(d.username || '').trim().toLowerCase(), pw = String(d.password || '').trim();
    if (!/^[a-z0-9_.-]{3,30}$/.test(un)) throw new Error('Username 3-30 karakter: huruf kecil, angka, titik, garis bawah, atau strip.');
    if (pw.length < 8) throw new Error('Password minimal 8 karakter.');
    if (ROLES_REGISTER.indexOf(d.role) < 0) throw new Error('Peran tidak valid.');
    if (!T(d.nama) || !T(d.ruangan)) throw new Error('Nama dan ruangan wajib diisi.');
    const nip = T(d.nip, 30).replace(/\s+/g, ' ');
    if (nip && !/^[0-9][0-9 ]{5,29}$/.test(nip)) throw new Error('NIP hanya berisi angka (boleh spasi), minimal 6 digit. Kosongkan jika bukan pegawai.');
    if (findUser(un)) throw new Error('Username sudah terdaftar.');
    if (isDeletedUsername(un)) throw new Error('Username ini pernah dipakai dan tidak dapat didaftarkan ulang. Gunakan username lain.');
    const us = sh('Users');
    if (!String(us.getRange(1, 8).getValue()).trim()) us.getRange(1, 8).setValue('nip').setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
    // Hanya Staf yang langsung aktif (Approved). Karu, Komite, KetuaKomite, Manajemen, Direktur tetap menunggu Admin.
    const autoApprove = ROLES_AUTO_APPROVE.indexOf(d.role) >= 0;
    const stAkun = autoApprove ? 'Approved' : 'Pending';
    us.appendRow([un, hashPw(pw), T(d.nama, 80), d.role, T(d.ruangan, 60), stAkun, T(d.email, 100)]);
    us.getRange(us.getLastRow(), 8).setNumberFormat('@').setValue(nip); // teks agar angka 0 di depan tidak hilang
    audit('USER:' + un, { username: un, nama: T(d.nama, 80), role: d.role }, 'Pendaftaran akun', '', stAkun, autoApprove ? 'Otomatis disetujui' : 'Menunggu persetujuan Admin');
    sendTelegramNotification(
      '<b>PENDAFTARAN AKUN BARU - e-IKP RSUD Merauke</b>\n' +
      'Nama: ' + escHtml(T(d.nama, 80)) + '\n' +
      'Username: ' + escHtml(un) + '\n' +
      'Peran: ' + escHtml(d.role) + '\n' +
      'Unit: ' + escHtml(T(d.ruangan, 60)) + '\n' +
      'Status: ' + (autoApprove ? 'Aktif (otomatis disetujui)' : '<b>PENDING</b> - menunggu persetujuan Admin'),
      ['admin', 'komite']);
    return ok({
      approved: autoApprove,
      message: autoApprove
        ? 'Pendaftaran berhasil. Akun Anda langsung aktif, silakan masuk.'
        : 'Pendaftaran berhasil. Silakan tunggu persetujuan Administrator.'
    });
  });
}

// ---------- Admin: pengguna ----------
function apiGetUsers(token) {
  try {
    auth(token, ['Admin']);
    const d = sh('Users').getDataRange().getValues();
    return ok({ users: d.slice(1).map(r => ({ username: String(r[0]).trim(), nama: String(r[2]).trim(), role: String(r[3]).trim(), ruangan: String(r[4]).trim(), status: String(r[5]).trim(), nip: String(r[7] || '').trim(), email: String(r[6] || '').trim() })) });
  } catch (e) { return fail(e); }
}
function apiUpdateUserStatus(token, username, newStatus, newRole) {
  return locked(() => {
    const a = auth(token, ['Admin']), u = findUser(username);
    if (!u) throw new Error('Pengguna tidak ditemukan.');
    if (u.username.toLowerCase() === a.username.toLowerCase()) throw new Error('Anda tidak dapat mengubah akun sendiri.');
    if (newStatus) {
      if (['Approved', 'Rejected', 'Pending'].indexOf(newStatus) < 0) throw new Error('Status tidak valid.');
      sh('Users').getRange(u.row, 6).setValue(newStatus);
      audit('USER:' + u.username, a, 'Status akun', u.status, newStatus);
    }
    if (newRole) {
      if (ROLES_ALL.indexOf(newRole) < 0) throw new Error('Peran tidak valid.');
      sh('Users').getRange(u.row, 4).setValue(newRole);
      audit('USER:' + u.username, a, 'Peran akun', u.role, newRole);
    }
    return ok({ message: 'Data pengguna berhasil diperbarui.' });
  });
}

// Daftar username yang pernah dihapus (arsip). Mencegah username dipakai ulang oleh orang lain,
// sehingga laporan lama atas nama username tersebut tidak otomatis terlihat oleh pemilik akun baru.
function deletedUsersSheet() {
  const ss = getDbSpreadsheet();
  let s = ss.getSheetByName('Users_Dihapus');
  if (!s) {
    s = ss.insertSheet('Users_Dihapus');
    const h = ['username', 'nama', 'role', 'ruangan', 'dihapus_oleh', 'timestamp'];
    s.appendRow(h);
    s.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  }
  return s;
}
function isDeletedUsername(un) {
  const s = getDbSpreadsheet().getSheetByName('Users_Dihapus');
  if (!s || s.getLastRow() < 2) return false;
  un = String(un).trim().toLowerCase();
  return s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().some(r => String(r[0]).trim().toLowerCase() === un);
}

// Admin mengubah data akun terdaftar. Username TIDAK dapat diubah (dipakai sebagai kunci laporan & tanda tangan).
// Akun sendiri diubah lewat menu Profil Saya.
function apiAdminUpdateUser(token, username, d) {
  return locked(() => {
    const a = auth(token, ['Admin']), u = findUser(username);
    if (!u) throw new Error('Pengguna tidak ditemukan.');
    if (u.username.toLowerCase() === a.username.toLowerCase()) throw new Error('Untuk mengubah akun Anda sendiri, gunakan menu Profil Saya.');
    d = d || {};
    const nama = T(d.nama, 80), ruangan = T(d.ruangan, 60), email = T(d.email, 100);
    const role = String(d.role || u.role).trim();
    const nip = T(d.nip, 30).replace(/\s+/g, ' ');
    const pw = String(d.password || '').trim();
    if (!nama) throw new Error('Nama lengkap wajib diisi.');
    if (!ruangan) throw new Error('Unit / ruangan wajib diisi.');
    if (ROLES_ALL.indexOf(role) < 0) throw new Error('Peran tidak valid.');
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Format email tidak valid.');
    if (nip && !/^[0-9][0-9 ]{5,29}$/.test(nip)) throw new Error('NIP hanya berisi angka (boleh spasi), minimal 6 digit. Kosongkan jika bukan pegawai.');
    if (pw && pw.length < 8) throw new Error('Password baru minimal 8 karakter.');
    const us = sh('Users');
    if (!String(us.getRange(1, 8).getValue()).trim()) us.getRange(1, 8).setValue('nip').setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
    us.getRange(u.row, 3).setValue(nama);
    us.getRange(u.row, 4).setValue(role);
    us.getRange(u.row, 5).setValue(ruangan);
    us.getRange(u.row, 7).setValue(email);
    us.getRange(u.row, 8).setNumberFormat('@').setValue(nip);
    if (pw) us.getRange(u.row, 2).setValue(hashPw(pw));
    const det = [];
    if (nama !== u.nama) det.push('nama');
    if (role !== u.role) det.push('peran ' + u.role + '→' + role);
    if (ruangan !== u.ruangan) det.push('ruangan');
    if (email !== u.email) det.push('email');
    if (nip !== u.nip) det.push('nip');
    if (pw) det.push('password direset');
    audit('USER:' + u.username, a, 'Edit akun', '', '', det.join(', ') || 'tanpa perubahan');
    return ok({ message: 'Data akun "' + u.username + '" berhasil diperbarui.' });
  });
}

// Admin menghapus akun. Laporan, lampiran, dan jejak audit milik akun tersebut TETAP disimpan (arsip mutu & SNARS).
// Sesi aktif akun yang dihapus langsung tidak berlaku (auth() mencari akun di sheet Users setiap permintaan).
function apiDeleteUser(token, username) {
  return locked(() => {
    const a = auth(token, ['Admin']), u = findUser(username);
    if (!u) throw new Error('Pengguna tidak ditemukan.');
    if (u.username.toLowerCase() === a.username.toLowerCase()) throw new Error('Anda tidak dapat menghapus akun sendiri.');
    if (u.role === 'Admin') {
      const admins = sh('Users').getDataRange().getValues().slice(1)
        .filter(r => String(r[3]).trim() === 'Admin' && String(r[5]).trim().toLowerCase() === 'approved').length;
      if (admins <= 1) throw new Error('Tidak dapat menghapus Administrator terakhir.');
    }
    deletedUsersSheet().appendRow([u.username, u.nama, u.role, u.ruangan, a.username, new Date()]);
    sh('Users').deleteRow(u.row);
    audit('USER:' + u.username, a, 'Hapus akun', u.status, '', u.nama + ' (' + u.role + ', ' + u.ruangan + ')');
    return ok({ message: 'Akun "' + u.username + '" berhasil dihapus.' });
  });
}

// ---------- Admin: pertanyaan formulir ----------
function apiGetFormQuestions(token) {
  try {
    auth(token);
    const d = sh('Form_Questions').getDataRange().getValues();
    const q = d.slice(1).map(r => ({ id: r[0], label: r[1], type: r[2], options: r[3], required: r[4] === 'true' || r[4] === true, order: r[5] }));
    q.sort((a, b) => a.order - b.order);
    return ok({ questions: q });
  } catch (e) { return fail(e); }
}
function apiSaveFormQuestion(token, q) {
  return locked(() => {
    const a = auth(token, ['Admin']);
    if (!T(q.label)) throw new Error('Label wajib diisi.');
    if (['text', 'textarea', 'select', 'radio', 'datetime-local', 'date', 'time', 'number', 'heading'].indexOf(q.type) < 0) throw new Error('Tipe input tidak valid.');
    const s = sh('Form_Questions'), d = s.getDataRange().getValues();
    if (q.id) {
      for (let i = 1; i < d.length; i++) if (d[i][0] === q.id) {
        s.getRange(i + 1, 2, 1, 4).setValues([[T(q.label, 200), q.type, T(q.options, 500), String(q.type !== 'heading' && q.required !== false)]]);
        audit('FORM', a, 'Ubah pertanyaan', '', '', T(q.label, 200));
        return ok({ message: 'Pertanyaan berhasil diperbarui.' });
      }
      throw new Error('Pertanyaan tidak ditemukan.');
    }
    s.appendRow(['q_' + Date.now(), T(q.label, 200), q.type, T(q.options, 500), String(q.type !== 'heading' && q.required !== false), Math.max(0, ...d.slice(1).map(r => Number(r[5]) || 0)) + 1]);
    audit('FORM', a, 'Tambah pertanyaan', '', '', T(q.label, 200));
    return ok({ message: 'Pertanyaan baru berhasil ditambahkan.' });
  });
}
function apiDeleteFormQuestion(token, qId) {
  return locked(() => {
    const a = auth(token, ['Admin']);
    if (qId === 'q_jenis_insiden') throw new Error('Pertanyaan "Jenis Insiden" adalah pertanyaan inti dan tidak dapat dihapus (teksnya boleh diubah).');
    const s = sh('Form_Questions'), d = s.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) if (d[i][0] === qId) { audit('FORM', a, 'Hapus pertanyaan', '', '', T(d[i][1], 200)); s.deleteRow(i + 1); return ok({ message: 'Pertanyaan berhasil dihapus.' }); }
    throw new Error('Pertanyaan tidak ditemukan.');
  });
}

// Simpan urutan baru: ids = daftar id pertanyaan dari atas ke bawah
function apiReorderFormQuestions(token, ids) {
  return locked(() => {
    const a = auth(token, ['Admin']);
    const s = sh('Form_Questions'), d = s.getDataRange().getValues();
    if (!Array.isArray(ids) || ids.length !== d.length - 1) throw new Error('Daftar urutan tidak lengkap. Muat ulang halaman.');
    const pos = {}; ids.forEach((id, i) => pos[id] = i + 1);
    const col = d.slice(1).map(r => { if (!pos[r[0]]) throw new Error('Urutan tidak cocok. Muat ulang halaman.'); return [pos[r[0]]]; });
    s.getRange(2, 6, col.length, 1).setValues(col);
    audit('FORM', a, 'Ubah urutan pertanyaan');
    return ok({ message: 'Urutan pertanyaan disimpan.' });
  });
}


// ---------- Referensi: Jenis Insiden yang harus dilaporkan ----------
// Sheet "Jenis_Insiden" dibuat otomatis (berisi 5 jenis bawaan) saat pertama kali dibuka. Semua peran boleh melihat; hanya Admin yang dapat menambah/mengubah/menghapus.
const JENIS_INFO_HEADER = ['id', 'kode', 'nama', 'inggris', 'definisi', 'contoh_json', 'catatan', 'urutan'];
const JENIS_INFO_SEED = [
 {
  "id": "ji_kpcs",
  "kode": "KPCS",
  "nama": "Kondisi Potensial Cedera Signifikan",
  "inggris": "",
  "definisi": "Kejadian Potensial Cedera/KPCS adalah suatu kondisi (selain dari proses penyakit atau kondisi pasien itu sendiri) yang berpotensi menyebabkan cedera atau kejadian sentinel.",
  "contoh": [
   "Lantai licin di koridor atau kamar mandi tanpa adanya tanda peringatan bahaya",
   "Penyimpanan obat dengan kategori LASA (Look Alike, Sound Alike / mirip rupa dan ucapan) yang diletakkan berdampingan tanpa pengamanan atau pelabelan khusus.",
   "Peralatan medis darurat di ruang perawatan (seperti alat kejut jantung / defibrilator, monitor tensi, Ventilator, syringe pump, EKG, Tensi, Oksimeter, Suction, termometer) yang berada dalam posisi siaga tetapi belum pernah diperiksa atau mengalami kerusakan komponen",
   "Kabel alat kesehatan yang melintang di lantai ruang tindakan dan berpotensi membuat petugas atau orang lain tersandung.",
   "Alat cateter CDL kosong",
   "Stok reagen kosong"
  ],
  "catatan": "",
  "order": 1
 },
 {
  "id": "ji_knc",
  "kode": "KNC",
  "nama": "Kejadian Nyaris Cedera",
  "inggris": "Near Missed",
  "definisi": "Kejadian Nyaris Cedera/KNC adalah terjadinya insiden yang belum terpapar ke pasien.",
  "contoh": [
   "Perawat salah mengambil obat yang mirip tersebut, tetapi sadar dan membatalkannya sesaat sebelum menyuntikkan obat ke pasien",
   "Pasien alergi makanan/obat tertentu tapi tidak terpasang stiker penanda alergi digelang identitas pasien.",
   "Salah penulisan obat/dosis namun di ketahui farmasis dan dikonfirmasi ke dokter",
   "Semua perbedaan besar antara diagnosis pre-operasi dan post-operasi; Ketidaksesuaian diagnosa medik pre-operasi dan post-operasi.",
   "Pasien dipindahkan dari IGD ke ruangan rawat tidak disertakan formulir transfer antar ruangan",
   "Keterlambatan pemberian hasil lab karena alat rusak/alasan lainnya",
   "Delayed pemberian obat vit K injeksi karena persediaan kosong"
  ],
  "catatan": "",
  "order": 2
 },
 {
  "id": "ji_ktc",
  "kode": "KTC",
  "nama": "Kejadian Tidak Cedera",
  "inggris": "No Harm Incident",
  "definisi": "Kejadian Tidak Cedera/KTC adalah terjadinya insiden yang sudah terpapar ke pasien tetapi tidak menimbulkan cedera.",
  "contoh": [
   "Pasien dilakukan RO ulang karena hasil rontgen sebelumnya hilang/alasan lainnya",
   "USG harus diulang karena hasil tidak lengkap",
   "Sampel darah hilang/lisis di laboratorium, sehingga di ambil darah ulang kembali pada pasien.",
   "Hasil pemeriksaan darah tertukar karena kesalahan penempelan barcode/penulisan nama",
   "Pasien BAB/BAK di meja operasi",
   "Memasukkan cairan/ obat2an tertentu kedalam cairan irigasi folley kateter"
  ],
  "catatan": "",
  "order": 3
 },
 {
  "id": "ji_ktd",
  "kode": "KTD",
  "nama": "Kejadian Tidak Diharapkan",
  "inggris": "Harmful Incident, Adverse Event",
  "definisi": "Suatu kejadian yang mengakibatkan cedera yang tidak diharapkan pada pasien karena suatu tindakan (“commission”) atau karena tidak bertindak (“omission”), bukan karena “underlying disease” atau kondisi pasien.",
  "contoh": [
   "Semua reaksi transfusi yang sudah dikonfirmasi.\nKejadian reaksi transfusi pada saat kegiatan transfusi darah",
   "Semua kejadian serius akibat efek samping obat: Insiden serius akibat efek samping obat, contoh:\nObat yang salah sudah disuntikkan ke pasien, atau sehingga pasien mengalami syok atau reaksi alergi yang berat.",
   "Semua kesalahan pengobatan yang signifikan.\nContoh: Kesalahan dispensing obat oleh farmasi",
   "Semua perbedaan besar antara diagnosis pre-operasi dan post-operasi; Ketidaksesuaian diagnosa medik pre-operasi dan post-operasi.",
   "Efek samping atau pola efek samping selama sedasi moderat atau mendalam dan pemakaian anastesi.\nContohnya: Kejadian de-saturasi O2 pada saat durante anastesi pasien dengan general anasthesia.",
   "Kejadian-kejadian lain seperti kejadian infeksi rumah sakit.\nMisalnya: Identifikasi pasien potensial wabah berdasarkan gejala seperti Covid 19",
   "Pasien memutuskan isoman karena salah menerima hasil PCR positif yang seharusnya negatif.",
   "Terjadi infeksi pasca operasi fistel",
   "Terjadi phlebitis, terjadi ekstravasasi saat pemberian obat",
   "Pasien kejang karena obat tertunda",
   "Aspirasi pada pasien anak/dewasa saat di rawat di ruangan perawatan/IGD",
   "Pasien jatuh dari tempat tidur yang telah dipasang pagar pengaman baik itu dalam kondisi sadar maupun tidak",
   "Efek samping dari tindakan medis.\nMisalnya:\n- Dislokasi rahang saat pembersihan karang gigi\n- Perdarahan saat pemasangan kateter, atau dari selang CDL akibat malfungsi"
  ],
  "catatan": "Dispensing adalah rangkaian proses pelayanan kefarmasian yang dimulai dari penerimaan resep, pengkajian resep, penyiapan atau peracikan obat, pemberian etiket, pemeriksaan akhir, hingga penyerahan obat disertai pemberian informasi atau edukasi kepada pasien.",
  "order": 4
 },
 {
  "id": "ji_sentinel",
  "kode": "Sentinel",
  "nama": "Kejadian Sentinel",
  "inggris": "Sentinel Event",
  "definisi": "Kejadian sentinel adalah kejadian yang tidak berhubungan dengan perjalanan penyakit pasien/penyakit yang mendasarinya yang dapat menyebabkan kematian atau cedera permanen pada pasien. Pemilihan kata “sentinel” terkait dengan keseriusan cedera yang terjadi (misalnya amputasi pada kaki yang salah, dan sebagainya) sehingga pencarian fakta terhadap kejadian ini mengungkapkan adanya masalah yang serius pada kebijakan dan prosedur yang berlaku.",
  "contoh": [
   "Bunuh diri oleh pasien yang sedang dirawat / ditatalaksana / menerima pelayanan di unit yang selalu memiliki staf sepanjang hari, atau dalam waktu 72 jam (3 hari) setelah pemulangan pasien termasuk dari unit gawat darurat (UGD) RSUD.",
   "Kematian bayi cukup bulan yang tidak diantisipasi",
   "Bayi dipulangkan kepada orang tua yang salah",
   "Penculikan pasien yang sedang menerima perawatan, tata laksana, dan pelayanan",
   "Kaburnya pasien (atau pulang tanpa izin) dari unit perawatan yang selalu dijaga oleh staf sepanjang hari (termasuk UGD), yang menyebabkan kematian, cedera permanen, atau cedera sementara derajat berat bagi pasien tersebut;",
   "Reaksi transfusi hemolitik yang melibatkan pemberian darah atau produk darah dengan inkompatibilitas golongan darah mayor (ABO, Rh, kelompok darah lainnya)",
   "Pemerkosaan, kekerasan di tempat kerja seperti penyerangan (menyebabkan kematian atau kehilangan fungsi tubuh yang permanen) atau pembunuhan (yang disengaja) atas pasien, anggota staf, dokter, mahasiswa kedokteran, siswa latihan, pengunjung atau vendor pihak ketiga ketika berada dalam lingkungan rumah sakit",
   "Salah lokasi, salah prosedur, salah pasien ketika operasi.",
   "Tertinggalnya benda asing dalam tubuh pasien secara tidak sengaja setelah suatu tindakan invasif, termasuk saat operasi.",
   "Hiperbilirubinemia neonatal berat (bilirubin >30 mg/dL)",
   "Fluoroskopi berkepanjangan dengan dosis kumulatif >1.500 rad pada satu medan tunggal atau pemberian radioterapi ke area tubuh yang salah/pemberian radioterapi >25% melebihi dosis radioterapi yang direncanakan",
   "Kebakaran, lidah api/asap/uap panas, atau pijaran yang tidak diantisipasi selama satu episode perawatan pasien",
   "Semua kematian ibu intrapartum (terkait dengan proses persalinan); atau",
   "Morbiditas ibu derajat berat (terutama tidak berhubungan dengan perjalanan alamiah penyakit pasien atau kondisi lain yang mendasari) terjadi pada pasien dan menyebabkan cedera permanen atau cedera sementara derajat berat",
   "Penularan penyakit yang kronik atau fatal akibat infus darah atau produk darah atau transplantasi organ atau jaringan yang terkontaminasi;"
  ],
  "catatan": "",
  "order": 5
 }
];
function jenisInfoSheet() {
  const ss = getDbSpreadsheet();
  let s = ss.getSheetByName('Jenis_Insiden');
  if (!s) {
    s = ss.insertSheet('Jenis_Insiden');
    s.appendRow(JENIS_INFO_HEADER);
    s.getRange(1, 1, 1, JENIS_INFO_HEADER.length).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
    JENIS_INFO_SEED.forEach(j => s.appendRow([j.id, j.kode, j.nama, j.inggris, j.definisi, JSON.stringify(j.contoh), j.catatan, j.order]));
  }
  return s;
}
function jenisInfoRow(r) {
  let contoh = [];
  try { contoh = JSON.parse(r[5] || '[]'); } catch (e) {}
  return { id: r[0], kode: r[1], nama: r[2], inggris: r[3], definisi: r[4], contoh: contoh, catatan: r[6], order: Number(r[7]) || 0 };
}
function apiGetJenisInsiden(token) {
  try {
    auth(token);
    const d = jenisInfoSheet().getDataRange().getValues();
    const list = d.slice(1).filter(r => r[0]).map(jenisInfoRow);
    list.sort((a, b) => a.order - b.order);
    return ok({ jenis: list });
  } catch (e) { return fail(e); }
}
function apiSaveJenisInsiden(token, j) {
  return locked(() => {
    const a = auth(token, ['Admin']);
    j = j || {};
    const kode = T(j.kode, 40), nama = T(j.nama, 150), definisi = T(j.definisi, 3000);
    if (!kode) throw new Error('Kode jenis insiden wajib diisi.');
    if (!nama) throw new Error('Nama jenis insiden wajib diisi.');
    if (!definisi) throw new Error('Definisi wajib diisi.');
    const contoh = (Array.isArray(j.contoh) ? j.contoh : []).map(x => T(x, 1500)).filter(Boolean).slice(0, 60);
    if (!contoh.length) throw new Error('Minimal satu contoh kasus wajib diisi.');
    const s = jenisInfoSheet(), d = s.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) if (d[i][0] !== j.id && String(d[i][1]).toLowerCase() === kode.toLowerCase()) throw new Error('Kode \"' + kode + '\" sudah dipakai jenis insiden lain.');
    const row = [kode, nama, T(j.inggris, 150), definisi, JSON.stringify(contoh), T(j.catatan, 1500)];
    if (j.id) {
      for (let i = 1; i < d.length; i++) if (d[i][0] === j.id) {
        s.getRange(i + 1, 2, 1, 6).setValues([row]);
        audit('JENIS', a, 'Ubah jenis insiden', '', '', kode + ' - ' + nama);
        return ok({ message: 'Jenis insiden berhasil diperbarui.' });
      }
      throw new Error('Jenis insiden tidak ditemukan.');
    }
    const urutan = Math.max(0, ...d.slice(1).map(r => Number(r[7]) || 0)) + 1;
    s.appendRow(['ji_' + Date.now()].concat(row, [urutan]));
    audit('JENIS', a, 'Tambah jenis insiden', '', '', kode + ' - ' + nama);
    return ok({ message: 'Jenis insiden baru berhasil ditambahkan.' });
  });
}
function apiDeleteJenisInsiden(token, id) {
  return locked(() => {
    const a = auth(token, ['Admin']);
    const s = jenisInfoSheet(), d = s.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) if (d[i][0] === id) {
      audit('JENIS', a, 'Hapus jenis insiden', '', '', T(d[i][1], 40) + ' - ' + T(d[i][2], 150));
      s.deleteRow(i + 1);
      return ok({ message: 'Jenis insiden berhasil dihapus.' });
    }
    throw new Error('Jenis insiden tidak ditemukan.');
  });
}

// ---------- Laporan ----------
function jenisFromAnswer(t) {
  t = String(t || '');
  return /sentinel/i.test(t) ? 'Sentinel' : /KTD|tidak diharapkan/i.test(t) ? 'KTD' : /KTC|tidak cedera/i.test(t) ? 'KTC'
    : /KNC|nyaris/i.test(t) ? 'KNC' : /KPCS|KPC|potensi/i.test(t) ? 'KPCS' : '';
}
function findReport(id) {
  const s = sh('Laporan_IKP'), d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) if (d[i][0] === id) return { sheet: s, row: i + 1, r: d[i] };
  throw new Error('Laporan tidak ditemukan.');
}
function canSee(u, r) {
  return ROLES_VIEW_ALL.indexOf(u.role) >= 0 ||
    (u.role === 'Karu' && r[4] === u.ruangan) ||
    (u.role === 'Staf' && String(r[2]).toLowerCase() === u.username.toLowerCase());
}

function apiSubmitLaporanIKP(token, p) {
  return locked(() => {
    const u = auth(token);
    const dj = p.data_json || {};
    const qrows = sh('Form_Questions').getDataRange().getValues().slice(1);
    const miss = qrows.filter(r => r[2] !== 'heading' && (r[4] === true || r[4] === 'true') && !String(dj[r[1]] == null ? '' : dj[r[1]]).trim());
    if (miss.length) throw new Error('Isian wajib belum lengkap: ' + miss.map(r => r[1]).slice(0, 3).join('; '));
    const jq = qrows.find(r => r[0] === 'q_jenis_insiden');
    const pj = normJenis(p.jenis_insiden);
    let jenis = (pj && JENIS.indexOf(pj) >= 0) ? pj : '';
    if (!jenis && jq) jenis = jenisFromAnswer(dj[jq[1]]);
    if (!jenis) {
      for (const k in dj) {
        const candidate = jenisFromAnswer(dj[k]);
        if (candidate) { jenis = candidate; break; }
      }
    }
    if (!jenis) throw new Error('Jenis insiden belum dipilih.');
    const js = JSON.stringify(dj);
    if (js.length > 30000) throw new Error('Isian terlalu panjang.');
    const id = 'IKP-' + Date.now(), st = 'Menunggu Grading Karu';
    sh('Laporan_IKP').appendRow([id, new Date(), u.username, u.nama, u.ruangan, jenis, js, '', '', '', '', st, new Date(), u.username]);
    audit(id, u, 'Laporan dibuat', '', st, jenis);
    const tinggi = (jenis === 'KTD' || jenis === 'Sentinel');
    sendTelegramNotification(
      (tinggi ? '<b>INSIDEN BARU - RISIKO TINGGI (' + jenis + ')</b>' : '<b>LAPORAN INSIDEN BARU (' + jenis + ')</b>') + '\n' +
      'ID: ' + id + '\n' +
      'Unit: ' + escHtml(u.ruangan) + '\n' +
      'Pelapor: ' + escHtml(u.nama) + '\n' +
      'Jenis: ' + jenis + '\n' +
      'Waktu: ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') + '\n' +
      (tinggi ? 'Mohon segera ditindaklanjuti oleh Komite Mutu.' : 'Status: Menunggu Grading Karu.'),
      ['admin', 'komite']);
    if (tinggi) {
      sendEmailNotification('Laporan Risiko Tinggi (' + jenis + ') - ' + id,
        'Laporan insiden berisiko tinggi masuk dari unit <b>' + escHtml(u.ruangan) + '</b> oleh <b>' + escHtml(u.nama) + '</b> (jenis <b>' + jenis + '</b>). Membutuhkan tindakan cepat Komite Mutu.');
    }
    return ok({ message: 'Laporan berhasil dikirim dengan ID: ' + id, report_id: id });
  });
}

function apiSubmitGradingKaru(token, id, grading, investigasi) {
  return locked(() => {
    const u = auth(token, ['Karu']);
    if (GRADES.indexOf(grading) < 0) throw new Error('Grading tidak valid.');
    const f = findReport(id);
    if (f.r[4] !== u.ruangan) throw new Error('Laporan bukan dari unit Anda.');
    if (f.r[11] !== 'Menunggu Grading Karu') throw new Error('Laporan tidak sedang menunggu grading Karu.');
    const next = 'Menunggu Regrading Komite';
    f.sheet.getRange(f.row, 8, 1, 2).setValues([[grading, T(investigasi, 3000)]]);
    f.sheet.getRange(f.row, 12, 1, 3).setValues([[next, new Date(), u.username]]);
    audit(id, u, 'Grading Karu: ' + grading, f.r[11], next);
    if (grading === 'Kuning' || grading === 'Merah') {
      sendEmailNotification('Grading Karu ' + grading.toUpperCase() + ' - ' + id,
        'Kepala Ruangan memberi <b>Grading ' + grading + '</b> untuk laporan <b>' + id + '</b> dari unit <b>' + escHtml(f.r[4]) + '</b>. Mohon segera diproses oleh Komite Mutu.');
    }
    return ok({ message: 'Grading dan investigasi Karu berhasil disimpan.' });
  });
}

function apiSubmitRegradingKomite(token, id, regrading, rekomendasi) {
  return locked(() => {
    const u = auth(token, ROLES_KOMITE);
    if (GRADES.indexOf(regrading) < 0) throw new Error('Grading tidak valid.');
    const f = findReport(id);
    if (f.r[11] !== 'Menunggu Regrading Komite') throw new Error('Laporan tidak sedang menunggu regrading Komite.');
    const next = (regrading === 'Kuning' || regrading === 'Merah') ? ST_RCA : ST_SELESAI; // ditentukan server
    f.sheet.getRange(f.row, 10, 1, 2).setValues([[regrading, T(rekomendasi, 3000)]]);
    f.sheet.getRange(f.row, 12, 1, 3).setValues([[next, new Date(), u.username]]);
    audit(id, u, 'Regrading Komite: ' + regrading, f.r[11], next);
    return ok({ message: 'Regrading Komite berhasil disimpan. Status: ' + next });
  });
}

function apiSubmitRCA(token, d) {
  return locked(() => {
    const u = auth(token, ROLES_KOMITE);
    const f = findReport(d.report_id);
    const revisi = f.r[11] === ST_REVISI;
    if (f.r[11] !== ST_RCA && !revisi) throw new Error('Laporan tidak dalam tahap RCA.');
    if (!T(d.judul_rca) || !T(d.why1) || !T(d.root_cause) || !T(d.corrective_action) || !T(d.pic) || !d.target_selesai) {
      throw new Error('Field wajib RCA belum lengkap (judul, Why 1, akar masalah, tindakan korektif, PIC, dan tenggat).');
    }
    const rs = rcaSheet(), data = rs.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) if (data[i][1] === d.report_id) { rowIdx = i + 1; break; }
    if (rowIdx > 0 && !revisi) throw new Error('RCA untuk laporan ini sudah ada.');
    const rcaId = rowIdx > 0 ? data[rowIdx - 1][0] : 'RCA-' + Date.now();
    const vals = [rcaId, d.report_id, f.r[9] || f.r[7], T(d.judul_rca, 200), T(d.ringkasan_kronologi), T(d.masalah_utama, 500),
      T(d.why1, 500), T(d.why2, 500), T(d.why3, 500), T(d.why4, 500), T(d.why5, 500), T(d.root_cause), T(d.contributing_factors),
      T(d.corrective_action), T(d.preventive_action), T(d.pic, 100), String(d.target_selesai), ST_DIREKSI, new Date(), T(d.indikator_keberhasilan, 500)];
    if (rowIdx > 0) rs.getRange(rowIdx, 1, 1, vals.length).setValues([vals]); else rs.appendRow(vals);
    // RCA + rencana perbaikan langsung dilaporkan kepada Direksi untuk evaluasi, pembelajaran & umpan balik sistem.
    f.sheet.getRange(f.row, 12, 1, 3).setValues([[ST_DIREKSI, new Date(), u.username]]);
    audit(d.report_id, u, (revisi ? 'RCA direvisi dan rencana perbaikan dikirim ulang ke Direksi (' : 'RCA dan rencana perbaikan dilaporkan ke Direksi (') + rcaId + ')', f.r[11], ST_DIREKSI);
    const msg = 'ID: ' + d.report_id + ' | Unit: ' + escHtml(f.r[4]) + ' | Risiko: ' + escHtml(f.r[9] || f.r[7]);
    sendTelegramNotification(
      '<b>' + (revisi ? 'RCA REVISI' : 'RCA BARU') + ' - MENUNGGU REVIEW DIREKSI</b>\n' + msg + '\n' +
      'Judul RCA: ' + escHtml(T(d.judul_rca, 120)) + '\n' +
      'Akar masalah: ' + escHtml(T(d.root_cause, 200)) + '\n' +
      'Rencana perbaikan: ' + escHtml(T(d.corrective_action, 200)) + '\n' +
      'PIC: ' + escHtml(T(d.pic, 100)) + ' | Tenggat: ' + escHtml(String(d.target_selesai)) + '\n' +
      'Mohon Direksi mengevaluasi dan memberikan rekomendasi/umpan balik perbaikan sistem.',
      ['direksi', 'komite', 'admin']);
    sendEmailNotification('RCA menunggu review Direksi - ' + d.report_id,
      'Komite Mutu melaporkan hasil <b>RCA</b> dan <b>rencana perbaikan (corrective action)</b> untuk laporan <b>' + d.report_id + '</b> dari unit <b>' + escHtml(f.r[4]) +
      '</b>. Mohon Direktur melakukan evaluasi, pembelajaran, dan memberi rekomendasi perbaikan sistem melalui aplikasi e-IKP.', direkturEmails());
    return ok({ message: 'RCA dan rencana perbaikan berhasil disimpan dan dilaporkan kepada Direksi. Status: ' + ST_DIREKSI });
  });
}

// Direktur mengevaluasi RCA + rencana perbaikan: keputusan, evaluasi, pembelajaran organisasi, umpan balik perbaikan sistem.
function apiSubmitReviewDireksi(token, d) {
  return locked(() => {
    const u = auth(token, ROLES_DIREKSI);
    d = d || {};
    const f = findReport(d.report_id);
    if (f.r[11] !== ST_DIREKSI) throw new Error('Laporan tidak sedang menunggu review Direksi.');
    const keputusan = String(d.keputusan || '').trim();
    if (REVIEW_KEPUTUSAN.indexOf(keputusan) < 0) throw new Error('Keputusan review tidak valid.');
    const revisi = keputusan === 'Dikembalikan untuk Revisi';
    if (!T(d.evaluasi)) throw new Error('Evaluasi Direksi wajib diisi.');
    if (revisi && !T(d.rekomendasi_sistem)) throw new Error('Arahan revisi / umpan balik wajib diisi bila RCA dikembalikan.');
    if (!revisi && !T(d.pembelajaran)) throw new Error('Pembelajaran (lessons learned) wajib diisi.');
    let rcaId = '';
    rcaSheet().getDataRange().getValues().slice(1).forEach(r => { if (r[1] === d.report_id) rcaId = r[0]; });
    if (!rcaId) throw new Error('Data RCA untuk laporan ini belum ada.');
    const next = revisi ? ST_REVISI : ST_SELESAI;
    reviewSheet(true).appendRow(['RVW-' + Date.now(), d.report_id, rcaId, new Date(), u.username, u.nama, keputusan,
      T(d.evaluasi, 3000), T(d.pembelajaran, 3000), T(d.rekomendasi_sistem, 3000)]);
    f.sheet.getRange(f.row, 12, 1, 3).setValues([[next, new Date(), u.username]]);
    audit(d.report_id, u, 'Review Direksi: ' + keputusan, f.r[11], next);
    sendTelegramNotification(
      '<b>REVIEW DIREKSI: ' + escHtml(keputusan.toUpperCase()) + '</b>\n' +
      'ID: ' + d.report_id + ' | Unit: ' + escHtml(f.r[4]) + '\n' +
      (revisi ? 'RCA dikembalikan ke Komite Mutu untuk direvisi.\nArahan: ' + escHtml(T(d.rekomendasi_sistem, 250))
              : 'Pembelajaran: ' + escHtml(T(d.pembelajaran, 250)) + '\nLaporan ditutup (Selesai).'),
      ['komite', 'admin', 'direksi']);
    return ok({ message: revisi ? 'RCA dikembalikan ke Komite Mutu untuk revisi.' : 'Review Direksi tersimpan. Laporan dinyatakan Selesai.' });
  });
}

function apiGetReviewDireksi(token, id) {
  try {
    const u = auth(token), f = findReport(id);
    if (!canSee(u, f.r)) throw new Error('Akses ditolak.');
    const list = reviewRows(id);
    return ok({ data: list.length ? list[list.length - 1] : null, history: list });
  } catch (e) { return fail(e); }
}

function apiGetRCAByReportId(token, id) {
  try {
    const u = auth(token), f = findReport(id);
    if (!canSee(u, f.r)) throw new Error('Akses ditolak.');
    const d = rcaSheet().getDataRange().getValues();
    for (let i = 1; i < d.length; i++) if (d[i][1] === id) {
      const o = {};
      RCA_K.forEach((k, j) => o[k] = fmt(d[i][j], k === 'target_selesai' ? 'dd/MM/yyyy' : null));
      o.status = o.status || '';
      return ok({ data: o });
    }
    return { success: false, message: 'Belum ada data RCA.' };
  } catch (e) { return fail(e); }
}

function apiGetAudit(token, id) {
  try {
    auth(token, ['Komite', 'KetuaKomite', 'Manajemen', 'Direktur', 'Admin']);
    const d = sh('Audit_Trail').getDataRange().getValues();
    return ok({ data: d.slice(1).filter(r => r[1] === id).map(r => ({ timestamp: fmt(r[2]), actor_name: r[4], actor_role: r[5], action: r[6], from_status: r[7], to_status: r[8], details: r[9] })) });
  } catch (e) { return fail(e); }
}

// ---------- Dashboard (filter periode & tren bulanan nyata) ----------
function apiGetDashboardData(token, months) {
  try {
    const u = auth(token);
    months = [1, 3, 6, 12].indexOf(Number(months)) < 0 ? 3 : Number(months);
    const now = new Date(), from = new Date(now.getFullYear(), now.getMonth() - months + 1, 1);
    const labels = [], idx = {};
    for (let k = 0; k < months; k++) {
      const d = new Date(from.getFullYear(), from.getMonth() + k, 1);
      idx[d.getFullYear() + '-' + d.getMonth()] = k;
      labels.push(BLN[d.getMonth()] + ' ' + String(d.getFullYear()).slice(2));
    }
    const tren = labels.map(() => 0), jc = { KPCS: 0, KNC: 0, KTC: 0, KTD: 0, Sentinel: 0 };
    const jg = {}; ['Biru', 'Hijau', 'Kuning', 'Merah', 'Belum'].forEach(k => jg[k] = JENIS.map(() => 0));
    const cqs = sh('Form_Questions').getDataRange().getValues().slice(1)
      .filter(r => (r[2] === 'radio' || r[2] === 'select') && r[0] !== 'q_jenis_insiden')
      .sort((x, y) => x[5] - y[5])
      .map(r => { const o = String(r[3]).split(',').map(x => x.trim()).filter(Boolean); return { id: r[0], label: r[1], opts: o, counts: o.map(() => 0), other: 0, empty: 0 }; });
    const uc = {};
    // Nama Kepala Ruangan per unit (dari sheet Users: peran Karu, status Approved)
    const karuMap = {}, karuNip = {}, pelaporNip = {}, userInfo = {};
    let ketua = { nama: '', nip: '' };
    sh('Users').getDataRange().getValues().slice(1).forEach(r => {
      if (String(r[5]).trim().toLowerCase() !== 'approved') return;
      pelaporNip[String(r[0]).trim().toLowerCase()] = String(r[7] || '').trim();
      userInfo[String(r[0]).trim().toLowerCase()] = { nama: String(r[2]).trim(), nip: String(r[7] || '').trim(), role: String(r[3]).trim() };
      if (String(r[3]).trim() === 'Karu') {
        const un = String(r[4]).trim();
        if (un && !karuMap[un]) { karuMap[un] = String(r[2]).trim(); karuNip[un] = String(r[7] || '').trim(); }
      }
      if (String(r[3]).trim() === 'KetuaKomite' && !ketua.nama) ketua = { nama: String(r[2]).trim(), nip: String(r[7] || '').trim() };
    });
    // Petugas Komite Mutu = pengguna berperan 'Komite' yang terakhir memproses laporan (regrading/RCA), dibaca dari updated_by
    const kom = r => { if (GRADES.indexOf(r[9]) < 0) return null; const x = userInfo[String(r[13]).trim().toLowerCase()]; return x && x.role === 'Komite' ? x : null; };
    // Review Direksi terakhir per laporan (untuk status alur kerja, panel detail, dan lembar cetak)
    const revMap = {};
    try {
      const rsh = reviewSheet(false);
      if (rsh && rsh.getLastRow() > 1) rsh.getDataRange().getValues().slice(1).forEach(r => { revMap[r[1]] = { keputusan: r[6], nama: r[5], waktu: fmt(r[3]) }; });
    } catch (e) {}
    const m = { total: 0, karu: 0, komite: 0, direksi: 0, selesai: 0 }, out = [];
    sh('Laporan_IKP').getDataRange().getValues().slice(1).forEach(r => {
      const t = new Date(r[1]);
      if (isNaN(t) || t < from || !canSee(u, r)) return;
      r[5] = normJenis(r[5]);   // data lama 'KPC' dibaca sebagai 'KPCS'
      let j = {}; try { j = JSON.parse(r[6]); } catch (e) {}
      out.push({ id: r[0], timestamp: fmt(t), username_pelapor: r[2], nama_pelapor: r[3], ruangan: r[4], jenis_insiden: r[5], data_json: j,
        grading_karu: r[7], investigasi_sederhana_karu: r[8], regrading_komite: r[9], rekomendasi_komite: r[10], status: r[11], updated_at: fmt(r[12]), updated_by: r[13],
        nama_karu: karuMap[String(r[4]).trim()] || '', nip_karu: karuNip[String(r[4]).trim()] || '', nama_ketua: ketua.nama, nip_ketua: ketua.nip,
        nip_pelapor: pelaporNip[String(r[2]).trim().toLowerCase()] || '',
        keputusan_direksi: (revMap[r[0]] || {}).keputusan || '', nama_direktur: (revMap[r[0]] || {}).nama || '', tgl_review_direksi: (revMap[r[0]] || {}).waktu || '',
        nama_komite: (kom(r) || {}).nama || '', nip_komite: (kom(r) || {}).nip || '' });
      m.total++;
      if (r[11] === 'Menunggu Grading Karu') m.karu++;
      if (r[11] === ST_KOMITE || r[11] === ST_RCA || r[11] === ST_REVISI) m.komite++;
      if (r[11] === ST_DIREKSI) m.direksi++;
      if (r[11] === ST_SELESAI) m.selesai++;
      if (jc[r[5]] !== undefined) jc[r[5]]++;
      const gr = GRADES.indexOf(r[9]) >= 0 ? r[9] : (GRADES.indexOf(r[7]) >= 0 ? r[7] : 'Belum'); // grading akhir: Komite, jika belum maka Karu
      if (JENIS.indexOf(r[5]) >= 0) jg[gr][JENIS.indexOf(r[5])]++;
      cqs.forEach(c => { const v = String(j[c.label] == null ? '' : j[c.label]).trim(), i = c.opts.indexOf(v); if (!v) c.empty++; else if (i >= 0) c.counts[i]++; else c.other++; });
      uc[r[4]] = (uc[r[4]] || 0) + 1;
      tren[idx[t.getFullYear() + '-' + t.getMonth()]]++;
    });
    const keep = JENIS.map(x => x !== 'KPCS' || jc.KPCS > 0); // KPCS tidak ada di formulir internal; tampil hanya jika ada datanya
    const jl = JENIS.filter((x, i) => keep[i]);
    Object.keys(jg).forEach(k => jg[k] = jg[k].filter((v, i) => keep[i]));
    const qCharts = cqs.map(c => { const l = c.opts.slice(), n = c.counts.slice(); if (c.other) { l.push('Lainnya'); n.push(c.other); } if (c.empty) { l.push('Tidak diisi'); n.push(c.empty); } return { id: c.id, label: c.label, labels: l, counts: n }; })
      .filter(c => c.counts.some(v => v > 0));
    const ul = Object.keys(uc).sort((x, y) => uc[y] - uc[x]);
    if (ul.length) qCharts.push({ id: 'unit', label: 'Jumlah Laporan per Unit / Ruangan', labels: ul, counts: ul.map(k => uc[k]) });
    out.reverse();
    return ok({ data: out, metrics: m, chartData: { jenisLabels: jl, jenisByGrade: jg, jenisCounts: jc, qCharts: qCharts, trenLabels: labels, trenCounts: tren } });
  } catch (e) { return fail(e); }
}

// ---------- Tanda tangan digital ----------
const SLOT_TTD = ['karu', 'komite', 'ketua'];
function isPngDataUrl(v) {
  v = String(v || '');
  if (v.length > 45000 || !/^data:image\/png;base64,[A-Za-z0-9+\/=]+$/.test(v)) return false;
  try { const b = Utilities.base64Decode(v.split(',')[1]); return b.length > 8 && (b[0] & 255) === 137 && b[1] === 80 && b[2] === 78 && b[3] === 71; } catch (e) { return false; }
}
function ttdSheet() {
  const ss = getDbSpreadsheet();
  let s = ss.getSheetByName('TTD_Dokumen');
  if (!s) {
    s = ss.insertSheet('TTD_Dokumen');
    const h = ['ttd_id', 'report_id', 'slot', 'username', 'nama', 'role', 'timestamp', 'image'];
    s.appendRow(h);
    s.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  }
  return s;
}
// Simpan / hapus tanda tangan milik pengguna yang sedang login (image kosong = hapus)
function apiSaveMySignature(token, image) {
  return locked(() => {
    const u = auth(token), us = sh('Users');
    if (!String(us.getRange(1, 9).getValue()).trim()) us.getRange(1, 9).setValue('ttd').setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
    if (!image) {
      us.getRange(u.row, 9).setValue('');
      audit('USER:' + u.username, u, 'Hapus tanda tangan digital');
      return ok({ message: 'Tanda tangan tersimpan telah dihapus.' });
    }
    if (!isPngDataUrl(image)) throw new Error('Gambar tanda tangan tidak valid atau terlalu besar. Coba gambar lebih sederhana.');
    us.getRange(u.row, 9).setNumberFormat('@').setValue(image);
    audit('USER:' + u.username, u, 'Simpan tanda tangan digital');
    return ok({ message: 'Tanda tangan digital berhasil disimpan.' });
  });
}
function apiGetMySignature(token) {
  try { return ok({ image: auth(token).ttd || '' }); } catch (e) { return fail(e); }
}
function signSlotFor(u, r) {
  if (u.role === 'Karu' && u.ruangan === r[4] && GRADES.indexOf(r[7]) >= 0) return 'karu';
  if (r[11] === 'Selesai' && GRADES.indexOf(r[9]) >= 0) {
    if (u.role === 'Komite') return 'komite';
    if (u.role === 'KetuaKomite') return 'ketua';
  }
  return '';
}
function readSigRows(id) {
  const d = ttdSheet().getDataRange().getValues(), out = [];
  for (let i = 1; i < d.length; i++) if (d[i][1] === id) out.push({ row: i + 1, slot: d[i][2], username: d[i][3], nama: d[i][4], role: d[i][5], waktu: d[i][6], image: d[i][7] });
  return out;
}
function apiGetSignatures(token, id) {
  try {
    const u = auth(token), f = findReport(id);
    if (!canSee(u, f.r)) throw new Error('Akses ditolak.');
    const sigs = {};
    // Staf pelapor: otomatis memakai tanda tangan pelapor (jika sudah tersimpan)
    const rep = findUser(f.r[2]);
    if (rep && rep.ttd && isPngDataUrl(rep.ttd)) sigs.staf = { nama: rep.nama, nip: rep.nip, waktu: fmt(f.r[1]), image: rep.ttd, auto: true };
    let mine = '';
    readSigRows(id).forEach(x => {
      if (SLOT_TTD.indexOf(x.slot) < 0 || !isPngDataUrl(x.image)) return;
      const su = findUser(x.username);
      sigs[x.slot] = { nama: x.nama, nip: su ? su.nip : '', waktu: fmt(x.waktu), image: x.image };
      if (String(x.username).toLowerCase() === u.username.toLowerCase()) mine = x.slot;
    });
    const slot = signSlotFor(u, f.r);
    return ok({ signatures: sigs, can_sign: (slot && !sigs[slot]) ? slot : '', my_slot: mine, has_ttd: !!u.ttd });
  } catch (e) { return fail(e); }
}
function apiSignDocument(token, id, slot) {
  return locked(() => {
    const u = auth(token), f = findReport(id);
    if (!canSee(u, f.r)) throw new Error('Akses ditolak.');
    if (SLOT_TTD.indexOf(slot) < 0 || signSlotFor(u, f.r) !== slot) throw new Error('Anda tidak berwenang menandatangani bagian ini, atau dokumen belum siap ditandatangani.' + ((slot === 'komite' || slot === 'ketua') ? ' (Lembar Komite hanya dapat ditandatangani setelah status laporan Selesai.)' : ''));
    if (!isPngDataUrl(u.ttd)) throw new Error('Anda belum menyimpan tanda tangan digital. Buka menu profil > Tanda Tangan Digital Saya.');
    if (readSigRows(id).some(x => x.slot === slot)) throw new Error('Bagian ini sudah ditandatangani.');
    ttdSheet().appendRow(['TTD-' + Utilities.getUuid().slice(0, 8), id, slot, u.username, u.nama, u.role, new Date(), u.ttd]);
    audit(id, u, 'Tanda tangan digital: ' + slot, '', '', u.nama);
    return ok({ message: 'Dokumen berhasil ditandatangani secara digital.' });
  });
}
function apiRemoveSignature(token, id) {
  return locked(() => {
    const u = auth(token), f = findReport(id);
    if (!canSee(u, f.r)) throw new Error('Akses ditolak.');
    const mine = readSigRows(id).filter(x => String(x.username).toLowerCase() === u.username.toLowerCase());
    if (!mine.length) throw new Error('Anda belum menandatangani dokumen ini.');
    mine.sort((a, b) => b.row - a.row).forEach(x => ttdSheet().deleteRow(x.row));
    audit(id, u, 'Batalkan tanda tangan digital', '', '', u.nama);
    return ok({ message: 'Tanda tangan Anda pada dokumen ini telah dibatalkan.' });
  });
}

// ---------- Email ----------
function sendEmailNotification(subject, bodyContent, extraTo) {
  try {
    let to = 'komitemutu@rsudmerauke.go.id';
    const s = sh('App_Settings');
    if (s) s.getDataRange().getValues().forEach(r => { if (r[0] === 'komite_email' && r[1]) to = r[1]; });
    if (extraTo) to = to + ',' + extraTo;   // mis. email Direktur untuk laporan RCA yang menunggu review
    MailApp.sendEmail({ to: to, subject: '[ALERT IKP RSUD MERAUKE] ' + subject,
      htmlBody: '<div style="font-family:Arial,sans-serif;padding:20px;border:1px solid #e2e8f0;border-radius:8px"><h2 style="color:#1e3a8a">Pemberitahuan Sistem e-IKP RSUD Merauke</h2><p>' + bodyContent + '</p><hr><small style="color:#64748b">Pesan otomatis Sistem e-IKP RSUD Merauke.</small></div>' });
  } catch (err) { Logger.log('Gagal mengirim email: ' + err.message); }
}

// ---------- Lampiran (unggah dokumen manual per laporan) ----------
// File disimpan PRIVAT di folder Drive "Lampiran e-IKP RSUD Merauke" (milik akun pemilik script);
// akses unduh hanya lewat server setelah cek hak lihat laporan (canSee). Metadata di sheet Lampiran_IKP.
const LAMPIRAN_MAX_BYTES = 5 * 1024 * 1024;   // 5 MB per file
const LAMPIRAN_MAX_PER_LAPORAN = 10;
const LAMPIRAN_KATEGORI = ['Lembar Pelaporan', 'Investigasi Sederhana', 'Lainnya'];
const LAMPIRAN_UPLOAD_ROLES = ['Staf', 'Karu', 'Komite', 'KetuaKomite', 'Admin'];
const LAMPIRAN_KATEGORI_KARU = ['Karu', 'Komite', 'KetuaKomite', 'Admin'];  // peran yang boleh mengunggah "Investigasi Sederhana"
const LAMPIRAN_HEADER = ['lampiran_id', 'report_id', 'kategori', 'nama_file', 'mime', 'ukuran', 'file_id', 'username', 'nama', 'role', 'timestamp'];

function lampiranSheet() {
  const ss = getDbSpreadsheet();
  let s = ss.getSheetByName('Lampiran_IKP');
  if (!s) {
    s = ss.insertSheet('Lampiran_IKP');
    s.appendRow(LAMPIRAN_HEADER);
    s.getRange(1, 1, 1, LAMPIRAN_HEADER.length).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  }
  return s;
}
function lampiranFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('LAMPIRAN_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = DriveApp.getFoldersByName('Lampiran e-IKP RSUD Merauke');
  const fo = it.hasNext() ? it.next() : DriveApp.createFolder('Lampiran e-IKP RSUD Merauke');
  props.setProperty('LAMPIRAN_FOLDER_ID', fo.getId());
  return fo;
}
// Tipe file ditentukan dari isi (magic bytes), bukan dari nama/mime kiriman klien.
function sniffLampiran(bytes) {
  const b = bytes.map(x => (x < 0 ? x + 256 : x));
  if (b.length > 4 && b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46) return { mime: 'application/pdf', ext: 'pdf' };
  if (b.length > 3 && b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF) return { mime: 'image/jpeg', ext: 'jpg' };
  if (b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return { mime: 'image/png', ext: 'png' };
  return null;
}
function lampiranRows(reportId) {
  const d = lampiranSheet().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < d.length; i++) if (d[i][1] === reportId) out.push({ row: i + 1, r: d[i] });
  return out;
}
function lampiranPub(r) {
  return { lampiran_id: r[0], report_id: r[1], kategori: r[2], nama_file: r[3], mime: r[4], ukuran: Number(r[5]) || 0,
    nama: r[8], role: r[9], timestamp: fmt(r[10]), username: r[7] };
}

function apiGetLampiran(token, id) {
  try {
    const u = auth(token), f = findReport(id);
    if (!canSee(u, f.r)) throw new Error('Anda tidak berhak melihat laporan ini.');
    const me = u.username.toLowerCase();
    const list = lampiranRows(id).map(x => {
      const p = lampiranPub(x.r);
      p.can_delete = (String(x.r[7]).toLowerCase() === me && LAMPIRAN_UPLOAD_ROLES.indexOf(u.role) >= 0 &&
        !((u.role === 'Staf' || u.role === 'Karu') && f.r[11] === 'Selesai')) || u.role === 'Admin';
      return p;
    });
    const canUp = LAMPIRAN_UPLOAD_ROLES.indexOf(u.role) >= 0 && !((u.role === 'Staf' || u.role === 'Karu') && f.r[11] === 'Selesai');
    return ok({ files: list, can_upload: canUp, can_upload_investigasi: LAMPIRAN_KATEGORI_KARU.indexOf(u.role) >= 0,
      max_mb: LAMPIRAN_MAX_BYTES / 1048576, max_files: LAMPIRAN_MAX_PER_LAPORAN });
  } catch (e) { return fail(e); }
}

// base64 = isi file tanpa prefix "data:...;base64,"
function apiUploadLampiran(token, id, kategori, namaFile, base64) {
  return locked(() => {
    const u = auth(token, LAMPIRAN_UPLOAD_ROLES), f = findReport(id);
    if (!canSee(u, f.r)) throw new Error('Anda tidak berhak mengunggah ke laporan ini.');
    if ((u.role === 'Staf' || u.role === 'Karu') && f.r[11] === 'Selesai') throw new Error('Laporan sudah Selesai; unggahan hanya dapat dilakukan Komite/Admin.');
    if (LAMPIRAN_KATEGORI.indexOf(kategori) < 0) throw new Error('Kategori lampiran tidak valid.');
    if (kategori === 'Investigasi Sederhana' && LAMPIRAN_KATEGORI_KARU.indexOf(u.role) < 0) throw new Error('Lampiran Investigasi Sederhana hanya dapat diunggah Karu/Komite/Admin.');
    if (lampiranRows(id).length >= LAMPIRAN_MAX_PER_LAPORAN) throw new Error('Batas ' + LAMPIRAN_MAX_PER_LAPORAN + ' lampiran per laporan tercapai.');
    const b64 = String(base64 || '').replace(/^data:[^,]*,/, '').replace(/\s+/g, '');
    if (!b64) throw new Error('File kosong.');
    if (b64.length > Math.ceil(LAMPIRAN_MAX_BYTES * 4 / 3) + 8) throw new Error('Ukuran file melebihi ' + (LAMPIRAN_MAX_BYTES / 1048576) + ' MB.');
    const bytes = Utilities.base64Decode(b64);
    if (!bytes.length || bytes.length > LAMPIRAN_MAX_BYTES) throw new Error('Ukuran file melebihi ' + (LAMPIRAN_MAX_BYTES / 1048576) + ' MB.');
    const kind = sniffLampiran(bytes);
    if (!kind) throw new Error('Format file tidak didukung. Gunakan PDF, JPG, atau PNG.');
    let nm = String(namaFile || 'lampiran').replace(/[\\\/:*?"<>|\r\n]+/g, '_').replace(/\.[A-Za-z0-9]{1,5}$/, '').trim().slice(0, 80) || 'lampiran';
    nm = nm + '.' + kind.ext;
    const lid = 'LMP-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
    const blob = Utilities.newBlob(bytes, kind.mime, id + '_' + lid + '_' + nm);
    const file = lampiranFolder().createFile(blob);
    lampiranSheet().appendRow([lid, id, kategori, nm, kind.mime, bytes.length, file.getId(), u.username, u.nama, u.role, new Date()]);
    audit(id, u, 'Unggah lampiran: ' + kategori, '', '', nm + ' (' + Math.round(bytes.length / 1024) + ' KB)');
    // Notifikasi Telegram ke Komite & Admin bila lampiran diunggah oleh Staf/Karu (unggahan Komite/Admin sendiri tidak diberitahukan).
    if (u.role === 'Staf' || u.role === 'Karu') {
      sendTelegramNotification(
        '<b>LAMPIRAN BARU DIUNGGAH</b>\n' +
        'ID Laporan: ' + escHtml(id) + '\n' +
        'Unit: ' + escHtml(f.r[4]) + '\n' +
        'Kategori: ' + escHtml(kategori) + '\n' +
        'File: ' + escHtml(nm) + ' (' + Math.round(bytes.length / 1024) + ' KB)\n' +
        'Diunggah oleh: ' + escHtml(u.nama) + ' (' + escHtml(u.role) + ')\n' +
        'Waktu: ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'),
        ['admin', 'komite']);
    }
    return ok({ message: 'Lampiran berhasil diunggah.', lampiran_id: lid });
  });
}

function apiDownloadLampiran(token, lampiranId) {
  try {
    const u = auth(token);
    const d = lampiranSheet().getDataRange().getValues();
    const row = d.slice(1).find(r => r[0] === lampiranId);
    if (!row) throw new Error('Lampiran tidak ditemukan.');
    const f = findReport(row[1]);
    if (!canSee(u, f.r)) throw new Error('Anda tidak berhak mengunduh lampiran ini.');
    const blob = DriveApp.getFileById(row[6]).getBlob();
    return ok({ nama_file: row[3], mime: row[4], data: Utilities.base64Encode(blob.getBytes()) });
  } catch (e) { return fail(e); }
}

function apiDeleteLampiran(token, lampiranId) {
  return locked(() => {
    const u = auth(token), s = lampiranSheet(), d = s.getDataRange().getValues();
    let idx = -1;
    for (let i = 1; i < d.length; i++) if (d[i][0] === lampiranId) { idx = i; break; }
    if (idx < 0) throw new Error('Lampiran tidak ditemukan.');
    const row = d[idx], f = findReport(row[1]);
    const own = String(row[7]).toLowerCase() === u.username.toLowerCase() && LAMPIRAN_UPLOAD_ROLES.indexOf(u.role) >= 0;
    if (!(u.role === 'Admin' || (own && canSee(u, f.r) && !((u.role === 'Staf' || u.role === 'Karu') && f.r[11] === 'Selesai')))) {
      throw new Error('Anda tidak berwenang menghapus lampiran ini.');
    }
    try { DriveApp.getFileById(row[6]).setTrashed(true); } catch (e) { Logger.log('Gagal menghapus file Drive: ' + e.message); }
    s.deleteRow(idx + 1);
    audit(row[1], u, 'Hapus lampiran', '', '', row[3]);
    return ok({ message: 'Lampiran dihapus.' });
  });
}

// ---------- Telegram (Bot API) ----------
// Konfigurasi (sekali saja, jalankan dari editor Apps Script):
//   setupTelegram('TOKEN_BOT', 'CHAT_ID_KOMITE', 'CHAT_ID_ADMIN', 'CHAT_ID_DIREKSI')   // argumen ke-4 (Direksi) opsional
//   -> argumen 2 = chat ID Komite Mutu, argumen 3 = chat ID Admin (boleh banyak, pisah koma;
//      boleh chat ID pribadi atau grup; ID grup diawali tanda minus).
//   Token disimpan di Script Properties; chat ID disimpan di sheet App_Settings.
// Kegagalan kirim Telegram TIDAK menggagalkan pendaftaran/laporan (hanya dicatat di log).
function setupTelegram(token, chatKomite, chatAdmin, chatDireksi) {
  if (token) PropertiesService.getScriptProperties().setProperty('TG_TOKEN', String(token).trim());
  const s = sh('App_Settings');
  const setKey = (k, v) => {
    const d = s.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) if (d[i][0] === k) { s.getRange(i + 1, 2).setNumberFormat('@').setValue(v); return; }
    s.appendRow([k, '']);
    s.getRange(s.getLastRow(), 1, 1, 2).setNumberFormat('@').setValues([[k, v]]);
  };
  if (chatKomite != null) setKey('tg_komite', String(chatKomite));
  if (chatAdmin != null) setKey('tg_admin', String(chatAdmin));
  if (chatDireksi != null) setKey('tg_direksi', String(chatDireksi));   // opsional: chat ID Direksi (RCA & review)
  return 'Telegram dikonfigurasi. Chat ID Komite: ' + chatKomite + ' | Admin: ' + chatAdmin + ' | Direksi: ' + chatDireksi;
}
// Pesan memakai HTML Telegram (<b>); isi dinamis WAJIB di-escape dengan escHtml().
function sendTelegramNotification(message, groups) {
  try {
    if (!String(message || '').trim()) { Logger.log('Telegram dilewati: isi pesan kosong (fungsi ini dipanggil tanpa pesan; jalankan ujiTelegram untuk menguji).'); return; }
    const token = PropertiesService.getScriptProperties().getProperty('TG_TOKEN');
    if (!token) { Logger.log('Telegram dilewati: TG_TOKEN belum diatur (jalankan setupTelegram).'); return; }
    const cfg = {}, s = sh('App_Settings');
    if (s) s.getDataRange().getValues().forEach(r => { cfg[String(r[0])] = String(r[1] || ''); });
    const targets = {};
    (groups || ['admin', 'komite']).forEach(g => {
      String(cfg['tg_' + g] || '').split(/[,;\s]+/).forEach(id => { if (/^-?\d+$/.test(id)) targets[id] = true; });
    });
    const list = Object.keys(targets);
    if (!list.length) { Logger.log('Telegram dilewati: chat ID tujuan belum diatur.'); return; }
    list.forEach(chatId => {
      const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ chat_id: chatId, text: message, parse_mode: 'HTML', disable_web_page_preview: true }),
        muteHttpExceptions: true
      });
      Logger.log('Telegram -> ' + chatId + ': ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
    });
  } catch (err) { Logger.log('Gagal mengirim Telegram: ' + err.message); }
}

// ---------- Notifikasi Telegram otomatis untuk semua aktivitas (dipanggil dari audit) ----------
// Aksi di bawah ini TIDAK dikirim lewat sini karena sudah punya pesan Telegram sendiri yang lebih lengkap
// (Pendaftaran akun, Laporan dibuat, Unggah lampiran), atau terlalu pribadi/berisik (ubah profil, tanda tangan tersimpan).
const TG_SKIP = ['Pendaftaran akun', 'Laporan dibuat', 'Ubah profil', 'Simpan tanda tangan digital', 'Hapus tanda tangan digital'];
// Aksi RCA & Review Direksi sudah punya pesan Telegram khusus di apiSubmitRCA / apiSubmitReviewDireksi.
const TG_SKIP_PREFIX = ['Unggah lampiran', 'RCA dan rencana perbaikan dilaporkan', 'RCA direvisi', 'Review Direksi'];
function notifAktivitas(id, u, action, from, to, details) {
  action = String(action || '');
  if (TG_SKIP.indexOf(action) >= 0 || TG_SKIP_PREFIX.some(p => action.indexOf(p) === 0)) return;
  id = String(id || '');
  const isLaporan = id.indexOf('IKP-') === 0;
  const isAkun = id.indexOf('USER:') === 0;
  const icon = isLaporan ? '\uD83D\uDCCB' : (isAkun ? '\uD83D\uDC64' : '\u2699\uFE0F');
  let msg = icon + ' <b>AKTIVITAS BARU - e-IKP RSUD Merauke</b>\n' +
    'Aksi: <b>' + escHtml(action) + '</b>\n';
  if (isLaporan) {
    msg += 'ID Laporan: ' + escHtml(id) + '\n';
    try { const f = findReport(id); msg += 'Unit: ' + escHtml(f.r[4]) + ' | Jenis: ' + escHtml(f.r[5]) + '\n'; } catch (e) {}
  } else if (isAkun) {
    msg += 'Akun: ' + escHtml(id.slice(5)) + '\n';
  }
  if (from || to) msg += 'Status: ' + escHtml(from || '-') + ' \u2192 ' + escHtml(to || '-') + '\n';
  if (details) msg += 'Keterangan: ' + escHtml(String(details).slice(0, 300)) + '\n';
  msg += 'Oleh: ' + escHtml(u.nama || u.username || '-') + (u.role ? ' (' + escHtml(u.role) + ')' : '') + '\n' +
    'Waktu: ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
  // Alur laporan -> Komite + Admin; pengelolaan akun/formulir/keamanan -> Admin saja.
  const keDireksi = isLaporan && (action.indexOf('RCA') === 0 || action.indexOf('Review Direksi') === 0);
  sendTelegramNotification(msg, isLaporan ? (keDireksi ? ['admin', 'komite', 'direksi'] : ['admin', 'komite']) : ['admin']);
}

// JALANKAN dari editor untuk menguji: cek token, cek chat ID, kirim pesan uji, lalu lihat hasilnya di Eksekusi > Log.
function ujiTelegram() {
  const token = PropertiesService.getScriptProperties().getProperty('TG_TOKEN');
  if (!token) { Logger.log('GAGAL: TG_TOKEN belum diatur. Jalankan setupTelegram(\'TOKEN_BOT\', null, null) dulu.'); return; }
  const me = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getMe', { muteHttpExceptions: true }).getContentText());
  Logger.log(me.ok ? 'Token VALID. Bot: @' + me.result.username : 'Token TIDAK VALID: ' + JSON.stringify(me));
  const cfg = {}, s = sh('App_Settings');
  if (s) s.getDataRange().getValues().forEach(r => { cfg[String(r[0])] = String(r[1] || ''); });
  Logger.log('Chat ID admin: ' + (cfg.tg_admin || '(kosong)') + ' | komite: ' + (cfg.tg_komite || '(kosong)'));
  sendTelegramNotification('\u2705 <b>Uji notifikasi e-IKP RSUD Merauke</b>\nJika pesan ini sampai, Telegram sudah berfungsi.\nWaktu: ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('Selesai. Periksa baris "Telegram -> ..." di atas: kode 200 = terkirim; 400/403 = chat ID salah atau bot belum di-/start / belum masuk grup.');
}

// ---------- Setup & migrasi (jalankan manual dari editor) ----------
function setupDatabase(targetSpreadsheetId) {
  if (targetSpreadsheetId) {
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', String(targetSpreadsheetId).trim());
  }
  const ss = getDbSpreadsheet();
  const defs = {
    Users: ['username', 'password', 'nama', 'role', 'ruangan', 'status', 'email', 'nip', 'ttd'],
    Form_Questions: ['id', 'label', 'type', 'options', 'required', 'order'],
    Laporan_IKP: ['id', 'timestamp', 'username_pelapor', 'nama_pelapor', 'ruangan', 'jenis_insiden', 'data_json', 'grading_karu', 'investigasi_sederhana_karu', 'regrading_komite', 'rekomendasi_komite', 'status', 'updated_at', 'updated_by'],
    RCA_IKP: ['rca_id', 'report_id', 'risk_level', 'judul_rca', 'ringkasan_kronologi', 'masalah_utama', 'why1', 'why2', 'why3', 'why4', 'why5', 'root_cause', 'contributing_factors', 'corrective_action', 'preventive_action', 'pic', 'target_selesai', 'status', 'created_at', 'indikator_keberhasilan'],
    Audit_Trail: ['audit_id', 'report_id', 'timestamp', 'actor_username', 'actor_name', 'actor_role', 'action', 'from_status', 'to_status', 'details'],
    App_Settings: ['key', 'value'],
    TTD_Dokumen: ['ttd_id', 'report_id', 'slot', 'username', 'nama', 'role', 'timestamp', 'image'],
    Lampiran_IKP: ['lampiran_id', 'report_id', 'kategori', 'nama_file', 'mime', 'ukuran', 'file_id', 'username', 'nama', 'role', 'timestamp'],
    Review_Direksi: REVIEW_HEADER
  };
  Object.keys(defs).forEach(n => {
    if (ss.getSheetByName(n)) return;
    const s = ss.insertSheet(n); s.appendRow(defs[n]);
    s.getRange(1, 1, 1, defs[n].length).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  });
  const us = ss.getSheetByName('Users');
  if (us.getLastRow() === 1) {
    // Inisialisasi akun default dengan password ter-hash
    us.appendRow(['admin', hashPw('admin123'), 'Administrator Utama', 'Admin', 'Komite Mutu', 'Approved', 'komitemutu@rsudmerauke.go.id']);
    us.appendRow(['komite', hashPw('komite123'), 'dr. Siti Rahma, Sp.A', 'Komite', 'Komite Mutu & KP', 'Approved', 'komitemutu@rsudmerauke.go.id']);
    us.appendRow(['karu', hashPw('karu123'), 'Ns. Yohanis, S.Kep', 'Karu', 'IGD Ponek', 'Approved', 'karu_igd@rsudmerauke.go.id']);
    us.appendRow(['staf', hashPw('staf123'), 'Bdn. Maria Kaize, A.Md.Keb', 'Staf', 'IGD Ponek', 'Approved', 'maria@rsudmerauke.go.id']);
    Logger.log('Akun default berhasil dibuat: admin/admin123, komite/komite123, karu/karu123, staf/staf123');
  }
  const fs = ss.getSheetByName('Form_Questions');
  if (fs.getLastRow() === 1) {
    FORM_INTERNAL.forEach(q => fs.appendRow(q));
  }
  const st = ss.getSheetByName('App_Settings');
  if (st.getLastRow() === 1) [['app_title', 'RSUD Merauke'], ['app_tagline', 'Melayani dengan hati'], ['komite_email', 'komitemutu@rsudmerauke.go.id'], ['tg_komite', ''], ['tg_admin', ''], ['tg_direksi', ''], ['direksi_email', '']].forEach(r => st.appendRow(r));
  jenisInfoSheet();   // buat sheet Jenis_Insiden + isi 5 jenis bawaan bila belum ada
  Logger.log('setupDatabase selesai! Database URL: ' + ss.getUrl());
}

// JALANKAN SEKALI (opsional, direkomendasikan): memperbarui istilah KPC -> KPCS pada data yang sudah tersimpan:
// kolom jenis_insiden di Laporan_IKP, jawaban "Kondisi Potensial Cedera/KPC" pada data_json, dan opsi di Form_Questions.
// Tanpa menjalankan ini pun dashboard tetap menampilkan KPCS (data lama dinormalisasi saat dibaca).
function migrasiKPCkeKPCS() {
  const OLD = 'Kondisi Potensial Cedera/KPC', NEW = 'Kondisi Potensial Cedera Signifikan/KPCS';
  const ls = sh('Laporan_IKP'), ld = ls.getDataRange().getValues();
  let nJenis = 0, nJson = 0;
  for (let i = 1; i < ld.length; i++) {
    if (String(ld[i][5]).trim() === 'KPC') { ls.getRange(i + 1, 6).setValue('KPCS'); nJenis++; }
    const js = String(ld[i][6]);
    if (js.indexOf(OLD) >= 0) { ls.getRange(i + 1, 7).setValue(js.split(OLD).join(NEW)); nJson++; }
  }
  const fs = sh('Form_Questions'), fd = fs.getDataRange().getValues();
  let nOpt = 0;
  for (let i = 1; i < fd.length; i++) {
    const o = String(fd[i][3]);
    if (o.indexOf(OLD) >= 0) { fs.getRange(i + 1, 4).setValue(o.split(OLD).join(NEW)); nOpt++; }
  }
  Logger.log('Migrasi KPC -> KPCS selesai. Jenis diubah: ' + nJenis + ' | data_json diubah: ' + nJson + ' | opsi formulir diubah: ' + nOpt);
}

// JALANKAN SEKALI setelah deploy v2: meng-hash semua password teks biasa di sheet Users.
function migratePasswords() {
  const s = sh('Users'), d = s.getDataRange().getValues(), weak = [];
  let n = 0;
  for (let i = 1; i < d.length; i++) {
    const p = String(d[i][1]).trim();
    if (p.indexOf('sha256$') === 0) continue;
    if (p === 'admin123' || p === '123456') weak.push(d[i][0]);
    s.getRange(i + 1, 2).setValue(hashPw(p)); n++;
  }
  Logger.log(n + ' password di-hash. Akun masih memakai password bawaan (WAJIB diganti): ' + (weak.join(', ') || '-'));
}

// JALANKAN SEKALI: mengganti seluruh pertanyaan formulir dengan "Formulir Laporan Insiden Keselamatan Pasien (Internal) 2026".
// Pertanyaan lama dicadangkan ke sheet "Backup_Form_<tanggal>". Laporan yang sudah ada tidak terpengaruh.
function terapkanFormulirInternal() {
  const ss = getDbSpreadsheet(), s = ss.getSheetByName('Form_Questions');
  s.copyTo(ss).setName('Backup_Form_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss'));
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
  s.getRange(2, 1, FORM_INTERNAL.length, 6).setValues(FORM_INTERNAL);
  Logger.log(FORM_INTERNAL.length + ' pertanyaan diterapkan.');
}

// LANGKAH 2: cari chat ID (jalankan setelah penerima kirim /start ke bot)
function langkah2_cariChatId() {
  const t = PropertiesService.getScriptProperties().getProperty('TG_TOKEN');
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + t + '/getUpdates', { muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText());
  (j.result || []).forEach(u => {
    const m = u.message || u.my_chat_member || u.channel_post;
    const c = m && m.chat;
    if (c) Logger.log(c.id + ' | ' + c.type + ' | ' + (c.title || c.first_name || '') + ' ' + (c.username || ''));
  });
}

// LANGKAH 3: simpan chat ID penerima (token tidak perlu diisi lagi)
function langkah3_setup() {
  Logger.log(setupTelegram(null, '-5359256197', '6650520932'));
}

// JALANKAN SEKALI (opsional): laporan Kuning/Merah yang RCA-nya sudah tersimpan SEBELUM fitur Review Direksi ada
// (status 'Selesai', belum pernah direview) dikirim ke Direksi supaya ikut dievaluasi. Aman dijalankan ulang.
function kirimRCALamaKeDireksi() {
  const rcaIds = {};
  rcaSheet().getDataRange().getValues().slice(1).forEach(r => { rcaIds[r[1]] = true; });
  const reviewed = {};
  const rs = reviewSheet(false);
  if (rs && rs.getLastRow() > 1) rs.getDataRange().getValues().slice(1).forEach(r => { reviewed[r[1]] = true; });
  const s = sh('Laporan_IKP'), d = s.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < d.length; i++) {
    if (d[i][11] === ST_SELESAI && rcaIds[d[i][0]] && !reviewed[d[i][0]]) {
      s.getRange(i + 1, 12, 1, 3).setValues([[ST_DIREKSI, new Date(), 'sistem']]);
      sh('Audit_Trail').appendRow(['AUD-' + Utilities.getUuid().slice(0, 8), d[i][0], new Date(), 'sistem', 'Sistem', 'Admin', 'RCA lama dikirim ke Direksi (migrasi)', ST_SELESAI, ST_DIREKSI, '']);
      n++;
    }
  }
  Logger.log(n + ' laporan dengan RCA lama dikirim ke antrean review Direksi.');
}
