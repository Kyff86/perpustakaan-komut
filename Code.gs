/**
 * PERPUSTAKAAN KOMITE MUTU - Backend API (Google Apps Script)
 * Cara pakai: lihat README.md. Jalankan setup() sekali, lalu Deploy > Web app.
 */
const P = PropertiesService.getScriptProperties();
const SHEETS = {
  Users: ['id','username','password_hash','salt','role','nama_lengkap','email','status','must_change_password','last_login'],
  Files: ['id','nomor_dokumen','nama_file','kategori_id','deskripsi','versi','tanggal_review','file_id','mime_type','ukuran_kb','uploaded_by','tanggal_upload','status'],
  Kategori: ['id','nama_kategori','drive_folder_id','status'],
  Settings: ['key','value'],
  AuditLog: ['timestamp','user','aksi','detail']
};
const ROLES = ['admin','staf','viewer'];

/* ===== SETUP (jalankan sekali dari editor) ===== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  P.setProperty('SS_ID', ss.getId());
  Object.keys(SHEETS).forEach(n => {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    s.getRange(1, 1, 1, SHEETS[n].length).setValues([SHEETS[n]]).setFontWeight('bold');
    s.setFrozenRows(1);
  });
  const root = DriveApp.createFolder('Perpustakaan Komite Mutu');
  P.setProperty('ROOT_ID', root.getId());
  P.setProperty('BACKUP_ID', DriveApp.createFolder('Backup Perpustakaan Komut').getId());
  const def = { email_notifikasi: '', notifikasi_aktif: 'tidak', max_ukuran_mb: '10', tipe_diizinkan: 'pdf,docx,xlsx,pptx,jpg,png' };
  Object.keys(def).forEach(k => add('Settings', { key: k, value: def[k] }));
  const pw = Utilities.getUuid().slice(0, 10), salt = Utilities.getUuid();
  add('Users', { id: Utilities.getUuid(), username: 'admin', password_hash: hash(pw, salt), salt, role: 'admin',
    nama_lengkap: 'Administrator', email: '', status: 'aktif', must_change_password: 'ya' });
  Logger.log('SETUP SELESAI. Login: admin / ' + pw + '  (wajib ganti saat login pertama)');
}

/** Jalankan sekali: pasang trigger backup harian & pengingat review */
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupHarian').timeBased().everyDays(1).atHour(2).create();
  ScriptApp.newTrigger('pengingatReview').timeBased().everyDays(1).atHour(7).create();
}

/* ===== HELPER ===== */
const ss = () => SpreadsheetApp.openById(P.getProperty('SS_ID'));
const hash = (p, s) => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s + p));
// Sanitasi: buang tag HTML, batasi panjang, cegah formula injection
const clean = (s, n) => {
  s = String(s == null ? '' : s).replace(/<[^>]*>/g, '').trim().slice(0, n || 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
};
function rows(n) {
  const v = ss().getSheetByName(n).getDataRange().getValues(), h = v.shift();
  return v.map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; });
}
function add(n, o) { ss().getSheetByName(n).appendRow(SHEETS[n].map(k => o[k] === undefined ? '' : o[k])); }
function upd(n, row, o) {
  const s = ss().getSheetByName(n);
  SHEETS[n].forEach((k, j) => { if (o[k] !== undefined) s.getRange(row, j + 1).setValue(o[k]); });
}
function settings() { const o = {}; rows('Settings').forEach(r => o[r.key] = String(r.value)); return o; }
function log(user, aksi, detail) { add('AuditLog', { timestamp: new Date(), user, aksi, detail: clean(detail, 300) }); }
function withLock(fn) {
  const l = LockService.getScriptLock();
  try { l.waitLock(20000); } catch (e) { throw new Error('Sistem sedang sibuk, coba lagi sebentar.'); }
  try { return fn(); } finally { l.releaseLock(); }
}
const pub = u => ({ id: u.id, username: u.username, role: u.role, nama_lengkap: u.nama_lengkap, email: u.email,
  status: u.status, must_change_password: u.must_change_password === 'ya' });
function auth(token) {
  const id = CacheService.getScriptCache().get('s_' + token);
  const u = id && rows('Users').find(x => x.id === id && x.status === 'aktif');
  if (!u) throw new Error('Sesi berakhir, silakan login kembali.');
  return u;
}
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
function notify(subject, body) {
  try { const s = settings(); if (s.notifikasi_aktif === 'ya' && s.email_notifikasi) MailApp.sendEmail(s.email_notifikasi, subject, body); }
  catch (e) { console.error(e); } // kegagalan email tidak boleh membatalkan proses
}

/* ===== ROUTER ===== */
function doGet() { return out({ ok: true, data: 'API Perpustakaan Komite Mutu aktif' }); }
function doPost(e) {
  try {
    const q = JSON.parse(e.postData.contents), A = ACT[q.action];
    if (!A) throw new Error('Aksi tidak dikenal');
    const u = q.action === 'login' ? null : auth(q.token);
    if (A.roles && !A.roles.includes(u.role)) throw new Error('Akses ditolak');
    const run = () => A.fn(q, u);
    return out({ ok: true, data: A.lock ? withLock(run) : run() });
  } catch (err) { return out({ ok: false, error: err.message }); }
}

/* ===== AKSI ===== */
const ADM = ['admin'], TULIS = ['admin', 'staf'];
const ACT = {
  login: { fn: q => {
    const c = CacheService.getScriptCache(), name = clean(q.username, 50), k = 'f_' + name;
    if (+c.get(k) >= 5) throw new Error('Terlalu banyak percobaan. Coba lagi 15 menit.');
    const u = rows('Users').find(x => x.username === name && x.status === 'aktif');
    if (!u || u.password_hash !== hash(String(q.password), u.salt)) {
      c.put(k, String((+c.get(k) || 0) + 1), 900);
      throw new Error('Username atau password salah');
    }
    c.remove(k);
    const t = Utilities.getUuid() + Utilities.getUuid();
    c.put('s_' + t, u.id, 21600); // sesi 6 jam
    upd('Users', u._row, { last_login: new Date() });
    log(u.username, 'login', '');
    return { token: t, user: pub(u) };
  } },
  logout: { fn: q => { CacheService.getScriptCache().remove('s_' + q.token); return true; } },
  changePassword: { fn: (q, u) => {
    if (String(q.password).length < 8) throw new Error('Password minimal 8 karakter');
    const salt = Utilities.getUuid();
    upd('Users', u._row, { password_hash: hash(String(q.password), salt), salt, must_change_password: 'tidak' });
    log(u.username, 'ganti_password', ''); return true;
  } },
  data: { fn: () => ({
    files: rows('Files').filter(f => f.status === 'aktif').map(f => { delete f._row; delete f.file_id; return f; }),
    kategori: rows('Kategori').filter(k => k.status === 'aktif').map(k => ({ id: k.id, nama_kategori: k.nama_kategori }))
  }) },
  upload: { roles: TULIS, lock: true, fn: (q, u) => {
    const s = settings(), max = +s.max_ukuran_mb || 10, bytes = Utilities.base64Decode(q.data);
    if (bytes.length > max * 1048576) throw new Error('Ukuran berkas melebihi ' + max + ' MB');
    const ext = String(q.filename).split('.').pop().toLowerCase();
    if (!s.tipe_diizinkan.split(',').includes(ext)) throw new Error('Tipe berkas .' + ext + ' tidak diizinkan');
    const kat = rows('Kategori').find(k => k.id === q.kategori_id && k.status === 'aktif');
    if (!kat) throw new Error('Kategori tidak valid');
    const nama = clean(q.nama_file, 150), versi = clean(q.versi || '1.0', 10);
    if (!nama) throw new Error('Nama dokumen wajib diisi');
    if (rows('Files').some(f => f.status === 'aktif' && f.nama_file === nama && f.kategori_id === kat.id && String(f.versi) === versi))
      throw new Error('Dokumen dengan nama & versi yang sama sudah ada');
    const nomor = clean(q.nomor_dokumen, 50);
    const file = DriveApp.getFolderById(kat.drive_folder_id)
      .createFile(Utilities.newBlob(bytes, q.mime || 'application/octet-stream', (nomor ? nomor + '_' : '') + nama + '_v' + versi + '.' + ext));
    add('Files', { id: Utilities.getUuid(), nomor_dokumen: nomor, nama_file: nama, kategori_id: kat.id, deskripsi: clean(q.deskripsi, 500),
      versi, tanggal_review: q.tanggal_review || '', file_id: file.getId(), mime_type: q.mime, ukuran_kb: Math.round(bytes.length / 1024),
      uploaded_by: u.nama_lengkap, tanggal_upload: new Date(), status: 'aktif' });
    log(u.username, 'upload', nama);
    notify('[Perpustakaan Komut] Berkas baru: ' + nama, 'Berkas "' + nama + '" (' + kat.nama_kategori + ') diunggah oleh ' + u.nama_lengkap + '.');
    return true;
  } },
  updateFile: { roles: TULIS, lock: true, fn: (q, u) => {
    const f = rows('Files').find(x => x.id === q.id && x.status === 'aktif');
    if (!f) throw new Error('Berkas tidak ditemukan');
    if (u.role !== 'admin' && f.uploaded_by !== u.nama_lengkap) throw new Error('Anda hanya dapat mengedit berkas milik sendiri');
    upd('Files', f._row, { nama_file: clean(q.nama_file, 150) || f.nama_file, deskripsi: clean(q.deskripsi, 500) });
    log(u.username, 'edit_berkas', f.nama_file); return true;
  } },
  deleteFile: { roles: ADM, lock: true, fn: (q, u) => { // soft delete + pindah ke folder Backup
    const f = rows('Files').find(x => x.id === q.id);
    if (!f) throw new Error('Berkas tidak ditemukan');
    upd('Files', f._row, { status: 'dihapus' });
    try { DriveApp.getFileById(f.file_id).moveTo(DriveApp.getFolderById(P.getProperty('BACKUP_ID'))); } catch (e) {}
    log(u.username, 'hapus_berkas', f.nama_file); return true;
  } },
  download: { fn: (q, u) => {
    const f = rows('Files').find(x => x.id === q.id && x.status === 'aktif');
    if (!f) throw new Error('Berkas tidak ditemukan');
    const b = DriveApp.getFileById(f.file_id).getBlob();
    log(u.username, 'unduh', f.nama_file);
    return { name: b.getName(), mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
  } },
  saveKategori: { roles: ADM, lock: true, fn: (q, u) => {
    const nama = clean(q.nama_kategori, 80);
    if (!nama) throw new Error('Nama kategori wajib diisi');
    const all = rows('Kategori').filter(k => k.status === 'aktif');
    if (all.some(k => k.nama_kategori.toLowerCase() === nama.toLowerCase() && k.id !== q.id)) throw new Error('Nama kategori sudah ada');
    if (q.id) {
      const k = all.find(x => x.id === q.id); if (!k) throw new Error('Kategori tidak ditemukan');
      upd('Kategori', k._row, { nama_kategori: nama });
      try { DriveApp.getFolderById(k.drive_folder_id).setName(nama); } catch (e) {}
    } else {
      const fid = DriveApp.getFolderById(P.getProperty('ROOT_ID')).createFolder(nama).getId();
      add('Kategori', { id: Utilities.getUuid(), nama_kategori: nama, drive_folder_id: fid, status: 'aktif' });
    }
    log(u.username, 'simpan_kategori', nama); return true;
  } },
  deleteKategori: { roles: ADM, lock: true, fn: (q, u) => {
    if (rows('Files').some(f => f.status === 'aktif' && f.kategori_id === q.id))
      throw new Error('Kategori masih dipakai berkas. Pindahkan atau hapus berkasnya dahulu.');
    const k = rows('Kategori').find(x => x.id === q.id); if (!k) throw new Error('Kategori tidak ditemukan');
    upd('Kategori', k._row, { status: 'nonaktif' }); log(u.username, 'hapus_kategori', k.nama_kategori); return true;
  } },
  listUsers: { roles: ADM, fn: () => rows('Users').map(pub) },
  saveUser: { roles: ADM, lock: true, fn: (q, u) => {
    const role = ROLES.includes(q.role) ? q.role : 'staf', nama = clean(q.nama_lengkap, 80);
    if (q.id) { // ubah role / status / reset password
      const t = rows('Users').find(x => x.id === q.id); if (!t) throw new Error('User tidak ditemukan');
      const o = { role, nama_lengkap: nama || t.nama_lengkap, email: clean(q.email, 100), status: q.status === 'nonaktif' ? 'nonaktif' : 'aktif' };
      if (t.id === u.id && (o.status === 'nonaktif' || role !== 'admin')) throw new Error('Anda tidak dapat menurunkan/menonaktifkan akun sendiri');
      if (q.password) { if (String(q.password).length < 8) throw new Error('Password minimal 8 karakter');
        o.salt = Utilities.getUuid(); o.password_hash = hash(String(q.password), o.salt); o.must_change_password = 'ya'; }
      upd('Users', t._row, o);
    } else {
      const un = clean(q.username, 50).toLowerCase();
      if (!/^[a-z0-9._-]{3,}$/.test(un)) throw new Error('Username minimal 3 karakter (huruf/angka/._-)');
      if (rows('Users').some(x => x.username === un)) throw new Error('Username sudah dipakai');
      if (String(q.password).length < 8) throw new Error('Password minimal 8 karakter');
      const salt = Utilities.getUuid();
      add('Users', { id: Utilities.getUuid(), username: un, password_hash: hash(String(q.password), salt), salt, role, nama_lengkap: nama || un,
        email: clean(q.email, 100), status: 'aktif', must_change_password: 'ya' });
    }
    log(u.username, 'simpan_user', q.username || q.id); return true;
  } },
  getSettings: { roles: ADM, fn: () => settings() },
  saveSettings: { roles: ADM, lock: true, fn: (q, u) => {
    const v = { email_notifikasi: clean(q.email_notifikasi, 300), notifikasi_aktif: q.notifikasi_aktif === 'ya' ? 'ya' : 'tidak',
      max_ukuran_mb: String(Math.min(Math.max(+q.max_ukuran_mb || 10, 1), 20)) };
    rows('Settings').forEach(r => { if (v[r.key] !== undefined) upd('Settings', r._row, { value: v[r.key] }); });
    log(u.username, 'simpan_pengaturan', ''); return true;
  } },
  audit: { roles: ADM, fn: () => rows('AuditLog').slice(-100).reverse().map(r => ({ timestamp: r.timestamp, user: r.user, aksi: r.aksi, detail: r.detail })) }
};

/* ===== TRIGGER OTOMATIS ===== */
function backupHarian() {
  const f = DriveApp.getFolderById(P.getProperty('BACKUP_ID'));
  DriveApp.getFileById(P.getProperty('SS_ID')).makeCopy('Backup_' + Utilities.formatDate(new Date(), 'Asia/Jayapura', 'yyyy-MM-dd'), f);
  const list = []; const it = f.searchFiles("title contains 'Backup_'");
  while (it.hasNext()) list.push(it.next());
  list.sort((a, b) => b.getDateCreated() - a.getDateCreated()).slice(10).forEach(x => x.setTrashed(true)); // simpan 10 terakhir
}
function pengingatReview() {
  const batas = new Date(Date.now() + 30 * 864e5), now = new Date();
  const due = rows('Files').filter(f => f.status === 'aktif' && f.tanggal_review instanceof Date && f.tanggal_review <= batas);
  if (!due.length) return;
  notify('[Perpustakaan Komut] ' + due.length + ' dokumen perlu direview',
    due.map(f => '- ' + f.nama_file + ' (review: ' + Utilities.formatDate(f.tanggal_review, 'Asia/Jayapura', 'dd/MM/yyyy') + (f.tanggal_review < now ? ', TERLAMBAT' : '') + ')').join('\n'));
}
