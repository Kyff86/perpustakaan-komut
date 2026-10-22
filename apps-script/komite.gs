// ==========================================================
// komite.gs — Modul Komite Mutu dan Keselamatan Pasien (SOP, SOP per ruangan, Semua SOP, Bookmark SOP, Jadwal, Perpustakaan)
// Memakai ulang auth()/locked()/ok()/fail()/fmt()/T()/sh()/audit() dari code.gs (e-IKP).
// Semua fungsi publik diawali "api" dan menerima token sesi sebagai argumen pertama.
// ==========================================================
const K_SHEETS = {
  Kegiatan: ['id', 'tanggal', 'nama_kegiatan', 'keterangan', 'bukti_url', 'tahun'],
  Master_Komite: ['id', 'nama_komite'],
  Perpustakaan_Folder: ['id', 'nama_folder', 'icon_class'],
  Perpustakaan_File: ['id', 'folder_id', 'nama_file', 'file_url', 'tanggal_upload'],
  SOP_Master: ['id', 'judul_sop', 'unit_pengusul', 'nomor_dokumen', 'status_terkini', 'tanggal_update', 'file_url'],
  SOP_Logs: ['id', 'sop_id', 'tahapan', 'catatan', 'diperbarui_oleh', 'waktu'],
  Laporan_Mutu: ['id', 'tanggal', 'periode', 'unit', 'indikator', 'numerator', 'denominator', 'persen', 'target', 'analisis', 'dibuat_oleh'],
  SOP_Bookmark: ['id', 'username', 'sop_id', 'waktu']
};
const K_ROLES_WRITE = ['Komite', 'KetuaKomite', 'Admin'];
const K_MAX_BYTES = 5 * 1024 * 1024;
const SOP_VALID_YEARS = 3;
const SOP_STAGE_ROLES = { Substansi: ['Komite', 'KetuaKomite', 'Admin'], Legalitas: ['TU', 'Admin'], Final: ['Direktur', 'Admin'] };
const SOP_NEXT = { Substansi: 'Legalitas', Legalitas: 'Final', Final: 'Approved' };
const SOP_DRAFT_ROLES = ['Karu', 'Admin'];

// Sheet dibuat otomatis saat pertama dipakai (database e-IKP lama tidak diubah).
function ks(name) {
  const ss = getDbSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    const h = K_SHEETS[name];
    s.appendRow(h);
    s.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#166534').setFontColor('#ffffff');
    if (name === 'Perpustakaan_Folder') [['Akreditasi RS', 'fa-shield-halved'], ['SOP & Kebijakan', 'fa-file-contract'], ['Insiden Keselamatan', 'fa-triangle-exclamation'], ['Indikator Mutu', 'fa-chart-line']]
      .forEach((r, i) => s.appendRow(['fd_' + (i + 1), r[0], r[1]]));
    if (name === 'Master_Komite') ['Komite Mutu & Keselamatan Pasien', 'Komite Medik', 'Komite Keperawatan'].forEach((n, i) => s.appendRow(['km_' + (i + 1), n]));
  }
  return s;
}
function setupKomiteDatabase() { Object.keys(K_SHEETS).forEach(ks); return 'Sheet modul Komite Mutu siap: ' + Object.keys(K_SHEETS).join(', '); }
function kRows(name) { const s = ks(name); return s.getLastRow() < 2 ? [] : s.getRange(2, 1, s.getLastRow() - 1, K_SHEETS[name].length).getValues(); }
function kFind(name, id) { const d = ks(name).getDataRange().getValues(); for (let i = 1; i < d.length; i++) if (String(d[i][0]) === String(id)) return { row: i + 1, r: d[i] }; return null; }
function kDay(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || ''); }
function kDriveSave(fileName, mime, base64) {
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > K_MAX_BYTES) throw new Error('Ukuran file maksimal 5 MB.');
  const it = DriveApp.getFoldersByName('Komite Mutu Drive');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('Komite Mutu Drive');
  const f = folder.createFile(Utilities.newBlob(bytes, mime || 'application/octet-stream', String(fileName || 'berkas').slice(0, 150)));
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return f.getUrl();
}
function kTrash(url) {
  const m = /[-\w]{25,}/.exec(String(url || ''));
  if (m) { try { DriveApp.getFileById(m[0]).setTrashed(true); } catch (e) {} }
}

// ---------- Beranda / Dashboard Mutu ----------
function apiKomiteHome(token) {
  try {
    const u = auth(token), now = new Date();
    const lim = new Date(now.getFullYear() - SOP_VALID_YEARS, now.getMonth(), now.getDate());
    const st = { aktif: 0, kadaluarsa: 0, proses: 0, draft: 0, menunggu_saya: 0 };
    kRows('SOP_Master').forEach(r => {
      const s = String(r[4]);
      if (s === 'Approved') { const t = new Date(r[5]); (!isNaN(t) && t < lim) ? st.kadaluarsa++ : st.aktif++; }
      else if (SOP_STAGE_ROLES[s]) { st.proses++; if (SOP_STAGE_ROLES[s].indexOf(u.role) >= 0) st.menunggu_saya++; }
      else st.draft++;
    });
    const t0 = Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), t1 = Utilities.formatDate(new Date(now.getTime() + 7 * 864e5), TZ, 'yyyy-MM-dd');
    const minggu = kRows('Kegiatan').map(r => ({ id: r[0], tanggal: kDay(r[1]), nama: r[2], keterangan: r[3] }))
      .filter(k => k.tanggal >= t0 && k.tanggal <= t1).sort((a, b) => a.tanggal < b.tanggal ? -1 : 1);
    return ok({ stats: st, minggu: minggu });
  } catch (e) { return fail(e); }
}

// ---------- Jadwal Kegiatan ----------
function apiKegList(token) {
  try {
    auth(token);
    return ok({ items: kRows('Kegiatan').map(r => ({ id: r[0], tanggal: kDay(r[1]), nama: r[2], keterangan: r[3], bukti_url: r[4], tahun: r[5] })) });
  } catch (e) { return fail(e); }
}
function apiKegSave(token, d) {
  return locked(() => {
    const u = auth(token, K_ROLES_WRITE); d = d || {};
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d.tanggal || ''))) throw new Error('Tanggal tidak valid.');
    if (!T(d.nama, 200)) throw new Error('Nama kegiatan wajib diisi.');
    const bukti = d.base64 ? kDriveSave(d.fileName, d.fileType, d.base64) : '';
    const s = ks('Kegiatan'), f = d.id ? kFind('Kegiatan', d.id) : null;
    if (d.id && !f) throw new Error('Kegiatan tidak ditemukan.');
    if (f) s.getRange(f.row, 1, 1, 6).setValues([[f.r[0], d.tanggal, T(d.nama, 200), T(d.keterangan, 1000), bukti || f.r[4], Number(d.tanggal.slice(0, 4))]]);
    else s.appendRow(['kg_' + Date.now(), d.tanggal, T(d.nama, 200), T(d.keterangan, 1000), bukti, Number(d.tanggal.slice(0, 4))]);
    s.getRange(2, 2, Math.max(1, s.getLastRow() - 1), 1).setNumberFormat('@');
    audit('KEGIATAN', u, f ? 'Ubah kegiatan' : 'Tambah kegiatan', '', '', T(d.nama, 120));
    return ok({ message: 'Kegiatan berhasil disimpan.' });
  });
}
function apiKegDelete(token, id) {
  return locked(() => {
    const u = auth(token, K_ROLES_WRITE), f = kFind('Kegiatan', id);
    if (!f) throw new Error('Kegiatan tidak ditemukan.');
    kTrash(f.r[4]); ks('Kegiatan').deleteRow(f.row);
    audit('KEGIATAN', u, 'Hapus kegiatan', '', '', String(f.r[2]).slice(0, 120));
    return ok({ message: 'Kegiatan dihapus.' });
  });
}

// ---------- Perpustakaan ----------
function apiLibList(token) {
  try {
    auth(token);
    return ok({
      folders: kRows('Perpustakaan_Folder').map(r => ({ id: r[0], nama: r[1], icon: r[2] || 'fa-folder' })),
      files: kRows('Perpustakaan_File').map(r => ({ id: r[0], folder_id: r[1], nama: r[2], url: r[3], tanggal: kDay(r[4]) }))
    });
  } catch (e) { return fail(e); }
}
function apiLibSaveFolder(token, d) {
  return locked(() => {
    const u = auth(token, K_ROLES_WRITE); d = d || {};
    const nama = T(d.nama, 80), icon = /^fa-[a-z0-9-]{2,40}$/.test(String(d.icon || '')) ? d.icon : 'fa-folder';
    if (!nama) throw new Error('Nama folder wajib diisi.');
    const f = d.id ? kFind('Perpustakaan_Folder', d.id) : null;
    if (f) ks('Perpustakaan_Folder').getRange(f.row, 2, 1, 2).setValues([[nama, icon]]);
    else ks('Perpustakaan_Folder').appendRow(['fd_' + Date.now(), nama, icon]);
    audit('PERPUSTAKAAN', u, f ? 'Ubah folder' : 'Tambah folder', '', '', nama);
    return ok({ message: 'Folder disimpan.' });
  });
}
function apiLibDeleteFolder(token, id) {
  return locked(() => {
    const u = auth(token, K_ROLES_WRITE), f = kFind('Perpustakaan_Folder', id);
    if (!f) throw new Error('Folder tidak ditemukan.');
    if (kRows('Perpustakaan_File').some(r => String(r[1]) === String(id))) throw new Error('Folder masih berisi berkas. Hapus berkasnya terlebih dahulu.');
    ks('Perpustakaan_Folder').deleteRow(f.row);
    audit('PERPUSTAKAAN', u, 'Hapus folder', '', '', String(f.r[1]));
    return ok({ message: 'Folder dihapus.' });
  });
}
function apiLibUpload(token, d) {
  return locked(() => {
    const u = auth(token, K_ROLES_WRITE); d = d || {};
    if (!kFind('Perpustakaan_Folder', d.folder_id)) throw new Error('Folder tidak ditemukan.');
    if (!d.base64) throw new Error('Pilih berkas terlebih dahulu.');
    const url = kDriveSave(d.fileName, d.fileType, d.base64), nama = T(d.nama, 150) || T(d.fileName, 150);
    ks('Perpustakaan_File').appendRow(['fl_' + Date.now(), d.folder_id, nama, url, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')]);
    audit('PERPUSTAKAAN', u, 'Unggah berkas', '', '', nama);
    return ok({ message: 'Berkas berhasil diunggah.' });
  });
}
function apiLibDeleteFile(token, id) {
  return locked(() => {
    const u = auth(token, K_ROLES_WRITE), f = kFind('Perpustakaan_File', id);
    if (!f) throw new Error('Berkas tidak ditemukan.');
    kTrash(f.r[3]); ks('Perpustakaan_File').deleteRow(f.row);
    audit('PERPUSTAKAAN', u, 'Hapus berkas', '', '', String(f.r[2]));
    return ok({ message: 'Berkas dihapus.' });
  });
}

// ---------- Alur SOP: Draft (Unit) -> Substansi (Komite) -> Legalitas (TU) -> Final (Direktur) -> Approved ----------
function sopPub(r, u) {
  const s = String(r[4]), t = new Date(r[5]);
  const exp = s === 'Approved' && !isNaN(t) && t < new Date(new Date().getFullYear() - SOP_VALID_YEARS, new Date().getMonth(), new Date().getDate());
  const own = (u.role === 'Admin') || (u.role === 'Karu' && String(r[2]) === u.ruangan);
  return { id: r[0], judul: r[1], unit: r[2], nomor: r[3], status: s, kadaluarsa: exp, tanggal: fmt(r[5], 'dd/MM/yyyy HH:mm'), file_url: r[6],
    can_act: !!(SOP_STAGE_ROLES[s] && SOP_STAGE_ROLES[s].indexOf(u.role) >= 0), can_submit: !!(own && SOP_DRAFT_ROLES.indexOf(u.role) >= 0 && (s === 'Draft' || s === 'Rejected')) };
}
function apiSopList(token) {
  try {
    const u = auth(token);
    const items = kRows('SOP_Master').filter(r => u.role !== 'Karu' || String(r[2]) === u.ruangan).map(r => sopPub(r, u)).reverse();
    return ok({ items: items });
  } catch (e) { return fail(e); }
}
function apiSopLogs(token, id) {
  try {
    const u = auth(token), f = kFind('SOP_Master', id);
    if (!f || (u.role === 'Karu' && String(f.r[2]) !== u.ruangan)) throw new Error('Akses ditolak.');
    return ok({ logs: kRows('SOP_Logs').filter(r => String(r[1]) === String(id)).map(r => ({ tahap: r[2], catatan: r[3], oleh: r[4], waktu: fmt(r[5]) })).reverse() });
  } catch (e) { return fail(e); }
}
function sopLog(id, tahap, catatan, u) { ks('SOP_Logs').appendRow(['lg_' + Utilities.getUuid().slice(0, 8), id, tahap, T(catatan, 1000), u.nama + ' (' + u.role + ')', new Date()]); }
function apiSopCreate(token, d) {
  return locked(() => {
    const u = auth(token, SOP_DRAFT_ROLES); d = d || {};
    if (!T(d.judul, 200)) throw new Error('Judul SOP wajib diisi.');
    const unit = u.role === 'Admin' ? T(d.unit, 80) : u.ruangan;
    if (!unit) throw new Error('Unit pengusul wajib diisi.');
    const url = d.base64 ? kDriveSave(d.fileName, d.fileType, d.base64) : '';
    const id = 'SOP-' + Date.now();
    ks('SOP_Master').appendRow([id, T(d.judul, 200), unit, '', 'Draft', new Date(), url]);
    sopLog(id, 'Draft', 'Draf SOP dibuat.', u);
    return ok({ message: 'Draf SOP tersimpan. Kirim ke Review Substansi bila sudah siap.', id: id });
  });
}
// action: submit | approve | reject. Dilindungi LockService (locked) agar aksi bersamaan tidak bentrok.
function apiSopAction(token, id, action, catatan, nomor, file) {
  return locked(() => {
    const u = auth(token), f = kFind('SOP_Master', id);
    if (!f) throw new Error('SOP tidak ditemukan.');
    const s = String(f.r[4]), sh0 = ks('SOP_Master'); let next, nom = f.r[3], url = f.r[6];
    if (action === 'submit') {
      if (!(s === 'Draft' || s === 'Rejected')) throw new Error('SOP tidak dalam tahap draf.');
      if (SOP_DRAFT_ROLES.indexOf(u.role) < 0 || (u.role === 'Karu' && String(f.r[2]) !== u.ruangan)) throw new Error('Akses ditolak.');
      if (file && file.base64) url = kDriveSave(file.fileName, file.fileType, file.base64);
      if (!url) throw new Error('Unggah file draf SOP terlebih dahulu.');
      next = 'Substansi';
    } else if (action === 'approve' || action === 'reject') {
      if (!SOP_STAGE_ROLES[s] || SOP_STAGE_ROLES[s].indexOf(u.role) < 0) throw new Error('SOP tidak sedang menunggu tindakan peran Anda.');
      if (action === 'reject') { if (!T(catatan)) throw new Error('Catatan penolakan wajib diisi.'); next = 'Rejected'; }
      else {
        if (s === 'Legalitas') { nom = T(nomor, 80); if (!nom) throw new Error('Nomor dokumen resmi wajib diisi.'); }
        next = SOP_NEXT[s];
      }
    } else throw new Error('Aksi tidak dikenal.');
    sh0.getRange(f.row, 4, 1, 4).setValues([[nom, next, new Date(), url]]);
    sopLog(id, s + ' → ' + next, (action === 'approve' && next === 'Approved') ? (catatan || 'Disahkan Direktur (Published).') : catatan, u);
    audit(id, u, 'SOP ' + action + ': ' + s + ' → ' + next, s, next, String(f.r[1]).slice(0, 100));
    try { sendTelegramNotification('<b>SOP ' + escHtml(next.toUpperCase()) + '</b>\n' + escHtml(String(f.r[1])) + '\nUnit: ' + escHtml(String(f.r[2])), ['komite', 'admin']); } catch (e) {}
    return ok({ message: next === 'Approved' ? 'SOP disahkan dan berstatus Published/Active.' : 'SOP diproses. Status: ' + next });
  });
}

// ---------- Pelaporan Mutu (indikator ruangan) ----------
function mutuPub(r) { return { id: r[0], tanggal: kDay(r[1]), periode: String(r[2]), unit: r[3], indikator: r[4], num: r[5], den: r[6], persen: r[7], target: r[8], analisis: r[9], oleh: r[10] }; }
function apiMutuList(token) {
  try {
    const u = auth(token);
    return ok({ items: kRows('Laporan_Mutu').filter(r => u.role !== 'Karu' || String(r[3]) === u.ruangan).map(mutuPub).reverse() });
  } catch (e) { return fail(e); }
}
function apiMutuSave(token, d) {
  return locked(() => {
    const u = auth(token, ['Karu', 'Komite', 'KetuaKomite', 'Admin']); d = d || {};
    const unit = u.role === 'Karu' ? u.ruangan : T(d.unit, 80), ind = T(d.indikator, 200);
    const num = Number(d.numerator), den = Number(d.denominator), tgt = Number(d.target);
    if (!unit || !ind) throw new Error('Unit dan indikator wajib diisi.');
    if (!/^\d{4}-\d{2}$/.test(String(d.periode || ''))) throw new Error('Periode tidak valid.');
    if (!isFinite(num) || !isFinite(den) || den <= 0 || num < 0 || num > den) throw new Error('Validasi data: numerator harus 0 s.d. denominator, dan denominator > 0.');
    if (!isFinite(tgt) || tgt < 0 || tgt > 100) throw new Error('Target harus 0–100 (%).');
    const pct = Math.round(num / den * 10000) / 100;
    if (pct < tgt && !T(d.analisis)) throw new Error('Capaian di bawah target: analisis / rencana tindak lanjut wajib diisi.');
    ks('Laporan_Mutu').appendRow(['mt_' + Date.now(), Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), d.periode, unit, ind, num, den, pct, tgt, T(d.analisis, 1500), u.nama]);
    audit('MUTU', u, 'Laporan mutu', '', '', unit + ' / ' + ind.slice(0, 80));
    return ok({ message: 'Data mutu tersimpan. Capaian ' + pct + '% (target ' + tgt + '%).' });
  });
}
function apiMutuDelete(token, id) {
  return locked(() => {
    const u = auth(token, ['Komite', 'KetuaKomite', 'Admin']), f = kFind('Laporan_Mutu', id);
    if (!f) throw new Error('Data tidak ditemukan.');
    ks('Laporan_Mutu').deleteRow(f.row);
    return ok({ message: 'Data mutu dihapus.' });
  });
}


// ---------- Pustaka SOP: Semua SOP, SOP per ruangan, Bookmark (hanya SOP berstatus Approved/Published) ----------
function apiSopPublished(token) {
  try {
    const u = auth(token), un = String(u.username).toLowerCase(), bm = {};
    kRows('SOP_Bookmark').forEach(r => { if (String(r[1]).toLowerCase() === un) bm[String(r[2])] = true; });
    const items = kRows('SOP_Master').filter(r => String(r[4]) === 'Approved').map(r => {
      const p = sopPub(r, u); p.bm = !!bm[String(r[0])]; return p;
    }).sort((a, b) => String(a.unit).localeCompare(String(b.unit), 'id') || String(a.judul).localeCompare(String(b.judul), 'id'));
    return ok({ items: items });
  } catch (e) { return fail(e); }
}
// Tombol bookmark: panggilan pertama menyimpan, panggilan kedua menghapus. Bookmark bersifat pribadi (per akun).
function apiSopBookmark(token, id) {
  return locked(() => {
    const u = auth(token), f = kFind('SOP_Master', id);
    if (!f || String(f.r[4]) !== 'Approved') throw new Error('SOP tidak ditemukan atau belum dipublikasikan.');
    const s = ks('SOP_Bookmark'), d = s.getDataRange().getValues(), un = String(u.username).toLowerCase();
    for (let i = d.length - 1; i >= 1; i--) {
      if (String(d[i][1]).toLowerCase() === un && String(d[i][2]) === String(id)) { s.deleteRow(i + 1); return ok({ bookmarked: false }); }
    }
    s.appendRow(['bm_' + Date.now(), u.username, String(id), new Date()]);
    return ok({ bookmarked: true });
  });
}
