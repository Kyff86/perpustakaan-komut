/**
 * Pelaporan Mutu (E-Mutu) RSUD Merauke - dijadikan modul di dalam Sistem Komite RS (mutu.gs)
 * Versi 2.6 - preview langsung pada editor template + opsi target 'tanpa tanda' (none); Versi 2.5 - edit Profil & Prioritas (Master_Profil), target 'angka saja' pada template; Versi 2.4 - tambah/ubah/hapus ruangan (saveUnit, deleteUnit) + simpan/restore database; isiIndikatorBawaan dihapus
 */

var SESSION_TTL = 21600; // 6 jam (detik)

// Kolom sheet Master_Indikator (formula_tipe: 'ge' = capaian harus >= target, 'le' = capaian harus <= target)
var IND_HEADERS = ['id_indikator', 'id_unit', 'nama_indikator', 'kategori', 'target_persen', 'formula_tipe', 'standar_teks', 'numerator_desc', 'denominator_desc', 'id_form'];
var IND_KATEGORI = ['INM', 'IMP RS', 'IMP Unit', 'SPM', 'PPI', 'Lainnya']; // urutan baku tampilan

/* ============================================================
 * ENTRY POINT
 * ============================================================ */


/* ============================================================
 * HELPER UMUM
 * ============================================================ */


/* ============================================================
 * JEMBATAN LOGIN: E-Mutu memakai sesi & akun e-IKP (auth() di code.gs).
 * Peran e-IKP dipetakan ke peran E-Mutu:
 *   Admin / Komite / KetuaKomite -> 'Komite Mutu'  (kelola master, template, validasi)
 *   Karu                          -> 'PJ Mutu Unit' (lapor untuk ruangannya sendiri)
 *   peran lain                    -> 'Staf RS'      (hanya melihat dashboard & profil)
 * ============================================================ */
function mutuRole_(role) {
  if (role === 'Admin' || role === 'Komite' || role === 'KetuaKomite') return 'Komite Mutu';
  if (role === 'Karu') return 'PJ Mutu Unit';
  return 'Staf RS';
}

function requireAuth_(token) {
  var u;
  try { u = auth(token); } catch (e) { throw new Error('Sesi berakhir. Silakan login ulang.'); }
  var unit = String(u.ruangan || '').trim();
  try { // samakan penulisan ruangan akun dengan Master_Unit (abaikan beda huruf besar/kecil)
    var key = unit.toLowerCase();
    getSheetObjects_(getSheet_('Master_Unit')).forEach(function(r) {
      if (String(r.nama_unit).trim().toLowerCase() === key) unit = String(r.nama_unit).trim();
    });
  } catch (e) {}
  return { id_user: u.username, username: u.username, email: u.email || '', nama_lengkap: u.nama, role: mutuRole_(u.role), unit_kerja: unit };
}

function requireRole_(user, roles) {
  if (roles.indexOf(user.role) === -1) throw new Error('Anda tidak memiliki akses untuk tindakan ini.');
}

function getSheet_(name) {
  var s = getDbSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('Sheet "' + name + '" belum ada. Jalankan setupDatabase terlebih dahulu.');
  return s;
}

function getSheetObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  var headers = values[0] || [];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var o = { _row: i + 1 };
    for (var j = 0; j < headers.length; j++) o[headers[j]] = values[i][j];
    out.push(o);
  }
  return out;
}

// Date tidak bisa dikirim ke browser lewat google.script.run, ubah ke teks
function cellStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  return (v === null || v === undefined) ? '' : String(v);
}

function bulanStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  return String(v || '');
}

// Cegah formula injection (=, +, -, @ di awal teks) & batasi panjang
function safeText_(s, max) {
  s = String(s === null || s === undefined ? '' : s).slice(0, max || 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function newId_(prefix) {
  return prefix + Utilities.formatDate(new Date(), 'Asia/Jayapura', 'yyMMddHHmmss') + Math.floor(100 + Math.random() * 900);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ============================================================
 * INISIALISASI DATABASE (aman dijalankan ulang: hanya menambah yang belum ada)
 * ============================================================ */
function setupMutuDatabase() {
  var ss = getDbSpreadsheet();

  var sheetsDefinition = [
    { name: 'Master_Unit', headers: ['id_unit', 'nama_unit', 'kategori_unit'] },
    { name: 'Master_Indikator', headers: IND_HEADERS },
    { name: 'Laporan_Harian', headers: ['id_harian', 'tanggal', 'id_unit', 'id_indikator', 'numerator', 'denominator', 'created_by', 'timestamp'] },
    { name: 'Laporan_Bulanan', headers: ['id_bulanan', 'bulan_tahun', 'id_unit', 'id_indikator', 'total_numerator', 'total_denominator', 'capaian_persen', 'analisis_masalah', 'rencana_tindak_lanjut', 'status_validasi', 'catatan_validator', 'created_by', 'validated_by', 'timestamp'] },
    // --- BARU: modul Form Pelaporan ---
    { name: 'Master_Form', headers: ['id_form', 'kode', 'judul', 'kategori', 'definisi_json', 'status_aktif', 'updated_by', 'updated_at'] },
    { name: 'Laporan_Form', headers: ['id_laporan', 'id_form', 'nama_form', 'bulan_tahun', 'unit', 'data_json', 'meta_json', 'num', 'denum', 'capaian_persen', 'status', 'catatan_validator', 'created_by', 'validated_by', 'created_at', 'updated_at'] }
  ];

  sheetsDefinition.forEach(function(item) {
    var sheet = ss.getSheetByName(item.name);
    if (!sheet) sheet = ss.insertSheet(item.name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(item.headers);
      var headerRange = sheet.getRange(1, 1, 1, item.headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#0284c7');
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  });

  // Kolom bulan_tahun harus teks agar "2026-10" tidak diubah Sheets menjadi tanggal
  var lap = ss.getSheetByName('Laporan_Form');
  lap.getRange(2, 4, Math.max(lap.getMaxRows() - 1, 1), 1).setNumberFormat('@');


  var unitSheet = ss.getSheetByName('Master_Unit');
  if (unitSheet.getLastRow() === 1) {
    unitSheet.appendRow(['UNT001', 'IGD', 'Pelayanan Medis']);
    unitSheet.appendRow(['UNT002', 'Rawat Inap', 'Pelayanan Medis']);
    unitSheet.appendRow(['UNT003', 'Laboratorium', 'Penunjang Medis']);
  }

  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  getIndikatorSheet_();
  return 'Inisialisasi sheet Pelaporan Mutu (E-Mutu) berhasil!';
}

/**
 * Jalankan SEKALI dari editor Apps Script: menambahkan daftar 28 ruangan
 * sesuai Keputusan Direktur RSUD Kelas C Merauke No: 400.7.1/0311.a/RSMRK/III/2026
 * tentang Penetapan Indikator Mutu Prioritas Rumah Sakit Tahun 2026.
 */
function tambahUnitDariDina() {
  var sheet = getSheet_('Master_Unit');
  var existing = {};
  getSheetObjects_(sheet).forEach(function(o) { existing[String(o.nama_unit).toLowerCase()] = true; });
  var daftar = [
    ['IGD', 'Pelayanan Medis'],
    ['Ranap Rusa II', 'Rawat Inap'],
    ['Ranap Kus-Kus', 'Rawat Inap'],
    ['Ranap Boha', 'Rawat Inap'],
    ['Ranap Maleo', 'Rawat Inap'],
    ['Ranap Cenderawasih', 'Rawat Inap'],
    ['Ruang Urip', 'Rawat Inap'],
    ['Ranap Kanguru', 'Rawat Inap'],
    ['Ranap Mambruk', 'Rawat Inap'],
    ['Ranap Rusa I', 'Rawat Inap'],
    ['PICU', 'Pelayanan Medis'],
    ['Ruang Kasuari', 'Rawat Inap'],
    ['Ruang Hemodialisa', 'Penunjang Medis'],
    ['IBS', 'Pelayanan Medis'],
    ['ICU', 'Pelayanan Medis'],
    ['Rajal A', 'Rawat Jalan'],
    ['Rajal B', 'Rawat Jalan'],
    ['Ponek', 'Pelayanan Medis'],
    ['Laboratorium', 'Penunjang Medis'],
    ['Farmasi', 'Penunjang Medis'],
    ['BDRS', 'Penunjang Medis'],
    ['CSSD', 'Penunjang'],
    ['Radiologi', 'Penunjang Medis'],
    ['Gizi', 'Penunjang Medis'],
    ['Pemulasaran Jenasah', 'Penunjang'],
    ['Admisi', 'Penunjang'],
    ['Rekam Medik', 'Penunjang'],
    ['Surveilans', 'Komite / Tim'],
    ['Komite Mutu', 'Komite / Tim'],
    ['PPI', 'Komite / Tim'],
    ['PIPP', 'Komite / Tim'],
    ['Pelayanan', 'Manajemen']
  ];
  var n = sheet.getLastRow(); // jumlah baris saat ini (termasuk header)
  var tambah = 0;
  daftar.forEach(function(d) {
    if (existing[d[0].toLowerCase()]) return;
    n++; tambah++;
    sheet.appendRow(['UNT' + ('000' + n).slice(-3), d[0], d[1]]);
  });
  return tambah + ' unit baru ditambahkan sesuai SK Penetapan 2026.';
}

/* ============================================================
 * AUTENTIKASI
 * ============================================================ */








/* ============================================================
 * MASTER UNIT
 * ============================================================ */
function getMasterUnits(token) {
  requireAuth_(token);
  return getSheetObjects_(getSheet_('Master_Unit')).map(function(o) {
    return { id_unit: cellStr_(o.id_unit), nama_unit: cellStr_(o.nama_unit), kategori_unit: cellStr_(o.kategori_unit) };
  }).filter(function(u) { return u.nama_unit; }).sort(function(a, b) {
    // urutan abjad (A-Z) agar daftar ruangan otomatis rapi saat ada penambahan/penghapusan
    return a.nama_unit.localeCompare(b.nama_unit, 'id', { numeric: true, sensitivity: 'base' });
  });
}

// Ubah nama ruangan juga memperbarui teks nama di Laporan_Form (kolom unit) dan Users (kolom unit_kerja)
function renameUnitRefs_(lama, baru) {
  [['Laporan_Form', 'unit']].forEach(function(p) {
    var sh = getDbSpreadsheet().getSheetByName(p[0]);
    if (!sh || sh.getLastRow() < 2) return;
    var c = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf(p[1]) + 1;
    if (!c) return;
    var rg = sh.getRange(2, c, sh.getLastRow() - 1, 1), v = rg.getValues(), ch = false;
    v.forEach(function(r) { if (String(r[0]) === lama) { r[0] = baru; ch = true; } });
    if (ch) rg.setValues(v);
  });
}

function countUnitRefs_(sheetName, colName, nama) {
  var sh = getDbSpreadsheet().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return 0;
  var c = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf(colName) + 1;
  if (!c) return 0;
  var n = 0;
  sh.getRange(2, c, sh.getLastRow() - 1, 1).getValues().forEach(function(r) { if (String(r[0]) === nama) n++; });
  return n;
}

// Tambah ruangan (tanpa id_unit) atau ubah nama/kategori ruangan (dengan id_unit) - hanya Komite Mutu
function saveUnit(token, d) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  return withLock_(function() {
    d = d || {};
    var nama = safeText_(String(d.nama_unit || '').trim(), 100);
    if (!nama) throw new Error('Nama ruangan wajib diisi.');
    var kat = safeText_(String(d.kategori_unit || '').trim(), 100);
    var sh = getSheet_('Master_Unit');
    var found = null, dup = false, max = 0;
    getSheetObjects_(sh).forEach(function(o) {
      if (d.id_unit && String(o.id_unit) === String(d.id_unit)) found = o;
    });
    getSheetObjects_(sh).forEach(function(o) {
      if (String(o.nama_unit).trim().toLowerCase() === nama.toLowerCase() && (!found || o._row !== found._row)) dup = true;
      var m = /^UNT(\d+)$/.exec(String(o.id_unit));
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
    if (d.id_unit && !found) throw new Error('Ruangan tidak ditemukan.');
    if (dup) throw new Error('Ruangan "' + nama + '" sudah ada.');
    if (found) {
      var lama = String(found.nama_unit);
      sh.getRange(found._row, 2, 1, 2).setValues([[nama, kat]]);
      if (lama !== nama) renameUnitRefs_(lama, nama);
      return { success: true, id_unit: String(found.id_unit), nama_unit: nama };
    }
    var id = 'UNT' + ('000' + (max + 1)).slice(-3);
    sh.appendRow([id, nama, kat]);
    return { success: true, id_unit: id, nama_unit: nama };
  });
}

// Hapus ruangan beserta indikatornya. Ditolak bila masih ada laporan tersimpan di ruangan itu.
function deleteUnit(token, id) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  return withLock_(function() {
    var sh = getSheet_('Master_Unit');
    var unit = null;
    getSheetObjects_(sh).forEach(function(o) { if (String(o.id_unit) === String(id)) unit = o; });
    if (!unit) throw new Error('Ruangan tidak ditemukan.');
    var nama = String(unit.nama_unit);
    var nLap = countUnitRefs_('Laporan_Form', 'unit', nama);
    if (nLap) throw new Error('Ruangan "' + nama + '" masih memiliki ' + nLap + ' laporan tersimpan. Hapus laporannya terlebih dahulu atau cukup ubah namanya.');
    var ish = getIndikatorSheet_(), iobjs = getSheetObjects_(ish), removed = 0;
    for (var i = iobjs.length - 1; i >= 0; i--) {
      if (String(iobjs[i].id_unit) === String(id)) { ish.deleteRow(iobjs[i]._row); removed++; }
    }
    sh.deleteRow(unit._row);
    return { success: true, indikator_dihapus: removed };
  });
}

/* ============================================================
 * MASTER INDIKATOR MUTU PER RUANGAN
 * ============================================================ */
// Sheet dibuat otomatis bila belum ada; kolom baru ditambahkan bila sheet lama masih 6 kolom.
function getIndikatorSheet_() {
  var ss = getDbSpreadsheet();
  var sh = ss.getSheetByName('Master_Indikator');
  if (!sh) sh = ss.insertSheet('Master_Indikator');
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var cur = sh.getLastRow() === 0 ? [] : sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var changed = false;
  IND_HEADERS.forEach(function(h, i) {
    if (String(cur[i] || '') !== h) { sh.getRange(1, i + 1).setValue(h); changed = true; }
  });
  if (changed || sh.getLastRow() <= 1) {
    var hr = sh.getRange(1, 1, 1, IND_HEADERS.length);
    hr.setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function unitMaps_() {
  var byId = {}, byName = {};
  getSheetObjects_(getSheet_('Master_Unit')).forEach(function(o) {
    var id = cellStr_(o.id_unit), nm = cellStr_(o.nama_unit);
    if (!id || !nm) return;
    byId[id] = nm; byName[nm.toLowerCase()] = id;
  });
  return { byId: byId, byName: byName };
}

function indRow_(o, units) {
  return {
    id_indikator: cellStr_(o.id_indikator),
    id_unit: cellStr_(o.id_unit),
    nama_unit: units.byId[cellStr_(o.id_unit)] || '',
    nama_indikator: cellStr_(o.nama_indikator),
    kategori: cellStr_(o.kategori),
    target_persen: (o.target_persen === '' || o.target_persen === null || o.target_persen === undefined) ? '' : Number(o.target_persen),
    formula_tipe: cellStr_(o.formula_tipe) === 'le' ? 'le' : 'ge',
    standar_teks: cellStr_(o.standar_teks),
    numerator_desc: cellStr_(o.numerator_desc),
    denominator_desc: cellStr_(o.denominator_desc),
    id_form: cellStr_(o.id_form)
  };
}

function getMasterIndikator(token) {
  requireAuth_(token);
  var units = unitMaps_();
  var rows = getSheetObjects_(getIndikatorSheet_()).filter(function(o) {
    return o.id_indikator && o.nama_indikator;
  }).map(function(o) { return indRow_(o, units); });
  // Urutkan: INM -> IMP RS -> IMP Unit -> lainnya (urutan semula dipertahankan dalam satu kategori)
  var rank = function(k) { var x = IND_KATEGORI.indexOf(k || 'Lainnya'); return x < 0 ? 99 : x; };
  return rows.map(function(r, n) { return [r, n]; }).sort(function(a, b) {
    return (rank(a[0].kategori) - rank(b[0].kategori)) || (a[1] - b[1]);
  }).map(function(x) { return x[0]; });
}

function cleanIndikator_(d, units) {
  d = d || {};
  var idUnit = String(d.id_unit || '');
  if (!idUnit && d.nama_unit) idUnit = units.byName[String(d.nama_unit).toLowerCase()] || '';
  if (!idUnit || !units.byId[idUnit]) throw new Error('Ruangan tidak ditemukan.');
  var nama = safeText_(String(d.nama_indikator || '').trim(), 300);
  if (!nama) throw new Error('Nama indikator wajib diisi.');
  var kat = IND_KATEGORI.indexOf(d.kategori) >= 0 ? d.kategori : 'Lainnya';
  var tgt = d.target_persen;
  if (tgt === '' || tgt === null || tgt === undefined) tgt = '';
  else {
    tgt = Number(String(tgt).replace(',', '.'));
    if (isNaN(tgt) || tgt < 0 || tgt > 100000) throw new Error('Target harus berupa angka.');
  }
  return {
    id_unit: idUnit, nama: nama, kategori: kat, target: tgt,
    tipe: d.formula_tipe === 'le' ? 'le' : 'ge',
    standar: safeText_(d.standar_teks, 100),
    num: safeText_(d.numerator_desc, 400),
    den: safeText_(d.denominator_desc, 400),
    form: /^[A-Za-z0-9_\-]{0,40}$/.test(String(d.id_form || '')) ? String(d.id_form || '') : ''
  };
}

// ID indikator otomatis: awalan per kategori + nomor urut, mis. INM-001, IMPRS-001, IMPUNIT-001
var IND_ID_PREFIX = { 'INM': 'INM', 'IMP RS': 'IMPRS', 'IMP Unit': 'IMPUNIT', 'SPM': 'SPM', 'PPI': 'PPI', 'Lainnya': 'LAIN' };

function indIdPrefix_(kat) { return IND_ID_PREFIX[kat] || 'LAIN'; }

function indIdFormat_(prefix, n) { return prefix + '-' + ('000' + n).slice(-3); }

// Nomor urut terbesar yang sudah dipakai untuk awalan tertentu
function nextIndSeq_(objs, prefix) {
  var max = 0, re = new RegExp('^' + prefix + '-(\\d+)$');
  objs.forEach(function(o) {
    var m = re.exec(String(o.id_indikator));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return max;
}

// Tambah (tanpa id_indikator) atau ubah (dengan id_indikator) satu indikator pada satu ruangan
function saveIndikator(token, d) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  return withLock_(function() {
    var sh = getIndikatorSheet_();
    var units = unitMaps_();
    var c = cleanIndikator_(d, units);
    var objs = getSheetObjects_(sh);
    var found = null, dup = false;
    objs.forEach(function(o) {
      if (d.id_indikator && String(o.id_indikator) === String(d.id_indikator)) found = o;
      if (String(o.id_unit) === c.id_unit && String(o.nama_indikator).trim().toLowerCase() === c.nama.toLowerCase() &&
          String(o.id_indikator) !== String(d.id_indikator || '')) dup = true;
    });
    if (dup) throw new Error('Indikator dengan nama yang sama sudah ada di ruangan ini.');
    if (d.id_indikator && !found) throw new Error('Indikator tidak ditemukan.');
    // ID otomatis: tetap bila indikator hanya diubah; dibuat baru bila indikator baru atau kategorinya berpindah
    var id;
    if (found && cellStr_(found.kategori) === c.kategori) id = String(found.id_indikator);
    else { var pf = indIdPrefix_(c.kategori); id = indIdFormat_(pf, nextIndSeq_(objs, pf) + 1); }
    var row = [id, c.id_unit, c.nama, c.kategori, c.target, c.tipe, c.standar, c.num, c.den, c.form];
    var r = found ? found._row : sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, row.length).setValues([row]);
    return { success: true, id_indikator: id };
  });
}

// Susun ulang seluruh ID indikator: berurutan INM -> IMP RS -> IMP Unit -> lainnya (lalu urutan ruangan)
function renumberIndikatorIds(token) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  return withLock_(function() {
    var sh = getIndikatorSheet_();
    var objs = getSheetObjects_(sh).filter(function(o) { return o.id_indikator && o.nama_indikator; });
    var unitOrder = {};
    getSheetObjects_(getSheet_('Master_Unit')).forEach(function(u, n) { unitOrder[cellStr_(u.id_unit)] = n; });
    var rank = function(k) { var x = IND_KATEGORI.indexOf(k || 'Lainnya'); return x < 0 ? 99 : x; };
    var list = objs.map(function(o) { return { o: o, k: cellStr_(o.kategori) || 'Lainnya' }; });
    list.sort(function(a, b) {
      return (rank(a.k) - rank(b.k)) ||
        ((unitOrder[cellStr_(a.o.id_unit)] || 0) - (unitOrder[cellStr_(b.o.id_unit)] || 0)) || (a.o._row - b.o._row);
    });
    var cnt = {}, newId = {};
    list.forEach(function(x) {
      var pf = indIdPrefix_(x.k);
      cnt[pf] = (cnt[pf] || 0) + 1;
      newId[x.o._row] = indIdFormat_(pf, cnt[pf]);
    });
    var last = sh.getLastRow();
    if (last < 2) return { success: true, jumlah: 0 };
    var rng = sh.getRange(2, 1, last - 1, 1), vals = rng.getValues();
    for (var i = 0; i < vals.length; i++) { if (newId[i + 2]) vals[i][0] = newId[i + 2]; }
    rng.setValues(vals);
    return { success: true, jumlah: list.length };
  });
}

// Kurangi: hapus satu indikator dari ruangan
function deleteIndikator(token, id) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  return withLock_(function() {
    var sh = getIndikatorSheet_();
    var objs = getSheetObjects_(sh);
    for (var i = objs.length - 1; i >= 0; i--) {
      if (String(objs[i].id_indikator) === String(id)) { sh.deleteRow(objs[i]._row); return { success: true }; }
    }
    throw new Error('Indikator tidak ditemukan.');
  });
}

/* ============================================================
 * TEMPLATE FORM (dapat diedit Komite Mutu)
 * ============================================================ */
function getFormTemplates(token) {
  requireAuth_(token);
  var out = [];
  getSheetObjects_(getSheet_('Master_Form')).forEach(function(o) {
    try {
      var def = JSON.parse(o.definisi_json);
      def.id = String(o.id_form);
      def.aktif = String(o.status_aktif) !== 'Nonaktif';
      out.push(def);
    } catch (e) { /* abaikan baris rusak */ }
  });
  return out;
}

function saveFormTemplate(token, def) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  if (!def || !/^[A-Za-z0-9_\-]{2,40}$/.test(def.id || '')) throw new Error('ID template harus 2-40 karakter (huruf, angka, - atau _).');
  if (!def.judul || !String(def.judul).trim()) throw new Error('Judul form wajib diisi.');
  if (!Array.isArray(def.kolom) || def.kolom.length < 1 || def.kolom.length > 30) throw new Error('Jumlah kolom harus 1 sampai 30.');

  def.judul = safeText_(def.judul, 200);
  def.kode = safeText_(def.kode, 40);
  def.targetOp = (def.targetOp === 'le' || def.targetOp === 'num' || def.targetOp === 'none') ? def.targetOp : 'ge'; // acuan target: ge (>=), le (<=), none (tanpa tanda, tetap %), atau num (angka saja)
  var json = JSON.stringify(def);
  if (json.length > 45000) throw new Error('Definisi template terlalu besar.');

  return withLock_(function() {
    var sheet = getSheet_('Master_Form');
    var objs = getSheetObjects_(sheet);
    var found = null;
    objs.forEach(function(o) { if (String(o.id_form) === def.id) found = o; });
    var row = [def.id, def.kode, def.judul, safeText_(def.kategori, 50), json, def.aktif === false ? 'Nonaktif' : 'Aktif', user.username, new Date()];
    if (found) sheet.getRange(found._row, 1, 1, row.length).setValues([row]);
    else sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    return { success: true };
  });
}

function deleteFormTemplate(token, id) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  return withLock_(function() {
    var sheet = getSheet_('Master_Form');
    var objs = getSheetObjects_(sheet);
    for (var i = objs.length - 1; i >= 0; i--) {
      if (String(objs[i].id_form) === String(id)) sheet.deleteRow(objs[i]._row);
    }
    return { success: true };
  });
}

/* ============================================================
 * PROFIL & PRIORITAS (penetapan SK, kamus profil indikator, indikator per unit)
 * Sheet Master_Profil dibuat otomatis. tipe: 'sk' (penetapan SK), 'ind' (indikator + profil), 'unit' (daftar kode indikator per ruangan)
 * ============================================================ */
var PROFIL_HEADERS = ['id_item', 'tipe', 'data_json', 'updated_by', 'updated_at'];
var PROFIL_TIPE = ['sk', 'ind', 'unit'];

function getProfilSheet_() {
  var ss = getDbSpreadsheet();
  var sh = ss.getSheetByName('Master_Profil');
  if (!sh) {
    sh = ss.insertSheet('Master_Profil');
    sh.appendRow(PROFIL_HEADERS);
    sh.getRange(1, 1, 1, PROFIL_HEADERS.length).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getProfilData(token) {
  requireAuth_(token);
  var out = [];
  getSheetObjects_(getProfilSheet_()).forEach(function(o) {
    try {
      var key = String(o.id_item), pos = key.indexOf(':');
      out.push({ tipe: String(o.tipe), id: key.slice(pos + 1), data: JSON.parse(o.data_json) });
    } catch (e) { /* abaikan baris rusak */ }
  });
  return out;
}

function saveProfilItem(token, tipe, id, data) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  if (PROFIL_TIPE.indexOf(tipe) === -1) throw new Error('Jenis data profil tidak valid.');
  id = safeText_(id, 80);
  if (!id) throw new Error('ID data profil wajib diisi.');
  if (tipe === 'ind') {
    if (!data || !data.k || String(data.k) !== id) throw new Error('Kode indikator tidak sesuai.');
    if (!data.n || !String(data.n).trim()) throw new Error('Nama indikator wajib diisi.');
    if (IND_KATEGORI.indexOf(data.kat) === -1) throw new Error('Kategori indikator tidak valid.');
    data.a = (data.a === 'le' || data.a === 'num' || data.a === 'none') ? data.a : 'ge';
  } else if (tipe === 'unit') {
    if (!Array.isArray(data) || data.length > 200) throw new Error('Daftar indikator unit tidak valid.');
  } else if (!data || typeof data !== 'object') {
    throw new Error('Data penetapan SK tidak valid.');
  }
  var json = JSON.stringify(data);
  if (json.length > 45000) throw new Error('Data terlalu besar untuk disimpan.');

  return withLock_(function() {
    var sheet = getProfilSheet_();
    var key = tipe + ':' + id, found = null;
    getSheetObjects_(sheet).forEach(function(o) { if (String(o.id_item) === key) found = o; });
    var row = [key, tipe, json, user.username, new Date()];
    if (found) sheet.getRange(found._row, 1, 1, row.length).setValues([row]);
    else sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    return { success: true };
  });
}

function deleteProfilItem(token, tipe, id) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  return withLock_(function() {
    var sheet = getProfilSheet_(), key = tipe + ':' + id;
    var objs = getSheetObjects_(sheet);
    for (var i = objs.length - 1; i >= 0; i--) {
      if (String(objs[i].id_item) === key) sheet.deleteRow(objs[i]._row);
    }
    return { success: true };
  });
}

/* ============================================================
 * LAPORAN FORM
 * ============================================================ */
function reportMeta_(o) {
  return {
    id_laporan: cellStr_(o.id_laporan),
    id_form: cellStr_(o.id_form),
    nama_form: cellStr_(o.nama_form),
    bulan_tahun: bulanStr_(o.bulan_tahun),
    unit: cellStr_(o.unit),
    num: Number(o.num) || 0,
    denum: Number(o.denum) || 0,
    capaian_persen: Number(o.capaian_persen) || 0,
    status: cellStr_(o.status),
    catatan_validator: cellStr_(o.catatan_validator),
    created_by: cellStr_(o.created_by),
    updated_at: cellStr_(o.updated_at)
  };
}

function saveFormReport(token, rep) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu', 'PJ Mutu Unit']);
  if (!rep || !rep.id_form) throw new Error('Template form tidak valid.');
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(rep.bulan_tahun || '')) throw new Error('Periode bulan tidak valid.');
  var unit = safeText_(rep.unit, 100);
  if (!unit) throw new Error('Unit/ruangan wajib diisi.');
  if (user.role === 'PJ Mutu Unit' && unit !== user.unit_kerja) {
    throw new Error('Anda hanya dapat melaporkan untuk unit ' + user.unit_kerja + '.');
  }
  if (!Array.isArray(rep.rows) || rep.rows.length > 300) throw new Error('Data baris tidak valid (maksimal 300 baris).');

  var dataJson = JSON.stringify(rep.rows);
  var metaJson = JSON.stringify(rep.meta || {});
  if (dataJson.length > 45000 || metaJson.length > 5000) throw new Error('Data terlalu besar untuk disimpan.');

  var num = Math.max(0, Number(rep.num) || 0);
  var denum = Math.max(0, Number(rep.denum) || 0);
  var capaian = rep.angka === true ? Math.round(num * 100) / 100 : (denum > 0 ? Math.round(num / denum * 10000) / 100 : 0); // angka === true: template "angka saja"
  var aksi = rep.aksi === 'ajukan' ? 'ajukan' : 'simpan';

  return withLock_(function() {
    var sheet = getSheet_('Laporan_Form');
    var objs = getSheetObjects_(sheet);
    var existing = null;
    if (rep.id_laporan) objs.forEach(function(o) { if (String(o.id_laporan) === String(rep.id_laporan)) existing = o; });

    var status = aksi === 'ajukan' ? 'Diajukan' : 'Draft';
    var id = rep.id_laporan, createdBy = user.username, createdAt = new Date(), validatedBy = '', catatan = '';

    if (existing) {
      if (user.role !== 'Komite Mutu' && cellStr_(existing.unit) !== user.unit_kerja) throw new Error('Anda tidak berhak mengubah laporan unit lain.');
      if (existing.status === 'Tervalidasi') {
        if (user.role !== 'Komite Mutu') throw new Error('Laporan sudah tervalidasi dan terkunci.');
        if (aksi !== 'ajukan') status = 'Tervalidasi';
      }
      id = existing.id_laporan; createdBy = existing.created_by; createdAt = existing.created_at;
      validatedBy = existing.validated_by; catatan = existing.catatan_validator;
      if (status !== 'Tervalidasi') { validatedBy = ''; }
    } else {
      id = newId_('LF');
    }

    var row = [id, safeText_(rep.id_form, 60), safeText_(rep.nama_form, 250), rep.bulan_tahun, unit, dataJson, metaJson,
               num, denum, capaian, status, catatan, createdBy, validatedBy, createdAt, new Date()];
    var r = existing ? existing._row : sheet.getLastRow() + 1;
    sheet.getRange(r, 4).setNumberFormat('@');
    sheet.getRange(r, 1, 1, row.length).setValues([row]);
    return { success: true, id_laporan: id, status: status, capaian_persen: capaian };
  });
}

function listFormReports(token, filter) {
  requireAuth_(token);
  filter = filter || {};
  var months = parseInt(filter.months, 10) || 0;
  var cutoff = '';
  if (months > 0) {
    var d = new Date();
    d.setDate(1);
    d.setMonth(d.getMonth() - (months - 1));
    cutoff = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  var out = [];
  getSheetObjects_(getSheet_('Laporan_Form')).forEach(function(o) {
    var m = reportMeta_(o);
    if (!m.id_laporan) return;
    if (filter.unit && filter.unit !== 'ALL' && m.unit !== filter.unit) return;
    if (filter.status && m.status !== filter.status) return;
    if (cutoff && m.bulan_tahun < cutoff) return;
    out.push(m);
  });
  out.sort(function(a, b) {
    if (a.bulan_tahun !== b.bulan_tahun) return a.bulan_tahun < b.bulan_tahun ? 1 : -1;
    return a.updated_at < b.updated_at ? 1 : -1;
  });
  return out.slice(0, 500);
}

// Data capaian setahun seluruh ruangan (tanpa isi baris), untuk dashboard Capaian Mutu Tahunan
function getCapaianTahunan(token, tahun) {
  requireAuth_(token);
  tahun = String(tahun || '').trim();
  if (!/^\d{4}$/.test(tahun)) throw new Error('Tahun tidak valid.');
  var out = [];
  getSheetObjects_(getSheet_('Laporan_Form')).forEach(function(o) {
    var m = reportMeta_(o);
    if (!m.id_laporan || m.bulan_tahun.slice(0, 4) !== tahun) return;
    out.push({ unit: m.unit, bulan_tahun: m.bulan_tahun, id_form: m.id_form, num: m.num, denum: m.denum, status: m.status });
  });
  return out;
}

// Rekap bulanan satu ruangan: semua laporan form (beserta isi barisnya) pada bulan tertentu
function getRekapBulananUnit(token, unit, bulan) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu', 'PJ Mutu Unit']);
  unit = String(unit || '').trim();
  if (!unit) throw new Error('Ruangan wajib dipilih.');
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(String(bulan || ''))) throw new Error('Periode bulan tidak valid.');
  if (user.role !== 'Komite Mutu' && unit !== user.unit_kerja) throw new Error('Anda hanya dapat melihat formulir bulanan unit ' + user.unit_kerja + '.');
  var out = [];
  getSheetObjects_(getSheet_('Laporan_Form')).forEach(function(o) {
    var m = reportMeta_(o);
    if (!m.id_laporan || m.unit !== unit || m.bulan_tahun !== bulan) return;
    try { m.rows = JSON.parse(o.data_json); } catch (e) { m.rows = []; }
    try { m.meta = JSON.parse(o.meta_json); } catch (e) { m.meta = {}; }
    out.push(m);
  });
  return out;
}

function getFormReport(token, id) {
  var user = requireAuth_(token);
  var found = null;
  getSheetObjects_(getSheet_('Laporan_Form')).forEach(function(o) { if (String(o.id_laporan) === String(id)) found = o; });
  if (!found) throw new Error('Laporan tidak ditemukan.');
  if (user.role !== 'Komite Mutu' && cellStr_(found.unit) !== user.unit_kerja) throw new Error('Anda tidak berhak membuka laporan unit lain.');
  var rep = reportMeta_(found);
  try { rep.rows = JSON.parse(found.data_json); } catch (e) { rep.rows = []; }
  try { rep.meta = JSON.parse(found.meta_json); } catch (e) { rep.meta = {}; }
  return rep;
}

function deleteFormReport(token, id) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu', 'PJ Mutu Unit']);
  return withLock_(function() {
    var sheet = getSheet_('Laporan_Form');
    var objs = getSheetObjects_(sheet);
    for (var i = 0; i < objs.length; i++) {
      if (String(objs[i].id_laporan) === String(id)) {
        if (user.role !== 'Komite Mutu') {
          if (cellStr_(objs[i].unit) !== user.unit_kerja) throw new Error('Anda tidak berhak menghapus laporan unit lain.');
          if (objs[i].status !== 'Draft' && objs[i].status !== 'Ditolak') throw new Error('Hanya laporan berstatus Draft/Ditolak yang dapat dihapus.');
        }
        sheet.deleteRow(objs[i]._row);
        return { success: true };
      }
    }
    throw new Error('Laporan tidak ditemukan.');
  });
}

function setValidasiForm(token, id, status, catatan) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  if (['Tervalidasi', 'Ditolak'].indexOf(status) === -1) throw new Error('Status tidak valid.');
  return withLock_(function() {
    var sheet = getSheet_('Laporan_Form');
    var objs = getSheetObjects_(sheet);
    for (var i = 0; i < objs.length; i++) {
      if (String(objs[i].id_laporan) === String(id)) {
        sheet.getRange(objs[i]._row, 11).setValue(status);                      // status
        sheet.getRange(objs[i]._row, 12).setValue(safeText_(catatan, 500));      // catatan_validator
        sheet.getRange(objs[i]._row, 14).setValue(user.username);                // validated_by
        sheet.getRange(objs[i]._row, 16).setValue(new Date());                   // updated_at
        return { success: true };
      }
    }
    throw new Error('Laporan tidak ditemukan.');
  });
}

/* ============================================================
 * SIMPAN & RESTORE DATABASE
 * Salinan disimpan di sheet "Backup_DB" (dipecah per 45.000 karakter karena batas isi 1 sel).
 * Yang disalin: Master_Unit, Master_Indikator, Master_Form, Laporan_Form, Laporan_Bulanan, Laporan_Harian.
 * Sheet Users (akun & password) sengaja tidak disalin agar restore tidak mengembalikan password lama.
 * ============================================================ */
var BACKUP_SHEETS = ['Master_Unit', 'Master_Indikator', 'Master_Form', 'Master_Profil', 'Laporan_Form', 'Laporan_Bulanan', 'Laporan_Harian'];
var BACKUP_HEADERS = ['id_backup', 'tanggal', 'bulan', 'waktu', 'dibuat_oleh', 'keterangan', 'ringkasan', 'chunk_no', 'chunk_total', 'data_json'];
var BACKUP_CHUNK = 45000;

function getBackupSheet_() {
  var ss = getDbSpreadsheet();
  var sh = ss.getSheetByName('Backup_DB');
  if (!sh) {
    sh = ss.insertSheet('Backup_DB');
    sh.appendRow(BACKUP_HEADERS);
    sh.getRange(1, 1, 1, BACKUP_HEADERS.length).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function writeBackup_(user, ket) {
  var ss = getDbSpreadsheet();
  var snap = {}, ringkas = [];
  BACKUP_SHEETS.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() === 0) return;
    var vals = sh.getDataRange().getValues();
    snap[name] = vals;
    ringkas.push(name.replace(/_/g, ' ') + ': ' + (vals.length - 1));
  });
  var json = JSON.stringify(snap);
  var now = new Date();
  var tgl = Utilities.formatDate(now, 'Asia/Jayapura', 'yyyy-MM-dd');
  var wkt = Utilities.formatDate(now, 'Asia/Jayapura', 'HH:mm:ss');
  var id = newId_('BK');
  var total = Math.max(1, Math.ceil(json.length / BACKUP_CHUNK));
  var oleh = user.nama_lengkap || user.username || '';
  var ring = ringkas.join(', ');
  var rows = [];
  for (var i = 0; i < total; i++) {
    rows.push([id, tgl, tgl.slice(0, 7), wkt, oleh, safeText_(ket, 200), ring, i + 1, total, json.substr(i * BACKUP_CHUNK, BACKUP_CHUNK)]);
  }
  var sh2 = getBackupSheet_();
  var rg = sh2.getRange(sh2.getLastRow() + 1, 1, rows.length, BACKUP_HEADERS.length);
  rg.setNumberFormat('@'); // semua teks, agar tanggal/bulan/potongan JSON tidak diubah Sheets
  rg.setValues(rows);
  return { success: true, id_backup: id, tanggal: tgl, waktu: wkt, ringkasan: ring };
}

function simpanDatabase(token, keterangan) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  return withLock_(function() { return writeBackup_(user, keterangan); });
}

// filter: { bulan: 'yyyy-MM' } atau { tanggal: 'yyyy-MM-dd' } (tanggal diutamakan); kosong = semua
function listBackupDB(token, filter) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  filter = filter || {};
  var sh = getBackupSheet_();
  if (sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  var out = [];
  vals.forEach(function(v) {
    if (String(v[7]) !== '1') return; // satu baris per backup (bagian pertama)
    var tgl = cellStr_(v[1]).slice(0, 10), bln = bulanStr_(v[2]), wkt = String(v[3] instanceof Date ? Utilities.formatDate(v[3], Session.getScriptTimeZone(), 'HH:mm:ss') : v[3]);
    if (filter.tanggal) { if (tgl !== String(filter.tanggal)) return; }
    else if (filter.bulan) { if (bln !== String(filter.bulan)) return; }
    out.push({ id_backup: String(v[0]), tanggal: tgl, bulan: bln, waktu: wkt, dibuat_oleh: String(v[4]), keterangan: String(v[5]), ringkasan: String(v[6]) });
  });
  out.sort(function(a, b) { return (a.tanggal + a.waktu) < (b.tanggal + b.waktu) ? 1 : -1; });
  return out.slice(0, 100);
}

function restoreCell_(v) {
  if (typeof v === 'string') {
    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z$/.test(v)) return new Date(v); // tanggal hasil JSON
    if (v.charAt(0) === '=') return "'" + v; // jangan jadi formula
  }
  return v;
}

function restoreDatabase(token, idBackup) {
  var user = requireAuth_(token);
  requireRole_(user, ['Komite Mutu']);
  return withLock_(function() {
    var sh = getBackupSheet_();
    if (sh.getLastRow() < 2) throw new Error('Belum ada salinan database.');
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, BACKUP_HEADERS.length).getValues();
    var parts = [], total = 0;
    vals.forEach(function(v) {
      if (String(v[0]) !== String(idBackup)) return;
      parts[Number(v[7]) - 1] = String(v[9]);
      total = Number(v[8]);
    });
    if (!total) throw new Error('Salinan database tidak ditemukan.');
    for (var i = 0; i < total; i++) if (parts[i] === undefined) throw new Error('Salinan database tidak lengkap (bagian ' + (i + 1) + ' hilang).');
    var snap;
    try { snap = JSON.parse(parts.join('')); } catch (e) { throw new Error('Salinan database rusak dan tidak bisa dibaca.'); }

    writeBackup_(user, 'Otomatis sebelum restore ' + idBackup); // pengaman: kondisi saat ini disimpan dulu

    var ss = getDbSpreadsheet(), done = [];
    BACKUP_SHEETS.forEach(function(name) {
      var data = snap[name];
      if (!data || !data.length) return;
      var target = ss.getSheetByName(name) || ss.insertSheet(name);
      var cols = data[0].length;
      var rows = data.map(function(r) { return r.map(restoreCell_); });
      target.clearContents();
      if (target.getMaxRows() < rows.length) target.insertRowsAfter(target.getMaxRows(), rows.length - target.getMaxRows());
      if (target.getMaxColumns() < cols) target.insertColumnsAfter(target.getMaxColumns(), cols - target.getMaxColumns());
      data[0].forEach(function(h, c) { if (h === 'bulan_tahun') target.getRange(1, c + 1, target.getMaxRows(), 1).setNumberFormat('@'); });
      target.getRange(1, 1, rows.length, cols).setValues(rows);
      done.push(name.replace(/_/g, ' ') + ': ' + (rows.length - 1));
    });
    return { success: true, ringkasan: done.join(', ') };
  });
}
