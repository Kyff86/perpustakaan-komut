# Pasang Perpustakaan Komite Mutu
1. Buat Google Sheet kosong > Ekstensi > Apps Script. Tempel isi `Code.gs`.
2. Jalankan `setup()` (izinkan akses). Lihat Logger: username `admin` + password acak. Lalu jalankan `setupTriggers()`.
3. Deploy > New deployment > Web app: Execute as **Me**, Who has access **Anyone**. Salin URL `/exec`.
4. Di `index.html` ganti `API` dengan URL tersebut.
5. Unggah `index.html, sw.js, manifest.json, icon-*.png` ke GitHub Pages / Netlify (HTTPS).
6. Buka situs, login, ganti password, tambah kategori & user. Pasang ke layar utama.
Catatan: jangan bagikan Google Sheet ke staf; semua akses lewat aplikasi.
