# Perpustakaan Komite Mutu

1. Google Sheet > Extensions > Apps Script: tempel isi `Code.gs`. Jalankan `setup()` (catat password admin di Execution log), lalu `setupTriggers()`.
2. Deploy > New deployment > Web app: Execute as **Me**, Who has access **Anyone**. Salin URL `/exec`.
3. Di `index.html` ganti nilai `API` dengan URL tersebut.
4. Unggah `index.html`, `manifest.json`, `sw.js`, `icon-192.png`, `icon-512.png` ke repo GitHub (aktifkan Pages). **Jangan** unggah `Code.gs` ke repo publik.
5. Lupa password admin: jalankan `resetAdmin()` dan lihat Execution log.
