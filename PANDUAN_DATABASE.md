# Menghubungkan kembali database lama

## 1. Cari ID spreadsheet lama
Buka Google Sheet database lama. ID ada di URL: docs.google.com/spreadsheets/d/**ID_INI**/edit

## 2. Di project Apps Script
Pastikan ada file: code.gs, api.gs, mutu.gs, komite.gs (+ Index.html).
Buka code.gs, tambahkan fungsi ini lalu Run SEKALI (ganti ID):

    function hubungkanDatabaseLama() {
      Logger.log(setSpreadsheetId('ID_SPREADSHEET_LAMA'));
    }

Setujui izin akses bila diminta.

## 3. Deploy ulang
Deploy > Manage deployments > Edit (ikon pensil) > Version: New version > Deploy.
Atur: Execute as = Me, Who has access = Anyone. Salin URL yang berakhiran /exec.

## 4. Isi config.js di GitHub
window.KMKP_API_URL = "URL_WEB_APP_/exec";
