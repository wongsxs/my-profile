# 🚀 FULL GOOGLE APPS SCRIPT (GAS) WEB APP - ABSENSI KEGIATAN GAMKI

Aplikasi web **Absensi Kegiatan GAMKI** kini 100% berjalan sepenuhnya di dalam **Google Apps Script (GAS)**, terintegrasi langsung dengan **Google Sheets** sebagai basis data dan **Google Drive** sebagai penyimpanan foto bukti.

---

## 🌟 Keunggulan Versi Full GAS:
1. **Tidak Membutuhkan Hosting/Server Tambahan**: Gratis di-host langsung oleh Google.
2. **Real-Time Data**:
   - Form Absensi -> Langsung menyimpan ke Google Sheets & Google Drive.
   - Riwayat Absensi -> Membaca langsung dari Google Sheets.
   - Ringkasan Statistik -> Dihitung otomatis dari Google Sheets.
   - Admin Panel -> Menambah, mengedit, dan menghapus data anggota & presensi langsung pada Google Sheets.

---

## 📁 Berkas Berada di Folder `gas/`:

1. [`gas/Code.gs`](file:///c:/Users/Jaja/Documents/project%20absensi/gas/Code.gs) -> Backend Apps Script (Spreadsheet CRUD & Upload Foto Drive)
2. [`gas/Index.html`](file:///c:/Users/Jaja/Documents/project%20absensi/gas/Index.html) -> Tampilan 4 Tab Lengkap (Form, Riwayat, Statistik, Admin Panel)
3. [`gas/Style.html`](file:///c:/Users/Jaja/Documents/project%20absensi/gas/Style.html) -> Desain CSS Glassmorphism GAMKI
4. [`gas/JavaScript.html`](file:///c:/Users/Jaja/Documents/project%20absensi/gas/JavaScript.html) -> Integrasi `google.script.run` 100% Backend-Frontend

---

## 🛠️ Cara Deploy ke Google Apps Script:

1. Buka [Google Sheets](https://sheets.google.com) Anda -> Buat dokumen baru **"Database Absensi GAMKI"**.
2. Klik **Ekstensi (Extensions)** -> **Apps Script**.
3. Buat 4 berkas di editor Apps Script:
   - `Code.gs` (Salin isi dari [`gas/Code.gs`](file:///c:/Users/Jaja/Documents/project%20absensi/gas/Code.gs))
   - `Index.html` (Salin isi dari [`gas/Index.html`](file:///c:/Users/Jaja/Documents/project%20absensi/gas/Index.html))
   - `Style.html` (Salin isi dari [`gas/Style.html`](file:///c:/Users/Jaja/Documents/project%20absensi/gas/Style.html))
   - `JavaScript.html` (Salin isi dari [`gas/JavaScript.html`](file:///c:/Users/Jaja/Documents/project%20absensi/gas/JavaScript.html))
4. Klik **Terapkan (Deploy)** -> **Penerapan Baru (New deployment)**.
5. Pilih type **Aplikasi web (Web app)**.
6. Setel *Akses (Who has access)* ke **Siapa saja (Anyone)**.
7. Klik **Terapkan (Deploy)** -> Izinkan Akses (*Authorize access*).
8. Salin **URL Web App** dan bagikan ke pengurus/anggota!

---

## 🔑 Informasi PIN Admin Panel:
- **PIN Admin**: **`1962`**
- Tab **Admin Panel** memungkinkan Anda mengedit/menambah nama anggota dan memodifikasi baris presensi langsung ke Google Sheets.
