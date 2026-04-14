# InvenTrack - Sistem Pendataan Laptop BPS Kab. Boyolali

## 🚀 Cara Setup

### 1. Buat Google Spreadsheet
1. Buka [Google Drive](https://drive.google.com)
2. Buat **Spreadsheet baru**
3. Copy **ID Spreadsheet** dari URL:
   ```
   https://docs.google.com/spreadsheets/d/[ID_INI]/edit
   ```

### 2. Setup Google Apps Script
1. Di Spreadsheet, klik **Extensions > Apps Script**
2. Hapus semua kode di editor
3. Copy-paste seluruh isi file `google-apps-script.js` ke editor
4. Ganti `SPREADSHEET_ID` di baris paling atas dengan ID spreadsheet Anda
5. Klik **▶ Run** pada fungsi `setupSheets` (jalankan sekali)
   - Klik "Review permissions" → pilih akun Google → "Advanced" → "Go to..."
6. Setelah berhasil, spreadsheet akan memiliki 4 sheet:
   - **Data Laptop**: ID, MERK, TYPE, NOP, STATUS
   - **DATA PEGAWAI**: NO, NIP, NAMA, TIM(DIVISI)
   - **Data Peminjaman**: ID, LAPTOP_ID, NAMA_PEMINJAM, NIP, DIVISI, KEPERLUAN, DESKRIPSI_KEPERLUAN, TGL_PINJAM, TGL_KEMBALI_RENCANA, STATUS
   - **Data Pengembalian**: ID, PEMINJAMAN_ID, LAPTOP_ID, NAMA_PEMINJAM, TGL_PINJAM, TGL_KEMBALI_RENCANA, TGL_REALISASI_PENGEMBALIAN, KONDISI, CATATAN

### 3. Deploy Web App
1. Di Apps Script Editor, klik **Deploy > New Deployment**
2. Klik ⚙ Settings → pilih **Web App**
3. Isi:
   - Description: `InvenTrack API`
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Klik **Deploy**
5. **Copy URL** yang muncul (format: `https://script.google.com/macros/s/.../exec`)

### 4. Konfigurasi Aplikasi
1. Buka file `js/config.js`
2. Paste URL Web App ke `APPS_SCRIPT_URL`:
   ```javascript
   APPS_SCRIPT_URL: 'https://script.google.com/macros/s/.../exec',
   ```

### 5. Isi Data Awal di Spreadsheet
1. Buka spreadsheet
2. Isi sheet **Data Laptop** dengan data laptop yang ada
3. Isi sheet **DATA PEGAWAI** dengan data pegawai

### 6. Jalankan Aplikasi
- Buka `index.html` di browser, atau
- Gunakan live server (VS Code: "Go Live")

---

## 📂 Struktur File

```
PROJEK1/
├── index.html              # Halaman utama
├── google-apps-script.js   # Kode untuk Google Apps Script
├── SETUP.md                # Panduan setup
├── css/
│   └── style.css           # Stylesheet
└── js/
    ├── config.js           # Konfigurasi URL & sheet names
    └── app.js              # Logika aplikasi utama
```

## 📊 Alur Kerja

### Dashboard
- Menampilkan statistik: Total Laptop, Tersedia, Dipinjam, Rusak
- Button cepat: Peminjaman & Pengembalian
- Tabel riwayat peminjaman (bisa difilter: laptop, nama, rentang tanggal)

### Peminjaman
1. Pilih laptop (hanya yang tersedia)
2. Pilih nama pegawai (otomatis isi divisi)
3. Pilih keperluan + deskripsi
4. Isi tanggal pinjam & perkiraan kembali
5. Submit → data masuk ke sheet Peminjaman, status laptop jadi "Dipinjam"

### Pengembalian
1. Pilih nama pegawai dari dropdown
2. Muncul daftar peminjaman aktif pegawai tersebut
3. Klik "Pilih" pada laptop yang dikembalikan
4. Muncul detail peminjaman + form tanggal realisasi
5. Submit → data masuk ke sheet Pengembalian, status laptop kembali "Tersedia" (atau "Rusak")
