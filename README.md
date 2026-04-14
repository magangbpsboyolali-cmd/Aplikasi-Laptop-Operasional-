# Aplikasi Laptop Operasional - BPS

Sistem dashboard web untuk pendataan laptop dan peralatan inventaris kantor dengan fitur QR Code dan integrasi Google Spreadsheet.

## 🎯 Fitur Utama

1. **Dashboard** - Tampilan ringkasan statistik inventaris
2. **Peminjaman** - Form dan daftar peminjaman laptop
3. **Pengembalian** - Proses pengembalian dengan laporan kondisi
4. **Daftar Inventaris** - Kelola data laptop dan peralatan
5. **Laporan Kerusakan** - Catat dan tracking kerusakan
6. **QR Code Generator** - Generate QR untuk akses cepat
7. **Riwayat** - History peminjaman dengan export CSV

## 📁 Struktur File

```
projek1/
├── index.html              # Halaman utama dashboard
├── css/
│   └── style.css           # Stylesheet
├── js/
│   ├── config.js           # Konfigurasi (URL Google Apps Script)
│   └── app.js              # Logic aplikasi
├── google-apps-script.js   # Kode untuk Google Apps Script
└── README.md               # Dokumentasi ini
```

## 🚀 Cara Menjalankan

### Mode Local (Testing)
1. Buka file `index.html` langsung di browser
2. Data akan disimpan di Local Storage browser
3. Cocok untuk testing dan demo

### Mode dengan Google Spreadsheet

#### Langkah 1: Buat Google Spreadsheet
1. Buka [Google Sheets](https://sheets.google.com)
2. Buat spreadsheet baru
3. Buat 4 sheet dengan nama:
   - `Inventaris`
   - `Peminjaman`
   - `Kerusakan`
   - `Riwayat`
4. Copy ID spreadsheet dari URL (bagian antara `/d/` dan `/edit`)

#### Langkah 2: Setup Google Apps Script
1. Di spreadsheet, buka menu **Extensions > Apps Script**
2. Hapus kode default dan paste semua kode dari file `google-apps-script.js`
3. Ganti `SPREADSHEET_ID` dengan ID spreadsheet Anda
4. Simpan project (Ctrl+S)
5. Jalankan fungsi `initializeSheets()` sekali untuk setup header

#### Langkah 3: Deploy Web App
1. Klik **Deploy > New deployment**
2. Klik ikon gear ⚙️, pilih **Web app**
3. Isi konfigurasi:
   - Description: `Inventaris API`
   - Execute as: `Me`
   - Who has access: `Anyone`
4. Klik **Deploy**
5. Authorize akses jika diminta
6. Copy URL Web App yang diberikan

#### Langkah 4: Konfigurasi Website
1. Buka file `js/config.js`
2. Ganti `YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL` dengan URL Web App
3. Ganti `YOUR_SPREADSHEET_ID` dengan ID spreadsheet
4. Set `USE_LOCAL_STORAGE = false`

## 📱 Penggunaan QR Code

1. Buka menu **QR Code** di dashboard
2. Download QR Code untuk:
   - Form Peminjaman
   - Form Pengembalian
   - Laporan Kerusakan
3. Cetak dan tempel QR di laptop/area kantor
4. Karyawan dapat scan untuk akses cepat

## 💡 Tips

- **Testing**: Gunakan Local Storage dulu untuk testing
- **Backup**: Export data secara berkala dari menu Riwayat
- **Mobile**: Website responsive, bisa diakses via HP
- **QR Per Laptop**: Generate QR khusus per laptop di menu QR Code

## 🔧 Troubleshooting

### Data tidak tersimpan ke Spreadsheet
- Pastikan URL Apps Script sudah benar
- Cek izin akses di Apps Script sudah di-authorize
- Buka Console browser (F12) untuk melihat error

### QR Code tidak muncul
- Pastikan koneksi internet aktif (library QRCode dari CDN)
- Coba refresh halaman

### Layout berantakan di mobile
- Gunakan sidebar toggle untuk menyembunyikan menu
- Scroll horizontal untuk melihat tabel lengkap

## 📝 Customization

### Mengubah bagian/divisi
Edit options di `index.html` pada select dengan id `bagian`:
```html
<option value="Nama Bagian">Nama Bagian</option>
```

### Mengubah jenis kerusakan
Edit checkbox di `index.html` pada bagian `kerusakanGroup`

### Mengubah warna tema
Edit variabel CSS di `css/style.css`:
```css
:root {
    --primary: #4F46E5;  /* Warna utama */
    ...
}
```

## 📞 Support

Untuk bantuan lebih lanjut, hubungi tim IT atau admin sistem.

---
© 2026 Sistem Pendataan Inventaris - BPS
