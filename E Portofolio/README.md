# E-Portofolio PPG Prajabatan

Aplikasi web E-Portofolio untuk Program Profesi Guru (PPG) Prajabatan. Aplikasi ini dirancang untuk mendokumentasikan perjalanan profesional, pencapaian, dan pengembangan kompetensi guru calon profesional.

## Fitur

- **Profil Peserta**: Menampilkan informasi profil peserta PPG Prajabatan
- **Riwayat Pendidikan**: Timeline pendidikan dari SMA hingga PPG
- **Pengalaman Mengajar**: Dokumentasi pengalaman mengajar dan PPL
- **Prestasi & Sertifikat**: Showcase pencapaian dan sertifikat
- **Dokumen Portofolio**: Repository dokumen RPP, laporan, materi pembelajaran, dll.
- **Formulir Kontak**: Fitur komunikasi interaktif
- **Desain Responsif**: Tampilan optimal di berbagai perangkat
- **Animasi Modern**: Transisi dan efek visual yang menarik

## Teknologi

- HTML5
- CSS3 (dengan Flexbox dan Grid)
- Vanilla JavaScript
- Font Awesome (untuk ikon)
- Google Fonts (Poppins)

## Cara Menjalankan di Localhost

### Opsi 1: Menggunakan Python (Rekomendasi)

Jika Anda memiliki Python terinstall:

1. Buka terminal/command prompt
2. Navigasi ke folder proyek:
   ```bash
   cd "c:\Users\HP\Documents\TUGAS PPG PRAJAB\E Portofolio"
   ```
3. Jalankan server Python:
   
   Untuk Python 3:
   ```bash
   python -m http.server 8000
   ```
   
   Atau untuk Python 2:
   ```bash
   python -m SimpleHTTPServer 8000
   ```

4. Buka browser dan akses: `http://localhost:8000`

### Opsi 2: Menggunakan Node.js

Jika Anda memiliki Node.js terinstall:

1. Install http-server secara global:
   ```bash
   npm install -g http-server
   ```

2. Navigasi ke folder proyek:
   ```bash
   cd "c:\Users\HP\Documents\TUGAS PPG PRAJAB\E Portofolio"
   ```

3. Jalankan server:
   ```bash
   http-server -p 8000
   ```

4. Buka browser dan akses: `http://localhost:8000`

### Opsi 3: Menggunakan PHP

Jika Anda memiliki PHP terinstall:

1. Navigasi ke folder proyek:
   ```bash
   cd "c:\Users\HP\Documents\TUGAS PPG PRAJAB\E Portofolio"
   ```

2. Jalankan server PHP:
   ```bash
   php -S localhost:8000
   ```

3. Buka browser dan akses: `http://localhost:8000`

### Opsi 4: Membuka Langsung di Browser

Anda juga dapat membuka file `index.html` langsung di browser tanpa server, namun beberapa fitur mungkin tidak berfungsi sepenuhnya.

1. Buka File Explorer
2. Navigasi ke folder proyek
3. Double-click pada file `index.html`

## Struktur File

```
E Portofolio/
├── index.html          # Halaman utama
├── styles.css          # File styling CSS
├── script.js           # File JavaScript untuk interaktivitas
└── README.md           # Dokumentasi proyek
```

## Kustomisasi

### Mengubah Informasi Profil

Edit file `index.html` dan cari bagian berikut:

- **Nama dan Foto**: Di dalam `<div class="profile-card">`
- **Informasi Kontak**: Di dalam `<div class="info-grid">`
- **Email dan Telepon**: Di bagian contact section

### Mengubah Warna Tema

Edit file `styles.css` dan ubah variabel CSS di bagian `:root`:

```css
:root {
    --primary-color: #2563eb;      /* Warna utama */
    --secondary-color: #1e40af;    /* Warna sekunder */
    --accent-color: #f59e0b;       /* Warna aksen */
    /* ... variabel lainnya */
}
```

### Menambah Dokumen

Edit file `index.html` di bagian `<section id="documents">` dan tambahkan card baru sesuai format yang ada.

## Fitur Interaktif

- **Navigasi Smooth Scroll**: Klik menu untuk scroll ke bagian yang diinginkan
- **Mobile Menu**: Menu responsif untuk tampilan mobile
- **Animasi Scroll**: Card muncul dengan animasi saat di-scroll
- **Skill Bar Animation**: Progress bar teranimasi saat terlihat
- **Formulir Kontak**: Form dengan validasi dasar
- **Hover Effects**: Efek hover pada card dan tombol

## Browser Support

- Chrome (Rekomendasi)
- Firefox
- Safari
- Edge
- Opera

## Catatan

- Aplikasi ini menggunakan CDN untuk Font Awesome dan Google Fonts, pastikan terkoneksi internet
- Untuk penggunaan produksi, disarankan untuk meng-host file CSS/JS/Font secara lokal
- Formulir kontak saat ini hanya menampilkan alert, untuk penggunaan nyata perlu backend server

## Dukungan

Jika mengalami masalah atau pertanyaan, silakan hubungi pengembang atau cek dokumentasi teknologi yang digunakan.

---

© 2024 E-Portofolio PPG Prajabatan
