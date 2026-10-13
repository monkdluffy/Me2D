# Canido Doc Manager (SMAP AutoDoc Studio)

A desktop application designed for automated document generation, batch templating, variable injection, image/stamp/signature placement, batch find-and-replace, and comprehensive PDF manipulation.

---

## 🚀 Fitur Utama

- **SMAP Automation & Templating**:
  - Pengisian otomatis variabel template Word (`.docx`) berdasarkan form input (Badan Usaha, Personil SMAP, Logbook Audit, Rekan Bisnis, dll.).
  - Penempatan aset gambar dinamis (Logo, Cap/Stempel, Tanda Tangan).
  - Batch Find & Replace (pencarian dan penggantian kata/kalimat khusus ke seluruh template Word & PDF).
  - Ekspor/Impor data profil (JSON & Excel).
  - Konversi Word ke PDF dengan fidelitas tinggi (dukungan native MS Word COM interop & fallback).

- **PDF Toolkit**:
  - Konversi DOCX ke PDF.
  - Penambahan Watermark teks & gambar pada PDF.
  - Penggabungan (*Merge*) beberapa file PDF menjadi satu dokumen utuh.
  - Pemisahan (*Split*) file PDF berdasarkan rentang halaman.

- **Image Toolkit**:
  - Penghapus latar belakang gambar (*Background Removal*) berbasis AI (rembg).
  - Ekspor gambar transparan beresolusi tinggi (300 DPI).

- **Database History**:
  - Penyimpanan riwayat dokumen yang diproses menggunakan database lokal SQLite.

---

## 🛠️ Persyaratan Sistem & Instalasi

### 1. Prasyarat
- **Python**: Versi 3.10 atau lebih baru.
- **Git**: Untuk manajemen versi dan update repository.
- **Microsoft Word** (opsional, disarankan untuk konversi PDF dengan kualitas maksimal di Windows).

### 2. Menjalankan dari Source Code
```bash
# Clone repository
git clone <URL_REPOSITORY_ANDA>
cd 0001_canido_doc_manager

# Buat virtual environment
python -m venv .venv

# Aktivasi virtual environment (Windows)
.venv\Scripts\activate

# Install dependensi
pip install -r requirements.txt

# Jalankan aplikasi
python app.py
```

### 3. Menjalankan Unit Tests
```bash
python -m unittest test_suite.py
```

---

## 📦 Build Installer (Windows)

Aplikasi dapat dikompilasi menjadi executable mandiri (.exe) menggunakan PyInstaller dan Inno Setup:
```bash
python build.py
```
Installer akan dihasilkan di folder `dist_installer/`.

---

## 🔄 Rencana Auto-Update (Next Phase)

Aplikasi dirancang agar dapat menerapkan mekanisme auto-update bagi pengguna akhir melalui:
1. **GitHub Releases API**: Memeriksa tag versi rilis terbaru secara berkala.
2. **Updater Executable**: Mengunduh paket patch/installer baru dari GitHub Release dan memperbarui aplikasi secara otomatis saat pengguna membuka aplikasi.

---

## 📄 Lisensi
Hak Cipta © 2026 Canido Sertifikasi. Seluruh hak cipta dilindungi undang-undang.
