# VoterMail Engine v1.5.0

[![Build & Release](https://github.com/Smurffahh/votermail-engine/actions/workflows/ci_release.yml/badge.svg)](https://github.com/Smurffahh/votermail-engine/actions/workflows/ci_release.yml)
[![Latest Release](https://img.shields.io/github/v/release/Smurffahh/votermail-engine?color=blue&label=Latest%20Release)](https://github.com/Smurffahh/votermail-engine/releases/latest)
[![Version](https://img.shields.io/badge/version-1.5.0-blue.svg)](https://github.com/Smurffahh/votermail-engine)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![Author: Danang Satria](https://img.shields.io/badge/Author-Danang%20Satria-purple.svg)](https://github.com/Smurffahh)
[![Python: 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)

Aplikasi desktop berbasis Python untuk otomatisasi mail-merge ribuan Surat Pemberitahuan Pemilih (Model C6 / Surat Panggilan) dari file Excel DPT ke dalam satu file PDF terpadu per Kadus/RW dengan jaminan presisi tata letak dokumen Microsoft Word asli, audit integritas PDF otomatis, paket portabel Windows (.EXE) siap pakai, serta workflow CI/CD GitHub Actions dengan auto-release.

---

## Unduh Aplikasi Siap Pakai (Windows .EXE Portabel — Tanpa Instal Python)

Untuk panitia di lapangan yang menginginkan aplikasi langsung pakai tanpa perlu menginstal Python atau dependensi:

1. Kunjungi halaman rilis resmi: **[Unduh VoterMail Engine Rilis Terbaru](https://github.com/Smurffahh/votermail-engine/releases/latest)**.
2. Unduh file arsip: `VoterMail_Engine_v1.5.0_Windows_x64_Portable.zip`.
3. Ekstrak (unzip) arsip tersebut ke folder komputer Anda.
4. Pastikan **Microsoft Word** telah terpasang pada Windows Anda.
5. Klik ganda **`VoterMail_Engine.exe`** untuk langsung menjalankan aplikasi dalam 0 detik!
6. Buka berkas `PANDUAN_PANITIA.txt` di dalam folder untuk petunjuk lengkap 4 langkah mudah.

---

## Fitur Utama

- **Paket Distribusi Portabel Windows (.EXE) & Otomasi CI/CD Release**  
  Aplikasi dapat dikompilasi menjadi satu folder portabel (`--onedir`) via PyInstaller. Setiap kali versi baru dinaikkan di `app/config/version.py` dan di-push ke GitHub, workflow GitHub Actions akan otomatis menguji seluruh fungsionalitas dan menerbitkan rilis GitHub baru beserta file arsip `.zip`.

- **First-Run Auto-Dependency Bootstrapper**  
  Jika dijalankan dari source code Python pada komputer baru, sistem otomatis mendeteksi pustaka yang belum terinstal, lalu mengunduh dan memasang paket resmi via `pip` dengan dialog progres visual yang ramah.

- **Pre-Flight Excel DPT Scanner**  
  Memindai seluruh sheet wilayah (Kadus/RW) secara otomatis, mendeteksi header tabel pada baris ke-11, memvalidasi integritas data pemilih tanpa membaca NIK ke memori log, menstandarkan jenis kelamin (`L` $\to$ `Laki-laki`, `P` $\to$ `Perempuan`), dan memformat usia (`21` $\to$ `21 Tahun`).
  
- **Non-Destructive MS Word COM Mail-Merge**  
  Mengotomasi dokumen Microsoft Word asli (`.doc`, `.docx`, `.rtf`) melalui Windows COM Automation (`win32com.client`) dengan format layout 2-Up (2 surat pemilih per lembar F4/A4). Mempertahankan seluruh elemen vector shape, garis pemisah horizontal, dan pita sudut *PANITIA PEMILIHAN* tanpa merusak struktur XML dokumen.

- **Dynamic Schedule Extraction & Template Synchronization**  
  Mengekstrak teks jadwal (`Hari, tanggal`, `Tempat`, `Waktu`) langsung dari Slot 1 secara dinamis dan menyinkronkannya ke Slot 2 jika belum terisi, menghindari hardcode tanggal di kode program.

- **Live File Change Watcher & Structural Diff Analyzer**  
  Mendeteksi otomatis setiap perubahan eksternal pada file Word atau Excel DPT (via hashing SHA-256 dan timestamp file). Menampilkan badge status Amber reaktif, menginvalidasi cache usang, serta mencetak laporan perbandingan perubahan secara mendalam.

- **Toggleable PDF Verification Engine (On/Off) & Official Metadata**  
  Mesin audit pasca-cetak berbasis `pypdf` yang memverifikasi kesesuaian jumlah halaman PDF per Kadus ($\lceil N / 2 \rceil$) serta menguji keterbacaan teks nama pemilih dan nomor urut DPT. Menyisipkan metadata resmi (`/Title`, `/Author: Danang Satria (VoterMail Engine)`, `/Creator`, `/Producer`) ke dalam dokumen PDF yang dihasilkan.

- **Modern Slate Dark/Light Mode Desktop UI & About Modal**  
  Antarmuka desktop responsif berbasis CustomTkinter dengan arsitektur multi-threaded, metrik penghitungan lembar dan pemilih *real-time*, tombol pembatalan proses yang aman, footer bar dengan credit developer, serta pop-up modal "Tentang & Riwayat Fitur".

---

## Struktur Direktori

```text
votermail-engine/
├── .github/
│   └── workflows/
│       └── ci_release.yml       # Otomasi pengujian CI, build .exe, & GitHub Auto-Release
├── Jalankan_VoterMail.bat       # Launcher 1-klik untuk mode source Python
├── Build_EXE_Lokal.bat          # Script 1-klik untuk merakit .EXE portabel secara lokal
├── main.py                      # Bootstrapper dependensi & titik masuk aplikasi
├── requirements.txt             # Daftar dependensi resmi Python (termasuk pyinstaller)
├── README.md                    # Dokumentasi teknis proyek
├── .gitignore                   # Proteksi data sensitif DPT, dist/, build/, & output
├── app/
│   ├── __init__.py
│   ├── config/
│   │   ├── __init__.py
│   │   ├── settings.py          # Konfigurasi path dinamis & adaptasi mode frozen (.exe)
│   │   ├── theme.py             # Design tokens Slate Dark & Light Mode
│   │   └── version.py           # Single Source of Truth versi & feature registry
│   ├── core/
│   │   ├── __init__.py
│   │   ├── security.py          # Sanitasi nama file, validasi path, & pembersihan temp
│   │   ├── word_preparer.py     # Pemeriksa & penyuntik tag dokumen Word non-destruktif
│   │   ├── excel_scanner.py     # Pre-flight scanner & validator data pemilih DPT
│   │   ├── mail_merger.py       # Engine mail-merge Word COM 3-tier & PDF metadata
│   │   └── pdf_verifier.py      # Engine verifikasi & audit integritas file PDF
│   └── ui/
│       ├── __init__.py
│       ├── main_window.py       # Window utama CustomTkinter & footer credit
│       └── components.py        # Komponen UI modular (MetricCard, StatusBadge, AboutModal)
├── data/
│   ├── templates/               # Tempat master dokumen Word & Excel (.gitkeep)
│   ├── output/                  # Direktori output PDF per Kadus (.gitkeep)
│   └── temp/                    # Direktori kerja sementara (.gitkeep)
├── logs/                        # Log audit aplikasi (.gitkeep)
├── scripts/
│   ├── build_exe.py             # Perakitan paket portabel .EXE & kompresi .ZIP
│   ├── bump_version.py          # Utilitas CLI penambahan fitur & SemVer bumper
│   ├── generate_sample_data.py  # Generator dataset simulasi DPT 5 Kadus publik figur
│   └── verify_preview_end_to_end.py # Skrip verifikasi end-to-end pratinjau PDF
└── tests/
    └── test_mailmerge.py        # Test suite komprehensif aman untuk CI (16 unit tests)
```

---

## Persyaratan Sistem

1. **Sistem Operasi**: Windows 10 atau Windows 11 (64-bit).
2. **Microsoft Word**: Microsoft Word (2010 / 2013 / 2016 / 2019 / 2021 / Office 365) terinstal secara lokal untuk menjalankan engine konversi PDF berbasis COM.
3. **Python (Opsional, hanya jika menjalankan dari source)**: Versi 3.10 hingga 3.13.

---

## Panduan Penggunaan Developer

### Cara 1: Menjalankan via Launcher 1-Klik
Klik ganda file `Jalankan_VoterMail.bat`.

### Cara 2: Menjalankan via Terminal / CLI
```powershell
# Jalankan aplikasi langsung
python main.py
```

### Cara 3: Merakit Paket Portabel .EXE Sendiri Secara Lokal
```powershell
# Opsi A: Klik ganda file Build_EXE_Lokal.bat
# Opsi B: Jalankan skrip builder via terminal
python scripts/build_exe.py
```
Hasil paket portabel akan tersedia di `dist/VoterMail_Engine/` dan arsip zip di `dist/VoterMail_Engine_v1.5.0_Windows_x64_Portable.zip`.

### Menjalankan Unit Test (16 Tests — CI Safe)
```powershell
python -m unittest discover -s tests -v
```

### Menambahkan Fitur Baru & Bump Versi
Setiap kali Anda menaikkan versi melalui perintah berikut dan mendorong commit ke branch `main`, GitHub Actions akan otomatis menguji dan menerbitkan paket rilis `.exe` baru ke GitHub Releases:
```powershell
python scripts/bump_version.py --version 1.6.0 --title "Judul Rilis Baru" --desc "Rincian fitur yang ditambahkan..."
git add .
git commit -m "feat(release): bump version to v1.6.0"
git push origin main
```

---

## Keamanan Data & Privasi (Zero PII Leak)

Repository ini dikonfigurasi dengan aturan `.gitignore` ketat:
- Seluruh file spreadsheet DPT asli yang memuat NIK warga dan dokumen dinas di dalam `data/templates/` diblokir dari Git staging.
- Hanya dataset dummy simulasi (`SAMPLE_DPT_DUMMY.xlsx` dan `SAMPLE_SURAT_DUMMY.docx`) yang dilacak ke repositori publik.
- File biner hasil kompilasi lokal (`build/`, `dist/`, `*.spec`) diblokir dari Git staging.
- File PDF hasil cetak pada `data/output/` dan dokumen kerja di `data/temp/` tidak pernah dilacak ke repositori publik.
- Konfigurasi lokal komputer (`config.json`) diblokir agar path direktori absolut lokal pengguna tidak terekspos.

---

## Credits & Lisensi

- **Created & Engineered by**: **Danang Satria** ([@Smurffahh](https://github.com/Smurffahh))
- **GitHub Repository**: [https://github.com/Smurffahh/votermail-engine](https://github.com/Smurffahh/votermail-engine)
- **Lisensi**: [MIT License](LICENSE) — Official Civic Automation Tool.
