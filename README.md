# VoterMail Engine v1.4.0

[![Version](https://img.shields.io/badge/version-1.4.0-blue.svg)](https://github.com/Smurffahh/votermail-engine)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![Author: Smurffahh](https://img.shields.io/badge/Author-Smurffahh-purple.svg)](https://github.com/Smurffahh)
[![Python: 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)

Aplikasi desktop berbasis Python untuk otomatisasi mail-merge ribuan Surat Pemberitahuan Pemilih (Model C6 / Surat Panggilan) dari file Excel DPT ke dalam satu file PDF terpadu per Kadus/RW dengan jaminan presisi tata letak dokumen Microsoft Word asli, audit integritas PDF otomatis, serta sistem instalasi dependensi otomatis pada eksekusi pertama.

---

## Fitur Utama

- **First-Run Auto-Dependency Bootstrapper**  
  Mendeteksi komponen yang belum terinstal pada komputer baru secara otomatis sebelum aplikasi dimuat. Mengunduh dan memasang paket resmi via `pip` dengan dialog progres visual tanpa perlu menjalankan perintah CLI manual.

- **Pre-Flight Excel DPT Scanner**  
  Memindai seluruh sheet wilayah (Kadus/RW) secara otomatis, mendeteksi header tabel pada baris ke-11, memvalidasi integritas data pemilih tanpa membaca NIK ke memori log, menstandarkan jenis kelamin (`L` $\to$ `Laki-laki`, `P` $\to$ `Perempuan`), dan memformat usia (`21` $\to$ `21 Tahun`).
  
- **Non-Destructive MS Word COM Mail-Merge**  
  Mengotomasi dokumen Microsoft Word asli (`.doc`, `.docx`, `.rtf`) melalui Windows COM Automation (`win32com.client`) dengan format layout 2-Up (2 surat pemilih per lembar F4/A4). Mempertahankan seluruh elemen vector shape, garis pemisah horizontal, dan pita sudut *PANITIA PEMILIHAN* tanpa merusak struktur XML dokumen.

- **Dynamic Schedule Extraction & Template Synchronization**  
  Mengekstrak teks jadwal (`Hari, tanggal`, `Tempat`, `Waktu`) langsung dari Slot 1 secara dinamis dan menyinkronkannya ke Slot 2 jika belum terisi, menghindari hardcode tanggal di kode program.

- **Live File Change Watcher & Structural Diff Analyzer**  
  Mendeteksi otomatis setiap perubahan eksternal pada file Word atau Excel DPT (via hashing SHA-256 dan timestamp file). Menampilkan badge status Amber reaktif, menginvalidasi cache usang, serta mencetak laporan perbandingan perubahan (jadwal acara, selisih pemilih per sheet, penambahan/penghapusan kolom) secara mendalam.

- **Toggleable PDF Verification Engine (On/Off) & Official Metadata**  
  Mesin audit pasca-cetak berbasis `pypdf` yang memverifikasi kesesuaian jumlah halaman PDF per Kadus ($\lceil N / 2 \rceil$) serta menguji keterbacaan teks nama pemilih dan nomor urut DPT. Menyisipkan metadata resmi (`/Title`, `/Author: Smurffahh`, `/Creator`, `/Producer`) ke dalam dokumen PDF yang dihasilkan.

- **Modern Slate Dark/Light Mode Desktop UI & About Modal**  
  Antarmuka desktop responsif berbasis CustomTkinter dengan arsitektur multi-threaded, metrik penghitungan lembar dan pemilih *real-time*, tombol pembatalan proses yang aman, footer bar dengan credit developer, serta pop-up modal "Tentang & Riwayat Fitur".

---

## Struktur Direktori

```text
votermail-engine/
├── Jalankan_VoterMail.bat       # Launcher 1-klik untuk Windows
├── main.py                      # Bootstrapper dependensi & titik masuk aplikasi
├── requirements.txt             # Daftar dependensi resmi Python
├── README.md                    # Dokumentasi teknis proyek
├── .gitignore                   # Proteksi data sensitif DPT, file temp, & output
├── app/
│   ├── __init__.py
│   ├── config/
│   │   ├── __init__.py
│   │   ├── settings.py          # Konfigurasi path dinamis (BASE_DIR)
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
│   ├── bump_version.py          # Utilitas CLI penambahan fitur & SemVer bumper
│   ├── generate_sample_data.py  # Generator dataset simulasi DPT 5 Kadus publik figur
│   └── verify_preview_end_to_end.py # Skrip verifikasi end-to-end pratinjau PDF
└── tests/
    └── test_mailmerge.py        # Test suite komprehensif (15 unit tests)
```

---

## Persyaratan Sistem

1. **Sistem Operasi**: Windows 10 atau Windows 11 (64-bit).
2. **Python**: Versi 3.10 hingga 3.13 (Pastikan dicentang *"Add Python to PATH"* saat instalasi).
3. **Microsoft Office**: Microsoft Word (2010 / 2013 / 2016 / 2019 / 2021 / Office 365) terinstal secara lokal untuk menjalankan engine konversi PDF berbasis COM.

---

## Panduan Penggunaan Cepat

### Cara 1: Menggunakan Launcher 1-Klik (Rekomendasi Pengguna Awam)
1. Buka folder `votermail-engine`.
2. Klik ganda file `Jalankan_VoterMail.bat`.
3. Jika dijalankan pertama kali di komputer baru, aplikasi akan otomatis memasang semua paket yang diperlukan melalui tampilan visual ramah.

### Cara 2: Menjalankan via Terminal / CLI
```powershell
# 1. Masuk ke direktori
cd votermail-engine

# 2. Jalankan aplikasi (auto-installer aktif jika ada paket kurang)
python main.py
```

### Menjalankan Unit Test (15 Tests)
```powershell
python -m unittest tests/test_mailmerge.py
```

### Menambahkan Fitur Baru & Bump Versi
```powershell
python scripts/bump_version.py --version 1.5.0 --title "Judul Fitur Baru" --desc "Rincian implementasi fitur..."
```

---

## Keamanan Data & Privasi (Zero PII Leak)

Repository ini dikonfigurasi dengan aturan `.gitignore` ketat:
- Seluruh file spreadsheet DPT asli yang memuat NIK warga dan dokumen dinas di dalam `data/templates/` diblokir dari Git staging.
- Hanya dataset dummy simulasi (`SAMPLE_DPT_DUMMY.xlsx` dan `SAMPLE_SURAT_DUMMY.docx`) yang dilacak ke repositori publik.
- File PDF hasil cetak pada `data/output/` dan dokumen kerja di `data/temp/` tidak pernah dilacak ke repositori publik.
- Konfigurasi lokal komputer (`config.json`) diblokir agar path direktori absolut lokal pengguna tidak terekspos.

---

## Credits & Lisensi

- **Author & Developer**: **Smurffahh** ([GitHub Profile](https://github.com/Smurffahh))
- **Repository**: [https://github.com/Smurffahh/votermail-engine](https://github.com/Smurffahh/votermail-engine)
- **Lisensi**: [MIT License](LICENSE) — Official Civic Automation Tool.
