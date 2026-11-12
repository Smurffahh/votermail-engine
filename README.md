# VoterMail Engine

Aplikasi desktop berbasis Python untuk otomatisasi mail-merge ribuan Surat Pemberitahuan Pemilih (Model C6 / Surat Panggilan) dari file Excel DPT ke dalam satu file PDF terpadu per Kadus/RW dengan jaminan presisi tata letak dokumen Microsoft Word asli dan audit integritas PDF otomatis.

---

## Fitur Utama

- **Pre-Flight Excel DPT Scanner**  
  Memindai seluruh sheet wilayah (Kadus/RW) secara otomatis, mendeteksi header tabel pada baris ke-11, memvalidasi integritas data pemilih tanpa membaca NIK ke memori log, menstandarkan jenis kelamin (`L` $\to$ `Laki-laki`, `P` $\to$ `Perempuan`), dan memformat usia (`21` $\to$ `21 Tahun`).
  
- **Non-Destructive MS Word COM Mail-Merge**  
  Mengotomasi dokumen Microsoft Word asli (`.doc`, `.docx`, `.rtf`) melalui Windows COM Automation (`win32com.client`) dengan format layout 2-Up (2 surat pemilih per lembar F4/A4). Mempertahankan seluruh elemen vector shape, garis pemisah horizontal, dan pita sudut *PANITIA PEMILIHAN* tanpa merusak struktur XML dokumen.

- **Toggleable PDF Verification Engine (On/Off)**  
  Mesin audit pasca-cetak berbasis `pypdf` yang memverifikasi kesesuaian jumlah halaman PDF per Kadus ($\lceil N / 2 \rceil$) serta menguji keterbacaan teks nama pemilih dan nomor urut DPT. Hasil audit diekspor langsung ke spreadsheet laporan verifikasi.

- **Modern Slate Dark/Light Mode Desktop UI**  
  Antarmuka desktop responsif berbasis CustomTkinter dengan arsitektur multi-threaded (antarmuka tetap responsif selama proses konversi berlangsung), metrik penghitungan lembar dan pemilih *real-time*, serta tombol pembatalan proses yang aman.

---

## Struktur Direktori

```text
votermail-engine/
├── main.py                      # Titik masuk utama aplikasi
├── requirements.txt             # Daftar dependensi Python
├── README.md                    # Dokumentasi teknis proyek
├── .gitignore                   # Proteksi data sensitif DPT, file temp, & output
├── app/
│   ├── __init__.py
│   ├── config/
│   │   ├── __init__.py
│   │   ├── settings.py          # Konfigurasi path dinamis (BASE_DIR)
│   │   └── theme.py             # Design tokens Slate Dark & Light Mode
│   ├── core/
│   │   ├── __init__.py
│   │   ├── security.py          # Sanitasi nama file, validasi path, & pembersihan temp
│   │   ├── word_preparer.py     # Pemeriksa & penyuntik tag dokumen Word non-destruktif
│   │   ├── excel_scanner.py     # Pre-flight scanner & validator data pemilih DPT
│   │   ├── mail_merger.py       # Engine mail-merge Word COM 3-tier fallback
│   │   └── pdf_verifier.py      # Engine verifikasi & audit integritas file PDF
│   └── ui/
│       ├── __init__.py
│       ├── main_window.py       # Window utama CustomTkinter
│       └── components.py        # Komponen UI modular (MetricCard, StatusBadge, KadusRow)
├── data/
│   ├── templates/               # Tempat master dokumen Word & Excel (.gitkeep)
│   ├── output/                  # Direktori output PDF per Kadus (.gitkeep)
│   └── temp/                    # Direktori kerja sementara (.gitkeep)
├── logs/                        # Log audit aplikasi (.gitkeep)
├── scripts/
│   ├── generate_sample_data.py  # Generator dataset simulasi DPT 5 Kadus
│   └── verify_preview_end_to_end.py # Skrip verifikasi end-to-end pratinjau PDF
└── tests/
    └── test_mailmerge.py        # Test suite komprehensif (10 unit tests)
```

---

## Persyaratan Sistem

1. **Sistem Operasi**: Windows 10 atau Windows 11 (64-bit).
2. **Python**: Versi 3.10 hingga 3.13.
3. **Microsoft Office**: Microsoft Word (2010 / 2013 / 2016 / 2019 / 2021 / Office 365) terinstal secara lokal untuk menjalankan engine konversi PDF berbasis COM.

---

## Panduan Instalasi & Penggunaan

### 1. Clone Repository
```powershell
git clone https://github.com/Smurffahh/votermail-engine.git
cd votermail-engine
```

### 2. Buat & Aktifkan Virtual Environment (Opsional)
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

### 3. Instal Dependensi
```powershell
pip install -r requirements.txt
```

### 4. Buat Data DPT Simulasi (Untuk Pengujian)
Jika belum memiliki file DPT resmi, buat data simulasi (3 Kadus @ 20 pemilih) dengan menjalankan:
```powershell
python -c "from scripts.generate_sample_data import create_dpt_karangjengkol_sample; create_dpt_karangjengkol_sample('data/templates/SAMPLE_DPT_DUMMY.xlsx', voters_per_sheet=20)"
```

### 5. Jalankan Aplikasi
```powershell
python main.py
```

### 6. Menjalankan Unit Test
```powershell
python -m unittest tests/test_mailmerge.py
```

---

## Keamanan Data & Privasi (Zero PII Leak)

Repository ini dikonfigurasi dengan aturan `.gitignore` ketat:
- Seluruh file spreadsheet DPT asli yang memuat NIK warga dan dokumen dinas di dalam `data/templates/` diblokir dari Git staging.
- File PDF hasil cetak pada `data/output/` dan dokumen kerja di `data/temp/` tidak pernah dilacak ke repositori publik.
- Konfigurasi lokal komputer (`config.json`) diblokir agar path direktori absolut lokal pengguna tidak terekspos.

---

## Lisensi

Didistribusikan di bawah Lisensi MIT. Bebas digunakan dan dikembangkan untuk keperluan otomasi administrasi pemilihan umum dan pilkades.
