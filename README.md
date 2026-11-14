# VoterMail Engine

Aplikasi desktop Windows berbasis Python untuk mengotomatisasi pembuatan ribuan Surat Pemberitahuan Pemilih dari file Excel DPT menjadi 1 file PDF siap cetak per wilayah (Kadus/RW) menggunakan tata letak asli Microsoft Word.

[![Version](https://img.shields.io/badge/version-1.8.0-10b981?style=flat-square)](https://github.com/Smurffahh/votermail-engine/releases/latest)
[![Build & Release](https://img.shields.io/github/actions/workflow/status/Smurffahh/votermail-engine/ci_release.yml?branch=main&style=flat-square)](https://github.com/Smurffahh/votermail-engine/actions/workflows/ci_release.yml)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11%20%E2%80%A2%20MS%20Word%20COM-334155?style=flat-square)
[![License: MIT](https://img.shields.io/badge/license-MIT-64748b?style=flat-square)](LICENSE)

---

## Mengapa alat ini dibuat

- **Cetak per wilayah tanpa ribuan file tercecer.** Seluruh pemilih dalam satu Kadus/RW digabungkan otomatis menjadi 1 file PDF utuh siap cetak per wilayah. Mendukung format 2 surat per lembar (2-Up) maupun 1 surat per lembar (1-Up).
- **Scanner Excel adaptif & kamus sinonim header.** Mengenali variasi nama kolom secara otomatis (`Nama Pemilih`, `Nama Warga`, `L/P`, `Gender`, `Usia`, `Domisili`). Tersedia tombol **Atur Kolom** untuk menyesuaikan pemetaan secara interaktif jika format tabel tidak standar.
- **Dukungan template fleksibel.** Mendukung dokumen Word standar tanpa tag (injeksi otomatis), format 1 surat per lembar, serta dokumen dengan tag manual `{{...}}`. Kolom tambahan (seperti TPS atau Dusun) otomatis tersedia sebagai variabel mail-merge.
- **Zero PII leak & keamanan privasi mutlak.** Kolom sensitif (NIK, Nomor KK, KTP) otomatis diblokir permanen dari pemetaan, memori aplikasi, dan file keluaran. Pemrosesan berjalan 100% lokal tanpa koneksi internet.
- **Keutuhan desain Word & verifikasi PDF otomatis.** Mempertahankan elemen grafis/vektor dokumen Word asli (pita sudut `PANITIA PEMILIHAN` dan potongan Tanda Terima) serta memeriksa ulang jumlah halaman dan kecocokan data pemilih di setiap PDF yang dihasilkan.

---

## Cara pakai cepat (untuk panitia pemilihan)

Panduan berikut ditujukan bagi operator pemilihan di lapangan tanpa perlu menginstal Python atau dependensi:

1. **Unduh aplikasi**: Buka halaman **[Releases Terbaru](https://github.com/Smurffahh/votermail-engine/releases/latest)**, lalu unduh file `VoterMail_Engine_v1.8.0_Windows_x64_Portable.zip`.
2. **Ekstrak folder**: Ekstrak file `.zip` tersebut ke folder mana saja di komputer Windows Anda. Pastikan Microsoft Word telah terpasang di komputer.
3. **Jalankan aplikasi**: Klik dua kali **`VoterMail_Engine.exe`**.
4. **Proses cetak 3 langkah di dalam aplikasi**:
   - **Sumber dokumen**: Pilih file Word Surat Panggilan dan file Excel DPT. Aplikasi otomatis mendeteksi format tata letak Word (1-Up atau 2-Up) dan memindai kolom data Excel. Jika kolom belum sesuai, klik **Atur Kolom**.
   - **Pilih wilayah**: Centang Kadus/RW yang ingin dicetak, lalu periksa pratinjau pemilih pada tabel sisi kanan.
   - **Buat PDF**: Klik tombol **Buat File PDF Sekarang**. Hasil akhir PDF per Kadus beserta laporan audit verifikasinya langsung tersimpan di folder `data/output/`.

---

## Spesifikasi format dokumen

Aplikasi membaca data pemilih dari Excel DPT dan memetakan setiap baris ke tag template Word secara otomatis:

| Variabel Surat | Sinonim Kolom Excel yang Dikenali | Transformasi Otomatis |
| :--- | :--- | :--- |
| `Nomor` / `nomor_1` | `No`, `Nomor`, `No Urut`, `ID` | Urut otomatis jika kolom tidak tersedia di Excel. |
| `Nama` / `nama_1` | `Nama Pemilih`, `Nama Lengkap`, `Nama Warga`, `Nama` | Validasi nama kosong & perapian spasi berlebih. |
| `JK` / `jk_1` | `Jenis Kelamin`, `L/P`, `Gender`, `Kelamin` | `L` menjadi `Laki-laki`, `P` menjadi `Perempuan`. |
| `Umur` / `umur_1` | `Usia (Tahun)`, `Umur`, `Usia`, `Age` | Angka `21` otomatis diformat menjadi `21 Tahun`. |
| `Alamat` / `alamat_1` | `Alamat`, `Tempat Tinggal`, `Domisili`, `Jalan` | Opsi lengkap dengan `RT/RW` atau hanya alamat pokok. |
| Variabel Kustom | `TPS`, `Tempat Lahir`, `Keterangan`, dll. | Tersedia sebagai tag `{{tps}}`, `{{tempat_lahir}}`, dll. |

Kolom NIK dan Nomor KK diblokir secara permanen oleh sistem keamanan aplikasi demi melindungi data kependudukan warga.

---

## Panduan developer dan arsitektur proyek

<details>
<summary><strong>Panduan menjalankan dari kode sumber (Python)</strong></summary>

### Prasyarat lingkungan
- Sistem operasi Windows 10 atau 11 (64-bit).
- Python versi 3.10 hingga 3.13.
- Microsoft Word (2010 hingga Office 365) terinstal secara lokal untuk engine Word COM.

### Menjalankan aplikasi
- **Cara termudah**: Klik dua kali `Jalankan_VoterMail.bat`. Fitur first-run bootstrapper di `main.py` akan otomatis mendeteksi dan memasang dependensi dari `requirements.txt` jika belum terinstal.
- **Cara via terminal**:
  ```bash
  git clone https://github.com/Smurffahh/votermail-engine.git
  cd votermail-engine
  python scripts/generate_sample_data.py
  python main.py
  ```

### Menjalankan pengujian otomatis (22 tests)
```bash
python -m pytest tests/test_mailmerge.py -v
```

### Merakit paket .EXE portabel secara lokal
```bash
# Opsi A: Klik dua kali file batch
Build_EXE_Lokal.bat

# Opsi B: Jalankan skrip builder via terminal
python scripts/build_exe.py
```
Paket hasil build akan tersimpan di direktori `dist/VoterMail_Engine/` beserta arsip `.zip`.

</details>

<details>
<summary><strong>Struktur direktori dan arsitektur modul</strong></summary>

### Struktur direktori proyek
```text
votermail-engine/
├── app/
│   ├── config/          # Pengaturan path, tema dual-tuple, & registri versi
│   ├── core/            # Engine Word COM, Excel scanner adaptif, PDF verifier, & sanitasi keamanan
│   └── ui/              # Window CustomTkinter, komponen modular, modal pemetaan kolom, & tabel pratinjau
├── data/
│   ├── output/          # Direktori hasil berkas PDF per Kadus
│   ├── templates/       # Berkas sampel dummy Word & Excel
│   └── temp/            # Direktori kerja sementara pemrosesan dokumen
├── design-system/       # Dokumentasi token warna dan arsitektur layout UI
├── scripts/             # Builder .exe, generator data dummy, & bump version
└── tests/               # Test suite otomatis (18 unit tests, CI safe)
```

### Menaikkan versi rilis & auto-release CI/CD
Setiap kali versi dinaikkan menggunakan utilitas CLI berikut dan didorong ke branch `main`, GitHub Actions (`.github/workflows/ci_release.yml`) akan menjalankan seluruh pengujian unit, merakit biner `.exe`, dan menerbitkan rilis GitHub baru secara otomatis:

```bash
python scripts/bump_version.py --version 1.8.0 --title "Judul Rilis Baru" --desc "Rincian fitur yang ditambahkan..."
git add .
git commit -m "feat(release): bump version to v1.8.0"
git push origin main
```

</details>

---

## Kredit dan lisensi

- **Dikembangkan oleh**: **Danang Satria** ([@Smurffahh](https://github.com/Smurffahh))
- **Repositori**: [github.com/Smurffahh/votermail-engine](https://github.com/Smurffahh/votermail-engine)
- **Lisensi**: Dirilis di bawah lisensi [MIT](LICENSE).
