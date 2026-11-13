# VoterMail Engine

Aplikasi desktop Windows berbasis Python untuk mengotomatisasi pembuatan ribuan Surat Pemberitahuan Pemilih dari file Excel DPT menjadi 1 file PDF siap cetak per wilayah (Kadus/RW) menggunakan tata letak asli Microsoft Word.

[![Version](https://img.shields.io/badge/version-1.7.1-10b981?style=flat-square)](https://github.com/Smurffahh/votermail-engine/releases/latest)
[![Build & Release](https://img.shields.io/github/actions/workflow/status/Smurffahh/votermail-engine/ci_release.yml?branch=main&style=flat-square)](https://github.com/Smurffahh/votermail-engine/actions/workflows/ci_release.yml)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11%20%E2%80%A2%20MS%20Word%20COM-334155?style=flat-square)
[![License: MIT](https://img.shields.io/badge/license-MIT-64748b?style=flat-square)](LICENSE)

---

## Mengapa alat ini dibuat

- **Cetak per wilayah tanpa ribuan file tercecer.** Seluruh pemilih dalam satu Kadus/RW digabungkan otomatis menjadi 1 file PDF utuh dengan format 2 surat per lembar F4 (Slot Atas dan Slot Bawah).
- **Format data pemilih dirapikan otomatis.** Kode jenis kelamin `L`/`P` pada Excel DPT diubah otomatis menjadi `Laki-laki`/`Perempuan`, angka usia ditambahkan satuan `Tahun`, dan kolom alamat digabungkan dengan `RT/RW` tanpa perlu mengedit file Excel asli.
- **Keutuhan desain Word & verifikasi PDF otomatis.** Mempertahankan elemen grafis/vektor dokumen Word asli (termasuk pita sudut `PANITIA PEMILIHAN` dan potongan Tanda Terima) serta memeriksa ulang jumlah halaman dan kecocokan nama pemilih di setiap PDF yang dihasilkan.

---

## Cara pakai cepat (untuk panitia pemilihan)

Panduan berikut ditujukan bagi operator pemilihan di lapangan tanpa perlu menginstal Python atau dependensi:

1. **Unduh aplikasi**: Buka halaman **[Releases Terbaru](https://github.com/Smurffahh/votermail-engine/releases/latest)**, lalu unduh file `VoterMail_Engine_v1.7.1_Windows_x64_Portable.zip`.
2. **Ekstrak folder**: Ekstrak file `.zip` tersebut ke folder mana saja di komputer Windows Anda. Pastikan Microsoft Word telah terpasang di komputer.
3. **Jalankan aplikasi**: Klik dua kali **`VoterMail_Engine.exe`**.
4. **Proses cetak 3 langkah di dalam aplikasi**:
   - **Sumber dokumen**: Pilih file Word Surat Panggilan dan file Excel DPT. Aplikasi otomatis memeriksa kelengkapan tag Word dan memindai daftar sheet Kadus di Excel.
   - **Pilih wilayah**: Centang Kadus/RW yang ingin dicetak, lalu periksa pratinjau 5 pemilih pertama pada tabel sisi kanan.
   - **Buat PDF**: Klik tombol **Buat File PDF Sekarang**. Hasil akhir PDF per Kadus beserta laporan verifikasinya langsung tersimpan di folder `data/output/`.

---

## Spesifikasi format dokumen

Aplikasi membaca data pemilih dari Excel DPT dan memetakan setiap baris ke tag template Word secara berpasangan:

| Field pada surat Word | Sumber kolom Excel DPT | Transformasi otomatis saat cetak |
| :--- | :--- | :--- |
| `Nomor` & `Nomor DPT` | Kolom A (`Nomor Urut`) | Diisi berurutan untuk Pemilih Atas (`_1`) dan Pemilih Bawah (`_2`). |
| `Nama Pemilih` | Kolom C (`Nama Pemilih`) | Pembersihan spasi berlebih dan validasi nama kosong. |
| `Laki-laki/Perempuan` | Kolom H (`Jenis Kelamin`) | `L` menjadi `Laki-laki`, `P` menjadi `Perempuan`. |
| `Umur` | Kolom F (`Usia (Tahun)`) | Angka `21` diformat menjadi `21 Tahun`. |
| `Alamat` | Kolom I, J, K (`Alamat`, `RT`, `RW`) | Opsi lengkap (`KARANGJENGKOL, RT 001 / RW 001`) atau hanya nama dusun/desa. |

Kolom B (NIK) diabaikan sepenuhnya oleh sistem untuk menjamin privasi data warga dan pemrosesan berjalan 100% secara offline.

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

### Menjalankan pengujian otomatis (18 tests)
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
│   ├── core/            # Engine Word COM, Excel scanner, PDF verifier, & sanitasi keamanan
│   └── ui/              # Window CustomTkinter, komponen modular, & tabel pratinjau
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
