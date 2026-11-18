# VoterMail Engine — Master Design System

Dokumentasi arsitektur desain visual, token warna, tipografi, dan disiplin antarmuka pengguna untuk **VoterMail Engine**. Standar ini dibangun dengan inspirasi utilitas desktop modern (*Linear + Fluent 2 Desktop Utility*), mengutamakan efisiensi pemrosesan data bervolume tinggi, kejelasan status berkas, dan kenyamanan visual panitia Pilkades.

---

## 1. Filosofi & Prinsip Utama

1. **Single Primary Accent Discipline**
   Hanya tombol tindakan eksekusi utama yang menggunakan aksen warna mencolok (*DaisyUI Garden Rose-Pink* `#FE0075` pada Mode Terang dan *DaisyUI Night Sky-Blue* `#38BDF8` pada Mode Gelap). Seluruh tombol kontrol sekunder, dropdown, dan kartu file menggunakan warna netral bersahaja (*warm neutral umber* `#291E00` pada Light Mode dan *slate neutral* `#1E293B` pada Dark Mode) agar mata pengguna terfokus pada status dan aksi berikutnya.
2. **Concentric Border Radius Architecture (`outer = inner + padding`)**
   Setiap kontainer bersarang mematuhi hukum konsentris optik 4-tingkat:
   - Kontainer Terluar / Kartu Utama / Modal Window: `corner_radius = 16px`
   - Kontainer Dalam / Sub-kartu / Kotak Progress: `corner_radius = 8px`
   - Elemen Kontrol Interaktif (Tombol Aksi, Dropdown OptionMenu, Input Field): `corner_radius = 8px`
   - Elemen Mikro (Badge Status, Indikator Versi, Checkbox): `corner_radius = 6px`
3. **Structural Hairline Borders**
   Gunakan garis batas struktural tipis `1px` (`#2E3C56` pada Dark Mode, `#BFBDBD` pada Light Mode) untuk memisahkan hierarki kartu, bukan bayangan tebal (*heavy drop-shadows*) yang membuat antarmuka kotor.
4. **Disiplin Tipografi & Tabular Numbers**
   - Font Antarmuka Utama: `Segoe UI` (bobot 700 untuk judul seksi, 600 untuk subjudul/kontrol, 400 untuk body teks). Seluruh subjudul dan header tabel menggunakan *Sentence case* alami (tanpa ALL CAPS yang kaku).
   - Font Numerik & Log Monospace: `Consolas` untuk nomor versi, jumlah pemilih, persentase progres, durasi waktu, dan log audit (`tabular-nums` demi mencegah pergeseran/jitter saat angka berubah dinamis).
5. **Anti-AI Slop & UX Copy Discipline**
   - Dilarang menggunakan gradien ungu-biru generik.
   - Dilarang menggunakan teks copywriting klise ("solusi revolusioner", "senantiasa", "menyelami").
   - Seluruh teks aksi menggunakan kata kerja spesifik (contoh: "Mulai Buat Surat PDF", "Pilih Template Word", "Pilih Berkas DPT").
   - Pesan kesalahan wajib memuat 3 elemen: Apa yang terjadi + Penyebab masalah + Langkah konkret perbaikan.
   - Selalu sediakan 8-state pada elemen interaktif.

---

## 2. Palet Token Warna & Arsitektur Dual-Tuple (DaisyUI Garden & Night)

Definisi token warna tersentralisasi pada `app/config/theme.py` menggunakan format **Native Dual-Tuple Color Tokens** bawaan CustomTkinter `(LIGHT_HEX, DARK_HEX)` yang dipetakan langsung dari palet DaisyUI "garden" (Mode Terang) dan DaisyUI "night" (Mode Gelap). Format ini memungkinkan `ctk.set_appearance_mode(mode)` mengganti tema secara instan (zero flicker/lag) dan mempertahankan kontras WCAG AA (>= 4.5:1) pada kedua mode tanpa perlu manipulasi konfigurasi manual.

### Tabel Token Dual-Tuple (Garden, Night)
| Token | Nilai Light (Garden) | Nilai Dark (Night) | Fungsi / Penggunaan |
| :--- | :--- | :--- | :--- |
| `bg_root` | `#D4D2D2` | `#090E1A` | Latar belakang jendela aplikasi utama (Base-200) |
| `bg_card` | `#E9E7E7` | `#0F172A` | Latar belakang kartu utama (Base-100) |
| `bg_surface` | `#F5F4F4` | `#1E293B` | Latar belakang sub-kartu, textbox log, & input |
| `bg_elevated` | `#DFDDDD` | `#162033` | Kotak Progress Bar & Header Tabel Pratinjau |
| `bg_hover` | `#D0CECE` | `#26354D` | Highlight saat baris tabel di-hover kursor |
| `border_card` | `#BFBDBD` | `#2E3C56` | Hairline border struktural 1px kartu utama |
| `border_surface` | `#B2AFAF` | `#334460` | Border sub-kartu, tombol sekunder, & input field |
| `text_primary` | `#100F0F` | `#C8CAD0` | Hitam pekat (Garden) & Muted soft white (Night) |
| `text_secondary` | `#3D3A3A` | `#94A3B8` | Slate secondary (Garden) & Slate 400 (Night) |
| `text_muted` | `#6E6A6A` | `#64748B` | Timestamp, placeholder, dan teks status netral |
| `btn_secondary_bg` | `#DCDADA` | `#1E293B` | Latar tombol kontrol sekunder (Pilih, Cek, Reset) |
| `btn_secondary_hover` | `#CECBCB` | `#334155` | Efek hover tombol kontrol sekunder |
| `btn_secondary_border`| `#B2AFAF` | `#475569` | Border penegas keterpencetan (*affordance*) |
| `btn_secondary_text`  | `#100F0F` | `#E2E8F0` | Teks tombol kontrol sekunder |
| `btn_accent_soft_bg`  | `#E8D6DD` | `#1E1F3B` | Latar tombol "Pratinjau 1 Halaman" & "Pilih File" |
| `btn_accent_soft_hover`| `#D9C0CB` | `#2B2D54` | Efek hover tombol aksen lembut |
| `btn_accent_soft_border`| `#8E4162` | `#818CF8` | Border aksen berry (Garden) / indigo (Night) |
| `btn_accent_soft_text` | `#8E4162` | `#818CF8` | Teks tombol aksen lembut |
| `btn_neutral_bg` | `#291E00` | `#1E293B` | Tombol utilitas netral ("Bersihkan", "Ubah Folder") |
| `btn_neutral_hover` | `#3D2D00` | `#334155` | Efek hover tombol netral |
| `btn_neutral_border` | `#1A1300` | `#334155` | Border tombol netral |
| `btn_neutral_text` | `#E9E7E7` | `#CDD0D6` | Teks tombol netral |
| `btn_primary_bg` | `#FE0075` | `#38BDF8` | Tombol CTA utama ("Mulai Buat Surat PDF") |
| `btn_primary_hover` | `#D90064` | `#0EA5E9` | Efek hover tombol CTA utama |
| `btn_primary_text` | `#FFFFFF` | `#001019` | Teks tombol CTA utama |
| `btn_danger_bg` | `#FFE4E6` | `#311018` | Tombol darurat pembatalan |
| `btn_danger_hover` | `#FECDD3` | `#4A1824` | Efek hover tombol darurat |
| `btn_danger_text` | `#C81E2B` | `#FB7185` | Teks tombol darurat |
| `btn_disabled_bg` | `#DCDADA` | `#162033` | Latar tombol saat dinonaktifkan |
| `btn_disabled_text` | `#8F8B8B` | `#4B5563` | Teks tombol saat dinonaktifkan |
| `status_ready_bg` | `#D1FAE5` | `#0D2E2B` | Latar badge siap cetak / Zero PII pill |
| `status_ready_text` | `#007A4F` | `#2DD4BF` | Teks badge siap cetak / Zero PII pill |
| `status_warn_bg` | `#FEF3C7` | `#2E2410` | Latar badge peringatan |
| `status_warn_text` | `#926B00` | `#F4BF50` | Teks badge peringatan |
| `status_error_bg` | `#FFE4E6` | `#311018` | Latar badge error validasi |
| `status_error_text` | `#C81E2B` | `#FB7185` | Teks badge error validasi |
| `status_info_bg` | `#E0F2FE` | `#0C2538` | Latar badge informasi filter & sistem |
| `status_info_text` | `#0077AD` | `#38BDF8` | Teks badge informasi filter & sistem |

---

## 3. Disiplin 8-State Komponen Interaktif

Seluruh komponen interaktif (tombol, baris Kadus, checkbox, tab log) menerapkan status visual:

1. **Default**: Tampilan tenang dengan garis batas 1px, latar belakang netral sesuai level permukaan.
2. **Hover**: Peningkatan kecerahan permukaan (`bg_surface` -> `bg_hover`) dalam 120ms.
3. **Active / Pressed**: Respon visual instan dengan penekanan warna aksen.
4. **Focus-Visible**: Garis batas jelas tanpa outline browser default yang pecah.
5. **Disabled**: Opasitas teks diturunkan (`text_muted`), warna tombol menjadi redup, kursor tidak mengaktifkan aksi.
6. **Loading / Processing**: Tombol aksi utama berubah label menjadi status dinamis ("Memproses...") disertai progress bar aktif.
7. **Error**: Badge dan label status menampilkan latar belakang khusus kesalahan dengan teks merah ramah pengguna (`status_error`).
8. **Success**: Badge hijau segar (`status_ready`) disertai ringkasan hasil audit validasi.

---

## 4. Pola Microcopy (UX Writing)

Pesan kesalahan dan peringatan wajib mengikuti pola 3 bagian:
`[Apa yang terjadi] + [Penyebab masalah] + [Langkah konkret penyelesaian]`

Contoh Penerapan:
- **File Terkunci**: *"Berkas PDF KADUS 1 sedang dibuka di aplikasi lain. Tutup jendela pembaca PDF tersebut, lalu klik 'Buat File PDF Sekarang' kembali."*
- **File Word Hilang**: *"Berkas template Word tidak ditemukan pada direktori yang dipilih. Periksa kembali lokasi berkas atau pilih ulang melalui tombol 'Pilih File Template'."*
- **Ukuran Melebihi Batas**: *"Ukuran berkas Excel terlalu besar (162 MB, batas maksimal 150 MB). Periksa kembali berkas yang Anda pilih."*

---

## 5. Tata Letak & Skalabilitas

- **Resolusi Minimum**: `1180x820` piksel.
- **Resolusi Bawaan**: `1340x940` piksel, adaptif untuk monitor resolusi HD, Full HD (1080p), hingga 4K.
- **Arsitektur 2-Row Kartu Sumber Dokumen (Anti-Clipping)**:
  - **Baris Atas (Informasi & Status Berkas)**: Menampilkan nama dokumen di sisi kiri dan `StatusBadge` di sisi kanan, diikuti path berkas berlebar 100% dengan `text_secondary` dan truncating aman.
  - **Baris Bawah (Dedicated Action Toolbar)**: Bar horizontal khusus tombol aksi dengan lebar proporsional (`115px` untuk tombol utama "Pilih File...", `95px-100px` untuk "Cek/Scan Ulang", `110px` untuk "Pratinjau PDF"). Struktur ini menjamin label tombol tidak akan pernah terpotong (*zero clipping*) berapa pun panjang nama berkas atau path folder.
- **Pintasan Keyboard**:
  - `F5`: Memperbarui dan memindai ulang template Word serta berkas Excel secara bersamaan.
  - `Escape`: Menutup modal dialog ("Tentang Aplikasi & Riwayat Versi").

---

## 6. Arsitektur Seksi 3 & Indikator Filter Cetak (FilterStatusBanner)

### Perataan Vertikal Label (Strict Vertical Alignment)
Untuk menciptakan ritme visual yang kokoh dan mudah dipindai, seluruh baris kontrol formulir pada kartu **3. Pengaturan & Cetak PDF** menggunakan perataan kiri label dengan lebar tetap `135px` (`anchor="w"`):
- **Baris 1 (Direktori Output)**: Label `135px` -> Entry Path -> Tombol `Ubah Folder` (`105px`) -> Tombol `Buka Folder` (`105px`).
- **Baris 2 (Format Alamat & Verifikasi)**: Label `135px` -> Dropdown Mode Alamat (`280px`) -> Switch `Periksa Keakuratan PDF Otomatis` (rata kanan).
- **Baris 3 (Kontrol Filter Cetak)**: Label `135px` -> Dropdown Mode Filter (`230px`) -> Input Query Filter (expand fleksibel) -> Tombol Reset `✕` (`32x32px`).
- **Header Seksi 3**: Tombol sekunder `Pratinjau 1 Halaman` diletakkan di sisi kanan atas berdampingan dengan `badge_mailing`.

### Disiplin 4-Status `FilterStatusBanner`
Tepat di bawah baris kontrol filter, `FilterStatusBanner` menyajikan umpan balik seketika (*instant zero-latency feedback*) kepada panitia:
1. **Mode Semua Pemilih (`all`)**: Warna Info (`status_info_bg`/`status_info_text`), pesan: `STATUS CETAK: Semua Pemilih • Total {total} Pemilih ({kadus_count} Kadus)`. Input query terkunci dengan placeholder informatif dan tombol reset dinonaktifkan.
2. **Mode Spesifik Belum Diisi (`empty`)**: Warna Warning (`status_warn_bg`/`status_warn_text`), pesan: `FILTER BELUM DIISI: Ketik kriteria filter pada kolom di atas untuk menyaring pemilih`.
3. **Filter Aktif & Ada Pemilih Cocok (`active`)**: Warna Ready (`status_ready_bg`/`status_ready_text`), pesan: `FILTER DITERAPKAN: [{mode}: {query}] • Menyaring {count} dari {total} Pemilih ({kadus_count} Kadus)`. Tombol eksekusi utama mengadopsi label dinamis: `Cetak PDF Terfilter ({count} Pemilih)`.
4. **Filter Aktif Tanpa Data Cocok (`no_match`)**: Warna Error (`status_error_bg`/`status_error_text`), pesan: `TIDAK ADA DATA COCOK: 0 pemilih ditemukan untuk filter tersebut. Periksa kembali ketikan Anda`. Tombol eksekusi utama dinonaktifkan dengan teks `Tidak Ada Pemilih Cocok`.

---

## 7. Arsitektur Monospace Log Sistem & Syntax Highlighting

Sistem log pada tab **Log Sistem** dan **Hasil Cek PDF** menggunakan tata letak monospace berlebar tetap (`Consolas` 11pt) dengan format terstandarisasi:
```text
HH:MM:SS │ KATEGORI   │ Pesan aktivitas sistem
```
- **Tag Warna Dinamis**: Warna kategori dipetakan via `CTkTextbox._textbox.tag_config` (`ts` muted, `div` subtle border, `cat_info` biru, `cat_success` hijau, `cat_warn` kuning, `cat_error` merah).
- **Tabel Pratinjau 5 Pemilih Monospace**: Menggunakan karakter box-drawing Unicode (`┌─┬─┐`, `│`, `├─┼─┤`, `└─┴─┘`) dengan pemotongan teks rapi (`...`) jika melebihi lebar kolom (`No`: 4, `Nama`: 24, `L/P`: 4, `Usia`: 5, `Alamat`: 24).
- **Penyederhanaan Path**: Path absolut yang panjang diringkas menjadi path relatif dari workspace atau prefix `...\subfolder\file.ext` agar tidak merusak keterbacaan baris log.
- **Dukungan Alih Tema**: Seluruh tag warna diperbarui secara otomatis saat pengguna beralih antara Mode Terang (Garden) dan Mode Gelap (Night).
