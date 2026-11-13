# VoterMail Engine — Master Design System

Dokumentasi arsitektur desain visual, token warna, tipografi, dan disiplin antarmuka pengguna untuk **VoterMail Engine**. Standar ini dibangun dengan inspirasi utilitas desktop modern (*Linear + Fluent 2 Desktop Utility*), mengutamakan efisiensi pemrosesan data bervolume tinggi, kejelasan status berkas, dan kenyamanan visual panitia Pilkades.

---

## 1. Filosofi & Prinsip Utama

1. **Single Primary Accent Discipline**
   Hanya tombol tindakan eksekusi utama yang menggunakan aksen warna mencolok (*Emerald Green* `#10B981`). Seluruh tombol kontrol sekunder, dropdown, dan kartu file menggunakan warna netral bersahaja (*neutral slate*) agar mata pengguna terfokus pada status dan aksi berikutnya.
2. **Concentric Border Radius**
   Setiap kontainer bersarang mematuhi hukum konsentris optik:
   - Kontainer terluar / Kartu Utama: `corner_radius = 12px`
   - Kontainer bagian dalam / Baris Item / Input: `corner_radius = 8px`
   - Tombol Pill / Badge status: `corner_radius = 6px`
3. **Structural Hairline Borders**
   Gunakan garis batas struktural tipis `1px` (`#1E293B` pada Dark Mode, `#E2E8F0` pada Light Mode) untuk memisahkan hierarki kartu, bukan bayangan tebal (*heavy drop-shadows*) yang membuat antarmuka kotor.
4. **Disiplin Tipografi & Tabular Numbers**
   - Font Antarmuka Utama: `Segoe UI` (bobot 700 untuk judul, 600 untuk subjudul/kontrol, 400 untuk body teks).
   - Font Numerik & Log Monospace: `Consolas` untuk nomor versi, jumlah pemilih, persentase progres, durasi waktu, dan log audit (`tabular-nums` demi kesejajaran visual).
5. **Anti-AI Slop Discipline**
   - Dilarang menggunakan gradien ungu-biru generik.
   - Dilarang menggunakan teks copywriting klise ("solusi revolusioner", "senantiasa", "menyelami").
   - Selalu sediakan 8-state pada elemen interaktif.

---

## 2. Palet Token Warna & Arsitektur Dual-Tuple (v1.7.1)

Definisi token warna tersentralisasi pada `app/config/theme.py` menggunakan format **Native Dual-Tuple Color Tokens** bawaan CustomTkinter `(LIGHT_HEX, DARK_HEX)`. Format ini memungkinkan `ctk.set_appearance_mode(mode)` mengganti tema secara instan (zero flicker/lag) dan mempertahankan kontras WCAG AA (>= 4.5:1) pada Mode Terang maupun Mode Gelap tanpa perlu manipulasi konfigurasi manual.

### Tabel Token Dual-Tuple (Light, Dark)
| Token | Nilai Light Mode | Nilai Dark Mode | Fungsi / Penggunaan |
| :--- | :--- | :--- | :--- |
| `bg_root` | `#F1F5F9` | `#090D16` | Latar belakang jendela aplikasi utama |
| `bg_card` | `#FFFFFF` | `#111726` | Latar belakang kartu utama (Header, Card 1-4, Footer) |
| `bg_surface` | `#F8FAFC` | `#161F30` | Latar belakang sub-kartu, textbox log, dan input |
| `bg_elevated` | `#F1F5F9` | `#1E293B` | Kotak Progress Bar & Header Tabel Pratinjau |
| `bg_hover` | `#E2E8F0` | `#1E293B` | Highlight saat baris tabel di-hover kursor |
| `border_card` | `#E2E8F0` | `#1E293B` | Hairline border struktural 1px kartu utama |
| `border_surface` | `#CBD5E1` | `#26334D` | Border sub-kartu, tombol sekunder, & input field |
| `text_primary` | `#0F172A` | `#F8FAFC` | Slate 900 pekat (Light) & Putih bersih (Dark) |
| `text_secondary` | `#334155` | `#94A3B8` | Slate 700 (Light) & Slate 400 (Dark) |
| `text_muted` | `#64748B` | `#64748B` | Timestamp, placeholder, dan teks status netral |
| `btn_secondary_bg` | `#EEF2F6` | `#1E293B` | Latar tombol kontrol sekunder (Pilih, Reset, Cek) |
| `btn_secondary_hover` | `#E2E8F0` | `#334155` | Efek hover tombol kontrol sekunder |
| `btn_secondary_border`| `#94A3B8` | `#3B82F6` | Border penegas keterpencetan (*affordance*) |
| `btn_secondary_text`  | `#0F172A` | `#F8FAFC` | Teks tombol kontrol sekunder |
| `btn_accent_soft_bg`  | `#EFF6FF` | `#172554` | Latar tombol "Pilih File" & "Pratinjau PDF" |
| `btn_accent_soft_hover`| `#DBEAFE` | `#1E3A8A` | Efek hover tombol lembut |
| `btn_accent_soft_border`| `#60A5FA` | `#2563EB` | Border aksen biru lembut |
| `btn_accent_soft_text` | `#1D4ED8` | `#93C5FD` | Teks tombol aksen biru |
| `btn_primary_bg` | `#10B981` | `#10B981` | Tombol CTA utama ("Buat File PDF Sekarang") |
| `btn_primary_hover` | `#059669` | `#059669` | Efek hover tombol CTA utama |
| `btn_primary_text` | `#FFFFFF` | `#FFFFFF` | Teks tombol CTA utama |
| `btn_danger_bg` | `#FEE2E2` | `#7F1D1D` | Tombol darurat pembatalan |
| `btn_danger_hover` | `#FECACA` | `#991B1B` | Efek hover tombol darurat |
| `btn_danger_text` | `#B91C1C` | `#FCA5A5` | Teks tombol darurat |
| `btn_disabled_bg` | `#E2E8F0` | `#161F30` | Latar tombol saat dinonaktifkan |
| `btn_disabled_text` | `#94A3B8` | `#475569` | Teks tombol saat dinonaktifkan |
| `status_ready_bg` | `#ECFDF5` | `#064E3B` | Latar badge siap cetak / Zero PII pill |
| `status_ready_text` | `#047857` | `#34D399` | Teks badge siap cetak / Zero PII pill |
| `status_warn_bg` | `#FFFBEB` | `#78350F` | Latar badge peringatan |
| `status_warn_text` | `#B45309` | `#FBBF24` | Teks badge peringatan |
| `status_error_bg` | `#FEF2F2` | `#7F1D1D` | Latar badge error validasi |
| `status_error_text` | `#DC2626` | `#F87171` | Teks badge error validasi |

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
