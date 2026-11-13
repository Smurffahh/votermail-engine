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

## 2. Palet Token Warna

Definisi token warna tersentralisasi pada `app/config/theme.py`:

### A. Linear Dark Mode (Default)
| Token | Nilai Hex | Fungsi / Penggunaan |
| :--- | :--- | :--- |
| `bg_root` | `#090D16` | Latar belakang jendela aplikasi utama |
| `bg_card` | `#111726` | Latar belakang kartu utama (Header, File Control, Status, Log) |
| `bg_surface` | `#161F30` | Latar belakang elemen bersarang (Metric bar, list sheet, input box) |
| `bg_hover` | `#1E293B` | Warna highlight saat kursor berada di atas baris tabel / item |
| `border_subtle` | `#1E293B` | Garis batas tipis 1px pada kartu dan pemisah modul |
| `text_primary` | `#F8FAFC` | Teks utama, judul, dan angka metrik penting |
| `text_secondary` | `#94A3B8` | Teks label pendukung, path file, dan deskripsi |
| `text_muted` | `#64748B` | Timestamp, teks nonaktif, dan placeholder |
| `primary_action` | `#10B981` | Tombol CTA utama ("Buat File PDF Sekarang") |
| `primary_action_hover` | `#059669` | Efek hover tombol CTA utama |
| `danger_action` | `#EF4444` | Tombol darurat ("Batalkan Proses") |
| `danger_action_hover` | `#DC2626` | Efek hover tombol darurat |
| `status_ready` | `#34D399` | Indikator badge siap cetak / validasi sukses |
| `status_warning` | `#FBBF24` | Indikator badge perubahan file / peringatan |
| `status_error` | `#F87171` | Indikator badge error / kegagalan validasi |
| `status_neutral` | `#94A3B8` | Indikator badge menunggu berkas / status netral |

### B. Clean Light Mode
| Token | Nilai Hex | Fungsi / Penggunaan |
| :--- | :--- | :--- |
| `bg_root` | `#F1F5F9` | Latar belakang jendela aplikasi utama |
| `bg_card` | `#FFFFFF` | Latar belakang kartu utama |
| `bg_surface` | `#F8FAFC` | Latar belakang elemen bersarang |
| `bg_hover` | `#E2E8F0` | Highlight hover baris tabel |
| `border_subtle` | `#E2E8F0` | Garis batas tipis 1px |
| `text_primary` | `#0F172A` | Teks utama dan judul |
| `text_secondary` | `#475569` | Teks pendukung |
| `text_muted` | `#64748B` | Timestamp dan teks nonaktif |
| `primary_action` | `#10B981` | Tombol CTA utama |
| `primary_action_hover` | `#059669` | Efek hover tombol CTA utama |
| `status_ready` | `#059669` | Indikator badge sukses |
| `status_warning` | `#D97706` | Indikator badge peringatan |
| `status_error` | `#DC2626` | Indikator badge error |

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
- **Pintasan Keyboard**:
  - `F5`: Memperbarui dan memindai ulang template Word serta berkas Excel secara bersamaan.
  - `Escape`: Menutup modal dialog ("Tentang Aplikasi & Riwayat Versi").
