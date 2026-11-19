# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

import re
from typing import Any

__app_name__ = "VoterMail Engine"
__author__ = "Danang Satria"
__github_user__ = "Smurffahh"
__repository__ = "https://github.com/Smurffahh/votermail-engine"
__copyright__ = "Copyright (c) 2026 Danang Satria"
__license__ = "MIT License — Official Civic Automation Tool"

def parse_version_tuple(version_str: str) -> tuple[int, ...]:
    cleaned = re.sub(r"^[^\d]*", "", str(version_str or "").strip())
    parts = [int(p) for p in re.findall(r"\d+", cleaned)]
    return tuple(parts) if parts else (0, 0, 0)

FEATURE_REGISTRY: list[dict[str, Any]] = [
    {
        "version": "1.0.0",
        "date": "2026-11-10",
        "title": "Rilis Awal Engine Mail-Merge",
        "description": "Rilis awal engine Mail-Merge Word COM & Excel DPT per Kadus/RW.",
    },
    {
        "version": "1.1.0",
        "date": "2026-11-11",
        "title": "Pre-Flight Scanner & PDF Verifier",
        "description": "Penambahan Pre-Flight Excel Scanner, auto-transform L/P & Usia, serta PDF Verifier (On/Off).",
    },
    {
        "version": "1.2.0",
        "date": "2026-11-11",
        "title": "UI Slate Dark Mode & Refactor Presisi",
        "description": "Perombakan UI Slate Dark Mode, perbaikan keutuhan elemen vektor Word, & arsitektur folder modular.",
    },
    {
        "version": "1.3.0",
        "date": "2026-11-12",
        "title": "Live Change Watcher & Dataset Simulasi",
        "description": "Penambahan Live File Change Watcher (Diff Analyzer Word & Excel) & Generator Data Dummy Publik Figur.",
    },
    {
        "version": "1.4.0",
        "date": "2026-11-12",
        "title": "Auto-Dependency Installer & Author Credits",
        "description": 'Penambahan First-Run Auto-Dependency Installer, Sistem Credit "Danang Satria" & About Modal, serta Dynamic Feature Versioning.',
    },
    {
        "version": "1.5.0",
        "date": "2026-11-12",
        "title": "Paket Portabel .EXE & GitHub Actions Auto-Release",
        "description": "Penambahan Paket Distribusi Portabel Windows (.EXE) via PyInstaller (--onedir) & Otomasi CI/CD GitHub Actions Auto-Release.",
    },
    {
        "version": "1.6.0",
        "date": "2026-11-13",
        "title": "Modern Minimalist UI & Dynamic Voter Preview",
        "description": "Perombakan antarmuka Modern Minimalis (Linear/Fluent 2 Style), Auto-Scan instan berkas, Tabel Pratinjau Pemilih dinamis, dan tombol interaktif buka PDF per Kadus.",
    },
    {
        "version": "1.7.0",
        "date": "2026-11-13",
        "title": "Hardening Stabilitas COM, Sanitasi Keamanan, & Sistem Desain",
        "description": "Optimalisasi menyeluruh stabilitas MS Word COM, penguatan keamanan sanitasi path & perangkat Windows, penyempurnaan microcopy & desain sistem UI/UX, serta perluasan cakupan pengujian otomatis.",
    },
    {
        "version": "1.7.1",
        "date": "2026-11-13",
        "title": "Dual-Tuple Theme Tokens, Kontras Mode Terang, & Toolbar Anti-Clipping",
        "description": "Perbaikan transisi tema instan berbasis Dual-Tuple Colors, penyempurnaan kontras Mode Terang, perbaikan tata letak tombol Sumber Dokumen anti-clipping, dan modernisasi kontrol pengubah tema.",
    },
    {
        "version": "1.8.0",
        "date": "2026-11-14",
        "title": "Adaptive Word Engine (1-Up/2-Up), Smart Excel Scanner, & Pemetaan Kolom",
        "description": "Mesin cerdas pemindaian header Excel dengan kamus sinonim dan proteksi Zero PII Leak, deteksi otomatis tata letak Word 1-Up & 2-Up serta dukungan tag manual {{...}}, modal pemetaan kolom interaktif, dan estimasi kebutuhan kertas dinamis.",
    },
    {
        "version": "1.8.1",
        "date": "2026-11-17",
        "title": "Fix Build docxcompose & Progres Bar Granular Real-Time",
        "description": "Perbaikan bug file template docxcompose yang hilang pada versi .EXE (Errno 2) dan penyempurnaan progress bar real-time untuk pemrosesan ribuan data.",
    },
    {
        "version": "1.8.2",
        "date": "2026-11-18",
        "title": "Ikon Resmi VoterMail Engine & Identitas Taskbar Windows",
        "description": "Penambahan ikon resmi VoterMail Engine pada jendela utama, seluruh jendela pop-up (Tentang & Atur Kolom), Taskbar Windows, serta file executable (.EXE).",
    },
    {
        "version": "1.8.3",
        "date": "2026-11-18",
        "title": "Opsi Format Alamat Paten Berbasis Nama Sheet & Live Preview",
        "description": "Penambahan opsi format alamat paten 'Dusun [1..N], RT / RW, Desa Karangjengkol' berbasis nama sheet dengan pembaruan pratinjau instan.",
    },
    {
        "version": "1.8.4",
        "date": "2026-11-18",
        "title": "Stopwatch Waktu Berjalan Real-Time pada Progress Bar",
        "description": "Penggantian indikator Estimasi Sisa menjadi penghitung Waktu Berjalan (Live Stopwatch) real-time yang terus berdetak hingga seluruh proses PDF selesai.",
    },
    {
        "version": "1.9.0",
        "date": "2026-11-18",
        "title": "Cetak Ulang Selektif (Filter Nomor Urut, RT, & Nama Pemilih)",
        "description": "Penambahan fitur Cetak Ulang Selektif (Filter Nomor Urut DPT, Filter per RT, dan Pencarian Nama Pemilih) dengan sinkronisasi pratinjau real-time dan penamaan PDF aman.",
    },
    {
        "version": "1.10.0",
        "date": "2026-11-18",
        "title": "Single-Version Release Notes & Auto-Updater dengan Opsi Update Nanti",
        "description": "Penyederhanaan catatan rilis GitHub ke versi terbaru, penambahan auto-update checker saat startup dengan modal notifikasi dan opsi tunda, serta penginstal update otomatis yang menjaga keamanan file pengguna di folder data/.",
        "highlights": [
            "Penyederhanaan catatan GitHub Releases agar hanya menampilkan daftar pembaruan pada versi terbaru.",
            "Penambahan fitur Auto-Update Checker saat aplikasi dibuka dengan pop-up notifikasi pembaruan dan opsi 'Update Sekarang' atau 'Nanti Saja'.",
            "Penambahan mekanisme pengunduh & pemasang pembaruan otomatis (.EXE & Git) yang menjaga keamanan file data pemilih di folder data/.",
        ],
    },
    {
        "version": "1.9.1",
        "date": "2026-11-18",
        "title": "Hardening Keamanan Template Word, Sanitasi Sel Excel, & Modern Notification System",
        "description": "Penguatan keamanan template Word (Jinja2 Tag Guard) & sanitasi sel Excel, penerjemah error ramah pengguna (Zero Stack-Trace Leak), serta penggantian seluruh pop-up bawaan Windows dengan sistem Modern Notification Modal & Ringkasan Selesai Cetak.",
        "highlights": [
            "Penerapan Jinja2 Template Tag Whitelist pada dokumen Word untuk mencegah eksploitasi Server/Client-Side Template Injection.",
            "Sanitasi otomatis error rumus Excel (#REF!, #VALUE!, #N/A) dan pembatasan panjang sel 250 karakter untuk mencegah MS Word membeku.",
            "Penerjemah error terpusat (Zero Stack-Trace Leak) yang menyajikan panduan Apa + Mengapa + Solusi konkret tanpa path internal mentah.",
            "Penggantian 100% tkinter.messagebox dengan ModernNotificationModal (Success, Warning, Error, Confirm) dan InAppToastBanner.",
        ],
    },
    {
        "version": "1.10.1",
        "date": "2026-11-18",
        "title": "Sistem Tema DaisyUI (Garden & Night), FilterStatusBanner, & Monospace Log",
        "description": "Perombakan sistem tema visual dengan skema DaisyUI Garden & Night, penataan ulang formulir Seksi 3 dengan label teratur 135px & tombol Pratinjau 1 Halaman di header, penambahan FilterStatusBanner 4-status, serta formatting log sistem monospace dengan tag warna dinamis.",
        "highlights": [
            "Penerapan palet DaisyUI Garden (Mode Terang) dan DaisyUI Night (Mode Gelap) dengan token radius konsentris 16px, 8px, 8px, 6px.",
            "Penataan ulang kartu '3. Pengaturan & Cetak PDF' dengan perataan vertikal label 135px dan tombol 'Pratinjau 1 Halaman' di header kanan atas.",
            "Penambahan FilterStatusBanner dengan umpan balik visual 4-status (Semua Pemilih, Belum Diisi, Aktif Cocok, Tidak Ada Data Cocok).",
            "Pemformatan log sistem monospace berlebar tetap dengan tag warna kategori dinamis dan tabel pratinjau ASCII 5 pemilih pertama.",
        ],
    },
    {
        "version": "1.9.2",
        "date": "2026-11-18",
        "title": "Penyelarasan Desain Pratinjau Pemilih & Tombol Log Kontekstual",
        "description": "Penyelarasan desain tabel Pratinjau Pemilih agar serasi dengan Log Sistem & Hasil Cek PDF (penambahan sub-header status sampel, baris header berbingkai, struktur baris bergaris tipis dengan badge wilayah KADUS, penyembunyian otomatis scrollbar pada 5 sampel, serta tombol Bersihkan log kontekstual).",
        "highlights": [
            "Penyelarasan bahasa desain VoterPreviewTable dengan tab Log Sistem dan Hasil Cek PDF.",
            "Penambahan sub-header status bar menampilkan ringkasan jumlah sampel terpilih dan mode alamat aktif.",
            "Tabel 6 kolom terstruktur dengan pemisahan Alamat dan badge Wilayah KADUS tersendiri.",
            "Penyembunyian scrollbar otomatis saat sampel berjumlah 5 baris untuk tampilan yang bersih.",
            "Tombol 'Bersihkan' log sekarang kontekstual dan hanya muncul ketika tab Log Sistem aktif.",
        ],
    },
    {
        "version": "1.9.3",
        "date": "2026-11-19",
        "title": "Migrasi Tipografi ke Inter & Space Mono dengan Pemuat Font Offline",
        "description": "Pembaruan tipografi antarmuka menggunakan font Inter pada seluruh elemen UI dan font Space Mono pada Log Sistem serta angka metrik, dilengkapi pemuat font offline otomatis.",
        "highlights": [
            "Penerapan font Inter sebagai font utama seluruh komponen antarmuka antarmuka (UI Sans).",
            "Penerapan font Space Mono untuk log sistem, detail angka audit, dan angka metrik tabular.",
            "Pemuat font offline otomatis (load_bundled_fonts) mendaftarkan file .ttf ke memori privat Windows GDI (FR_PRIVATE).",
            "Penyelarasan ruang horizontal kolom dan kontrol anti-clipping untuk karakter Space Mono yang lebih tegas.",
        ],
    },
    {
        "version": "1.9.4",
        "date": "2026-11-19",
        "title": "Audit Kebersihan Kode, Modularisasi Modal UI, & Efisiensi Aset",
        "description": "Audit kebenaran kode mendalam, pembersihan kode mati & komentar berlebih, modularisasi komponen modal ke app/ui/modals.py, eliminasi file font redundan (-876 KB), dan penutupan resources file handle secara deterministik.",
        "highlights": [
            "Modularisasi seluruh jendela pop-up modal (About, ColumnMapper, UpdateNotification, ModernNotification) ke app/ui/modals.py.",
            "Pembersihan komentar WHAT dan section headers berlebih pada komponen antarmuka sesuai standar comment-rule.",
            "Eliminasi file font variabel yang tidak terpakai sehingga memangkas ukuran aset hingga 876 KB.",
            "Penutupan handle berkas Excel (openpyxl & pandas) secara deterministik untuk mencegah resource leak di Windows.",
            "Pemeliharaan 100% atribusi resmi pengembang Danang Satria pada seluruh metadata, header, dan antarmuka.",
        ],
    },
]

def get_latest_release_entry() -> dict[str, Any]:
    if FEATURE_REGISTRY:
        return FEATURE_REGISTRY[-1]
    return {
        "version": "1.0.0",
        "date": "2026-11-10",
        "title": "Rilis Awal Engine Mail-Merge",
        "description": "Rilis awal engine Mail-Merge Word COM & Excel DPT per Kadus/RW.",
        "highlights": [],
    }

def get_current_version() -> str:
    if FEATURE_REGISTRY:
        return FEATURE_REGISTRY[-1]["version"]
    return "1.0.0"

__version__ = get_current_version()
