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
        "version": "1.10.1",
        "date": "2026-11-18",
        "title": "Hardening Keamanan Word, Sanitasi Excel, Tema DaisyUI, & Notifikasi Modern",
        "description": "Penguatan keamanan template Word (Jinja2 Tag Guard) & sanitasi sel Excel, penerapan sistem tema DaisyUI (Garden & Night) dengan FilterStatusBanner, penerjemah error ramah pengguna (Zero Stack-Trace Leak), serta sistem Modern Notification Modal.",
        "highlights": [
            "Penerapan Jinja2 Template Tag Whitelist pada dokumen Word untuk mencegah eksploitasi Template Injection.",
            "Sanitasi otomatis error rumus Excel (#REF!, #VALUE!, #N/A) dan pembatasan panjang sel 250 karakter.",
            "Penerapan palet tema DaisyUI Garden (Mode Terang) dan DaisyUI Night (Mode Gelap) dengan FilterStatusBanner.",
            "Penerjemah error terpusat (Zero Stack-Trace Leak) dan ModernNotificationModal terpadu (Success, Warning, Error, Confirm).",
        ],
    },
    {
        "version": "1.10.2",
        "date": "2026-11-18",
        "title": "Penyelarasan Desain Pratinjau Pemilih & Tombol Log Kontekstual",
        "description": "Penyelarasan desain tabel Pratinjau Pemilih agar serasi dengan Log Sistem & Hasil Cek PDF (bar status sampel, header berbingkai, badge wilayah KADUS, perbaikan scrollbar, dan tombol kontekstual).",
        "highlights": [
            "Penyelarasan bahasa desain VoterPreviewTable dengan tab Log Sistem dan Hasil Cek PDF.",
            "Penambahan sub-header status bar menampilkan ringkasan jumlah sampel terpilih dan mode alamat aktif.",
            "Tabel 6 kolom terstruktur dengan pemisahan Alamat dan badge Wilayah KADUS tersendiri.",
            "Penyembunyian scrollbar otomatis saat sampel berjumlah 5 baris untuk tampilan yang bersih.",
            "Tombol 'Bersihkan' log sekarang kontekstual dan hanya muncul ketika tab Log Sistem aktif.",
        ],
    },
    {
        "version": "1.10.3",
        "date": "2026-11-19",
        "title": "Migrasi Tipografi ke Inter & Space Mono dengan Pemuat Font Offline",
        "description": "Pembaruan tipografi antarmuka ke font Inter (seluruh UI) dan Space Mono (Log Sistem & angka metrik) lengkap dengan pemuat font offline otomatis.",
        "highlights": [
            "Penerapan font Inter sebagai font utama seluruh komponen antarmuka antarmuka (UI Sans).",
            "Penerapan font Space Mono untuk log sistem, detail angka audit, dan angka metrik tabular.",
            "Pemuat font offline otomatis (load_bundled_fonts) mendaftarkan file .ttf ke memori privat Windows GDI (FR_PRIVATE).",
            "Penyelarasan ruang horizontal kolom dan kontrol anti-clipping untuk karakter Space Mono yang lebih tegas.",
        ],
    },
    {
        "version": "1.10.4",
        "date": "2026-11-19",
        "title": "Audit Kebersihan Kode, Modularisasi Modal UI, & Efisiensi Aset",
        "description": "Audit menyeluruh kebenaran kode, penghapusan kode/aset tidak terpakai, modularisasi app/ui/modals.py, standardisasi komentar minimalis, serta perbaikan ASCII & CRLF pada launcher .bat.",
        "highlights": [
            "Modularisasi seluruh jendela pop-up modal ke app/ui/modals.py dengan re-export backward-compatible.",
            "Pembersihan komentar WHAT dan section headers berlebih pada komponen antarmuka sesuai standar comment-rule.",
            "Eliminasi file font variabel yang tidak terpakai sehingga memangkas ukuran aset hingga 876 KB.",
            "Penutupan handle berkas Excel (openpyxl & pandas) secara deterministik untuk mencegah resource leak di Windows.",
            "Perbaikan format akhiran baris Windows CRLF dan ASCII murni pada script launcher dan auto-updater .bat.",
        ],
    },
    {
        "version": "1.11.0",
        "date": "2026-11-19",
        "title": "Standardisasi SemVer Timeline & Monotonic Version Guardrails",
        "description": "Standardisasi kronologi Semantic Versioning (SemVer) menaik secara ketat hingga v1.11.0, penambahan validator anti-regresi versi (Monotonic Version Guard) di seluruh sistem, dan rilis konsolidasi stabil resmi.",
        "highlights": [
            "Standardisasi Semantic Versioning (SemVer) menaik secara ketat hingga v1.11.0 agar seluruh pengguna versi sebelumnya (v1.9.x maupun v1.10.0) langsung terdeteksi oleh Auto-Updater.",
            "Penambahan mekanisme validasi anti-regresi versi (Monotonic Version Guard) pada registry versi, skrip bump versi, unit test, dan pipeline CI/CD GitHub Actions.",
            "Konsolidasi seluruh pembaruan keamanan, notifikasi modern, tipografi Inter & Space Mono, modularisasi UI, serta perbaikan launcher Windows dalam satu rilis stabil.",
        ],
    },
    {
        "version": "1.11.1",
        "date": "2026-11-22",
        "title": "Perbaikan Duplikasi Teks Horizontal 'PANITIA PEMILIHAN' & Sanitasi OOXML",
        "description": "Mengeliminasi teks duplikat horizontal 'PANITIA PEMILIHAN' pada pratinjau dan hasil cetak PDF melalui eliminasi textbox sintetis Word COM serta sanitasi berkas OOXML terpadu (penomoran ulang ID grafis DrawingML/VML dan normalisasi markup AlternateContent).",
        "highlights": [
            "Eliminasi textbox horizontal sintetis pada proses konversi template .doc ke .docx.",
            "Sanitasi penomoran unik ID grafis (wp:docPr, wps:cNvPr, v:shape) lintas halaman hasil merge docxcompose.",
            "Normalisasi AlternateContent dan preservasi namespace mc:Ignorable (w14 wp14) untuk isolasi Word COM PDF.",
            "Pembersihan otomatis elemen shape hantu tanpa isi/fill/stroke pada pipeline template.",
        ],
    },
    {
        "version": "1.11.2",
        "date": "2026-11-22",
        "title": "Optimasi Tata Letak Dashboard, Penyelarasan Kolom Pratinjau, & Pemformatan Angka Titik Indonesia",
        "description": "Pengurangan tinggi bar metrik wilayah dan ekspansi daftar dusun hingga 5 baris tanpa scroll, penyelarasan presisi kolom VoterPreviewTable, pemformatan angka ribuan titik Indonesia (3.505), serta humanisasi teks operasional.",
        "highlights": [
            "Pengurangan tinggi CleanMetricBar menjadi 58px dan ekspansi daftar Kadus agar seluruh 5 baris langsung terlihat tanpa scroll.",
            "Penyelarasan grid dan eliminasi offset 6px pada VoterPreviewTable sehingga judul kolom lurus sempurna dengan baris data.",
            "Penerapan format angka ribuan dengan pemisah titik (3.505 pemilih) di seluruh antarmuka dashboard dan modal konfirmasi.",
            "Humanisasi teks status cetak dan badge keamanan 'Data Aman • 100% Offline' tanpa kata serba kapital berlebih.",
        ],
    },
    {
        "version": "1.12.0",
        "date": "2026-11-22",
        "title": "Sistem Smart In-Place Patch Updater, Graceful Shutdown, & Notifikasi Melayang",
        "description": "Arsitektur auto-updater modular dengan dukungan paket Patch ringan (~4-5 MB), verifikasi integritas checksum SHA-256 otomatis, graceful shutdown dengan penanganan proses aktif, atomic backup dan auto-rollback, serta floating notification toast card berbasis Windows Forms di pojok kanan bawah layar.",
        "highlights": [
            "Deteksi cerdas perubahan dependensi (requirements.txt hash) untuk memilih mode unduh Patch (~4,2 MB) vs Full Portable (~31,5 MB).",
            "Validasi integritas berkas via streaming SHA-256 dan perlindungan Zip-Slip sebelum ekstraksi.",
            "Graceful shutdown terpadu dengan dialog konfirmasi jika proses pembuatan surat sedang berjalan dan penyimpanan konfigurasi otomatis.",
            "Pemasangan terpisah tanpa jendela konsol hitam (detached PowerShell) dengan UI notifikasi melayang di pojok kanan bawah (Windows Forms Toast Card).",
            "Atomic backup dan rollback otomatis jika proses penimpaan file gagal serta proteksi 100% data pemilih dan template.",
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

def validate_version_registry(
    registry: list[dict[str, Any]] | None = None,
    current_version: str | None = None,
) -> bool:
    reg = registry if registry is not None else FEATURE_REGISTRY
    cur_ver = current_version if current_version is not None else __version__

    if not reg:
        raise ValueError("FEATURE_REGISTRY tidak boleh kosong.")

    semver_pattern = re.compile(r"^\d+\.\d+\.\d+$")
    seen_versions: set[str] = set()
    prev_tuple: tuple[int, ...] | None = None
    prev_version_str: str | None = None

    for i, entry in enumerate(reg):
        ver_str = str(entry.get("version", "")).strip()
        if not semver_pattern.match(ver_str):
            raise ValueError(
                f"Format versi tidak valid pada indeks {i}: '{ver_str}'. Wajib mengikuti pola SemVer mayor.minor.patch (contoh: 1.11.0)."
            )

        if ver_str in seen_versions:
            raise ValueError(f"Ditemukan duplikasi versi SemVer pada registry: '{ver_str}'.")
        seen_versions.add(ver_str)

        curr_tuple = parse_version_tuple(ver_str)
        if prev_tuple is not None and curr_tuple <= prev_tuple:
            raise ValueError(
                f"Regresi versi terdeteksi pada indeks {i}: 'v{ver_str}' tidak lebih besar dari versi sebelumnya 'v{prev_version_str}'. "
                "Urutan FEATURE_REGISTRY wajib naik secara monotonik ketat."
            )

        prev_tuple = curr_tuple
        prev_version_str = ver_str

    last_ver = str(reg[-1].get("version", "")).strip()
    if cur_ver != last_ver:
        raise ValueError(
            f"Versi aktif __version__ ('{cur_ver}') tidak cocok dengan entri terakhir FEATURE_REGISTRY ('{last_ver}')."
        )

    return True

validate_version_registry()

