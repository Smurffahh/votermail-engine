# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__app_name__ = "VoterMail Engine"
__author__ = "Danang Satria"
__github_user__ = "Smurffahh"
__repository__ = "https://github.com/Smurffahh/votermail-engine"
__copyright__ = "Copyright (c) 2026 Danang Satria"
__license__ = "MIT License — Official Civic Automation Tool"

FEATURE_REGISTRY: list[dict[str, str]] = [
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
]

def get_current_version() -> str:
    if FEATURE_REGISTRY:
        return FEATURE_REGISTRY[-1]["version"]
    return "1.0.0"

__version__ = get_current_version()
