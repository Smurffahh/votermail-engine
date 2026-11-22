# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import zipfile

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from app.core.updater import compute_deps_hash, compute_file_sha256
from app.config.version import (
    __app_name__,
    __author__,
    __repository__,
    __version__,
)
BUILD_DIR = BASE_DIR / "build"
DIST_DIR = BASE_DIR / "dist"
PACKAGE_DIR = DIST_DIR / "VoterMail_Engine"
ASSETS_DIR = BASE_DIR / "app" / "assets"
ICON_ICO_PATH = ASSETS_DIR / "votermail.ico"
ICON_PNG_PATH = ASSETS_DIR / "votermail.png"
BUILD_META_PATH = ASSETS_DIR / "build_meta.json"

def generate_build_meta() -> Path:
    meta_data = {
        "version": __version__,
        "deps_hash": compute_deps_hash(BASE_DIR / "requirements.txt"),
    }
    BUILD_META_PATH.write_text(json.dumps(meta_data, indent=2), encoding="utf-8")
    print(f"  [+] Menghasilkan metadata build: {BUILD_META_PATH.name} (deps_hash: {meta_data['deps_hash'][:8]}...)")
    return BUILD_META_PATH

def ensure_assets() -> None:
    generate_build_meta()
    if not ICON_ICO_PATH.exists() or not ICON_PNG_PATH.exists():
        print("[*] Menghasilkan aset ikon aplikasi via scripts/generate_icon.py...")
        from scripts.generate_icon import generate_votermail_icon
        generate_votermail_icon(ASSETS_DIR)

    fonts_dir = ASSETS_DIR / "fonts"
    if not fonts_dir.exists() or len(list(fonts_dir.glob("*.ttf"))) < 8:
        print("[*] Menyiapkan aset font offline via scripts/setup_fonts.py...")
        from scripts.setup_fonts import setup_bundled_fonts
        setup_bundled_fonts(fonts_dir)

def clean_build_artifacts() -> None:
    print(f"[*] Membersihkan folder artefak lama (build/, dist/)...")
    if BUILD_DIR.exists():
        shutil.rmtree(BUILD_DIR, ignore_errors=True)
    if DIST_DIR.exists():
        shutil.rmtree(DIST_DIR, ignore_errors=True)
    for spec_file in BASE_DIR.glob("*.spec"):
        try:
            spec_file.unlink(missing_ok=True)
        except Exception:
            pass

def run_pyinstaller() -> None:
    ensure_assets()
    print(f"[*] Menjalankan PyInstaller untuk merakit {__app_name__} v{__version__} (--onedir)...")
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--name=VoterMail_Engine",
        "--onedir",
        "--windowed",
        "--noconfirm",
        "--clean",
        "--icon",
        str(ICON_ICO_PATH),
        "--add-data",
        f"{ASSETS_DIR}{os.pathsep}app/assets",
        "--collect-all=customtkinter",
        "--hidden-import=win32timezone",
        "--hidden-import=win32com",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=pywintypes",
        "--hidden-import=openpyxl",
        "--hidden-import=docxtpl",
        "--hidden-import=docxcompose",
        "--collect-data=docxcompose",
        "--hidden-import=pypdf",
        "--hidden-import=pydantic",
        str(BASE_DIR / "main.py"),
    ]
    subprocess.check_call(cmd, cwd=str(BASE_DIR))

def prepare_distribution_tree() -> None:
    print("[*] Menyiapkan struktur direktori data dan panduan operasional...")
    exe_file = PACKAGE_DIR / "VoterMail_Engine.exe"
    if not exe_file.exists():
        raise FileNotFoundError(f"File biner utama tidak ditemukan: {exe_file}")

    templates_dir = PACKAGE_DIR / "data" / "templates"
    output_dir = PACKAGE_DIR / "data" / "output"
    temp_dir = PACKAGE_DIR / "data" / "temp"
    logs_dir = PACKAGE_DIR / "logs"

    for d in [templates_dir, output_dir, temp_dir, logs_dir]:
        d.mkdir(parents=True, exist_ok=True)
        gitkeep = d / ".gitkeep"
        gitkeep.touch(exist_ok=True)

    src_templates = BASE_DIR / "data" / "templates"
    allowed_dummy_files = ["SAMPLE_DPT_DUMMY.xlsx", "SAMPLE_SURAT_DUMMY.docx"]
    for fname in allowed_dummy_files:
        src_file = src_templates / fname
        if src_file.exists():
            shutil.copy2(src_file, templates_dir / fname)
            print(f"  [+] Menyalin sampel dummy: {fname}")

    # Verifikasi keamanan anti-kebocoran PII
    for item in templates_dir.iterdir():
        lower = item.name.lower()
        if "karangjengkol" in lower or "surat panggilan.doc" in lower:
            item.unlink(missing_ok=True)
            raise RuntimeError(f"Pelanggaran keamanan: File riil {item.name} terdeteksi di dist! Proses dibatalkan.")

    dist_assets = PACKAGE_DIR / "app" / "assets"
    dist_assets.mkdir(parents=True, exist_ok=True)
    if ICON_ICO_PATH.exists():
        shutil.copy2(ICON_ICO_PATH, dist_assets / "votermail.ico")
        print(f"  [+] Menyalin cadangan ikon aplikasi: {dist_assets / 'votermail.ico'}")
    if ICON_PNG_PATH.exists():
        shutil.copy2(ICON_PNG_PATH, dist_assets / "votermail.png")
        print(f"  [+] Menyalin cadangan PNG logo aplikasi: {dist_assets / 'votermail.png'}")
    if BUILD_META_PATH.exists():
        shutil.copy2(BUILD_META_PATH, dist_assets / "build_meta.json")
        print(f"  [+] Menyalin metadata build ke dist: {dist_assets / 'build_meta.json'}")

    dist_fonts = dist_assets / "fonts"
    dist_fonts.mkdir(parents=True, exist_ok=True)
    src_fonts = ASSETS_DIR / "fonts"
    if src_fonts.exists():
        for ttf in src_fonts.glob("*.ttf"):
            shutil.copy2(ttf, dist_fonts / ttf.name)
        print(f"  [+] Menyalin cadangan font offline ke dist: {dist_fonts}")

    guide_path = PACKAGE_DIR / "PANDUAN_PANITIA.txt"
    guide_content = f"""================================================================================
{__app_name__.upper()} v{__version__} — PANDUAN PENGGUNAAN CEPAT UNTUK PANITIA
Dikembangkan oleh: {__author__} ({__repository__})
================================================================================

Aplikasi ini bersifat portabel (tanpa instalasi Python).
Cukup pastikan Microsoft Word terinstal pada Windows Anda untuk ekspor PDF.

--------------------------------------------------------------------------------
LANGKAH PENGGUNAAN (4 TAHAP MUDAH):
--------------------------------------------------------------------------------

1. JALANKAN APLIKASI
   Klik ganda pada 'VoterMail_Engine.exe' di dalam folder ini.
   Aplikasi akan langsung membuka jendela utama.

2. PENYIAPAN BERKAS (WORD & EXCEL DPT)
   Anda dapat meletakkan berkas di folder 'data/templates/' atau memilihnya lewat tombol UI:
   a. File Word Master: Berkas .doc atau .docx surat panggilan (format 2 surat per lembar).
   b. File Excel DPT: Berkas .xlsx berisi sheet data per Kadus/RW (kolom No, Nama, L/P, Umur, Alamat).
   Gunakan tombol 'Cek & Siapkan Template' dan 'Scan & Validasi Excel'.

3. PRATINJAU 1 LEMBAR (PDF)
   Klik tombol 'Pratinjau 1 Lembar (PDF)'.
   Aplikasi akan menguji kompilasi 2 nama pertama dari Excel ke format PDF resmi.
   Periksa hasil pratinjau yang terbuka otomatis di pembaca PDF Anda.

4. CETAK MASSAL PER KADUS
   Pilih sheet/Kadus yang ingin dicetak, lalu klik 'Mulai Cetak Massal per Kadus'.
   Hasil PDF per wilayah akan otomatis tersimpan rapi di folder:
   data/output/Surat_Pemberitahuan_[Nama_Kadus].pdf

--------------------------------------------------------------------------------
CATATAN TEKNIS & KEAMANAN:
--------------------------------------------------------------------------------
- Privasi 100% Offline: Seluruh proses dilakukan lokal di PC tanpa koneksi internet.
- Mode Alamat: Dapat disetel 'Lengkap (Alamat, RT / RW)' atau 'Hanya Alamat'.
- Laporan Audit Verifikasi: Otomatis memverifikasi kesesuaian jumlah halaman dan
  nama pemilih antara Excel DPT dan PDF hasil cetak.

Bantuan teknis & pembaruan versi: {__repository__}
================================================================================
"""
    guide_path.write_text(guide_content, encoding="utf-8")
    print(f"  [+] Menulis panduan panitia: {guide_path.name}")

def create_portable_zip() -> Path:
    zip_name = f"VoterMail_Engine_v{__version__}_Windows_x64_Portable.zip"
    zip_path = DIST_DIR / zip_name
    print(f"[*] Mengompresi paket portabel ke arsip: {zip_path.name}...")

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(PACKAGE_DIR):
            for file in files:
                file_full = Path(root) / file
                rel_path = file_full.relative_to(DIST_DIR)
                zf.write(file_full, arcname=str(rel_path))

    if not zip_path.exists() or zip_path.stat().st_size == 0:
        raise RuntimeError(f"Gagal membuat file ZIP distribusi: {zip_path} tidak ditemukan atau 0 byte.")

    size_mb = zip_path.stat().st_size / (1024 * 1024)
    print(f"[✓] Paket portabel berhasil dirakit: {zip_path.name} ({size_mb:.2f} MB)")
    return zip_path

def create_patch_zip(full_zip_path: Path) -> Path:
    patch_name = f"VoterMail_Engine_v{__version__}_Patch.zip"
    patch_path = DIST_DIR / patch_name
    print(f"[*] Mengompresi paket patch ringan ke arsip: {patch_path.name}...")

    patch_files: list[Path] = []

    exe_file = PACKAGE_DIR / "VoterMail_Engine.exe"
    if exe_file.exists():
        patch_files.append(exe_file)

    assets_folder = PACKAGE_DIR / "app" / "assets"
    if assets_folder.exists():
        for root, _, files in os.walk(assets_folder):
            for f in files:
                patch_files.append(Path(root) / f)

    guide_file = PACKAGE_DIR / "PANDUAN_PANITIA.txt"
    if guide_file.exists():
        patch_files.append(guide_file)

    internal_dir = PACKAGE_DIR / "_internal"
    if internal_dir.exists():
        for root, _, files in os.walk(internal_dir):
            for f in files:
                fp = Path(root) / f
                rel_parts = fp.relative_to(internal_dir).parts
                if f == "base_library.zip" or "app" in rel_parts:
                    patch_files.append(fp)

    with zipfile.ZipFile(patch_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_full in patch_files:
            rel_path = file_full.relative_to(DIST_DIR)
            zf.write(file_full, arcname=str(rel_path))

    if not patch_path.exists() or patch_path.stat().st_size == 0:
        raise RuntimeError(f"Gagal membuat file ZIP patch: {patch_path} tidak ditemukan atau 0 byte.")

    size_mb = patch_path.stat().st_size / (1024 * 1024)
    print(f"[✓] Paket patch berhasil dirakit: {patch_path.name} ({size_mb:.2f} MB)")
    return patch_path

def create_update_manifest(full_zip: Path, patch_zip: Path) -> Path:
    manifest_path = DIST_DIR / "update_manifest.json"
    deps_hash = compute_deps_hash(BASE_DIR / "requirements.txt")
    full_sha = compute_file_sha256(full_zip)
    patch_sha = compute_file_sha256(patch_zip)

    manifest_data = {
        "version": __version__,
        "deps_hash": deps_hash,
        "patch": {
            "filename": patch_zip.name,
            "size_bytes": patch_zip.stat().st_size,
            "sha256": patch_sha,
        },
        "full": {
            "filename": full_zip.name,
            "size_bytes": full_zip.stat().st_size,
            "sha256": full_sha,
        },
    }

    manifest_path.write_text(json.dumps(manifest_data, indent=2), encoding="utf-8")
    print(f"[✓] Manifest pembaruan berhasil dibuat: {manifest_path.name}")
    return manifest_path

def main() -> None:
    print(f"\n========================================================")
    print(f" VoterMail Engine v{__version__} — Portable EXE Packaging Tool")
    print(f" Developer: {__author__}")
    print(f"========================================================\n")
    clean_build_artifacts()
    run_pyinstaller()
    prepare_distribution_tree()
    full_zip = create_portable_zip()
    patch_zip = create_patch_zip(full_zip)
    manifest_file = create_update_manifest(full_zip, patch_zip)
    print(f"\n[SELESAI] Seluruh 3 artefak rilis siap didistribusikan:")
    print(f"  1. Paket Penuh: {full_zip.name} ({full_zip.stat().st_size / (1024*1024):.2f} MB)")
    print(f"  2. Paket Patch: {patch_zip.name} ({patch_zip.stat().st_size / (1024*1024):.2f} MB)")
    print(f"  3. Manifest:    {manifest_file.name}\n")

if __name__ == "__main__":
    main()
