# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import io
import os
from pathlib import Path
import sys
import urllib.request
import zipfile

BASE_DIR = Path(__file__).resolve().parent.parent
FONTS_DIR = BASE_DIR / "app" / "assets" / "fonts"

SPACE_MONO_BASE_URL = "https://raw.githubusercontent.com/google/fonts/main/ofl/spacemono"
SPACE_MONO_FILES = [
    "SpaceMono-Regular.ttf",
    "SpaceMono-Bold.ttf",
    "SpaceMono-Italic.ttf",
    "SpaceMono-BoldItalic.ttf",
]

INTER_ZIP_URL = "https://github.com/rsms/inter/releases/download/v4.1/Inter-4.1.zip"
INTER_TARGET_FILES = {
    "extras/ttf/Inter-Regular.ttf": "Inter-Regular.ttf",
    "extras/ttf/Inter-Medium.ttf": "Inter-Medium.ttf",
    "extras/ttf/Inter-SemiBold.ttf": "Inter-SemiBold.ttf",
    "extras/ttf/Inter-Bold.ttf": "Inter-Bold.ttf",
}
INTER_VARIABLE_URL = "https://raw.githubusercontent.com/google/fonts/main/ofl/inter/Inter%5Bopsz%2Cwght%5D.ttf"

def download_file(url: str, dest_path: Path, min_bytes: int = 10240) -> bool:
    if dest_path.exists() and dest_path.stat().st_size >= min_bytes:
        return True
    print(f"  [-] Mengunduh: {dest_path.name}...")
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) VoterMail-FontSetup/1.9.3"}
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = resp.read()
    if len(data) < min_bytes:
        raise ValueError(f"File {dest_path.name} terlalu kecil ({len(data)} bytes), unduhan kemungkinan korup.")
    dest_path.write_bytes(data)
    print(f"  [✓] Berhasil disimpan: {dest_path.name} ({len(data):,} bytes)")
    return True

def setup_space_mono(target_dir: Path) -> None:
    print("[*] Menyiapkan font Space Mono (Regular, Bold, Italic, BoldItalic)...")
    for fname in SPACE_MONO_FILES:
        url = f"{SPACE_MONO_BASE_URL}/{fname}"
        out_file = target_dir / fname
        download_file(url, out_file)

def setup_inter(target_dir: Path) -> None:
    print("[*] Menyiapkan font Inter (Regular, Medium, SemiBold, Bold)...")
    needed = [out_name for out_name in INTER_TARGET_FILES.values() if not (target_dir / out_name).exists() or (target_dir / out_name).stat().st_size < 10240]
    
    if needed:
        print(f"  [-] Mengunduh arsip rilis resmi Inter ({len(needed)} varian dibutuhkan)...")
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) VoterMail-FontSetup/1.9.3"}
        req = urllib.request.Request(INTER_ZIP_URL, headers=headers)
        with urllib.request.urlopen(req, timeout=60) as resp:
            zip_bytes = resp.read()
        
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            for zip_internal, out_name in INTER_TARGET_FILES.items():
                dest_file = target_dir / out_name
                if zip_internal in z.namelist():
                    data = z.read(zip_internal)
                    dest_file.write_bytes(data)
                    print(f"  [✓] Berhasil diekstrak: {out_name} ({len(data):,} bytes)")
    else:
        print("  [✓] Seluruh varian statis Inter sudah tersedia.")

    var_font_path = target_dir / "Inter-VariableFont_opsz,wght.ttf"
    if not var_font_path.exists() or var_font_path.stat().st_size < 10240:
        try:
            download_file(INTER_VARIABLE_URL, var_font_path)
        except Exception as e:
            print(f"  [!] Catatan: Font variabel opsional dilewati ({e})")

def setup_bundled_fonts(target_dir: Path | None = None) -> Path:
    dest = target_dir or FONTS_DIR
    dest.mkdir(parents=True, exist_ok=True)
    print(f"=== Menyiapkan Aset Font Offline VoterMail Engine ===")
    print(f"Target Direktori: {dest}")
    setup_space_mono(dest)
    setup_inter(dest)
    print("=== Persiapan Font Offline Selesai ===")
    return dest

if __name__ == "__main__":
    setup_bundled_fonts()
