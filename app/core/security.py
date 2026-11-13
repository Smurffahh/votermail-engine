# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import os
from pathlib import Path
import re
import shutil

RESERVED_WINDOWS_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}

def to_win_abs_path(path_val: str | Path) -> str:
    return os.path.normpath(os.path.abspath(str(path_val)))

def sanitize_filename(name: str) -> str:
    cleaned = str(name).strip()
    cleaned = re.sub(r"[\x00-\x1f\x7f-\x9f\u2000-\u200f\u2028-\u202f\u2060-\u206f]", "", cleaned)
    cleaned = re.sub(r"[/\\]+", "_", cleaned)

    # Check reserved Windows device names before '.' is replaced with '_'
    raw_base = re.split(r"[\.\s_/\\]", cleaned)[0].upper()
    is_reserved = raw_base in RESERVED_WINDOWS_NAMES

    cleaned = re.sub(r"[^a-zA-Z0-9_\-\s]", "_", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = re.sub(r"_+", "_", cleaned).strip(" _.")
    if not cleaned:
        cleaned = "DOKUMEN_OUTPUT"

    if is_reserved or cleaned.upper() in RESERVED_WINDOWS_NAMES:
        cleaned = f"WILAYAH_{cleaned}"
    cleaned = cleaned.rstrip(". ")
    return cleaned or "DOKUMEN_OUTPUT"

def validate_safe_output_path(base_dir: str | Path, filename: str) -> Path:
    resolved_base = Path(base_dir).resolve()
    target_path = (resolved_base / filename).resolve()
    if not target_path.is_relative_to(resolved_base):
        raise PermissionError(f"Akses path traversal dicegah: {filename}")
    return target_path

def sanitize_cell_text(val: str) -> str:
    if val and val[0] in ("=", "+", "-", "@", "\t", "\r"):
        return f"'{val}"
    return val

def check_file_accessible_for_writing(file_path: str | Path) -> tuple[bool, str]:
    p = Path(file_path)
    if not p.exists():
        return True, ""
    try:
        with open(p, "a+b"):
            pass
        return True, ""
    except (PermissionError, OSError) as exc:
        return False, f"Berkas '{p.name}' sedang dibuka atau dikunci aplikasi lain. Tutup berkas tersebut terlebih dahulu."

def safe_delete_file(file_path: str | Path) -> bool:
    try:
        p = Path(file_path)
        if p.exists() and p.is_file():
            p.unlink()
            return True
        return False
    except Exception:
        return False

def clean_directory(dir_path: str | Path) -> None:
    p = Path(dir_path)
    if p.exists() and p.is_dir():
        for item in p.iterdir():
            try:
                if item.is_file():
                    item.unlink(missing_ok=True)
                elif item.is_dir():
                    shutil.rmtree(item, ignore_errors=True)
            except Exception:
                pass
