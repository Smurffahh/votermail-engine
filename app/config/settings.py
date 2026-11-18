# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import json
import os
from pathlib import Path
import shutil
import sys
from typing import Any
from pydantic import BaseModel, Field

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
    BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
    if not (BUNDLE_DIR / "app" / "assets" / "votermail.ico").exists() and (BASE_DIR / "app" / "assets" / "votermail.ico").exists():
        BUNDLE_DIR = BASE_DIR
else:
    BASE_DIR = Path(__file__).resolve().parent.parent.parent
    BUNDLE_DIR = BASE_DIR

ASSETS_DIR = BUNDLE_DIR / "app" / "assets"
APP_ICON_ICO = ASSETS_DIR / "votermail.ico"
APP_ICON_PNG = ASSETS_DIR / "votermail.png"

DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = DATA_DIR / "templates"
OUTPUT_DIR = DATA_DIR / "output"
TEMP_DIR = DATA_DIR / "temp"
LOGS_DIR = BASE_DIR / "logs"
CONFIG_FILE_PATH = BASE_DIR / "config.json"

for d in [DATA_DIR, TEMPLATES_DIR, OUTPUT_DIR, TEMP_DIR, LOGS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

def get_default_word_path() -> str:
    real_doc = TEMPLATES_DIR / "SURAT PANGGILAN.doc"
    if real_doc.exists():
        return str(real_doc.resolve())
    sample_doc = TEMPLATES_DIR / "SAMPLE_SURAT_DUMMY.docx"
    if sample_doc.exists():
        return str(sample_doc.resolve())
    return ""

ADDRESS_FORMAT_FULL = "Lengkap (Alamat, RT / RW)"
ADDRESS_FORMAT_DUSUN_PATENT = "Dusun, RT / RW, Desa Karangjengkol"
ADDRESS_FORMAT_SHORT = "Ringkas (Hanya Alamat / Dusun)"

ADDRESS_FORMAT_OPTIONS = [
    ADDRESS_FORMAT_FULL,
    ADDRESS_FORMAT_DUSUN_PATENT,
    ADDRESS_FORMAT_SHORT,
]
DEFAULT_ADDRESS_FORMAT = ADDRESS_FORMAT_FULL
PATENTED_VILLAGE_NAME = "Desa Karangjengkol"

FILTER_MODE_ALL = "Semua Pemilih (Standar)"
FILTER_MODE_NUMBER = "Nomor Urut Tertentu"
FILTER_MODE_RT = "Khusus RT Tertentu"
FILTER_MODE_NAME = "Cari Nama Pemilih"

FILTER_MODE_OPTIONS = [
    FILTER_MODE_ALL,
    FILTER_MODE_NUMBER,
    FILTER_MODE_RT,
    FILTER_MODE_NAME,
]
DEFAULT_FILTER_MODE = FILTER_MODE_ALL

def get_default_excel_path() -> str:
    real_excel = TEMPLATES_DIR / "DAFTAR PEMILIH TETAP DESA KARANGJENGKOL.xlsx"
    if real_excel.exists():
        return str(real_excel.resolve())
    sample_excel = TEMPLATES_DIR / "SAMPLE_DPT_DUMMY.xlsx"
    if sample_excel.exists():
        return str(sample_excel.resolve())
    return ""

class AppConfig(BaseModel):
    word_path: str = Field(default_factory=get_default_word_path)
    excel_path: str = Field(default_factory=get_default_excel_path)
    output_dir: str = Field(default_factory=lambda: str(OUTPUT_DIR.resolve()))
    address_mode: str = DEFAULT_ADDRESS_FORMAT
    enable_pdf_verification: bool = True
    theme_mode: str = "Dark"

def _is_safe_local_path(path_str: str) -> bool:
    if not path_str or not isinstance(path_str, str):
        return False
    if path_str.startswith(("\\\\", "//")):
        return False
    p = Path(path_str)
    try:
        if p.is_symlink():
            return False
        return p.exists()
    except Exception:
        return False

def load_app_config() -> AppConfig:
    if CONFIG_FILE_PATH.exists():
        try:
            if CONFIG_FILE_PATH.is_symlink():
                raise PermissionError("config.json berupa symlink tidak diizinkan.")
            if CONFIG_FILE_PATH.stat().st_size > 64 * 1024:
                raise ValueError("Ukuran config.json melebihi batas 64 KB.")

            with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)

            excel = data.get("excel_path", "")
            if not _is_safe_local_path(excel):
                data["excel_path"] = get_default_excel_path()

            word = data.get("word_path", "")
            if not _is_safe_local_path(word):
                data["word_path"] = get_default_word_path()

            out = data.get("output_dir", "")
            if not _is_safe_local_path(out):
                data["output_dir"] = str(OUTPUT_DIR.resolve())

            addr = data.get("address_mode", "")
            if addr not in ADDRESS_FORMAT_OPTIONS:
                data["address_mode"] = DEFAULT_ADDRESS_FORMAT

            return AppConfig(**data)
        except Exception:
            pass
    cfg = AppConfig()
    save_app_config(cfg)
    return cfg

def save_app_config(cfg: AppConfig) -> bool:
    try:
        with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg.model_dump(), f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False

def clean_temp_dir() -> None:
    if TEMP_DIR.exists():
        for item in TEMP_DIR.iterdir():
            if item.name == ".gitkeep":
                continue
            try:
                if item.is_file():
                    item.unlink(missing_ok=True)
                elif item.is_dir():
                    shutil.rmtree(item, ignore_errors=True)
            except Exception:
                pass
