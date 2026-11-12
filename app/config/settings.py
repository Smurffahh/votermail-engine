# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Copyright (c) 2026 Smurffahh (https://github.com/Smurffahh/votermail-engine)

__author__ = "Smurffahh"

import json
import os
from pathlib import Path
import shutil
from typing import Any
from pydantic import BaseModel, Field

BASE_DIR = Path(__file__).resolve().parent.parent.parent
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
    address_mode: str = "Lengkap (Alamat, RT / RW)"
    enable_pdf_verification: bool = True
    theme_mode: str = "Dark"

def load_app_config() -> AppConfig:
    if CONFIG_FILE_PATH.exists():
        try:
            with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)

            excel = data.get("excel_path", "")
            if not excel or not os.path.exists(excel):
                data["excel_path"] = get_default_excel_path()

            word = data.get("word_path", "")
            if not word or not os.path.exists(word):
                data["word_path"] = get_default_word_path()

            out = data.get("output_dir", "")
            if not out or not os.path.exists(out):
                data["output_dir"] = str(OUTPUT_DIR.resolve())

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
            try:
                if item.is_file():
                    item.unlink(missing_ok=True)
                elif item.is_dir():
                    shutil.rmtree(item, ignore_errors=True)
            except Exception:
                pass
