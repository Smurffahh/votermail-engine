# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import os
from pathlib import Path
import re
import shutil

from typing import Any

RESERVED_WINDOWS_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}

EXCEL_FORMULA_ERRORS = {
    "#REF!", "#REF",
    "#DIV/0!", "#DIV/0",
    "#VALUE!", "#VALUE",
    "#NAME?", "#NAME",
    "#N/A", "#NA",
    "#NUM!", "#NUM",
    "#NULL!", "#NULL",
}

SAFE_TAG_REGEX = re.compile(r"^[a-zA-Z_][a-zA-Z0-9_]{0,63}$")
BLOCK_TAG_REGEX = re.compile(r"\{%.*?%\}", re.DOTALL)
VAR_TAG_REGEX = re.compile(r"\{\{(.*?)\}\}", re.DOTALL)

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

def validate_safe_template_tags(content_or_path: str | Path) -> tuple[bool, str]:
    text = ""
    if isinstance(content_or_path, (str, Path)):
        p = Path(content_or_path)
        if p.is_file() and p.suffix.lower() == ".docx":
            try:
                import docx
                doc = docx.Document(str(p))
                paras = []
                for child in doc.element.body:
                    if child.tag.endswith("p"):
                        paras.append(child.text or "")
                    elif child.tag.endswith("tbl"):
                        for row in child.xpath(".//w:tr"):
                            for cell in row.xpath(".//w:tc"):
                                for t in cell.xpath(".//w:t"):
                                    paras.append(t.text or "")
                text = " ".join(paras)
            except Exception:
                text = str(content_or_path)
        else:
            text = str(content_or_path)
    else:
        text = str(content_or_path)

    if BLOCK_TAG_REGEX.search(text) or "{%" in text:
        return False, "Ditemukan tag blok eksekusi Jinja2 ({% ... %}) yang tidak diizinkan demi keamanan sistem."

    var_matches = VAR_TAG_REGEX.findall(text)
    for tag_content in var_matches:
        cleaned = tag_content.strip()
        if "__" in cleaned:
            return False, f"Tag template '{{{{{cleaned}}}}}' tidak aman karena mengandung atribut dunder (__)."
        if any(ch in cleaned for ch in (".", "(", ")", "[", "]", "|", ":", ";", "=", "'", '"')):
            return False, f"Tag template '{{{{{cleaned}}}}}' tidak aman atau mengandung karakter terlarang."
        if not SAFE_TAG_REGEX.match(cleaned):
            return False, f"Tag template '{{{{{cleaned}}}}}' tidak valid. Gunakan hanya huruf, angka, dan garis bawah."

    return True, ""

def sanitize_cell_text(val: Any, max_len: int = 250) -> str:
    if val is None:
        return ""
    text = str(val).strip()
    if not text:
        return ""

    if text.upper() in EXCEL_FORMULA_ERRORS:
        return ""

    if text[0] in ("=", "+", "-", "@", "\t", "\r"):
        text = f"'{text}"

    if len(text) > max_len:
        text = text[:max_len - 3] + "..."

    return text

def translate_exception_to_user_message(exc: Exception | str, context_label: str = "") -> dict[str, str]:
    raw_str = str(exc or "")
    cls_name = type(exc).__name__ if isinstance(exc, Exception) else "Exception"
    ctx_prefix = f" pada {context_label}" if context_label else ""
    ctx_name = context_label or "dokumen yang sedang diproses"

    sanitized_tech = re.sub(r"[A-Za-z]:\\[^:\n\r\t]+", "[PATH_INTERNAL]", raw_str)
    technical_detail = f"[{cls_name}] {sanitized_tech}"

    lower_raw = raw_str.lower()
    lower_cls = cls_name.lower()

    if "permissionerror" in lower_cls or "[errno 13]" in lower_raw or "permission denied" in lower_raw or "being used by another process" in lower_raw:
        return {
            "title": "File Sedang Terbuka di Aplikasi Lain",
            "summary": f"Sistem tidak dapat menyimpan atau mengakses file{ctx_prefix} karena file tersebut sedang dibuka di program lain.",
            "solution": "Tutup file PDF, Word, atau Excel terkait di aplikasi pembaca Anda, lalu klik 'Mulai Buat Surat PDF' kembali.",
            "technical_detail": technical_detail,
        }

    if "filenotfounderror" in lower_cls or "[errno 2]" in lower_raw or "no such file" in lower_raw:
        if any(k in lower_raw for k in ("docxcompose", "templates", "_internal", "custom.xml")):
            return {
                "title": "Komponen Template Internal Tidak Lengkap",
                "summary": "Salah satu file pendukung pembuat dokumen tidak ditemukan di folder aplikasi.",
                "solution": "Pastikan seluruh isi folder ZIP aplikasi telah diekstrak secara utuh (jangan menjalankan aplikasi langsung dari dalam file ZIP), lalu coba kembali.",
                "technical_detail": technical_detail,
            }
        return {
            "title": "File atau Folder Tidak Ditemukan",
            "summary": f"File sumber untuk {ctx_name} tidak ditemukan di lokasi yang dipilih.",
            "solution": "Pastikan file belum dipindahkan atau pilih ulang file melalui kartu '1. Sumber Dokumen'.",
            "technical_detail": technical_detail,
        }

    if "com_error" in lower_cls or "pywintypes" in lower_raw or "word.application" in lower_raw:
        return {
            "title": "Gagal Terhubung ke Microsoft Word",
            "summary": f"Microsoft Word tidak merespons atau menolak perintah saat memproses {ctx_name}.",
            "solution": "Pastikan Microsoft Word sudah terpasang dan teraktivasi di komputer ini, tutup jendela dialog Word yang mungkin terbuka di latar belakang, lalu ulangi proses.",
            "technical_detail": technical_detail,
        }

    if "badzipfile" in lower_cls or "not a zip file" in lower_raw or "password" in lower_raw or "corrupt" in lower_raw:
        return {
            "title": "Format File Tidak Dapat Dibaca",
            "summary": f"File {ctx_name} rusak, terkunci password, atau bukan dokumen Word/Excel yang valid.",
            "solution": "Buka file tersebut di Microsoft Office, pastikan tidak dikunci password, simpan ulang (Save As) sebagai .docx atau .xlsx, lalu muat kembali.",
            "technical_detail": technical_detail,
        }

    if "memoryerror" in lower_cls or "out of memory" in lower_raw:
        return {
            "title": "Memori Komputer Penuh",
            "summary": "Memori sistem tidak mencukupi untuk menggabungkan dokumen dalam satu proses sekaligus.",
            "solution": "Tutup aplikasi lain yang berat atau gunakan fitur 'Filter Cetak' untuk mencetak per beberapa RT/nomor urut.",
            "technical_detail": technical_detail,
        }

    return {
        "title": "Proses Belum Berhasil Diselesaikan",
        "summary": f"Terjadi kendala teknis saat memproses {ctx_name}.",
        "solution": "Silakan coba ulangi beberapa saat lagi. Jika masalah berlanjut, periksa tab 'Log Sistem' untuk rincian lebih lanjut.",
        "technical_detail": technical_detail,
    }

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
