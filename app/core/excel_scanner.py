# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

from dataclasses import dataclass, field
from datetime import datetime
import hashlib
import math
import os
from pathlib import Path
import re
from typing import Any
import pandas as pd
from pydantic import BaseModel, Field

@dataclass
class ExcelSnapshot:
    file_path: str
    sha256: str
    mtime: float
    size: int
    sheets_summary: dict[str, int]
    total_voters: int
    total_sheets: int
    invalid_rows_count: int
    is_valid: bool = True

def get_excel_file_fingerprint(file_path: str) -> tuple[str, float, int]:
    p = Path(file_path)
    if not p.exists():
        return ("", 0.0, 0)
    mtime = p.stat().st_mtime
    size = p.stat().st_size
    hasher = hashlib.sha256()
    with open(p, "rb") as f:
        while chunk := f.read(65536):
            hasher.update(chunk)
    return (hasher.hexdigest(), mtime, size)

def compute_excel_snapshot(scan_res: "PreFlightScanResult", file_path: str) -> ExcelSnapshot:
    sha256, mtime, size = get_excel_file_fingerprint(file_path)
    sheets_summary = {
        s.sheet_name: s.total_voters
        for s in scan_res.sheet_summaries
        if s.is_voter_sheet
    }
    return ExcelSnapshot(
        file_path=file_path,
        sha256=sha256,
        mtime=mtime,
        size=size,
        sheets_summary=sheets_summary,
        total_voters=scan_res.total_voters,
        total_sheets=scan_res.total_sheets,
        invalid_rows_count=len(scan_res.all_issues),
        is_valid=scan_res.is_valid,
    )

def analyze_excel_diff(old_snap: ExcelSnapshot | None, new_snap: ExcelSnapshot) -> list[str]:
    lines: list[str] = []
    file_name = os.path.basename(new_snap.file_path)
    mod_time_str = datetime.fromtimestamp(new_snap.mtime).strftime("%d/%m/%Y %H:%M:%S")
    lines.append(f"[DIFF EXCEL] File Excel terdeteksi diperbarui: {file_name}")
    lines.append(f"[DIFF EXCEL] Waktu Simpan: {mod_time_str} | Ukuran: {new_snap.size:,} bytes")

    if old_snap is None:
        lines.append("[DIFF EXCEL] Snapshot awal data Excel DPT berhasil dibuat.")
        lines.append(f"[DIFF EXCEL] Total Pemilih: {new_snap.total_voters:,} orang di {new_snap.total_sheets} Sheet Kadus")
        for s_name, count in new_snap.sheets_summary.items():
            lines.append(f" - {s_name}: {count:,} Pemilih")
        return lines

    old_sheets = set(old_snap.sheets_summary.keys())
    new_sheets = set(new_snap.sheets_summary.keys())
    added = new_sheets - old_sheets
    removed = old_sheets - new_sheets
    if added:
        lines.append(f"[DIFF EXCEL] Sheet Kadus Baru: {', '.join(sorted(added))}")
    if removed:
        lines.append(f"[DIFF EXCEL] Sheet Kadus Dihapus: {', '.join(sorted(removed))}")

    count_changed = False
    for s_name in sorted(new_sheets.intersection(old_sheets)):
        old_c = old_snap.sheets_summary[s_name]
        new_c = new_snap.sheets_summary[s_name]
        if old_c != new_c:
            count_changed = True
            diff = new_c - old_c
            diff_sign = f"+{diff}" if diff > 0 else str(diff)
            lines.append(f"[DIFF EXCEL] {s_name}: {old_c:,} -> {new_c:,} Pemilih ({diff_sign} pemilih)")

    if old_snap.total_voters != new_snap.total_voters:
        diff_total = new_snap.total_voters - old_snap.total_voters
        diff_total_sign = f"+{diff_total}" if diff_total > 0 else str(diff_total)
        lines.append(f"[DIFF EXCEL] Total Pemilih: {old_snap.total_voters:,} -> {new_snap.total_voters:,} ({diff_total_sign} pemilih)")
    elif not count_changed and not added and not removed:
        lines.append("[DIFF EXCEL] Jumlah pemilih pada tiap Sheet Kadus tetap sama.")

    if old_snap.invalid_rows_count != new_snap.invalid_rows_count:
        lines.append(f"[DIFF EXCEL] Baris Bermasalah: {old_snap.invalid_rows_count} -> {new_snap.invalid_rows_count}")

    status_str = "VALID (SIAP MAILING)" if new_snap.is_valid else "PERLU PERBAIKAN"
    lines.append(f"[DIFF EXCEL] Status Validasi Excel: {status_str}")
    return lines

class VoterData(BaseModel):
    nomor: str
    nama: str
    jenis_kelamin: str
    umur: str
    alamat: str
    sheet_name: str
    wilayah: str
    row_index: int
    extra_fields: dict[str, str] = Field(default_factory=dict)

class ScanIssue(BaseModel):
    sheet_name: str
    row_index: int
    column_name: str
    issue: str

class SheetScanSummary(BaseModel):
    sheet_name: str
    wilayah_text: str
    is_voter_sheet: bool = True
    is_selected: bool = True
    total_voters: int = 0
    expected_pages: int = 0
    issues: list[ScanIssue] = Field(default_factory=list)
    voters: list[VoterData] = Field(default_factory=list)
    header_row_index: int = 0
    schema_type: str = "STANDAR_DPT"
    schema_message: str = ""
    column_mapping: dict[str, int] = Field(default_factory=dict)
    available_columns: list[tuple[int, str]] = Field(default_factory=list)

class PreFlightScanResult(BaseModel):
    is_valid: bool = True
    total_sheets: int = 0
    sheet_summaries: list[SheetScanSummary] = Field(default_factory=list)
    total_voters: int = 0
    total_pages: int = 0
    all_issues: list[ScanIssue] = Field(default_factory=list)
    preview_samples: list[str] = Field(default_factory=list)
    schema_type: str = "STANDAR_DPT"
    schema_message: str = ""
    column_mapping: dict[str, int] = Field(default_factory=dict)
    available_columns: list[tuple[int, str]] = Field(default_factory=list)

def transform_gender(val: Any) -> str:
    cleaned = str(val or "").strip().upper()
    if cleaned in ("L", "LK") or cleaned.startswith("L"):
        return "Laki-laki"
    if cleaned in ("P", "PR") or cleaned.startswith("P"):
        return "Perempuan"
    return str(val or "").strip().title()

def transform_age(val: Any) -> str:
    raw = str(val or "").strip()
    if raw.endswith(".0"):
        raw = raw[:-2]
    if raw.isdigit():
        return f"{raw} Tahun"
    if raw and "tahun" not in raw.lower():
        return f"{raw} Tahun"
    return raw

def transform_address(alamat_val: Any, rt_val: Any, rw_val: Any, mode: str = "lengkap") -> str:
    alamat = str(alamat_val or "").strip()
    if mode != "lengkap":
        return alamat
    rt_s = str(rt_val or "").strip()
    rw_s = str(rw_val or "").strip()
    if rt_s or rw_s:
        rt_fmt = rt_s.zfill(3) if rt_s.isdigit() else rt_s
        rw_fmt = rw_s.zfill(3) if rw_s.isdigit() else rw_s
        return f"{alamat}, RT {rt_fmt} / RW {rw_fmt}".strip(", ")
    return alamat

SYNONYMS: dict[str, list[str]] = {
    "no": ["no", "nomor", "no.", "no urut", "nomor urut", "id", "num"],
    "nama": ["nama", "nama lengkap", "nama pemilih", "nama warga", "warga"],
    "jk": ["jk", "jenis kelamin", "l/p", "gender", "kelamin", "sex"],
    "umur": ["usia", "umur", "age", "thn", "usia (tahun)"],
    "alamat": ["alamat", "tempat tinggal", "domisili", "jalan", "dusun", "dukuh"],
    "rt": ["rt", "r.t", "r.t."],
    "rw": ["rw", "r.w", "r.w."],
}

BLOCKED_PII_PATTERNS = [
    r"\bnik\b",
    r"no\.?\s*kk",
    r"\bnokk\b",
    r"nomor\s+induk\s+kependudukan",
    r"kartu\s+keluarga",
    r"nomor\s+kk",
    r"\bktp\b",
]

def is_blocked_pii(col_name: str) -> bool:
    name_clean = str(col_name or "").strip().lower()
    return any(re.search(pat, name_clean) for pat in BLOCKED_PII_PATTERNS)

def to_snake_case(text: str) -> str:
    s = re.sub(r"[^\w\s]", "", str(text or ""))
    s = re.sub(r"\s+", "_", s.strip().lower())
    return s

def _detect_header_and_mapping(
    df: pd.DataFrame,
    custom_map: dict[str, int] | None = None,
) -> tuple[int | None, dict[str, int], int, list[tuple[int, str]], str, str]:
    best_row_idx: int | None = None
    best_score = -1
    best_matches: dict[str, int] = {}

    max_search_rows = min(25, len(df))
    for r_idx in range(max_search_rows):
        row_vals = [str(df.iloc[r_idx, c]).strip() for c in range(df.shape[1])]
        matches: dict[str, int] = {}
        for c_idx, cell in enumerate(row_vals):
            if not cell or is_blocked_pii(cell):
                continue
            cl = cell.lower().strip()
            for cat, syns in SYNONYMS.items():
                if cat not in matches:
                    for s in syns:
                        if cl == s or cl.startswith(s + " ") or cl.endswith(" " + s) or f"({s})" in cl:
                            matches[cat] = c_idx
                            break
        score = len(matches)
        if "nama" in matches and "no" in matches:
            score += 2
        elif "nama" in matches:
            score += 1

        if score > best_score:
            best_score = score
            best_row_idx = r_idx
            best_matches = matches

    if best_score < 2 or best_row_idx is None:
        for r_idx in range(min(20, len(df))):
            row_vals = [str(df.iloc[r_idx, c]).strip() for c in range(min(12, df.shape[1]))]
            if len(row_vals) >= 3 and row_vals[0] == "1" and row_vals[2] == "3":
                best_row_idx = max(0, r_idx - 1)
                best_matches = {
                    "no": 0, "nama": 2, "umur": 5, "jk": 7, "alamat": 8, "rt": 9, "rw": 10
                }
                break

    if best_row_idx is None:
        return None, {}, 0, [], "PERLU_PEMETAAN_MANUAL", "Struktur header tidak terdeteksi."

    available_columns: list[tuple[int, str]] = []
    for c_idx in range(df.shape[1]):
        val = str(df.iloc[best_row_idx, c_idx]).strip()
        available_columns.append((c_idx, val if val else f"Kolom {c_idx + 1}"))

    col_map = dict(best_matches)
    if custom_map:
        for k, v in custom_map.items():
            if v < df.shape[1]:
                col_name = str(df.iloc[best_row_idx, v]).strip()
                if not is_blocked_pii(col_name):
                    col_map[k] = v

    is_standard = (
        col_map.get("no") == 0
        and col_map.get("nama") == 2
        and col_map.get("umur") == 5
        and col_map.get("jk") == 7
        and col_map.get("alamat") == 8
        and col_map.get("rt") == 9
        and col_map.get("rw") == 10
    )

    if is_standard:
        schema_type = "STANDAR_DPT"
        schema_message = "Format DPT Standar Karangjengkol (Otomatis)"
    elif col_map.get("nama") is not None:
        schema_type = "KUSTOM_TERPETAKAN"
        schema_message = f"Format DPT Kustom Terpetakan (Header Baris {best_row_idx + 1})"
    else:
        schema_type = "PERLU_PEMETAAN_MANUAL"
        schema_message = "Kolom Nama Pemilih belum terpetakan. Silakan gunakan tombol 'Atur Kolom'."

    data_start_row = best_row_idx + 1
    if data_start_row < len(df):
        next_vals = [str(df.iloc[data_start_row, c]).strip() for c in range(min(12, df.shape[1]))]
        numeric_count = sum(1 for c, v in enumerate(next_vals) if v == str(c + 1))
        if numeric_count >= 3 or (len(next_vals) >= 3 and next_vals[0] == "1" and next_vals[2] == "3"):
            data_start_row = best_row_idx + 2

    return best_row_idx, col_map, data_start_row, available_columns, schema_type, schema_message

def _find_wilayah_text(df: pd.DataFrame, default_sheet_name: str) -> str:
    for r_idx in range(min(10, len(df))):
        row_vals = [str(df.iloc[r_idx, c]).strip() for c in range(min(12, df.shape[1]))]
        for c_idx, cell_v in enumerate(row_vals):
            if "WILAYAH" in cell_v.upper():
                next_cell = row_vals[c_idx + 1] if c_idx + 1 < len(row_vals) else ""
                clean_target = next_cell if next_cell else cell_v
                clean_text = clean_target.replace("WILAYAH", "").replace(":", "").strip()
                if clean_text:
                    return clean_text
    return default_sheet_name

def scan_and_validate_excel(
    file_path: str,
    address_mode: str = "lengkap",
    custom_column_map: dict[str, int] | None = None,
) -> PreFlightScanResult:
    if not os.path.exists(file_path):
        return PreFlightScanResult(
            is_valid=False,
            all_issues=[ScanIssue(sheet_name="", row_index=0, column_name="File", issue="File Excel tidak ditemukan.")],
        )

    file_size = os.path.getsize(file_path)
    if file_size > 150 * 1024 * 1024:
        return PreFlightScanResult(
            is_valid=False,
            all_issues=[
                ScanIssue(
                    sheet_name="",
                    row_index=0,
                    column_name="File",
                    issue=f"Ukuran berkas Excel terlalu besar ({file_size / (1024 * 1024):.1f} MB, batas maksimal 150 MB). Periksa kembali berkas yang Anda pilih.",
                )
            ],
        )

    try:
        with open(file_path, "rb") as f:
            f.read(1024)
    except (PermissionError, OSError):
        return PreFlightScanResult(
            is_valid=False,
            all_issues=[
                ScanIssue(
                    sheet_name="",
                    row_index=0,
                    column_name="File",
                    issue=f"Berkas Excel '{Path(file_path).name}' sedang dibuka atau dikunci aplikasi lain. Tutup Microsoft Excel atau aplikasi yang membuka berkas ini, lalu coba lagi.",
                )
            ],
        )

    xl = pd.ExcelFile(file_path)
    summaries: list[SheetScanSummary] = []
    total_all_voters = 0
    total_all_pages = 0
    all_issues: list[ScanIssue] = []
    preview_samples: list[str] = []
    top_schema_type = "STANDAR_DPT"
    top_schema_msg = ""
    top_col_mapping: dict[str, int] = {}
    top_available_cols: list[tuple[int, str]] = []

    for sheet_name in xl.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str, keep_default_na=False)
        wilayah_name = _find_wilayah_text(df, sheet_name)

        if any(k in sheet_name.lower() for k in ("rekap", "summary", "cover", "petunjuk", "instruksi")):
            summaries.append(
                SheetScanSummary(
                    sheet_name=sheet_name,
                    wilayah_text=wilayah_name,
                    is_voter_sheet=False,
                    is_selected=False,
                    header_row_index=0,
                    schema_type="REKAP",
                    schema_message="Sheet rekapitulasi / non-pemilih diabaikan.",
                    column_mapping={},
                    available_columns=[],
                )
            )
            continue

        header_idx, col_map, data_start_row, avail_cols, schema_type, schema_msg = _detect_header_and_mapping(
            df, custom_map=custom_column_map
        )

        if header_idx is None:
            summaries.append(
                SheetScanSummary(
                    sheet_name=sheet_name,
                    wilayah_text=wilayah_name,
                    is_voter_sheet=False,
                    is_selected=False,
                    header_row_index=0,
                    schema_type="PERLU_PEMETAAN_MANUAL",
                    schema_message="Header tabel pemilih tidak terdeteksi.",
                    column_mapping={},
                    available_columns=[],
                )
            )
            continue

        if not top_col_mapping:
            top_schema_type = schema_type
            top_schema_msg = schema_msg
            top_col_mapping = col_map
            top_available_cols = avail_cols

        # Map non-PII and non-core columns to extra_fields
        extra_fields_cols: dict[str, int] = {}
        mapped_indices = set(col_map.values())
        for c in range(df.shape[1]):
            if c not in mapped_indices:
                c_name = str(df.iloc[header_idx, c]).strip()
                if c_name and not is_blocked_pii(c_name):
                    snake_k = to_snake_case(c_name)
                    if snake_k and not is_blocked_pii(snake_k):
                        extra_fields_cols[snake_k] = c

        voters: list[VoterData] = []
        sheet_issues: list[ScanIssue] = []
        recap_keywords = ("jumlah pemilih", "status perkawinan", "laki-laki", "perempuan", "total", "b =", "s =", "p =")

        no_col = col_map.get("no")
        nama_col = col_map.get("nama")
        jk_col = col_map.get("jk")
        umur_col = col_map.get("umur")
        alamat_col = col_map.get("alamat")
        rt_col = col_map.get("rt")
        rw_col = col_map.get("rw")

        for r_idx in range(data_start_row, len(df)):
            row = [str(df.iloc[r_idx, c]).strip() for c in range(df.shape[1])]
            no_val = row[no_col] if no_col is not None and no_col < len(row) else ""
            nama_val = row[nama_col] if nama_col is not None and nama_col < len(row) else ""

            first_cols_combined = " ".join(row[:min(5, len(row))]).lower()
            if any(k in first_cols_combined for k in recap_keywords):
                continue
            if not no_val and not nama_val:
                continue

            excel_row_num = r_idx + 1

            if no_col is not None:
                if no_val.isdigit():
                    final_no = str(int(no_val))
                elif nama_val and not any(k in first_cols_combined for k in recap_keywords):
                    final_no = no_val if no_val else str(len(voters) + 1)
                else:
                    continue
            else:
                if not nama_val:
                    continue
                final_no = str(len(voters) + 1)

            if not nama_val:
                sheet_issues.append(
                    ScanIssue(
                        sheet_name=sheet_name,
                        row_index=excel_row_num,
                        column_name="Nama Pemilih",
                        issue="Nama pemilih kosong.",
                    )
                )

            raw_jk = row[jk_col] if jk_col is not None and jk_col < len(row) else ""
            jk_val = transform_gender(raw_jk)

            raw_umur = row[umur_col] if umur_col is not None and umur_col < len(row) else ""
            umur_val = transform_age(raw_umur)

            raw_alamat = row[alamat_col] if alamat_col is not None and alamat_col < len(row) else ""
            raw_rt = row[rt_col] if rt_col is not None and rt_col < len(row) else ""
            raw_rw = row[rw_col] if rw_col is not None and rw_col < len(row) else ""
            alamat_val = transform_address(raw_alamat, raw_rt, raw_rw, mode=address_mode)

            voter_extra: dict[str, str] = {}
            for snake_k, col_i in extra_fields_cols.items():
                if col_i < len(row):
                    cell_extra = str(row[col_i]).strip()
                    if not is_blocked_pii(cell_extra):
                        voter_extra[snake_k] = cell_extra

            voter = VoterData(
                nomor=final_no,
                nama=nama_val,
                jenis_kelamin=jk_val,
                umur=umur_val,
                alamat=alamat_val,
                sheet_name=sheet_name,
                wilayah=wilayah_name,
                row_index=excel_row_num,
                extra_fields=voter_extra,
            )
            voters.append(voter)

        voter_count = len(voters)
        pages = math.ceil(voter_count / 2) if voter_count > 0 else 0

        summary = SheetScanSummary(
            sheet_name=sheet_name,
            wilayah_text=wilayah_name,
            is_voter_sheet=voter_count > 0,
            is_selected=voter_count > 0,
            total_voters=voter_count,
            expected_pages=pages,
            issues=sheet_issues,
            voters=voters,
            header_row_index=header_idx,
            schema_type=schema_type,
            schema_message=schema_msg,
            column_mapping=col_map,
            available_columns=avail_cols,
        )
        summaries.append(summary)
        all_issues.extend(sheet_issues)

        if summary.is_voter_sheet:
            total_all_voters += voter_count
            total_all_pages += pages

            if not preview_samples and voters:
                for v in voters[:5]:
                    preview_samples.append(
                        f"No: {v.nomor} | Nama: {v.nama} | JK: {v.jenis_kelamin} | "
                        f"Umur: {v.umur} | Alamat: {v.alamat}"
                    )

    voter_sheets = [s for s in summaries if s.is_voter_sheet]
    is_valid = len(voter_sheets) > 0 and len(all_issues) == 0

    return PreFlightScanResult(
        is_valid=is_valid,
        total_sheets=len(voter_sheets),
        sheet_summaries=summaries,
        total_voters=total_all_voters,
        total_pages=total_all_pages,
        all_issues=all_issues,
        preview_samples=preview_samples,
        schema_type=top_schema_type,
        schema_message=top_schema_msg,
        column_mapping=top_col_mapping,
        available_columns=top_available_cols,
    )
