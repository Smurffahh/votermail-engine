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

class PreFlightScanResult(BaseModel):
    is_valid: bool = True
    total_sheets: int = 0
    sheet_summaries: list[SheetScanSummary] = Field(default_factory=list)
    total_voters: int = 0
    total_pages: int = 0
    all_issues: list[ScanIssue] = Field(default_factory=list)
    preview_samples: list[str] = Field(default_factory=list)

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

def scan_and_validate_excel(file_path: str, address_mode: str = "lengkap") -> PreFlightScanResult:
    if not os.path.exists(file_path):
        return PreFlightScanResult(
            is_valid=False,
            all_issues=[ScanIssue(sheet_name="", row_index=0, column_name="File", issue="File Excel tidak ditemukan.")],
        )

    xl = pd.ExcelFile(file_path)
    summaries: list[SheetScanSummary] = []
    total_all_voters = 0
    total_all_pages = 0
    all_issues: list[ScanIssue] = []
    preview_samples: list[str] = []

    for sheet_name in xl.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str, keep_default_na=False)
        wilayah_name = _find_wilayah_text(df, sheet_name)

        # Detect index row where column A is '1' and column C is '3' (standard row 11)
        index_row_idx: int | None = None
        for r_idx in range(min(20, len(df))):
            row_vals = [str(df.iloc[r_idx, c]).strip() for c in range(min(12, df.shape[1]))]
            if len(row_vals) >= 3 and row_vals[0] == "1" and row_vals[2] == "3":
                index_row_idx = r_idx
                break

        if index_row_idx is None:
            summaries.append(
                SheetScanSummary(
                    sheet_name=sheet_name,
                    wilayah_text=wilayah_name,
                    is_voter_sheet=False,
                    is_selected=False,
                )
            )
            continue

        col_map = {
            "no": 0,
            "nama": 2,
            "umur": 5,
            "jk": 7,
            "alamat": 8,
            "rt": 9,
            "rw": 10,
        }

        start_voter_row = index_row_idx + 1
        voters: list[VoterData] = []
        sheet_issues: list[ScanIssue] = []

        recap_keywords = ("jumlah pemilih", "status perkawinan", "laki-laki", "perempuan", "total", "b =", "s =", "p =")

        for r_idx in range(start_voter_row, len(df)):
            row = [str(df.iloc[r_idx, c]).strip() for c in range(df.shape[1])]
            no_val = row[col_map["no"]] if col_map["no"] < len(row) else ""
            nama_val = row[col_map["nama"]] if col_map["nama"] < len(row) else ""

            first_cols_combined = " ".join(row[:4]).lower()
            if any(k in first_cols_combined for k in recap_keywords):
                continue
            if not no_val and not nama_val:
                continue
            if not no_val.isdigit():
                continue

            excel_row_num = r_idx + 1

            if not nama_val:
                sheet_issues.append(
                    ScanIssue(
                        sheet_name=sheet_name,
                        row_index=excel_row_num,
                        column_name="Nama Pemilih",
                        issue="Nama pemilih kosong.",
                    )
                )

            raw_jk = row[col_map["jk"]] if col_map["jk"] < len(row) else ""
            jk_val = transform_gender(raw_jk)

            raw_umur = row[col_map["umur"]] if col_map["umur"] < len(row) else ""
            umur_val = transform_age(raw_umur)

            raw_alamat = row[col_map["alamat"]] if col_map["alamat"] < len(row) else ""
            raw_rt = row[col_map["rt"]] if col_map["rt"] < len(row) else ""
            raw_rw = row[col_map["rw"]] if col_map["rw"] < len(row) else ""
            alamat_val = transform_address(raw_alamat, raw_rt, raw_rw, mode=address_mode)

            voter = VoterData(
                nomor=no_val,
                nama=nama_val,
                jenis_kelamin=jk_val,
                umur=umur_val,
                alamat=alamat_val,
                sheet_name=sheet_name,
                wilayah=wilayah_name,
                row_index=excel_row_num,
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
    )
