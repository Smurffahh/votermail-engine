import math
import os
import re
from typing import Any
import pandas as pd
from pydantic import BaseModel, Field

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
