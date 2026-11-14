# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import os
from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
from pydantic import BaseModel, Field
from pypdf import PdfReader

from app.core.excel_scanner import VoterData
from app.core.security import sanitize_cell_text, validate_safe_output_path

class VerificationAuditItem(BaseModel):
    sheet_name: str
    baris_excel: int
    nomor_urut: str
    nama_pemilih: str
    halaman_pdf: int
    status: str
    keterangan: str

class VerificationResult(BaseModel):
    sheet_name: str
    pdf_path: str
    is_valid: bool = True
    expected_voters: int = 0
    matched_voters: int = 0
    accuracy_percent: float = 0.0
    expected_pages: int = 0
    actual_pages: int = 0
    audit_summary: str = ""
    audit_items: list[VerificationAuditItem] = Field(default_factory=list)

def verify_kadus_pdf(
    pdf_path: str,
    sheet_name: str,
    voters: list[VoterData],
    layout_mode: str = "2-UP",
) -> VerificationResult:
    result = VerificationResult(
        sheet_name=sheet_name,
        pdf_path=pdf_path,
        expected_voters=len(voters),
    )

    if not os.path.exists(pdf_path):
        result.is_valid = False
        result.audit_summary = f"[GAGAL] File PDF tidak ditemukan: {os.path.basename(pdf_path)}"
        return result

    try:
        reader = PdfReader(pdf_path)
        result.actual_pages = len(reader.pages)
    except Exception as exc:
        result.is_valid = False
        result.audit_summary = f"[ERROR CORRUPT] Gagal membaca PDF: {exc}"
        return result

    import math
    if layout_mode == "1-UP":
        expected_pages = len(voters)
    else:
        expected_pages = math.ceil(len(voters) / 2) if voters else 0
    result.expected_pages = expected_pages

    if result.actual_pages != expected_pages:
        result.is_valid = False

    page_texts: list[str] = []
    for page in reader.pages:
        try:
            page_texts.append(page.extract_text() or "")
        except Exception:
            page_texts.append("")

    matched = 0
    for idx, voter in enumerate(voters):
        target_page_idx = idx if layout_mode == "1-UP" else idx // 2
        target_page_num = target_page_idx + 1

        if target_page_idx >= len(page_texts):
            result.is_valid = False
            result.audit_items.append(
                VerificationAuditItem(
                    sheet_name=sheet_name,
                    baris_excel=voter.row_index,
                    nomor_urut=voter.nomor,
                    nama_pemilih=voter.nama,
                    halaman_pdf=target_page_num,
                    status="HALAMAN_HILANG",
                    keterangan=f"Halaman {target_page_num} tidak ditemukan di dalam PDF",
                )
            )
            continue

        content = page_texts[target_page_idx]
        content_upper = content.upper()

        nama_ok = (not voter.nama) or (voter.nama.strip().upper() in content_upper)
        jk_ok = (not voter.jenis_kelamin) or (voter.jenis_kelamin in content)
        umur_ok = (not voter.umur) or (voter.umur in content)

        if nama_ok and jk_ok and umur_ok:
            matched += 1
            result.audit_items.append(
                VerificationAuditItem(
                    sheet_name=sheet_name,
                    baris_excel=voter.row_index,
                    nomor_urut=voter.nomor,
                    nama_pemilih=voter.nama,
                    halaman_pdf=target_page_num,
                    status="TERVERIFIKASI",
                    keterangan="Data pemilih, gender, dan umur terverifikasi di halaman PDF",
                )
            )
        else:
            result.is_valid = False
            missing = []
            if not nama_ok:
                missing.append("Nama")
            if not jk_ok:
                missing.append("Jenis Kelamin")
            if not umur_ok:
                missing.append("Umur")

            result.audit_items.append(
                VerificationAuditItem(
                    sheet_name=sheet_name,
                    baris_excel=voter.row_index,
                    nomor_urut=voter.nomor,
                    nama_pemilih=voter.nama,
                    halaman_pdf=target_page_num,
                    status="DATA_TIDAK_LENGKAP",
                    keterangan=f"Elemen teks tidak terdeteksi di PDF: {', '.join(missing)}",
                )
            )

    result.matched_voters = matched
    if len(voters) > 0:
        result.accuracy_percent = round((matched / len(voters)) * 100, 2)
    else:
        result.accuracy_percent = 100.0

    if result.is_valid and result.accuracy_percent == 100.0:
        result.audit_summary = (
            f"[VERIFIED - DATA 100% SESUAI] {sheet_name}: "
            f"{matched}/{len(voters)} pemilih terverifikasi di {result.actual_pages} halaman."
        )
    else:
        result.audit_summary = (
            f"[PERINGATAN AUDIT] {sheet_name}: "
            f"Akurasi {result.accuracy_percent}% ({matched}/{len(voters)} cocok, "
            f"{len(voters) - matched} anomali)."
        )

    return result

def export_audit_excel(audit_results: list[VerificationResult], output_dir: str) -> str | None:
    all_rows = []
    for res in audit_results:
        for item in res.audit_items:
            all_rows.append({
                "Sheet / Kadus": item.sheet_name,
                "Baris Excel": item.baris_excel if item.baris_excel > 0 else "-",
                "Nomor Urut": item.nomor_urut,
                "Nama Pemilih": sanitize_cell_text(item.nama_pemilih),
                "Halaman PDF": item.halaman_pdf,
                "Status Verifikasi": item.status,
                "Keterangan": item.keterangan,
            })

    if not all_rows:
        return None

    df = pd.DataFrame(all_rows)
    report_path = validate_safe_output_path(output_dir, "Laporan_Cek_PDF.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Audit Verifikasi PDF"

    ws.append([
        "Sheet / Kadus", "Baris Excel", "Nomor Urut",
        "Nama Pemilih", "Halaman PDF", "Status Verifikasi", "Keterangan"
    ])

    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    fill_ok = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_warn = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    font_data = Font(name="Segoe UI", size=10)

    for row_idx, r in enumerate(all_rows, start=2):
        ws.append([
            r["Sheet / Kadus"],
            r["Baris Excel"],
            r["Nomor Urut"],
            r["Nama Pemilih"],
            r["Halaman PDF"],
            r["Status Verifikasi"],
            r["Keterangan"],
        ])
        status = r["Status Verifikasi"]
        row_fill = fill_ok if status == "TERVERIFIKASI" else fill_warn

        for col_idx in range(1, 8):
            c = ws.cell(row=row_idx, column=col_idx)
            c.font = font_data
            c.border = thin_border
            if col_idx in (2, 3, 5, 6):
                c.alignment = Alignment(horizontal="center")
            if col_idx == 6:
                c.fill = row_fill

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(report_path)
    return str(report_path)
