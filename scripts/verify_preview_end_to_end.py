# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import os
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pypdf
import fitz

from app.config.settings import (
    BASE_DIR,
    OUTPUT_DIR,
    get_default_word_path,
    get_default_excel_path,
)
from app.core.word_preparer import check_and_prepare_word_template
from app.core.excel_scanner import scan_and_validate_excel
from app.core.mail_merger import generate_preview_via_word
from app.core.security import to_win_abs_path

def is_ms_word_installed() -> bool:
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        w = win32com.client.DispatchEx("Word.Application")
        w.Visible = False
        w.Quit()
        pythoncom.CoUninitialize()
        return True
    except Exception:
        return False

def main():
    if not is_ms_word_installed():
        print("[SKIP] Microsoft Word COM tidak terinstal di lingkungan ini. Melewati pengujian Word COM.")
        return

    word_file = get_default_word_path()
    excel_file = get_default_excel_path()
    preview_pdf = str(OUTPUT_DIR / "Pratinjau_Surat_Word_1_Halaman.pdf")

    print("[1] Cek & Siapkan Template Word...")
    word_res = check_and_prepare_word_template(word_file)
    print("Word Check is_valid:", word_res.is_valid)
    print("Word Badge:", word_res.status_badge)
    print("Prepared docx:", word_res.prepared_docx_path)
    assert word_res.is_valid, f"Word check failed: {word_res.message}"

    print("\n[2] Scan Excel DPT...")
    excel_res = scan_and_validate_excel(excel_file, address_mode="lengkap")
    print("Excel Scan is_valid:", excel_res.is_valid)
    print("Total voters:", excel_res.total_voters)
    sample_voters = []
    for s in excel_res.sheet_summaries:
        if s.voters:
            sample_voters = s.voters[:2]
            break
    print(f"Sample voters: {sample_voters[0].nama} & {sample_voters[1].nama}")

    print("\n[3] Generate Preview via Word COM...")
    ok, msg, final_pdf = generate_preview_via_word(
        template_docx_path=word_res.prepared_docx_path,
        sample_voters=sample_voters,
        output_pdf_path=preview_pdf,
    )
    print("Preview ok:", ok)
    print("Preview msg:", msg)
    print("Final PDF path:", final_pdf)
    assert ok, f"Preview failed: {msg}"
    assert os.path.exists(final_pdf), "Final PDF file not found!"

    print("\n[4] Audit & Verifikasi PDF...")
    reader = pypdf.PdfReader(final_pdf)
    page_count = len(reader.pages)
    print(f"Page Count: {page_count} (Expect exactly 1)")
    assert page_count == 1, f"Expected 1 page, got {page_count}!"

    text = reader.pages[0].extract_text()
    
    checklist = [
        ("PANITIA PEMILIHAN (case-insensitive)", lambda t: t.upper().count("PANITIA PEMILIHAN") >= 2),
        ("Dikeluarkan pada tanggal", lambda t: t.count("Dikeluarkan pada tanggal") >= 2),
        ("Karangjengkol, ....................................", lambda t: t.count("Karangjengkol, ....................................") >= 2),
        ("Penerima/Pemilih,", lambda t: t.count("Penerima/Pemilih,") >= 2),
        ("..............................................", lambda t: t.count("..............................................") >= 2),
        ("DAFTAR PEMILIH TETAP NOMOR : 1", lambda t: t.count("DAFTAR PEMILIH TETAP NOMOR : 1") >= 1),
        ("DAFTAR PEMILIH TETAP NOMOR : 2", lambda t: t.count("DAFTAR PEMILIH TETAP NOMOR : 2") >= 1),
        ("ALIF MUBAROK", lambda t: t.count("ALIF MUBAROK") >= 1),
        ("HAYATUNISA", lambda t: t.count("HAYATUNISA") >= 1),
        ("21 Tahun", lambda t: t.count("21 Tahun") >= 1),
        ("19 Tahun", lambda t: t.count("19 Tahun") >= 1),
        ("KARANGJENGKOL, RT 001 / RW 001", lambda t: t.count("KARANGJENGKOL, RT 001 / RW 001") >= 2),
        ("Tidak boleh diwakilkan.", lambda t: t.count("Tidak boleh diwakilkan.") >= 2),
    ]

    all_passed = True
    for label, check_fn in checklist:
        passed = check_fn(text)
        print(f"  [{'PASS' if passed else 'FAIL'}] {label}")
        if not passed:
            all_passed = False

    assert all_passed, "Checklist verification failed!"

    # Render PNG for visual proof
    doc_fitz = fitz.open(final_pdf)
    pix = doc_fitz[0].get_pixmap(dpi=150)
    png_out = str(OUTPUT_DIR / "Pratinjau_Surat_Word_1_Halaman_Verifikasi.png")
    pix.save(png_out)
    print(f"\n[5] Rendered visual preview image: {png_out}")
    print("All end-to-end verifications PASSED 100%!")

if __name__ == "__main__":
    main()
