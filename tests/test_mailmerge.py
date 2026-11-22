# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import functools
import os
from pathlib import Path
import shutil
import sys
from typing import Any, Callable
import unittest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

py_base = Path(sys.executable).parent
tcl_dir = py_base / "tcl" / "tcl8.6"
tk_dir = py_base / "tcl" / "tk8.6"
if tcl_dir.exists():
    os.environ["TCL_LIBRARY"] = str(tcl_dir)
if tk_dir.exists():
    os.environ["TK_LIBRARY"] = str(tk_dir)

import time

def create_test_ctk_root() -> Any:
    import customtkinter as ctk
    for attempt in range(4):
        try:
            return ctk.CTk()
        except Exception:
            if attempt == 3:
                raise
            time.sleep(0.3)

from pypdf import PdfReader

@functools.lru_cache(maxsize=1)
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

def create_synthetic_test_pdf(file_path: str, pages_text: list[str]) -> str:
    page_count = len(pages_text)
    page_ids = [3 + i * 2 for i in range(page_count)]
    content_ids = [4 + i * 2 for i in range(page_count)]
    font_id = 3 + page_count * 2

    lines = ['%PDF-1.4']
    offsets = {}

    def add_obj(num, content):
        offsets[num] = sum(len(line.encode('latin1')) + 1 for line in lines)
        lines.append(f'{num} 0 obj')
        lines.append(content)
        lines.append('endobj')

    kids_str = ' '.join(f'{pid} 0 R' for pid in page_ids)
    add_obj(1, '<</Type /Catalog /Pages 2 0 R>>')
    add_obj(2, f'<</Type /Pages /Kids [{kids_str}] /Count {page_count}>>')

    for i, text in enumerate(pages_text):
        pid = page_ids[i]
        cid = content_ids[i]
        stream_data = f'BT /F1 12 Tf 50 700 Td ({text}) Tj ET\n'
        slen = len(stream_data.encode('latin1'))
        add_obj(pid, f'<</Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents {cid} 0 R /Resources <</Font <</F1 {font_id} 0 R>>>>>>')
        add_obj(cid, f'<</Length {slen}>>\nstream\n{stream_data}endstream')

    add_obj(font_id, '<</Type /Font /Subtype /Type1 /BaseFont /Helvetica>>')

    xref_offset = sum(len(line.encode('latin1')) + 1 for line in lines)
    total_objs = font_id + 1
    lines.append('xref')
    lines.append(f'0 {total_objs}')
    lines.append('0000000000 65535 f ')
    for i in range(1, total_objs):
        lines.append(f'{offsets[i]:010d} 00000 n ')
    lines.append('trailer')
    lines.append(f'<</Size {total_objs} /Root 1 0 R>>')
    lines.append('startxref')
    lines.append(f'{xref_offset}')
    lines.append('%%EOF')

    data = '\n'.join(lines).encode('latin1')
    with open(file_path, 'wb') as f:
        f.write(data)
    return file_path

from app.config.settings import TEMPLATES_DIR, DATA_DIR, TEMP_DIR
from app.core.excel_scanner import (
    VoterData,
    scan_and_validate_excel,
    transform_address,
    transform_age,
    transform_gender,
)
from app.core.mail_merger import (
    generate_preview_via_word,
    merge_voters_to_pdf_via_word_com,
)
from app.core.pdf_verifier import (
    VerificationResult,
    export_audit_excel,
    verify_kadus_pdf,
)
from app.core.security import (
    check_file_accessible_for_writing,
    sanitize_cell_text,
    sanitize_filename,
    validate_safe_output_path,
)
from app.core.word_preparer import (
    WordCheckResult,
    check_and_prepare_word_template,
)
from scripts import generate_sample_data

class TestMailMergeApp(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.test_dir = os.path.abspath("./test_e2e_output")
        os.makedirs(cls.test_dir, exist_ok=True)
        cls.sample_dpt_file = os.path.abspath("./sample_dpt_karangjengkol.xlsx")
        generate_sample_data.create_dpt_karangjengkol_sample(cls.sample_dpt_file, voters_per_sheet=25)

        real_dpt_candidate = TEMPLATES_DIR / "DAFTAR PEMILIH TETAP DESA KARANGJENGKOL.xlsx"
        cls.real_dpt_file = str(real_dpt_candidate.resolve()) if real_dpt_candidate.exists() else ""

        real_doc_candidate = TEMPLATES_DIR / "SURAT PANGGILAN.doc"
        cls.real_doc_file = str(real_doc_candidate.resolve()) if real_doc_candidate.exists() else ""

        ready_docx_candidate = TEMP_DIR / "template_ready.docx"
        if ready_docx_candidate.exists():
            try:
                os.remove(ready_docx_candidate)
            except Exception:
                pass
        cls.ready_docx_file = ""

    @classmethod
    def tearDownClass(cls) -> None:
        if os.path.exists(cls.test_dir):
            shutil.rmtree(cls.test_dir, ignore_errors=True)
        if hasattr(cls, "sample_dpt_file") and os.path.exists(cls.sample_dpt_file):
            try:
                os.remove(cls.sample_dpt_file)
            except Exception:
                pass

    def test_01_security_sanitization_and_paths(self) -> None:
        self.assertEqual(sanitize_filename("KADUS 1 / Krajan"), "KADUS 1 _ Krajan")
        self.assertEqual(sanitize_filename("RW:05\\RT:02*?"), "RW_05_RT_02")
        self.assertTrue(sanitize_filename("CON").startswith("WILAYAH_CON"))
        self.assertTrue(sanitize_filename("PRN").startswith("WILAYAH_PRN"))
        self.assertTrue(sanitize_filename("NUL.pdf").startswith("WILAYAH_NUL"))
        self.assertTrue(sanitize_filename("COM1").startswith("WILAYAH_COM1"))
        self.assertTrue(sanitize_filename("LPT3").startswith("WILAYAH_LPT3"))
        self.assertEqual(sanitize_filename("KADUS\u202e1\x00_test"), "KADUS1_test")
        self.assertEqual(sanitize_filename("KADUS 1...   "), "KADUS 1")
        self.assertEqual(sanitize_filename("..."), "DOKUMEN_OUTPUT")
        self.assertEqual(sanitize_filename("../evil/path"), "evil_path")

        safe_path = validate_safe_output_path(self.test_dir, "document.pdf")
        self.assertTrue(safe_path.is_relative_to(Path(self.test_dir).resolve()))

        with self.assertRaises(PermissionError):
            validate_safe_output_path(self.test_dir, "../../etc/passwd.pdf")

        self.assertEqual(sanitize_cell_text("=CMD('calc')"), "'=CMD('calc')")
        self.assertEqual(sanitize_cell_text("+123"), "'+123")
        self.assertEqual(sanitize_cell_text("Normal Text"), "Normal Text")

        accessible, msg = check_file_accessible_for_writing(Path(self.test_dir) / "check_non_existent.pdf")
        self.assertTrue(accessible)
        self.assertEqual(msg, "")

    def test_02_transformations(self) -> None:
        self.assertEqual(transform_gender("L"), "Laki-laki")
        self.assertEqual(transform_gender("lk"), "Laki-laki")
        self.assertEqual(transform_gender("LAKI-LAKI"), "Laki-laki")
        self.assertEqual(transform_gender("P"), "Perempuan")
        self.assertEqual(transform_gender("pr"), "Perempuan")
        self.assertEqual(transform_gender("PEREMPUAN"), "Perempuan")

        self.assertEqual(transform_age("21"), "21 Tahun")
        self.assertEqual(transform_age("19.0"), "19 Tahun")
        self.assertEqual(transform_age("52 Tahun"), "52 Tahun")

        self.assertEqual(
            transform_address("KARANGJENGKOL", "1", "1", mode="lengkap"),
            "KARANGJENGKOL, RT 001 / RW 001",
        )
        self.assertEqual(
            transform_address("KARANGJENGKOL", "1", "1", mode="hanya_alamat"),
            "KARANGJENGKOL",
        )

    def test_03_excel_scanner_on_generated_sample(self) -> None:
        self.assertTrue(os.path.exists(self.sample_dpt_file))
        res = scan_and_validate_excel(self.sample_dpt_file, address_mode="lengkap")

        self.assertTrue(res.is_valid)
        self.assertEqual(res.total_sheets, 5)
        self.assertEqual(res.total_voters, 125)
        self.assertEqual(res.total_pages, 65)
        self.assertEqual(len(res.all_issues), 0)

        for s in res.sheet_summaries:
            self.assertEqual(s.total_voters, 25)
            self.assertEqual(s.expected_pages, 13)
            self.assertTrue(s.is_voter_sheet)
            self.assertTrue(s.is_selected)
            self.assertEqual(len(s.issues), 0)

            v1 = s.voters[0]
            self.assertEqual(v1.nomor, "1")
            self.assertIn("Tahun", v1.umur)
            self.assertIn(v1.jenis_kelamin, ("Laki-laki", "Perempuan"))
            self.assertIn("RT 00", v1.alamat)

    def test_04_excel_scanner_on_real_dpt(self) -> None:
        if not os.path.exists(self.real_dpt_file):
            self.skipTest("File bahan DPT real tidak tersedia")

        res = scan_and_validate_excel(self.real_dpt_file, address_mode="lengkap")
        self.assertTrue(res.is_valid)
        self.assertEqual(res.total_sheets, 5)
        self.assertEqual(res.total_voters, 3505)
        self.assertEqual(res.total_pages, 1754)
        self.assertEqual(len(res.all_issues), 0)

        sheet_counts = {s.sheet_name: s.total_voters for s in res.sheet_summaries if s.is_voter_sheet}
        self.assertEqual(sheet_counts["KADUS 1"], 685)
        self.assertEqual(sheet_counts["KADUS 2"], 563)
        self.assertEqual(sheet_counts["KADUS 3"], 862)
        self.assertEqual(sheet_counts["KADUS 4"], 878)
        self.assertEqual(sheet_counts["KADUS 5"], 517)

        rekap_summary = next(s for s in res.sheet_summaries if s.sheet_name == "REKAP")
        self.assertFalse(rekap_summary.is_voter_sheet)
        self.assertFalse(rekap_summary.is_selected)

    def test_05_word_checker_missing_and_real_doc(self) -> None:
        missing_res = check_and_prepare_word_template("file_tidak_ada.docx", work_dir=self.test_dir)
        self.assertFalse(missing_res.is_valid)
        self.assertEqual(missing_res.status_badge, "Perlu Perbaikan")

        if not os.path.exists(self.real_doc_file):
            self.skipTest("File bahan SURAT PANGGILAN.doc tidak tersedia")

        res = check_and_prepare_word_template(self.real_doc_file, work_dir=self.test_dir)
        self.assertTrue(res.is_valid)
        self.assertTrue("Siap Mailing" in res.status_badge)
        self.assertEqual(res.layout_mode, "2-UP")
        self.assertTrue(os.path.exists(res.prepared_docx_path))

        expected_fields = ["Alamat", "Laki-laki/Perempuan", "Nama Pemilih", "Nomor", "Umur"]
        self.assertEqual(res.detected_fields_slot1, expected_fields)
        self.assertEqual(res.detected_fields_slot2, expected_fields)

    def test_05b_word_checker_on_sample_dummy_docx(self) -> None:
        sample_docx = TEMPLATES_DIR / "SAMPLE_SURAT_DUMMY.docx"
        self.assertTrue(sample_docx.exists())

        dummy_work = os.path.join(self.test_dir, "dummy_work")
        res = check_and_prepare_word_template(str(sample_docx), work_dir=dummy_work)
        self.assertTrue(res.is_valid)
        self.assertTrue("Siap Mailing" in res.status_badge)
        self.assertEqual(res.layout_mode, "2-UP")
        self.assertTrue(os.path.exists(res.prepared_docx_path))

        expected_fields = ["Alamat", "Laki-laki/Perempuan", "Nama Pemilih", "Nomor", "Umur"]
        self.assertEqual(res.detected_fields_slot1, expected_fields)
        self.assertEqual(res.detected_fields_slot2, expected_fields)

    def test_06_mail_merger_pairwise_and_odd_voters_via_word(self) -> None:
        if not is_ms_word_installed():
            self.skipTest("Microsoft Word COM tidak terinstal di lingkungan ini (CI runner)")

        tpl_path = self.ready_docx_file if os.path.exists(self.ready_docx_file) else os.path.join(self.test_dir, "template_ready.docx")
        if not os.path.exists(tpl_path):
            sample_docx = str(TEMPLATES_DIR / "SAMPLE_SURAT_DUMMY.docx")
            res = check_and_prepare_word_template(sample_docx, work_dir=self.test_dir)
            tpl_path = res.prepared_docx_path

        voters = [
            VoterData(nomor="1", nama="ALIF MUBAROK", jenis_kelamin="Laki-laki", umur="21 Tahun", alamat="KARANGJENGKOL, RT 001 / RW 001", sheet_name="KADUS 1", wilayah="KADUS 1", row_index=12),
            VoterData(nomor="2", nama="HAYATUNISA", jenis_kelamin="Perempuan", umur="19 Tahun", alamat="KARANGJENGKOL, RT 001 / RW 001", sheet_name="KADUS 1", wilayah="KADUS 1", row_index=13),
            VoterData(nomor="3", nama="TARSUN", jenis_kelamin="Laki-laki", umur="52 Tahun", alamat="KARANGJENGKOL, RT 001 / RW 001", sheet_name="KADUS 1", wilayah="KADUS 1", row_index=14),
        ]
        out_pdf = os.path.join(self.test_dir, "test_word_3_voters.pdf")
        ok, pages, msg = merge_voters_to_pdf_via_word_com(
            template_docx_path=tpl_path,
            voters=voters,
            sheet_name="KADUS 1",
            wilayah_name="KADUS 1",
            output_pdf_path=out_pdf,
        )
        self.assertTrue(ok)
        self.assertEqual(pages, 2)
        self.assertTrue(os.path.exists(out_pdf))

        reader = PdfReader(out_pdf)
        self.assertEqual(len(reader.pages), 2)

        p1_text = reader.pages[0].extract_text()
        p2_text = reader.pages[1].extract_text()

        self.assertIn("ALIF MUBAROK", p1_text)
        self.assertIn("HAYATUNISA", p1_text)
        self.assertNotIn("TARSUN", p1_text)

        self.assertIn("TARSUN", p2_text)

    def test_07_word_preview_generation(self) -> None:
        if not is_ms_word_installed():
            self.skipTest("Microsoft Word COM tidak terinstal di lingkungan ini (CI runner)")

        tpl_path = self.ready_docx_file if os.path.exists(self.ready_docx_file) else os.path.join(self.test_dir, "template_ready.docx")
        if not os.path.exists(tpl_path):
            sample_docx = str(TEMPLATES_DIR / "SAMPLE_SURAT_DUMMY.docx")
            res = check_and_prepare_word_template(sample_docx, work_dir=self.test_dir)
            tpl_path = res.prepared_docx_path

        out_preview = os.path.join(self.test_dir, "test_word_preview.pdf")
        res = generate_preview_via_word(tpl_path, [], out_preview)
        ok = res[0]
        self.assertTrue(ok)
        self.assertTrue(os.path.exists(out_preview))

        reader = PdfReader(out_preview)
        self.assertEqual(len(reader.pages), 1)
        text = reader.pages[0].extract_text()
        self.assertTrue("SURAT PEMBERITAHUAN" in text)
        self.assertIn("ALIF MUBAROK", text)

    def test_08_pdf_verifier_success_and_audit_export(self) -> None:
        voters = [
            VoterData(nomor="1", nama="ALIF MUBAROK", jenis_kelamin="Laki-laki", umur="21 Tahun", alamat="KARANGJENGKOL, RT 001 / RW 001", sheet_name="KADUS 1", wilayah="KADUS 1", row_index=12),
            VoterData(nomor="2", nama="HAYATUNISA", jenis_kelamin="Perempuan", umur="19 Tahun", alamat="KARANGJENGKOL, RT 001 / RW 001", sheet_name="KADUS 1", wilayah="KADUS 1", row_index=13),
        ]
        out_pdf = os.path.join(self.test_dir, "test_word_verify_success.pdf")

        if is_ms_word_installed():
            tpl_path = self.ready_docx_file if os.path.exists(self.ready_docx_file) else os.path.join(self.test_dir, "template_ready.docx")
            if not os.path.exists(tpl_path):
                sample_docx = str(TEMPLATES_DIR / "SAMPLE_SURAT_DUMMY.docx")
                res = check_and_prepare_word_template(sample_docx, work_dir=self.test_dir)
                tpl_path = res.prepared_docx_path

            merge_voters_to_pdf_via_word_com(
                template_docx_path=tpl_path,
                voters=voters,
                sheet_name="KADUS 1",
                wilayah_name="KADUS 1",
                output_pdf_path=out_pdf,
            )
        else:
            create_synthetic_test_pdf(out_pdf, ["ALIF MUBAROK HAYATUNISA Laki-laki Perempuan 21 Tahun 19 Tahun KARANGJENGKOL, RT 001 / RW 001"])

        res = verify_kadus_pdf(out_pdf, "KADUS 1", voters)
        self.assertTrue(res.is_valid)
        self.assertEqual(res.accuracy_percent, 100.0)
        self.assertEqual(res.matched_voters, 2)
        self.assertEqual(res.actual_pages, 1)
        self.assertIn("100% SESUAI", res.audit_summary)

        excel_report = export_audit_excel([res], self.test_dir)
        self.assertIsNotNone(excel_report)
        self.assertTrue(os.path.exists(excel_report))

    def test_09_pdf_verifier_discrepancy_detection(self) -> None:
        voters_actual = [
            VoterData(nomor="1", nama="ALIF MUBAROK", jenis_kelamin="Laki-laki", umur="21 Tahun", alamat="KARANGJENGKOL", sheet_name="KADUS 1", wilayah="KADUS 1", row_index=12),
        ]
        out_pdf = os.path.join(self.test_dir, "test_verify_discrepancy.pdf")

        if is_ms_word_installed():
            tpl_path = self.ready_docx_file if os.path.exists(self.ready_docx_file) else os.path.join(self.test_dir, "template_ready.docx")
            if not os.path.exists(tpl_path):
                sample_docx = str(TEMPLATES_DIR / "SAMPLE_SURAT_DUMMY.docx")
                res = check_and_prepare_word_template(sample_docx, work_dir=self.test_dir)
                tpl_path = res.prepared_docx_path

            merge_voters_to_pdf_via_word_com(
                template_docx_path=tpl_path,
                voters=voters_actual,
                sheet_name="KADUS 1",
                wilayah_name="KADUS 1",
                output_pdf_path=out_pdf,
            )
        else:
            create_synthetic_test_pdf(out_pdf, ["ALIF MUBAROK Laki-laki 21 Tahun KARANGJENGKOL"])

        voters_tampered = [
            VoterData(nomor="1", nama="ALIF MUBAROK", jenis_kelamin="Laki-laki", umur="21 Tahun", alamat="KARANGJENGKOL", sheet_name="KADUS 1", wilayah="KADUS 1", row_index=12),
            VoterData(nomor="2", nama="PEMILIH FIKTIF", jenis_kelamin="Perempuan", umur="99 Tahun", alamat="UNKNOWN", sheet_name="KADUS 1", wilayah="KADUS 1", row_index=13),
        ]

        res = verify_kadus_pdf(out_pdf, "KADUS 1", voters_tampered)
        self.assertFalse(res.is_valid)
        self.assertLess(res.accuracy_percent, 100.0)
        self.assertEqual(res.matched_voters, 1)

    def test_10_app_config_persistence(self) -> None:
        from app.config.settings import load_app_config, save_app_config
        cfg = load_app_config()
        self.assertIsNotNone(cfg.output_dir)
        self.assertIn(cfg.theme_mode, ("Dark", "Light"))
        ok = save_app_config(cfg)
        self.assertTrue(ok)

    def test_11_word_diff_and_snapshots(self) -> None:
        from app.core.word_preparer import (
            WordSnapshot,
            compute_word_snapshot,
            analyze_word_diff,
        )
        sample_word = TEMPLATES_DIR / "SAMPLE_SURAT_DUMMY.docx"
        self.assertTrue(sample_word.exists())

        snap1 = compute_word_snapshot(
            str(sample_word),
            str(sample_word),
            is_valid=True,
            detected_slot1=["Nomor", "Nama Pemilih"],
            detected_slot2=["Nomor", "Nama Pemilih"],
        )
        self.assertEqual(snap1.file_path, str(sample_word))
        self.assertTrue(snap1.is_valid)
        self.assertGreater(snap1.size, 0)

        diff1 = analyze_word_diff(None, snap1)
        self.assertTrue(any("Snapshot awal" in line for line in diff1))

        snap2 = WordSnapshot(
            file_path=snap1.file_path,
            sha256="fake_sha",
            mtime=snap1.mtime + 10,
            size=snap1.size + 100,
            fields_slot1=["Nomor", "Nama Pemilih", "Umur"],
            fields_slot2=snap1.fields_slot2,
            schedule_summary={"hari": "Jumat, 11 Desember 2026"},
            ketua="RATNO",
            paragraphs_count=snap1.paragraphs_count + 1,
            is_valid=True,
        )
        diff2 = analyze_word_diff(snap1, snap2)
        self.assertTrue(any("Hari/Tanggal" in line for line in diff2))
        self.assertTrue(any("Slot 1" in line for line in diff2))

    def test_12_excel_diff_and_snapshots(self) -> None:
        from app.core.excel_scanner import (
            ExcelSnapshot,
            compute_excel_snapshot,
            analyze_excel_diff,
        )
        sample_excel = TEMPLATES_DIR / "SAMPLE_DPT_DUMMY.xlsx"
        self.assertTrue(sample_excel.exists())

        scan_res = scan_and_validate_excel(str(sample_excel))
        snap1 = compute_excel_snapshot(scan_res, str(sample_excel))
        self.assertEqual(snap1.total_voters, 125)
        self.assertEqual(snap1.total_sheets, 5)

        diff1 = analyze_excel_diff(None, snap1)
        self.assertTrue(any("Snapshot awal" in line for line in diff1))

        snap2 = ExcelSnapshot(
            file_path=snap1.file_path,
            sha256="fake_sha",
            mtime=snap1.mtime + 10,
            size=snap1.size + 500,
            sheets_summary={"KADUS 1": 28, "KADUS 2": 25, "KADUS 6": 10},
            total_voters=138,
            total_sheets=6,
            invalid_rows_count=0,
            is_valid=True,
        )
        diff2 = analyze_excel_diff(snap1, snap2)
        self.assertTrue(any("KADUS 1" in line and "+3" in line for line in diff2))
        self.assertTrue(any("KADUS 6" in line for line in diff2))

    def test_13_version_and_credits(self) -> None:
        from app.config.version import (
            __app_name__,
            __author__,
            __github_user__,
            __repository__,
            __version__,
            FEATURE_REGISTRY,
            get_current_version,
            validate_version_registry,
        )
        self.assertEqual(__app_name__, "VoterMail Engine")
        self.assertEqual(__author__, "Danang Satria")
        self.assertEqual(__github_user__, "Smurffahh")
        self.assertTrue(__repository__.startswith("https://github.com/Smurffahh"))
        self.assertEqual(__version__, get_current_version())
        self.assertEqual(FEATURE_REGISTRY[-1]["version"], __version__)
        self.assertGreaterEqual(len(FEATURE_REGISTRY), 10)
        self.assertTrue(validate_version_registry())

    def test_14_dependency_bootstrapper(self) -> None:
        import main
        missing = main.get_missing_packages()
        self.assertEqual(len(missing), 0)

    def test_15_pdf_metadata(self) -> None:
        from pypdf import PdfWriter, PdfReader
        from app.core.mail_merger import apply_pdf_metadata
        from app.config.version import __version__, __author__, __repository__

        test_pdf = os.path.join(self.test_dir, "test_metadata.pdf")
        writer = PdfWriter()
        writer.add_blank_page(width=612, height=792)
        with open(test_pdf, "wb") as f:
            writer.write(f)

        apply_pdf_metadata(test_pdf, kadus_name="KADUS 1")

        reader = PdfReader(test_pdf)
        meta = reader.metadata
        self.assertIn("KADUS 1", meta.title)
        self.assertIn(__author__, meta.author)
        self.assertIn(f"v{__version__}", meta.creator)
        self.assertIn(__repository__, meta.producer)

    def test_16_ui_components_and_theme(self) -> None:
        from app.config.theme import DARK_THEME, LIGHT_THEME, get_theme
        from app.ui.components import (
            DocumentSourceRow,
            CleanMetricBar,
            KadusRowItem,
            VoterPreviewTable,
            AboutModal,
        )
        self.assertEqual(DARK_THEME.primary_action, ("#FE0075", "#38BDF8"))
        self.assertEqual(LIGHT_THEME.primary_action, ("#FE0075", "#38BDF8"))
        self.assertEqual(DARK_THEME.corner_radius_card, 16)
        self.assertEqual(DARK_THEME.corner_radius_inner, 8)
        self.assertEqual(DARK_THEME.corner_radius_control, 8)
        self.assertEqual(DARK_THEME.corner_radius_badge, 6)
        self.assertEqual(get_theme("Dark").name, "Dark")
        self.assertEqual(get_theme("Light").name, "Light")

    def test_17_boundary_and_lock_handling(self) -> None:
        from unittest.mock import patch

        with patch("os.path.getsize", return_value=160 * 1024 * 1024):
            res_xl = scan_and_validate_excel(self.sample_dpt_file)
            self.assertFalse(res_xl.is_valid)
            self.assertTrue(any("terlalu besar" in iss.issue.lower() for iss in res_xl.all_issues))

            res_doc = check_and_prepare_word_template(self.sample_dpt_file)
            self.assertFalse(res_doc.is_valid)
            self.assertTrue("terlalu besar" in res_doc.message.lower())

        locked_file = Path(self.test_dir) / "locked_file.txt"
        locked_file.write_text("sample content")

        if os.name == "nt":
            import ctypes
            GENERIC_READ = 0x80000000
            GENERIC_WRITE = 0x40000000
            FILE_SHARE_NONE = 0
            OPEN_EXISTING = 3
            handle = ctypes.windll.kernel32.CreateFileW(
                str(locked_file.resolve()),
                GENERIC_READ | GENERIC_WRITE,
                FILE_SHARE_NONE,
                None,
                OPEN_EXISTING,
                0,
                None,
            )
            try:
                is_accessible, lock_err = check_file_accessible_for_writing(locked_file)
                self.assertFalse(is_accessible)
                self.assertIn("dikunci", lock_err.lower())
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)

            is_accessible2, _ = check_file_accessible_for_writing(locked_file)
            self.assertTrue(is_accessible2)

    def test_18_custom_excel_scanner_synonyms_and_pii_blocking(self) -> None:
        import pandas as pd
        custom_excel = os.path.join(self.test_dir, "custom_dpt.xlsx")
        df = pd.DataFrame([
            ["No", "Nama Lengkap", "Gender", "Usia", "Domisili", "TPS", "Nomor Induk Kependudukan (NIK)"],
            ["1", "SITI AMINAH", "P", "28", "Dusun Krajan RT 01 RW 02", "TPS 03", "3303079999990001"],
            ["2", "AGUS SANTOSO", "L", "35", "Dusun Krajan RT 02 RW 02", "TPS 03", "3303079999990002"],
        ])
        with pd.ExcelWriter(custom_excel) as writer:
            df.to_excel(writer, sheet_name="DUSUN 1", header=False, index=False)

        res = scan_and_validate_excel(custom_excel)
        self.assertTrue(res.is_valid)
        self.assertEqual(res.schema_type, "KUSTOM_TERPETAKAN")
        self.assertEqual(res.total_voters, 2)
        summary = res.sheet_summaries[0]
        self.assertEqual(summary.voters[0].nama, "SITI AMINAH")
        self.assertEqual(summary.voters[0].jenis_kelamin, "Perempuan")
        self.assertEqual(summary.voters[0].umur, "28 Tahun")
        self.assertEqual(summary.voters[0].extra_fields.get("tps"), "TPS 03")
        self.assertNotIn("nik", summary.voters[0].extra_fields)
        self.assertNotIn("nomor_induk_kependudukan_nik", summary.voters[0].extra_fields)

    def test_19_word_checker_1up_and_manual_tags(self) -> None:
        import docx
        custom_1up_path = os.path.join(self.test_dir, "template_1up.docx")
        doc = docx.Document()
        doc.add_paragraph("SURAT UNDANGAN PEMILIHAN (1 PEMILIH / LEMBAR)")
        doc.add_paragraph("Nomor : ")
        doc.add_paragraph("Nama Pemilih : ")
        doc.add_paragraph("Jenis Kelamin : ")
        doc.add_paragraph("Usia : ")
        doc.add_paragraph("Alamat : ")
        doc.save(custom_1up_path)

        res_1up = check_and_prepare_word_template(custom_1up_path, work_dir=self.test_dir)
        self.assertTrue(res_1up.is_valid)
        self.assertEqual(res_1up.layout_mode, "1-UP")
        self.assertEqual(res_1up.template_type, "KUSTOM_1UP")
        self.assertTrue("Siap Mailing" in res_1up.status_badge)

        manual_path = os.path.join(self.test_dir, "template_manual.docx")
        doc_manual = docx.Document()
        doc_manual.add_paragraph("SURAT RESMI DENGAN TAG MANUAL")
        doc_manual.add_paragraph("Kepada Yth. {{nama}} (No: {{nomor}})")
        doc_manual.add_paragraph("Alamat: {{alamat}}")
        doc_manual.save(manual_path)

        res_manual = check_and_prepare_word_template(manual_path, work_dir=self.test_dir)
        self.assertTrue(res_manual.is_valid)
        self.assertEqual(res_manual.layout_mode, "1-UP")
        self.assertEqual(res_manual.template_type, "MANUAL_TAGS")
        self.assertTrue("Siap Mailing" in res_manual.status_badge)

    def test_20_mail_merger_1up_mode(self) -> None:
        if not is_ms_word_installed():
            self.skipTest("MS Word COM tidak tersedia di lingkungan pengujian ini")

        import docx
        custom_1up_path = os.path.join(self.test_dir, "template_1up_ready.docx")
        doc = docx.Document()
        doc.add_paragraph("SURAT RESMI: {{nama_1}}")
        doc.add_paragraph("Nomor: {{nomor_1}}")
        doc.add_paragraph("Jenis Kelamin: {{jk_1}}")
        doc.add_paragraph("Umur: {{umur_1}}")
        doc.add_paragraph("Alamat: {{alamat_1}}")
        doc.save(custom_1up_path)

        out_pdf = os.path.join(self.test_dir, "output_1up.pdf")
        voters = [
            VoterData(
                nomor="1",
                nama="SURIPTO",
                jenis_kelamin="Laki-laki",
                umur="40 Tahun",
                alamat="RT 01 RW 01",
                sheet_name="KADUS 1",
                wilayah="KADUS 1",
                row_index=1,
            ),
            VoterData(
                nomor="2",
                nama="WARSIH",
                jenis_kelamin="Perempuan",
                umur="38 Tahun",
                alamat="RT 01 RW 01",
                sheet_name="KADUS 1",
                wilayah="KADUS 1",
                row_index=2,
            ),
        ]

        ok, pages, msg = merge_voters_to_pdf_via_word_com(
            template_docx_path=custom_1up_path,
            voters=voters,
            sheet_name="KADUS 1",
            wilayah_name="KADUS 1",
            output_pdf_path=out_pdf,
            layout_mode="1-UP",
        )
        self.assertTrue(ok)
        self.assertEqual(pages, 2)
        self.assertTrue(os.path.exists(out_pdf))

        ver = verify_kadus_pdf(out_pdf, "KADUS 1", voters, layout_mode="1-UP")
        self.assertTrue(ver.is_valid)
        self.assertEqual(ver.expected_pages, 2)
        self.assertEqual(ver.actual_pages, 2)
        self.assertEqual(ver.matched_voters, 2)

    def test_21_column_mapper_helpers(self) -> None:
        from app.ui.components import get_col_letter
        from app.core.excel_scanner import is_blocked_pii, to_snake_case

        self.assertEqual(get_col_letter(0), "A")
        self.assertEqual(get_col_letter(25), "Z")
        self.assertEqual(get_col_letter(26), "AA")
        self.assertEqual(get_col_letter(27), "AB")

        self.assertTrue(is_blocked_pii("NIK"))
        self.assertTrue(is_blocked_pii("Nomor Induk Kependudukan (NIK)"))
        self.assertTrue(is_blocked_pii("No. KK"))
        self.assertFalse(is_blocked_pii("Nama Lengkap"))
        self.assertFalse(is_blocked_pii("Tempat Lahir"))

        self.assertEqual(to_snake_case("Tempat Lahir"), "tempat_lahir")
        self.assertEqual(to_snake_case("TPS / Wilayah"), "tps_wilayah")

    def test_22_multisheet_consistency_and_single_sheet_grouping(self) -> None:
        from app.core.excel_scanner import scan_and_validate_excel
        import openpyxl

        sample_excel = os.path.join(PROJECT_ROOT, "data", "templates", "SAMPLE_DPT_DUMMY.xlsx")
        res = scan_and_validate_excel(sample_excel)
        self.assertTrue(res.is_valid)
        self.assertEqual(res.total_sheets, 5)
        self.assertTrue(res.is_uniform_multisheet)
        self.assertIn("5 Sheet Seragam", res.cross_sheet_uniformity)
        self.assertTrue(res.is_standard_dpt)

        # Non-voter sheet skipping (REKAP & Blank)
        rekap_test_path = os.path.join(self.test_dir, "test_rekap_skip.xlsx")
        wb = openpyxl.Workbook()
        ws_voter = wb.active
        ws_voter.title = "KADUS 1"
        ws_voter.append(["No", "Nama", "Alamat"])
        ws_voter.append(["1", "BUDI SANTOSO", "RT 01 RW 01"])

        ws_rekap = wb.create_sheet(title="REKAPITULASI")
        ws_rekap.append(["Jumlah Pemilih", "Laki-laki", "Perempuan"])
        ws_rekap.append(["100", "50", "50"])

        ws_blank = wb.create_sheet(title="SheetKosong")
        wb.save(rekap_test_path)
        wb.close()

        res_rekap = scan_and_validate_excel(rekap_test_path)
        self.assertTrue(res_rekap.is_valid)
        self.assertEqual(res_rekap.total_sheets, 1)
        self.assertEqual(res_rekap.sheet_summaries[0].sheet_name, "KADUS 1")
        self.assertFalse(res_rekap.sheet_summaries[1].is_voter_sheet)
        self.assertFalse(res_rekap.sheet_summaries[2].is_voter_sheet)

        # Single-sheet with region grouping (KADUS column)
        single_kadus_path = os.path.join(self.test_dir, "test_single_sheet_grouping.xlsx")
        wb_single = openpyxl.Workbook()
        ws_s = wb_single.active
        ws_s.title = "DPT GABUNGAN"
        ws_s.append(["No", "Nama", "Alamat", "KADUS"])
        ws_s.append(["1", "AGUS SETIAWAN", "RT 01 RW 01", "KADUS 1"])
        ws_s.append(["2", "BAMBANG PAMUNGKAS", "RT 02 RW 01", "KADUS 1"])
        ws_s.append(["3", "CITRA KIRANA", "RT 01 RW 02", "KADUS 2"])
        ws_s.append(["4", "DEWI SARTIKA", "RT 02 RW 02", "KADUS 2"])
        wb_single.save(single_kadus_path)
        wb_single.close()

        res_single = scan_and_validate_excel(single_kadus_path)
        self.assertTrue(res_single.is_valid)
        self.assertEqual(res_single.total_sheets, 2)
        sheet_names = [s.sheet_name for s in res_single.sheet_summaries if s.is_voter_sheet]
        self.assertIn("KADUS 1", sheet_names)
        self.assertIn("KADUS 2", sheet_names)
        self.assertEqual(res_single.total_voters, 4)

    def test_23_standard_format_compatibility_validator(self) -> None:
        from app.core.excel_scanner import scan_and_validate_excel
        from app.core.word_preparer import check_and_prepare_word_template
        import docx

        sample_excel = os.path.join(PROJECT_ROOT, "data", "templates", "SAMPLE_DPT_DUMMY.xlsx")
        res_xl = scan_and_validate_excel(sample_excel)
        self.assertTrue(res_xl.is_standard_dpt)
        self.assertTrue(res_xl.format_status_text.startswith("✓ Sesuai Format Standar DPT"))

        sample_word = os.path.join(PROJECT_ROOT, "data", "templates", "SAMPLE_SURAT_DUMMY.docx")
        res_word = check_and_prepare_word_template(sample_word)
        self.assertTrue(res_word.is_standard_format)
        self.assertEqual(res_word.format_status_text, "✓ Sesuai Format Standar (2 Surat/Lembar • 12 Tag Lengkap)")

        # Custom 1-UP Word
        custom_doc_path = os.path.join(self.test_dir, "custom_1up_test.docx")
        cdoc = docx.Document()
        cdoc.add_paragraph("SURAT UNDANGAN KUSTOM")
        cdoc.add_paragraph("Nomor: {{nomor}}")
        cdoc.add_paragraph("Nama: {{nama}}")
        cdoc.add_paragraph("Alamat: {{alamat}}")
        cdoc.save(custom_doc_path)

        res_custom_word = check_and_prepare_word_template(custom_doc_path)
        self.assertTrue(res_custom_word.is_valid)
        self.assertFalse(res_custom_word.is_standard_format)
        self.assertTrue(res_custom_word.format_status_text.startswith("• Format Kustom Terdeteksi"))

    def test_24_application_assets_and_icon_helper(self) -> None:
        from app.config.settings import APP_ICON_ICO, APP_ICON_PNG
        from app.ui.components import apply_window_icon

        self.assertTrue(APP_ICON_ICO.exists(), f"APP_ICON_ICO harus ada di {APP_ICON_ICO}")
        self.assertGreater(APP_ICON_ICO.stat().st_size, 0, "Ukuran APP_ICON_ICO harus > 0 bytes")

        self.assertTrue(APP_ICON_PNG.exists(), f"APP_ICON_PNG harus ada di {APP_ICON_PNG}")
        self.assertGreater(APP_ICON_PNG.stat().st_size, 0, "Ukuran APP_ICON_PNG harus > 0 bytes")

        class MockWindow:
            def __init__(self) -> None:
                self.icon_calls: list[str] = []
                self.scheduled_calls: list[tuple[int, Any]] = []

            def iconbitmap(self, path: str) -> None:
                self.icon_calls.append(path)

            def winfo_exists(self) -> bool:
                return True

            def after(self, ms: int, fn: Any) -> None:
                self.scheduled_calls.append((ms, fn))

        mock = MockWindow()
        apply_window_icon(mock)

        self.assertEqual(len(mock.icon_calls), 1)
        self.assertEqual(mock.icon_calls[0], str(APP_ICON_ICO.resolve()))
        self.assertEqual(len(mock.scheduled_calls), 2)
        self.assertEqual(mock.scheduled_calls[0][0], 220)
        self.assertEqual(mock.scheduled_calls[1][0], 300)

        for _, cb in mock.scheduled_calls:
            cb()
        self.assertEqual(len(mock.icon_calls), 3)

        # Test robustness against non-Tk objects
        apply_window_icon(None)
        apply_window_icon(object())

    def test_25_custom_dusun_patent_address_format(self) -> None:
        from app.config.settings import (
            ADDRESS_FORMAT_OPTIONS,
            ADDRESS_FORMAT_FULL,
            ADDRESS_FORMAT_DUSUN_PATENT,
            ADDRESS_FORMAT_SHORT,
            DEFAULT_ADDRESS_FORMAT,
            PATENTED_VILLAGE_NAME,
        )
        from app.core.excel_scanner import (
            format_dusun_from_sheet,
            build_voter_address,
            VoterData,
        )

        # 1. Test format_dusun_from_sheet
        self.assertEqual(format_dusun_from_sheet("KADUS 1"), "Dusun 1")
        self.assertEqual(format_dusun_from_sheet("KADUS 5"), "Dusun 5")
        self.assertEqual(format_dusun_from_sheet("DUSUN 03"), "Dusun 3")
        self.assertEqual(format_dusun_from_sheet("Sheet2"), "Dusun 2")
        self.assertEqual(format_dusun_from_sheet("KADUS II"), "Dusun II")
        self.assertEqual(format_dusun_from_sheet(""), "Dusun 1")

        # 2. Test default address format constant
        self.assertEqual(DEFAULT_ADDRESS_FORMAT, ADDRESS_FORMAT_FULL)
        self.assertEqual(ADDRESS_FORMAT_OPTIONS[0], "Lengkap (Alamat, RT / RW)")
        self.assertIn(ADDRESS_FORMAT_DUSUN_PATENT, ADDRESS_FORMAT_OPTIONS)
        self.assertIn(ADDRESS_FORMAT_SHORT, ADDRESS_FORMAT_OPTIONS)

        # 3. Test build_voter_address with dict and VoterData
        sample_dict = {
            "alamat_raw": "Jalan Melati No 4",
            "rt_raw": "1",
            "rw_raw": "1",
            "sheet_name": "KADUS 1",
        }

        # Mode 1: Default Lengkap (Alamat, RT / RW)
        addr_default = build_voter_address(sample_dict)
        self.assertEqual(addr_default, "Jalan Melati No 4, RT 001 / RW 001")
        addr_full = build_voter_address(sample_dict, address_format=ADDRESS_FORMAT_FULL)
        self.assertEqual(addr_full, "Jalan Melati No 4, RT 001 / RW 001")

        # Mode 2: Dusun, RT / RW, Desa Karangjengkol (Patent Format)
        addr_patent = build_voter_address(sample_dict, address_format=ADDRESS_FORMAT_DUSUN_PATENT)
        self.assertEqual(addr_patent, "Dusun 1, RT 001 / RW 001, Desa Karangjengkol")

        # Mode 2 with empty alamat
        sample_empty_addr = {
            "alamat_raw": "",
            "rt_raw": "02",
            "rw_raw": "05",
            "sheet_name": "KADUS 3",
        }
        addr_patent_empty = build_voter_address(sample_empty_addr, address_format=ADDRESS_FORMAT_DUSUN_PATENT)
        self.assertEqual(addr_patent_empty, "Dusun 3, RT 002 / RW 005, Desa Karangjengkol")

        # Mode 3: Ringkas
        addr_short = build_voter_address(sample_dict, address_format=ADDRESS_FORMAT_SHORT)
        self.assertEqual(addr_short, "Jalan Melati No 4")

        # 4. Test VoterData model instance
        voter = VoterData(
            nomor="1",
            nama="Budi Santoso",
            jenis_kelamin="Laki-laki",
            umur="25 Tahun",
            alamat="KARANGJENGKOL, RT 001 / RW 001",
            sheet_name="KADUS 1",
            wilayah="KADUS 1",
            row_index=5,
            alamat_raw="KARANGJENGKOL",
            rt_raw="1",
            rw_raw="1",
            rt_formatted="001",
            rw_formatted="001",
            dusun_label="Dusun 1",
        )

        self.assertEqual(
            build_voter_address(voter, address_format=ADDRESS_FORMAT_DUSUN_PATENT),
            "Dusun 1, RT 001 / RW 001, Desa Karangjengkol",
        )
        self.assertEqual(
            build_voter_address(voter, address_format=ADDRESS_FORMAT_FULL),
            "KARANGJENGKOL, RT 001 / RW 001",
        )

    def test_26_format_elapsed_time(self) -> None:
        from app.ui.main_window import format_elapsed_time

        self.assertEqual(format_elapsed_time(0), "00:00")
        self.assertEqual(format_elapsed_time(0.4), "00:00")
        self.assertEqual(format_elapsed_time(59), "00:59")
        self.assertEqual(format_elapsed_time(59.9), "00:59")
        self.assertEqual(format_elapsed_time(60), "01:00")
        self.assertEqual(format_elapsed_time(131), "02:11")
        self.assertEqual(format_elapsed_time(3599), "59:59")
        self.assertEqual(format_elapsed_time(3600), "01:00:00")
        self.assertEqual(format_elapsed_time(3665), "01:01:05")
        self.assertEqual(format_elapsed_time(7325), "02:02:05")
        self.assertEqual(format_elapsed_time(-10), "00:00")

    def test_27_selective_reprint_filter_parsers_and_functions(self) -> None:
        from app.config.settings import (
            FILTER_MODE_ALL,
            FILTER_MODE_NUMBER,
            FILTER_MODE_RT,
            FILTER_MODE_NAME,
            FILTER_MODE_OPTIONS,
            DEFAULT_FILTER_MODE,
        )
        from app.core.excel_scanner import (
            parse_number_filter_query,
            parse_rt_filter_query,
            parse_name_filter_query,
            apply_voter_filter,
            build_filter_filename_suffix,
            VoterData,
        )

        self.assertEqual(DEFAULT_FILTER_MODE, FILTER_MODE_ALL)
        self.assertIn(FILTER_MODE_NUMBER, FILTER_MODE_OPTIONS)
        self.assertIn(FILTER_MODE_RT, FILTER_MODE_OPTIONS)
        self.assertIn(FILTER_MODE_NAME, FILTER_MODE_OPTIONS)

        # 1. Test parse_number_filter_query
        self.assertEqual(
            parse_number_filter_query("15, 42, 100-105"),
            {15, 42, 100, 101, 102, 103, 104, 105},
        )
        self.assertEqual(parse_number_filter_query("No. 1, 3-4"), {1, 3, 4})
        self.assertEqual(parse_number_filter_query("100 - 102; 200"), {100, 101, 102, 200})
        self.assertEqual(parse_number_filter_query("105-100"), {100, 101, 102, 103, 104, 105})
        self.assertEqual(len(parse_number_filter_query("1-50000")), 10001)
        self.assertEqual(parse_number_filter_query(""), set())

        # 2. Test parse_rt_filter_query
        self.assertEqual(parse_rt_filter_query("3"), {"3"})
        self.assertEqual(parse_rt_filter_query("03"), {"3"})
        self.assertEqual(parse_rt_filter_query("003"), {"3"})
        self.assertEqual(parse_rt_filter_query("RT 003"), {"3"})
        self.assertEqual(parse_rt_filter_query("1, 2, RT 003"), {"1", "2", "3"})
        self.assertEqual(parse_rt_filter_query(""), set())

        # 3. Test parse_name_filter_query
        self.assertEqual(
            parse_name_filter_query("ADE SETIAWAN, ADNAN"),
            ["ade setiawan", "adnan"],
        )
        self.assertEqual(parse_name_filter_query("Budi"), ["budi"])
        self.assertEqual(parse_name_filter_query(""), [])

        # 4. Test apply_voter_filter
        v1 = VoterData(
            nomor="1",
            nama="Budi Santoso",
            jenis_kelamin="Laki-laki",
            umur="25 Tahun",
            alamat="KARANGJENGKOL, RT 001 / RW 001",
            sheet_name="KADUS 1",
            wilayah="KADUS 1",
            row_index=5,
            rt_formatted="001",
        )
        v2 = VoterData(
            nomor="3",
            nama="Ade Setiawan",
            jenis_kelamin="Laki-laki",
            umur="30 Tahun",
            alamat="KARANGJENGKOL, RT 003 / RW 001",
            sheet_name="KADUS 1",
            wilayah="KADUS 1",
            row_index=6,
            rt_formatted="003",
        )
        v3 = VoterData(
            nomor="4",
            nama="Siti Adnan",
            jenis_kelamin="Perempuan",
            umur="28 Tahun",
            alamat="KARANGJENGKOL, RT 003 / RW 001",
            sheet_name="KADUS 1",
            wilayah="KADUS 1",
            row_index=7,
            rt_formatted="003",
        )
        voters = [v1, v2, v3]

        self.assertEqual(len(apply_voter_filter(voters, FILTER_MODE_ALL, "")), 3)
        self.assertEqual(len(apply_voter_filter(voters, FILTER_MODE_NUMBER, "1, 3-4")), 3)
        self.assertEqual(apply_voter_filter(voters, FILTER_MODE_NUMBER, "3"), [v2])
        self.assertEqual(apply_voter_filter(voters, FILTER_MODE_RT, "3"), [v2, v3])
        self.assertEqual(apply_voter_filter(voters, FILTER_MODE_RT, "03"), [v2, v3])
        self.assertEqual(apply_voter_filter(voters, FILTER_MODE_RT, "003"), [v2, v3])
        self.assertEqual(apply_voter_filter(voters, FILTER_MODE_RT, "RT 003"), [v2, v3])
        self.assertEqual(apply_voter_filter(voters, FILTER_MODE_RT, "1"), [v1])
        self.assertEqual(apply_voter_filter(voters, FILTER_MODE_NAME, "ade"), [v2])
        self.assertEqual(apply_voter_filter(voters, FILTER_MODE_NAME, "adnan, budi"), [v1, v3])
        self.assertEqual(apply_voter_filter(voters, FILTER_MODE_NAME, "Zorro"), [])

        # 5. Test build_filter_filename_suffix
        self.assertEqual(build_filter_filename_suffix(FILTER_MODE_ALL, ""), "")
        self.assertEqual(build_filter_filename_suffix(FILTER_MODE_NUMBER, "15, 42"), "_CETAK_NO_15_42")
        self.assertEqual(build_filter_filename_suffix(FILTER_MODE_NUMBER, "100-115"), "_CETAK_NO_100-115")
        self.assertEqual(build_filter_filename_suffix(FILTER_MODE_RT, "003"), "_RT_003")
        self.assertEqual(build_filter_filename_suffix(FILTER_MODE_NAME, "Ade"), "_FILTER_NAMA_ADE")

    def test_28_updater_version_parsing_and_release_notes(self) -> None:
        from app.config.version import (
            parse_version_tuple,
            get_latest_release_entry,
            __version__,
        )

        self.assertEqual(parse_version_tuple("v1.10.0"), (1, 10, 0))
        self.assertEqual(parse_version_tuple("1.9.0"), (1, 9, 0))
        self.assertTrue(parse_version_tuple("1.10.0") > parse_version_tuple("1.9.0"))
        self.assertTrue(parse_version_tuple("v2.0.0") > parse_version_tuple("1.10.0"))
        self.assertEqual(parse_version_tuple("1.10.0"), parse_version_tuple("v1.10.0"))
        self.assertEqual(parse_version_tuple(""), (0, 0, 0))

        latest = get_latest_release_entry()
        self.assertEqual(latest.get("version"), __version__)
        self.assertIn("highlights", latest)
        self.assertGreater(len(latest["highlights"]), 0)

    def test_29_updater_url_validation_and_security(self) -> None:
        from app.core.updater import validate_update_url

        # Valid HTTPS official URLs
        self.assertTrue(
            validate_update_url("https://github.com/Smurffahh/votermail-engine/releases/download/v1.10.0/VoterMail_Engine_v1.10.0_Windows_x64_Portable.zip")
        )
        self.assertTrue(
            validate_update_url("https://objects.githubusercontent.com/github-production-release-asset-2e65be/12345/VoterMail_Engine.zip")
        )
        self.assertTrue(
            validate_update_url("https://api.github.com/repos/Smurffahh/votermail-engine/releases/latest")
        )

        # Invalid or untrusted URLs
        self.assertFalse(validate_update_url("http://github.com/Smurffahh/votermail-engine/releases/latest"))
        self.assertFalse(validate_update_url("https://evil.com/malicious.zip"))
        self.assertFalse(validate_update_url("https://github.com/attacker/malicious-repo/releases/latest"))
        self.assertFalse(validate_update_url(""))
        self.assertFalse(validate_update_url(None))

    def test_30_safe_zip_extraction(self) -> None:
        import tempfile
        import zipfile
        from app.core.updater import safe_extract_zip

        temp_dir = Path(tempfile.mkdtemp(prefix="test_zip_"))
        try:
            valid_zip = temp_dir / "valid.zip"
            with zipfile.ZipFile(valid_zip, "w") as zf:
                zf.writestr("test_file.txt", "isi valid")
                zf.writestr("subfolder/inner.txt", "isi inner")

            target_dir = temp_dir / "extracted"
            extracted = safe_extract_zip(valid_zip, target_dir)
            self.assertEqual(len(extracted), 2)
            self.assertTrue((target_dir / "test_file.txt").exists())
            self.assertTrue((target_dir / "subfolder" / "inner.txt").exists())
            self.assertEqual((target_dir / "test_file.txt").read_text(encoding="utf-8"), "isi valid")

            # Zip Slip exploit prevention test
            slip_zip = temp_dir / "slip.zip"
            with zipfile.ZipFile(slip_zip, "w") as zf:
                zf.writestr("../evil.txt", "muatan berbahaya")

            with self.assertRaises(PermissionError):
                safe_extract_zip(slip_zip, target_dir)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def test_31_check_for_updates_mock(self) -> None:
        import json
        from unittest.mock import patch, MagicMock
        from app.core.updater import check_for_updates

        # Scenario 1: New version available
        mock_payload = {
            "tag_name": "v2.0.0",
            "name": "VoterMail Engine v2.0.0",
            "body": "- Fitur baru besar",
            "html_url": "https://github.com/Smurffahh/votermail-engine/releases/tag/v2.0.0",
            "published_at": "2026-11-20T10:00:00Z",
            "assets": [
                {
                    "name": "VoterMail_Engine_v2.0.0_Windows_x64_Portable.zip",
                    "browser_download_url": "https://github.com/Smurffahh/votermail-engine/releases/download/v2.0.0/VoterMail_Engine_v2.0.0_Windows_x64_Portable.zip",
                }
            ],
        }

        mock_resp = MagicMock()
        mock_resp.status = 200
        mock_resp.read.return_value = json.dumps(mock_payload).encode("utf-8")
        mock_resp.__enter__.return_value = mock_resp

        with patch("urllib.request.urlopen", return_value=mock_resp):
            res = check_for_updates(current_version="1.10.0")
            self.assertTrue(res["update_available"])
            self.assertEqual(res["latest_version"], "2.0.0")
            self.assertEqual(
                res["download_url"],
                "https://github.com/Smurffahh/votermail-engine/releases/download/v2.0.0/VoterMail_Engine_v2.0.0_Windows_x64_Portable.zip",
            )
            self.assertIsNone(res["error"])

        # Scenario 2: Already latest version
        mock_payload_same = {
            "tag_name": "v1.10.0",
            "name": "VoterMail Engine v1.10.0",
            "body": "Versi stabil saat ini",
            "assets": [],
        }
        mock_resp_same = MagicMock()
        mock_resp_same.status = 200
        mock_resp_same.read.return_value = json.dumps(mock_payload_same).encode("utf-8")
        mock_resp_same.__enter__.return_value = mock_resp_same

        with patch("urllib.request.urlopen", return_value=mock_resp_same):
            res_same = check_for_updates(current_version="1.10.0")
            self.assertFalse(res_same["update_available"])

        # Scenario 3: Offline / Network timeout
        import urllib.error
        with patch("urllib.request.urlopen", side_effect=urllib.error.URLError("Connection timed out")):
            res_err = check_for_updates(current_version="1.10.0")
            self.assertFalse(res_err["update_available"])
            self.assertIsNotNone(res_err["error"])

    def test_32_security_hardening_jinja2_and_excel_sanitizer(self) -> None:
        from app.core.security import validate_safe_template_tags, sanitize_cell_text

        # 1. Test validate_safe_template_tags - Safe standard tags
        safe_standard = "Surat Pemanggilan untuk {{nama_1}}, beralamat di {{alamat}}, TPS: {{tps}}."
        is_safe, msg = validate_safe_template_tags(safe_standard)
        self.assertTrue(is_safe)
        self.assertEqual(msg, "")

        self.assertTrue(validate_safe_template_tags("No: {{nomor_1}} Nama: {{nama}} RT: {{rt}}")[0])

        # 2. Test validate_safe_template_tags - Dangerous execution / injection tags
        self.assertFalse(validate_safe_template_tags("{{ ''.__class__ }}")[0])
        self.assertFalse(validate_safe_template_tags("{{ config.items() }}")[0])
        self.assertFalse(validate_safe_template_tags("{% import os %}")[0])
        self.assertFalse(validate_safe_template_tags("{% if true %} Hello {% endif %}")[0])
        self.assertFalse(validate_safe_template_tags("{{ __import__('os').system('calc') }}")[0])
        self.assertFalse(validate_safe_template_tags("{{ a[0] }}")[0])
        self.assertFalse(validate_safe_template_tags("{{ x | upper }}")[0])

        # 3. Test sanitize_cell_text - Formula error values
        for err_val in ("#REF!", "#ref!", "#VALUE!", "#value!", "#N/A", "#n/a", "#DIV/0!", "#NAME?", "#NUM!"):
            self.assertEqual(sanitize_cell_text(err_val), "")

        # 4. Test sanitize_cell_text - Max length constraint (250 chars)
        long_text = "B" * 300
        sanitized_long = sanitize_cell_text(long_text)
        self.assertEqual(len(sanitized_long), 250)
        self.assertTrue(sanitized_long.endswith("..."))
        self.assertEqual(sanitized_long[:10], "BBBBBBBBBB")

        # 5. Normal text remains intact
        self.assertEqual(sanitize_cell_text("Budi Santoso"), "Budi Santoso")

    def test_33_exception_translation_and_modern_notifications(self) -> None:
        from app.core.security import translate_exception_to_user_message
        from app.ui.components import ModernNotificationModal, InAppToastBanner
        from app.config.theme import ThemeTokens
        import customtkinter as ctk

        # 1. PermissionError translation (No raw internal path leak in summary/solution)
        perm_exc = PermissionError("[Errno 13] Permission denied: 'D:\\secret\\KADUS_1.pdf'")
        res_perm = translate_exception_to_user_message(perm_exc, context_label="Wilayah Kadus 1")
        self.assertIn("File Sedang Terbuka", res_perm["title"])
        self.assertTrue("program lain" in res_perm["summary"] or "aplikasi lain" in res_perm["summary"])
        self.assertIn("Tutup file", res_perm["solution"])
        self.assertNotIn("D:\\secret", res_perm["summary"])
        self.assertNotIn("D:\\secret", res_perm["solution"])
        self.assertIn("[PermissionError]", res_perm["technical_detail"])

        # 2. FileNotFoundError translation (docxcompose ZIP extraction guide)
        fnf_exc = FileNotFoundError("[Errno 2] No such file or directory: 'D:\\...\\_internal\\docxcompose\\templates\\custom.xml'")
        res_fnf = translate_exception_to_user_message(fnf_exc)
        self.assertIn("Komponen Template Internal", res_fnf["title"])
        self.assertIn("ZIP", res_fnf["solution"])

        # 3. pywintypes COM error translation
        com_exc = Exception("pywintypes.com_error: (-2147023170, 'The remote procedure call failed.', None, None)")
        res_com = translate_exception_to_user_message(com_exc, context_label="Surat Panggilan")
        self.assertIn("Microsoft Word", res_com["title"])
        self.assertIn("tidak merespons", res_com["summary"])

        # 4. ModernNotificationModal & InAppToastBanner headless test
        root = create_test_ctk_root()
        root.withdraw()
        try:
            theme = ThemeTokens()
            toast = InAppToastBanner(root, theme=theme)
            toast.show("Notifikasi tes berhasil", variant="success")
            self.assertEqual(toast.label_msg.cget("text"), "Notifikasi tes berhasil")
            toast.hide()

            modal_succ = ModernNotificationModal(
                parent=root,
                variant="success",
                title="Selesai Cetak",
                summary="Semua surat berhasil dibuat.",
                info_box={"Wilayah": "Kadus 1", "Total": "100 Surat"},
                theme=theme,
            )
            self.assertEqual(modal_succ.variant, "success")
            modal_succ.destroy()

            modal_err = ModernNotificationModal(
                parent=root,
                variant="error",
                title="Gagal",
                summary="Gagal memproses",
                solution="Silakan coba lagi",
                technical_detail="Err code 500",
                theme=theme,
            )
            self.assertEqual(modal_err.variant, "error")
            modal_err.destroy()
        finally:
            root.destroy()

    def test_34_daisyui_theme_filter_banner_and_structured_logs(self) -> None:
        from app.config.theme import THEME, ThemeTokens, DARK_THEME, LIGHT_THEME
        from app.ui.components import FilterStatusBanner
        from app.ui.main_window import shorten_path, format_voter_preview_table
        import customtkinter as ctk

        # 1. DaisyUI Garden & Night token verification
        self.assertEqual(THEME["bg_root"], ("#D4D2D2", "#090E1A"))
        self.assertEqual(THEME["bg_card"], ("#E9E7E7", "#0F172A"))
        self.assertEqual(THEME["btn_primary_bg"], ("#FE0075", "#38BDF8"))
        self.assertEqual(THEME["btn_neutral_bg"], ("#291E00", "#1E293B"))

        theme_dark = ThemeTokens()
        self.assertEqual(theme_dark.corner_radius_card, 16)
        self.assertEqual(theme_dark.corner_radius_inner, 8)
        self.assertEqual(theme_dark.corner_radius_control, 8)
        self.assertEqual(theme_dark.corner_radius_badge, 6)

        # 2. FilterStatusBanner states verification
        root = create_test_ctk_root()
        root.withdraw()
        try:
            banner = FilterStatusBanner(root, theme=theme_dark)
            self.assertEqual(banner.current_state, "all")

            # State 'all'
            banner.set_status("all", "Siap cetak: Semua pemilih • 150 pemilih (2 Kadus)")
            self.assertEqual(banner.current_state, "all")
            self.assertEqual(banner.label_icon.cget("text"), "ℹ")
            self.assertIn("semua pemilih", banner.label_msg.cget("text").lower())

            # State 'empty'
            banner.set_status("empty", "FILTER BELUM DIISI: Ketik kriteria filter")
            self.assertEqual(banner.current_state, "empty")
            self.assertEqual(banner.label_icon.cget("text"), "⚠")
            self.assertIn("FILTER BELUM DIISI", banner.label_msg.cget("text"))

            # State 'active'
            banner.set_status("active", "FILTER DITERAPKAN: [Nomor Urut: 1-10] • Menyaring 10 Pemilih")
            self.assertEqual(banner.current_state, "active")
            self.assertEqual(banner.label_icon.cget("text"), "✓")
            self.assertIn("FILTER DITERAPKAN", banner.label_msg.cget("text"))

            # State 'no_match'
            banner.set_status("no_match", "TIDAK ADA DATA COCOK: 0 pemilih ditemukan")
            self.assertEqual(banner.current_state, "no_match")
            self.assertEqual(banner.label_icon.cget("text"), "✕")
            self.assertIn("TIDAK ADA DATA COCOK", banner.label_msg.cget("text"))

            # Theme refresh
            banner.update_theme(LIGHT_THEME)
            self.assertEqual(banner.theme.name, "Light")
        finally:
            root.destroy()

        # 3. shorten_path verification
        cwd_file = Path.cwd() / "app" / "config" / "theme.py"
        rel_short = shorten_path(cwd_file)
        self.assertTrue("app" in rel_short and "theme.py" in rel_short)
        self.assertFalse(str(Path.cwd().resolve()) in rel_short)

        deep_file = "D:\\Pilkades2026\\DesaKarangjengkol\\Dokumen\\C1\\KADUS_01.pdf"
        short_deep = shorten_path(deep_file)
        self.assertTrue(short_deep.startswith("...\\") or short_deep.startswith(".../"))
        self.assertTrue(short_deep.endswith("KADUS_01.pdf"))

        # 4. format_voter_preview_table verification
        samples = [
            "No: 1 | Nama: ADE SETIAWAN | JK: L | Umur: 32 | Alamat: DUSUN 01 RT 001 RW 001",
            "No: 125 | Nama: MUHAMMAD ABDURRAHMAN WAHID AL-FATIH | JK: L | Umur: 45 | Alamat: DUSUN 02 RT 003 RW 002",
        ]
        table_lines = format_voter_preview_table(samples)
        self.assertTrue(any("┌" in l for l in table_lines))
        self.assertTrue(any("│ No   │" in l for l in table_lines))
        self.assertTrue(any("ADE SETIAWAN" in l for l in table_lines))
        self.assertTrue(any("MUHAMMAD ABDURRAHMAN ..." in l for l in table_lines))
        self.assertTrue(any("└" in l for l in table_lines))

    def test_35_redesigned_voter_preview_table(self) -> None:
        import customtkinter as ctk
        from app.ui.components import VoterPreviewTable
        from app.core.excel_scanner import VoterData
        from app.config.theme import LIGHT_THEME, DARK_THEME
        from app.config.settings import DEFAULT_ADDRESS_FORMAT

        root = create_test_ctk_root()
        root.withdraw()
        try:
            table = VoterPreviewTable(root, LIGHT_THEME)

            self.assertIn("Menunggu data", table.lbl_status_summary.cget("text"))
            self.assertIn("Format:", table.lbl_format_info.cget("text"))

            sample_voters = [
                VoterData(
                    nomor=str(i + 1),
                    nama=f"PEMILIH TEST {i + 1}",
                    jenis_kelamin="Laki-laki" if i % 2 == 0 else "Perempuan",
                    umur=f"{25 + i} Tahun",
                    alamat=f"RT 00{i + 1} RW 001",
                    rt=f"00{i + 1}",
                    rw="001",
                    sheet_name="KADUS 01",
                    wilayah="KADUS 01",
                    row_index=i + 10,
                )
                for i in range(5)
            ]

            table.set_voters(sample_voters, address_format=DEFAULT_ADDRESS_FORMAT, is_filtered=False, total_count=1250)
            self.assertIn("Pratinjau Sampel:", table.lbl_status_summary.cget("text"))
            self.assertIn("5 dari 1.250", table.lbl_status_summary.cget("text"))
            self.assertEqual(len(table.rows_container.winfo_children()), 5)

            header_labels = [c.cget("text") for c in table.header_row.winfo_children() if isinstance(c, ctk.CTkLabel)]
            for col_name in ["NO", "NAMA PEMILIH", "L/P", "USIA", "ALAMAT PEMILIH", "WILAYAH"]:
                self.assertIn(col_name, header_labels)

            first_row = table.rows_container.winfo_children()[0]
            row_labels = [c.cget("text") for c in first_row.winfo_children() if isinstance(c, ctk.CTkLabel)]
            self.assertIn("01", row_labels)
            self.assertIn("PEMILIH TEST 1", row_labels)
            self.assertIn("Laki-laki", row_labels)
            self.assertIn("25 Tahun", row_labels)

            badge_frame = [c for c in first_row.winfo_children() if isinstance(c, ctk.CTkFrame)][0]
            badge_label = [c for c in badge_frame.winfo_children() if isinstance(c, ctk.CTkLabel)][0]
            self.assertEqual(badge_label.cget("text"), "KADUS 01")

            table.update_address_format("RT / RW Saja (RT ... / RW ...)")
            self.assertEqual(table.cached_address_format, "RT / RW Saja (RT ... / RW ...)")
            self.assertIn("RT / RW Saja", table.lbl_format_info.cget("text"))

            table.set_voters(
                sample_voters[:2],
                address_format=DEFAULT_ADDRESS_FORMAT,
                is_filtered=True,
                total_count=1250,
                matched_count=45,
            )
            self.assertIn("Pratinjau Filter:", table.lbl_status_summary.cget("text"))
            self.assertIn("2 dari 45 Pemilih Cocok", table.lbl_status_summary.cget("text"))
            self.assertEqual(len(table.rows_container.winfo_children()), 2)

            more_voters = [
                VoterData(
                    nomor=str(i + 1),
                    nama=f"PEMILIH FILTER {i + 1}",
                    jenis_kelamin="Laki-laki" if i % 2 == 0 else "Perempuan",
                    umur=f"{30 + i} Tahun",
                    alamat=f"RT 00{i + 1} RW 002",
                    rt=f"00{i + 1}",
                    rw="002",
                    sheet_name="KADUS 02",
                    wilayah="KADUS 02",
                    row_index=i + 10,
                )
                for i in range(10)
            ]
            table.set_voters(more_voters, address_format=DEFAULT_ADDRESS_FORMAT, is_filtered=True, matched_count=10)
            self.assertEqual(len(table.rows_container.winfo_children()), 10)

            table.set_voters(
                [],
                address_format=DEFAULT_ADDRESS_FORMAT,
                is_filtered=True,
                matched_count=0,
                empty_message="Tidak ada pemilih yang cocok dengan filter tersebut.",
            )
            self.assertIn("0 Pemilih Cocok", table.lbl_status_summary.cget("text"))
            self.assertEqual(len(table.rows_container.winfo_children()), 1)

            table.update_theme(DARK_THEME)
            self.assertEqual(table.theme.name, "Dark")

            # Contextual clear log button behavior simulation
            btn_clear = ctk.CTkButton(root, text="Bersihkan")
            self.assertEqual(btn_clear.winfo_manager(), "")

            btn_clear.pack(side="left")
            self.assertEqual(btn_clear.winfo_manager(), "pack")

            btn_clear.pack_forget()
            self.assertEqual(btn_clear.winfo_manager(), "")
        finally:
            root.destroy()

    def test_36_typography_inter_and_space_mono_fonts(self) -> None:
        import customtkinter as ctk
        from app.config.theme import (
            FONT_FAMILY_UI,
            FONT_FAMILY_MONO,
            load_bundled_fonts,
            get_ui_font,
            get_mono_font,
            DARK_THEME,
            LIGHT_THEME,
        )
        from app.config.settings import FONTS_DIR

        self.assertEqual(FONT_FAMILY_UI, "Inter")
        self.assertEqual(FONT_FAMILY_MONO, "Space Mono")
        self.assertEqual(DARK_THEME.font_family_main, "Inter")
        self.assertEqual(DARK_THEME.font_family_mono, "Space Mono")
        self.assertEqual(LIGHT_THEME.font_family_main, "Inter")
        self.assertEqual(LIGHT_THEME.font_family_mono, "Space Mono")

        self.assertTrue(FONTS_DIR.exists())
        required_fonts = [
            "Inter-Regular.ttf",
            "Inter-Medium.ttf",
            "Inter-SemiBold.ttf",
            "Inter-Bold.ttf",
            "SpaceMono-Regular.ttf",
            "SpaceMono-Bold.ttf",
            "SpaceMono-Italic.ttf",
            "SpaceMono-BoldItalic.ttf",
        ]
        for font_name in required_fonts:
            font_path = FONTS_DIR / font_name
            self.assertTrue(font_path.exists(), f"Font file missing: {font_name}")
            size_kb = font_path.stat().st_size / 1024
            self.assertGreater(size_kb, 10, f"Font file {font_name} is too small: {size_kb:.2f} KB")

        root = create_test_ctk_root()
        root.withdraw()
        try:
            loaded = load_bundled_fonts()
            self.assertIsInstance(loaded, int)

            ui_font = get_ui_font(size=14, weight="bold")
            self.assertEqual(ui_font.cget("family"), "Inter")
            self.assertEqual(ui_font.cget("size"), 14)
            self.assertEqual(ui_font.cget("weight"), "bold")

            mono_font = get_mono_font(size=11)
            self.assertEqual(mono_font.cget("family"), "Space Mono")
            self.assertEqual(mono_font.cget("size"), 11)
        finally:
            root.destroy()

    def test_37_batch_files_crlf_and_ascii_clean(self) -> None:
        batch_files = ["Jalankan_VoterMail.bat", "Build_EXE_Lokal.bat"]
        for bname in batch_files:
            bpath = PROJECT_ROOT / bname
            self.assertTrue(bpath.exists(), f"Batch launcher missing: {bname}")
            raw = bpath.read_bytes()

            self.assertFalse(raw.startswith(b"\xef\xbb\xbf"), f"{bname} contains UTF-8 BOM!")
            self.assertIn(b"\r\n", raw, f"{bname} does not contain CRLF line endings!")
            bare_lf = raw.replace(b"\r\n", b"")
            self.assertNotIn(b"\n", bare_lf, f"{bname} contains bare LF line endings!")

            try:
                decoded = raw.decode("ascii")
            except UnicodeDecodeError as exc:
                self.fail(f"{bname} contains non-ASCII characters: {exc}")

            self.assertNotIn("chcp 65001", decoded, f"{bname} should not use chcp 65001")
            self.assertIn("PYTHONIOENCODING", decoded, f"{bname} should configure PYTHONIOENCODING")

        gitattr_path = PROJECT_ROOT / ".gitattributes"
        self.assertTrue(gitattr_path.exists(), ".gitattributes file is missing")
        gitattr_text = gitattr_path.read_text(encoding="utf-8")
        self.assertIn("*.bat text eol=crlf", gitattr_text)
        self.assertIn("*.cmd text eol=crlf", gitattr_text)

        # 6. Verify updater.py temporary batch script format
        from app.config.settings import TEMP_DIR
        test_source = TEMP_DIR / "test_stage"
        test_app = TEMP_DIR / "test_app"
        test_exe = test_app / "VoterMail_Engine.exe"
        bat_lines = [
            "@echo off",
            "setlocal",
            "echo Memasang pembaruan VoterMail Engine...",
            "ping 127.0.0.1 -n 3 >nul",
            f'if exist "{test_source}\\_internal" (',
            f'    xcopy /E /Y /I "{test_source}\\_internal" "{test_app}\\_internal" >nul',
            ")",
            f'if exist "{test_source}\\VoterMail_Engine.exe" (',
            f'    copy /Y "{test_source}\\VoterMail_Engine.exe" "{test_app}\\VoterMail_Engine.exe" >nul',
            ")",
            f'if exist "{test_source}\\PANDUAN_PANITIA.txt" (',
            f'    copy /Y "{test_source}\\PANDUAN_PANITIA.txt" "{test_app}\\PANDUAN_PANITIA.txt" >nul',
            ")",
            f'rmdir /S /Q "{test_source}" >nul',
            f'start "" "{test_exe}"',
            'del "%~f0"',
            "",
        ]
        sample_bat = "\r\n".join(bat_lines).encode("ascii")
        self.assertIn(b"\r\n", sample_bat)
        self.assertNotIn(b"\n", sample_bat.replace(b"\r\n", b""))
        self.assertNotIn(b"chcp 65001", sample_bat)

    def test_38_ui_modularization_modals_and_comment_hygiene(self) -> None:
        from app.config.version import __version__, FEATURE_REGISTRY, __author__
        from app.config.settings import FONTS_DIR
        import app.ui.components as comp
        import app.ui.modals as modals

        self.assertEqual(__version__, "1.12.1")
        self.assertEqual(FEATURE_REGISTRY[-1]["version"], "1.12.1")
        self.assertEqual(__author__, "Danang Satria")

        # Verify modals and helpers exist in modals.py and re-exported in components.py
        for attr in ["AboutModal", "ColumnMapperModal", "UpdateNotificationModal", "ModernNotificationModal", "get_col_letter", "apply_window_icon"]:
            self.assertTrue(hasattr(modals, attr), f"Missing {attr} in app.ui.modals")
            self.assertTrue(hasattr(comp, attr), f"Missing {attr} in app.ui.components re-export")
            self.assertIs(getattr(modals, attr), getattr(comp, attr), f"{attr} mismatch between modules")

        # Verify author credits preservation
        self.assertEqual(getattr(modals, "__author__", ""), "Danang Satria")
        self.assertEqual(getattr(comp, "__author__", ""), "Danang Satria")

        # Verify font asset hygiene (variable font eliminated, static fonts preserved)
        redundant_var_font = FONTS_DIR / "Inter-VariableFont_opsz,wght.ttf"
        self.assertFalse(redundant_var_font.exists(), "Inter-VariableFont_opsz,wght.ttf should be deleted to save 876 KB")

        expected_static_fonts = [
            "Inter-Regular.ttf",
            "Inter-Medium.ttf",
            "Inter-SemiBold.ttf",
            "Inter-Bold.ttf",
            "SpaceMono-Regular.ttf",
            "SpaceMono-Bold.ttf",
            "SpaceMono-Italic.ttf",
            "SpaceMono-BoldItalic.ttf",
        ]
        for fname in expected_static_fonts:
            self.assertTrue((FONTS_DIR / fname).exists(), f"Static font asset missing: {fname}")

    def test_39_monotonic_semver_guardrails_and_updater_compatibility(self) -> None:
        from unittest.mock import patch, MagicMock
        import json
        from app.config.version import (
            __version__,
            FEATURE_REGISTRY,
            validate_version_registry,
            parse_version_tuple,
        )
        from app.core.updater import check_for_updates

        # 1. Full Monotonic Guardrail verification on active registry
        self.assertTrue(validate_version_registry())

        # 2. Negative Scenario: Duplicate versions in registry
        dup_reg = [
            {"version": "1.0.0", "title": "A", "description": "A"},
            {"version": "1.0.0", "title": "B", "description": "B"},
        ]
        with self.assertRaises(ValueError) as ctx_dup:
            validate_version_registry(dup_reg, current_version="1.0.0")
        self.assertIn("duplikasi versi", str(ctx_dup.exception))

        # 3. Negative Scenario: Regression in versioning (e.g. 1.10.0 -> 1.9.1)
        regress_reg = [
            {"version": "1.9.0", "title": "A", "description": "A"},
            {"version": "1.10.0", "title": "B", "description": "B"},
            {"version": "1.9.1", "title": "C", "description": "C"},
        ]
        with self.assertRaises(ValueError) as ctx_regress:
            validate_version_registry(regress_reg, current_version="1.9.1")
        self.assertIn("Regresi versi terdeteksi", str(ctx_regress.exception))

        # 4. Negative Scenario: Invalid SemVer format
        bad_format_reg = [
            {"version": "1.10", "title": "A", "description": "A"},
        ]
        with self.assertRaises(ValueError) as ctx_fmt:
            validate_version_registry(bad_format_reg, current_version="1.10")
        self.assertIn("Format versi tidak valid", str(ctx_fmt.exception))

        # 5. Negative Scenario: Active version mismatch with last registry entry
        with self.assertRaises(ValueError) as ctx_mismatch:
            validate_version_registry(FEATURE_REGISTRY, current_version="9.9.9")
        self.assertIn("tidak cocok", str(ctx_mismatch.exception))

        # 6. Auto-Updater compatibility: Old versions (1.9.4, 1.10.0, 1.10.4, 1.11.0, 1.11.1, 1.11.2, 1.12.0) detect 1.12.1
        mock_payload_1_12_1 = {
            "tag_name": "v1.12.1",
            "name": "VoterMail Engine v1.12.1",
            "body": "Rilis Resmi v1.12.1",
            "assets": [
                {
                    "name": "VoterMail_Engine_v1.12.1_Windows_x64_Portable.zip",
                    "browser_download_url": "https://github.com/Smurffahh/votermail-engine/releases/download/v1.12.1/VoterMail_Engine_v1.12.1_Windows_x64_Portable.zip",
                }
            ],
            "html_url": "https://github.com/Smurffahh/votermail-engine/releases/tag/v1.12.1",
        }
        mock_resp = MagicMock()
        mock_resp.status = 200
        mock_resp.read.return_value = json.dumps(mock_payload_1_12_1).encode("utf-8")
        mock_resp.__enter__.return_value = mock_resp

        with patch("urllib.request.urlopen", return_value=mock_resp):
            # User from v1.9.4
            res_194 = check_for_updates(current_version="1.9.4")
            self.assertTrue(res_194["update_available"])
            self.assertEqual(res_194["latest_version"], "1.12.1")

            # User from v1.10.0
            res_1100 = check_for_updates(current_version="1.10.0")
            self.assertTrue(res_1100["update_available"])
            self.assertEqual(res_1100["latest_version"], "1.12.1")

            # User from v1.10.4
            res_1104 = check_for_updates(current_version="1.10.4")
            self.assertTrue(res_1104["update_available"])
            self.assertEqual(res_1104["latest_version"], "1.12.1")

            # User from v1.11.0
            res_1110 = check_for_updates(current_version="1.11.0")
            self.assertTrue(res_1110["update_available"])
            self.assertEqual(res_1110["latest_version"], "1.12.1")

            # User from v1.11.1
            res_1111 = check_for_updates(current_version="1.11.1")
            self.assertTrue(res_1111["update_available"])
            self.assertEqual(res_1111["latest_version"], "1.12.1")

            # User from v1.11.2
            res_1112 = check_for_updates(current_version="1.11.2")
            self.assertTrue(res_1112["update_available"])
            self.assertEqual(res_1112["latest_version"], "1.12.1")

            # User from v1.12.0
            res_1120 = check_for_updates(current_version="1.12.0")
            self.assertTrue(res_1120["update_available"])
            self.assertEqual(res_1120["latest_version"], "1.12.1")

            # User already on v1.12.1
            res_1121 = check_for_updates(current_version="1.12.1")
            self.assertFalse(res_1121["update_available"])

            # User on future build v1.13.0
            res_future = check_for_updates(current_version="1.13.0")
            self.assertFalse(res_future["update_available"])

    def test_40_ooxml_shape_sanitization_and_no_ghost_text(self) -> None:
        import zipfile
        import tempfile
        from lxml import etree
        from app.core.mail_merger import sanitize_merged_ooxml_shapes, generate_preview_via_word
        from app.core.word_preparer import check_and_prepare_word_template
        from app.config.settings import get_default_word_path, OUTPUT_DIR

        # 1. Test OOXML sanitization helper with duplicate shapes, AlternateContent, and empty ghost shapes
        test_xml = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            b'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            b'xmlns:v="urn:schemas-microsoft-com:vml" '
            b'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
            b'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">\n'
            b'<w:body>\n'
            b'  <w:p><w:r><w:pict>'
            b'    <v:shape id="_x0000_s1026" type="#_x0000_t202" filled="f" stroked="f">'
            b'      <v:textbox><w:txbxContent><w:p><w:r><w:t></w:t></w:r></w:p></w:txbxContent></v:textbox>'
            b'    </v:shape>'
            b'  </w:pict></w:r></w:p>\n'
            b'  <w:p><w:r><w:pict>'
            b'    <v:shape id="_x0000_s1026" type="#_x0000_t6" filled="t" stroked="f">'
            b'      <v:textpath string="PANITIA PEMILIHAN"/>'
            b'    </v:shape>'
            b'  </w:pict></w:r></w:p>\n'
            b'  <w:p><w:r><w:pict>'
            b'    <v:shape id="_x0000_s1026" type="#_x0000_t6" filled="t" stroked="f">'
            b'      <v:textpath string="PANITIA PEMILIHAN"/>'
            b'    </v:shape>'
            b'  </w:pict></w:r></w:p>\n'
            b'  <w:p><w:r><w:drawing><wp:docPr id="1" name="Shape 1"/></w:drawing></w:r></w:p>\n'
            b'  <w:p><w:r><w:drawing><wp:docPr id="1" name="Shape 1"/></w:drawing></w:r></w:p>\n'
            b'  <w:p><w:r><mc:AlternateContent>'
            b'    <mc:Choice Requires="w14"><w:drawing><wps:cNvPr id="2"/><w:t>PANITIA PEMILIHAN</w:t></w:drawing></mc:Choice>'
            b'    <mc:Fallback><w:pict><w:t>PANITIA PEMILIHAN</w:t></w:pict></mc:Fallback>'
            b'  </mc:AlternateContent></w:r></w:p>\n'
            b'</w:body></w:document>'
        )

        with tempfile.TemporaryDirectory() as td:
            dummy_docx = Path(td) / "test_ooxml.docx"
            with zipfile.ZipFile(dummy_docx, "w") as z:
                z.writestr("word/document.xml", test_xml)

            sanitize_merged_ooxml_shapes(str(dummy_docx))

            with zipfile.ZipFile(dummy_docx, "r") as z:
                res_xml = z.read("word/document.xml")
                root = etree.fromstring(res_xml)
                ns = {
                    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
                    "v": "urn:schemas-microsoft-com:vml",
                    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                }
                # Check root mc:Ignorable contains w14 and wp14
                ign_val = root.attrib.get("{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable", "")
                self.assertIn("w14", ign_val)
                self.assertIn("wp14", ign_val)

                # Check empty ghost shape was removed
                shapes = root.xpath("//v:shape", namespaces=ns)
                self.assertEqual(len(shapes), 2)
                shape_ids = [s.attrib.get("id") for s in shapes]
                self.assertEqual(len(shape_ids), len(set(shape_ids)))

                # Check docPr IDs are unique
                docprs = root.xpath("//wp:docPr", namespaces=ns)
                self.assertEqual(len(docprs), 2)
                docpr_ids = [d.attrib.get("id") for d in docprs]
                self.assertEqual(len(docpr_ids), len(set(docpr_ids)))

                # Check fallback duplicate text was neutralized
                fallback_t = root.xpath("//mc:AlternateContent/mc:Fallback//w:t", namespaces=ns)
                for t in fallback_t:
                    self.assertEqual(t.text or "", "")

        # 2. Check that template preparation does not inject synthetic horizontal textboxes
        def_word = get_default_word_path()
        if os.path.exists(def_word) and is_ms_word_installed():
            word_res = check_and_prepare_word_template(str(def_word))
            self.assertTrue(word_res.is_valid)
            with zipfile.ZipFile(word_res.prepared_docx_path, "r") as z:
                xml_content = z.read("word/document.xml").decode("utf-8")
                # There should be exactly 4 occurrences of "PANITIA PEMILIHAN" in v:textpath
                self.assertEqual(xml_content.count("PANITIA PEMILIHAN"), 4)

            # 3. Verify PDF preview has no duplicate horizontal ghost text
            preview_pdf = str(OUTPUT_DIR / "test_no_ghost_preview.pdf")
            ok, msg, final_pdf = generate_preview_via_word(
                template_docx_path=word_res.prepared_docx_path,
                sample_voters=[],
                output_pdf_path=preview_pdf,
            )
            self.assertTrue(ok, f"Preview failed: {msg}")
            if os.path.exists(final_pdf):
                try:
                    import pypdf
                    reader = pypdf.PdfReader(final_pdf)
                    text = reader.pages[0].extract_text()
                    self.assertEqual(text.count("Panitia Pemilihan Kepala Desa Karangjengkol"), 2)
                finally:
                    try:
                        os.unlink(final_pdf)
                    except Exception:
                        pass

    def test_41_format_int_id_and_layout_proportions(self) -> None:
        from app.config.settings import format_int_id
        from app.ui.components import CleanMetricBar, VoterPreviewTable, PREVIEW_COLUMNS_SPEC
        from app.config.theme import get_theme
        from app.core.excel_scanner import VoterData

        # 1. format_int_id verification
        self.assertEqual(format_int_id(None), "0")
        self.assertEqual(format_int_id(0), "0")
        self.assertEqual(format_int_id(5), "5")
        self.assertEqual(format_int_id(3505), "3.505")
        self.assertEqual(format_int_id(1129), "1.129")
        self.assertEqual(format_int_id(2257), "2.257")
        self.assertEqual(format_int_id("1000000"), "1.000.000")
        self.assertEqual(format_int_id("invalid"), "invalid")

        # 2. UI layout proportions & table columns spec
        self.assertEqual(len(PREVIEW_COLUMNS_SPEC), 6)
        col_names = [spec[0] for spec in PREVIEW_COLUMNS_SPEC]
        self.assertEqual(col_names, [0, 1, 2, 3, 4, 5])

        theme = get_theme("Dark")
        root = create_test_ctk_root()
        root.withdraw()
        try:
            bar = CleanMetricBar(root, theme=theme)
            self.assertEqual(bar.cget("height"), 58)
            bar.update_metrics(5, 3505, 1753)
            self.assertEqual(bar.lbl_v1.cget("text"), "5 Kadus")
            self.assertEqual(bar.lbl_v2.cget("text"), "3.505 Orang")
            self.assertEqual(bar.lbl_v3.cget("text"), "1.753 Lembar F4")

            table = VoterPreviewTable(root, theme=theme)
            self.assertEqual(table.rows_container.cget("corner_radius"), 0)
            self.assertEqual(table.rows_container.cget("border_width"), 0)

            sample_voters = [
                VoterData(nomor="1", nama="Budi Santoso", jenis_kelamin="L", umur="35", alamat="RT 01 RW 01", sheet_name="KADUS 1", wilayah="KADUS 1", row_index=2, extra_fields={}),
                VoterData(nomor="2", nama="Siti Rahayu", jenis_kelamin="P", umur="28", alamat="RT 02 RW 01", sheet_name="KADUS 1", wilayah="KADUS 1", row_index=3, extra_fields={}),
            ]
            table.set_voters(sample_voters, is_filtered=False, total_count=3505)
            self.assertIn("3.505", table.lbl_status_summary.cget("text"))

            table.set_voters(sample_voters[:1], is_filtered=True, matched_count=1129)
            self.assertIn("1.129", table.lbl_status_summary.cget("text"))
        finally:
            root.destroy()

    def test_42_smart_in_place_patch_updater_and_integrity_verification(self) -> None:
        import hashlib
        import json
        import subprocess
        import tempfile
        import zipfile
        from unittest.mock import patch, MagicMock
        from app.core.updater import (
            compute_deps_hash,
            compute_file_sha256,
            format_file_size,
            check_for_updates,
            stage_update_package,
            generate_patch_script,
            launch_detached_patcher,
        )
        from app.ui.main_window import MainWindow

        # 1. Dependency Hash computation invariance
        req_1 = "customtkinter==5.2.2\npython-docx==1.1.2\n# Comment\n\npypdf==5.1.0\n"
        req_2 = "# Header comment\npypdf==5.1.0\n\npython-docx==1.1.2\ncustomtkinter==5.2.2\n"
        req_3 = "customtkinter==5.2.2\npython-docx==1.2.0\npypdf==5.1.0\n"
        h1 = compute_deps_hash(req_1)
        h2 = compute_deps_hash(req_2)
        h3 = compute_deps_hash(req_3)
        self.assertEqual(h1, h2)
        self.assertNotEqual(h1, h3)

        # 2. File size formatting with Indonesian comma
        self.assertEqual(format_file_size(4_390_000), "4,2 MB")
        self.assertEqual(format_file_size(33_000_000), "31,5 MB")
        self.assertEqual(format_file_size(850_000), "830 KB")
        self.assertEqual(format_file_size(500), "500 B")

        # 3. Check for updates: Patch vs Full resolution via update_manifest.json
        manifest_data = {
            "version": "1.12.0",
            "deps_hash": h1,
            "assets": {
                "full_zip": {
                    "name": "VoterMail_Engine_v1.12.0_Windows_x64_Portable.zip",
                    "size": 33_000_000,
                    "sha256": "aaaa1111",
                },
                "patch_zip": {
                    "name": "VoterMail_Engine_v1.12.0_Windows_x64_Patch.zip",
                    "size": 4_390_000,
                    "sha256": "bbbb2222",
                },
            },
        }

        mock_release = {
            "tag_name": "v1.12.0",
            "name": "VoterMail Engine v1.12.0",
            "body": "Notes v1.12.0",
            "assets": [
                {
                    "name": "update_manifest.json",
                    "browser_download_url": "https://github.com/Smurffahh/votermail-engine/releases/download/v1.12.0/update_manifest.json",
                },
                {
                    "name": "VoterMail_Engine_v1.12.0_Windows_x64_Portable.zip",
                    "browser_download_url": "https://github.com/Smurffahh/votermail-engine/releases/download/v1.12.0/VoterMail_Engine_v1.12.0_Windows_x64_Portable.zip",
                    "size": 33_000_000,
                },
                {
                    "name": "VoterMail_Engine_v1.12.0_Windows_x64_Patch.zip",
                    "browser_download_url": "https://github.com/Smurffahh/votermail-engine/releases/download/v1.12.0/VoterMail_Engine_v1.12.0_Windows_x64_Patch.zip",
                    "size": 4_390_000,
                },
            ],
            "html_url": "https://github.com/Smurffahh/votermail-engine/releases/tag/v1.12.0",
        }

        def mock_urlopen_router(req, *args, **kwargs):
            url = req.full_url if hasattr(req, "full_url") else str(req)
            resp = MagicMock()
            resp.status = 200
            if "update_manifest.json" in url:
                resp.read.return_value = json.dumps(manifest_data).encode("utf-8")
            else:
                resp.read.return_value = json.dumps(mock_release).encode("utf-8")
            resp.__enter__.return_value = resp
            return resp

        # Case A: Dependencies match -> Resolves PATCH mode
        with patch("urllib.request.urlopen", side_effect=mock_urlopen_router), \
             patch("app.core.updater.get_local_build_meta", return_value={"deps_hash": h1}):
            info_patch = check_for_updates(current_version="1.11.2")
            self.assertTrue(info_patch["update_available"])
            self.assertEqual(info_patch["update_mode"], "PATCH")
            self.assertIn("Patch.zip", info_patch["download_url"])
            self.assertEqual(info_patch["expected_sha256"], "bbbb2222")
            self.assertEqual(info_patch["download_size_bytes"], 4_390_000)

        # Case B: Dependencies changed -> Resolves FULL mode
        with patch("urllib.request.urlopen", side_effect=mock_urlopen_router), \
             patch("app.core.updater.get_local_build_meta", return_value={"deps_hash": "different_hash"}):
            info_full = check_for_updates(current_version="1.11.2")
            self.assertTrue(info_full["update_available"])
            self.assertEqual(info_full["update_mode"], "FULL")
            self.assertIn("Portable.zip", info_full["download_url"])
            self.assertEqual(info_full["expected_sha256"], "aaaa1111")
            self.assertEqual(info_full["download_size_bytes"], 33_000_000)

        # 4. stage_update_package SHA-256 verification and anti-Zip-Slip
        with tempfile.TemporaryDirectory() as td:
            pkg_zip_path = Path(td) / "test_pkg.zip"
            with zipfile.ZipFile(pkg_zip_path, "w") as z:
                z.writestr("VoterMail_Engine.exe", b"MZ_MOCK_EXE")
                z.writestr("app/assets/build_meta.json", b'{"version": "1.12.0"}')

            pkg_bytes = pkg_zip_path.read_bytes()
            real_sha = hashlib.sha256(pkg_bytes).hexdigest()

            def mock_dl_resp(req, *args, **kwargs):
                resp = MagicMock()
                resp.status = 200
                resp.headers = {"Content-Length": str(len(pkg_bytes))}
                resp.read.side_effect = [pkg_bytes, b""]
                resp.__enter__.return_value = resp
                return resp

            # Subcase: Corrupted bytes -> Fails integrity
            with patch("urllib.request.urlopen", side_effect=mock_dl_resp):
                ok_corrupt, msg_corrupt, _, _ = stage_update_package(
                    download_url="https://github.com/Smurffahh/votermail-engine/releases/download/v1.12.0/test_pkg.zip",
                    target_version="1.12.0_test_bad",
                    expected_sha256="wrong_checksum_000000000000000000000000000000000000000000000000000000",
                    update_mode="PATCH",
                )
                self.assertFalse(ok_corrupt)
                self.assertIn("integritas", msg_corrupt.lower())
                self.assertIn("SHA-256", msg_corrupt)

            # Subcase: Valid SHA-256 -> Success
            with patch("urllib.request.urlopen", side_effect=mock_dl_resp):
                ok_valid, msg_valid, s_dir, src_dir = stage_update_package(
                    download_url="https://github.com/Smurffahh/votermail-engine/releases/download/v1.12.0/test_pkg.zip",
                    target_version="1.12.0_test_good",
                    expected_sha256=real_sha,
                    update_mode="PATCH",
                )
                self.assertTrue(ok_valid, f"Staging failed: {msg_valid}")
                self.assertIsNotNone(s_dir)
                self.assertTrue((src_dir / "VoterMail_Engine.exe").exists())

        # 5. PowerShell script generation: pure ASCII, CRLF, and AST parse validity
        with tempfile.TemporaryDirectory() as td:
            ps1_path = generate_patch_script(td)
            self.assertTrue(ps1_path.exists())
            raw_bytes = ps1_path.read_bytes()
            self.assertTrue(all(b < 128 for b in raw_bytes))
            self.assertIn(b"\r\n", raw_bytes)
            self.assertNotIn(b"\n", raw_bytes.replace(b"\r\n", b""))

            ps_cmd = [
                "powershell.exe",
                "-NoProfile",
                "-Command",
                f"$errs = $null; [System.Management.Automation.Language.Parser]::ParseFile('{str(ps1_path)}', [ref]$null, [ref]$errs); if ($errs.Count -gt 0) {{ $errs | ForEach-Object {{ Write-Error $_.Message }}; exit 1 }} else {{ exit 0 }}",
            ]
            res_ps = subprocess.run(ps_cmd, capture_output=True, text=True)
            self.assertEqual(res_ps.returncode, 0, f"PowerShell syntax error: {res_ps.stderr}")

        # 6. Verify MainWindow shutdown methods exist
        self.assertTrue(hasattr(MainWindow, "_prepare_shutdown_for_update"))
        self.assertTrue(hasattr(MainWindow, "_shutdown_and_exit_for_update"))

    def test_43_cross_drive_flashdisk_updater_hardening(self) -> None:
        import json
        import subprocess
        import tempfile
        from unittest.mock import patch
        from app.core.updater import generate_patch_script, launch_detached_patcher

        with tempfile.TemporaryDirectory() as td:
            ps1_path = generate_patch_script(td)
            self.assertTrue(ps1_path.exists())
            ps1_text = ps1_path.read_text(encoding="ascii")

            # 1. Assert Copy-Item with -LiteralPath and -Force on all 3 phases (Backup, Apply, Rollback)
            self.assertIn("Copy-Item -LiteralPath $targetFile -Destination $backupTarget -Force", ps1_text)
            self.assertIn("Copy-Item -LiteralPath $f.FullName -Destination $targetFile -Force", ps1_text)
            self.assertIn("Copy-Item -LiteralPath $bSrc -Destination $tDst -Force", ps1_text)

            # 2. Assert Move-Item is NOT used anywhere
            self.assertNotIn("Move-Item", ps1_text)

            # 3. Assert ReadOnly clearing logic exists for target file and rollback destination
            self.assertIn("(Get-Item -LiteralPath $targetFile -Force).IsReadOnly = $false", ps1_text)
            self.assertIn("(Get-Item -LiteralPath $tDst -Force).IsReadOnly = $false", ps1_text)

            # 4. Assert friendly storage error message for write-protected or full storage
            expected_msg = "Pembaruan dibatalkan karena media penyimpanan penuh atau terkunci (write-protected). Versi sebelumnya telah dipulihkan."
            self.assertIn(expected_msg, ps1_text)

            # 5. Assert pure ASCII and CRLF line endings
            raw_bytes = ps1_path.read_bytes()
            self.assertTrue(all(b < 128 for b in raw_bytes))
            self.assertIn(b"\r\n", raw_bytes)
            self.assertNotIn(b"\n", raw_bytes.replace(b"\r\n", b""))

            # 6. Validate PowerShell syntax via AST parser
            ps_cmd = [
                "powershell.exe",
                "-NoProfile",
                "-Command",
                f"$errs = $null; [System.Management.Automation.Language.Parser]::ParseFile('{str(ps1_path)}', [ref]$null, [ref]$errs); if ($errs.Count -gt 0) {{ $errs | ForEach-Object {{ Write-Error $_.Message }}; exit 1 }} else {{ exit 0 }}",
            ]
            res_ps = subprocess.run(ps_cmd, capture_output=True, text=True)
            self.assertEqual(res_ps.returncode, 0, f"PowerShell AST syntax error: {res_ps.stderr}")

            # 7. Simulation of bracketed path on flash drive (e.g. F:\Pilkades [2026]\VoterMail_Engine)
            bracket_app_dir = r"F:\Pilkades [2026]\VoterMail_Engine"
            bracket_exe = r"F:\Pilkades [2026]\VoterMail_Engine\VoterMail_Engine.exe"
            staging_d = Path(td) / "staging"
            staging_d.mkdir(parents=True, exist_ok=True)
            backup_d = Path(td) / "backup"

            with patch("subprocess.Popen") as mock_popen:
                ok, msg = launch_detached_patcher(
                    parent_pid=9999,
                    app_dir=bracket_app_dir,
                    exe_path=bracket_exe,
                    staging_dir=staging_d,
                    backup_dir=backup_d,
                    update_mode="PATCH",
                    target_version="1.12.1",
                )
                self.assertTrue(ok, f"Launch failed: {msg}")
                self.assertTrue(mock_popen.called)

            job_json = Path(td) / "patch_job.json"
            self.assertTrue(job_json.exists())
            job_data = json.loads(job_json.read_text(encoding="utf-8"))
            self.assertIn("[2026]", job_data["app_dir"])
            self.assertIn("[2026]", job_data["exe_path"])

            # Verify PowerShell ConvertFrom-Json parses the bracketed path correctly
            ps_json_test = [
                "powershell.exe",
                "-NoProfile",
                "-Command",
                f"$raw = [System.IO.File]::ReadAllText('{str(job_json)}', [System.Text.Encoding]::UTF8); $j = $raw | ConvertFrom-Json; if ($j.app_dir -like '*[2026]*') {{ exit 0 }} else {{ exit 1 }}",
            ]
            res_json = subprocess.run(ps_json_test, capture_output=True, text=True)
            self.assertEqual(res_json.returncode, 0, "PowerShell failed to parse bracketed path JSON")

if __name__ == "__main__":
    unittest.main()
