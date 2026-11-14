# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import functools
import os
from pathlib import Path
import shutil
import sys
import unittest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

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
        cls.ready_docx_file = str(ready_docx_candidate.resolve()) if ready_docx_candidate.exists() else ""

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
        self.assertIn("SURAT PEMBERITAHUAN UNTUK MEMBERIKAN SUARA", text)
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
        )
        self.assertEqual(__app_name__, "VoterMail Engine")
        self.assertEqual(__author__, "Danang Satria")
        self.assertEqual(__github_user__, "Smurffahh")
        self.assertTrue(__repository__.startswith("https://github.com/Smurffahh"))
        self.assertEqual(__version__, "1.8.0")
        self.assertEqual(get_current_version(), "1.8.0")
        self.assertGreaterEqual(len(FEATURE_REGISTRY), 9)
        self.assertEqual(FEATURE_REGISTRY[-1]["version"], "1.7.1")

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
        self.assertEqual(DARK_THEME.primary_action, ("#10B981", "#10B981"))
        self.assertEqual(LIGHT_THEME.primary_action, ("#10B981", "#10B981"))
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

if __name__ == "__main__":
    unittest.main()
