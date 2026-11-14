# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import math
import os
from pathlib import Path
import shutil
import tempfile
import threading
import time
from typing import Callable
import docx
from docxcompose.composer import Composer
from docxtpl import DocxTemplate
from pypdf import PdfReader, PdfWriter
import pythoncom
import win32com.client

from app.config.settings import TEMP_DIR
from app.config.version import __version__, __author__, __repository__
from app.core.excel_scanner import VoterData
from app.core.security import check_file_accessible_for_writing, safe_delete_file, to_win_abs_path

def apply_pdf_metadata(pdf_path: str, kadus_name: str = "") -> None:
    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        title_str = f"Surat Pemberitahuan Pemilih - {kadus_name}" if kadus_name else "Surat Pemberitahuan Pemilih"
        metadata = {
            "/Title": title_str,
            "/Author": f"{__author__} (VoterMail Engine)",
            "/Creator": f"VoterMail Engine v{__version__} — Created by {__author__}",
            "/Producer": f"VoterMail Engine ({__repository__})",
        }
        writer.add_metadata(metadata)

        temp_pdf = f"{pdf_path}.meta.tmp"
        with open(temp_pdf, "wb") as f:
            writer.write(f)

        shutil.move(temp_pdf, pdf_path)
    except Exception:
        pass

def export_docx_to_pdf_via_word(docx_path: str, pdf_path: str) -> tuple[bool, str, str]:
    abs_docx = to_win_abs_path(docx_path)
    abs_pdf = to_win_abs_path(pdf_path)

    if not os.path.exists(abs_docx):
        return False, f"File docx tidak ditemukan: {abs_docx}", abs_pdf

    Path(abs_pdf).parent.mkdir(parents=True, exist_ok=True)

    accessible, lock_err = check_file_accessible_for_writing(abs_pdf)
    if not accessible:
        stem = Path(abs_pdf).stem
        ext = Path(abs_pdf).suffix
        ts = int(time.time())
        fallback_pdf = to_win_abs_path(Path(abs_pdf).parent / f"{stem}_{ts}{ext}")
        fallback_ok, _ = check_file_accessible_for_writing(fallback_pdf)
        if fallback_ok:
            abs_pdf = fallback_pdf
        else:
            return False, lock_err, abs_pdf

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(abs_docx, False, True, False)

        try:
            doc.Repaginate()
        except Exception:
            pass

        exported = False
        last_error = ""

        for attempt in range(3):
            # Tier 1: Positional ExportAsFixedFormat
            try:
                doc.ExportAsFixedFormat(abs_pdf, 17)
                exported = True
                break
            except Exception as e1:
                last_error = f"Tier 1 (ExportAsFixedFormat): {e1}"

            # Tier 2: SaveAs2 with FileFormat=17
            try:
                doc.SaveAs2(abs_pdf, 17)
                exported = True
                break
            except Exception as e2:
                last_error += f" | Tier 2 (SaveAs2): {e2}"

            # Tier 3: SaveAs with FileFormat=17
            try:
                doc.SaveAs(abs_pdf, 17)
                exported = True
                break
            except Exception as e3:
                last_error += f" | Tier 3 (SaveAs): {e3}"

            if attempt < 2:
                time.sleep(0.5)

        if not exported:
            return False, f"Gagal mengekspor PDF melalui MS Word: {last_error}", abs_pdf

        return True, "Berhasil", abs_pdf
    except Exception as exc:
        return False, f"Word COM error: {exc}", abs_pdf
    finally:
        if doc is not None:
            try:
                doc.Close(0)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

def merge_voters_to_pdf_via_word_com(
    template_docx_path: str,
    voters: list[VoterData],
    sheet_name: str,
    wilayah_name: str,
    output_pdf_path: str,
    progress_callback: Callable[[int, int, str], None] | None = None,
    cancel_event: threading.Event | None = None,
    layout_mode: str = "2-UP",
) -> tuple[bool, int, str]:
    if not voters:
        return False, 0, "Daftar pemilih kosong."

    abs_template = to_win_abs_path(template_docx_path)
    abs_target_pdf = to_win_abs_path(output_pdf_path)
    Path(abs_target_pdf).parent.mkdir(parents=True, exist_ok=True)

    accessible, lock_err = check_file_accessible_for_writing(abs_target_pdf)
    if not accessible:
        return False, 0, lock_err

    temp_dir = tempfile.mkdtemp(prefix="mailmerge_work_", dir=str(TEMP_DIR))
    total_voters = len(voters)
    pages_count = total_voters if layout_mode == "1-UP" else math.ceil(total_voters / 2)

    try:
        page_doc_files: list[str] = []

        if layout_mode == "1-UP":
            for p_idx, v1 in enumerate(voters):
                if cancel_event is not None and cancel_event.is_set():
                    return False, 0, "Proses mailing dibatalkan pengguna."

                ctx = {
                    "nomor_1": v1.nomor,
                    "nama_1": v1.nama,
                    "jk_1": v1.jenis_kelamin,
                    "umur_1": v1.umur,
                    "alamat_1": v1.alamat,
                    "nomor": v1.nomor,
                    "nama": v1.nama,
                    "jk": v1.jenis_kelamin,
                    "umur": v1.umur,
                    "alamat": v1.alamat,
                }
                for k, v in v1.extra_fields.items():
                    ctx[k] = v
                    ctx[f"{k}_1"] = v

                tpl = DocxTemplate(abs_template)
                tpl.render(ctx)
                page_path = os.path.join(temp_dir, f"page_{p_idx}.docx")
                tpl.save(page_path)
                page_doc_files.append(page_path)

                if progress_callback:
                    progress_callback(p_idx + 1, total_voters, sheet_name)
        else:
            pairs: list[tuple[VoterData, VoterData | None]] = []
            for i in range(0, total_voters, 2):
                v1 = voters[i]
                v2 = voters[i + 1] if (i + 1) < total_voters else None
                pairs.append((v1, v2))

            for p_idx, (v1, v2) in enumerate(pairs):
                if cancel_event is not None and cancel_event.is_set():
                    return False, 0, "Proses mailing dibatalkan pengguna."

                ctx = {
                    "nomor_1": v1.nomor,
                    "nama_1": v1.nama,
                    "jk_1": v1.jenis_kelamin,
                    "umur_1": v1.umur,
                    "alamat_1": v1.alamat,
                    "nomor_2": v2.nomor if v2 else "",
                    "nama_2": v2.nama if v2 else "",
                    "jk_2": v2.jenis_kelamin if v2 else "",
                    "umur_2": v2.umur if v2 else "",
                    "alamat_2": v2.alamat if v2 else "",
                }
                for k, v in v1.extra_fields.items():
                    ctx[f"{k}_1"] = v
                if v2:
                    for k, v in v2.extra_fields.items():
                        ctx[f"{k}_2"] = v

                tpl = DocxTemplate(abs_template)
                tpl.render(ctx)
                page_path = os.path.join(temp_dir, f"page_{p_idx}.docx")
                tpl.save(page_path)
                page_doc_files.append(page_path)

                current_done = min((p_idx + 1) * 2, total_voters)
                if progress_callback:
                    progress_callback(current_done, total_voters, sheet_name)

        if not page_doc_files:
            return False, 0, "Tidak ada halaman yang berhasil dibuat."

        base_doc = docx.Document(page_doc_files[0])
        composer = Composer(base_doc)
        for next_page_file in page_doc_files[1:]:
            if cancel_event is not None and cancel_event.is_set():
                return False, 0, "Proses mailing dibatalkan pengguna."
            if layout_mode == "1-UP":
                composer.doc.add_page_break()
            composer.append(docx.Document(next_page_file))

        merged_docx_path = os.path.join(temp_dir, "merged_kadus.docx")
        composer.save(merged_docx_path)

        if cancel_event is not None and cancel_event.is_set():
            return False, 0, "Proses mailing dibatalkan pengguna."

        ok, msg, final_pdf = export_docx_to_pdf_via_word(merged_docx_path, abs_target_pdf)
        if not ok:
            if os.path.exists(abs_target_pdf):
                safe_delete_file(abs_target_pdf)
            return False, 0, msg

        apply_pdf_metadata(final_pdf, kadus_name=wilayah_name or sheet_name)
        return True, pages_count, "Berhasil"
    except Exception as exc:
        if os.path.exists(abs_target_pdf):
            safe_delete_file(abs_target_pdf)
        return False, 0, str(exc)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

def generate_preview_via_word(
    template_docx_path: str,
    sample_voters: list[VoterData],
    output_pdf_path: str,
    layout_mode: str = "2-UP",
) -> tuple[bool, str, str]:
    if not sample_voters:
        sample_voters = [
            VoterData(
                nomor="1",
                nama="ALIF MUBAROK",
                jenis_kelamin="Laki-laki",
                umur="21 Tahun",
                alamat="KARANGJENGKOL, RT 001 / RW 001",
                sheet_name="KADUS 1",
                wilayah="KADUS 1",
                row_index=12,
            ),
            VoterData(
                nomor="2",
                nama="HAYATUNISA",
                jenis_kelamin="Perempuan",
                umur="19 Tahun",
                alamat="KARANGJENGKOL, RT 001 / RW 001",
                sheet_name="KADUS 1",
                wilayah="KADUS 1",
                row_index=13,
            ),
        ]
    elif len(sample_voters) == 1 and layout_mode == "2-UP":
        sample_voters = [
            sample_voters[0],
            VoterData(
                nomor="2",
                nama="HAYATUNISA",
                jenis_kelamin="Perempuan",
                umur="19 Tahun",
                alamat="KARANGJENGKOL, RT 001 / RW 001",
                sheet_name=sample_voters[0].sheet_name,
                wilayah=sample_voters[0].wilayah,
                row_index=13,
            ),
        ]

    limit = 1 if layout_mode == "1-UP" else 2
    ok, pages, msg = merge_voters_to_pdf_via_word_com(
        template_docx_path=template_docx_path,
        voters=sample_voters[:limit],
        sheet_name=sample_voters[0].sheet_name,
        wilayah_name=sample_voters[0].wilayah,
        output_pdf_path=output_pdf_path,
        layout_mode=layout_mode,
    )
    final_pdf = to_win_abs_path(output_pdf_path)
    if ok and os.path.exists(final_pdf):
        apply_pdf_metadata(final_pdf, kadus_name="Pratinjau")
    return ok, msg, final_pdf
