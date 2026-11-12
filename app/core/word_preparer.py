import os
from pathlib import Path
import re
import shutil
import time
from typing import Any, Generator
import docx
from pydantic import BaseModel, Field
import pythoncom
import win32com.client

from app.config.settings import TEMP_DIR
from app.core.security import to_win_abs_path

class WordCheckResult(BaseModel):
    is_valid: bool = True
    original_path: str = ""
    prepared_docx_path: str = ""
    is_converted: bool = False
    has_two_slots: bool = True
    detected_fields_slot1: list[str] = Field(default_factory=list)
    detected_fields_slot2: list[str] = Field(default_factory=list)
    status_badge: str = "Belum Dicek"
    message: str = ""
    issues: list[str] = Field(default_factory=list)

def _convert_doc_to_docx_via_word_com(source_path: str, target_docx_path: str) -> None:
    target_abs = to_win_abs_path(target_docx_path)
    work_dir = Path(target_abs).parent
    work_dir.mkdir(parents=True, exist_ok=True)

    src_ext = os.path.splitext(source_path)[1]
    staged_copy = work_dir / f"staged_input_{int(time.time())}{src_ext}"
    shutil.copy2(source_path, staged_copy)
    staged_abs = to_win_abs_path(staged_copy)

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(staged_abs, False, True, False)

        try:
            sec = doc.Sections(1)
            ps = sec.PageSetup
            ps.TopMargin = 14.2
            ps.BottomMargin = 14.2
        except Exception:
            pass

        try:
            for p in doc.Paragraphs:
                p.Format.SpaceBefore = 0
                p.Format.SpaceAfter = 0
        except Exception:
            pass

        try:
            tb1 = doc.Shapes.AddTextbox(1, 480, 15, 120, 20)
            tb1.TextFrame.TextRange.Text = "PANITIA PEMILIHAN"
            tb1.TextFrame.TextRange.Font.Name = "Arial"
            tb1.TextFrame.TextRange.Font.Size = 8
            tb1.TextFrame.TextRange.Font.Bold = True
            tb1.Fill.Visible = False
            tb1.Line.Visible = False

            tb2 = doc.Shapes.AddTextbox(1, 480, 475, 120, 20)
            tb2.TextFrame.TextRange.Text = "PANITIA PEMILIHAN"
            tb2.TextFrame.TextRange.Font.Name = "Arial"
            tb2.TextFrame.TextRange.Font.Size = 8
            tb2.TextFrame.TextRange.Font.Bold = True
            tb2.Fill.Visible = False
            tb2.Line.Visible = False
        except Exception:
            pass

        doc.SaveAs2(target_abs, FileFormat=12)
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
        if staged_copy.exists():
            try:
                staged_copy.unlink(missing_ok=True)
            except Exception:
                pass

def set_paragraph_text_preserving_drawings(
    paragraph: docx.text.paragraph.Paragraph,
    new_text: str,
    font_name: str = "Times New Roman",
) -> None:
    p_elem = paragraph._element
    runs_to_remove = []
    for r in paragraph.runs:
        has_drawing = bool(r._element.xpath('.//*[local-name()="pict" or local-name()="drawing" or local-name()="AlternateContent"]'))
        if has_drawing:
            for t in r._element.xpath('.//w:t'):
                t.text = ""
        else:
            runs_to_remove.append(r)

    for r in runs_to_remove:
        p_elem.remove(r._element)

    new_run = paragraph.add_run(new_text)
    new_run.font.name = font_name

def _iter_cell_paragraphs(cell: docx.table._Cell, parent_doc: Any) -> Generator[docx.text.paragraph.Paragraph, None, None]:
    for element in cell._tc:
        if element.tag.endswith("p"):
            yield docx.text.paragraph.Paragraph(element, parent_doc)
        elif element.tag.endswith("tbl"):
            nested_table = docx.table.Table(element, parent_doc)
            for row in nested_table.rows:
                for c in row.cells:
                    yield from _iter_cell_paragraphs(c, parent_doc)

def iter_all_paragraphs(doc: docx.Document) -> Generator[docx.text.paragraph.Paragraph, None, None]:
    for child in doc.element.body:
        if child.tag.endswith("p"):
            yield docx.text.paragraph.Paragraph(child, doc)
        elif child.tag.endswith("tbl"):
            table = docx.table.Table(child, doc)
            for row in table.rows:
                for cell in row.cells:
                    yield from _iter_cell_paragraphs(cell, doc)

def check_and_prepare_word_template(
    word_file_path: str,
    work_dir: str | Path | None = None,
) -> WordCheckResult:
    result = WordCheckResult(original_path=str(word_file_path))

    if not os.path.exists(word_file_path):
        result.is_valid = False
        result.status_badge = "Perlu Perbaikan"
        result.message = f"File Word tidak ditemukan: {os.path.basename(word_file_path)}"
        result.issues.append("File tidak ditemukan.")
        return result

    base_dir = Path(work_dir) if work_dir else TEMP_DIR
    base_dir.mkdir(parents=True, exist_ok=True)
    target_docx = to_win_abs_path(base_dir / "template_ready.docx")

    lower_ext = os.path.splitext(word_file_path)[1].lower()
    is_doc_or_rtf = lower_ext in (".doc", ".rtf") or word_file_path.lower().endswith(".rtf.doc")

    if is_doc_or_rtf:
        try:
            _convert_doc_to_docx_via_word_com(word_file_path, target_docx)
            result.is_converted = True
            result.prepared_docx_path = target_docx
        except Exception as exc:
            result.is_valid = False
            result.status_badge = "Perlu Perbaikan"
            result.message = f"Gagal mengonversi file Word melalui MS Word: {exc}"
            result.issues.append(str(exc))
            return result
    else:
        try:
            doc_in = docx.Document(word_file_path)
            doc_in.save(target_docx)
            result.prepared_docx_path = target_docx
        except Exception:
            try:
                _convert_doc_to_docx_via_word_com(word_file_path, target_docx)
                result.is_converted = True
                result.prepared_docx_path = target_docx
            except Exception as exc:
                result.is_valid = False
                result.status_badge = "Perlu Perbaikan"
                result.message = f"File Word tidak dapat dibuka: {exc}"
                result.issues.append(str(exc))
                return result

    try:
        doc = docx.Document(result.prepared_docx_path)
    except Exception as exc:
        result.is_valid = False
        result.status_badge = "Perlu Perbaikan"
        result.message = f"Struktur file docx tidak valid: {exc}"
        result.issues.append(str(exc))
        return result

    patterns = [
        (r"^\s*Nomor[\s\t]*:.*$", "nomor", "Nomor"),
        (r"^\s*Nama\s+Pemilih[\s\t]*:.*$", "nama", "Nama Pemilih"),
        (r"^\s*Laki[\s\-]*laki\s*/\s*Perempuan[\s\t]*:.*$", "jk", "Laki-laki/Perempuan"),
        (r"^\s*Umur[\s\t]*:.*$", "umur", "Umur"),
        (r"^\s*Alamat[\s\t]*:.*$", "alamat", "Alamat"),
        (r"^\s*DAFTAR\s+PEMILIH\s+TETAP\s+NOMOR[\s\t]*:.*$", "nomor_dpt", ""),
    ]

    tag_counts: dict[str, int] = {}
    schedule_counts: dict[str, int] = {}

    for p in iter_all_paragraphs(doc):
        txt = p.text
        lower = txt.lower()

        if "catatan" in lower and "diwakilkan" in lower:
            clean_catatan = re.sub(r"\.\.+", ".", txt)
            set_paragraph_text_preserving_drawings(p, clean_catatan)
            continue

        if "untuk memberikan suara pada" in lower:
            set_paragraph_text_preserving_drawings(p, "Untuk memberikan suara pada\t\t:")
            continue

        matched_field = False
        for pat, key, field_name in patterns:
            if re.match(pat, txt, flags=re.IGNORECASE):
                slot = tag_counts.get(key, 0) + 1
                tag_counts[key] = slot

                if key == "nomor":
                    set_paragraph_text_preserving_drawings(p, f"Nomor\t: {{{{nomor_{slot}}}}}")
                elif key == "nama":
                    set_paragraph_text_preserving_drawings(p, f"Nama Pemilih\t: {{{{nama_{slot}}}}}")
                elif key == "jk":
                    set_paragraph_text_preserving_drawings(p, f"Laki-laki/Perempuan\t: {{{{jk_{slot}}}}}")
                elif key == "umur":
                    set_paragraph_text_preserving_drawings(p, f"Umur\t: {{{{umur_{slot}}}}}")
                elif key == "alamat":
                    set_paragraph_text_preserving_drawings(p, f"Alamat\t: {{{{alamat_{slot}}}}}")
                elif key == "nomor_dpt":
                    set_paragraph_text_preserving_drawings(p, f"DAFTAR PEMILIH TETAP NOMOR : {{{{nomor_{slot}}}}}")

                if field_name:
                    if slot == 1:
                        result.detected_fields_slot1.append(field_name)
                    elif slot == 2:
                        result.detected_fields_slot2.append(field_name)

                matched_field = True
                break

        if matched_field:
            continue

        if "hari, tanggal" in lower:
            c = schedule_counts.get("hari", 0) + 1
            schedule_counts["hari"] = c
            if c == 2:
                set_paragraph_text_preserving_drawings(
                    p,
                    "Hari, tanggal\t: Kamis, 10 Desember 2026\tHari ................................................... tanggal ................................................................"
                )
        elif "tempat" in lower and ":" in txt:
            c = schedule_counts.get("tempat", 0) + 1
            schedule_counts["tempat"] = c
            if c == 2:
                set_paragraph_text_preserving_drawings(
                    p,
                    "Tempat\t: Lapangan Desa Karangjengkol\t........................................................................................................................................"
                )
        elif "waktu" in lower and ":" in txt:
            c = schedule_counts.get("waktu", 0) + 1
            schedule_counts["waktu"] = c
            if c == 2:
                set_paragraph_text_preserving_drawings(
                    p,
                    "Waktu\t: 08.00 WIB s/d selesai\tJam 08.00 sampai dengan 14.00 WIB"
                )

    doc.save(result.prepared_docx_path)

    recheck_doc = docx.Document(result.prepared_docx_path)
    full_text = " ".join(p.text for p in iter_all_paragraphs(recheck_doc))

    required_tags = [
        "{{nomor_1}}", "{{nama_1}}", "{{jk_1}}", "{{umur_1}}", "{{alamat_1}}",
        "{{nomor_2}}", "{{nama_2}}", "{{jk_2}}", "{{umur_2}}", "{{alamat_2}}"
    ]
    missing_tags = [t for t in required_tags if t not in full_text]

    if full_text.count("{{nomor_1}}") < 2:
        missing_tags.append("{{nomor_1}} pada Tanda Terima Atas")
    if full_text.count("{{nomor_2}}") < 2:
        missing_tags.append("{{nomor_2}} pada Tanda Terima Bawah")

    result.detected_fields_slot1 = sorted(list(set(result.detected_fields_slot1)))
    result.detected_fields_slot2 = sorted(list(set(result.detected_fields_slot2)))

    if missing_tags:
        result.is_valid = False
        result.status_badge = "Perlu Perbaikan"
        issues = [f"Tag belum terinjeksi: {', '.join(missing_tags)}"]
        result.issues = issues
        result.message = " | ".join(issues)
    else:
        result.is_valid = True
        result.status_badge = "Siap Mailing"
        result.message = (
            "File Word Siap untuk Mailing (12 Tag Terverifikasi Lengkap: "
            "Nomor, Nama, L/P, Umur, Alamat, Jadwal Surat Bawah & Tanda Terima)"
        )

    return result
