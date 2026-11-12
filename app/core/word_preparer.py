from dataclasses import dataclass, field
from datetime import datetime
import hashlib
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

@dataclass
class WordSnapshot:
    file_path: str
    sha256: str
    mtime: float
    size: int
    fields_slot1: list[str]
    fields_slot2: list[str]
    schedule_summary: dict[str, str]
    ketua: str
    paragraphs_count: int
    is_valid: bool = True

def get_file_fingerprint(file_path: str) -> tuple[str, float, int]:
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

def _extract_val_from_line(line_txt: str, prefix: str) -> str:
    parts = line_txt.split("\t")
    for p in parts:
        if ":" in p and not p.lower().startswith(prefix.lower()):
            v = p.lstrip(":").strip()
            if v and not v.startswith("..."):
                return v
        elif ":" in p:
            v = p.split(":", 1)[1].strip()
            if v and not v.startswith("..."):
                return v
    return ""

def compute_word_snapshot(
    docx_path: str,
    original_file_path: str,
    is_valid: bool = True,
    detected_slot1: list[str] | None = None,
    detected_slot2: list[str] | None = None,
) -> WordSnapshot:
    sha256, mtime, size = get_file_fingerprint(original_file_path)
    schedule: dict[str, str] = {}
    ketua = ""
    paragraphs_count = 0

    if os.path.exists(docx_path):
        try:
            doc = docx.Document(docx_path)
            paragraphs = list(iter_all_paragraphs(doc))
            paragraphs_count = len(paragraphs)
            for p in paragraphs:
                txt = p.text.strip()
                tl = txt.lower()
                if "hari, tanggal" in tl and "hari" not in schedule:
                    schedule["hari"] = _extract_val_from_line(txt, "Hari, tanggal")
                elif "tempat" in tl and ":" in txt and "tempat" not in schedule:
                    schedule["tempat"] = _extract_val_from_line(txt, "Tempat")
                elif "waktu" in tl and ":" in txt and "waktu" not in schedule:
                    schedule["waktu"] = _extract_val_from_line(txt, "Waktu")
                elif "ratno" in tl:
                    ketua = "RATNO"
        except Exception:
            pass

    return WordSnapshot(
        file_path=original_file_path,
        sha256=sha256,
        mtime=mtime,
        size=size,
        fields_slot1=sorted(list(set(detected_slot1 or []))),
        fields_slot2=sorted(list(set(detected_slot2 or []))),
        schedule_summary=schedule,
        ketua=ketua,
        paragraphs_count=paragraphs_count,
        is_valid=is_valid,
    )

def analyze_word_diff(old_snap: WordSnapshot | None, new_snap: WordSnapshot) -> list[str]:
    lines: list[str] = []
    file_name = os.path.basename(new_snap.file_path)
    mod_time_str = datetime.fromtimestamp(new_snap.mtime).strftime("%d/%m/%Y %H:%M:%S")
    lines.append(f"[DIFF WORD] File Word terdeteksi diperbarui: {file_name}")
    lines.append(f"[DIFF WORD] Waktu Simpan: {mod_time_str} | Ukuran: {new_snap.size:,} bytes")

    if old_snap is None:
        lines.append("[DIFF WORD] Snapshot awal template Word berhasil dimuat.")
        if new_snap.schedule_summary:
            lines.append(f"[DIFF WORD] Jadwal Acara: {new_snap.schedule_summary.get('hari', '-')} | Tempat: {new_snap.schedule_summary.get('tempat', '-')}")
        lines.append(f"[DIFF WORD] Field Slot 1: {', '.join(new_snap.fields_slot1) or 'Kosong'}")
        lines.append(f"[DIFF WORD] Field Slot 2: {', '.join(new_snap.fields_slot2) or 'Kosong'}")
        return lines

    schedule_changed = False
    for k, label in [("hari", "Hari/Tanggal"), ("tempat", "Tempat"), ("waktu", "Waktu")]:
        old_v = old_snap.schedule_summary.get(k, "")
        new_v = new_snap.schedule_summary.get(k, "")
        if old_v != new_v:
            schedule_changed = True
            lines.append(f"[DIFF WORD] Perubahan {label}: '{old_v}' -> '{new_v}'")

    if not schedule_changed:
        lines.append("[DIFF WORD] Teks jadwal acara (Hari, Tempat, Waktu) tidak mengalami perubahan.")

    if old_snap.fields_slot1 != new_snap.fields_slot1:
        lines.append(f"[DIFF WORD] Perubahan Field Slot 1: {old_snap.fields_slot1} -> {new_snap.fields_slot1}")
    if old_snap.fields_slot2 != new_snap.fields_slot2:
        lines.append(f"[DIFF WORD] Perubahan Field Slot 2: {old_snap.fields_slot2} -> {new_snap.fields_slot2}")

    if old_snap.paragraphs_count != new_snap.paragraphs_count:
        p_diff = new_snap.paragraphs_count - old_snap.paragraphs_count
        p_sign = f"+{p_diff}" if p_diff > 0 else str(p_diff)
        lines.append(f"[DIFF WORD] Jumlah Paragraf: {old_snap.paragraphs_count} -> {new_snap.paragraphs_count} ({p_sign})")

    status_str = "SIAP MAILING" if new_snap.is_valid else "PERLU PERBAIKAN"
    lines.append(f"[DIFF WORD] Status Validasi Template: {status_str}")
    return lines

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

    # Pre-scan Slot 1 to dynamically extract schedule values and column 2 text
    extracted_schedule: dict[str, str] = {
        "hari": "Kamis, 10 Desember 2026",
        "tempat": "Lapangan Desa Karangjengkol",
        "waktu": "08.00 WIB s/d selesai",
        "col2_hari": "Hari ................................................... tanggal ................................................................",
        "col2_tempat": "........................................................................................................................................",
        "col2_waktu": "Jam 08.00 sampai dengan 14.00 WIB",
    }
    seen_schedule_keys: set[str] = set()

    for p in iter_all_paragraphs(doc):
        t = p.text.strip()
        tl = t.lower()
        if "hari, tanggal" in tl and "hari" not in seen_schedule_keys:
            seen_schedule_keys.add("hari")
            val = _extract_val_from_line(t, "Hari, tanggal")
            if val:
                extracted_schedule["hari"] = val
            parts = t.split("\t")
            if len(parts) > 2 and parts[2].strip():
                extracted_schedule["col2_hari"] = parts[2].strip()
        elif "tempat" in tl and ":" in t and "tempat" not in seen_schedule_keys:
            seen_schedule_keys.add("tempat")
            val = _extract_val_from_line(t, "Tempat")
            if val:
                extracted_schedule["tempat"] = val
            parts = t.split("\t")
            if len(parts) > 2 and parts[2].strip():
                extracted_schedule["col2_tempat"] = parts[2].strip()
        elif "waktu" in tl and ":" in t and "waktu" not in seen_schedule_keys:
            seen_schedule_keys.add("waktu")
            val = _extract_val_from_line(t, "Waktu")
            if val:
                extracted_schedule["waktu"] = val
            parts = t.split("\t")
            if len(parts) > 2 and parts[2].strip():
                extracted_schedule["col2_waktu"] = parts[2].strip()

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
                slot2_val = _extract_val_from_line(txt, "Hari, tanggal")
                if not slot2_val:
                    set_paragraph_text_preserving_drawings(
                        p,
                        f"Hari, tanggal\t: {extracted_schedule['hari']}\t{extracted_schedule['col2_hari']}"
                    )
        elif "tempat" in lower and ":" in txt:
            c = schedule_counts.get("tempat", 0) + 1
            schedule_counts["tempat"] = c
            if c == 2:
                slot2_val = _extract_val_from_line(txt, "Tempat")
                if not slot2_val:
                    set_paragraph_text_preserving_drawings(
                        p,
                        f"Tempat\t: {extracted_schedule['tempat']}\t{extracted_schedule['col2_tempat']}"
                    )
        elif "waktu" in lower and ":" in txt:
            c = schedule_counts.get("waktu", 0) + 1
            schedule_counts["waktu"] = c
            if c == 2:
                slot2_val = _extract_val_from_line(txt, "Waktu")
                if not slot2_val:
                    set_paragraph_text_preserving_drawings(
                        p,
                        f"Waktu\t: {extracted_schedule['waktu']}\t{extracted_schedule['col2_waktu']}"
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
