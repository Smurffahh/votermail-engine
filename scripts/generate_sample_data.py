# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

from pathlib import Path
import random
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
import openpyxl

CELEBRITY_SHEETS = {
    "KADUS 1": [
        ("RAFFI FARID AHMAD", "L", 37, "S", "KARANGJENGKOL", "001", "001"),
        ("NAGITA SLAVINA MARIANA TENGKER", "P", 36, "S", "KARANGJENGKOL", "001", "001"),
        ("DIAN SASTROWARDOYO", "P", 42, "S", "KARANGJENGKOL", "002", "001"),
        ("REZA RAHADIAN MATULESSY", "L", 37, "B", "KARANGJENGKOL", "002", "001"),
        ("NAJWA SHIHAB", "P", 47, "S", "KARANGJENGKOL", "003", "001"),
        ("MAUDY AYUNDA", "P", 30, "S", "KARANGJENGKOL", "003", "001"),
        ("IKO UWAIS", "L", 41, "S", "KARANGJENGKOL", "004", "001"),
        ("JOE TASLIM", "L", 43, "S", "KARANGJENGKOL", "004", "001"),
        ("DEDDY CORBUZIER", "L", 48, "S", "KARANGJENGKOL", "001", "001"),
        ("VINCENT RYAN ROMPIES", "L", 44, "S", "KARANGJENGKOL", "002", "001"),
        ("DEDDY MAHENDRA DESTA", "L", 47, "P", "KARANGJENGKOL", "002", "001"),
        ("PRILLY LATUCONSINA", "P", 28, "B", "KARANGJENGKOL", "003", "001"),
        ("RADITYA DIKA", "L", 40, "S", "KARANGJENGKOL", "003", "001"),
        ("CINTA LAURA KIEHL", "P", 31, "B", "KARANGJENGKOL", "004", "001"),
        ("PEVITA PEARCE", "P", 32, "S", "KARANGJENGKOL", "004", "001"),
        ("CHICCO JERIKHO", "L", 40, "S", "KARANGJENGKOL", "001", "001"),
        ("TARA BASRO", "P", 34, "S", "KARANGJENGKOL", "002", "001"),
        ("ISYANA SARASVATI", "P", 31, "S", "KARANGJENGKOL", "002", "001"),
        ("AFGAN SYAHREZA", "L", 35, "B", "KARANGJENGKOL", "003", "001"),
        ("BUNGA CITRA LESTARI", "P", 41, "S", "KARANGJENGKOL", "003", "001"),
        ("ARIEL NOAH", "L", 43, "P", "KARANGJENGKOL", "004", "001"),
        ("SURYA INSOMNIA", "L", 38, "S", "KARANGJENGKOL", "004", "001"),
        ("WENDI CAGUR", "L", 45, "S", "KARANGJENGKOL", "001", "001"),
        ("ANDRE TAULANY", "L", 50, "S", "KARANGJENGKOL", "002", "001"),
        ("ERNEST PRAKASA", "L", 42, "S", "KARANGJENGKOL", "003", "001"),
    ],
    "KADUS 2": [
        ("AGNEZ MO", "P", 38, "B", "DUSUN KARANGJENGKOL", "001", "002"),
        ("TULUS RUSYDI", "L", 37, "B", "DUSUN KARANGJENGKOL", "001", "002"),
        ("RAISA ANDRIANA", "P", 34, "S", "DUSUN KARANGJENGKOL", "002", "002"),
        ("HAMISH DAUD", "L", 44, "S", "DUSUN KARANGJENGKOL", "002", "002"),
        ("NICHOLAS SAPUTRA", "L", 40, "B", "DUSUN KARANGJENGKOL", "003", "002"),
        ("CHELSEA ISLAN", "P", 29, "S", "DUSUN KARANGJENGKOL", "003", "002"),
        ("IQBAAL RAMADHAN", "L", 25, "B", "DUSUN KARANGJENGKOL", "004", "002"),
        ("ANGGA YUNANDA", "L", 24, "B", "DUSUN KARANGJENGKOL", "004", "002"),
        ("SHENINA CINNAMON", "P", 25, "B", "DUSUN KARANGJENGKOL", "001", "002"),
        ("JEFRI NICHOL", "L", 25, "B", "DUSUN KARANGJENGKOL", "001", "002"),
        ("ADIPATI DOLKEN", "L", 33, "S", "DUSUN KARANGJENGKOL", "002", "002"),
        ("YURA YUNITA", "P", 33, "S", "DUSUN KARANGJENGKOL", "002", "002"),
        ("DONNE MAULA", "L", 35, "S", "DUSUN KARANGJENGKOL", "003", "002"),
        ("EKO PATRIO", "L", 54, "S", "DUSUN KARANGJENGKOL", "003", "002"),
        ("SULE PURNOMO", "L", 48, "P", "DUSUN KARANGJENGKOL", "004", "002"),
        ("RIZKY FEBIAN", "L", 26, "S", "DUSUN KARANGJENGKOL", "004", "002"),
        ("MAHALINI RAHARJA", "P", 24, "S", "DUSUN KARANGJENGKOL", "001", "002"),
        ("ATTA HALILINTAR", "L", 30, "S", "DUSUN KARANGJENGKOL", "001", "002"),
        ("AUREL HERMANSYAH", "P", 26, "S", "DUSUN KARANGJENGKOL", "002", "002"),
        ("ANANG HERMANSYAH", "L", 55, "S", "DUSUN KARANGJENGKOL", "002", "002"),
        ("ASHANTY HASTUTI", "P", 40, "S", "DUSUN KARANGJENGKOL", "003", "002"),
        ("JUDIKA SIHOTANG", "L", 46, "S", "DUSUN KARANGJENGKOL", "003", "002"),
        ("ROSSA ROSLINA", "P", 46, "P", "DUSUN KARANGJENGKOL", "004", "002"),
        ("TITI DJ", "P", 58, "P", "DUSUN KARANGJENGKOL", "004", "002"),
        ("ARMAND MAULANA", "L", 53, "S", "DUSUN KARANGJENGKOL", "001", "002"),
    ],
    "KADUS 3": [
        ("GREYSIA POLII", "P", 37, "S", "LEGOK", "001", "003"),
        ("APRIYANI RAHAYU", "P", 26, "B", "LEGOK", "001", "003"),
        ("ANTHONY SINISUKA GINTING", "L", 28, "S", "LEGOK", "002", "003"),
        ("JONATAN CHRISTIE", "L", 27, "S", "LEGOK", "002", "003"),
        ("TAUFIK HIDAYAT", "L", 43, "S", "LEGOK", "003", "003"),
        ("SUSI SUSANTI", "P", 53, "S", "LEGOK", "003", "003"),
        ("ALAN BUDIKUSUMA", "L", 56, "S", "LEGOK", "004", "003"),
        ("LILIYANA NATSIR", "P", 39, "B", "LEGOK", "004", "003"),
        ("TONTOWI AHMAD", "L", 37, "S", "LEGOK", "001", "003"),
        ("BAMBANG PAMUNGKAS", "L", 44, "S", "LEGOK", "001", "003"),
        ("BOAZ SOLOSSA", "L", 38, "S", "LEGOK", "002", "003"),
        ("PRATAMA ARHAN", "L", 23, "S", "LEGOK", "002", "003"),
        ("ASNAWI MANGKUALAM", "L", 25, "B", "LEGOK", "003", "003"),
        ("MARSELINO FERDINAN", "L", 20, "B", "LEGOK", "003", "003"),
        ("RIZKY RIDHO", "L", 23, "B", "LEGOK", "004", "003"),
        ("ERNANDO ARI SUTARYADI", "L", 22, "B", "LEGOK", "004", "003"),
        ("JAY IDZES", "L", 24, "B", "LEGOK", "001", "003"),
        ("THOM HAYE", "L", 29, "B", "LEGOK", "001", "003"),
        ("NATHAN TJOE-A-ON", "L", 23, "B", "LEGOK", "002", "003"),
        ("RAFAEL STRUICK", "L", 21, "B", "LEGOK", "002", "003"),
        ("IVAR JENNER", "L", 20, "B", "LEGOK", "003", "003"),
        ("SANDY WALSH", "L", 29, "S", "LEGOK", "003", "003"),
        ("JORDI AMAT", "L", 32, "S", "LEGOK", "004", "003"),
        ("MAARTEN PAES", "L", 26, "B", "LEGOK", "004", "003"),
        ("SHIN TAE-YONG", "L", 54, "S", "LEGOK", "001", "003"),
    ],
    "KADUS 4": [
        ("BUTET KARTARAREDJASA", "L", 63, "S", "KARANGJENGKOL", "001", "004"),
        ("SLAMET RAHARDJO", "L", 75, "S", "KARANGJENGKOL", "001", "004"),
        ("CHRISTINE HAKIM", "P", 68, "S", "KARANGJENGKOL", "002", "004"),
        ("JOKO ANWAR", "L", 48, "B", "KARANGJENGKOL", "002", "004"),
        ("MIRA LESMANA", "P", 60, "S", "KARANGJENGKOL", "003", "004"),
        ("RIRI RIZA", "L", 54, "S", "KARANGJENGKOL", "003", "004"),
        ("HANUNG BRAMANTYO", "L", 49, "S", "KARANGJENGKOL", "004", "004"),
        ("ZASKIA ADYA MECCA", "P", 37, "S", "KARANGJENGKOL", "004", "004"),
        ("LUKMAN SARDI", "L", 53, "S", "KARANGJENGKOL", "001", "004"),
        ("RINGGO AGUS RAHMAN", "L", 42, "S", "KARANGJENGKOL", "001", "004"),
        ("SABAI MORSCHECK", "P", 36, "S", "KARANGJENGKOL", "002", "004"),
        ("DION WIYOKO", "L", 40, "S", "KARANGJENGKOL", "002", "004"),
        ("GADING MARTEN", "L", 42, "P", "KARANGJENGKOL", "003", "004"),
        ("GISELLA ANASTASIA", "P", 34, "P", "KARANGJENGKOL", "003", "004"),
        ("DENNY SUMARGO", "L", 43, "S", "KARANGJENGKOL", "004", "004"),
        ("OLIVIA ALLAN", "P", 38, "S", "KARANGJENGKOL", "004", "004"),
        ("BOY WILLIAM", "L", 33, "B", "KARANGJENGKOL", "001", "004"),
        ("DANIEL MANANTA", "L", 43, "S", "KARANGJENGKOL", "001", "004"),
        ("INDRA BEKTI", "L", 47, "S", "KARANGJENGKOL", "002", "004"),
        ("IRFAN HAKIM", "L", 49, "S", "KARANGJENGKOL", "002", "004"),
        ("RAMZI GEZALI", "L", 48, "S", "KARANGJENGKOL", "003", "004"),
        ("GILANG DIRGA", "L", 35, "S", "KARANGJENGKOL", "003", "004"),
        ("KIKY SAPUTRI", "P", 31, "S", "KARANGJENGKOL", "004", "004"),
        ("MARSHEL WIDIANTO", "L", 28, "S", "KARANGJENGKOL", "004", "004"),
        ("BORIS BOKIR", "L", 36, "P", "KARANGJENGKOL", "001", "004"),
    ],
    "KADUS 5": [
        ("CAK LONTONG (LIES HARTONO)", "L", 54, "S", "LEGOK", "001", "005"),
        ("JARWO KWAT", "L", 57, "S", "LEGOK", "001", "005"),
        ("KOMENG (ALFIANSYAH)", "L", 54, "S", "LEGOK", "002", "005"),
        ("ADUL (ABDUL LATIEF)", "L", 41, "S", "LEGOK", "002", "005"),
        ("TARZAN (TOTO MULYADI)", "L", 79, "S", "LEGOK", "003", "005"),
        ("TESSY (KABUL BASUKI)", "L", 77, "S", "LEGOK", "003", "005"),
        ("NUNUNG SURYANI", "P", 61, "S", "LEGOK", "004", "005"),
        ("PARTO PATRIO", "L", 63, "S", "LEGOK", "004", "005"),
        ("AKRI PATRIO", "L", 55, "S", "LEGOK", "001", "005"),
        ("MANDRA NAIH", "L", 58, "S", "LEGOK", "001", "005"),
        ("RANO KARNO", "L", 64, "S", "LEGOK", "002", "005"),
        ("MAUDY KOESNAEDI", "P", 49, "S", "LEGOK", "002", "005"),
        ("CORNELIA AGATHA", "P", 51, "P", "LEGOK", "003", "005"),
        ("MAT SOLAR", "L", 61, "S", "LEGOK", "003", "005"),
        ("NANI WIJAYA", "P", 68, "S", "LEGOK", "004", "005"),
        ("FANNY FADILLAH", "L", 57, "S", "LEGOK", "004", "005"),
        ("DODIT MULYANTO", "L", 39, "B", "LEGOK", "001", "005"),
        ("INDRA JEGEL", "L", 35, "S", "LEGOK", "001", "005"),
        ("OZA RANGKUTI", "L", 33, "B", "LEGOK", "002", "005"),
        ("BENIDICTIVITY", "L", 34, "B", "LEGOK", "002", "005"),
        ("AWECAN", "L", 38, "S", "LEGOK", "003", "005"),
        ("BEBI ROMEO", "L", 50, "S", "LEGOK", "003", "005"),
        ("MEISYA SIREGAR", "P", 45, "S", "LEGOK", "004", "005"),
        ("MONA RATULIU", "P", 42, "S", "LEGOK", "004", "005"),
        ("INDRA BRASCO", "L", 50, "S", "LEGOK", "001", "005"),
    ],
}

def generate_sample_dpt_excel(output_path: str, voters_per_sheet: int = 25) -> str:
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    if default_sheet is not None:
        wb.remove(default_sheet)

    for sheet_name, voters in CELEBRITY_SHEETS.items():
        ws = wb.create_sheet(title=sheet_name)

        ws.cell(1, 1, "DAFTAR PEMILIH TETAP")
        ws.cell(3, 1, "PEMILIHAN KEPALA DESA KARANGJENGKOL KECAMATAN KUTASARI")
        ws.cell(4, 1, "KABUPATEN PURBALINGGA TAHUN 2026")

        ws.cell(6, 2, "DESA")
        ws.cell(6, 3, ": KARANGJENGKOL")

        ws.cell(7, 2, "KECAMATAN")
        ws.cell(7, 3, ": KUTASARI")

        ws.cell(8, 2, "KABUPATEN")
        ws.cell(8, 3, ": PURBALINGGA")
        ws.cell(8, 9, "WILAYAH")
        ws.cell(8, 10, f": {sheet_name}")

        headers = [
            "Nomor Urut",
            "Nomor Induk Kependudukan (NIK)",
            "Nama Pemilih",
            "Tempat Lahir",
            "Tanggal Lahir",
            "Usia (Tahun)",
            "Status Perkawinan (B/S/P)",
            "Jenis Kelamin",
            "Alamat",
            "RT",
            "RW",
            "Keterangan",
        ]
        for c_idx, h_text in enumerate(headers, start=1):
            ws.cell(10, c_idx, h_text)

        for c_idx in range(1, 13):
            ws.cell(11, c_idx, c_idx)

        curr_row = 12
        male_count = 0
        female_count = 0

        for v_idx, item in enumerate(voters, start=1):
            nama, jk, age, status, alamat, rt, rw = item
            if jk == "L":
                male_count += 1
            else:
                female_count += 1

            b_year = 2026 - age
            b_day = (v_idx % 28) + 1
            b_month = (v_idx % 12) + 1
            b_date_str = f"{b_day:02d}|{b_month:02d}|{b_year}"
            dummy_nik = f"330307999999{v_idx:04d}"

            ws.cell(curr_row, 1, v_idx)
            ws.cell(curr_row, 2, dummy_nik)
            ws.cell(curr_row, 3, nama)
            ws.cell(curr_row, 4, "PURBALINGGA")
            ws.cell(curr_row, 5, b_date_str)
            ws.cell(curr_row, 6, age)
            ws.cell(curr_row, 7, status)
            ws.cell(curr_row, 8, jk)
            ws.cell(curr_row, 9, alamat)
            ws.cell(curr_row, 10, rt)
            ws.cell(curr_row, 11, rw)
            ws.cell(curr_row, 12, None)

            curr_row += 1

        curr_row += 1
        ws.cell(curr_row, 2, "Jumlah Pemilih ")
        ws.cell(curr_row, 4, "Status Perkawinan:")
        ws.cell(curr_row, 7, "Disahkan di")
        ws.cell(curr_row, 8, ": Karangjengkol")
        curr_row += 1

        ws.cell(curr_row, 2, "Laki-laki ")
        ws.cell(curr_row, 3, male_count)
        ws.cell(curr_row, 4, "B = Belum Nikah")
        ws.cell(curr_row, 7, "Pada Tanggal")
        ws.cell(curr_row, 8, ": 5 November 2026")
        curr_row += 1

        ws.cell(curr_row, 2, "Perempuan")
        ws.cell(curr_row, 3, female_count)
        ws.cell(curr_row, 4, "S = Sudah Nikah")
        curr_row += 1

        ws.cell(curr_row, 2, "Total")
        ws.cell(curr_row, 3, len(voters))
        ws.cell(curr_row, 4, "P = Pernah Nikah")
        ws.cell(curr_row, 6, "PANITIA PELAKSANA PEMILIHAN KEPALA DESA")
        curr_row += 2

        ws.cell(curr_row, 6, 1)
        ws.cell(curr_row, 7, "Ketua")
        ws.cell(curr_row, 8, "Ratno")
        curr_row += 1

        ws.cell(curr_row, 6, 2)
        ws.cell(curr_row, 7, "Sekretaris")
        ws.cell(curr_row, 8, "Danang Satria Utomo")

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()
    return output_path

def generate_sample_word_dummy(output_path: str) -> str:
    doc = docx.Document()

    # Configure Section for F4 (Folio 215 x 330 mm = 8.5 x 13.0 inches)
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(13.0)
    section.top_margin = Pt(14.2)
    section.bottom_margin = Pt(14.2)
    section.left_margin = Pt(28.0)
    section.right_margin = Pt(28.0)

    def _add_slot(slot_idx: int) -> None:
        p_hdr = doc.add_paragraph()
        p_hdr.paragraph_format.space_before = Pt(0)
        p_hdr.paragraph_format.space_after = Pt(2)
        p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p_hdr.add_run("PANITIA PEMILIHAN KEPALA DESA KARANGJENGKOL\nKECAMATAN KUTASARI KABUPATEN PURBALINGGA\n")
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(10)
        r1.font.bold = True

        r2 = p_hdr.add_run("SURAT PEMBERITAHUAN PEMUNGUTAN SUARA KEPADA PEMILIH")
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(11)
        r2.font.bold = True
        r2.font.underline = True

        p_no = doc.add_paragraph()
        p_no.paragraph_format.space_before = Pt(0)
        p_no.paragraph_format.space_after = Pt(2)
        r_no = p_no.add_run(f"Nomor\t: {{{{nomor_{slot_idx}}}}}")
        r_no.font.name = "Times New Roman"
        r_no.font.size = Pt(9.5)

        p_intro = doc.add_paragraph()
        p_intro.paragraph_format.space_before = Pt(0)
        p_intro.paragraph_format.space_after = Pt(2)
        r_intro = p_intro.add_run(
            "Bersama ini diberitahukan bahwa Panitia Pemilihan Kepala Desa Karangjengkol "
            "mengundang Saudara/i untuk memberikan suara:"
        )
        r_intro.font.name = "Times New Roman"
        r_intro.font.size = Pt(9)

        fields = [
            ("Nama Pemilih", f"{{{{nama_{slot_idx}}}}}"),
            ("Laki-laki/Perempuan", f"{{{{jk_{slot_idx}}}}}"),
            ("Umur", f"{{{{umur_{slot_idx}}}}}"),
            ("Alamat", f"{{{{alamat_{slot_idx}}}}}"),
        ]
        for f_label, f_tag in fields:
            pf = doc.add_paragraph()
            pf.paragraph_format.space_before = Pt(0)
            pf.paragraph_format.space_after = Pt(1)
            rf = pf.add_run(f"{f_label}\t: {f_tag}")
            rf.font.name = "Times New Roman"
            rf.font.size = Pt(9.5)

        schedules = [
            ("Hari, tanggal", "Kamis, 10 Desember 2026", "Hari ................................................... tanggal ................................................................"),
            ("Tempat", "Lapangan Desa Karangjengkol", "........................................................................................................................................"),
            ("Waktu", "08.00 WIB s/d selesai", "Jam 08.00 sampai dengan 14.00 WIB"),
        ]
        for s_label, s_val, s_col2 in schedules:
            ps = doc.add_paragraph()
            ps.paragraph_format.space_before = Pt(0)
            ps.paragraph_format.space_after = Pt(1)
            rs = ps.add_run(f"{s_label}\t: {s_val}\t{s_col2}")
            rs.font.name = "Times New Roman"
            rs.font.size = Pt(9)

        p_cat = doc.add_paragraph()
        p_cat.paragraph_format.space_before = Pt(1)
        p_cat.paragraph_format.space_after = Pt(2)
        rcat = p_cat.add_run("CATATAN\t: Untuk memberikan suara, surat pemberitahuan ini harus dibawa oleh pemilih sendiri.")
        rcat.font.name = "Times New Roman"
        rcat.font.size = Pt(8.5)
        rcat.font.italic = True

        p_ttd = doc.add_paragraph()
        p_ttd.paragraph_format.space_before = Pt(0)
        p_ttd.paragraph_format.space_after = Pt(0)
        p_ttd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rttd = p_ttd.add_run("Panitia Pemilihan Kepala Desa Karangjengkol\nKetua,\n\n\nRATNO")
        rttd.font.name = "Times New Roman"
        rttd.font.size = Pt(9)
        rttd.font.bold = True

        p_rcpt_hdr = doc.add_paragraph()
        p_rcpt_hdr.paragraph_format.space_before = Pt(2)
        p_rcpt_hdr.paragraph_format.space_after = Pt(1)
        r_rcpt_hdr = p_rcpt_hdr.add_run(f"DAFTAR PEMILIH TETAP NOMOR : {{{{nomor_{slot_idx}}}}}")
        r_rcpt_hdr.font.name = "Times New Roman"
        r_rcpt_hdr.font.size = Pt(9)
        r_rcpt_hdr.font.bold = True

        p_date = doc.add_paragraph()
        p_date.paragraph_format.space_before = Pt(0)
        p_date.paragraph_format.space_after = Pt(0)
        p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_date = p_date.add_run("Karangjengkol, ....................................\nPenerima/Pemilih,\n\n\n..............................................")
        r_date.font.name = "Times New Roman"
        r_date.font.size = Pt(8.5)

    _add_slot(1)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(4)
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_div = p_div.add_run("- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
    r_div.font.name = "Times New Roman"
    r_div.font.size = Pt(8)
    r_div.font.color.rgb = RGBColor(150, 150, 150)

    _add_slot(2)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    return output_path

create_dpt_karangjengkol_sample = generate_sample_dpt_excel

if __name__ == "__main__":
    base_templates = Path(__file__).resolve().parent.parent / "data" / "templates"
    excel_out = base_templates / "SAMPLE_DPT_DUMMY.xlsx"
    word_out = base_templates / "SAMPLE_SURAT_DUMMY.docx"

    generate_sample_dpt_excel(str(excel_out))
    print(f"Generated sample Excel: {excel_out}")

    generate_sample_word_dummy(str(word_out))
    print(f"Generated sample Word: {word_out}")
