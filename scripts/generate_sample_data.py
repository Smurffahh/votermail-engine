import os
import random
import openpyxl

FIRST_NAMES = [
    "Agus", "Bambang", "Budi", "Cahyo", "Dedi", "Eko", "Fajar", "Guruh", "Hadi", "Iwan",
    "Joko", "Kurniawan", "Lukman", "Mustofa", "Nanang", "Oki", "Prasetyo", "Rahmat", "Slamet", "Tri",
    "Untung", "Wahyu", "Yanto", "Zainal", "Anisa", "Citra", "Dewi", "Erna", "Fitri", "Gita",
    "Haryati", "Indah", "Kartika", "Lestari", "Maya", "Nurul", "Putri", "Ratna", "Siti", "Triana",
    "Utami", "Wulandari", "Yuni", "Zulfa", "Ahmad", "Aris", "Bayu", "Danang", "Endang", "Farida"
]

LAST_NAMES = [
    "Santoso", "Wijaya", "Kusuma", "Saputra", "Wibowo", "Hidayat", "Setiawan", "Pratama",
    "Nugroho", "Purnomo", "Utomo", "Suryono", "Handayani", "Rahayu", "Lestari", "Safitri",
    "Kurniawati", "Anggraeni", "Mahendra", "Firmansyah", "Gunawan", "Subekti", "Mulyono", "Hartono"
]

def create_dpt_karangjengkol_sample(file_path: str = "sample_dpt_karangjengkol.xlsx", voters_per_sheet: int = 25) -> str:
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    if default_sheet is not None:
        wb.remove(default_sheet)

    random.seed(42)
    sheets_info = [
        ("KADUS 1", "001"),
        ("KADUS 2", "002"),
        ("KADUS 3", "003"),
    ]

    for sheet_name, rw_code in sheets_info:
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

        # Skip row 11 containing column sequence numbers
        for c_idx in range(1, 13):
            ws.cell(11, c_idx, c_idx)

        curr_row = 12
        male_count = 0
        female_count = 0

        for v_idx in range(1, voters_per_sheet + 1):
            gender = "L" if random.random() > 0.45 else "P"
            if gender == "L":
                male_count += 1
            else:
                female_count += 1

            first = random.choice(FIRST_NAMES)
            last = random.choice(LAST_NAMES)
            full_name = f"{first} {last}".upper()

            age = random.randint(18, 75)
            b_day = random.randint(1, 28)
            b_month = random.randint(1, 12)
            b_year = 2026 - age
            b_date_str = f"{b_day:02d}|{b_month:02d}|{b_year}"

            rt_code = f"{(v_idx % 4) + 1:03d}"
            dummy_nik = f"330307{b_day:02d}{b_month:02d}{str(b_year)[2:]}000{v_idx:02d}"

            ws.cell(curr_row, 1, v_idx)
            ws.cell(curr_row, 2, dummy_nik)
            ws.cell(curr_row, 3, full_name)
            ws.cell(curr_row, 4, "PURBALINGGA")
            ws.cell(curr_row, 5, b_date_str)
            ws.cell(curr_row, 6, age)
            ws.cell(curr_row, 7, "S" if age > 25 else "B")
            ws.cell(curr_row, 8, gender)
            ws.cell(curr_row, 9, "KARANGJENGKOL")
            ws.cell(curr_row, 10, rt_code)
            ws.cell(curr_row, 11, rw_code)
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
        ws.cell(curr_row, 3, voters_per_sheet)
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

    wb.save(file_path)
    wb.close()
    return file_path

if __name__ == "__main__":
    out_file = create_dpt_karangjengkol_sample("sample_dpt_karangjengkol.xlsx")
    print(f"Sample DPT Karangjengkol berhasil dibuat: {out_file}")
