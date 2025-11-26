import openpyxl
import os

FILE_NAME = 'data_beasiswa.xlsx'

def tampil_laporan_distribusi():
    if not os.path.exists(FILE_NAME):
        print("File data beasiswa tidak ditemukan.")
        return

    workbook = openpyxl.load_workbook(FILE_NAME)

    if 'Pemberian' not in workbook.sheetnames:
        print("Belum ada data pemberian beasiswa.")
        return

    sheet = workbook['Pemberian']

    if sheet.max_row == 1:
        print("Belum ada data distribusi beasiswa.")
        return

    print("\n=== Laporan Distribusi Beasiswa ===")
    print("NISN | Kode Beasiswa | Tanggal Penerimaan")
    print("------------------------------------------")

    for row in sheet.iter_rows(min_row=2, values_only=True):
        print(row)

def menu_laporan():
    while True:
        print("\nMenu Laporan:")
        print("1. Laporan Distribusi Beasiswa")
        print("2. Kembali ke Menu Utama")

        pilihan = input("Pilih menu: ")

        if pilihan == '1':
            tampil_laporan_distribusi()
        elif pilihan == '2':
            break
        else:
            print("Pilihan tidak valid.")