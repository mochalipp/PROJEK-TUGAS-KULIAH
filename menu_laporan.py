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
    for row in sheet.iter_rows(min_row=1, values_only=True):
        print("{:<15} {:<20} {:<15} {:<30}".format(*row))
        print("-" * 80)

def menu_laporan():
    while True:
        print("\n === MENU LAPORAN ===")
        print("1. Laporan Distribusi Beasiswa")
        print("2. Kembali ke Menu Utama")

        pilihan = input("Pilih menu: ")

        if pilihan == '1':
            tampil_laporan_distribusi()
        elif pilihan == '2':
            break
        else:
            print("Pilihan tidak valid.")