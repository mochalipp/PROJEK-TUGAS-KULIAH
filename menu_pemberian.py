import openpyxl
import os
from datetime import datetime

FILE_NAME = 'data_beasiswa.xlsx'

def create_sheet_if_not_exists(workbook, sheet_name, header=None):
    if sheet_name not in workbook.sheetnames:
        sheet = workbook.create_sheet(sheet_name)
        if header:
            sheet.append(header)
    return workbook[sheet_name]

def pemberian_beasiswa():
    nisn = input("Masukkan NISN Siswa: ")
    kode = input("Masukkan Kode Beasiswa: ")
    tanggal = datetime.today().strftime("%Y-%m-%d")

    if not os.path.exists(FILE_NAME):
        print("File data tidak ditemukan.")
        return

    workbook = openpyxl.load_workbook(FILE_NAME)

    if 'Siswa' not in workbook.sheetnames or 'Beasiswa' not in workbook.sheetnames:
        print("Data siswa atau beasiswa tidak ada.")
        return

    siswa_sheet = workbook['Siswa']
    bea_sheet = workbook['Beasiswa']

    # Validasi siswa
    siswa_valid = any(row[0].value == nisn for row in siswa_sheet.iter_rows(min_row=2))
    if not siswa_valid:
        print("Siswa tidak terdaftar.")
        return

    # Validasi beasiswa & kuota
    bea_row = None
    for row in bea_sheet.iter_rows(min_row=2):
        if row[0].value == kode:
            bea_row = row
            break

    if bea_row is None:
        print("Beasiswa tidak ditemukan.")
        return

    kuota = int(bea_row[5].value)

    if kuota <= 0:
        print("Kuota beasiswa HABIS!")
        return

    # Simpan ke sheet pemberian
    pemberian_sheet = create_sheet_if_not_exists(
        workbook, 'Pemberian', ['NISN', 'Kode Beasiswa', 'Tanggal']
    )
    pemberian_sheet.append([nisn, kode, tanggal])

    # Kurangi kuota
    kuota -= 1
    bea_row[5].value = kuota

    # Update status jika kuota habis
    if kuota == 0:
        bea_row[6].value = "Habis"
    else:
        bea_row[6].value = "Tersedia"

    workbook.save(FILE_NAME)
    print("Beasiswa berhasil diberikan!")

# TAMPIL PEMBERIAN
def tampil_data_pemberian():
    if not os.path.exists(FILE_NAME):
        print("File tidak ditemukan.")
        return

    workbook = openpyxl.load_workbook(FILE_NAME)

    if 'Pemberian' not in workbook.sheetnames:
        print("Belum ada pemberian.")
        return

    sheet = workbook['Pemberian']

    print("\n=== DATA PEMBERIAN BEASISWA ===")
    for row in sheet.iter_rows(min_row=1, values_only=True):
        print("{:<15} {:<20} {:<15}".format(*row))
        print("-" * 50)
        
# MENU PEMBERIAN
def menu_pemberian():
    while True:
        print("\n=== MENU PEMBERIAN BEASISWA ===")
        print("1. Tambahkan Pemberian Beasiswa")
        print("2. Tampilkan Pemberian Beasiswa")
        print("3. Kembali ke Menu Utama")

        pilihan = input("Pilih menu: ")

        if pilihan == '1': 
            pemberian_beasiswa()
        elif pilihan == '2': 
            tampil_data_pemberian()
        elif pilihan == '3': 
            break
        else: 
            print("Pilihan tidak valid!")