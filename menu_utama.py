from menu_beasiswa import menu_beasiswa
from menu_siswa import menu_siswa
from menu_pemberian import menu_pemberian
from menu_laporan import menu_laporan

# Menu utama
def menu_utama():
    while True:
        print("\n=== MENU " \
        "UTAMA SISTEM PEMBERIAN BEASISWA ===")
        print("1. Data Siswa Berhak Beasiswa")
        print("2. Jenis Beasiswa")
        print("3. Pemberian Beasiswa")
        print("4. Laporan Beasiswa")
        print("5. Keluar")

        pilihan = input("Pilih menu (1-5): ")

        if pilihan == '1':
            menu_siswa()
        elif pilihan == '2':
            menu_beasiswa()
        elif pilihan == '3':
            menu_pemberian()
        elif pilihan == '4':
            menu_laporan()
        elif pilihan == '5':
            print("Terima kasih telah menggunakan sistem pemberian beasiswa.")
            break
        else:
            print("Input tidak valid! Silakan pilih menu 1–5.")

if __name__ == "__main__":
    menu_utama()