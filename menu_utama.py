from menu_beasiswa import menu_beasiswa
from menu_siswa import menu_siswa
from menu_pemberian import menu_pemberian
from menu_laporan import menu_laporan

# Menu utama
def menu_utama():  # fungsi untuk menampilkan menu utama
    while True:  # jika While True, maka program akan terus berjalan hingga dihentikan 
        print("\n=== MENU UTAMA SISTEM PEMBERIAN BEASISWA ===")
        print("1. Data Siswa Penerima Beasiswa")
        print("2. Data Beasiswa")
        print("3. Data Pemberian Beasiswa")
        print("4. Data Laporan Beasiswa")
        print("5. Keluar")

        # mengambil input dari pengg1una
        pilihan = input("Pilih menu (1-5): ")

        # mengarahkan ke menu sesuai pilihan pengguna
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