import calendar
import pandas as pd

def buat_dan_ekspor_kalender_2025_ke_excel():
    """
    Membuat kalender tahun 2025 dan mengekspornya ke file Excel.
    Setiap bulan akan berada di sheet terpisah.
    """
    tahun = 2025
    nama_file_excel = f"Kalender_{tahun}_gemini.xlsx"

    # Membuat objek ExcelWriter untuk menulis ke file Excel
    with pd.ExcelWriter(nama_file_excel, engine='openpyxl') as writer:
        for bulan_num in range(1, 13):
            # Mendapatkan nama bulan
            nama_bulan = calendar.month_name[bulan_num]

            # Mendapatkan string kalender untuk bulan ini
            # calendar.month().splitlines() akan memecah string menjadi baris-baris
            # Kita buang baris pertama (nama bulan dan tahun) dan baris kedua (nama hari)
            # Karena kita akan membuat header sendiri
            cal_str_lines = calendar.month(tahun, bulan_num).splitlines()
            
            # Baris pertama adalah nama bulan dan tahun, baris kedua adalah nama hari (Sen, Sel, ...)
            # Kita ambil baris ketiga dan seterusnya (tanggal-tanggal)
            data_tanggal = [line.strip().split() for line in cal_str_lines[2:] if line.strip()]

            # Mendefinisikan nama-nama hari untuk header kolom
            # calendar.day_abbr mengembalikan singkatan hari (Mon, Tue, ...)
            # calendar.day_name mengembalikan nama hari lengkap (Monday, Tuesday, ...)
            # Kita bisa pilih mana yang lebih diinginkan. Di sini saya pakai singkatan.
            nama_hari_kolom = [calendar.day_abbr[i] for i in range(7)]

            # Membuat DataFrame Pandas
            df_bulan = pd.DataFrame(data_tanggal, columns=nama_hari_kolom)

            # Menyimpan DataFrame ke sheet di Excel
            # sheet_name adalah nama sheet di Excel
            df_bulan.to_excel(writer, sheet_name=nama_bulan, index=False)
            
            print(f"Bulan {nama_bulan} telah diekspor ke sheet '{nama_bulan}'.")

    print(f"\nKalender tahun {tahun} berhasil diekspor ke file '{nama_file_excel}'.")

if __name__ == "__main__":
    buat_dan_ekspor_kalender_2025_ke_excel()