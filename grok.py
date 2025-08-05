import calendar
import pandas as pd
from datetime import datetime

# Fungsi untuk membuat data kalender untuk satu bulan
def get_month_data(year, month):
    cal = calendar.monthcalendar(year, month)
    month_name = calendar.month_name[month]
    data = []
    
    # Header hari
    days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    
    # Mengisi data untuk setiap minggu
    for week in cal:
        week_data = []
        for day in week:
            if day == 0:
                week_data.append('')  # Kosong untuk hari yang tidak ada
            else:
                week_data.append(day)
        data.append(week_data)
    
    return month_name, data, days

# Membuat list untuk menyimpan semua data kalender
all_data = []

# Mengumpulkan data untuk semua bulan di tahun 2025
for month in range(1, 13):
    month_name, month_data, days = get_month_data(2025, month)
    
    # Menambahkan nama bulan sebagai header
    all_data.append([f'{month_name} 2025'] + ['']*6)  # Header bulan
    all_data.append(days)  # Header hari
    
    # Menambahkan data tanggal
    for week in month_data:
        all_data.append(week)
    
    # Menambahkan baris kosong antar bulan
    all_data.append(['']*7)

# Membuat DataFrame
df = pd.DataFrame(all_data, columns=['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'])

# Menyimpan ke file Excel
output_file = 'kalendar_2025_grok.xlsx'
df.to_excel(output_file, index=False)

print(f"Kalender 2025 telah diekspor ke {output_file}")