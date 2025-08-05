import calendar
import pandas as pd

# Tahun kalender
year = 2025

# Membuat dictionary untuk menyimpan data kalender
calendar_data = {"Bulan": [], "Tanggal": [], "Hari": []}

# Membuat kalender untuk setiap bulan dan setiap hari
for month in range(1, 13):
    month_calendar = calendar.monthcalendar(year, month)
    for week in month_calendar:
        for day in week:
            if day != 0:
                calendar_data["Bulan"].append(month)
                calendar_data["Tanggal"].append(day)
                # Mendapatkan nama hari
                day_name = calendar.day_name[calendar.weekday(year, month, day)]
                calendar_data["Hari"].append(day_name)

# Membuat DataFrame dari data kalender
df_calendar = pd.DataFrame(calendar_data)

# Menyimpan DataFrame ke file Excel
excel_filename = "kalender_2025_perplexity.xlsx"
df_calendar.to_excel(excel_filename, index=False)
