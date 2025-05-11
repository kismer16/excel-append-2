import pandas as pd
import glob
import openpyxl

# Az Excel fájlok elérési útja
file_path = 'H:/dp_pdf/innpro/keszteszt/merge/*.xlsx'  # Cseréld le a megfelelő elérési útra



# Az összes Excel fájl beolvasása
all_files = glob.glob(file_path)

# Lista a DataFrame-ek tárolására
dataframes = []

# Minden fájl beolvasása és hozzáadása a listához
for file in all_files:
    df = pd.read_excel(file,dtype=str)

    dataframes.append(df)

# Az összes DataFrame összefűzése
combined_df = pd.concat(dataframes, ignore_index=True)

# Az összefűzött DataFrame mentése egy új Excel fájlba
combined_df.to_excel('H:/dp_pdf/innpro/keszteszt/merge/combined_file.xlsx', index=False)