import pandas as pd
import os
import xlwings as xw
from xlwings.utils import col_name as get_col_name
from collections import defaultdict
import datetime

print("--- PROSES 4: PEMBARUAN KOLOM TANGGAL JT ---")

giro_file = 'Giro_temp.xlsx'
target_file = 'ARVIEWER.xlsm'
sheet_name = 'Source'

if not os.path.exists(giro_file):
    print(f"ERROR: File '{giro_file}' tidak ditemukan!")
    exit()
if not os.path.exists(target_file):
    print(f"ERROR: File '{target_file}' tidak ditemukan!")
    exit()

indo_months = {
    'Jan': 1,  'Feb': 2,  'Peb': 2,  'Mar': 3,  'Apr': 4,  'Mei': 5,  'Jun': 6,
    'Jul': 7,  'Agu': 8,  'Ags': 8,  'Sep': 9,  'Okt': 10, 'Nov': 11, 'Nop': 11, 'Des': 12,
    
    'jan': 1,  'feb': 2,  'mar': 3,  'apr': 4,  'mei': 5,  'jun': 6,
    'jul': 7,  'agu': 8,  'ags': 8,  'sep': 9,  'okt': 10, 'nov': 11, 'nop': 11, 'des': 12
}

def clean_invoice_str(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        return s[:-2]
    return s

print(f"Membaca data referensi dari {giro_file}...")
df_giro = pd.read_excel(giro_file)

giro_groups = defaultdict(list)
for _, row in df_giro.iterrows():
    no_faktur_so = clean_invoice_str(row['No. Faktur. (SO)'])
    tgl_cek = row['Tgl Cek']
    if no_faktur_so and not pd.isna(tgl_cek):
        giro_groups[no_faktur_so].append(tgl_cek)

mapping_jt = {}
for no_faktur, tgl_list in giro_groups.items():
    parsed_dates = []
    
    for tgl in tgl_list:
        if isinstance(tgl, (pd.Timestamp, datetime.datetime)):
            d_int, m_int, y_int = tgl.day, tgl.month, tgl.year
            d_str = str(d_int).zfill(2)
            m_short = str(m_int).zfill(2)
            y_short = str(y_int)[-2:]
            parsed_dates.append((y_int, m_int, d_int, d_str, m_short, y_short))
        else:
            parts = str(tgl).strip().split()
            if len(parts) == 3:
                try:
                    d_int = int(parts[0])
                    d_str = parts[0].zfill(2)
                    m_name = parts[1]
                    m_int = indo_months.get(m_name, 1)
                    m_short = str(m_int).zfill(2)
                    y_int = int(parts[2])
                    y_short = parts[2][-2:]
                    parsed_dates.append((y_int, m_int, d_int, d_str, m_short, y_short))
                except ValueError:
                    continue
                    
    if not parsed_dates:
        continue
        
    parsed_dates.sort(key=lambda x: (x[0], x[1], x[2]))
    
    date_by_month_year = defaultdict(list)
    for item in parsed_dates:
        y_int, m_int, d_int, d_str, m_short, y_short = item
        group_key = (y_int, m_int, m_short, y_short)
        if d_str not in date_by_month_year[group_key]:
            date_by_month_year[group_key].append(d_str)
            
    group_strings = []
    for key in sorted(date_by_month_year.keys()):
        y_int, m_int, m_short, y_short = key
        days_str = ",".join(date_by_month_year[key])
        group_strings.append(f"{days_str}/{m_short}/{y_short}")
        
    mapping_jt[no_faktur] = "JT " + " & ".join(group_strings)

full_path_target = os.path.abspath(target_file)
app = None

try:
    print("Membuka aplikasi Excel secara background...")
    app = xw.App(visible=False)
    app.display_alerts = False
    
    print(f"Membuka file {target_file}...")
    wb = app.books.open(full_path_target)
    ws = wb.sheets[sheet_name]
    
    last_cell = ws.range('B' + str(ws.cells.last_cell.row)).end('up')
    last_row = last_cell.row
    
    if last_row < 4:
        print("Peringatan: Tidak ditemukan data transaksi dari baris 4 ke bawah.")
        wb.close()
        exit()
        
    headers = ws.range('A3').expand('right').value
    if 'Tanggal JT' not in headers:
        print("ERROR: Judul kolom 'Tanggal JT' tidak ditemukan pada baris 3!")
        wb.close()
        exit()
        
    jt_col_idx = headers.index('Tanggal JT') + 1
    jt_col_letter = get_col_name(jt_col_idx)
    
    faktur_range = ws.range(f'B4:B{last_row}').value
    if not isinstance(faktur_range, list):
        faktur_range = [faktur_range]
        
    print("Mencocokkan nomor faktur dan menyusun nilai baru...")
    hasil_kolom_jt = []
    for fktr in faktur_range:
        fktr_cleaned = clean_invoice_str(fktr)
        nilai_jt = mapping_jt.get(fktr_cleaned, "")
        hasil_kolom_jt.append([nilai_jt])
        
    print(f"Menulis data rekapan ke kolom {jt_col_letter}...")
    ws.range(f'{jt_col_letter}4:{jt_col_letter}{last_row}').value = hasil_kolom_jt
    
    print("Menyimpan hasil perubahan file...")
    wb.save()
    wb.close()
    print("\nPROSES BERHASIL! Kolom 'Tanggal JT' berhasil diperbarui dengan aman.")
    
except Exception as e:
    print(f"TERJADI ERROR SAAT PROSES EXECUTION: {e}")
finally:
    if app:
        app.quit()
        print("Aplikasi Excel latar belakang berhasil ditutup.")