import pandas as pd
import os
import xlwings as xw
import time

from xlwings.utils import col_name as get_col_name 

print(f"--- PROSES 1: PEMBERSIHAN DATA (CLEANING) ---")

file_path = 'ExportFile.xls'

if not os.path.exists(file_path):
    print(f"File '{file_path}' tidak ditemukan. Pastikan file ada.")
    exit()

def load_dataset(path):
    try:
        return pd.read_excel(path, header=3)
    except:
        pass
    try:
        return pd.read_csv(path, header=3, sep=',', engine='python', encoding='ISO-8859-1')
    except:
        return None

df = load_dataset(file_path)

if df is None:
    print("Gagal membaca file ExportFile.xls.")
    exit()

target_indices = [2, 3, 4, 7, 8, 10, 11, 12, 13, 14]
df_clean = df.iloc[:, target_indices].copy()

new_columns = [
    'Kode Pelanggan', 
    'No. Faktur',     
    'Tgl Faktur',     
    'Jatuh Tempo',    
    'Nilai Faktur',   
    'Sisa Piutang',   
    'Umur JT',        
    'Nama Pelanggan', 
    'Nama Penjual',   
    'Nama Kontak'    
]
df_clean.columns = new_columns

df_clean['Kode Pelanggan'] = df_clean['Kode Pelanggan'].ffill()
df_clean = df_clean.dropna(subset=['No. Faktur'])

def format_clean(val):
    if pd.isna(val):
        return ""
    s = str(val)
    if s.endswith('.0'):
        return s[:-2]
    if s.endswith(',00'):
        return s[:-3]
    return s

cols_to_clean = ['Kode Pelanggan', 'Nilai Faktur', 'Sisa Piutang']
for col in cols_to_clean:
    df_clean[col] = df_clean[col].apply(format_clean)

indo_months_in = {
    'Jan': 'Jan', 'Feb': 'Feb', 'Mar': 'Mar', 'Apr': 'Apr', 'Mei': 'May', 'Jun': 'Jun',
    'Jul': 'Jul', 'Agu': 'Aug', 'Sep': 'Sep', 'Okt': 'Oct', 'Nop': 'Nov', 'Des': 'Dec'
}

def clean_and_format_date(x):
    if pd.isna(x) or str(x).strip() == '':
        return None
    date_str = str(x)
    for indo, eng in indo_months_in.items():
        if indo in date_str:
            date_str = date_str.replace(indo, eng)
            break
    try:
        return pd.to_datetime(date_str, dayfirst=True, errors='coerce')
    except:
        return None

df_clean['Tgl Faktur'] = df_clean['Tgl Faktur'].apply(clean_and_format_date)
df_clean['Jatuh Tempo'] = df_clean['Jatuh Tempo'].apply(clean_and_format_date)

if 'Tgl Faktur' in df_clean.columns:
    idx_tgl = df_clean.columns.get_loc('Tgl Faktur')
    df_clean.insert(idx_tgl + 1, 'SS', '')

if 'Jatuh Tempo' in df_clean.columns:
    idx_jt = df_clean.columns.get_loc('Jatuh Tempo')
    df_clean.insert(idx_jt + 1, 'SS', '', allow_duplicates=True)

df_clean.reset_index(drop=True, inplace=True)

output_filename = f"EXPORT_Sementara.xlsx"
df_clean.to_excel(output_filename, index=False)
print(f"Data bersih siap ({len(df_clean)} baris).")


print(f"\n--- PROSES 2: UPDATE ARVIEWER.xlsm ---")
target_file = 'ARVIEWER.xlsm'
sheet_name = 'Source'
full_path_target = os.path.abspath(target_file)

if not os.path.exists(full_path_target):
    print(f"ERROR: File {target_file} tidak ditemukan!")
    exit()

app = None
try:
    print("Membuka aplikasi Excel...")
    app = xw.App(visible=False) 
    app.display_alerts = False 
    
    print(f"Membuka workbook {target_file}...")
    wb = app.books.open(full_path_target)
    
    try:
        ws = wb.sheets[sheet_name]
    except:
        print(f"Sheet '{sheet_name}' tidak ditemukan!")
        wb.close()
        app.quit()
        exit()
        
    last_cell = ws.range('A' + str(ws.cells.last_cell.row)).end('up')
    last_row_existing = last_cell.row
    
    if last_row_existing >= 4:
        print(f"Membersihkan data lama...")
        ws.range(f'A4:K{last_row_existing}').clear_contents()
        
    print("Menyalin data baru...")
    if len(df_clean) > 0:
        ws.range('A4').options(index=False, header=False).value = df_clean
        
        print("Menerapkan format tanggal Indonesia...")
        new_last_row = 4 + len(df_clean) - 1
        
        for idx, header_text in enumerate(df_clean.columns):
            if 'Tgl' in str(header_text) or 'Jatuh Tempo' in str(header_text):
            	
                excel_col_num = 1 + idx
                excel_col_letter = get_col_name(excel_col_num) 
                
                rng = ws.range(f"{excel_col_letter}4:{excel_col_letter}{new_last_row}")
                rng.number_format = '[$-421]dd mmm yyyy;@'

    print("Data berhasil disalin dan diformat.")
    
    wb.save()
    wb.close()
    print("File disimpan.")

except Exception as e:
    print(f"TERJADI ERROR: {e}")
finally:
    if app:
        app.quit()
        print("Aplikasi Excel ditutup.")

print(f"\n--- PROSES 3: MENGHAPUS FILE SEMENTARA ---")
time.sleep(1)
files_to_remove = [file_path, output_filename]

for f in files_to_remove:
    try:
        if os.path.exists(f):
            os.remove(f)
            print(f"Berhasil menghapus: {f}")
    except Exception as e:
        print(f"Gagal menghapus {f}: {e}")

print("\nSEMUA PROSES SELESAI!")    
