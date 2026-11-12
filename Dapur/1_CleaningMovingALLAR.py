import os
import sys
import time
import warnings
import pandas as pd
import xlwings as xw
from xlwings.utils import col_name as get_col_name


def load_dataset(path):
    try:
        return pd.read_excel(path, header=3, engine='calamine')
    except Exception:
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            try:
                return pd.read_excel(path, header=3, engine='openpyxl')
            except Exception:
                try:
                    return pd.read_excel(path, header=3, engine='xlrd')
                except Exception:
                    return None


def format_clean(val):
    if pd.isna(val):
        return ""
    s = str(val)
    if s.endswith('.0'):
        return s[:-2]
    if s.endswith(',00'):
        return s[:-3]
    return s


def clean_and_format_date(x):
    indo_months_in = {
        'Jan': 'Jan', 'Feb': 'Feb', 'Mar': 'Mar', 'Apr': 'Apr', 'Mei': 'May', 'Jun': 'Jun',
        'Jul': 'Jul', 'Agu': 'Aug', 'Sep': 'Sep', 'Okt': 'Oct', 'Nop': 'Nov', 'Des': 'Dec',
        'Peb': 'Feb', 'Ags': 'Aug',
        'jan': 'Jan', 'feb': 'Feb', 'mar': 'Mar', 'apr': 'Apr', 'mei': 'May', 'jun': 'Jun',
        'jul': 'Jul', 'agu': 'Aug', 'ags': 'Aug', 'sep': 'Sep', 'okt': 'Oct', 'nop': 'Nov',
        'nov': 'Nov', 'des': 'Dec'
    }
    if pd.isna(x) or str(x).strip() == '':
        return None
    date_str = str(x)
    for indo, eng in indo_months_in.items():
        if indo in date_str:
            date_str = date_str.replace(indo, eng)
            break
    try:
        return pd.to_datetime(date_str, dayfirst=True, errors='coerce')
    except Exception:
        return None


print("--> PROSES 1: PEMBERSIHAN DATA (CLEANING)")

file_path = 'Piutang.xls'

if not os.path.exists(file_path):
    print(f"--> File '{file_path}' tidak ditemukan. Pastikan file ada.")
    sys.exit()

df = load_dataset(file_path)

if df is None:
    print("--> Gagal membaca file Piutang.xls.")
    sys.exit()

target_indices = [2, 3, 5, 9, 11, 14, 16, 18, 20, 22]
df_clean = df.iloc[:, target_indices].copy()

df_clean.columns = [
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

df_clean['Kode Pelanggan'] = df_clean['Kode Pelanggan'].ffill()
df_clean = df_clean.dropna(subset=['No. Faktur'])

for col in ['Kode Pelanggan', 'Nilai Faktur', 'Sisa Piutang']:
    df_clean[col] = df_clean[col].apply(format_clean)

df_clean['Tgl Faktur'] = df_clean['Tgl Faktur'].apply(clean_and_format_date)
df_clean['Jatuh Tempo'] = df_clean['Jatuh Tempo'].apply(clean_and_format_date)

if 'Tgl Faktur' in df_clean.columns:
    idx_tgl = df_clean.columns.get_loc('Tgl Faktur')
    df_clean.insert(idx_tgl + 1, 'SS', '')

if 'Jatuh Tempo' in df_clean.columns:
    idx_jt = df_clean.columns.get_loc('Jatuh Tempo')
    df_clean.insert(idx_jt + 1, 'SS', '', allow_duplicates=True)

df_clean = df_clean.reset_index(drop=True)

output_filename = "Piutang_temp.xlsx"
df_clean.to_excel(output_filename, index=False)
print(f"--> Data bersih siap ({len(df_clean)} baris).")

print("--> PROSES 2: UPDATE ARVIEWER.xlsm")
target_file = 'ARVIEWER.xlsm'
sheet_name = 'Source'
full_path_target = os.path.abspath(target_file)

if not os.path.exists(full_path_target):
    print(f"--> ERROR: File {target_file} tidak ditemukan!")
    sys.exit()

app = None
try:
    print("--> Membuka aplikasi Excel...")
    app = xw.App(visible=False)
    app.display_alerts = False

    print(f"--> Membuka workbook {target_file}...")
    wb = app.books.open(full_path_target)

    try:
        ws = wb.sheets[sheet_name]
    except Exception:
        print(f"--> Sheet '{sheet_name}' tidak ditemukan!")
        wb.close()
        app.quit()
        sys.exit()

    last_cell = ws.range('A' + str(ws.cells.last_cell.row)).end('up')
    last_row_existing = last_cell.row

    if last_row_existing >= 4:
        print("--> Membersihkan data lama...")
        ws.range(f'A4:M{last_row_existing}').clear_contents()

    print("--> Menyalin data baru...")
    if len(df_clean) > 0:
        ws.range('A4').options(index=False, header=False).value = df_clean

        print("--> Menerapkan format tanggal Indonesia...")
        new_last_row = 4 + len(df_clean) - 1

        for idx, header_text in enumerate(df_clean.columns):
            if 'Tgl' in str(header_text) or 'Jatuh Tempo' in str(header_text):
                excel_col_num = 1 + idx
                excel_col_letter = get_col_name(excel_col_num)
                rng = ws.range(f"{excel_col_letter}4:{excel_col_letter}{new_last_row}")
                rng.number_format = '[$-421]dd mmm yyyy;@'

    print("--> Data berhasil disalin dan diformat.")

    wb.save()
    wb.close()
    print("--> File disimpan.")

except Exception as e:
    print(f"--> TERJADI ERROR: {e}")
finally:
    if app:
        app.quit()
        print("--> Aplikasi Excel ditutup.")

print("--> PROSES 3: MENGHAPUS FILE SEMENTARA")
time.sleep(1)

for f in [file_path, output_filename]:
    try:
        if os.path.exists(f):
            os.remove(f)
            print(f"--> Berhasil menghapus: {f}")
    except Exception as e:
        print(f"--> Gagal menghapus {f}: {e}")

print("--> SEMUA PROSES SELESAI!")