import pandas as pd
import configparser
from openpyxl.utils import get_column_letter

print("--> Membaca konfigurasi dari config.conf")

config = configparser.ConfigParser()
config.read('config.conf')

kontak_filters = [x.strip() for x in config.get('FILTER', 'Nama Kontak', fallback='').split(',')]
penjual_filters = [x.strip() for x in config.get('FILTER', 'Nama Penjual', fallback='').split(',')]

print("--> Mengekstrak data dari ARVIEWER.xlsm")

df = pd.read_excel('ARVIEWER.xlsm', sheet_name='Source', header=2, usecols="B:L")
df.columns = ['No. Faktur', 'Tgl Faktur', 'AA', 'Jatuh Tempo', 'SS', 'Nilai Faktur', 'Sisa Piutang', 'Umur JT', 'Nama Pelanggan', 'Nama Penjual', 'Nama Kontak']

if kontak_filters and kontak_filters[0]:
    df = df[df['Nama Kontak'].fillna('').astype(str).str.contains('|'.join(kontak_filters), case=False, na=False)]

if penjual_filters and penjual_filters[0]:
    df = df[df['Nama Penjual'].fillna('').astype(str).str.contains('|'.join(penjual_filters), case=False, na=False)]

print("--> Mengurutkan data")

df = df.sort_values(by=['Nama Pelanggan', 'No. Faktur'])

print("--> Memproses data untuk Sheet Lookup")

df['Nilai Faktur'] = pd.to_numeric(df['Nilai Faktur'], errors='coerce').fillna(0)
df['Sisa Piutang'] = pd.to_numeric(df['Sisa Piutang'], errors='coerce').fillna(0)

lookup_data = []
for pelanggan, group in df.groupby('Nama Pelanggan', sort=False):
    for _, row in group.iterrows():
        lookup_data.append(row.to_dict())
    
    subtotal = {col: '' for col in df.columns}
    subtotal['Nilai Faktur'] = group['Nilai Faktur'].sum()
    subtotal['Sisa Piutang'] = group['Sisa Piutang'].sum()
    
    lookup_data.append(subtotal)
    lookup_data.append({col: '' for col in df.columns})

df_lookup = pd.DataFrame(lookup_data)

print("--> Memformat tanggal")

def format_tanggal(val):
    try:
        if pd.isna(val) or val == '':
            return ''
        if isinstance(val, (int, float)):
            return pd.to_datetime(val, unit='D', origin='1899-12-30').strftime('%d %b %Y')
        return pd.to_datetime(val).strftime('%d %b %Y')
    except Exception:
        return val

df['Tgl Faktur'] = df['Tgl Faktur'].apply(format_tanggal)
df['Jatuh Tempo'] = df['Jatuh Tempo'].apply(format_tanggal)

df_lookup['Tgl Faktur'] = df_lookup['Tgl Faktur'].apply(format_tanggal)
df_lookup['Jatuh Tempo'] = df_lookup['Jatuh Tempo'].apply(format_tanggal)

print("--> Menyimpan hasil ke Hasil_Analisis.xlsx")

with pd.ExcelWriter('Hasil_Analisis_JT_IRCZN.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Source', index=False)
    df_lookup.to_excel(writer, sheet_name='Lookup', index=False)
    
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for col_idx, col in enumerate(worksheet.columns, 1):
            col_letter = get_column_letter(col_idx)
            col_header = str(col[0].value)
            
            if col_header == 'Keterangan':
                worksheet.column_dimensions[col_letter].width = 25
            else:
                max_length = 0
                for cell in col:
                    try:
                        if cell.value is not None:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                worksheet.column_dimensions[col_letter].width = max_length + 2
            
            if col_header in ['Nilai Faktur', 'Sisa Piutang']:
                for cell in col[1:]:
                    if cell.value != '' and cell.value is not None:
                        cell.number_format = '#,##0'

print("--> Seluruh proses berhasil diselesaikan")