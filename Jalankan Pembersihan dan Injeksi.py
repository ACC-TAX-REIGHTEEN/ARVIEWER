import os
import shutil
import subprocess
import glob
import sys

dapur_dir = "Dapur"

required_root_files = [
    "ExportFile.xls", 
    "Giro.xls", 
    "ARVIEWER.xlsm"
]

required_dapur_files = [
    "__init__.py",
    "1_CleaningMovingALLAR.py",
    "2_AddGiroDate.py",
    "3_InjectGiroDtl2SS.py"
]

scripts_to_run = [
    "1_CleaningMovingALLAR.py",
    "2_AddGiroDate.py",
    "3_InjectGiroDtl2SS.py"
]

print("--- PROSES UTAMA: ORKESTRASI & EKSEKUSI DATA ---")

print("--> Memeriksa ketersediaan file data utama...")
missing_root_files = []
for f in required_root_files:
    if not os.path.exists(f):
        missing_root_files.append(f)

if missing_root_files:
    print("\n[GAGAL] Proses dihentikan. File data berikut tidak ditemukan di folder utama:")
    for item in missing_root_files:
        print(f"   - {item}")
    print("Silakan lengkapi file di atas terlebih dahulu.")
    input("\n--> Tekan Enter untuk keluar...")
    sys.exit()

print("--> Memeriksa ketersediaan skrip di folder Dapur...")
missing_dapur_items = []
if not os.path.exists(dapur_dir):
    missing_dapur_items.append(dapur_dir)
else:
    for f in required_dapur_files:
        file_path = os.path.join(dapur_dir, f)
        if not os.path.exists(file_path):
            missing_dapur_items.append(file_path)

if missing_dapur_items:
    print("\n[GAGAL] Proses dihentikan. Folder 'Dapur' atau skrip di dalamnya tidak lengkap:")
    for item in missing_dapur_items:
        print(f"   - {item}")
    print("Silakan pastikan semua skrip pembersihan dan injeksi berada di dalam folder Dapur.")
    input("\n--> Tekan Enter untuk keluar...")
    sys.exit()

print("--> Memindahkan file data utama ke folder Dapur untuk diproses...")
try:
    for f in required_root_files:
        src_path = f
        dest_path = os.path.join(dapur_dir, f)
        
        if os.path.exists(dest_path):
            os.remove(dest_path)
            
        shutil.move(src_path, dest_path)
        print(f"    * Berhasil memindahkan {f} -> {dapur_dir}/")
except Exception as e:
    print(f"\n[ERROR] Gagal memindahkan file ke folder Dapur: {e}")
    input("\n--> Tekan Enter untuk keluar...")
    sys.exit()

current_dir = os.getcwd()
os.chdir(dapur_dir)

try:
    for script in scripts_to_run:
        print(f"\n--> Menjalankan skrip: {script}...")
        subprocess.run([sys.executable, script], check=True)
except subprocess.CalledProcessError:
    print(f"\n[ERROR] Terjadi kesalahan saat menjalankan {script}. Proses dihentikan otomatis.")
    os.chdir(current_dir)
    input("\n--> Tekan Enter untuk keluar...")
    sys.exit()

os.chdir(current_dir)

print("\n--> Mengembalikan file hasil pembaruan ke folder utama...")
arviewer_in_dapur = os.path.join(dapur_dir, "ARVIEWER.xlsm")

if os.path.exists(arviewer_in_dapur):
    if os.path.exists("ARVIEWER.xlsm"):
        os.remove("ARVIEWER.xlsm")
    shutil.move(arviewer_in_dapur, "ARVIEWER.xlsm")
    print("    * File ARVIEWER.xlsm berhasil dipindahkan kembali ke folder utama.")
else:
    print("\n[PERINGATAN] File ARVIEWER.xlsm tidak ditemukan di folder Dapur setelah proses selesai!")

print("--> Membersihkan file sisa (.xls, .xlsx, .xlsm) di folder Dapur...")
for ext in ['*.xls', '*.xlsx', '*.xlsm']:
    for file in glob.glob(os.path.join(dapur_dir, ext)):
        try:
            os.remove(file)
            print(f"    * Menghapus file sisa: {os.path.basename(file)}")
        except Exception as e:
            print(f"    * Gagal menghapus file sisa {file}: {e}")

print("\n=== SEMUA PROSES BERHASIL DISELESAIKAN DENGAN AMAN! ===")
input("--> Tekan Enter untuk menutup jendela...")