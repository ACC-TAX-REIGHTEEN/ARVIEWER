# 📊 ARVIEWER — Data Refresh Pipeline

> **Pembaruan data AR dashboard Excel otomatis — dari ekspor Accurate & rekap Giro ke ARVIEWER.xlsm siap pakai dalam sekali klik**

Pipeline Python tiga langkah yang membaca ekspor AR dari Accurate (`Piutang.xls`) dan rekap pembayaran Giro (`Giro.xls`), membersihkan dan menstrukturisasi data, lalu menyuntikkannya langsung ke sheet `Source` di **`ARVIEWER.xlsm`** — sebuah workbook Excel macro-enabled yang menjadi dashboard/viewer piutang tim — termasuk pengisian otomatis kolom `Tanggal JT` berdasarkan pencocokan nomor faktur dengan tanggal cek giro yang ada.

---

## 📋 Daftar Isi

- [Gambaran Umum](#-gambaran-umum)
- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Kerja Pipeline](#-alur-kerja-pipeline)
- [Detail Tiap Skrip](#-detail-tiap-skrip)
  - [Skrip 1 — Cleaning & Inject ke Source](#skrip-1--cleaning--inject-ke-source)
  - [Skrip 2 — Bersihkan Data Giro](#skrip-2--bersihkan-data-giro)
  - [Skrip 3 — Inject Tanggal JT ke ARVIEWER](#skrip-3--inject-tanggal-jt-ke-arviewer)
- [Format Data Input](#-format-data-input)
- [Output: Perubahan pada ARVIEWER.xlsm](#-output-perubahan-pada-arviewerxlsm)
- [Format Tanggal JT (Giro)](#-format-tanggal-jt-giro)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## 🗂️ Gambaran Umum

`ARVIEWER.xlsm` adalah dashboard Excel macro-enabled yang digunakan tim untuk melihat, memfilter, dan menganalisis data piutang. Agar informasinya selalu terkini, datanya perlu diperbarui secara berkala dari dua sumber:

1. **`Piutang.xls`** — ekspor laporan AR dari Accurate (berisi faktur, saldo, tanggal jatuh tempo, nama pelanggan, dst.)
2. **`Giro.xls`** — rekap pembayaran giro/cek yang diterima (berisi nomor faktur, bank, tanggal cair)

Pipeline ini mengotomasi seluruh proses pembaruan tersebut: membersihkan kedua file sumber, menyuntikkan data bersih ke sheet `Source` di dalam `ARVIEWER.xlsm`, dan mengisi kolom `Tanggal JT` dengan tanggal giro yang dicocokkan per nomor faktur — semua tanpa membuka Excel secara manual.

---

## ✨ Fitur Utama

- **Satu perintah, tiga proses** — Cukup jalankan satu orkestrator; ketiga skrip pembersihan dan injeksi berjalan berurutan secara otomatis.
- **Seleksi kolom dari ekspor Accurate** — Dari puluhan kolom di `Piutang.xls`, hanya 10 kolom relevan yang diambil berdasarkan indeks posisi, dengan penamaan ulang yang rapi.
- **Forward-fill kode pelanggan** — Menangani sel tergabung (merged cell) di kolom `Kode Pelanggan` pada ekspor Accurate dengan mengisi ke bawah secara otomatis.
- **Normalisasi bulan Indonesia** — Mengonversi nama bulan bahasa Indonesia (Jan, Peb, Mar, Mei, Agu, Okt, Nop, Des) ke format yang dapat diparsing Python.
- **Format tanggal Indonesia di Excel** — Setelah data disuntikkan, kolom tanggal diformat dengan locale Indonesia (`[$-421]dd mmm yyyy`) via xlwings sehingga tampilan di ARVIEWER tetap konsisten.
- **Pencocokan Giro per faktur** — Tanggal cek giro dicocokkan ke nomor faktur, diurutkan kronologis, dikelompokkan per bulan, dan ditulis ke kolom `Tanggal JT` dalam format ringkas yang mudah dibaca.
- **Penanganan multi-giro** — Satu faktur dengan lebih dari satu tanggal cek (giro cicilan/bertahap) ditangani otomatis: semua tanggal digabung dengan pemisah `&`.
- **Injeksi langsung ke .xlsm via xlwings** — Data dimasukkan ke workbook macro-enabled tanpa merusak VBA, macro, atau format yang ada di sheet-sheet lain.
- **Auto-cleanup menyeluruh** — File sumber (`Piutang.xls`, `Giro.xls`) dan file sementara dihapus otomatis setelah proses selesai; `ARVIEWER.xlsm` dikembalikan ke folder utama.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl xlwings xlrd
```

| Library | Digunakan di | Kegunaan |
|---|---|---|
| `pandas` | Skrip 1, 2, 3 | Baca `.xls`, bersihkan, transformasi data |
| `xlwings` | Skrip 1, 3 | Buka/tulis `.xlsm` dan terapkan format via Excel COM |
| `openpyxl` | Skrip 1 (via pandas) | Engine penulisan `.xlsx` sementara |
| `xlrd` | Skrip 1, 2 | Baca file legacy `.xls` dari Accurate |
| `collections`, `datetime`, `os`, `shutil`, `subprocess`, `glob`, `time` | Semua | Standard library |

### Aplikasi wajib
- **Microsoft Excel** — **wajib terinstall** di komputer yang menjalankan skrip ini. `xlwings` beroperasi dengan memanggil Excel secara background (COM automation) untuk membuka, menulis, dan menyimpan `.xlsm`. Tanpa Excel, proses akan gagal di Skrip 1 dan Skrip 3.

> **Catatan `xlrd`:** Untuk membaca `.xls` (format Excel lama), pastikan versi xlrd yang terinstall kompatibel:
> ```bash
> pip install "xlrd>=1.0.0,<2.0.0"
> ```

---

## 📁 Struktur Folder & File

```
📦 ARVIEWER/
│
├── 📄 Jalankan Pembersihan dan Injeksi.py ← Orkestrator utama. Jalankan ini
│
├── 📄 Piutang.xls           ← [INPUT] Ekspor AR dari Accurate (taruh di sini)
├── 📄 Giro.xls                 ← [INPUT] Rekap giro/cek masuk (taruh di sini)
├── 📄 ARVIEWER.xlsm            ← [INPUT+OUTPUT] Workbook dashboard (taruh di sini)
│
└── 📁 Dapur/                   ← Folder pipeline (jangan diubah strukturnya)
    ├── 📄 __init__.py
    ├── 📄 1_CleaningMovingALLAR.py  ← Bersihkan ExportFile + inject ke Source sheet
    ├── 📄 2_AddGiroDate.py          ← Bersihkan Giro.xls
    └── 📄 3_InjectGiroDtl2SS.py     ← Cocokkan & inject Tanggal JT ke ARVIEWER
```

> ⚠️ **`ARVIEWER.xlsm` bersifat input sekaligus output** — file ini dimodifikasi in-place (datanya diperbarui), bukan diganti atau dibuat ulang. Semua macro, formula, dan sheet lain di dalamnya tetap utuh.

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan tiga file input

Letakkan ketiga file berikut di folder utama (sejajar dengan `Jalankan Pembersihan dan Injeksi.py`):

| File | Sumber |
|---|---|
| `Piutang.xls` | Export laporan AR dari Accurate (nama **harus persis**) |
| `Giro.xls` | Rekap pembayaran giro/cek dari tim keuangan (nama **harus persis**) |
| `ARVIEWER.xlsm` | Workbook dashboard yang akan diperbarui (nama **harus persis**) |

> ⚠️ **Nama file bersifat eksak dan case-sensitive** — jangan ganti nama ketiga file di atas.

### Langkah 2 — Pastikan Excel tidak membuka ARVIEWER.xlsm

Tutup `ARVIEWER.xlsm` jika sedang terbuka di Excel. File yang sedang terbuka tidak dapat dimodifikasi oleh xlwings dan akan menyebabkan error.

### Langkah 3 — Jalankan

```bash
python "Jalankan Pembersihan dan Injeksi.py"
```

atau klik dua kali file tersebut jika Python sudah terasosiasi di sistem.

### Langkah 4 — Pantau progress

```
--- PROSES UTAMA: ORKESTRASI & EKSEKUSI DATA ---
--> Memeriksa ketersediaan file data utama...
--> Memeriksa ketersediaan skrip di folder Dapur...
--> Memindahkan file data utama ke folder Dapur untuk diproses...
    * Berhasil memindahkan Piutang.xls -> Dapur/
    * Berhasil memindahkan Giro.xls -> Dapur/
    * Berhasil memindahkan ARVIEWER.xlsm -> Dapur/

--> Menjalankan skrip: 1_CleaningMovingALLAR.py...
--- PROSES 1: PEMBERSIHAN DATA (CLEANING) ---
Data bersih siap (NNN baris).
--- PROSES 2: UPDATE ARVIEWER.xlsm ---
Membuka aplikasi Excel...
...
File disimpan.
--- PROSES 3: MENGHAPUS FILE SEMENTARA ---

--> Menjalankan skrip: 2_AddGiroDate.py...
--> File Giro_temp.xlsx telah berhasil dibuat...

--> Menjalankan skrip: 3_InjectGiroDtl2SS.py...
--- PROSES 4: PEMBARUAN KOLOM TANGGAL JT ---
...
PROSES BERHASIL! Kolom 'Tanggal JT' berhasil diperbarui dengan aman.

--> Mengembalikan file hasil pembaruan ke folder utama...
    * File ARVIEWER.xlsm berhasil dipindahkan kembali ke folder utama.
--> Membersihkan file sisa di folder Dapur...

=== SEMUA PROSES BERHASIL DISELESAIKAN DENGAN AMAN! ===
--> Tekan Enter untuk menutup jendela...
```

### Langkah 5 — Buka ARVIEWER.xlsm

`ARVIEWER.xlsm` sudah kembali di folder utama dengan data yang diperbarui. Buka dan gunakan seperti biasa.

---

## 🔄 Alur Kerja Pipeline

```
[Mulai: Jalankan Pembersihan dan Injeksi.py]
   │
   ├─── Validasi file root
   │       Cek Piutang.xls ada
   │       Cek Giro.xls ada
   │       Cek ARVIEWER.xlsm ada
   │       Jika kurang → tampilkan daftar & berhenti
   │
   ├─── Validasi folder Dapur/
   │       Cek folder ada
   │       Cek semua 4 skrip ada
   │       Jika kurang → tampilkan daftar & berhenti
   │
   ├─── PINDAHKAN (bukan salin) ketiga file → Dapur/
   │       Piutang.xls → Dapur/ExportFile.xls
   │       Giro.xls → Dapur/Giro.xls
   │       ARVIEWER.xlsm → Dapur/ARVIEWER.xlsm
   │
   ├─── Pindah ke direktori Dapur/ → jalankan skrip
   │
   ├─── [Skrip 1] 1_CleaningMovingALLAR.py
   │       ├─ Proses 1: Bersihkan Piutang.xls → Piutang_temp.xlsx
   │       ├─ Proses 2: Buka ARVIEWER.xlsm via xlwings (background)
   │       │            Hapus data lama A4:K{N} → Tulis data baru mulai A4
   │       │            Format kolom tanggal → Simpan & tutup
   │       └─ Proses 3: Hapus Piutang.xls & Piutang_temp.xlsx
   │
   ├─── [Skrip 2] 2_AddGiroDate.py
   │       └─ Bersihkan Giro.xls → Giro_temp.xlsx
   │
   ├─── [Skrip 3] 3_InjectGiroDtl2SS.py
   │       ├─ Proses 4: Bangun mapping {No. Faktur → "JT DD/MM/YY & ..."}
   │       │            dari Giro_temp.xlsx
   │       ├─ Buka ARVIEWER.xlsm via xlwings (background)
   │       ├─ Baca No. Faktur dari kolom B (B4:B{N})
   │       ├─ Tulis Tanggal JT per baris → kolom "Tanggal JT"
   │       └─ Simpan & tutup
   │
   ├─── Kembali ke direktori utama
   ├─── Pindahkan ARVIEWER.xlsm dari Dapur/ → folder utama
   └─── Hapus semua .xls/.xlsx/.xlsm sisa dari Dapur/

=== SELESAI ===
```

---

## 🔍 Detail Tiap Skrip

### Skrip 1 — Cleaning & Inject ke Source

**Input:** `Piutang.xls` (header ada di baris ke-4, index 3)

**Kolom yang dipilih** dari ekspor Accurate (berdasarkan posisi indeks, bukan nama):

| Indeks kolom asli | Nama kolom baru | Keterangan |
|---|---|---|
| 2 | `Kode Pelanggan` | Kode unik pelanggan (di-ffill dari sel merge) |
| 3 | `No. Faktur` | Nomor faktur (baris tanpa ini → dihapus) |
| 5 | `Tgl Faktur` | Tanggal faktur (diparsing + diformat Indonesia) |
| — | `SS` | *Kolom spacer kosong* (disisipkan setelah Tgl Faktur) |
| 9 | `Jatuh Tempo` | Tanggal jatuh tempo (diparsing + diformat Indonesia) |
| — | `SS` | *Kolom spacer kosong* (disisipkan setelah Jatuh Tempo) |
| 11 | `Nilai Faktur` | Nilai faktur (dibersihkan dari `.0` / `,00`) |
| 14 | `Sisa Piutang` | Sisa piutang (dibersihkan dari `.0` / `,00`) |
| 16 | `Umur JT` | Umur piutang sejak jatuh tempo |
| 18 | `Nama Pelanggan` | Nama lengkap pelanggan |
| 20 | `Nama Penjual` | Nama sales/penjual |
| 22 | `Nama Kontak` | Kontak pelanggan |

> Dua kolom `SS` (blank) disisipkan sebagai spacer agar kolom-kolom di ARVIEWER.xlsm tetap sejajar dengan formula/referensi yang sudah ada di sheet lain.

**Proses pembersihan:**

```
Piutang.xls
  ├─ Baca dengan header baris ke-4 (index 3)
  ├─ Pilih 10 kolom berdasarkan indeks posisi
  ├─ Forward-fill Kode Pelanggan (isi baris kosong akibat merge cell Accurate)
  ├─ Drop baris di mana No. Faktur = NaN
  ├─ Hapus artefak desimal (.0 → hilang, ,00 → hilang)
  │   pada: Kode Pelanggan, Nilai Faktur, Sisa Piutang
  ├─ Parse tanggal Indonesia → datetime:
  │   Mei→May, Agu→Aug, Okt→Oct, Nop→Nov, Des→Dec, Peb→Feb
  ├─ Sisipkan kolom "SS" (blank) setelah Tgl Faktur dan Jatuh Tempo
  └─ Simpan sementara → EXPORT_Sementara.xlsx
```

**Injeksi ke ARVIEWER.xlsm:**

```
Buka ARVIEWER.xlsm via xlwings (Excel background, tidak terlihat)
  ├─ Akses sheet 'Source'
  ├─ Cari baris terakhir berisi data → Hapus konten A4:K{last_row}
  ├─ Tulis df_clean mulai A4 (tanpa header, tanpa index)
  ├─ Format kolom bertipe tanggal dengan: [$-421]dd mmm yyyy;@
  │   (locale Indonesia, contoh: 15 Jan 2025)
  └─ Simpan & tutup workbook → tutup Excel
```

---

### Skrip 2 — Bersihkan Data Giro

**Input:** `Giro.xls` (tanpa header baku — dibaca `header=None`)

Skrip memfilter baris yang memiliki minimal **9 kolom tidak kosong** untuk membuang baris total, subtotal, dan baris kosong yang umum ada di ekspor Giro dari sistem keuangan.

**Kolom yang dipetakan (posisi 0–8):**

| Posisi | Nama Kolom |
|---|---|
| 0 | `No. Pelanggan` |
| 1 | `Nama Pelanggan` |
| 2 | `Tgl Faktur` |
| 3 | `No. Faktur. (SO)` ← kunci pencocokan dengan ARVIEWER |
| 4 | `No. Form` |
| 5 | `Total Diterima` |
| 6 | `Nilai terima` |
| 7 | `Nama Bank` |
| 8 | `Tgl Cek` ← tanggal cair giro, dimasukkan ke Tanggal JT |

Kolom angka (`Total Diterima`, `Nilai terima`) dibersihkan dari koma desimal dan dikonversi ke numerik.

**Output:** `Giro_temp.xlsx`

---

### Skrip 3 — Inject Tanggal JT ke ARVIEWER

**Input:** `Giro_temp.xlsx` + `ARVIEWER.xlsm`

Skrip membangun `mapping_jt` — kamus `{No. Faktur: "JT ..."}` — lalu menyuntikkan nilainya ke kolom `Tanggal JT` di sheet `Source` ARVIEWER.

**Logika pembangunan mapping:**

```
Untuk setiap No. Faktur di Giro_temp.xlsx:
  ├─ Kumpulkan semua Tgl Cek yang terkait
  ├─ Parse ke komponen (hari, bulan, tahun) — dukung Timestamp & string Indonesia
  ├─ Urutkan kronologis
  ├─ Kelompokkan per bulan+tahun:
  │   Jika banyak hari dalam satu bulan → gabung dengan koma: "01,15/01/25"
  │   Jika satu hari saja → "15/01/25"
  └─ Gabung antar bulan dengan " & ": "15/01/25 & 10,22/02/25"
  Prefix: "JT " + string gabungan
```

**Proses injeksi:**

```
Buka ARVIEWER.xlsm via xlwings (background)
  ├─ Cari kolom 'Tanggal JT' di baris header (baris 3)
  ├─ Baca semua No. Faktur dari kolom B (B4:B{last_row})
  ├─ Per faktur:
  │   lookup mapping_jt → jika cocok: tulis nilai "JT ..."
  │                      → jika tidak ada di Giro: tulis "" (kosong)
  └─ Simpan & tutup
```

---

## 📋 Format Data Input

### `Piutang.xls`

File ekspor laporan AR dari Accurate. **Persyaratan kritis:**
- Format file: `.xls` (bukan `.xlsx`)
- Header kolom **berada di baris ke-4** (baris 1–3 biasanya berisi judul laporan dan info periode)
- Kolom yang dibutuhkan berada di **indeks posisi** 2, 3, 5, 9, 11, 14, 16, 18, 20, 22

> Jika struktur ekspor Accurate berubah (urutan kolom bergeser), indeks pada `target_indices` di Skrip 1 perlu disesuaikan manual.

### `Giro.xls`

Rekap pembayaran giro/cek. **Persyaratan:**
- Format file: `.xls`
- Tidak memiliki baris header yang baku — baris data valid diidentifikasi dari jumlah kolom tidak kosong (≥ 9)
- Kolom ke-4 (indeks 3) harus berisi **No. Faktur (SO)** yang sama persis dengan `No. Faktur` di `Piutang.xls` agar pencocokan berhasil

### `ARVIEWER.xlsm`

Workbook Excel macro-enabled yang menjadi dashboard. **Persyaratan struktur:**
- Sheet bernama **`Source`** harus ada
- Baris **3** adalah baris header kolom di sheet `Source`, dan harus mengandung teks `Tanggal JT`
- Data dimulai dari **baris 4**
- Kolom **B** berisi `No. Faktur` (digunakan untuk pencocokan dengan data Giro)

---

## 📤 Output: Perubahan pada ARVIEWER.xlsm

Setelah pipeline selesai, dua hal berubah di dalam sheet `Source` ARVIEWER.xlsm:

**1. Data AR diperbarui (baris 4 ke bawah, kolom A–K)**

| Kolom | Isi |
|---|---|
| A | Kode Pelanggan |
| B | No. Faktur |
| C | Tgl Faktur (format `[$-421]dd mmm yyyy`) |
| D | SS *(kosong / spacer)* |
| E | Jatuh Tempo (format `[$-421]dd mmm yyyy`) |
| F | SS *(kosong / spacer)* |
| G | Nilai Faktur |
| H | Sisa Piutang |
| I | Umur JT |
| J | Nama Pelanggan |
| K | Nama Penjual |

Data lama dari run sebelumnya dihapus terlebih dahulu sebelum data baru ditulis.

**2. Kolom `Tanggal JT` diisi dengan tanggal giro per faktur**

Setiap baris yang nomor fakturnya ada di `Giro.xls` akan terisi nilai seperti:

| Contoh nilai | Arti |
|---|---|
| `JT 15/01/25` | Satu cek giro cair 15 Januari 2025 |
| `JT 01,22/01/25` | Dua cek giro cair di bulan Januari: tgl 1 & 22 |
| `JT 15/01/25 & 10/02/25` | Dua cek giro di bulan berbeda |
| `JT 01,15/01/25 & 10,28/02/25` | Dua bulan, masing-masing dua cek |
| *(kosong)* | Tidak ada data giro untuk faktur ini |

---

## 📅 Format Tanggal JT (Giro)

Format nilai `Tanggal JT` dirancang ringkas agar mudah dibaca di sel Excel:

```
"JT " + [hari],[hari]/[bulan]/[tahun-2digit] & [hari]/[bulan]/[tahun-2digit] & ...
```

**Aturan penggabungan:**
- Tanggal diurutkan dari yang paling awal
- Jika ada beberapa cek cair dalam **bulan yang sama**: hari-hari digabung dengan koma sebelum `/bulan/tahun`
- Jika ada cek di **bulan berbeda**: blok bulan dipisahkan dengan ` & `
- Tahun ditampilkan 2 digit (`25` untuk 2025)
- Bulan ditampilkan 2 digit (`01` untuk Januari)

**Contoh:**

```
Tanggal cek masuk:  1 Jan 2025, 15 Jan 2025, 10 Feb 2025
Hasil mapping:     "JT 01,15/01/25 & 10/02/25"
```

---

## 🛠️ Troubleshooting

### ❌ `[GAGAL] File berikut tidak ditemukan di folder utama`
Salah satu atau lebih dari `Piutang.xls`, `Giro.xls`, `ARVIEWER.xlsm` belum ada di folder utama. Pastikan ketiga file ada dengan nama **persis** (termasuk ekstensi `.xls` dan `.xlsm`).

### ❌ `TERJADI ERROR: ... PermissionError ...` (di Skrip 1 atau 3)
`ARVIEWER.xlsm` sedang terbuka di Excel. Tutup file tersebut terlebih dahulu, lalu jalankan ulang pipeline.

### ❌ `Sheet 'Source' tidak ditemukan!`
Sheet di dalam `ARVIEWER.xlsm` yang menjadi target data tidak bernama `Source`. Buka `ARVIEWER.xlsm` dan periksa nama sheet — pastikan ada sheet bernama persis `Source`.

### ❌ `ERROR: Judul kolom 'Tanggal JT' tidak ditemukan pada baris 3!`
Header di baris 3 sheet `Source` tidak mengandung teks `Tanggal JT`. Pastikan kolom tersebut ada di baris 3 dan namanya persis `Tanggal JT`.

### ❌ `Gagal membaca file Piutang.xls`
Kemungkinan penyebab: (1) file rusak atau terproteksi, (2) `xlrd` versi terlalu baru. Coba install `xlrd` versi kompatibel:
```bash
pip install "xlrd>=1.0.0,<2.0.0"
```

### ❌ Kolom tanggal di ARVIEWER tampil sebagai angka (serial date)
Terjadi jika Excel tidak menerapkan format tanggal. Skrip menggunakan xlwings untuk memformat ulang kolom tanggal dengan `[$-421]dd mmm yyyy;@`. Pastikan xlwings dan Excel sudah terinstall dengan benar dan berjalan tanpa error.

### ❌ Kolom `Tanggal JT` kosong semua setelah proses
Berarti tidak ada nomor faktur di `Giro.xls` yang cocok dengan nomor di `Piutang.xls`. Periksa format `No. Faktur. (SO)` di Giro — pastikan nilainya identik (termasuk trailing `.0` — skrip membuang `.0` dari keduanya sebelum membandingkan).

### ❌ Proses berhenti di tengah, file terlanjur dipindahkan ke Dapur/
Jika pipeline gagal setelah file dipindahkan ke `Dapur/`, file tidak akan dikembalikan otomatis. Ambil manual ketiga file dari dalam folder `Dapur/` kembali ke folder utama, lalu perbaiki penyebab error sebelum menjalankan ulang.

### ❌ `xlwings` tidak bisa menemukan instalasi Excel
Pastikan Microsoft Excel sudah terinstall di komputer. xlwings tidak mendukung LibreOffice Calc untuk operasi `.xlsm`.

---

## 📌 Catatan Penting

- **File input dipindahkan, bukan disalin** — Orkestrator menggunakan `shutil.move()`. Ketiga file (`Piutang.xls`, `Giro.xls`, `ARVIEWER.xlsm`) akan **hilang dari folder utama** selama proses berlangsung dan baru dikembalikan (`ARVIEWER.xlsm`) atau dihapus (dua file `.xls`) setelah selesai. **Simpan backup** jika data sumber masih dibutuhkan setelah proses.
- **Hanya `ARVIEWER.xlsm` yang dikembalikan** — `Piutang.xls` dan `Giro.xls` dihapus otomatis setelah data berhasil diproses. Ini disengaja karena datanya sudah tersimpan di dalam ARVIEWER.
- **Data lama selalu terhapus dulu** — Setiap kali pipeline dijalankan, isi `A4:K{N}` di sheet `Source` dihapus sepenuhnya sebelum data baru ditulis. Ini memastikan tidak ada data lama yang bercampur.
- **Jangan jalankan dua instance sekaligus** — xlwings membuka Excel secara background; menjalankan pipeline dua kali bersamaan dapat menyebabkan konflik file.
- **Jangan ubah nama kolom `Tanggal JT` di ARVIEWER** — Skrip 3 mencari kolom ini berdasarkan nama teks eksak di baris 3. Jika diubah, injeksi Giro tidak akan berjalan.
- **Struktur kolom `Source` harus konsisten** — Jika kolom di sheet `Source` ditambah atau digeser, kolom spacer `SS` dan urutan A–K akan bergeser juga. Sesuaikan `target_indices` di Skrip 1 jika ekspor Accurate berubah strukturnya.
- **Makro dan sheet lain di ARVIEWER tidak terpengaruh** — xlwings hanya memodifikasi konten sel di sheet `Source`. Semua VBA module, formula di sheet lain, chart, dan formatting lain tetap utuh.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
