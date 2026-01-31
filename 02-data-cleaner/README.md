# Data Cleaner

Membersihkan dan menstandardisasi data Excel berdasarkan konfigurasi YAML.

## Cerita

Kamu bekerja sebagai data analyst di perusahaan retail. Tim marketing baru saja selesai event promo dan mengumpulkan data pelanggan dari berbagai sumber: form online, pendaftaran di booth, dan input manual dari SPG.

Masalahnya? Data yang masuk berantakan:

```
| Nama Lengkap        | Tanggal Lahir | No HP              |
|---------------------|---------------|--------------------|
|   BUDI SANTOSO      | 15/01/1990    | 081234567890       |
| siti nurhaliza      | 1985-03-20    | +62 812 3456 7891  |
| AHMAD   DAHLAN      | 05 Mei 2000   | 0813-4567-8901     |
|   dewi  lestari     | 12-12-1995    | 62 814 567 8902    |
| BUDI SANTOSO        | 15/01/1990    | 081234567890       |  <- duplikat!
|                     | 01-01-2000    | 0817 4567 8905     |  <- nama kosong!
```

Masalah yang sering ditemui:
- **Nama**: Ada yang UPPERCASE, lowercase, spasi berlebih
- **Tanggal**: Format campur aduk (DD/MM/YYYY, YYYY-MM-DD, "05 Mei 2000")
- **No HP**: Ada yang pakai +62, 62, 0, dengan/tanpa separator
- **Duplikat**: Data yang sama diinput berkali-kali
- **Data kosong**: Field penting tidak diisi

**Cara manual:**
1. Sort data, cari duplikat, hapus satu-satu
2. Find & replace untuk standardisasi nama
3. Pakai formula Excel untuk format tanggal
4. Text to columns + concatenate untuk nomor HP
5. Filter dan hapus baris kosong

Waktu: **30-60 menit** untuk 500 baris. Dan harus diulang setiap ada data baru.

**Dengan script ini:** Definisikan aturan sekali di config, jalankan kapan saja. 5 detik selesai.

## Instalasi

```bash
pip install -r requirements.txt
```

## Cara Menggunakan

1. Siapkan file Excel yang mau dibersihkan
2. Edit `config.yaml` sesuai kebutuhan
3. Jalankan script:

```bash
python data_cleaner.py
```

Atau dengan config file custom:

```bash
python data_cleaner.py config_custom.yaml
```

## Konfigurasi

Edit file `config.yaml` untuk mengatur aturan cleaning:

```yaml
input: sample/data_kotor.xlsx
output: output/data_bersih.xlsx

cleaning:
  # Standardisasi nama
  nama:
    kolom: "Nama Lengkap"
    format: "title"  # title, upper, lower
    trim: true

  # Standardisasi format tanggal
  tanggal:
    kolom: "Tanggal Lahir"
    format: "%d-%m-%Y"

  # Standardisasi nomor HP
  telepon:
    kolom: "No HP"
    format: "0xxx-xxxx-xxxx"

  # Hapus duplikat
  duplikat:
    kolom: ["Nama Lengkap", "No HP"]

  # Hapus baris kosong
  hapus_kosong:
    kolom: ["Nama Lengkap"]
```

## Demo dengan Sample Data

Generate sample data kotor untuk testing:

```bash
python generate_sample.py
python data_cleaner.py
```

Output:
```
Membaca file: sample/data_kotor.xlsx
Total baris: 10

Proses cleaning:
  Nama: 9 data di-standardisasi
  Tanggal: 10 data di-format ke %d-%m-%Y
  Telepon: 10 data di-format
  Duplikat: 1 baris dihapus
  Baris kosong: 0 baris dihapus

Berhasil! Data bersih disimpan ke: output/data_bersih.xlsx
Total baris setelah cleaning: 9
```

## Fitur Cleaning

### 1. Standardisasi Nama
- Hapus spasi berlebih di awal, akhir, dan tengah
- Format: Title Case, UPPERCASE, atau lowercase

### 2. Standardisasi Tanggal
- Support berbagai format input (DD/MM/YYYY, YYYY-MM-DD, dll)
- Support format Indonesia ("05 Mei 2000", "30 Juni 1988")
- Output format bisa dikustomisasi

### 3. Standardisasi Nomor Telepon
- Normalisasi prefix (+62, 62, 0)
- Format output: `0xxx-xxxx-xxxx` atau `+62xxx-xxxx-xxxx`

### 4. Hapus Duplikat
- Deteksi duplikat berdasarkan kombinasi kolom
- Keep baris pertama, hapus sisanya

### 5. Hapus Baris Kosong
- Hapus baris jika kolom tertentu kosong

## Catatan Penting

- Kolom yang tidak ada di config akan dibiarkan apa adanya
- Jika kolom tidak ditemukan di file, proses cleaning untuk kolom tersebut di-skip
- File output akan di-overwrite jika sudah ada

## Pengembangan Selanjutnya

Fitur yang bisa ditambahkan:
- [ ] Validasi format email
- [ ] Standardisasi alamat
- [ ] Mapping nilai (misal: "JKT" -> "Jakarta")
- [ ] Export laporan cleaning ke file terpisah
- [ ] Support multiple sheet

## Blog

[Link ke artikel blog] *(coming soon)*
