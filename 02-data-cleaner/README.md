# 02 · Data Cleaner

> Bab 2 kisah [Bu Sari](../README.md): Selasa, data pelanggan dari form online harus
> dirapikan sebelum dipakai tim marketing.

Membersihkan dan menstandardisasi data Excel berdasarkan konfigurasi YAML.

## Cerita

Data pelanggan PT Maju Jaya datang dari banyak sumber: form online, input manual SPG, dan
export sistem lama. Hasilnya campur aduk:

```
| Nama Lengkap     | Tanggal Lahir | No HP              |
|------------------|---------------|--------------------|
|   BUDI SANTOSO   | 15/01/1990    | 081234567890       |
| siti nurhaliza   | 1985-03-20    | +62 812 3456 7891  |
| AHMAD   DAHLAN   | 05 Mei 2000   | 0813-4567-8901     |
|   dewi  lestari  | 12-12-1995    | 62 814 567 8902    |
| BUDI SANTOSO     | 15/01/1990    | 081234567890       |  <- duplikat
|                  | 01-01-2000    | 0817 4567 8905     |  <- nama kosong
```

Masalah yang ditangani:

- **Nama** — huruf besar/kecil tidak konsisten, spasi berlebih
- **Tanggal** — format campur (`DD/MM/YYYY`, `YYYY-MM-DD`, `05 Mei 2000`)
- **No HP** — awalan `+62`, `62`, `0`, dengan/tanpa pemisah; nol di depan bisa hilang
  jika Excel menyimpan nomor sebagai angka
- **Duplikat** — baris yang sama diinput berkali-kali
- **Data kosong** — kolom penting tidak diisi

## Instalasi

```bash
pip install -r requirements.txt
```

## Cara Menggunakan

```bash
python data_cleaner.py                 # config.yaml di folder utility
python data_cleaner.py config.yaml     # positional
python data_cleaner.py -c config.yaml  # opsi eksplisit
```

Path file config dari CLI mengikuti current working directory. Path `input` dan `output`
di dalam config di-resolve relatif terhadap folder utility, atau bisa diisi absolut.

## Konfigurasi

```yaml
input: sample/data_kotor.xlsx
output: output/data_bersih.xlsx

cleaning:
  nama:
    kolom: "Nama Lengkap"
    format: "title"   # title, upper, lower
    trim: true

  tanggal:
    kolom: "Tanggal Lahir"
    format: "%d-%m-%Y"

  telepon:
    kolom: "No HP"
    format: "0xxx-xxxx-xxxx"   # atau "+62xxx-xxxx-xxxx"

  duplikat:
    kolom: ["Nama Lengkap", "No HP"]

  hapus_kosong:
    kolom: ["Nama Lengkap"]
```

## Demo (hasil run nyata)

```bash
python generate_sample.py
python data_cleaner.py
```

```
Membaca file: /path/ke/repo/02-data-cleaner/sample/data_kotor.xlsx
Total baris: 10

Proses cleaning:
  Nama: 9 data di-standardisasi
  Tanggal: 10 data di-format ke %d-%m-%Y
  Telepon: 10 data di-format
  Duplikat: 1 baris dihapus
  Baris kosong: 1 baris dihapus

Berhasil! Data bersih disimpan ke: /path/ke/repo/02-data-cleaner/output/data_bersih.xlsx
Total baris setelah cleaning: 8
```

## Format Nomor HP

Nomor dinormalisasi ke awalan `0`, lalu diformat **tanpa memotong digit**:

| Input | Digit hasil normalisasi | Output |
|-------|-------------------------|--------|
| `081234567890` | `081234567890` (12) | `0812-3456-7890` |
| `+62 812 3456 7891` | `081234567891` (12) | `0812-3456-7891` |
| `62 814 567 8902` | `08145678902` (11) | `0814-567-8902` |

Aturan:

- 12 digit → `0xxx-xxxx-xxxx`
- 11 digit → `0xxx-xxx-xxxx`
- panjang lain (10–15) → grup 4 digit dari kiri, sisa digit tetap disertakan
- di luar 10–15 digit → nilai asli dipertahankan dan dihitung dalam peringatan

Kolom telepon dibaca dan ditulis sebagai **teks** (`dtype=str` + number format `@`) supaya
angka nol di depan tidak berubah menjadi angka.

## Catatan

- Kolom yang tidak ada di config dibiarkan apa adanya.
- Jika kolom yang diminta tidak ditemukan, cleaning untuk kolom itu dilewati dengan
  peringatan.
- File output ditimpa jika sudah ada.

## Roadmap Singkat

- [ ] Validasi format email
- [ ] Mapping nilai (mis. `JKT` → `Jakarta`)
- [ ] Export laporan cleaning terpisah
