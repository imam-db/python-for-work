# 05 · Data Validator

> Bab 5 kisah [Bu Sari](../README.md): Jumat, data pendaftaran harus lolos pemeriksaan
> sebelum masuk sistem.

Validasi data Excel berdasarkan rules yang didefinisikan di config YAML.

## Cerita

Data pendaftaran PT Maju Jaya datang dari form online, input manual SPG, dan import sistem
lain. Sebelum diproses, kualitasnya harus dicek: email yang formatnya salah, nomor HP yang
kurang digit, tanggal lahir yang tidak masuk akal, field wajib yang kosong, NIK yang bukan
16 digit, dan status di luar pilihan yang tersedia.

Daripada memfilter dan memeriksa satu per satu, rules didefinisikan sekali di YAML. Setiap
temuan dilaporkan lengkap dengan nomor baris, nilai, dan pesannya.

## Instalasi

```bash
pip install -r requirements.txt
```

## Cara Menggunakan

```bash
# pakai config.yaml di folder utility
python data_validator.py

# file dan output lewat CLI
python data_validator.py data.xlsx --output report.xlsx
```

Path dari CLI mengikuti current working directory. Path `input` dan `output` di dalam config
di-resolve relatif terhadap folder utility.

## Opsi CLI

| Opsi | Shortcut | Deskripsi |
|------|----------|-----------|
| `input` | – | File Excel input (positional) |
| `--output` | `-o` | Export hasil ke file Excel |
| `--config` | `-c` | File config YAML (default `config.yaml`) |

## Rules yang Didukung

| Rule | Keterangan | Contoh config |
|------|-----------|---------------|
| `required` | Tidak boleh kosong | `- type: required` |
| `email` | Format email valid | `- type: email` |
| `phone` | Panjang digit min–max | `- type: phone`<br>`  min_digits: 10`<br>`  max_digits: 15` |
| `date_range` | Tanggal dalam rentang | `- type: date_range`<br>`  min: "1950-01-01"` |
| `number_range` | Angka dalam rentang | `- type: number_range`<br>`  min: 17`<br>`  max: 65` |
| `regex` | Pola kustom | `- type: regex`<br>`  pattern: "^\\d{16}$"` |
| `in_list` | Nilai dari daftar | `- type: in_list`<br>`  values: ["Aktif", "Nonaktif"]` |
| `unique` | Tidak boleh duplikat | `- type: unique` |

Satu kolom boleh punya beberapa rules, dijalankan berurutan.

## Demo (hasil run nyata)

```bash
python generate_sample.py
python data_validator.py
```

```
Membaca file: /path/ke/repo/05-data-validator/sample/data_pendaftaran.xlsx
Total baris: 12
Kolom: Nama Lengkap, Email, No HP, Tanggal Lahir, Umur, NIK, Status

Memvalidasi data...

==================================================
LAPORAN VALIDASI
==================================================

Ringkasan:
  Total baris : 12
  Valid       : 5
  Error       : 7 baris (17 masalah)

--- DETAIL ERROR (17) ---
  Baris 4: Nama Lengkap tidak boleh kosong
  Baris 9: Nama Lengkap tidak boleh kosong
  Baris 6: Email tidak boleh kosong
  Baris 3: Format email tidak valid: dewi@email
  Baris 8: Format email tidak valid: joko@@email.com
  Baris 3: No HP kurang dari 10 digit: 08123456
  Baris 7: No HP kurang dari 10 digit: 0812
  Baris 4: Tanggal sebelum 1950-01-01: 1940-01-01
  Baris 6: Tanggal setelah 2010-12-31: 2015-06-10
  Baris 10: Tanggal setelah 2010-12-31: 2020-01-01
  Baris 6: Nilai kurang dari 17: 15
  Baris 9: Nilai lebih dari 65: 70
  Baris 4: NIK harus 16 digit angka
  Baris 7: NIK harus 16 digit angka
  Baris 10: NIK harus 16 digit angka
  Baris 6: Nilai "Tidak Valid" tidak valid. Harus salah satu dari: ['Aktif', 'Nonaktif', 'Pending']
  Baris 9: Nilai "Batal" tidak valid. Harus salah satu dari: ['Aktif', 'Nonaktif', 'Pending']

Laporan Excel disimpan ke: /path/ke/repo/05-data-validator/output/validation_report.xlsx
```

## Nomor HP Tetap Teks

Kolom yang punya rule `phone` dibaca sebagai teks, sehingga nilai di laporan tetap
`08123456`, bukan `8123456`. Panjang digit diperiksa setelah normalisasi awalan `62`/`8`
menjadi `0`, dengan batas bawah `min_digits` dan batas atas `max_digits` (default 10–15).
Laporan Excel juga menulis kolom tersebut sebagai teks (number format `@`).

## Output Excel

File laporan berisi 3 sheet:

1. **Summary** — jumlah baris, valid, baris error, dan total masalah
2. **Detail Error** — baris, kolom, nilai, rule, dan pesan
3. **Data** — data asli dengan highlight merah pada sel yang bermasalah

## Contoh Config Lengkap

Ada di [`config.yaml`](./config.yaml) pada folder ini, termasuk `min_digits` dan
`max_digits` untuk kolom `No HP`.

## Catatan

- Jika kolom yang diminta tidak ditemukan, validasi kolom tersebut dilewati dengan
  peringatan.
- Sel kosong tidak divalidasi, kecuali oleh rule `required`.
- Rules dijalankan berurutan sesuai urutan di YAML.

## Roadmap Singkat

- [ ] Validasi antar kolom (mis. tanggal mulai < tanggal selesai)
- [ ] Import rules dari file terpisah
- [ ] Export hanya baris yang error
