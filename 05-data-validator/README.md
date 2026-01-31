# Data Validator

Validasi data Excel berdasarkan rules yang didefinisikan di config YAML.

## Cerita

Kamu terima data pendaftaran dari berbagai sumber: form online, input manual SPG, import dari sistem lain. Sebelum diproses, perlu dicek dulu kualitas datanya:

- Email yang formatnya salah
- Nomor HP yang kurang digit
- Tanggal lahir yang tidak masuk akal
- Field wajib yang kosong
- NIK yang tidak 16 digit
- Status yang tidak sesuai pilihan

**Cara manual:** Filter, sort, cek satu-satu, bikin catatan. Untuk 500 baris bisa 1 jam, dan sering ada yang kelewatan.

**Dengan script ini:** Definisikan rules sekali di config, jalankan kapan saja. Langsung dapat laporan lengkap.

## Instalasi

```bash
pip install -r requirements.txt
```

## Cara Menggunakan

### Via Config (Default)

Edit `config.yaml` untuk definisikan rules, lalu:
```bash
python data_validator.py
```

### Via CLI

```bash
python data_validator.py data.xlsx --output report.xlsx
```

## Validation Rules

### 1. required
Field tidak boleh kosong.
```yaml
Nama:
  - type: required
```

### 2. email
Format email harus valid.
```yaml
Email:
  - type: email
```

### 3. phone
Nomor telepon minimal X digit.
```yaml
No HP:
  - type: phone
    min_digits: 10
```

### 4. date_range
Tanggal harus dalam range tertentu.
```yaml
Tanggal Lahir:
  - type: date_range
    min: "1950-01-01"
    max: "2010-12-31"
```

### 5. number_range
Angka harus dalam range min-max.
```yaml
Umur:
  - type: number_range
    min: 17
    max: 65
```

### 6. regex
Validasi dengan custom pattern.
```yaml
NIK:
  - type: regex
    pattern: "^\\d{16}$"
    message: "NIK harus 16 digit angka"
```

### 7. in_list
Nilai harus salah satu dari list.
```yaml
Status:
  - type: in_list
    values: ["Aktif", "Nonaktif", "Pending"]
```

### 8. unique
Nilai tidak boleh duplikat.
```yaml
Email:
  - type: unique
```

## Demo dengan Sample Data

```bash
python generate_sample.py
python data_validator.py
```

Output:
```
Membaca file: sample/data_pendaftaran.xlsx
Total baris: 12
Kolom: Nama Lengkap, Email, No HP, Tanggal Lahir, Umur, NIK, Status

Memvalidasi data...

==================================================
LAPORAN VALIDASI
==================================================

Ringkasan:
  Total baris : 12
  Valid       : 4
  Error       : 8 baris (15 masalah)

--- DETAIL ERROR (15) ---
  Baris 3: Nama Lengkap tidak boleh kosong
  Baris 8: Nama Lengkap tidak boleh kosong
  Baris 2: Format email tidak valid: dewi@email
  Baris 5: Email tidak boleh kosong
  ...

Laporan Excel disimpan ke: output/validation_report.xlsx
```

## Output Excel Report

File Excel output berisi 3 sheet:

1. **Summary** - Ringkasan jumlah valid/error
2. **Detail Error** - List semua error dengan baris, kolom, nilai, dan pesan
3. **Data** - Data asli dengan highlight merah pada cell yang error

## CLI Options

| Option | Shortcut | Deskripsi |
|--------|----------|-----------|
| `input` | - | File Excel input (positional) |
| `--output` | `-o` | Export hasil ke file Excel |
| `--config` | `-c` | File config custom |

## Contoh Config Lengkap

```yaml
input: data/pendaftaran.xlsx
output: output/validation_report.xlsx

rules:
  Nama Lengkap:
    - type: required

  Email:
    - type: required
    - type: email
    - type: unique

  No HP:
    - type: required
    - type: phone
      min_digits: 10

  Tanggal Lahir:
    - type: date_range
      min: "1950-01-01"
      max: "2010-12-31"

  Umur:
    - type: number_range
      min: 17
      max: 65

  NIK:
    - type: regex
      pattern: "^\\d{16}$"
      message: "NIK harus 16 digit angka"

  Status:
    - type: in_list
      values: ["Aktif", "Nonaktif", "Pending"]
```

## Catatan Penting

- Satu kolom bisa punya multiple rules
- Rules dijalankan berurutan
- Jika kolom tidak ditemukan di file, validasi untuk kolom tersebut di-skip
- Cell kosong tidak divalidasi (kecuali rule `required`)

## Pengembangan Selanjutnya

- [ ] Custom error message per rule
- [ ] Validasi antar kolom (misal: tanggal_mulai < tanggal_selesai)
- [ ] Import rules dari file terpisah
- [ ] Export hanya baris yang error

## Blog

[Link ke artikel blog] *(coming soon)*
