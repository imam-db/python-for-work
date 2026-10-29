# 05 · Data Validator

Validate Excel rows against rules defined in a YAML config.

## Story

Andi handles program administration at **Lembaga Pelatihan Cendekia**, a training center that
runs classes with partner schools. Registration files arrive from several sources, and the
CRM rejects an entire upload when a single row is bad: malformed emails, phone numbers that
are too short, impossible birth dates, missing names, NIK values that are not 16 digits, and
statuses outside the allowed list.

Instead of checking columns one by one, Andi defines the rules once in YAML. Every finding is
reported with its Excel row number, the value, and a message — before anything reaches the
CRM.

The sample file has 12 registration rows with seven columns:

```text
Nama Lengkap · Email · No HP · Tanggal Lahir · Umur · NIK · Status
```

## Install

```bash
pip install -r requirements.txt
```

## Usage

```bash
# use config.yaml in the utility folder
python data_validator.py

# pass input and output on the command line
python data_validator.py data.xlsx --output report.xlsx
```

Command-line paths resolve against the current working directory. `input` and `output` inside
the config resolve against the utility folder.

## Options

| Option | Short | Description |
|--------|-------|-------------|
| `input` | – | Input Excel file (positional) |
| `--output` | `-o` | Export the result to an Excel file |
| `--config` | `-c` | YAML config file (default `config.yaml`) |

## Supported rules

| Rule | Meaning | Config example |
|------|---------|----------------|
| `required` | Value must not be empty | `- type: required` |
| `email` | Valid email format | `- type: email` |
| `phone` | Digit count within min–max | `- type: phone`<br>`  min_digits: 10`<br>`  max_digits: 15` |
| `date_range` | Date within a range | `- type: date_range`<br>`  min: "1950-01-01"` |
| `number_range` | Number within a range | `- type: number_range`<br>`  min: 17`<br>`  max: 65` |
| `regex` | Custom pattern | `- type: regex`<br>`  pattern: "^\\d{16}$"` |
| `in_list` | Value from a fixed list | `- type: in_list`<br>`  values: ["Aktif", "Nonaktif"]` |
| `unique` | Value must not repeat | `- type: unique` |

A column can have several rules; they run in the order written.

## Example configuration

```yaml
input: sample/data_pendaftaran.xlsx
output: output/validation_report.xlsx

rules:
  Nama Lengkap:
    - type: required

  Email:
    - type: required
    - type: email

  No HP:
    - type: required
    - type: phone
      min_digits: 10
      max_digits: 15

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

## Demo output

```bash
python generate_sample.py
python data_validator.py
```

```text
Membaca file: .../05-data-validator/sample/data_pendaftaran.xlsx
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

Laporan Excel disimpan ke: .../05-data-validator/output/validation_report.xlsx
```

The run checks **12 rows** and finds **5 valid rows** and **7 rows with errors**, for **17
findings** in total.

## Phone numbers stay text

Columns with a `phone` rule are read as text, so values such as `08123456` keep their leading
zero in the report. Digit length is checked after normalizing a `62`/`8` prefix to `0`, using
`min_digits` and `max_digits` (default 10–15). The exported workbook also writes those columns
as text (number format `@`).

## Excel report

The report contains three sheets:

1. **Summary** — total rows, valid rows, error rows, and total findings
2. **Detail Error** — row, column, value, rule, and message
3. **Data** — the original data with problem cells highlighted in red

## Notes

- If a configured column is missing, that column's rules are skipped with a warning.
- Empty cells are not validated, except by the `required` rule.
- Rules run in the order they appear in the YAML file.

## Ideas

- Validate across columns (for example, start date before end date).
- Load rules from a separate, reusable file.
- Export only the failing rows.
