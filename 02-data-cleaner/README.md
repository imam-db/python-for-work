# 02 · Data Cleaner

Clean and standardize names, dates, and phone numbers in an Excel file using YAML rules.

## Story

Nina is a marketing analyst on the promo team at **Toko Sumber Makmur**, a mid-size retailer.
Weekend booth staff collect customer signups on shared laptops, and by Monday the list looks
like it survived a typing contest: names in random caps, birth dates in three formats, and
phone numbers that lost their leading zero somewhere in Excel. Nina needs a dependable list
before the SMS campaign goes out.

A typical sheet before cleaning:

| Nama Lengkap | Tanggal Lahir | No HP |
|--------------|---------------|-------|
| `  BUDI SANTOSO ` | `15/01/1990` | `081234567890` |
| `siti nurhaliza` | `1985-03-20` | `+62 812 3456 7891` |
| `AHMAD   DAHLAN` | `05 Mei 2000` | `0813-4567-8901` |
| `dewi  lestari` | `12-12-1995` | `62 814 567 8902` |
| `BUDI SANTOSO` | `15/01/1990` | `081234567890` ← duplicate |
| *(blank)* | `01-01-2000` | `0817 4567 8905` |

## Install

Dependencies live in the repository root:

```bash
pip install -r requirements.txt
# or, for the pfw command too:
pip install -e .
```

## Usage

```bash
python data_cleaner.py                 # uses config.yaml in the utility folder
python data_cleaner.py config.yaml     # positional config path
python data_cleaner.py -c config.yaml  # explicit option
```

The same options are available through `pfw`:

```bash
pfw clean                              # uses config.yaml in the utility folder
pfw clean -c my_config.yaml
```

Config paths passed on the command line resolve against the current working directory.
Inside the config, `input` and `output` resolve against the utility folder unless absolute.

## Configuration

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
    format: "0xxx-xxxx-xxxx"   # or "+62xxx-xxxx-xxxx"

  duplikat:
    kolom: ["Nama Lengkap", "No HP"]

  hapus_kosong:
    kolom: ["Nama Lengkap"]
```

## Demo output

```bash
python generate_sample.py
python data_cleaner.py
```

```text
Membaca file: .../02-data-cleaner/sample/data_kotor.xlsx
Total baris: 10

Proses cleaning:
  Nama: 9 data di-standardisasi
  Tanggal: 10 data di-format ke %d-%m-%Y
  Telepon: 10 data di-format
  Duplikat: 1 baris dihapus
  Baris kosong: 1 baris dihapus

Berhasil! Data bersih disimpan ke: .../02-data-cleaner/output/data_bersih.xlsx
Total baris setelah cleaning: 8
```

So 10 dirty rows become **8 clean rows**: one duplicate and one row with a missing name are
removed.

## Phone formatting

Numbers are normalized to a leading `0` and formatted **without dropping digits**:

| Input | Normalized digits | Output |
|-------|-------------------|--------|
| `081234567890` | `081234567890` (12) | `0812-3456-7890` |
| `+62 812 3456 7891` | `081234567891` (12) | `0812-3456-7891` |
| `62 814 567 8902` | `08145678902` (11) | `0814-567-8902` |

Rules:

- 12 digits → `0xxx-xxxx-xxxx`
- 11 digits → `0xxx-xxx-xxxx`
- other lengths within 10–15 → groups of four from the left, remainder included
- outside 10–15 digits → the original value is kept and counted in a warning

Phone columns are read and written as **text** (`dtype=str` plus number format `@`) so a
leading zero never turns into a number.

## Notes

- Columns not mentioned in the config are left untouched.
- If a configured column is missing, that cleaning step is skipped with a warning.
- An existing output file is overwritten.

## Ideas

- Validate email format as an optional cleaning step.
- Map shorthand values (for example `JKT` → `Jakarta`).
- Export a separate cleaning report.
