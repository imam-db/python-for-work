# 04 · Excel Comparator

Compare two Excel files and report exactly what changed.

## Story

Citra works in finance at **PT Arta Graha**, a building-materials supplier. A colleague sends
back the month-end workbook with the note "a few corrections — please check." There is no
change log, the sheet is hundreds of rows long, and approving it means knowing which cells
actually moved. Scrolling two files side by side is how mistakes slip through.

The sample has ten rows in each version, with changes spread across totals and statuses:

| No | Change in the revised file |
|----|----------------------------|
| 1 | Total 5000000 → 5500000 |
| 2 | Status Pending → Lunas |
| 4 | Status Cicilan → Lunas |
| 5 | Total 3100000 → 3600000 |
| 6 | Status Pending → Lunas |
| 8 | Status Pending → Lunas |
| 9, 10 | Rows deleted |
| 11, 12 | Rows added |

## Install

Dependencies live in the repository root:

```bash
pip install -r requirements.txt
# or, for the pfw command too:
pip install -e .
```

## Usage

```bash
# use config.yaml in the utility folder
python excel_comparator.py

# pass the two files on the command line
python excel_comparator.py file_lama.xlsx file_baru.xlsx

# match rows by a key column (recommended)
python excel_comparator.py file_lama.xlsx file_baru.xlsx --key "No"

# also export the report to Excel
python excel_comparator.py file_lama.xlsx file_baru.xlsx --key "No" -o laporan.xlsx
```

The same options are available through `pfw`:

```bash
pfw compare file_lama.xlsx file_baru.xlsx --key "No" -o laporan.xlsx
```

Command-line paths resolve against the current working directory. `file_old`, `file_new`, and
`output` inside the config resolve against the utility folder.

## Options

| Option | Short | Description |
|--------|-------|-------------|
| `file_old` | – | Original Excel file (positional) |
| `file_new` | – | Revised Excel file (positional) |
| `--key` | `-k` | Key column used to match rows |
| `--output` | `-o` | Export the report to an Excel file |
| `--config` | `-c` | YAML config file (default `config.yaml`) |

## Demo output

```bash
python generate_sample.py
python excel_comparator.py
```

```text
File lama : .../04-excel-comparator/sample/data_lama.xlsx
File baru : .../04-excel-comparator/sample/data_revisi.xlsx
Key column: No

File lama: 10 baris, 5 kolom
File baru: 10 baris, 5 kolom

==================================================
LAPORAN PERBANDINGAN
==================================================

Ringkasan:
  - Baris baru     : 2
  - Baris dihapus  : 2
  - Cell berubah   : 6

--- PERUBAHAN (6) ---
  [No=1] Total: 5000000 -> 5500000
  [No=2] Status: Pending -> Lunas
  [No=4] Status: Cicilan -> Lunas
  [No=5] Total: 3100000 -> 3600000
  [No=6] Status: Pending -> Lunas
  [No=8] Status: Pending -> Lunas

--- BARIS BARU (2) ---
  [No=11]
  [No=12]

--- BARIS DIHAPUS (2) ---
  [No=9]
  [No=10]

Laporan Excel disimpan ke: .../04-excel-comparator/output/laporan_perbandingan.xlsx
```

The verified result is **2 added rows, 2 deleted rows, and 6 changed cells**.

## Comparison modes

**With a key column (recommended)** — rows are matched by the key value, so reordered rows
are not reported as differences:

```bash
python excel_comparator.py old.xlsx new.xlsx --key "No"
```

**Without a key column** — rows are compared by position (row 1 against row 1).

## Excel report

The exported workbook contains (the output folder is created automatically if needed):

1. **Summary** — counts of added, deleted, and changed items
2. **Perubahan** — changed cells (yellow header)
3. **Baris Baru** — added rows (green header)
4. **Baris Dihapus** — deleted rows (red header)

## Duplicate keys

The key column should be unique. If duplicates exist, the script still runs using the
**first** matching row and prints a warning listing the duplicate keys:

```text
Peringatan: file lama kolom kunci 'No' duplikat: 1, 2; memakai baris pertama
```

## Notes

- Both files should have the same column structure.
- Comparison is case-sensitive.
- Added and deleted rows are sorted numerically when the key is numeric.

## Ideas

- Ignore selected columns.
- Allow a numeric tolerance for amount comparisons.
- Export the report as HTML or PDF.
