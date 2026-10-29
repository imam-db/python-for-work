# 01 · Merge Excel

Combine every `.xlsx` workbook in a folder into a single file.

## Story

Raka works in operations at **PT Sinar Teknologi Nusantara**, an IT distributor whose Jakarta
head office collects a sales workbook from every city branch at the start of each month. The
files all share the same columns, but they arrive separately from Bandung, Surabaya, and the
rest. Before the Monday directors' meeting, Raka needs one sheet — not a folder of tabs.

The sample uses three branch files with **272 transactions** in total:

```text
sample/
├── cabang_jakarta.xlsx    (120 transactions)
├── cabang_bandung.xlsx    (72 transactions)
└── cabang_surabaya.xlsx   (80 transactions)
```

Columns are identical across the files:

| No | Tanggal | Nama Sales | Produk | Qty | Harga Satuan | Total | Status | Metode Bayar |
|----|---------|------------|--------|-----|--------------|-------|--------|--------------|
| 1 | 2024-01-05 | Budi Santoso | Laptop ASUS ROG | 2 | 16500000 | 33000000 | Lunas | Transfer |
| 2 | 2024-01-05 | Dewi Lestari | Monitor LG 27 inch | 5 | 4200000 | 21000000 | Cicilan | Kartu Kredit |

## Install

```bash
pip install -r requirements.txt
```

## Usage

```bash
python merge_excel.py <input_folder> [output_file]
```

| Argument | Required | Description |
|----------|----------|-------------|
| `input_folder` | yes | Folder containing `.xlsx` files |
| `output_file` | no | Merged output path. Default: `output/hasil_gabungan.xlsx` next to the script |

Command-line paths resolve against the current working directory, so the script can be run
from anywhere:

```bash
# from the utility folder
python merge_excel.py ./sample

# from the repository root
python 01-merge-excel/merge_excel.py ./01-merge-excel/sample
```

## Demo output

```bash
python generate_sample.py
python merge_excel.py ./sample
```

```text
Ditemukan 3 file Excel di folder ./sample
- Memproses: cabang_bandung.xlsx (72 baris)
- Memproses: cabang_jakarta.xlsx (120 baris)
- Memproses: cabang_surabaya.xlsx (80 baris)

Berhasil! Total 272 baris dari 3 file digabungkan ke .../01-merge-excel/output/hasil_gabungan.xlsx
```

## How merging works

- A **Cabang** column is added from the filename: `cabang_jakarta.xlsx` → `Jakarta`.
- A file with **0 rows** is skipped and reported as a warning.
- If column sets differ, the file is still included (outer concat) and a warning lists the
  missing and extra columns. Rows from that file get `NaN` in columns they do not have, so
  the gap is visible.
- Files are processed in sorted filename order for consistent output.

Example warning:

```text
- PERINGATAN: kolom cabang_bandung.xlsx berbeda (hilang: ['Status'], tambahan: -);
  tetap digabung, kolom yang tidak ada akan berisi NaN
```

## Notes

- Only the first sheet of each workbook is read.
- Legacy `.xls` files are not supported; convert them to `.xlsx` first.
- The `No` column keeps the original per-file numbering; it is not renumbered after merging.

## Ideas

- Renumber `No` after the merge.
- Filter input files by filename pattern.
- Read multiple sheets per workbook.
