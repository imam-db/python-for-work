# 03 · Excel Splitter

Split one large Excel file into separate files based on the values of a column.

## Story

Pak Hendra coordinates the regional sales team at **PT Sumber Pangan Abadi**, a packaged-food
distributor. Every month the ERP produces one national export containing transactions from
all branches. The file is fine for accounting, but five branch managers do not want to filter
a hundred rows themselves — they want one file each, delivered before their Monday review.

The sample export mixes five branches in a single sheet:

```text
laporan_nasional.xlsx
├── 100 transaction rows
├── Columns: No, Tanggal, Cabang, Nama Sales, Produk, Qty, Harga Satuan, Total, Status, Metode Bayar
└── Branch values: Jakarta, Bandung, Surabaya, Medan, Makassar
```

## Install

```bash
pip install -r requirements.txt
```

## Usage

```bash
# use config.yaml in the utility folder
python excel_splitter.py

# pass input file and split column on the command line
python excel_splitter.py laporan.xlsx --kolom "Cabang"

# override part of the config
python excel_splitter.py --kolom "Bulan" --prefix "report_"
```

Command-line paths resolve against the current working directory. `input` and
`output_folder` inside the config resolve against the utility folder.

## Options

| Option | Short | Description |
|--------|-------|-------------|
| `input` | – | Input Excel file (positional) |
| `--kolom` | `-k` | Column to split on |
| `--output` | `-o` | Output folder |
| `--prefix` | `-p` | Prefix for output filenames |
| `--suffix` | `-s` | Suffix for output filenames |
| `--no-header` | – | Do not include the header row in output files |
| `--config` | `-c` | YAML config file (default `config.yaml`) |

## Demo output

```bash
python generate_sample.py
python excel_splitter.py
```

```text
Membaca file: .../03-excel-splitter/sample/laporan_nasional.xlsx
Total baris: 100

Memecah berdasarkan kolom: Cabang
Ditemukan 5 grup

  - Bandung.xlsx: 20 baris
  - Jakarta.xlsx: 23 baris
  - Makassar.xlsx: 22 baris
  - Medan.xlsx: 15 baris
  - Surabaya.xlsx: 20 baris

Berhasil! 5 file dibuat di folder '.../03-excel-splitter/output'
Total baris: 100
```

The five files account for all 100 rows: Jakarta 23, Makassar 22, Bandung 20, Surabaya 20,
and Medan 15. The sample generator uses `random.seed(42)`, so this distribution is identical
on every run.

## Empty split values

Rows where the split column is empty (`NaN` or an empty string) are **not dropped**. They go
into a single bucket named `_kosong.xlsx`, and the count is reported:

```text
  - _kosong.xlsx: 2 baris

  2 baris tanpa nilai 'Cabang' masuk ke _kosong.xlsx
```

## YAML configuration

```yaml
input: sample/laporan_nasional.xlsx
output_folder: output/
split_by: "Cabang"
prefix: ""
suffix: ""
include_header: true
```

## Notes

- Characters that are invalid in filenames are replaced with `_`.
- Command-line options always override the config file.
- If the split column is missing, the script lists the columns that are available.
- Only the first sheet of the input workbook is read.

## Ideas

- Split on a combination of columns.
- Export groups as CSV or JSON.
- Pre-filter rows before splitting.
