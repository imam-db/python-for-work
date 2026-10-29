# split

Split one tabular file into separate files based on the values of a column.

## Story

Pak Hendra coordinates regional sales at **PT Sumber Pangan Abadi**. The ERP dumps one
national export; five branch managers each want their own file before Monday review. The
sample of **100** rows becomes **5** branch workbooks (measured on sample data).

## Input

A single tabular file (xlsx/csv/parquet) and the column to split on (`--kolom`). Optional
YAML defaults live in `configs/split.yaml`. Rows with blank split values go to `_kosong.xlsx`.
Pass `--no-header` to omit the header row in each output file (still written through
`pfw.common.io`).

## Command

```bash
pfw sample split -o ./sample/laporan_nasional.xlsx
pfw split ./sample/laporan_nasional.xlsx --kolom Cabang -o ./output
```

## Output

One file per distinct column value in the output folder, with optional prefix/suffix.
