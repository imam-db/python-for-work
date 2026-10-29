# clean

Standardize names, dates, and phone numbers; drop duplicates and empty rows via YAML rules.

## Story

Nina is a marketing analyst at **Toko Sumber Makmur**. Weekend booth signups arrive with
random caps, three date formats, and phone numbers that lost their leading zero in Excel.
She needs a dependable list before the SMS campaign goes out. On the sample sheet, **10**
dirty rows become **8** clean rows (measured on sample data).

## Input

An Excel/CSV/Parquet file plus a YAML config under `configs/clean.yaml` (or `-c`) describing
which columns to clean and how.

## Command

```bash
pfw sample clean -o ./sample/data_kotor.xlsx
pfw clean -c configs/clean.yaml
# or point input/output inside a custom YAML
```

## Output

A cleaned workbook with title-cased names, normalized dates, dash-formatted phones that keep
every digit, duplicates removed, and empty name rows dropped.
