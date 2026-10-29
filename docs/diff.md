# diff

Compare two tabular file versions and report exactly what changed.

## Story

Citra works in finance at **PT Arta Graha**. A colleague returns the month-end file with
"a few corrections — please check," and no change log. She needs added, deleted, and changed
cells before approving. The sample shows **2** added, **2** deleted, and **6** changed cells
(measured on sample data).

## Input

Two tabular files (old and new). Prefer a key column (`--key`) for stable row matching;
without it, comparison is by row position.

## Command

```bash
pfw sample diff -o ./sample
pfw diff ./sample/data_lama.xlsx ./sample/data_revisi.xlsx --key No -o ./output/laporan.xlsx
```

## Output

A console report plus an optional highlighted spreadsheet report (Summary, Changes, New Rows,
Deleted Rows).
