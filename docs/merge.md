# merge

Combine every `.xlsx` workbook in a folder into a single file.

## Story

Raka works in operations at **PT Sinar Teknologi Nusantara**, an IT distributor. Each city
branch sends its own sales workbook at month-start; before the Monday directors' meeting he
needs one sheet, not a folder of tabs. The sample merges three branch files into **272**
transactions (measured on the included sample data).

## Input

A folder of `.xlsx` files that share (or nearly share) the same columns. Empty files are
skipped; mismatched columns still merge with a warning and NaN fill.

## Command

```bash
pfw sample merge -o ./sample
pfw merge ./sample ./output/hasil_gabungan.xlsx
```

## Output

One workbook with an added `Cabang` column derived from each file name, plus a console
summary of rows merged and any skipped or mismatched files.
