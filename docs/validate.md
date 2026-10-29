# validate

Validate rows against YAML rules before import, with a spreadsheet findings report.

## Story

Andi handles program admin at **Lembaga Pelatihan Cendekia**. Partner schools send
registration files; the CRM rejects whole uploads for one bad email, short phone, impossible
birth date, or NIK that is not 16 digits. He defines rules once in YAML and gets every
finding with its Excel row number. The sample has **12** rows → **5** valid, **7** with
errors, **17** findings (measured on sample data).

## Why custom instead of Pandera or Great Expectations?

Those libraries are excellent for full data-quality platforms, but they pull large dependency
trees and lean toward DataFrame schemas or warehouse workflows. This utility stays a single
lightweight YAML file plus pandas/openpyxl: Indonesian phone digit rules, NIK regex, and an
spreadsheet report with highlighted cells — enough for ops hand-offs without standing up a DQ
stack. Reach for Pandera/GE when you need typed schemas, suite versioning, or warehouse
orchestration.

## Input

A registration-style sheet and `configs/validate.yaml` (or `-c`) listing per-column rules:
`required`, `email`, `phone`, `date_range`, `number_range`, `regex`, `in_list`, `unique`.

## Command

```bash
pfw sample validate -o ./sample/data_pendaftaran.xlsx
pfw validate ./sample/data_pendaftaran.xlsx -c configs/validate.yaml -o ./output/report.xlsx
```

## Output

Console summary plus an optional multi-sheet spreadsheet report (Summary, Detail Error, Data with
error cells highlighted). Phone columns stay text so leading zeros survive.
