# Python for Work Utilities

Five small Python utilities for everyday Excel work — the repetitive cleanup that quietly
eats an afternoon. They are built for analysts, admins, and operations staff who already know
Excel but would rather not copy-paste the same monthly report twenty times.

Each utility lives in its own folder, uses pandas + openpyxl, ships synthetic sample data,
and can be tried in one command. Shared helpers (config loading, phone normalization,
text-safe Excel I/O) live in `common/`. Requires Python 3.10+.

## Utilities

| # | Utility | What it does | Sample result |
|---|---------|--------------|---------------|
| 1 | [01-merge-excel](./01-merge-excel/) | Merge every workbook in a folder into one file | 3 files → **272 rows** (72 + 120 + 80) |
| 2 | [02-data-cleaner](./02-data-cleaner/) | Standardize names, dates, and phones; drop duplicates and empty rows | 10 dirty rows → **8 clean rows** (1 duplicate + 1 empty removed) |
| 3 | [03-excel-splitter](./03-excel-splitter/) | Split one workbook into per-value files | 100 rows → **5 files** (Jakarta 23, Makassar 22, Bandung 20, Surabaya 20, Medan 15) |
| 4 | [04-excel-comparator](./04-excel-comparator/) | Diff two workbook versions cell by cell | **2 added**, **2 deleted**, **6 cells changed** |
| 5 | [05-data-validator](./05-data-validator/) | Validate rows against YAML rules before import | 12 rows → **5 valid**, **7 with errors**, **17 findings** |

The numbers above come from running each demo on the included sample data, not from
estimates.

## Quickstart

```bash
git clone <repository-url>
cd python-for-work

python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

pip install -r requirements.txt

# Try one demo
cd 01-merge-excel
python generate_sample.py
python merge_excel.py ./sample
```

Every utility has its own `generate_sample.py`, so no Excel file needs to be prepared
beforehand. Generators are safe to run from the utility folder or from the repo root:

```bash
python 03-excel-splitter/generate_sample.py
python 03-excel-splitter/excel_splitter.py            # uses 03-excel-splitter/config.yaml
```

By default, config paths resolve against the utility folder while paths passed on the
command line resolve against your current working directory.

## Repository layout

```text
common/                Shared helpers (config, Excel I/O, phone, logging)
01-merge-excel/        Merge workbooks from a folder
02-data-cleaner/       Standardize and deduplicate rows
03-excel-splitter/     Split one workbook by a column
04-excel-comparator/   Diff two workbook versions
05-data-validator/     Validate rows against YAML rules
tests/                 pytest suite for all utilities
.github/workflows/     CI: ruff + pytest
```

## Development and tests

```bash
pip install -r requirements.txt pytest ruff
ruff check .
pytest -q
```

The GitHub Actions workflow runs the same lint and test commands on every push and pull
request.

## License

MIT License. See [LICENSE](./LICENSE).
