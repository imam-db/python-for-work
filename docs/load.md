# load

Load a tabular file into DuckDB (local file, default) or Postgres.

## Story

Sari runs fulfillment ops at **Gudang Nusantara Retail**. Every evening the WMS drops a
CSV of that day's picks; finance wants the same rows in a queryable store before the 8 a.m.
standup. She loads the extract into a local DuckDB file in replace mode on weekdays, and
upserts weekend corrections on `sku`. The sample has **20** sales rows (measured on sample
data).

## Input

One `.xlsx`, `.csv`, or `.parquet` file. DuckDB needs only a file path (`--db`). Postgres
reads its URL from an environment variable (`PFW_DATABASE_URL` by default) — never from a
hardcoded string or CLI flag that would leak credentials into shell history.

## Command

```bash
pfw sample load -o ./sample/sales_daily.csv
pfw load ./sample/sales_daily.csv --table sales_daily --db ./output/pfw.duckdb --mode replace

# Upsert corrections on sku
pfw load ./sample/sales_daily.csv --table sales_daily --mode upsert --keys sku

# Postgres (optional extra). Export the URL first; never commit it.
# export PFW_DATABASE_URL='postgresql://user:pass@host:5432/dbname'
# pip install 'python-for-work[postgres]'
# pfw load ./sample/sales_daily.csv --table sales_daily --backend postgres --mode append
```

## Modes

| Mode | Behavior |
|------|----------|
| `replace` | Drop the table if it exists, recreate, insert |
| `append` | Create the table if missing, then insert |
| `upsert` | Create with a primary key if missing, then upsert on `--keys` |

## Output

Rows written to the chosen backend, plus a console summary of how many rows were written and
the table's new total.
