"""Tests for pfw.load (DuckDB in CI; Postgres skipped without env)."""

import os

import duckdb
import pandas as pd
import pytest

from pfw.load import load_table
from pfw.sample import build_load_data, generate_load

POSTGRES_ENV = "PFW_DATABASE_URL"


def _frame():
    return pd.DataFrame(
        {
            "sku": ["A1", "A2", "A3"],
            "qty": [1, 2, 3],
            "warehouse": ["Jakarta", "Bandung", "Surabaya"],
        }
    )


def test_replace_creates_duckdb_table(tmp_path):
    source = tmp_path / "sales.csv"
    _frame().to_csv(source, index=False)
    db = tmp_path / "pfw.duckdb"

    result = load_table(source, "sales_daily", db_path=db, mode="replace")

    assert result["written"] == 3
    assert result["total"] == 3
    conn = duckdb.connect(str(db))
    try:
        rows = conn.execute("SELECT sku, qty FROM sales_daily ORDER BY sku").fetchall()
    finally:
        conn.close()
    assert rows == [("A1", 1), ("A2", 2), ("A3", 3)]


def test_append_adds_rows(tmp_path):
    source = tmp_path / "sales.csv"
    _frame().to_csv(source, index=False)
    db = tmp_path / "pfw.duckdb"

    load_table(source, "sales_daily", db_path=db, mode="replace")
    result = load_table(source, "sales_daily", db_path=db, mode="append")

    assert result["written"] == 3
    assert result["total"] == 6


def test_upsert_updates_existing_keys(tmp_path):
    source = tmp_path / "sales.csv"
    _frame().to_csv(source, index=False)
    db = tmp_path / "pfw.duckdb"

    load_table(source, "sales_daily", db_path=db, mode="upsert", keys=["sku"])

    updated = pd.DataFrame(
        {
            "sku": ["A2", "A4"],
            "qty": [99, 4],
            "warehouse": ["Medan", "Makassar"],
        }
    )
    updated_path = tmp_path / "sales_update.csv"
    updated.to_csv(updated_path, index=False)

    result = load_table(
        updated_path, "sales_daily", db_path=db, mode="upsert", keys=["sku"]
    )

    assert result["written"] == 2
    assert result["total"] == 4
    conn = duckdb.connect(str(db))
    try:
        rows = dict(
            conn.execute("SELECT sku, qty FROM sales_daily ORDER BY sku").fetchall()
        )
    finally:
        conn.close()
    assert rows == {"A1": 1, "A2": 99, "A3": 3, "A4": 4}


def test_upsert_requires_keys(tmp_path):
    source = tmp_path / "sales.csv"
    _frame().to_csv(source, index=False)

    with pytest.raises(ValueError, match="requires --keys"):
        load_table(source, "sales_daily", db_path=tmp_path / "x.duckdb", mode="upsert")


def test_invalid_table_name_rejected(tmp_path):
    source = tmp_path / "sales.csv"
    _frame().to_csv(source, index=False)

    with pytest.raises(ValueError, match="Invalid table name"):
        load_table(source, "sales-daily", db_path=tmp_path / "x.duckdb")


def test_sample_has_20_rows():
    frame = build_load_data()
    assert len(frame) == 20
    assert list(frame.columns) == [
        "sku",
        "product",
        "qty",
        "unit_price",
        "total",
        "warehouse",
    ]


def test_generate_load_writes_csv(tmp_path):
    path = generate_load(tmp_path / "sales_daily.csv")
    assert path.exists()
    assert len(pd.read_csv(path)) == 20


@pytest.mark.skipif(
    not os.environ.get(POSTGRES_ENV),
    reason=f"{POSTGRES_ENV} not set",
)
def test_postgres_replace_when_url_present(tmp_path):
    source = tmp_path / "sales.csv"
    _frame().to_csv(source, index=False)

    result = load_table(
        source,
        "pfw_sales_daily_ci",
        backend="postgres",
        mode="replace",
        postgres_env=POSTGRES_ENV,
    )
    assert result["backend"] == "postgres"
    assert result["written"] == 3
    assert result["total"] == 3
