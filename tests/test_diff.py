"""Tests for pfw.diff."""

import logging

import pandas as pd

from pfw.diff import compare_dataframes, compare_excel
from pfw.sample import generate_diff


def test_compare_creates_missing_output_dir(tmp_path):
    file_old, file_new = generate_diff(tmp_path / "sample")

    output = tmp_path / "belum" / "ada" / "laporan.xlsx"
    result = compare_excel(str(file_old), str(file_new), "No", str(output))

    assert result["summary"] == {
        "added_count": 2,
        "deleted_count": 2,
        "changed_count": 6,
    }
    assert output.exists()

    deleted_keys = [item["key"] for item in result["deleted"]]
    assert deleted_keys == ["9", "10"]


def test_duplicate_keys_emit_warning(caplog):
    df_old = pd.DataFrame({"No": [1, 1, 2], "Total": [10, 11, 20]})
    df_new = pd.DataFrame({"No": [1, 2], "Total": [10, 20]})

    with caplog.at_level(logging.WARNING):
        result = compare_dataframes(df_old, df_new, "No")

    assert result["summary"]["changed_count"] == 0
    assert any("duplicate" in record.message.lower() for record in caplog.records)
