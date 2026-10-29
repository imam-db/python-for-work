"""Test untuk utility 04-excel-comparator."""

import logging

import excel_comparator
import pandas as pd


def test_compare_creates_missing_output_dir(tmp_path, load_utility):
    generate_sample = load_utility("04-excel-comparator", "generate_sample.py")
    file_old, file_new = generate_sample.generate_sample_data(tmp_path / "sample")

    output = tmp_path / "belum" / "ada" / "laporan.xlsx"
    result = excel_comparator.compare_excel(str(file_old), str(file_new), "No", str(output))

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
        result = excel_comparator.compare_dataframes(df_old, df_new, "No")

    assert result["summary"]["changed_count"] == 0
    assert any("duplikat" in record.message.lower() for record in caplog.records)
