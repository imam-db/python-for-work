"""Tests for pfw.common.io."""

import pandas as pd
import pytest

from pfw.common.io import read_table, write_table

PHONE = "081234567890"


@pytest.mark.parametrize("suffix", [".xlsx", ".csv", ".parquet"])
def test_roundtrip_preserves_leading_zero(tmp_path, suffix):
    path = tmp_path / f"phone{suffix}"

    write_table(
        pd.DataFrame({"No HP": [PHONE], "Nama": ["Budi"]}),
        path,
        text_columns=("No HP",),
    )
    df = read_table(path, text_columns=("No HP",))

    assert str(df["No HP"].iloc[0]) == PHONE


def test_csv_write_has_no_index_column(tmp_path):
    path = tmp_path / "data.csv"

    write_table(pd.DataFrame({"Nama": ["Budi"]}), path)

    assert path.read_text(encoding="utf-8").splitlines() == ["Nama", "Budi"]


def test_unknown_extension_raises(tmp_path):
    path = tmp_path / "data.txt"

    with pytest.raises(ValueError, match="tidak didukung"):
        read_table(path)

    with pytest.raises(ValueError, match="tidak didukung"):
        write_table(pd.DataFrame({"a": [1]}), path)
