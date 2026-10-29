"""Tests for pfw.split."""

import pandas as pd

from pfw.sample import build_split_data, generate_split
from pfw.split import split_excel


def test_empty_split_values_go_to_empty_bucket(tmp_path):
    source = tmp_path / "laporan.xlsx"
    df = pd.DataFrame(
        {
            "Cabang": ["Jakarta", "Jakarta", "", None, "Bandung"],
            "Total": [10, 20, 30, 40, 50],
        }
    )
    df.to_excel(source, index=False)

    output = tmp_path / "output"
    results = split_excel(str(source), "Cabang", str(output))

    rows = {item["file"]: item["rows"] for item in results}
    assert rows["Jakarta.xlsx"] == 2
    assert rows["Bandung.xlsx"] == 1
    assert rows["_kosong.xlsx"] == 2

    empty_bucket = pd.read_excel(output / "_kosong.xlsx")
    assert empty_bucket["Total"].tolist() == [30, 40]
    assert empty_bucket["Total"].sum() == 70


def test_no_header_writes_through_io(tmp_path):
    source = tmp_path / "laporan.xlsx"
    pd.DataFrame(
        {
            "Cabang": ["Jakarta", "Bandung"],
            "Total": [10, 20],
        }
    ).to_excel(source, index=False)

    nested = tmp_path / "nested" / "out"
    results = split_excel(
        str(source),
        "Cabang",
        str(nested),
        include_header=False,
    )

    assert len(results) == 2
    jakarta = nested / "Jakarta.xlsx"
    assert jakarta.exists()
    frame = pd.read_excel(jakarta, header=None)
    assert list(frame.iloc[0]) == ["Jakarta", 10]
    assert len(frame) == 1


def test_seeded_sample_is_reproducible():
    first = build_split_data()
    second = build_split_data()

    assert first.equals(second)
    assert len(first) == 100


def test_generate_split_writes_all_rows(tmp_path):
    path = generate_split(tmp_path / "laporan.xlsx")

    frame = pd.read_excel(path)

    assert len(frame) == 100
    assert set(frame["Cabang"]) == {"Jakarta", "Bandung", "Surabaya", "Medan", "Makassar"}
