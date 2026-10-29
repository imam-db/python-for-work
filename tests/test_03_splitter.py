"""Test untuk utility 03-excel-splitter."""

import excel_splitter
import pandas as pd


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
    results = excel_splitter.split_excel(str(source), "Cabang", str(output))

    rows = {item["file"]: item["rows"] for item in results}
    assert rows["Jakarta.xlsx"] == 2
    assert rows["Bandung.xlsx"] == 1
    assert rows["_kosong.xlsx"] == 2

    empty_bucket = pd.read_excel(output / "_kosong.xlsx")
    assert empty_bucket["Total"].tolist() == [30, 40]
    assert empty_bucket["Total"].sum() == 70


def test_seeded_sample_is_reproducible(load_utility):
    generate_sample = load_utility("03-excel-splitter", "generate_sample.py")

    first = generate_sample.build_sample_data()
    second = generate_sample.build_sample_data()

    assert first.equals(second)
    assert len(first) == 100
