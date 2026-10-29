"""Tests for pfw.merge."""

import logging
from pathlib import Path

import pandas as pd

from pfw.merge import merge_excel_files
from pfw.sample import generate_merge

ROW = {
    "No": 1,
    "Tanggal": "2024-01-05",
    "Nama Sales": "Budi Santoso",
    "Produk": "Laptop ASUS ROG",
    "Qty": 2,
    "Harga Satuan": 16500000,
    "Total": 33000000,
    "Status": "Lunas",
    "Metode Bayar": "Transfer",
}


def write_branch(folder: Path, name: str, rows: int = 1, drop_column: str | None = None):
    data = [dict(ROW, No=index + 1) for index in range(rows)]
    df = pd.DataFrame(data)
    if drop_column:
        df = df.drop(columns=[drop_column])
    df.to_excel(folder / f"{name}.xlsx", index=False)


def test_merge_two_files(tmp_path):
    source = tmp_path / "inbox"
    source.mkdir()
    write_branch(source, "cabang_jakarta")
    write_branch(source, "cabang_bandung")

    output = tmp_path / "output" / "gabungan.xlsx"
    total = merge_excel_files(str(source), str(output))

    assert total == 2
    merged = pd.read_excel(output)
    assert len(merged) == 2
    assert set(merged["Cabang"]) == {"Jakarta", "Bandung"}


def test_schema_mismatch_warns_and_still_merges(tmp_path, caplog):
    source = tmp_path / "inbox"
    source.mkdir()
    write_branch(source, "cabang_jakarta")
    write_branch(source, "cabang_bandung", drop_column="Status")

    output = tmp_path / "output" / "gabungan.xlsx"
    with caplog.at_level(logging.WARNING):
        total = merge_excel_files(str(source), str(output))

    assert total == 2
    assert any("kolom" in record.message.lower() for record in caplog.records)

    merged = pd.read_excel(output)
    bandung = merged[merged["Cabang"] == "Bandung"]
    assert bandung["Status"].isna().all()


def test_empty_file_is_skipped_with_warning(tmp_path, caplog):
    source = tmp_path / "inbox"
    source.mkdir()
    write_branch(source, "cabang_jakarta")
    pd.DataFrame(columns=list(ROW)).to_excel(source / "cabang_kosong.xlsx", index=False)

    output = tmp_path / "output" / "gabungan.xlsx"
    with caplog.at_level(logging.WARNING):
        total = merge_excel_files(str(source), str(output))

    assert total == 1
    assert any("kosong" in record.message.lower() for record in caplog.records)


def test_generate_sample_writes_to_default_folder(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    written = generate_merge()

    assert len(written) == 3
    for path in written:
        resolved = Path(path).resolve()
        assert resolved.parent == tmp_path / "sample"
        assert resolved.exists()


def test_generated_sample_has_272_rows(tmp_path):
    written = generate_merge(tmp_path / "sample")

    total = sum(len(pd.read_excel(path)) for path in written)

    assert total == 272
