"""Tests for pfw.clean."""

import pandas as pd
import yaml

from pfw.clean import clean_data, clean_telepon
from pfw.common.io import read_table
from pfw.common.phone import format_phone, normalize_phone
from pfw.sample import generate_clean


def test_normalize_and_format_keep_every_digit():
    digits = normalize_phone("62 814 567 8902")

    assert digits == "08145678902"
    assert format_phone(digits) == "0814-567-8902"
    assert format_phone(digits).replace("-", "") == digits


def test_clean_telepon_preserves_invalid_values():
    df = pd.DataFrame({"No HP": ["12345", "081234567890", "+62 812 3456 7891"]})
    stats = {}

    result = clean_telepon(df, {"kolom": "No HP"}, stats)

    assert result["No HP"].tolist() == ["12345", "0812-3456-7890", "0812-3456-7891"]
    assert stats["telepon_invalid"] == 1


def test_clean_data_pipeline(tmp_path):
    input_path = generate_clean(tmp_path / "sample" / "data_kotor.xlsx")
    output_path = tmp_path / "output" / "data_bersih.xlsx"

    config = {
        "input": str(input_path),
        "output": str(output_path),
        "cleaning": {
            "nama": {"kolom": "Nama Lengkap", "format": "title", "trim": True},
            "tanggal": {"kolom": "Tanggal Lahir", "format": "%d-%m-%Y"},
            "telepon": {"kolom": "No HP", "format": "0xxx-xxxx-xxxx"},
            "duplikat": {"kolom": ["Nama Lengkap", "No HP"]},
            "hapus_kosong": {"kolom": ["Nama Lengkap"]},
        },
    }
    config_path = tmp_path / "config.yaml"
    config_path.write_text(yaml.safe_dump(config), encoding="utf-8")

    df = clean_data(config_path)

    assert len(df) == 8
    assert "Budi Santoso" in df["Nama Lengkap"].tolist()
    assert "" not in df["Nama Lengkap"].tolist()

    phones = df["No HP"].tolist()
    assert "0812-3456-7890" in phones
    assert "0814-567-8902" in phones
    assert all(phone.count("-") >= 1 for phone in phones if phone)

    written = read_table(output_path, text_columns=("No HP",))
    assert written["No HP"].iloc[0].startswith("0")
