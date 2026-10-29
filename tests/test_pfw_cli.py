"""Tests for the unified ``pfw`` CLI."""

import os
import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest
import yaml

ROOT = Path(__file__).resolve().parents[1]

UTILITY_COMMANDS = ["merge", "clean", "split", "compare", "validate"]


def run_pfw(*args, cwd=None) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    env["PYTHONPATH"] = os.pathsep.join(
        filter(None, [str(ROOT), env.get("PYTHONPATH")]),
    )
    return subprocess.run(
        [sys.executable, "-m", "pfw", *map(str, args)],
        cwd=str(cwd or ROOT),
        env=env,
        capture_output=True,
        text=True,
    )


@pytest.mark.parametrize("command", UTILITY_COMMANDS)
def test_subcommand_help_exits_zero(command):
    result = run_pfw(command, "--help")

    assert result.returncode == 0
    assert "usage" in result.stdout.lower()


def test_sample_help_exits_zero():
    result = run_pfw("sample", "--help")

    assert result.returncode == 0
    assert "usage" in result.stdout.lower()
    for name in UTILITY_COMMANDS:
        assert name in result.stdout


def test_merge_sample(tmp_path):
    sample_dir = tmp_path / "sample"
    assert run_pfw("sample", "merge", "-o", sample_dir).returncode == 0

    output = tmp_path / "merged.xlsx"
    result = run_pfw("merge", sample_dir, output)

    assert result.returncode == 0, result.stderr
    assert len(pd.read_excel(output)) == 272


def test_clean_sample(tmp_path):
    dirty = tmp_path / "data_kotor.xlsx"
    assert run_pfw("sample", "clean", "-o", dirty).returncode == 0

    output = tmp_path / "data_bersih.xlsx"
    config = {
        "input": str(dirty),
        "output": str(output),
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

    result = run_pfw("clean", "-c", config_path, cwd=tmp_path)

    assert result.returncode == 0, result.stderr
    assert len(pd.read_excel(output)) == 8


def test_split_sample(tmp_path):
    source = tmp_path / "laporan.xlsx"
    assert run_pfw("sample", "split", "-o", source).returncode == 0

    output = tmp_path / "split"
    result = run_pfw("split", source, "--kolom", "Cabang", "--output", output)

    assert result.returncode == 0, result.stderr
    written = {path.stem for path in output.glob("*.xlsx")}
    assert written == {"Jakarta", "Bandung", "Surabaya", "Medan", "Makassar"}


def test_compare_sample(tmp_path):
    sample_dir = tmp_path / "sample"
    assert run_pfw("sample", "compare", "-o", sample_dir).returncode == 0

    report = tmp_path / "report.xlsx"
    result = run_pfw(
        "compare",
        sample_dir / "data_lama.xlsx",
        sample_dir / "data_revisi.xlsx",
        "--key",
        "No",
        "-o",
        report,
    )

    assert result.returncode == 0, result.stderr
    assert report.exists()
    assert "Baris baru     : 2" in result.stdout
    assert "Cell berubah   : 6" in result.stdout


def test_validate_sample(tmp_path):
    source = tmp_path / "data_pendaftaran.xlsx"
    assert run_pfw("sample", "validate", "-o", source).returncode == 0

    report = tmp_path / "report.xlsx"
    result = run_pfw("validate", source, "--output", report)

    assert result.returncode == 0, result.stderr
    assert report.exists()
    assert "Error       : 7 baris (17 masalah)" in result.stdout


def test_classic_script_invocation_still_works():
    result = subprocess.run(
        [sys.executable, str(ROOT / "01-merge-excel" / "merge_excel.py"), "--help"],
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0
    assert "input_folder" in result.stdout
