"""
Data Cleaner - Membersihkan data Excel berdasarkan konfigurasi YAML
"""

import argparse
import sys
from pathlib import Path

import pandas as pd
from dateutil import parser as date_parser

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from common.config import load_config, resolve_path  # noqa: E402
from common.excel import read_excel, write_excel  # noqa: E402
from common.logging_setup import get_logger, setup_logging  # noqa: E402
from common.phone import format_phone, normalize_phone  # noqa: E402

log = get_logger("data_cleaner")


# Mapping bulan Indonesia ke angka
BULAN_INDONESIA = {
    "januari": "01",
    "februari": "02",
    "maret": "03",
    "april": "04",
    "mei": "05",
    "juni": "06",
    "juli": "07",
    "agustus": "08",
    "september": "09",
    "oktober": "10",
    "november": "11",
    "desember": "12",
    "jan": "01",
    "feb": "02",
    "mar": "03",
    "apr": "04",
    "jun": "06",
    "jul": "07",
    "agu": "08",
    "ags": "08",
    "sep": "09",
    "okt": "10",
    "nov": "11",
    "des": "12",
}


def convert_indonesian_date(date_str: str) -> str:
    """Konversi tanggal format Indonesia ke format standar."""
    date_str = str(date_str).strip().lower()

    for bulan_indo, bulan_num in BULAN_INDONESIA.items():
        if bulan_indo in date_str:
            date_str = date_str.replace(bulan_indo, bulan_num)
            parts = date_str.split()
            if len(parts) == 3:
                return f"{parts[0]}-{parts[1]}-{parts[2]}"
    return date_str


def clean_nama(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Standardisasi format nama."""
    kolom = config.get("kolom")
    if kolom not in df.columns:
        log.warning("  Peringatan: Kolom '%s' tidak ditemukan, skip cleaning nama", kolom)
        return df

    format_type = config.get("format", "title")
    trim = config.get("trim", True)

    original = df[kolom].copy()

    if trim:
        df[kolom] = df[kolom].astype(str).str.strip()
        df[kolom] = df[kolom].str.replace(r"\s+", " ", regex=True)

    if format_type == "title":
        df[kolom] = df[kolom].str.title()
    elif format_type == "upper":
        df[kolom] = df[kolom].str.upper()
    elif format_type == "lower":
        df[kolom] = df[kolom].str.lower()

    changed = int((original != df[kolom]).sum())
    stats["nama_fixed"] = changed
    log.info("  Nama: %s data di-standardisasi", changed)

    return df


def clean_tanggal(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Standardisasi format tanggal."""
    kolom = config.get("kolom")
    if kolom not in df.columns:
        log.warning("  Peringatan: Kolom '%s' tidak ditemukan, skip cleaning tanggal", kolom)
        return df

    output_format = config.get("format", "%d-%m-%Y")
    fixed_count = 0

    def parse_date(val):
        nonlocal fixed_count
        if pd.isna(val) or str(val).strip() == "":
            return val
        try:
            val_converted = convert_indonesian_date(str(val))
            parsed = date_parser.parse(val_converted, dayfirst=True)
            fixed_count += 1
            return parsed.strftime(output_format)
        except (ValueError, TypeError, OverflowError):
            return val

    df[kolom] = df[kolom].apply(parse_date)
    stats["tanggal_fixed"] = fixed_count
    log.info("  Tanggal: %s data di-format ke %s", fixed_count, output_format)

    return df


def clean_telepon(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Standardisasi format nomor telepon tanpa memotong digit."""
    kolom = config.get("kolom")
    if kolom not in df.columns:
        log.warning("  Peringatan: Kolom '%s' tidak ditemukan, skip cleaning telepon", kolom)
        return df

    output_format = config.get("format", "0xxx-xxxx-xxxx")
    fixed_count = 0
    invalid_count = 0

    def format_phone_value(val):
        nonlocal fixed_count, invalid_count
        if pd.isna(val) or str(val).strip() == "":
            return val

        digits = normalize_phone(val)
        if digits is None:
            invalid_count += 1
            return val

        if output_format == "+62xxx-xxxx-xxxx":
            formatted = "+62" + format_phone(digits)[1:]
        else:
            formatted = format_phone(digits)

        fixed_count += 1
        return formatted

    df[kolom] = df[kolom].apply(format_phone_value)
    stats["telepon_fixed"] = fixed_count
    stats["telepon_invalid"] = invalid_count

    log.info("  Telepon: %s data di-format", fixed_count)
    if invalid_count:
        log.warning(
            "  Peringatan: %s nomor tidak valid (harus 10-15 digit), nilai asli dipertahankan",
            invalid_count,
        )

    return df


def remove_duplicates(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Hapus baris duplikat berdasarkan kolom tertentu."""
    kolom = config.get("kolom", [])

    missing = [k for k in kolom if k not in df.columns]
    if missing:
        log.warning("  Peringatan: Kolom %s tidak ditemukan, skip hapus duplikat", missing)
        return df

    before = len(df)
    df = df.drop_duplicates(subset=kolom, keep="first")
    removed = before - len(df)

    stats["duplikat_removed"] = removed
    log.info("  Duplikat: %s baris dihapus", removed)

    return df


def remove_empty(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Hapus baris dengan kolom kosong."""
    kolom = config.get("kolom", [])

    missing = [k for k in kolom if k not in df.columns]
    if missing:
        log.warning("  Peringatan: Kolom %s tidak ditemukan, skip hapus kosong", missing)
        return df

    before = len(df)
    df = df.dropna(subset=kolom)
    for k in kolom:
        df = df[df[k].astype(str).str.strip() != ""]
    removed = before - len(df)

    stats["empty_removed"] = removed
    log.info("  Baris kosong: %s baris dihapus", removed)

    return df


def phone_columns(cleaning: dict) -> tuple[str, ...]:
    """Ambil nama kolom telepon dari konfigurasi cleaning."""
    telepon = cleaning.get("telepon") or {}
    kolom = telepon.get("kolom")
    return (kolom,) if kolom else ()


def clean_data(config_path: str) -> pd.DataFrame:
    """Main function untuk membersihkan data."""
    config_file = Path(config_path)
    if not config_file.exists():
        raise FileNotFoundError(f"File konfigurasi '{config_path}' tidak ditemukan")

    config = load_config(config_file)
    script_dir = Path(__file__).parent

    if "input" not in config or "output" not in config:
        raise ValueError("Config harus berisi 'input' dan 'output'")

    input_file = resolve_path(script_dir, config["input"])
    output_file = resolve_path(script_dir, config["output"])
    cleaning = config.get("cleaning", {})
    text_columns = phone_columns(cleaning)

    if not input_file.exists():
        raise FileNotFoundError(f"File '{input_file}' tidak ditemukan")

    log.info("Membaca file: %s", input_file)
    df = read_excel(input_file, text_columns=text_columns)
    log.info("Total baris: %s", len(df))

    stats = {}
    log.info("")
    log.info("Proses cleaning:")

    if "nama" in cleaning:
        df = clean_nama(df, cleaning["nama"], stats)
    if "tanggal" in cleaning:
        df = clean_tanggal(df, cleaning["tanggal"], stats)
    if "telepon" in cleaning:
        df = clean_telepon(df, cleaning["telepon"], stats)
    if "duplikat" in cleaning:
        df = remove_duplicates(df, cleaning["duplikat"], stats)
    if "hapus_kosong" in cleaning:
        df = remove_empty(df, cleaning["hapus_kosong"], stats)

    write_excel(df, output_file, text_columns=text_columns)

    log.info("")
    log.info("Berhasil! Data bersih disimpan ke: %s", output_file)
    log.info("Total baris setelah cleaning: %s", len(df))

    return df


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Membersihkan data Excel berdasarkan aturan di config YAML."
    )
    parser.add_argument(
        "config_file",
        nargs="?",
        help="File konfigurasi YAML (default: config.yaml di folder utility)",
    )
    parser.add_argument(
        "--config",
        "-c",
        dest="config_option",
        help="File konfigurasi YAML (alternatif dari positional)",
    )
    return parser


def main(argv: list[str] | None = None) -> None:
    setup_logging()
    args = build_parser().parse_args(argv)

    explicit_config = args.config_option or args.config_file
    if explicit_config:
        config_path = Path(explicit_config)
        if not config_path.is_absolute():
            config_path = Path.cwd() / config_path
    else:
        config_path = Path(__file__).parent / "config.yaml"

    try:
        clean_data(str(config_path))
    except (FileNotFoundError, ValueError) as exc:
        log.error("Error: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
