"""Clean and standardize names, dates, and phone numbers using YAML rules."""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd
from dateutil import parser as date_parser

from pfw.common.config import load_config, resolve_path
from pfw.common.io import read_table, write_table
from pfw.common.logging_setup import get_logger, setup_logging
from pfw.common.phone import format_phone, normalize_phone

log = get_logger("clean")

DEFAULT_CONFIG = Path("configs") / "clean.yaml"

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
    """Convert an Indonesian-style date string to a parseable form."""
    date_str = str(date_str).strip().lower()

    for bulan_indo, bulan_num in BULAN_INDONESIA.items():
        if bulan_indo in date_str:
            date_str = date_str.replace(bulan_indo, bulan_num)
            parts = date_str.split()
            if len(parts) == 3:
                return f"{parts[0]}-{parts[1]}-{parts[2]}"
    return date_str


def clean_nama(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Standardize name formatting."""
    kolom = config.get("kolom")
    if kolom not in df.columns:
        log.warning("  Warning: column '%s' not found, skipping name cleaning", kolom)
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
    log.info("  Names: %s value(s) standardized", changed)

    return df


def clean_tanggal(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Standardize date formatting."""
    kolom = config.get("kolom")
    if kolom not in df.columns:
        log.warning("  Warning: column '%s' not found, skipping date cleaning", kolom)
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
    log.info("  Dates: %s value(s) formatted to %s", fixed_count, output_format)

    return df


def clean_telepon(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Standardize phone numbers without truncating digits."""
    kolom = config.get("kolom")
    if kolom not in df.columns:
        log.warning("  Warning: column '%s' not found, skipping phone cleaning", kolom)
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

    log.info("  Phones: %s value(s) formatted", fixed_count)
    if invalid_count:
        log.warning(
            "  Warning: %s number(s) invalid (must be 10-15 digits); original values kept",
            invalid_count,
        )

    return df


def remove_duplicates(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Drop duplicate rows based on the configured columns."""
    kolom = config.get("kolom", [])

    missing = [k for k in kolom if k not in df.columns]
    if missing:
        log.warning("  Warning: column(s) %s not found, skipping duplicate removal", missing)
        return df

    before = len(df)
    df = df.drop_duplicates(subset=kolom, keep="first")
    removed = before - len(df)

    stats["duplikat_removed"] = removed
    log.info("  Duplicates: %s row(s) removed", removed)

    return df


def remove_empty(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Drop rows that have empty values in the configured columns."""
    kolom = config.get("kolom", [])

    missing = [k for k in kolom if k not in df.columns]
    if missing:
        log.warning("  Warning: column(s) %s not found, skipping empty-row removal", missing)
        return df

    before = len(df)
    df = df.dropna(subset=kolom)
    for k in kolom:
        df = df[df[k].astype(str).str.strip() != ""]
    removed = before - len(df)

    stats["empty_removed"] = removed
    log.info("  Empty rows: %s row(s) removed", removed)

    return df


def phone_columns(cleaning: dict) -> tuple[str, ...]:
    """Return phone column names from the cleaning config."""
    telepon = cleaning.get("telepon") or {}
    kolom = telepon.get("kolom")
    return (kolom,) if kolom else ()


def clean_data(config_path: str | Path) -> pd.DataFrame:
    """Clean data using a YAML configuration.

    ``input`` and ``output`` paths in the config are resolved relative to the
    current working directory (absolute paths are left unchanged).
    """
    config_file = Path(config_path)
    if not config_file.exists():
        raise FileNotFoundError(f"Config file '{config_path}' not found")

    config = load_config(config_file)

    if "input" not in config or "output" not in config:
        raise ValueError("Config must include 'input' and 'output'")

    input_file = resolve_path(Path.cwd(), config["input"])
    output_file = resolve_path(Path.cwd(), config["output"])
    cleaning = config.get("cleaning", {})
    text_columns = phone_columns(cleaning)

    if not input_file.exists():
        raise FileNotFoundError(f"File '{input_file}' not found")

    log.info("Reading file: %s", input_file)
    df = read_table(input_file, text_columns=text_columns)
    log.info("Total rows: %s", len(df))

    stats: dict = {}
    log.info("")
    log.info("Cleaning steps:")

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

    write_table(df, output_file, text_columns=text_columns)

    log.info("")
    log.info("Done! Clean data saved to: %s", output_file)
    log.info("Total rows after cleaning: %s", len(df))

    return df


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--config",
        "-c",
        help="YAML config file (default: configs/clean.yaml)",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pfw clean",
        description="Clean tabular data using rules from a YAML config.",
    )
    add_arguments(parser)
    return parser


def run(args: argparse.Namespace) -> int:
    setup_logging()
    config_path = Path(args.config) if args.config else DEFAULT_CONFIG
    try:
        clean_data(config_path)
    except (FileNotFoundError, ValueError) as exc:
        log.error("Error: %s", exc)
        return 1
    return 0


def main(argv: list[str] | None = None) -> int:
    return run(build_parser().parse_args(argv))
