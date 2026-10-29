"""Generate deterministic sample data for every utility."""

from __future__ import annotations

import argparse
import random
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

from pfw.common.io import write_table
from pfw.common.logging_setup import get_logger, setup_logging

log = get_logger("sample")

DEFAULT_SAMPLE_DIR = Path("sample")

MERGE_PRODUCTS = [
    ("Laptop ASUS ROG", 15000000, 18000000),
    ("Laptop Lenovo ThinkPad", 12000000, 15000000),
    ("Monitor LG 27 inch", 3500000, 4500000),
    ("Monitor Samsung 24 inch", 2500000, 3200000),
    ("Keyboard Mechanical Logitech", 800000, 1200000),
    ("Mouse Wireless Logitech", 350000, 500000),
    ("Printer Epson L3210", 2800000, 3500000),
    ("Printer HP LaserJet", 4500000, 5500000),
    ("UPS APC 1100VA", 1500000, 1800000),
    ("Router WiFi TP-Link", 450000, 650000),
    ("Webcam Logitech C920", 1200000, 1500000),
    ("Headset Jabra Evolve", 2000000, 2500000),
    ("SSD Samsung 1TB", 1200000, 1500000),
    ("RAM DDR4 16GB", 600000, 800000),
    ("Kabel HDMI 2m", 50000, 100000),
]

MERGE_SALES_NAMES = [
    "Budi Santoso",
    "Dewi Lestari",
    "Ahmad Fauzi",
    "Siti Rahayu",
    "Rudi Hermawan",
    "Nina Kartika",
    "Eko Prasetyo",
    "Maya Sari",
    "Doni Wijaya",
    "Rina Putri",
]

MERGE_BRANCHES = {
    "cabang_jakarta": (80, 120),
    "cabang_bandung": (50, 80),
    "cabang_surabaya": (60, 90),
}

SPLIT_BRANCHES = ["Jakarta", "Bandung", "Surabaya", "Medan", "Makassar"]
SPLIT_SALES = [
    "Budi Santoso",
    "Dewi Lestari",
    "Ahmad Dahlan",
    "Siti Nurhaliza",
    "Rudi Hartono",
]
SPLIT_PRODUCTS = [
    ("Laptop ASUS ROG", 15000000),
    ("Monitor LG 27 inch", 4200000),
    ("Keyboard Mechanical", 1500000),
    ("Mouse Gaming", 800000),
    ("Headset Wireless", 1200000),
    ("Webcam HD", 900000),
    ("SSD 1TB", 1800000),
    ("RAM 16GB", 1200000),
]
SPLIT_STATUS = ["Lunas", "Cicilan", "Pending"]
SPLIT_METHODS = ["Transfer", "Cash", "Kartu Kredit", "Tempo 30 Hari"]


def build_merge_data(
    num_rows: int, start_date: datetime, rng: random.Random
) -> list[dict]:
    """Generate sales rows for one branch."""
    rows = []
    for index in range(num_rows):
        date = start_date + timedelta(days=rng.randint(0, 30))
        produk, min_price, max_price = rng.choice(MERGE_PRODUCTS)
        qty = rng.randint(1, 20)
        harga = rng.randint(min_price, max_price)
        rows.append(
            {
                "No": index + 1,
                "Tanggal": date.strftime("%Y-%m-%d"),
                "Nama Sales": rng.choice(MERGE_SALES_NAMES),
                "Produk": produk,
                "Qty": qty,
                "Harga Satuan": harga,
                "Total": qty * harga,
                "Status": rng.choice(["Lunas", "Lunas", "Lunas", "Cicilan", "Pending"]),
                "Metode Bayar": rng.choice(
                    ["Transfer", "Cash", "Kartu Kredit", "Tempo 30 Hari"]
                ),
            }
        )
    return rows


def generate_merge(output_dir: str | Path | None = None) -> list[Path]:
    """Generate one workbook per branch and return the written paths."""
    sample_dir = Path(output_dir) if output_dir else DEFAULT_SAMPLE_DIR
    sample_dir.mkdir(parents=True, exist_ok=True)

    rng = random.Random(42)
    start_date = datetime(2024, 1, 1)
    written = []

    for cabang, (min_rows, max_rows) in MERGE_BRANCHES.items():
        num_rows = rng.randint(min_rows, max_rows)
        frame = pd.DataFrame(build_merge_data(num_rows, start_date, rng))
        path = sample_dir / f"{cabang}.xlsx"
        write_table(frame, path)
        written.append(path)
        log.info("%s.xlsx: %s transaksi", cabang, num_rows)

    log.info("Done!")
    return written


def build_clean_data() -> pd.DataFrame:
    """Build a sheet with names, dates, and phones that need cleaning."""
    data = {
        "Nama Lengkap": [
            "  BUDI SANTOSO  ",
            "siti nurhaliza",
            "AHMAD   DAHLAN",
            "dewi  lestari",
            "Rudi Hartono",
            "BUDI SANTOSO",
            "  maya   sari  ",
            "",
            "joko widodo",
            "MEGAWATI SOEKARNO",
        ],
        "Tanggal Lahir": [
            "15/01/1990",
            "1985-03-20",
            "05 Mei 2000",
            "12-12-1995",
            "2001/07/25",
            "15/01/1990",
            "30 Juni 1988",
            "01-01-2000",
            "1975-08-17",
            "25/12/1980",
        ],
        "No HP": [
            "081234567890",
            "+62 812 3456 7891",
            "0813-4567-8901",
            "62 814 567 8902",
            "0815.4567.8903",
            "081234567890",
            "+62-816-4567-8904",
            "0817 4567 8905",
            "0818-4567-8906",
            "081945678907",
        ],
        "Email": [
            "budi@email.com",
            "siti@email.com",
            "ahmad@email.com",
            "dewi@email.com",
            "rudi@email.com",
            "budi@email.com",
            "maya@email.com",
            "",
            "joko@email.com",
            "mega@email.com",
        ],
        "Kota": [
            "Jakarta",
            "Bandung",
            "Surabaya",
            "Yogyakarta",
            "Semarang",
            "Jakarta",
            "Medan",
            "Makassar",
            "Denpasar",
            "Palembang",
        ],
    }
    return pd.DataFrame(data)


def generate_clean(output_path: str | Path | None = None) -> Path:
    """Write the dirty sample workbook with phone numbers kept as text."""
    path = Path(output_path) if output_path else DEFAULT_SAMPLE_DIR / "data_kotor.xlsx"
    frame = build_clean_data()
    write_table(frame, path, text_columns=("No HP",))
    log.info("Sample data kotor berhasil dibuat: %s", path)
    log.info("Total baris: %s", len(frame))
    return path


def build_split_data() -> pd.DataFrame:
    """Generate a national sales report (deterministic, seed 42)."""
    rng = random.Random(42)
    start_date = datetime(2024, 1, 1)
    rows = []

    for index in range(100):
        cabang = rng.choice(SPLIT_BRANCHES)
        sales = rng.choice(SPLIT_SALES)
        produk, harga = rng.choice(SPLIT_PRODUCTS)
        qty = rng.randint(1, 5)
        tanggal = start_date + timedelta(days=rng.randint(0, 30))
        rows.append(
            {
                "No": index + 1,
                "Tanggal": tanggal.strftime("%Y-%m-%d"),
                "Cabang": cabang,
                "Nama Sales": sales,
                "Produk": produk,
                "Qty": qty,
                "Harga Satuan": harga,
                "Total": qty * harga,
                "Status": rng.choice(SPLIT_STATUS),
                "Metode Bayar": rng.choice(SPLIT_METHODS),
            }
        )

    return pd.DataFrame(rows)


def generate_split(output_path: str | Path | None = None) -> Path:
    """Write the national report sample workbook."""
    path = (
        Path(output_path) if output_path else DEFAULT_SAMPLE_DIR / "laporan_nasional.xlsx"
    )
    frame = build_split_data()
    write_table(frame, path)
    log.info("Sample data berhasil dibuat: %s", path)
    log.info("Total baris: %s", len(frame))
    log.info("")
    log.info("Distribusi per cabang:")
    for cabang, count in frame["Cabang"].value_counts().items():
        log.info("  - %s: %s transaksi", cabang, count)
    return path


def build_diff_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Generate two versions of the same report with known differences."""
    data_lama = {
        "No": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
        "Nama": [
            "Budi Santoso",
            "Dewi Lestari",
            "Ahmad Dahlan",
            "Siti Nurhaliza",
            "Rudi Hartono",
            "Maya Sari",
            "Joko Widodo",
            "Mega Wati",
            "Andi Pratama",
            "Rina Susanti",
        ],
        "Cabang": [
            "Jakarta",
            "Bandung",
            "Surabaya",
            "Medan",
            "Makassar",
            "Jakarta",
            "Bandung",
            "Surabaya",
            "Medan",
            "Makassar",
        ],
        "Total": [
            5000000,
            3500000,
            4200000,
            2800000,
            3100000,
            4500000,
            3800000,
            2900000,
            3300000,
            4100000,
        ],
        "Status": [
            "Lunas",
            "Pending",
            "Lunas",
            "Cicilan",
            "Lunas",
            "Pending",
            "Lunas",
            "Pending",
            "Lunas",
            "Cicilan",
        ],
    }

    data_revisi = {
        "No": [1, 2, 3, 4, 5, 6, 7, 8, 11, 12],
        "Nama": [
            "Budi Santoso",
            "Dewi Lestari",
            "Ahmad Dahlan",
            "Siti Nurhaliza",
            "Rudi Hartono",
            "Maya Sari",
            "Joko Widodo",
            "Mega Wati",
            "Bambang Susilo",
            "Citra Dewi",
        ],
        "Cabang": [
            "Jakarta",
            "Bandung",
            "Surabaya",
            "Medan",
            "Makassar",
            "Jakarta",
            "Bandung",
            "Surabaya",
            "Semarang",
            "Yogyakarta",
        ],
        "Total": [
            5500000,
            3500000,
            4200000,
            2800000,
            3600000,
            4500000,
            3800000,
            2900000,
            2700000,
            3200000,
        ],
        "Status": [
            "Lunas",
            "Lunas",
            "Lunas",
            "Lunas",
            "Lunas",
            "Lunas",
            "Lunas",
            "Lunas",
            "Lunas",
            "Pending",
        ],
    }

    return pd.DataFrame(data_lama), pd.DataFrame(data_revisi)


def generate_diff(output_dir: str | Path | None = None) -> tuple[Path, Path]:
    """Write the old and revised sample workbooks."""
    sample_dir = Path(output_dir) if output_dir else DEFAULT_SAMPLE_DIR
    sample_dir.mkdir(parents=True, exist_ok=True)

    df_lama, df_revisi = build_diff_data()
    path_lama = sample_dir / "data_lama.xlsx"
    path_revisi = sample_dir / "data_revisi.xlsx"
    write_table(df_lama, path_lama)
    write_table(df_revisi, path_revisi)

    log.info("Sample data berhasil dibuat:")
    log.info("  - %s (%s baris)", path_lama, len(df_lama))
    log.info("  - %s (%s baris)", path_revisi, len(df_revisi))
    return path_lama, path_revisi


def build_validate_data() -> pd.DataFrame:
    """Build registration data containing known validation errors."""
    data = {
        "Nama Lengkap": [
            "Budi Santoso",
            "Dewi Lestari",
            "",
            "Siti Nurhaliza",
            "Rudi Hartono",
            "Maya Sari",
            "Joko Widodo",
            None,
            "Andi Pratama",
            "Rina Susanti",
            "Bambang Susilo",
            "Citra Dewi",
        ],
        "Email": [
            "budi@email.com",
            "dewi@email",
            "ahmad@email.com",
            "siti@email.com",
            "",
            "maya@email.com",
            "joko@@email.com",
            "mega@email.com",
            "andi@email.com",
            "rina@email.com",
            "bambang@email.com",
            "citra@email.com",
        ],
        "No HP": [
            "081234567890",
            "08123456",
            "081234567892",
            "081234567893",
            "081234567894",
            "0812",
            "081234567896",
            "081234567897",
            "081234567898",
            "081234567899",
            "081234567800",
            "081234567801",
        ],
        "Tanggal Lahir": [
            "1990-05-15",
            "1985-03-20",
            "1940-01-01",
            "1995-12-25",
            "2015-06-10",
            "1988-08-08",
            "1975-11-11",
            "1992-02-29",
            "2020-01-01",
            "1980-07-04",
            "1998-09-17",
            "2005-04-22",
        ],
        "Umur": [30, 35, 25, 28, 15, 32, 45, 70, 22, 40, 26, 19],
        "NIK": [
            "3201234567890001",
            "3201234567890002",
            "320123456789",
            "3201234567890004",
            "3201234567890005",
            "32012345678900067",
            "3201234567890007",
            "3201234567890008",
            "320123ABC7890009",
            "3201234567890010",
            "3201234567890011",
            "3201234567890012",
        ],
        "Status": [
            "Aktif",
            "Nonaktif",
            "Pending",
            "Aktif",
            "Tidak Valid",
            "Aktif",
            "Nonaktif",
            "Batal",
            "Pending",
            "Aktif",
            "Nonaktif",
            "Pending",
        ],
    }
    return pd.DataFrame(data)


def generate_validate(output_path: str | Path | None = None) -> Path:
    """Write the registration sample workbook with text columns preserved."""
    path = (
        Path(output_path)
        if output_path
        else DEFAULT_SAMPLE_DIR / "data_pendaftaran.xlsx"
    )
    frame = build_validate_data()
    write_table(frame, path, text_columns=("No HP", "NIK"))
    log.info("Sample data berhasil dibuat: %s", path)
    log.info("Total baris: %s", len(frame))
    return path


GENERATORS = {
    "merge": generate_merge,
    "clean": generate_clean,
    "split": generate_split,
    "diff": generate_diff,
    "validate": generate_validate,
}


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "name", choices=list(GENERATORS), help="Utility to generate data for"
    )
    parser.add_argument(
        "-o", "--output", help="Output folder or file (default: ./sample)"
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pfw sample",
        description="Generate sample data for a utility.",
    )
    add_arguments(parser)
    return parser


def run(args: argparse.Namespace) -> int:
    setup_logging()
    GENERATORS[args.name](args.output)
    return 0


def main(argv: list[str] | None = None) -> int:
    return run(build_parser().parse_args(argv))
