"""Merge every workbook in a folder into a single Excel file."""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from pfw.common.io import read_table, write_table
from pfw.common.logging_setup import get_logger, setup_logging

log = get_logger("merge")

DEFAULT_OUTPUT = Path("output") / "hasil_gabungan.xlsx"


def merge_excel_files(input_folder: str, output_file: str) -> int:
    """Gabungkan semua file Excel dalam folder menjadi satu file.

    File kosong dilewati. File dengan kolom berbeda tetap digabung (outer
    concat) tetapi memunculkan peringatan agar NaN-nya terlihat jelas.
    """
    folder = Path(input_folder)

    if not folder.exists():
        raise FileNotFoundError(f"Folder '{input_folder}' tidak ditemukan")

    excel_files = sorted(folder.glob("*.xlsx"))
    if not excel_files:
        raise FileNotFoundError(f"Tidak ada file Excel di folder '{input_folder}'")

    log.info("Ditemukan %s file Excel di folder %s", len(excel_files), input_folder)

    all_data: list[pd.DataFrame] = []
    total_rows = 0
    reference_columns: list[str] | None = None
    skipped_empty = 0
    mismatched = 0

    for file in excel_files:
        df = read_table(file)

        if df.empty:
            log.warning("- Dilewati: %s (file kosong, 0 baris)", file.name)
            skipped_empty += 1
            continue

        columns = list(df.columns)
        if reference_columns is None:
            reference_columns = columns
        elif columns != reference_columns:
            missing = [col for col in reference_columns if col not in columns]
            extra = [col for col in columns if col not in reference_columns]
            log.warning(
                "- PERINGATAN: kolom %s berbeda (hilang: %s, tambahan: %s); "
                "tetap digabung, kolom yang tidak ada akan berisi NaN",
                file.name,
                missing or "-",
                extra or "-",
            )
            mismatched += 1

        nama_cabang = file.stem.replace("cabang_", "").title()
        df["Cabang"] = nama_cabang
        row_count = len(df)
        total_rows += row_count
        log.info("- Memproses: %s (%s baris)", file.name, row_count)
        all_data.append(df)

    if not all_data:
        raise ValueError("Semua file Excel kosong, tidak ada data untuk digabung")

    merged_df = pd.concat(all_data, ignore_index=True, sort=False)
    write_table(merged_df, output_file)

    log.info("")
    log.info(
        "Berhasil! Total %s baris dari %s file digabungkan ke %s",
        total_rows,
        len(all_data),
        output_file,
    )
    if skipped_empty:
        log.info("%s file kosong dilewati", skipped_empty)
    if mismatched:
        log.warning("%s file memiliki kolom berbeda (lihat peringatan di atas)", mismatched)

    return total_rows


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("input_folder", help="Folder berisi file Excel (.xlsx)")
    parser.add_argument(
        "output_file",
        nargs="?",
        help="File hasil gabungan (default: output/hasil_gabungan.xlsx)",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pfw merge",
        description="Menggabungkan semua file .xlsx dalam satu folder menjadi satu file.",
    )
    add_arguments(parser)
    return parser


def run(args: argparse.Namespace) -> int:
    setup_logging()
    try:
        merge_excel_files(args.input_folder, args.output_file or str(DEFAULT_OUTPUT))
    except (FileNotFoundError, ValueError) as exc:
        log.error("Error: %s", exc)
        return 1
    return 0


def main(argv: list[str] | None = None) -> int:
    return run(build_parser().parse_args(argv))
