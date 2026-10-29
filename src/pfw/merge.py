"""Merge every tabular file in a folder into a single output file."""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from pfw.common.io import SUPPORTED_SUFFIXES, read_table, write_table
from pfw.common.logging_setup import get_logger, setup_logging

log = get_logger("merge")

DEFAULT_OUTPUT = Path("output") / "hasil_gabungan.xlsx"


def _discover_files(folder: Path) -> list[Path]:
    files: list[Path] = []
    for path in folder.iterdir():
        if path.is_file() and path.suffix.lower() in SUPPORTED_SUFFIXES:
            files.append(path)
    return sorted(files, key=lambda item: item.name.lower())


def merge_tables(input_folder: str, output_file: str) -> int:
    """Merge every xlsx, csv, or parquet file in a folder into one file.

    Empty files are skipped. Files with different columns still merge (outer
    concat) but emit a warning so the resulting NaNs are obvious.
    """
    folder = Path(input_folder)

    if not folder.exists():
        raise FileNotFoundError(f"Folder '{input_folder}' not found")

    table_files = _discover_files(folder)
    if not table_files:
        supported = ", ".join(SUPPORTED_SUFFIXES)
        raise FileNotFoundError(
            f"No tabular files ({supported}) found in folder '{input_folder}'"
        )

    log.info("Found %s tabular file(s) in folder %s", len(table_files), input_folder)

    all_data: list[pd.DataFrame] = []
    total_rows = 0
    reference_columns: list[str] | None = None
    skipped_empty = 0
    mismatched = 0

    for file in table_files:
        df = read_table(file)

        if df.empty:
            log.warning("- Skipped: %s (empty file, 0 rows)", file.name)
            skipped_empty += 1
            continue

        columns = list(df.columns)
        if reference_columns is None:
            reference_columns = columns
        elif columns != reference_columns:
            missing = [col for col in reference_columns if col not in columns]
            extra = [col for col in columns if col not in reference_columns]
            log.warning(
                "- WARNING: columns in %s differ (missing: %s, extra: %s); "
                "still merging, absent columns will be NaN",
                file.name,
                missing or "-",
                extra or "-",
            )
            mismatched += 1

        branch_name = file.stem.replace("cabang_", "").title()
        df["Cabang"] = branch_name
        row_count = len(df)
        total_rows += row_count
        log.info("- Processing: %s (%s rows)", file.name, row_count)
        all_data.append(df)

    if not all_data:
        raise ValueError("All tabular files are empty; nothing to merge")

    merged_df = pd.concat(all_data, ignore_index=True, sort=False)
    write_table(merged_df, output_file)

    log.info("")
    log.info(
        "Done! Merged %s rows from %s file(s) into %s",
        total_rows,
        len(all_data),
        output_file,
    )
    if skipped_empty:
        log.info("%s empty file(s) skipped", skipped_empty)
    if mismatched:
        log.warning("%s file(s) had mismatched columns (see warnings above)", mismatched)

    return total_rows


def merge_excel_files(input_folder: str, output_file: str) -> int:
    """Backward-compatible alias for :func:`merge_tables`."""
    return merge_tables(input_folder, output_file)


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "input_folder",
        help="Folder of tabular files (.xlsx, .csv, .parquet)",
    )
    parser.add_argument(
        "output_file",
        nargs="?",
        help="Merged output path (default: output/hasil_gabungan.xlsx)",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pfw merge",
        description=(
            "Merge every .xlsx, .csv, or .parquet file in a folder into one file."
        ),
    )
    add_arguments(parser)
    return parser


def run(args: argparse.Namespace) -> int:
    setup_logging()
    try:
        merge_tables(args.input_folder, args.output_file or str(DEFAULT_OUTPUT))
    except (FileNotFoundError, ValueError) as exc:
        log.error("Error: %s", exc)
        return 1
    return 0


def main(argv: list[str] | None = None) -> int:
    return run(build_parser().parse_args(argv))
