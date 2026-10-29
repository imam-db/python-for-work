"""Split a tabular file into one file per value of a column."""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from pfw.common.config import load_config
from pfw.common.io import read_table, write_table
from pfw.common.logging_setup import get_logger, setup_logging

log = get_logger("split")

DEFAULT_CONFIG = Path("configs") / "split.yaml"
EMPTY_BUCKET_NAME = "_kosong"


def sanitize_filename(name: str) -> str:
    """Strip characters that are invalid in file names."""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = str(name).replace(char, "_")
    return name.strip()


def split_excel(
    input_file: str,
    split_by: str,
    output_folder: str = "output",
    prefix: str = "",
    suffix: str = "",
    include_header: bool = True,
) -> list[dict]:
    """Split a tabular file by the values of one column.

    Rows with a blank/NaN split value are written to ``_kosong.xlsx``
    (groupby uses ``dropna=False``).
    """
    input_path = Path(input_file)

    if not input_path.exists():
        raise FileNotFoundError(f"File '{input_file}' not found")

    log.info("Reading file: %s", input_file)
    df = read_table(input_path)
    log.info("Total rows: %s", len(df))

    if split_by not in df.columns:
        raise ValueError(
            f"Column '{split_by}' not found. "
            f"Available columns: {', '.join(map(str, df.columns))}"
        )

    output_path = Path(output_folder)
    output_path.mkdir(parents=True, exist_ok=True)

    groups = df.groupby(split_by, dropna=False)

    # Merge NaN and blank strings into one bucket so they do not overwrite each other.
    buckets: dict[str, list[pd.DataFrame]] = {}
    for group_name, group_df in groups:
        if pd.isna(group_name) or str(group_name).strip() == "":
            bucket_name = EMPTY_BUCKET_NAME
        else:
            bucket_name = sanitize_filename(str(group_name))
        buckets.setdefault(bucket_name, []).append(group_df)

    log.info("")
    log.info("Splitting by column: %s", split_by)
    log.info("Found %s group(s)", len(buckets))
    log.info("")

    results = []
    empty_rows = 0
    for bucket_name, frames in buckets.items():
        group_df = pd.concat(frames, ignore_index=True) if len(frames) > 1 else frames[0]
        filename = f"{prefix}{bucket_name}{suffix}.xlsx"
        filepath = output_path / filename
        write_table(group_df, filepath, header=include_header)

        is_empty = bucket_name == EMPTY_BUCKET_NAME
        if is_empty:
            empty_rows += len(group_df)

        results.append({"name": bucket_name, "file": filename, "rows": len(group_df)})
        log.info("  - %s: %s rows", filename, len(group_df))

    if empty_rows:
        log.info("")
        log.info(
            "  %s row(s) without a '%s' value went to %s.xlsx",
            empty_rows,
            split_by,
            EMPTY_BUCKET_NAME,
        )

    log.info("")
    log.info("Done! Created %s file(s) in folder '%s'", len(results), output_folder)
    log.info("Total rows: %s", sum(r["rows"] for r in results))

    return results


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("input", nargs="?", help="Input tabular file")
    parser.add_argument("--kolom", "-k", help="Column name to split by")
    parser.add_argument("--output", "-o", help="Output folder")
    parser.add_argument("--prefix", "-p", help="Prefix for output file names")
    parser.add_argument("--suffix", "-s", help="Suffix for output file names")
    parser.add_argument(
        "--no-header",
        action="store_true",
        help="Omit the header row in each output file",
    )
    parser.add_argument(
        "--config",
        "-c",
        help="YAML config file (default: configs/split.yaml)",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pfw split",
        description="Split a tabular file by the values of one column.",
    )
    add_arguments(parser)
    return parser


def run(args: argparse.Namespace) -> int:
    setup_logging()
    config_path = Path(args.config) if args.config else DEFAULT_CONFIG
    config = load_config(config_path)

    if args.input:
        input_file = Path(args.input)
    elif config.get("input"):
        input_file = Path(config["input"])
    else:
        input_file = None

    output_folder = (
        Path(args.output) if args.output else Path(config.get("output_folder", "output"))
    )
    split_by = args.kolom or config.get("split_by")
    prefix = args.prefix if args.prefix is not None else config.get("prefix", "")
    suffix = args.suffix if args.suffix is not None else config.get("suffix", "")
    include_header = config.get("include_header", True) and not args.no_header

    if not input_file:
        log.error("Error: input file is required (via CLI or config)")
        return 1

    if not split_by:
        log.error("Error: split column is required (via --kolom or config)")
        return 1

    try:
        split_excel(
            input_file=str(input_file),
            split_by=split_by,
            output_folder=str(output_folder),
            prefix=prefix,
            suffix=suffix,
            include_header=include_header,
        )
    except (FileNotFoundError, ValueError) as exc:
        log.error("Error: %s", exc)
        return 1
    return 0


def main(argv: list[str] | None = None) -> int:
    return run(build_parser().parse_args(argv))
