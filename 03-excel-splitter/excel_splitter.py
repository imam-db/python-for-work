"""
Excel Splitter - Memecah file Excel berdasarkan nilai kolom tertentu
Mendukung konfigurasi via YAML dan CLI arguments (hybrid)
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from common.config import load_config, resolve_config_path, resolve_path  # noqa: E402
from common.logging_setup import get_logger, setup_logging  # noqa: E402

log = get_logger("excel_splitter")

EMPTY_BUCKET_NAME = "_kosong"


def sanitize_filename(name: str) -> str:
    """Bersihkan nama file dari karakter yang tidak valid."""
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
    """
    Memecah file Excel berdasarkan nilai kolom tertentu.

    Baris dengan nilai split kosong/NaN tetap disertakan ke file
    ``_kosong.xlsx`` (groupby memakai ``dropna=False``).
    """
    input_path = Path(input_file)

    if not input_path.exists():
        raise FileNotFoundError(f"File '{input_file}' tidak ditemukan")

    log.info("Membaca file: %s", input_file)
    df = pd.read_excel(input_path)
    log.info("Total baris: %s", len(df))

    if split_by not in df.columns:
        raise ValueError(
            f"Kolom '{split_by}' tidak ditemukan. "
            f"Kolom yang tersedia: {', '.join(map(str, df.columns))}"
        )

    output_path = Path(output_folder)
    output_path.mkdir(parents=True, exist_ok=True)

    groups = df.groupby(split_by, dropna=False)

    # Gabungkan NaN dan string kosong ke satu bucket supaya tidak saling timpa.
    buckets: dict[str, list[pd.DataFrame]] = {}
    for group_name, group_df in groups:
        if pd.isna(group_name) or str(group_name).strip() == "":
            bucket_name = EMPTY_BUCKET_NAME
        else:
            bucket_name = sanitize_filename(str(group_name))
        buckets.setdefault(bucket_name, []).append(group_df)

    log.info("")
    log.info("Memecah berdasarkan kolom: %s", split_by)
    log.info("Ditemukan %s grup", len(buckets))
    log.info("")

    results = []
    empty_rows = 0
    for bucket_name, frames in buckets.items():
        group_df = pd.concat(frames, ignore_index=True) if len(frames) > 1 else frames[0]
        filename = f"{prefix}{bucket_name}{suffix}.xlsx"
        filepath = output_path / filename
        group_df.to_excel(filepath, index=False, header=include_header)

        is_empty = bucket_name == EMPTY_BUCKET_NAME
        if is_empty:
            empty_rows += len(group_df)

        results.append({"name": bucket_name, "file": filename, "rows": len(group_df)})
        log.info("  - %s: %s baris", filename, len(group_df))

    if empty_rows:
        log.info("")
        log.info(
            "  %s baris tanpa nilai '%s' masuk ke %s.xlsx",
            empty_rows,
            split_by,
            EMPTY_BUCKET_NAME,
        )

    log.info("")
    log.info("Berhasil! %s file dibuat di folder '%s'", len(results), output_folder)
    log.info("Total baris: %s", sum(r["rows"] for r in results))

    return results


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Memecah file Excel berdasarkan nilai kolom tertentu."
    )
    parser.add_argument("input", nargs="?", help="File Excel input")
    parser.add_argument("--kolom", "-k", help="Nama kolom untuk split")
    parser.add_argument("--output", "-o", help="Folder output")
    parser.add_argument("--prefix", "-p", help="Prefix nama file output")
    parser.add_argument("--suffix", "-s", help="Suffix nama file output")
    parser.add_argument("--no-header", action="store_true", help="Tidak sertakan header")
    parser.add_argument(
        "--config",
        "-c",
        default="config.yaml",
        help="File konfigurasi YAML (default: config.yaml)",
    )
    return parser


def main(argv: list[str] | None = None) -> None:
    setup_logging()
    parser = build_parser()
    args = parser.parse_args(argv)

    script_dir = Path(__file__).parent
    config_path = resolve_config_path(args.config, script_dir)
    config = load_config(config_path)

    # CLI selalu menang; path dari CLI relatif terhadap cwd.
    if args.input:
        input_file = Path(args.input)
    else:
        input_file = resolve_path(script_dir, config["input"]) if config.get("input") else None

    if args.output:
        output_folder = Path(args.output)
    else:
        output_folder = resolve_path(script_dir, config.get("output_folder", "output"))

    split_by = args.kolom or config.get("split_by")
    prefix = args.prefix if args.prefix is not None else config.get("prefix", "")
    suffix = args.suffix if args.suffix is not None else config.get("suffix", "")
    include_header = config.get("include_header", True) and not args.no_header

    if not input_file:
        log.error("Error: File input harus diisi (via CLI atau config.yaml)")
        parser.print_help()
        sys.exit(1)

    if not split_by:
        log.error("Error: Kolom split harus diisi (via --kolom atau config.yaml)")
        parser.print_help()
        sys.exit(1)

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
        sys.exit(1)


if __name__ == "__main__":
    main()
