"""
Merge Excel - Menggabungkan beberapa file Excel menjadi satu file
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from common.logging_setup import get_logger, setup_logging  # noqa: E402

log = get_logger("merge_excel")


def merge_excel_files(input_folder: str, output_file: str) -> int:
    """Menggabungkan semua file Excel dalam folder menjadi satu file.

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

    all_data = []
    total_rows = 0
    reference_columns = None
    skipped_empty = 0
    mismatched = 0

    for file in excel_files:
        df = pd.read_excel(file)

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

        # Ambil nama cabang, hapus prefix "cabang_" jika ada
        nama_cabang = file.stem.replace("cabang_", "").title()
        df["Cabang"] = nama_cabang
        row_count = len(df)
        total_rows += row_count
        log.info("- Memproses: %s (%s baris)", file.name, row_count)
        all_data.append(df)

    if not all_data:
        raise ValueError("Semua file Excel kosong, tidak ada data untuk digabung")

    merged_df = pd.concat(all_data, ignore_index=True, sort=False)

    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    merged_df.to_excel(output_path, index=False)

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


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Menggabungkan semua file .xlsx dalam satu folder menjadi satu file."
    )
    parser.add_argument("input_folder", help="Folder berisi file Excel (.xlsx)")
    parser.add_argument(
        "output_file",
        nargs="?",
        help="File hasil gabungan (default: output/hasil_gabungan.xlsx di folder utility)",
    )
    return parser


def main() -> None:
    setup_logging()
    args = build_parser().parse_args()

    script_dir = Path(__file__).parent
    output_file = args.output_file or str(script_dir / "output" / "hasil_gabungan.xlsx")

    try:
        merge_excel_files(args.input_folder, output_file)
    except (FileNotFoundError, ValueError) as exc:
        log.error("Error: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
