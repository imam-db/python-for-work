"""
Generate sample data untuk testing Excel Comparator
Membuat 2 file: data_lama.xlsx dan data_revisi.xlsx dengan beberapa perbedaan
"""

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from common.logging_setup import setup_logging  # noqa: E402


def build_sample_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Generate 2 DataFrame dengan perbedaan untuk testing."""
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
        "No": [1, 2, 3, 4, 5, 6, 7, 8, 11, 12],  # No 9,10 dihapus, 11,12 baru
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
            5500000,  # Berubah dari 5000000
            3500000,
            4200000,
            2800000,
            3600000,  # Berubah dari 3100000
            4500000,
            3800000,
            2900000,
            2700000,  # Baru
            3200000,  # Baru
        ],
        "Status": [
            "Lunas",
            "Lunas",  # Berubah dari Pending
            "Lunas",
            "Lunas",  # Berubah dari Cicilan
            "Lunas",
            "Lunas",  # Berubah dari Pending
            "Lunas",
            "Lunas",  # Berubah dari Pending
            "Lunas",  # Baru
            "Pending",  # Baru
        ],
    }

    return pd.DataFrame(data_lama), pd.DataFrame(data_revisi)


def generate_sample_data(output_dir=None) -> tuple[Path, Path]:
    """Generate 2 file Excel dengan perbedaan dan kembalikan path-nya."""
    log = setup_logging()
    script_dir = Path(__file__).parent
    sample_dir = Path(output_dir) if output_dir else script_dir / "sample"
    sample_dir.mkdir(parents=True, exist_ok=True)

    df_lama, df_revisi = build_sample_data()

    path_lama = sample_dir / "data_lama.xlsx"
    path_revisi = sample_dir / "data_revisi.xlsx"

    df_lama.to_excel(path_lama, index=False)
    df_revisi.to_excel(path_revisi, index=False)

    log.info("Sample data berhasil dibuat:")
    log.info("  - %s (%s baris)", path_lama, len(df_lama))
    log.info("  - %s (%s baris)", path_revisi, len(df_revisi))

    log.info("")
    log.info("Perbedaan yang dibuat:")
    log.info("  - No 1: Total berubah 5.000.000 -> 5.500.000")
    log.info("  - No 2: Status berubah Pending -> Lunas")
    log.info("  - No 4: Status berubah Cicilan -> Lunas")
    log.info("  - No 5: Total berubah 3.100.000 -> 3.600.000")
    log.info("  - No 6: Status berubah Pending -> Lunas")
    log.info("  - No 8: Status berubah Pending -> Lunas")
    log.info("  - No 9, 10: Dihapus")
    log.info("  - No 11, 12: Baris baru")

    return path_lama, path_revisi


if __name__ == "__main__":
    generate_sample_data()
