"""
Generate sample data kotor untuk testing Data Cleaner
"""

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from common.excel import write_excel  # noqa: E402
from common.logging_setup import setup_logging  # noqa: E402


def build_dirty_data() -> pd.DataFrame:
    """Bangun DataFrame dengan berbagai masalah untuk di-clean."""
    data = {
        "Nama Lengkap": [
            "  BUDI SANTOSO  ",  # Extra spaces, uppercase
            "siti nurhaliza",  # Lowercase
            "AHMAD   DAHLAN",  # Multiple spaces
            "dewi  lestari",  # Multiple spaces, lowercase
            "Rudi Hartono",  # Sudah benar
            "BUDI SANTOSO",  # Duplikat
            "  maya   sari  ",  # Extra spaces
            "",  # Kosong
            "joko widodo",  # Lowercase
            "MEGAWATI SOEKARNO",  # Uppercase
        ],
        "Tanggal Lahir": [
            "15/01/1990",  # Format slash
            "1985-03-20",  # Format ISO
            "05 Mei 2000",  # Format Indonesia
            "12-12-1995",  # Format dash
            "2001/07/25",  # Format slash year first
            "15/01/1990",  # Duplikat
            "30 Juni 1988",  # Format Indonesia
            "01-01-2000",  # Placeholder
            "1975-08-17",  # Format ISO
            "25/12/1980",  # Format slash
        ],
        "No HP": [
            "081234567890",  # Tanpa format
            "+62 812 3456 7891",  # Format +62 dengan spasi
            "0813-4567-8901",  # Sudah ada dash
            "62 814 567 8902",  # Format 62 tanpa + (11 digit setelah normalisasi)
            "0815.4567.8903",  # Format dengan titik
            "081234567890",  # Duplikat
            "+62-816-4567-8904",  # Format +62 dengan dash
            "0817 4567 8905",  # Format dengan spasi
            "0818-4567-8906",  # Sudah ada dash
            "081945678907",  # Tanpa format
        ],
        "Email": [
            "budi@email.com",
            "siti@email.com",
            "ahmad@email.com",
            "dewi@email.com",
            "rudi@email.com",
            "budi@email.com",  # Duplikat
            "maya@email.com",
            "",  # Kosong
            "joko@email.com",
            "mega@email.com",
        ],
        "Kota": [
            "Jakarta",
            "Bandung",
            "Surabaya",
            "Yogyakarta",
            "Semarang",
            "Jakarta",  # Duplikat
            "Medan",
            "Makassar",
            "Denpasar",
            "Palembang",
        ],
    }

    return pd.DataFrame(data)


def generate_dirty_data(output_path=None) -> Path:
    """Generate sample data kotor sebagai teks (nomor HP tetap diawali 0)."""
    log = setup_logging()
    script_dir = Path(__file__).parent
    path = Path(output_path) if output_path else script_dir / "sample" / "data_kotor.xlsx"

    df = build_dirty_data()
    write_excel(df, path, text_columns=("No HP",))

    log.info("Sample data kotor berhasil dibuat: %s", path)
    log.info("Total baris: %s", len(df))
    return path


if __name__ == "__main__":
    generate_dirty_data()
