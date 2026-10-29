"""
Generate sample data untuk testing Data Validator
Data sengaja dibuat dengan berbagai error untuk testing
"""

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from common.excel import write_excel  # noqa: E402
from common.logging_setup import setup_logging  # noqa: E402


def build_sample_data() -> pd.DataFrame:
    """Bangun data dengan berbagai error untuk testing validasi."""
    data = {
        "Nama Lengkap": [
            "Budi Santoso",
            "Dewi Lestari",
            "",  # Error: required
            "Siti Nurhaliza",
            "Rudi Hartono",
            "Maya Sari",
            "Joko Widodo",
            None,  # Error: required
            "Andi Pratama",
            "Rina Susanti",
            "Bambang Susilo",
            "Citra Dewi",
        ],
        "Email": [
            "budi@email.com",
            "dewi@email",  # Error: format email
            "ahmad@email.com",
            "siti@email.com",
            "",  # Error: required
            "maya@email.com",
            "joko@@email.com",  # Error: format email
            "mega@email.com",
            "andi@email.com",
            "rina@email.com",
            "bambang@email.com",
            "citra@email.com",
        ],
        "No HP": [
            "081234567890",
            "08123456",  # Error: kurang digit
            "081234567892",
            "081234567893",
            "081234567894",
            "0812",  # Error: kurang digit
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
            "1940-01-01",  # Error: sebelum 1950
            "1995-12-25",
            "2015-06-10",  # Error: setelah 2010
            "1988-08-08",
            "1975-11-11",
            "1992-02-29",
            "2020-01-01",  # Error: setelah 2010
            "1980-07-04",
            "1998-09-17",
            "2005-04-22",
        ],
        "Umur": [
            30,
            35,
            25,
            28,
            15,  # Error: kurang dari 17
            32,
            45,
            70,  # Error: lebih dari 65
            22,
            40,
            26,
            19,
        ],
        "NIK": [
            "3201234567890001",
            "3201234567890002",
            "320123456789",  # Error: kurang dari 16 digit
            "3201234567890004",
            "3201234567890005",
            "32012345678900067",  # Error: lebih dari 16 digit
            "3201234567890007",
            "3201234567890008",
            "320123ABC7890009",  # Error: ada huruf
            "3201234567890010",
            "3201234567890011",
            "3201234567890012",
        ],
        "Status": [
            "Aktif",
            "Nonaktif",
            "Pending",
            "Aktif",
            "Tidak Valid",  # Error: tidak dalam list
            "Aktif",
            "Nonaktif",
            "Batal",  # Error: tidak dalam list
            "Pending",
            "Aktif",
            "Nonaktif",
            "Pending",
        ],
    }

    return pd.DataFrame(data)


def generate_sample_data(output_path=None) -> Path:
    """Generate sample data (nomor HP ditulis sebagai teks)."""
    log = setup_logging()
    script_dir = Path(__file__).parent
    path = Path(output_path) if output_path else script_dir / "sample" / "data_pendaftaran.xlsx"

    df = build_sample_data()
    write_excel(df, path, text_columns=("No HP", "NIK"))

    log.info("Sample data berhasil dibuat: %s", path)
    log.info("Total baris: %s", len(df))

    log.info("")
    log.info("Error yang sengaja dibuat:")
    log.info("  - Baris 3, 8: Nama Lengkap kosong")
    log.info("  - Baris 2, 7: Format email tidak valid")
    log.info("  - Baris 5: Email kosong")
    log.info("  - Baris 2, 6: No HP kurang digit")
    log.info("  - Baris 3: Tanggal lahir sebelum 1950")
    log.info("  - Baris 5, 9: Tanggal lahir setelah 2010")
    log.info("  - Baris 5: Umur kurang dari 17")
    log.info("  - Baris 8: Umur lebih dari 65")
    log.info("  - Baris 3, 6, 9: NIK tidak 16 digit")
    log.info("  - Baris 5, 8: Status tidak valid")

    return path


if __name__ == "__main__":
    generate_sample_data()
