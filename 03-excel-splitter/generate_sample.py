"""
Generate sample data laporan nasional untuk testing Excel Splitter
"""

import random
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from common.logging_setup import setup_logging  # noqa: E402


def build_sample_data() -> pd.DataFrame:
    """Generate data laporan penjualan nasional (deterministik, seed 42)."""
    random.seed(42)

    cabang_list = ["Jakarta", "Bandung", "Surabaya", "Medan", "Makassar"]
    sales_list = ["Budi Santoso", "Dewi Lestari", "Ahmad Dahlan", "Siti Nurhaliza", "Rudi Hartono"]
    produk_list = [
        ("Laptop ASUS ROG", 15000000),
        ("Monitor LG 27 inch", 4200000),
        ("Keyboard Mechanical", 1500000),
        ("Mouse Gaming", 800000),
        ("Headset Wireless", 1200000),
        ("Webcam HD", 900000),
        ("SSD 1TB", 1800000),
        ("RAM 16GB", 1200000),
    ]
    status_list = ["Lunas", "Cicilan", "Pending"]
    metode_list = ["Transfer", "Cash", "Kartu Kredit", "Tempo 30 Hari"]

    data = []
    start_date = datetime(2024, 1, 1)

    for i in range(100):
        cabang = random.choice(cabang_list)
        sales = random.choice(sales_list)
        produk, harga = random.choice(produk_list)
        qty = random.randint(1, 5)
        tanggal = start_date + timedelta(days=random.randint(0, 30))

        data.append(
            {
                "No": i + 1,
                "Tanggal": tanggal.strftime("%Y-%m-%d"),
                "Cabang": cabang,
                "Nama Sales": sales,
                "Produk": produk,
                "Qty": qty,
                "Harga Satuan": harga,
                "Total": qty * harga,
                "Status": random.choice(status_list),
                "Metode Bayar": random.choice(metode_list),
            }
        )

    return pd.DataFrame(data)


def generate_sample_data(output_path=None) -> Path:
    """Generate sample data dan kembalikan path file."""
    log = setup_logging()
    script_dir = Path(__file__).parent
    path = Path(output_path) if output_path else script_dir / "sample" / "laporan_nasional.xlsx"

    df = build_sample_data()
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(path, index=False)

    log.info("Sample data berhasil dibuat: %s", path)
    log.info("Total baris: %s", len(df))
    log.info("")
    log.info("Distribusi per cabang:")
    for cabang, count in df["Cabang"].value_counts().items():
        log.info("  - %s: %s transaksi", cabang, count)

    return path


if __name__ == "__main__":
    generate_sample_data()
