"""
Script untuk generate sample data Excel yang realistis
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

random.seed(42)

produk_list = [
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

sales_names = [
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

# Cabang dengan range jumlah transaksi per bulan
cabang_data = {
    "cabang_jakarta": (80, 120),
    "cabang_bandung": (50, 80),
    "cabang_surabaya": (60, 90),
}


def generate_data(num_rows, start_date):
    data = []
    for i in range(num_rows):
        date = start_date + timedelta(days=random.randint(0, 30))
        produk, min_price, max_price = random.choice(produk_list)
        qty = random.randint(1, 20)
        harga = random.randint(min_price, max_price)
        total = qty * harga
        sales = random.choice(sales_names)

        data.append(
            {
                "No": i + 1,
                "Tanggal": date.strftime("%Y-%m-%d"),
                "Nama Sales": sales,
                "Produk": produk,
                "Qty": qty,
                "Harga Satuan": harga,
                "Total": total,
                "Status": random.choice(["Lunas", "Lunas", "Lunas", "Cicilan", "Pending"]),
                "Metode Bayar": random.choice(
                    ["Transfer", "Cash", "Kartu Kredit", "Tempo 30 Hari"]
                ),
            }
        )
    return data


def generate_sample(output_dir=None):
    """Generate file sample per cabang dan kembalikan daftar path."""
    log = setup_logging()
    random.seed(42)
    script_dir = Path(__file__).parent
    sample_dir = Path(output_dir) if output_dir else script_dir / "sample"
    sample_dir.mkdir(parents=True, exist_ok=True)

    start_date = datetime(2024, 1, 1)
    written = []

    for cabang, (min_rows, max_rows) in cabang_data.items():
        num_rows = random.randint(min_rows, max_rows)
        data = generate_data(num_rows, start_date)
        df = pd.DataFrame(data)
        output_path = sample_dir / f"{cabang}.xlsx"
        df.to_excel(output_path, index=False)
        written.append(output_path)
        log.info("%s.xlsx: %s transaksi", cabang, num_rows)

    log.info("Done!")
    return written


if __name__ == "__main__":
    generate_sample()
