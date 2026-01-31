"""
Script untuk generate sample data Excel yang realistis
"""

import pandas as pd
import random
from datetime import datetime, timedelta

random.seed(42)

produk_list = [
    ('Laptop ASUS ROG', 15000000, 18000000),
    ('Laptop Lenovo ThinkPad', 12000000, 15000000),
    ('Monitor LG 27 inch', 3500000, 4500000),
    ('Monitor Samsung 24 inch', 2500000, 3200000),
    ('Keyboard Mechanical Logitech', 800000, 1200000),
    ('Mouse Wireless Logitech', 350000, 500000),
    ('Printer Epson L3210', 2800000, 3500000),
    ('Printer HP LaserJet', 4500000, 5500000),
    ('UPS APC 1100VA', 1500000, 1800000),
    ('Router WiFi TP-Link', 450000, 650000),
    ('Webcam Logitech C920', 1200000, 1500000),
    ('Headset Jabra Evolve', 2000000, 2500000),
    ('SSD Samsung 1TB', 1200000, 1500000),
    ('RAM DDR4 16GB', 600000, 800000),
    ('Kabel HDMI 2m', 50000, 100000),
]

sales_names = [
    'Budi Santoso', 'Dewi Lestari', 'Ahmad Fauzi', 'Siti Rahayu', 
    'Rudi Hermawan', 'Nina Kartika', 'Eko Prasetyo', 'Maya Sari', 
    'Doni Wijaya', 'Rina Putri'
]

# Cabang dengan range jumlah transaksi per bulan
cabang_data = {
    'cabang_jakarta': (80, 120),
    'cabang_bandung': (50, 80),
    'cabang_surabaya': (60, 90),
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
        
        data.append({
            'No': i + 1,
            'Tanggal': date.strftime('%Y-%m-%d'),
            'Nama Sales': sales,
            'Produk': produk,
            'Qty': qty,
            'Harga Satuan': harga,
            'Total': total,
            'Status': random.choice(['Lunas', 'Lunas', 'Lunas', 'Cicilan', 'Pending']),
            'Metode Bayar': random.choice(['Transfer', 'Cash', 'Kartu Kredit', 'Tempo 30 Hari'])
        })
    return data


if __name__ == "__main__":
    start_date = datetime(2024, 1, 1)
    
    for cabang, (min_rows, max_rows) in cabang_data.items():
        num_rows = random.randint(min_rows, max_rows)
        data = generate_data(num_rows, start_date)
        df = pd.DataFrame(data)
        df.to_excel(f'01-merge-excel/sample/{cabang}.xlsx', index=False)
        print(f'{cabang}.xlsx: {num_rows} transaksi')
    
    print('Done!')
