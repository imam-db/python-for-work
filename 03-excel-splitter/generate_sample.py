"""
Generate sample data laporan nasional untuk testing Excel Splitter
"""

import pandas as pd
import random
from pathlib import Path
from datetime import datetime, timedelta


def generate_sample_data():
    """Generate data laporan penjualan nasional."""
    
    cabang_list = ['Jakarta', 'Bandung', 'Surabaya', 'Medan', 'Makassar']
    sales_list = ['Budi Santoso', 'Dewi Lestari', 'Ahmad Dahlan', 'Siti Nurhaliza', 'Rudi Hartono']
    produk_list = [
        ('Laptop ASUS ROG', 15000000),
        ('Monitor LG 27 inch', 4200000),
        ('Keyboard Mechanical', 1500000),
        ('Mouse Gaming', 800000),
        ('Headset Wireless', 1200000),
        ('Webcam HD', 900000),
        ('SSD 1TB', 1800000),
        ('RAM 16GB', 1200000),
    ]
    status_list = ['Lunas', 'Cicilan', 'Pending']
    metode_list = ['Transfer', 'Cash', 'Kartu Kredit', 'Tempo 30 Hari']
    
    data = []
    start_date = datetime(2024, 1, 1)
    
    # Generate 100 transaksi
    for i in range(100):
        cabang = random.choice(cabang_list)
        sales = random.choice(sales_list)
        produk, harga = random.choice(produk_list)
        qty = random.randint(1, 5)
        tanggal = start_date + timedelta(days=random.randint(0, 30))
        
        data.append({
            'No': i + 1,
            'Tanggal': tanggal.strftime('%Y-%m-%d'),
            'Cabang': cabang,
            'Nama Sales': sales,
            'Produk': produk,
            'Qty': qty,
            'Harga Satuan': harga,
            'Total': qty * harga,
            'Status': random.choice(status_list),
            'Metode Bayar': random.choice(metode_list)
        })
    
    df = pd.DataFrame(data)
    
    # Simpan ke folder sample
    script_dir = Path(__file__).parent
    output_path = script_dir / 'sample' / 'laporan_nasional.xlsx'
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    df.to_excel(output_path, index=False)
    
    print(f"Sample data berhasil dibuat: {output_path}")
    print(f"Total baris: {len(df)}")
    print(f"\nDistribusi per cabang:")
    for cabang, count in df['Cabang'].value_counts().items():
        print(f"  - {cabang}: {count} transaksi")


if __name__ == "__main__":
    generate_sample_data()
