"""
Generate sample data untuk testing Excel Comparator
Membuat 2 file: data_lama.xlsx dan data_revisi.xlsx dengan beberapa perbedaan
"""

import pandas as pd
from pathlib import Path


def generate_sample_data():
    """Generate 2 file Excel dengan perbedaan untuk testing."""
    
    # Data lama
    data_lama = {
        'No': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
        'Nama': [
            'Budi Santoso', 'Dewi Lestari', 'Ahmad Dahlan', 'Siti Nurhaliza',
            'Rudi Hartono', 'Maya Sari', 'Joko Widodo', 'Mega Wati',
            'Andi Pratama', 'Rina Susanti'
        ],
        'Cabang': [
            'Jakarta', 'Bandung', 'Surabaya', 'Medan', 'Makassar',
            'Jakarta', 'Bandung', 'Surabaya', 'Medan', 'Makassar'
        ],
        'Total': [
            5000000, 3500000, 4200000, 2800000, 3100000,
            4500000, 3800000, 2900000, 3300000, 4100000
        ],
        'Status': [
            'Lunas', 'Pending', 'Lunas', 'Cicilan', 'Lunas',
            'Pending', 'Lunas', 'Pending', 'Lunas', 'Cicilan'
        ]
    }
    
    # Data revisi (dengan beberapa perubahan)
    data_revisi = {
        'No': [1, 2, 3, 4, 5, 6, 7, 8, 11, 12],  # No 9,10 dihapus, 11,12 baru
        'Nama': [
            'Budi Santoso', 'Dewi Lestari', 'Ahmad Dahlan', 'Siti Nurhaliza',
            'Rudi Hartono', 'Maya Sari', 'Joko Widodo', 'Mega Wati',
            'Bambang Susilo', 'Citra Dewi'  # 2 orang baru
        ],
        'Cabang': [
            'Jakarta', 'Bandung', 'Surabaya', 'Medan', 'Makassar',
            'Jakarta', 'Bandung', 'Surabaya', 'Semarang', 'Yogyakarta'  # 2 cabang baru
        ],
        'Total': [
            5500000,  # Berubah dari 5000000
            3500000,
            4200000,
            2800000,
            3600000,  # Berubah dari 3100000
            4500000,
            3800000,
            2900000,
            2700000,  # Baru
            3200000   # Baru
        ],
        'Status': [
            'Lunas',
            'Lunas',    # Berubah dari Pending
            'Lunas',
            'Lunas',    # Berubah dari Cicilan
            'Lunas',
            'Lunas',    # Berubah dari Pending
            'Lunas',
            'Lunas',    # Berubah dari Pending
            'Lunas',    # Baru
            'Pending'   # Baru
        ]
    }
    
    df_lama = pd.DataFrame(data_lama)
    df_revisi = pd.DataFrame(data_revisi)
    
    # Simpan ke folder sample
    script_dir = Path(__file__).parent
    sample_dir = script_dir / 'sample'
    sample_dir.mkdir(parents=True, exist_ok=True)
    
    path_lama = sample_dir / 'data_lama.xlsx'
    path_revisi = sample_dir / 'data_revisi.xlsx'
    
    df_lama.to_excel(path_lama, index=False)
    df_revisi.to_excel(path_revisi, index=False)
    
    print(f"Sample data berhasil dibuat:")
    print(f"  - {path_lama} ({len(df_lama)} baris)")
    print(f"  - {path_revisi} ({len(df_revisi)} baris)")
    
    print("\nPerbedaan yang dibuat:")
    print("  - No 1: Total berubah 5.000.000 → 5.500.000")
    print("  - No 2: Status berubah Pending → Lunas")
    print("  - No 4: Status berubah Cicilan → Lunas")
    print("  - No 5: Total berubah 3.100.000 → 3.600.000")
    print("  - No 6: Status berubah Pending → Lunas")
    print("  - No 8: Status berubah Pending → Lunas")
    print("  - No 9, 10: Dihapus")
    print("  - No 11, 12: Baris baru")


if __name__ == "__main__":
    generate_sample_data()
