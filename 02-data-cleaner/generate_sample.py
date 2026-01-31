"""
Generate sample data kotor untuk testing Data Cleaner
"""

import pandas as pd
from pathlib import Path


def generate_dirty_data():
    """Generate data dengan berbagai masalah untuk di-clean."""
    
    data = {
        'Nama Lengkap': [
            '  BUDI SANTOSO  ',      # Extra spaces, uppercase
            'siti nurhaliza',         # Lowercase
            'AHMAD   DAHLAN',         # Multiple spaces
            'dewi  lestari',          # Multiple spaces, lowercase
            'Rudi Hartono',           # Sudah benar
            'BUDI SANTOSO',           # Duplikat
            '  maya   sari  ',        # Extra spaces
            '',                       # Kosong
            'joko widodo',            # Lowercase
            'MEGAWATI SOEKARNO',      # Uppercase
        ],
        'Tanggal Lahir': [
            '15/01/1990',             # Format slash
            '1985-03-20',             # Format ISO
            '05 Mei 2000',            # Format Indonesia
            '12-12-1995',             # Format dash
            '2001/07/25',             # Format slash year first
            '15/01/1990',             # Duplikat
            '30 Juni 1988',           # Format Indonesia
            '01-01-2000',             # Placeholder
            '1975-08-17',             # Format ISO
            '25/12/1980',             # Format slash
        ],
        'No HP': [
            '081234567890',           # Tanpa format
            '+62 812 3456 7891',      # Format +62 dengan spasi
            '0813-4567-8901',         # Sudah ada dash
            '62 814 567 8902',        # Format 62 tanpa +
            '0815.4567.8903',         # Format dengan titik
            '081234567890',           # Duplikat
            '+62-816-4567-8904',      # Format +62 dengan dash
            '0817 4567 8905',         # Format dengan spasi
            '0818-4567-8906',         # Sudah ada dash
            '081945678907',           # Tanpa format
        ],
        'Email': [
            'budi@email.com',
            'siti@email.com',
            'ahmad@email.com',
            'dewi@email.com',
            'rudi@email.com',
            'budi@email.com',         # Duplikat
            'maya@email.com',
            '',                       # Kosong
            'joko@email.com',
            'mega@email.com',
        ],
        'Kota': [
            'Jakarta',
            'Bandung',
            'Surabaya',
            'Yogyakarta',
            'Semarang',
            'Jakarta',                # Duplikat
            'Medan',
            'Makassar',
            'Denpasar',
            'Palembang',
        ]
    }
    
    df = pd.DataFrame(data)
    
    # Simpan ke folder sample
    script_dir = Path(__file__).parent
    output_path = script_dir / 'sample' / 'data_kotor.xlsx'
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    df.to_excel(output_path, index=False)
    print(f"Sample data kotor berhasil dibuat: {output_path}")
    print(f"Total baris: {len(df)}")
    print("\nPreview data:")
    print(df.to_string())


if __name__ == "__main__":
    generate_dirty_data()
