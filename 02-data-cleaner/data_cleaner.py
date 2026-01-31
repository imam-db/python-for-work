"""
Data Cleaner - Membersihkan data Excel berdasarkan konfigurasi YAML
"""

import sys
import re
import pandas as pd
import yaml
from pathlib import Path
from dateutil import parser as date_parser


# Mapping bulan Indonesia ke angka
BULAN_INDONESIA = {
    'januari': '01', 'februari': '02', 'maret': '03', 'april': '04',
    'mei': '05', 'juni': '06', 'juli': '07', 'agustus': '08',
    'september': '09', 'oktober': '10', 'november': '11', 'desember': '12',
    'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
    'jun': '06', 'jul': '07', 'agu': '08', 'ags': '08',
    'sep': '09', 'okt': '10', 'nov': '11', 'des': '12'
}


def convert_indonesian_date(date_str: str) -> str:
    """Konversi tanggal format Indonesia ke format standar."""
    date_str = str(date_str).strip().lower()
    
    for bulan_indo, bulan_num in BULAN_INDONESIA.items():
        if bulan_indo in date_str:
            # Ganti nama bulan dengan angka
            date_str = date_str.replace(bulan_indo, bulan_num)
            # Parse: "05 05 2000" -> "05-05-2000"
            parts = date_str.split()
            if len(parts) == 3:
                return f"{parts[0]}-{parts[1]}-{parts[2]}"
    return date_str


def load_config(config_path: str) -> dict:
    """Load konfigurasi dari file YAML."""
    with open(config_path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def clean_nama(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Standardisasi format nama."""
    kolom = config.get('kolom')
    if kolom not in df.columns:
        print(f"  Peringatan: Kolom '{kolom}' tidak ditemukan, skip cleaning nama")
        return df
    
    format_type = config.get('format', 'title')
    trim = config.get('trim', True)
    
    original = df[kolom].copy()
    
    if trim:
        df[kolom] = df[kolom].astype(str).str.strip()
        df[kolom] = df[kolom].str.replace(r'\s+', ' ', regex=True)
    
    if format_type == 'title':
        df[kolom] = df[kolom].str.title()
    elif format_type == 'upper':
        df[kolom] = df[kolom].str.upper()
    elif format_type == 'lower':
        df[kolom] = df[kolom].str.lower()
    
    changed = (original != df[kolom]).sum()
    stats['nama_fixed'] = changed
    print(f"  Nama: {changed} data di-standardisasi")
    
    return df


def clean_tanggal(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Standardisasi format tanggal."""
    kolom = config.get('kolom')
    if kolom not in df.columns:
        print(f"  Peringatan: Kolom '{kolom}' tidak ditemukan, skip cleaning tanggal")
        return df
    
    output_format = config.get('format', '%d-%m-%Y')
    fixed_count = 0
    
    def parse_date(val):
        nonlocal fixed_count
        if pd.isna(val) or str(val).strip() == '':
            return val
        try:
            # Coba konversi format Indonesia dulu
            val_converted = convert_indonesian_date(str(val))
            parsed = date_parser.parse(val_converted, dayfirst=True)
            fixed_count += 1
            return parsed.strftime(output_format)
        except:
            return val
    
    df[kolom] = df[kolom].apply(parse_date)
    stats['tanggal_fixed'] = fixed_count
    print(f"  Tanggal: {fixed_count} data di-format ke {output_format}")
    
    return df


def clean_telepon(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Standardisasi format nomor telepon."""
    kolom = config.get('kolom')
    if kolom not in df.columns:
        print(f"  Peringatan: Kolom '{kolom}' tidak ditemukan, skip cleaning telepon")
        return df
    
    output_format = config.get('format', '0xxx-xxxx-xxxx')
    fixed_count = 0
    
    def format_phone(val):
        nonlocal fixed_count
        if pd.isna(val) or str(val).strip() == '':
            return val
        
        # Hapus semua karakter non-digit
        digits = re.sub(r'\D', '', str(val))
        
        # Handle +62 atau 62 di awal
        if digits.startswith('62'):
            digits = '0' + digits[2:]
        
        # Format sesuai pattern
        if len(digits) >= 10:
            if output_format == '0xxx-xxxx-xxxx':
                formatted = f"{digits[:4]}-{digits[4:8]}-{digits[8:12]}"
            elif output_format == '+62xxx-xxxx-xxxx':
                formatted = f"+62{digits[1:4]}-{digits[4:8]}-{digits[8:12]}"
            else:
                formatted = digits
            fixed_count += 1
            return formatted
        return val
    
    df[kolom] = df[kolom].apply(format_phone)
    stats['telepon_fixed'] = fixed_count
    print(f"  Telepon: {fixed_count} data di-format")
    
    return df


def remove_duplicates(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Hapus baris duplikat berdasarkan kolom tertentu."""
    kolom = config.get('kolom', [])
    
    # Validasi kolom ada
    missing = [k for k in kolom if k not in df.columns]
    if missing:
        print(f"  Peringatan: Kolom {missing} tidak ditemukan, skip hapus duplikat")
        return df
    
    before = len(df)
    df = df.drop_duplicates(subset=kolom, keep='first')
    removed = before - len(df)
    
    stats['duplikat_removed'] = removed
    print(f"  Duplikat: {removed} baris dihapus")
    
    return df


def remove_empty(df: pd.DataFrame, config: dict, stats: dict) -> pd.DataFrame:
    """Hapus baris dengan kolom kosong."""
    kolom = config.get('kolom', [])
    
    # Validasi kolom ada
    missing = [k for k in kolom if k not in df.columns]
    if missing:
        print(f"  Peringatan: Kolom {missing} tidak ditemukan, skip hapus kosong")
        return df
    
    before = len(df)
    df = df.dropna(subset=kolom)
    # Juga hapus string kosong
    for k in kolom:
        df = df[df[k].astype(str).str.strip() != '']
    removed = before - len(df)
    
    stats['empty_removed'] = removed
    print(f"  Baris kosong: {removed} baris dihapus")
    
    return df


def clean_data(config_path: str) -> None:
    """Main function untuk membersihkan data."""
    
    # Load config
    config = load_config(config_path)
    script_dir = Path(__file__).parent
    
    input_file = script_dir / config['input']
    output_file = script_dir / config['output']
    cleaning = config.get('cleaning', {})
    
    # Validasi input
    if not input_file.exists():
        print(f"Error: File '{input_file}' tidak ditemukan")
        sys.exit(1)
    
    print(f"Membaca file: {input_file}")
    df = pd.read_excel(input_file)
    print(f"Total baris: {len(df)}")
    
    stats = {}
    print("\nProses cleaning:")
    
    # Jalankan cleaning sesuai config
    if 'nama' in cleaning:
        df = clean_nama(df, cleaning['nama'], stats)
    
    if 'tanggal' in cleaning:
        df = clean_tanggal(df, cleaning['tanggal'], stats)
    
    if 'telepon' in cleaning:
        df = clean_telepon(df, cleaning['telepon'], stats)
    
    if 'duplikat' in cleaning:
        df = remove_duplicates(df, cleaning['duplikat'], stats)
    
    if 'hapus_kosong' in cleaning:
        df = remove_empty(df, cleaning['hapus_kosong'], stats)
    
    # Simpan hasil
    output_file.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_file, index=False)
    
    print(f"\nBerhasil! Data bersih disimpan ke: {output_file}")
    print(f"Total baris setelah cleaning: {len(df)}")


if __name__ == "__main__":
    config_file = sys.argv[1] if len(sys.argv) > 1 else "config.yaml"
    clean_data(config_file)
