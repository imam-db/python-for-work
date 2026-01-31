"""
Excel Splitter - Memecah file Excel berdasarkan nilai kolom tertentu
Mendukung konfigurasi via YAML dan CLI arguments (hybrid)
"""

import sys
import argparse
import pandas as pd
import yaml
from pathlib import Path


def load_config(config_path: str) -> dict:
    """Load konfigurasi dari file YAML."""
    config_file = Path(config_path)
    if config_file.exists():
        with open(config_file, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}
    return {}


def sanitize_filename(name: str) -> str:
    """Bersihkan nama file dari karakter yang tidak valid."""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = str(name).replace(char, '_')
    return name.strip()


def split_excel(
    input_file: str,
    split_by: str,
    output_folder: str = "output",
    prefix: str = "",
    suffix: str = "",
    include_header: bool = True
) -> None:
    """
    Memecah file Excel berdasarkan nilai kolom tertentu.
    
    Args:
        input_file: Path ke file Excel input
        split_by: Nama kolom untuk split
        output_folder: Folder output
        prefix: Prefix untuk nama file output
        suffix: Suffix untuk nama file output
        include_header: Sertakan header di setiap file
    """
    input_path = Path(input_file)
    
    if not input_path.exists():
        print(f"Error: File '{input_file}' tidak ditemukan")
        sys.exit(1)
    
    print(f"Membaca file: {input_file}")
    df = pd.read_excel(input_file)
    print(f"Total baris: {len(df)}")
    
    if split_by not in df.columns:
        print(f"Error: Kolom '{split_by}' tidak ditemukan")
        print(f"Kolom yang tersedia: {', '.join(df.columns)}")
        sys.exit(1)
    
    # Buat output folder
    output_path = Path(output_folder)
    output_path.mkdir(parents=True, exist_ok=True)
    
    # Group by kolom yang dipilih
    groups = df.groupby(split_by)
    
    print(f"\nMemecah berdasarkan kolom: {split_by}")
    print(f"Ditemukan {len(groups)} grup\n")
    
    results = []
    for group_name, group_df in groups:
        # Buat nama file
        safe_name = sanitize_filename(str(group_name))
        filename = f"{prefix}{safe_name}{suffix}.xlsx"
        filepath = output_path / filename
        
        # Simpan ke file
        group_df.to_excel(filepath, index=False, header=include_header)
        
        results.append({
            'name': group_name,
            'file': filename,
            'rows': len(group_df)
        })
        print(f"  - {filename}: {len(group_df)} baris")
    
    print(f"\nBerhasil! {len(results)} file dibuat di folder '{output_folder}'")
    
    # Summary
    total_rows = sum(r['rows'] for r in results)
    print(f"Total baris: {total_rows}")


def main():
    parser = argparse.ArgumentParser(
        description='Memecah file Excel berdasarkan nilai kolom tertentu'
    )
    parser.add_argument('input', nargs='?', help='File Excel input')
    parser.add_argument('--kolom', '-k', help='Nama kolom untuk split')
    parser.add_argument('--output', '-o', help='Folder output')
    parser.add_argument('--prefix', '-p', help='Prefix nama file output')
    parser.add_argument('--suffix', '-s', help='Suffix nama file output')
    parser.add_argument('--no-header', action='store_true', help='Tidak sertakan header')
    parser.add_argument('--config', '-c', default='config.yaml', help='File konfigurasi YAML')
    
    args = parser.parse_args()
    
    # Load config dari YAML
    script_dir = Path(__file__).parent
    config = load_config(script_dir / args.config)
    
    # Override dengan CLI arguments (CLI lebih prioritas)
    input_file = args.input or config.get('input')
    split_by = args.kolom or config.get('split_by')
    output_folder = args.output or config.get('output_folder', 'output')
    prefix = args.prefix if args.prefix is not None else config.get('prefix', '')
    suffix = args.suffix if args.suffix is not None else config.get('suffix', '')
    include_header = not args.no_header if args.no_header else config.get('include_header', True)
    
    # Resolve path relatif terhadap script directory
    if input_file and not Path(input_file).is_absolute():
        input_file = str(script_dir / input_file)
    if not Path(output_folder).is_absolute():
        output_folder = str(script_dir / output_folder)
    
    # Validasi
    if not input_file:
        print("Error: File input harus diisi (via CLI atau config.yaml)")
        parser.print_help()
        sys.exit(1)
    
    if not split_by:
        print("Error: Kolom split harus diisi (via --kolom atau config.yaml)")
        parser.print_help()
        sys.exit(1)
    
    split_excel(
        input_file=input_file,
        split_by=split_by,
        output_folder=output_folder,
        prefix=prefix,
        suffix=suffix,
        include_header=include_header
    )


if __name__ == "__main__":
    main()
