"""
Merge Excel - Menggabungkan beberapa file Excel menjadi satu file
"""

import sys
import pandas as pd
from pathlib import Path


def merge_excel_files(input_folder: str, output_file: str) -> None:
    """Menggabungkan semua file Excel dalam folder menjadi satu file."""
    
    folder = Path(input_folder)
    
    if not folder.exists():
        print(f"Error: Folder '{input_folder}' tidak ditemukan")
        sys.exit(1)
    
    excel_files = list(folder.glob("*.xlsx"))
    
    if not excel_files:
        print(f"Error: Tidak ada file Excel di folder '{input_folder}'")
        sys.exit(1)
    
    print(f"Ditemukan {len(excel_files)} file Excel di folder {input_folder}")
    
    all_data = []
    total_rows = 0
    
    for file in excel_files:
        df = pd.read_excel(file)
        # Ambil nama cabang, hapus prefix "cabang_" jika ada
        nama_cabang = file.stem.replace("cabang_", "").title()
        df['Cabang'] = nama_cabang
        row_count = len(df)
        total_rows += row_count
        print(f"- Memproses: {file.name} ({row_count} baris)")
        all_data.append(df)
    
    merged_df = pd.concat(all_data, ignore_index=True)
    
    # Buat folder output jika belum ada
    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    merged_df.to_excel(output_file, index=False)
    
    print(f"\nBerhasil! Total {total_rows} baris digabungkan ke {output_file}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Penggunaan: python merge_excel.py <folder_input> [file_output]")
        print("Contoh: python merge_excel.py ./sample")
        print("        python merge_excel.py ./sample ./output/custom_output.xlsx")
        sys.exit(1)
    
    input_folder = sys.argv[1]
    
    # Default output di folder yang sama dengan script
    script_dir = Path(__file__).parent
    default_output = script_dir / "output" / "hasil_gabungan.xlsx"
    output_file = sys.argv[2] if len(sys.argv) > 2 else str(default_output)
    
    merge_excel_files(input_folder, output_file)
