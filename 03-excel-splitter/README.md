# Excel Splitter

Memecah satu file Excel besar menjadi beberapa file berdasarkan nilai kolom tertentu.

## Cerita

Kamu bekerja sebagai admin di perusahaan distribusi dengan 15 cabang. Setiap bulan, sistem ERP generate satu file laporan penjualan nasional yang berisi semua transaksi dari semua cabang.

```
laporan_nasional_jan2024.xlsx
├── 1000+ baris transaksi
├── Kolom: No, Tanggal, Cabang, Sales, Produk, Qty, Total, Status
└── Data dari 15 cabang tercampur
```

Manager tiap cabang minta laporan masing-masing untuk evaluasi tim. Kamu harus:
1. Filter data per cabang
2. Copy ke file baru
3. Save dengan nama cabang
4. Ulangi 15 kali

**Cara manual:** 15-20 menit, rawan salah filter atau lupa cabang.

**Dengan script ini:** 5 detik. Otomatis pecah jadi 15 file sesuai nama cabang.

## Instalasi

```bash
pip install -r requirements.txt
```

## Cara Menggunakan

### Via Config (Default)

Edit `config.yaml`:
```yaml
input: data/laporan.xlsx
split_by: "Cabang"
prefix: "laporan_"
```

Jalankan:
```bash
python excel_splitter.py
```

### Via CLI Arguments

```bash
python excel_splitter.py laporan.xlsx --kolom "Cabang"
```

### Hybrid (Config + Override)

Setting default di config, override yang perlu via CLI:
```bash
python excel_splitter.py --kolom "Bulan" --prefix "report_"
```

## CLI Options

| Option | Shortcut | Deskripsi |
|--------|----------|-----------|
| `input` | - | File Excel input (positional) |
| `--kolom` | `-k` | Nama kolom untuk split |
| `--output` | `-o` | Folder output |
| `--prefix` | `-p` | Prefix nama file output |
| `--suffix` | `-s` | Suffix nama file output |
| `--no-header` | - | Tidak sertakan header |
| `--config` | `-c` | File config custom |

## Demo dengan Sample Data

```bash
python generate_sample.py
python excel_splitter.py
```

Output:
```
Membaca file: sample/laporan_nasional.xlsx
Total baris: 100

Memecah berdasarkan kolom: Cabang
Ditemukan 5 grup

  - Jakarta.xlsx: 23 baris
  - Bandung.xlsx: 18 baris
  - Surabaya.xlsx: 21 baris
  - Medan.xlsx: 19 baris
  - Makassar.xlsx: 19 baris

Berhasil! 5 file dibuat di folder 'output'
```

## Contoh Penggunaan

**Split berdasarkan cabang:**
```bash
python excel_splitter.py laporan.xlsx -k "Cabang"
```

**Split berdasarkan bulan dengan prefix:**
```bash
python excel_splitter.py laporan.xlsx -k "Bulan" -p "sales_" -s "_2024"
# Output: sales_Januari_2024.xlsx, sales_Februari_2024.xlsx, ...
```

**Split ke folder custom:**
```bash
python excel_splitter.py laporan.xlsx -k "Region" -o "./reports/per_region"
```

## Konfigurasi YAML

```yaml
input: sample/laporan_nasional.xlsx
output_folder: output/
split_by: "Cabang"
prefix: ""
suffix: ""
include_header: true
```

## Catatan Penting

- Nama file output diambil dari nilai kolom (karakter invalid otomatis di-replace dengan `_`)
- CLI arguments selalu override config YAML
- Jika kolom tidak ditemukan, script akan tampilkan daftar kolom yang tersedia

## Pengembangan Selanjutnya

- [ ] Split berdasarkan multiple kolom
- [ ] Export ke format lain (CSV, JSON)
- [ ] Filter data sebelum split
- [ ] Template nama file custom

## Blog

[Link ke artikel blog] *(coming soon)*
