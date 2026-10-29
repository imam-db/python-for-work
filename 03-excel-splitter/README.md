# 03 · Excel Splitter

> Bab 3 kisah [Bu Sari](../README.md): Rabu, laporan nasional dari ERP harus dipecah per
> cabang untuk evaluasi tim.

Memecah satu file Excel besar menjadi beberapa file berdasarkan nilai kolom tertentu.

## Cerita

ERP PT Maju Jaya menghasilkan satu file laporan penjualan nasional setiap bulan. Semua
transaksi dari semua cabang tercampur di dalamnya, sementara manager tiap cabang hanya butuh
bagiannya sendiri.

```
laporan_nasional.xlsx
├── 100 baris transaksi (pada data contoh)
├── Kolom: No, Tanggal, Cabang, Nama Sales, Produk, Qty, Harga Satuan, Total, Status, Metode Bayar
└── Data 5 cabang tercampur
```

Script ini memecah file tersebut menjadi satu file per cabang, lengkap dengan nama file yang
aman dipakai di sistem operasi mana pun.

## Instalasi

```bash
pip install -r requirements.txt
```

## Cara Menggunakan

```bash
# pakai config.yaml di folder utility
python excel_splitter.py

# file dan kolom lewat CLI
python excel_splitter.py laporan.xlsx --kolom "Cabang"

# override sebagian setting config
python excel_splitter.py --kolom "Bulan" --prefix "report_"
```

Path dari CLI mengikuti current working directory. Path `input` dan `output_folder` di
dalam config di-resolve relatif terhadap folder utility.

## Opsi CLI

| Opsi | Shortcut | Deskripsi |
|------|----------|-----------|
| `input` | – | File Excel input (positional) |
| `--kolom` | `-k` | Nama kolom untuk split |
| `--output` | `-o` | Folder output |
| `--prefix` | `-p` | Prefix nama file output |
| `--suffix` | `-s` | Suffix nama file output |
| `--no-header` | – | Tidak sertakan header di file hasil |
| `--config` | `-c` | File config YAML (default `config.yaml`) |

## Demo (hasil run nyata)

```bash
python generate_sample.py
python excel_splitter.py
```

```
Membaca file: /path/ke/repo/03-excel-splitter/sample/laporan_nasional.xlsx
Total baris: 100

Memecah berdasarkan kolom: Cabang
Ditemukan 5 grup

  - Bandung.xlsx: 20 baris
  - Jakarta.xlsx: 23 baris
  - Makassar.xlsx: 22 baris
  - Medan.xlsx: 15 baris
  - Surabaya.xlsx: 20 baris

Berhasil! 5 file dibuat di folder '/path/ke/repo/03-excel-splitter/output'
Total baris: 100
```

Sample dibuat dengan `random.seed(42)`, jadi distribusi di atas selalu sama setiap kali
`generate_sample.py` dijalankan.

## Nilai Split Kosong

Baris dengan kolom split kosong (`NaN` atau string kosong) **tidak dibuang**. Semua baris
tersebut masuk ke satu bucket bernama `_kosong.xlsx`, dan jumlahnya dicatat:

```
  - _kosong.xlsx: 2 baris

  2 baris tanpa nilai 'Cabang' masuk ke _kosong.xlsx
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

## Catatan

- Karakter yang tidak valid untuk nama file diganti `_`.
- CLI selalu menang atas config.
- Jika kolom split tidak ditemukan, script menampilkan daftar kolom yang tersedia.

## Roadmap Singkat

- [ ] Split berdasarkan beberapa kolom
- [ ] Export ke CSV/JSON
- [ ] Filter data sebelum split
