# Python for Work Utilities

**English summary**

Five small Python utilities for everyday Excel work: merging branch reports, cleaning
customer data, splitting a national report by branch, comparing two file versions, and
validating registrations against YAML rules. Every utility runs from its own folder, uses
pandas/openpyxl, ships synthetic sample data, and can be tried in one command. Shared
helpers (config loading, phone normalization, text-safe Excel I/O) live in `common/`.
Install with `pip install -r requirements.txt`; tests run with `pytest`. The narrative
below is in Indonesian, per the project's primary documentation language.

---

# Kisah Bu Sari di PT Maju Jaya

Bu Sari adalah satu-satunya admin di PT Maju Jaya, distributor peralatan IT dengan 15
cabang. Setiap awal bulan, inbox-nya penuh file Excel: laporan penjualan dari tiap cabang,
data pelanggan dari tim marketing, laporan nasional dari ERP, dan revisi file dari finance.
Semua pekerjaan itu sebenarnya sederhana, hanya saja repetitif dan rawan salah ketik.

Repository ini adalah lima bab dari minggu kerja Bu Sari. Setiap utility menyelesaikan satu
masalah nyata, dengan data contoh sintetis supaya bisa langsung dicoba tanpa menyentuh data
asli perusahaan.

1. **Senin** — 15 file laporan cabang harus jadi satu file untuk direksi.
2. **Selasa** — data pelanggan dari form online masih berantakan: nama huruf besar-kecil,
   tanggal campur format, nomor HP kehilangan angka nol di depan.
3. **Rabu** — satu laporan nasional harus dipecah per cabang untuk evaluasi tim.
4. **Kamis** — finance mengirim file revisi; Bu Sari harus tahu persis apa yang berubah.
5. **Jumat** — data pendaftaran harus divalidasi sebelum masuk sistem.

## Daftar Utilities

| Bab | Utility | Masalah | Hasil pada data contoh |
|-----|---------|---------|------------------------|
| 1 | [01-merge-excel](./01-merge-excel/) | Menggabungkan banyak file laporan | 3 file → 272 baris dalam satu file |
| 2 | [02-data-cleaner](./02-data-cleaner/) | Menstandardisasi nama, tanggal, nomor HP | 10 baris kotor → 8 baris bersih |
| 3 | [03-excel-splitter](./03-excel-splitter/) | Memecah laporan berdasarkan kolom | 100 baris → 5 file cabang |
| 4 | [04-excel-comparator](./04-excel-comparator/) | Membandingkan dua versi file | 2 baris baru, 2 dihapus, 6 sel berubah |
| 5 | [05-data-validator](./05-data-validator/) | Memeriksa kualitas data dengan rules | 12 baris: 5 valid, 7 bermasalah (17 temuan) |

## Before / After

Angka di bawah berasal dari menjalankan demo pada data contoh di repository ini, bukan
estimasi.

| Pekerjaan | Cara manual | Hasil script |
|-----------|-------------|--------------|
| Gabung laporan cabang | buka tiap file, copy-paste, ulangi | 272 baris dari 3 file jadi satu `.xlsx` |
| Bersihkan data pelanggan | find & replace, hapus duplikat satu per satu | 10 baris → 8 baris; 1 duplikat dan 1 baris kosong terhapus |
| Pecah laporan nasional | filter per cabang lalu save-as | 100 baris → 5 file (Jakarta 23, Makassar 22, Bandung 20, Surabaya 20, Medan 15) |
| Bandingkan revisi | scroll dua file berdampingan | 2 baris baru, 2 baris dihapus, 6 sel berubah |
| Validasi pendaftaran | cek kolom satu per satu | 17 temuan di 7 baris, lengkap dengan nomor barisnya |

Contoh nomor HP yang tadinya bermasalah:

```
62 814 567 8902  ->  0814-567-8902   (semua 11 digit tetap ada)
081234567890     ->  0812-3456-7890  (nol di depan tidak hilang)
```

## Quickstart

```bash
git clone <url-repository-ini>
cd python-for-work

python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

pip install -r requirements.txt

# Coba salah satu demo
cd 01-merge-excel
python generate_sample.py
python merge_excel.py ./sample
```

Setiap utility punya `generate_sample.py` sendiri, jadi tidak perlu menyiapkan file Excel
apa pun. Generator aman dijalankan dari folder utility maupun dari root repository:

```bash
python 03-excel-splitter/generate_sample.py
```

## Struktur Repository

```
common/              Helper bersama (config, Excel teks, telepon, logging)
01-merge-excel/      Gabung banyak file Excel
02-data-cleaner/     Bersihkan dan standardisasi data
03-excel-splitter/   Pecah file berdasarkan kolom
04-excel-comparator/ Bandingkan dua file
05-data-validator/   Validasi data dengan rules YAML
tests/               pytest untuk semua utility
.github/workflows/   CI: ruff + pytest
```

## Testing

```bash
pip install pytest ruff
ruff check .
pytest -q
```

CI menjalankan perintah yang sama pada setiap push dan pull request.

## Lisensi

MIT License. Lihat [LICENSE](./LICENSE).
