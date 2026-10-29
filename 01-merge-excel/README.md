# 01 · Merge Excel

> Bab 1 kisah [Bu Sari](../README.md): Senin pagi, 15 laporan cabang harus jadi satu file
> untuk direksi.

Menggabungkan beberapa file Excel dalam satu folder menjadi satu file.

## Cerita

PT Maju Jaya punya 15 cabang. Setiap awal bulan, masing-masing cabang mengirim laporan
penjualan dalam format Excel yang sama. Bu Sari harus menggabungkannya sebelum rapat
direksi.

Untuk contoh di repository ini, tersedia 3 file cabang (Jakarta, Bandung, Surabaya) dengan
total **272 transaksi**:

```
sample/
├── cabang_jakarta.xlsx    (120 transaksi)
├── cabang_bandung.xlsx    (72 transaksi)
└── cabang_surabaya.xlsx   (80 transaksi)
```

Setiap file berisi kolom yang sama:

| No | Tanggal | Nama Sales | Produk | Qty | Harga Satuan | Total | Status | Metode Bayar |
|----|---------|------------|--------|-----|--------------|-------|--------|--------------|
| 1 | 2024-01-05 | Budi Santoso | Laptop ASUS ROG | 2 | 16.500.000 | 33.000.000 | Lunas | Transfer |
| 2 | 2024-01-05 | Dewi Lestari | Monitor LG 27 inch | 5 | 4.200.000 | 21.000.000 | Cicilan | Kartu Kredit |
| ... | ... | ... | ... | ... | ... | ... | ... | ... |

## Instalasi

```bash
pip install -r requirements.txt
```

## Cara Menggunakan

```bash
python merge_excel.py <folder_input> [file_output]
```

- `folder_input` — folder berisi file `.xlsx` (wajib)
- `file_output` — file hasil gabungan (opsional). Default:
  `output/hasil_gabungan.xlsx` di folder utility ini.

Path dari CLI mengikuti current working directory, jadi bisa dijalankan dari mana saja:

```bash
# dari folder utility
python merge_excel.py ./sample

# dari root repository
python 01-merge-excel/merge_excel.py ./01-merge-excel/sample
```

## Demo (hasil run nyata)

```bash
python generate_sample.py
python merge_excel.py ./sample
```

```
Ditemukan 3 file Excel di folder ./sample
- Memproses: cabang_bandung.xlsx (72 baris)
- Memproses: cabang_jakarta.xlsx (120 baris)
- Memproses: cabang_surabaya.xlsx (80 baris)

Berhasil! Total 272 baris dari 3 file digabungkan ke /path/ke/repo/01-merge-excel/output/hasil_gabungan.xlsx
```

## Perilaku Penggabungan

- Kolom **Cabang** ditambahkan otomatis dari nama file (`cabang_jakarta.xlsx` → `Jakarta`).
- File dengan **0 baris** dilewati dan dicatat sebagai peringatan.
- Jika kolom antar file **berbeda**, file tetap digabung (outer join) dan muncul peringatan
  berisi kolom yang hilang/tambahan. Baris dari file tersebut akan berisi `NaN` pada kolom
  yang tidak dimilikinya, sehingga masalahnya terlihat jelas.
- File diproses dalam urutan nama (sort) supaya hasilnya konsisten.

Contoh peringatan:

```
- PERINGATAN: kolom cabang_bandung.xlsx berbeda (hilang: ['Status'], tambahan: -);
  tetap digabung, kolom yang tidak ada akan berisi NaN
```

## Catatan

- Hanya membaca sheet pertama.
- Format `.xls` (Excel lama) belum didukung; konversi dulu ke `.xlsx`.
- Nomor urut kolom `No` mengikuti file asli, belum di-reset setelah merge.

## Roadmap Singkat

- [ ] Reset nomor urut setelah merge
- [ ] Filter file berdasarkan pola nama
- [ ] Dukungan multiple sheet
