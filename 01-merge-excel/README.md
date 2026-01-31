# Merge Excel

Menggabungkan beberapa file Excel dalam satu folder menjadi satu file.

## Cerita

Kamu bekerja sebagai admin di distributor peralatan IT yang punya 15 cabang di seluruh Indonesia. Setiap akhir bulan, masing-masing cabang mengirimkan laporan penjualan dalam format Excel ke email kamu.

```
inbox/
├── laporan_cabang_jakarta.xlsx      (120 transaksi)
├── laporan_cabang_bandung.xlsx      (72 transaksi)
├── laporan_cabang_surabaya.xlsx     (80 transaksi)
├── laporan_cabang_medan.xlsx        (65 transaksi)
├── laporan_cabang_makassar.xlsx     (45 transaksi)
├── laporan_cabang_semarang.xlsx     (58 transaksi)
├── laporan_cabang_yogyakarta.xlsx   (42 transaksi)
├── laporan_cabang_palembang.xlsx    (38 transaksi)
├── laporan_cabang_denpasar.xlsx     (55 transaksi)
├── laporan_cabang_balikpapan.xlsx   (35 transaksi)
├── laporan_cabang_manado.xlsx       (28 transaksi)
├── laporan_cabang_pontianak.xlsx    (32 transaksi)
├── laporan_cabang_banjarmasin.xlsx  (30 transaksi)
├── laporan_cabang_pekanbaru.xlsx    (40 transaksi)
└── laporan_cabang_lampung.xlsx      (36 transaksi)
```

Total: **776 transaksi** tersebar di 15 file.

Setiap file berisi data dengan format yang sama:

| No | Tanggal | Nama Sales | Produk | Qty | Harga Satuan | Total | Status | Metode Bayar |
|----|---------|------------|--------|-----|--------------|-------|--------|--------------|
| 1 | 2024-01-05 | Budi Santoso | Laptop ASUS ROG | 2 | 16.500.000 | 33.000.000 | Lunas | Transfer |
| 2 | 2024-01-05 | Dewi Lestari | Monitor LG 27 inch | 5 | 4.200.000 | 21.000.000 | Cicilan | Kartu Kredit |
| ... | ... | ... | ... | ... | ... | ... | ... | ... |

Manager minta kamu gabungkan semua data ini jadi satu file untuk:
- Analisis penjualan nasional
- Laporan ke direksi
- Input ke sistem ERP

**Cara manual:**
1. Buka file pertama
2. Buka file kedua, select all (skip header), copy
3. Paste di file pertama
4. Ulangi untuk 13 file lainnya
5. Save

Waktu: sekitar **10-15 menit**. Tidak lama, tapi:
- Kerjaan repetitif yang harus dilakukan **setiap bulan**
- Rawan human error (salah select, lupa skip header, double paste)
- Makin banyak cabang = makin lama dan makin rawan error
- Waktu kamu lebih berharga untuk kerjaan yang butuh mikir

**Dengan script ini:** 5 detik. Tinggal jalankan, selesai, lanjut kerjaan lain.

## Instalasi

```bash
pip install -r requirements.txt
```

## Cara Menggunakan

```bash
python merge_excel.py <folder_input> <file_output>
```

**Contoh:**

```bash
# Taruh semua file Excel di satu folder
python merge_excel.py ./laporan_januari laporan_gabungan_jan2024.xlsx
```

## Demo dengan Sample Data

Repository ini menyertakan 3 sample file (dari 15 cabang) untuk testing:

```bash
python merge_excel.py ./sample hasil_gabungan.xlsx
```

Output:
```
Ditemukan 3 file Excel di folder ./sample
- Memproses: cabang_bandung.xlsx (72 baris)
- Memproses: cabang_jakarta.xlsx (120 baris)
- Memproses: cabang_surabaya.xlsx (80 baris)

Berhasil! Total 272 baris digabungkan ke hasil_gabungan.xlsx
```

## Struktur Data Sample

Setiap file Excel berisi kolom:
- **No** - Nomor urut transaksi
- **Tanggal** - Tanggal transaksi (format: YYYY-MM-DD)
- **Nama Sales** - Nama sales yang handle transaksi
- **Produk** - Nama produk yang dijual
- **Qty** - Jumlah unit
- **Harga Satuan** - Harga per unit (dalam Rupiah)
- **Total** - Total harga (Qty x Harga Satuan)
- **Status** - Status pembayaran (Lunas/Cicilan/Pending)
- **Metode Bayar** - Metode pembayaran (Transfer/Cash/Kartu Kredit/Tempo 30 Hari)

## Catatan Penting

- Semua file Excel **harus memiliki struktur kolom yang sama**
- Script hanya membaca **sheet pertama** dari setiap file
- Kolom "No" akan mengikuti file asli (tidak di-reset) - bisa ditambahkan fitur renumber jika perlu
- File dengan format `.xls` (Excel lama) tidak didukung, convert dulu ke `.xlsx`

## Pengembangan Selanjutnya

Fitur yang bisa ditambahkan:
- [ ] Reset nomor urut setelah merge
- [ ] Filter file berdasarkan pattern (misal: `*_januari_*.xlsx`)
- [ ] Support multiple sheet
- [ ] Validasi struktur kolom sebelum merge
- [ ] Export ke format lain (CSV, Google Sheets)

## Blog

[Link ke artikel blog] *(coming soon)*
