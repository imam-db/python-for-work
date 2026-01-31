# Excel Comparator

Membandingkan 2 file Excel dan menampilkan perbedaannya secara detail.

## Cerita

Kamu terima email dari tim finance: "Tolong cek file revisi ini, ada beberapa koreksi data."

File-nya 500 baris. Mereka tidak jelaskan baris mana yang berubah. Cara manual:
1. Buka kedua file side by side
2. Scroll dan bandingkan satu-satu
3. Catat perbedaan di notepad

Waktu: 30-60 menit, mata perih, rawan kelewatan.

**Dengan script ini:** 5 detik, langsung dapat laporan lengkap:
- Cell mana yang berubah (dari apa ke apa)
- Baris mana yang baru ditambahkan
- Baris mana yang dihapus

## Instalasi

```bash
pip install -r requirements.txt
```

## Cara Menggunakan

### Via Config (Default)

Edit `config.yaml`:
```yaml
file_old: data/file_lama.xlsx
file_new: data/file_revisi.xlsx
key_column: "No"
output: output/laporan.xlsx
```

Jalankan:
```bash
python excel_comparator.py
```

### Via CLI Arguments

```bash
python excel_comparator.py file_lama.xlsx file_baru.xlsx
```

### Dengan Key Column

Bandingkan berdasarkan kolom ID (bukan posisi baris):
```bash
python excel_comparator.py file_lama.xlsx file_baru.xlsx --key "No"
```

### Export ke Excel

```bash
python excel_comparator.py file_lama.xlsx file_baru.xlsx --output laporan.xlsx
```

## CLI Options

| Option | Shortcut | Deskripsi |
|--------|----------|-----------|
| `file_old` | - | File Excel lama (positional) |
| `file_new` | - | File Excel baru (positional) |
| `--key` | `-k` | Kolom kunci untuk matching baris |
| `--output` | `-o` | Export hasil ke file Excel |
| `--config` | `-c` | File config custom |

## Demo dengan Sample Data

```bash
python generate_sample.py
python excel_comparator.py
```

Output:
```
File lama : sample/data_lama.xlsx
File baru : sample/data_revisi.xlsx
Key column: No

File lama: 10 baris, 5 kolom
File baru: 10 baris, 5 kolom

==================================================
LAPORAN PERBANDINGAN
==================================================

Ringkasan:
  - Baris baru     : 2
  - Baris dihapus  : 2
  - Cell berubah   : 6

--- PERUBAHAN (6) ---
  [No=1] Total: 5000000 → 5500000
  [No=2] Status: Pending → Lunas
  [No=4] Status: Cicilan → Lunas
  [No=5] Total: 3100000 → 3600000
  [No=6] Status: Pending → Lunas
  [No=8] Status: Pending → Lunas

--- BARIS BARU (2) ---
  [No=11]
  [No=12]

--- BARIS DIHAPUS (2) ---
  [No=9]
  [No=10]

Laporan Excel disimpan ke: output/laporan_perbandingan.xlsx
```

## Output Excel Report

Jika menggunakan `--output`, akan generate file Excel dengan sheet:

1. **Summary** - Ringkasan jumlah perubahan
2. **Perubahan** - Detail cell yang berubah (highlight kuning)
3. **Baris Baru** - Data baris yang ditambahkan (highlight hijau)
4. **Baris Dihapus** - Data baris yang dihapus (highlight merah)

## Mode Perbandingan

### 1. Dengan Key Column (Recommended)

```bash
python excel_comparator.py old.xlsx new.xlsx --key "ID"
```

Matching baris berdasarkan nilai kolom ID. Cocok untuk:
- Data yang baris-nya bisa berubah posisi
- Data dengan primary key (No, ID, NIK, dll)

### 2. Tanpa Key Column

```bash
python excel_comparator.py old.xlsx new.xlsx
```

Matching berdasarkan posisi baris (baris 1 vs baris 1, dst). Cocok untuk:
- Data yang urutan barisnya tetap
- File yang hanya ada perubahan nilai, bukan penambahan/penghapusan baris

## Catatan Penting

- Kedua file harus memiliki struktur kolom yang sama
- Key column harus unik (tidak ada duplikat)
- Perbandingan bersifat case-sensitive

## Pengembangan Selanjutnya

- [ ] Support multiple sheet
- [ ] Ignore kolom tertentu
- [ ] Threshold untuk perbandingan angka (toleransi selisih)
- [ ] Export ke format lain (HTML, PDF)

## Blog

[Link ke artikel blog] *(coming soon)*
