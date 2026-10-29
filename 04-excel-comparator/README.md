# 04 · Excel Comparator

> Bab 4 kisah [Bu Sari](../README.md): Kamis, finance mengirim file revisi dan Bu Sari harus
> tahu persis sel mana yang berubah.

Membandingkan 2 file Excel dan menampilkan perbedaannya secara detail.

## Cerita

Email dari finance singkat: "Tolong cek file revisi ini, ada beberapa koreksi." Filenya
ratusan baris dan tidak dijelaskan baris mana yang berubah. Membandingkan dua file secara
manual berarti scroll berdampingan dan mencatat perbedaan satu per satu.

Script ini menghasilkan laporan:

- sel mana yang berubah (nilai lama → nilai baru)
- baris mana yang baru ditambahkan
- baris mana yang dihapus

## Instalasi

```bash
pip install -r requirements.txt
```

## Cara Menggunakan

```bash
# pakai config.yaml di folder utility
python excel_comparator.py

# dua file lewat CLI
python excel_comparator.py file_lama.xlsx file_baru.xlsx

# matching berdasarkan kolom kunci (disarankan)
python excel_comparator.py file_lama.xlsx file_baru.xlsx --key "No"

# sekaligus export laporan Excel
python excel_comparator.py file_lama.xlsx file_baru.xlsx --key "No" -o laporan.xlsx
```

Path dari CLI mengikuti current working directory. Path `file_old`, `file_new`, dan `output`
di dalam config di-resolve relatif terhadap folder utility.

## Opsi CLI

| Opsi | Shortcut | Deskripsi |
|------|----------|-----------|
| `file_old` | – | File Excel lama (positional) |
| `file_new` | – | File Excel baru (positional) |
| `--key` | `-k` | Kolom kunci untuk matching baris |
| `--output` | `-o` | Export hasil ke file Excel |
| `--config` | `-c` | File config YAML (default `config.yaml`) |

## Demo (hasil run nyata)

```bash
python generate_sample.py
python excel_comparator.py
```

```
File lama : /path/ke/repo/04-excel-comparator/sample/data_lama.xlsx
File baru : /path/ke/repo/04-excel-comparator/sample/data_revisi.xlsx
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
  [No=1] Total: 5000000 -> 5500000
  [No=2] Status: Pending -> Lunas
  [No=4] Status: Cicilan -> Lunas
  [No=5] Total: 3100000 -> 3600000
  [No=6] Status: Pending -> Lunas
  [No=8] Status: Pending -> Lunas

--- BARIS BARU (2) ---
  [No=11]
  [No=12]

--- BARIS DIHAPUS (2) ---
  [No=9]
  [No=10]

Laporan Excel disimpan ke: /path/ke/repo/04-excel-comparator/output/laporan_perbandingan.xlsx
```

## Output Excel

Sheet yang dihasilkan (folder output dibuat otomatis jika belum ada):

1. **Summary** — jumlah baris baru, dihapus, dan sel berubah
2. **Perubahan** — detail sel yang berubah (highlight kuning)
3. **Baris Baru** — data baris baru (highlight hijau)
4. **Baris Dihapus** — data baris yang dihapus (highlight merah)

## Mode Perbandingan

**Dengan key column** (disarankan) — baris dicocokkan berdasarkan nilai kolom kunci, jadi
perubahan urutan baris tidak dianggap perbedaan:

```bash
python excel_comparator.py old.xlsx new.xlsx --key "No"
```

**Tanpa key column** — baris dibandingkan berdasarkan posisi (baris 1 vs baris 1).

## Kunci Duplikat

Kolom kunci sebaiknya unik. Jika ada nilai duplikat, script tetap berjalan memakai baris
**pertama** yang cocok, dan menampilkan peringatan berisi daftar kuncinya:

```
Peringatan: file lama kolom kunci 'No' duplikat: 1, 2; memakai baris pertama
```

## Catatan

- Kedua file sebaiknya memiliki struktur kolom yang sama.
- Perbandingan bersifat case-sensitive.
- Daftar baris baru/dihapus diurutkan secara numerik bila kuncinya berupa angka.

## Roadmap Singkat

- [ ] Abaikan kolom tertentu
- [ ] Toleransi selisih untuk kolom angka
- [ ] Export HTML/PDF
