# Alat Deteksi Foto: Watermark dan File Double

Dua script Python untuk memeriksa foto, saling melengkapi. Keduanya hanya
membaca foto, tidak pernah mengubah, memindah, atau menghapus file asli.
Opsi `--pisah` hanya menyalin file ke subfolder.

| Tool | File | Fungsi |
| --- | --- | --- |
| Deteksi watermark | `../alat-watermark/deteksi_watermark.py` | Memisahkan foto yang ada capnya (GPS Map Camera / Timemark) dari yang tidak |
| Deteksi file double | `../alat-file-double/deteksi_file_double.py` | Mencari file yang duplikat persis dan foto yang mirip |

Detail lengkap masing-masing ada di README di foldernya sendiri:
`../alat-watermark/README.md` dan `../alat-file-double/README.md`.

## Kebutuhan

- Python 3 untuk keduanya.
- Alat watermark butuh tambahan: Tesseract OCR (`tesseract` di PATH) dan
  Pillow. Karena itu alat watermark **tidak bisa jalan di server
  kartarsurabaya.web.id** (di sana tidak ada tesseract); alat ini jalan di
  VM ini atau di PC yang sudah terpasang tesseract.
- Alat file double butuh tambahan: Pillow saja (tanpa tesseract), jadi
  lebih ringan.
- Laporan Excel di keduanya butuh openpyxl; tanpa itu, laporan CSV tetap
  dibuat (`pip install Pillow openpyxl`, Ubuntu/Debian:
  `sudo apt install tesseract-ocr`).

## Cara pakai

```bash
# 1. Deteksi watermark: pisahkan foto bercap dari yang tidak
python3 ../alat-watermark/deteksi_watermark.py /path/ke/folder_foto --pisah

# 2. Deteksi file double: cari duplikat persis + foto mirip, lalu pisahkan
python3 ../alat-file-double/deteksi_file_double.py /path/ke/folder_foto --mirip --pisah
```

Kedua perintah bisa dijalankan berurutan pada folder yang sama. Urutan yang
disarankan: deteksi file double dulu (rapikan duplikat), lalu deteksi
watermark pada foto yang tersisa.

## Hasil

Keduanya menulis laporan CSV dan Excel di dalam folder foto, dan dengan
`--pisah` menyalin file ke subfolder:

- Alat watermark: `watermark/`, `bukan_watermark/`, `ragu_ragu/`, plus
  `hasil_deteksi_watermark.csv` dan `.xlsx`.
- Alat file double: `duplikat_persis/grup_01/...`,
  `foto_mirip/grup_01/...`, dan `bersih/` (file yang tidak duplikat),
  plus `hasil_file_double.csv`, `hasil_file_double_bersih.csv`, dan
  `.xlsx` (pemindaian rekursif, jadi subfolder ikut diperiksa; subfolder
  hasil `--pisah` dan laporan dari proses sebelumnya dilewati otomatis).

## Aturan keputusan

**Watermark** (dari OCR teks cap): sinyal yang dihitung adalah hari,
bulan, tanggal, waktu, alamat, kata GPS, dan koordinat. Skor 4 ke atas
berarti Watermark, skor 3 berarti Ragu-ragu (perlu cek mata), skor 2 ke
bawah berarti Bukan watermark. Jenis stempel ditebak dari sinyal: ada
koordinat/kata GPS berarti GPS Map Camera, ada nama bulan tanpa sinyal
GPS berarti Timemark.

**File double**: duplikat persis ditentukan hash MD5 (sama berarti isi
identik persis). Foto mirip ditentukan dHash 64-bit, algoritma yang sama
dengan halaman "Cek Foto Double" di admin muskel; jarak Hamming 0
berarti hash identik, ambang bawaan 5 (sama dengan bawaan halaman admin).

## Validasi yang sudah dilakukan

- Watermark: aturan skor diturunkan dari 431 foto kegiatan muskel
  (2026-11-02): 0 foto salah kelompok, hanya 6 foto skor 3 yang masuk
  Ragu-ragu (2 berwatermark, 4 tidak), tebakan jenis benar 57 dari 57.
  Uji end-to-end 4 foto buatan (cap GPS dengan dan tanpa bar hitam, cap
  Timemark, foto bersih) semuanya benar, dan foto asli terbukti utuh.
- File double: uji sintetis menghasilkan 1 grup persis (3 file termasuk
  satu di subfolder; file yang ukurannya sama tapi isinya beda tidak
  ikut), 1 grup mirip (asli, versi diperkecil, versi dikompres, versi
  PNG), foto yang memang beda tidak ikut, dan proses ulang tetap aman
  karena salinan dilewati. dHash Python dibanding helper PHP di admin
  untuk foto yang sama beda 2 bit dari 64, masih jauh di bawah ambang 5,
  jadi pengelompokannya sama.

## Keterbatasan

- OCR bisa gagal membaca cap yang buram atau kecil; foto seperti itu
  cenderung jatuh ke Ragu-ragu atau Bukan watermark.
- Mode foto mirip hanya untuk gambar; dokumen (PDF, Word) hanya
  terdeteksi lewat duplikat persis.
- File yang gagal dibaca (rusak, 0 byte) dilaporkan terpisah, bukan
  dihapus atau diubah.
