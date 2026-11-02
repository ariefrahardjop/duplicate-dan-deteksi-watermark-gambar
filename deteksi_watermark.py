#!/usr/bin/env python3
"""Deteksi watermark pada foto (cap GPS Map Camera / Timemark).

Cara kerja: setiap foto dibaca teksnya dengan OCR (tesseract), lalu dihitung
sinyal-sinyal cap: hari, bulan, tanggal, waktu, alamat, kata GPS, dan
koordinat. Jumlah sinyal = skor. Foto juga diukur strip gelap bawahnya
(sebagai info tambahan, bukan penentu).

Keputusan (divalidasi pada 431 foto kegiatan muskel, 2026-11-02):
  skor >= 4  -> Watermark  (55/55 benar, 0 salah pada data validasi)
  skor == 3  -> Ragu-ragu (perlu cek mata; pada data validasi 2 dari 6
                foto skor 3 ternyata berwatermark, 4 tidak)
  skor <= 2  -> Bukan watermark (370/370 benar pada data validasi)

Jenis stempel (tebakan terbaik, keputusan utama tetap kolom Keputusan):
  ada sinyal koordinat / kata GPS -> GPS Map Camera
  ada nama bulan tanpa sinyal GPS  -> Timemark
  selain itu (berwatermark)        -> GPS Map Camera

Tool ini TIDAK PERNAH mengubah foto asli: foto hanya dibaca. Opsi --pisah
menyalin (bukan memindah) foto ke subfolder watermark/, bukan_watermark/,
dan ragu_ragu/. Laporan selalu ditulis ke CSV dan Excel (bila openpyxl ada).

Pemakaian:
  python3 deteksi_watermark.py <folder_foto>
  python3 deteksi_watermark.py <folder_foto> --pisah
  python3 deteksi_watermark.py <folder_foto> --output hasil.xlsx --csv hasil.csv

Kebutuhan: python3, tesseract (perintah `tesseract` di PATH), Pillow.
Excel butuh openpyxl; tanpa itu, laporan CSV tetap dibuat.
"""

import argparse
import csv
import os
import re
import shutil
import subprocess
import sys
import tempfile

from PIL import Image

DAY = r"\b(senin|selasa|rabu|kamis|jumat|jum\'?at|sabtu|minggu)\b"
MONTH = r"\b(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember)\b"
DATE = r"\b\d{1,2}[-/ ]\d{1,2}[-/ ]\d{2,4}\b"
TIME = r"\b\d{1,2}[.:]\d{2}[.:]\d{2}\b|\b\d{1,2}:\d{2}\b"
ADDR = r"\b(jl\.?|jalan|kec\.?|kecamatan|kel\.?|kelurahan|surabaya|rt\.?|rw\.?|no\.?)\b"
GPSW = r"\b(gps|map camera|latitude|longitude|lintang|bujur)\b"
COORD = r"-?\d{2,3}\.\d{4,}"
GROUPS = [('hari', DAY), ('bulan', MONTH), ('tanggal', DATE), ('waktu', TIME),
          ('alamat', ADDR), ('gps_kata', GPSW), ('koordinat', COORD)]

EKSTENSI = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff', '.webp'}


def ocr(path, psm):
    """Jalankan tesseract, kembalikan teks (string kosong bila gagal)."""
    try:
        r = subprocess.run(['tesseract', path, 'stdout', '--psm', str(psm)],
                           capture_output=True, text=True, timeout=180)
        return r.stdout or ''
    except Exception:
        return ''


def cari_sinyal(teks):
    low = teks.lower()
    return sorted(nama for nama, pat in GROUPS if re.search(pat, low))


def bar_gelap(path):
    """Fraksi piksel sangat gelap di strip bawah 12% + keseragamannya.

    Hanya info tambahan: foto malam atau latar gelap bisa punya strip gelap
    tanpa watermark, jadi nilai ini tidak dipakai untuk keputusan.
    """
    try:
        with Image.open(path) as im:
            im = im.convert('L')
            w, h = im.size
            strip = im.crop((0, int(h * 0.88), w, h))
            px = list(strip.getdata())
            gelap = sum(1 for p in px if p < 45) / max(1, len(px))
            bw = max(1, w // 10)
            blok = []
            for i in range(10):
                seg = [p for x, p in enumerate(px) if (x % w) // bw == i]
                blok.append(sum(1 for p in seg if p < 45) / max(1, len(seg)))
            merata = sum(1 for b in blok if b > 0.6) / 10
            return round(gelap, 3), round(merata, 2)
    except Exception:
        return 0.0, 0.0


def klasifikasi(skor, sinyal):
    """Kembalikan (keputusan, jenis)."""
    if skor >= 4:
        keputusan = 'Watermark'
    elif skor == 3:
        keputusan = 'Ragu-ragu'
    else:
        keputusan = 'Bukan watermark'
    if keputusan == 'Bukan watermark':
        jenis = '-'
    elif 'koordinat' in sinyal or 'gps_kata' in sinyal:
        jenis = 'GPS Map Camera'
    elif 'bulan' in sinyal:
        jenis = 'Timemark'
    else:
        jenis = 'GPS Map Camera'
    return keputusan, jenis


def proses_foto(path):
    """Proses satu foto, kembalikan dict hasil."""
    try:
        with Image.open(path) as im:
            im.verify()
    except Exception:
        return {'file': os.path.basename(path), 'skor': -1, 'sinyal': [],
                'bar_gelap': 0.0, 'bar_merata': 0.0,
                'keputusan': 'Gagal dibaca', 'jenis': '-'}
    teks = ocr(path, 11)
    sinyal = cari_sinyal(teks)
    skor = len(sinyal)
    # Pass kedua (strip bawah diperbesar) hanya untuk skor rendah.
    if skor < 3:
        try:
            with Image.open(path) as im:
                w, h = im.size
                strip = im.crop((0, int(h * 0.70), w, h))
                strip = strip.resize((strip.width * 2, strip.height * 2))
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                    sp = tmp.name
                strip.save(sp)
            try:
                sinyal = sorted(set(sinyal) | set(cari_sinyal(ocr(sp, 6))))
                skor = len(sinyal)
            finally:
                if os.path.exists(sp):
                    os.remove(sp)
        except Exception:
            pass
    gelap, merata = bar_gelap(path)
    keputusan, jenis = klasifikasi(skor, sinyal)
    return {'file': os.path.basename(path), 'skor': skor,
            'sinyal': sinyal, 'bar_gelap': gelap, 'bar_merata': merata,
            'keputusan': keputusan, 'jenis': jenis}


def tulis_csv(path, hasil):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['No', 'Nama File', 'Skor', 'Sinyal', 'Bar Gelap',
                    'Bar Merata', 'Keputusan', 'Jenis Stempel'])
        for i, h in enumerate(hasil, 1):
            w.writerow([i, h['file'], h['skor'], ','.join(h['sinyal']),
                        h['bar_gelap'], h['bar_merata'],
                        h['keputusan'], h['jenis']])


def tulis_excel(path, hasil):
    try:
        import openpyxl
    except ImportError:
        return False
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Detail Foto'
    ws.append(['No', 'Nama File', 'Skor', 'Sinyal', 'Bar Gelap',
               'Bar Merata', 'Keputusan', 'Jenis Stempel'])
    for i, h in enumerate(hasil, 1):
        ws.append([i, h['file'], h['skor'], ','.join(h['sinyal']),
                   h['bar_gelap'], h['bar_merata'],
                   h['keputusan'], h['jenis']])
    ws2 = wb.create_sheet('Ringkasan')
    from collections import Counter
    c = Counter(h['keputusan'] for h in hasil)
    ws2.append(['Keterangan', 'Jumlah'])
    ws2.append(['Total foto diproses', len(hasil)])
    for k in ['Watermark', 'Ragu-ragu', 'Bukan watermark', 'Gagal dibaca']:
        ws2.append([k, c.get(k, 0)])
    ws2.append(['Catatan', 'Skor>=4 Watermark; skor 3 Ragu-ragu (cek mata); '
                            'skor<=2 Bukan watermark. Foto asli tidak diubah.'])
    wb.save(path)
    return True


def main():
    ap = argparse.ArgumentParser(description='Deteksi watermark pada foto.')
    ap.add_argument('folder', help='Folder berisi foto yang akan diperiksa')
    ap.add_argument('--output', default=None, help='File Excel hasil (baku: '
                    '<folder>/hasil_deteksi_watermark.xlsx)')
    ap.add_argument('--csv', default=None, help='File CSV hasil (baku: '
                    '<folder>/hasil_deteksi_watermark.csv)')
    ap.add_argument('--pisah', action='store_true',
                    help='Salin foto ke subfolder watermark/, bukan_watermark/, '
                         'ragu_ragu/ di dalam folder input (asli tidak dipindah)')
    args = ap.parse_args()
    folder = args.folder
    if not os.path.isdir(folder):
        print(f'Folder tidak ditemukan: {folder}', file=sys.stderr)
        return 1
    daftar = sorted(f for f in os.listdir(folder)
                    if os.path.isfile(os.path.join(folder, f))
                    and os.path.splitext(f)[1].lower() in EKSTENSI)
    if not daftar:
        print('Tidak ada file gambar di folder itu.', file=sys.stderr)
        return 1
    hasil = []
    for n, nama in enumerate(daftar, 1):
        h = proses_foto(os.path.join(folder, nama))
        hasil.append(h)
        print(f"[{n}/{len(daftar)}] {nama} skor={h['skor']} "
              f"sinyal={','.join(h['sinyal'])} -> {h['keputusan']}", flush=True)
    out_xlsx = args.output or os.path.join(folder, 'hasil_deteksi_watermark.xlsx')
    out_csv = args.csv or os.path.join(folder, 'hasil_deteksi_watermark.csv')
    tulis_csv(out_csv, hasil)
    ok_xlsx = tulis_excel(out_xlsx, hasil)
    if args.pisah:
        peta = {'Watermark': 'watermark', 'Bukan watermark': 'bukan_watermark',
                'Ragu-ragu': 'ragu_ragu', 'Gagal dibaca': 'gagal_dibaca'}
        for h in hasil:
            tujuan = os.path.join(folder, peta[h['keputusan']])
            os.makedirs(tujuan, exist_ok=True)
            shutil.copy2(os.path.join(folder, h['file']),
                         os.path.join(tujuan, h['file']))
    from collections import Counter
    c = Counter(h['keputusan'] for h in hasil)
    print(f"\nSELESAI: {len(hasil)} foto | Watermark={c.get('Watermark', 0)} | "
          f"Ragu-ragu={c.get('Ragu-ragu', 0)} | "
          f"Bukan watermark={c.get('Bukan watermark', 0)} | "
          f"Gagal dibaca={c.get('Gagal dibaca', 0)}")
    print(f'CSV: {out_csv}')
    print(f"Excel: {out_xlsx if ok_xlsx else '(openpyxl tidak ada, Excel dilewati)'}")
    if args.pisah:
        print('Foto disalin (bukan dipindah) ke subfolder watermark/, '
              'bukan_watermark/, ragu_ragu/.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
