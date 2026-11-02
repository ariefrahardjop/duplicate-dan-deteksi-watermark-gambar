#!/usr/bin/env python3
"""Deteksi file double (duplikat) dalam satu folder, dengan Python.

Dua mode:
  1. Duplikat persis (selalu jalan): file yang byte-nya identik, dideteksi
     lewat hash MD5. Ini pasti: MD5 sama berarti isi file sama persis.
  2. Foto mirip (--mirip): gambar yang isinya mirip walau sudah diubah
     ukuran/dikompres ulang, dideteksi lewat dHash 64-bit (algoritma yang
     sama dengan halaman "Cek Foto Double" di admin muskel: kecilkan ke
     9x8, grayscale, bandingkan piksel bersebelahan). Dua foto dianggap
     satu grup bila jarak Hamming hash-nya <= ambang (--jarak, bawaan 5).

Tool ini TIDAK PERNAH menghapus, memindah, atau mengubah file asli.
Opsi --pisah hanya menyalin (bukan memindah) file-file duplikat ke
subfolder duplikat_persis/ dan foto_mirip/ untuk ditinjau manual.
Folder bersih adalah hasil akhir yang sudah bersih dari duplikat:
berisi file yang tidak masuk grup mana pun, ditambah satu perwakilan
dari tiap grup duplikat (anggota pertama grup seperti tercantum di
laporan). Folder grup tetap berisi semua anggota sebagai arsip peninjauan.

Pemakaian:
  python3 deteksi_file_double_bersih.py <folder>
  python3 deteksi_file_double_bersih.py <folder> --mirip
  python3 deteksi_file_double_bersih.py <folder> --mirip --jarak 3 --pisah
  python3 deteksi_file_double_bersih.py <folder> --no-rekursif --output hasil.xlsx --csv hasil.csv

Bawaan: pemindaian rekursif (termasuk subfolder). Laporan CSV dan Excel
ditulis di dalam folder input, ditambah CSV tersendiri untuk file bersih.
Folder hasil --pisah dan laporan dari proses sebelumnya dilewati agar
salinan tidak ikut terhitung duplikat.

Kebutuhan: python3, Pillow. Excel butuh openpyxl; tanpa itu, CSV tetap dibuat.
"""

import argparse
import csv
import hashlib
import os
import shutil
import sys

from PIL import Image

IMG_EXTS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tif', '.tiff'}
LEWATI_DIR = {'duplikat_persis', 'foto_mirip', 'bersih',
              'watermark', 'bukan_watermark', 'ragu_ragu'}


def md5_file(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        while True:
            blok = f.read(1024 * 1024)
            if not blok:
                break
            h.update(blok)
    return h.hexdigest()


def dhash_file(path):
    """dHash 64-bit, sama dengan helper dhash_helper.php di admin muskel.

    Kecilkan ke 9x8, grayscale (0.299R + 0.587G + 0.114B), tiap baris
    bandingkan piksel kiri vs kanan (1 bila kiri lebih terang). Hasil:
    16 karakter hex, atau None bila gambar tidak bisa dibaca.
    """
    try:
        with Image.open(path) as im:
            if im.mode in ('RGBA', 'LA') or \
                    (im.mode == 'P' and 'transparency' in im.info):
                latar = Image.new('RGB', im.size, (0, 0, 0))
                latar.paste(im, mask=im.convert('RGBA').split()[3])
                im = latar
            elif im.mode != 'RGB':
                im = im.convert('RGB')
            kecil = im.resize((9, 8), Image.LANCZOS)
            px = list(kecil.convert('L').getdata())
            bit = []
            for y in range(8):
                for x in range(8):
                    bit.append('1' if px[y * 9 + x] > px[y * 9 + x + 1] else '0')
            return '%016x' % int(''.join(bit), 2)
    except Exception:
        return None


def hamming(a, b):
    return bin(int(a, 16) ^ int(b, 16)).count('1')


class UnionFind:
    def __init__(self, n):
        self.ortu = list(range(n))

    def cari(self, x):
        while self.ortu[x] != x:
            self.ortu[x] = self.ortu[self.ortu[x]]
            x = self.ortu[x]
        return x

    def gabung(self, a, b):
        ra, rb = self.cari(a), self.cari(b)
        if ra != rb:
            self.ortu[rb] = ra


def kumpulkan_file(folder, rekursif, lewati_file=frozenset()):
    hasil = []
    if rekursif:
        for akar, dirs, files in os.walk(folder):
            dirs[:] = sorted(d for d in dirs if d not in LEWATI_DIR)
            for nama in sorted(files):
                p = os.path.join(akar, nama)
                if os.path.abspath(p) in lewati_file:
                    continue
                hasil.append((p, os.path.relpath(p, folder)))
    else:
        for nama in sorted(os.listdir(folder)):
            p = os.path.join(folder, nama)
            if os.path.isfile(p) and os.path.abspath(p) not in lewati_file:
                hasil.append((p, nama))
    return hasil


def tulis_csv(path, baris):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['Jenis', 'Grup', 'File', 'Ukuran (byte)', 'Hash',
                    'Jarak ke referensi'])
        for b in baris:
            w.writerow(b)


def tulis_csv_bersih(path, bersih_final):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['No', 'File', 'Ukuran (byte)', 'MD5', 'Perwakilan grup'])
        for i, (a, wak) in enumerate(bersih_final, 1):
            w.writerow([i, a['rel'], a['ukuran'], a['hash'], wak])


def tulis_excel(path, grup_persis, grup_mirip, bersih, ringkasan):
    try:
        import openpyxl
    except ImportError:
        return False
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Duplikat Persis'
    ws.append(['Grup', 'File', 'Ukuran (byte)', 'MD5'])
    for g in grup_persis:
        for item in g['anggota']:
            ws.append([g['nama'], item['rel'], item['ukuran'], item['hash']])
    if grup_mirip is not None:
        ws2 = wb.create_sheet('Foto Mirip')
        ws2.append(['Grup', 'File', 'Ukuran (byte)', 'dHash',
                    'Jarak ke referensi'])
        for g in grup_mirip:
            for item in g['anggota']:
                ws2.append([g['nama'], item['rel'], item['ukuran'],
                            item['hash'], item['jarak']])
    wsb = wb.create_sheet('File Bersih')
    wsb.append(['No', 'File', 'Ukuran (byte)', 'MD5', 'Perwakilan grup'])
    for i, (a, wak) in enumerate(bersih, 1):
        wsb.append([i, a['rel'], a['ukuran'], a['hash'], wak])
    ws3 = wb.create_sheet('Ringkasan')
    for k, v in ringkasan:
        ws3.append([k, v])
    wb.save(path)
    return True


def main():
    ap = argparse.ArgumentParser(description='Deteksi file double.')
    ap.add_argument('folder', help='Folder yang akan dipindai')
    ap.add_argument('--no-rekursif', action='store_true',
                    help='Hanya folder itu sendiri, tanpa subfolder')
    ap.add_argument('--mirip', action='store_true',
                    help='Tambah deteksi foto mirip (dHash) selain duplikat persis')
    ap.add_argument('--jarak', type=int, default=5,
                    help='Ambang jarak dHash untuk foto mirip (bawaan 5)')
    ap.add_argument('--pisah', action='store_true',
                    help='Salin grup duplikat ke subfolder duplikat_persis/ '
                         'dan foto_mirip/ (asli tidak dipindah)')
    ap.add_argument('--output', default=None,
                    help='File Excel hasil (bawaan <folder>/hasil_file_double.xlsx)')
    ap.add_argument('--csv', default=None,
                    help='File CSV hasil (bawaan <folder>/hasil_file_double.csv)')
    args = ap.parse_args()
    folder = args.folder
    if not os.path.isdir(folder):
        print(f'Folder tidak ditemukan: {folder}', file=sys.stderr)
        return 1
    jarak = max(0, min(64, args.jarak))

    # Laporan ditulis di folder ini juga, jadi dikecualikan dari pemindaian
    # (laporan akan ditimpa; laporan lama tidak boleh ikut terhitung).
    out_csv = args.csv or os.path.join(folder, 'hasil_file_double.csv')
    out_xlsx = args.output or os.path.join(folder, 'hasil_file_double.xlsx')
    akar_csv, ekst_csv = os.path.splitext(out_csv)
    out_csv_bersih = akar_csv + '_bersih' + ekst_csv
    lewati_file = {os.path.abspath(p) for p in
                   (out_csv, out_xlsx, out_csv_bersih,
                    os.path.join(folder, 'hasil_file_double.csv'),
                    os.path.join(folder, 'hasil_file_double.xlsx'),
                    os.path.join(folder, 'hasil_file_double_bersih.csv'))}
    daftar = kumpulkan_file(folder, not args.no_rekursif, lewati_file)
    print(f'Memindai {len(daftar)} file di {folder} ...', flush=True)

    # --- Duplikat persis (MD5) ---
    gagal = []
    per_md5 = {}
    for n, (p, rel) in enumerate(daftar, 1):
        if n % 100 == 0 or n == len(daftar):
            print(f'  hash {n}/{len(daftar)}', flush=True)
        try:
            h = md5_file(p)
            per_md5.setdefault(h, []).append(
                {'path': p, 'rel': rel, 'ukuran': os.path.getsize(p), 'hash': h})
        except Exception:
            gagal.append(rel)
    grup_persis = []
    for h, anggota in per_md5.items():
        if len(anggota) >= 2:
            anggota.sort(key=lambda a: a['rel'])
            grup_persis.append({'nama': '', 'anggota': anggota})
    grup_persis.sort(key=lambda g: (-len(g['anggota']), g['anggota'][0]['rel']))
    for i, g in enumerate(grup_persis, 1):
        g['nama'] = f'P{i:02d}'

    # --- Foto mirip (dHash) ---
    grup_mirip = None
    gagal_mirip = []
    if args.mirip:
        gambar = []
        for p, rel in daftar:
            if os.path.splitext(rel)[1].lower() not in IMG_EXTS:
                continue
            h = dhash_file(p)
            if h is None:
                gagal_mirip.append(rel)
                continue
            gambar.append({'path': p, 'rel': rel,
                           'ukuran': os.path.getsize(p), 'hash': h})
        gambar.sort(key=lambda a: a['rel'])
        uf = UnionFind(len(gambar))
        for i in range(len(gambar)):
            for j in range(i + 1, len(gambar)):
                if hamming(gambar[i]['hash'], gambar[j]['hash']) <= jarak:
                    uf.gabung(i, j)
        per_akar = {}
        for i, g in enumerate(gambar):
            per_akar.setdefault(uf.cari(i), []).append(g)
        grup_mirip = []
        for anggota in per_akar.values():
            if len(anggota) >= 2:
                anggota.sort(key=lambda a: a['rel'])
                ref = anggota[0]['hash']
                for a in anggota:
                    a['jarak'] = hamming(a['hash'], ref)
                anggota.sort(key=lambda a: (a['jarak'], a['rel']))
                grup_mirip.append({'nama': '', 'anggota': anggota})
        grup_mirip.sort(key=lambda g: (-len(g['anggota']), g['anggota'][0]['rel']))
        for i, g in enumerate(grup_mirip, 1):
            g['nama'] = f'M{i:02d}'

    # --- File bersih: berhasil dibaca tapi tidak masuk grup mana pun ---
    # (gambar yang gagal dihitung dHash tidak dihitung bersih bila --mirip,
    # karena kemiripannya tidak diketahui; ia dilaporkan terpisah.)
    rel_grup = set()
    for g in grup_persis:
        rel_grup.update(a['rel'] for a in g['anggota'])
    if grup_mirip:
        rel_grup.update(a['rel'] for g in grup_mirip for a in g['anggota'])
    set_gagal_mirip = set(gagal_mirip)
    bersih = []
    for anggota in per_md5.values():
        for a in anggota:
            if a['rel'] not in rel_grup and a['rel'] not in set_gagal_mirip:
                bersih.append(a)
    bersih.sort(key=lambda a: a['rel'])

    # --- Perwakilan grup: satu file per grup ikut ke folder bersih ---
    # Perwakilan = anggota pertama grup seperti tercantum di laporan
    # (grup persis: urutan nama file; grup mirip: urutan jarak lalu nama).
    # Satu file hanya satu baris di bersih walau mewakili beberapa grup.
    perwakilan = {}
    for g in grup_persis:
        perwakilan.setdefault(g['anggota'][0]['rel'], []).append(g['nama'])
    if grup_mirip:
        for g in grup_mirip:
            perwakilan.setdefault(g['anggota'][0]['rel'], []).append(g['nama'])
    by_rel = {}
    for anggota in per_md5.values():
        for a in anggota:
            by_rel[a['rel']] = a
    bersih_final = [(a, '-') for a in bersih]
    for rel, nama_grup in sorted(perwakilan.items()):
        bersih_final.append((by_rel[rel], ', '.join(nama_grup)))
    bersih_final.sort(key=lambda t: t[0]['rel'])

    # --- Laporan ---
    baris_csv = []
    for g in grup_persis:
        for a in g['anggota']:
            baris_csv.append(['Duplikat persis', g['nama'], a['rel'],
                              a['ukuran'], a['hash'], ''])
    if grup_mirip:
        for g in grup_mirip:
            for a in g['anggota']:
                baris_csv.append(['Foto mirip', g['nama'], a['rel'],
                                  a['ukuran'], a['hash'], a['jarak']])
    tulis_csv(out_csv, baris_csv)
    tulis_csv_bersih(out_csv_bersih, bersih_final)
    ringkasan = [
        ('Total file dipindai', len(daftar)),
        ('Grup duplikat persis', len(grup_persis)),
        ('File dalam grup duplikat persis',
         sum(len(g['anggota']) for g in grup_persis)),
    ]
    if grup_mirip is not None:
        ringkasan += [
            ('Grup foto mirip', len(grup_mirip)),
            ('File dalam grup foto mirip',
             sum(len(g['anggota']) for g in grup_mirip)),
            ('Ambang jarak dHash', jarak),
        ]
    ringkasan.append(('File bersih (hasil akhir)', len(bersih_final)))
    if args.mirip:
        ringkasan.append(('Arti bersih', 'File tanpa duplikat + satu '
                                        'perwakilan dari tiap grup '
                                        '(persis & mirip)'))
    else:
        ringkasan.append(('Arti bersih', 'File tanpa duplikat persis + '
                                        'satu perwakilan dari tiap grup '
                                        'duplikat persis (kemiripan visual '
                                        'tidak diperiksa tanpa --mirip)'))
    ringkasan.append(('Gambar gagal dihitung dHash', len(gagal_mirip)))
    for rel in gagal_mirip:
        ringkasan.append(('Gagal dHash', rel))
    ringkasan.append(('File gagal dibaca', len(gagal)))
    for rel in gagal:
        ringkasan.append(('Gagal dibaca', rel))
    ringkasan.append(('Catatan', 'MD5 sama = isi identik persis. dHash = kemiripan '
                                 'visual; jarak 0 identik, makin kecil makin mirip. '
                                 'File asli tidak diubah/dihapus.'))
    ok_xlsx = tulis_excel(out_xlsx, grup_persis, grup_mirip, bersih_final, ringkasan)

    # --- Pisah (salinan saja) ---
    if args.pisah:
        for g in grup_persis:
            for a in g['anggota']:
                t = os.path.join(folder, 'duplikat_persis', g['nama'], a['rel'])
                os.makedirs(os.path.dirname(t), exist_ok=True)
                shutil.copy2(a['path'], t)
        if grup_mirip:
            for g in grup_mirip:
                for a in g['anggota']:
                    t = os.path.join(folder, 'foto_mirip', g['nama'], a['rel'])
                    os.makedirs(os.path.dirname(t), exist_ok=True)
                    shutil.copy2(a['path'], t)
        for a, _wak in bersih_final:
            t = os.path.join(folder, 'bersih', a['rel'])
            os.makedirs(os.path.dirname(t), exist_ok=True)
            shutil.copy2(a['path'], t)

    print(f"\nSELESAI: {len(daftar)} file dipindai | "
          f"duplikat persis: {len(grup_persis)} grup "
          f"({sum(len(g['anggota']) for g in grup_persis)} file)"
          + (f" | foto mirip: {len(grup_mirip)} grup "
             f"({sum(len(g['anggota']) for g in grup_mirip)} file)"
             if grup_mirip is not None else '')
          + f" | bersih: {len(bersih_final)} file"
          + f" | gagal dibaca: {len(gagal)}")
    print(f'CSV: {out_csv}')
    print(f'CSV bersih: {out_csv_bersih}')
    print(f"Excel: {out_xlsx if ok_xlsx else '(openpyxl tidak ada, Excel dilewati)'}")
    if args.pisah:
        print('Grup disalin (bukan dipindah) ke subfolder duplikat_persis/, '
              'foto_mirip/, dan bersih/.')
    if gagal:
        print('File gagal dibaca:', ', '.join(gagal))
    return 0


if __name__ == '__main__':
    sys.exit(main())
