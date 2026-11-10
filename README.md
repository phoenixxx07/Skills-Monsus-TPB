# Monsus Rekonsiliasi — Skill untuk AI (Claude Code)

Skill reusable untuk melakukan **Monitoring Khusus (Monsus)** Kawasan Berikat / TPB:
merekonsiliasi data barang antara **IT Inventory perusahaan** dengan **data resmi
deklarasi Bea Cukai** (CEISA/TPB).

Dibuat untuk dipakai sebagai *Claude Code skill* — cukup taruh folder ini (atau isi
`SKILL.md`-nya) di `~/.claude/skills/monsus-rekonsiliasi/` atau `<project>/.claude/skills/`,
lalu minta Claude melakukan analisis Monsus dengan mengunggah data IT Inventory dan
data Bea Cukai.

> Termasuk modul kasus khusus **insiden kebakaran/barang musnah**: audit selisih klaim
> Stock Opname pasca-insiden + penghitungan estimasi tagihan pungutan (BM/PPN/PPh Impor)
> atas selisihnya.

## Struktur

```
SKILL.md                              # entry point skill, baca ini dulu
references/
  alur_kategori_barang.md             # aturan alur per kategori barang (BB, WIP, BJ, Modal, Scrap)
  pemetaan_tabel_bc.md                 # pemetaan jenis dokumen BC -> tabel database
  struktur_sheet_output.md            # struktur sheet Excel hasil rekonsiliasi reguler
  kasus_khusus_kebakaran.md           # metodologi kasus khusus: audit selisih klaim + tagihan pungutan
scripts/
  build_rekap_monsus.py               # script generator Excel rekonsiliasi reguler (config-driven)
  kasus_khusus_kebakaran/
    audit_selisih_kebakaran.py        # langkah 1: deteksi selisih klaim SO pasca-insiden
    hitung_tagihan_pungutan.py        # langkah 2: estimasi tagihan BM/PPN/PPh atas selisih
examples/                             # (opsional) data contoh/fiktif untuk uji coba
```

## Cara pakai cepat

1. Install dependency: `pip install -r requirements.txt`
2. Siapkan folder data berisi file IT Inventory (lihat `SKILL.md` bagian 2 untuk daftar
   file yang dibutuhkan).
3. Edit `CONFIG` di `scripts/build_rekap_monsus.py`: nama perusahaan, NPWP, periode,
   path folder data.
4. Set kredensial database lewat environment variable (lihat komentar di bagian atas
   script) — **jangan** hardcode kredensial ke dalam file yang di-commit.
5. Jalankan:
   ```
   python scripts/build_rekap_monsus.py
   ```
6. Hasil: file Excel rekonsiliasi dengan sheet `Ringkasan Monsus`, `Sandingan Dok
   Masuk`, `Sandingan Dok Keluar`.

## Mengembangkan skill ini

Repo ini khusus untuk topik Monsus. Kalau ke depan ada spesifikasi teknis baru
(jenis dokumen BC lain, kategori barang baru, perubahan struktur laporan IT
Inventory, kasus khusus lain selain kebakaran, dll.), tambahkan/update di:

- `references/` — untuk pengetahuan/aturan baru (buat file `.md` baru per topik)
- `scripts/` — untuk logika pemrosesan baru
- `SKILL.md` — update ringkasan & langkah eksekusi agar tetap jadi entry point yang akurat

## Keamanan

Jangan commit ke repo publik:
- Kredensial database (host, user, password) — gunakan environment variable
- Data riil milik perusahaan klien (nama, NPWP, data transaksi) — gunakan data fiktif
  di `examples/` jika perlu contoh
