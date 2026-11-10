---
name: monsus-rekonsiliasi
description: >-
  Melakukan analisis Monitoring Khusus (Monsus) Kawasan Berikat / TPB: merekonsiliasi
  data barang antara IT Inventory perusahaan (Laporan Mutasi Barang & Laporan Per
  Dokumen) dengan data resmi deklarasi Bea Cukai (CEISA/TPB: BC 2.3, BC 4.0, BC 2.7,
  BC 2.5, BC 3.0, BC 4.1, BC 2.6.1/2.6.2). Gunakan skill ini ketika user mengunggah
  data IT Inventory (Laporan Mutasi Bahan Baku/Barang Jadi/Barang Sisa/Mesin, Laporan
  Stock Opname, Laporan Pemasukan/Pengeluaran Per Dokumen) dan/atau data Bea Cukai,
  lalu meminta dibuatkan rekonsiliasi/penyandingan/monsus/audit kesesuaian saldo buku.
  TIDAK mencakup kasus khusus kebakaran/klaim barang terbakar (musnah) — itu di luar
  lingkup skill ini dan butuh penanganan tersendiri.
---

# Monsus (Monitoring Khusus) Kawasan Berikat — Rekonsiliasi IT Inventory vs Data Bea Cukai

Skill ini membungkus pola kerja standar untuk mengaudit/merekonsiliasi data barang
perusahaan pemilik fasilitas Kawasan Berikat / TPB (Tempat Penimbunan Berikat) antara:

1. **IT Inventory perusahaan** (buku mutasi barang + laporan per dokumen pabean), dan
2. **Data resmi deklarasi Bea Cukai** (database CEISA/TPB).

Cakupan skill ini adalah rekonsiliasi **umum** (saldo, dokumen, kuantitas). Kasus
khusus seperti investigasi kebakaran/klaim barang musnah **tidak** dibahas di sini —
itu memerlukan analisis forensik tambahan (perbandingan SO pasca-insiden vs saldo
buku) yang di luar lingkup skill ini.

## 1. Konsep Inti: 3 Pilar Data

```
        ┌───────────────────────────────┐
        │  1. Laporan Mutasi Barang     │
        │  (Buku Mutasi IT Inventory)  │
        └───────────────┬───────────────┘
                         │
                         ▼
┌───────────────────────────────┐   PENYANDINGAN    ┌───────────────────────────────┐
│ 2. Laporan Per Dokumen (IT)   │◄──────────────────►│ 3. Data Dokumen Bea Cukai     │
│    (Transaksi pabean internal)│    REKONSILIASI    │    (Database CEISA & TPB)     │
└───────────────────────────────┘                     └───────────────────────────────┘
```

Rumus dasar: **Saldo Akhir Buku = Saldo Awal (Stock Opname) + Pemasukan − Pengeluaran**,
dihitung tiga kali (dari Mutasi IT, dari Per Dokumen IT, dan dari Data BC) untuk melihat
seberapa selaras ketiganya.

Detail lengkap aturan alur per kategori barang (Bahan Baku, WIP, Barang Jadi, Barang
Modal, Scrap) ada di `references/alur_kategori_barang.md`.

Pemetaan jenis dokumen BC ke tabel database ada di `references/pemetaan_tabel_bc.md`.

## 2. Input yang Dibutuhkan dari User

Minta user menyediakan (jika belum diunggah):

**A. Data IT Inventory (wajib, biasanya file `.xls`/`.xlsx` hasil export sistem inventory):**
- Laporan Mutasi Bahan Baku
- Laporan Mutasi Barang Jadi
- Laporan Mutasi Barang Sisa (Scrap)
- Laporan Mutasi Mesin/Barang Modal (jika ada aktivitas barang modal)
- Laporan Pemasukan Barang Per Dokumen
- Laporan Pengeluaran Barang Per Dokumen
- Laporan Hasil Stock Opname (saldo awal periode) — opsional tapi sangat disarankan

**B. Data Bea Cukai — sumbernya fleksibel, pilih salah satu:**
- **Akses database CEISA/TPB langsung** (jika tersedia) — lihat
  `references/pemetaan_tabel_bc.md` untuk nama tabel & kolom filter per jenis dokumen BC.
- **Upload Excel/CSV** yang disiapkan user sendiri (paling umum jika user tidak punya
  akses DB) — per jenis dokumen (BC 2.3, BC 4.0, BC 2.7, BC 2.5, BC 3.0, BC 4.1,
  BC 2.6.1/2.6.2), satu file per jenis ATAU satu file Excel gabungan dengan satu sheet
  per jenis dokumen. Kolom minimal yang dibutuhkan: `no_daftar`, `no_aju`, `tgl_daftar`,
  `kode_barang`, `uraian_barang`, `jumlah`, `satuan`, `nilai_idr`, `mitra`.

Kedua sumber ini **tidak harus konsisten di semua jenis dokumen** — misalnya BC 2.3 bisa
diambil dari DB sementara BC 2.6.1/2.6.2 (subkontrak) diupload manual via Excel kalau
tabelnya belum tersedia di DB. Sesuaikan `CONFIG["USE_DB"]` dan isi file per jenis BC
yang datanya belum ada di DB.

**C. Identitas & periode:**
- Nama perusahaan, NPWP, No. SKEP Kawasan Berikat (jika relevan)
- Rentang tanggal periode monitoring (tanggal mulai & akhir)

Jika salah satu data di atas kosong untuk kategori tertentu (misal perusahaan tidak
punya aktivitas Barang Modal), lewati kategori itu — jangan dipaksakan.

## 3. Langkah Eksekusi

1. **Susun konfigurasi** — isi `CONFIG` di `scripts/build_rekap_monsus.py` (nama
   perusahaan, kata kunci pencarian nama/NPWP, periode, path folder data, kredensial
   DB jika pakai akses langsung — jangan commit kredensial asli ke git, gunakan
   environment variable atau file `.env` lokal yang di-gitignore).
2. **Pastikan nama kolom cocok** dengan format laporan user. Nama kolom laporan
   mutasi/per-dokumen bisa sedikit berbeda antar sistem inventory — cek header file
   yang diunggah (`pd.read_html(...)[0].columns`) sebelum menjalankan script penuh,
   sesuaikan mapping kolom di bagian `LOAD MUTASI` / `LOAD PER DOKUMEN IT` jika perlu.
3. **Jalankan script** untuk menghasilkan file Excel rekonsiliasi:
   ```
   python scripts/build_rekap_monsus.py
   ```
4. **Periksa hasil** di sheet-sheet berikut (lihat `references/struktur_sheet_output.md`
   untuk detail kolom tiap sheet):
   - `Ringkasan Monsus` — ringkasan saldo per kategori + rekap dokumen masuk/keluar
   - `Analisis Alur & Saldo Buku` — rincian matematis per kategori barang
   - `Sandingan Dok Masuk` — detail per nomor dokumen pabean masuk & status match
   - `Sandingan Dok Keluar` — detail per nomor dokumen pabean keluar & status match
   - `Detail Saldo Per Item` — rekonsiliasi level item barang (opsional, untuk data besar)
5. **Soroti temuan** yang butuh perhatian: dokumen yang hanya ada di satu sisi
   (`Hanya di Dokumen IT` / `Hanya di Data BC`), selisih kuantitas signifikan, atau
   pengeluaran mutasi IT tanpa dokumen pabean pendamping (lihat catatan khusus Bahan
   Baku → WIP di bagian 4 referensi alur kategori).
6. **Sampaikan ringkasan ke user** dalam bentuk naratif singkat: persentase kesesuaian
   dokumen per jenis BC, saldo akhir per kategori (IT vs BC), dan item/dokumen anomali
   yang perlu ditindaklanjuti.

## 4. Hal yang Perlu Diperhatikan

- **Satuan tidak selalu sama** antar sumber (misal MTR di IT vs KGM di deklarasi BC
  untuk barang jadi yang dikonversi berat) — jangan langsung simpulkan "selisih" tanpa
  cek satuan dan faktor konversi dulu.
- **Bahan Baku → WIP** biasanya tidak punya dokumen pabean pengeluaran resmi (karena
  cuma perpindahan internal ke produksi). Jika "Per Dokumen IT" atau "Data BC" tercatat
  0 untuk pengeluaran kategori ini, samakan dengan angka Mutasi IT (bukan dianggap
  selisih/anomali).
- **Keselarasan internal** (Mutasi IT vs Per Dokumen IT) dan **keselarasan eksternal**
  (Per Dokumen IT vs Data BC) adalah dua pertanyaan berbeda — laporkan keduanya secara
  terpisah, jangan dicampur jadi satu angka.
- **Jangan masukkan kredensial database, nama perusahaan, NPWP, atau data klien nyata**
  ke dalam file apa pun di folder skill ini sebelum di-push ke repo publik. Gunakan
  `examples/` hanya untuk data contoh yang sudah di-anonimkan/fiktif.
- **Di luar lingkup**: analisis klaim barang terbakar/musnah pasca-insiden (butuh
  perbandingan Stock Opname pasca-insiden vs saldo buku sebelum insiden + pencacahan
  fisik forensik) — itu kasus khusus, tangani di luar skill ini.
