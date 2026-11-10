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
  Mencakup juga modul kasus khusus insiden kebakaran/barang musnah: audit selisih
  klaim Stock Opname pasca-insiden dan penghitungan estimasi tagihan pungutan (BM/PPN/PPh
  Impor) atas selisihnya — lihat references/kasus_khusus_kebakaran.md.
---

# Monsus (Monitoring Khusus) Kawasan Berikat — Rekonsiliasi IT Inventory vs Data Bea Cukai

Skill ini membungkus pola kerja standar untuk mengaudit/merekonsiliasi data barang
perusahaan pemilik fasilitas Kawasan Berikat / TPB (Tempat Penimbunan Berikat) antara:

1. **IT Inventory perusahaan** (buku mutasi barang + laporan per dokumen pabean), dan
2. **Data resmi deklarasi Bea Cukai** (database CEISA/TPB).

Cakupan utama skill ini adalah rekonsiliasi **umum** (saldo, dokumen, kuantitas).
Jika perusahaan mengalami insiden (kebakaran/barang musnah) dan mengajukan klaim
Stock Opname pasca-insiden ke Bea Cukai, gunakan modul tambahan **Kasus Khusus
Kebakaran** (bagian 5 di bawah) untuk mencari selisih klaim dan menghitung estimasi
tagihan pungutan atas selisih tersebut.

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

## 2. Data Minimal (Checklist Wajib Sebelum Mulai)

Sebelum mulai mengerjakan monsus, pastikan kelompok data A–D di bawah ini sudah
tersedia. Kalau belum lengkap, **minta dulu ke user** — jangan mulai menghitung dari
data yang kosong, parsial, atau diasumsikan.

**A. IT Inventory — Laporan tiap jenis barang, Mutasi Masuk & Keluar (wajib, biasanya
file `.xls`/`.xlsx` hasil export sistem inventory):**
- Mutasi **Bahan Baku & Penolong**
- Mutasi **Barang Jadi** (Finished Goods)
- Mutasi **Barang Sisa** (Scrap/Limbah)
- Mutasi **Barang Modal** (Mesin & Parts) — jika ada aktivitas barang modal di periode itu
- Laporan Pemasukan Barang Per Dokumen
- Laporan Pengeluaran Barang Per Dokumen
- Laporan Hasil Stock Opname (saldo awal tiap kategori di atas untuk periode yang
  dianalisis) — sangat disarankan, dibutuhkan untuk menghitung saldo akhir buku

**B. Data WIP (Barang Dalam Proses)** — sering terlewat karena banyak sistem inventory
tidak punya laporan mutasi WIP yang berdiri sendiri seperti kategori lain. Minimal
butuh salah satu dari:
- Laporan Mutasi WIP tersendiri (kalau sistem user punya), ATAU
- Kombinasi: **Saldo Awal WIP** (dari Stock Opname) + **rekap pemakaian bahan baku ke
  produksi** (masuk WIP) + **rekap hasil produksi** yang ditransfer menjadi Barang Jadi
  dan Scrap (keluar WIP), per periode monitoring yang sama.

Tanpa data WIP, saldo akhir Barang Jadi & Scrap tidak bisa divalidasi penuh (lihat
alur kategori di `references/alur_kategori_barang.md`) — jangan asumsikan saldo WIP
nol atau konstan tanpa konfirmasi ke user.

**C. Data Bea Cukai (wajib, sumbernya fleksibel, pilih salah satu):**
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

**D. Identitas & periode (wajib):**
- Nama perusahaan, NPWP, No. SKEP Kawasan Berikat (jika relevan)
- Rentang tanggal periode monitoring (tanggal mulai & akhir)

Jika salah satu data di atas kosong untuk kategori tertentu karena memang tidak relevan
(misal perusahaan tidak punya aktivitas Barang Modal sama sekali di periode itu), lewati
kategori itu — ini beda dengan data yang relevan tapi belum diupload, yang harus diminta
dulu ke user (lihat aturan di bawah).

### Kalau permintaan user lebih spesifik dari data minimal di atas

Kalau di tengah analisis user minta sesuatu yang lebih detail/spesifik dari data minimal
yang sudah tersedia (misal: rincian per item barang, breakdown subkontrak BC 2.6.1/2.6.2
yang belum diupload, data BOM/komposisi untuk kasus kebakaran di bagian 5, filter per
mitra/pemasok tertentu, dsb.) dan **data pendukungnya belum ada**:

- **Jangan menebak, mengekstrapolasi, atau mengisi dengan asumsi/nol** supaya analisis
  tetap "jalan".
- **Hentikan dan sampaikan ke user secara spesifik** data apa yang kurang (nama
  file/kolom yang dibutuhkan), lalu minta mereka mengunggah/menyediakannya dulu.
- Lanjutkan analisis permintaan spesifik itu setelah data tersebut tersedia. Bagian
  rekonsiliasi umum yang datanya sudah lengkap tetap bisa dikerjakan/dilaporkan duluan.

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

## 5. Kasus Khusus: Insiden Kebakaran / Barang Musnah

Jika user menyebut ada insiden (kebakaran, bencana, dll.) dan klaim Stock Opname
pasca-insiden ke Bea Cukai, jalankan dua langkah tambahan setelah rekonsiliasi reguler
di atas selesai (lihat detail metodologi & rumus di `references/kasus_khusus_kebakaran.md`):

1. **Audit selisih klaim** —
   `scripts/kasus_khusus_kebakaran/audit_selisih_kebakaran.py`. Membandingkan saldo
   buku sebelum insiden (hasil rekonsiliasi reguler, 3 versi: Mutasi IT/Per Dok IT/Data
   BC) dengan klaim SO pasca-insiden dan hasil pencacahan fisik bersama. Output: sheet
   `Detail Perbandingan Item SO` dengan status per item (`Terbakar Sesuai Mutasi IT`,
   `INDIKASI PENGGELEMBUNGAN KLAIM TERBAKAR`, `INDIKASI BARANG HILANG SEBELUM INSIDEN`).
2. **Hitung tagihan pungutan atas selisih** —
   `scripts/kasus_khusus_kebakaran/hitung_tagihan_pungutan.py`. Untuk item yang
   berstatus "INDIKASI..." dari langkah 1, hitung estimasi BM + PPN Impor + PPh Impor
   yang terutang atas kuantitas selisihnya — dasar nilai barang beda per kategori
   (Barang Modal: nilai pabean per unit dari BC 2.3/4.0; Barang Jadi/WIP: nilai bahan
   baku per KG dari BOM/resep komposisi). Output: sheet `Tagihan Pungutan Selisih`
   dengan kolom formula Excel yang bisa ditelusuri + rincian sumber nilai per baris.

Sampaikan ke user dua angka utama dari modul ini: **total kuantitas selisih per
kategori barang** dan **total estimasi tagihan pungutan (Rp)**, plus item mana saja
yang paling signifikan kontribusinya ke total tagihan.
