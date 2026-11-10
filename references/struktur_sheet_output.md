# Struktur Sheet Output Excel

File Excel hasil rekonsiliasi berisi sheet-sheet berikut (sheet investigasi kebakaran
dari pola aslinya **tidak** disertakan di sini — itu di luar lingkup skill ini).

## Sheet 1: `Ringkasan Monsus`

- **Tabel 1 — Komparasi Saldo Akhir per Kategori**: Saldo Awal SO, Saldo Akhir
  (Mutasi IT / Per Dokumen IT / Data BC), Selisih, % Keselarasan, catatan evaluasi.
- **Tabel 2 — Evaluasi Keselarasan Internal**: Masuk/Keluar (Mutasi IT vs Per
  Dokumen IT) per kategori, status, catatan audit.
- **Tabel 3 — Rincian 3 Pilar Data**: gabungan Masuk & Keluar (Mutasi, Per Dokumen,
  Data BC) per kategori dalam satu baris, lengkap dengan saldo akhir tiap versi.
- **Tabel 4 — Rekapitulasi Dokumen Pemasukan**: per jenis BC masuk — jumlah dokumen
  (IT vs BC), jumlah item, total kuantitas, % kesesuaian dokumen, status evaluasi.
- **Tabel 5 — Rekapitulasi Dokumen Pengeluaran**: sama seperti Tabel 4, untuk arah
  keluar.

## Sheet 2: `Analisis Alur & Saldo Buku`

Rincian matematis per kategori barang, baris demi baris komponen pembentuk saldo:

```
(   ) Saldo Awal (Stock Opname)
(+) Pemasukan [sumber & rincian per jenis dokumen]
(-) Pengeluaran [sumber & rincian per jenis dokumen]
(=) Saldo Akhir Buku
```

dihitung untuk kolom "Data IT (Mutasi/Dok IT)" vs "Data Bea Cukai", dengan kolom
selisih dan catatan metode/asumsi di setiap baris.

## Sheet 3: `Sandingan Dok Masuk`

Satu baris per nomor dokumen pabean masuk (hasil `outer merge` antara rekap dokumen
IT dan rekap dokumen Data BC berdasarkan nomor dokumen), kolom:

- Nomor dokumen, jenis BC, tanggal, nama mitra
- Jumlah item (IT) & (BC), total kuantitas (IT) & (BC), selisih kuantitas
- Status: `Match Sesuai` / `Selisih Kuantitas` / `Hanya di Dokumen IT` / `Hanya di Data BC`

## Sheet 4: `Sandingan Dok Keluar`

Struktur sama seperti Sheet 3, untuk dokumen pabean arah keluar.

## Sheet 5: `Detail Saldo Per Item` (opsional — aktifkan jika data per-item dibutuhkan)

Satu baris per kombinasi kode barang unik, menyerupai tampilan Stock Opname
komparatif:

1. **Identitas**: No, Kategori Barang, Kode Barang, Nama/Uraian Barang, Satuan.
2. **Saldo Awal**: dari Laporan Hasil Stock Opname periode sebelumnya.
3. **Pemasukan**: kolom `Mutasi IT`, `Per Dokumen IT`, `Data Dokumen BC`.
4. **Pengeluaran**: kolom `Mutasi IT`, `Per Dokumen IT`, `Data Dokumen BC` (ingat
   perlakuan khusus Bahan Baku→WIP, lihat `alur_kategori_barang.md`).
5. **Saldo Akhir Buku**: dihitung dinamis `= Saldo Awal + Masuk - Keluar` untuk
   masing-masing dari 3 versi data.
6. **Selisih**: `Selisih (Dok IT - Data BC)` dan `Selisih (Mutasi IT - Dok IT)`.
7. **Status Rekonsiliasi**: lihat daftar status di `alur_kategori_barang.md`.

Sheet ini bisa jadi sangat besar (ribuan baris) untuk perusahaan dengan ragam item
banyak — pertimbangkan hanya menampilkan item dengan selisih di atas threshold
tertentu jika file jadi terlalu besar untuk dibuka user.

## Konvensi Styling (opsional, konsisten dengan contoh di `scripts/`)

- Header: bold putih di atas latar navy (`#1F4E79`) untuk tabel utama, biru pudar
  (`#2E75B6`) untuk sub-tabel.
- Baris selang-seling (zebra striping) warna abu sangat muda (`#F2F4F7`).
- Baris highlight perhatian/anomali: latar oranye pudar (`#FCE4D6`), font merah tua
  (`#9C0006`).
- Baris total: latar biru pudar (`#D9E1F2`) dengan border atas/bawah tegas.
