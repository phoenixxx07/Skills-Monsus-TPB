# Aturan Alur Barang per Kategori

Setiap kategori barang punya aturan sumber saldo awal dan arah pemasukan/pengeluaran
yang berbeda. Pahami ini sebelum mengklasifikasikan dokumen sebagai "match" atau
"selisih" — banyak "selisih" yang sebenarnya cuma perbedaan alur pencatatan, bukan
kesalahan.

| No | Kategori Barang | Sumber Saldo Awal | Alur Pemasukan (+) | Alur Pengeluaran (−) | Catatan & Asumsi |
|:---|:---|:---|:---|:---|:---|
| 1 | **Bahan Baku & Penolong** | Hasil Stock Opname (SO) Awal Periode | Dokumen Pemasukan IT & Data BC (BC 2.3, BC 4.0, BC 2.7 Masuk) | Asumsi ditransfer ke proses produksi (**WIP**), kecuali ada penjualan lokal (BC 2.5) kategori bahan baku atau pengeluaran kembali (BC 4.1) | Jika tidak ada BC 2.5/4.1 kategori bahan baku, seluruh pengeluaran mutasi dianggap konsumsi produksi (tidak perlu dokumen pabean pendamping). |
| 2 | **Barang Dalam Proses (WIP)** | Hasil Stock Opname (SO) Awal Periode | Ditransfer dari pemakaian Bahan Baku | Dikeluarkan setelah proses menjadi Barang Jadi dan Scrap | WIP adalah jembatan konversi internal; tidak pernah punya dokumen pabean sendiri. |
| 3 | **Barang Jadi (Finished Goods)** | Hasil Stock Opname (SO) Awal Periode | Ditransfer dari hasil produksi WIP (pemasukan di mutasi Barang Jadi) | Dokumen IT & Data BC (BC 3.0 Ekspor, BC 2.7 Antar-KB, BC 2.5 Jual Lokal) | Produk siap jual/ekspor. |
| 4 | **Barang Modal (Mesin & Parts)** | Hasil Stock Opname (SO) Awal Periode | Dokumen Pemasukan IT & Data BC (BC 2.3 & BC 4.0 kategori mesin/spare part, BC 2.6.2 masuk subkon) | Dokumen Pengeluaran IT & Data BC (BC 2.6.1 keluar subkon, BC 4.1 pengembalian kemasan/mesin) | Pengeluaran di mutasi IT sering mencatat pemakaian suku cadang internal yang tidak selalu punya dokumen pabean. |
| 5 | **Barang Sisa (Scrap / Limbah)** | Hasil Stock Opname (SO) Awal Periode | Timbulan dari proses produksi di lini WIP (pemasukan di mutasi Scrap) | Dokumen Pengeluaran IT & Data BC (BC 2.5 penjualan scrap lokal, BC 4.1 limbah/kemasan) | Satuan umumnya KG/KGM. |

## Perlakuan Khusus: Bahan Baku → WIP (tidak ada dokumen pabean)

Karena pengeluaran bahan baku ke lantai produksi (WIP) **tidak dibuatkan dokumen
pabean pengeluaran resmi**, saat menghitung saldo akhir buku versi "Per Dokumen IT"
dan versi "Data BC":

1. Jika pada kolom pengeluaran "Per Dokumen IT" tercatat `0.00`, samakan angkanya
   dengan angka pengeluaran "Mutasi IT".
2. Jika pada kolom pengeluaran "Data BC" tercatat `0.00`, samakan juga angkanya
   dengan angka pengeluaran "Mutasi IT".

Tanpa perlakuan ini, saldo akhir versi dokumen/BC akan meledak positif secara tidak
realistis (karena hanya pemasukan yang tercatat, pengeluaran dianggap nol).

## Dua Jenis Pertanyaan Keselarasan — Jangan Dicampur

1. **Keselarasan Eksternal** (yang jadi fokus utama Monsus): *Per Dokumen IT vs Data
   BC*. Ini mengukur apakah dokumen pabean yang dicatat perusahaan cocok dengan
   database resmi Bea Cukai.
2. **Keselarasan Internal**: *Mutasi IT (buku gudang/produksi) vs Per Dokumen IT
   (rekap dokumen pabean)*. Ini mengukur apakah pencatatan gudang internal konsisten
   dengan dokumen pabean — selisih di sini **bukan otomatis masalah compliance**,
   bisa jadi karena mutasi internal mencakup transfer antar-gudang, pemakaian
   maintenance, sample/reject yang belum didokumentasikan, dsb. Tandai sebagai
   "catatan audit" untuk ditelusuri lebih lanjut, bukan langsung dianggap pelanggaran.

## Status Rekonsiliasi per Item (opsional, untuk sheet detail per item)

- `Match Sesuai (BC & IT)` — Mutasi, Dokumen IT, dan Data BC selaras sempurna.
- `Match BC (Selisih Mutasi Internal)` — Dokumen IT dan Data BC cocok 100%, tapi
  mutasi internal mencatat proses pabrik tambahan (transfer gudang, dsb).
- `Selisih Data BC vs IT` — ada perbedaan kuantitas antara dokumen pabean perusahaan
  dengan data resmi Bea Cukai — ini yang paling perlu ditindaklanjuti.
- `Selisih Pabean & Mutasi` — perbedaan terjadi di kedua level sekaligus.
