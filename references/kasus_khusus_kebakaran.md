# Kasus Khusus: Audit Klaim Barang Terbakar/Musnah + Tagihan atas Selisih

Modul tambahan ini dipakai **setelah** rekonsiliasi reguler (`scripts/build_rekap_monsus.py`)
selesai dan perusahaan mengalami insiden (kebakaran, bencana, dll.) yang menyebabkan
sebagian barang di Kawasan Berikat/TPB musnah, lalu perusahaan mengajukan klaim
berdasarkan Stock Opname (SO) pasca-insiden ke Bea Cukai.

Tujuan modul ini: **(1) mencari selisih** antara klaim perusahaan dan saldo buku riil,
dan **(2) menghitung estimasi tagihan pungutan** (Bea Masuk + PPN Impor + PPh Impor)
atas selisih tersebut.

## Kapan dipakai

- Hanya jika ada insiden barang musnah dan perusahaan mengajukan klaim SO ke Bea Cukai.
- Membutuhkan output rekonsiliasi reguler lebih dulu (saldo buku per item dari Mutasi IT,
  Per Dokumen IT, dan Data BC) sebagai baseline "saldo buku sebelum insiden".

## Langkah 1 — Audit Selisih Klaim (`scripts/kasus_khusus_kebakaran/audit_selisih_kebakaran.py`)

### Input tambahan yang dibutuhkan dari user
- **Laporan Hasil Stock Opname Pasca-Insiden**: `Saldo Buku (SO Klaim)` per item yang
  diajukan perusahaan ke Bea Cukai.
- **Hasil Pencacahan Fisik Bersama**: `Fisik Selamat` per item (idealnya dicacah bersama
  petugas Bea Cukai di lokasi insiden).
- Tanggal insiden & tanggal SO pasca-insiden diajukan (untuk label laporan).

### Logika perhitungan

Untuk setiap item barang:

```
Saldo Buku Sebelum Insiden  = dari rekonsiliasi reguler (3 versi: Mutasi IT, Per Dok IT, Data BC)
Klaim Terbakar (SO Klaim)   = Fisik Selamat − Saldo Buku (SO Klaim yang diajukan ke BC)
Estimasi Terbakar Riil      = Fisik Selamat − Saldo Buku Sebelum Insiden   (dihitung utk 3 versi)
Selisih vs Mutasi IT        = Klaim Terbakar (SO Klaim) − Saldo Buku Mutasi IT Sebelum Insiden
```

### Interpretasi `Selisih vs Mutasi IT`

- **Mendekati nol** → klaim perusahaan konsisten dengan buku riil, tidak ada temuan.
- **Positif signifikan (Klaim > Saldo Buku Riil)** → **indikasi penggelembungan klaim
  barang terbakar**: ada kemungkinan barang sebenarnya sudah keluar sebelum insiden
  (terjual, dipindah, dll.) tapi ikut dimasukkan dalam klaim barang musnah. Ini jadi
  dasar temuan `INDIKASI PENGGELEMBUNGAN KLAIM TERBAKAR`.
- **Negatif signifikan (Klaim < Saldo Buku Riil)** → **indikasi barang hilang sebelum
  insiden**: saldo buku riil sebelum insiden sudah lebih tinggi dari yang diklaim,
  berarti ada barang yang "menghilang" dari pencatatan sebelum insiden terjadi, bukan
  karena insiden itu sendiri.

### Status hasil per item

- `Barang Selamat / Utuh (Fisik = Buku Mutasi)`
- `Terbakar Sesuai Mutasi IT` (selisih kecil, dalam toleransi)
- `INDIKASI PENGGELEMBUNGAN KLAIM TERBAKAR (Klaim Lebih +XX)`
- `INDIKASI BARANG HILANG SEBELUM KEBAKARAN`

Output: sheet Excel `Detail Perbandingan Item SO` berisi semua item dengan kolom
identitas, saldo buku 3 versi, klaim SO, fisik selamat, selisih, dan status.

## Langkah 2 — Hitung Tagihan Pungutan atas Selisih (`scripts/kasus_khusus_kebakaran/hitung_tagihan_pungutan.py`)

Untuk **item yang berstatus selisih** (indikasi penggelembungan klaim atau barang hilang
sebelum insiden) dari Langkah 1, hitung estimasi tagihan pungutan yang terutang — karena
barang tersebut secara fisik sudah keluar dari Kawasan Berikat/TPB tanpa penyelesaian
kewajiban pabean (ekspor resmi, pembayaran BM/PPN/PPh, dll.), sehingga terutang sebagai
pengeluaran ke TLDDP tanpa dokumen.

### Dasar nilai barang (basis pungutan) — berbeda per kategori

| Kategori | Dasar Nilai Barang | Sumber Data |
|:---|:---|:---|
| **Barang Modal (Mesin & Parts)** | Nilai Pabean (CIF Rupiah) rata-rata per unit dari deklarasi BC 2.3/BC 4.0 riil milik perusahaan, dikali selisih qty | Query/lookup ke tabel `bc_23` / `db_40` (atau file lookup `nilai_pabean_per_unit.csv` jika tidak ada akses DB), kolom: kode barang, HS Code, Jumlah Satuan, Cif Rupiah/Harga Penyerahan, Tarif BM/PPN/PPh, Nilai Fasilitas |
| **Barang Jadi & WIP (berbasis komposisi bahan)** | Nilai bahan baku (misal yarn/benang untuk produk tekstil) per KG dari resep komposisi (Bill of Materials/BOM), dikali selisih qty yang sudah dikonversi ke KG | File BOM/resep komposisi yang disiapkan user (per kode barang: daftar bahan baku penyusun, berat KG per baris, nilai Rupiah per baris, HS Code & tarif BM/PPN/PPh bahan baku tersebut) |

> **Penting — konversi satuan**: jika kategori barang punya satuan berbeda dari basis
> nilai (misal Barang Jadi dalam MTR tapi nilai bahan dalam KG), **jangan** percaya
> begitu saja kolom "faktor konversi" yang sudah ada di file sumber tanpa verifikasi —
> turunkan sendiri faktor konversi dari dua kolom yang berasal dari baris data yang sama
> (misal Total KG komposisi ÷ Qty klaim pada baris itu), karena kolom konversi bawaan
> sering tidak konsisten satuannya antar item. Jika sudah sama-sama KG (misal WIP),
> faktor konversi = 1 (tidak perlu dikonversi).

### Rumus pungutan

```
Nilai Barang (Rp)  = Nilai per Unit/KG  ×  Selisih Qty (setelah konversi satuan jika perlu)
BM (Rp)             = Nilai Barang  ×  Tarif BM%
PPN Impor (Rp)      = (Nilai Barang + BM)  ×  Tarif PPN%
PPh Impor (Rp)      = (Nilai Barang + BM)  ×  Tarif PPh%
Total Tagihan (Rp)  = BM + PPN Impor + PPh Impor
```

Tarif efektif untuk Barang Jadi/WIP (yang bisa punya campuran beberapa HS Code bahan
baku) dihitung sebagai tarif blended: `total BM bahan baku di BOM / total nilai bahan
baku di BOM`, begitu juga untuk PPN dan PPh (basis PPN/PPh = Nilai Barang + BM).

### Output

Sheet Excel `Tagihan Pungutan Selisih` — satu baris per item berselisih, dengan kolom
formula Excel (bukan angka statis) untuk BM/PPN/PPh/Total, supaya bisa ditelusuri
langsung dari sel Excel-nya. Sertakan kolom "Sumber & Cara Memperoleh Nilai" dan
"Sumber & Cara Memperoleh Angka Konversi" secara eksplisit per baris — ini penting untuk
keperluan audit/pembuktian ke Bea Cukai, jangan hanya menampilkan angka hasil akhir.

Item yang kategorinya tidak masuk dua basis di atas, atau tidak ditemukan datanya di
sumber lookup, ditandai statusnya `Tidak Tersedia` dengan catatan alasannya — jangan
dipaksakan dihitung dengan asumsi kosong/nol.

### Catatan Scrap

Kategori Barang Sisa (Scrap) biasanya **tidak** disertakan dalam penghitungan tagihan
ini kecuali diminta lain oleh user — nilai scrap relatif kecil dan sulit divaluasi
secara andal dengan metode di atas. Konfirmasi ke user sebelum mengecualikan.
