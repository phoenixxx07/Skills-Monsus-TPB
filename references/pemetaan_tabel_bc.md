# Pemetaan Jenis Dokumen BC ke Tabel Database (contoh struktur CEISA/TPB)

> Nama host, nama database, dan kredensial di tabel ini **harus diisi ulang sesuai
> environment masing-masing** — jangan hardcode IP/password server internal ke dalam
> repo publik. Gunakan environment variable (lihat `scripts/build_rekap_monsus.py`,
> bagian `DB_CONFIG`).

| Jenis Dokumen | Fungsi Transaksi | Nama Tabel (contoh) | Kolom Filter Perusahaan | Kolom Tanggal | Arah |
|:---|:---|:---|:---|:---|:---|
| **BC 2.3** | Impor dari Luar Daerah Pabean (LDP) | `bc_23` | `Nama Pengusaha`, `Id Pengusaha` | `Tanggal Daftar` | Masuk |
| **BC 4.0** | Pemasukan Lokal dari TLDDP ke TPB | `db_40` | `Nama Pengusaha`, `Id Pengusaha` | `Tanggal Daftar` | Masuk |
| **BC 2.7 Masuk** | Pemasukan Antar-TPB / Kawasan Berikat | `db_27_penerima` | `Nama Penerima`, `Id Penerima` | `Tanggal Daftar` | Masuk |
| **BC 2.6.2** | Pemasukan Kembali Hasil Subkontrak | `subkon_dokumen` (arah masuk) | `nama_pengusaha`, `id_pengusaha` | `tanggal_daftar` | Masuk |
| **BC 2.5** | Pengeluaran dari TPB ke TLDDP (Lokal/Scrap) | `db_25_barang` | `Nama Pengusaha`, `Id Pengusaha` | `Tanggal Daftar` | Keluar |
| **BC 2.7 Keluar** | Pengeluaran Antar-TPB / Kawasan Berikat | `db_27_pengusaha` | `Nama Pengusaha`, `Id Pengusaha` | `Tanggal Daftar` | Keluar |
| **BC 3.0** | Ekspor Barang Jadi ke Luar Daerah Pabean | `db_30` | `Nama Eksportir`, `Npwp Eksportir` | `Tanggal Peb` | Keluar |
| **BC 4.1** | Pengeluaran Barang/Kemasan Bekas ke TLDDP | `db_41` | `Nama Pengusaha`, `Id Pengusaha` | `Tanggal Daftar` | Keluar |
| **BC 2.6.1** | Pengeluaran untuk Subkontrak ke TLDDP | `subkon_dokumen` (arah keluar) | `nama_pengusaha`, `id_pengusaha` | `tanggal_daftar` | Keluar |

## Catatan Query

- Filter perusahaan sebaiknya pakai kombinasi `LIKE` nama (pakai `%NAMA%`) **DAN**
  `LIKE` NPWP/ID pengusaha yang sudah dinormalisasi (hilangkan titik, strip, spasi),
  supaya tidak kehilangan data karena variasi penulisan nama perusahaan.
- Filter tanggal pakai kolom tanggal dokumen (`Tanggal Daftar` / `Tanggal Peb`),
  bukan tanggal entry/input sistem.
- **Sumber data BC bersifat fleksibel** — tidak harus dari database. Jika user tidak
  punya/tidak mau memberi akses DB langsung, minta mereka upload data per jenis dokumen
  sebagai file Excel (`.xlsx`/`.xls`) atau CSV, dengan kolom minimal: `no_daftar`,
  `no_aju`, `tgl_daftar`, `kode_barang`, `uraian_barang`, `jumlah`, `satuan`,
  `nilai_idr` (opsional), `mitra` (importir/eksportir/pemasok/penerima). Ini bisa berupa
  satu file per jenis BC, atau satu file Excel dengan satu sheet per jenis BC. Campuran
  juga boleh: sebagian jenis dokumen dari DB, sebagian dari file upload, tergantung mana
  yang tersedia untuk kasus tersebut.
- Nomor dokumen (`no_daftar`/`no_peb`) sebaiknya dinormalisasi (misal `zfill` ke 6
  digit, buang suffix desimal) sebelum dipakai sebagai kunci join dengan data IT,
  karena format penulisan nomor dokumen antar sistem sering berbeda (misal
  `123456.0` vs `000123`).

## Jika ada jenis dokumen BC lain (BC 2.8, BC 4.2, dst.)

Tambahkan baris baru di tabel ini saat menemui jenis dokumen lain, lalu tambahkan
query-nya di `scripts/build_rekap_monsus.py` dengan pola yang sama (SELECT kolom
standar: jenis_bc, no_daftar, no_aju, tgl_daftar, kode_barang, uraian_barang, jumlah,
satuan, nilai_idr, mitra).
