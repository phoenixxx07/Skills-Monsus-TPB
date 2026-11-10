"""
Template reusable untuk Monitoring Khusus (Monsus) Kawasan Berikat / TPB.
Merekonsiliasi IT Inventory perusahaan (Mutasi Barang + Per Dokumen) dengan
Data Dokumen Bea Cukai (CEISA/TPB).

Cara pakai:
1. Isi CONFIG di bawah (nama perusahaan, periode, path folder data).
2. Pilih sumber data Bea Cukai — fleksibel, dua opsi:
   a. Akses DB langsung: CONFIG["USE_DB"] = True, isi kredensial lewat environment
      variable (jangan hardcode):
        set MONSUS_DB_HOST_MASTER=...   (PowerShell: $env:MONSUS_DB_HOST_MASTER="...")
        set MONSUS_DB_HOST_TPB=...
        set MONSUS_DB_USER=...
        set MONSUS_DB_PASSWORD=...
   b. Upload Excel/CSV manual (tanpa akses DB): CONFIG["USE_DB"] = False, lalu siapkan
      salah satu dari:
        - satu file per jenis BC: export_BC_23.xlsx (atau .xls/.csv), export_BC_40.xlsx, dst.
        - ATAU satu file Excel gabungan dengan satu sheet per jenis BC, namanya diisi
          di CONFIG["BC_COMBINED_FILE"], nama sheet = nama jenis BC (mis. "BC 2.3").
      Kolom wajib tiap file/sheet: jenis_bc, no_daftar, no_aju, tgl_daftar, kode_barang,
      uraian_barang, jumlah, satuan, nilai_idr, mitra (lihat load_bc_from_files()).
3. Jalankan: python scripts/build_rekap_monsus.py

Tidak mencakup analisis kasus khusus (misal klaim barang terbakar) — itu di luar
lingkup template ini.
"""

import os
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIG — ganti sesuai perusahaan yang dianalisis
# ==============================================================================
CONFIG = {
    "NAMA_PERUSAHAAN": "PT. CONTOH PERUSAHAAN",
    "NAMA_SEARCH_KEY": "%CONTOH%",            # LIKE pattern untuk nama di DB
    "NPWP": "00.000.000.0-000.000",
    "NPWP_SEARCH_KEY": "%000000000000%",      # LIKE pattern NPWP/ID ternormalisasi
    "SKEP_KB": "",                             # nomor SKEP Kawasan Berikat (jika ada)
    "JENIS_FASILITAS": "Kawasan Berikat",
    "PERIODE_START": "2026-01-01",
    "PERIODE_END": "2026-12-31",
    "PERIODE_LABEL": "1 Januari 2026 s.d. 31 Desember 2026",
    "FOLDER_DATA": "data",                     # folder berisi file IT Inventory
    "OUTPUT_FILE": "Rekap_Penyandingan_Monsus.xlsx",
    "USE_DB": True,                            # False jika data BC disediakan via file upload (Excel/CSV)
    "BC_COMBINED_FILE": "",                    # opsional: nama file Excel gabungan (1 sheet per jenis BC) di FOLDER_DATA, dipakai saat USE_DB=False
    "SERTAKAN_SHEET_DETAIL_PER_ITEM": False,   # True jika butuh rincian per item barang
}

DB_CONFIG = {
    "HOST_MASTER": os.environ.get("MONSUS_DB_HOST_MASTER", ""),
    "HOST_TPB": os.environ.get("MONSUS_DB_HOST_TPB", ""),
    "USER": os.environ.get("MONSUS_DB_USER", ""),
    "PASSWORD": os.environ.get("MONSUS_DB_PASSWORD", ""),
    "PORT": int(os.environ.get("MONSUS_DB_PORT", "3306")),
    "DB_MASTER": os.environ.get("MONSUS_DB_NAME_MASTER", "hs_code"),
    "DB_TPB": os.environ.get("MONSUS_DB_NAME_TPB", "db_tpb"),
}

folder = CONFIG["FOLDER_DATA"]
output_excel = os.path.join(folder, CONFIG["OUTPUT_FILE"])

print(f"Starting Monsus reconciliation for {CONFIG['NAMA_PERUSAHAAN']} ({CONFIG['PERIODE_LABEL']})...")

MUTASI_NUMERIC_COLS = ['Saldo Awal', 'Pemasukan', 'Pengeluaran', 'Penyesuaian', 'Saldo Akhir', 'Stock Opname']

# Kategori barang: nama file mutasi -> label kategori di laporan
MUTASI_FILES = {
    "Bahan Baku & Penolong": "LapKPBC-Mutasi-BahanBaku.xls",
    "Barang Dalam Proses (WIP)": "LapKPBC-Mutasi-WIP.xls",
    "Barang Jadi (Finished Goods)": "LapKPBC-Mutasi-BarangJadi.xls",
    "Barang Modal (Mesin & Parts)": "LapKPBC-Mutasi-Mesin.xls",
    "Barang Sisa (Scrap / Limbah)": "LapKPBC-Mutasi-BarangSisa.xls",
}

PER_DOKUMEN_IN_FILE = "LAPKPPBCPemasukanBarang Per Dokumen.xls"
PER_DOKUMEN_OUT_FILE = "LAPKPPBCPengeluaranBarang Per Dokumen.xls"


def load_mutasi_file(path):
    """Load satu file mutasi, buang baris TOTAL, numerik-kan kolom standar."""
    if not os.path.exists(path):
        return None
    df_raw = pd.read_html(path)[0]
    df = df_raw[df_raw['Kode Brg'] != 'TOTAL'].copy()
    for c in MUTASI_NUMERIC_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
    return df


def load_all_mutasi():
    """Return dict kategori -> dataframe mutasi (kategori dilewati jika file tidak ada)."""
    result = {}
    for kategori, fname in MUTASI_FILES.items():
        df = load_mutasi_file(os.path.join(folder, fname))
        if df is not None:
            result[kategori] = df
        else:
            print(f"  [skip] {fname} tidak ditemukan, kategori '{kategori}' dilewati.")
    return result


def load_per_dokumen():
    """Load laporan per dokumen masuk & keluar dari IT Inventory."""
    df_in = pd.read_html(os.path.join(folder, PER_DOKUMEN_IN_FILE))[0]
    df_out = pd.read_html(os.path.join(folder, PER_DOKUMEN_OUT_FILE))[0]
    df_in['Jumlah_num'] = pd.to_numeric(df_in['Jumlah'], errors='coerce').fillna(0)
    df_out['Jumlah_num'] = pd.to_numeric(df_out['Jumlah'], errors='coerce').fillna(0)
    df_in['bc_no_clean'] = df_in['BCNo'].astype(str).str.split('.').str[0].str.zfill(6)
    df_out['bc_no_clean'] = df_out['BC No'].astype(str).str.split('.').str[0].str.zfill(6)
    df_in['Tanggal BC'] = pd.to_datetime(df_in['Tanggal BC'], errors='coerce')
    df_out['Tanggal BC'] = pd.to_datetime(df_out['Tanggal BC'], errors='coerce')
    return df_in, df_out


# ==============================================================================
# QUERY BC — satu dict per jenis dokumen: (sql_template, host_key, db_key)
# Sesuaikan nama tabel/kolom dengan referensi: references/pemetaan_tabel_bc.md
# ==============================================================================
BC_QUERIES = {
    "BC 2.3": dict(host="HOST_MASTER", db="DB_MASTER", table="bc_23",
        cols=dict(no_daftar="`Nomor Daftar`", no_aju="`Nomor Aju`", tgl="`Tanggal Daftar`",
                  kode="`Kode Barang`", uraian="`Uraian Barang`", jumlah="`Jumlah Satuan`",
                  satuan="`Kode Satuan`", nilai="`Cif Rupiah`", mitra="`Nama Pemasok`"),
        filter_nama="`Nama Pengusaha`", filter_id="`Id Pengusaha`", tgl_filter="`Tanggal Daftar`", arah="masuk"),
    "BC 4.0": dict(host="HOST_TPB", db="DB_TPB", table="db_40",
        cols=dict(no_daftar="`Nomor Daftar`", no_aju="`Nomor Aju`", tgl="`Tanggal Daftar`",
                  kode="`Kode Barang`", uraian="`Uraian Barang`", jumlah="`Jumlah Satuan`",
                  satuan="`Kode Satuan`", nilai="`Harga Penyerahan Brg`", mitra="`Nama Pengirim`"),
        filter_nama="`Nama Pengusaha`", filter_id="`Id Pengusaha`", tgl_filter="`Tanggal Daftar`", arah="masuk"),
    "BC 2.7 Masuk": dict(host="HOST_TPB", db="DB_TPB", table="db_27_penerima",
        cols=dict(no_daftar="`Nomor Daftar`", no_aju="`Nomor Aju`", tgl="`Tanggal Daftar`",
                  kode="`Kode Barang`", uraian="`Uraian Barang`", jumlah="`Jumlah Satuan`",
                  satuan="`Kode Satuan`", nilai="`Harga Penyerahan Brg`", mitra="`Nama Pengusaha`"),
        filter_nama="`Nama Penerima`", filter_id="`Id Penerima`", tgl_filter="`Tanggal Daftar`", arah="masuk"),
    "BC 2.5": dict(host="HOST_TPB", db="DB_TPB", table="db_25_barang",
        cols=dict(no_daftar="`Nomor Daftar`", no_aju="`Nomor Aju`", tgl="`Tanggal Daftar`",
                  kode="`Kode Barang`", uraian="`Uraian Barang`", jumlah="`Jumlah Satuan`",
                  satuan="`Kode Satuan`", nilai="`Harga Penyerahan Brg`", mitra="`Nama Penerima`"),
        filter_nama="`Nama Pengusaha`", filter_id="`Id Pengusaha`", tgl_filter="`Tanggal Daftar`", arah="keluar"),
    "BC 2.7 Keluar": dict(host="HOST_TPB", db="DB_TPB", table="db_27_pengusaha",
        cols=dict(no_daftar="`Nomor Daftar`", no_aju="`Nomor Aju`", tgl="`Tanggal Daftar`",
                  kode="`Kode Barang`", uraian="`Uraian Barang`", jumlah="`Jumlah Satuan`",
                  satuan="`Kode Satuan`", nilai="`Harga Penyerahan Brg`", mitra="`Nama Penerima`"),
        filter_nama="`Nama Pengusaha`", filter_id="`Id Pengusaha`", tgl_filter="`Tanggal Daftar`", arah="keluar"),
    "BC 3.0": dict(host="HOST_TPB", db="DB_TPB", table="db_30",
        cols=dict(no_daftar="`Nomor Peb`", no_aju="`No Aju`", tgl="`Tanggal Peb`",
                  kode="`KODE_BARANG`", uraian="`Uraian Barang`", jumlah="`Jumlah Satuan`",
                  satuan="`Jenis Satuan`", nilai="`Nilai Dev Idr Brg`", mitra="`Nama Pembeli`"),
        filter_nama="`Nama Eksportir`", filter_id="`Npwp Eksportir`", tgl_filter="`Tanggal Peb`", arah="keluar"),
    "BC 4.1": dict(host="HOST_TPB", db="DB_TPB", table="db_41",
        cols=dict(no_daftar="`Nomor Daftar`", no_aju="`Nomor Aju`", tgl="`Tanggal Daftar`",
                  kode="`Kode Barang`", uraian="`Uraian Barang`", jumlah="`Jumlah Satuan`",
                  satuan="`Kode Satuan`", nilai="`Harga Penyerahan Brg`", mitra="`Nama Penerima`"),
        filter_nama="`Nama Pengusaha`", filter_id="`Id Pengusaha`", tgl_filter="`Tanggal Daftar`", arah="keluar"),
    # BC 2.6.1 / 2.6.2 (subkontrak): sesuaikan nama tabel & kolom dengan sistem masing-masing.
}


def build_query(jenis_bc, spec):
    c = spec["cols"]
    return f"""
    SELECT
        '{jenis_bc}' as jenis_bc,
        LPAD({c['no_daftar']}, 6, '0') as no_daftar,
        {c['no_aju']} as no_aju,
        {c['tgl']} as tgl_daftar,
        {c['kode']} as kode_barang,
        {c['uraian']} as uraian_barang,
        {c['jumlah']} as jumlah,
        {c['satuan']} as satuan,
        {c['nilai']} as nilai_idr,
        {c['mitra']} as mitra
    FROM {spec['table']}
    WHERE ({spec['filter_nama']} LIKE '{CONFIG["NAMA_SEARCH_KEY"]}' OR {spec['filter_id']} LIKE '{CONFIG["NPWP_SEARCH_KEY"]}')
      AND {spec['tgl_filter']} BETWEEN '{CONFIG["PERIODE_START"]}' AND '{CONFIG["PERIODE_END"]}'
    """


def load_bc_from_db():
    import pymysql
    conns = {}
    for key in ("HOST_MASTER", "HOST_TPB"):
        host = DB_CONFIG[key]
        if not host:
            continue
    conn_master = pymysql.connect(host=DB_CONFIG["HOST_MASTER"], user=DB_CONFIG["USER"],
                                   password=DB_CONFIG["PASSWORD"], database=DB_CONFIG["DB_MASTER"],
                                   port=DB_CONFIG["PORT"])
    conn_tpb = pymysql.connect(host=DB_CONFIG["HOST_TPB"], user=DB_CONFIG["USER"],
                                password=DB_CONFIG["PASSWORD"], database=DB_CONFIG["DB_TPB"],
                                port=DB_CONFIG["PORT"])
    conns = {"HOST_MASTER": conn_master, "HOST_TPB": conn_tpb}

    results = {}
    for jenis_bc, spec in BC_QUERIES.items():
        conn = conns[spec["host"]]
        sql = build_query(jenis_bc, spec)
        results[jenis_bc] = pd.read_sql(sql, conn)
        print(f"  Loaded {jenis_bc}: {len(results[jenis_bc])} baris")

    for conn in conns.values():
        conn.close()
    return results


BC_FILE_SCHEMA_COLS = ["jenis_bc", "no_daftar", "no_aju", "tgl_daftar", "kode_barang",
                        "uraian_barang", "jumlah", "satuan", "nilai_idr", "mitra"]


def _read_bc_table(path):
    """Baca satu file data BC, format csv/xlsx/xls, kolom wajib: lihat BC_FILE_SCHEMA_COLS."""
    if path.lower().endswith(".csv"):
        df = pd.read_csv(path, dtype={"no_daftar": str})
    else:
        df = pd.read_excel(path, dtype={"no_daftar": str})
    missing = [c for c in BC_FILE_SCHEMA_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"File {path} tidak punya kolom wajib: {missing}. "
                          f"Kolom yang dibutuhkan: {BC_FILE_SCHEMA_COLS}")
    return df


def load_bc_from_files():
    """Fallback/alternatif jika tidak ada akses DB langsung: data Bea Cukai disediakan
    user sebagai file upload (Excel .xlsx/.xls atau .csv), satu file per jenis dokumen BC,
    ATAU satu file Excel dengan satu sheet per jenis dokumen BC (nama sheet = nama jenis BC,
    misal "BC 2.3", "BC 4.0", dst).

    Pola pencarian file per jenis BC, dicoba berurutan:
      1. export_<JENIS_BC>.xlsx / .xls / .csv   (mis. export_BC_2.3.xlsx)
      2. sheet bernama <JENIS_BC> di BC_COMBINED_FILE (jika CONFIG mengarah ke file gabungan)
    """
    results = {}
    combined_path = os.path.join(folder, CONFIG.get("BC_COMBINED_FILE", "")) if CONFIG.get("BC_COMBINED_FILE") else None
    combined_sheets = {}
    if combined_path and os.path.exists(combined_path):
        combined_sheets = pd.read_excel(combined_path, sheet_name=None, dtype={"no_daftar": str})

    for jenis_bc in BC_QUERIES.keys():
        loaded = False
        for ext in (".xlsx", ".xls", ".csv"):
            fname = os.path.join(folder, f"export_{jenis_bc.replace(' ', '_').replace('.', '')}{ext}")
            if os.path.exists(fname):
                results[jenis_bc] = _read_bc_table(fname)
                print(f"  Loaded {jenis_bc} dari file: {len(results[jenis_bc])} baris ({fname})")
                loaded = True
                break
        if not loaded and jenis_bc in combined_sheets:
            df = combined_sheets[jenis_bc]
            missing = [c for c in BC_FILE_SCHEMA_COLS if c not in df.columns]
            if missing:
                raise ValueError(f"Sheet '{jenis_bc}' di {combined_path} tidak punya kolom wajib: {missing}")
            results[jenis_bc] = df
            print(f"  Loaded {jenis_bc} dari sheet gabungan: {len(df)} baris")
            loaded = True
        if not loaded:
            print(f"  [skip] Tidak ditemukan sumber data untuk {jenis_bc} "
                  f"(cek export_{jenis_bc.replace(' ', '_')}.xlsx/.xls/.csv atau sheet di BC_COMBINED_FILE)")
            results[jenis_bc] = pd.DataFrame(columns=BC_FILE_SCHEMA_COLS)
    return results


ARAH_MASUK = [k for k, v in BC_QUERIES.items() if v["arah"] == "masuk"]
ARAH_KELUAR = [k for k, v in BC_QUERIES.items() if v["arah"] == "keluar"]


def aggregate_per_dokumen(df, group_col, item_col, mitra_col):
    return df.groupby(group_col).agg(
        jenis_bc=('Jenis BC', 'first'),
        tgl_bc=('Tanggal BC', 'first'),
        mitra=(mitra_col, 'first'),
        jml_item=(item_col, 'count'),
        total_qty=('Jumlah_num', 'sum'),
    ).reset_index()


def aggregate_bc(df_dict, jenis_list):
    frames = [df_dict[j] for j in jenis_list if j in df_dict and not df_dict[j].empty]
    if not frames:
        return pd.DataFrame(columns=["no_daftar", "jenis_bc_db", "tgl_db", "mitra_db",
                                      "no_aju_db", "jml_item_db", "total_qty_db", "total_nilai_db"])
    df_all = pd.concat(frames, ignore_index=True)
    return df_all.groupby('no_daftar').agg(
        jenis_bc_db=('jenis_bc', 'first'),
        tgl_db=('tgl_daftar', 'first'),
        mitra_db=('mitra', 'first'),
        no_aju_db=('no_aju', 'first'),
        jml_item_db=('kode_barang', 'count'),
        total_qty_db=('jumlah', 'sum'),
        total_nilai_db=('nilai_idr', 'sum'),
    ).reset_index()


def merge_dok(agg_it, agg_db):
    merged = pd.merge(agg_it, agg_db, left_on='bc_no_clean', right_on='no_daftar', how='outer')
    merged['no_daftar_final'] = merged['bc_no_clean'].fillna(merged['no_daftar'])
    merged['jenis_bc_final'] = merged['jenis_bc'].fillna(merged['jenis_bc_db'])
    merged['tgl_final'] = merged['tgl_bc'].fillna(merged['tgl_db'])
    merged['mitra_final'] = merged['mitra'].fillna(merged['mitra_db'])
    merged['total_qty_it'] = merged['total_qty'].fillna(0)
    merged['total_qty_db'] = merged['total_qty_db'].fillna(0)
    merged['selisih_qty'] = merged['total_qty_it'] - merged['total_qty_db']
    merged['status'] = np.where(
        merged['total_qty'].isna(), 'Hanya di Data BC',
        np.where(merged['total_qty_db'] == 0, 'Hanya di Dokumen IT',
        np.where(np.abs(merged['selisih_qty']) < 0.01, 'Match Sesuai', 'Selisih Kuantitas'))
    )
    return merged


def saldo_kategori(mutasi_dfs, df_in, df_out, df_bc_in, df_bc_out):
    """Hitung saldo awal/pemasukan/pengeluaran/akhir per kategori untuk 3 versi data:
    Mutasi IT, Per Dokumen IT, Data BC — berdasarkan kode barang yang tercatat di
    file mutasi tiap kategori."""
    rows = []
    for kategori, df_mut in mutasi_dfs.items():
        kode_set = set(df_mut['Kode Brg'].astype(str).str.strip())
        satuan = df_mut['Satuan'].iloc[0] if 'Satuan' in df_mut.columns and len(df_mut) else ""

        saldo_awal = df_mut['Stock Opname'].sum() if 'Stock Opname' in df_mut.columns else df_mut['Saldo Awal'].sum()
        masuk_mutasi = df_mut['Pemasukan'].sum()
        keluar_mutasi = df_mut['Pengeluaran'].sum()
        saldo_akhir_mutasi = df_mut['Saldo Akhir'].sum() if 'Saldo Akhir' in df_mut.columns else (saldo_awal + masuk_mutasi - keluar_mutasi)

        masuk_dok_it = df_in[df_in['Kode Brg'].astype(str).str.strip().isin(kode_set)]['Jumlah_num'].sum() if df_in is not None else 0.0
        keluar_dok_it = df_out[df_out['Kode Brg'].astype(str).str.strip().isin(kode_set)]['Jumlah_num'].sum() if df_out is not None else 0.0

        masuk_bc = df_bc_in[df_bc_in['kode_barang'].astype(str).str.strip().isin(kode_set)]['jumlah'].sum() if not df_bc_in.empty else 0.0
        keluar_bc = df_bc_out[df_bc_out['kode_barang'].astype(str).str.strip().isin(kode_set)]['jumlah'].sum() if not df_bc_out.empty else 0.0

        # Perlakuan khusus: pengeluaran tanpa dokumen pabean (misal Bahan Baku -> WIP)
        if keluar_dok_it == 0:
            keluar_dok_it = keluar_mutasi
        if keluar_bc == 0:
            keluar_bc = keluar_mutasi

        saldo_akhir_dok_it = saldo_awal + masuk_dok_it - keluar_dok_it
        saldo_akhir_bc = saldo_awal + masuk_bc - keluar_bc

        rows.append(dict(
            kategori=kategori, satuan=satuan, saldo_awal=saldo_awal,
            masuk_mutasi=masuk_mutasi, keluar_mutasi=keluar_mutasi, saldo_akhir_mutasi=saldo_akhir_mutasi,
            masuk_dok_it=masuk_dok_it, keluar_dok_it=keluar_dok_it, saldo_akhir_dok_it=saldo_akhir_dok_it,
            masuk_bc=masuk_bc, keluar_bc=keluar_bc, saldo_akhir_bc=saldo_akhir_bc,
            selisih_akhir=saldo_akhir_dok_it - saldo_akhir_bc,
        ))
    return pd.DataFrame(rows)


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    os.makedirs(folder, exist_ok=True)

    print("1. Loading data mutasi IT Inventory...")
    mutasi_dfs = load_all_mutasi()

    print("2. Loading data per dokumen IT Inventory...")
    df_in, df_out = load_per_dokumen()

    print("3. Loading data Bea Cukai...")
    bc_data = load_bc_from_db() if CONFIG["USE_DB"] else load_bc_from_files()

    print("4. Menyandingkan dokumen masuk & keluar...")
    agg_in_it = aggregate_per_dokumen(df_in, 'bc_no_clean', 'Kode Brg', 'Pemasok')
    agg_out_it = aggregate_per_dokumen(df_out, 'bc_no_clean', 'Kode Brg', 'Customer')
    agg_in_bc = aggregate_bc(bc_data, ARAH_MASUK)
    agg_out_bc = aggregate_bc(bc_data, ARAH_KELUAR)
    merged_in = merge_dok(agg_in_it, agg_in_bc)
    merged_out = merge_dok(agg_out_it, agg_out_bc)

    print("Status dokumen masuk:")
    print(merged_in['status'].value_counts())
    print("Status dokumen keluar:")
    print(merged_out['status'].value_counts())

    print("5. Menghitung saldo per kategori...")
    df_bc_in_all = pd.concat([bc_data[j] for j in ARAH_MASUK if j in bc_data and not bc_data[j].empty], ignore_index=True) if any(j in bc_data for j in ARAH_MASUK) else pd.DataFrame(columns=["kode_barang", "jumlah"])
    df_bc_out_all = pd.concat([bc_data[j] for j in ARAH_KELUAR if j in bc_data and not bc_data[j].empty], ignore_index=True) if any(j in bc_data for j in ARAH_KELUAR) else pd.DataFrame(columns=["kode_barang", "jumlah"])
    df_saldo = saldo_kategori(mutasi_dfs, df_in, df_out, df_bc_in_all, df_bc_out_all)

    print("6. Menyusun file Excel...")
    write_excel(df_saldo, merged_in, merged_out)

    print(f"Selesai. File tersimpan di: {output_excel}")


# ==============================================================================
# EXCEL OUTPUT
# ==============================================================================
def write_excel(df_saldo, merged_in, merged_out):
    wb = openpyxl.Workbook()
    font_family = "Segoe UI"
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    title_font = Font(name=font_family, size=14, bold=True, color="1F4E79")
    bold_font = Font(name=font_family, size=10, bold=True)
    regular_font = Font(name=font_family, size=10)
    italic_font = Font(name=font_family, size=9, italic=True, color="595959")
    zebra_fill = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    thin_border = Border(*[Side(style='thin', color='D9D9D9')] * 4)
    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    def write_table(ws, start_row, title, headers, rows, number_cols=None):
        number_cols = number_cols or []
        ws.cell(row=start_row, column=1, value=title).font = Font(name=font_family, size=11, bold=True, color="1F4E79")
        r = start_row + 1
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=r, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center
            cell.border = thin_border
        for row_data in rows:
            r += 1
            for col_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=r, column=col_idx, value=val)
                cell.font = regular_font
                cell.border = thin_border
                if col_idx in number_cols:
                    cell.alignment = align_right
                    cell.number_format = '#,##0.00'
            if r % 2 == 1:
                for c in range(1, len(headers) + 1):
                    ws.cell(row=r, column=c).fill = zebra_fill
        return r + 2

    # Sheet 1: Ringkasan Monsus
    ws1 = wb.active
    ws1.title = "Ringkasan Monsus"
    ws1['A1'] = "LAPORAN PENYANDINGAN DATA MONITORING KHUSUS (MONSUS)"
    ws1['A1'].font = title_font
    ws1['A2'] = f"{CONFIG['NAMA_PERUSAHAAN']} ({CONFIG['JENIS_FASILITAS']})"
    ws1['A2'].font = Font(name=font_family, size=12, bold=True, color="2E75B6")
    ws1['A3'] = f"NPWP: {CONFIG['NPWP']} | SKEP: {CONFIG['SKEP_KB']} | PERIODE: {CONFIG['PERIODE_LABEL']}"
    ws1['A3'].font = italic_font

    headers_saldo = ["Kategori Barang", "Satuan", "Saldo Awal", "Saldo Akhir (Mutasi IT)",
                      "Saldo Akhir (Per Dok IT)", "Saldo Akhir (Data BC)", "Selisih (Dok IT - BC)"]
    rows_saldo = [[r.kategori, r.satuan, r.saldo_awal, r.saldo_akhir_mutasi,
                   r.saldo_akhir_dok_it, r.saldo_akhir_bc, r.selisih_akhir]
                  for r in df_saldo.itertuples()]
    next_row = write_table(ws1, 5, "TABEL 1: KOMPARASI SALDO AKHIR PER KATEGORI (DATA IT vs DATA BEA CUKAI)",
                            headers_saldo, rows_saldo, number_cols=[3, 4, 5, 6, 7])

    def rekap_dokumen(merged, label):
        total_dok_it = merged['total_qty'].notna().sum()
        total_dok_bc = merged['total_qty_db'].ne(0).sum()
        match = (merged['status'] == 'Match Sesuai').sum()
        return [["TOTAL " + label, total_dok_it, total_dok_bc, match,
                  f"{match / len(merged) * 100:.1f}%" if len(merged) else "0%"]]

    next_row = write_table(ws1, next_row, "TABEL 2: REKAP DOKUMEN MASUK (IT vs BC)",
                            ["Ringkasan", "Jml Dok IT", "Jml Dok BC", "Jml Match", "% Match"],
                            rekap_dokumen(merged_in, "PEMASUKAN"))
    write_table(ws1, next_row, "TABEL 3: REKAP DOKUMEN KELUAR (IT vs BC)",
                ["Ringkasan", "Jml Dok IT", "Jml Dok BC", "Jml Match", "% Match"],
                rekap_dokumen(merged_out, "PENGELUARAN"))

    for col in ws1.columns:
        max_len = max((len(str(c.value or '')) for c in col), default=10)
        ws1.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 3, 12), 45)

    # Sheet 2: Sandingan Dok Masuk
    ws2 = wb.create_sheet("Sandingan Dok Masuk")
    headers_dok = ["No Dokumen", "Jenis BC", "Tanggal", "Mitra", "Jml Item (IT)", "Jml Item (BC)",
                    "Qty (IT)", "Qty (BC)", "Selisih Qty", "Status"]
    rows_dok = [[r.no_daftar_final, r.jenis_bc_final, r.tgl_final, r.mitra_final,
                 r.jml_item if pd.notna(r.jml_item) else 0, r.jml_item_db if pd.notna(r.jml_item_db) else 0,
                 r.total_qty_it, r.total_qty_db, r.selisih_qty, r.status]
                for r in merged_in.itertuples()]
    write_table(ws2, 1, "SANDINGAN DOKUMEN PEMASUKAN (IT vs DATA BC)", headers_dok, rows_dok,
                number_cols=[5, 6, 7, 8, 9])

    # Sheet 3: Sandingan Dok Keluar
    ws3 = wb.create_sheet("Sandingan Dok Keluar")
    rows_dok_out = [[r.no_daftar_final, r.jenis_bc_final, r.tgl_final, r.mitra_final,
                      r.jml_item if pd.notna(r.jml_item) else 0, r.jml_item_db if pd.notna(r.jml_item_db) else 0,
                      r.total_qty_it, r.total_qty_db, r.selisih_qty, r.status]
                     for r in merged_out.itertuples()]
    write_table(ws3, 1, "SANDINGAN DOKUMEN PENGELUARAN (IT vs DATA BC)", headers_dok, rows_dok_out,
                number_cols=[5, 6, 7, 8, 9])

    for ws in (ws2, ws3):
        for col in ws.columns:
            max_len = max((len(str(c.value or '')) for c in col), default=10)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 3, 12), 40)

    wb.save(output_excel)


if __name__ == "__main__":
    main()
