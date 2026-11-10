"""
Kasus Khusus: Audit Selisih Klaim Barang Terbakar/Musnah.

Dipakai SETELAH rekonsiliasi reguler (scripts/build_rekap_monsus.py) selesai, saat
perusahaan mengalami insiden (kebakaran/musnah) dan mengajukan klaim Stock Opname (SO)
pasca-insiden ke Bea Cukai.

Lihat penjelasan metodologi lengkap di: references/kasus_khusus_kebakaran.md

Input yang dibutuhkan (semua per kode barang):
  1. Saldo buku sebelum insiden, 3 versi (Mutasi IT / Per Dokumen IT / Data BC) --
     ambil dari hasil scripts/build_rekap_monsus.py (fungsi saldo_kategori / sheet
     "Detail Saldo Per Item" jika diaktifkan), diteruskan sebagai df_saldo_item.
  2. File "Laporan Hasil Stock Opname Pasca-Insiden" -- kolom: Kode Barang,
     Nama Barang, Satuan, Saldo Buku (SO Klaim), Fisik Selamat (Pencacahan).

Output: Excel dengan sheet "Detail Perbandingan Item SO" -- satu baris per item,
berisi saldo buku 3 versi sebelum insiden, klaim SO, fisik selamat, selisih, dan
status temuan.
"""

import os
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CONFIG = {
    "FOLDER_DATA": "data",
    "FILE_SALDO_SEBELUM_INSIDEN": "saldo_item_sebelum_insiden.xlsx",  # output dari rekonsiliasi reguler
    "FILE_SO_PASCA_INSIDEN": "laporan_so_pasca_insiden.xlsx",
    "OUTPUT_FILE": "Audit_Selisih_Klaim_Insiden.xlsx",
    "TOLERANSI_QTY": 0.01,       # selisih di bawah ini dianggap "selamat/sesuai", bukan temuan
    "LABEL_INSIDEN": "Insiden [isi: kebakaran/bencana lain] tanggal [isi tanggal]",
    "TANGGAL_SO_DIAJUKAN": "",
}

# Kolom wajib di FILE_SALDO_SEBELUM_INSIDEN:
#   kode_barang, nama_barang, kategori, satuan,
#   saldo_akhir_mutasi_it, saldo_akhir_per_dok_it, saldo_akhir_data_bc
# Kolom wajib di FILE_SO_PASCA_INSIDEN:
#   kode_barang, saldo_buku_so_klaim, fisik_selamat


def klasifikasi_status(selisih_vs_mutasi, toleransi):
    if pd.isna(selisih_vs_mutasi):
        return "Data Tidak Lengkap"
    if abs(selisih_vs_mutasi) <= toleransi:
        return "Terbakar Sesuai Mutasi IT"
    if selisih_vs_mutasi > toleransi:
        return f"INDIKASI PENGGELEMBUNGAN KLAIM TERBAKAR (Klaim Lebih +{selisih_vs_mutasi:,.2f})"
    return f"INDIKASI BARANG HILANG SEBELUM INSIDEN (Klaim Kurang {selisih_vs_mutasi:,.2f})"


def main():
    folder = CONFIG["FOLDER_DATA"]
    df_saldo = pd.read_excel(os.path.join(folder, CONFIG["FILE_SALDO_SEBELUM_INSIDEN"]))
    df_so = pd.read_excel(os.path.join(folder, CONFIG["FILE_SO_PASCA_INSIDEN"]))

    df_saldo["kode_barang"] = df_saldo["kode_barang"].astype(str).str.strip()
    df_so["kode_barang"] = df_so["kode_barang"].astype(str).str.strip()

    df = pd.merge(df_saldo, df_so, on="kode_barang", how="outer", suffixes=("", "_so"))

    for c in ["saldo_akhir_mutasi_it", "saldo_akhir_per_dok_it", "saldo_akhir_data_bc",
              "saldo_buku_so_klaim", "fisik_selamat"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["klaim_terbakar_so"] = df["fisik_selamat"] - df["saldo_buku_so_klaim"]
    df["estimasi_terbakar_mutasi_it"] = df["fisik_selamat"] - df["saldo_akhir_mutasi_it"]
    df["estimasi_terbakar_dok_it"] = df["fisik_selamat"] - df["saldo_akhir_per_dok_it"]
    df["estimasi_terbakar_data_bc"] = df["fisik_selamat"] - df["saldo_akhir_data_bc"]

    df["selisih_vs_mutasi_it"] = df["klaim_terbakar_so"] - df["estimasi_terbakar_mutasi_it"]
    # setara: df["saldo_buku_so_klaim"] - df["saldo_akhir_mutasi_it"], lihat referensi

    df["status"] = df["selisih_vs_mutasi_it"].apply(
        lambda v: klasifikasi_status(v, CONFIG["TOLERANSI_QTY"]))

    write_excel(df, folder)

    flagged = df[~df["status"].isin(["Terbakar Sesuai Mutasi IT", "Data Tidak Lengkap"])]
    print(f"Total item: {len(df):,} | Item dengan temuan (perlu ditelusuri Langkah 2): {len(flagged):,}")
    print(flagged["status"].value_counts())
    print("\nSimpan hasil ini (atau filter item flagged-nya) sebagai input "
          "FILE_ITEM_BERSELISIH untuk scripts/kasus_khusus_kebakaran/hitung_tagihan_pungutan.py")


def write_excel(df, folder):
    out_path = os.path.join(folder, CONFIG["OUTPUT_FILE"])
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Detail Perbandingan Item SO"

    font_family = "Segoe UI"
    title_font = Font(name=font_family, size=13, bold=True, color="1F4E79")
    header_font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    warn_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    warn_font = Font(name=font_family, size=10, bold=True, color="9C0006")
    regular_font = Font(name=font_family, size=10)
    thin = Side(style='thin', color='D9D9D9')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws['A1'] = f"AUDIT SELISIH KLAIM BARANG TERBAKAR/MUSNAH — {CONFIG['LABEL_INSIDEN']}"
    ws['A1'].font = title_font
    ws['A2'] = ("Selisih vs Mutasi IT > 0 = indikasi penggelembungan klaim; < 0 = indikasi barang "
                "hilang sebelum insiden. Item dengan temuan jadi input Langkah 2 (hitung tagihan pungutan).")
    ws['A2'].font = Font(name=font_family, size=9, italic=True, color="595959")

    headers = ["Kode Barang", "Nama Barang", "Kategori", "Satuan",
               "Saldo Akhir (Mutasi IT) Sebelum Insiden", "Saldo Akhir (Per Dok IT) Sebelum Insiden",
               "Saldo Akhir (Data BC) Sebelum Insiden",
               "Saldo Buku (SO Klaim)", "Fisik Selamat (Pencacahan)",
               "Klaim Terbakar (SO)", "Estimasi Terbakar (Mutasi IT)",
               "Estimasi Terbakar (Per Dok IT)", "Estimasi Terbakar (Data BC)",
               "Selisih vs Mutasi IT", "Status Temuan"]
    r0 = 4
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=r0, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    ws.row_dimensions[r0].height = 36

    cols_order = ["kode_barang", "nama_barang", "kategori", "satuan",
                  "saldo_akhir_mutasi_it", "saldo_akhir_per_dok_it", "saldo_akhir_data_bc",
                  "saldo_buku_so_klaim", "fisik_selamat",
                  "klaim_terbakar_so", "estimasi_terbakar_mutasi_it",
                  "estimasi_terbakar_dok_it", "estimasi_terbakar_data_bc",
                  "selisih_vs_mutasi_it", "status"]
    for i, row in df.iterrows():
        r = r0 + 1 + i
        is_flag = "INDIKASI" in str(row.get("status", ""))
        for c, colname in enumerate(cols_order, 1):
            val = row.get(colname, "")
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = border
            cell.font = warn_font if is_flag else regular_font
            if is_flag:
                cell.fill = warn_fill
            if c in range(5, 15):
                cell.alignment = Alignment(horizontal='right')
                cell.number_format = '#,##0.00'

    widths = [18, 32, 24, 10, 20, 20, 18, 16, 16, 16, 18, 18, 16, 16, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(out_path)
    print(f"Tersimpan: {out_path}")


if __name__ == "__main__":
    main()
