"""
Kasus Khusus: Hitung Estimasi Tagihan Pungutan (BM + PPN Impor + PPh Impor)
atas Item Berselisih (hasil Langkah 1: audit_selisih_kebakaran.py).

Lihat metodologi lengkap di: references/kasus_khusus_kebakaran.md

Dasar nilai barang dua jalur, pilih sesuai kategori item:
  A. Barang Modal (mesin/parts) -> Nilai Pabean (CIF) per unit dari deklarasi BC 2.3/BC 4.0
  B. Barang Jadi / WIP (berbasis komposisi bahan baku) -> Nilai bahan baku per KG dari
     file resep/BOM (Bill of Materials), + faktor konversi satuan yang DITURUNKAN sendiri
     dari data BOM (bukan dipercaya mentah dari kolom konversi bawaan file sumber).

Input:
  1. FILE_ITEM_BERSELISIH: hasil Langkah 1 yang sudah difilter ke item berstatus
     "INDIKASI ..." saja (kolom minimal: kode_barang, nama_barang, kategori, satuan,
     selisih_vs_mutasi_it, status).
  2. FILE_NILAI_PABEAN_BARANG_MODAL (opsional, untuk kategori Barang Modal): kode_barang,
     hs_code, qty_total, cif_rp_total, bm_tarif_pct, ppn_tarif_pct, pph_tarif_pct
     -> bisa diisi dari query manual ke tabel bc_23/db_40, atau dari export Excel/CSV.
  3. FILE_BOM_BARANG_JADI / FILE_BOM_WIP (opsional, untuk kategori berbasis komposisi):
     satu baris per (kode_barang, baris_bahan_baku): kode_barang, qty_klaim_basis,
     komposisi_kg, nilai_bahan_rp, hs_code, bm_tarif_pct, ppn_tarif_pct, pph_tarif_pct.
     Nilai & tarif harus SUDAH dalam Rupiah (konversi kurs dilakukan sebelum file ini
     disiapkan, atau tambahkan kolom kurs/rate dan konversi di load_bom()).

Output: Excel "Tagihan Pungutan Selisih.xlsx", sheet "Tagihan Pungutan Selisih",
kolom hasil (BM/PPN/PPh/Total) ditulis sebagai FORMULA Excel agar bisa ditelusuri.
"""

import os
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CONFIG = {
    "FOLDER_DATA": "data",
    "FILE_ITEM_BERSELISIH": "item_berselisih.xlsx",
    "FILE_NILAI_PABEAN_BARANG_MODAL": "nilai_pabean_barang_modal.xlsx",
    "FILE_BOM_BARANG_JADI": "bom_barang_jadi.xlsx",
    "FILE_BOM_WIP": "bom_wip.xlsx",
    "KATEGORI_BARANG_MODAL": "Barang Modal (Mesin & Parts)",
    "KATEGORI_KOMPOSISI_BAHAN": ["Barang Jadi (Finished Goods)", "Barang Dalam Proses (WIP)"],
    "KATEGORI_DIKECUALIKAN": ["Barang Sisa (Scrap / Limbah)"],  # konfirmasi ke user sebelum eksekusi
    "OUTPUT_FILE": "Tagihan_Pungutan_Selisih.xlsx",
}


def load_nilai_pabean_barang_modal(folder):
    path = os.path.join(folder, CONFIG["FILE_NILAI_PABEAN_BARANG_MODAL"])
    if not os.path.exists(path):
        return {}
    df = pd.read_excel(path, dtype={"kode_barang": str})
    df["nilai_pabean_per_unit"] = df["cif_rp_total"] / df["qty_total"].replace(0, np.nan)
    return df.set_index("kode_barang").to_dict("index")


def load_bom(folder, fname):
    """Agregasi file BOM (satu baris per bahan baku penyusun) jadi satu entri per
    kode_barang: total komposisi KG, total nilai bahan (Rp), HS code gabungan, dan
    tarif efektif (blended) BM/PPN/PPh."""
    path = os.path.join(folder, fname)
    if not os.path.exists(path):
        return {}
    df = pd.read_excel(path, dtype={"kode_barang": str})
    out = {}
    for kode, g in df.groupby("kode_barang"):
        comp_kg = g["komposisi_kg"].sum()
        nilai_rp = g["nilai_bahan_rp"].sum()
        bm_rp = (g["nilai_bahan_rp"] * g["bm_tarif_pct"] / 100.0).sum()
        basis_ppn = nilai_rp + bm_rp
        ppn_rp = (g["nilai_bahan_rp"] * g.get("ppn_tarif_pct", 0) / 100.0).sum()
        pph_rp = (g["nilai_bahan_rp"] * g.get("pph_tarif_pct", 0) / 100.0).sum()
        qty_klaim_basis = g["qty_klaim_basis"].iloc[0] if "qty_klaim_basis" in g.columns else None
        out[kode] = dict(
            comp_kg=comp_kg, nilai_rp=nilai_rp, bm_rp=bm_rp, ppn_rp=ppn_rp, pph_rp=pph_rp,
            hs_set=sorted(set(g["hs_code"].dropna().astype(str))),
            qty_klaim_basis=qty_klaim_basis,
            bm_tarif_pct=(bm_rp / nilai_rp * 100.0) if nilai_rp else 0.0,
            ppn_tarif_pct=(ppn_rp / basis_ppn * 100.0) if basis_ppn else 0.0,
            pph_tarif_pct=(pph_rp / basis_ppn * 100.0) if basis_ppn else 0.0,
        )
    return out


def hitung_barang_modal(kode, selisih, lookup):
    info = lookup.get(str(kode).strip())
    if not info or not info.get("qty_total") or pd.isna(info.get("nilai_pabean_per_unit")):
        return None
    nilai_barang = info["nilai_pabean_per_unit"] * abs(selisih)
    return dict(
        metode="BC 2.3/4.0: Nilai Pabean (CIF Rp) per unit x Selisih Qty",
        hs_code=info.get("hs_code", ""),
        nilai_barang_rp=nilai_barang,
        bm_tarif_pct=info.get("bm_tarif_pct", 0.0),
        ppn_tarif_pct=info.get("ppn_tarif_pct", 0.0),
        pph_tarif_pct=info.get("pph_tarif_pct", 0.0),
        konversi_factor=1.0, konversi_satuan="-",
        sumber_nilai=(f"Nilai Pabean/unit Rp{info['nilai_pabean_per_unit']:,.2f} "
                      f"(CIF total Rp{info['cif_rp_total']:,.0f} / qty total {info['qty_total']:,.2f}) "
                      f"x selisih {abs(selisih):,.2f} = Rp{nilai_barang:,.2f}"),
    )


def hitung_komposisi_bahan(kode, selisih, satuan, bom_lookup, perlu_konversi_kg):
    rec = bom_lookup.get(str(kode).strip())
    if not rec or rec["comp_kg"] <= 0 or rec["nilai_rp"] <= 0:
        return None
    if perlu_konversi_kg:
        if not rec.get("qty_klaim_basis") or rec["qty_klaim_basis"] <= 0:
            return None
        konv = rec["comp_kg"] / rec["qty_klaim_basis"]
        selisih_kg = abs(selisih) * konv
        konv_satuan = f"KG/{satuan}"
        sumber_konv = (f"Komposisi KG total {rec['comp_kg']:,.4f} / Qty basis klaim BOM "
                       f"{rec['qty_klaim_basis']:,.2f} {satuan} = {konv:.6f} KG/{satuan}")
    else:
        konv = 1.0
        selisih_kg = abs(selisih)
        konv_satuan = "KG/KG"
        sumber_konv = "Satuan sudah KG, tidak perlu konversi"

    nilai_per_kg = rec["nilai_rp"] / rec["comp_kg"]
    nilai_barang = nilai_per_kg * selisih_kg
    return dict(
        metode="BOM: Nilai Bahan Baku per KG x Selisih (dikonversi KG dari data BOM)",
        hs_code=", ".join(rec["hs_set"]),
        nilai_barang_rp=nilai_barang,
        bm_tarif_pct=rec["bm_tarif_pct"], ppn_tarif_pct=rec["ppn_tarif_pct"], pph_tarif_pct=rec["pph_tarif_pct"],
        konversi_factor=konv, konversi_satuan=konv_satuan,
        sumber_nilai=(f"Nilai bahan/KG Rp{nilai_per_kg:,.2f} (total nilai bahan Rp{rec['nilai_rp']:,.2f} "
                      f"/ komposisi {rec['comp_kg']:,.4f} KG) x {selisih_kg:,.4f} KG "
                      f"({sumber_konv}) = Rp{nilai_barang:,.2f}"),
    )


def main():
    folder = CONFIG["FOLDER_DATA"]
    df_flag = pd.read_excel(os.path.join(folder, CONFIG["FILE_ITEM_BERSELISIH"]), dtype={"kode_barang": str})
    df_flag = df_flag[~df_flag["kategori"].isin(CONFIG["KATEGORI_DIKECUALIKAN"])].copy()

    lookup_modal = load_nilai_pabean_barang_modal(folder)
    bom_bj = load_bom(folder, CONFIG["FILE_BOM_BARANG_JADI"])
    bom_wip = load_bom(folder, CONFIG["FILE_BOM_WIP"])

    results = []
    for _, item in df_flag.iterrows():
        kat = item["kategori"]
        selisih = float(item.get("selisih_vs_mutasi_it", 0) or 0)
        base = dict(kode=item["kode_barang"], nama=item.get("nama_barang", ""),
                    kategori=kat, satuan=item.get("satuan", ""), selisih_qty=selisih,
                    status=item.get("status", ""))

        info = None
        if kat == CONFIG["KATEGORI_BARANG_MODAL"]:
            info = hitung_barang_modal(item["kode_barang"], selisih, lookup_modal)
        elif kat in CONFIG["KATEGORI_KOMPOSISI_BAHAN"]:
            is_wip = "WIP" in kat or "Dalam Proses" in kat
            bom_lookup = bom_wip if is_wip else bom_bj
            info = hitung_komposisi_bahan(item["kode_barang"], selisih, item.get("satuan", ""),
                                            bom_lookup, perlu_konversi_kg=not is_wip)

        if info is None:
            base.update(dict(metode="Tidak Tersedia", hs_code="", nilai_barang_rp=0.0,
                              bm_tarif_pct=0.0, ppn_tarif_pct=0.0, pph_tarif_pct=0.0,
                              konversi_factor=None, konversi_satuan="-",
                              sumber_nilai="Data lookup (nilai pabean / BOM) tidak ditemukan untuk kode ini"))
        else:
            base.update(info)
        results.append(base)

    df_res = pd.DataFrame(results)
    write_excel(df_res, folder)

    tersedia = df_res[df_res["metode"] != "Tidak Tersedia"]
    print(f"Total item dihitung: {len(df_res):,} | Tersedia datanya: {len(tersedia):,}")
    print(tersedia.groupby("kategori")["nilai_barang_rp"].sum() if len(tersedia) else "Tidak ada data.")


def write_excel(df_res, folder):
    out_path = os.path.join(folder, CONFIG["OUTPUT_FILE"])
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tagihan Pungutan Selisih"

    font_family = "Segoe UI"
    title_font = Font(name=font_family, size=12, bold=True, color="1F4E79")
    header_font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    na_font = Font(name=font_family, size=9, italic=True, color="808080")
    na_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    regular_font = Font(name=font_family, size=10)
    thin = Side(style='thin', color='D9D9D9')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws['A1'] = "ESTIMASI TAGIHAN PUNGUTAN (BM + PPN IMPOR + PPH IMPOR) ATAS ITEM BERSELISIH"
    ws['A1'].font = title_font
    ws['A2'] = ("Kolom BM/PPN/PPh/Total adalah FORMULA Excel yang merujuk kolom lain di baris yang "
                "sama -- klik selnya untuk melihat rumus & nilai sumber. Item 'Tidak Tersedia' berarti "
                "data lookup (nilai pabean/BOM) belum ditemukan untuk kode tersebut.")
    ws['A2'].font = Font(name=font_family, size=9, italic=True, color="595959")
    ws.row_dimensions[2].height = 30

    headers = ["No", "Kategori", "Kode Barang", "Nama Barang", "Satuan", "Selisih Qty",
               "Metode Valuasi", "HS Code", "Angka Konversi", "Satuan Konversi",
               "Nilai Barang/Pabean - Dasar Pungutan (Rp)", "Sumber & Rincian Nilai Barang",
               "Tarif BM (%)", "Tarif PPN (%)", "Tarif PPh (%)",
               "Estimasi BM (Rp)", "Estimasi PPN (Rp)", "Estimasi PPh (Rp)",
               "Total Estimasi Pungutan (Rp)", "Status Audit"]
    r0 = 4
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=r0, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    ws.row_dimensions[r0].height = 40

    df_res = df_res.sort_values(["kategori", "nilai_barang_rp"], ascending=[True, False]).reset_index(drop=True)
    for i, r in df_res.iterrows():
        rr = r0 + 1 + i
        vals = [i + 1, r["kategori"], r["kode"], r["nama"], r["satuan"], r["selisih_qty"],
                r["metode"], r["hs_code"], r["konversi_factor"], r["konversi_satuan"],
                r["nilai_barang_rp"], r["sumber_nilai"],
                r["bm_tarif_pct"], r["ppn_tarif_pct"], r["pph_tarif_pct"],
                f"=K{rr}*M{rr}/100", f"=(K{rr}+P{rr})*N{rr}/100", f"=(K{rr}+P{rr})*O{rr}/100",
                f"=P{rr}+Q{rr}+R{rr}", r["status"]]
        ws.append(vals)
        is_na = r["metode"] == "Tidak Tersedia"
        for c in range(1, 21):
            cell = ws.cell(row=rr, column=c)
            cell.border = border
            cell.font = na_font if is_na else regular_font
            if is_na:
                cell.fill = na_fill
            if c in [6, 9, 11, 13, 14, 15, 16, 17, 18, 19]:
                cell.alignment = Alignment(horizontal='right')
                cell.number_format = '#,##0.0000' if c == 9 else '#,##0.00'

    tot_row = r0 + 1 + len(df_res)
    ws.cell(row=tot_row, column=2, value="TOTAL").font = Font(name=font_family, bold=True)
    for c, colname in [(6, "selisih_qty"), (11, "nilai_barang_rp")]:
        cell = ws.cell(row=tot_row, column=c, value=float(pd.to_numeric(df_res[colname], errors="coerce").sum()))
        cell.font = Font(name=font_family, bold=True)
        cell.number_format = '#,##0.00'
    for c, letter in [(16, "P"), (17, "Q"), (18, "R"), (19, "S")]:
        cell = ws.cell(row=tot_row, column=c, value=f"=SUM({letter}{r0+1}:{letter}{tot_row-1})")
        cell.font = Font(name=font_family, bold=True)
        cell.number_format = '#,##0.00'

    widths = [5, 26, 20, 32, 10, 14, 44, 16, 14, 14, 20, 56, 10, 10, 10, 16, 16, 16, 18, 42]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(out_path)
    print(f"Tersimpan: {out_path}")


if __name__ == "__main__":
    main()
