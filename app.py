import io
import streamlit as st
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Dashboard Peramalan PO Konveksi", layout="wide")

NAMA_BULAN = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

def hitung_tahun_bulan(periode: int):
    total_bulan = (2024 * 12 + 2) + (periode - 1)
    tahun = total_bulan // 12
    bulan = NAMA_BULAN[total_bulan % 12]
    return tahun, bulan

def buat_file_excel_rapi(df: pd.DataFrame, judul_sheet: str, judul_laporan: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = judul_sheet[:31]

    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_hdr = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10, color="1E293B")
    font_pred = Font(name="Calibri", size=10, bold=True, color="1E293B")

    fill_navy = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    fill_slate = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    fill_pred = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    thin = Side(border_style="thin", color="CBD5E1")
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)

    jml_kolom = len(df.columns)
    last_col_letter = get_column_letter(jml_kolom)

    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = judul_laporan
    ws["A1"].font = font_title
    ws["A1"].fill = fill_navy
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 28

    ws.row_dimensions[3].height = 24
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=3, column=col_idx, value=col_name)
        cell.font = font_hdr
        cell.fill = fill_slate
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_cell

    kolom_teks = {"Bulan", "Status Data"}
    kolom_tanpa_ribuan = {"Tahun", "Periode"}

    for row_idx, row_data in enumerate(df.itertuples(index=False), start=4):
        ws.row_dimensions[row_idx].height = 20
        status_val = row_data[-1] if "Status Data" in df.columns else ""
        is_prediksi = (status_val == "Prediksi")

        for col_idx, val in enumerate(row_data, start=1):
            col_name = df.columns[col_idx - 1]
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = border_cell
            cell.font = font_pred if is_prediksi else font_data

            if is_prediksi:
                cell.fill = fill_pred
            elif row_idx % 2 == 1:
                cell.fill = fill_zebra

            if col_name in kolom_tanpa_ribuan or col_name == "Status Data":
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name in kolom_teks:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0"

    for col_idx, col_name in enumerate(df.columns, start=1):
        col_let = get_column_letter(col_idx)
        max_len = max(
            len(str(col_name)),
            df[col_name].astype(str).map(len).max() if not df.empty else 0
        )
        ws.column_dimensions[col_let].width = max(max_len + 5, 13)

    ws.freeze_panes = "A4"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

def inisialisasi_data():
    if "df_bulanan" not in st.session_state:
        po_bulanan = [
            974, 924, 1622, 1301, 736, 710, 1611, 975, 1075, 972,
            1420, 1071, 1137, 1213, 1630, 1411, 877, 1657, 2175, 1539, 1135, 798,
            1207, 1231, 807, 1380, 886, 807, 942, 1052, 2043, 1728
        ]
        rows = []
        for idx, val in enumerate(po_bulanan, start=1):
            thn, bln = hitung_tahun_bulan(idx)
            rows.append({
                "Tahun": thn,
                "Bulan": bln,
                "Periode": idx,
                "Jumlah PO": int(val),
                "Status Data": "Aktual"
            })
        st.session_state.df_bulanan = pd.DataFrame(rows)

    if "df_produk" not in st.session_state:
        data_prod = {
            "WORKJACKET": [
                281, 207, 575, 301, 193, 36, 345, 210, 0, 0,
                579, 313, 237, 862, 217, 455, 228, 259, 391, 511, 547, 339,
                687, 931, 265, 519, 329, 582, 720, 393, 1503, 1396
            ],
            "VARSITY": [
                251, 544, 422, 818, 310, 560, 1203, 441, 594, 300,
                523, 658, 790, 351, 562, 597, 509, 312, 432, 254, 491, 303,
                333, 160, 100, 104, 204, 75, 70, 378, 100, 87
            ],
            "COACH": [
                330, 90, 573, 134, 233, 101, 26, 151, 0, 122,
                37, 64, 110, 0, 0, 0, 0, 0, 89, 197, 97, 71,
                73, 24, 0, 15, 0, 0, 22, 150, 82, 137
            ]
        }
        rows = []
        for idx in range(1, 33):
            thn, bln = hitung_tahun_bulan(idx)
            rows.append({
                "Tahun": thn,
                "Bulan": bln,
                "Periode": idx,
                "WORKJACKET": data_prod["WORKJACKET"][idx - 1],
