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
                "VARSITY": data_prod["VARSITY"][idx - 1],
                "COACH": data_prod["COACH"][idx - 1],
                "Status Data": "Aktual"
            })
        st.session_state.df_produk = pd.DataFrame(rows)
        st.session_state.daftar_produk = ["WORKJACKET", "VARSITY", "COACH"]

    if "mode_edit_bulanan" not in st.session_state:
        st.session_state.mode_edit_bulanan = False
    if "mode_edit_produk" not in st.session_state:
        st.session_state.mode_edit_produk = False

def buat_prediksi_3_bulan(df_aktual: pd.DataFrame, kolom_target: list):
    X = df_aktual[["Periode"]].values
    periode_terakhir = int(df_aktual["Periode"].max())
    periode_depan = np.arange(periode_terakhir + 1, periode_terakhir + 4).reshape(-1, 1)

    rows_pred = []
    for p in periode_depan.flatten():
        thn, bln = hitung_tahun_bulan(int(p))
        rows_pred.append({
            "Tahun": thn,
            "Bulan": bln,
            "Periode": int(p),
            "Status Data": "Prediksi"
        })

    df_pred = pd.DataFrame(rows_pred)

    for col in kolom_target:
        y = df_aktual[col].fillna(0).values
        model = LinearRegression()
        model.fit(X, y)
        pred = model.predict(periode_depan)
        df_pred[col] = np.maximum(0, np.round(pred)).astype(int)

    kolom_urut = ["Tahun", "Bulan", "Periode"] + kolom_target + ["Status Data"]
    return df_pred[kolom_urut]

def warnai_baris_prediksi(row):
    if row["Status Data"] == "Prediksi":
        return ["background-color: #FEF9C3; color: #1E293B; font-weight: bold"] * len(row)
    return [""] * len(row)

inisialisasi_data()

# --- SIDEBAR NAVIGASI & UPLOAD EXCEL ---
st.sidebar.title("Menu Navigasi")
halaman = st.sidebar.radio(
    "Pilih Halaman:",
    ["1. PO Bulanan", "2. Peramalan Produk Utama"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("Import Data dari Excel / CSV")
uploaded_file = st.sidebar.file_uploader(
    "Upload file template Excel (.xlsx) atau CSV",
    type=["xlsx", "csv"]
)

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_up = pd.read_csv(uploaded_file)
        else:
            df_up = pd.read_excel(uploaded_file)

        df_up.columns = [str(c).strip() for c in df_up.columns]
        if "Tahun" in df_up.columns:
            df_up["Tahun"] = df_up["Tahun"].ffill()

        if "Periode" in df_up.columns:
            df_up = df_up.dropna(subset=["Periode"]).copy()
            df_up["Periode"] = df_up["Periode"].astype(int)

        mode_import = st.sidebar.radio(
            "Pilih cara memproses file:",
            ["Tambahkan data baru ke dashboard", "Ganti seluruh data dengan file ini"]
        )

        if st.sidebar.button("Proses File Excel", width="stretch"):
            if "Jumlah PO" in df_up.columns:
                df_b_baru = df_up.dropna(subset=["Jumlah PO"]).copy()
                if "Status Data" in df_b_baru.columns:
                    df_b_baru = df_b_baru[df_b_baru["Status Data"] != "Prediksi"]

                rows_b = []
                for _, r in df_b_baru.iterrows():
                    p = int(r["Periode"])
                    thn, bln = hitung_tahun_bulan(p)
                    rows_b.append({
                        "Tahun": int(r["Tahun"]) if "Tahun" in r and pd.notna(r["Tahun"]) else thn,
                        "Bulan": str(r["Bulan"]) if "Bulan" in r and pd.notna(r["Bulan"]) else bln,
                        "Periode": p,
                        "Jumlah PO": int(r["Jumlah PO"]),
                        "Status Data": "Aktual"
                    })
                df_b_clean = pd.DataFrame(rows_b)

                if mode_import == "Ganti seluruh data dengan file ini":
                    st.session_state.df_bulanan = df_b_clean
                else:
                    gabung_b = pd.concat(
                        [st.session_state.df_bulanan, df_b_clean],
                        ignore_index=True
                    )
                    gabung_b = gabung_b.drop_duplicates(subset=["Periode"], keep="last")
                    gabung_b = gabung_b.sort_values("Periode").reset_index(drop=True)
                    st.session_state.df_bulanan = gabung_b

            kolom_non_produk = {"Tahun", "Bulan", "Periode", "Jumlah PO", "Status Data"}
            kolom_prod_file = [c for c in df_up.columns if c not in kolom_non_produk]

            if len(kolom_prod_file) > 0:
                df_p_baru = df_up.dropna(subset=kolom_prod_file, how="all").copy()
                if "Status Data" in df_p_baru.columns:
                    df_p_baru = df_p_baru[df_p_baru["Status Data"] != "Prediksi"]

                rows_p = []
                for _, r in df_p_baru.iterrows():
                    p = int(r["Periode"])
                    thn, bln = hitung_tahun_bulan(p)
                    item = {
                        "Tahun": int(r["Tahun"]) if "Tahun" in r and pd.notna(r["Tahun"]) else thn,
                        "Bulan": str(r["Bulan"]) if "Bulan" in r and pd.notna(r["Bulan"]) else bln,
                        "Periode": p,
                        "Status Data": "Aktual"
                    }
                    for kp in kolom_prod_file:
                        item[kp] = int(r[kp]) if pd.notna(r[kp]) else 0
                    rows_p.append(item)

                df_p_clean = pd.DataFrame(rows_p)
                if mode_import == "Ganti seluruh data dengan file ini":
                    st.session_state.daftar_produk = kolom_prod_file
                    kolom_akhir = ["Tahun", "Bulan", "Periode"] + kolom_prod_file + ["Status Data"]
                    st.session_state.df_produk = df_p_clean[kolom_akhir]
                else:
                    for kp in kolom_prod_file:
                        if kp not in st.session_state.daftar_produk:
                            st.session_state.daftar_produk.append(kp)
                            st.session_state.df_produk[kp] = 0
                    gabung_p = pd.concat(
                        [st.session_state.df_produk, df_p_clean],
                        ignore_index=True
                    ).fillna(0)
                    for kp in st.session_state.daftar_produk:
                        gabung_p[kp] = gabung_p[kp].astype(int)
                    gabung_p = gabung_p.drop_duplicates(subset=["Periode"], keep="last")
                    gabung_p = gabung_p.sort_values("Periode").reset_index(drop=True)
                    kolom_akhir = ["Tahun", "Bulan", "Periode"] + st.session_state.daftar_produk + ["Status Data"]
                    st.session_state.df_produk = gabung_p[kolom_akhir]

            st.sidebar.success("Data berhasil diproses!")
            st.rerun()

    except Exception as e:
        st.sidebar.error(f"Gagal membaca file: {e}")

# ==============================================================================
# HALAMAN 1: PO BULANAN
# ==============================================================================
if halaman == "1. PO Bulanan":
    st.title("Peramalan Total Purchase Order (PO) Bulanan")
    st.caption("Metode: Regresi Linear Sederhana (Proyeksi 3 Periode ke Depan)")

    df_aktual_b = st.session_state.df_bulanan.copy()
    df_pred_b = buat_prediksi_3_bulan(df_aktual_b, ["Jumlah PO"])
    df_gabung_b = pd.concat([df_aktual_b, df_pred_b], ignore_index=True)

    cols_metric = st.columns(3)
    for i, row_p in df_pred_b.iterrows():
        with cols_metric[i]:
            st.metric(
                label=f"Prediksi {row_p['Bulan']} {row_p['Tahun']} (P-{row_p['Periode']})",
                value=f"{row_p['Jumlah PO']:,} pcs"
            )

    fig_b = go.Figure()
    fig_b.add_trace(go.Scatter(
        x=df_aktual_b["Periode"],
        y=df_aktual_b["Jumlah PO"],
        mode="lines+markers",
        name="Data Aktual",
        line=dict(color="#1E3A5F", width=2.5),
        hovertemplate="Periode %{x}<br>Jumlah PO: %{y:,}<extra></extra>"
    ))

    df_garis_pred = pd.concat([df_aktual_b.iloc[[-1]], df_pred_b], ignore_index=True)
    fig_b.add_trace(go.Scatter(
        x=df_garis_pred["Periode"],
        y=df_garis_pred["Jumlah PO"],
        mode="lines+markers",
        name="Prediksi (3 Bulan)",
        line=dict(color="#D97706", width=2.5, dash="dash"),
        marker=dict(size=9),
        hovertemplate="Periode %{x}<br>Prediksi PO: %{y:,}<extra></extra>"
    ))
    fig_b.update_layout(
        title="Grafik Aktual vs Prediksi PO Bulanan",
        xaxis_title="Periode (Bulan ke-)",
        yaxis_title="Jumlah PO (pcs)",
        hovermode="x unified",
        height=400
    )
    st.plotly_chart(fig_b, width="stretch")

    col_btn1, col_btn2, _ = st.columns([1.2, 1.5, 3])
    with col_btn1:
        if st.button("Edit Data Aktual", key="btn_edit_b", width="stretch"):
            st.session_state.mode_edit_bulanan = not st.session_state.mode_edit_bulanan
    with col_btn2:
        excel_bytes_b = buat_file_excel_rapi(
            df_gabung_b,
            "PO Bulanan",
            "Laporan Aktual & Peramalan PO Bulanan Konveksi"
        )
        st.download_button(
            label="Download Laporan Excel",
            data=excel_bytes_b,
            file_name="Laporan_Peramalan_PO_Bulanan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch"
        )

    if st.session_state.mode_edit_bulanan:
        st.info("Ubah nilai pada kolom 'Jumlah PO' atau tambah baris periode baru di bawah ini, lalu klik Simpan.")
        df_edited_b = st.data_editor(
            df_aktual_b[["Periode", "Jumlah PO"]],
            num_rows="dynamic",
            width="stretch",
            key="editor_bulanan"
        )
        if st.button("Simpan Perubahan Data Bulanan", type="primary"):
            rows_new = []
            for _, r in df_edited_b.dropna(subset=["Periode", "Jumlah PO"]).iterrows():
                p = int(r["Periode"])
                thn, bln = hitung_tahun_bulan(p)
                rows_new.append({
                    "Tahun": thn,
                    "Bulan": bln,
                    "Periode": p,
                    "Jumlah PO": int(r["Jumlah PO"]),
                    "Status Data": "Aktual"
                })
            df_sorted_b = pd.DataFrame(rows_new).sort_values("Periode").reset_index(drop=True)
            st.session_state.df_bulanan = df_sorted_b
            st.session_state.mode_edit_bulanan = False
            st.rerun()

    st.subheader("Tabel Data Aktual & Hasil Peramalan")
    st.dataframe(
        df_gabung_b.style.apply(warnai_baris_prediksi, axis=1).format({"Jumlah PO": "{:,}"}),
        width="stretch",
        hide_index=True
    )

# ==============================================================================
# HALAMAN 2: PERAMALAN PRODUK UTAMA
# ==============================================================================
elif halaman == "2. Peramalan Produk Utama":
    st.title("Peramalan PO per Kategori Produk Utama")
    st.caption("Metode: Regresi Linear Sederhana per Item Produk (Proyeksi 3 Periode ke Depan)")

    daftar_prod = st.session_state.daftar_produk
    df_aktual_p = st.session_state.df_produk.copy()
    df_pred_p = buat_prediksi_3_bulan(df_aktual_p, daftar_prod)
    df_gabung_p = pd.concat([df_aktual_p, df_pred_p], ignore_index=True)

    produk_pilihan = st.multiselect(
        "Pilih Produk untuk Ditampilkan di Grafik:",
        options=daftar_prod,
        default=daftar_prod
    )

    fig_p = go.Figure()
    warna_produk = ["#1E3A5F", "#0D9488", "#D97706", "#DC2626", "#7C3AED"]

    for idx_p, prod in enumerate(produk_pilihan):
        warna = warna_produk[idx_p % len(warna_produk)]
        fig_p.add_trace(go.Scatter(
            x=df_aktual_p["Periode"],
            y=df_aktual_p[prod],
            mode="lines+markers",
            name=f"{prod} (Aktual)",
            line=dict(color=warna, width=2)
        ))
        df_garis_p = pd.concat([df_aktual_p.iloc[[-1]], df_pred_p], ignore_index=True)
        fig_p.add_trace(go.Scatter(
            x=df_garis_p["Periode"],
            y=df_garis_p[prod],
            mode="lines+markers",
            name=f"{prod} (Prediksi)",
            line=dict(color=warna, width=2, dash="dash")
        ))

    fig_p.update_layout(
        title="Grafik Aktual vs Prediksi per Produk",
        xaxis_title="Periode (Bulan ke-)",
        yaxis_title="Jumlah PO (pcs)",
        hovermode="x unified",
        height=420
    )
    st.plotly_chart(fig_p, width="stretch")

    st.subheader("Ringkasan Prediksi 3 Bulan ke Depan")
    st.dataframe(
        df_pred_p.style.apply(warnai_baris_prediksi, axis=1).format({c: "{:,}" for c in daftar_prod}),
        width="stretch",
        hide_index=True
    )

    col_p1, col_p2, _ = st.columns([1.2, 1.5, 3])
    with col_p1:
        if st.button("Edit Data Produk", key="btn_edit_p", width="stretch"):
            st.session_state.mode_edit_produk = not st.session_state.mode_edit_produk
    with col_p2:
        excel_bytes_p = buat_file_excel_rapi(
            df_gabung_p,
            "Peramalan Produk",
            "Laporan Aktual & Peramalan PO per Produk Utama"
        )
        st.download_button(
            label="Download Laporan Excel",
            data=excel_bytes_p,
            file_name="Laporan_Peramalan_Produk_Utama.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch"
        )

    if st.session_state.mode_edit_produk:
        st.info("Edit jumlah pesanan per produk di bawah ini, lalu klik Simpan.")
        df_edited_p = st.data_editor(
            df_aktual_p[["Periode"] + daftar_prod],
            num_rows="dynamic",
            width="stretch",
            key="editor_produk"
        )
        if st.button("Simpan Perubahan Data Produk", type="primary"):
            rows_new_p = []
            for _, r in df_edited_p.dropna(subset=["Periode"]).iterrows():
                p = int(r["Periode"])
                thn, bln = hitung_tahun_bulan(p)
                row_item = {
                    "Tahun": thn,
                    "Bulan": bln,
                    "Periode": p,
                    "Status Data": "Aktual"
                }
                for prod in daftar_prod:
                    row_item[prod] = int(r[prod]) if pd.notna(r[prod]) else 0
                rows_new_p.append(row_item)
            df_sorted_p = pd.DataFrame(rows_new_p).sort_values("Periode").reset_index(drop=True)
            st.session_state.df_produk = df_sorted_p
            st.session_state.mode_edit_produk = False
            st.rerun()

    st.subheader("Tabel Lengkap Aktual & Peramalan Produk")
    st.dataframe(
        df_gabung_p.style.apply(warnai_baris_prediksi, axis=1).format({c: "{:,}" for c in daftar_prod}),
        width="stretch",
        hide_index=True
    )
# END_OF_FILE
