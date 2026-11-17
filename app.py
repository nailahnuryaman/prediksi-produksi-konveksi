import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def buat_file_excel(df: pd.DataFrame, nama_sheet: str, kolom_angka: list) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nama_sheet)
        ws = writer.sheets[nama_sheet]

        # Definisi Style
        header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        pred_fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
        pred_font = Font(color="1E293B", bold=True)
        thin_side = Side(border_style="thin", color="CBD5E1")
        border_style = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        # Styling Header (Baris 1)
        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_style
        ws.row_dimensions[1].height = 24

        # Styling Baris Data & Highlight Baris Prediksi
        idx_status = df.columns.get_loc("Status Data") + 1
        for row_idx in range(2, len(df) + 2):
            is_prediksi = ws.cell(row=row_idx, column=idx_status).value == "Prediksi"
            for col_idx, col_name in enumerate(df.columns, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = border_style

                if is_prediksi:
                    cell.fill = pred_fill
                    cell.font = pred_font

                if col_name in kolom_angka:
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_name in ["Tahun", "Periode", "Status Data"]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        # Auto-adjust lebar kolom & Freeze header
        for col_idx, col_name in enumerate(df.columns, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = max(
                df[col_name].astype(str).map(len).max(),
                len(str(col_name))
            ) + 4
            ws.column_dimensions[col_letter].width = max_len

        ws.freeze_panes = "A2"

    return output.getvalue()
