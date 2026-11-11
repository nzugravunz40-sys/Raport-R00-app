import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Rapoarte R00 Automatizate", page_icon="📦", layout="centered"
)

st.title("📦 Generator Rapoarte R00")
st.write(
    "Încarcă fișierul tău Excel de transport (`Transport List` și `888 Report`) pentru a genera automat raportul centralizat."
)

uploaded_file = st.file_uploader(
    "Alege fișierul Excel", type=["xlsx", "xls"]
)

if uploaded_file is not None:
  try:
    with st.spinner("Se procesează datele..."):
      df_transport = pd.read_excel(uploaded_file, sheet_name="Transport List")
      df_888 = pd.read_excel(uploaded_file, sheet_name="888 Report")

      parents = df_transport["Unnamed: 10"].dropna().unique()

      summary_list = []
      for p in parents:
        children = df_transport[df_transport["Unnamed: 10"] == p]["Nr."].tolist()
        rep = df_888[df_888["ID Container"].isin(children)]
        total_qty = int(rep["Cantitate"].sum())
        total_pos = len(rep)
        summary_list.append({
            "Parent_R00": p,
            "Num_Children": len(children),
            "Total_Pozitii": total_pos,
            "Total_Cantitate": total_qty,
        })

      summary_df = pd.DataFrame(summary_list).sort_values(
          by="Total_Cantitate", ascending=False
      )
      ordered_parents = summary_df["Parent_R00"].tolist()

      wb = openpyxl.Workbook()
      wb.remove(wb.active)

      ws_summary = wb.create_sheet(title="Sumar Pachete R00")

      header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
      header_fill = PatternFill(
          start_color="1F4E79", end_color="1F4E79", fill_type="solid"
      )
      border_thin = Border(
          left=Side(style="thin", color="D9D9D9"),
          right=Side(style="thin", color="D9D9D9"),
          top=Side(style="thin", color="D9D9D9"),
          bottom=Side(style="thin", color="D9D9D9"),
      )
      total_fill = PatternFill(
          start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
      )
      total_font = Font(name="Calibri", size=11, bold=True)

      ws_summary.append([
          "Nr. crt.",
          "R00 Principal",
          "Număr Sub-R00",
          "Total Poziții",
          "Total Cantitate",
      ])
      for col in range(1, 6):
        cell = ws_summary.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

      summary_data = []
      row_idx = 1

      for parent in ordered_parents:
        sub = df_transport[df_transport["Unnamed: 10"] == parent]
        if len(sub) == 0:
          continue

        sheet_title = (
            str(parent)[-11:] if len(str(parent)) > 11 else str(parent)
        )
        ws = wb.create_sheet(title=f"R00_{sheet_title}")

        ws.merge_cells("A1:D1")
        title_cell = ws["A1"]
        title_cell.value = f"Raport Pachet Principal: {parent}"
        title_cell.font = Font(
            name="Calibri", size=13, bold=True, color="1F4E79"
        )
        title_cell.alignment = Alignment(horizontal="left", vertical="center")

        headers = [
            "Sub-R00 (Container Component)",
            "Număr Poziții (Linii)",
            "Cantitate Totală",
            "R00 Principal",
        ]
        ws.append([])
        ws.append(headers)

        for col in range(1, 5):
          cell = ws.cell(row=3, column=col)
          cell.font = header_font
          cell.fill = header_fill
          cell.alignment = Alignment(horizontal="center", vertical="center")

        p_positions = 0
        p_qty = 0

        for idx, r in sub.iterrows():
          child = r["Nr."]
          child_rep = df_888[df_888["ID Container"] == child]
          pos_count = len(child_rep)
          qty_sum = int(child_rep["Cantitate"].sum())

          p_positions += pos_count
          p_qty += qty_sum

          ws.append([child, pos_count, qty_sum, parent])
          curr_row = ws.max_row
          ws.cell(row=curr_row, column=1).alignment = Alignment(
              horizontal="center"
          )
          ws.cell(row=curr_row, column=2).alignment = Alignment(
              horizontal="right"
          )
          ws.cell(row=curr_row, column=3).alignment = Alignment(
              horizontal="right"
          )
          ws.cell(row=curr_row, column=4).alignment = Alignment(
              horizontal="center"
          )

          for c in range(1, 5):
            ws.cell(row=curr_row, column=c).border = border_thin

        ws.append(["TOTAL PACHET", p_positions, p_qty, parent])
        tot_row = ws.max_row
        for c in range(1, 5):
          cell = ws.cell(row=tot_row, column=c)
          cell.font = total_font
          cell.fill = total_fill
          cell.border = border_thin
        ws.cell(row=tot_row, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=tot_row, column=2).alignment = Alignment(horizontal="right")
        ws.cell(row=tot_row, column=3).alignment = Alignment(horizontal="right")
        ws.cell(row=tot_row, column=4).alignment = Alignment(horizontal="center")

        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["C"].width = 20
        ws.column_dimensions["D"].width = 25

        summary_data.append([row_idx, parent, len(sub), p_positions, p_qty])
        row_idx += 1

      for s_row in summary_data:
        ws_summary.append(s_row)
        c_row = ws_summary.max_row
        ws_summary.cell(row=c_row, column=1).alignment = Alignment(
            horizontal="center"
        )
        ws_summary.cell(row=c_row, column=2).alignment = Alignment(
            horizontal="center"
        )
        ws_summary.cell(row=c_row, column=3).alignment = Alignment(
            horizontal="right"
        )
        ws_summary.cell(row=c_row, column=4).alignment = Alignment(
            horizontal="right"
        )
        ws_summary.cell(row=c_row, column=5).alignment = Alignment(
            horizontal="right"
        )
        for c in range(1, 6):
          ws_summary.cell(row=c_row, column=c).border = border_thin

      tot_sum_pos = sum(x[3] for x in summary_data)
      tot_sum_qty = sum(x[4] for x in summary_data)
      ws_summary.append([
          "TOTAL GENERAL",
          "-",
          sum(x[2] for x in summary_data),
          tot_sum_pos,
          tot_sum_qty,
      ])
      s_tot_row = ws_summary.max_row
      for c in range(1, 6):
        cell = ws_summary.cell(row=s_tot_row, column=c)
        cell.font = total_font
        cell.fill = total_fill
        cell.border = border_thin
      ws_summary.cell(row=s_tot_row, column=1).alignment = Alignment(
          horizontal="center"
      )
      ws_summary.cell(row=s_tot_row, column=2).alignment = Alignment(
          horizontal="center"
      )
      ws_summary.cell(row=s_tot_row, column=3).alignment = Alignment(
          horizontal="right"
      )
      ws_summary.cell(row=s_tot_row, column=4).alignment = Alignment(
          horizontal="right"
      )
      ws_summary.cell(row=s_tot_row, column=5).alignment = Alignment(
          horizontal="right"
      )

      ws_summary.column_dimensions["A"].width = 12
      ws_summary.column_dimensions["B"].width = 25
      ws_summary.column_dimensions['C'].width = 18
      ws_summary.column_dimensions['D'].width = 18
      ws_summary.column_dimensions['E'].width = 18

      output = io.BytesIO()
      wb.save(output)
      output.seek(0)

      st.success("Raportul a fost generat cu succes!")
      st.download_button(
          label="📥 Descarcă Fișierul Excel Centralizat",
          data=output,
          file_name="Rapoarte_R00_Centralizate.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
  except Exception as e:
    st.error(f"A apărut o eroare la procesarea fișierului: {e}")
