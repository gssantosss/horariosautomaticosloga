import streamlit as st
import pandas as pd
import io

st.title("Ajustar Planilha de Horários")

uploaded_file = st.file_uploader("Faça upload do arquivo Excel", type=["xlsx"])

if uploaded_file is not None:
    # lê a planilha
    df = pd.read_excel(uploaded_file)

    # normaliza a coluna de horários
    if "HORARIOTER" in df.columns:
        df["HORARIOTER"] = pd.to_datetime(df["HORARIOTER"], errors="coerce").dt.strftime("%H:%M")

    # força a ordem fixa das colunas
    colunas_ordem = ["HORARIOTER", "ORDEMTER"]
    df = df[colunas_ordem]

    # salva em memória
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Planilha Ajustada")

        workbook  = writer.book
        worksheet = writer.sheets["Planilha Ajustada"]

        # formata a coluna HORARIOTER para exibir como hh:mm
        col_idx = df.columns.get_loc("HORARIOTER")
        cell_format = workbook.add_format({"num_format": "hh:mm"})
        worksheet.set_column(col_idx, col_idx, 8, cell_format)

    # botão de download
    st.download_button(
        label="Baixar planilha ajustada",
        data=buffer.getvalue(),
        file_name="planilha_ajustada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
