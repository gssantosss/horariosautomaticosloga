import pandas as pd
import streamlit as st
from io import BytesIO
import os

st.title("Ajuste e Visualização de Horários")

uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    dias = ["SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]
    horario_cols = [f"HORARIO{dia}" for dia in dias if f"HORARIO{dia}" in df.columns]

    # Converte colunas HORARIO para datetime (hora e minuto)
    for col in horario_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col].astype(str).str.strip(), format='%H:%M:%S', errors='coerce')

    # Cria uma versão para exibição com os horários como texto HH:MM e células vazias
    df_display = df.copy()
    for col in horario_cols:
        if col in df_display.columns:
            df_display[col] = df_display[col].apply(lambda x: x.strftime('%H:%M') if pd.notnull(x) else "")

    st.write("Planilha com horários convertidos:")
    st.dataframe(df_display)

    # Exporta para Excel com formato de hora
    output = BytesIO()
    original_name = uploaded_file.name
    name, ext = os.path.splitext(original_name)
    novo_nome = f"{name}_ajustado.xlsx"

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Ajustado")
        workbook = writer.book
        worksheet = writer.sheets["Ajustado"]

        time_format = workbook.add_format({'num_format': 'hh:mm'})

        for col in horario_cols:
            if col in df.columns:
                col_idx = df.columns.get_loc(col)
                worksheet.set_column(col_idx, col_idx, 12, time_format)

    output.seek(0)

    st.success("✅ Ajuste concluído!")
    st.download_button(
        label="⬇️ Baixar planilha ajustada",
        data=output,
        file_name=novo_nome,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
