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

    # Converte colunas HORARIO para texto 'HH:MM' e deixa vazios os nulos
    for col in horario_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col].astype(str).str.strip(), format='%H:%M:%S', errors='coerce')
            df[col] = df[col].apply(lambda x: x.strftime('%H:%M') if pd.notnull(x) else "")

    # Exibe a versão formatada no Streamlit
    st.write("Planilha com horários convertidos:")
    st.dataframe(df)

    # Exporta para Excel com os horários como texto
    output = BytesIO()
    original_name = uploaded_file.name
    name, ext = os.path.splitext(original_name)
    novo_nome = f"{name}_ajustado.xlsx"

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Ajustado")
        workbook = writer.book
        worksheet = writer.sheets["Ajustado"]

        # Define largura das colunas de horário
        for col in horario_cols:
            if col in df.columns:
                col_idx = df.columns.get_loc(col)
                worksheet.set_column(col_idx, col_idx, 12)  # largura da coluna

    output.seek(0)

    st.success("✅ Ajuste concluído!")
    st.download_button(
        label="⬇️ Baixar planilha ajustada",
        data=output,
        file_name=novo_nome,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
