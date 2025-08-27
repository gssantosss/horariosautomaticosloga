import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Ordenar Horários", layout="wide")
st.title("🕒 Ordenar Horários do Maior para o Menor")
st.write("Faça upload da planilha para ordenar os horários.")

# Upload do arquivo
uploaded_file = st.file_uploader("📂 Carregue sua planilha (Excel)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.subheader("📊 Dados Originais")
    st.dataframe(df)

    # Identifica colunas que contêm horários
    horario_cols = [col for col in df.columns if col.startswith("HORARIO")]

    # Verifica se há colunas de horário
    if not horario_cols:
        st.error("Nenhuma coluna de horário encontrada. Verifique o nome das colunas.")
    else:
        # Ordena os horários do maior para o menor
        for col in horario_cols:
            # Tenta converter para datetime
            try:
                df[col] = pd.to_datetime(df[col], format='%H:%M', errors='coerce')  # Converte para datetime
                if df[col].isnull().all():
                    st.warning(f"A coluna '{col}' não contém horários válidos.")
                else:
                    df[col] = df[col].sort_values(ascending=False).reset_index(drop=True)  # Ordena do maior para o menor
            except Exception as e:
                st.error(f"Erro ao processar a coluna '{col}': {e}")

        st.subheader("📊 Horários Ordenados do Maior para o Menor")
        st.dataframe(df[horario_cols])

        # Salvar Excel em memória
        output = BytesIO()
        original_filename = uploaded_file.name.split(".")[0]
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            sheet_name = f"{original_filename}_ordenado"[:31]
            df.to_excel(writer, index=False, sheet_name=sheet_name)
        output.seek(0)

        corrected_filename = f"{original_filename}_ordenado.xlsx"

        st.download_button(
            label="⬇️ Baixar arquivo ordenado",
            data=output,
            file_name=corrected_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
