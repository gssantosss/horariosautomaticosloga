import streamlit as st
import pandas as pd

st.title("📂 Upload de Planilha e Identificação de Colunas de Horário")

# Upload do arquivo
uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])

if uploaded_file:
    # Lê a planilha
    df = pd.read_excel(uploaded_file)
    
    # Filtra só colunas que começam com HORARIO
    horario_cols = [col for col in df.columns if col.upper().startswith("HORARIO")]
    
    if horario_cols:
        st.success("✅ Colunas de horário encontradas:")
        st.write(horario_cols)
    else:
        st.warning("⚠️ Nenhuma coluna de horário encontrada na planilha.")
