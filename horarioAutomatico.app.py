import streamlit as st
import pandas as pd

st.title("📂 Upload de Planilha e Identificação de Colunas de Horário Preenchidas")

# Upload do arquivo
uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])

if uploaded_file:
    # Lê a planilha
    df = pd.read_excel(uploaded_file)
    
    # Filtra colunas que começam com HORARIO
    horario_cols = [col for col in df.columns if col.upper().startswith("HORARIO")]
    
    # Mantém só colunas que têm pelo menos um valor não nulo
    horario_cols_validas = [col for col in horario_cols if df[col].notna().any()]
    
    if horario_cols_validas:
        st.success("✅ Colunas de horário preenchidas encontradas:")
        st.write(horario_cols_validas)
    else:
        st.warning("⚠️ Nenhuma coluna de horário preenchida encontrada na planilha.")
