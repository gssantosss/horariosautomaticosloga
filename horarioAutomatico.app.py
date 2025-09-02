import streamlit as st
import pandas as pd

st.title("📂 Colunas HORARIO preenchidas")

uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # Identifica colunas HORARIO com pelo menos 1 valor
    horario_cols = [col for col in df.columns if col.upper().startswith("HORARIO")]
    horario_cols = [col for col in horario_cols if df[col].notna().any()]
    
    if not horario_cols:
        st.write("❌ Nenhuma coluna HORARIO preenchida encontrada.")
    else:
        # Seleciona apenas essas colunas preenchidas
        df_horarios = df[horario_cols]
        st.subheader("📋 Colunas HORARIO preenchidas")
        st.dataframe(df_horarios)
