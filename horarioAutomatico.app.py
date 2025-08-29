import streamlit as st
import pandas as pd

uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    st.write("### Tipos das colunas após leitura:")
    st.write(df.dtypes)
    
    dias = ["SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]
    horario_cols = [f"HORARIO{dia}" for dia in dias if f"HORARIO{dia}" in df.columns]
    
    for col in horario_cols:
        st.write(f"### Valores originais da coluna {col}:")
        st.write(df[col].head(10))
        
        # Tente converter para datetime com formato %H:%M
        try:
            converted = pd.to_datetime(df[col].astype(str).str.strip(), format='%H:%M', errors='coerce')
            st.write(f"### Valores convertidos da coluna {col}:")
            st.write(converted.head(10))
            st.write(f"### Quantidade de valores nulos após conversão em {col}: {converted.isna().sum()}")
        except Exception as e:
            st.error(f"Erro ao converter a coluna {col}: {e}")
