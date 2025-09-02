import streamlit as st
import pandas as pd

st.title("📂 Colunas HORARIO preenchidas + Menor horário")

uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # Detecta colunas HORARIO com pelo menos 1 valor
    horario_cols = [col for col in df.columns if col.upper().startswith("HORARIO")]
    horario_cols = [col for col in horario_cols if df[col].notna().any()]
    
    if not horario_cols:
        st.write("❌ Nenhuma coluna HORARIO preenchida encontrada.")
    else:
        # Exibe colunas HORARIO preenchidas
        df_horarios = df[horario_cols]
        st.subheader("📋 Colunas HORARIO preenchidas")
        st.dataframe(df_horarios)
        
        # Calcula menor horário de cada coluna
        menores = {}
        for col in horario_cols:
            series = df[col]
            
            # Se for float (fração do dia do Excel), converte
            if pd.api.types.is_float_dtype(series):
                temp = pd.to_timedelta(series, unit='d') + pd.Timestamp('1899-12-30')
            else:
                # Tenta converter strings para datetime
                temp = pd.to_datetime(series, errors='coerce')
            
            # Pega menor horário válido
            if temp.notna().any():
                menores[col] = temp.min().strftime("%H:%M")
            else:
                menores[col] = "Sem valor"
        
        st.subheader("⏱ Menor horário de cada coluna HORARIO")
        st.table(pd.DataFrame([menores]))
