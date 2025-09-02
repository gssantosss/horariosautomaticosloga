import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta

st.title("📂 Colunas HORARIO preenchidas + Menor horário")

uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])

def parse_excel_time(val):
    """Converte valores de Excel (float), datetime.time ou string para datetime"""
    if pd.isna(val):
        return None
    if isinstance(val, float):  # Excel fraction
        return datetime(1899, 12, 30) + timedelta(days=val)
    if isinstance(val, time):
        return datetime.combine(datetime.today(), val)
    if isinstance(val, datetime):
        return val
    if isinstance(val, str):
        try:
            return datetime.strptime(val.strip(), "%H:%M:%S")
        except:
            try:
                return datetime.strptime(val.strip(), "%H:%M")
            except:
                return None
    return None

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # Detecta colunas HORARIO preenchidas
    horario_cols = [col for col in df.columns if col.upper().startswith("HORARIO")]
    horario_cols = [col for col in horario_cols if df[col].notna().any()]
    
    if not horario_cols:
        st.write("❌ Nenhuma coluna HORARIO preenchida encontrada.")
    else:
        st.subheader("📋 Colunas HORARIO preenchidas")
        st.dataframe(df[horario_cols])
        
        menores = {}
        for col in horario_cols:
            # Normaliza todos os valores da coluna
            temp = df[col].apply(parse_excel_time)
            temp = temp.dropna()
            if not temp.empty:
                menores[col] = temp.min().strftime("%H:%M")
            else:
                menores[col] = "Sem valor"
        
        st.subheader("⏱ Menor horário de cada coluna HORARIO")
        st.table(pd.DataFrame([menores]))
