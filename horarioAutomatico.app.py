import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta

st.title("📂 Colunas HORARIO preenchidas + Menor e Maior horário + Ordenação")

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
        # Normaliza todos os horários
        for col in horario_cols:
            df[col] = df[col].apply(parse_excel_time)
        
        # Cria dicionário com menor e maior horário
        extremos = {}
        for col in horario_cols:
            temp = df[col].dropna()
            if not temp.empty:
                extremos[col] = {
                    "Menor": temp.min().strftime("%H:%M"),
                    "Maior": temp.max().strftime("%H:%M")
                }
            else:
                extremos[col] = {"Menor": "Sem valor", "Maior": "Sem valor"}
        
        # Exibe tabela com menor e maior horário
        st.subheader("⏱ Menor e Maior horário de cada coluna HORARIO")
        st.table(pd.DataFrame(extremos).T)  # Transposta para ficar colunas como HORARIOxxx
        
        # Ordena cada coluna HORARIO individualmente (crescente)
        df_sorted = df.copy()
        for col in horario_cols:
            df_sorted = df_sorted.sort_values(by=col, na_position='last')
        
        st.subheader("📋 Colunas HORARIO preenchidas - Ordenadas por cada coluna")
        st.dataframe(df_sorted[horario_cols])
