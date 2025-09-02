import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta

st.title("📂 Colunas HORARIO preenchidas + Ordenação individual + Gaps")

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

min_gap = timedelta(minutes=10)

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
        
        # --- Menor e Maior horário ---
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
        
        st.subheader("⏱ Menor e Maior horário de cada coluna HORARIO")
        st.table(pd.DataFrame(extremos).T)
        
        # --- Horários antes de gaps acima de 10 minutos ---
        gaps_dict = {}
        for col in horario_cols:
            temp = sorted([h for h in df[col] if h is not None])
            gaps_before = []
            for i in range(1, len(temp)):
                if temp[i] - temp[i-1] > min_gap:
                    gaps_before.append(temp[i-1].strftime("%H:%M"))
            gaps_dict[col] = gaps_before if gaps_before else ["Nenhum gap >10min"]
        
        st.subheader("⚡ Horários antes de gaps > 10 minutos")
        st.table(pd.DataFrame(gaps_dict, index=[0]))
        
        # --- Ordena cada coluna individualmente ---
        df_sorted = df.copy()
        for col in horario_cols:
            filled = df_sorted[col].dropna().sort_values(ascending=True).reset_index(drop=True)
            sorted_col = pd.Series([pd.NaT]*len(df_sorted))
            sorted_col[:len(filled)] = filled
            df_sorted[col] = sorted_col.dt.strftime("%H:%M")
        
        st.subheader("📋 Colunas HORARIO preenchidas - Ordenadas individualmente")
        st.dataframe(df_sorted[horario_cols])
