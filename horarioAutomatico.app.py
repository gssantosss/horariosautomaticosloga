import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta

st.title("📂 Colunas HORARIO preenchidas + Mini Tabela de Gaps Separados")

uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])

def parse_excel_time(val):
    if pd.isna(val):
        return None
    if isinstance(val, float):
        return datetime(1899, 12, 30) + timedelta(days=val)
    if isinstance(val, time):
        return datetime.combine(datetime.today(), val)
    if isinstance(val, datetime):
        return val
    if isinstance(val, str):
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(val.strip(), fmt)
            except:
                continue
        return None
    return None

def gap_times(series, min_gap_minutes=10):
    """Retorna listas de horários antes e depois de gaps maiores que min_gap_minutes"""
    filled = series.dropna().sort_values()
    before_gap, after_gap = [], []
    for i in range(1, len(filled)):
        diff = (filled.iloc[i] - filled.iloc[i-1]).total_seconds() / 60
        if diff > min_gap_minutes:
            before_gap.append(filled.iloc[i-1].strftime("%H:%M"))
            after_gap.append(filled.iloc[i].strftime("%H:%M"))
    return before_gap, after_gap

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

        # Mini tabela com Menor, Antes/Depois dos Gaps, Maior
        mini_data = {}
        max_gaps = 0  # para padronizar número de colunas

        for col in horario_cols:
            filled = df[col].dropna().sort_values()
            before_gap, after_gap = gap_times(filled, min_gap_minutes=10)
            max_gaps = max(max_gaps, len(before_gap))

            mini_data[col] = {
                "Menor": filled.min().strftime("%H:%M") if not filled.empty else "Sem valor",
                "Maior": filled.max().strftime("%H:%M") if not filled.empty else "Sem valor",
                "Antes_gap": before_gap,
                "Depois_gap": after_gap
            }

        # Monta DataFrame final da mini tabela
        tabela_rows = {}
        for col, info in mini_data.items():
            row = [info["Menor"]]
            for i in range(max_gaps):
                # preenche com vazio se não tiver gap
                row.append(info["Antes_gap"][i] if i < len(info["Antes_gap"]) else "")
                row.append(info["Depois_gap"][i] if i < len(info["Depois_gap"]) else "")
            row.append(info["Maior"])
            tabela_rows[col] = row

        # Cria DataFrame e ajusta índice 1 a x
        mini_tabela = pd.DataFrame(tabela_rows).transpose()
        mini_tabela.index = range(1, len(mini_tabela)+1)

        st.subheader("⏱ Mini Tabela de Horários e Gaps (Separados)")
        st.write(f"Cont.Valores: {df.shape[0]}")
        st.dataframe(mini_tabela)

        # Ordena cada coluna HORARIO individualmente (crescente)
        df_sorted = df.copy()
        for col in horario_cols:
            filled = df_sorted[col].dropna().sort_values(ascending=True).reset_index(drop=True)
            sorted_col = pd.Series([pd.NaT]*len(df_sorted))
            sorted_col[:len(filled)] = filled
            df_sorted[col] = sorted_col.dt.strftime("%H:%M")

        st.subheader("📋 Colunas HORARIO preenchidas - Ordenadas individualmente")
        st.dataframe(df_sorted[horario_cols])
