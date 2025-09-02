import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta

st.title("📂 Colunas HORARIO preenchidas + Mini Tabela de Gaps")

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

# Função para pegar horários antes/depois dos gaps
def gap_times(series, min_gap_minutes=10):
    """Retorna lista de horários antes e depois dos gaps maiores que min_gap_minutes"""
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

        # Mini tabela com Menor, Maior, Antes/Depois dos Gaps
        mini_data = {}
        cont_valores_total = df[horario_cols[0]].dropna().shape[0]

        for col in horario_cols:
            filled = df[col].dropna().sort_values()
            before_gap, after_gap = gap_times(filled, min_gap_minutes=10)

            # Cria lista com menor, horários antes/depois dos gaps e maior
            row = []
            row.append(filled.min().strftime("%H:%M") if not filled.empty else "Sem valor")
            # Adiciona horários antes e depois dos gaps
            for bg, ag in zip(before_gap, after_gap):
                row.append(f"Antes: {bg} / Depois: {ag}")
            row.append(filled.max().strftime("%H:%M") if not filled.empty else "Sem valor")
            # Preenche no mini_data
            mini_data[col] = row

        # Converte para DataFrame
        mini_tabela = pd.DataFrame.from_dict(mini_data, orient='index').transpose()

        # Adiciona cont.valores
        mini_tabela["Cont.Valores"] = [cont_valores_total] * len(mini_tabela)

        # Ajusta índice de 1 a x
        mini_tabela.index = range(1, len(mini_tabela)+1)

        st.subheader("⏱ Mini Tabela de Horários e Gaps")
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
