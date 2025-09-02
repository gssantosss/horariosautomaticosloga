import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta

st.title("📊 Mini tabela HORARIO + PROCV + Contagem de Valores")

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
        try:
            return datetime.strptime(val.strip(), "%H:%M:%S")
        except:
            try:
                return datetime.strptime(val.strip(), "%H:%M")
            except:
                return None
    return None

def get_gap_times(temp, min_gap=10):
    """Retorna listas de horários antes e depois de gaps"""
    before_gap, after_gap = [], []
    for i in range(1, len(temp)):
        diff = (temp[i] - temp[i-1]).total_seconds() / 60
        if diff > min_gap:
            before_gap.append(temp[i-1])
            after_gap.append(temp[i])
    return before_gap, after_gap

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Detecta colunas HORARIO preenchidas
    horario_cols = [col for col in df.columns if col.upper().startswith("HORARIO")]
    horario_cols = [col for col in horario_cols if df[col].notna().any()]

    if not horario_cols:
        st.write("❌ Nenhuma coluna HORARIO preenchida encontrada.")
    else:
        for col in horario_cols:
            df[col] = df[col].apply(parse_excel_time)

        mini_tabela = {}
        max_gaps = 0

        # Monta mini tabela por coluna
        for col in horario_cols:
            ordem_col = col.replace("HORARIO", "ORDEM")
            temp = df[[col, ordem_col]].dropna().sort_values(by=col).reset_index(drop=True)

            if temp.empty:
                mini_tabela[col] = []
                mini_tabela[col + "_ORDEM"] = []
                mini_tabela[col + "_AntesGap"] = []
                mini_tabela[col + "_DepoisGap"] = []
                continue

            # Menor e Maior
            menor = temp[col].iloc[0]
            maior = temp[col].iloc[-1]

            # Horários antes e depois dos gaps
            before_gap, after_gap = get_gap_times(temp[col], min_gap=10)
            max_gaps = max(max_gaps, len(before_gap))

            # Lista final
            horarios = [menor] + before_gap + after_gap + [maior]
            ordens = [temp[ordem_col].iloc[0]]  # Menor horário
            # PROCV para gaps
            for h in before_gap + after_gap:
                ordens.append(temp.loc[temp[col] == h, ordem_col].iloc[0])
            ordens.append(temp[ordem_col].iloc[-1])  # Maior horário

            # Antes e depois separados
            mini_tabela[col] = [h.strftime("%H:%M") for h in horarios]
            mini_tabela[col + "_ORDEM"] = ordens
            mini_tabela[col + "_AntesGap"] = [h.strftime("%H:%M") for h in before_gap] + [""] * (len(horarios) - len(before_gap))
            mini_tabela[col + "_DepoisGap"] = [h.strftime("%H:%M") for h in after_gap] + [""] * (len(horarios) - len(after_gap))

        # Normaliza comprimento das listas
        max_len = max(len(v) for v in mini_tabela.values())
        for k in mini_tabela:
            while len(mini_tabela[k]) < max_len:
                mini_tabela[k].append("")

        # Adiciona coluna Cont.Valores (total de linhas do DF)
        mini_tabela["Cont.Valores"] = [df.shape[0]] * max_len

        mini_df = pd.DataFrame(mini_tabela)
        mini_df.index = range(1, len(mini_df)+1)

        st.subheader("📊 Mini tabela HORARIO + ORDEM + Antes/Depois dos Gaps + Cont.Valores")
        st.dataframe(mini_df)
