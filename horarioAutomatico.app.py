import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta

st.title("⏱ Mini tabela: Menor, Gaps e Maior horário de cada HORARIO")

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

        # Cria mini tabela
        mini_tabela = {}
        for col in horario_cols:
            temp = df[col].dropna().sort_values().reset_index(drop=True)
            if temp.empty:
                mini_tabela[col] = ["Sem valor"]
                continue

            # Menor horário
            linha = [temp.iloc[0].strftime("%H:%M")]

            # Horários antes de gaps > 10 min
            for i in range(1, len(temp)):
                diff = (temp.iloc[i] - temp.iloc[i-1]).total_seconds() / 60  # minutos
                if diff > 10:
                    linha.append(temp.iloc[i-1].strftime("%H:%M"))

            # Maior horário
            linha.append(temp.iloc[-1].strftime("%H:%M"))

            mini_tabela[col] = linha

        # Transforma em DataFrame com índice começando em 1
        max_len = max(len(v) for v in mini_tabela.values())
        for k in mini_tabela:
            while len(mini_tabela[k]) < max_len:
                mini_tabela[k].append("")

        mini_df = pd.DataFrame(mini_tabela)
        mini_df.index = range(1, len(mini_df)+1)

        st.subheader("📊 Menor horário, horários antes de gaps >10min e maior horário")
        st.dataframe(mini_df)
