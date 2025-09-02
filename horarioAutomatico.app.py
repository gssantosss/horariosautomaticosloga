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
        for col in horario_cols:
            ordem_col = col.replace("HORARIO", "ORDEM")
            temp = df[[col, ordem_col]].dropna().sort_values(by=col).reset_index(drop=True)

            if temp.empty:
                mini_tabela[col] = []
                mini_tabela[col + "_ORDEM"] = []
                continue

            # Menor horário
            horarios = [temp[col].iloc[0]]
            ordens = [temp[ordem_col].iloc[0]]

            # Horários antes e depois de gaps >10min
            for i in range(1, len(temp)):
                diff = (temp[col].iloc[i] - temp[col].iloc[i-1]).total_seconds() / 60
                if diff > 10:
                    horarios.append(temp[col].iloc[i-1])  # antes do gap
                    ordens.append(temp[ordem_col].iloc[i-1])
                    horarios.append(temp[col].iloc[i])    # depois do gap
                    ordens.append(temp[ordem_col].iloc[i])

            # Maior horário
            horarios.append(temp[col].iloc[-1])
            ordens.append(temp[ordem_col].iloc[-1])

            mini_tabela[col] = [h.strftime("%H:%M") for h in horarios]
            mini_tabela[col + "_ORDEM"] = ordens

        # Normaliza comprimento das listas
        max_len = max(len(v) for v in mini_tabela.values())
        for k in mini_tabela:
            while len(mini_tabela[k]) < max_len:
                mini_tabela[k].append("")

        # Adiciona coluna Cont.Valores
        cont_valores = [len([v for v in mini_tabela[horario_cols[0]] if v != ""])] * max_len
        mini_tabela["Cont.Valores"] = cont_valores

        mini_df = pd.DataFrame(mini_tabela)
        mini_df.index = range(1, len(mini_df)+1)

        st.subheader("📊 Mini tabela HORARIO + ORDEM + Cont.Valores")
        st.dataframe(mini_df)
