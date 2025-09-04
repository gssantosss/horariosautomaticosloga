import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta

st.title("📊 Mini tabela HORARIO + PROCV para ORDEM")

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
        for fmt in ["%H:%M:%S", "%H:%M"]:
            try:
                return datetime.strptime(val.strip(), fmt)
            except:
                pass
    return None

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Detecta colunas HORARIO preenchidas
    horario_cols = [col for col in df.columns if col.upper().startswith("HORARIO")]
    horario_cols = [col for col in horario_cols if df[col].notna().any()]

    if not horario_cols:
        st.write("❌ Nenhuma coluna HORARIO preenchida encontrada.")
    else:
        # Normaliza horários
        for col in horario_cols:
            df[col] = df[col].apply(parse_excel_time)

        mini_tabela = {}

        max_gaps = 0  # para alinhar linhas entre colunas
        for col in horario_cols:
            ordem_col = col.replace("HORARIO", "ORDEM")
            temp = df[[col, ordem_col]].dropna().sort_values(by=col).reset_index(drop=True)

            horarios_fmt = []
            ordens_fmt = []

            if not temp.empty:
                # INÍCIO
                horarios_fmt.append(temp[col].iloc[0].strftime("%H:%M"))
                ordens_fmt.append(temp[ordem_col].iloc[0])

                # GAPs
                gap_count = 0
                for i in range(1, len(temp)):
                    diff = (temp[col].iloc[i] - temp[col].iloc[i-1]).total_seconds() / 60
                    if diff > 10:
                        gap_count += 1
                        # ANTERIOR GAP
                        horarios_fmt.append(temp[col].iloc[i-1].strftime("%H:%M"))
                        ordens_fmt.append(temp[ordem_col].iloc[i-1])
                        # POSTERIOR GAP
                        horarios_fmt.append(temp[col].iloc[i].strftime("%H:%M"))
                        ordens_fmt.append(temp[ordem_col].iloc[i])

                # FINAL
                horarios_fmt.append(temp[col].iloc[-1].strftime("%H:%M"))
                ordens_fmt.append(temp[ordem_col].iloc[-1])

                if gap_count > max_gaps:
                    max_gaps = gap_count

            mini_tabela[col] = horarios_fmt
            mini_tabela[ordem_col] = ordens_fmt

        # constrói índice da tabela (linhas fixas)
        index_labels = ["INÍCIO"]
        for g in range(1, max_gaps+1):
            index_labels.append(f"ANTERIOR GAP{g}")
            index_labels.append(f"POSTERIOR GAP{g}")
        index_labels.append("FINAL")

        # normaliza todas as colunas pro mesmo tamanho
        max_len = len(index_labels)
        for k in mini_tabela:
            while len(mini_tabela[k]) < max_len:
                mini_tabela[k].append("")

        mini_df = pd.DataFrame(mini_tabela, index=index_labels)

        st.subheader("📊 Mini tabela HORARIO + ORDEM (modelo fixo)")
        st.dataframe(mini_df)
