import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import timedelta

st.set_page_config(page_title="Análise de Horários", layout="wide")
st.title("🕒 Análise de Horários da Coleta")
st.write("Faça upload da planilha e veja os horários preenchidos, menor, maior e horários antes de gaps acima de 10 minutos.")

# Upload do arquivo
uploaded_file = st.file_uploader("📂 Carregue sua planilha (Excel)", type=["xlsx"])

# Definir tamanho mínimo do gap
min_gap = timedelta(minutes=10)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Filtrar colunas HORARIO preenchidas
    horario_cols = [col for col in df.columns if col.startswith("HORARIO") and df[col].notna().any()]
    if not horario_cols:
        st.warning("Nenhuma coluna HORARIO preenchida encontrada.")
    else:
        st.subheader("📋 Colunas HORARIO preenchidas")
        st.write(horario_cols)

        # Converte para datetime.time
        for col in horario_cols:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.time

        # Tabela com horários ordenados individualmente
        st.subheader("⏱ Horários ordenados individualmente")
        df_sorted = df.copy()
        for col in horario_cols:
            horarios_col = [h for h in df[col] if pd.notna(h)]
            horarios_col.sort()
            df_sorted[col] = pd.Series(horarios_col + [None]*(len(df_sorted)-len(horarios_col)))

        # Exibe com índice começando em 1
        df_sorted.index = range(1, len(df_sorted)+1)
        st.dataframe(df_sorted[horario_cols])

        # Menor e maior horário de cada coluna
        st.subheader("📌 Menor e maior horário de cada coluna")
        menor_maior = {}
        for col in horario_cols:
            horarios_col = [h for h in df[col] if pd.notna(h)]
            menor_maior[col] = {
                "Menor horário": min(horarios_col),
                "Maior horário": max(horarios_col)
            }
        st.json(menor_maior)

        # Horários antes dos gaps acima de 10 minutos
        st.subheader(f"⚡ Horários antes de gaps maiores que {int(min_gap.total_seconds()/60)} minutos")
        gaps_info = {}
        for col in horario_cols:
            horarios_col = sorted([h for h in df[col] if pd.notna(h)])
            gaps_col = []
            for i in range(1, len(horarios_col)):
                delta = (pd.Timestamp.combine(pd.Timestamp.today(), horarios_col[i]) - 
                         pd.Timestamp.combine(pd.Timestamp.today(), horarios_col[i-1]))
                if delta > min_gap:
                    gaps_col.append(horarios_col[i-1])
            gaps_info[col] = gaps_col
        st.json(gaps_info)
