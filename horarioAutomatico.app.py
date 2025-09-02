import streamlit as st
import pandas as pd

st.title("Resumo de Horários com Gaps 🕒")

uploaded_file = st.file_uploader("📂 Escolha a planilha Excel", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    st.subheader("📋 Planilha carregada")
    st.dataframe(df.head())

    dias = ["SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]
    gap_threshold_min = 10  # mínimo de 10 minutos

    resumo_texto = ""

    for dia in dias:
        col_horario = f"HORARIO{dia}"
        if col_horario in df.columns:
            horarios = pd.to_datetime(df[col_horario], errors='coerce').dropna().sort_values()
            if horarios.empty:
                continue

            menor = horarios.min().strftime("%H:%M")
            maior = horarios.max().strftime("%H:%M")

            diffs = horarios.diff().dt.total_seconds() / 60
            gap_indices = diffs[diffs >= gap_threshold_min].index

            gaps_txt = []
            for i, idx in enumerate(gap_indices, start=1):
                antes = horarios.loc[idx - 1].strftime("%H:%M")
                depois = horarios.loc[idx].strftime("%H:%M")
                gaps_txt.append(f"Gap{i}: {antes} → {depois}")

            resumo_texto += f"{dia}: Menor horário: {menor} | "
            resumo_texto += " | ".join(gaps_txt) + " | " if gaps_txt else ""
            resumo_texto += f"Maior horário: {maior}\n"

    st.subheader("📑 Resumo Final")
    st.text(resumo_texto)
