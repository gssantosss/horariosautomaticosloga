import streamlit as st
import pandas as pd

st.title("Análise de Horários por Dia com Gaps 🕒")

uploaded_file = st.file_uploader("📂 Escolha a planilha Excel", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    st.subheader("📋 Planilha carregada")
    st.dataframe(df.head())

    dias = ["SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]
    gap_threshold_min = st.number_input("Tamanho mínimo do gap (minutos)", value=60, min_value=1)

    results = []

    for dia in dias:
        col_horario = f"HORARIO{dia}"
        if col_horario in df.columns:
            horarios = pd.to_datetime(df[col_horario], errors='coerce').dropna().sort_values()
            if horarios.empty:
                continue

            menor = horarios.min()
            maior = horarios.max()

            # Detecta todos os gaps maiores que threshold
            diffs = horarios.diff().dt.total_seconds() / 60  # em minutos
            gap_indices = diffs[diffs >= gap_threshold_min].index

            # Lista todos os pares antes/depois do gap
            gaps = []
            for idx in gap_indices:
                antes = horarios.loc[idx - 1]
                depois = horarios.loc[idx]
                gaps.append(f"{antes.strftime('%H:%M')} → {depois.strftime('%H:%M')}")

            results.append({
                "Dia": dia,
                "Menor Horário": menor.strftime("%H:%M"),
                "Gaps": ", ".join(gaps) if gaps else "-",
                "Maior Horário": maior.strftime("%H:%M")
            })

    st.subheader("📊 Resumo de Horários por Dia")
    st.dataframe(pd.DataFrame(results))
