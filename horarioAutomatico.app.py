import streamlit as st
import pandas as pd
import numpy as np
import io
import os

st.title("🕐 Correção de Horários de Coleta")

# Upload do arquivo
uploaded_file = st.file_uploader("Carregue aqui a planilha Excel (.xlsx)", type=["xlsx", "xls"])

# Configuração do limite de gap
limite_gap = st.number_input("Defina o limite máximo de gap (em minutos)", min_value=1, value=10, step=1)

if uploaded_file:
    # Lê o arquivo
    df = pd.read_excel(uploaded_file)

    st.subheader("📋 Pré-visualização dos dados originais")
    st.dataframe(df.head())

    # Garante que a coluna "HORARIO" esteja no formato datetime
    if "HORARIO" not in df.columns:
        st.error("A planilha precisa ter uma coluna chamada 'HORARIO'")
    else:
        df["HORARIO"] = pd.to_datetime(df["HORARIO"], errors="coerce")

        # Ordena internamente os horários (não altera a ordem da planilha original)
        horarios_ordenados = df["HORARIO"].dropna().sort_values().reset_index(drop=True)

        # Calcula diferenças
        diffs = horarios_ordenados.diff().dt.total_seconds().div(60)

        # Encontra onde tem gaps
        gaps = diffs > limite_gap

        # Corrige horários bugados entre gaps
        horarios_corrigidos = horarios_ordenados.copy()
        start = 0
        for i, gap in enumerate(gaps):
            if gap:
                end = i  # posição do último antes do gap
                # corrige somente os miolos (mantém start e end fixos)
                if end - start > 2:  
                    bloco = horarios_ordenados[start:end+1]
                    novos = pd.Series(pd.date_range(start=bloco.iloc[0], end=bloco.iloc[-1], periods=len(bloco)))
                    horarios_corrigidos[start+1:end] = novos[1:-1].values
                start = end

        # Substitui no df original sem mudar ordem
        df_corrigido = df.copy()
        idx_ordenados = df["HORARIO"].dropna().sort_values().index
        df_corrigido.loc[idx_ordenados, "HORARIO"] = horarios_corrigidos.values

        st.subheader("📋 Dados corrigidos")
        st.dataframe(df_corrigido.head())

        # Monta nome do arquivo
        nome_original = uploaded_file.name
        nome_corrigido = os.path.splitext(nome_original)[0] + "_corrigido.xlsx"

        # Salva em buffer
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_corrigido.to_excel(writer, index=False)
        st.download_button("⬇️ Baixar planilha corrigida", 
                           data=output.getvalue(),
                           file_name=nome_corrigido,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    



