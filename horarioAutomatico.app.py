import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.set_page_config(page_title="Ajuste de Horários", layout="wide")

st.title("🕒 Ajuste Automático de Horários da Coleta")

st.write("Faça upload da planilha, ajuste os horários de acordo com a ordem e baixe o resultado.")

# Upload do arquivo
uploaded_file = st.file_uploader("📂 Carregue sua planilha (Excel)", type=["xlsx"])

# Input do tempo mínimo de pausa
pause_threshold = st.number_input(
    "Tempo mínimo de pausa (minutos)", 
    min_value=1, max_value=120, value=10
)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("📊 Pré-visualização dos dados originais")
    st.dataframe(df.head())

    # Processamento
    new_df = df.copy()

    for col in df.columns:
        if col.startswith("ORDEM"):
            # Acha o prefixo do dia (ex: TER, SAB)
            dia = col.replace("ORDEM", "")
            ordem_col = col
            horario_col = f"HORARIO{dia}"

            if horario_col not in df.columns:
                continue

            # Seleciona subset
            subset = df[[ordem_col, horario_col]].dropna().copy()

            if subset.empty:
                continue

            # Ordena por ordem
            subset = subset.sort_values(by=ordem_col)

            # Pega o primeiro e último horário originais
            inicio = pd.to_datetime(subset[horario_col].min())
            fim = pd.to_datetime(subset[horario_col].max())

            total_itens = len(subset)
            if total_itens <= 1:
                continue

            # Calcula intervalo base (distribuição linear)
            intervalo_base = (fim - inicio) / (total_itens - 1)

            # Gera novos horários
            horarios = [inicio]
            for i in range(1, total_itens):
                proximo = horarios[-1] + intervalo_base

                # Se o gap original era maior que o threshold, pula esse tempo
                gap_original = pd.to_datetime(subset[horario_col].iloc[i]) - pd.to_datetime(subset[horario_col].iloc[i-1])
                if gap_original >= timedelta(minutes=pause_threshold):
                    proximo = horarios[-1] + gap_original

                horarios.append(proximo)

            # Atualiza no DF final
            new_df.loc[subset.index, horario_col] = [h.strftime("%H:%M") for h in horarios]

    st.subheader("✅ Dados ajustados")
    st.dataframe(new_df.head())

    # Botão para baixar
    st.download_button(
        label="📥 Baixar planilha ajustada",
        data=new_df.to_excel(index=False, engine="openpyxl"),
        file_name="planilha_ajustada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")