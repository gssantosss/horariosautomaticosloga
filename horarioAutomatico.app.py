import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

st.set_page_config(page_title="Ajuste de Horários", layout="wide")

st.title("Ajuste de Horários de Coleta")

st.write("Faça upload da planilha, os horários serão ajustados automaticamente.")

# Upload do arquivo
uploaded_file = st.file_uploader("📂 Carregue aqui sua planilha em Excel (.xlsx)", type=["xlsx"])

# Input do tempo mínimo de pausa
pause_threshold = st.number_input(
    "Considerar pausas a partir de: (minutos)", 
    min_value=1, max_value=120, value=10
)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("📊 Pré-visualização dos dados originais")
    st.dataframe(df.head())

    # Copia para trabalhar
    new_df = df.copy()

    # Loop em todos os pares ORDEM/HORARIO
    for col in df.columns:
        if col.startswith("ORDEM"):
            dia = col.replace("ORDEM", "")
            ordem_col = col
            horario_col = f"HORARIO{dia}"

            if horario_col not in df.columns:
                continue  # se não tiver coluna HORARIO correspondente, pula

            # Subset de ordem + horário
            subset = df[[ordem_col, horario_col]].dropna().copy()
            if subset.empty:
                continue

            # Ordena pelo número da ordem
            subset = subset.sort_values(by=ordem_col)
            
            # Converte todos os horários da coluna para datetime (baseado no mesmo dia fictício)
            subset[horario_col] = subset[horario_col].apply(
                lambda t: pd.to_datetime(str(t), format="%H:%M:%S")
                if pd.notnull(t) else pd.NaT
            )
            
            # Agora sim pega o início e fim
            inicio = subset[horario_col].min()
            fim = subset[horario_col].max()
            
            total_itens = len(subset)
            if total_itens <= 1:
                continue

            intervalo_base = (fim - inicio) / (total_itens - 1)

            # Gera os novos horários
            horarios = [inicio]
            for i in range(1, total_itens):
                proximo = horarios[-1] + intervalo_base

                gap_original = (
                    pd.to_datetime(subset[horario_col].iloc[i]) 
                    - pd.to_datetime(subset[horario_col].iloc[i-1])
                )
                if gap_original >= timedelta(minutes=pause_threshold):
                    proximo = horarios[-1] + gap_original

                horarios.append(proximo)

            # Atualiza no DF final (só a coluna de horário, não mexe na ordem)
            new_df.loc[subset.index, horario_col] = [h.strftime("%H:%M") for h in horarios]

    st.subheader("✅ Dados ajustados")
    st.dataframe(new_df.head())

    # Botão para baixar
    output = BytesIO()
    new_df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    st.download_button(
        label="📥 Baixar planilha ajustada",
        data=output,
        file_name="planilha_ajustada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

