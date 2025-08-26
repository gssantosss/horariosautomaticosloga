import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

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

    # Copia para trabalhar
    new_df = df.copy()
    
   # Loop em todos os pares HORARIO/ORDEM
    for col in df.columns:
        if col.startswith("HORARIO"):
            dia = col.replace("HORARIO", "")
            horario_col = col
            ordem_col = f"ORDEM{dia}"

            if ordem_col not in df.columns:
                continue  # se não tiver coluna ORDEM correspondente, pula
            # Subset de ordem + horário
            subset = df[[horario_col, ordem_col]].dropna().copy()
            if subset.empty:
                continue

            # Ordena pelo número da ordem
            subset = subset.sort_values(by=ordem_col)

            inicio = pd.to_datetime(subset[horario_col].min())
            fim = pd.to_datetime(subset[horario_col].max())

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

    # Salvar em memória o Excel corrigido
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet_name = f"{original_filename}_corrigido"
        sheet_name = sheet_name[:31]  # garante no máximo 31 caracteres
    
        df.to_excel(writer, index=False, sheet_name=sheet_name)

    # Nome final do arquivo
    corrected_filename = f"{original_filename}_corrigido.xlsx"

    # Botão de download
    st.download_button(
        label="⬇️ Baixar arquivo corrigido",
        data=output,
        file_name=corrected_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",)
