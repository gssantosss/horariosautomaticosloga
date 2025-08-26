import streamlit as st
import pandas as pd
import io
import os

st.title("🕒Correção Automática de Horários")

uploaded_file = st.file_uploader("Faça upload do arquivo Excel", type=["xlsx"])

# Configuração do limite de gap
limite_gap = st.number_input("Defina o limite máximo de gap (em minutos)", min_value=1, value=10, step=1)

if uploaded_file:
    # Pegando o nome original do arquivo (sem extensão)
    original_filename = os.path.splitext(uploaded_file.name)[0]

    # Lendo Excel
    df = pd.read_excel(uploaded_file)

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
            new_df.loc[subset.index, horario_col] = [h.strftime("%H:%M") for h in horarios]]

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
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


