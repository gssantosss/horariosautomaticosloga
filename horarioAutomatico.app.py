import streamlit as st
import pandas as pd
from datetime import timedelta
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

    new_df = df.copy()

    for col in df.columns:
        if col.startswith("HORARIO"):
            dia = col.replace("HORARIO", "")
            horario_col = col
            ordem_col = f"ORDEM{dia}"

            if ordem_col not in df.columns:
                continue

            # Converte coluna de horário para datetime, valores inválidos viram NaT
            df[horario_col] = pd.to_datetime(df[horario_col], format='%H:%M', errors='coerce')
            subset = df[[horario_col, ordem_col]].dropna().copy()
            if subset.empty:
                continue

            # Ordena os horários
            subset = subset.sort_values(by=horario_col)

            # Gera os novos horários respeitando os gaps
            novos_horarios = []
            for i in range(len(subset)):
                if i == 0:
                    novos_horarios.append(subset[horario_col].iloc[i])  # Mantém o primeiro horário
                elif i == len(subset) - 1:
                    novos_horarios.append(subset[horario_col].iloc[i])  # Mantém o último horário
                else:
                    # Calcula o próximo horário respeitando o gap
                    gap = subset[horario_col].iloc[i] - subset[horario_col].iloc[i - 1]
                    if gap >= timedelta(minutes=pause_threshold):
                        # Se o gap for maior que o limite, mantém o horário original
                        novos_horarios.append(subset[horario_col].iloc[i])
                    else:
                        # Caso contrário, ajusta o horário para o próximo disponível
                        proximo_horario = novos_horarios[-1] + timedelta(minutes=pause_threshold)
                        # Garante que o próximo horário não ultrapasse o último horário
                        if proximo_horario < subset[horario_col].iloc[i + 1]:
                            novos_horarios.append(proximo_horario)
                        else:
                            novos_horarios.append(subset[horario_col].iloc[i])  # Mantém o horário original se ultrapassar

            # Atualiza no DF final
            new_df.loc[subset.index, horario_col] = [h.strftime("%H:%M") for h in novos_horarios]

    # Verifica se os horários foram corrigidos
    st.subheader("📊 Horários Corrigidos")
    st.dataframe(new_df[[col for col in new_df.columns if col.startswith("HORARIO")]])

    # Salvar Excel em memória
    output = BytesIO()
    original_filename = uploaded_file.name.split(".")[0]
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet_name = f"{original_filename}_corrigido"[:31]
        new_df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)

    corrected_filename = f"{original_filename}_corrigido.xlsx"

    st.download_button(
        label="⬇️ Baixar arquivo corrigido",
        data=output,
        file_name=corrected_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
