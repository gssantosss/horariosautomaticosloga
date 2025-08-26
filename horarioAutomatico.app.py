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

    # Lista de dias da semana que aparecem na planilha
    dias = ["SEG", "TER", "QUA", "QUI", "SEX", "SAB"]

    # Função para corrigir horários dentro de cada grupo de ordem
    def corrigir_horarios(sub_df):
        horarios = sub_df.copy()
        horarios = horarios.sort_values("ordem")
        horarios["horario_corrigido"] = horarios["horario"].ffill()
        return horarios

    # Loop pelos dias e corrigir
    for dia in dias:
        ordem_col = f"ORDEM{dia}"
        horario_col = f"HORARIO{dia}"

        if ordem_col in df.columns and horario_col in df.columns:
            temp = df[[ordem_col, horario_col]].rename(
                columns={ordem_col: "ordem", horario_col: "horario"}
            )
            corrigido = corrigir_horarios(temp)
            df[horario_col] = corrigido["horario_corrigido"]

    # Salvar em memória o Excel corrigido
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # 👇 nome da aba agora é "original_corrigida"
        df.to_excel(writer, index=False, sheet_name=f"{original_filename}_corrigida")
    output.seek(0)

    # Nome final do arquivo
    corrected_filename = f"{original_filename}_corrigido.xlsx"

    # Botão de download
    st.download_button(
        label="⬇️ Baixar arquivo corrigido",
        data=output,
        file_name=corrected_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",)
