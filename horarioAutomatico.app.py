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

    # Verifica os nomes das colunas
    st.write("Nomes das colunas:", df.columns)

    # Verifica o conteúdo do DataFrame
    st.write("Conteúdo do DataFrame:", df.head())

    # Lista de dias da semana que aparecem na planilha
    dias = ["SEG", "TER", "QUA", "QUI", "SEX", "SAB"]

    # Função para corrigir horários dentro de cada grupo de ordem
    def corrigir_horarios(sub_df):
        horarios = sub_df.copy()
        horarios = horarios.sort_values("ordem")
        
        # Converte os horários para datetime
        horarios["horario"] = pd.to_datetime(horarios["horario"], format='%H:%M', errors='coerce')
        
        # Corrige os horários com base no limite de gap
        for i in range(1, len(horarios)):
            # Calcula a diferença em minutos
            gap = (horarios.iloc[i]["horario"] - horarios.iloc[i - 1]["horario"]).total_seconds() / 60
            
            # Se o gap for maior que o limite, ajusta o horário
            if gap > limite_gap:
                horarios.iloc[i]["horario"] = horarios.iloc[i - 1]["horario"] + pd.Timedelta(minutes=limite_gap)

        # Formata os horários de volta para string
        horarios["horario_corrigido"] = horarios["horario"].dt.strftime('%H:%M')
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
