import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# Upload do arquivo
uploaded_file = st.file_uploader("Envie sua planilha", type=["xlsx"])

if uploaded_file is not None:
    # Ler Excel
    df = pd.read_excel(uploaded_file)

    # Selecionar apenas colunas de horário
    horario_cols = [col for col in df.columns if col.startswith("HORARIO")]

    # Criar um DataFrame novo só com as colunas de horários ordenadas
    horarios_ordenados = pd.DataFrame()

    for col in horario_cols:
        # Pega valores não nulos
        horarios_raw = df[col].dropna().astype(str).tolist()

        horarios = []
        for h in horarios_raw:
            try:
                t = datetime.strptime(h.strip(), "%H:%M")
                # se for depois da meia-noite, joga pro "dia seguinte"
                if t.hour < 6:
                    t = t + timedelta(days=1)
                horarios.append(t)
            except:
                pass

        # Ordena
        horarios_sorted = sorted(horarios)

        # Converte de volta pro formato hh:mm
        horarios_fmt = [dt.strftime("%H:%M") for dt in horarios_sorted]

        # Salva no dataframe novo
        horarios_ordenados[col] = pd.Series(horarios_fmt)

    st.write("### Colunas de horários ordenadas individualmente")
    st.dataframe(horarios_ordenados)
