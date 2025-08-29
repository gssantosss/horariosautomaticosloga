import pandas as pd
import streamlit as st
from io import BytesIO
import os

st.title("Ajuste de Horários - Virada da Noite 🌙➡️☀️")

uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    st.write("📋 Planilha original carregada:")
    st.dataframe(df.head())

    dias = ["SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]

    for dia in dias:
        col_horario = f"HORARIO{dia}"
        col_ordem = f"ORDEM{dia}"

        if col_horario in df.columns and col_ordem in df.columns:
            mask_valid = df[col_horario].notna() & df[col_ordem].notna()
            if mask_valid.any():
                # Converter para datetime para cálculo
                t = pd.to_datetime(df.loc[mask_valid, col_horario], format="%H:%M", errors="coerce")

                # Regra da virada
                has_night = (t.dt.hour >= 18).any()
                has_early = (t.dt.hour < 10).any()
                t_adj = t.mask(t.dt.hour < 10, t + pd.Timedelta(days=1)) if (has_night and has_early) else t

                # Ordenar horários
                sorted_times = t_adj.sort_values().reset_index(drop=True)

                # Criar mapa ORDEM passo3 -> horário datetime
                ordem_passo3 = range(1, len(sorted_times) + 1)
                mapa_horario = dict(zip(ordem_passo3, sorted_times))

                # Reatribuir mantendo ordem original
                df.loc[mask_valid, col_horario] = df.loc[mask_valid, col_ordem].map(mapa_horario)

    # Preparar download usando datetime_format para Excel hh:mm
    output = BytesIO()
    original_name = uploaded_file.name
    name, ext = os.path.splitext(original_name)
    novo_nome = f"{name}_ajustado.xlsx"

    with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='hh:mm') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    st.success("✅ Ajuste concluído!")
    st.download_button(
        label="⬇️ Baixar planilha ajustada",
        data=output,
        file_name=novo_nome,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
