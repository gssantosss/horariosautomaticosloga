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
            # Selecionar apenas linhas com dados
            mask_valid = df[col_horario].notna() & df[col_ordem].notna()
            if mask_valid.any():
                # 1) Garantir que a coluna de horário seja string
                horarios = df.loc[mask_valid, col_horario].astype(str)

                # 2) Converter para datetime apenas para cálculo
                t = pd.to_datetime(horarios, format="%H:%M", errors="coerce")

                # 3) Regra da virada
                has_night = (t.dt.hour >= 18).any()
                has_early = (t.dt.hour < 10).any()
                t_adj = t.mask(t.dt.hour < 10, t + pd.Timedelta(days=1)) if (has_night and has_early) else t

                # 4) Ordenar horários ajustados
                sorted_times = t_adj.sort_values().reset_index(drop=True)

                # 5) Criar mapa ORDEM passo3 -> horário formatado HH:MM
                ordem_passo3 = range(1, len(sorted_times) + 1)
                mapa_horario = dict(zip(ordem_passo3, sorted_times.dt.strftime("%H:%M")))

                # 6) Reatribuir horários mantendo a ordem original
                df.loc[mask_valid, col_horario] = df.loc[mask_valid, col_ordem].map(mapa_horario)

    # 7) Preparar download
    output = BytesIO()
    original_name = uploaded_file.name
    name, ext = os.path.splitext(original_name)
    novo_nome = f"{name}_ajustado.xlsx"
    df.to_excel(output, index=False)
    output.seek(0)

    st.success("✅ Ajuste concluído!")
    st.download_button(
        label="⬇️ Baixar planilha ajustada",
        data=output,
        file_name=novo_nome,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
