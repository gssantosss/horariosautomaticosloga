import pandas as pd
import streamlit as st
from io import BytesIO
import os

st.title("Ajuste de Horários")

def excel_time_to_datetime(t):
    return pd.to_timedelta(t, unit='d') + pd.Timestamp('1899-12-30')

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
                valores = df.loc[mask_valid, col_horario]

                if pd.api.types.is_float_dtype(valores):
                    t = valores.apply(excel_time_to_datetime)
                else:
                    t = pd.to_datetime(valores, errors='coerce')

                has_night = (t.dt.hour >= 18).any()
                has_early = (t.dt.hour < 10).any()
                t_adj = t.mask(t.dt.hour < 10, t + pd.Timedelta(days=1)) if (has_night and has_early) else t

                aux = df.loc[mask_valid, [col_ordem]].copy()
                aux['horario_ajustado'] = t_adj.values
                aux = aux.sort_values('horario_ajustado').reset_index()
                aux['nova_ordem'] = range(1, len(aux) + 1)

                mapa_ordem_horario = dict(zip(aux['nova_ordem'], aux['horario_ajustado']))

                df.loc[mask_valid, col_horario] = df.loc[mask_valid, col_ordem].map(mapa_ordem_horario)

    # Converte colunas HORARIO para datetime.time (apenas hora:minuto:segundo)
    for dia in dias:
        col_horario = f"HORARIO{dia}"
        if col_horario in df.columns:
            df[col_horario] = pd.to_datetime(df[col_horario], errors='coerce').dt.time

    st.dataframe(df.head())
 
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
