import pandas as pd
import streamlit as st
from io import BytesIO
import os

st.title("Ajuste de Horários - Virada da Noite 🌙➡️☀️")

def excel_time_to_datetime(t):
    # Converte número decimal do Excel (fração do dia) para Timestamp datetime
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

                # Converte fração de dia (Excel) ou datetime
                if pd.api.types.is_float_dtype(valores):
                    t = valores.apply(excel_time_to_datetime)
                else:
                    t = pd.to_datetime(valores, errors='coerce')

                # Ajusta virada da noite
                has_night = (t.dt.hour >= 18).any()
                has_early = (t.dt.hour < 10).any()
                t_adj = t.mask(t.dt.hour < 10, t + pd.Timedelta(days=1)) if (has_night and has_early) else t

                # Mantém ORDEM original e atualiza apenas horários
                df.loc[mask_valid, col_horario] = t_adj.values

    # --- Preview Streamlit: mostra apenas HH:MM ---
    df_preview = df.copy()
    for col in df_preview.columns:
        if col.startswith("HORARIO"):
            df_preview[col] = pd.to_datetime(df_preview[col], errors='coerce').dt.strftime("%H:%M")

    st.subheader("✅ Dados ajustados")
    st.dataframe(df_preview.head())

    # --- Preparar download Excel: mantém formato hora ---
    for col in df.columns:
        if col.startswith("HORARIO"):
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.time

    output = BytesIO()
    original_name = uploaded_file.name
    name, ext = os.path.splitext(original_name)
    novo_nome = f"{name}_ajustado.xlsx"

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Ajustado")
        workbook = writer.book
        worksheet = writer.sheets["Ajustado"]

        # Formata colunas HORARIO como hh:mm
        for i, col in enumerate(df.columns):
            if col.startswith("HORARIO"):
                worksheet.set_column(i, i, 8, workbook.add_format({"num_format": "hh:mm"}))

    output.seek(0)
    st.success("✅ Ajuste concluído!")
    st.download_button(
        label="📥 Baixar planilha ajustada",
        data=output,
        file_name=novo_nome,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
