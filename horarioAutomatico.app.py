import pandas as pd
import streamlit as st
from io import BytesIO
import os
from st_aggrid import AgGrid, GridOptionsBuilder

st.title("Ajuste de Horários - Virada da Noite 🌙➡️☀️")

def excel_time_to_datetime(t):
    # Converte fração de dia do Excel em Timestamp datetime
    return pd.to_timedelta(t, unit='d') + pd.Timestamp('1899-12-30')

uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    dias = ["SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]

    for dia in dias:
        col_horario = f"HORARIO{dia}"
        col_ordem = f"ORDEM{dia}"

        if col_horario in df.columns and col_ordem in df.columns:
            mask_valid = df[col_horario].notna() & df[col_ordem].notna()

            if mask_valid.any():
                valores = df.loc[mask_valid, col_horario]

                # Converte float (fração do dia) ou string para datetime
                if pd.api.types.is_float_dtype(valores):
                    t = valores.apply(excel_time_to_datetime)
                else:
                    t = pd.to_datetime(valores, errors='coerce')

                # Detecta virada da noite
                has_night = (t.dt.hour >= 18).any()
                has_early = (t.dt.hour < 10).any()

                t_adj = t.copy()
                if has_night and has_early:
                    t_adj[t.dt.hour < 10] += pd.Timedelta(days=1)

                # Ordena pelos horários ajustados
                aux = df.loc[mask_valid, [col_ordem]].copy()
                aux['horario_ajustado'] = t_adj.values
                aux = aux.sort_values('horario_ajustado').reset_index(drop=True)
                aux['nova_ordem'] = range(1, len(aux)+1)

                # Atualiza o horário original com horário ajustado
                df.loc[mask_valid, col_horario] = aux['horario_ajustado'].values

    # --- Preview com AgGrid ---
    df_preview = df.copy()
    
    # Converte todas as colunas HORARIO para datetime
    for dia in dias:
        col_horario = f"HORARIO{dia}"
        if col_horario in df_preview.columns:
            df_preview[col_horario] = pd.to_datetime(df_preview[col_horario], errors='coerce')

    gb = GridOptionsBuilder.from_dataframe(df_preview)
    # Configura todas as colunas HORARIO como tipo hora, com relógio
    for dia in dias:
        col_horario = f"HORARIO{dia}"
        if col_horario in df_preview.columns:
            gb.configure_column(
                col_horario, 
                type=["timeColumnFilter"], 
                valueFormatter="(params.value) ? new Date(params.value).toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'}) : ''"
            )
    gridOptions = gb.build()
    st.subheader("📊 Planilha ajustada (Preview com relógio):")
    AgGrid(df_preview, gridOptions=gridOptions)

    # DOWNLOAD: mantém datetime para Excel interpretar como hora
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
        label="⬇️ Baixar planilha ajustada",
        data=output,
        file_name=novo_nome,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
