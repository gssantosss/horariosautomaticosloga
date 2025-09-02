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

    # Processamento
    new_df = df.copy()
    dias = ["SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]

    for dia in dias:
        ordem_col = f"ORDEM{dia}"
        horario_col = f"HORARIO{dia}"

        if ordem_col in df.columns and horario_col in df.columns:
            mask_valid = df[ordem_col].notna() & df[horario_col].notna()
            if mask_valid.any():
                subset = df.loc[mask_valid, [ordem_col, horario_col]].copy()

                # Converte horários para datetime
                subset[horario_col] = pd.to_datetime(subset[horario_col].astype(str), errors='coerce')

                # Detecta virada da noite
                has_night = (subset[horario_col].dt.hour >= 18).any()
                has_early = (subset[horario_col].dt.hour < 10).any()
                if has_night and has_early:
                    subset.loc[subset[horario_col].dt.hour < 10, horario_col] += pd.Timedelta(days=1)

                # Ordena por horário ajustado
                subset = subset.sort_values(by=horario_col).reset_index()
                subset['nova_ordem'] = range(1, len(subset)+1)

                # Atualiza horário e ordem no dataframe final
                new_df.loc[subset['index'], horario_col] = subset[horario_col].values
                new_df.loc[subset['index'], ordem_col] = subset['nova_ordem'].values

    # --- Preview: transforma para HH:MM apenas para exibir no Streamlit ---
    df_preview = new_df.copy()
    for col in df_preview.columns:
        if col.startswith("HORARIO"):
            df_preview[col] = pd.to_datetime(df_preview[col], errors='coerce').dt.strftime("%H:%M")

    st.subheader("✅ Dados ajustados")
    st.dataframe(df_preview.head())

    # --- Download: converte para datetime.time pra Excel mostrar apenas hora ---
    for col in new_df.columns:
        if col.startswith("HORARIO"):
            new_df[col] = pd.to_datetime(new_df[col], errors='coerce').dt.time

    output = BytesIO()
    nome_arquivo = uploaded_file.name.replace(".xlsx", "_ajustada.xlsx")
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        new_df.to_excel(writer, index=False, sheet_name="Ajustado")
        workbook = writer.book
        worksheet = writer.sheets["Ajustado"]

        # Formata colunas HORARIO como hh:mm
        for i, col in enumerate(new_df.columns):
            if col.startswith("HORARIO"):
                worksheet.set_column(i, i, 8, workbook.add_format({"num_format": "hh:mm"}))

    output.seek(0)
    st.download_button(
        label="📥 Baixar planilha ajustada",
        data=output,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
