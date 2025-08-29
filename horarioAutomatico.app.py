import pandas as pd
import streamlit as st
from io import BytesIO
import os

st.title("Ajuste de Horários - Virada da Noite 🌙➡️☀️")

# Converte número excel (fração do dia) para Timestamp com base 1900-01-01
def excel_time_to_datetime_series(s):
    # s: Series numeric (fração do dia)
    return pd.to_timedelta(s.astype(float), unit="d") + pd.Timestamp("1900-01-01")

uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])
if uploaded_file is not None:
    # lê mantendo inferência normal
    df = pd.read_excel(uploaded_file)

    st.write("📋 Planilha original carregada:")
    st.dataframe(df.head())

    dias = ["SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]
    original_cols = df.columns.tolist()

    for dia in dias:
        col_horario = f"HORARIO{dia}"
        col_ordem = f"ORDEM{dia}"

        if col_horario in df.columns and col_ordem in df.columns:
            # linhas com ambos preenchidos
            mask_valid = df[col_horario].notna() & df[col_ordem].notna()
            if mask_valid.any():
                valores = df.loc[mask_valid, col_horario]

                # --- converter para datetime corretamente dependendo do dtype ---
                if pd.api.types.is_numeric_dtype(valores):
                    # caso Excel tenha salvo como fração do dia (float)
                    t = excel_time_to_datetime_series(valores)
                elif pd.api.types.is_datetime64_any_dtype(valores):
                    t = pd.to_datetime(valores)
                else:
                    # strings ou objetos: tenta parsear como HH:MM (mais seguro)
                    t = pd.to_datetime(valores.astype(str), format="%H:%M", errors="coerce")

                # --- regra da virada ---
                has_night = (t.dt.hour >= 18).any()
                has_early = (t.dt.hour < 10).any()
                if has_night and has_early:
                    t_adj = t.where(~(t.dt.hour < 10), t + pd.Timedelta(days=1))
                else:
                    t_adj = t

                # --- ordenar horários ajustados (passo 3) ---
                sorted_times = t_adj.sort_values().reset_index(drop=True)  # Series de Timestamps

                # --- criar mapa 1..n -> horário (Timestamp) ---
                mapa_horario = dict(zip(range(1, len(sorted_times) + 1), sorted_times))

                # --- mapear usando os valores da coluna ORDEM (garantir int) ---
                # converte a coluna ORDEM das linhas válidas para int (ex: '2'|'2.0' -> 2)
                ord_series = pd.to_numeric(df.loc[mask_valid, col_ordem], errors="coerce").astype(int)
                df.loc[mask_valid, col_horario] = ord_series.map(mapa_horario)
                # fim do processamento do dia

    # garantir que as colunas HORARIO* estejam em dtype datetime64 (para o ExcelWriter formatar)
    for dia in dias:
        col_horario = f"HORARIO{dia}"
        if col_horario in df.columns:
            df[col_horario] = pd.to_datetime(df[col_horario], errors="coerce")

    # manter ordem original das colunas religiosamente
    df = df[original_cols]

    st.dataframe(df.head())
    
    # preparar o arquivo para download com formatação hh:mm
    output = BytesIO()
    original_name = uploaded_file.name
    name, ext = os.path.splitext(original_name)
    novo_nome = f"{name}_ajustado.xlsx"

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="hh:mm") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    st.success("✅ Ajuste concluído!")
    st.download_button(
        label="⬇️ Baixar planilha ajustada",
        data=output,
        file_name=novo_nome,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

