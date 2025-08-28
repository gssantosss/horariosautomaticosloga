import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

def parse_horario(h):
    """Converte valores de hora (string, datetime, float do Excel) em datetime,
    considerando pós-meia-noite como dia seguinte"""
    if pd.isna(h):
        return None
    try:
        if isinstance(h, datetime):  # já é datetime
            t = h
        elif hasattr(h, "hour") and hasattr(h, "minute"):  # datetime.time
            t = datetime(1900, 1, 1, h.hour, h.minute)
        elif isinstance(h, (int, float)):  # número do Excel (fração do dia)
            total_min = int(round(h * 24 * 60))
            horas, minutos = divmod(total_min, 60)
            t = datetime(1900, 1, 1, horas % 24, minutos % 60)
        else:  # string "HH:MM"
            h_str = str(h).strip()
            t = datetime.strptime(h_str, "%H:%M")

        # Se for depois da meia-noite até 05:59, joga pro "dia seguinte"
        if t.hour < 6:
            t = t + timedelta(days=1)
        return t
    except:
        return None

# Streamlit
st.title("⏰ Ordenador de Horários")

uploaded_file = st.file_uploader("Envie sua planilha", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Selecionar colunas de horário
    horario_cols = [col for col in df.columns if col.startswith("HORARIO")]

    # Copiar dataframe original
    df_ordenado = df.copy()

    for col in horario_cols:
        parsed = df[col].apply(parse_horario)

        # Ordenar mantendo NaN no final
        sorted_vals = parsed.dropna().sort_values().reset_index(drop=True)

        # Recolocar no mesmo tamanho da coluna original
        reordered = pd.Series(
            [val.strftime("%H:%M") if val else None for val in sorted_vals]
        )
        reordered = reordered.reindex(range(len(df)))  # garante mesmo tamanho
        df_ordenado[col] = reordered

    st.write("### Planilha com colunas de horário ordenadas")
    st.dataframe(df_ordenado)

    # Botão de download em Excel
    towrite = io.BytesIO()
    df_ordenado.to_excel(towrite, index=False)
    towrite.seek(0)
    st.download_button(
        label="📥 Baixar Excel ordenado",
        data=towrite,
        file_name="planilha_ordenada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
