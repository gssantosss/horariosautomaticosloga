import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO
import os

st.set_page_config(page_title="Correção Automática de Horários", layout="wide")
st.title("🕒 Correção Automática de Horários (por ORDEM)")

uploaded_file = st.file_uploader("Carregue sua planilha (.xlsx)", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    pause_threshold = st.number_input(
        "Considerar pausa a partir de (minutos)", min_value=1, max_value=240, value=10, step=1
    )
with col2:
    overnight_cutoff_h = st.number_input(
        "Janela para considerar virada de madrugada (horas)",
        min_value=1, max_value=12, value=6, step=1,
        help="Se um horário cair para trás em mais de X horas, trata como madrugada (ex.: 23:50 → 00:30). Quedas menores são tratadas como erro e serão só 'encostadas' no anterior."
    )

def _to_time_series(s: pd.Series) -> pd.Series:
    """Converte qualquer formato de hora em pandas datetime (data fictícia)."""
    # Converte python time, strings 'HH:MM' / 'HH:MM:SS' e datetimes
    # Evita números (excel serial) propositalmente — seus dados são HH:MM / time
    return pd.to_datetime(s.astype(str), errors="coerce")

def normalize_over_midnight_by_order(times: pd.Series, overnight_cutoff_h: int) -> pd.Series:
    """
    Recebe uma série de datetimes (mesmo dia fictício), já na SEQUÊNCIA DA ORDEM,
    e normaliza somando +1 dia quando a queda for 'grande' (madrugada).
    Quedas pequenas (<= cutoff) são tratadas como bug (sem somar dia).
    Retorna timestamps não decrescentes, exceto bugs que serão ajustados depois.
    """
    base_date = pd.Timestamp("2000-01-01")
    day_offset = 0
    out = []
    prev_tod = None

    for t in times:
        if pd.isna(t):
            out.append(pd.NaT)
            prev_tod = None
            continue

        tod = pd.Timestamp(t).time()
        cur = pd.Timestamp.combine((base_date + pd.Timedelta(days=day_offset)).date(), tod)

        if out and prev_tod is not None:
            # queda (hora do dia menor que a anterior)?
            if tod < prev_tod:
                drop = (pd.Timestamp.combine(base_date.date(), prev_tod)
                        - pd.Timestamp.combine(base_date.date(), tod))
                if drop >= pd.Timedelta(hours=overnight_cutoff_h):
                    # é madrugada → soma um dia
                    day_offset += 1
                    cur = cur + pd.Timedelta(days=1)
                # se não, é bug → não soma dia (vamos tratar no clamp)
        out.append(cur)
        prev_tod = tod

    return pd.Series(out, index=times.index)

def clamp_monotonic_and_fix_extremes(times_norm: pd.Series) -> pd.Series:
    """
    Garante: 1º = menor horário real, último = maior horário real.
    Sequência não decrescente (clamp: cada item >= anterior).
    Não inventa intervalos; gaps permanecem do jeito que vieram.
    """
    if times_norm.empty:
        return times_norm

    tmin = times_norm.min()
    tmax = times_norm.max()

    adj = times_norm.copy()

    # força primeiro = mínimo real
    adj.iloc[0] = tmin

    # forward clamp (não deixa voltar no tempo)
    for i in range(1, len(adj)):
        if pd.isna(adj.iloc[i]):
            adj.iloc[i] = adj.iloc[i-1]
        elif adj.iloc[i] < adj.iloc[i-1]:
            adj.iloc[i] = adj.iloc[i-1]

    # força último = máximo real (último horário real do setor)
    adj.iloc[-1] = tmax

    # garante não decrescente até o fim (caso raro de NaT intermediário)
    for i in range(1, len(adj)):
        if adj.iloc[i] < adj.iloc[i-1]:
            adj.iloc[i] = adj.iloc[i-1]

    return adj

def adjust_day(df: pd.DataFrame, ordem_col: str, horario_col: str, overnight_cutoff_h: int) -> pd.Series:
    """
    Ajusta apenas a coluna de horário de um dia específico,
    respeitando a coluna ORDEM correspondente. NÃO mexe na coluna de ordem.
    Retorna a série formatada '%H:%M' para atribuição no DF original.
    """
    # Subset com índices originais (não reordenamos o DF global)
    subset = df[[ordem_col, horario_col]].copy()

    # Série de horários em datetime (hora do dia), sem mexer em coluna original
    times_raw = _to_time_series(subset[horario_col])

    # Ordena VIRTUALMENTE por ordem pra calcular sequência
    order_sorted_idx = subset.sort_values(by=ordem_col).index
    times_in_order = times_raw.loc[order_sorted_idx].reset_index(drop=True)

    # Normaliza madrugada
    times_norm = normalize_over_midnight_by_order(times_in_order, overnight_cutoff_h)

    # Ajusta monotonia e extremos
    times_adj = clamp_monotonic_and_fix_extremes(times_norm)

    # IMPORTANTE: não mexer nos gaps — nada de interpolar; só clamp em bugs.
    # (pause_threshold é só um parâmetro de detecção, aqui mantemos tudo)

    # Volta pros mesmos índices, agora apenas substituindo os horários
    # Formata como HH:MM
    times_adj_fmt = times_adj.dt.strftime("%H:%M")
    # Remonta na mesma ordem_sorted_idx
    out = pd.Series(index=order_sorted_idx, data=times_adj_fmt.values)
    # Reordena pra alinhar com o DF original (mesmos índices)
    out = out.reindex(subset.index)

    return out

if uploaded_file:
    original_filename = os.path.splitext(uploaded_file.name)[0]
    df = pd.read_excel(uploaded_file)

    st.subheader("Prévia dos dados (originais)")
    st.dataframe(df.head())

    new_df = df.copy()

    # Percorre TODOS os pares ORDEM/HORARIO existentes
    for col in df.columns:
        if not col.startswith("ORDEM"):
            continue
        dia = col.replace("ORDEM", "")
        ordem_col = col
        horario_col = f"HORARIO{dia}"

        if horario_col not in df.columns:
            continue  # se não houver par, pula

        # Ajusta apenas os horários desse dia
        ajustados = adjust_day(df, ordem_col, horario_col, overnight_cutoff_h)
        # Aplica no DF final (mesma posição de cada linha, sem reordenar nada)
        new_df.loc[ajustados.index, horario_col] = ajustados.values

    st.subheader("✅ Dados ajustados (sem mudar estrutura)")
    st.dataframe(new_df.head())

    # Salvar para download usando openpyxl
    output = BytesIO()
    sheet_name = f"{original_filename}_corrigida"
    sheet_name = sheet_name[:31]  # limite do Excel

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        new_df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)

    corrected_filename = f"{original_filename}_corrigido.xlsx"
    st.download_button(
        "📥 Baixar arquivo corrigido",
        data=output,
        file_name=corrected_filename,
        mime="application/vnd.openxmlformats-officed
