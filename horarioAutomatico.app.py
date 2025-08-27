import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import os

st.set_page_config(page_title="Correção Automática de Horários", layout="wide")
st.title("🕒 Correção Automática de Horários (mantém ordem e gaps)")

uploaded_file = st.file_uploader("Carregue sua planilha Excel (.xlsx)", type=["xlsx"])
limite_gap = st.number_input("Considerar pausa a partir de (minutos)", min_value=1, value=10, step=1)

def parse_time_to_dt(val, base_date):
    """Tenta converter valor (string/time/datetime) para datetime com data base_date."""
    if pd.isna(val):
        return pd.NaT
    # Usa pandas para parser flexível, em seguida normaliza para base_date com a hora correta
    parsed = pd.to_datetime(str(val), errors="coerce")
    if pd.isna(parsed):
        return pd.NaT
    t = parsed.time()
    return datetime.combine(base_date, t)

def process_pair(df, ordem_col, horario_col, gap_min):
    """
    Processa uma única coluna (ordem_col, horario_col) retornando os valores corrigidos
    para as posições onde existiam dados (não cria/remover linhas).
    """
    # Seleciona apenas linhas onde há ordens e horários (evita mexer em linhas vazias)
    mask = df[ordem_col].notna() & df[horario_col].notna()
    if not mask.any():
        return  # nada a processar

    sub = df.loc[mask, [ordem_col, horario_col]].copy()
    # Garante que ordem é numérico para ordenar corretamente
    sub[ordem_col] = pd.to_numeric(sub[ordem_col], errors="coerce")

    # Ordena pela ordem (A ORDEM É A BASE, nunca alteramos ela)
    sub = sub.sort_values(by=ordem_col)
    original_indices = sub.index.to_list()

    # Parser seguro dos horários para datetimes com data base
    base_date = datetime(1900, 1, 1).date()
    times = [parse_time_to_dt(v, base_date) for v in sub[horario_col].tolist()]

    # Ajusta virada de dia: garante sequência não-decrecente ao longo da ordem,
    # somando dias quando um horário parecer "voltar" (ex: 23:00 -> 01:30)
    for i in range(1, len(times)):
        if pd.isna(times[i]) or pd.isna(times[i-1]):
            continue
        # enquanto atual <= anterior, soma 1 dia (normalmente 1 soma já resolve)
        while times[i] <= times[i-1]:
            times[i] = times[i] + timedelta(days=1)

    # Calcula diferenças em minutos entre vizinhos (em ordem)
    diffs_min = []
    for i in range(len(times) - 1):
        if pd.isna(times[i]) or pd.isna(times[i+1]):
            diffs_min.append(float('inf'))
        else:
            diffs_min.append((times[i+1] - times[i]).total_seconds() / 60.0)

    # Particiona em blocos onde diffs >= gap_min (gaps respeitados)
    blocks = []
    start = 0
    for i, d in enumerate(diffs_min):
        if d >= gap_min:
            # bloco vai de start .. i (inclusive)
            blocks.append((start, i))
            start = i + 1
    # último bloco
    blocks.append((start, len(times) - 1))

    # Para cada bloco, interpola interior entre limites (mantendo limites fixos).
    # Se bloco tem 1 ou 2 pontos, não muda nada.
    corrected = times.copy()
    for (s, e) in blocks:
        # tamanho do bloco = e - s + 1
        if s >= e:
            continue
        if pd.isna(times[s]) or pd.isna(times[e]):
            # se extremos inválidos, pula (não tenta inventar)
            continue
        n = e - s
        if n < 2:
            continue  # só 2 pontos -> nada a interpolar
        total_span = times[e] - times[s]
        # intervalo base entre pontos do bloco
        step = total_span / n
        # atribui interior (mantendo extremos como estavam)
        for k in range(1, n):
            corrected[s + k] = times[s] + step * k

    # Por segurança: se algum horário computado ficou menor que um anterior (por float err),
    # corrigimos garantindo monotonicidade estrita não-decrescente (só soma micros se necessário)
    for i in range(1, len(corrected)):
        if pd.isna(corrected[i]) or pd.isna(corrected[i-1]):
            continue
        if corrected[i] <= corrected[i-1]:
            # garante ao menos +1 segundo
            corrected[i] = corrected[i-1] + timedelta(seconds=1)

    # Converte de volta para string HH:MM e escreve nos índices originais
    for idx, dtval in zip(original_indices, corrected):
        if pd.isna(dtval):
            # Se por alguma razão era inválido, deixa como estava
            continue
        # formata só hora (se passou da meia-noite e subiu dia, %H:%M mostra hora correta)
        df.at[idx, horario_col] = dtval.strftime("%H:%M")

# ---- UI / fluxo ----
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Erro ao ler o Excel: {e}")
        st.stop()

    st.subheader("Pré-visualização (não será alterada a estrutura)")
    st.dataframe(df.head())

    # Detecta automaticamente coluna de setor se houver (opcional)
    group_col = None
    candidate_cols = [c for c in df.columns if "SETOR" in c.upper()]
    if candidate_cols:
        group_col = candidate_cols[0]  # usa a primeira que encontrar

    dias = ["SEG", "TER", "QUA", "QUI", "SEX", "SAB"]

    # Faz cópia para modificar
    new_df = df.copy()

    # Processa por dia e por setor (se existir)
    for dia in dias:
        ordem_col = f"ORDEM{dia}"
        horario_col = f"HORARIO{dia}"
        if ordem_col not in new_df.columns or horario_col not in new_df.columns:
            continue

        if group_col:
            for gval in new_df[group_col].unique():
                mask = (new_df[group_col] == gval) & new_df[ordem_col].notna() & new_df[horario_col].notna()
                if mask.sum() == 0:
                    continue
                subset_idx = new_df.loc[mask].index
                # chama process_pair só no subconjunto (fazendo slice temporário)
                subdf = new_df.loc[subset_idx, [ordem_col, horario_col]].copy()
                # Para facilitar, monta um temp df com reindex ao chamar function
                temp_full = new_df.loc[subset_idx].copy()
                process_pair(temp_full, ordem_col, horario_col, limite_gap)
                # aplica resultados de volta
                new_df.loc[subset_idx, horario_col] = temp_full.loc[:, horario_col]
        else:
            # sem setor: processa toda a coluna
            mask = new_df[ordem_col].notna() & new_df[horario_col].notna()
            if mask.sum() == 0:
                continue
            temp_full = new_df.loc[mask].copy()
            process_pair(temp_full, ordem_col, horario_col, limite_gap)
            new_df.loc[mask, horario_col] = temp_full.loc[:, horario_col]

    st.subheader("Preview — Horários ajustados (somente valores de HORARIOxxx são alterados)")
    st.dataframe(new_df.head())

    # Prepara download (mesmo nome do arquivo + _corrigido.xlsx)
    original_filename = os.path.splitext(uploaded_file.name)[0]
    out_name = f"{original_filename}_corrigido.xlsx"
    # aba com nome curto (máx 31 chars)
    sheet_name = (original_filename + "_corrigida")[:31]

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        new_df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)

    st.download_button(
        label="⬇️ Baixar planilha corrigida",
        data=output.getvalue(),
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Pronto — horários ajustados conforme regras. Lembra: ordem das ruas não foi alterada.")
