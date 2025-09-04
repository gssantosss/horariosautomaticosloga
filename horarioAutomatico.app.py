# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from typing import Optional

# ------------------------------------------------------------
# Configuração da página
# ------------------------------------------------------------
st.set_page_config(
    page_title="Normalizador de Roteiro (HORARIO/ORDEM)",
    layout="wide",
)

DIAS = ['SEG','TER','QUA','QUI','SEX','SAB','DOM']
CAT_DIAS = pd.api.types.CategoricalDtype(categories=DIAS, ordered=True)

# ------------------------------------------------------------
# Utilitários
# ------------------------------------------------------------
def to_hhmm(x) -> str:
    """Converte 'HH:MM[:SS]' ou datetime/time -> 'hh:mm' (texto). Vazio se inválido/ausente."""
    if x is None:
        return ""
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return ""
    m = re.match(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?$', s)
    if m:
        hh = int(m.group(1)); mm = int(m.group(2))
        if 0 <= hh <= 23 and 0 <= mm <= 59:
            return f"{hh:02d}:{mm:02d}"
    try:
        t = pd.to_datetime(s, errors='raise').time()
        return f"{t.hour:02d}:{t.minute:02d}"
    except Exception:
        return ""

def valor_unico_ou_multiplos(df: pd.DataFrame, col: str) -> str:
    """Retorna o valor único da coluna (se houver), 'múltiplos' se >1, ou '—' se vazio/inexistente."""
    if df is None or not isinstance(df, pd.DataFrame) or col not in df.columns:
        return "—"
    vals = (
        df[col]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: s.ne("")]
        .unique()
        .tolist()
    )
    if len(vals) == 0:
        return "—"
    if len(vals) == 1:
        return vals[0]
    return "múltiplos"

def nome_setor(df_raw: pd.DataFrame, uploaded_name: Optional[str] = None) -> str:
    """Obtém o nome do setor pela coluna SETOR, ou tenta extrair algo tipo 'PR18' do nome do arquivo."""
    setor = valor_unico_ou_multiplos(df_raw, 'SETOR')
    if setor not in ("—", "múltiplos"):
        return setor
    if uploaded_name:
        m = re.search(r'\b([A-Z]{2}\d{1,3})\b', uploaded_name.upper())
        if m:
            return m.group(1)
    return setor

def selecionar_aba_dados(xls: pd.ExcelFile) -> str:
    """
    Seleciona automaticamente a aba que contenha qualquer coluna ORDEM* ou HORARIO*.
    Se não encontrar, retorna a primeira aba.
    """
    for sh in xls.sheet_names:
        try:
            header_df = pd.read_excel(xls, sheet_name=sh, nrows=0)
            cols_upper = [str(c).upper() for c in header_df.columns]
            if any(c.startswith("ORDEM") or c.startswith("HORARIO") for c in cols_upper):
                return sh
        except Exception:
            continue
    return xls.sheet_names[0]

def montar_excel_somente_agenda(agenda: pd.DataFrame) -> bytes:
    """Monta o Excel de saída, apenas com a aba 'agenda_por_dia'."""
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as xw:
        agenda.to_excel(xw, sheet_name='agenda_por_dia', index=False)
    bio.seek(0)
    return bio.read()

def construir_tabelas_por_dia(df_raw: pd.DataFrame) -> dict:
    """ 
    Monta tabelas por dia contendo registros com HORARIO preenchido:
    - HORARIO<dia> (texto hh:mm), ORDEM<dia> (Int64), OBS<dia> (vazio)
    - Ordena por HORARIO<dia> (crescente)
    - Retorna um dicionário { 'SEG': df_seg, 'TER': df_ter, ... } para todos os dias com dados
    """
    tabelas = {}
    for dia in DIAS:
        hcol = f'HORARIO{dia}'
        ocol = f'ORDEM{dia}'

        if hcol not in df_raw.columns and ocol not in df_raw.columns:
            continue

        ser_h = df_raw[hcol] if hcol in df_raw.columns else pd.Series([pd.NA]*len(df_raw), index=df_raw.index)
        ser_o = df_raw[ocol] if ocol in df_raw.columns else pd.Series([pd.NA]*len(df_raw), index=df_raw.index)

        horarios = ser_h.apply(to_hhmm)
        ordens = pd.to_numeric(ser_o, errors='coerce').astype('Int64')

        # Inclui linhas com HORARIO preenchido
        mask_com_horario = horarios.ne('')
        if not mask_com_horario.any():
            continue

        df_dia = pd.DataFrame({
            f'HORARIO{dia}': horarios[mask_com_horario].values,
            f'ORDEM{dia}'  : ordens[mask_com_horario].values,
        })

        df_dia[f'OBS{dia}'] = ''
        df_dia.sort_values(by=[f'HORARIO{dia}', f'ORDEM{dia}'], inplace=True, kind='stable')
        tabelas[dia] = df_dia.reset_index(drop=True)

    return tabelas
# ------------------------------------------------------------
# Processamento principal (normalização) - sem alterar df_raw
# ------------------------------------------------------------
def processar_df_sem_mutar(df: pd.DataFrame):
    """
    Constrói a 'agenda' (formato longo) com linhas COMPLETAS (HORARIO + ORDEM),
    sem adicionar colunas ao df original.
    Retorna apenas a 'agenda' (para manter tudo enxuto).
    """
    base_cols = [c for c in [
        'ID','SETOR','TIPOCOLETA','FREQUENCIA','TURNO',
        'TIPO','TITULO','PREPOSICAO','LOGRADOURO',
        'INICIO','FIM','DISTRITO','SUBPREFEITURA'
    ] if c in df.columns]

    linhas_ok = []
    n = len(df)

    for dia in DIAS:
        hcol, ocol, fcol = f'HORARIO{dia}', f'ORDEM{dia}', f'FORMACOLETA{dia}'

        # Séries independentes (não alteram df_raw)
        ser_h = df[hcol] if hcol in df.columns else pd.Series([pd.NA]*n, index=df.index)
        ser_o = df[ocol] if ocol in df.columns else pd.Series([pd.NA]*n, index=df.index)
        ser_f = df[fcol] if fcol in df.columns else pd.Series([pd.NA]*n, index=df.index)

        bloco = df[base_cols].copy()
        bloco['HORARIO'] = ser_h.apply(to_hhmm)                         # texto hh:mm
        bloco['ORDEM'] = pd.to_numeric(ser_o, errors='coerce').astype('Int64')
        bloco['FORMA_COLETA'] = pd.Series(ser_f, index=df.index).astype('string').str.strip().fillna('')
        bloco['DIA_SEMANA'] = dia

        has_h = bloco['HORARIO'].ne('')
        has_o = bloco['ORDEM'].notna()
        completo = has_h & has_o

        if completo.any():
            linhas_ok.append(bloco.loc[completo])

    agenda = pd.concat(linhas_ok, ignore_index=True) if linhas_ok else pd.DataFrame()

    # Ordenação final
    if not agenda.empty:
        if 'DIA_SEMANA' in agenda.columns:
            agenda['DIA_SEMANA'] = agenda['DIA_SEMANA'].astype(CAT_DIAS)
        sort_cols = [c for c in ['ID', 'DIA_SEMANA', 'ORDEM'] if c in agenda.columns]
        agenda.sort_values(by=sort_cols, inplace=True, kind='stable')

    return agenda

# ------------------------------------------------------------
# Cálculo: Qtde. de Pontos (1 por linha com qualquer ORDEM* preenchida)
# ------------------------------------------------------------
def calcular_qtde_pontos(df_raw: pd.DataFrame) -> int:
    ordem_cols = [f'ORDEM{d}' for d in DIAS if f'ORDEM{d}' in df_raw.columns]
    if not ordem_cols:
        return 0
    mask = pd.Series(False, index=df_raw.index)
    for col in ordem_cols:
        s = pd.to_numeric(df_raw[col], errors='coerce').notna()
        mask = mask | s
    return int(mask.sum())

# ------------------------------------------------------------
# Mini tabela: menor/maior horário por coluna HORARIO*
# ------------------------------------------------------------
def tabela_min_max_horarios(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Retorna uma tabela com: Coluna, Menor horário, Maior horário
    (ignora colunas HORARIO* que não possuam nenhum valor válido).
    Não modifica df_raw.
    """
    from typing import Optional

    def to_minutes(v) -> Optional[int]:
        if pd.isna(v):
            return None
        s = str(v).strip()
        if not s or s.lower() == 'nan':
            return None
        # tenta parse geral (aceita HH:MM:SS, datetime etc.)
        try:
            t = pd.to_datetime(s, errors='raise').time()
            return t.hour*60 + t.minute
        except Exception:
            m = re.match(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?$', s)
            if m:
                hh = int(m.group(1)); mm = int(m.group(2))
                if 0 <= hh <= 23 and 0 <= mm <= 59:
                    return hh*60 + mm
        return None

    hor_cols = [c for c in df_raw.columns if str(c).upper().startswith('HORARIO')]
    out = []
    for col in hor_cols:
        mins = [to_minutes(v) for v in df_raw[col].tolist()]
        mins = [m for m in mins if m is not None]  # só válidos
        if mins:
            mi, ma = min(mins), max(mins)
            out.append({
                "Coluna": col,
                "Menor horário": f"{mi//60:02d}:{mi%60:02d}",
                "Maior horário": f"{ma//60:02d}:{ma%60:02d}",
            })

    # mantém somente colunas HORARIO* com pelo menos um valor válido
    return pd.DataFrame(out)

# ------------------------------------------------------------
# Mini painel (apenas métricas)
# ------------------------------------------------------------
def render_mini_painel(df_raw: pd.DataFrame, agenda: pd.DataFrame, uploaded_name: Optional[str]):
    qt_pontos      = calcular_qtde_pontos(df_raw)
    setor_nome     = nome_setor(df_raw, uploaded_name)
    subprefeitura  = valor_unico_ou_multiplos(df_raw, 'SUBPREFEITURA')
    # Frequência exibida: prioriza detectada pela agenda; se vazio, usa coluna FREQUENCIA (se única)
    if not agenda.empty and 'DIA_SEMANA' in agenda.columns:
        freq_detectada = '/'.join([d for d in DIAS if agenda['DIA_SEMANA'].astype(str).eq(d).any()])
    else:
        freq_detectada = ''
    frequencia_exb = freq_detectada if freq_detectada.strip() else valor_unico_ou_multiplos(df_raw, 'FREQUENCIA')
    turno          = valor_unico_ou_multiplos(df_raw, 'TURNO')
    tipo_coleta    = valor_unico_ou_multiplos(df_raw, 'TIPOCOLETA')

    st.markdown("### 🔎 Visão geral do setor")
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("Qtde. de Pontos", f"{qt_pontos:,}".replace(",", "."))
    with c2: st.metric("Setor", setor_nome)
    with c3: st.metric("Subprefeitura", subprefeitura)

    c4, c5, c6 = st.columns(3)
    with c4: st.metric("Frequência", frequencia_exb if frequencia_exb != "—" else "")
    with c5: st.metric("Turno", turno if turno != "—" else "")
    with c6: st.metric("Tipo de coleta", tipo_coleta if tipo_coleta != "—" else "")

# ------------------------------------------------------------
# UI (enxuta): sem prévia, sem seletor de aba; inclui mini tabela de horários
# ------------------------------------------------------------
# === Bloco principal da UI (substitua o seu trecho por este) ===

st.title("Normalizador de Roteiro por Dia (HORARIO/ORDEM)")
st.caption("Faça upload da planilha (.xlsx) do setor. O app usa automaticamente a aba com colunas HORARIO*/ORDEM*. Interface limpa, sem prévias.")

uploaded_file = st.file_uploader("Selecione a planilha do setor (formato .xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 1) Carregar a aba correta e o df_raw
        xls = pd.ExcelFile(uploaded_file)
        aba_dados = selecionar_aba_dados(xls)  # escolhe automaticamente a aba de dados
        df_raw = pd.read_excel(uploaded_file, sheet_name=aba_dados)

        # 2) Processamento (sem alterar df_raw; sem colunas extras)
        agenda = processar_df_sem_mutar(df_raw)

        # 3) Painel principal (métricas)
        st.markdown("---")
        render_mini_painel(df_raw, agenda, getattr(uploaded_file, 'name', None))

        # 4) Mini tabela: menor/maior horário por coluna HORARIO*
        st.markdown("### ⏱️ Faixa de horários por coluna (HORARIO*)")
        tabela_h = tabela_min_max_horarios(df_raw)
        st.dataframe(tabela_h, use_container_width=True, hide_index=True)

        # 5) Prévia completa por dia (somente válidos)
        st.markdown("### 📋 Prévia por dia (somente horários e ordens válidos)")
        tabelas_por_dia = construir_tabelas_por_dia(df_raw)

        if tabelas_por_dia:
            for dia in DIAS:
                if dia in tabelas_por_dia:
                    st.markdown(f"**{dia}**")
                    st.dataframe(
                        tabelas_por_dia[dia],
                        use_container_width=True,
                        hide_index=True
                    )

        # (Download removido por enquanto, conforme combinado)

    except Exception as e:
        st.exception(e)
        st.error("Erro ao processar a prévia. Verifique o arquivo e o layout (HORARIO*/ORDEM*).")
else:
    st.info("👉 Faça o upload de um arquivo .xlsx para começar.")
