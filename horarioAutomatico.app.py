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
    # tenta outras formas (ex.: objetos datetime, números serializados etc.)
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

def frequencia_para_exibir(meta_df: Optional[pd.DataFrame], df_raw: pd.DataFrame) -> str:
    """Prioriza a frequência detectada; se vazia, usa a coluna FREQUENCIA (se única)."""
    if isinstance(meta_df, pd.DataFrame):
        try:
            freq_detectada = meta_df.loc[meta_df['chave'] == 'frequencia_detectada', 'valor'].iloc[0]
            if str(freq_detectada).strip():
                return str(freq_detectada)
        except Exception:
            pass
    f_raw = valor_unico_ou_multiplos(df_raw, 'FREQUENCIA')
    return f_raw if f_raw != "—" else "—"

def montar_excel(agenda, resumo, checagens, meta, parciais=None) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as xw:
        agenda.to_excel(xw, sheet_name='agenda_por_dia', index=False)
        resumo.to_excel(xw, sheet_name='resumo_dias', index=False)
        checagens.to_excel(xw, sheet_name='checagens', index=False)
        meta.to_excel(xw, sheet_name='meta', index=False)
        if parciais is not None and not parciais.empty:
            parciais.to_excel(xw, sheet_name='parciais_ignoradas', index=False)
    bio.seek(0)
    return bio.read()

# ------------------------------------------------------------
# Processamento principal
# ------------------------------------------------------------
def processar_df(
    df: pd.DataFrame,
    incluir_parciais_em_aba: bool = True,
    excluir_logradouro_vazio_da_agenda: bool = False,
):
    """
    - Constrói a 'agenda' (formato longo) com linhas COMPLETAS (HORARIO + ORDEM).
    - Mantém base de contexto: ID, SETOR, TIPOCOLETA, FREQUENCIA, TURNO, TIPO, TITULO,
      PREPOSICAO, LOGRADOURO, INICIO, FIM, DISTRITO, SUBPREFEITURA.
    - 'Pontos do setor' serão calculados *depois* como linhas da agenda onde LOGRADOURO ≠ vazio.
    - Retorna: agenda, resumo_dias, checagens, meta, parciais
    """
    base_cols = [c for c in [
        'ID','SETOR','TIPOCOLETA','FREQUENCIA','TURNO',
        'TIPO','TITULO','PREPOSICAO','LOGRADOURO',
        'INICIO','FIM','DISTRITO','SUBPREFEITURA'
    ] if c in df.columns]

    linhas_ok, linhas_parciais = [], []

    for dia in DIAS:
        hcol, ocol, fcol = f'HORARIO{dia}', f'ORDEM{dia}', f'FORMACOLETA{dia}'
        for c in (hcol, ocol, fcol):
            if c not in df.columns:
                df[c] = pd.NA

        bloco = df[base_cols + [hcol, ocol, fcol]].copy()
        bloco.rename(columns={hcol: 'HORARIO', ocol: 'ORDEM', fcol: 'FORMA_COLETA'}, inplace=True)
        bloco['DIA_SEMANA'] = dia

        bloco['HORARIO'] = bloco['HORARIO'].apply(to_hhmm)     # texto hh:mm
        bloco['ORDEM']   = pd.to_numeric(bloco['ORDEM'], errors='coerce').astype('Int64')
        bloco['FORMA_COLETA'] = bloco['FORMA_COLETA'].astype('string').str.strip().fillna('')

        has_h = bloco['HORARIO'].ne('')
        has_o = bloco['ORDEM'].notna()

        bloco['STATUS'] = np.where(has_h & has_o, 'completo',
                           np.where(has_h & ~has_o, 'so_horario',
                           np.where(~has_h & has_o, 'so_ordem', 'vazio')))

        linhas_ok.append(bloco[bloco['STATUS'] == 'completo'])

        if incluir_parciais_em_aba:
            linhas_parciais.append(bloco[bloco['STATUS'].isin(['so_horario', 'so_ordem'])])

    agenda = pd.concat(linhas_ok, ignore_index=True) if linhas_ok else pd.DataFrame()
    parciais = pd.concat(linhas_parciais, ignore_index=True) if (incluir_parciais_em_aba and linhas_parciais) else pd.DataFrame()

    # Opcional: excluir da agenda linhas com LOGRADOURO vazio (caso queira limpar o dataset principal)
    if excluir_logradouro_vazio_da_agenda and not agenda.empty and 'LOGRADOURO' in agenda.columns:
        agenda = agenda[agenda['LOGRADOURO'].fillna('').astype(str).str.strip().ne('')].copy()

    # Ordenação amigável
    if not agenda.empty:
        if 'DIA_SEMANA' in agenda.columns:
            agenda['DIA_SEMANA'] = agenda['DIA_SEMANA'].astype(CAT_DIAS)
        sort_cols = [c for c in ['ID', 'DIA_SEMANA', 'ORDEM'] if c in agenda.columns]
        agenda.sort_values(by=sort_cols, inplace=True, kind='stable')

    # Resumo por dia
    if not agenda.empty and 'DIA_SEMANA' in agenda.columns:
        resumo = (agenda.groupby('DIA_SEMANA', as_index=False)
                        .size()
                        .rename(columns={'size':'linhas_completas'}))
        resumo['DIA_SEMANA'] = resumo['DIA_SEMANA'].astype(CAT_DIAS)
        # adiciona dias faltantes
        faltantes = set(DIAS) - set(resumo['DIA_SEMANA'].astype(str))
        if faltantes:
            add = pd.DataFrame({'DIA_SEMANA': list(faltantes), 'linhas_completas':[0]*len(faltantes)})
            add['DIA_SEMANA'] = add['DIA_SEMANA'].astype(CAT_DIAS)
            resumo = pd.concat([resumo, add], ignore_index=True)
        resumo.sort_values('DIA_SEMANA', inplace=True)
    else:
        resumo = pd.DataFrame({'DIA_SEMANA': DIAS, 'linhas_completas': [0]*len(DIAS)})

    # Frequência detectada (a partir da agenda)
    if not agenda.empty and 'DIA_SEMANA' in agenda.columns:
        freq_detectada = '/'.join([d for d in DIAS if agenda['DIA_SEMANA'].astype(str).eq(d).any()])
    else:
        freq_detectada = ''

    # Checagens: monotonicidade da ORDEM por dia
    checagens = []
    for dia in DIAS:
        if agenda.empty:
            checagens.append({'DIA_SEMANA': dia, 'linhas': 0, 'ordem_nao_decrescente': None})
            continue
        sub = agenda[agenda['DIA_SEMANA'].astype(str) == dia]
        if sub.empty:
            checagens.append({'DIA_SEMANA': dia, 'linhas': 0, 'ordem_nao_decrescente': None})
            continue
        s = sub['ORDEM'].dropna().sort_values().to_numpy()
        ok = bool(np.all(np.diff(s) >= 0)) if len(s) else None
        checagens.append({'DIA_SEMANA': dia, 'linhas': int(sub.shape[0]), 'ordem_nao_decrescente': ok})
    checagens_df = pd.DataFrame(checagens)

    # Metadados
    freq_coluna = ''
    if 'FREQUENCIA' in df.columns and not df['FREQUENCIA'].isna().all():
        try:
            freq_coluna = str(df['FREQUENCIA'].dropna().astype(str).str.strip().unique()[0])
        except Exception:
            freq_coluna = ''
    meta = pd.DataFrame({
        'chave': ['frequencia_coluna','frequencia_detectada'],
        'valor': [freq_coluna, freq_detectada]
    })

    return agenda, resumo, checagens_df, meta, parciais

# ------------------------------------------------------------
# Mini painel (usa agenda + df_raw + meta)
# ------------------------------------------------------------
def render_mini_painel(df_raw: pd.DataFrame, agenda: Optional[pd.DataFrame],
                       meta: Optional[pd.DataFrame], uploaded_name: Optional[str]):
    if agenda is None or not isinstance(agenda, pd.DataFrame):
        st.info("Faça o upload e processe o arquivo para ver o painel.")
        return

    # Pontos do setor = linhas válidas: LOGRADOURO ≠ vazio e (HORARIO + ORDEM) presentes
    if not agenda.empty and 'LOGRADOURO' in agenda.columns:
        endereco_ok = agenda['LOGRADOURO'].fillna('').astype(str).str.strip().ne('')
        pontos_setor = int(agenda[endereco_ok].shape[0])
    else:
        pontos_setor = 0

    setor_nome      = nome_setor(df_raw, uploaded_name)
    subprefeitura   = valor_unico_ou_multiplos(df_raw, 'SUBPREFEITURA')
    frequencia_exib = frequencia_para_exibir(meta, df_raw)
    turno           = valor_unico_ou_multiplos(df_raw, 'TURNO')
    tipo_coleta     = valor_unico_ou_multiplos(df_raw, 'TIPOCOLETA')

    st.markdown("### 🔎 Visão geral do setor")
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("Pontos do setor", f"{pontos_setor:,}".replace(",", "."))
    with c2: st.metric("Setor", setor_nome)
    with c3: st.metric("Subprefeitura", subprefeitura)

    c4, c5, c6 = st.columns(3)
    with c4: st.metric("Frequência", frequencia_exib)
    with c5: st.metric("Turno", turno)
    with c6: st.metric("Tipo de coleta", tipo_coleta)

# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
st.title("Normalizador de Roteiro por Dia (HORARIO/ORDEM)")
st.caption("Faça upload da planilha (.xlsx) do setor. O app detecta os dias com HORARIO+ORDEM e normaliza a agenda.")

uploaded_file = st.file_uploader("Selecione a planilha do setor (formato .xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        abas = xls.sheet_names
        if not abas:
            st.error("Não foi possível ler as abas do arquivo. Verifique se o formato é .xlsx válido.")
            st.stop()

        aba_escolhida = st.selectbox("Escolha a aba com os dados", options=abas, index=0)
        df_raw = pd.read_excel(uploaded_file, sheet_name=aba_escolhida)

        st.subheader("Prévia dos dados brutos")
        st.dataframe(df_raw.head(12), use_container_width=True)

        # Opções do processamento
        st.markdown("### Opções")
        col1, col2 = st.columns(2)
        with col1:
            incluir_parciais = st.checkbox(
                "Incluir linhas parciais em aba separada (somente HORARIO ou somente ORDEM)",
                value=True
            )
        with col2:
            excluir_logradouro_vazio = st.checkbox(
                "Excluir da agenda linhas com LOGRADOURO vazio",
                value=False
            )

        # Processar
        agenda, resumo, checagens, meta, parciais = processar_df(
            df_raw,
            incluir_parciais_em_aba=incluir_parciais,
            excluir_logradouro_vazio_da_agenda=excluir_logradouro_vazio
        )

        # ---- MINI PAINEL (requisito da Gio)
        st.markdown("---")
        render_mini_painel(df_raw, agenda, meta, getattr(uploaded_file, 'name', None))

        # ---- Tabelas e download
        st.markdown("---")
        st.subheader("Agenda normalizada")
        if agenda.empty:
            st.warning("Nenhuma linha completa (HORARIO + ORDEM) foi encontrada. Verifique o arquivo/aba.")
        else:
            st.dataframe(agenda.head(100), use_container_width=True)

        colA, colB, colC = st.columns(3)
        with colA:
            st.markdown("**Resumo por dia**")
            st.dataframe(resumo, use_container_width=True)
        with colB:
            st.markdown("**Checagens**")
            st.dataframe(checagens, use_container_width=True)
        with colC:
            st.markdown("**Metadados**")
            st.dataframe(meta, use_container_width=True)

        if incluir_parciais and parciais is not None and not parciais.empty:
            with st.expander("Parciais (ignoras na agenda)"):
                st.dataframe(parciais.head(100), use_container_width=True)

        # Download Excel
        out_bytes = montar_excel(agenda, resumo, checagens, meta, parciais if incluir_parciais else None)
        st.download_button(
            label="⬇️ Baixar Excel normalizado",
            data=out_bytes,
            file_name="roteiro_normalizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.exception(e)
        st.error("Erro ao processar o arquivo. Confira se a estrutura está conforme o padrão (colunas HORARIO*/ORDEM* por dia).")
else:
    st.info("👉 Faça o upload de um arquivo .xlsx para começar.")
