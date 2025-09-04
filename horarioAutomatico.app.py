# =========================
# MINI PAINEL DE INDICADORES
# (colar após processar o df e gerar `agenda`, `resumo`, `checagens`, `meta`)
# =========================

import streamlit as st
import pandas as pd
from typing import Optional

def _valor_unico_ou_multiplos(df: pd.DataFrame, col: str) -> str:
    """Retorna o valor único da coluna (se houver), 'múltiplos' se >1, ou '—' se vazio/inexistente."""
    if col not in df.columns:
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

def _frequencia_para_exibir(meta_df: pd.DataFrame, df_raw: pd.DataFrame) -> str:
    """Prioriza a frequência detectada; se vazia, cai para a coluna FREQUENCIA do arquivo."""
    freq_detectada = ""
    try:
        freq_detectada = meta_df.loc[meta_df['chave'] == 'frequencia_detectada', 'valor'].iloc[0]
    except Exception:
        pass

    if freq_detectada and str(freq_detectada).strip():
        return str(freq_detectada)

    # fallback: usa o valor da coluna FREQUENCIA (se for único)
    f_raw = _valor_unico_ou_multiplos(df_raw, 'FREQUENCIA')
    return f_raw if f_raw != "—" else "—"

def _nome_setor(df_raw: pd.DataFrame, uploaded_name: Optional[str] = None) -> str:
    """Tenta obter o nome do setor pela coluna SETOR; se não houver, tenta inferir do nome do arquivo."""
    setor = _valor_unico_ou_multiplos(df_raw, 'SETOR')
    if setor != "—" and setor != "múltiplos":
        return setor
    # fallback: tenta extrair algo como 'PR18' do nome do arquivo
    if uploaded_name:
        import re
        m = re.search(r'\\b([A-Z]{2}\\d{1,3})\\b', uploaded_name.upper())
        if m:
            return m.group(1)
    return setor  # '—' ou 'múltiplos'

# ---- Contagem de pontos
# Por padrão, contamos cada linha válida (registro-dia) da agenda:
pontos_por_registro_dia = int(agenda.shape[0]) if not agenda.empty else 0

# Se preferir contar 1 ponto por endereço (ID), independentemente do dia:
pontos_por_endereco = int(agenda['ID'].nunique()) if ('ID' in agenda.columns and not agenda.empty) else 0

# Pequeno seletor (opcional) para o usuário escolher como contar:
contagem = st.radio(
    "Como você quer contar os pontos?",
    options=["Registro-dia (padrão)", "Por endereço (ID único)"],
    horizontal=True,
    index=0
)
pontos = pontos_por_registro_dia if contagem == "Registro-dia (padrão)" else pontos_por_endereco

# Extraímos os metadados solicitados
setor_nome      = _nome_setor(df, getattr(uploaded, 'name', None))
subprefeitura   = _valor_unico_ou_multiplos(df, 'SUBPREFEITURA')
frequencia_exib = _frequencia_para_exibir(meta, df)
turno           = _valor_unico_ou_multiplos(df, 'TURNO')
tipo_coleta     = _valor_unico_ou_multiplos(df, 'TIPOCOLETA')

# ---- Layout do mini painel
st.markdown("### 🔎 Visão geral do setor")
c1, c2, c3 = st.columns(3)
with c1:
    st.metric("Pontos do setor", f"{pontos:,}".replace(",", "."))
with c2:
    st.metric("Setor", setor_nome)
with c3:
    st.metric("Subprefeitura", subprefeitura)

c4, c5, c6 = st.columns(3)
with c4:
    st.metric("Frequência", frequencia_exib)
with c5:
    st.metric("Turno", turno)
with c6:
    st.metric("Tipo de coleta", tipo_coleta)
