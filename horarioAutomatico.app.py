import streamlit as st
import pandas as pd

st.title("📂 Menor e Maior Horário de Cada Coluna HORARIO")

uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    horario_cols = [col for col in df.columns if col.upper().startswith("HORARIO")]
    horario_cols_validas = [col for col in horario_cols if df[col].notna().any()]
    
    if not horario_cols_validas:
        st.warning("⚠️ Nenhuma coluna HORARIO preenchida encontrada.")
    else:
        resultados = []
        for col_horario in horario_cols_validas:
            dia = col_horario.replace("HORARIO", "")
            ordem_col = f"ORDEM{dia}"
            
            if ordem_col not in df.columns:
                st.warning(f"⚠️ Coluna {ordem_col} não encontrada para {col_horario}. Ignorando.")
                continue
            
            subset = df.loc[df[col_horario].notna(), [col_horario, ordem_col]].copy()
            
            # Converte para datetime seguro
            subset[col_horario] = pd.to_datetime(subset[col_horario], errors='coerce')
            subset = subset.dropna(subset=[col_horario])  # remove valores que não viraram datetime
            
            if subset.empty:
                continue
            
            menor_horario_idx = subset[col_horario].idxmin()
            menor_horario = subset.loc[menor_horario_idx, col_horario].time()
            ordem_menor = subset.loc[menor_horario_idx, ordem_col]
            
            maior_horario_idx = subset[col_horario].idxmax()
            maior_horario = subset.loc[maior_horario_idx, col_horario].time()
            ordem_maior = subset.loc[maior_horario_idx, ordem_col]
            
            resultados.append({
                "Dia": dia,
                "Menor Horário": menor_horario.strftime("%H:%M"),
                "ORDEM Menor": ordem_menor,
                "Maior Horário": maior_horario.strftime("%H:%M"),
                "ORDEM Maior": ordem_maior
            })
        
        st.subheader("📊 Menor e Maior Horário por coluna HORARIO")
        st.dataframe(pd.DataFrame(resultados))
