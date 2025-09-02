import streamlit as st
import pandas as pd

st.title("📂 Menor e Maior Horário de Cada Coluna HORARIO")

uploaded_file = st.file_uploader("Escolha a planilha Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # Identifica colunas HORARIO com pelo menos 1 valor
    horario_cols = [col for col in df.columns if col.upper().startswith("HORARIO")]
    horario_cols = [col for col in horario_cols if df[col].notna().any()]
    
    resultados = []

    for col_horario in horario_cols:
        dia = col_horario.replace("HORARIO", "")
        ordem_col = f"ORDEM{dia}"
        if ordem_col not in df.columns:
            continue
        
        subset = df[[col_horario, ordem_col]].dropna()
        
        # Mantém datetime64[ns] internamente
        subset[col_horario] = pd.to_datetime(subset[col_horario], errors='coerce')
        subset = subset.dropna(subset=[col_horario])
        
        if subset.empty:
            continue
        
        # menor horário
        menor_idx = subset[col_horario].idxmin()
        menor = subset.loc[menor_idx, col_horario]
        ordem_menor = subset.loc[menor_idx, ordem_col]

        # maior horário
        maior_idx = subset[col_horario].idxmax()
        maior = subset.loc[maior_idx, col_horario]
        ordem_maior = subset.loc[maior_idx, ordem_col]

        resultados.append({
            "Dia": dia,
            "Menor Horário": menor.strftime("%H:%M"),
            "ORDEM Menor": ordem_menor,
            "Maior Horário": maior.strftime("%H:%M"),
            "ORDEM Maior": ordem_maior
        })

    st.dataframe(pd.DataFrame(resultados))
