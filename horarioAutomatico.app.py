import pandas as pd

# Lê a planilha
df = pd.read_excel("upload.xlsx")

# Identifica as colunas de horário (todas menos ORDEM, se existir)
colunas_horario = [c for c in df.columns if c != "ORDEM"]

# Converte todas para datetime (ignorando data, só hora)
for col in colunas_horario:
    df[col] = pd.to_datetime(df[col].astype(str), format='%H:%M', errors='coerce')

# Ajuste para madrugadas: se um horário for menor que o primeiro horário do dia, soma 1 dia
for idx, row in df.iterrows():
    horarios = row[colunas_horario].dropna().sort_values()
    if len(horarios) > 0:
        base = horarios.iloc[0]
        for col in colunas_horario:
            if pd.notna(row[col]) and row[col] < base:
                df.at[idx, col] = row[col] + pd.Timedelta(days=1)

# Reordena os horários dentro da linha
for idx, row in df.iterrows():
    valores = row[colunas_horario].dropna().sort_values().values
    for i, col in enumerate(colunas_horario[:len(valores)]):
        df.at[idx, col] = valores[i]

# Salva de volta
df.to_excel("saida.xlsx", index=False)
