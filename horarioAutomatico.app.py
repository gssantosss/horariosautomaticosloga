import pandas as pd

# Carregar os dados
df = pd.read_excel("upload.xlsx")

# Pega só colunas que começam com HORARIO
horario_cols = [col for col in df.columns if col.startswith("HORARIO")]

for col in horario_cols:
    # Tira valores vazios e converte pra datetime
    horarios = pd.to_datetime(df[col], errors="coerce").dropna().sort_values().reset_index(drop=True)
    
    if horarios.empty:
        continue
    
    menor = horarios.iloc[0].strftime("%H:%M")
    maior = horarios.iloc[-1].strftime("%H:%M")
    
    gaps = []
    for i in range(len(horarios) - 1):
        diff = (horarios[i+1] - horarios[i]).total_seconds() / 60
        if diff >= 10:  # gap mínimo de 10 minutos
            antes = horarios[i].strftime("%H:%M")
            depois = horarios[i+1].strftime("%H:%M")
            gaps.append((antes, depois))
    
    # Monta saída
    output = f"{col}: Menor horário: {menor} "
    for j, (antes, depois) in enumerate(gaps, 1):
        output += f"| Horário antes do gap{j}: {antes} | Horário depois do gap{j}: {depois} "
    output += f"| Horário final: {maior}"
    
    print(output)
