import os
import pandas as pd

# nome do arquivo esperado
file_name = "upload.xlsx"

# diretório atual onde o script está rodando
print("Diretório atual:", os.getcwd())
print("Arquivos encontrados:", os.listdir())

# checa se o arquivo existe
if not os.path.exists(file_name):
    print(f"❌ Arquivo '{file_name}' não encontrado no diretório atual.")
    print("Coloque o arquivo no mesmo diretório ou use o caminho completo.")
    exit()

# lê a planilha
df = pd.read_excel(file_name)
print("✅ Planilha carregada com sucesso!")
print("Colunas encontradas:", df.columns.tolist())

# filtrar só colunas que começam com "HORARIO"
horario_cols = [col for col in df.columns if col.upper().startswith("HORARIO")]

if not horario_cols:
    print("❌ Nenhuma coluna de horário encontrada no arquivo.")
    exit()

# transformar horários em datetime
for col in horario_cols:
    df[col] = pd.to_datetime(df[col], errors="coerce").dt.time

# função para analisar cada coluna de horários
def analisar_horarios(series):
    series = series.dropna().sort_values()
    if series.empty:
        return None
    
    horarios = pd.to_datetime(series.astype(str))
    menor = horarios.min().strftime("%H:%M")
    maior = horarios.max().strftime("%H:%M")

    gaps = []
    for i in range(len(horarios) - 1):
        diff = (horarios.iloc[i+1] - horarios.iloc[i]).total_seconds() / 60
        if diff >= 10:  # só gaps >= 10 min
            gaps.append(f"{horarios.iloc[i].strftime('%H:%M')} → {horarios.iloc[i+1].strftime('%H:%M')} ({int(diff)} min)")

    return {
        "menor": menor,
        "maior": maior,
        "gaps": gaps
    }

# rodar análise por coluna
for col in horario_cols:
    resultado = analisar_horarios(df[col])
    print(f"\n📅 {col}:")
    if resultado:
        print(f" - Menor horário: {resultado['menor']}")
        print(f" - Maior horário: {resultado['maior']}")
        if resultado['gaps']:
            print(" - Gaps ≥ 10 min:")
            for g in resultado['gaps']:
                print(f"   {g}")
        else:
            print(" - Nenhum gap ≥ 10 min encontrado")
    else:
        print(" - Sem horários válidos")
