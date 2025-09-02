import pandas as pd
import datetime as dt

# Função para converter valores em horário real
def converter_para_horario(valor):
    try:
        if pd.isna(valor):
            return None

        # Caso venha como número decimal do Excel (ex: 0.75 = 18:00)
        if isinstance(valor, (int, float)):
            horas = int(valor * 24)
            minutos = int((valor * 24 * 60) % 60)
            return dt.time(horas, minutos)

        # Caso já seja datetime/time
        if isinstance(valor, dt.time):
            return valor
        if isinstance(valor, dt.datetime):
            return valor.time()

        # Caso venha como string "HH:MM"
        if isinstance(valor, str):
            return dt.datetime.strptime(valor.strip(), "%H:%M").time()

    except:
        return None
    return None

# Carregar planilha
df = pd.read_excel("upload.xlsx")

# Garantir colunas certas
df = df.rename(columns={df.columns[0]: "HORARIO", df.columns[1]: "ORDEM"})

# Converte a coluna de horários
df["HORARIO"] = df["HORARIO"].apply(converter_para_horario)

# Ajustar virada da noite
horarios = []
for h in df["HORARIO"]:
    if h is None:
        horarios.append(None)
    else:
        horarios.append(dt.datetime.combine(dt.date.today(), h))

df["HORARIO_DT"] = horarios

# Detecta virada da noite (se tem horários tarde e cedo misturados)
if any(h and h.hour >= 18 for h in df["HORARIO_DT"]) and any(h and h.hour < 10 for h in df["HORARIO_DT"]):
    df.loc[df["HORARIO_DT"].dt.hour < 10, "HORARIO_DT"] += dt.timedelta(days=1)

# Ordena pelos horários ajustados
df = df.sort_values(by="HORARIO_DT").reset_index(drop=True)

# Reatribui nova ordem
df["ORDEM"] = range(1, len(df) + 1)

# Ajusta coluna final de horário (só o time, que o Excel entende)
df["HORARIO"] = df["HORARIO_DT"].dt.time

# Preview antes de salvar
print("\nPrévia do resultado final:")
print(df[["HORARIO", "ORDEM"]].head(10))

# Salvar em nova planilha
df[["HORARIO", "ORDEM"]].to_excel("saida.xlsx", index=False)
print("\nPlanilha 'saida.xlsx' gerada com sucesso!")
