import pandas as pd
from tkinter import Tk, filedialog, messagebox
from openpyxl.styles import Font

# Esconde a janela principal
root = Tk()
root.withdraw()

arquivo = filedialog.askopenfilename(
    title="Selecione o CSV do Uber",
    filetypes=[("CSV", "*.csv")]
)

if not arquivo:
    exit()

df = pd.read_csv(arquivo)

# Colunas principais (serão usadas apenas se existirem)
colunas = [
    "request_timestamp_local",
    "city_name",
    "begintrip_address",
    "dropoff_address",
    "product_type_name",
    "trip_distance_miles",
    "trip_duration_seconds",
    "fare_amount",
    "trip_status"
]

colunas_existentes = [c for c in colunas if c in df.columns]

planilha = df[colunas_existentes].copy()

# Renomeia colunas
nomes = {
    "request_timestamp_local": "Data/Hora",
    "city_name": "Cidade",
    "begintrip_address": "Origem",
    "dropoff_address": "Destino",
    "product_type_name": "Tipo",
    "trip_distance_miles": "Distância (milhas)",
    "trip_duration_seconds": "Duração (segundos)",
    "fare_amount": "Valor",
    "trip_status": "Status"
}

planilha.rename(columns=nomes, inplace=True)

# Converte datas
if "Data/Hora" in planilha.columns:
    planilha["Data/Hora"] = (
    pd.to_datetime(planilha["Data/Hora"], errors="coerce")
      .dt.tz_localize(None)
)

arquivo_saida = "Relatorio_Uber.xlsx"

with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:

    planilha.to_excel(
        writer,
        sheet_name="Viagens",
        index=False
    )

    resumo = []

    if "Valor" in planilha.columns:
        resumo.append(["Total Gasto", planilha["Valor"].sum()])
        resumo.append(["Média por Corrida", planilha["Valor"].mean()])
        resumo.append(["Maior Corrida", planilha["Valor"].max()])
        resumo.append(["Menor Corrida", planilha["Valor"].min()])

    resumo.append(["Quantidade de Viagens", len(planilha)])

    resumo_df = pd.DataFrame(
        resumo,
        columns=["Indicador", "Valor"]
    )

    resumo_df.to_excel(
        writer,
        sheet_name="Resumo",
        index=False
    )

    wb = writer.book

    for aba in wb.sheetnames:
        ws = wb[aba]

        for cell in ws[1]:
            cell.font = Font(bold=True)

print("Planilha criada com sucesso!")

messagebox.showinfo(
    "Concluído",
    f"Arquivo salvo como:\n{arquivo_saida}"
)