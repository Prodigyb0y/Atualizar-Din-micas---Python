import time
from openpyxl import load_workbook

# Caminho para a planilha (substitua por seu caminho)
caminho_planilha = r"C:\caminho\para\sua\planilha.xlsx"

# Carrega a planilha
workbook = load_workbook(caminho_planilha)

# Acessa a planilha "Geral"
ws = workbook.worksheets["Geral"]

# Busca pela tabela dinâmica
for pivot_table in ws.pivot_tables:
    if pivot_table.name == "TabelaDinamica1":
        # Atualiza a tabela dinâmica
        pivot_table.refresh_table()
        time.sleep(2)
        print("Tabela Dinâmica atualizada!")
        break