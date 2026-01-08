import pandas as pd

# Caminho da planilha
excel_path = 'Informações TJMG_CEINFO.xlsx'

# Carrega todas as abas
xls = pd.ExcelFile(excel_path)
print('Abas encontradas:')
print(xls.sheet_names)

# Para cada aba, mostra as primeiras linhas e colunas
for sheet in xls.sheet_names:
    print(f'\n--- Aba: {sheet} ---')
    df = pd.read_excel(xls, sheet)
    print('Colunas:', list(df.columns))
    print(df.head())
