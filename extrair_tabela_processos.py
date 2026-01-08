import pandas as pd

# Caminho da planilha
excel_path = 'Informações TJMG_CEINFO.xlsx'

# Lê a aba "Movimentação Processual"
df = pd.read_excel(excel_path, sheet_name='Movimentação Processual')

# Seleciona as linhas 2 a 8 (índice 1 a 7, pois pandas é 0-based)
df_processos = df.iloc[1:8]

# Exibe as colunas e os dados selecionados
print('Colunas:', list(df_processos.columns))
print(df_processos)

# Converte para lista de tuplas (como no report_data.py)
dados_tabela_processos = [tuple(row) for row in df_processos.values]

print('\nExemplo de dados_tabela_processos:')
for item in dados_tabela_processos:
    print(item)
