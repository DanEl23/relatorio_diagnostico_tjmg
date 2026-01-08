import pandas as pd
import math

# Caminho da planilha
excel_path = 'Informações TJMG_CEINFO.xlsx'

# Lê a aba "Movimentação Processual"
df = pd.read_excel(excel_path, sheet_name='Movimentação Processual', header=None)

# HEADER_MERGE: linha 2 (índice 1)
header_merge = tuple(['HEADER_MERGE'] + [str(x) if not (isinstance(x, float) and math.isnan(x)) else '' for x in df.iloc[1].values[0:8]])

# SUB_HEADER: linha 3 (índice 2)
sub_header = tuple(['SUB_HEADER'] + [str(x) if not (isinstance(x, float) and math.isnan(x)) else '' for x in df.iloc[2].values[0:8]])

# DATA_ROW e TOTAL_ROW: linhas 4 a 7 (índices 3 a 6)
data_rows = []
for i in range(3, 7):
    row = df.iloc[i].values[0:8]
    tipo = 'TOTAL_ROW' if str(row[0]).strip().upper() == 'TOTAL' else 'DATA_ROW'
    data_rows.append(tuple([tipo] + [str(x) if not (isinstance(x, float) and math.isnan(x)) else '' for x in row]))

# Junta tudo
result = [header_merge, sub_header] + data_rows

print('dados_tabela_processos = [')
for item in result:
    print(f'    {item},')
print(']')
