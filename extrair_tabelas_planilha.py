import pandas as pd
import math
from typing import List, Tuple, Any

EXCEL_PATH = 'Informações TJMG_CEINFO.xlsx'

# Função genérica para extrair tabela de uma aba, dado o range de linhas e colunas

def extrair_tabela(sheet: str, start_row: int, end_row: int, col_range: slice, tipo_linha: List[str]=None, tipo_total: str=None) -> List[Tuple[Any,...]]:
    """
    Extrai uma tabela de uma aba do Excel e retorna uma lista de tuplas.
    - sheet: nome da aba
    - start_row, end_row: intervalo de linhas (0-based, inclusive start, exclusive end)
    - col_range: fatia de colunas (ex: slice(0,8))
    - tipo_linha: lista de tipos para cada linha (ex: ['HEADER_MERGE', 'SUB_HEADER', 'DATA_ROW', ...])
    - tipo_total: tipo para linha de total (ex: 'TOTAL_ROW')
    """
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet, header=None)
    result = []
    for idx, i in enumerate(range(start_row, end_row)):
        row = df.iloc[i][col_range].values
        # Define o tipo da linha
        if tipo_linha:
            tipo = tipo_linha[idx] if idx < len(tipo_linha) else 'DATA_ROW'
        elif tipo_total and str(row[0]).strip().upper() == 'TOTAL':
            tipo = tipo_total
        else:
            tipo = 'DATA_ROW'
        # Converte valores para string, exceto NaN
        tuple_row = tuple([tipo] + [str(x) if not (isinstance(x, float) and math.isnan(x)) else '' for x in row])
        result.append(tuple_row)
    return result

# Exemplo de uso para dados_tabela_processos
if __name__ == '__main__':
    # PROCESSOS DISTRIBUÍDOS: linhas 2 a 7 (índices 1 a 6), 8 colunas
    tipos = ['HEADER_MERGE', 'SUB_HEADER', 'DATA_ROW', 'DATA_ROW', 'DATA_ROW', 'DATA_ROW']
    processos = extrair_tabela('Movimentação Processual', 1, 7, slice(0,8), tipo_linha=tipos)
    print('dados_tabela_processos = [')
    for item in processos:
        print(f'    {item},')
    print(']')

    # JULGAMENTOS: linhas 9 a 15 (índices 8 a 15), 8 colunas
    tipos_julg = ['HEADER_MERGE', 'SUB_HEADER', 'DATA_ROW', 'DATA_ROW', 'DATA_ROW', 'DATA_ROW', 'TOTAL_ROW']
    julgamentos = extrair_tabela('Movimentação Processual', 8, 15, slice(0,8), tipo_linha=tipos_julg)
    print('\ndados_tabela_julgamentos = [')
    for item in julgamentos:
        print(f'    {item},')
    print(']')

    # ACERVO: linhas 16 a 22 (índices 15 a 22), 8 colunas
    tipos_acervo = ['HEADER_MERGE', 'SUB_HEADER', 'DATA_ROW', 'DATA_ROW', 'DATA_ROW', 'DATA_ROW', 'TOTAL_ROW']
    acervo = extrair_tabela('Movimentação Processual', 15, 22, slice(0,8), tipo_linha=tipos_acervo)
    print('\ndados_tabela_acervo = [')
    for item in acervo:
        print(f'    {item},')
    print(']')

    # Adicione aqui chamadas para outras tabelas conforme necessário
