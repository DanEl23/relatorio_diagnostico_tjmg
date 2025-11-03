# data_loader.py

import pandas as pd

def get_context(data_source_path):
    """
    Carrega todos os dados do Excel e prepara o dicionário de 
    contexto (context) para o Jinja2.
    """
    
    # 1. Carregar Variáveis Globais
    df_globais = pd.read_excel(data_source_path, sheet_name='Variaveis_Globais')
    # Converte a tabela Chave/Valor em um dicionário simples
    context = dict(zip(df_globais['Chave'], df_globais['Valor']))

    # 2. Carregar Tabelas (convertendo para lista de dicionários)
    df_mov_dist = pd.read_excel(data_source_path, sheet_name='Mov_Distribuido')
    context['tabela_mov_distribuido'] = df_mov_dist.to_dict('records')
    
    df_orc_2024 = pd.read_excel(data_source_path, sheet_name='Orcamento2024')
    context['tabela_orcamento_2024'] = df_orc_2024.to_dict('records')
    
    # 3. Carregar Dados para Gráficos (DataFrames brutos)
    df_justica_numeros = pd.read_excel(data_source_path, sheet_name='JusticaNumeros')
    context['stats_justica_numeros'] = df_justica_numeros
    
    print("Contexto de dados carregado.")
    return context