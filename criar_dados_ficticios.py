# criar_dados_ficticios.py

import pandas as pd
import os

print("Gerando arquivo de dados fictícios: 'Dados_Fonte_Ficticios_TJMG.xlsx'...")

# Garantir que o diretório 'temp' exista para os gráficos
if not os.path.exists("temp"):
    os.makedirs("temp")

# 1. Dados para Variáveis Globais (Ex: Títulos, rodapés)
df_globais = pd.DataFrame([
    {"Chave": "ano_relatorio", "Valor": 2025},
    {"Chave": "ano_exercicio", "Valor": 2024},
    {"Chave": "data_extracao", "Valor": "31/10/2024"},
    {"Chave": "total_magistrados", "Valor": 985},
    {"Chave": "total_servidores", "Valor": 14720},
    {"Chave": "posicao_ranking_transparencia", "Valor": "12º"},
    {"Chave": "igovtic_jud_nota", "Valor": 89.5}
])

# 2. Dados Fictícios para Tabela 06 (Movimentação - Distribuídos)
df_mov_dist = pd.DataFrame([
    {"Instancia": "Justiça Comum", "2020": 1200000, "2021": 1350000, "2022": 1400000, "2023": 1500000, "2024_Parcial": 1100000},
    {"Instancia": "Juizado Especial", "2020": 540000, "2021": 530000, "2022": 600000, "2023": 620000, "2024_Parcial": 500000},
    {"Instancia": "Turma Recursal", "2020": 110000, "2021": 120000, "2022": 130000, "2023": 125000, "2024_Parcial": 100000}
])

# 3. Dados Fictícios para Tabela 09/10 (Orçamento 2024)
df_orc_2024 = pd.DataFrame([
    {"Acao": "7006 - Proventos e Pensões", "Despesa_Realizada_RS": 2535040959.40},
    {"Acao": "2053 - Remuneração (Ativos)", "Despesa_Realizada_RS": 1353944848.00},
    {"Acao": "4257 - Encargos e Benefícios", "Despesa_Realizada_RS": 870500100.20},
    {"Acao": "2004 - Manutenção e Serviços", "Despesa_Realizada_RS": 460200300.00},
    {"Acao": "1001 - Investimentos (TIC)", "Despesa_Realizada_RS": 150100000.00}
])

# 4. Dados Fictícios para Gráficos (Ex: Justiça em Números - Tabela 12)
df_justica_numeros = pd.DataFrame({
    "Ano": [2020, 2021, 2022, 2023, 2024],
    "Taxa_Congestionamento_Total": [70.1, 69.5, 69.0, 68.8, 68.5],
    "Indice_Atendimento_Demanda": [102.5, 103.1, 104.0, 103.5, 105.0],
    "Casos_Novos": [1850000, 2000000, 2130000, 2245000, 1700000]
})

# Salvar no arquivo Excel
with pd.ExcelWriter('Dados_Fonte_Ficticios_TJMG.xlsx') as writer:
    df_globais.to_excel(writer, sheet_name='Variaveis_Globais', index=False)
    df_mov_dist.to_excel(writer, sheet_name='Mov_Distribuido', index=False)
    df_orc_2024.to_excel(writer, sheet_name='Orcamento2024', index=False)
    df_justica_numeros.to_excel(writer, sheet_name='JusticaNumeros', index=False)

print("Arquivo 'Dados_Fonte_Ficticios_TJMG.xlsx' gerado com sucesso.")