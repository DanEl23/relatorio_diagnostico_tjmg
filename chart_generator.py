# chart_generator.py

import matplotlib.pyplot as plt
import os

# Define uma paleta de cores institucional fictícia (tons de azul/cinza)
TJMG_BLUE = "#004a91"
TJMG_GRAY = "#6c757d"
TJMG_ACCENT = "#fca311"

def create_all_charts(stats_df):
    """
    Recebe o DataFrame 'stats_justica_numeros' e gera os gráficos 
    necessários, salvando-os em /temp.
    """
    
    print("Gerando gráficos...")
    chart_paths = {}

    # --- Gráfico 01: Taxa de Congestionamento ---
    try:
        plt.figure(figsize=(10, 6))
        plt.bar(
            stats_df['Ano'], 
            stats_df['Taxa_Congestionamento_Total'], 
            color=TJMG_BLUE,
            label='Taxa de Congestionamento'
        )
        
        plt.title('Gráfico 01 (Fictício) - Taxa de Congestionamento Total')
        plt.ylabel('Percentual (%)')
        plt.xlabel('Ano')
        plt.ylim(60, 80) # Fixar eixo Y para comparação
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.tight_layout()
        
        chart_path = os.path.join('temp', 'grafico_01_congestionamento.png')
        plt.savefig(chart_path, dpi=300)
        plt.close()
        
        chart_paths['grafico_01_congestionamento'] = chart_path
        print(f"Gráfico 01 salvo em: {chart_path}")

    except Exception as e:
        print(f"Erro ao gerar Gráfico 01: {e}")

    # (Adicionar aqui outras funções para gerar Gráfico 02, 03, etc.)

    return chart_paths