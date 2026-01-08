import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import sys

# --- CONFIGURAÇÕES ---
# Nome do arquivo exatamente como está na sua pasta
ARQUIVO_EXCEL = 'Informações TJMG_ASPLAG resposta do CEINFO.xlsx'
NOME_ABA = 'Estrutura e Força de Trabalho'

# Itens exatos que queremos buscar (Texto da coluna de descrição)
ITENS_ALVO = [
    "Fóruns digitais(*)",
    "Central de Processo Eletrônico instalada",
    "Salas de depoimento especial instaladas"
]

CORES = ['#004a80', '#e67e22', '#27ae60'] # Azul, Laranja, Verde

def carregar_dados_excel(caminho, aba):
    try:
        # Lê o Excel sem cabeçalho (header=None) para trazer tudo como dados brutos
        df = pd.read_excel(caminho, sheet_name=aba, header=None)
        
        dados_encontrados = []

        # Para cada item que queremos plotar...
        for item in ITENS_ALVO:
            # Procura em qual célula (linha, coluna) o texto exato está
            # A função .stack() transforma o DataFrame em uma série longa para facilitar a busca
            posicoes = df[df.isin([item])].stack().index
            
            if not posicoes.empty:
                linha, col = posicoes[0] # Pega a primeira ocorrência (linha, coluna)
                
                # Assume que o valor numérico está na coluna IMEDIATAMENTE à direita (col + 1)
                try:
                    valor = df.iloc[linha, col + 1]
                    dados_encontrados.append({'Item': item, 'Quantidade': valor})
                except IndexError:
                    print(f"Aviso: Encontrei '{item}' mas não há coluna à direita com valor.")
            else:
                print(f"Aviso: Não encontrei o texto exato: '{item}' na planilha.")
        
        if not dados_encontrados:
            return pd.DataFrame()
            
        return pd.DataFrame(dados_encontrados)

    except FileNotFoundError:
        print(f"ERRO CRÍTICO: Não encontrei o arquivo '{caminho}'.")
        print("Verifique se o nome do arquivo está correto e se você está rodando o script na mesma pasta dele.")
        sys.exit()
    except Exception as e:
        print(f"Erro ao processar Excel: {e}")
        return pd.DataFrame()

def desenhar_elemento(ax, y, label, valor, cor):
    # Ajuste fino para labels muito longos (quebra de linha opcional)
    label_display = label
    if len(label) > 35: # Se for muito grande, diminui a fonte ou ajusta
        font_size_text = 10
    else:
        font_size_text = 12

    # 1. Retângulo (Texto)
    rect = patches.FancyBboxPatch(
        (0.05, y), 0.6, 0.15,
        boxstyle="round,pad=0.02,rounding_size=0.05",
        ec="none", fc=cor, alpha=0.9
    )
    ax.add_patch(rect)
    
    ax.text(
        0.08, y + 0.075, label_display,
        color='white', fontsize=font_size_text, fontweight='bold',
        va='center', ha='left'
    )
    
    # 2. Círculo (Valor)
    circle_center_x = 0.65
    circle_center_y = y + 0.075
    
    circle = patches.Circle(
        (circle_center_x, circle_center_y), 0.09,
        ec='white', fc=cor, linewidth=3
    )
    ax.add_patch(circle)
    
    ax.text(
        circle_center_x, circle_center_y, str(valor),
        color='white', fontsize=14, fontweight='bold',
        va='center', ha='center'
    )

# --- EXECUÇÃO ---
print(f"Lendo arquivo: {ARQUIVO_EXCEL}...")
df_plot = carregar_dados_excel(ARQUIVO_EXCEL, NOME_ABA)

if not df_plot.empty:
    # Garante a ordem que você pediu
    df_plot['Item'] = pd.Categorical(df_plot['Item'], categories=ITENS_ALVO, ordered=True)
    df_plot = df_plot.sort_values('Item')
    
    print("Dados carregados com sucesso:")
    print(df_plot)

    # Configuração da Figura
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, len(df_plot) * 0.3)
    ax.axis('off')
    
    y_pos = (len(df_plot) - 1) * 0.25
    
    for i, (index, row) in enumerate(df_plot.iterrows()):
        cor = CORES[i % len(CORES)]
        desenhar_elemento(ax, y_pos, row['Item'], row['Quantidade'], cor)
        y_pos -= 0.25

    NOME_SAIDA = 'resultado_tjmg_excel.png'
    plt.tight_layout()
    plt.savefig(NOME_SAIDA, dpi=300, bbox_inches='tight', transparent=True)
    print(f"\nImagem gerada: {NOME_SAIDA}")
    plt.show()
else:
    print("\nNenhum dado foi encontrado. Verifique se os nomes dos itens (ITENS_ALVO) estão idênticos aos da planilha.")