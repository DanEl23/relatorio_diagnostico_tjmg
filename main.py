# main.py

import data_loader
import chart_generator
import report_builder
import os
import sys

# --- Configuração ---
DATA_SOURCE_PATH = 'Dados_Fonte_Ficticios_TJMG.xlsx'
TEMPLATE_PATH = 'Relatorio_Template_Teste.docx'
OUTPUT_PATH = 'Relatorio_Gerado_TESTE.docx'

def run_automation():
    print("--- Iniciando Processo de Automação de Relatório ---")
    
    # 1. Validar arquivos de entrada
    if not os.path.exists(DATA_SOURCE_PATH):
        print(f"Erro: Arquivo de dados '{DATA_SOURCE_PATH}' não encontrado.")
        print("Por favor, execute 'python criar_dados_ficticios.py' primeiro.")
        sys.exit(1)
        
    if not os.path.exists(TEMPLATE_PATH):
        print(f"Erro: Arquivo de template '{TEMPLATE_PATH}' não encontrado.")
        print("Por favor, crie este arquivo conforme as instruções (Passo 3).")
        sys.exit(1)

    # 2. Carregar e preparar dados
    context = data_loader.get_context(DATA_SOURCE_PATH)

    # 3. Gerar imagens de gráficos
    # Os gráficos usam o DataFrame 'stats_justica_numeros' do contexto
    chart_paths = chart_generator.create_all_charts(
        context['stats_justica_numeros']
    )
    # Adiciona os caminhos das imagens ao contexto
    context['images'] = chart_paths

    # 4. Renderizar o template DOCX
    rendered_docx = report_builder.render(TEMPLATE_PATH, context, OUTPUT_PATH)

    if rendered_docx:
        print("--- Processo Concluído com Sucesso ---")
        print(f"Arquivo final gerado: {rendered_docx}")
        print("(Nota: O Sumário e campos automáticos não são atualizados sem o 'post_processor' com win32com).")
    else:
        print("--- Processo Falhou ---")

if __name__ == "__main__":
    run_automation()