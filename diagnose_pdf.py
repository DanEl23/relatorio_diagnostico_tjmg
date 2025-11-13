import fitz  # PyMuPDF
import os

# =====================================================================
#                        DIAGNÓSTICO DE PDF
# =====================================================================
PDF_PATH = "justica-em-numeros-2024.pdf"

# PÁGINAS A SEREM ANALISADAS (números de página como aparecem no PDF, 1-indexed)
# Exemplos:
#   [1, 2, 3] - Analisa páginas 1, 2 e 3
#   [50, 100, 150] - Analisa páginas 50, 100 e 150
#   range(1, 11) - Analisa páginas 1 a 10
PAGES_TO_ANALYZE = list(range(40, 60))  # AJUSTE AQUI AS PÁGINAS QUE DESEJA VER

def diagnose_pdf_text(pdf_path: str, page_numbers: list):
    """
    Analisa páginas específicas do PDF para identificar padrões de texto.
    Útil para descobrir como os títulos/legendas estão formatados.
    
    Args:
        pdf_path: Caminho do arquivo PDF
        page_numbers: Lista de números de página (1-indexed) para analisar
    """
    
    print(f"Abrindo PDF: {pdf_path}")
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"ERRO: Não foi possível abrir o PDF: {e}")
        return

    total_pages = len(doc)
    print(f"Total de páginas: {total_pages}")
    print(f"Páginas a analisar: {page_numbers}")
    print("=" * 80)
    
    for page_num_display in page_numbers:
        # Converter para índice 0-based
        page_num = page_num_display - 1
        
        # Validar se a página existe
        if page_num < 0 or page_num >= total_pages:
            print(f"\n⚠️  AVISO: Página {page_num_display} não existe (total: {total_pages})")
            continue
        
        page = doc.load_page(page_num)
        page_dict = page.get_text("dict")
        
        print(f"\n{'='*80}")
        print(f"PÁGINA {page_num + 1}")
        print(f"{'='*80}\n")
        
        # Coletar todos os textos com suas propriedades
        text_items = []
        
        for block in page_dict.get('blocks', []):
            if block.get('type') == 0:  # Bloco de texto
                for line in block.get('lines', []):
                    for span in line['spans']:
                        text = span['text'].strip()
                        if text:  # Ignorar textos vazios
                            text_items.append({
                                'text': text,
                                'size': round(span['size'], 2),
                                'color': span['color'],
                                'font': span['font'],
                                'y_position': round(line['bbox'][1], 2)
                            })
        
        # Procurar por padrões de "Figura"
        print("🔍 Textos contendo 'Figura' ou 'Gráfico':\n")
        found_figure = False
        
        for item in text_items:
            if 'figura' in item['text'].lower() or 'gráfico' in item['text'].lower():
                found_figure = True
                print(f"  📌 Texto: '{item['text']}'")
                print(f"     Tamanho: {item['size']} pt")
                print(f"     Cor: {item['color']} (hex: #{item['color']:06x})")
                print(f"     Fonte: {item['font']}")
                print(f"     Posição Y: {item['y_position']}")
                print()
        
        if not found_figure:
            print("  ⚠️  Nenhum texto com 'Figura' ou 'Gráfico' encontrado nesta página.\n")
        
        # Mostrar amostra de todos os textos (primeiros 10)
        print("📋 Amostra dos primeiros 10 textos da página:\n")
        for i, item in enumerate(text_items[:10]):
            print(f"  {i+1}. '{item['text'][:60]}{'...' if len(item['text']) > 60 else ''}'")
            print(f"     Size: {item['size']}pt | Color: #{item['color']:06x} | Font: {item['font']}")
            print()
        
        if len(text_items) > 10:
            print(f"  ... e mais {len(text_items) - 10} textos nesta página.\n")
    
    doc.close()
    print("\n" + "="*80)
    print("✅ Diagnóstico concluído!")
    print("\nUSE ESTAS INFORMAÇÕES para ajustar:")
    print("  - TITLE_FONT_SIZE (tamanho do título)")
    print("  - TITLE_TEXT_COLOR (cor do título)")
    print("  - TITLE_REGEX (padrão do texto)")

# --- EXECUÇÃO ---
if __name__ == "__main__":
    if not os.path.exists(PDF_PATH):
        print(f"ERRO: O arquivo '{PDF_PATH}' não foi encontrado.")
    else:
        diagnose_pdf_text(PDF_PATH, PAGES_TO_ANALYZE)

