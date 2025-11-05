import fitz # PyMuPDF
import os

# --- CONFIGURAÇÕES ---
PDF_PATH = "justica-em-numeros-2025.pdf"
# Defina o intervalo de páginas que você deseja analisar (baseado no número da página do PDF, começando em 1)
START_PAGE = 557 
END_PAGE = 559 
# ---------------------

def analyze_text_metadata_range(pdf_path: str, start_page: int, end_page: int):
    """
    Analisa o PDF para encontrar tamanhos de fonte e cores únicos dentro de um intervalo de páginas.
    """
    if not os.path.exists(pdf_path):
        print(f"ERRO: O arquivo '{pdf_path}' não foi encontrado.")
        return

    print(f"Iniciando análise de metadados de texto em '{pdf_path}'...")
    print(f"Analisando páginas de {start_page} a {end_page}...")
    
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"ERRO ao abrir PDF: {e}")
        return

    # Converte o número da página para o índice base zero do PyMuPDF
    start_index = start_page - 1 
    end_index = end_page # O PyMuPDF usa 'até', então o final é inclusivo se usarmos o índice final + 1
    
    # Validação de intervalo
    if start_index < 0 or start_index >= len(doc) or end_index > len(doc):
        print("ERRO: O intervalo de páginas configurado é inválido.")
        doc.close()
        return

    # { (size, color): [exemplos_de_texto, count] }
    unique_properties = {}

    for page_num in range(start_index, end_index):
        page = doc.load_page(page_num)
        
        page_dict = page.get_text("dict")
        
        for block in page_dict.get('blocks', []):
            if block.get('type') == 0: # Bloco de texto
                for line in block.get('lines', []):
                    for span in line.get('spans', []):
                        
                        size = round(span.get('size', 0), 2)
                        color = span.get('color', 0)
                        text = span.get('text', '').strip()
                        
                        key = (size, color)
                        
                        if len(text) > 3: 
                            if key not in unique_properties:
                                hex_color = f"#{color:06x}" 
                                unique_properties[key] = {
                                    "size": size,
                                    "color_dec": color,
                                    "color_hex": hex_color,
                                    "examples": set(),
                                    "count": 0
                                }
                            
                            unique_properties[key]["count"] += 1
                            if len(unique_properties[key]["examples"]) < 3:
                                unique_properties[key]["examples"].add(text)

    doc.close()

    print("\n" + "="*50)
    print(f"RESUMO DAS PROPRIEDADES DE TEXTO ENCONTRADAS (Páginas {start_page}-{end_page})")
    print("="*50)

    # Ordena as propriedades pelo tamanho da fonte
    sorted_properties = sorted(unique_properties.values(), key=lambda x: x['size'], reverse=True)

    for prop in sorted_properties:
        print(f"\nFONT SIZE: {prop['size']} (Total Ocorrências: {prop['count']})")
        print(f"COR (DECIMAL): {prop['color_dec']}")
        print(f"COR (HEX): {prop['color_hex']}")
        print(f"EXEMPLOS: {list(prop['examples'])}")
        print("-" * 30)

    print("\nCapture FONT SIZE e COR DECIMAL para a Legenda e o Corpo do Texto.")

# --- EXECUÇÃO ---

if __name__ == "__main__":
    # Altere as configurações no topo do arquivo
    analyze_text_metadata_range(PDF_PATH, START_PAGE, END_PAGE)