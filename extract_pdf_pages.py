"""
Script para extrair páginas específicas de um PDF e criar um novo arquivo.
Extrai intervalos de páginas do documento "Justiça em Números 2025".
"""

import fitz  # PyMuPDF
import os

# =====================================================================
#                        CONFIGURAÇÕES
# =====================================================================
PDF_INPUT = "justica-em-numeros-2025.pdf"
PDF_OUTPUT = "justica-em-numeros-2025-selecionado.pdf"

# Intervalos de páginas a serem extraídos (1-indexed, como aparecem no PDF)
# Nota: PyMuPDF usa índice 0-based, então subtraímos 1
PAGE_RANGES = [
    (1, 71),      # Páginas 1 a 71
    (98, 133),    # Páginas 98 a 133
    (241, 285),   # Páginas 241 a 285
    (316, 398),   # Páginas 316 a 398
    (587, None),  # Página 587 até o final (None = última página)
]

# =====================================================================


def extract_pdf_pages(input_path, output_path, page_ranges):
    """
    Extrai páginas específicas de um PDF e cria um novo arquivo.
    
    Args:
        input_path: Caminho do PDF de entrada
        output_path: Caminho do PDF de saída
        page_ranges: Lista de tuplas (início, fim) com números de página (1-indexed)
                     Use None no fim para indicar última página
    """
    
    print(f"Abrindo PDF: {input_path}")
    
    try:
        doc = fitz.open(input_path)
    except Exception as e:
        print(f"ERRO: Não foi possível abrir o PDF: {e}")
        return
    
    total_pages = len(doc)
    print(f"Total de páginas no documento: {total_pages}")
    print("-" * 60)
    
    # Criar novo documento PDF
    output_doc = fitz.open()
    
    total_extracted = 0
    
    # Processar cada intervalo
    for start, end in page_ranges:
        # Converter para índice 0-based
        start_idx = start - 1
        
        # Se end é None, usar última página
        if end is None:
            end_idx = total_pages - 1
            end = total_pages
        else:
            end_idx = end - 1
        
        # Validar intervalos
        if start_idx < 0 or start_idx >= total_pages:
            print(f"⚠️  AVISO: Página inicial {start} está fora do intervalo (1-{total_pages})")
            continue
        
        if end_idx < 0 or end_idx >= total_pages:
            print(f"⚠️  AVISO: Página final {end} está fora do intervalo (1-{total_pages})")
            continue
        
        if start_idx > end_idx:
            print(f"⚠️  AVISO: Intervalo inválido ({start}-{end})")
            continue
        
        # Extrair páginas
        num_pages = end_idx - start_idx + 1
        print(f"Extraindo páginas {start} a {end} ({num_pages} páginas)...")
        
        # Inserir páginas no novo documento
        output_doc.insert_pdf(doc, from_page=start_idx, to_page=end_idx)
        
        total_extracted += num_pages
    
    # Salvar novo PDF
    print("-" * 60)
    print(f"Salvando PDF com {total_extracted} páginas extraídas...")
    
    try:
        output_doc.save(output_path)
        output_doc.close()
        print(f"✅ PDF criado com sucesso: {output_path}")
    except Exception as e:
        print(f"ERRO ao salvar PDF: {e}")
        return
    
    doc.close()
    print("-" * 60)
    print("Resumo:")
    print(f"  - Documento original: {total_pages} páginas")
    print(f"  - Páginas extraídas: {total_extracted} páginas")
    print(f"  - Arquivo de saída: {output_path}")


# --- EXECUÇÃO ---
if __name__ == "__main__":
    if not os.path.exists(PDF_INPUT):
        print(f"ERRO: O arquivo '{PDF_INPUT}' não foi encontrado.")
        print("Certifique-se de que o arquivo está no mesmo diretório do script.")
    else:
        extract_pdf_pages(PDF_INPUT, PDF_OUTPUT, PAGE_RANGES)
        print("\n✨ Extração concluída!")
