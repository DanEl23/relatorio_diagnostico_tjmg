import os
import re
import fitz # PyMuPDF
from google.cloud import documentai_v1 as documentai

# --- CONFIGURAÇÕES ---
# Substitua pelos seus dados do Google Cloud
PROJECT_ID = "diagnostico-tjmg"  # Ex: "nome-do-seu-projeto-12345"
LOCATION = "us-central1" # Ou a localização que você escolheu (ex: "us-central1")
PROCESSOR_ID = "e481e5e2c61d5fc4" # O ID do processador de OCR que você criou

# Arquivo PDF de entrada e pasta de saída
PDF_PATH = "justica-em-numeros-2025.pdf"
OUTPUT_DIR = "export_images"
# ---------------------

def extract_figures_with_documentai(pdf_path: str, project_id: str, location: str, processor_id: str, output_dir: str):
    """
    Chama a Document AI para extrair o layout, e então usa PyMuPDF para recortar 
    a figura nomeada pela legenda.
    """
    
    # 1. Configuração do Cliente da Document AI
    client = documentai.DocumentProcessorServiceClient(
        client_options={"api_endpoint": f"{location}-documentai.googleapis.com"}
    )
    
    processor_name = client.processor_path(project_id, location, processor_id)

    # Cria a pasta de saída se ela não existir
    os.makedirs(output_dir, exist_ok=True)
    
    # Carrega o PDF e o codifica em base64 para o Document AI
    with open(pdf_path, "rb") as image:
        image_content = image.read()
    
    raw_document = documentai.RawDocument(
        content=image_content,
        mime_type="application/pdf",
    )
    
    # Configure a solicitação de processamento
    request = documentai.ProcessRequest(
        name=processor_name,
        raw_document=raw_document,
    )

    # 2. Chamada da API Document AI
    print("Chamando Document AI para processar o PDF...")
    result = client.process_document(request=request)
    document = result.document
    print("Processamento concluído. Analisando o layout...")

    # Abre o PDF com PyMuPDF para a etapa de recorte (necessário para extrair a imagem)
    pdf_doc = fitz.open(pdf_path)
    extracted_count = 0

    # Itera pelas páginas do resultado da Document AI
    for page_index, page_docai in enumerate(document.pages):
        pdf_page = pdf_doc.load_page(page_index)
        
        # O Document AI fornece as 'tokens' (palavras) e suas coordenadas
        # O PyMuPDF, embora útil para o recorte, tem coordenadas baseadas em pixels.
        # As coordenadas do Document AI estão normalizadas (0 a 1000).

        # Usaremos as coordenadas normalizadas (0-1000) do Document AI para a lógica:
        for block in page_docai.blocks:
            
            # Tenta encontrar o texto completo do bloco no documento
            block_text = get_text_from_layout(document, block.layout)

            # 3. Filtragem e Associalçao de Legendas (via Regex)
            # Regex para encontrar "Figura X" ou "FIGURA X"
            match_legenda = re.search(r"^(figura\s+\d+|fig\.\s*\d+)", block_text, re.IGNORECASE)
            
            if match_legenda:
                # Extrai o nome limpo para uso como nome do arquivo
                fig_name = match_legenda.group(0).replace(":", "").strip()
                
                # Coordenadas normalizadas (0-1000) da legenda
                legenda_bbox_norm = get_bbox_from_layout(block.layout)
                
                if legenda_bbox_norm:
                    # 4. Inferência da Figura Associada (Lógica de Proximidade)
                    # Esta é uma HEURÍSTICA SIMPLES: procura um retângulo de figura logo ACIMA
                    # (y é menor no topo da página) ou logo ABAIXO da legenda.
                    
                    # Para simplificar o template, vamos procurar uma figura em uma área genérica
                    # acima da legenda. Em um código de produção, você teria que iterar
                    # sobre todos os elementos visuais (images, tables, etc.) retornados.
                    
                    # Devido à complexidade de encontrar a *imagem binária* no JSON do OCR
                    # (que prioriza texto e layout), o método mais confiável é:
                    # Procurar a região retangular que está logo acima da legenda.
                    
                    # HEURÍSTICA: Assumimos que a figura ocupa 80% da largura acima da legenda
                    # e tem altura de 30% da página.
                    
                    # Coordenadas do PyMuPDF (necessárias para o recorte)
                    page_width = pdf_page.rect.width
                    page_height = pdf_page.rect.height
                    
                    # Converte Y normalizado (0-1000) para Y de pixel do PyMuPDF
                    # y_inferior_figura = topo da legenda (legenda_bbox_norm[1] * page_height / 1000)
                    # y_superior_figura = 30% acima do topo da legenda (assumindo figura grande)
                    
                    # Recortar uma região retangular acima da legenda (Heurística de Layout)
                    y_top_norm = 100 # Inicia em 10% do topo da página
                    y_bottom_norm = legenda_bbox_norm[1] # Termina no topo da legenda
                    
                    # Converte coordenadas normalizadas (0-1000) para coordenadas de pixel (PyMuPDF)
                    clip_rect_px = fitz.Rect(
                        page_width * 0.05, # Margem de 5% à esquerda
                        y_top_norm * page_height / 1000,
                        page_width * 0.95, # Margem de 5% à direita
                        y_bottom_norm * page_height / 1000 
                    )
                    
                    # 5. Recorte e Salvar
                    try:
                        # Renderiza a área da figura em um pixmap (bitmap)
                        pix = pdf_page.get_pixmap(clip=clip_rect_px)
                        
                        # Nome do arquivo final
                        filename = f"{fig_name}_{page_index+1}.png"
                        output_path = os.path.join(output_dir, filename)
                        
                        # Salva o recorte como PNG
                        pix.save(output_path)
                        
                        print(f"-> Figura salva: {filename}")
                        extracted_count += 1
                        
                    except Exception as e:
                        print(f"Erro ao salvar figura {fig_name}: {e}")

    pdf_doc.close()
    print("-" * 30)
    print(f"Extração concluída. Total de figuras identificadas: {extracted_count}")


# --- FUNÇÕES AUXILIARES ---

def get_text_from_layout(document: documentai.Document, layout: documentai.Document.Page.Layout) -> str:
    """Extrai o texto associado a um layout/bloco."""
    text_segments = []
    for segment in layout.text_anchor.text_segments:
        start_index = int(segment.start_index)
        end_index = int(segment.end_index)
        text_segments.append(document.text[start_index:end_index])
    return "".join(text_segments)

def get_bbox_from_layout(layout: documentai.Document.Page.Layout) -> list:
    """Extrai o Bounding Box (coordenadas) de um layout normalizado (0-1000)."""
    if layout.bounding_box:
        vertices = layout.bounding_box.normalized_vertices
        # Coordenadas são retornadas como [x_min, y_min, x_max, y_max]
        if vertices:
            x_min = min(v.x for v in vertices)
            y_min = min(v.y for v in vertices)
            x_max = max(v.x for v in vertices)
            y_max = max(v.y for v in vertices)
            return [x_min * 1000, y_min * 1000, x_max * 1000, y_max * 1000] # Converte para 0-1000
    return None

# --- EXECUÇÃO ---

if __name__ == "__main__":
    # Verifique se todas as variáveis de configuração foram definidas
    if any(var in globals() and not globals()[var] for var in ['PROJECT_ID', 'PROCESSOR_ID']):
        print("ERRO: Por favor, preencha as variáveis PROJECT_ID e PROCESSOR_ID.")
    else:
        # Certifique-se de que o PDF de teste está na mesma pasta
        if not os.path.exists(PDF_PATH):
             print(f"AVISO: O arquivo de teste '{PDF_PATH}' não foi encontrado. Crie um PDF com este nome para testar.")
        else:
            extract_figures_with_documentai(PDF_PATH, PROJECT_ID, LOCATION, PROCESSOR_ID, OUTPUT_DIR)