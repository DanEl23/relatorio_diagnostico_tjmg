import fitz # PyMuPDF
import re
import os

# =====================================================================
#                        CONFIGURAÇÕES CHAVE
# =====================================================================
PDF_PATH = "justica-em-numeros-2024.pdf"
OUTPUT_DIR = "graficos_extraidos_por_titulo"

# PROPRIEDADES DO TÍTULO (baseado no diagnóstico real do PDF)
TITLE_FONT_SIZE = 9.0  # Títulos usam 9pt
TITLE_TEXT_COLOR = 5066063  # Cor cinza escuro (#4d4d4f)
TITLE_FONT_NAME = "AzoSans-Bold"  # Fonte específica dos títulos
TITLE_SIZE_TOLERANCE = 0.2  # Tolerância para aceitar tamanhos próximos (8.8-9.2pt)

# PROPRIEDADES DO TEXTO DO CORPO (para definir o limite da figura)
BODY_TEXT_COLOR = 0  # Cor preta (#000000) para textos normais
BODY_TEXT_MIN_SIZE = 8.0  # Textos com 8pt ou mais

# PROPRIEDADES DO RODAPÉ (Para limitar figuras em páginas únicas)
FOOTER_FONT_SIZE = 12.0  # Número da página
FOOTER_TEXT_COLOR = 5789816  # Cor do rodapé (#586878)
FOOTER_TEXT_MIN_Y_RATIO = 0.95  # Rodapés aparecem nos últimos 5% da página

# CONFIGURAÇÃO DE QUALIDADE
DPI_FACTOR = 3.0 # Fator de zoom: 3.0 (equivalente a 300 DPI) para alta qualidade
# =====================================================================


def extract_figures_by_title(pdf_path: str, output_dir: str):
    """
    Extrai figuras usando o TÍTULO (Figura XX -) ao invés da legenda.
    O título aparece ANTES da figura, então capturamos da figura até o próximo texto.
    """
    
    print(f"Abrindo PDF: {pdf_path}")
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"ERRO: Não foi possível abrir o PDF: {e}")
        return

    os.makedirs(output_dir, exist_ok=True)
    extracted_count = 0
    
    # Regex para capturar "Figura" no título (formato: "Figura 1 - Título")
    TITLE_REGEX = r"(?i)^(Figura\s+\d+\.?\s*-?\s*)(.*)"
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        page_dict = page.get_text("dict")
        title_info = None  # Guarda info do título (nome e posição Y final)
        zoom_matrix = fitz.Matrix(DPI_FACTOR, DPI_FACTOR)
        
        # --- PRÉ-ANÁLISE DE RODAPÉS ---
        footer_limit_y0 = page.rect.y1
        
        for block in page_dict.get('blocks', []):
            if block.get('type') == 0:
                for line in block.get('lines', []):
                    first_span = line['spans'][0] if line['spans'] else None
                    if not first_span: continue

                    size = round(first_span['size'], 2)
                    color = first_span['color']
                    
                    is_footer = (
                        abs(size - FOOTER_FONT_SIZE) < 0.1 and 
                        color == FOOTER_TEXT_COLOR and 
                        line['bbox'][1] > page.rect.y1 * FOOTER_TEXT_MIN_Y_RATIO
                    )
                    
                    if is_footer and line['bbox'][1] < footer_limit_y0:
                        footer_limit_y0 = line['bbox'][1]

        # --- LOOP PRINCIPAL: ENCONTRAR TÍTULO E LIMITE ---
        for block_idx, block in enumerate(page_dict.get('blocks', [])):
            if block.get('type') == 0:
                for line in block.get('lines', []):
                    
                    first_span = line['spans'][0] if line['spans'] else None
                    if not first_span: continue
                        
                    size = round(first_span['size'], 2)
                    color = first_span['color']
                    font = first_span['font']
                    text = first_span['text'].strip()
                    
                    match = re.search(TITLE_REGEX, text)
                    
                    # --- CONDIÇÃO DE TÍTULO ---
                    is_title = (
                        match and 
                        color == TITLE_TEXT_COLOR and 
                        abs(size - TITLE_FONT_SIZE) <= TITLE_SIZE_TOLERANCE and
                        font == TITLE_FONT_NAME and
                        len(text) > 8  # Mínimo de caracteres para ser título válido
                    )
                    
                    if is_title:
                        # Salvar figura anterior se existir
                        if title_info is not None:
                            # A figura vai do título anterior até ANTES deste novo título
                            clip_y_end = line['bbox'][1] - 5  # 5px de margem
                            y_start = title_info['y_end'] + 5  # Começa após o título
                            
                            clip_rect = fitz.Rect(
                                page.rect.x0, y_start, page.rect.x1, clip_y_end
                            )
                            
                            if clip_rect.height > 10 and clip_rect.width > 10:
                                try:
                                    pix = page.get_pixmap(matrix=zoom_matrix, clip=clip_rect)
                                    
                                    filename = f"{title_info['name']}_Pg{page_num + 1}.png"
                                    output_path = os.path.join(output_dir, filename)
                                    pix.save(output_path)
                                    
                                    print(f"-> Gráfico salvo (300 DPI): {filename}")
                                    extracted_count += 1
                                except Exception as e:
                                    print(f"Erro ao salvar figura {title_info['name']}: {e}")
                        
                        # Processar novo título
                        label = match.group(1).strip()
                        title = match.group(2).strip()
                        
                        # Substituir "Figura" por "Gráfico" no nome de salvamento
                        label = label.replace("Figura", "Gráfico").replace("figura", "Gráfico").replace("Fig.", "Graf.")
                        # Remover hífen extra
                        label = label.rstrip('-').strip()
                        
                        fig_name_raw = f"{label} - {title}" if title else label
                        fig_name = re.sub(r'[\\/:*?"<>|]', '', fig_name_raw)[:80].strip()
                        
                        title_info = {
                            'name': fig_name,
                            'y_end': line['bbox'][3]  # Posição Y final do título
                        }
                        
                        print(f"Página {page_num + 1}: Título Encontrado: {fig_name}")
                        continue
        
        # --- TRATAR ÚLTIMA FIGURA DA PÁGINA ---
        if title_info is not None:
            # Usar limite do rodapé ou margem fixa
            final_clip_y_end = footer_limit_y0 - 5 if footer_limit_y0 < page.rect.y1 else page.rect.y1 - 20
            
            y_start = title_info['y_end'] + 5
            
            clip_rect = fitz.Rect(page.rect.x0, y_start, page.rect.x1, final_clip_y_end)
            
            if clip_rect.height > 10:
                try:
                    pix = page.get_pixmap(matrix=zoom_matrix, clip=clip_rect)
                    filename = f"{title_info['name']}_Pg{page_num + 1}_FINAL.png"
                    output_path = os.path.join(output_dir, filename)
                    pix.save(output_path)
                    
                    status = "Limite: Rodapé" if footer_limit_y0 < page.rect.y1 else "Limite: Margem Fixa"
                    print(f"-> Gráfico salvo (300 DPI): {filename} ({status})")
                    extracted_count += 1
                except Exception as e:
                    print(f"Erro ao salvar figura {title_info['name']} no final da Pág {page_num + 1}: {e}")
                    
            # Resetar para próxima página
            title_info = None

    doc.close()
    print("-" * 50)
    print(f"Extração concluída. Total de gráficos identificados e salvos: {extracted_count}")

# --- EXECUÇÃO ---

if __name__ == "__main__":
    if not os.path.exists(PDF_PATH):
         print(f"ERRO: O arquivo de teste '{PDF_PATH}' não foi encontrado.")
    else:
        extract_figures_by_title(PDF_PATH, OUTPUT_DIR)
