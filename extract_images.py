import fitz # PyMuPDF
import re
import os

# =====================================================================
#                        CONFIGURAÇÕES CHAVE
# =====================================================================
PDF_PATH = "justica-em-numeros-2025.pdf"
OUTPUT_DIR = "figuras_extraidas_metadados_alta_qualidade"

# PROPRIEDADES DA LEGENDA (Baseado no seu diagnóstico)
LEGEND_FONT_SIZE = 11.0
LEGEND_TEXT_COLOR = 37509 # Cor verde/ciano (#009285)

# PROPRIEDADES DO TEXTO DO CORPO (para definir o limite da figura)
BODY_TEXT_COLOR = 2301728 # Cor escura (#231f20)
BODY_TEXT_MIN_SIZE = 11.9 # Qualquer texto significativo acima de 11.9 será considerado corpo

# PROPRIEDADES DO RODAPÉ (Para limitar figuras em páginas únicas)
FOOTER_FONT_SIZE = 9.0 
FOOTER_TEXT_COLOR = 2301728 
FOOTER_TEXT_MIN_Y_RATIO = 0.9 # Rodapés aparecem tipicamente abaixo de 90% da altura da página

# CONFIGURAÇÃO DE QUALIDADE
DPI_FACTOR = 3.0 # Fator de zoom: 3.0 (equivalente a 300 DPI) para alta qualidade
# =====================================================================


def extract_figures_by_metadata(pdf_path: str, output_dir: str):
    """
    Extrai figuras usando metadados de texto para definir limites precisos.
    Garante alta qualidade (300 DPI) e remove espaços em branco de rodapés 
    em páginas de gráficos únicos.
    """
    
    print(f"Abrindo PDF: {pdf_path}")
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"ERRO: Não foi possível abrir o PDF: {e}")
        return

    os.makedirs(output_dir, exist_ok=True)
    extracted_count = 0
    
    LEGEND_REGEX = r"(?i)\b(Figura\s+\d+\.?|Fig\.\s*\d+\.?)(.*)"
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        page_dict = page.get_text("dict")
        last_legend_y1 = -1 
        zoom_matrix = fitz.Matrix(DPI_FACTOR, DPI_FACTOR)
        
        # --- NOVO: Pré-análise de Rodapés (Para Limitar o Fim da Página) ---
        footer_limit_y0 = page.rect.y1 # Limite padrão: fundo total
        
        # Encontra o elemento de rodapé mais alto na página para usá-lo como limite
        for block in page_dict.get('blocks', []):
            if block.get('type') == 0:
                for line in block.get('lines', []):
                    first_span = line['spans'][0] if line['spans'] else None
                    if not first_span: continue

                    size = round(first_span['size'], 2)
                    color = first_span['color']
                    
                    # Condição: É texto de rodapé (baseado em tamanho, cor e posição)
                    is_footer = (
                        abs(size - FOOTER_FONT_SIZE) < 0.1 and 
                        color == FOOTER_TEXT_COLOR and 
                        line['bbox'][1] > page.rect.y1 * FOOTER_TEXT_MIN_Y_RATIO
                    )
                    
                    if is_footer and line['bbox'][1] < footer_limit_y0:
                        footer_limit_y0 = line['bbox'][1] # Captura o topo do rodapé mais alto


        # 2. Loop Principal: Encontrar Legenda e Limite
        for block_idx, block in enumerate(page_dict.get('blocks', [])):
            if block.get('type') == 0:
                for line in block.get('lines', []):
                    
                    first_span = line['spans'][0] if line['spans'] else None
                    if not first_span: continue
                        
                    size = round(first_span['size'], 2)
                    color = first_span['color']
                    text = first_span['text']
                    
                    match = re.search(LEGEND_REGEX, text.strip())
                    
                    # --- Condição de Legenda ---
                    is_legend = (
                        match and 
                        color == LEGEND_TEXT_COLOR and 
                        abs(size - LEGEND_FONT_SIZE) < 0.1
                    )
                    
                    if is_legend:
                        last_legend_y1 = line['bbox'][3]
                        
                        label = match.group(1).strip()
                        title = match.group(2).strip()
                        fig_name_raw = f"{label} {title}"
                        fig_name = re.sub(r'[\\/:*?"<>|.]', '', fig_name_raw)[:80].strip()
                        
                        print(f"Página {page_num + 1}: Legenda Encontrada por Metadado: {fig_name}")
                        continue
                        
                    # --- Condição de Limite Inferior (BODY_TEXT) ---
                    if last_legend_y1 != -1:
                        
                        is_body_text = (
                            line['bbox'][1] > last_legend_y1 and 
                            color == BODY_TEXT_COLOR and 
                            size >= BODY_TEXT_MIN_SIZE and
                            len(text.strip()) > 20
                        )
                        
                        if is_body_text:
                            # BODY_TEXT ENCONTRADO: Recorte Normal (Resolve a figura anterior)
                            
                            clip_y_end = line['bbox'][1] - 1
                            y_start = last_legend_y1 + 1
                            
                            clip_rect = fitz.Rect(
                                page.rect.x0, y_start, page.rect.x1, clip_y_end
                            )
                            
                            # 3. Recorte e Salvar
                            if clip_rect.height > 10 and clip_rect.width > 10:
                                try:
                                    pix = page.get_pixmap(matrix=zoom_matrix, clip=clip_rect)
                                    
                                    filename = f"{fig_name}_Pg{page_num + 1}.png"
                                    output_path = os.path.join(output_dir, filename)
                                    pix.save(output_path)
                                    
                                    print(f"-> Figura salva (300 DPI): {filename}")
                                    extracted_count += 1
                                except Exception as e:
                                    print(f"Erro ao salvar figura {fig_name}: {e}")
                            
                            # Resetar para a próxima legenda
                            last_legend_y1 = -1
                            
        # 4. Tratar Última Figura da Página / Figura Única (USANDO LIMITE DE RODAPÉ)
        if last_legend_y1 != -1:
            
            # Se a figura vai até o final (sem BODY_TEXT abaixo), usamos o limite do rodapé.
            # O limite será o topo do rodapé, menos uma margem, OU o final da página.
            
            # Se o rodapé foi detectado, final_clip_y_end será o topo dele - 5px.
            final_clip_y_end = footer_limit_y0 - 5 if footer_limit_y0 < page.rect.y1 else page.rect.y1 - 20
            
            y_start = last_legend_y1 + 1
            
            clip_rect = fitz.Rect(page.rect.x0, y_start, page.rect.x1, final_clip_y_end)
            
            if clip_rect.height > 10:
                try:
                    pix = page.get_pixmap(matrix=zoom_matrix, clip=clip_rect)
                    filename = f"{fig_name}_Pg{page_num + 1}_FINAL.png"
                    output_path = os.path.join(output_dir, filename)
                    pix.save(output_path)
                    
                    status = "Limite: Rodapé" if footer_limit_y0 < page.rect.y1 else "Limite: Margem Fixa"
                    print(f"-> Figura salva (300 DPI): {filename} ({status})")
                    extracted_count += 1
                except Exception as e:
                    print(f"Erro ao salvar figura {fig_name} no final da Pág {page_num + 1}: {e}")

    doc.close()
    print("-" * 50)
    print(f"Extração concluída. Total de figuras identificadas e salvas: {extracted_count}")

# --- EXECUÇÃO ---

if __name__ == "__main__":
    if not os.path.exists(PDF_PATH):
         print(f"ERRO: O arquivo de teste '{PDF_PATH}' não foi encontrado.")
    else:
        extract_figures_by_metadata(PDF_PATH, OUTPUT_DIR)