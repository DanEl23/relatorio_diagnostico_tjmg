import json
import re 
import os 
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT 
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

# --- 1. IMPORTAÇÃO DOS DADOS EXTERNOS ---
try:
    from report_data import dados_tabela_atos, dados_tabela_areas, dados_tabela_estrutura, dados_tabela_comarcas, dados_tabela_nucleos, dados_tabela_historicos, MAPA_IMAGENS
except ImportError:
    print("!!! ERRO CRÍTICO: Não foi possível encontrar o arquivo 'report_data.py'.")
    print("!!! Certifique-se de que 'report_data.py' está no mesmo diretório.")
    exit()

# --- 2. DEFINIÇÃO DAS EXPRESSÕES REGULARES ---
PATTERN_SUMARIO = r'^\s*(\d+(?:\.\d+)*)\.?\s+(.*?)\s*[\. ]*\d+$' 
PATTERN_CONTEUDO = r'^\s*(\d+(\.\d+)*)\.?\s*(.*)$'
PATTERN_LEGENDA = r'^(Figura|Gráfico)\s+\d+' 

# --- 3. DADOS BRUTOS (HARDCODED) ---
pass

# --- 4. FUNÇÕES AUXILIARES (PAGINAÇÃO, ALINHAMENTO, XML) ---
def configurar_margens(documento, superior_cm, esquerda_cm, direita_cm, inferior_cm):
    """ Define as margens da seção principal do documento em centímetros. """
    # Assume que estamos trabalhando na primeira seção do documento
    section = documento.sections[0]
    
    section.top_margin = Cm(superior_cm)
    section.left_margin = Cm(esquerda_cm)
    section.right_margin = Cm(direita_cm)
    section.bottom_margin = Cm(inferior_cm)
    
    print(f"Margens definidas: Superior={superior_cm}cm, Esquerda={esquerda_cm}cm.")


def set_row_height_exact(row, height_twips):
    """ Define a altura exata da linha usando XML (twips). """
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(height_twips))
    trHeight.set(qn('w:hRule'), 'exact') 
    
    for existing_trHeight in trPr.findall(qn('w:trHeight')):
        trPr.remove(existing_trHeight)
        
    trPr.append(trHeight)


def set_cell_bottom_border(cell):
    """ Adiciona uma borda inferior sólida (preta, 0.5pt) a uma célula específica. """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    
    bottom_border = OxmlElement('w:bottom') 
    bottom_border.set(qn('w:val'), 'single') 
    bottom_border.set(qn('w:sz'), '4')       
    bottom_border.set(qn('w:color'), '000000') 
    
    tcBorders.append(bottom_border)
    tcPr.append(tcBorders)


def add_page_number(footer):
    """ Adiciona o código de campo de paginação (PAGE) a um rodapé (footer). """
    p = footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.text = 'PAGE' 
    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')

    run_begin = p.add_run()
    run_begin.element.append(fldChar_begin)
    run_instr = p.add_run()
    run_instr.element.append(instrText)
    run_end = p.add_run()
    run_end.element.append(fldChar_end)


def set_cell_vertical_alignment(cell, align):
    """Define o alinhamento vertical de uma célula usando w:vAlign (XML)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), align) 
    tcPr.append(vAlign)


def limpar_espacamento_lista(paragraph):
    """ Remove o espaçamento extra de um parágrafo de lista (Tabela 01) usando XML. """
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    
    pPr = paragraph._element.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), '240') 
    spacing.set(qn('w:lineRule'), 'auto')
    
    if hasattr(pPr, 'spacing'):
        try:
            pPr.remove(pPr.spacing)
        except:
            pass
    pPr.append(spacing)

def set_group_top_border(cell):
    """ Adiciona uma borda superior sólida (preta, 0.5pt) a uma célula específica. """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    
    top_border = OxmlElement('w:top') 
    top_border.set(qn('w:val'), 'single') 
    top_border.set(qn('w:sz'), '4')       
    top_border.set(qn('w:color'), '000000') 
    
    tcBorders.append(top_border)
    tcPr.append(tcBorders)

# --- 5. FUNÇÕES DE PROCESSAMENTO E CRIAÇÃO ---

def extrair_sumario_para_json(caminho_arquivo_docx, pattern_regex):
    """Lê o DOCX (Sumário) e extrai títulos (Gerador de Chave Limpa)."""
    try:
        documento_sumario = Document(caminho_arquivo_docx)
    except Exception as e:
        print(f"!!! ERRO CRÍTICO: Não foi possível carregar o arquivo: {caminho_arquivo_docx}")
        print(f"!!! Detalhe do Erro: {e}")
        return [] 

    estrutura_do_relatorio = [] 

    for paragrafo in documento_sumario.paragraphs:
        texto_completo = paragrafo.text
        texto_limpo = texto_completo.replace('\xa0', ' ').strip().replace('SUMÁRIO', '')
        
        if not texto_limpo:
            continue
        
        match = re.search(pattern_regex, texto_limpo, re.IGNORECASE)
        
        if match:
            prefixo_completo = match.group(1).strip()
            texto_titulo = match.group(2).strip()
            level = len(prefixo_completo.split('.'))
            
            texto_final_com_numero = f"{prefixo_completo} {texto_titulo}"

            if level >= 1:
                estrutura_do_relatorio.append({
                    "tipo": "TITULO",
                    "level": level,
                    "texto": texto_final_com_numero 
                })
    return estrutura_do_relatorio


def extrair_conteudo_mapeado(caminho_arquivo_docx, pattern_titulo, pattern_legenda):
    """
    Lê o DOCX de conteúdo, mapeia parágrafos e identifica marcadores.
    (Atualizado para Tabela 05)
    """
    print(f"Iniciando extração de conteúdo de: {caminho_arquivo_docx}")
    
    conteudo_mapeado = {}
    chave_titulo_atual = None
    
    try:
        documento_conteudo = Document(caminho_arquivo_docx)
    except Exception as e:
        print(f"!!! ERRO CRÍTICO: Não foi possível carregar o arquivo de CONTEÚDO: {caminho_arquivo_docx}")
        print(f"!!! Detalhe do Erro: {e}")
        return {} 

    for paragrafo in documento_conteudo.paragraphs:
        texto = paragrafo.text.replace('\xa0', ' ').strip()
        
        if not texto:
            continue 

        match_titulo = re.search(pattern_titulo, texto, re.IGNORECASE)
        
        if match_titulo:
            prefixo = match_titulo.group(1).strip()
            titulo = match_titulo.group(3).strip()
            chave_titulo_atual = f"{prefixo} {titulo}" 
            conteudo_mapeado[chave_titulo_atual] = []
            
        elif chave_titulo_atual:
            
            if texto == "[INSERIR_TABELA_ATOS_NORMATIVOS]":
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "TABELA_ATOS",
                    "dados": dados_tabela_atos 
                })
            
            elif texto == "[INSERIR_TABELA_AREAS]":
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "TABELA_AREAS",
                    "dados": dados_tabela_areas
                })
            
            elif texto == "[INSERIR_TABELA_ESTRUTURA]":
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "TABELA_ESTRUTURA",
                    "dados": dados_tabela_estrutura
                })

            elif texto == "[INSERIR_TABELA_COMARCAS]":
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "TABELA_COMARCAS",
                    "dados": dados_tabela_comarcas
                })
            
            elif texto == "[INSERIR_TABELA_NUCLEOS]":
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "TABELA_NUCLEOS",
                    "dados": dados_tabela_nucleos
                })
                
            elif texto == "[INSERIR_TABELA_HISTORICOS]":
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "TABELA_HISTORICOS",
                    "dados": dados_tabela_historicos
                })

            elif re.search(pattern_legenda, texto, re.IGNORECASE):
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "FIGURA",
                    "legenda_completa": texto 
                })
  
            else:
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "PARAGRAFO",
                    "texto": texto
                })

    print(f"Extração de conteúdo concluída. {len(conteudo_mapeado)} títulos mapeados.")
    return conteudo_mapeado


def customizar_estilos_titulo(documento):
    """Aplica formatação personalizada (Copiado da sua base)."""
    
    style_h1 = documento.styles['Heading 1']
    font_h1 = style_h1.font
    font_h1.name = 'Calibri' 
    font_h1.size = Pt(18) 
    font_h1.color.rgb = RGBColor(162, 22, 18) 
    font_h1.all_caps = True 
    font_h1.bold = True
    p_format_h1 = style_h1.paragraph_format
    p_format_h1.space_after = Pt(12)
    
    style_h2 = documento.styles['Heading 2']
    font_h2 = style_h2.font
    font_h2.name = 'Calibri' 
    font_h2.size = Pt(16)
    font_h2.color.rgb = RGBColor(162, 22, 18) 
    font_h2.bold = True
    font_h2.all_caps = False
    p_format_h2 = style_h2.paragraph_format
    p_format_h2.space_after = Pt(10)
    
    style_h3 = documento.styles['Heading 3']
    font_h3 = style_h3.font
    font_h3.name = 'Calibri' 
    font_h3.size = Pt(15.5)
    font_h3.color.rgb = RGBColor(162, 22, 18) 
    font_h3.bold = True
    p_format_h3 = style_h3.paragraph_format
    p_format_h3.space_after = Pt(8)

    style_normal = documento.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Calibri'
    font_normal.size = Pt(12) 

# --- COMPONENTE: TABELA 01 (ATOS) ---
def adicionar_tabela_atos(document, dados):
    """ Cria e estiliza a Tabela de Atos Normativos. """
    
    COR_CABECALHO_RGB = RGBColor(127, 127, 127)   
    COR_CABECALHO_HEX = '7F7F7F'                  
    COR_BRANCO_RGB = RGBColor(255, 255, 255)       
    COR_PRETO_RGB = RGBColor(0, 0, 0) 
    COR_CINZA_CLARO_HEX = 'EEEEEE'                 

    TAMANHO_FONTE_PADRAO = Pt(12) 
    FONTE = 'Calibri'
    
    table = document.add_table(rows=1, cols=len(dados[0]))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.indent = Cm(1)
    
    col_widths = [Cm(5), Cm(16)]
    for i, width in enumerate(col_widths):
        table.columns[i].width = width

    for i, row_data in enumerate(dados):
        if i > 0:
            row = table.add_row()
        else:
            row = table.rows[0]
            
        tr = row._tr 
        trPr = tr.get_or_add_trPr() 

        if i == 0:
            tblHeader = OxmlElement('w:tblHeader')
            trPr.append(tblHeader)
        else:
            cantSplit = OxmlElement('w:cantSplit')
            cantSplit.set(qn('w:val'), '0') 
            trPr.append(cantSplit)
            
        for j, cell_data in enumerate(row_data):
            cell = row.cells[j]
            set_cell_vertical_alignment(cell, 'center') 
            cell.text = "" 
            
            lines = cell_data.split('\n')
            is_first_content_line = True

            for k, line in enumerate(lines):
                line = line.strip()
                if not line:
                    continue 

                is_list_item = line.startswith('ü')
                
                if is_first_content_line:
                    current_paragraph = cell.paragraphs[0]
                    is_first_content_line = False
                else:
                    current_paragraph = cell.add_paragraph()

                text_to_insert = line.replace('ü', '').strip() if is_list_item else line
                run = current_paragraph.add_run(text_to_insert)

                if is_list_item:
                    current_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    current_paragraph.style = 'List Bullet' 
                    limpar_espacamento_lista(current_paragraph)
                    
                elif i == 0:
                    current_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    current_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT 
                    current_paragraph.paragraph_format.space_after = Pt(8) 
                    current_paragraph.paragraph_format.line_spacing = 1
                
                run.font.name = FONTE
                run.font.size = TAMANHO_FONTE_PADRAO
                
                if i == 0:
                    run.font.color.rgb = COR_BRANCO_RGB 
                    run.bold = True
                else:
                    run.font.color.rgb = COR_PRETO_RGB 
                    run.bold = False
            
            if i == 0:
                shading_elm = OxmlElement('w:shd')
                shading_elm.set(qn('w:fill'), COR_CABECALHO_HEX)
                cell._tc.get_or_add_tcPr().append(shading_elm)
            else:
                if i % 2 == 0:
                   shading_elm = OxmlElement('w:shd')
                   shading_elm.set(qn('w:fill'), COR_CINZA_CLARO_HEX)
                   cell._tc.get_or_add_tcPr().append(shading_elm)

    p_titulo_tabela = document.add_paragraph(style='Normal')
    p_titulo_tabela.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_titulo_tabela.paragraph_format.space_before = Pt(6)
    run_titulo = p_titulo_tabela.add_run("Tabela 01 - Atos Normativos referentes à Estrutura do TJMG. Fonte: Portal TJMG")
    run_titulo.bold = False 
    run_titulo.font.name = FONTE
    run_titulo.font.size = Pt(8)
    p_titulo_tabela.paragraph_format.space_after = Pt(12) 


# --- COMPONENTE: TABELA 02 (ÁREAS) ---
def adicionar_tabela_areas(document, dados):
    """ Cria a Tabela 02 com alinhamento esquerdo, largura de 3cm, bordas de grupo e espaçamento de linha 1.0. """
    
    COR_HEADER_MAIN = '7F7F7F'    
    COR_HEADER_GROUP = 'D9D9D9'  
    COR_LINHA_ZEBRADA = 'EEEEEE'
    COR_BRANCO_RGB = RGBColor(255, 255, 255)       
    COR_PRETO_RGB = RGBColor(0, 0, 0) 
    TAMANHO_FONTE_PADRAO = Pt(11) 
    FONTE = 'Calibri'
    
    table = document.add_table(rows=0, cols=2)
    
    try:
        table.style = 'Normal Table'
    except KeyError:
        print("Aviso: Estilo 'Normal Table' não encontrado. Revertendo para 'Table Grid'.")
        table.style = 'Table Grid' 
    
    col_widths = [Cm(14.5), Cm(3.0)] 
    table.columns[0].width = col_widths[0]
    table.columns[1].width = col_widths[1]

    data_row_index = 0 

    for i, row_data in enumerate(dados):
        
        tipo_linha = row_data[0]
        col1_texto = row_data[1]
        col2_texto = row_data[2]
        
        row = table.add_row()
        
        tr = row._tr 
        trPr = tr.get_or_add_trPr() 
        if tipo_linha.startswith("HEADER"):
            tblHeader = OxmlElement('w:tblHeader')
            trPr.append(tblHeader)
        else:
            cantSplit = OxmlElement('w:cantSplit')
            cantSplit.set(qn('w:val'), '0') 
            trPr.append(cantSplit)
        
        cell1 = row.cells[0]
        cell2 = row.cells[1]
        
        if tipo_linha == "HEADER_MAIN":
            cell1.merge(cell2)
            cell = cell1
            cell.text = col1_texto
            
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), COR_HEADER_MAIN)
            cell._tc.get_or_add_tcPr().append(shading_elm)
            
            set_cell_vertical_alignment(cell, 'center')
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            p_format = p.paragraph_format
            p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p_format.space_before = Pt(0)
            p_format.space_after = Pt(0)
            
            run = p.runs[0]
            run.font.color.rgb = COR_BRANCO_RGB
            run.bold = True
            run.font.name = FONTE
            run.font.size = TAMANHO_FONTE_PADRAO
            
        elif tipo_linha == "HEADER_GROUP_SIGLA":
            for j, cell in enumerate([cell1, cell2]):
                cell.text = col1_texto if j == 0 else col2_texto
                shading_elm = OxmlElement('w:shd')
                shading_elm.set(qn('w:fill'), COR_HEADER_GROUP)
                cell._tc.get_or_add_tcPr().append(shading_elm)
                
                set_cell_vertical_alignment(cell, 'center')
                p = cell.paragraphs[0]
                run = p.runs[0]
                run.font.color.rgb = COR_PRETO_RGB
                run.bold = True
                run.font.name = FONTE
                run.font.size = TAMANHO_FONTE_PADRAO
                
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
                set_group_top_border(cell) 

                p_format = p.paragraph_format
                p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p_format.space_before = Pt(0)
                p_format.space_after = Pt(0)

        elif tipo_linha == "HEADER_GROUP_MERGED":
            cell1.merge(cell2)
            cell = cell1
            cell.text = col1_texto
            
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), COR_HEADER_GROUP)
            cell._tc.get_or_add_tcPr().append(shading_elm)
            
            set_cell_vertical_alignment(cell, 'center')
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.runs[0]
            run.font.color.rgb = COR_PRETO_RGB
            run.bold = True
            run.font.name = FONTE
            run.font.size = TAMANHO_FONTE_PADRAO
            
            set_group_top_border(cell) 

            p_format = p.paragraph_format
            p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p_format.space_before = Pt(0)
            p_format.space_after = Pt(0)

        elif tipo_linha == "DATA_MERGED":
            data_row_index += 1
            cell1.merge(cell2)
            cell = cell1
            cell.text = col1_texto
            
            set_cell_vertical_alignment(cell, 'center')
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.runs[0]
            run.font.color.rgb = COR_PRETO_RGB
            run.font.name = FONTE
            run.font.size = TAMANHO_FONTE_PADRAO

            p_format = p.paragraph_format
            p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p_format.space_before = Pt(0)
            p_format.space_after = Pt(0)

            if data_row_index % 2 == 0:
                shading_elm = OxmlElement('w:shd')
                shading_elm.set(qn('w:fill'), COR_LINHA_ZEBRADA)
                cell._tc.get_or_add_tcPr().append(shading_elm)

        elif tipo_linha == "DATA_SPLIT":
            data_row_index += 1
            
            cell1.text = col1_texto
            set_cell_vertical_alignment(cell1, 'center')
            p1 = cell1.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run1 = p1.runs[0]
            run1.font.color.rgb = COR_PRETO_RGB
            run1.font.name = FONTE
            run1.font.size = TAMANHO_FONTE_PADRAO
            
            p1_format = p1.paragraph_format
            p1_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p1_format.space_before = Pt(0)
            p1_format.space_after = Pt(0)

            cell2.text = col2_texto
            set_cell_vertical_alignment(cell2, 'center')
            p2 = cell2.paragraphs[0]
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run2 = p2.runs[0]
            run2.font.color.rgb = COR_PRETO_RGB
            run2.font.name = FONTE
            run2.font.size = TAMANHO_FONTE_PADRAO
            
            p2_format = p2.paragraph_format
            p2_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p2_format.space_before = Pt(0)
            p2_format.space_after = Pt(0)
            
            if data_row_index % 2 == 0:
                for cell in [cell1, cell2]:
                    shading_elm = OxmlElement('w:shd')
                    shading_elm.set(qn('w:fill'), COR_LINHA_ZEBRADA)
                    cell._tc.get_or_add_tcPr().append(shading_elm)

    p_titulo_tabela = document.add_paragraph(style='Normal')
    p_titulo_tabela.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_titulo_tabela.paragraph_format.space_before = Pt(6)
    
    run_titulo = p_titulo_tabela.add_run("Tabela 02 - Principais áreas da Secretaria do TJMG. Fonte: Portal TJMG")
    
    run_titulo.bold = False 
    run_titulo.font.name = FONTE
    run_titulo.font.size = Pt(8)
    p_titulo_tabela.paragraph_format.space_after = Pt(12) 

# --- COMPONENTE: TABELA 03 (ESTRUTURA - 1 COLUNA) ---
def adicionar_tabela_estrutura(document, dados):
    """
    Cria a Tabela 03 (Estrutura) com 1 coluna.
    """
    
    COR_HEADER_MAIN = '7F7F7F'    
    COR_HEADER_GROUP = 'D9D9D9'  
    COR_LINHA_ZEBRADA = 'EEEEEE'
    COR_BRANCO_RGB = RGBColor(255, 255, 255)       
    COR_PRETO_RGB = RGBColor(0, 0, 0) 
    TAMANHO_FONTE_PADRAO = Pt(11) 
    FONTE = 'Calibri'
    
    table = document.add_table(rows=0, cols=1) 
    
    try:
        table.style = 'Normal Table'
    except KeyError:
        print("Aviso: Estilo 'Normal Table' não encontrado. Revertendo para 'Table Grid'.")
        table.style = 'Table Grid' 
    
    col_widths = [Cm(17.5)] 
    table.columns[0].width = col_widths[0]

    data_row_index = 0 

    for i, row_data in enumerate(dados):
        
        tipo_linha = row_data[0]
        col1_texto = row_data[1]
        
        row = table.add_row()
        
        tr = row._tr 
        trPr = tr.get_or_add_trPr() 
        if tipo_linha.startswith("HEADER"):
            tblHeader = OxmlElement('w:tblHeader')
            trPr.append(tblHeader)
        else:
            cantSplit = OxmlElement('w:cantSplit')
            cantSplit.set(qn('w:val'), '0') 
            trPr.append(cantSplit)
        
        cell = row.cells[0] 
        cell.text = col1_texto
        
        set_cell_vertical_alignment(cell, 'center')
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT 
        
        p_format = p.paragraph_format
        p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p_format.space_before = Pt(0)
        p_format.space_after = Pt(0)
        
        run = p.runs[0]
        run.font.name = FONTE
        run.font.size = TAMANHO_FONTE_PADRAO
        run.font.color.rgb = COR_PRETO_RGB
        run.bold = False

        if tipo_linha == "HEADER_MAIN":
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), COR_HEADER_MAIN)
            cell._tc.get_or_add_tcPr().append(shading_elm)
            run.font.color.rgb = COR_BRANCO_RGB
            run.bold = True
            
        elif tipo_linha == "HEADER_GROUP_MERGED":
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), COR_HEADER_GROUP)
            cell._tc.get_or_add_tcPr().append(shading_elm)
            run.bold = False
            set_group_top_border(cell) 

        elif tipo_linha == "DATA_MERGED":
            data_row_index += 1
            # Não aplicamos zebrado nesta tabela

    p_titulo_tabela = document.add_paragraph(style='Normal')
    p_titulo_tabela.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_titulo_tabela.paragraph_format.space_before = Pt(6)
    
    run_titulo = p_titulo_tabela.add_run("Tabela 03 - Estruturas para a Prestação Jurisdicional na Segunda Instância. Fonte: Portal TJMG")
    
    run_titulo.bold = False 
    run_titulo.font.name = FONTE
    run_titulo.font.size = Pt(8)
    p_titulo_tabela.paragraph_format.space_after = Pt(25)

# --- COMPONENTE: TABELA 04 (COMARCAS) ---
def adicionar_tabela_comarcas(document, dados):
    """
    Cria a Tabela 04 (Comarcas) com 4 colunas e cabeçalho mesclado.
    """
    
    COR_HEADER_MAIN = '7F7F7F'    
    COR_LINHA_ZEBRADA = 'D9D9D9' # Cor funcional para zebrado
    COR_BRANCO_RGB = RGBColor(255, 255, 255)       
    COR_PRETO_RGB = RGBColor(0, 0, 0) 
    TAMANHO_FONTE_PADRAO = Pt(11) 
    FONTE = 'Calibri'
    
    table = document.add_table(rows=0, cols=4)
    table.space_after = Pt(20) 
    
    try:
        table.style = 'Normal Table'
    except KeyError:
        print("Aviso: Estilo 'Normal Table' não encontrado. Revertendo para 'Table Grid'.")
        table.style = 'Table Grid' 
    
    col_widths = [Cm(4.375), Cm(4.375), Cm(4.375), Cm(4.375)] 
    for i, width in enumerate(col_widths):
        table.columns[i].width = width

    data_row_index = 0 

    for i, row_data in enumerate(dados):
        
        tipo_linha = row_data[0]
        row = table.add_row()
        
        tr = row._tr 
        trPr = tr.get_or_add_trPr() 
        if tipo_linha.startswith("HEADER"):
            tblHeader = OxmlElement('w:tblHeader')
            trPr.append(tblHeader)
        else:
            cantSplit = OxmlElement('w:cantSplit')
            cantSplit.set(qn('w:val'), '0') 
            trPr.append(cantSplit)
        
        cells = row.cells 

        if tipo_linha == "HEADER_MERGE_4":
            col1_texto = row_data[1]
            cell = cells[0].merge(cells[1]).merge(cells[2]).merge(cells[3])
            cell.text = col1_texto
            
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), COR_HEADER_MAIN)
            cell._tc.get_or_add_tcPr().append(shading_elm)
            
            set_cell_vertical_alignment(cell, 'center')
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            p_format = p.paragraph_format
            p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p_format.space_before = Pt(0)
            p_format.space_after = Pt(0)
            
            run = p.runs[0]
            run.font.color.rgb = COR_BRANCO_RGB
            run.bold = True
            run.font.name = FONTE
            run.font.size = TAMANHO_FONTE_PADRAO

        elif tipo_linha == "DATA_4_COL":
            data_row_index += 1
            
            for j in range(4):
                cell = cells[j]
                cell.text = row_data[j+1] 
                
                set_cell_vertical_alignment(cell, 'center')
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT 
                
                p_format = p.paragraph_format
                p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p_format.space_before = Pt(0)
                p_format.space_after = Pt(0)
                
                run = p.runs[0]
                run.font.name = FONTE
                run.font.size = TAMANHO_FONTE_PADRAO
                run.font.color.rgb = COR_PRETO_RGB
                run.bold = False

                if data_row_index % 2 != 0: 
                    shading_elm = OxmlElement('w:shd')
                    shading_elm.set(qn('w:fill'), COR_LINHA_ZEBRADA)
                    cell._tc.get_or_add_tcPr().append(shading_elm)

    p_titulo_tabela = document.add_paragraph(style='Normal')
    p_titulo_tabela.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_titulo_tabela.paragraph_format.space_before = Pt(6)
    
    run_titulo = p_titulo_tabela.add_run("Tabela 04 - Comarcas Instaladas. Fonte: Portal TJMG")
    
    run_titulo.bold = False 
    run_titulo.font.name = FONTE
    run_titulo.font.size = Pt(8)
    p_titulo_tabela.paragraph_format.space_before = Pt(20)
    p_titulo_tabela.paragraph_format.space_after = Pt(30)

# --- NOVO COMPONENTE: TABELA 05 (NÚCLEOS) ---
def adicionar_tabela_nucleos(document, dados):
    """
    Cria a Tabela 05 (Núcleos) com 1 coluna. Começa com o estilo de grupo e 
    as linhas de dados são brancas (sem zebrado).
    """
    
    COR_HEADER_GROUP = 'D9D9D9'  
    COR_PRETO_RGB = RGBColor(0, 0, 0) 
    TAMANHO_FONTE_PADRAO = Pt(11) 
    FONTE = 'Calibri'
    
    table = document.add_table(rows=0, cols=1) 
    
    try:
        table.style = 'Normal Table'
    except KeyError:
        print("Aviso: Estilo 'Normal Table' não encontrado. Revertendo para 'Table Grid'.")
        table.style = 'Table Grid' 
    
    col_widths = [Cm(17.5)] 
    table.columns[0].width = col_widths[0]

    for i, row_data in enumerate(dados):
        
        tipo_linha = row_data[0]
        col1_texto = row_data[1]
        
        row = table.add_row()
        
        tr = row._tr 
        trPr = tr.get_or_add_trPr() 
        if tipo_linha.startswith("HEADER"):
            tblHeader = OxmlElement('w:tblHeader')
            trPr.append(tblHeader)
        else:
            cantSplit = OxmlElement('w:cantSplit')
            cantSplit.set(qn('w:val'), '0') 
            trPr.append(cantSplit)
        
        cell = row.cells[0] 
        cell.text = col1_texto
        
        set_cell_vertical_alignment(cell, 'center')
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT 
        
        p_format = p.paragraph_format
        p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p_format.space_before = Pt(0)
        p_format.space_after = Pt(0)
        
        run = p.runs[0]
        run.font.name = FONTE
        run.font.size = TAMANHO_FONTE_PADRAO
        run.font.color.rgb = COR_PRETO_RGB
        run.bold = False

        # --- TIPO 1: Cabeçalho de Grupo (Primeira Linha) ---
        if tipo_linha == "HEADER_GROUP_MERGED":
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), COR_HEADER_GROUP)
            cell._tc.get_or_add_tcPr().append(shading_elm)
            run.bold = True
            set_group_top_border(cell) 

        # --- TIPO 2: Dados (Linhas seguintes) ---
        elif tipo_linha == "DATA_MERGED":
            # Nenhuma cor de fundo é aplicada (mantendo-se branca/transparente)
            pass

    # --- LEGENDA/FONTE (Tabela 05) ---
    p_titulo_tabela = document.add_paragraph(style='Normal')
    p_titulo_tabela.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_titulo_tabela.paragraph_format.space_before = Pt(6)
    
    run_titulo = p_titulo_tabela.add_run("Tabela 05 - Relação dos Núcleos de Justiça 4.0 da Primeira Instância. Fonte: Infoguia")
    
    run_titulo.bold = False 
    run_titulo.font.name = FONTE
    run_titulo.font.size = Pt(8)
    p_titulo_tabela.paragraph_format.space_after = Pt(30)


# --- NOVO COMPONENTE: TABELA 06 (DADOS HISTÓRICOS) ---
def adicionar_tabela_historicos(document, dados):
    """
    Cria a Tabela 06 (Dados Históricos), aplicando cor de destaque à
    coluna do ano mais recente (2024), margens e altura de linha fixa.
    """
    
    # --- VARIÁVEIS DE COR E ESTILO ---
    COR_HEADER_PRINCIPAL = '7F7F7F'    
    COR_LINHA_ZEBRADA = 'D9D9D9'       
    COR_BRANCO_RGB = RGBColor(255, 255, 255)       
    COR_PRETO_RGB = RGBColor(0, 0, 0) 
    TAMANHO_FONTE_PADRAO = Pt(12) 
    FONTE = 'Calibri'
    NUM_COLUNAS = 7
    
    # NOVAS CORES (Hexadecimais)
    COR_SUB_HEADER_COLUNA = '44546A' 
    COR_DADOS_COLUNA = 'D5DCE4'      
    COLUNA_DESTAQUE_INDEX = 5        # Coluna "2024"
    
    ALTURA_LINHA_TWIPS = 284 
    
    # --- ESTRUTURA E LARGURA ---
    table = document.add_table(rows=0, cols=NUM_COLUNAS) 
    
    try:
        table.style = 'Normal Table' 
    except KeyError:
        print("Aviso: Estilo 'Normal Table' não encontrado. Revertendo para 'Table Grid'.")
        table.style = 'Table Grid' 
    
    # Larguras de Coluna (3.2cm + 6 * 2.133cm = ~16cm área útil)
    col_widths = [Cm(3.2)] + [Cm(2.133)] * 6
    for i, width in enumerate(col_widths):
        table.columns[i].width = width

    data_row_index = 0 

    for i, row_data_full in enumerate(dados):
        
        tipo_linha = row_data_full[0]
        row_data = row_data_full[1:] 
        
        row = table.add_row()
        set_row_height_exact(row, ALTURA_LINHA_TWIPS)
        
        tr = row._tr 
        trPr = tr.get_or_add_trPr() 
        if tipo_linha in ["HEADER_MERGE", "SUB_HEADER"]:
            tblHeader = OxmlElement('w:tblHeader')
            trPr.append(tblHeader)
        else:
            cantSplit = OxmlElement('w:cantSplit')
            cantSplit.set(qn('w:val'), '0') 
            trPr.append(cantSplit)
        
        cells = row.cells 

        # --- TIPO 1: CABEÇALHO PRINCIPAL (Merscla 7 Colunas) ---
        if tipo_linha == "HEADER_MERGE":
            
            cell = cells[0].merge(cells[1]).merge(cells[2]).merge(cells[3]).merge(cells[4]).merge(cells[5]).merge(cells[6])
            cell.text = row_data[0]
            
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), COR_HEADER_PRINCIPAL)
            cell._tc.get_or_add_tcPr().append(shading_elm)
            
            set_cell_vertical_alignment(cell, 'center')
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.runs[0]
            run.font.color.rgb = COR_BRANCO_RGB
            run.bold = True
            run.font.name = FONTE
            run.font.size = Pt(12) 
        
        # --- TIPO 2: SUB-CABEÇALHO (Instância, 2020...) ---
        elif tipo_linha == "SUB_HEADER":
            
            for j in range(NUM_COLUNAS):
                cell = cells[j]
                cell.text = row_data[j]
                
                set_cell_vertical_alignment(cell, 'center')
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.runs[0]
                run.font.color.rgb = COR_PRETO_RGB
                run.bold = True
                run.font.name = FONTE
                run.font.size = Pt(12) 
                
                # --- Destaque: Coluna 2024 ---
                if j == COLUNA_DESTAQUE_INDEX:
                    shading_elm = OxmlElement('w:shd')
                    shading_elm.set(qn('w:fill'), COR_SUB_HEADER_COLUNA)
                    cell._tc.get_or_add_tcPr().append(shading_elm)
                    run.font.color.rgb = COR_BRANCO_RGB 
                
                set_cell_bottom_border(cell)

        # --- TIPO 3 & 4: LINHAS DE DADOS E TOTAL ---
        elif tipo_linha in ["DATA_ROW", "TOTAL_ROW"]:
            
            is_total_row = (tipo_linha == "TOTAL_ROW")
            if not is_total_row:
                data_row_index += 1 

            for j, cell_data in enumerate(row_data):
                cell = cells[j]
                cell.text = cell_data
                
                cell_align = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_vertical_alignment(cell, 'center')
                p = cell.paragraphs[0]
                p.alignment = cell_align
                
                run = p.runs[0]
                run.font.name = FONTE
                run.font.size = TAMANHO_FONTE_PADRAO
                run.font.color.rgb = COR_PRETO_RGB
                run.bold = is_total_row
                
                # --- Lógica de Sombreamento (Prioridade: Coluna > Total/Zebrado) ---
                current_shading_color = None

                if is_total_row or (data_row_index % 2 != 0): 
                    current_shading_color = COR_LINHA_ZEBRADA 
                
                if j == COLUNA_DESTAQUE_INDEX:
                    current_shading_color = COR_DADOS_COLUNA
                
                if current_shading_color:
                    shading_elm = OxmlElement('w:shd')
                    shading_elm.set(qn('w:fill'), current_shading_color)
                    cell._tc.get_or_add_tcPr().append(shading_elm)
                # --- Fim da Lógica de Sombreamento ---

                p_format = p.paragraph_format
                p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p_format.space_before = Pt(0)
                p_format.space_after = Pt(0)
                

    # --- LEGENDA/FONTE (Tabela 06) ---
    p_titulo_tabela = document.add_paragraph(style='Normal')
    p_titulo_tabela.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_titulo_tabela.paragraph_format.space_before = Pt(6)
    
    run_titulo = p_titulo_tabela.add_run("Tabela 06 - Número de processos distribuídos. Fonte: Centro de Informações para a Gestão Institucional – CEINFO")
    
    run_titulo.bold = False 
    run_titulo.font.name = FONTE
    run_titulo.font.size = Pt(8)
    p_titulo_tabela.paragraph_format.space_after = Pt(12)


def aplicar_estilo_capa(paragrafo, texto, tamanho_pt):
    """Aplica o estilo de fonte Bahnschrift com um tamanho específico."""
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragrafo.add_run(texto)
    run.font.name = 'Bahnschrift SemiCondensed' 
    run.font.size = Pt(tamanho_pt)
    run.bold = True                          

# --- 6. EXECUÇÃO E GERAÇÃO DO DOCUMENTO ---

if __name__ == "__main__":
    
    print("--- INICIANDO GERADOR DE RELATÓRIO COMPLETO ---")
    
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    CAMINHO_SUMARIO = os.path.join(SCRIPT_DIR, "Sumario_Modelo.docx")
    CAMINHO_CONTEUDO = os.path.join(SCRIPT_DIR, "Conteudo_Fonte.docx")
    CAMINHO_SAIDA = os.path.join(SCRIPT_DIR, "export/Relatorio_Final_Completo.docx")

    estrutura_final = extrair_sumario_para_json(CAMINHO_SUMARIO, PATTERN_SUMARIO)
    conteudo_mapeado = extrair_conteudo_mapeado(CAMINHO_CONTEUDO, PATTERN_CONTEUDO, PATTERN_LEGENDA)

    if not estrutura_final:
        print("Execução interrompida. 'estrutura_final' (Sumário) está vazia.")
        exit()

    document = Document()

    configurar_margens(document, 3.0, 3.0, 2.0, 2.0) 

    customizar_estilos_titulo(document)
    footer = document.sections[0].footer
    add_page_number(footer)

    # --- INÍCIO DA SEÇÃO: CRIAÇÃO DA CAPA ---
    print("Criando a Capa...")
    
    p_titulo1 = document.add_paragraph()
    aplicar_estilo_capa(p_titulo1, "RELATÓRIO DIAGNÓSTICO DO TRIBUNAL", 20)
    p_titulo1.paragraph_format.space_before = Pt(48)
    p_titulo2 = document.add_paragraph()
    aplicar_estilo_capa(p_titulo2, "DE JUSTIÇA DO ESTADO DE MINAS GERAIS – TJMG", 20)
    p_ano = document.add_paragraph()
    aplicar_estilo_capa(p_ano, "2025", 20)
    p_ano.paragraph_format.space_before = Pt(48)
    p_ano.paragraph_format.space_after = Pt(48)
    p_plan = document.add_paragraph()
    aplicar_estilo_capa(p_plan, "PLANEJAMENTO ESTRATÉGICO 2021-2026", 20)
    p_plan.paragraph_format.space_after = Pt(280)
    p_setor = document.add_paragraph()
    aplicar_estilo_capa(p_setor, 'DEPLAG - TJMG', 14)
    p_data = document.add_paragraph()
    aplicar_estilo_capa(p_data, "JANEIRO DE 2025", 14)

    document.add_page_break()
    # --- FIM DA CAPA ---

    # --- INÍCIO DA SEÇÃO: SUMÁRIO (Simplificado) ---
    document.add_heading('Sumário', level=1) 
    print("Criando sumário estático...")

    for elemento in estrutura_final:
        if elemento['tipo'] == 'TITULO':
            level = elemento['level']
            texto = elemento['texto']
            
            p = document.add_paragraph(style='Normal')
            run = p.add_run(texto)
            run.bold = True
            run.font.name = 'Calibri'
            
            p_format = p.paragraph_format
            p_format.line_spacing = 1.25 
            p_format.space_after = Pt(8) 
            
            if level == 2:
                p_format.left_indent = Inches(0.2)
            elif level == 3:
                p_format.left_indent = Inches(0.4)
            else:
                p_format.left_indent = Inches(0) 

    document.add_page_break()
    # --- FIM DO SUMÁRIO ---

    # --- SEÇÃO: CORPO DO DOCUMENTO (COM PROCESSADOR DE CONTEÚDO) ---
    print("Gerando corpo do relatório com títulos e conteúdo...")

    for elemento in estrutura_final:
        if elemento['tipo'] == 'TITULO':
            
            texto_chave = elemento['texto'] 
            level = elemento['level']
            
            if level == 1:
                texto_para_imprimir = texto_chave.replace(" ", ". ", 1)
            else:
                texto_para_imprimir = texto_chave
                
            # document.add_heading(texto_para_imprimir, level=level)
            p = document.add_heading(texto_para_imprimir, level=level)
            p_format = p.paragraph_format
            p_format.space_after = Pt(20)

            titulo_chave = elemento['texto'] 
            
            if titulo_chave in conteudo_mapeado:
                
                for bloco in conteudo_mapeado[titulo_chave]:
                    
                    # --- PROCESSADOR DE BLOCOS ---
                    
                    if bloco['tipo'] == 'PARAGRAFO':
                        p = document.add_paragraph(bloco['texto'])
                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        p_format = p.paragraph_format
                        p_format.line_spacing = 1.5
                        p_format.space_after = Pt(8) 
                    
                    elif bloco['tipo'] == 'FIGURA':
                        legenda_completa = bloco['legenda_completa']
                        
                        partes = legenda_completa.split("Fonte:")
                        legenda_chave = partes[0].strip() 
                        
                        if len(partes) > 1:
                            texto_fonte = f"Fonte: {partes[1].strip()}"
                        else:
                            texto_fonte = ""

                        print(f"--- Processando FIGURA: {legenda_chave}")
                        
                        if legenda_chave in MAPA_IMAGENS:
                            caminho_imagem_relativo = MAPA_IMAGENS[legenda_chave]
                            caminho_imagem_abs = os.path.join(SCRIPT_DIR, caminho_imagem_relativo)
                            
                            try:
                                document.add_picture(caminho_imagem_abs, width=Cm(16.5)) 
                                document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                
                            except FileNotFoundError:
                                print(f"!!! AVISO: Imagem não encontrada em: {caminho_imagem_abs}")
                                document.add_paragraph(f"[ERRO: IMAGEM NÃO ENCONTRADA: {caminho_imagem_relativo}]")
                            except Exception as e:
                                print(f"!!! ERRO ao inserir imagem: {e}")

                        else:
                            print(f"!!! AVISO: Legenda '{legenda_chave}' não encontrada no MAPA_IMAGENS.")
                            document.add_paragraph(f"[ERRO: MAPEAMENTO DE IMAGEM AUSENTE PARA: {legenda_chave}]")

                        p_legenda = document.add_paragraph()
                        run_legenda = p_legenda.add_run(legenda_chave)
                        run_legenda.font.name = 'Calibri'
                        run_legenda.font.size = Pt(8) 
                        run_legenda.bold = False 
                        p_legenda.alignment = WD_ALIGN_PARAGRAPH.LEFT 
                        p_legenda.paragraph_format.space_before = Pt(6) 
                        p_legenda.paragraph_format.space_after = Pt(0) 

                        if texto_fonte:
                            p_fonte = document.add_paragraph()
                            run_fonte = p_fonte.add_run(texto_fonte)
                            run_fonte.font.name = 'Calibri'
                            run_fonte.font.size = Pt(8) 
                            p_fonte.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            p_fonte.paragraph_format.space_after = Pt(12)

                    elif bloco['tipo'] == 'TABELA_ATOS':
                        print(f"--- Inserindo Tabela 01 (Atos) para {titulo_chave} ---")
                        adicionar_tabela_atos(document, bloco['dados'])
                    
                    elif bloco['tipo'] == 'TABELA_AREAS':
                        print(f"--- Inserindo Tabela 02 (Áreas) para {titulo_chave} ---")
                        adicionar_tabela_areas(document, bloco['dados'])
                    
                    elif bloco['tipo'] == 'TABELA_ESTRUTURA':
                        print(f"--- Inserindo Tabela 03 (Estrutura) para {titulo_chave} ---")
                        adicionar_tabela_estrutura(document, bloco['dados'])

                    elif bloco['tipo'] == 'TABELA_COMARCAS':
                        print(f"--- Inserindo Tabela 04 (Comarcas) para {titulo_chave} ---")
                        adicionar_tabela_comarcas(document, bloco['dados'])
                    
                    elif bloco['tipo'] == 'TABELA_NUCLEOS':
                        print(f"--- Inserindo Tabela 05 (Núcleos) para {titulo_chave} ---")
                        adicionar_tabela_nucleos(document, bloco['dados'])

                    elif bloco['tipo'] == 'TABELA_HISTORICOS':
                        print(f"--- Inserindo Tabela 06 (Dados Históricos) para {titulo_chave} ---")
                        adicionar_tabela_historicos(document, bloco['dados'])

    # --- FIM DA SEÇÃO ---

    # --- SALVAR O DOCUMENTO ---
    try:
        document.save(CAMINHO_SAIDA)
        print(f"Documento '{CAMINHO_SAIDA}' gerado com sucesso!")
    except PermissionError:
        print(f"!!! ERRO DE PERMISSÃO: Não foi possível salvar o arquivo '{CAMINHO_SAIDA}'.")
        print("!!! Verifique se o arquivo não está aberto no Word.")
    except Exception as e:
        print(f"!!! ERRO INESPERADO AO SALVAR: {e}")