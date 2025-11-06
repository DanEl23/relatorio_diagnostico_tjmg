import json
import re 
import os 
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

# --- 1. DEFINIÇÃO DAS EXPRESSÕES REGULARES ---

PATTERN_SUMARIO = r'^\s*(\d+(?:\.\d+)*)\.?\s+(.*?)\s*[\. ]*\d+$' 
PATTERN_CONTEUDO = r'^\s*(\d+(\.\d+)*)\.?\s*(.*)$'
PATTERN_LEGENDA = r'^(Figura|Gráfico)\s+\d+' 

# --- 2. DADOS BRUTOS (HARDCODED) ---

dados_tabela_atos = [
    # Cabeçalho
    ("ATO NORMATIVO", "ESTRUTURA"),
    ("Lei Complementar nº 59/2001", "Contém a organização e a divisão judiciárias do Estado de Minas Gerais."),
    ("Resolução do Tribunal Pleno nº 03/2012", "Contém o Regimento Interno do Tribunal de Justiça."),
    ("Resolução nº 518/2007", "Dispõe sobre os níveis hierárquicos e as atribuições gerais das unidades organizacionais que integram a Secretaria do Tribunal de Justiça do Estado de Minas Gerais."),
    ("Resolução nº 522/2007", 
     "Dispõe sobre a Superintendência Administrativa:\n"
     "ü Superintendente Administrativo;\n"
     "ü Diretoria Executiva da Gestão de Bens, Serviços e Patrimônio;\n"
     "ü Diretoria Executiva de Engenharia e Gestão Predial;\n"
     "ü Diretoria Executiva de Informática."),
    ("Resolução nº 557/2008", "Dispõe sobre a criação da Comissão Estadual Judiciária de Adoção, CEJA-MG."),
    ("Resolução nº 640/2010", "Cria a Coordenadoria da Infância e da Juventude."),
    ("Resolução nº 673/2011", "Cria a Coordenadoria da Mulher em Situação de Violência Doméstica e Familiar."),
    ("Resolução nº 821/2016", "Dispõe sobre a reestruturação da Corregedoria Geral de Justiça."),
    ("Resolução nº 862/2017", "Dispõe sobre a estrutura e o funcionamento da Ouvidoria do Tribunal de Justiça do Estado de Minas Gerais."),
    ("Resolução nº 873/2018", "Dispõe sobre a estrutura e o funcionamento do Núcleo Permanente de Métodos de Solução de Conflitos, da Superintendência da Gestão de Inovação e do órgão jurisdicional da Secretaria do Tribunal de Justiça diretamente vinculado à Terceira Vice-Presidência, e estabelece normas para a instalação dos Centros Judiciários de Solução de Conflitos e Cidadania."),
    ("Resolução nº 877/2018", "Instala, \"ad referendum\" do Órgão Especial, a 19ª Câmara Cível no Tribunal de Justiça."),
    ("Resolução n° 878/2018", "Referenda a instalação da Câmara de que trata o art. 7º da Lei Complementar estadual nº 146, de 9 de janeiro de 2018, promovida pela Resolução nº 877, de 29 de junho de 2018."),
    ("Resolução nº 886/2019", "Determina a instalação da 8ª Câmara Criminal no Tribunal de Justiça."),
    ("Resolução nº 893/2019", "Determina a instalação da 20ª Câmara Cível no Tribunal de Justiça."),
    ("Resolução n° 969/2021", 
     "Dispõe sobre os Comitês de Assessoramento à Presidência, estabelece a estrutura e o funcionamento das unidades organizacionais da Secretaria do Tribunal de Justiça diretamente vinculadas ou subordinadas à Presidência:\n"
     "ü Comitê de Governança e Gestão Estratégica;\n"
     "ü Comitê Executivo de Gestão Institucional;\n"
     "ü Comitê Institucional de Inteligência;\n"
     "ü Comitê de Monitoramento e Suporte à Prestação Jurisdicional;\n"
     "ü Comitê de Tecnologia da Informação e Comunicação;\n"
     "ü Comitê Gestor de Segurança da Informação;\n"
     "ü Comitê Gestor da Política Judiciária para a Primeira Infância; (Alínea acrescentada pela Resolução do Órgão Especial nº 1052/2023).\n"
     "ü Comitê Gestor Regional de Primeira Instância. (Alínea acrescentada pela Resolução do Órgão Especial nº 1063/2023).\n"
     "ü Secretaria de Governança e Gestão Estratégica;\n"
     "ü Diretoria Executiva de Comunicação;\n"
     "ü Gabinete de Segurança Institucional;\n"
     "ü Diretoria Executiva de Planejamento Orçamentário e Qualidade na Gestão Institucional;\n"
     "ü Gerência de Suporte aos Juizados Especiais;\n"
     "ü Secretaria do Órgão Especial;\n"
     "ü Assessoria de Precatórios;\n"
     "ü Secretaria de Auditoria Interna;\n"
     "ü Memória do Judiciário."),
    ("Resolução nº 971/2021", "Institui o Programa de Justiça Restaurativa e dispõe sobre a estrutura e funcionamento do Comitê de Justiça Restaurativa - COMJUR e da Central de Apoio à Justiça Restaurativa – CEAJUR."),
    ("Resolução nº 977/2021", "Determina a instalação da Vigésima Primeira Câmara Cível e da Nona Câmara Criminal, a especialização de Câmaras no Tribunal de Justiça."),
    ("Resolução nº 979/2021", "Dispõe sobre a estrutura organizacional e o regulamento da Escola Judicial Desembargador Edésio Fernandes - EJEF."),
    ("Resolução nº 1053/2023", "Dispõe sobre a Superintendência Judiciária."),
    ("Resolução nº 1062/2023", "Altera a Resolução do Órgão Especial nº 979, de 17 de novembro de 2021, que \"Dispõe sobre a estrutura organizacional e o regulamento da Escola Judicial Desembargador Edésio Fernandes - EJEF\"."),
    ("Resolução nº 1063/2023", "Dispõe sobre a organização e o funcionamento do Comitê Gestor Regional de Primeira Instância no âmbito do Poder Judiciário do Estado de Minas Gerais."),
    ("Resolução nº 1066/2023", "Dispõe sobre a estrutura e o funcionamento do Grupo de Monitoramento e Fiscalização do Sistema Carcerário e Socioeducativo - GMF no âmbito do Tribunal de Justiça do Estado de Minas Gerais."),
    ("Resolução nº 1079/2024", "Altera a Resolução do Órgão Especial nº 979, de 17 de novembro de 2021, que \"Dispõe sobre a estrutura organizacional e o regulamento da Escola Judicial Desembargador Edésio Fernandes - EJEF."),
    ("Resolução nº 1080/2024", "Institui o Regulamento da Escola Judicial Desembargador Edésio Fernandes - EJEF."),
    ("Resolução nº 1086/2024", "Altera a Resolução do Órgão Especial nº 1.010, de 29 de agosto de 2020, que \"Dispõe sobre a implementação, a estrutura e o funcionamento dos \"Núcleos de Justiça 4.0\" e dá outras providências\", e altera a Resolução do Órgão Especial nº 1.053, de 20 de setembro de 2023, que \"Dispõe sobre a Superintendência Judiciária e dá outras providências\".")
]

# --- 3. FUNÇÕES AUXILIARES (PAGINAÇÃO, ALINHAMENTO, XML) ---

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
    """ Remove o espaçamento extra de um parágrafo de lista usando XML. """
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

# --- 4. FUNÇÕES DE PROCESSAMENTO E CRIAÇÃO ---

def extrair_sumario_para_json(caminho_arquivo_docx, pattern_regex):
    """Lê o DOCX modelo (Sumário) e extrai títulos numerados."""
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
    Lê o documento DOCX de conteúdo, mapeia parágrafos e identifica marcadores.
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

        # 1. É um TÍTULO?
        match_titulo = re.search(pattern_titulo, texto, re.IGNORECASE)
        
        if match_titulo:
            prefixo = match_titulo.group(1).strip()
            titulo = match_titulo.group(3).strip()
            chave_titulo_atual = f"{prefixo} {titulo}"
            conteudo_mapeado[chave_titulo_atual] = []
            
        elif chave_titulo_atual:
            # 2. NÃO é um TÍTULO. É um conteúdo.
            
            if texto == "[INSERIR_TABELA_ATOS_NORMATIVOS]":
                # 2A. É O GATILHO DA TABELA
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "TABELA_ATOS",
                    "dados": dados_tabela_atos 
                })
                
            elif re.search(pattern_legenda, texto, re.IGNORECASE):
                # 2B. É UMA FIGURA/GRÁFICO
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "FIGURA",
                    "legenda": texto,
                    "caminho_imagem": "placeholders/figura_xx.png" 
                })
            else:
                # 2C. É um PARÁGRAFO comum.
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "PARAGRAFO",
                    "texto": texto
                })

    print(f"Extração de conteúdo concluída. {len(conteudo_mapeado)} títulos mapeados.")
    return conteudo_mapeado


def customizar_estilos_titulo(documento):
    """Aplica formatação personalizada nos estilos de Título e no 'Normal'."""
    
    # --- Estilos H1, H2, H3 (com Fonte Calibri) ---
    style_h1 = documento.styles['Heading 1']
    font_h1 = style_h1.font
    font_h1.name = 'Calibri' 
    font_h1.size = Pt(20) 
    font_h1.color.rgb = RGBColor(162, 22, 18) 
    font_h1.all_caps = True 
    font_h1.bold = True
    
    style_h2 = documento.styles['Heading 2']
    font_h2 = style_h2.font
    font_h2.name = 'Calibri' 
    font_h2.size = Pt(17)
    font_h2.color.rgb = RGBColor(162, 22, 18) 
    font_h2.bold = True
    font_h2.all_caps = False
    
    style_h3 = documento.styles['Heading 3']
    font_h3 = style_h3.font
    font_h3.name = 'Calibri' 
    font_h3.size = Pt(15.5)
    font_h3.color.rgb = RGBColor(162, 22, 18) 
    font_h3.bold = True

    # --- Estilo "Normal" (Corpo do Texto) ---
    style_normal = documento.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Calibri'
    font_normal.size = Pt(12) 

def adicionar_tabela_atos(document, dados):
    """
    Cria e estiliza a Tabela de Atos Normativos com formatação específica,
    incluindo repetição de cabeçalho e permissão de quebra de linha.
    """
    
    # --- VARIÁVEIS DE COR E ESTILO ---
    COR_CABECALHO_RGB = RGBColor(127, 127, 127)   
    COR_CABECALHO_HEX = '7F7F7F'                  
    COR_BRANCO_RGB = RGBColor(255, 255, 255)       
    COR_PRETO_RGB = RGBColor(0, 0, 0) 
    COR_CINZA_CLARO_HEX = 'EEEEEE'                 

    TAMANHO_FONTE_PADRAO = Pt(12) 
    FONTE = 'Calibri'
    
    table = document.add_table(rows=1, cols=len(dados[0]))
    table.style = 'Table Grid'
    
    # Definir Larguras de Coluna (Cm) (Valores Corrigidos)
    col_widths = [Cm(5), Cm(16)]
    for i, width in enumerate(col_widths):
        table.columns[i].width = width

    # Processar Linhas e Estilização
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

    # --- LEGENDA/FONTE (CORRIGIDA) ---
    p_titulo_tabela = document.add_paragraph(style='Normal')
    p_titulo_tabela.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_titulo_tabela.paragraph_format.space_before = Pt(6)
    
    run_titulo = p_titulo_tabela.add_run("Tabela 01 - Atos Normativos referentes à Estrutura do TJMG. Fonte: Portal TJMG")
    
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

# --- 5. EXECUÇÃO E GERAÇÃO DO DOCUMENTO ---

if __name__ == "__main__":
    
    print("--- INICIANDO GERADOR DE RELATÓRIO COMPLETO ---")
    
    # Define o caminho absoluto para os arquivos
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    CAMINHO_SUMARIO = os.path.join(SCRIPT_DIR, "Sumario_Modelo.docx")
    CAMINHO_CONTEUDO = os.path.join(SCRIPT_DIR, "Conteudo_Fonte.docx")
    CAMINHO_SAIDA = os.path.join(SCRIPT_DIR, "export/Relatorio_Final_Completo.docx")

    # 1. Extrai o roteiro (O Esqueleto)
    estrutura_final = extrair_sumario_para_json(CAMINHO_SUMARIO, PATTERN_SUMARIO)

    # 2. Extrai o conteúdo (O Recheio)
    conteudo_mapeado = extrair_conteudo_mapeado(CAMINHO_CONTEUDO, PATTERN_CONTEUDO, PATTERN_LEGENDA)

    if not estrutura_final:
        print("Execução interrompida. 'estrutura_final' (Sumário) está vazia.")
        exit()

    # 3. Cria o novo documento
    document = Document() 

    # 4. Aplica Estilos e Paginação
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
    document.add_heading('Sumário \n', level=1) 
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
            document.add_heading(elemento['texto']+'\n', level=elemento['level'])
            
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
                        print(f"--- Encontrado marcador de FIGURA: {bloco['legenda']}")
                        document.add_paragraph(f"[PLACEHOLDER DE IMAGEM AQUI: {bloco['caminho_imagem']}]")
                        document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        p_legenda = document.add_paragraph(bloco['legenda'])
                        p_legenda.alignment = WD_ALIGN_PARAGRAPH.LEFT 
                        p_legenda.paragraph_format.space_before = Pt(6) 
                        p_legenda.paragraph_format.space_after = Pt(12) 
                        run_legenda = p_legenda.runs[0]
                        run_legenda.font.name = 'Calibri'
                        run_legenda.font.size = Pt(8)

                    elif bloco['tipo'] == 'TABELA_ATOS':
                        print(f"--- Inserindo Tabela de Atos Normativos (Estilizada) para {titulo_chave} ---")
                        adicionar_tabela_atos(document, bloco['dados'])

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