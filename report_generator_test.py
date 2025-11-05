import json
import re 
import os # Para lidar com os caminhos de arquivo
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

# --- 1. DEFINIÇÃO DAS EXPRESSÕES REGULARES ---

# PATTERN_SUMARIO: Para o 'Sumario_Modelo.docx' (espera números de página no final)
PATTERN_SUMARIO = r'^\s*(\d+(?:\.\d+)*)\.?\s+(.*?)\s*[\. ]*\d+$' 

# PATTERN_CONTEUDO: Para o 'Conteudo_Fonte.docx' (robusta, não espera números de página)
PATTERN_CONTEUDO = r'^\s*(\d+(\.\d+)*)\.?\s*(.*)$'

# PATTERN_LEGENDA: Identifica legendas que começam com "Figura" ou "Gráfico".
PATTERN_LEGENDA = r'^(Figura|Gráfico|Tabela)\s+\d+'


# --- FUNÇÃO DE PAGINAÇÃO (Rodapé) ---
def add_page_number(footer):
    """
    Adiciona o código de campo de paginação (PAGE) a um rodapé (footer).
    """
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

# --- FUNÇÃO DE EXTRAÇÃO DO SUMÁRIO ---
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
        
        # Usa a PATTERN_SUMARIO
        match = re.search(pattern_regex, texto_limpo, re.IGNORECASE)
        
        if match:
            prefixo_completo = match.group(1).strip()
            texto_titulo = match.group(2).strip() # Grupo 2 para esta pattern
            level = len(prefixo_completo.split('.'))
            
            # Chave: "1. INTRODUÇÃO"
            texto_final_com_numero = f"{prefixo_completo} {texto_titulo}"

            if level >= 1:
                estrutura_do_relatorio.append({
                    "tipo": "TITULO",
                    "level": level,
                    "texto": texto_final_com_numero
                })

    return estrutura_do_relatorio


# --- FUNÇÃO DE EXTRAÇÃO DE CONTEÚDO (COM DETECÇÃO DE FIGURAS) ---
def extrair_conteudo_mapeado(caminho_arquivo_docx, pattern_titulo, pattern_legenda):
    """
    Lê o documento DOCX de conteúdo, mapeia parágrafos e identifica legendas de figuras.
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
            # SIM, é um título. (Lógica de captura validada)
            prefixo = match_titulo.group(1).strip()
            titulo = match_titulo.group(3).strip()
            chave_titulo_atual = f"{prefixo} {titulo}"
            conteudo_mapeado[chave_titulo_atual] = []
            
        elif chave_titulo_atual:
            # 2. NÃO é um TÍTULO. É um conteúdo.
            
            match_legenda = re.search(pattern_legenda, texto, re.IGNORECASE)
            
            if match_legenda:
                # 2A. É uma FIGURA/GRÁFICO (Capturado do Conteudo_Fonte)
                # O CAMINHO_IMAGEM aqui é um PLACEHOLDER! Ele não existe ainda,
                # mas é o que usaremos para indicar que é preciso inserir uma imagem.
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "FIGURA",
                    "legenda": texto,
                    "caminho_imagem": "placeholders/figura_01.png" # <--- CONVENÇÃO FUTURA
                })
            else:
                # 2B. É um PARÁGRAFO comum.
                conteudo_mapeado[chave_titulo_atual].append({
                    "tipo": "PARAGRAFO",
                    "texto": texto
                })

    print(f"Extração de conteúdo concluída. {len(conteudo_mapeado)} títulos mapeados.")
    return conteudo_mapeado


# --- FUNÇÃO DE CUSTOMIZAÇÃO DE ESTILOS ---
def customizar_estilos_titulo(documento):
    """Aplica formatação personalizada nos estilos de Título (H1, H2, H3)"""
    
    # --- Customização H1 (Heading 1) ---
    style_h1 = documento.styles['Heading 1']
    font_h1 = style_h1.font
    font_h1.name = 'Calibri'
    font_h1.size = Pt(20) 
    font_h1.color.rgb = RGBColor(162, 22, 18) # Vermelho Escuro
    font_h1.all_caps = True 
    font_h1.bold = True
    
    # --- Customização H2 (Heading 2) ---
    style_h2 = documento.styles['Heading 2']
    font_h2 = style_h2.font
    font_h2.name = 'Calibri'
    font_h2.size = Pt(17)
    font_h2.color.rgb = RGBColor(162, 22, 18) # Vermelho Escuro
    font_h2.bold = True
    font_h2.all_caps = False
    
    # --- Customização H3 (Heading 3) ---
    style_h3 = documento.styles['Heading 3']
    font_h3 = style_h3.font
    font_h3.name = 'Calibri'
    font_h3.size = Pt(15.5)
    font_h3.color.rgb = RGBColor(162, 22, 18) # Vermelho Escuro
    font_h3.bold = True

    style_normal = documento.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Calibri'
    font_normal.size = Pt(12)

# --- 3. EXECUÇÃO E GERAÇÃO DO DOCUMENTO ---

# Define o caminho absoluto para os arquivos
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_SUMARIO = os.path.join(SCRIPT_DIR, "Sumario_Modelo.docx")
CAMINHO_CONTEUDO = os.path.join(SCRIPT_DIR, "Conteudo_Fonte.docx")

# 3.1. Extrai o roteiro (O Esqueleto)
estrutura_final = extrair_sumario_para_json(CAMINHO_SUMARIO, PATTERN_SUMARIO)

# 3.2. Extrai o conteúdo (O Recheio)
conteudo_mapeado = extrair_conteudo_mapeado(CAMINHO_CONTEUDO, PATTERN_CONTEUDO, PATTERN_LEGENDA)

# Se a estrutura estiver vazia, para aqui
if not estrutura_final:
    print("Execução interrompida. 'estrutura_final' (Sumário) está vazia.")
    exit()

# 3.3. Cria o novo documento
document = Document() 

# 3.4. Aplica Estilos e Paginação
customizar_estilos_titulo(document)
footer = document.sections[0].footer
add_page_number(footer)

# --- INÍCIO DA SEÇÃO: CRIAÇÃO DA CAPA ---
def aplicar_estilo_capa(paragrafo, texto, tamanho_pt):
    """Aplica o estilo de fonte Bahnschrift com um tamanho específico."""
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragrafo.add_run(texto)
    run.font.name = 'Bahnschrift SemiCondensed' 
    run.font.size = Pt(tamanho_pt)
    run.bold = True                          

print("Criando a Capa (versão TJMG)...")

# (Logo comentado)
# try:
#    CAMINHO_LOGO = os.path.join(SCRIPT_DIR, "resources/capa_relatorio.png")
#    document.add_picture(CAMINHO_LOGO, width=Cm(5.0))
#    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
# except FileNotFoundError:
#    print(f"AVISO: Arquivo '{CAMINHO_LOGO}' não encontrado. Pulando o logo.")
# except Exception as e:
#    print(f"ERRO ao adicionar a imagem: {e}")

# Textos da Capa
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

print("Criando sumário estático (simplificado, sem pontos)...")
for elemento in estrutura_final:
    if elemento['tipo'] == 'TITULO':
        level = elemento['level']
        texto = elemento['texto']
        
        p = document.add_paragraph(style='Normal')
        run = p.add_run(texto)
        run.bold = True

        run.font.name = 'Calibri'
        
        p_format = p.paragraph_format
        p_format.line_spacing = 1.5 
        p_format.space_after = Pt(6) 
        
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
    # 1. Sempre adiciona o Título (H1, H2, H3)
    if elemento['tipo'] == 'TITULO':
        document.add_heading(elemento['texto'], level=elemento['level'])
        
        titulo_chave = elemento['texto']
        if titulo_chave in conteudo_mapeado:
            
            for bloco in conteudo_mapeado[titulo_chave]:
                
                # --- PROCESSADOR DE BLOCOS (COM FIGURA) ---
                
                if bloco['tipo'] == 'PARAGRAFO':
                    # Aplica a formatação de corpo de texto
                    p = document.add_paragraph(bloco['texto'])
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p_format = p.paragraph_format
                    p_format.line_spacing = 1.5
                    p_format.space_after = Pt(8) 
                
                elif bloco['tipo'] == 'FIGURA':
                    print(f"--- Encontrado marcador de FIGURA: {bloco['legenda']}")
                    #
                    # LÓGICA DE INSERÇÃO DE IMAGEM (PRÓXIMO PASSO DE DETALHE)
                    # Por agora, apenas adicionamos a legenda formatada.
                    #
                    
                    # 1. Insere a imagem (Por enquanto, um marcador de texto)
                    document.add_paragraph(f"[PLACEHOLDER DE IMAGEM AQUI: {bloco['caminho_imagem']}]")
                    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # 2. Insere a Legenda
                    p_legenda = document.add_paragraph(bloco['legenda'])
                    p_legenda.alignment = WD_ALIGN_PARAGRAPH.CENTER # Legendas costumam ser centralizadas
                    p_legenda.paragraph_format.space_before = Pt(6) # Espaço após a imagem
                    p_legenda.paragraph_format.space_after = Pt(12) # Espaço antes do próximo texto
                    
                    # 3. Aplica o estilo de fonte Calibri (tamanho padrão será 12, pois é Normal)
                    run_legenda = p_legenda.runs[0]
                    run_legenda.font.name = 'Calibri'
                    run_legenda.font.size = Pt(8)
                
# --- FIM DA SEÇÃO ---

# --- SALVAR O DOCUMENTO ---
try:
    # (Use r-strings ou barras normais para o caminho de saída)
    CAMINHO_SAIDA = os.path.join(SCRIPT_DIR, "export/Relatorio_Final_Completo.docx")
    document.save(CAMINHO_SAIDA)
    print(f"Documento '{CAMINHO_SAIDA}' gerado com sucesso!")
except PermissionError:
    print(f"!!! ERRO DE PERMISSÃO: Não foi possível salvar o arquivo '{CAMINHO_SAIDA}'.")
    print("!!! Verifique se o arquivo não está aberto no Word.")
except Exception as e:
    print(f"!!! ERRO INESPERADO AO SALVAR: {e}")