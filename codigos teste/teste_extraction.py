import re
import os
import json
from docx import Document

# --- 1. DEPENDÊNCIAS ---

# Nova PATTERN ROBUSTA: Usa \s* (zero ou mais espaços)
PATTERN_CONTEUDO = r'^\s*(\d+(\.\d+)*)\.?\s*(.*)$'

# --- 2. A FUNÇÃO QUE QUEREMOS TESTAR ---

def extrair_conteudo_mapeado(caminho_arquivo_docx, pattern_regex):
    """
    Lê um documento DOCX de conteúdo e mapeia parágrafos para os títulos.
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

        # 1. Este parágrafo é um TÍTULO?
        match = re.search(pattern_regex, texto, re.IGNORECASE)
        
        if match:
            # SIM, é um título.
            
            # --- CORREÇÃO IMPORTANTE NA LÓGICA DA CHAVE ---
            # Group 1: O prefixo numérico (ex: "1" ou "3.2.1")
            prefixo = match.group(1).strip()
            
            # Group 3: O texto do título (ex: "INTRODUÇÃO")
            # (Group 2 é usado internamente pelo regex para os subníveis)
            titulo = match.group(3).strip()
            
            # Define a nova chave (ex: "1 INTRODUÇÃO")
            # Esta lógica de chave DEVE ser idêntica à da função extrair_sumario_para_json
            chave_titulo_atual = f"{prefixo} {titulo}"
            
            conteudo_mapeado[chave_titulo_atual] = []
            
        elif chave_titulo_atual:
            # NÃO, não é um título.
            conteudo_mapeado[chave_titulo_atual].append({
                "tipo": "PARAGRAFO",
                "texto": texto
            })
        
    print(f"Extração concluída. {len(conteudo_mapeado)} títulos mapeados.")
    return conteudo_mapeado

# --- 3. BLOCO EXECUTOR DO TESTE ---

if __name__ == "__main__":
    
    print("--- INICIANDO TESTE DA FUNÇÃO 'extrair_conteudo_mapeado' (v_robusta) ---")
    
    try:
        SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
        CAMINHO_CONTEUDO = os.path.join(SCRIPT_DIR, "Conteudo_Fonte.docx")
        
        print(f"Procurando arquivo de conteúdo em: {CAMINHO_CONTEUDO}")

        # Chama a função com a NOVA PATTERN ROBUSTA
        conteudo_mapeado_resultado = extrair_conteudo_mapeado(CAMINHO_CONTEUDO, PATTERN_CONTEUDO)
        
        if conteudo_mapeado_resultado:
            print("\n--- RESULTADO (DICIONÁRIO MAPEADO) ---")
            print(json.dumps(conteudo_mapeado_resultado, indent=4, ensure_ascii=False))
        else:
            print("\n--- RESULTADO ---")
            print("A função retornou um dicionário vazio.")

    except Exception as e:
        print(f"\n--- ERRO INESPERADO NO TESTE ---")
        print(e)
        
    print("\n--- TESTE CONCLUÍDO ---")