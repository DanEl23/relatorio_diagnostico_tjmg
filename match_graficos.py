"""
Script para fazer a correspondência entre os gráficos do relatório (Contudo_Fonte)
e os arquivos de imagem extraídos do PDF Justiça em Números.

Utiliza o arquivo dicionario_graficos.json que mapeia:
- Chave: Nome do gráfico no relatório (ex: "Gráfico 1")
- Valor: Nome do gráfico no PDF original (ex: "Gráfico 78")
"""

import os
import json
import re
from difflib import SequenceMatcher

# =====================================================================
#                        CONFIGURAÇÕES
# =====================================================================
JSON_PATH = "dicionario_graficos.json"
IMAGES_DIR = "graficos_extraidos_por_titulo"
OUTPUT_JSON = "mapeamento_graficos_completo.json"

# =====================================================================
#                        FUNÇÕES DE SIMILARIDADE
# =====================================================================

def extrair_numero_grafico(texto):
    """
    Extrai o número do gráfico de um texto.
    Ex: "Gráfico 78" → 78
    Ex: "Gráfico 78 - Descrição..." → 78
    """
    match = re.search(r'Gráfico\s+(\d+)', texto, re.IGNORECASE)
    if match:
        return int(match.group(1))
    return None

def calcular_similaridade(str1, str2):
    """
    Calcula a similaridade entre duas strings usando SequenceMatcher.
    Retorna um valor entre 0.0 (totalmente diferente) e 1.0 (idêntico).
    """
    return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()

def encontrar_arquivo_por_numero(numero_grafico, arquivos_disponiveis):
    """
    Encontra o arquivo de imagem correspondente ao número do gráfico.
    
    Args:
        numero_grafico (int): Número do gráfico procurado
        arquivos_disponiveis (list): Lista de nomes de arquivos disponíveis
    
    Returns:
        str or None: Nome do arquivo encontrado ou None se não encontrado
    """
    # Padrão: "Gráfico {numero} - Descrição_PgXXX_FINAL.png"
    padrao = rf'^Gráfico\s+{numero_grafico}\s+-\s+'
    
    for arquivo in arquivos_disponiveis:
        if re.match(padrao, arquivo, re.IGNORECASE):
            return arquivo
    
    return None

def encontrar_melhor_correspondencia(numero_grafico, arquivos_disponiveis, threshold=0.8):
    """
    Encontra a melhor correspondência usando similaridade de strings.
    Útil quando a busca exata falha.
    
    Args:
        numero_grafico (int): Número do gráfico procurado
        arquivos_disponiveis (list): Lista de nomes de arquivos disponíveis
        threshold (float): Limite mínimo de similaridade (0.0 a 1.0)
    
    Returns:
        tuple: (nome_arquivo, score_similaridade) ou (None, 0.0)
    """
    busca = f"Gráfico {numero_grafico}"
    melhor_arquivo = None
    melhor_score = 0.0
    
    for arquivo in arquivos_disponiveis:
        # Extrair apenas o início do nome do arquivo (até a descrição)
        inicio_arquivo = arquivo.split('_Pg')[0] if '_Pg' in arquivo else arquivo
        score = calcular_similaridade(busca, inicio_arquivo)
        
        if score > melhor_score and score >= threshold:
            melhor_score = score
            melhor_arquivo = arquivo
    
    return melhor_arquivo, melhor_score

# =====================================================================
#                        FUNÇÃO PRINCIPAL
# =====================================================================

def mapear_graficos(json_path, images_dir, output_json):
    """
    Cria um mapeamento completo entre os gráficos do relatório e os arquivos extraídos.
    
    Estrutura do JSON de saída:
    {
        "Gráfico 1": {
            "grafico_original": "Gráfico 78",
            "arquivo_encontrado": "Gráfico 78 - Taxa de congestionamento..._Pg162_FINAL.png",
            "caminho_completo": "graficos_extraidos_por_titulo/Gráfico 78 - ...",
            "status": "encontrado" | "nao_encontrado" | "numero_invalido"
        },
        ...
    }
    """
    # Carregar dicionário JSON
    if not os.path.exists(json_path):
        print(f"ERRO: Arquivo JSON não encontrado: {json_path}")
        return
    
    with open(json_path, 'r', encoding='utf-8') as f:
        dicionario = json.load(f)
    
    # Listar arquivos disponíveis no diretório
    if not os.path.exists(images_dir):
        print(f"ERRO: Diretório de imagens não encontrado: {images_dir}")
        return
    
    arquivos_disponiveis = [
        f for f in os.listdir(images_dir) 
        if f.lower().endswith('.png')
    ]
    
    print(f"Total de gráficos no dicionário: {len(dicionario)}")
    print(f"Total de arquivos de imagem disponíveis: {len(arquivos_disponiveis)}")
    print("-" * 70)
    
    # Criar mapeamento
    mapeamento = {}
    estatisticas = {
        "encontrados": 0,
        "nao_encontrados": 0,
        "numeros_invalidos": 0
    }
    
    for grafico_relatorio, grafico_original in dicionario.items():
        resultado = {
            "grafico_original": grafico_original,
            "arquivo_encontrado": None,
            "caminho_completo": None,
            "status": None,
            "similaridade": None
        }
        
        # Extrair número do gráfico original
        numero = extrair_numero_grafico(grafico_original)
        
        if numero is None:
            resultado["status"] = "numero_invalido"
            estatisticas["numeros_invalidos"] += 1
            print(f"⚠️  {grafico_relatorio} → {grafico_original} (número inválido)")
        else:
            # Tentar encontrar arquivo por correspondência exata
            arquivo = encontrar_arquivo_por_numero(numero, arquivos_disponiveis)
            
            if arquivo:
                resultado["arquivo_encontrado"] = arquivo
                resultado["caminho_completo"] = os.path.join(images_dir, arquivo)
                resultado["status"] = "encontrado"
                resultado["similaridade"] = 1.0
                estatisticas["encontrados"] += 1
                print(f"✅ {grafico_relatorio} → {grafico_original} → {arquivo}")
            else:
                # Tentar busca por similaridade
                arquivo_similar, score = encontrar_melhor_correspondencia(
                    numero, arquivos_disponiveis, threshold=0.6
                )
                
                if arquivo_similar:
                    resultado["arquivo_encontrado"] = arquivo_similar
                    resultado["caminho_completo"] = os.path.join(images_dir, arquivo_similar)
                    resultado["status"] = "encontrado_similar"
                    resultado["similaridade"] = round(score, 2)
                    estatisticas["encontrados"] += 1
                    print(f"⚠️  {grafico_relatorio} → {grafico_original} → {arquivo_similar} (similaridade: {score:.2f})")
                else:
                    resultado["status"] = "nao_encontrado"
                    estatisticas["nao_encontrados"] += 1
                    print(f"❌ {grafico_relatorio} → {grafico_original} (arquivo não encontrado)")
        
        mapeamento[grafico_relatorio] = resultado
    
    # Salvar resultado em JSON
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(mapeamento, f, ensure_ascii=False, indent=4)
    
    # Exibir estatísticas
    print("-" * 70)
    print("RESUMO:")
    print(f"  ✅ Encontrados: {estatisticas['encontrados']}")
    print(f"  ❌ Não encontrados: {estatisticas['nao_encontrados']}")
    print(f"  ⚠️  Números inválidos: {estatisticas['numeros_invalidos']}")
    print(f"\nArquivo de mapeamento salvo em: {output_json}")
    
    return mapeamento

# =====================================================================
#                        FUNÇÕES AUXILIARES
# =====================================================================

def buscar_grafico(mapeamento, nome_grafico_relatorio):
    """
    Busca um gráfico específico no mapeamento.
    
    Args:
        mapeamento (dict): Dicionário de mapeamento completo
        nome_grafico_relatorio (str): Nome do gráfico no relatório (ex: "Gráfico 1")
    
    Returns:
        dict or None: Informações do gráfico ou None se não encontrado
    """
    return mapeamento.get(nome_grafico_relatorio)

def listar_nao_encontrados(mapeamento):
    """
    Lista todos os gráficos que não foram encontrados.
    
    Args:
        mapeamento (dict): Dicionário de mapeamento completo
    
    Returns:
        list: Lista de nomes de gráficos não encontrados
    """
    return [
        nome for nome, info in mapeamento.items()
        if info["status"] == "nao_encontrado"
    ]

def listar_encontrados(mapeamento):
    """
    Lista todos os gráficos que foram encontrados com sucesso.
    
    Args:
        mapeamento (dict): Dicionário de mapeamento completo
    
    Returns:
        dict: Dicionário com gráficos encontrados
    """
    return {
        nome: info for nome, info in mapeamento.items()
        if info["status"] in ["encontrado", "encontrado_similar"]
    }

# =====================================================================
#                        EXECUÇÃO
# =====================================================================

if __name__ == "__main__":
    print("=" * 70)
    print("CORRESPONDÊNCIA DE GRÁFICOS - Relatório x Imagens Extraídas")
    print("=" * 70)
    print()
    
    mapeamento = mapear_graficos(JSON_PATH, IMAGES_DIR, OUTPUT_JSON)
    
    if mapeamento:
        print("\n" + "=" * 70)
        print("EXEMPLOS DE USO:")
        print("=" * 70)
        
        # Exemplo 1: Buscar gráfico específico
        print("\n1. Buscar gráfico específico:")
        print("   grafico = buscar_grafico(mapeamento, 'Gráfico 1')")
        print("   print(grafico['caminho_completo'])")
        
        # Exemplo 2: Listar não encontrados
        nao_encontrados = listar_nao_encontrados(mapeamento)
        print(f"\n2. Gráficos não encontrados: {len(nao_encontrados)}")
        if nao_encontrados:
            print("   " + ", ".join(nao_encontrados[:5]))
            if len(nao_encontrados) > 5:
                print(f"   ... e mais {len(nao_encontrados) - 5}")
        
        # Exemplo 3: Listar encontrados
        encontrados = listar_encontrados(mapeamento)
        print(f"\n3. Gráficos encontrados: {len(encontrados)}")
