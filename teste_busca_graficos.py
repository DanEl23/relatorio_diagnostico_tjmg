"""
Teste rápido para verificar a função de busca automática de gráficos
"""

import json
import os

# Carregar mapeamento
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MAPEAMENTO_GRAFICOS_PATH = os.path.join(SCRIPT_DIR, "mapeamento_graficos_completo.json")

print("=" * 70)
print("TESTE DE BUSCA AUTOMÁTICA DE GRÁFICOS")
print("=" * 70)

# Carregar JSON
if os.path.exists(MAPEAMENTO_GRAFICOS_PATH):
    with open(MAPEAMENTO_GRAFICOS_PATH, 'r', encoding='utf-8') as f:
        MAPEAMENTO_GRAFICOS = json.load(f)
    print(f"✅ Mapeamento carregado: {len(MAPEAMENTO_GRAFICOS)} gráficos\n")
else:
    print(f"❌ Arquivo não encontrado: {MAPEAMENTO_GRAFICOS_PATH}")
    exit()

def buscar_caminho_grafico(legenda_chave):
    """Busca o caminho de um gráfico fazendo a tradução do dicionário"""
    if legenda_chave in MAPEAMENTO_GRAFICOS:
        info = MAPEAMENTO_GRAFICOS[legenda_chave]
        if info.get("status") == "encontrado" and info.get("caminho_completo"):
            grafico_original = info.get("grafico_original", "")
            caminho = info["caminho_completo"]
            if not os.path.isabs(caminho):
                caminho = os.path.join(SCRIPT_DIR, caminho)
            
            if os.path.exists(caminho):
                return caminho, grafico_original
    return None, None

# Testar alguns gráficos
graficos_teste = ["Gráfico 1", "Gráfico 2", "Gráfico 5", "Gráfico 10", "Gráfico 18"]

print("TESTE DE BUSCA COM TRADUÇÃO:")
print("-" * 70)
print("Formato: Conteudo_Fonte → dicionario_graficos.json → Arquivo Extraído")
print("-" * 70)

for grafico in graficos_teste:
    caminho, grafico_original = buscar_caminho_grafico(grafico)
    
    if caminho:
        nome_arquivo = os.path.basename(caminho)
        existe = "✅ EXISTE" if os.path.exists(caminho) else "❌ NÃO EXISTE"
        print(f"\n{grafico}:")
        print(f"  └─ Tradução: {grafico_original}")
        print(f"  └─ Status: {existe}")
        print(f"  └─ Arquivo: {nome_arquivo}")
    else:
        info = MAPEAMENTO_GRAFICOS.get(grafico, {})
        status = info.get("status", "não_mapeado")
        grafico_original = info.get("grafico_original", "N/A")
        print(f"\n{grafico}:")
        print(f"  └─ Tradução: {grafico_original}")
        print(f"  └─ Status: ❌ {status}")

# Estatísticas
print("-" * 70)
print("ESTATÍSTICAS:")
print("-" * 70)

encontrados = sum(1 for g in MAPEAMENTO_GRAFICOS.values() if g.get("status") == "encontrado")
nao_encontrados = sum(1 for g in MAPEAMENTO_GRAFICOS.values() if g.get("status") == "nao_encontrado")
invalidos = sum(1 for g in MAPEAMENTO_GRAFICOS.values() if g.get("status") == "numero_invalido")

print(f"Total no mapeamento: {len(MAPEAMENTO_GRAFICOS)}")
print(f"  ✅ Encontrados: {encontrados}")
print(f"  ❌ Não encontrados: {nao_encontrados}")
print(f"  ⚠️  Números inválidos: {invalidos}")

print("\n" + "=" * 70)
print("GRÁFICOS COM NÚMEROS INVÁLIDOS (precisam ser preenchidos):")
print("=" * 70)

for nome, info in MAPEAMENTO_GRAFICOS.items():
    if info.get("status") == "numero_invalido":
        print(f"  {nome}")
