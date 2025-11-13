"""
Teste rápido para verificar a extração de "Gráfico X" de legendas completas
"""

import re

def extrair_grafico_simples(legenda_completa):
    """Extrai apenas 'Gráfico X' da legenda completa"""
    match = re.match(r'^(Gráfico\s+\d+)', legenda_completa, re.IGNORECASE)
    if match:
        return match.group(1)
    return legenda_completa

# Testes
legendas_teste = [
    "Gráfico 11 - Percentual de Magistrados(as) no Poder Judiciário.",
    "Gráfico 1 - Taxa de congestionamento total",
    "Gráfico 123 - Título muito longo com várias palavras",
    "Gráfico 5",
    "Figura 10 - Alguma coisa",
]

print("=" * 70)
print("TESTE DE EXTRAÇÃO DE NÚMERO DO GRÁFICO")
print("=" * 70)

for legenda in legendas_teste:
    resultado = extrair_grafico_simples(legenda)
    print(f"\nOriginal: {legenda}")
    print(f"Extraído: {resultado}")
    print(f"Match: {'✅' if 'Gráfico' in resultado and len(resultado) < 15 else '❌'}")
