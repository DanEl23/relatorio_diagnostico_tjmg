"""
EXEMPLO DE USO DE RECUO PERSONALIZADO PARA IMAGENS

Este arquivo mostra como configurar recuos personalizados para cada imagem
no arquivo report_data.py usando o dicionário MAPA_IMAGENS.

IMPORTANTE: O recuo funciona apenas com imagens definidas no MAPA_IMAGENS,
não com as imagens do mapeamento automático (mapeamento_graficos_completo.json).
"""

# ============================================================================
# ESTRUTURA DO MAPA_IMAGENS COM SUPORTE A RECUO
# ============================================================================

MAPA_IMAGENS = {
    # Formato 1: Apenas caminho (SEM recuo personalizado)
    "Figura 1": "JN_images/figura1.png",
    
    # Formato 2: Dicionário com largura (SEM recuo personalizado)
    "Figura 2": {
        "caminho": "JN_images/figura2.png",
        "width": 14.0  # Largura em cm
    },
    
    # Formato 3: Dicionário com largura E recuo personalizado
    "Figura 3": {
        "caminho": "JN_images/figura3.png",
        "width": 16.5,   # Largura em cm
        "indent": -1.15  # Recuo em cm (negativo = para esquerda)
    },
    
    # Exemplos de diferentes recuos:
    "Gráfico 1": {
        "caminho": "JN_images/grafico1.png",
        "width": 15.0,
        "indent": -0.5   # Pequeno recuo para esquerda
    },
    
    "Gráfico 2": {
        "caminho": "JN_images/grafico2.png",
        "width": 14.0,
        "indent": 0.0    # Sem recuo (alinhamento normal)
    },
    
    "Gráfico 3": {
        "caminho": "JN_images/grafico3.png",
        "width": 17.0,
        "indent": -1.15  # Recuo igual ao da tabela 12
    },
    
    "Gráfico 4": {
        "caminho": "JN_images/grafico4.png",
        "width": 13.0,
        "indent": 1.0    # Recuo para DIREITA (positivo)
    },
}

# ============================================================================
# VALORES DE RECUO RECOMENDADOS
# ============================================================================

"""
VALORES POSITIVOS (recuo para DIREITA):
  0.5cm  - Pequeno recuo
  1.0cm  - Recuo médio
  2.0cm  - Recuo grande

VALORES NEGATIVOS (recuo para ESQUERDA):
  -0.5cm - Pequeno recuo para esquerda
  -1.15cm - Igual ao recuo da Tabela 12 (Justiça em Números)
  -2.0cm - Recuo grande para esquerda

VALOR ZERO:
  0.0cm  - Sem recuo (comportamento padrão)
"""

# ============================================================================
# COMO USAR NO report_data.py
# ============================================================================

"""
1. Abra o arquivo report_data.py

2. Localize o dicionário MAPA_IMAGENS

3. Para cada imagem que precisa de recuo personalizado, use o formato:

   "Nome da Imagem": {
       "caminho": "pasta/arquivo.png",
       "width": 16.5,      # Largura em cm
       "indent": -1.15     # Recuo em cm (pode ser negativo)
   }

4. Salve o arquivo e execute report_generator_test.py

5. O log mostrará:
   ✅ Imagem inserida com recuo de -1.15cm: arquivo.png
"""

# ============================================================================
# CASOS DE USO COMUNS
# ============================================================================

"""
USO 1: Imagem muito larga que precisa "sair" da margem
-------------------------------------------------------
Problema: Gráfico tem 18cm de largura e não cabe na área padrão
Solução: Aplicar recuo negativo de -1.15cm

MAPA_IMAGENS = {
    "Gráfico Wide": {
        "caminho": "graficos/grafico_largo.png",
        "width": 18.0,
        "indent": -1.15
    }
}


USO 2: Alinhar imagem com tabela que tem recuo
-----------------------------------------------
Problema: Tabela 12 tem recuo de -1.15cm, quero que imagem alinhe com ela
Solução: Usar o mesmo recuo na imagem

MAPA_IMAGENS = {
    "Gráfico Alinhado": {
        "caminho": "graficos/grafico_alinhado.png",
        "width": 16.5,
        "indent": -1.15  # Mesmo recuo da tabela
    }
}


USO 3: Imagem menor centralizada com deslocamento
--------------------------------------------------
Problema: Imagem pequena centralizada, mas quero deslocar um pouco
Solução: Usar recuo positivo pequeno

MAPA_IMAGENS = {
    "Figura Pequena": {
        "caminho": "figuras/fig_pequena.png",
        "width": 10.0,
        "indent": 0.5  # Desloca 0.5cm para direita
    }
}
"""

# ============================================================================
# LIMITAÇÕES IMPORTANTES
# ============================================================================

"""
⚠️ ATENÇÃO:

1. O recuo personalizado funciona APENAS para imagens definidas no MAPA_IMAGENS
   do arquivo report_data.py

2. Imagens encontradas pelo mapeamento automático (mapeamento_graficos_completo.json)
   NÃO suportam recuo personalizado (usam recuo padrão = 0.0)

3. Para usar recuo personalizado em gráficos do Justiça em Números:
   - Adicione manualmente no MAPA_IMAGENS do report_data.py
   - Use o formato de dicionário com "indent"
   
4. O alinhamento do parágrafo (CENTER) é aplicado ANTES do recuo,
   então o recuo desloca a partir da posição centralizada
"""

# ============================================================================
# EXEMPLO COMPLETO NO report_data.py
# ============================================================================

"""
# No arquivo report_data.py, adicione ou modifique o MAPA_IMAGENS:

MAPA_IMAGENS = {
    # Gráficos do Justiça em Números com recuo personalizado
    "Gráfico 1": {
        "caminho": "graficos_extraidos_por_titulo/Gráfico 78 - Taxa de congestionamento...png",
        "width": 17.0,
        "indent": -1.15
    },
    
    "Gráfico 2": {
        "caminho": "graficos_extraidos_por_titulo/Gráfico 61 - Tempo de giro...png",
        "width": 16.0,
        "indent": -0.8
    },
    
    # Outras imagens sem recuo
    "Figura 1": "JN_images/estrutura_organizacional.png",
    
    # Imagem com largura personalizada mas sem recuo
    "Figura 2": {
        "caminho": "JN_images/mapa.png",
        "width": 14.0
    }
}
"""

print("=" * 80)
print("DOCUMENTAÇÃO: Recuo Personalizado para Imagens")
print("=" * 80)
print("\n📖 Este arquivo contém exemplos e documentação sobre como usar")
print("   o recurso de recuo personalizado para imagens no relatório.")
print("\n📝 Para mais informações, leia os comentários acima.")
print("\n✅ Funcionalidade implementada com sucesso!")
print("=" * 80)
