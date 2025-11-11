# report_data.py
# Este arquivo armazena todos os dados brutos (hardcoded)
# para manter o script principal limpo.

# --- DADOS TABELA 01 (ATOS) ---
dados_tabela_atos = [
    # Cabeçalho
    ("Ato Normativo", "Estrutura"),
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

# --- DADOS TABELA 02 (ÁREAS) ---
dados_tabela_areas = [
    # Tipo, Col 1, Col 2
    ("HEADER_MAIN", "DENOMINAÇÃO", ""),
    ("DATA_MERGED", "Comitê Estratégico de Gestão Institucional", ""),
    ("DATA_MERGED", "Comitê Gestor de Segurança da Informação", ""),
    ("DATA_MERGED", "Comitê Institucional de Inteligência", ""),
    ("DATA_MERGED", "Comitê de Governança e Gestão Estratégica", ""),
    ("DATA_MERGED", "Comitê de Monitoramento e Suporte à Prestação Jurisdicional", ""),
    ("DATA_MERGED", "Comitê de Tecnologia da Informação e Comunicação", ""),
    ("HEADER_GROUP_SIGLA", "SUPERINTENDÊNCIA ADMINISTRATIVA", "SIGLA"),
    ("DATA_SPLIT", "Diretoria Executiva de Administração de Recursos Humanos", "DEARHU"),
    ("DATA_SPLIT", "Diretoria Executiva de Comunicação", "DIRCOM"),
    ("DATA_SPLIT", "Diretoria Executiva de Engenharia e Gestão Predial", "DENGEP"),
    ("DATA_SPLIT", "Diretoria Executiva de Finanças e Execução Orçamentária", "DIRFIN"),
    ("DATA_SPLIT", "Diretoria Executiva de Gestão de Bens, Serviços e Patrimônio", "DIRSEP"),
    ("DATA_SPLIT", "Diretoria Executiva de Informática", "DIRTEC"),
    ("DATA_SPLIT", "Diretoria Executiva de Planejamento Orçamentário e Qualidade na Gestão Institucional", "DEPLAG"),
    ("DATA_SPLIT", "Gabinete de Segurança Institucional", "GSI"),
    ("DATA_SPLIT", "Secretaria de Auditoria Interna", "SECAUD"),
    ("DATA_SPLIT", "Secretaria de Governança e Gestão Estratégica", "SEGOVE"),
    ("DATA_SPLIT", "Secretaria do Órgão Especial", "SEOESP"),
    ("HEADER_GROUP_MERGED", "SUPERINTENDÊNCIA DO 1º VICE-PRESIDENTE", ""),
    ("DATA_SPLIT", "Diretoria Executiva de Suporte à Prestação Jurisdicional", "DIRSUP"),
    ("DATA_SPLIT", "Secretaria de Padronização e Acompanhamento da Gestão Judiciária", "SEPAD"),
    ("HEADER_GROUP_MERGED", "SUPERINTENDÊNCIA DO 2º VICE-PRESIDENTE", ""),
    ("DATA_SPLIT", "Diretoria Executiva de Desenvolvimento de Pessoas", "DIRDEP"),
    ("DATA_SPLIT", "Diretoria Executiva de Gestão da Informação Documental", "DIRGED"),
    ("HEADER_GROUP_MERGED", "SUPERINTENDÊNCIA DO 3º VICE-PRESIDENTE", ""),
    ("DATA_SPLIT", "Assessoria de Gestão da Inovação", "AGIN"),
    ("DATA_SPLIT", "Núcleo Permanente de Métodos Consensuais de Solução de Conflitos", "NUPEMEC"),
    ("HEADER_GROUP_MERGED", "CORREGEDORIA-GERAL DE JUSTIÇA", ""),
    ("DATA_SPLIT", "Diretoria Executiva de Atividade Correcional", "DIRCOR"),
    ("DATA_SPLIT", "Secretaria de Suporte ao Planejamento e à Gestão da Primeira Instância", "SEPLAN")
]


dados_tabela_estrutura = [
    # Tipo, Col 1
    ("HEADER_MAIN", "ESTRUTURAS PARA A PRESTAÇÃO JURISDICIONAL NA SEGUNDA INSTÂNCIA"),
    
    ("HEADER_GROUP_MERGED", "Câmaras Cíveis"),
    ("DATA_MERGED", "01ª Câmara Cível"),
    ("DATA_MERGED", "02ª Câmara Cível"),
    ("DATA_MERGED", "03ª Câmara Cível"),
    ("DATA_MERGED", "04ª Câmara Cível Especializada"),
    ("DATA_MERGED", "05ª Câmara Cível"),
    ("DATA_MERGED", "06ª Câmara Cível"),
    ("DATA_MERGED", "07ª Câmara Cível"),
    ("DATA_MERGED", "08ª Câmara Cível Especializada"),
    ("DATA_MERGED", "09ª Câmara Cível"),
    ("DATA_MERGED", "10ª Câmara Cível"),
    ("DATA_MERGED", "11ª Câmara Cível"),
    ("DATA_MERGED", "12ª Câmara Cível"),
    ("DATA_MERGED", "13ª Câmara Cível"),
    ("DATA_MERGED", "14ª Câmara Cível"),
    ("DATA_MERGED", "15ª Câmara Cível"),
    ("DATA_MERGED", "16ª Câmara Cível Especializada"),
    ("DATA_MERGED", "17ª Câmara Cível"),
    ("DATA_MERGED", "18ª Câmara Cível"),
    ("DATA_MERGED", "19ª Câmara Cível"),
    ("DATA_MERGED", "20ª Câmara Cível"),
    ("DATA_MERGED", "21ª Câmara Cível Especializada"),

    ("HEADER_GROUP_MERGED", "Câmaras Criminais"),
    ("DATA_MERGED", "01ª Câmara Criminal"),
    ("DATA_MERGED", "02ª Câmara Criminal"),
    ("DATA_MERGED", "03ª Câmara Criminal"),
    ("DATA_MERGED", "04ª Câmara Criminal"),
    ("DATA_MERGED", "05ª Câmara Criminal"),
    ("DATA_MERGED", "06ª Câmara Criminal"),
    ("DATA_MERGED", "07ª Câmara Criminal"),
    ("DATA_MERGED", "08ª Câmara Criminal"),
    ("DATA_MERGED", "09ª Câmara Criminal"),

    ("HEADER_GROUP_MERGED", "Justiça 4.0 - Cível"),
    ("DATA_MERGED", "Câmara Justiça 4.0 - Especializada Cível - 4"),
    ("DATA_MERGED", "Câmara Justiça 4.0 - Especializada Cível - 8"),
    ("DATA_MERGED", "Câmara Justiça 4.0 - Cível - 18"),
    
    ("HEADER_GROUP_MERGED", "Justiça 4.0 - Criminal"),
    ("DATA_MERGED", "Câmara Justiça 4.0 - Especializada Criminal"),

]

dados_tabela_comarcas = [
    # Tipo, Col 1, Col 2, Col 3, Col 4
    ("HEADER_MERGE_4", "COMARCAS INSTALADAS", "", "", ""),
    ("DATA_4_COL", "Abaeté", "Abre Campo", "Açucena", "Águas Formosas"),
    ("DATA_4_COL", "Aimorés", "Aiuruoca", "Além Paraíba", "Alvinópolis"),
    ("DATA_4_COL", "Andradas", "Andrelândia", "Alfenas", "Almenara"),
    ("DATA_4_COL", "Areado", "Arinos", "Alpinópolis", "Alto Rio Doce"),
    ("DATA_4_COL", "Araçuaí", "Araguari", "Araxá", "Arcos"),
    ("DATA_4_COL", "Baependi", "Bambuí", "Barão de Cocais", "Barbacena"),
    ("DATA_4_COL", "Barroso", "Belo Horizonte", "Belo Vale", "Betim"),
    ("DATA_4_COL", "Bicas", "Boa Esperança", "Bocaiúva", "Bom Despacho"),
    ("DATA_4_COL", "Bom Sucesso", "Bonfim", "Bonfinópolis de Minas", "Borda da Mata"),
    ("DATA_4_COL", "Botelhos", "Brasília de Minas", "Brazópolis", "Brumadinho"),
    ("DATA_4_COL", "Bueno Brandão", "Buenópolis", "Buritis", "Cabo Verde"),
    ("DATA_4_COL", "Cachoeira de Minas", "Caeté", "Caldas", "Camanducaia"),
    ("DATA_4_COL", "Cambuí", "Cambuquira", "Campanha", "Campestre"),
    ("DATA_4_COL", "Campina Verde", "Campo Belo", "Campos Altos", "Campos Gerais"),
    ("DATA_4_COL", "Canápolis", "Candeias", "Capelinha", "Capinópolis"),
    ("DATA_4_COL", "Carandaí", "Carangola", "Caratinga", "Carlos Chagas"),
    ("DATA_4_COL", "Carmo da Mata", "Carmo de Minas", "Carmo do Cajuru", "Carmo do Paranaíba"),
    ("DATA_4_COL", "Carmo do Rio Claro", "Carmópolis de Minas", "Cássia", "Cataguases"),
    ("DATA_4_COL", "Caxambu", "Cláudio", "Conceição das Alagoas", "Conceição do Mato Dentro"),
    ("DATA_4_COL", "Conceição do Rio Verde", "Congonhas", "Conquista", "Conselheiro Lafaiete"),
    ("DATA_4_COL", "Conselheiro Pena", "Contagem", "Coração de Jesus", "Corinto"),
    ("DATA_4_COL", "Coromandel", "Coronel Fabriciano", "Cristina", "Cruzília"),
    ("DATA_4_COL", "Curvelo", "Diamantina ", "Divino", "Divinópolis"),
    ("DATA_4_COL", "Dores do Indaiá", "Elói Mendes", "Entre-Rios de Minas", "Ervália"),
    ("DATA_4_COL", "Esmeraldas", "Espera Feliz", "Espinosa", "Estrela do Sul"),
    ("DATA_4_COL", "Eugenópolis", "Extrema", "Ferros", "Formiga"),
    ("DATA_4_COL", "Francisco Sá", "Frutal", "Galiléia", "Governador Valadares"),
    ("DATA_4_COL", "Grão-Mogol", "Guanhães", "Guapé", "Guaranésia"),
    ("DATA_4_COL", "Guarani", "Guaxupé", "Ibiá", "Ibiraci"),
    ("DATA_4_COL", "Ibirité", "Igarapé", "Iguatama", "Inhapim"),
    ("DATA_4_COL", "Ipanema", "Ipatinga", "Itabira", "Itabirito"),
    ("DATA_4_COL", "Itaguara", "Itajubá", "Itamarandiba", "Itambacuri"),
    ("DATA_4_COL", "Itamogi", "Itamonte", "Itanhandu", "Itanhomi"),
    ("DATA_4_COL", "Itapagipe", "Itapecerica", "Itaúna", "Ituiutaba"),
    ("DATA_4_COL", "Itumirim", "Iturama", "Jaboticatubas", "Jacinto"),
    ("DATA_4_COL", "Jacuí", "Jacutinga", "Jaíba", "Janaúba"),
    ("DATA_4_COL", "Januária", "Jequeri", "Jequitinhonha", "João Monlevade"),
    ("DATA_4_COL", "João Pinheiro", "Juatuba", "Juíz de Fora", "Lagoa da Prata"),
    ("DATA_4_COL", "Lagoa Santa", "Lajinha", "Lambari", "Lavras"),
    ("DATA_4_COL", "Leopoldina", "Lima Duarte", "Luz", "Machado"),
    ("DATA_4_COL", "Malacacheta", "Manga", "Manhuaçu", "Manhumirim"),
    ("DATA_4_COL", "Mantena", "Mar de Espanha", "Mariana", "Martinho Campos"),
    ("DATA_4_COL", "Mateus Leme", "Matias Barbosa", "Matozinhos", "Medina"),
    ("DATA_4_COL", "Mercês", "Mesquita", "Minas Novas", "Miradouro"),
    ("DATA_4_COL", "Miraí", "Montalvânia", "Monte Alegre de Minas", "Monte Azul"),
    ("DATA_4_COL", "Monte Belo", "Monte Carmelo", "Monte Santo de Minas", "Monte Sião"),
    ("DATA_4_COL", "Montes Claros", "Morada Nova de Minas", "Muriaé", "Mutum"),
    ("DATA_4_COL", "Muzambinho", "Natércia", "Nepomuceno", "Nova Era"),
    ("DATA_4_COL", "Nova Lima", "Nova Ponte", "Nova Resende", "Nova Serrana"),
    ("DATA_4_COL", "Novo Cruzeiro", "Oliveira", "Ouro Branco", "Ouro Fino"),
    ("DATA_4_COL", "Ouro Preto", "Palma", "Pará de Minas", "Paracatu"),
    ("DATA_4_COL", "Paraguaçu", "Paraisópolis", "Paraopeba ", "Passa Quatro"),
    ("DATA_4_COL", "Passa Tempo", "Passos", "Patos de Minas", "Patrocínio"),
    ("DATA_4_COL", "Peçanha", "Pedra Azul", "Pedralva", "Pedro Leopoldo"),
    ("DATA_4_COL", "Perdizes", "Perdões", "Piranga", "Pirapetinga"),
    ("DATA_4_COL", "Pirapora", "Pitangui", "Piumhi", "Poço Fundo"),
    ("DATA_4_COL", "Poços de Caldas", "Pompéu", "Ponte Nova", "Porteirinha"),
    ("DATA_4_COL", "Pouso Alegre", "Prados", "Prata", "Pratápolis"),
    ("DATA_4_COL", "Presidente Olegário", "Raul Soares", "Resende Costa", "Resplendor"),
    ("DATA_4_COL", "Ribeirão das Neves", "Rio Casca", "Rio Novo", "Rio Paranaíba"),
    ("DATA_4_COL", "Rio Pardo de Minas", "Rio Piracicaba", "Rio Pomba", "Rio Preto"),
    ("DATA_4_COL", "Rio Vermelho", "Sabará", "Sabinópolis", "Sacramento"),
    ("DATA_4_COL", "Salinas", "Santa Bárbara", "Santa Luzia", "Santa Maria do Suaçuí"),
    ("DATA_4_COL", "Santa Rita de Caldas", "Santa Rita do Sapucaí", "Santa Vitória", "Santo Antônio do Monte"),
    ("DATA_4_COL", "Santos Dumont", "São Domingos do Prata", "São Francisco", "São Gonçalo do Sapucaí"),
    ("DATA_4_COL", "São Gotardo", "São João da Ponte", "São João Del Rei", "São João do Paraíso"),
    ("DATA_4_COL", "São João Evangelista", "São João Nepomuceno", "São Lourenço", "São Romão"),
    ("DATA_4_COL", "São Roque de Minas", "São Sebastião do Paraíso", "Senador Firmino", "Serro"),
    ("DATA_4_COL", "Sete Lagoas", "Silvianópolis", "Taiobeiras", "Tarumirim"),
    ("DATA_4_COL", "Teixeiras", "Teófilo Otoni", "Timóteo", "Tiros"),
    ("DATA_4_COL", "Tombos", "Três Corações", "Três Marias", "Três pontas"),
    ("DATA_4_COL", "Tupaciguara", "Turmalina", "Ubá", "Uberaba"),
    ("DATA_4_COL", "Uberlândia", "Unaí", "Varginha", "Várzea da Palma"),
    ("DATA_4_COL", "Vazante", "Vespasiano", "Viçosa", "Virginópolis"),
    ("DATA_4_COL", "Visconde do Rio Branco", "", "", "")
]

dados_tabela_nucleos = [
    # Tipo, Col 1
    ("HEADER_GROUP_MERGED", "Núcleos de Justiça 4.0 – 1ª Instância"), 
    ("DATA_MERGED", "Núcleo de Justiça 4.0 - Cooperação Judiciária"),
    ("DATA_MERGED", "Núcleo de Justiça 4.0 – Cível"),
    ("DATA_MERGED", "Núcleo de Justiça 4.0 – Criminal"),
    ("DATA_MERGED", "Núcleo de Justiça 4.0 - Fazenda Pública"),
    ("DATA_MERGED", "CEMES - Central de Execução de Medidas de Segurança 4.0")
]

dados_tabela_processos = [
    # Tipo, Col 1, Col 2, Col 3, Col 4, Col 5, Col 6, Col 7
    ("HEADER_MERGE", "PROCESSOS DISTRIBUÍDOS", "", "", "", "", "", ""),
    ("SUB_HEADER", "Instância", "2020", "2021", "2022", "2023", "2024", "Média"),
    ("DATA_ROW", "Justiça Comum", "1.191.628", "1.365.924", "1.565.819", "1.710.153", "1.675.686", "1.501.842"),
    ("DATA_ROW", "Juizado Especial", "534.375", "536.797", "558.504", "622.683", "661.356", "582.743"),
    ("DATA_ROW", "Turma Recursal", "56.088", "84.268", "84.215", "93.299", "103.728", "84.320"),
    ("DATA_ROW", "2º Grau", "199.457", "222.614", "227.760", "271.256", "334.528", "251.123"),
    ("TOTAL_ROW", "Total", "1.981.548", "2.209.603", "2.436.298", "2.697.391", "2.775.298", "2.420.028")
]

dados_tabela_julgamentos = [
    # Tipo, Col 1, Col 2, Col 3, Col 4, Col 5, Col 6, Col 7
    ("HEADER_MERGE", "JULGAMENTOS", "", "", "", "", "", ""),
    ("SUB_HEADER", "Instância", "2020", "2021", "2022", "2023", "2024", "Média"),
    ("DATA_ROW", "Justiça Comum", "878.705", "1.015.223", "1.185.589", "1.320.950", "1.412.397", "1.162.573"),
    ("DATA_ROW", "Juizado Especial", "460.286", "636.208", "810.834", "932.469", "920.189", "751.997"),
    ("DATA_ROW", "Turma Recursal", "878.705", "67.797", "77.926", "105.764", "117.904", "249.619"),
    ("DATA_ROW", "2º Grau", "52.746", "225.454", "236.418", "275.286", "337.993", "225.579"),
    ("TOTAL_ROW", "Total", "2.270.442", "1.944.682", "2.310.767", "2.634.469", "2.788.483", "2.389.769")
]

dados_tabela_acervo = [
    # Tipo, Col 1, Col 2, Col 3, Col 4, Col 5, Col 6, Col 7
    ("HEADER_MERGE", "ACERVO DE FEITOS ATIVOS NO ÚLTIMO DIA DO ANO", "", "", "", "", "", ""),
    ("SUB_HEADER", "Instância", "2020", "2021", "2022", "2023", "2024", "Média"),
    ("DATA_ROW", "Justiça Comum", "4.255.163", "4.152.223", "4.233.968", "4.140.228", "4.042.435", "4.164.803"),
    ("DATA_ROW", "Juizado Especial", "1.125.405", "1.125.081", "1.053.185", "963.386", "922.153", "1.037.842"),
    ("DATA_ROW", "Turma Recursal", "41.272", "67.940", "69.541", "76.573", "87.801", "68.625"),
    ("DATA_ROW", "2º Grau", "224.715", "232.448", "224.156", "220.826", "206.944", "221.818"),
    ("TOTAL_ROW", "Total", "5.646.555", "5.577.692", "5.580.850", "5.401.013", "5.259.333", "5.493.089")
]

# report_data.py (Versão Atualizada)

# --- 9. DADOS TABELA 09 (ORÇAMENTO) ---
TITULO_TABELA_ORCAMENTO = "Unidade Orçamentária 1031 – TJMG | Despesa Realizada por Ação Orçamentária – 2024"

dados_tabela_orcamento = [
    # A linha HEADER_MERGE foi removida daqui!
    ("SUB_HEADER", "AÇÃO ORÇAMENTÁRIA", "DESPESA REALIZADA 2024 (R$)", "", "", "", "", ""),
    ("DATA_ROW", "7004 - Precatórios e Sentenças Judiciárias", "-", "", "", "", "", ""),
    ("DATA_ROW", "7006 - Proventos de Inativos Civis e Pensionistas", "2.535.040.959,40", "", "", "", "", ""),
    ("DATA_ROW", "2053 - Remuneração de Magistrados da Ativa E Encargos Sociais", "1.353.944.848,00", "", "", "", "", ""),
    ("DATA_ROW", "2054 - Remuneração de Servidores da Ativa e Encargos Sociais", "5.448.469.921,18", "", "", "", "", ""),
    ("TOTAL_ROW", "TOTAL", "9.337.455.728,58", "", "", "", "", "")
]

TITULO_TABELA_ORCAMENTO_ACAO = "Unidade Orçamentária 4031 – FEPJ | Despesa Realizada por Ação Orçamentária – 2024"

dados_tabela_orcamento_acao = [
    # Tipo, Col 1, Col 2, Col 3, Col 4, Col 5, Col 6, Col 7
    ("SUB_HEADER", "AÇÃO ORÇAMENTÁRIA", "DESPESA REALIZADA 2024 (R$)", "", "", "", "", ""),
    ("DATA_ROW", "2025 - Gestão de Serviços De TIC", "192.893.477,53", "", "", "", "", ""),
    ("DATA_ROW", "2055 - Auxílios Concedidos a Magistrados e Servidores", "158.320.267,00", "", "", "", "", ""),
    ("DATA_ROW", "2091 - Obras e Gestão Predial", "520.565.747,04", "", "", "", "", ""),
    ("DATA_ROW", "2109 - Formação, Aperfeiçoamento e Desenvolvimento Contínuo De Pessoas", "4.421.386,45", "", "", "", "", ""),
    ("DATA_ROW", "4395 - Processamento Judiciário", "1.398.500.019,48", "", "", "", "", ""),
    ("TOTAL_ROW", "TOTAL", "2.274.700.897,50", "", "", "", "", "")
]

dados_tabela_orcamento_2025 = [
    # Título do Grupo 1 (UO 1031)
    ("GROUP_TITLE", "UO 1031 – TJMG", ""), 
    ("SUB_HEADER", "Ação Orçamentária", "Ação Orçamentária"),
    ("DATA_ROW", "7004 - Precatórios e Sentenças Judiciais", "7004 - Precatórios e Sentenças Judiciais"),
    ("DATA_ROW", "7006 - Proventos de Inativos Civis e Pensionistas", "7006 - Proventos de Inativos Civis e Pensionistas"),
    ("DATA_ROW", "2053 - Remuneração de Magistrados da Ativa", "2053 - Remuneração de Magistrados da Ativa"),
    ("DATA_ROW", "2054 - Remuneração de Servidores da Ativa", "2054 - Remuneração de Servidores da Ativa"),
    ("TOTAL_ROW", "VALOR TOTAL – UO 1031", "VALOR TOTAL – UO 1031"), # Fim do primeiro bloco
    
    # Título do Grupo 2 (UO 4031) - Começa na linha seguinte
    ("GROUP_TITLE", "UO 4031 – FEPJ", ""), 
    ("SUB_HEADER", "Ação Orçamentária", "Ação Orçamentária"),
    ("DATA_ROW", "2025 - Gestão de Serviços De TIC", "2025 - Gestão de Serviços De TIC"),
    ("DATA_ROW", "2055 - Auxílios Concedidos a Magistrados", "2055 - Auxílios Concedidos a Magistrados"),
    ("DATA_ROW", "2091 - Obras e Gestão Predial", "2091 - Obras e Gestão Predial"),
    ("DATA_ROW", "2109 - Formação, Aperfeiçoamento e Desenvolvimento Contínuo De Pessoas", "2109 - Formação, Aperfeiçoamento e Desenvolvimento Contínuo De Pessoas"),
    ("DATA_ROW", "4395 - Processamento Judiciário", "4395 - Processamento Judiciário"),
    ("TOTAL_ROW", "VALOR TOTAL – UO 4031", "VALOR TOTAL – UO 4031"),
]

dados_tabela_cidades = [
    # Tipo, Col 1, Col 2, Col 3, Col 4
    ("DATA_ROW", "Brasília de Minas", "Caeté", "Frutal", "Itajubá"),
    ("DATA_ROW", "João Monlevade", "Manhuaçu", "Mariana", "Monte Carmelo"),
    ("DATA_ROW", "Montes Claros", "Muriaé", "Nova Serrana", "Ponte Nova"),
    ("DATA_ROW", "Porteirinha", "Uberaba", "Unaí", "Vazante"),
]

# --- MAPA DE RECURSOS DE IMAGEM ---
MAPA_IMAGENS = {
    # Chave: "Legenda" (sem a Fonte)
    # Valor: "caminho/relativo/do_arquivo.png"
    
    "Figura 01 - Informações sobre o Estado de Minas Gerais.": "canvas_images/figura_01.png",
    "Figura 02 - Síntese da estrutura na área fim.": "canvas_images/figura_02.png",
    "Figura 03 - Novas estruturas na área fim.": "canvas_images/figura_03.png",
    "Figura 04 - Força de Trabalho.": "canvas_images/figura_04.png",
    "Figura 05 - Colaboradores da Justiça.": "canvas_images/figura_05.png",
    "Figura 06 – Força de Trabalho e Colaboradores na área de TI.": "canvas_images/figura_06.png",
    "Figura 07 - Instalações prediais do TJMG.":"canvas_images/figura_07.png",
    "Figura 08 - Desempenho da ação por programa (Unidade 1031).":"canvas_images/figura_08.png",
    "Figura 09 - Desempenho da ação por programa (Unidade 4031).":"canvas_images/figura_09.png"
    # Adicionar outras figuras/gráficos aqui
    # "Gráfico 01 - ...": "resources/grafico_01.png",
}