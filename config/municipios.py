# ==============================================================================
# Mapeamento das 7 Regiões Territoriais do Programa Planeja+
# Fonte: Plano de Trabalho Planeja+ (Quadro 1, pág. 3)
# Total: 26 municípios nos estados de ES, RJ e SP
# ==============================================================================

REGIOES = {
    "Regional I - Espírito Santo": {
        "uf": ["ES"],
        "municipios": [
            "Piúma",
            "Marataízes",
            "Itapemirim",
            "Presidente Kennedy"
        ]
    },
    "Regional II - Norte Fluminense": {
        "uf": ["RJ"],
        "municipios": [
            "São Francisco de Itabapoana",
            "São João da Barra",
            "Campos dos Goytacazes",
            "Quissamã"
        ]
    },
    "Regional III - Macaé / Rio das Ostras": {
        "uf": ["RJ"],
        "municipios": [
            "Carapebus",
            "Macaé",
            "Rio das Ostras",
            "Casimiro de Abreu"
        ]
    },
    "Regional IV - Região dos Lagos": {
        "uf": ["RJ"],
        "municipios": [
            "Saquarema",
            "Araruama",
            "Arraial do Cabo",
            "Cabo Frio",
            "Armação dos Búzios"
        ]
    },
    "Regional V - Grande Niterói": {
        "uf": ["RJ"],
        "municipios": [
            "Maricá",
            "Niterói",
            "Guapimirim"
        ]
    },
    "Regional VI - Costa Verde / Litoral Norte SP": {
        "uf": ["RJ", "SP"],
        "nota": (
            "Atenção: Paraty pertence ao RJ e Caraguatatuba/Ilhabela pertencem ao SP. "
            "Ações vinculadas a políticas estaduais do RJ devem considerar Paraty separadamente."
        ),
        "municipios": [
            "Paraty",
            "Caraguatatuba",
            "Ilhabela"
        ]
    },
    "Regional VII - Litoral Sul de SP": {
        "uf": ["SP"],
        "municipios": [
            "Iguape",
            "Ilha Comprida",
            "Cananéia"
        ]
    }
}

# Lista plana e ordenada de todos os municípios (para widgets de filtro)
TODOS_MUNICIPIOS = sorted([
    municipio
    for dados in REGIOES.values()
    for municipio in dados["municipios"]
])

# Lookup rápido: nome do município -> nome da regional
MUNICIPIO_PARA_REGIAO = {
    municipio: nome_regiao
    for nome_regiao, dados in REGIOES.items()
    for municipio in dados["municipios"]
}

# Lookup rápido: nome do município -> UF (usa a UF principal da regional)
MUNICIPIO_PARA_UF = {
    "Piúma": "ES", "Marataízes": "ES", "Itapemirim": "ES", "Presidente Kennedy": "ES",
    "São Francisco de Itabapoana": "RJ", "São João da Barra": "RJ",
    "Campos dos Goytacazes": "RJ", "Quissamã": "RJ",
    "Carapebus": "RJ", "Macaé": "RJ", "Rio das Ostras": "RJ", "Casimiro de Abreu": "RJ",
    "Saquarema": "RJ", "Araruama": "RJ", "Arraial do Cabo": "RJ",
    "Cabo Frio": "RJ", "Armação dos Búzios": "RJ",
    "Maricá": "RJ", "Niterói": "RJ", "Guapimirim": "RJ",
    "Paraty": "RJ",
    "Caraguatatuba": "SP", "Ilhabela": "SP",
    "Iguape": "SP", "Ilha Comprida": "SP", "Cananéia": "SP"
}