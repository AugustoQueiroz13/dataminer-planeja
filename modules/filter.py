# ==============================================================================
# filter.py
# Filtragem de DataFrames pandas (usado para XLSX e PDF).
# Para CSV, a filtragem é feita diretamente no DuckDB via reader.py.
# ==============================================================================

import pandas as pd
from typing import Optional
from config.municipios import REGIOES, TODOS_MUNICIPIOS


def get_municipios_por_regiao(regioes_selecionadas: list) -> list:
    """
    Dado uma lista de nomes de regionais selecionadas pelo usuário,
    retorna a lista plana de municípios correspondentes.
    Se 'Todas as Regionais' estiver na lista, retorna todos os 26.
    """
    if not regioes_selecionadas or "Todas as Regionais" in regioes_selecionadas:
        return TODOS_MUNICIPIOS

    municipios = []
    for nome_regiao in regioes_selecionadas:
        if nome_regiao in REGIOES:
            municipios.extend(REGIOES[nome_regiao]["municipios"])

    return sorted(list(set(municipios)))


def filtrar_dataframe(
    df: pd.DataFrame,
    municipios: list,
    col_municipio: str,
    col_receita: str,
    termo_receita: Optional[str] = None,
    col_ano: Optional[str] = None,
    anos: Optional[list] = None
) -> pd.DataFrame:
    """
    Aplica filtros a um DataFrame pandas já carregado.
    Usado para arquivos XLSX e tabelas extraídas de PDF.

    Parâmetros:
        df            - DataFrame com os dados brutos
        municipios    - lista de nomes de municípios a manter
        col_municipio - nome da coluna com o município
        col_receita   - nome da coluna com o tipo/fonte de receita
        termo_receita - texto parcial para buscar no tipo de receita
        col_ano       - nome da coluna de ano (opcional)
        anos          - lista de anos a manter (opcional)

    Retorna DataFrame filtrado.
    """
    resultado = df.copy()

    # Filtro de municípios
    if municipios:
        resultado = resultado[resultado[col_municipio].isin(municipios)]

    # Filtro de receita por texto parcial
    if termo_receita and termo_receita.strip():
        mascara = resultado[col_receita].astype(str).str.lower().str.contains(
            termo_receita.lower(), na=False
        )
        resultado = resultado[mascara]

    # Filtro de ano
    if col_ano and anos:
        resultado = resultado[resultado[col_ano].astype(str).isin([str(a) for a in anos])]

    return resultado.reset_index(drop=True)