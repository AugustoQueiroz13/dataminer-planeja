# ==============================================================================
# connectors/ppa.py
# Conector para arquivos de PPA (Plano Plurianual) municipal.
#
# O PPA define as diretrizes, objetivos e metas para um período de 4 anos.
# Estrutura típica: Eixo, Objetivo, Programa, Ação, Meta, Produto,
#                   Unidade de Medida, Valor Previsto por Exercício.
#
# Notas:
#   - PPAs cobrem 4 exercícios (ex: 2022-2025), então a coluna de valor
#     pode estar dividida por ano.
#   - A estrutura varia significativamente entre municípios.
#   - Arquivos de PPA costumam ser mais verbosos que LOA e SICONFI.
# ==============================================================================

import pandas as pd
from connectors.siconfi import _detectar_por_padroes


class PPA:
    """
    Conector para arquivos de Plano Plurianual (PPA) municipal.
    """

    TIPO = "PPA"

    PADROES = {
        "municipio": [
            "município", "municipio", "entidade", "ente", "prefeitura",
            "poder", "órgão", "orgao"
        ],
        "ibge": [
            "cód. ibge", "cod. ibge", "código ibge", "codigo ibge",
            "ibge", "cod_ibge"
        ],
        "eixo": [
            "eixo", "eixo estratégico", "eixo estrategico",
            "eixo temático", "eixo tematico", "diretriz"
        ],
        "objetivo": [
            "objetivo", "objetivo estratégico", "objetivo estrategico"
        ],
        "codigo": [
            "programa", "cod. programa", "código programa", "codigo programa",
            "ação", "acao", "cod. ação", "cod. acao", "código ação",
            "codigo acao", "projeto", "atividade", "operação", "operacao",
            "iniciativa", "meta", "código meta", "codigo meta",
            "cod. meta", "nr. ação", "nr. programa"
        ],
        "descricao": [
            "descrição", "descricao", "denominação", "denominacao",
            "nome", "especificação", "especificacao", "título", "titulo",
            "objetivo", "finalidade"
        ],
        "produto": [
            "produto", "entrega", "resultado esperado"
        ],
        "unidade_medida": [
            "unidade de medida", "unidade medida", "unidade", "und."
        ],
        "valor": [
            "valor previsto", "valor total", "total", "dotação total",
            "dotacao total", "previsão total", "previsao total",
            "recurso", "valor", "montante"
        ],
        "exercicio": [
            "exercício", "exercicio", "período", "periodo", "vigência",
            "vigencia", "ano"
        ],
        "orgao": [
            "órgão", "orgao", "secretaria", "unidade gestora",
            "ug", "responsável", "responsavel"
        ],
        "fonte_recurso": [
            "fonte", "fonte recurso", "fonte de recurso",
            "origem recurso", "cod. fonte"
        ]
    }

    @classmethod
    def detectar_colunas(cls, colunas_df: list) -> dict:
        """
        Detecta automaticamente as colunas do arquivo PPA
        com base nos padrões de nomes esperados.
        """
        return _detectar_por_padroes(cls.PADROES, colunas_df)

    @classmethod
    def aplicar_tipos(cls, df: pd.DataFrame, mapeamento: dict) -> pd.DataFrame:
        """
        Converte colunas de valor para numérico.
        O PPA pode ter múltiplas colunas de valor (uma por exercício).
        """
        resultado = df.copy()

        col_valor = mapeamento.get("valor")
        if col_valor and col_valor in resultado.columns:
            resultado[col_valor] = (
                resultado[col_valor]
                .astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
                .str.replace("R$", "", regex=False)
                .str.strip()
            )
            resultado[col_valor] = pd.to_numeric(resultado[col_valor], errors="coerce")

        return resultado