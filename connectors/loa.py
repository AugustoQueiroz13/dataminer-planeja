# ==============================================================================
# connectors/loa.py
# Conector para arquivos de LOA (Lei Orçamentária Anual) municipal.
#
# A LOA detalha as receitas e despesas previstas para o exercício.
# As colunas variam entre municípios, mas seguem a estrutura da
# classificação orçamentária brasileira (STN/SOF).
#
# Estruturas comuns de LOA:
#   Receitas: Código, Especificação, Previsão Inicial, Previsão Atualizada
#   Despesas: Órgão, Unidade, Função, Subfunção, Programa, Ação,
#             Natureza Despesa, Dotação Inicial, Dotação Atualizada
# ==============================================================================

import pandas as pd
from connectors.siconfi import _detectar_por_padroes


class LOA:
    """
    Conector para arquivos de Lei Orçamentária Anual (LOA) municipal.
    Cobre tanto a parte de receitas quanto de despesas da LOA.
    """

    TIPO = "LOA"

    PADROES = {
        "municipio": [
            "município", "municipio", "entidade", "ente", "prefeitura",
            "órgão gestor", "orgao gestor", "poder"
        ],
        "ibge": [
            "cód. ibge", "cod. ibge", "código ibge", "codigo ibge",
            "ibge", "cod_ibge"
        ],
        "orgao": [
            "órgão", "orgao", "unidade gestora", "ug", "secretaria",
            "cod. órgão", "cod. orgao", "código órgão"
        ],
        "funcao": [
            "função", "funcao", "cod. função", "cod. funcao",
            "código função", "codigo funcao"
        ],
        "subfuncao": [
            "subfunção", "subfuncao", "subfunção", "cod. subfunção",
            "cod. subfuncao"
        ],
        "programa": [
            "programa", "cod. programa", "código programa",
            "codigo programa", "nr. programa"
        ],
        "codigo": [
            "ação", "acao", "atividade", "projeto", "operação especial",
            "operacao especial", "cod. ação", "cod. acao", "código ação",
            "codigo acao", "natureza", "natureza despesa", "elemento",
            "cod. natureza", "rubrica", "classificação receita",
            "classificacao receita", "cod. receita", "codigo receita",
            "especificação", "especificacao", "fonte"
        ],
        "descricao": [
            "descrição", "descricao", "denominação", "denominacao",
            "especificação", "especificacao", "nome", "título", "titulo"
        ],
        "valor": [
            "dotação inicial", "dotacao inicial", "dotação", "dotacao",
            "previsão inicial", "previsao inicial", "previsão", "previsao",
            "dotação atualizada", "dotacao atualizada",
            "previsão atualizada", "previsao atualizada",
            "valor", "total", "montante"
        ],
        "valor_atualizado": [
            "dotação atualizada", "dotacao atualizada",
            "previsão atualizada", "previsao atualizada",
            "valor atualizado", "total atualizado"
        ],
        "exercicio": [
            "exercício", "exercicio", "ano", "período", "periodo"
        ],
        "fonte_recurso": [
            "fonte de recurso", "fonte recurso", "fonte recursos",
            "fonte", "cod. fonte", "codigo fonte"
        ]
    }

    @classmethod
    def detectar_colunas(cls, colunas_df: list) -> dict:
        """
        Detecta automaticamente as colunas do arquivo LOA
        com base nos padrões de nomes esperados.
        """
        return _detectar_por_padroes(cls.PADROES, colunas_df)

    @classmethod
    def aplicar_tipos(cls, df: pd.DataFrame, mapeamento: dict) -> pd.DataFrame:
        """
        Converte colunas de valor para numérico, tratando
        formatação brasileira (ponto como separador de milhar,
        vírgula como decimal).
        """
        resultado = df.copy()

        for campo_valor in ["valor", "valor_atualizado"]:
            col = mapeamento.get(campo_valor)
            if col and col in resultado.columns:
                resultado[col] = (
                    resultado[col]
                    .astype(str)
                    .str.replace(".", "", regex=False)
                    .str.replace(",", ".", regex=False)
                    .str.replace("R$", "", regex=False)
                    .str.strip()
                )
                resultado[col] = pd.to_numeric(resultado[col], errors="coerce")

        return resultado