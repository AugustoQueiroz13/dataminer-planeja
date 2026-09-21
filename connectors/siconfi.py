# ==============================================================================
# connectors/siconfi.py
# Conector para arquivos do SICONFI (Sistema de Informações Contábeis e
# Fiscais do Setor Público Brasileiro) do Tesouro Nacional.
#
# Colunas típicas do SICONFI:
#   Município, Cód. IBGE, UF, Populacao, Conta, Fonte da Receita, Receita
# ==============================================================================

import pandas as pd


class SICONFI:
    """
    Conector para arquivos exportados do SICONFI.
    Define os padrões de nomes de colunas e fornece
    a detecção automática para arquivos reais.
    """

    TIPO = "SICONFI"

    # Palavras-chave para detecção de cada campo
    # Ordem: da mais específica para a mais genérica
    PADROES = {
        "municipio": [
            "município", "municipio", "nome município", "nome municipio",
            "nome do município", "nome_municipio"
        ],
        "ibge": [
            "cód. ibge", "cod. ibge", "código ibge", "codigo ibge",
            "ibge", "cod_ibge", "codibge", "cód ibge"
        ],
        "uf": [
            "uf", "estado", "sigla uf", "sigla_uf", "sg_uf"
        ],
        "populacao": [
            "populacao", "população", "pop", "habitantes"
        ],
        "codigo": [
            "conta", "identificador conta", "cod conta", "código conta",
            "codigo conta", "cod. conta", "identificador_conta",
            "cod_conta", "fonte da receita", "fonte", "rubrica"
        ],
        "descricao": [
            "descrição", "descricao", "nome conta", "especificação",
            "especificacao", "nome da receita", "titulo", "título"
        ],
        "valor": [
            "receita", "valor", "receita realizada", "valor receita",
            "arrecadado", "arrecadação", "arrecadacao", "total"
        ],
        "exercicio": [
            "exercício", "exercicio", "ano", "ano exercício",
            "ano_exercicio", "competência", "competencia"
        ]
    }

    @classmethod
    def detectar_colunas(cls, colunas_df: list) -> dict:
        """
        Recebe a lista de colunas do DataFrame e retorna um dicionário
        com o mapeamento campo -> nome_da_coluna_encontrada.
        Campos não encontrados ficam com valor None.
        """
        return _detectar_por_padroes(cls.PADROES, colunas_df)

    @classmethod
    def aplicar_tipos(cls, df: pd.DataFrame, mapeamento: dict) -> pd.DataFrame:
        """
        Converte as colunas do DataFrame para os tipos corretos
        de acordo com o padrão SICONFI.
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


# ==============================================================================
# Função utilitária compartilhada entre conectores
# ==============================================================================

def _detectar_por_padroes(padroes: dict, colunas_df: list) -> dict:
    """
    Para cada campo definido em 'padroes', encontra a coluna do DataFrame
    que mais se aproxima das palavras-chave esperadas.

    Estratégia de pontuação:
        3 pontos: coluna é idêntica à palavra-chave (case-insensitive)
        2 pontos: coluna começa com a palavra-chave
        1 ponto:  palavra-chave está contida no nome da coluna
        0 pontos: sem correspondência
    """
    colunas_lower = {col.lower().strip(): col for col in colunas_df}
    mapeamento = {}

    for campo, palavras_chave in padroes.items():
        melhor_coluna = None
        melhor_pontuacao = 0

        for col_lower, col_original in colunas_lower.items():
            for palavra in palavras_chave:
                palavra_lower = palavra.lower().strip()

                if col_lower == palavra_lower:
                    pontuacao = 3
                elif col_lower.startswith(palavra_lower):
                    pontuacao = 2
                elif palavra_lower in col_lower:
                    pontuacao = 1
                else:
                    continue

                if pontuacao > melhor_pontuacao:
                    melhor_pontuacao = pontuacao
                    melhor_coluna = col_original

        mapeamento[campo] = melhor_coluna

    return mapeamento