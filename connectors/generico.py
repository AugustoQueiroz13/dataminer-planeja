# ==============================================================================
# connectors/generico.py
# Conector genérico para documentos orçamentários não mapeados
# (LDO, Balanços, relatórios avulsos, etc.).
#
# Estratégia: combina todos os padrões conhecidos dos outros conectores
# e tenta a melhor correspondência possível com as colunas reais.
# É o fallback quando nenhum conector específico se encaixa.
# ==============================================================================

import re
import pandas as pd
from connectors.siconfi import _detectar_por_padroes


class Generico:
    """
    Conector genérico com vocabulário amplo para detecção de colunas.
    Funciona para LDO, Balanços Orçamentários e outros documentos fiscais.
    """

    TIPO = "Genérico"

    # Vocabulário expandido que cobre todos os tipos de documentos orçamentários
    PADROES = {
        "municipio": [
            "município", "municipio", "entidade", "ente", "prefeitura",
            "poder", "órgão", "orgao", "nome município", "nome municipio"
        ],
        "ibge": [
            "cód. ibge", "cod. ibge", "ibge", "cod_ibge", "código ibge",
            "codigo ibge"
        ],
        "codigo": [
            # Classificação de receita
            "conta", "rubrica", "código", "codigo", "cod.", "cód.",
            "classificação", "classificacao", "identificador",
            # Receita
            "fonte da receita", "fonte receita", "cod. receita",
            "classificação receita",
            # Despesa
            "ação", "acao", "projeto", "atividade", "programa",
            "função", "funcao", "subfunção", "subfuncao",
            "natureza", "elemento", "subelemento",
            # Genérico
            "nr.", "número", "numero", "item", "código item"
        ],
        "descricao": [
            "descrição", "descricao", "especificação", "especificacao",
            "denominação", "denominacao", "nome", "título", "titulo",
            "discriminação", "discriminacao", "detalhamento"
        ],
        "valor": [
            "valor", "total", "montante", "dotação", "dotacao",
            "previsão", "previsao", "receita", "despesa",
            "arrecadado", "arrecadação", "arrecadacao",
            "realizado", "executado", "pago", "liquidado",
            "saldo", "recurso"
        ],
        "exercicio": [
            "exercício", "exercicio", "ano", "período", "periodo",
            "competência", "competencia", "vigência", "vigencia"
        ],
        "uf": [
            "uf", "estado", "sigla", "sg_uf"
        ],
        "fonte_recurso": [
            "fonte", "fonte recurso", "fonte de recurso",
            "origem", "cod. fonte"
        ]
    }

    @classmethod
    def detectar_colunas(cls, colunas_df: list) -> dict:
        """
        Tenta mapear as colunas do DataFrame para campos conhecidos
        usando o vocabulário expandido do conector genérico.
        """
        return _detectar_por_padroes(cls.PADROES, colunas_df)

    @classmethod
    def aplicar_tipos(cls, df: pd.DataFrame, mapeamento: dict) -> pd.DataFrame:
        """
        Converte a coluna de valor para numérico, se detectada.
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

    @classmethod
    def inspecionar_dataframe(cls, df: pd.DataFrame) -> dict:
        """
        Analisa o DataFrame e retorna informações de diagnóstico
        para ajudar o usuário a identificar as colunas manualmente
        quando a detecção automática for insuficiente.

        Retorna um dicionário com:
            - colunas_texto: colunas com conteúdo de texto
            - colunas_numericas: colunas com conteúdo numérico
            - colunas_mistas: colunas com conteúdo misto (podem ser códigos)
            - amostras: dicionário com até 3 valores de cada coluna
        """
        if df.empty:
            return {}

        colunas_texto = []
        colunas_numericas = []
        colunas_mistas = []

        for col in df.columns:
            serie = df[col].dropna().astype(str)
            if serie.empty:
                continue

            # Verifica se parece numérico
            numericos = pd.to_numeric(
                serie.str.replace(".", "").str.replace(",", ".").str.replace("R$", ""),
                errors="coerce"
            ).notna().sum()

            proporcao_numerica = numericos / len(serie)

            if proporcao_numerica > 0.85:
                colunas_numericas.append(col)
            elif proporcao_numerica > 0.3:
                colunas_mistas.append(col)
            else:
                colunas_texto.append(col)

        amostras = {
            col: df[col].dropna().astype(str).head(3).tolist()
            for col in df.columns
        }

        return {
            "colunas_texto": colunas_texto,
            "colunas_numericas": colunas_numericas,
            "colunas_mistas": colunas_mistas,
            "amostras": amostras,
            "total_linhas": len(df),
            "total_colunas": len(df.columns)
        }