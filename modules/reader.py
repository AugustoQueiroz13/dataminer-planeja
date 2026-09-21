# ==============================================================================
# reader.py
# Responsável por ler os dados de três tipos de arquivo:
#   - CSV (SICONFI): via DuckDB, sem carregar tudo na memória
#   - XLSX: via pandas + openpyxl
#   - PDF: extração de tabelas via pdfplumber
# ==============================================================================

import duckdb
import pandas as pd
import pdfplumber
import streamlit as st
from typing import Optional


# ==================== CSV via DuckDB ==========================================

@st.cache_data(show_spinner=False)
def get_csv_columns(csv_path: str) -> list:
    """
    Lê apenas o cabeçalho do CSV sem carregar os dados.
    Retorna lista com os nomes de todas as colunas.
    """
    con = duckdb.connect()
    resultado = con.execute(
        f"DESCRIBE SELECT * FROM read_csv_auto('{csv_path}', header=true, sample_size=100)"
    ).fetchdf()
    con.close()
    return resultado["column_name"].tolist()


@st.cache_data(show_spinner=False)
def get_valores_unicos(csv_path: str, nome_coluna: str) -> list:
    """
    Retorna os valores únicos de uma coluna sem carregar o arquivo completo.
    Útil para preencher filtros de ano, UF e tipo de receita.
    """
    con = duckdb.connect()
    resultado = con.execute(
        f'SELECT DISTINCT "{nome_coluna}" '
        f"FROM read_csv_auto('{csv_path}', header=true) "
        f'WHERE "{nome_coluna}" IS NOT NULL '
        f'ORDER BY "{nome_coluna}"'
    ).fetchdf()
    con.close()
    return resultado.iloc[:, 0].tolist()


def query_csv(
    csv_path: str,
    municipios: list,
    col_municipio: str,
    col_receita: str,
    termo_receita: Optional[str] = None,
    col_ano: Optional[str] = None,
    anos: Optional[list] = None
) -> pd.DataFrame:
    """
    Consulta o CSV com DuckDB usando filtros dinâmicos.
    Retorna apenas as linhas que correspondem aos filtros,
    sem nunca carregar o arquivo inteiro na memória.
    """
    con = duckdb.connect()

    # Monta lista de municípios para cláusula IN
    lista_municipios = ", ".join([f"'{m.replace(chr(39), chr(39)*2)}'" for m in municipios])
    clausulas = [f'"{col_municipio}" IN ({lista_municipios})']

    # Filtro por tipo de receita (busca parcial, sem distinção de maiúsculas)
    if termo_receita and termo_receita.strip():
        termo_seguro = termo_receita.replace("'", "''")
        clausulas.append(f'LOWER("{col_receita}") LIKE \'%{termo_seguro.lower()}%\'')

    # Filtro por ano
    if col_ano and anos:
        lista_anos = ", ".join([f"'{str(a)}'" for a in anos])
        clausulas.append(f'CAST("{col_ano}" AS VARCHAR) IN ({lista_anos})')

    where_sql = " AND ".join(clausulas)

    query = f"""
        SELECT *
        FROM read_csv_auto('{csv_path}', header=true)
        WHERE {where_sql}
    """

    resultado = con.execute(query).fetchdf()
    con.close()
    return resultado


# ==================== XLSX via pandas ========================================

def read_xlsx(arquivo) -> pd.DataFrame:
    """
    Lê um arquivo Excel e retorna um DataFrame pandas.
    Aceita caminho de arquivo ou objeto file-like (upload do Streamlit).
    """
    return pd.read_excel(arquivo, engine="openpyxl")


# ==================== PDF via pdfplumber =====================================

def read_pdf_tables(arquivo) -> list:
    """
    Extrai todas as tabelas de um PDF.
    Retorna lista de DataFrames, um por tabela encontrada em qualquer página.
    Tabelas sem cabeçalho usam índices numéricos como nome de coluna.
    """
    tabelas = []
    with pdfplumber.open(arquivo) as pdf:
        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            tabelas_da_pagina = pagina.extract_tables()
            for tabela in tabelas_da_pagina:
                if not tabela:
                    continue
                cabecalho = tabela[0]
                linhas = tabela[1:]
                # Garante que o cabeçalho não tenha valores None
                cabecalho_limpo = [
                    str(c) if c is not None else f"Col_{i}"
                    for i, c in enumerate(cabecalho)
                ]
                df = pd.DataFrame(linhas, columns=cabecalho_limpo)
                df["_pagina_origem"] = numero_pagina
                tabelas.append(df)
    return tabelas