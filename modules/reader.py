# ==============================================================================
# reader.py
# Leitura de dados de CSV (DuckDB), XLSX (pandas) e PDF (pdfplumber).
#
# Correções incluídas:
#   - utf-8-sig como primeiro encoding (remove BOM e corrige detecção de header)
#   - Município usa busca LIKE em vez de IN exato
#     (necessário pois SICONFI usa "Prefeitura Municipal de X" e não apenas "X")
# ==============================================================================

import duckdb
import pandas as pd
import pdfplumber
import streamlit as st
from typing import Optional

# utf-8-sig PRIMEIRO: remove o BOM de UTF-8, garantindo que a linha de
# cabeçalho seja lida corretamente pelo DuckDB.
ENCODINGS_CSV = ["utf-8-sig", "utf-8", "latin-1", "cp1252", "utf-16"]


# ==============================================================================
# Leitura de colunas do CSV
# ==============================================================================

@st.cache_data(show_spinner=False)
def get_csv_columns(csv_path: str) -> list:
    """
    Lê apenas o cabeçalho do CSV sem carregar os dados.
    Tenta múltiplos encodings automaticamente.
    utf-8-sig é tentado primeiro para remover BOM de arquivos do governo.
    """
    con = duckdb.connect()
    caminho = csv_path.replace("'", "''")

    for enc in ENCODINGS_CSV:
        try:
            resultado = con.execute(
                f"DESCRIBE SELECT * FROM read_csv_auto("
                f"'{caminho}', header=true, encoding='{enc}', "
                f"ignore_errors=true, sample_size=500"
                f")"
            ).fetchdf()
            colunas = resultado["column_name"].tolist()
            con.close()
            return colunas
        except Exception:
            continue

    con.close()
    return []


# ==============================================================================
# Leitura de valores únicos de uma coluna
# ==============================================================================

@st.cache_data(show_spinner=False)
def get_valores_unicos(csv_path: str, nome_coluna: str) -> list:
    """
    Retorna os valores únicos de uma coluna sem carregar o arquivo completo.
    Útil para preencher filtros de ano, UF e tipo de receita.
    """
    con = duckdb.connect()
    caminho = csv_path.replace("'", "''")
    col     = nome_coluna.replace("'", "''")

    for enc in ENCODINGS_CSV:
        try:
            resultado = con.execute(
                f'SELECT DISTINCT "{col}" '
                f"FROM read_csv_auto('{caminho}', header=true, "
                f"encoding='{enc}', ignore_errors=true) "
                f'WHERE "{col}" IS NOT NULL '
                f'ORDER BY "{col}"'
            ).fetchdf()
            con.close()
            return resultado.iloc[:, 0].tolist()
        except Exception:
            continue

    con.close()
    return []


# ==============================================================================
# Query principal no CSV via DuckDB
# ==============================================================================

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

    Busca de município usa LIKE em vez de IN exato, porque arquivos do SICONFI
    registram os municípios como "Prefeitura Municipal de X" e não apenas "X".
    A busca LIKE %guapimirim% encontra "Prefeitura Municipal de Guapimirim".
    """
    con     = duckdb.connect()
    caminho = csv_path.replace("'", "''")

    # Filtro de município com LIKE (busca parcial, case-insensitive)
    # Cobre tanto "Guapimirim" quanto "Prefeitura Municipal de Guapimirim"
    like_parts = " OR ".join([
        f'LOWER("{col_municipio}") LIKE \'%{m.lower().replace(chr(39), chr(39)*2)}%\''
        for m in municipios
    ])
    clausulas = [f"({like_parts})"]

    # Filtro por tipo de receita (busca parcial no campo de conta/código)
    if termo_receita and termo_receita.strip():
        termo = termo_receita.replace("'", "''")
        clausulas.append(
            f'(LOWER("{col_receita}") LIKE \'%{termo.lower()}%\')'
        )

    # Filtro por ano
    if col_ano and anos:
        lista_anos = ", ".join([f"'{str(a)}'" for a in anos])
        clausulas.append(f'CAST("{col_ano}" AS VARCHAR) IN ({lista_anos})')

    where_sql = " AND ".join(clausulas)

    for enc in ENCODINGS_CSV:
        try:
            query = (
                f"SELECT * FROM read_csv_auto("
                f"'{caminho}', header=true, encoding='{enc}', ignore_errors=true"
                f") WHERE {where_sql}"
            )
            resultado = con.execute(query).fetchdf()
            con.close()
            return resultado
        except Exception:
            continue

    con.close()
    return pd.DataFrame()


# ==============================================================================
# Leitura de XLSX
# ==============================================================================

def read_xlsx(arquivo) -> pd.DataFrame:
    """
    Lê um arquivo Excel e retorna DataFrame pandas.
    Aceita caminho de arquivo ou objeto file-like (upload Streamlit).
    """
    return pd.read_excel(arquivo, engine="openpyxl")


# ==============================================================================
# Leitura de PDF com pdfplumber
# ==============================================================================

def read_pdf_tables(arquivo) -> list:
    """
    Extrai todas as tabelas de um PDF com texto selecionável.
    Retorna lista de DataFrames, um por tabela encontrada.
    Para PDFs escaneados, use modules/ocr.py.
    """
    tabelas = []
    with pdfplumber.open(arquivo) as pdf:
        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            tabelas_da_pagina = pagina.extract_tables()
            for tabela in tabelas_da_pagina:
                if not tabela:
                    continue
                cabecalho = [
                    str(c) if c is not None else f"Col_{i}"
                    for i, c in enumerate(tabela[0])
                ]
                df = pd.DataFrame(tabela[1:], columns=cabecalho)
                df["_pagina_origem"] = numero_pagina
                tabelas.append(df)
    return tabelas