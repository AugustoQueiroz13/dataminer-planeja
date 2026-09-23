# ==============================================================================
# reader.py
# Leitura de dados de CSV (DuckDB), XLSX (pandas) e PDF (pdfplumber).
#
# Correção principal: get_csv_columns usa Python puro (não DuckDB) para
# ler o cabeçalho do arquivo. O Python com encoding='utf-8-sig' remove
# o BOM automaticamente, garantindo que os nomes das colunas sejam
# detectados corretamente (ex: Instituição, Conta, Valor).
# ==============================================================================

import csv
import duckdb
import pandas as pd
import pdfplumber
import streamlit as st
from typing import Optional

ENCODINGS_CSV  = ["utf-8-sig", "utf-8", "latin-1", "cp1252"]
SEPARADORES    = [",", ";", "\t"]


# ==============================================================================
# Leitura do cabeçalho do CSV via Python puro
# ==============================================================================

@st.cache_data(show_spinner=False)
def get_csv_columns(csv_path: str) -> list:
    """
    Lê os nomes das colunas do CSV usando Python puro (csv.reader).
    Usa utf-8-sig como primeiro encoding para remover o BOM automaticamente.

    Este método é mais confiável que o DuckDB para detectar o cabeçalho,
    pois o Python com utf-8-sig trata o BOM corretamente, evitando que
    valores de dados apareçam como nomes de colunas.

    Tenta combinações de encoding e separador até encontrar uma leitura
    válida (mais de 2 colunas no cabeçalho).
    """
    for enc in ENCODINGS_CSV:
        for sep in SEPARADORES:
            try:
                with open(csv_path, "r", encoding=enc, newline="") as f:
                    leitor = csv.reader(f, delimiter=sep)
                    primeira_linha = next(leitor)

                colunas = [c.strip().strip('"') for c in primeira_linha if c.strip()]
                if len(colunas) > 2:
                    return colunas

            except Exception:
                continue

    # Fallback: DuckDB (caso a leitura Python falhe)
    con = duckdb.connect()
    caminho = csv_path.replace("'", "''")
    for enc in ENCODINGS_CSV:
        try:
            resultado = con.execute(
                f"DESCRIBE SELECT * FROM read_csv_auto('{caminho}', header=true, "
                f"encoding='{enc}', ignore_errors=true, sample_size=100)"
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
    con     = duckdb.connect()
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

    Município: busca LIKE parcial (ex: 'guapimirim' encontra
    'Prefeitura Municipal de Guapimirim').

    Receita/Código: busca LIKE parcial no campo selecionado
    (ex: '1.7' encontra '1.7.2.2.52.0.0 - Cota-parte Royalties').
    """
    con     = duckdb.connect()
    caminho = csv_path.replace("'", "''")

    # Filtro de município: LIKE parcial para cobrir "Prefeitura Municipal de X"
    like_mun = " OR ".join([
        f"LOWER(\"{col_municipio}\") LIKE '%{m.lower().replace(chr(39), chr(39)*2)}%'"
        for m in municipios
    ])
    clausulas = [f"({like_mun})"]

    # Filtro de receita/código: LIKE parcial
    if termo_receita and termo_receita.strip():
        termo = termo_receita.strip().replace("'", "''")
        clausulas.append(
            f"LOWER(\"{col_receita}\") LIKE '%{termo.lower()}%'"
        )

    # Filtro de ano
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
    """Lê um arquivo Excel e retorna DataFrame pandas."""
    return pd.read_excel(arquivo, engine="openpyxl")


# ==============================================================================
# Leitura de PDF com pdfplumber
# ==============================================================================

def read_pdf_tables(arquivo) -> list:
    """
    Extrai tabelas de um PDF com texto selecionável.
    Para PDFs escaneados, use modules/ocr.py.
    """
    tabelas = []
    with pdfplumber.open(arquivo) as pdf:
        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            for tabela in pagina.extract_tables() or []:
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