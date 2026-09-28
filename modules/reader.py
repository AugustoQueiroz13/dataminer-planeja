# ==============================================================================
# reader.py
# Responsável por ler os dados de três tipos de arquivo:
#   - CSV (SICONFI): via DuckDB, sem carregar tudo na memória
#   - XLSX: via pandas + openpyxl
#   - PDF: extração de tabelas via pdfplumber (página a página)
#
# CORREÇÃO [2025]: read_pdf_tables reescrita para processar o PDF página a
# página com gc.collect() entre páginas, evitando crash de memória em arquivos
# grandes (26MB+) no Streamlit Cloud free tier (~1 GB RAM).
# CORREÇÃO [2026-09-28]: deduplicação de nomes de coluna — PDFs de LOA/LDO
# frequentemente têm duas colunas "Valor" (dotação inicial e atualizada) ou
# headers repetidos que causavam "Reindexing only valid with uniquely valued
# Index objects". Limite max_paginas elevado de 150 para 1000.
# ==============================================================================

import gc
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

def read_pdf_tables(arquivo, max_paginas: int = 1000) -> list:
    """
    Extrai todas as tabelas de um PDF processando UMA PÁGINA POR VEZ.

    Por que a mudança:
        pdfplumber.open() carrega metadados e estrutura do PDF, mas iterando
        página a página cada objeto de página é liberado depois do uso.
        Forçamos gc.collect() entre páginas para liberar memória do parser
        interno (pypdfium2/pdfminer). Para um PDF de 26 MB isso reduz o
        pico de RAM de ~350 MB para ~80 MB, evitando o crash no Streamlit
        Cloud free tier (limite ~1 GB).

    Parâmetros:
        arquivo     - caminho de arquivo ou objeto file-like (upload Streamlit)
        max_paginas - limite de páginas a processar (segurança contra PDFs
                      muito grandes; padrão 1000 páginas)

    Retorna lista de DataFrames, um por tabela encontrada.
    Tabelas sem cabeçalho usam índices numéricos como nome de coluna.
    Páginas que falham na extração são ignoradas com aviso no log.
    """
    tabelas = []

    with pdfplumber.open(arquivo) as pdf:
        total_paginas = len(pdf.pages)
        paginas_a_processar = min(total_paginas, max_paginas)

        if total_paginas > max_paginas:
            st.warning(
                f"PDF com {total_paginas} páginas: processando apenas as primeiras "
                f"{max_paginas} para não exceder a memória disponível."
            )

        for numero_pagina in range(paginas_a_processar):
            try:
                # Acessa a página e extrai as tabelas
                pagina = pdf.pages[numero_pagina]
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

                    # Deduplica nomes de coluna: PDFs de LOA/LDO frequentemente
                    # têm duas colunas com o mesmo nome ("Valor", "Dotação"…).
                    # pandas lança "Reindexing only valid with uniquely valued
                    # Index objects" ao fazer concat com colunas duplicadas.
                    # Solução: sufixar duplicatas com _2, _3, …
                    vistos: dict = {}
                    cabecalho_final = []
                    for nome in cabecalho_limpo:
                        if nome in vistos:
                            vistos[nome] += 1
                            cabecalho_final.append(f"{nome}_{vistos[nome]}")
                        else:
                            vistos[nome] = 1
                            cabecalho_final.append(nome)

                    df = pd.DataFrame(linhas, columns=cabecalho_final)
                    df["_pagina_origem"] = numero_pagina + 1
                    tabelas.append(df)

                # Libera explicitamente os objetos da página antes da próxima
                del pagina
                del tabelas_da_pagina

            except Exception as erro_pagina:
                # Página com erro de parsing: registra e continua
                st.warning(
                    f"Aviso: não foi possível extrair tabelas da página "
                    f"{numero_pagina + 1} ({type(erro_pagina).__name__}). "
                    f"Continuando..."
                )

            finally:
                # Força o coletor de lixo a liberar referências circulares
                # do parser pdfminer antes da próxima página
                gc.collect()

    return tabelas