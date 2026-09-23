# ==============================================================================
# loader.py
# Sincronização de planilhas a partir de uma pasta pública do Google Drive.
#
# Inclui remoção automática do BOM (Byte Order Mark) dos arquivos CSV.
# O BOM em UTF-8 (bytes 0xEF 0xBB 0xBF) faz o DuckDB não reconhecer
# a linha de cabeçalho, causando exibição de dados no lugar de nomes de colunas.
# ==============================================================================

import os
import streamlit as st

DATA_DIR          = "data"
GDRIVE_FOLDER_KEY = "GDRIVE_FOLDER_URL"


def sincronizar_e_listar_csvs() -> list:
    """
    Sincroniza os arquivos CSV de uma pasta pública do Google Drive
    com a pasta local data/. Retorna lista de caminhos locais disponíveis.
    """
    os.makedirs(DATA_DIR, exist_ok=True)

    try:
        folder_url = st.secrets.get(GDRIVE_FOLDER_KEY, None)
    except Exception:
        folder_url = None

    if folder_url:
        _sincronizar_pasta(folder_url)
    else:
        st.sidebar.info(
            "Configure GDRIVE_FOLDER_URL em .streamlit/secrets.toml "
            "para sincronizar planilhas automaticamente.",
            icon="ℹ️"
        )

    # Remove BOM de todos os CSVs para garantir leitura correta pelo DuckDB
    _remover_bom_csvs(DATA_DIR)

    return _listar_csvs_locais()


def _sincronizar_pasta(folder_url: str):
    """
    Usa gdown para baixar arquivos da pasta pública do Google Drive.
    Arquivos já existentes localmente são preservados.
    """
    try:
        import gdown
    except ImportError:
        st.sidebar.warning("Pacote 'gdown' não instalado. Execute: pip install gdown")
        return

    csvs_antes = set(_listar_csvs_locais())

    try:
        with st.spinner("Sincronizando planilhas com o Google Drive..."):
            gdown.download_folder(
                url=folder_url,
                output=DATA_DIR,
                quiet=True,
                use_cookies=False
            )

        csvs_depois = set(_listar_csvs_locais())
        novos       = csvs_depois - csvs_antes

        if novos:
            nomes = [os.path.basename(p) for p in novos]
            st.sidebar.success(
                f"Sincronizado: {len(novos)} arquivo(s) novo(s) — {', '.join(nomes)}",
                icon="✅"
            )

    except Exception as e:
        if csvs_antes:
            st.sidebar.warning(
                f"Não foi possível sincronizar com o Drive. "
                f"Usando {len(csvs_antes)} arquivo(s) local(is).",
                icon="⚠️"
            )
        else:
            st.sidebar.error(f"Erro ao acessar a pasta do Google Drive: {e}", icon="🚫")


def _remover_bom_csvs(pasta: str):
    """
    Remove o BOM (Byte Order Mark) UTF-8 do início dos arquivos CSV.

    O BOM é uma sequência de 3 bytes (0xEF 0xBB 0xBF) que alguns programas
    como Excel inserem no início de arquivos CSV. O DuckDB não consegue
    ignorar o BOM ao ler o cabeçalho, e acaba tratando a linha de cabeçalho
    como dados — exibindo valores da primeira linha no lugar dos nomes das colunas.

    Essa função corrige o arquivo em disco uma única vez, se o BOM estiver presente.
    """
    BOM_UTF8 = b"\xef\xbb\xbf"

    for raiz, _, arquivos in os.walk(pasta):
        for nome in arquivos:
            if not nome.lower().endswith(".csv"):
                continue

            caminho = os.path.join(raiz, nome)
            try:
                with open(caminho, "rb") as f:
                    primeiros_bytes = f.read(3)

                if primeiros_bytes == BOM_UTF8:
                    with open(caminho, "rb") as f:
                        conteudo_completo = f.read()
                    with open(caminho, "wb") as f:
                        f.write(conteudo_completo[3:])  # Grava sem os 3 bytes do BOM
            except Exception:
                continue


def _listar_csvs_locais() -> list:
    """
    Lista todos os arquivos CSV presentes na pasta data/ e subpastas.
    Retorna lista de caminhos ordenada pelo nome do arquivo.
    """
    csvs = []
    if not os.path.exists(DATA_DIR):
        return csvs

    for raiz, _, arquivos in os.walk(DATA_DIR):
        for nome in arquivos:
            if nome.lower().endswith(".csv"):
                csvs.append(os.path.join(raiz, nome))

    return sorted(csvs, key=lambda p: os.path.basename(p).lower())


def nome_amigavel_csv(caminho: str) -> str:
    """
    Converte o nome do arquivo CSV em rótulo legível para o seletor.
    Exemplo: 'dados 2024.csv' -> 'dados 2024'
    """
    return os.path.splitext(os.path.basename(caminho))[0]