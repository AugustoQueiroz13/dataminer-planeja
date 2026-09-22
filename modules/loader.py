# ==============================================================================
# loader.py
# Sincronização de planilhas a partir de uma pasta pública do Google Drive.
#
# Comportamento:
#   - Na inicialização do app, sincroniza automaticamente com a pasta do Drive
#   - Arquivos já existentes localmente são ignorados (não baixados novamente)
#   - Novos arquivos adicionados à pasta aparecem na próxima abertura do app
#   - Se a sincronização falhar, usa os arquivos já disponíveis localmente
#
# Configuração necessária em .streamlit/secrets.toml:
#   GDRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/SEU_ID_AQUI"
# ==============================================================================

import os
import streamlit as st

DATA_DIR         = "data"
GDRIVE_FOLDER_KEY = "GDRIVE_FOLDER_URL"


def sincronizar_e_listar_csvs() -> list:
    """
    Sincroniza os arquivos CSV de uma pasta pública do Google Drive
    com a pasta local data/. Retorna a lista de caminhos locais
    de todos os CSVs disponíveis.

    Na primeira execução, faz o download dos arquivos.
    Nas execuções seguintes, apenas verifica se há arquivos novos.
    Se a sincronização falhar, usa os arquivos já presentes localmente.
    """
    os.makedirs(DATA_DIR, exist_ok=True)

    # Lê a URL da pasta configurada em secrets.toml
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

    return _listar_csvs_locais()


def _sincronizar_pasta(folder_url: str):
    """
    Usa gdown para baixar arquivos da pasta pública do Google Drive.
    Arquivos já existentes localmente são preservados e não baixados novamente.
    """
    try:
        import gdown
    except ImportError:
        st.sidebar.warning(
            "Pacote 'gdown' não instalado. Execute: pip install gdown"
        )
        return

    csvs_antes = set(_listar_csvs_locais())

    try:
        with st.spinner("Sincronizando planilhas com o Google Drive..."):
            gdown.download_folder(
                url=folder_url,
                output=DATA_DIR,
                quiet=True,
                use_cookies=False,
            )
           
        csvs_depois = set(_listar_csvs_locais())
        novos = csvs_depois - csvs_antes

        if novos:
            nomes = [os.path.basename(p) for p in novos]
            st.sidebar.success(
                f"Sincronizado: {len(novos)} arquivo(s) novo(s) — "
                f"{', '.join(nomes)}",
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
            st.sidebar.error(
                f"Erro ao acessar a pasta do Google Drive: {e}",
                icon="🚫"
            )


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
    Converte o nome do arquivo CSV em um rótulo legível para o seletor.
    Exemplo: 'siconfi_receitas_2024.csv' -> 'siconfi_receitas_2024'
    """
    return os.path.splitext(os.path.basename(caminho))[0]