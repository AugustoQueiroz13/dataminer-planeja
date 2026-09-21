# ==============================================================================
# loader.py
# Responsável por encontrar o CSV local ou baixá-lo do Google Drive.
# O arquivo fica em cache em disco para não baixar toda vez.
#
# IMPORTANTE: esta função retorna None quando o arquivo não está disponível.
# O app.py verifica esse retorno antes de tentar qualquer consulta.
# Nunca chama st.stop() para não interromper o restante do app.
# ==============================================================================

import os
import streamlit as st

DATA_DIR    = "data"
CSV_FILENAME = "siconfi.csv"
CSV_PATH    = os.path.join(DATA_DIR, CSV_FILENAME)


def get_csv_path():
    """
    Verifica se o CSV já existe localmente em data/siconfi.csv.
    Se não encontrar, tenta baixar do Google Drive usando a URL
    configurada em .streamlit/secrets.toml (chave: GDRIVE_CSV_URL).

    Retorna o caminho local do arquivo CSV, ou None se não disponível.
    Nunca interrompe o app — o chamador decide o que fazer com None.
    """
    os.makedirs(DATA_DIR, exist_ok=True)

    # Arquivo já existe localmente
    if os.path.exists(CSV_PATH):
        tamanho_mb = os.path.getsize(CSV_PATH) / (1024 * 1024)
        st.sidebar.caption(f"SICONFI em disco ({tamanho_mb:.0f} MB)")
        return CSV_PATH

    # Verifica se há URL configurada
    try:
        gdrive_url = st.secrets.get("GDRIVE_CSV_URL", None)
    except Exception:
        gdrive_url = None

    if not gdrive_url:
        st.sidebar.info(
            "SICONFI não encontrado. Configure GDRIVE_CSV_URL "
            "em .streamlit/secrets.toml para habilitá-lo.",
            icon="ℹ️"
        )
        return None

    # Tenta baixar do Google Drive
    try:
        import gdown
    except ImportError:
        st.sidebar.warning("Pacote 'gdown' não instalado. Execute: pip install gdown")
        return None

    with st.spinner("Baixando SICONFI do Google Drive (primeira vez pode demorar)..."):
        try:
            gdown.download(gdrive_url, CSV_PATH, quiet=False, fuzzy=True)
        except Exception as erro:
            st.sidebar.warning(f"Falha no download do SICONFI: {erro}")
            return None

    if not os.path.exists(CSV_PATH):
        st.sidebar.warning(
            "Download do SICONFI falhou. Verifique se o link do Google Drive "
            "está configurado como 'qualquer pessoa com o link pode ver'."
        )
        return None

    st.sidebar.success("SICONFI baixado e salvo localmente.")
    return CSV_PATH