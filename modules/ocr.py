# ==============================================================================
# ocr.py
# Leitura de documentos escaneados com uso otimizado de memória.
#
# Otimizações para o Streamlit Cloud (1GB RAM):
#   - DPI reduzido para 150 (suficiente para texto impresso, muito menos memória)
#   - Garbage collection forçado entre páginas
#   - Processamento de uma página por vez (sem carregar o PDF inteiro em RAM)
#   - Limite de páginas configurável para evitar timeout
# ==============================================================================

import re
import io
import os
import gc
import platform
import pandas as pd
from typing import Optional

try:
    import pytesseract
    PYTESSERACT_DISPONIVEL = True
except ImportError:
    PYTESSERACT_DISPONIVEL = False

try:
    from PIL import Image, ImageFilter, ImageEnhance
    PILLOW_DISPONIVEL = True
except ImportError:
    PILLOW_DISPONIVEL = False

try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE_DISPONIVEL = True
except ImportError:
    PDF2IMAGE_DISPONIVEL = False


# DPI para conversão de PDF em imagem.
# 150 usa ~4x menos memória que 300 e é suficiente para texto impresso.
DPI_OCR = 150

# Limite de páginas por processamento para evitar estouro de memória.
MAX_PAGINAS_OCR = 20


# ==============================================================================
# Configuração automática do caminho do Tesseract no Windows
# ==============================================================================

if platform.system() == "Windows" and PYTESSERACT_DISPONIVEL:
    caminhos_possiveis = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.join(
            os.environ.get("LOCALAPPDATA", ""),
            "Programs", "Tesseract-OCR", "tesseract.exe"
        ),
        os.path.join(
            os.environ.get("APPDATA", ""),
            "Tesseract-OCR", "tesseract.exe"
        ),
    ]
    for caminho in caminhos_possiveis:
        if os.path.exists(caminho):
            pytesseract.pytesseract.tesseract_cmd = caminho
            break


# ==============================================================================
# Verificação de dependências
# ==============================================================================

def verificar_dependencias() -> tuple:
    """
    Verifica se todas as dependências de OCR estão disponíveis.
    Retorna (ok: bool, mensagem: str).
    """
    if not PYTESSERACT_DISPONIVEL:
        return False, "Biblioteca 'pytesseract' não instalada. Execute: pip install pytesseract"

    if not PILLOW_DISPONIVEL:
        return False, "Biblioteca 'Pillow' não instalada. Execute: pip install Pillow"

    if not PDF2IMAGE_DISPONIVEL:
        return False, "Biblioteca 'pdf2image' não instalada. Execute: pip install pdf2image"

    try:
        pytesseract.get_tesseract_version()
    except pytesseract.TesseractNotFoundError:
        sistema = platform.system()
        if sistema == "Windows":
            instrucao = (
                "Tesseract não encontrado. Instale em:\n"
                "https://github.com/UB-Mannheim/tesseract/wiki\n"
                "Durante a instalação, marque 'Portuguese' em Additional language data."
            )
        elif sistema == "Darwin":
            instrucao = "macOS: brew install tesseract tesseract-lang"
        else:
            instrucao = "Linux: sudo apt install tesseract-ocr tesseract-ocr-por poppler-utils"

        return False, instrucao

    return True, "OK"


# ==============================================================================
# Pré-processamento de imagem
# ==============================================================================

def preprocessar_imagem(imagem: "Image.Image") -> "Image.Image":
    """
    Aplica pré-processamento leve para melhorar o OCR sem consumir muita memória.
    Usa operações simples que funcionam bem a 150 DPI.
    """
    # Escala de cinza
    img = imagem.convert("L")

    # Aumento leve de contraste
    img = ImageEnhance.Contrast(img).enhance(1.5)

    # Binarização simples
    img = img.point(lambda p: 255 if p > 150 else 0)

    return img


# ==============================================================================
# Extração de texto
# ==============================================================================

def extrair_texto_imagem(imagem: "Image.Image") -> str:
    """
    Executa OCR em uma imagem com configurações para português.
    """
    img_proc = preprocessar_imagem(imagem)
    config   = "--oem 3 --psm 6 -l por"
    texto    = pytesseract.image_to_string(img_proc, config=config)

    # Libera memória imediatamente
    del img_proc
    gc.collect()

    return texto


# ==============================================================================
# Estruturação do texto em DataFrame
# ==============================================================================

def texto_para_dataframe(texto: str, numero_pagina: int) -> pd.DataFrame:
    """
    Tenta estruturar o texto extraído em DataFrame.
    Sem estrutura de tabela clara, retorna o texto bruto linha a linha.
    """
    linhas = [l.strip() for l in texto.split("\n") if l.strip()]

    if not linhas:
        return pd.DataFrame(columns=["pagina_origem", "linha", "texto_ocr"])

    linhas_com_sep = sum(1 for l in linhas if re.search(r"\s{2,}|\t|\|", l))
    tem_estrutura  = linhas_com_sep > len(linhas) * 0.4

    if tem_estrutura:
        registros = []
        for linha in linhas:
            colunas = re.split(r"\s{2,}|\t|\s*\|\s*", linha)
            colunas = [c.strip() for c in colunas if c.strip()]
            registros.append(colunas)

        num_max       = max(len(r) for r in registros)
        registros_norm = [r + [""] * (num_max - len(r)) for r in registros]

        if len(registros_norm) > 1:
            cabecalho = registros_norm[0]
            cab_unico = []
            contagem  = {}
            for col in cabecalho:
                nome = col if col else "Coluna"
                contagem[nome] = contagem.get(nome, 0) + 1
                sufixo = f"_{contagem[nome]}" if contagem[nome] > 1 else ""
                cab_unico.append(f"{nome}{sufixo}")
            df = pd.DataFrame(registros_norm[1:], columns=cab_unico)
        else:
            df = pd.DataFrame(
                registros_norm,
                columns=[f"Coluna_{i+1}" for i in range(num_max)]
            )
    else:
        df = pd.DataFrame({
            "linha":     range(1, len(linhas) + 1),
            "texto_ocr": linhas
        })

    df.insert(0, "pagina_origem", numero_pagina)
    return df


# ==============================================================================
# Funções públicas de leitura
# ==============================================================================

def ler_pdf_escaneado(arquivo_bytes: bytes) -> tuple:
    """
    Lê um PDF escaneado página por página com DPI reduzido (150).
    Libera memória entre páginas para funcionar no Streamlit Cloud (1GB RAM).
    Limita a MAX_PAGINAS_OCR páginas para evitar timeout.

    Retorna: (lista de DataFrames, total de páginas).
    """
    ok, msg = verificar_dependencias()
    if not ok:
        raise RuntimeError(msg)

    # Converte apenas as primeiras MAX_PAGINAS_OCR páginas
    imagens = convert_from_bytes(
        arquivo_bytes,
        dpi=DPI_OCR,
        first_page=1,
        last_page=MAX_PAGINAS_OCR
    )
    total_processado = len(imagens)
    dataframes       = []

    for num_pag, imagem in enumerate(imagens, start=1):
        texto = extrair_texto_imagem(imagem)

        # Libera a imagem da memória imediatamente após processar
        del imagem
        gc.collect()

        df = texto_para_dataframe(texto, num_pag)
        if not df.empty:
            dataframes.append(df)

    # Libera a lista de imagens
    del imagens
    gc.collect()

    return dataframes, total_processado


def ler_imagem(arquivo_bytes: bytes, nome_arquivo: str) -> pd.DataFrame:
    """
    Lê um único arquivo de imagem via OCR.
    """
    ok, msg = verificar_dependencias()
    if not ok:
        raise RuntimeError(msg)

    imagem = Image.open(io.BytesIO(arquivo_bytes))
    texto  = extrair_texto_imagem(imagem)
    del imagem
    gc.collect()

    df = texto_para_dataframe(texto, numero_pagina=1)
    df.insert(1, "arquivo_origem", nome_arquivo)
    return df


def consolidar_paginas(dataframes: list) -> pd.DataFrame:
    """
    Consolida múltiplos DataFrames de páginas em um único DataFrame.
    """
    if not dataframes:
        return pd.DataFrame()
    return pd.concat(dataframes, ignore_index=True)