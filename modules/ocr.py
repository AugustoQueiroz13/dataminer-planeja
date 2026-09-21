# ==============================================================================
# ocr.py
# Leitura de documentos escaneados (PDFs como imagens e arquivos de imagem).
#
# Pipeline de pré-processamento por imagem:
#   1. Conversão para escala de cinza
#   2. Redimensionamento se resolução baixa
#   3. Aumento de contraste
#   4. Nitidez (unsharp mask)
#   5. Binarização (preto e branco limpo)
#   6. OCR com Tesseract em português brasileiro
#
# Formatos suportados:
#   PDF (todas as páginas processadas individualmente)
#   PNG, JPG, JPEG, TIFF, BMP
# ==============================================================================

import re
import io
import os
import platform
import pandas as pd
from typing import Optional

# Importações condicionais para dar mensagens de erro claras
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


# ==============================================================================
# Configuração automática do caminho do Tesseract no Windows
# No Windows o Tesseract não é adicionado ao PATH automaticamente.
# Este bloco localiza o executável nos caminhos de instalação mais comuns.
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
    Aplica pipeline de pré-processamento para melhorar a qualidade do OCR.
    Funciona tanto para documentos limpos quanto para documentos com ruído.
    """
    # Passo 1: Escala de cinza
    img = imagem.convert("L")

    # Passo 2: Redimensionar se a resolução for baixa
    largura, altura = img.size
    if largura < 1500:
        fator = 2.0
        img = img.resize(
            (int(largura * fator), int(altura * fator)),
            Image.LANCZOS
        )

    # Passo 3: Aumento de contraste
    img = ImageEnhance.Contrast(img).enhance(2.0)

    # Passo 4: Nitidez
    img = img.filter(ImageFilter.UnsharpMask(radius=1, percent=150, threshold=3))

    # Passo 5: Binarização (preto e branco limpo)
    img = img.point(lambda p: 255 if p > 160 else 0)

    return img


# ==============================================================================
# Extração de texto por OCR
# ==============================================================================

def extrair_texto_imagem(imagem: "Image.Image") -> str:
    """
    Executa o OCR em uma imagem e retorna o texto extraído.
    Usa configurações otimizadas para documentos com tabelas.
    """
    img_processada = preprocessar_imagem(imagem)

    # --oem 3: usa o melhor motor disponível (LSTM + legacy)
    # --psm 6: assume bloco de texto uniforme (bom para tabelas e formulários)
    # -l por: dicionário de português
    config = "--oem 3 --psm 6 -l por"

    return pytesseract.image_to_string(img_processada, config=config)


# ==============================================================================
# Tentativa de estruturar o texto em tabela
# ==============================================================================

def texto_para_dataframe(texto: str, numero_pagina: int) -> pd.DataFrame:
    """
    Tenta estruturar o texto extraído pelo OCR em um DataFrame.
    Quando o texto não tem estrutura de tabela clara, retorna
    o conteúdo bruto linha a linha.
    """
    linhas = [l.strip() for l in texto.split("\n") if l.strip()]

    if not linhas:
        return pd.DataFrame(columns=["pagina_origem", "linha", "texto_ocr"])

    # Verifica se o texto tem estrutura de tabela
    linhas_com_separadores = sum(
        1 for l in linhas if re.search(r"\s{2,}|\t|\|", l)
    )
    tem_estrutura = linhas_com_separadores > len(linhas) * 0.4

    if tem_estrutura:
        registros = []
        for linha in linhas:
            colunas = re.split(r"\s{2,}|\t|\s*\|\s*", linha)
            colunas = [c.strip() for c in colunas if c.strip()]
            registros.append(colunas)

        num_max = max(len(r) for r in registros)
        registros_norm = [r + [""] * (num_max - len(r)) for r in registros]

        if len(registros_norm) > 1:
            cabecalho = registros_norm[0]
            cab_unico = []
            contagem = {}
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
    Lê um PDF escaneado página por página.
    Retorna: (lista de DataFrames, total de páginas).
    """
    ok, msg = verificar_dependencias()
    if not ok:
        raise RuntimeError(msg)

    imagens = convert_from_bytes(arquivo_bytes, dpi=300)
    total_paginas = len(imagens)
    dataframes = []

    for num_pag, imagem in enumerate(imagens, start=1):
        texto = extrair_texto_imagem(imagem)
        df = texto_para_dataframe(texto, num_pag)
        if not df.empty:
            dataframes.append(df)

    return dataframes, total_paginas


def ler_imagem(arquivo_bytes: bytes, nome_arquivo: str) -> pd.DataFrame:
    """
    Lê um único arquivo de imagem (PNG, JPG, TIFF, BMP) via OCR.
    """
    ok, msg = verificar_dependencias()
    if not ok:
        raise RuntimeError(msg)

    imagem = Image.open(io.BytesIO(arquivo_bytes))
    texto = extrair_texto_imagem(imagem)
    df = texto_para_dataframe(texto, numero_pagina=1)
    df.insert(1, "arquivo_origem", nome_arquivo)
    return df


def consolidar_paginas(dataframes: list) -> pd.DataFrame:
    """
    Consolida múltiplos DataFrames de páginas num único DataFrame.
    """
    if not dataframes:
        return pd.DataFrame()
    return pd.concat(dataframes, ignore_index=True)