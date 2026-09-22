# ==============================================================================
# ocr.py
# Leitura de documentos escaneados com processamento página por página.
#
# Estratégia de memória para Streamlit Cloud (1GB RAM):
#   - Uma página por vez: nunca carrega o PDF inteiro em memória
#   - DPI 150: boa qualidade para texto, usa 4x menos RAM que 300 DPI
#   - gc.collect() forçado após cada página
#   - Sem limite de páginas: documentos de qualquer tamanho são suportados
# ==============================================================================

import re
import io
import os
import gc
import platform
import pandas as pd

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
    from pdf2image import convert_from_bytes, pdfinfo_from_bytes
    PDF2IMAGE_DISPONIVEL = True
except ImportError:
    PDF2IMAGE_DISPONIVEL = False

# 150 DPI: bom para texto impresso, ~4x menos memória que 300 DPI
DPI_OCR = 150


# ==============================================================================
# Configuração automática do Tesseract no Windows
# ==============================================================================

if platform.system() == "Windows" and PYTESSERACT_DISPONIVEL:
    caminhos_possiveis = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.join(
            os.environ.get("LOCALAPPDATA", ""),
            "Programs", "Tesseract-OCR", "tesseract.exe"
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
    """Verifica se todas as dependências de OCR estão disponíveis."""
    if not PYTESSERACT_DISPONIVEL:
        return False, "Biblioteca 'pytesseract' não instalada."
    if not PILLOW_DISPONIVEL:
        return False, "Biblioteca 'Pillow' não instalada."
    if not PDF2IMAGE_DISPONIVEL:
        return False, "Biblioteca 'pdf2image' não instalada."
    try:
        pytesseract.get_tesseract_version()
    except pytesseract.TesseractNotFoundError:
        sistema = platform.system()
        if sistema == "Windows":
            return False, (
                "Tesseract não encontrado. Instale em:\n"
                "https://github.com/UB-Mannheim/tesseract/wiki\n"
                "Marque 'Portuguese' em Additional language data."
            )
        elif sistema == "Darwin":
            return False, "macOS: brew install tesseract tesseract-lang"
        else:
            return False, "Linux: sudo apt install tesseract-ocr tesseract-ocr-por"
    return True, "OK"


# ==============================================================================
# Pré-processamento de imagem
# ==============================================================================

def preprocessar_imagem(imagem: "Image.Image") -> "Image.Image":
    """
    Pré-processamento leve otimizado para 150 DPI.
    Converte para cinza, aumenta contraste e binariza.
    """
    img = imagem.convert("L")
    img = ImageEnhance.Contrast(img).enhance(1.5)
    img = img.point(lambda p: 255 if p > 150 else 0)
    return img


# ==============================================================================
# Extração de texto
# ==============================================================================

def extrair_texto_imagem(imagem: "Image.Image") -> str:
    """Executa OCR na imagem e libera memória em seguida."""
    img_proc = preprocessar_imagem(imagem)
    texto    = pytesseract.image_to_string(img_proc, config="--oem 3 --psm 6 -l por")
    del img_proc
    gc.collect()
    return texto


# ==============================================================================
# Estruturação do texto em DataFrame
# ==============================================================================

def texto_para_dataframe(texto: str, numero_pagina: int) -> pd.DataFrame:
    """
    Tenta estruturar o texto em tabela.
    Sem estrutura clara, retorna o texto bruto linha a linha.
    """
    linhas = [l.strip() for l in texto.split("\n") if l.strip()]

    if not linhas:
        return pd.DataFrame(columns=["pagina_origem", "linha", "texto_ocr"])

    linhas_com_sep = sum(1 for l in linhas if re.search(r"\s{2,}|\t|\|", l))
    tem_estrutura  = linhas_com_sep > len(linhas) * 0.4

    if tem_estrutura:
        registros = []
        for linha in linhas:
            cols = re.split(r"\s{2,}|\t|\s*\|\s*", linha)
            cols = [c.strip() for c in cols if c.strip()]
            registros.append(cols)

        num_max       = max(len(r) for r in registros)
        registros_norm = [r + [""] * (num_max - len(r)) for r in registros]

        if len(registros_norm) > 1:
            cab   = registros_norm[0]
            unico = []
            cont  = {}
            for c in cab:
                nome = c if c else "Coluna"
                cont[nome] = cont.get(nome, 0) + 1
                unico.append(f"{nome}_{cont[nome]}" if cont[nome] > 1 else nome)
            df = pd.DataFrame(registros_norm[1:], columns=unico)
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
    Lê um PDF escaneado processando UMA PÁGINA POR VEZ.
    Sem limite de páginas. Suporta documentos de qualquer tamanho.
    Libera memória entre cada página para funcionar no Streamlit Cloud.

    Retorna: (lista de DataFrames, total de páginas processadas).
    """
    ok, msg = verificar_dependencias()
    if not ok:
        raise RuntimeError(msg)

    # Obtém o total de páginas sem carregar as imagens
    try:
        info          = pdfinfo_from_bytes(arquivo_bytes)
        total_paginas = info["Pages"]
    except Exception:
        # Se pdfinfo falhar, usa convert para a primeira página e estima
        total_paginas = 999

    dataframes = []

    for num_pag in range(1, total_paginas + 1):
        try:
            # Converte apenas a página atual (sem carregar o PDF inteiro)
            imagens = convert_from_bytes(
                arquivo_bytes,
                dpi=DPI_OCR,
                first_page=num_pag,
                last_page=num_pag
            )

            if not imagens:
                break

            imagem = imagens[0]
            texto  = extrair_texto_imagem(imagem)

            # Libera imagem imediatamente
            del imagem, imagens
            gc.collect()

            df = texto_para_dataframe(texto, num_pag)
            if not df.empty:
                dataframes.append(df)

        except Exception:
            # Se uma página falhar, continua para a próxima
            gc.collect()
            continue

    return dataframes, num_pag


def ler_imagem(arquivo_bytes: bytes, nome_arquivo: str) -> pd.DataFrame:
    """Lê um único arquivo de imagem via OCR."""
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
    """Consolida DataFrames de múltiplas páginas em um único."""
    if not dataframes:
        return pd.DataFrame()
    return pd.concat(dataframes, ignore_index=True)