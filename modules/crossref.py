# ==============================================================================
# crossref.py
# Motor de cruzamento de dados por código orçamentário.
#
# Suporta os formatos padrão brasileiros de classificação:
#   Receita:              1.7.2.8.00.00  (6 níveis com pontos)
#   Despesa funcional:    04.122         (função.subfunção)
#   Despesa programática: 03.001.0001.2001 (unidade.programa.ação.subação)
#   Natureza da despesa:  3.3.90.39.00   (categoria.grupo.modalidade.elemento.subelemento)
#
# Modos de busca:
#   exato    - código idêntico ao informado
#   parcial  - código começa com o trecho informado (hierarquia orçamentária)
#   fuzzy    - busca aproximada, tolerante a variações de digitação e OCR
# ==============================================================================

import re
import pandas as pd
from typing import Optional

try:
    from rapidfuzz import fuzz
    RAPIDFUZZ_DISPONIVEL = True
except ImportError:
    RAPIDFUZZ_DISPONIVEL = False

MODOS_BUSCA = ["parcial", "exato", "fuzzy"]
LIMIAR_FUZZY = 80  # Pontuação mínima de similaridade (0 a 100)

# Colunas de proveniência adicionadas em cada linha do resultado
COLUNAS_PROVENIENCIA = [
    "_doc_id",
    "_municipio",
    "_cod_ibge",
    "_tipo_documento",
    "_exercicio",
    "_fonte_arquivo",
    "_formato",
    "_origem_ocr",
    "_pagina_origem",
    "_linha_origem",
    "_data_extracao"
]


# ==============================================================================
# Normalização de códigos para comparação
# ==============================================================================

def normalizar_codigo(codigo: str) -> str:
    """
    Prepara um código para comparação: remove espaços extras
    e converte para minúsculas. Preserva os pontos, pois eles
    fazem parte da hierarquia orçamentária.
    Exemplo: "  1.7.2.8.00.00 " -> "1.7.2.8.00.00"
    """
    return str(codigo).strip().lower().replace(" ", "")


def extrair_codigos_do_texto(texto: str) -> list:
    """
    Identifica e extrai todos os códigos orçamentários presentes
    em uma string de texto. Útil para células que contêm código
    e descrição juntos (ex: "1.7.2.8.00.00 - Royalties Petróleo").

    Padrões reconhecidos:
        - Receita:              1.7.2.8.00.00
        - Despesa funcional:    04.122
        - Despesa programática: 03.001.0001.2001
        - Natureza da despesa:  3.3.90.39.00
    """
    padroes = [
        r"\d{1,2}\.\d{1,3}\.\d{1,4}\.\d{1,4}",   # 4 grupos (programática)
        r"\d{1,2}\.\d{1,2}\.\d{2}\.\d{2}\.\d{2}",  # 5 grupos (natureza despesa)
        r"\d\.\d+\.\d+\.\d+\.\d+\.\d+",             # 6 grupos (receita)
        r"\d{2}\.\d{3}",                             # 2 grupos (funcional)
    ]
    codigos_encontrados = []
    for padrao in padroes:
        codigos_encontrados.extend(re.findall(padrao, str(texto)))
    return list(set(codigos_encontrados))


# ==============================================================================
# Verificação de correspondência de código
# ==============================================================================

def codigo_coincide(
    codigo_documento: str,
    codigo_busca: str,
    modo: str = "parcial"
) -> bool:
    """
    Verifica se um código de um documento corresponde ao código buscado.

    Modos:
        exato   - o código do documento deve ser idêntico ao buscado
        parcial - o código do documento começa com o código buscado,
                  ou o código buscado está contido no código do documento
                  (navega pela hierarquia orçamentária)
        fuzzy   - usa rapidfuzz para tolerância a erros de OCR e digitação
    """
    cod_doc = normalizar_codigo(codigo_documento)
    cod_busca = normalizar_codigo(codigo_busca)

    if not cod_doc or not cod_busca:
        return False

    if modo == "exato":
        return cod_doc == cod_busca

    elif modo == "parcial":
        # Verifica correspondência hierárquica em ambas as direções
        return cod_doc.startswith(cod_busca) or cod_busca.startswith(cod_doc)

    elif modo == "fuzzy":
        if not RAPIDFUZZ_DISPONIVEL:
            # Fallback para parcial se rapidfuzz não estiver disponível
            return cod_doc.startswith(cod_busca) or cod_busca in cod_doc
        return fuzz.partial_ratio(cod_busca, cod_doc) >= LIMIAR_FUZZY

    return False


# ==============================================================================
# Busca em um DataFrame individual
# ==============================================================================

def buscar_em_dataframe(
    df: pd.DataFrame,
    codigo_busca: str,
    col_codigo: Optional[str],
    modo: str = "parcial"
) -> pd.DataFrame:
    """
    Busca o código em um DataFrame. Se col_codigo for informada, busca
    apenas nessa coluna. Caso contrário, varre todas as colunas de texto
    tentando encontrar o código em qualquer campo.

    Retorna as linhas que contêm o código, com uma coluna extra
    '_linha_origem' indicando o número da linha no arquivo original.
    """
    if df.empty:
        return pd.DataFrame()

    df_busca = df.copy()
    df_busca["_linha_origem"] = df_busca.index + 1  # 1-indexado para o usuário

    mascaras = []

    if col_codigo and col_codigo in df_busca.columns:
        # Busca na coluna específica de código
        mascara = df_busca[col_codigo].astype(str).apply(
            lambda v: codigo_coincide(v, codigo_busca, modo)
        )
        mascaras.append(mascara)

        # Também tenta extrair códigos embutidos em texto misto
        mascara_embutida = df_busca[col_codigo].astype(str).apply(
            lambda v: any(
                codigo_coincide(c, codigo_busca, modo)
                for c in extrair_codigos_do_texto(v)
            )
        )
        mascaras.append(mascara_embutida)

    else:
        # Busca em todas as colunas de texto quando col_codigo não está definida
        colunas_texto = df_busca.select_dtypes(include=["object"]).columns
        for col in colunas_texto:
            mascara = df_busca[col].astype(str).apply(
                lambda v: codigo_coincide(v, codigo_busca, modo)
                or any(
                    codigo_coincide(c, codigo_busca, modo)
                    for c in extrair_codigos_do_texto(v)
                )
            )
            mascaras.append(mascara)

    if not mascaras:
        return pd.DataFrame()

    mascara_final = mascaras[0]
    for m in mascaras[1:]:
        mascara_final = mascara_final | m

    return df_busca[mascara_final].reset_index(drop=True)


# ==============================================================================
# Adição de metadados de proveniência ao resultado
# ==============================================================================

def adicionar_proveniencia(df: pd.DataFrame, doc: dict) -> pd.DataFrame:
    """
    Adiciona as colunas de rastreabilidade de origem a cada linha do resultado.
    Essas colunas começam com '_' para serem facilmente identificadas na saída.
    """
    from datetime import datetime
    resultado = df.copy()

    resultado.insert(0, "_doc_id", doc.get("id", ""))
    resultado.insert(1, "_municipio", doc.get("municipio", ""))
    resultado.insert(2, "_cod_ibge", "")  # preenchido pelo conector quando disponível
    resultado.insert(3, "_tipo_documento", doc.get("tipo_documento", ""))
    resultado.insert(4, "_exercicio", doc.get("exercicio", ""))
    resultado.insert(5, "_fonte_arquivo", doc.get("nome_arquivo", ""))
    resultado.insert(6, "_formato", doc.get("formato", ""))
    resultado.insert(7, "_origem_ocr", "Sim" if doc.get("origem_ocr") else "Não")
    resultado.insert(8, "_pagina_origem", resultado.get("pagina_origem", ""))
    resultado.insert(9, "_data_extracao", datetime.now().strftime("%d/%m/%Y %H:%M"))

    # Remove coluna pagina_origem duplicada se existir
    if "pagina_origem" in resultado.columns:
        resultado = resultado.drop(columns=["pagina_origem"])

    return resultado


# ==============================================================================
# Busca em um documento completo do registry
# ==============================================================================

def buscar_em_documento(
    codigo_busca: str,
    doc: dict,
    modo: str = "parcial"
) -> pd.DataFrame:
    """
    Busca um código em um documento registrado na biblioteca.
    Extrai o DataFrame do documento, executa a busca e adiciona
    os metadados de proveniência nas linhas encontradas.
    """
    df = doc.get("df")
    if df is None or df.empty:
        return pd.DataFrame()

    col_codigo = doc.get("col_codigo")
    df_resultado = buscar_em_dataframe(df, codigo_busca, col_codigo, modo)

    if df_resultado.empty:
        return pd.DataFrame()

    return adicionar_proveniencia(df_resultado, doc)


# ==============================================================================
# Função principal: cruzamento entre documentos
# ==============================================================================

def cruzar_codigo(
    codigo_busca: str,
    municipio: Optional[str] = None,
    modo: str = "parcial"
) -> pd.DataFrame:
    """
    Busca um código orçamentário em todos os documentos carregados na biblioteca.
    Se 'municipio' for informado, restringe a busca a esse município.

    Retorna um DataFrame consolidado com todos os registros encontrados
    em todos os documentos, com colunas de proveniência em cada linha.

    Parâmetros:
        codigo_busca  - código a ser buscado (ex: "1.7.2", "04.122", "03.001.0001")
        municipio     - filtro por município (None para buscar em todos)
        modo          - "parcial", "exato" ou "fuzzy"
    """
    from modules.registry import get_documentos_por_municipio, listar_documentos

    if municipio:
        documentos = get_documentos_por_municipio(municipio)
    else:
        documentos = listar_documentos()

    if not documentos:
        return pd.DataFrame()

    resultados = []
    log = []

    for doc in documentos:
        df_resultado = buscar_em_documento(codigo_busca, doc, modo)
        if not df_resultado.empty:
            resultados.append(df_resultado)
            log.append({
                "documento": doc.get("nome_arquivo"),
                "municipio": doc.get("municipio"),
                "tipo": doc.get("tipo_documento"),
                "registros_encontrados": len(df_resultado)
            })

    if not resultados:
        return pd.DataFrame()

    df_consolidado = pd.concat(resultados, ignore_index=True)
    return df_consolidado


def gerar_log_cruzamento(
    codigo_busca: str,
    municipio: Optional[str],
    modo: str,
    df_resultado: pd.DataFrame
) -> str:
    """
    Gera um texto de log resumindo o resultado do cruzamento.
    Usado na interface para exibir um resumo antes da tabela completa.
    """
    if df_resultado.empty:
        return f"Nenhum registro encontrado para o código '{codigo_busca}'."

    total = len(df_resultado)
    fontes = df_resultado["_fonte_arquivo"].nunique() if "_fonte_arquivo" in df_resultado else 0
    municipios = df_resultado["_municipio"].nunique() if "_municipio" in df_resultado else 0
    tipos = df_resultado["_tipo_documento"].unique().tolist() if "_tipo_documento" in df_resultado else []

    return (
        f"{total} registro(s) encontrado(s) para o código '{codigo_busca}' "
        f"em {fontes} arquivo(s) de {municipios} município(s). "
        f"Tipos de documento: {', '.join(tipos)}."
    )