# ==============================================================================
# registry.py
# Gerencia a Biblioteca de Documentos do DataMiner Planeja+.
#
# Responsabilidades:
#   - Registrar cada arquivo carregado com seus metadados
#   - Armazenar o DataFrame extraído junto com os metadados
#   - Permitir remoção de documentos individuais
#   - Fornecer consultas por município e tipo de documento
#   - Persistir em disco (local) ou manter em sessão (Streamlit Cloud)
# ==============================================================================

import os
import json
import uuid
import pandas as pd
import streamlit as st
from datetime import datetime
from typing import Optional

# Chave do session_state onde a biblioteca fica armazenada
CHAVE_BIBLIOTECA = "dataminer_biblioteca"

# Tipos de documento suportados
TIPOS_DOCUMENTO = ["SICONFI", "LOA", "PPA", "LDO", "Balanço Orçamentário", "Outro"]

# Formatos de arquivo suportados
FORMATOS_SUPORTADOS = {
    "csv":  "CSV",
    "xlsx": "Excel",
    "xls":  "Excel",
    "pdf":  "PDF",
    "png":  "Imagem",
    "jpg":  "Imagem",
    "jpeg": "Imagem",
    "tiff": "Imagem",
    "bmp":  "Imagem"
}


# ==============================================================================
# Inicialização da biblioteca na sessão
# ==============================================================================

def inicializar_biblioteca():
    """
    Garante que a biblioteca de documentos existe no session_state.
    Deve ser chamada no início de cada execução do app.py.
    """
    if CHAVE_BIBLIOTECA not in st.session_state:
        st.session_state[CHAVE_BIBLIOTECA] = {}


# ==============================================================================
# Estrutura de um registro de documento
# ==============================================================================

def criar_registro(
    nome_arquivo: str,
    municipio: str,
    tipo_documento: str,
    exercicio: str,
    formato: str,
    tamanho_bytes: int,
    df: pd.DataFrame,
    col_municipio: Optional[str] = None,
    col_codigo: Optional[str] = None,
    col_descricao: Optional[str] = None,
    col_valor: Optional[str] = None,
    origem_ocr: bool = False,
    paginas_totais: Optional[int] = None
) -> dict:
    """
    Cria um dicionário com todos os metadados de um documento carregado.
    O campo 'id' é gerado automaticamente e é único por documento.
    """
    return {
        "id": str(uuid.uuid4())[:8],
        "nome_arquivo": nome_arquivo,
        "municipio": municipio,
        "tipo_documento": tipo_documento,
        "exercicio": str(exercicio),
        "formato": formato,
        "tamanho_bytes": tamanho_bytes,
        "tamanho_legivel": _formatar_tamanho(tamanho_bytes),
        "data_carga": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "linhas": len(df),
        "colunas": len(df.columns),
        "origem_ocr": origem_ocr,
        "paginas_totais": paginas_totais,
        "col_municipio": col_municipio,
        "col_codigo": col_codigo,
        "col_descricao": col_descricao,
        "col_valor": col_valor,
        "df": df
    }


# ==============================================================================
# Operações CRUD da biblioteca
# ==============================================================================

def adicionar_documento(registro: dict):
    """
    Adiciona um documento à biblioteca da sessão.
    Usa o campo 'id' do registro como chave.
    """
    inicializar_biblioteca()
    st.session_state[CHAVE_BIBLIOTECA][registro["id"]] = registro


def remover_documento(doc_id: str):
    """Remove um documento da biblioteca pelo seu ID."""
    inicializar_biblioteca()
    if doc_id in st.session_state[CHAVE_BIBLIOTECA]:
        del st.session_state[CHAVE_BIBLIOTECA][doc_id]


def limpar_biblioteca():
    """Remove todos os documentos da biblioteca."""
    st.session_state[CHAVE_BIBLIOTECA] = {}


def listar_documentos() -> list:
    """
    Retorna lista de todos os documentos carregados,
    ordenados do mais recente para o mais antigo.
    Cada item inclui todos os metadados, mas não o DataFrame
    (para não pesar a iteração).
    """
    inicializar_biblioteca()
    documentos = list(st.session_state[CHAVE_BIBLIOTECA].values())
    return sorted(documentos, key=lambda d: d["data_carga"], reverse=True)


def listar_municipios_carregados() -> list:
    """
    Retorna lista ordenada dos municípios que têm ao menos
    um documento carregado na biblioteca.
    """
    municipios = {d["municipio"] for d in listar_documentos()}
    return sorted(municipios)


def get_documentos_por_municipio(municipio: str) -> list:
    """
    Retorna todos os documentos carregados de um município específico.
    Inclui o DataFrame completo de cada um.
    """
    inicializar_biblioteca()
    return [
        d for d in st.session_state[CHAVE_BIBLIOTECA].values()
        if d["municipio"] == municipio
    ]


def get_documento_por_id(doc_id: str) -> Optional[dict]:
    """Retorna um documento específico pelo ID, ou None se não existir."""
    inicializar_biblioteca()
    return st.session_state[CHAVE_BIBLIOTECA].get(doc_id)


def total_documentos() -> int:
    """Retorna o número total de documentos na biblioteca."""
    inicializar_biblioteca()
    return len(st.session_state[CHAVE_BIBLIOTECA])


# ==============================================================================
# Exibição da biblioteca como tabela de resumo (sem o DataFrame)
# ==============================================================================

def get_resumo_biblioteca() -> pd.DataFrame:
    """
    Retorna um DataFrame com os metadados de todos os documentos,
    formatado para exibição na interface. Não inclui os DataFrames internos.
    """
    documentos = listar_documentos()
    if not documentos:
        return pd.DataFrame()

    linhas = []
    for d in documentos:
        linhas.append({
            "ID": d["id"],
            "Arquivo": d["nome_arquivo"],
            "Município": d["municipio"],
            "Tipo": d["tipo_documento"],
            "Exercício": d["exercicio"],
            "Formato": d["formato"],
            "Tamanho": d["tamanho_legivel"],
            "Registros": f"{d['linhas']:,}",
            "OCR": "Sim" if d["origem_ocr"] else "Não",
            "Carregado em": d["data_carga"]
        })

    return pd.DataFrame(linhas)


# ==============================================================================
# Persistência em disco (para execução local)
# ==============================================================================

def salvar_metadata_em_disco(pasta_data: str = "data"):
    """
    Salva os metadados da biblioteca atual em disco (sem os DataFrames).
    Útil para reconstruir a lista de arquivos já processados em execuções futuras.
    Os arquivos em si precisam ainda estar na pasta data/ para serem relidos.
    """
    os.makedirs("registry", exist_ok=True)
    metadados = []
    for doc in listar_documentos():
        entrada = {k: v for k, v in doc.items() if k != "df"}
        metadados.append(entrada)

    caminho = os.path.join("registry", "sources.json")
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(metadados, f, ensure_ascii=False, indent=2)


# ==============================================================================
# Utilitários internos
# ==============================================================================

def _formatar_tamanho(bytes_val: int) -> str:
    """Converte bytes para string legível (KB, MB)."""
    if bytes_val < 1024:
        return f"{bytes_val} B"
    elif bytes_val < 1024 ** 2:
        return f"{bytes_val / 1024:.1f} KB"
    else:
        return f"{bytes_val / (1024 ** 2):.1f} MB"