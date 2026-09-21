# ==============================================================================
# app.py
# DataMiner Planeja+ — Interface Principal (v2.1)
# Associação Raízes | Programa Planeja+
#
# Para rodar localmente:
#   streamlit run app.py
#
# Para deploy no Streamlit Cloud:
#   1. Suba o projeto para um repositório GitHub (sem a pasta data/)
#   2. Acesse share.streamlit.io e conecte o repositório
#   3. Configure GDRIVE_CSV_URL em Secrets (para o arquivo SICONFI)
# ==============================================================================

import io
import pandas as pd
import streamlit as st
from datetime import datetime

from modules.loader import get_csv_path
from modules.reader import (
    get_csv_columns, get_valores_unicos, query_csv,
    read_xlsx, read_pdf_tables
)
from modules.ocr import (
    ler_pdf_escaneado, ler_imagem, consolidar_paginas,
    verificar_dependencias as verificar_ocr
)
from modules.filter import get_municipios_por_regiao
from modules.registry import (
    inicializar_biblioteca, adicionar_documento, remover_documento,
    limpar_biblioteca, listar_documentos, listar_municipios_carregados,
    criar_registro, total_documentos, TIPOS_DOCUMENTO
)
from modules.crossref import cruzar_codigo, gerar_log_cruzamento, MODOS_BUSCA
from modules.exporter import exportar_csv, exportar_xlsx, exportar_pdf, exportar_docx
from connectors import obter_conector
from config.municipios import REGIOES, TODOS_MUNICIPIOS


# ==============================================================================
# Configuração da Página
# ==============================================================================

st.set_page_config(
    page_title="DataMiner Planeja+",
    page_icon="⛏️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    div.stButton > button[kind="primary"] {
        background-color: #1f4e2f;
        color: white;
        border: none;
        font-weight: bold;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #2d6b41;
    }
    div.stDownloadButton > button {
        background-color: #1f4e2f;
        color: white;
        border: none;
        font-weight: bold;
        width: 100%;
    }
    div.stDownloadButton > button:hover {
        background-color: #2d6b41;
    }
    .aviso-vi {
        background: #fff8e1;
        border-left: 4px solid #f9a825;
        padding: 8px 12px;
        border-radius: 4px;
        font-size: 0.82em;
        color: #5d4037;
        margin-top: 6px;
    }
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# Inicialização do Estado da Sessão
# ==============================================================================

inicializar_biblioteca()

for chave, padrao in {
    "df_resultado":      None,
    "log_resultado":     None,
    "filtros_resultado": {},
    "processados_keys":  set(),
}.items():
    if chave not in st.session_state:
        st.session_state[chave] = padrao


# ==============================================================================
# Funções Auxiliares
# ==============================================================================

def processar_arquivo_upload(arquivo, forcar_ocr: bool = False) -> tuple:
    """
    Lê um arquivo carregado pelo file_uploader do Streamlit.
    Retorna (DataFrame, origem_ocr: bool, total_paginas: int).
    """
    extensao     = arquivo.name.rsplit(".", 1)[-1].lower()
    bytes_arquivo = arquivo.getvalue()

    if extensao == "csv":
        for enc in ["utf-8-sig", "utf-8", "latin-1"]:
            try:
                df = pd.read_csv(io.BytesIO(bytes_arquivo), encoding=enc)
                return df, False, 1
            except Exception:
                continue
        return pd.DataFrame(), False, 0

    elif extensao in ["xlsx", "xls"]:
        df = read_xlsx(io.BytesIO(bytes_arquivo))
        return df, False, 1

    elif extensao == "pdf":
        if forcar_ocr:
            ok, msg = verificar_ocr()
            if not ok:
                raise RuntimeError(msg)
            dfs, total = ler_pdf_escaneado(bytes_arquivo)
            return consolidar_paginas(dfs) if dfs else pd.DataFrame(), True, total
        else:
            tabelas = read_pdf_tables(io.BytesIO(bytes_arquivo))
            if tabelas:
                df = pd.concat(tabelas, ignore_index=True)
                return df, False, len(tabelas)
            return pd.DataFrame(), False, 0

    elif extensao in ["png", "jpg", "jpeg", "tiff", "bmp"]:
        ok, msg = verificar_ocr()
        if not ok:
            raise RuntimeError(msg)
        df = ler_imagem(bytes_arquivo, arquivo.name)
        return df, True, 1

    return pd.DataFrame(), False, 0


def acha_indice(colunas: list, termos: list) -> int:
    """Retorna o índice da primeira coluna que contém algum dos termos."""
    for i, col in enumerate(colunas):
        if any(t in col.lower() for t in termos):
            return i
    return 0


# ==============================================================================
# Cabeçalho da Página
# ==============================================================================

st.title("⛏️ DataMiner Planeja+")
st.caption(
    "Extração e Cruzamento de Dados Orçamentários Municipais  "
    "|  Associação Raízes  |  Programa Planeja+"
)
st.divider()


# ==============================================================================
# Sidebar: Upload e Biblioteca de Documentos
# ==============================================================================

with st.sidebar:

    st.header("📤 Carregar Documento")

    arquivo = st.file_uploader(
        "Selecione o arquivo",
        type=["csv", "xlsx", "xls", "pdf", "png", "jpg", "jpeg", "tiff", "bmp"],
        help="Suporte a CSV, Excel, PDF (texto ou escaneado) e imagens."
    )

    if arquivo is not None:
        arquivo_key = f"{arquivo.name}_{arquivo.size}"
        extensao    = arquivo.name.rsplit(".", 1)[-1].lower()
        ja_na_bib   = arquivo_key in st.session_state.processados_keys

        if ja_na_bib:
            st.success(f"✓ {arquivo.name[:32]} já está na biblioteca.")
        else:
            municipio_upload = st.selectbox(
                "Município",
                options=TODOS_MUNICIPIOS,
                help="Município ao qual este documento pertence."
            )
            tipo_doc_upload = st.selectbox(
                "Tipo de documento",
                options=TIPOS_DOCUMENTO
            )
            exercicio_upload = st.text_input(
                "Exercício (ano)",
                value=str(datetime.now().year),
                max_chars=9,
                help="Ex: 2024 | Para PPA: 2022-2025"
            )

            forcar_ocr = False
            if extensao in ["pdf", "png", "jpg", "jpeg", "tiff", "bmp"]:
                forcar_ocr = st.checkbox(
                    "Usar OCR (documento escaneado ou imagem)",
                    value=(extensao != "pdf"),
                    help=(
                        "Marque para PDFs escaneados e imagens. "
                        "Para PDFs com texto selecionável, deixe desmarcado."
                    )
                )

            if st.button(
                "⚙️ Processar e Adicionar à Biblioteca",
                type="primary",
                use_container_width=True
            ):
                with st.spinner(f"Processando {arquivo.name}..."):
                    try:
                        df_lido, origem_ocr, total_pag = processar_arquivo_upload(
                            arquivo=arquivo,
                            forcar_ocr=forcar_ocr
                        )

                        if df_lido.empty and not forcar_ocr and extensao == "pdf":
                            st.warning(
                                "O PDF não contém tabelas com texto selecionável. "
                                "Marque 'Usar OCR' e tente novamente."
                            )
                        elif df_lido.empty:
                            st.error(
                                "Não foi possível extrair dados deste arquivo. "
                                "Verifique se o formato é compatível."
                            )
                        else:
                            conector   = obter_conector(tipo_doc_upload)
                            mapeamento = conector.detectar_colunas(df_lido.columns.tolist())

                            registro = criar_registro(
                                nome_arquivo   = arquivo.name,
                                municipio      = municipio_upload,
                                tipo_documento = tipo_doc_upload,
                                exercicio      = exercicio_upload,
                                formato        = extensao.upper(),
                                tamanho_bytes  = arquivo.size,
                                df             = df_lido,
                                col_municipio  = mapeamento.get("municipio"),
                                col_codigo     = mapeamento.get("codigo"),
                                col_descricao  = mapeamento.get("descricao"),
                                col_valor      = mapeamento.get("valor"),
                                origem_ocr     = origem_ocr,
                                paginas_totais = total_pag
                            )
                            adicionar_documento(registro)
                            st.session_state.processados_keys.add(arquivo_key)

                            msg_ocr = f" ({total_pag} páginas via OCR)" if origem_ocr else ""
                            st.success(
                                f"✓ Adicionado. "
                                f"{len(df_lido):,} registros extraídos{msg_ocr}."
                            )
                            if mapeamento.get("codigo"):
                                st.caption(
                                    f"Coluna de código detectada: `{mapeamento['codigo']}`"
                                )
                            st.rerun()

                    except RuntimeError as e:
                        st.error(str(e))
                    except Exception as e:
                        st.error(f"Erro inesperado: {e}")

    st.divider()

    total        = total_documentos()
    municipios_bib = listar_municipios_carregados()

    st.header(f"📚 Biblioteca  ({total} doc{'s' if total != 1 else ''})")

    if total == 0:
        st.caption(
            "Nenhum documento carregado nesta sessão. "
            "Faça upload pelo painel acima."
        )
    else:
        col_m1, col_m2 = st.columns(2)
        col_m1.metric("Documentos", total)
        col_m2.metric("Municípios", len(municipios_bib))

        st.caption("Clique em ▶ para ver detalhes ou remover.")

        for doc in listar_documentos():
            icone_ocr = " 🔍" if doc["origem_ocr"] else ""
            with st.expander(f"📄 {doc['nome_arquivo'][:26]}{icone_ocr}"):
                st.markdown(f"**Município:** {doc['municipio']}")
                st.markdown(
                    f"**Tipo:** {doc['tipo_documento']}  "
                    f"|  **Exercício:** {doc['exercicio']}"
                )
                st.markdown(
                    f"**Registros:** {doc['linhas']:,}  "
                    f"|  **Tamanho:** {doc['tamanho_legivel']}"
                )
                st.markdown(
                    f"**Formato:** {doc['formato']}  "
                    f"|  **ID:** `{doc['id']}`"
                )
                if doc.get("col_codigo"):
                    st.caption(f"Coluna de código: `{doc['col_codigo']}`")
                if doc["origem_ocr"] and doc.get("paginas_totais"):
                    st.caption(f"OCR — {doc['paginas_totais']} página(s)")

                if st.button(
                    "🗑️ Remover",
                    key=f"rm_{doc['id']}",
                    use_container_width=True
                ):
                    remover_documento(doc["id"])
                    st.rerun()

        st.divider()
        if st.button("🗑️ Limpar biblioteca completa", use_container_width=True):
            limpar_biblioteca()
            st.session_state.processados_keys = set()
            st.session_state.df_resultado     = None
            st.session_state.log_resultado    = None
            st.rerun()


# ==============================================================================
# Área Principal: Abas
# ==============================================================================

tab_cruzamento, tab_siconfi = st.tabs([
    "🔗 Cruzamento de Dados",
    "🛢️ SICONFI (arquivo grande)"
])


# ==============================================================================
# Aba: Cruzamento de Dados
# ==============================================================================

with tab_cruzamento:

    if total_documentos() == 0:
        st.info(
            "Nenhum documento na biblioteca ainda. "
            "Use o painel lateral para carregar arquivos de LOA, PPA ou outros documentos.",
            icon="ℹ️"
        )
    else:
        st.subheader("🔗 Cruzamento por Código Orçamentário")
        st.caption(
            "Informe um código de receita ou despesa para buscá-lo em todos os "
            "documentos carregados. O resultado consolida registros de múltiplas "
            "fontes com informações de origem em cada linha."
        )

        c1, c2, c3 = st.columns([3, 2, 2])

        with c1:
            codigo_busca = st.text_input(
                "Código orçamentário",
                placeholder="Ex: 1.7.2.8  |  04.122  |  03.001.0001",
                help=(
                    "Receita: 1.7.2.8.00.00 | Despesa funcional: 04.122 | "
                    "Despesa programática: 03.001.0001.2001 | "
                    "Natureza: 3.3.90.39.00"
                )
            )

        with c2:
            opcoes_mun = ["Todos os municípios carregados"] + listar_municipios_carregados()
            mun_filtro = st.selectbox(
                "Município",
                options=opcoes_mun,
                help="Restringe a busca a um único município ou abrange todos."
            )

        with c3:
            modo_labels = {
                "parcial": "Parcial (hierarquia)",
                "exato":   "Exato",
                "fuzzy":   "Fuzzy (aproximado)"
            }
            modo_busca = st.radio(
                "Modo de busca",
                options=list(modo_labels.keys()),
                format_func=lambda m: modo_labels[m],
                horizontal=True,
                help=(
                    "Parcial: 1.7 encontra 1.7.2.8.00.00. "
                    "Exato: apenas correspondência total. "
                    "Fuzzy: tolerante a erros de OCR e digitação."
                )
            )

        mun_param = None if mun_filtro == "Todos os municípios carregados" else mun_filtro

        if st.button(
            "🔎 Cruzar código",
            type="primary",
            disabled=not (codigo_busca or "").strip()
        ):
            with st.spinner("Cruzando código entre os documentos da biblioteca..."):
                df_cruzado = cruzar_codigo(
                    codigo_busca=codigo_busca.strip(),
                    municipio=mun_param,
                    modo=modo_busca
                )
                log = gerar_log_cruzamento(
                    codigo_busca, mun_param, modo_busca, df_cruzado
                )

            st.session_state.df_resultado    = df_cruzado
            st.session_state.log_resultado   = log
            st.session_state.filtros_resultado = {
                "Código buscado": codigo_busca.strip(),
                "Município":      mun_param or "Todos",
                "Modo de busca":  modo_labels.get(modo_busca, modo_busca)
            }


# ==============================================================================
# Aba: SICONFI
# ==============================================================================

with tab_siconfi:

    st.subheader("🛢️ Busca Direta no SICONFI")
    st.caption(
        "Consulta o arquivo CSV do SICONFI (até 135MB) sem carregar tudo na memória, "
        "usando DuckDB para filtrar diretamente em disco."
    )

    sa, sb = st.columns(2)

    with sa:
        opcoes_reg  = ["Todas as Regionais"] + list(REGIOES.keys())
        regioes_sel = st.multiselect(
            "Regional(is)",
            options=opcoes_reg,
            default=["Todas as Regionais"]
        )
        if any("VI" in r for r in regioes_sel):
            st.markdown(
                '<div class="aviso-vi"><strong>Regional VI:</strong> '
                'Paraty (RJ) e Caraguatatuba/Ilhabela (SP) integram esta regional. '
                'Considere isso ao cruzar dados com políticas estaduais.</div>',
                unsafe_allow_html=True
            )

    with sb:
        muns_disp = get_municipios_por_regiao(regioes_sel)
        muns_sel  = st.multiselect(
            "Município(s)",
            options=["Todos"] + muns_disp,
            default=["Todos"]
        )

    muns_filtro_sic = (
        muns_disp
        if "Todos" in muns_sel or not muns_sel
        else muns_sel
    )

    termo_rec_sic = st.text_input(
        "Filtrar por tipo de receita",
        placeholder="Ex: royalt, participação especial, CFEM...",
        key="termo_siconfi"
    )

    # ------------------------------------------------------------------
    # Localiza o arquivo CSV.
    # get_csv_path() retorna None quando o arquivo não está disponível,
    # sem interromper o app. A aba exibe instruções nesse caso.
    # ------------------------------------------------------------------

    csv_path    = None
    colunas_sic = []

    with st.spinner("Verificando arquivo CSV do SICONFI..."):
        csv_path = get_csv_path()

    if csv_path is not None:
        try:
            colunas_sic = get_csv_columns(csv_path)
        except Exception as e:
            st.warning(f"Erro ao ler a estrutura do arquivo SICONFI: {e}")

    # Arquivo não disponível: mostra instruções e encerra a aba
    if csv_path is None:
        st.info(
            "O arquivo SICONFI não está disponível nesta sessão. "
            "Para habilitar esta aba, escolha uma das opções abaixo:\n\n"
            "**Opção A (local):** coloque o arquivo `siconfi.csv` "
            "na pasta `data/` do projeto e reinicie o app.\n\n"
            "**Opção B (Streamlit Cloud):** configure o link do Google Drive "
            "em `.streamlit/secrets.toml`:\n"
            "```\nGDRIVE_CSV_URL = \"https://drive.google.com/...\"\n```\n\n"
            "Enquanto isso, use a aba **Cruzamento de Dados** para testar "
            "o sistema com arquivos de LOA, PPA ou outros documentos.",
            icon="ℹ️"
        )

    # Arquivo disponível e colunas lidas com sucesso
    elif colunas_sic:

        with st.expander(
            "⚙️ Mapeamento de colunas (clique para ajustar se necessário)",
            expanded=False
        ):
            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                col_mun_sic = st.selectbox(
                    "Município", colunas_sic,
                    index=acha_indice(colunas_sic, ["munic"])
                )
                col_uf_sic = st.selectbox(
                    "UF", colunas_sic,
                    index=acha_indice(colunas_sic, ["uf", "estado"])
                )
            with sc2:
                col_rec_sic = st.selectbox(
                    "Fonte / Tipo de Receita", colunas_sic,
                    index=acha_indice(colunas_sic, ["fonte", "conta", "descri"])
                )
                col_val_sic = st.selectbox(
                    "Valor", colunas_sic,
                    index=acha_indice(colunas_sic, ["valor", "receita", "montante"])
                )
            with sc3:
                usar_ano_sic = st.checkbox("Filtrar por ano?", value=False)
                col_ano_sic  = None
                anos_sel_sic = None
                if usar_ano_sic:
                    col_ano_sic = st.selectbox(
                        "Coluna de Ano", colunas_sic,
                        index=acha_indice(colunas_sic, ["ano", "exerc"])
                    )
                    with st.spinner("Lendo anos disponíveis..."):
                        anos_disp = get_valores_unicos(csv_path, col_ano_sic)
                    anos_sel_sic = st.multiselect(
                        "Ano(s)", options=anos_disp, default=anos_disp
                    )

        st.info(
            f"Arquivo SICONFI localizado. "
            f"{len(colunas_sic)} colunas | "
            f"{len(muns_filtro_sic)} município(s) no filtro.",
            icon="ℹ️"
        )

        if st.button("🔎 Buscar no SICONFI", type="primary"):
            if not muns_filtro_sic:
                st.warning("Selecione ao menos um município ou regional.")
            else:
                with st.spinner("Consultando SICONFI com DuckDB..."):
                    df_sic = query_csv(
                        csv_path      = csv_path,
                        municipios    = muns_filtro_sic,
                        col_municipio = col_mun_sic,
                        col_receita   = col_rec_sic,
                        termo_receita = termo_rec_sic or None,
                        col_ano       = col_ano_sic,
                        anos          = anos_sel_sic
                    )

                st.session_state.df_resultado    = df_sic
                st.session_state.log_resultado   = (
                    f"{len(df_sic):,} registros encontrados no SICONFI "
                    f"para {len(muns_filtro_sic)} município(s)."
                    if not df_sic.empty
                    else "Nenhum registro encontrado para os filtros selecionados."
                )
                st.session_state.filtros_resultado = {
                    "Fonte": "SICONFI / Tesouro Nacional",
                    "Regional(is)": (
                        ", ".join(
                            r.split(" - ")[-1]
                            for r in regioes_sel
                            if r != "Todas as Regionais"
                        ) or "Todas"
                    ),
                    "Municípios": (
                        ", ".join(muns_filtro_sic[:5])
                        + ("..." if len(muns_filtro_sic) > 5 else "")
                    ),
                    "Tipo de Receita": termo_rec_sic or "Todos",
                    "Ano(s)": (
                        ", ".join(str(a) for a in anos_sel_sic)
                        if anos_sel_sic else "Todos"
                    )
                }


# ==============================================================================
# Seção de Resultados — aparece abaixo das abas quando há dados
# ==============================================================================

df_res      = st.session_state.get("df_resultado")
log_res     = st.session_state.get("log_resultado")
filtros_res = st.session_state.get("filtros_resultado", {})

if df_res is not None:
    st.divider()
    st.subheader("📊 Resultados")

    if log_res:
        if df_res.empty:
            st.warning(log_res)
        else:
            st.success(log_res)

    if not df_res.empty:

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Registros", f"{len(df_res):,}")
        m2.metric(
            "Municípios",
            df_res["_municipio"].nunique()
            if "_municipio" in df_res.columns
            else df_res.iloc[:, 0].nunique()
        )
        m3.metric(
            "Arquivos fonte",
            df_res["_fonte_arquivo"].nunique()
            if "_fonte_arquivo" in df_res.columns
            else "N/A"
        )
        m4.metric(
            "Tipos de doc.",
            df_res["_tipo_documento"].nunique()
            if "_tipo_documento" in df_res.columns
            else len(df_res.columns)
        )

        st.dataframe(df_res, use_container_width=True, height=420)

        st.divider()

        with st.expander("⚙️ Opções de exportação", expanded=False):
            op1, op2 = st.columns(2)
            with op1:
                incluir_prov = st.checkbox(
                    "Incluir colunas de origem nos arquivos",
                    value=True,
                    help=(
                        "Colunas que começam com '_' identificam de qual "
                        "arquivo e linha cada registro foi extraído."
                    )
                )
            with op2:
                modo_simples_exp = st.checkbox(
                    "Modo simples (sem cabeçalho institucional)",
                    value=False
                )

        st.subheader("⬇️ Exportar Resultados")
        nome_base = f"dataminer_{datetime.now().strftime('%Y%m%d_%H%M')}"

        bc1, bc2, bc3, bc4 = st.columns(4)

        with bc1:
            st.download_button(
                label="📥 CSV",
                data=exportar_csv(df_res, incluir_prov),
                file_name=f"{nome_base}.csv",
                mime="text/csv",
                use_container_width=True,
                help="Melhor para grandes volumes e integração com Excel."
            )

        with bc2:
            try:
                xlsx_bytes = exportar_xlsx(
                    df_res, filtros_res, modo_simples_exp, incluir_prov
                )
                st.download_button(
                    label="📊 Excel",
                    data=xlsx_bytes,
                    file_name=f"{nome_base}.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                    help="Planilha formatada com colunas de origem destacadas."
                )
            except Exception as e:
                st.error(f"Erro ao gerar Excel: {e}")

        with bc3:
            try:
                pdf_bytes = exportar_pdf(
                    df_res, filtros_res, modo_simples_exp, incluir_prov
                )
                st.download_button(
                    label="📄 PDF",
                    data=pdf_bytes,
                    file_name=f"{nome_base}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    help="Relatório com cabeçalho institucional Raízes/Planeja+."
                )
            except Exception as e:
                st.error(f"Erro ao gerar PDF: {e}")

        with bc4:
            try:
                docx_bytes = exportar_docx(
                    df_res, filtros_res, modo_simples_exp, incluir_prov
                )
                st.download_button(
                    label="📝 Word",
                    data=docx_bytes,
                    file_name=f"{nome_base}.docx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.wordprocessingml.document"
                    ),
                    use_container_width=True,
                    help="Documento Word para edição e incorporação em relatórios."
                )
            except Exception as e:
                st.error(f"Erro ao gerar Word: {e}")

        if len(df_res) >= 2000:
            st.caption(
                "O PDF exporta no máximo 2.000 linhas e o Word 1.000. "
                "Para volumes maiores, use CSV ou Excel."
            )