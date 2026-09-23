# ==============================================================================
# app.py
# DataMiner Planeja+ — Interface Principal (v4.0)
# Associação Raízes | Programa Planeja+
# ==============================================================================

import io
import os
import pandas as pd
import streamlit as st
from datetime import datetime

from modules.loader import sincronizar_e_listar_csvs, nome_amigavel_csv
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
from modules.crossref import cruzar_codigo, gerar_log_cruzamento
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
    div.stButton > button[kind="primary"]:hover { background-color: #2d6b41; }
    div.stDownloadButton > button {
        background-color: #1f4e2f;
        color: white;
        border: none;
        font-weight: bold;
        width: 100%;
    }
    div.stDownloadButton > button:hover { background-color: #2d6b41; }
    .aviso-vi {
        background: #fff8e1;
        border-left: 4px solid #f9a825;
        padding: 8px 12px;
        border-radius: 4px;
        font-size: 0.82em;
        color: #5d4037;
        margin-top: 6px;
    }
    .caixa-copiar {
        background: #1e3a2a;
        border: 1px solid #2d6b41;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 12px;
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
    "csvs_disponiveis":  [],
}.items():
    if chave not in st.session_state:
        st.session_state[chave] = padrao


# ==============================================================================
# Sincronização com o Google Drive (uma vez por sessão)
# ==============================================================================

if not st.session_state.csvs_disponiveis:
    st.session_state.csvs_disponiveis = sincronizar_e_listar_csvs()


# ==============================================================================
# Funções Auxiliares
# ==============================================================================

def processar_arquivo_upload(arquivo, forcar_ocr: bool = False) -> tuple:
    """Lê um arquivo carregado e retorna (DataFrame, origem_ocr, total_paginas)."""
    extensao      = arquivo.name.rsplit(".", 1)[-1].lower()
    bytes_arquivo = arquivo.getvalue()

    if extensao == "csv":
        for enc in ["utf-8-sig", "utf-8", "latin-1"]:
            try:
                return pd.read_csv(io.BytesIO(bytes_arquivo), encoding=enc), False, 1
            except Exception:
                continue
        return pd.DataFrame(), False, 0

    elif extensao in ["xlsx", "xls"]:
        return read_xlsx(io.BytesIO(bytes_arquivo)), False, 1

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
                return pd.concat(tabelas, ignore_index=True), False, len(tabelas)
            return pd.DataFrame(), False, 0

    elif extensao in ["png", "jpg", "jpeg", "tiff", "bmp"]:
        ok, msg = verificar_ocr()
        if not ok:
            raise RuntimeError(msg)
        return ler_imagem(bytes_arquivo, arquivo.name), True, 1

    return pd.DataFrame(), False, 0


def acha_indice(colunas: list, termos: list) -> int:
    """Retorna o índice da primeira coluna que contém algum dos termos."""
    for i, col in enumerate(colunas):
        if any(t in col.lower() for t in termos):
            return i
    return 0


def df_para_texto_copiavel(df: pd.DataFrame) -> str:
    """
    Converte o DataFrame para texto separado por tabulações (TSV).
    Ao colar no Excel ou Google Sheets, cada célula vai para a coluna certa.
    Remove colunas de proveniência (prefixo '_') para deixar mais limpo.
    """
    colunas_dados = [c for c in df.columns if not str(c).startswith("_")]
    df_limpo = df[colunas_dados]
    return df_limpo.to_csv(sep="\t", index=False, encoding="utf-8")


# ==============================================================================
# Cabeçalho: Logos empilhadas + Título centralizado
# ==============================================================================

col_logos, col_titulo = st.columns([1.3, 5])

with col_logos:
    if os.path.exists("assets/logo_planeja.png"):
        st.image("assets/logo_planeja.png", use_container_width=True)
    st.write("")
    if os.path.exists("assets/logo_pgp.png"):
        st.image("assets/logo_pgp.png", use_container_width=True)

with col_titulo:
    st.markdown("""
    <div style="padding-top: 12px;">
        <h1 style="margin-bottom: 6px; font-size: 2.2rem;">⛏️ DataMiner Planeja+</h1>
        <p style="text-align: center; color: #aaa; margin: 0; font-size: 1rem; line-height: 1.6;">
            Extração e Cruzamento de Dados Orçamentários Municipais
        </p>
        <p style="text-align: center; color: #888; margin: 0; font-size: 0.9rem; line-height: 1.6;">
            Associação Raízes | Programa Planeja+
        </p>
    </div>
    """, unsafe_allow_html=True)

st.divider()


# ==============================================================================
# Sidebar: Upload e Biblioteca de Documentos
# ==============================================================================

with st.sidebar:

    st.header("📤 Enviar Documento")
    st.caption("Carregue aqui os arquivos que deseja pesquisar: LOA, PPA, balanços, etc.")

    arquivo = st.file_uploader(
        "Selecione o arquivo",
        type=["csv", "xlsx", "xls", "pdf", "png", "jpg", "jpeg", "tiff", "bmp"],
        help=(
            "Formatos aceitos: CSV, Excel (.xlsx/.xls), PDF com texto "
            "selecionável ou escaneado (requer OCR), e imagens."
        )
    )

    if arquivo is not None:
        arquivo_key = f"{arquivo.name}_{arquivo.size}"
        extensao    = arquivo.name.rsplit(".", 1)[-1].lower()
        ja_na_bib   = arquivo_key in st.session_state.processados_keys

        if ja_na_bib:
            st.success(f"✓ {arquivo.name[:32]} já está na biblioteca.")
        else:
            municipio_upload = st.selectbox(
                "A qual município pertence este arquivo?",
                options=TODOS_MUNICIPIOS,
                help="Selecione o município do Planeja+ ao qual este documento se refere."
            )
            tipo_doc_upload = st.selectbox(
                "Tipo de documento",
                options=TIPOS_DOCUMENTO,
                help=(
                    "LOA: Lei Orçamentária Anual (receitas e despesas do exercício). "
                    "PPA: Plano Plurianual (4 anos). LDO: Lei de Diretrizes Orçamentárias. "
                    "SICONFI: planilha do Tesouro Nacional. Outro: demais documentos financeiros."
                )
            )
            exercicio_upload = st.text_input(
                "Ano do documento",
                value=str(datetime.now().year),
                max_chars=9,
                help="Informe o ano de referência. Para PPA, use o período: 2022-2025."
            )

            forcar_ocr = False
            if extensao in ["pdf", "png", "jpg", "jpeg", "tiff", "bmp"]:
                forcar_ocr = st.checkbox(
                    "Usar OCR (documento escaneado ou foto)",
                    value=(extensao != "pdf"),
                    help=(
                        "Marque para PDFs digitalizados em scanner ou fotografados. "
                        "PDFs gerados em computador não precisam de OCR."
                    )
                )

            if st.button("✅ Processar e adicionar", type="primary", use_container_width=True):
                with st.spinner(f"Processando {arquivo.name}..."):
                    try:
                        df_lido, origem_ocr, total_pag = processar_arquivo_upload(
                            arquivo=arquivo, forcar_ocr=forcar_ocr
                        )

                        if df_lido.empty and not forcar_ocr and extensao == "pdf":
                            st.warning(
                                "PDF sem tabelas detectáveis. "
                                "Marque 'Usar OCR' e tente novamente."
                            )
                        elif df_lido.empty:
                            st.error("Não foi possível extrair dados deste arquivo.")
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

                            msg_ocr = f" via OCR ({total_pag} páginas)" if origem_ocr else ""
                            st.success(f"✓ {len(df_lido):,} registros extraídos{msg_ocr}.")
                            st.rerun()

                    except RuntimeError as e:
                        st.error(str(e))
                    except Exception as e:
                        st.error(f"Erro: {e}")

    st.divider()

    # ---- Biblioteca ----
    total          = total_documentos()
    municipios_bib = listar_municipios_carregados()

    st.header(f"📚 Documentos carregados ({total})")

    if total == 0:
        st.caption("Nenhum documento ainda. Envie um arquivo acima para começar.")
    else:
        col_m1, col_m2 = st.columns(2)
        col_m1.metric("Documentos", total)
        col_m2.metric("Municípios", len(municipios_bib))

        for doc in listar_documentos():
            icone = "🔍" if doc["origem_ocr"] else "📄"
            with st.expander(f"{icone} {doc['nome_arquivo'][:26]}"):
                st.markdown(
                    f"**{doc['municipio']}** · {doc['tipo_documento']} {doc['exercicio']}"
                )
                st.caption(f"{doc['linhas']:,} registros · {doc['formato']}")
                if doc.get("col_codigo"):
                    st.caption(f"Coluna de código: `{doc['col_codigo']}`")
                if st.button("🗑️ Remover", key=f"rm_{doc['id']}", use_container_width=True):
                    remover_documento(doc["id"])
                    st.rerun()

        st.divider()
        if st.button("🗑️ Remover todos", use_container_width=True):
            limpar_biblioteca()
            st.session_state.processados_keys = set()
            st.session_state.df_resultado     = None
            st.rerun()


# ==============================================================================
# Área Principal: Abas
# ==============================================================================

tab_documentos, tab_siconfi, tab_ajuda = st.tabs([
    "🔍 Busca de Dados em Documentos enviados",
    "📂 Fonte de Dados SICONFI",
    "📖 Manual de Uso"
])

# ==============================================================================
# Aba: Manual de Uso
# ==============================================================================

with tab_ajuda:
    import streamlit.components.v1 as components

    manual_path = "assets/manual.html"

    if os.path.exists(manual_path):
        with open(manual_path, "r", encoding="utf-8") as f:
            html_manual = f.read()

        st.download_button(
            label="⬇️ Baixar manual (abrir no navegador e Ctrl+P para salvar como PDF)",
            data=html_manual.encode("utf-8"),
            file_name="Manual_DataMiner_Planeja.html",
            mime="text/html",
            use_container_width=False
        )

        components.html(html_manual, height=820, scrolling=True)
    else:
        st.info(
            "Manual não encontrado. Certifique-se de que o arquivo "
            "`manual.html` está na pasta `assets/` do projeto.",
            icon="📖"
        )
# ==============================================================================
# Aba: Busca de Dados em Documentos enviados
# ==============================================================================

with tab_documentos:

    if total_documentos() == 0:
        st.info(
            "Ainda não há documentos carregados. "
            "Use o painel à esquerda para enviar um arquivo de LOA, PPA ou outro documento.",
            icon="👈"
        )
    else:
        st.subheader("🔍 Buscar por código orçamentário")
        st.caption(
            "Digite um código de receita ou despesa para encontrá-lo nos documentos carregados. "
            "O resultado mostra todos os registros encontrados, com a fonte de cada um."
        )

        c1, c2, c3 = st.columns([3, 2, 2])

        with c1:
            codigo_busca = st.text_input(
                "Código para buscar",
                placeholder="Ex: 1.7   |   3390.39   |   04.122",
                help=(
                    "Digite o código que deseja encontrar. Exemplos:\n\n"
                    "1.7 → royalties de petróleo e gás (receita)\n\n"
                    "1.7.2.8.00.00 → código exato de royalties\n\n"
                    "3390.39 → outros serviços de terceiros (despesa)\n\n"
                    "04.122 → administração geral (funcional)\n\n"
                    "No modo Parcial, basta digitar os primeiros números. "
                    "1.7 encontra 1.7.2, 1.7.2.8.00.00 e todos os subcódigos."
                )
            )

        with c2:
            opcoes_mun = ["Todos os municípios"] + listar_municipios_carregados()
            mun_filtro = st.selectbox(
                "Município",
                options=opcoes_mun,
                help="Filtre por município ou busque em todos os documentos carregados."
            )

        with c3:
            modo_labels = {
                "parcial": "Parcial (recomendado)",
                "exato":   "Exato",
                "fuzzy":   "Aproximado"
            }
            modo_busca = st.radio(
                "Modo de busca",
                options=list(modo_labels.keys()),
                format_func=lambda m: modo_labels[m],
                horizontal=True,
                help=(
                    "Parcial: encontra o código e todos os seus subcódigos. "
                    "Exato: apenas o código digitado. "
                    "Aproximado: tolerante a erros de digitação."
                )
            )

        mun_param = None if mun_filtro == "Todos os municípios" else mun_filtro

        col_btn, col_dica = st.columns([2, 5])
        with col_btn:
            buscar = st.button(
                "🔎 Buscar agora",
                type="primary",
                use_container_width=True,
                disabled=not (codigo_busca or "").strip()
            )
        with col_dica:
            if not (codigo_busca or "").strip():
                st.caption("👆 Digite um código acima para habilitar a busca.")

        if buscar:
            with st.spinner("Buscando nos documentos..."):
                df_cruzado = cruzar_codigo(
                    codigo_busca=codigo_busca.strip(),
                    municipio=mun_param,
                    modo=modo_busca
                )
                log = gerar_log_cruzamento(
                    codigo_busca, mun_param, modo_busca, df_cruzado
                )

            st.session_state.df_resultado     = df_cruzado
            st.session_state.log_resultado    = log
            st.session_state.filtros_resultado = {
                "Código buscado": codigo_busca.strip(),
                "Município":      mun_param or "Todos",
                "Modo de busca":  modo_labels.get(modo_busca, modo_busca)
            }


# ==============================================================================
# Aba: Fonte de Dados SICONFI
# ==============================================================================

with tab_siconfi:

    st.subheader("📂 Fonte de Dados SICONFI")
    st.caption(
        "Consulta as planilhas do SICONFI sincronizadas do Google Drive. "
        "Use DuckDB para filtrar sem carregar o arquivo inteiro na memória."
    )

    csvs = st.session_state.csvs_disponiveis

    if st.button("🔄 Atualizar planilhas do Drive", use_container_width=False):
        with st.spinner("Sincronizando..."):
            st.session_state.csvs_disponiveis = sincronizar_e_listar_csvs()
            csvs = st.session_state.csvs_disponiveis
        st.rerun()

    if not csvs:
        st.info(
            "Nenhuma planilha disponível. Configure GDRIVE_FOLDER_URL em "
            ".streamlit/secrets.toml ou coloque arquivos CSV na pasta data/.",
            icon="ℹ️"
        )
    else:
        nomes_csv       = {nome_amigavel_csv(p): p for p in csvs}
        csv_escolhido   = st.selectbox(
            "Qual planilha deseja consultar?",
            options=list(nomes_csv.keys()),
            help="Escolha o ano da planilha SICONFI. Novas planilhas adicionadas ao Drive aparecem automaticamente."
        )
        csv_path = nomes_csv[csv_escolhido]

        st.divider()

        sa, sb = st.columns(2)
        with sa:
            opcoes_reg  = ["Todas as Regionais"] + list(REGIOES.keys())
            regioes_sel = st.multiselect(
                "Regional(is)",
                options=opcoes_reg,
                default=["Todas as Regionais"],
                help="Filtre pelos municípios de uma ou mais regionais do Planeja+."
            )
            if any("VI" in r for r in regioes_sel):
                st.markdown(
                    '<div class="aviso-vi"><strong>Regional VI:</strong> '
                    'Paraty (RJ) e Caraguatatuba/Ilhabela (SP) integram esta regional.</div>',
                    unsafe_allow_html=True
                )

        with sb:
            muns_disp = get_municipios_por_regiao(regioes_sel)
            muns_sel  = st.multiselect(
                "Município(s)",
                options=["Todos"] + muns_disp,
                default=["Todos"],
                help="Refine para um ou mais municípios específicos."
            )

        muns_filtro_sic = muns_disp if "Todos" in muns_sel or not muns_sel else muns_sel

        termo_rec_sic = st.text_input(
            "Filtrar por tipo de receita ou código",
            placeholder="Ex: royalt   |   1.7   |   participação especial   |   CFEM",
            key="termo_siconfi",
            help=(
                "Busca no campo de conta/código. Exemplos:\n\n"
                "royalt → todas as receitas de royalties\n\n"
                "1.7 → pelo código de royalties de petróleo\n\n"
                "participação → participações especiais\n\n"
                "CFEM → Compensação Financeira pela Exploração Mineral"
            )
        )

        try:
            colunas_sic = get_csv_columns(csv_path)
        except Exception as e:
            st.error(f"Erro ao ler colunas da planilha: {e}")
            colunas_sic = []

        if colunas_sic:
            with st.expander("⚙️ Ajuste de colunas (clique se os resultados estiverem errados)", expanded=False):
                st.caption(
                    "O sistema detectou as colunas automaticamente. "
                    "Ajuste aqui apenas se o resultado não corresponder ao esperado. "
                    "Para o SICONFI padrão: Município = Instituição, "
                    "Fonte de Receita = Conta, Valor = Valor."
                )
                sc1, sc2, sc3 = st.columns(3)
                with sc1:
                    col_mun_sic = st.selectbox("Coluna de Município", colunas_sic,
                                               index=acha_indice(colunas_sic, ["institui", "munic", "prefeit"]))
                    col_uf_sic  = st.selectbox("Coluna de UF", colunas_sic,
                                               index=acha_indice(colunas_sic, ["uf", "estado"]))
                with sc2:
                    col_rec_sic = st.selectbox("Coluna de Receita/Código", colunas_sic,
                                               index=acha_indice(colunas_sic, ["conta", "fonte", "descri"]))
                    col_val_sic = st.selectbox("Coluna de Valor", colunas_sic,
                                               index=acha_indice(colunas_sic, ["valor", "receita", "montante"]))
                with sc3:
                    usar_ano_sic = st.checkbox("Filtrar por ano?", value=False)
                    col_ano_sic  = None
                    anos_sel_sic = None
                    if usar_ano_sic:
                        col_ano_sic  = st.selectbox("Coluna de Ano", colunas_sic,
                                                    index=acha_indice(colunas_sic, ["ano", "exerc"]))
                        anos_disp    = get_valores_unicos(csv_path, col_ano_sic)
                        anos_sel_sic = st.multiselect("Ano(s)", options=anos_disp, default=anos_disp)

            st.info(
                f"Planilha: **{csv_escolhido}** | {len(muns_filtro_sic)} município(s) no filtro.",
                icon="ℹ️"
            )

            if st.button("🔎 Buscar na planilha SICONFI", type="primary", use_container_width=False):
                if not muns_filtro_sic:
                    st.warning("Selecione ao menos um município.")
                else:
                    with st.spinner("Consultando planilha..."):
                        df_sic = query_csv(
                            csv_path      = csv_path,
                            municipios    = muns_filtro_sic,
                            col_municipio = col_mun_sic,
                            col_receita   = col_rec_sic,
                            termo_receita = termo_rec_sic or None,
                            col_ano       = col_ano_sic,
                            anos          = anos_sel_sic
                        )

                    st.session_state.df_resultado  = df_sic
                    st.session_state.log_resultado = (
                        f"{len(df_sic):,} registros encontrados em **{csv_escolhido}** "
                        f"para {len(muns_filtro_sic)} município(s)."
                        if not df_sic.empty
                        else "Nenhum registro encontrado. Verifique os filtros aplicados."
                    )
                    st.session_state.filtros_resultado = {
                        "Fonte":         f"SICONFI / {csv_escolhido}",
                        "Regional(is)":  (
                            ", ".join(r.split(" - ")[-1] for r in regioes_sel if r != "Todas as Regionais")
                            or "Todas"
                        ),
                        "Municípios":    ", ".join(muns_filtro_sic[:5]) + ("..." if len(muns_filtro_sic) > 5 else ""),
                        "Tipo Receita":  termo_rec_sic or "Todos",
                        "Ano(s)":        ", ".join(str(a) for a in anos_sel_sic) if anos_sel_sic else "Todos"
                    }


# ==============================================================================
# Seção de Resultados
# ==============================================================================

df_res      = st.session_state.get("df_resultado")
log_res     = st.session_state.get("log_resultado")
filtros_res = st.session_state.get("filtros_resultado", {})

if df_res is not None:
    st.divider()
    st.subheader("📊 Resultado da Busca")

    if log_res:
        if not df_res.empty:
            st.success(log_res)
        else:
            st.warning(log_res)

    if not df_res.empty:

        # Métricas resumidas
        m1, m2, m3 = st.columns(3)
        m1.metric("Registros encontrados", f"{len(df_res):,}")
        m2.metric(
            "Municípios",
            df_res["_municipio"].nunique() if "_municipio" in df_res.columns
            else df_res.iloc[:, 0].nunique()
        )
        m3.metric(
            "Fontes consultadas",
            df_res["_fonte_arquivo"].nunique() if "_fonte_arquivo" in df_res.columns
            else "N/A"
        )

        st.divider()

        # ---- Copiar para planilha (destaque, expandido por padrão) ----
        with st.expander("📋 Copiar dados para sua planilha", expanded=True):
            st.markdown(
                "**Como usar:** clique na área de texto abaixo, "
                "pressione **Ctrl+A** para selecionar tudo e **Ctrl+C** para copiar. "
                "Em seguida, abra sua planilha no Excel ou Google Sheets e pressione **Ctrl+V**. "
                "Os dados vão direto para as colunas certas."
            )
            tsv = df_para_texto_copiavel(df_res)
            st.text_area(
                "Dados prontos para copiar:",
                value=tsv,
                height=180,
                label_visibility="collapsed",
                key="area_copia"
            )
            colunas_dados = [c for c in df_res.columns if not str(c).startswith("_")]
            st.caption(
                f"{len(df_res):,} linhas · {len(colunas_dados)} colunas · "
                f"Colunas de origem não incluídas nesta cópia (use o CSV para rastreabilidade completa)"
            )

        # ---- Tabela interativa ----
        st.markdown("**Visualizar tabela completa:**")
        st.dataframe(df_res, use_container_width=True, height=350)

        # ---- Downloads ----
        with st.expander("⬇️ Baixar arquivo completo", expanded=False):
            st.caption(
                "Use o download quando precisar do arquivo para enviar, arquivar "
                "ou trabalhar com mais dados do que é prático copiar manualmente."
            )

            nome_base = f"dataminer_{datetime.now().strftime('%Y%m%d_%H%M')}"

            with st.expander("⚙️ Opções", expanded=False):
                op1, op2 = st.columns(2)
                with op1:
                    incluir_prov = st.checkbox(
                        "Incluir colunas de origem",
                        value=True,
                        help="As colunas de origem (prefixo '_') mostram de qual arquivo cada dado veio."
                    )
                with op2:
                    modo_simples_exp = st.checkbox(
                        "Sem cabeçalho institucional no PDF/Word",
                        value=False
                    )

            bc1, bc2, bc3, bc4 = st.columns(4)

            with bc1:
                st.download_button(
                    label="📥 CSV",
                    data=exportar_csv(df_res, incluir_prov),
                    file_name=f"{nome_base}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    help="Abre no Excel. Melhor para grandes volumes."
                )

            with bc2:
                try:
                    st.download_button(
                        label="📊 Excel",
                        data=exportar_xlsx(df_res, filtros_res, modo_simples_exp, incluir_prov),
                        file_name=f"{nome_base}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        help="Planilha com formatação e cores."
                    )
                except Exception as e:
                    st.error(f"Excel: {e}")

            with bc3:
                try:
                    st.download_button(
                        label="📄 PDF",
                        data=exportar_pdf(df_res, filtros_res, modo_simples_exp, incluir_prov),
                        file_name=f"{nome_base}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        help="Relatório com cabeçalho institucional."
                    )
                except Exception as e:
                    st.error(f"PDF: {e}")

            with bc4:
                try:
                    st.download_button(
                        label="📝 Word",
                        data=exportar_docx(df_res, filtros_res, modo_simples_exp, incluir_prov),
                        file_name=f"{nome_base}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        help="Documento Word para incorporar em relatórios."
                    )
                except Exception as e:
                    st.error(f"Word: {e}")

            if len(df_res) >= 2000:
                st.caption("PDF exporta até 2.000 linhas e Word até 1.000. Use CSV para volumes maiores.")