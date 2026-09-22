# ==============================================================================
# exporter.py
# Geração de arquivos de saída do DataMiner Planeja+.
#
# Formatos: CSV, XLSX, PDF, DOCX
# ==============================================================================

import io
import pandas as pd
from datetime import datetime
from typing import Optional

COR_VERDE_ESCURO = (31, 78, 47)
COR_VERDE_CLARO  = (240, 247, 240)
COR_BRANCO       = (255, 255, 255)
COR_CINZA_TEXTO  = (100, 100, 100)

HEX_VERDE_ESCURO = "1F4E2F"
HEX_VERDE_CLARO  = "F0F7F0"
HEX_LARANJA_PROV = "FFF3E0"
HEX_BRANCO       = "FFFFFF"

MAX_LINHAS_PDF  = 2000
MAX_LINHAS_DOCX = 1000


# ==============================================================================
# Utilitários compartilhados
# ==============================================================================

def _separar_colunas_proveniencia(df: pd.DataFrame) -> tuple:
    colunas_dados = [c for c in df.columns if not str(c).startswith("_")]
    colunas_prov  = [c for c in df.columns if str(c).startswith("_")]
    return colunas_dados, colunas_prov


def _preparar_df_para_export(df: pd.DataFrame, incluir_proveniencia: bool = True) -> pd.DataFrame:
    colunas_dados, colunas_prov = _separar_colunas_proveniencia(df)
    if incluir_proveniencia:
        return df[colunas_dados + colunas_prov]
    return df[colunas_dados]


def _bloco_metadados(filtros: dict) -> list:
    linhas = [f"Gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}"]
    for chave, valor in filtros.items():
        if valor and str(valor).strip():
            linhas.append(f"{chave}: {valor}")
    return linhas


def _sanitizar_para_pdf(texto: str) -> str:
    """
    Substitui caracteres Unicode não suportados pela fonte Helvetica do fpdf2.
    O fpdf2 com Helvetica suporta apenas Latin-1 (ISO 8859-1).
    Comum em textos extraídos por OCR de documentos brasileiros.
    """
    mapa = {
        '\u2014': '-',    # — em dash
        '\u2013': '-',    # – en dash
        '\u2012': '-',    # ‒ figure dash
        '\u2011': '-',    # ‑ non-breaking hyphen
        '\u201c': '"',    # " aspas esquerdas
        '\u201d': '"',    # " aspas direitas
        '\u201e': '"',    # „ aspas baixas
        '\u2018': "'",    # ' aspas simples esquerda
        '\u2019': "'",    # ' aspas simples direita
        '\u2026': '...', # … reticências
        '\u2022': '*',    # • bullet
        '\u2023': '>',    # ‣ triangular bullet
        '\u25cf': '*',    # ● filled circle
        '\u25cb': 'o',    # ○ empty circle
        '\u2192': '->',   # → seta direita
        '\u2190': '<-',   # ← seta esquerda
        '\u00b0': 'o',    # ° grau
        '\u00b2': '2',    # ² superscript 2
        '\u00b3': '3',    # ³ superscript 3
        '\u00bd': '1/2',  # ½
        '\u00bc': '1/4',  # ¼
        '\u00be': '3/4',  # ¾
        '\u2265': '>=',   # ≥
        '\u2264': '<=',   # ≤
        '\u2260': '!=',   # ≠
        '\u00d7': 'x',    # × multiplicação
        '\u00f7': '/',    # ÷ divisão
        '\u00a0': ' ',    # non-breaking space
        '\u00ad': '-',    # soft hyphen
    }
    for char_unicode, substituto in mapa.items():
        texto = texto.replace(char_unicode, substituto)

    # Remove qualquer outro caractere fora de Latin-1
    return texto.encode('latin-1', errors='replace').decode('latin-1')


# ==============================================================================
# Exportação CSV
# ==============================================================================

def exportar_csv(df: pd.DataFrame, incluir_proveniencia: bool = True) -> bytes:
    df_export = _preparar_df_para_export(df, incluir_proveniencia)
    buffer = io.StringIO()
    df_export.to_csv(buffer, index=False, encoding="utf-8-sig")
    return buffer.getvalue().encode("utf-8-sig")


# ==============================================================================
# Exportação XLSX
# ==============================================================================

def exportar_xlsx(
    df: pd.DataFrame,
    filtros_aplicados: Optional[dict] = None,
    modo_simples: bool = False,
    incluir_proveniencia: bool = True,
    nome_aba: str = "DataMiner"
) -> bytes:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        raise RuntimeError("Biblioteca 'openpyxl' não instalada.")

    filtros   = filtros_aplicados or {}
    df_export = _preparar_df_para_export(df, incluir_proveniencia)

    wb = Workbook()
    ws = wb.active
    ws.title = nome_aba[:31]

    fonte_titulo        = Font(color=HEX_VERDE_ESCURO, bold=True, size=12)
    fonte_meta          = Font(size=9, color="606060", italic=True)
    fonte_header_tabela = Font(color="FFFFFF", bold=True, size=9)
    fonte_dados         = Font(size=9)
    fonte_prov          = Font(size=8, color="5D4037", italic=True)

    fill_verde_escuro = PatternFill("solid", fgColor=HEX_VERDE_ESCURO)
    fill_verde_claro  = PatternFill("solid", fgColor=HEX_VERDE_CLARO)
    fill_laranja_prov = PatternFill("solid", fgColor=HEX_LARANJA_PROV)
    fill_branco       = PatternFill("solid", fgColor=HEX_BRANCO)
    fill_cabecalho    = PatternFill("solid", fgColor="E8F5E9")

    borda = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC")
    )

    num_colunas = len(df_export.columns)
    linha_atual = 1

    if not modo_simples:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(num_colunas, 4))
        c = ws.cell(row=1, column=1, value="DataMiner Planeja+  |  Associação Raízes")
        c.font = fonte_titulo
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = fill_cabecalho
        ws.row_dimensions[1].height = 22
        linha_atual = 2

        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(num_colunas, 4))
        c2 = ws.cell(row=2, column=1, value="Extração e Cruzamento de Dados Orçamentários Municipais")
        c2.font = fonte_meta
        c2.alignment = Alignment(horizontal="center")
        linha_atual = 3

        for texto in _bloco_metadados(filtros):
            ws.merge_cells(start_row=linha_atual, start_column=1, end_row=linha_atual, end_column=max(num_colunas, 4))
            cm = ws.cell(row=linha_atual, column=1, value=texto)
            cm.font = fonte_meta
            cm.alignment = Alignment(horizontal="left", indent=1)
            linha_atual += 1

        linha_atual += 1

    linha_header = linha_atual

    for col_idx, nome_col in enumerate(df_export.columns, start=1):
        c = ws.cell(row=linha_header, column=col_idx)
        c.value = str(nome_col).lstrip("_").replace("_", " ").title()
        c.font = fonte_header_tabela
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borda
        c.fill = PatternFill("solid", fgColor="795548") if str(nome_col).startswith("_") else fill_verde_escuro

    ws.row_dimensions[linha_header].height = 30
    linha_atual = linha_header + 1

    for i, (_, linha_df) in enumerate(df_export.iterrows()):
        fill_linha = fill_verde_claro if i % 2 == 0 else fill_branco

        for col_idx, nome_col in enumerate(df_export.columns, start=1):
            c = ws.cell(row=linha_atual, column=col_idx)
            valor = linha_df[nome_col]
            c.value = "" if pd.isna(valor) else valor
            c.border = borda
            c.alignment = Alignment(vertical="center")

            if str(nome_col).startswith("_"):
                c.fill = fill_laranja_prov
                c.font = fonte_prov
            else:
                c.fill = fill_linha
                c.font = fonte_dados

        ws.row_dimensions[linha_atual].height = 16
        linha_atual += 1

    ws.cell(row=linha_atual, column=1, value=f"Total de registros: {len(df_export):,}").font = fonte_meta
    ws.freeze_panes = ws.cell(row=linha_header + 1, column=1)

    for col_idx, nome_col in enumerate(df_export.columns, start=1):
        letra = get_column_letter(col_idx)
        valores_col = df_export[nome_col].astype(str)
        largura = max(
            len(str(nome_col).lstrip("_")) + 4,
            int(valores_col.str.len().max()) if not valores_col.empty else 10
        )
        ws.column_dimensions[letra].width = min(max(largura, 10), 50)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ==============================================================================
# Exportação PDF
# ==============================================================================

def exportar_pdf(
    df: pd.DataFrame,
    filtros_aplicados: Optional[dict] = None,
    modo_simples: bool = False,
    incluir_proveniencia: bool = True,
    max_linhas: int = MAX_LINHAS_PDF
) -> bytes:
    try:
        from fpdf import FPDF
    except ImportError:
        raise RuntimeError("Biblioteca 'fpdf2' não instalada.")

    filtros   = filtros_aplicados or {}
    df_export = _preparar_df_para_export(df, incluir_proveniencia)

    if len(df_export) > max_linhas:
        df_export = df_export.head(max_linhas)

    class PDFRelatorio(FPDF):
        def header(self):
            if modo_simples:
                self.set_font("Helvetica", "I", 8)
                self.set_text_color(*COR_CINZA_TEXTO)
                self.cell(0, 6, "DataMiner Planeja+ / Associacao Raizes", ln=True, align="C")
                self.ln(2)
                return

            self.set_font("Helvetica", "B", 14)
            self.set_text_color(*COR_VERDE_ESCURO)
            self.cell(0, 8, "Associacao Raizes  |  Programa Planeja+", ln=True, align="C")

            self.set_font("Helvetica", "", 10)
            self.set_text_color(60, 60, 60)
            self.cell(0, 5, "DataMiner Planeja+ - Extracao e Cruzamento de Dados Orcamentarios",
                      ln=True, align="C")

            self.set_draw_color(*COR_VERDE_ESCURO)
            self.set_line_width(0.5)
            y = self.get_y() + 2
            self.line(10, y, 287, y)
            self.ln(5)

            self.set_font("Helvetica", "I", 8)
            self.set_text_color(*COR_CINZA_TEXTO)
            for linha in _bloco_metadados(filtros):
                self.cell(0, 4, _sanitizar_para_pdf(linha), ln=True)
            self.ln(3)

        def footer(self):
            if modo_simples:
                return
            self.set_y(-10)
            self.set_font("Helvetica", "I", 7)
            self.set_text_color(160, 160, 160)
            self.cell(0, 5,
                      f"Pagina {self.page_no()}  |  DataMiner Planeja+ / Associacao Raizes",
                      align="C")

    pdf = PDFRelatorio(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_margins(10, 10, 10)
    pdf.add_page()

    num_colunas  = len(df_export.columns)
    largura_util = 277
    largura_col  = max(18, min(largura_util // num_colunas, 70))

    # Cabeçalho da tabela
    pdf.set_font("Helvetica", "B", 7)
    for nome_col in df_export.columns:
        if str(nome_col).startswith("_"):
            pdf.set_fill_color(121, 85, 72)
        else:
            pdf.set_fill_color(*COR_VERDE_ESCURO)
        pdf.set_text_color(*COR_BRANCO)
        nome_exib = _sanitizar_para_pdf(str(nome_col).lstrip("_").replace("_", " ")[:20])
        pdf.cell(largura_col, 6, nome_exib, border=1, fill=True, align="C")
    pdf.ln()

    # Linhas de dados
    pdf.set_font("Helvetica", "", 6)
    for i, (_, linha_df) in enumerate(df_export.iterrows()):
        for nome_col in df_export.columns:
            valor_raw = str(linha_df[nome_col]) if pd.notna(linha_df[nome_col]) else ""
            valor     = _sanitizar_para_pdf(valor_raw[:28])

            if str(nome_col).startswith("_"):
                pdf.set_fill_color(255, 243, 224)
                pdf.set_text_color(93, 64, 55)
            elif i % 2 == 0:
                pdf.set_fill_color(*COR_VERDE_CLARO)
                pdf.set_text_color(30, 30, 30)
            else:
                pdf.set_fill_color(*COR_BRANCO)
                pdf.set_text_color(30, 30, 30)

            pdf.cell(largura_col, 5, valor, border=1, fill=True, align="L")
        pdf.ln()

    pdf.set_font("Helvetica", "I", 7)
    pdf.set_text_color(*COR_CINZA_TEXTO)
    pdf.ln(3)
    nota = f"Total de registros exportados: {len(df_export):,}"
    if len(df) > max_linhas:
        nota += f"  (limitado a {max_linhas:,} linhas para este PDF)"
    pdf.cell(0, 5, _sanitizar_para_pdf(nota), ln=True)

    return bytes(pdf.output())


# ==============================================================================
# Exportação DOCX
# ==============================================================================

def exportar_docx(
    df: pd.DataFrame,
    filtros_aplicados: Optional[dict] = None,
    modo_simples: bool = False,
    incluir_proveniencia: bool = True,
    max_linhas: int = MAX_LINHAS_DOCX
) -> bytes:
    try:
        from docx import Document
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
    except ImportError:
        raise RuntimeError("Biblioteca 'python-docx' não instalada.")

    filtros   = filtros_aplicados or {}
    df_export = _preparar_df_para_export(df, incluir_proveniencia)

    if len(df_export) > max_linhas:
        df_export = df_export.head(max_linhas)

    doc = Document()

    for secao in doc.sections:
        secao.left_margin   = Cm(1.5)
        secao.right_margin  = Cm(1.5)
        secao.top_margin    = Cm(2.0)
        secao.bottom_margin = Cm(2.0)
        secao.orientation   = 1

    if not modo_simples:
        p_titulo = doc.add_heading("DataMiner Planeja+  |  Associação Raízes", level=1)
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p_titulo.runs:
            run.font.color.rgb = RGBColor(*COR_VERDE_ESCURO)

        p_sub = doc.add_paragraph("Extração e Cruzamento de Dados Orçamentários Municipais")
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p_sub.runs:
            run.font.size      = Pt(10)
            run.font.color.rgb = RGBColor(80, 80, 80)

        doc.add_paragraph()

        for linha_meta in _bloco_metadados(filtros):
            p = doc.add_paragraph(linha_meta)
            if p.runs:
                p.runs[0].font.size      = Pt(9)
                p.runs[0].font.color.rgb = RGBColor(*COR_CINZA_TEXTO)
                p.runs[0].font.italic    = True

        doc.add_paragraph()

    num_colunas = len(df_export.columns)
    tabela = doc.add_table(rows=1, cols=num_colunas)
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabela.style     = "Table Grid"

    linha_header = tabela.rows[0]
    for col_idx, nome_col in enumerate(df_export.columns):
        celula    = linha_header.cells[col_idx]
        nome_exib = str(nome_col).lstrip("_").replace("_", " ").title()
        celula.text = nome_exib

        par = celula.paragraphs[0]
        run = par.runs[0] if par.runs else par.add_run(nome_exib)
        run.font.bold      = True
        run.font.size      = Pt(8)
        run.font.color.rgb = RGBColor(*COR_BRANCO)
        par.alignment      = WD_ALIGN_PARAGRAPH.CENTER

        cor_hex = "795548" if str(nome_col).startswith("_") else HEX_VERDE_ESCURO
        _cor_celula_docx(celula, cor_hex)

    for i, (_, linha_df) in enumerate(df_export.iterrows()):
        linha_tabela = tabela.add_row()

        for col_idx, nome_col in enumerate(df_export.columns):
            celula = linha_tabela.cells[col_idx]
            valor  = str(linha_df[nome_col]) if pd.notna(linha_df[nome_col]) else ""
            celula.text = valor

            par = celula.paragraphs[0]
            run = par.runs[0] if par.runs else par.add_run(valor)
            run.font.size = Pt(7)

            if str(nome_col).startswith("_"):
                _cor_celula_docx(celula, "FFF3E0")
                run.font.color.rgb = RGBColor(93, 64, 55)
                run.font.italic    = True
            elif i % 2 == 0:
                _cor_celula_docx(celula, HEX_VERDE_CLARO)

    doc.add_paragraph()
    nota = f"Total de registros: {len(df_export):,}"
    if len(df) > max_linhas:
        nota += f"  (limitado a {max_linhas:,} linhas)"
    p_nota = doc.add_paragraph(nota)
    if p_nota.runs:
        p_nota.runs[0].font.size      = Pt(8)
        p_nota.runs[0].font.italic    = True
        p_nota.runs[0].font.color.rgb = RGBColor(*COR_CINZA_TEXTO)

    if not modo_simples:
        rodape = doc.sections[0].footer
        par_rodape = rodape.paragraphs[0]
        par_rodape.text = (
            f"DataMiner Planeja+ / Associação Raízes  |  "
            f"Gerado em {datetime.now().strftime('%d/%m/%Y')}"
        )
        par_rodape.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in par_rodape.runs:
            run.font.size      = Pt(8)
            run.font.color.rgb = RGBColor(160, 160, 160)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _cor_celula_docx(celula, cor_hex: str):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tc_pr = celula._tc.get_or_add_tcPr()
    shd   = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  cor_hex.upper())
    tc_pr.append(shd)