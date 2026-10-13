"""
Exportador corporativo para planilhas Excel (.xlsx) e relatórios PDF Paisagem (.pdf).
Aplica regras financeiras cumulativas de formatação e proteção contra PermissionError.
"""
import os
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from core.data_processor import formatar_moeda_float, obter_estilo_financeiro
from core.logger_config import obter_logger

logger = obter_logger("RelatorioExporter")

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfgen import canvas
    REPORTLAB_DISPONIVEL = True
except ImportError:
    REPORTLAB_DISPONIVEL = False


if REPORTLAB_DISPONIVEL:
    class NumberedCanvas(canvas.Canvas):
        """Canvas dinâmico de 2 passos para numeração de páginas 'Página X de Y'."""
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._saved_page_states = []

        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            num_paginas = len(self._saved_page_states)
            for state in self._saved_page_states:
                self.__dict__.update(state)
                self.draw_page_number(num_paginas)
                canvas.Canvas.showPage(self)
            canvas.Canvas.save(self)

        def draw_page_number(self, total_paginas):
            self.saveState()
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.HexColor("#64748B"))
            largura, altura = landscape(A4)
            data_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            texto_esq = f"Relatório Gerado em: {data_atual} | Sistema de Gestão de Frotas (CE)"
            self.drawString(36, 20, texto_esq)
            texto_dir = f"Página {self._pageNumber} de {total_paginas}"
            self.drawRightString(largura - 36, 20, texto_dir)
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.5)
            self.line(36, 32, largura - 36, 32)
            self.restoreState()
else:
    NumberedCanvas = None


class FrotaRelatorioExporter:
    """Exportador para planilhas Excel formatadas e relatórios PDF em layout Paisagem."""

    @staticmethod
    def exportar_excel(df: pd.DataFrame, caminho_arquivo: str) -> str:
        """
        Gera e formata a planilha Excel de saída com estilo corporativo.
        Trata PermissionError salvando com timestamp para evitar perda de dados.
        Retorna o caminho efetivo do arquivo gravado.
        """
        caminho_efetivo = caminho_arquivo
        pasta_dir = os.path.dirname(os.path.abspath(caminho_efetivo))
        if pasta_dir:
            os.makedirs(pasta_dir, exist_ok=True)
        try:
            df.to_excel(caminho_efetivo, index=False, engine="openpyxl")
        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pasta, nome = os.path.split(caminho_arquivo)
            raiz, ext = os.path.splitext(nome)
            caminho_efetivo = os.path.join(pasta, f"{raiz}_{timestamp}{ext}")
            df.to_excel(caminho_efetivo, index=False, engine="openpyxl")

        wb = load_workbook(caminho_efetivo)
        ws = wb.active

        cor_cabecalho = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        fonte_linhas = Font(name="Calibri", size=10)
        alinhamento_centro = Alignment(horizontal="center", vertical="center")
        alinhamento_direita = Alignment(horizontal="right", vertical="center")
        alinhamento_esquerda = Alignment(horizontal="left", vertical="center")
        borda_fina = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        for cell in ws[1]:
            cell.fill = cor_cabecalho
            cell.font = fonte_cabecalho
            cell.alignment = alinhamento_centro

        cor_zebra = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
        col_lic_idx = None
        col_ipva_idx = None
        col_multas_idx = None
        col_total_idx = None
        col_status_idx = None

        for col_idx, col_name in enumerate(df.columns, start=1):
            cn = str(col_name).strip().upper()
            if "LICENCIAMENTO" in cn:
                col_lic_idx = col_idx
            elif "IPVA" in cn:
                col_ipva_idx = col_idx
            elif "MULTA" in cn:
                col_multas_idx = col_idx
            elif "TOTAL" in cn:
                col_total_idx = col_idx
            elif "STATUS" in cn:
                col_status_idx = col_idx

        for row_idx in range(2, ws.max_row + 1):
            is_zebra = (row_idx % 2 == 0)
            fill_padrao = cor_zebra if is_zebra else None

            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if fill_padrao:
                    cell.fill = fill_padrao
                cell.font = fonte_linhas
                cell.border = borda_fina
                val_txt = str(cell.value or "")

                is_col_financeira = col_idx in [col_lic_idx, col_ipva_idx, col_multas_idx, col_total_idx]
                val_txt_upper = val_txt.upper()

                if is_col_financeira:
                    if val_txt_upper == "PAGO":
                        cell.alignment = alinhamento_centro
                        # Fundo verde pastel suave e fonte verde escuro elegante
                        cell.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
                        cell.font = Font(name="Calibri", size=10, bold=True, color="166534")
                    elif "RESTRIÇÃO" in val_txt_upper or "RESTRICAO" in val_txt_upper:
                        cell.alignment = alinhamento_centro
                        # Fundo laranja pastel suave e fonte laranja escuro elegante
                        cell.fill = PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid")
                        cell.font = Font(name="Calibri", size=10, bold=True, color="9A3412")
                    else:
                        cell.alignment = alinhamento_direita
                        val_float = formatar_moeda_float(val_txt)
                        is_lic = (col_idx == col_lic_idx)
                        is_tot = (col_idx == col_total_idx)

                        bold, underline, cor_hex = obter_estilo_financeiro(val_float, is_licenciamento=is_lic)
                        if is_tot:
                            bold = True
                            if val_float > 0:
                                cor_hex = "#15803D"

                        sublinhado_estilo = 'single' if underline else None
                        cor_fonte = cor_hex.replace("#", "") if cor_hex else '000000'

                        cell.font = Font(
                            name="Calibri",
                            size=10,
                            bold=bold,
                            underline=sublinhado_estilo,
                            color=cor_fonte
                        )
                elif col_idx == col_status_idx:
                    cell.alignment = alinhamento_centro
                    if "RESTRIÇÃO" in val_txt_upper or "RESTRICAO" in val_txt_upper:
                        cell.fill = PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid")
                        cell.font = Font(name="Calibri", size=10, bold=True, color="9A3412")
                    elif "SEM DÉBITO" in val_txt_upper:
                        cell.font = Font(name="Calibri", size=10, bold=True, color="166534")
                    elif "COM DÉBITO" in val_txt_upper:
                        cell.font = Font(name="Calibri", size=10, bold=True, color="991B1B")
                    elif "FALHA" in val_txt_upper:
                        cell.font = Font(name="Calibri", size=10, bold=True, color="B45309")
                elif "PLACA" in str(ws.cell(row=1, column=col_idx).value or "").upper():
                    cell.alignment = alinhamento_centro
                    cell.font = Font(name="Calibri", size=10, bold=True)
                elif "RENAVAM" in str(ws.cell(row=1, column=col_idx).value or "").upper():
                    cell.alignment = alinhamento_centro
                else:
                    cell.alignment = alinhamento_esquerda

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or '')
                if len(val) > max_len:
                    max_len = len(val)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        ws.freeze_panes = "A2"
        wb.save(caminho_efetivo)
        logger.info(f"Relatório Excel gravado com sucesso: {caminho_efetivo}")
        return caminho_efetivo

    @staticmethod
    def exportar_pdf(
        df: pd.DataFrame,
        caminho_arquivo: str,
        totais: dict,
        filtros_ativos: list[str]
    ) -> str:
        """
        Gera relatório executivo em PDF no formato Paisagem (A4) com ReportLab.
        Aplica regras financeiras cumulativas de estilo e proteção de sobreescrita.
        """
        if not REPORTLAB_DISPONIVEL:
            raise RuntimeError("ReportLab não está instalado no ambiente.")

        caminho_efetivo = caminho_arquivo
        pasta_dir = os.path.dirname(os.path.abspath(caminho_efetivo))
        if pasta_dir:
            os.makedirs(pasta_dir, exist_ok=True)

        try:
            with open(caminho_efetivo, "ab"):
                pass
        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pasta, nome = os.path.split(caminho_arquivo)
            raiz, ext = os.path.splitext(nome)
            caminho_efetivo = os.path.join(pasta, f"{raiz}_{timestamp}{ext}")

        doc = SimpleDocTemplate(
            caminho_efetivo,
            pagesize=landscape(A4),
            leftMargin=36,
            rightMargin=36,
            topMargin=36,
            bottomMargin=45
        )

        styles = getSampleStyleSheet()
        estilo_titulo = ParagraphStyle('TituloDoc', parent=styles['Heading1'], fontSize=15, leading=18, textColor=colors.HexColor("#1E3A8A"), fontName="Helvetica-Bold", spaceAfter=2)
        estilo_sub = ParagraphStyle('SubDoc', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=colors.HexColor("#64748B"), spaceAfter=8)
        estilo_th = ParagraphStyle('Th', parent=styles['Normal'], fontSize=7.5, leading=9.5, fontName="Helvetica-Bold", textColor=colors.white, alignment=1)
        estilo_td = ParagraphStyle('Td', parent=styles['Normal'], fontSize=7, leading=9, textColor=colors.HexColor("#1F2937"), alignment=1)
        estilo_td_left = ParagraphStyle('TdLeft', parent=styles['Normal'], fontSize=7, leading=9, textColor=colors.HexColor("#1F2937"), alignment=0)
        estilo_td_right = ParagraphStyle('TdRight', parent=styles['Normal'], fontSize=7, leading=9, textColor=colors.HexColor("#1F2937"), alignment=2)

        elementos = []
        elementos.append(Paragraph("RELATÓRIO CONSOLIDADO DE DÉBITOS VEICULARES", estilo_titulo))
        elementos.append(Paragraph(
            f"Portais Consultados: Detran-CE & Sefaz-CE | Filtros Ativos: {', '.join(filtros_ativos)} | Total de Veículos: {len(df)}",
            estilo_sub
        ))

        # CARDS DE TOTAIS
        dados_cards = [[
            Paragraph(f"<b>VEÍCULOS</b><br/><font size=10 color='#1E3A8A'><b>{len(df)}</b></font>", estilo_td),
            Paragraph(f"<b>LICENCIAMENTO</b><br/><font size=10 color='#B45309'><b>{totais.get('lic', 'R$ 0,00')}</b></font>", estilo_td),
            Paragraph(f"<b>IPVA EM ABERTO</b><br/><font size=10 color='#B91C1C'><b>{totais.get('ipva', 'R$ 0,00')}</b></font>", estilo_td),
            Paragraph(f"<b>MULTAS DETRAN</b><br/><font size=10 color='#7C2D12'><b>{totais.get('multas', 'R$ 0,00')}</b></font>", estilo_td),
            Paragraph(f"<b>TOTAL A PAGAR</b><br/><font size=10 color='#15803D'><b>{totais.get('total', 'R$ 0,00')}</b></font>", estilo_td)
        ]]
        tabela_cards = Table(dados_cards, colWidths=[150]*5)
        tabela_cards.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F8FAFC")),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#CBD5E1")),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ]))
        elementos.append(tabela_cards)
        elementos.append(Spacer(1, 10))

        # TABELA DE DADOS
        colunas_df = list(df.columns)
        cabecalho = [Paragraph(str(col), estilo_th) for col in colunas_df]
        tabela_linhas = [cabecalho]

        largura_total_util = 770
        largura_por_col = largura_total_util / max(len(colunas_df), 1)
        larguras = [largura_por_col] * len(colunas_df)

        for _, row in df.iterrows():
            linha_celulas = []
            for col_nome in colunas_df:
                val = str(row[col_nome] if pd.notna(row[col_nome]) else "")
                val_upper = val.upper()
                col_upper = str(col_nome).upper()

                if "STATUS" in col_upper:
                    if "RESTRIÇÃO" in val_upper or "RESTRICAO" in val_upper:
                        txt_p = f"<font color='#C2410C'><b>{val}</b></font>"
                    elif "SEM DÉBITO" in val_upper:
                        txt_p = f"<font color='#166534'><b>{val}</b></font>"
                    elif "COM DÉBITO" in val_upper:
                        txt_p = f"<font color='#991B1B'><b>{val}</b></font>"
                    elif "FALHA" in val_upper:
                        txt_p = f"<font color='#B45309'><b>{val}</b></font>"
                    else:
                        txt_p = val
                    linha_celulas.append(Paragraph(txt_p, estilo_td))
                elif any(c in col_upper for c in ["LICENCIAMENTO", "IPVA", "MULTA", "TOTAL"]):
                    if val_upper == "PAGO":
                        txt_p = "<font color='#166534'><b>PAGO</b></font>"
                        linha_celulas.append(Paragraph(txt_p, estilo_td))
                    elif "RESTRIÇÃO" in val_upper or "RESTRICAO" in val_upper:
                        txt_p = "<font color='#C2410C'><b>Restrição</b></font>"
                        linha_celulas.append(Paragraph(txt_p, estilo_td))
                    else:
                        val_float = formatar_moeda_float(val)
                        is_lic = "LICENCIAMENTO" in col_upper
                        is_tot = "TOTAL" in col_upper
                        bold, underline, cor_hex = obter_estilo_financeiro(val_float, is_licenciamento=is_lic)
                        if is_tot:
                            bold = True
                            if val_float > 0:
                                cor_hex = "#15803D"

                        conteudo = val
                        if underline:
                            conteudo = f"<u>{conteudo}</u>"
                        if bold:
                            conteudo = f"<b>{conteudo}</b>"
                        if cor_hex:
                            conteudo = f"<font color='{cor_hex}'>{conteudo}</font>"

                        linha_celulas.append(Paragraph(conteudo, estilo_td_right))
                elif "PLACA" in col_upper:
                    linha_celulas.append(Paragraph(f"<b>{val}</b>", estilo_td))
                else:
                    linha_celulas.append(Paragraph(val, estilo_td_left))
            tabela_linhas.append(linha_celulas)

        tabela_dados = Table(tabela_linhas, colWidths=larguras, repeatRows=1)
        tabela_dados.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1F497D")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F8FAFC")]),
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
            ('INNERGRID', (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        elementos.append(tabela_dados)

        doc.build(elementos, canvasmaker=NumberedCanvas)
        logger.info(f"Relatório PDF gerado com sucesso: {caminho_efetivo}")
        return caminho_efetivo
