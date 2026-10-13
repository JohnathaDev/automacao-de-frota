"""
========================================================================================
Aplicativo GUI Profissional: Consulta de Frota - Detran/Sefaz CE (Versão 2.1)
========================================================================================
Sistema corporativo de automação e web scraping com Interface Gráfica nativa (Tkinter),
resiliência dinâmica de dados (Pandas), extração de Licenciamento, IPVA e Multas,
e geração de relatórios executivos em Excel (.xlsx) e PDF Paisagem (.pdf).

Melhorias da Versão 2.1:
1. Prevenção de Falsos Negativos: diferenciação clara entre veículos regulares (R$ 0,00)
   e erros/timeouts de conexão (status 'Falha na Consulta' destacado na interface).
2. Proteção contra Arquivo Excel/PDF Aberto (PermissionError): detecção de bloqueio e
   salvamento automático com timestamp para evitar perda de dados.
3. Garantia de 11 Dígitos no RENAVAM: preenchimento automático de zeros à esquerda (zfill).
4. Sanitização e Validação Prévia: remoção de linhas em branco e validação de padrão
   de placas (Mercosul e Tradicional) antes de iniciar as consultas.
5. Salvamento Incremental (Checkpoint): gravação contínua do progresso para proteger
   contra quedas de energia, cancelamentos prematuros ou falhas de rede.
6. Pasta de Destino Personalizada: botão e seletor para definir onde salvar os relatórios.
7. Reconsulta Inteligente: menu de contexto e botão rápido para reprocessar apenas
   veículos selecionados ou que apresentaram falha/timeout transitório.
========================================================================================
"""

import os
import sys
import re
import time
import subprocess
import threading
from datetime import datetime
import pandas as pd

# Manipulação e Formatação Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook

# Geração de PDF Executivo em Paisagem (ReportLab)
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfgen import canvas
    REPORTLAB_DISPONIVEL = True
except ImportError:
    REPORTLAB_DISPONIVEL = False

# Interface Gráfica
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Automação Web (Selenium)
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager


# Padrão oficial de placas veiculares brasileiras (Mercosul e Tradicional)
PADRAO_PLACA = re.compile(r"^[A-Z]{3}[0-9][A-Z0-9][0-9]{2}$|^[A-Z]{3}[0-9]{4}$")


# ======================================================================================
# SUPRESSÃO DE CONSOLE NO WINDOWS
# ======================================================================================

def ocultar_console():
    """
    Oculta a janela preta de console do Python no Windows, mantendo apenas
    a interface gráfica visível para o usuário final.
    """
    if sys.platform == "win32" and "--console" not in sys.argv:
        try:
            import ctypes
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE = 0
        except Exception:
            pass


# ======================================================================================
# UTILITÁRIOS FINANCEIROS
# ======================================================================================

def formatar_moeda_float(valor_str: str) -> float:
    """Converte valores monetários textuais (ex: 'R$ 188,96' ou '188.96') em float 188.96."""
    try:
        if not valor_str:
            return 0.0
        val_lower = str(valor_str).lower().strip()
        termos_ignorar = [
            "-", "n/a", "não consultado", "falha", "erro", "aguardando",
            "inválid", "não localizado", "timeout", "..."
        ]
        if any(termo in val_lower for termo in termos_ignorar):
            return 0.0
        limpo = val_lower.replace("r$", "").replace(".", "").replace(",", ".").strip()
        return float(limpo)
    except Exception:
        return 0.0


def formatar_float_moeda(valor: float) -> str:
    """Converte um float monetário em formato brasileiro: R$ 1.234,56."""
    try:
        formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return formatado
    except Exception:
        return "R$ 0,00"


def obter_estilo_financeiro(col_nome: str, valor_str: str) -> dict:
    """
    Aplica as regras corporativas de destaque financeiro (Excel e PDF):
    - Valor > R$ 1.000,00: Negrito (bold)
    - Valor > R$ 1.500,00: Negrito + Sublinhado (bold + underline)
    - Licenciamento > R$ 188,96 (multa/taxa embutida): Fonte Vermelha (#DC2626)
    Todas as regras são cumulativas caso haja sobreposição de características.
    """
    val_num = formatar_moeda_float(valor_str)
    col_low = str(col_nome).lower()
    val_clean = str(valor_str).strip()

    is_fin = val_clean.startswith("R$") or any(
        k in col_low for k in ["licenciamento", "ipva", "multas", "valor", "total"]
    )
    is_bold = is_fin and (val_num > 1000.00)
    is_underline = is_fin and (val_num > 1500.00)
    is_red = is_fin and ("licenciamento" in col_low) and (val_num > 188.96)

    return {
        "is_fin": is_fin,
        "val_num": val_num,
        "is_bold": is_bold,
        "is_underline": is_underline,
        "is_red": is_red
    }


# ======================================================================================
# 1. PROCESSAMENTO RESILIENTE DE DADOS (PANDAS)
# ======================================================================================

class FrotaDataProcessor:
    """
    Gerencia a leitura resiliente de planilhas de entrada, validação de dados
    e a montagem do DataFrame final, preservando colunas originais do cliente.
    """

    @staticmethod
    def localizar_colunas_chave(df: pd.DataFrame) -> tuple[str, str]:
        """
        Localiza dinamicamente as colunas correspondentes à Placa e ao RENAVAM,
        ignorando distinções de maiúsculas, espaços e variações comuns.
        """
        col_placa = None
        col_renavam = None

        for col in df.columns:
            nome_norm = str(col).strip().lower()
            if "placa" in nome_norm and not col_placa:
                col_placa = col
            if "renavam" in nome_norm and not col_renavam:
                col_renavam = col

        if not col_placa or not col_renavam:
            faltantes = []
            if not col_placa:
                faltantes.append("'Placa'")
            if not col_renavam:
                faltantes.append("'RENAVAM'")
            raise ValueError(
                f"A planilha selecionada não possui as seguintes colunas obrigatórias: {', '.join(faltantes)}.\n"
                f"Colunas encontradas: {list(df.columns)}"
            )

        return col_placa, col_renavam

    @classmethod
    def carregar_planilha(cls, caminho_arquivo: str) -> tuple[pd.DataFrame, str, str]:
        """
        Carrega a planilha Excel assegurando que:
        - Linhas vazias sejam removidas
        - Valores de RENAVAM não percam zeros à esquerda (zfill 11)
        - Placas sejam higienizadas
        """
        df = pd.read_excel(caminho_arquivo, dtype=str)
        if len(df) == 0:
            raise ValueError("O arquivo Excel selecionado não contém nenhum registro.")

        col_placa, col_renavam = cls.localizar_colunas_chave(df)

        # 1. Remove linhas onde placa ou renavam são nulos
        df = df.dropna(subset=[col_placa, col_renavam]).copy()

        # 2. Higienização de Placa (remove hífens, espaços e transforma em maiúsculo)
        df[col_placa] = (
            df[col_placa]
            .astype(str)
            .str.strip()
            .str.upper()
            .str.replace("-", "", regex=False)
            .str.replace(" ", "", regex=False)
        )

        # 3. Higienização de RENAVAM (preserva 11 dígitos com zeros à esquerda)
        def sanitizar_renavam(val):
            val_str = str(val).strip()
            if val_str.endswith(".0"):
                val_str = val_str[:-2]
            digitos = "".join(c for c in val_str if c.isdigit())
            if digitos and len(digitos) <= 11:
                return digitos.zfill(11)
            return digitos or val_str

        df[col_renavam] = df[col_renavam].apply(sanitizar_renavam)

        # 4. Remove linhas com campos 'NAN', 'NONE', vazios ou hífens
        df = df[
            ~df[col_placa].isin(["NAN", "NONE", "", "N/A", "-", "NULL"]) &
            ~df[col_renavam].isin(["NAN", "NONE", "", "N/A", "-", "NULL"])
        ].copy()

        if len(df) == 0:
            raise ValueError("Nenhum registro com Placa e RENAVAM válidos foi encontrado na planilha.")

        df = df.reset_index(drop=True)
        return df, col_placa, col_renavam

    @staticmethod
    def consolidar_resultados(
        df_original: pd.DataFrame,
        resultados_lic: list[str],
        resultados_ipva: list[str],
        resultados_multas: list[str],
        qtd_processada: int
    ) -> pd.DataFrame:
        """
        Constrói o DataFrame final preservando todas as colunas originais do usuário
        e adicionando as colunas oficiais de débitos e situação.
        Evita falsos negativos: identifica 'Falha na Consulta' e 'Placa Inválida'.
        """
        df_final = df_original.iloc[:qtd_processada].copy()

        df_final["Valor_Licenciamento"] = resultados_lic[:qtd_processada]
        df_final["Valor_IPVA"] = resultados_ipva[:qtd_processada]
        df_final["Valor_Multas"] = resultados_multas[:qtd_processada]

        # Determinação do Status consolidado
        status_lista = []
        for i in range(qtd_processada):
            lic = str(resultados_lic[i])
            ipva = str(resultados_ipva[i])
            multas = str(resultados_multas[i])

            tem_falha = any(
                termo in v.lower()
                for v in [lic, ipva, multas]
                for termo in ["falha", "erro", "timeout", "inválid", "não localizado"]
            )
            soma = (
                formatar_moeda_float(lic) +
                formatar_moeda_float(ipva) +
                formatar_moeda_float(multas)
            )

            if tem_falha:
                status_lista.append("Falha na Consulta")
            elif soma > 0:
                status_lista.append("Com Débitos Pendentes")
            else:
                status_lista.append("Sem Débitos")

        df_final["Status"] = status_lista
        df_final["Data_Consulta"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        return df_final


# ======================================================================================
# 2. MOTOR DE WEB SCRAPING OTIMIZADO (SELENIUM / DETRAN-CE / SEFAZ-CE)
# ======================================================================================

class FrotaScraperEngine:
    """
    Executa a automação de consultas no Detran-CE e Sefaz-CE com alta resiliência,
    sessão única unificada para Licenciamento e Multas e tratamento explícito de falhas.
    """

    def __init__(self, headless: bool = True):
        self.headless = headless
        self.driver: webdriver.Chrome | None = None
        self.wait: WebDriverWait | None = None

    def iniciar(self):
        """Inicializa o WebDriver com configurações silenciosas e robustas."""
        chrome_options = Options()
        if self.headless:
            chrome_options.add_argument("--headless=new")

        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        )

        os.environ["WDM_LOG"] = "0"
        os.environ["WDM_LOG_LEVEL"] = "0"

        service = Service(ChromeDriverManager().install())
        if sys.platform == "win32":
            service.creation_flags = subprocess.CREATE_NO_WINDOW

        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.wait = WebDriverWait(self.driver, 15)

    def fechar(self):
        """Encerra a instância do navegador de forma segura."""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None
            self.wait = None

    def _clicar_seguro(self, elemento) -> None:
        """Executa o clique com scrollIntoView e fallback JavaScript."""
        try:
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
            elemento.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", elemento)

    def consultar_detran(
        self,
        placa: str,
        renavam: str,
        consultar_lic: bool,
        consultar_multas: bool
    ) -> tuple[str, str]:
        """
        Consulta Detran-CE na mesma sessão:
        Extrai Licenciamento e/ou Multas com navegação unificada para velocidade máxima.
        Diferencia veículos regulares (R$ 0,00) de falhas/erros de conexão.
        """
        val_lic = "-" if not consultar_lic else "Falha Detran"
        val_multas = "-" if not consultar_multas else "Falha Detran"

        if not (consultar_lic or consultar_multas):
            return "-", "-"

        url_detran = "https://sistemas.detran.ce.gov.br/central/#veiculo"
        try:
            self.driver.get(url_detran)

            # 1. Clica no menu lateral de Licenciamento (ou Taxas/Multas)
            menu_btn = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Licenciamento') or contains(text(), 'Taxas / Multas')]"))
            )
            self._clicar_seguro(menu_btn)

            # 2. Localiza inputs visíveis (Placa e RENAVAM)
            self.wait.until(lambda d: len([i for i in d.find_elements(By.TAG_NAME, "input") if i.is_displayed()]) >= 2)
            inputs = [i for i in self.driver.find_elements(By.TAG_NAME, "input") if i.is_displayed()]

            inputs[0].clear()
            inputs[0].send_keys(str(placa).strip().upper())
            inputs[1].clear()
            inputs[1].send_keys(str(renavam).strip())

            # 3. Clica em CONFIRMAR
            btn_confirmar = self.wait.until(
                EC.presence_of_element_located((
                    By.XPATH, "//button[contains(text(), 'CONFIRMAR') or @id='submit-veiculo']"
                ))
            )
            self._clicar_seguro(btn_confirmar)
            time.sleep(1.2)

            # Checa se houve alerta de veículo inexistente/inválido
            alertas = self.driver.find_elements(
                By.XPATH, "//*[contains(text(), 'não encontrado') or contains(text(), 'Inexistente') or contains(text(), 'Dados incorretos')]"
            )
            if alertas:
                return ("Não Localizado" if consultar_lic else "-"), ("Não Localizado" if consultar_multas else "-")

            # Helper interno para fechar o modal compartilhado do Detran-CE
            def fechar_modal_se_aberto():
                try:
                    fechar_btns = self.driver.find_elements(
                        By.XPATH, "//*[@id='outros-servicos-modal']//button[contains(text(), 'Fechar') or contains(@class, 'close') or @data-dismiss='modal']"
                    )
                    for b in fechar_btns:
                        if b.is_displayed():
                            self._clicar_seguro(b)
                            time.sleep(0.5)
                            break
                except Exception:
                    pass

            # 4. Extração de LICENCIAMENTO (Executado PRIMEIRO em modal limpo)
            if consultar_lic:
                try:
                    link_lic_els = self.driver.find_elements(
                        By.XPATH, "//*[contains(text(), 'Clique AQUI se deseja imprimir seu licenciamento.')]"
                    )
                    if not link_lic_els:
                        val_lic = "R$ 0,00"
                    else:
                        self._clicar_seguro(link_lic_els[0])
                        wait_modal = WebDriverWait(self.driver, 10)
                        
                        # Aguarda o modal carregar especificamente o conteúdo do Licenciamento
                        wait_modal.until(EC.visibility_of_element_located((
                            By.XPATH, "//*[@id='outros-servicos-modal']//*[contains(text(), 'licenciamento') or contains(text(), 'Licenciamento')]"
                        )))
                        time.sleep(0.8)

                        el_total_lic = wait_modal.until(EC.visibility_of_element_located((
                            By.XPATH, "//*[@id='outros-servicos-modal']//tr[contains(., 'TOTAL')]/td[last()]"
                        )))
                        texto_lic = el_total_lic.text.strip()
                        val_lic = texto_lic if texto_lic else "R$ 0,00"
                except TimeoutException:
                    val_lic = "R$ 0,00"
                except Exception:
                    val_lic = "Falha Licenciamento"
                finally:
                    fechar_modal_se_aberto()

            # 5. Extração de MULTAS (Executado DEPOIS com modal anterior fechado)
            if consultar_multas:
                try:
                    sem_multas = self.driver.find_elements(
                        By.XPATH, "//*[contains(text(), 'não possui multas') or contains(text(), 'Nenhuma multa') or contains(text(), 'Não constam multas')]"
                    )
                    if sem_multas:
                        val_multas = "R$ 0,00"
                    else:
                        links_multa = self.driver.find_elements(
                            By.XPATH, "//*[contains(text(), 'multas, CLIQUE AQUI') or contains(text(), 'multa, CLIQUE AQUI')]"
                        )
                        if links_multa:
                            self._clicar_seguro(links_multa[0])
                            wait_modal = WebDriverWait(self.driver, 10)
                            
                            wait_modal.until(EC.visibility_of_element_located((
                                By.XPATH, "//*[@id='outros-servicos-modal']//table//tr"
                            )))
                            time.sleep(0.8)

                            linhas_multa = self.driver.find_elements(By.XPATH, "//*[@id='outros-servicos-modal']//table//tr")
                            soma_multas = 0.0
                            encontrou_multas = False

                            for linha in linhas_multa:
                                tds = linha.find_elements(By.TAG_NAME, "td")
                                if len(tds) >= 6:
                                    txt_valor = tds[-1].text.strip()
                                    if "R$" in txt_valor:
                                        soma_multas += formatar_moeda_float(txt_valor)
                                        encontrou_multas = True

                            if encontrou_multas:
                                val_multas = formatar_float_moeda(soma_multas)
                            else:
                                el_total = self.driver.find_elements(
                                    By.XPATH, "//*[@id='outros-servicos-modal']//tr[contains(., 'TOTAL')]/td[last()]"
                                )
                                if el_total and "R$" in el_total[0].text:
                                    val_multas = el_total[0].text.strip()
                                else:
                                    val_multas = "R$ 0,00"
                        else:
                            val_multas = "R$ 0,00"
                except TimeoutException:
                    val_multas = "R$ 0,00"
                except Exception:
                    val_multas = "Falha Multas"
                finally:
                    fechar_modal_se_aberto()

            return val_lic, val_multas

        except (TimeoutException, WebDriverException):
            return ("Falha Detran" if consultar_lic else "-"), ("Falha Detran" if consultar_multas else "-")
        except Exception:
            return ("Falha Detran" if consultar_lic else "-"), ("Falha Detran" if consultar_multas else "-")

    def consultar_sefaz_ipva(self, placa: str, renavam: str) -> str:
        """
        Consulta o saldo de IPVA em aberto na Sefaz-CE com detecção precoce de veículos quitados.
        Diferencia veículos regulares (R$ 0,00) de falhas/erros de conexão.
        """
        url_sefaz = "https://ipva.sefaz.ce.gov.br/#/impostos/emitir-dae"
        try:
            self.driver.get(url_sefaz)

            # 1. Preenche Placa
            xpath_placa = "//input[contains(@ng-model, 'placa') or @name='placa' or @formcontrolname='placa' or @id='mat-input-1']"
            inp_placa = self.wait.until(EC.visibility_of_element_located((By.XPATH, xpath_placa)))
            inp_placa.clear()
            inp_placa.send_keys(str(placa).strip().upper())

            # 2. Preenche RENAVAM
            xpath_renavam = "//input[contains(@ng-model, 'renavam') or @name='renavam' or @formcontrolname='renavam' or @id='mat-input-2']"
            inp_renavam = self.wait.until(EC.visibility_of_element_located((By.XPATH, xpath_renavam)))
            inp_renavam.clear()
            inp_renavam.send_keys(str(renavam).strip())

            # 3. Clica em Pesquisar
            xpath_btn = "//button[contains(translate(., 'pesquisar', 'PESQUISAR'), 'PESQUISAR') or contains(@class, 'search-button')]"
            btn_pesquisar = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn)))
            self._clicar_seguro(btn_pesquisar)

            # 4. Aguarda carregar resposta (débitos em aberto OU débitos quitados OU aviso de erro)
            def resposta_pronta(d):
                tem_aberto = d.find_elements(By.XPATH, "//*[contains(text(), 'Débitos de IPVA não quitados')]")
                quitado = d.find_elements(
                    By.XPATH, "//*[contains(text(), 'Débitos de IPVA Regularizados/Quitados') or contains(text(), 'Nenhum débito')]"
                )
                erro = d.find_elements(
                    By.XPATH, "//*[contains(text(), 'não encontrado') or contains(text(), 'inválido') or contains(text(), 'Não há registro')]"
                )
                return tem_aberto or quitado or erro

            self.wait.until(resposta_pronta)

            # Se acusar não encontrado
            if self.driver.find_elements(By.XPATH, "//*[contains(text(), 'não encontrado') or contains(text(), 'Não há registro')]"):
                return "Não Localizado"

            # Se não há tabela de não quitados, encerra como R$ 0,00 imediatamente
            if not self.driver.find_elements(By.XPATH, "//*[contains(text(), 'Débitos de IPVA não quitados')]"):
                return "R$ 0,00"

            # 5. Extrai saldo em aberto
            xpath_valor = "//td[contains(., 'Em aberto')]/preceding-sibling::td[contains(., 'R$')][1]"
            el_valor = self.wait.until(EC.visibility_of_element_located((By.XPATH, xpath_valor)))
            txt_ipva = el_valor.text.strip()
            return txt_ipva if txt_ipva else "R$ 0,00"

        except (TimeoutException, WebDriverException):
            return "Falha Sefaz"
        except Exception:
            return "Falha Sefaz"


# ======================================================================================
# 3. EXPORTAÇÃO EXECUTIVA (EXCEL E PDF PAISAGEM COM REPORTLAB)
# ======================================================================================

if REPORTLAB_DISPONIVEL:
    class NumberedCanvas(canvas.Canvas):
        """Canvas customizado para renderizar número total de páginas dinamicamente no rodapé."""
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._saved_page_states = []

        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            num_pages = len(self._saved_page_states)
            for state in self._saved_page_states:
                self.__dict__.update(state)
                self.draw_footer(num_pages)
                canvas.Canvas.showPage(self)
            canvas.Canvas.save(self)

        def draw_footer(self, page_count):
            self.saveState()
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.HexColor("#64748B"))
            largura, altura = landscape(A4)
            texto_esq = "Automação de Frota CE v2.1 | Relatório Oficial Consolidado (Detran-CE & Sefaz-CE)"
            texto_dir = f"Página {self._pageNumber} de {page_count}"
            self.drawString(25, 18, texto_esq)
            self.drawRightString(largura - 25, 18, texto_dir)
            self.setStrokeColor(colors.HexColor("#E2E8F0"))
            self.setLineWidth(0.5)
            self.line(25, 28, largura - 25, 28)
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

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                col_header = str(ws.cell(1, cell.column).value or "")
                val_str = str(cell.value or "")
                estilo = obter_estilo_financeiro(col_header, val_str)
                cell.border = borda_fina

                if estilo["is_fin"] and (estilo["val_num"] > 0 or val_str.startswith("R$")):
                    cell.alignment = alinhamento_direita
                    cor_hex = "DC2626" if estilo["is_red"] else "000000"
                    under = "single" if estilo["is_underline"] else None
                    cell.font = Font(
                        name="Calibri",
                        size=10,
                        bold=estilo["is_bold"],
                        underline=under,
                        color=cor_hex
                    )
                elif any(k in col_header.lower() for k in ["placa", "renavam", "status", "data"]):
                    cell.alignment = alinhamento_centro
                    cell.font = fonte_linhas
                else:
                    cell.alignment = alinhamento_esquerda
                    cell.font = fonte_linhas

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

        try:
            wb.save(caminho_efetivo)
        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pasta, nome = os.path.split(caminho_arquivo)
            raiz, ext = os.path.splitext(nome)
            caminho_efetivo = os.path.join(pasta, f"{raiz}_{timestamp}{ext}")
            wb.save(caminho_efetivo)

        return caminho_efetivo

    @staticmethod
    def exportar_pdf(df: pd.DataFrame, caminho_arquivo: str, totais: dict, filtros_ativos: list[str]) -> str:
        """
        Gera relatório PDF profissional em formato Paisagem (Landscape A4).
        Trata PermissionError salvando com timestamp para evitar perda de dados.
        Retorna o caminho efetivo do arquivo gravado.
        """
        if not REPORTLAB_DISPONIVEL:
            raise RuntimeError(
                "A biblioteca 'reportlab' não está instalada no ambiente Python.\n"
                "Para gerar relatórios em PDF, execute no terminal: pip install reportlab"
            )

        caminho_efetivo = caminho_arquivo
        pasta_dir = os.path.dirname(os.path.abspath(caminho_efetivo))
        if pasta_dir:
            os.makedirs(pasta_dir, exist_ok=True)

        def construir_story():
            styles = getSampleStyleSheet()
            style_titulo = ParagraphStyle(
                'Titulo',
                parent=styles['Heading1'],
                fontName='Helvetica-Bold',
                fontSize=15,
                leading=18,
                textColor=colors.HexColor('#1E3A8A'),
                spaceAfter=2
            )
            style_sub = ParagraphStyle(
                'SubTitulo',
                parent=styles['Normal'],
                fontName='Helvetica',
                fontSize=8.5,
                leading=11,
                textColor=colors.HexColor('#475569'),
                spaceAfter=10
            )
            style_th = ParagraphStyle(
                'Th',
                fontName='Helvetica-Bold',
                fontSize=8,
                leading=10,
                textColor=colors.white,
                alignment=1
            )
            style_td = ParagraphStyle(
                'Td',
                fontName='Helvetica',
                fontSize=7.5,
                leading=9.5,
                textColor=colors.HexColor('#1E293B'),
                alignment=0
            )
            style_td_center = ParagraphStyle('TdCenter', parent=style_td, alignment=1)
            style_td_right = ParagraphStyle('TdRight', parent=style_td, alignment=2)

            story = []

            # 1. Cabeçalho
            story.append(Paragraph("🚗 RELATÓRIO CONSOLIDADO DE DÉBITOS VEICULARES - FROTA CE", style_titulo))
            filtros_str = ", ".join(filtros_ativos) if filtros_ativos else "Todos"
            data_extracao = datetime.now().strftime("%d/%m/%Y às %H:%M:%S")
            sub_info = f"Emissão: {data_extracao}  |  Débitos Consultados: <b>{filtros_str}</b>  |  Total de Veículos: <b>{len(df)}</b>"
            story.append(Paragraph(sub_info, style_sub))

            # 2. Cards de Resumo Financeiro
            largura_util = landscape(A4)[0] - 50  # ~791.89 pt
            larg_card = largura_util / 5.0
            cards_data = [
                [
                    Paragraph("<b>VEÍCULOS</b><br/><font size=10 color='#1E3A8A'><b>" + str(len(df)) + "</b></font>", style_td_center),
                    Paragraph("<b>TOTAL LICENCIAMENTO</b><br/><font size=10 color='#B45309'><b>" + totais.get('lic', 'R$ 0,00') + "</b></font>", style_td_center),
                    Paragraph("<b>TOTAL IPVA</b><br/><font size=10 color='#B91C1C'><b>" + totais.get('ipva', 'R$ 0,00') + "</b></font>", style_td_center),
                    Paragraph("<b>TOTAL MULTAS</b><br/><font size=10 color='#7C2D12'><b>" + totais.get('multas', 'R$ 0,00') + "</b></font>", style_td_center),
                    Paragraph("<b>TOTAL GERAL</b><br/><font size=10 color='#15803D'><b>" + totais.get('total', 'R$ 0,00') + "</b></font>", style_td_center),
                ]
            ]
            t_cards = Table(cards_data, colWidths=[larg_card] * 5)
            t_cards.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
                ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#CBD5E1')),
                ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ]))
            story.append(t_cards)
            story.append(Spacer(1, 10))

            # 3. Tabela Detalhada com Colunas
            colunas = list(df.columns)
            num_cols = len(colunas)
            larg_col = largura_util / num_cols

            table_data = [[Paragraph(f"<b>{col}</b>", style_th) for col in colunas]]

            for _, row in df.iterrows():
                linha_celulas = []
                for col in colunas:
                    val = str(row[col] if pd.notna(row[col]) else "")
                    estilo = obter_estilo_financeiro(col, val)

                    if estilo["is_fin"] and (estilo["val_num"] > 0 or val.startswith("R$")):
                        est = style_td_right
                        val_html = val
                        if estilo["is_underline"]:
                            val_html = f"<u>{val_html}</u>"
                        if estilo["is_bold"]:
                            val_html = f"<b>{val_html}</b>"
                        if estilo["is_red"]:
                            val_html = f'<font color="#DC2626">{val_html}</font>'
                        linha_celulas.append(Paragraph(val_html, est))
                    elif any(k in col.lower() for k in ["placa", "renavam", "status", "data"]):
                        est = style_td_center
                        linha_celulas.append(Paragraph(val, est))
                    else:
                        est = style_td
                        linha_celulas.append(Paragraph(val, est))
                table_data.append(linha_celulas)

            t_frota = Table(table_data, colWidths=[larg_col] * num_cols, repeatRows=1)
            t_style = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]

            for i in range(1, len(table_data)):
                bg = colors.HexColor('#F8FAFC') if i % 2 == 1 else colors.white
                t_style.append(('BACKGROUND', (0, i), (-1, i), bg))

            t_frota.setStyle(TableStyle(t_style))
            story.append(t_frota)
            return story

        try:
            doc = SimpleDocTemplate(
                caminho_efetivo,
                pagesize=landscape(A4),
                leftMargin=25,
                rightMargin=25,
                topMargin=25,
                bottomMargin=35
            )
            doc.build(construir_story(), canvasmaker=NumberedCanvas)
        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pasta, nome = os.path.split(caminho_arquivo)
            raiz, ext = os.path.splitext(nome)
            caminho_efetivo = os.path.join(pasta, f"{raiz}_{timestamp}{ext}")
            doc = SimpleDocTemplate(
                caminho_efetivo,
                pagesize=landscape(A4),
                leftMargin=25,
                rightMargin=25,
                topMargin=25,
                bottomMargin=35
            )
            doc.build(construir_story(), canvasmaker=NumberedCanvas)

        return caminho_efetivo


# ======================================================================================
# 4. INTERFACE GRÁFICA PROFISSIONAL (TKINTER + TTK)
# ======================================================================================

class AppConsultaFrotaGUI:
    """Interface Gráfica com filtros dinâmicos, acompanhamento em tempo real e exportações."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Consulta de Frota - Detran/Sefaz CE (v2.1)")
        self.root.geometry("960x780")
        self.root.minsize(880, 700)

        # Variáveis de Configuração e Estado
        self.caminho_arquivo = tk.StringVar(value="")
        self.pasta_destino_var = tk.StringVar(value="")
        self.arquivo_excel_gerado = ""
        self.arquivo_pdf_gerado = ""

        # Filtros de Débito
        self.var_ipva = tk.BooleanVar(value=True)
        self.var_lic = tk.BooleanVar(value=True)
        self.var_multas = tk.BooleanVar(value=True)

        # Formato de Exportação ("excel", "pdf", "ambos")
        padrao_export = "ambos" if REPORTLAB_DISPONIVEL else "excel"
        self.var_formato_exportacao = tk.StringVar(value=padrao_export)

        # Opções Gerais
        self.var_headless = tk.BooleanVar(value=True)
        self.var_status = tk.StringVar(value="Pronto para iniciar. Selecione a planilha de frota.")

        # Métricas do Dashboard
        self.var_metricas_veiculos = tk.StringVar(value="0 / 0")
        self.var_metricas_lic = tk.StringVar(value="R$ 0,00")
        self.var_metricas_ipva = tk.StringVar(value="R$ 0,00")
        self.var_metricas_multas = tk.StringVar(value="R$ 0,00")
        self.var_metricas_total = tk.StringVar(value="R$ 0,00")

        # Controle de Execução e Memória de Resultados
        self.is_executando = False
        self.cancelar_solicitado = False
        self.scraper_engine: FrotaScraperEngine | None = None
        self.df_carregado: pd.DataFrame | None = None
        self.col_placa = "Placa"
        self.col_renavam = "RENAVAM"
        self.resultados_memoria: dict[int, dict] = {}
        self.indices_em_falha: list[int] = []

        # Modo Noturno / Dark Mode
        self.modo_noturno = False
        self.cards_metricas = []

        # Configuração Visual
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self._criar_layout()
        self._aplicar_tema()

        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar_janela)

    def _aplicar_tema(self):
        """Aplica tokens de cores e estilos suaves e elegantes para Modo Claro e Modo Noturno."""
        if self.modo_noturno:
            bg_root = "#191B22"       # Fundo grafite escuro suave
            bg_card = "#21242D"       # Superfície de cartões fosca e calma
            border_card = "#2F333F"   # Borda sutil de baixo contraste
            fg_title = "#E2E8F0"      # Título suave e nítido
            fg_sub = "#8E99A8"        # Subtítulo com opacidade agradável
            fg_text = "#CBD5E1"       # Texto claro confortável para leitura
            bg_tree = "#21242D"       # Fundo da tabela integrado ao cartão
            fg_tree = "#CBD5E1"
            bg_tree_hd = "#2A2E3B"    # Cabeçalho da tabela discreto
            fg_tree_hd = "#E2E8F0"
            bg_sel = "#353D52"        # Destaque de seleção suave (azul ardósia)
            fg_sel = "#F8FAFC"
            bg_tab = "#1B1D24"
            fg_tab = "#8E99A8"
            bg_tab_sel = "#21242D"
            fg_tab_sel = "#9BB3D6"    # Aba selecionada em tom ardósia pastel
            bg_log = "#181A20"        # Terminal escuro com contraste repousante
            fg_log = "#A5B4C7"        # Fonte mono cinza-azulada suave
            status_fg = "#9BB3D6"     # Status informativo sutil
            bg_btn_sec = "#2D323E"
            fg_btn_sec = "#CBD5E1"
            bg_btn_sec_act = "#393F4E"
            bg_btn_muted = "#2A2E3A"
            fg_btn_muted = "#CBD5E1"
            bg_btn_theme = "#2A2E3A"
            fg_btn_theme = "#CBD5E1"
            # Botões de ação em tons sóbrios e não ofuscantes
            bg_btn_pri = "#354A6B"
            bg_btn_pri_act = "#2C3E5A"
            bg_btn_act = "#2D5A43"
            bg_btn_act_act = "#244936"
            bg_btn_dng = "#6B3838"
            bg_btn_dng_act = "#582E2E"
        else:
            bg_root = "#F0F4F8"
            bg_card = "#FFFFFF"
            border_card = "#CBD5E1"
            fg_title = "#1A365D"
            fg_sub = "#4A5568"
            fg_text = "#2D3748"
            bg_tree = "#FFFFFF"
            fg_tree = "#2D3748"
            bg_tree_hd = "#E2E8F0"
            fg_tree_hd = "#1A365D"
            bg_sel = "#BEE3F8"
            fg_sel = "#1A365D"
            bg_tab = "#E2E8F0"
            fg_tab = "#4A5568"
            bg_tab_sel = "#FFFFFF"
            fg_tab_sel = "#1A365D"
            bg_log = "#1A202C"
            fg_log = "#E2E8F0"
            status_fg = "#1A365D"
            bg_btn_sec = "#4A5568"
            fg_btn_sec = "#FFFFFF"
            bg_btn_sec_act = "#2D3748"
            bg_btn_muted = "#E2E8F0"
            fg_btn_muted = "#2D3748"
            bg_btn_theme = "#E2E8F0"
            fg_btn_theme = "#1A365D"
            bg_btn_pri = "#2563EB"
            bg_btn_pri_act = "#1D4ED8"
            bg_btn_act = "#16A34A"
            bg_btn_act_act = "#15803D"
            bg_btn_dng = "#DC2626"
            bg_btn_dng_act = "#B91C1C"

        self.root.configure(bg=bg_root)

        # Configurações ttk gerais
        self.style.configure(".", background=bg_root, foreground=fg_text, font=("Segoe UI", 9))
        self.style.configure("TFrame", background=bg_root)
        self.style.configure("Card.TFrame", background=bg_card)
        self.style.configure("TLabel", background=bg_root, foreground=fg_text)
        self.style.configure("Card.TLabel", background=bg_card, foreground=fg_text)
        self.style.configure("Header.TLabel", background=bg_root, font=("Segoe UI", 14, "bold"), foreground=fg_title)
        self.style.configure("SubHeader.TLabel", background=bg_root, font=("Segoe UI", 9), foreground=fg_sub)

        # Cards / LabelFrames
        self.style.configure("Card.TLabelframe", background=bg_card, relief=tk.SOLID, borderwidth=1)
        self.style.configure("Card.TLabelframe.Label", font=("Segoe UI", 9, "bold"), foreground=fg_title, background=bg_root)

        # Checkbuttons e Radiobuttons
        self.style.configure("Card.TCheckbutton", background=bg_card, foreground=fg_text)
        self.style.map("Card.TCheckbutton", background=[("active", bg_card)], foreground=[("active", fg_title)])
        self.style.configure("Card.TRadiobutton", background=bg_card, foreground=fg_text)
        self.style.map("Card.TRadiobutton", background=[("active", bg_card)], foreground=[("active", fg_title)])

        # Botão Primário (Selecionar)
        self.style.configure(
            "Primary.TButton",
            font=("Segoe UI", 9, "bold"),
            background=bg_btn_pri,
            foreground="#FFFFFF",
            padding=(10, 6)
        )
        self.style.map("Primary.TButton", background=[("active", bg_btn_pri_act), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        # Botão de Ação (Iniciar)
        self.style.configure(
            "Action.TButton",
            font=("Segoe UI", 10, "bold"),
            background=bg_btn_act,
            foreground="#FFFFFF",
            padding=(14, 7)
        )
        self.style.map("Action.TButton", background=[("active", bg_btn_act_act), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        # Botão de Cancelar
        self.style.configure(
            "Danger.TButton",
            font=("Segoe UI", 9, "bold"),
            background=bg_btn_dng,
            foreground="#FFFFFF",
            padding=(10, 7)
        )
        self.style.map("Danger.TButton", background=[("active", bg_btn_dng_act), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        # Botões Secundários
        self.style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9, "bold"),
            background=bg_btn_sec,
            foreground=fg_btn_sec,
            padding=(9, 5)
        )
        self.style.map("Secondary.TButton", background=[("active", bg_btn_sec_act), ("disabled", "#4B5563" if self.modo_noturno else "#64748B")])

        # Botão Muted (Alternar Filtros)
        self.style.configure(
            "Muted.TButton",
            font=("Segoe UI", 8, "bold"),
            background=bg_btn_muted,
            foreground=fg_btn_muted,
            padding=(6, 3)
        )
        self.style.map("Muted.TButton", background=[("active", border_card)])

        # Botão Alternador de Tema (Dark / Light)
        self.style.configure(
            "Theme.TButton",
            font=("Segoe UI", 9, "bold"),
            background=bg_btn_theme,
            foreground=fg_btn_theme,
            padding=(10, 5)
        )
        self.style.map("Theme.TButton", background=[("active", border_card)])

        # Notebook / Abas
        self.style.configure("TNotebook", background=bg_root)
        self.style.configure("TNotebook.Tab", background=bg_tab, foreground=fg_tab, padding=(10, 4), font=("Segoe UI", 9, "bold"))
        self.style.map("TNotebook.Tab", background=[("selected", bg_tab_sel)], foreground=[("selected", fg_tab_sel)])

        # Tabela Treeview
        self.style.configure(
            "Treeview",
            background=bg_tree,
            fieldbackground=bg_tree,
            foreground=fg_tree,
            rowheight=25,
            font=("Segoe UI", 9)
        )
        self.style.configure(
            "Treeview.Heading",
            background=bg_tree_hd,
            foreground=fg_tree_hd,
            font=("Segoe UI", 9, "bold")
        )
        self.style.map("Treeview", background=[("selected", bg_sel)], foreground=[("selected", fg_sel)])

        # Tags da Treeview com cores suaves e descansadas
        if hasattr(self, "tree"):
            if self.modo_noturno:
                self.tree.tag_configure("pendente", foreground="#7D8799")
                self.tree.tag_configure("processando", background="#232B3B", foreground="#9BB3D6", font=("Segoe UI", 9, "bold"))
                self.tree.tag_configure("com_debito", background="#332225", foreground="#DDA0A8", font=("Segoe UI", 9, "bold"))
                self.tree.tag_configure("sem_debito", background="#1E2B24", foreground="#90BE9F")
                self.tree.tag_configure("falha", background="#332B1E", foreground="#D8B983", font=("Segoe UI", 9, "bold"))
            else:
                self.tree.tag_configure("pendente", foreground="#718096")
                self.tree.tag_configure("processando", background="#EBF8FF", foreground="#2B6CB0", font=("Segoe UI", 9, "bold"))
                self.tree.tag_configure("com_debito", background="#FFF5F5", foreground="#C53030", font=("Segoe UI", 9, "bold"))
                self.tree.tag_configure("sem_debito", background="#F0FFF4", foreground="#22543D")
                self.tree.tag_configure("falha", background="#FEF3C7", foreground="#B45309", font=("Segoe UI", 9, "bold"))

        # Log Textual
        if hasattr(self, "txt_log"):
            self.txt_log.configure(bg=bg_log, fg=fg_log, insertbackground="#FFFFFF")

        # Rótulo de Status
        if hasattr(self, "lbl_status"):
            self.lbl_status.configure(foreground=status_fg)

        # Rótulos de Caminho
        if hasattr(self, "lbl_caminho"):
            if self.df_carregado is not None:
                self.lbl_caminho.configure(foreground="#90BE9F" if self.modo_noturno else "#15803D")
            else:
                self.lbl_caminho.configure(foreground=fg_sub)

        # Atualização dos Mini-Cards de Métricas
        for card_frame, lbl_t, lbl_v, cor_l, cor_d in self.cards_metricas:
            card_frame.configure(bg=bg_card, highlightbackground=border_card)
            lbl_t.configure(bg=bg_card, fg=fg_sub)
            lbl_v.configure(bg=bg_card, fg=cor_d if self.modo_noturno else cor_l)

    def _alternar_modo_noturno(self):
        """Alterna entre Modo Claro e Modo Noturno em tempo real."""
        self.modo_noturno = not self.modo_noturno
        self.btn_tema.config(text="☀️ Modo Claro" if self.modo_noturno else "🌙 Modo Noturno")
        self._aplicar_tema()
        self._adicionar_log("Modo Noturno ativado." if self.modo_noturno else "Modo Claro ativado.")

    def _criar_layout(self):
        """Estrutura os painéis, filtros e controles na janela."""
        container = ttk.Frame(self.root, padding="12 12 12 12")
        container.pack(fill=tk.BOTH, expand=True)

        # 1. CABEÇALHO (Título + Alternador de Tema)
        header = ttk.Frame(container)
        header.pack(fill=tk.X, pady=(0, 8))

        lbl_tit = ttk.Label(header, text="🚗 Consulta de Frota - Detran/Sefaz CE (v2.1)", style="Header.TLabel")
        lbl_tit.pack(side=tk.LEFT)

        lbl_sub = ttk.Label(
            header,
            text="Automação Resiliente: IPVA, Licenciamento & Multas com Exportação PDF/Excel",
            style="SubHeader.TLabel"
        )
        lbl_sub.pack(side=tk.LEFT, padx=(12, 0), pady=(3, 0))

        # Botão Alternador de Modo Noturno
        self.btn_tema = ttk.Button(
            header,
            text="🌙 Modo Noturno",
            style="Theme.TButton",
            command=self._alternar_modo_noturno
        )
        self.btn_tema.pack(side=tk.RIGHT)

        # 2. CARD: SELEÇÃO, FILTROS E EXPORTAÇÃO
        frame_configs = ttk.LabelFrame(container, text=" Configurações e Filtros de Consulta ", style="Card.TLabelframe", padding="10 8 10 8")
        frame_configs.pack(fill=tk.X, pady=(0, 8))

        # Linha 1: Seleção de Arquivo
        linha_arq = ttk.Frame(frame_configs, style="Card.TFrame")
        linha_arq.pack(fill=tk.X, pady=(0, 5))

        self.btn_selecionar = ttk.Button(
            linha_arq,
            text="📁 Selecionar Planilha (.xlsx)",
            style="Primary.TButton",
            command=self._selecionar_arquivo
        )
        self.btn_selecionar.pack(side=tk.LEFT, padx=(0, 10))

        self.lbl_caminho = ttk.Label(
            linha_arq,
            text="Nenhuma planilha selecionada",
            font=("Segoe UI", 9, "italic"),
            foreground="#718096",
            style="Card.TLabel"
        )
        self.lbl_caminho.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Linha 1.5: Pasta de Destino Customizada (Melhoria 7)
        linha_dest = ttk.Frame(frame_configs, style="Card.TFrame")
        linha_dest.pack(fill=tk.X, pady=(0, 6))

        self.btn_pasta_destino = ttk.Button(
            linha_dest,
            text="📂 Pasta de Destino...",
            style="Secondary.TButton",
            command=self._selecionar_pasta_destino
        )
        self.btn_pasta_destino.pack(side=tk.LEFT, padx=(0, 10))

        self.lbl_pasta_destino = ttk.Label(
            linha_dest,
            text="Destino dos relatórios: Automático (mesma pasta da planilha)",
            font=("Segoe UI", 8),
            foreground="#64748B",
            style="Card.TLabel"
        )
        self.lbl_pasta_destino.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Linha 2: Filtros de Débito (Checkboxes)
        linha_filtros = ttk.Frame(frame_configs, style="Card.TFrame")
        linha_filtros.pack(fill=tk.X, pady=(0, 6))

        lbl_filtro = ttk.Label(linha_filtros, text="Consultar Débitos:", font=("Segoe UI", 9, "bold"), style="Card.TLabel")
        lbl_filtro.pack(side=tk.LEFT, padx=(0, 8))

        self.chk_ipva = ttk.Checkbutton(linha_filtros, text="IPVA (Sefaz-CE)", variable=self.var_ipva, style="Card.TCheckbutton")
        self.chk_ipva.pack(side=tk.LEFT, padx=(0, 10))

        self.chk_lic = ttk.Checkbutton(linha_filtros, text="Licenciamento (Detran-CE)", variable=self.var_lic, style="Card.TCheckbutton")
        self.chk_lic.pack(side=tk.LEFT, padx=(0, 10))

        self.chk_multas = ttk.Checkbutton(linha_filtros, text="Multas (Detran-CE)", variable=self.var_multas, style="Card.TCheckbutton")
        self.chk_multas.pack(side=tk.LEFT, padx=(0, 12))

        self.btn_toggle_filtros = ttk.Button(
            linha_filtros,
            text="☑️ Marcar/Desmarcar Todos",
            style="Muted.TButton",
            command=self._alternar_todos_filtros
        )
        self.btn_toggle_filtros.pack(side=tk.LEFT)

        # Linha 3: Formato de Exportação & Headless
        linha_opcoes = ttk.Frame(frame_configs, style="Card.TFrame")
        linha_opcoes.pack(fill=tk.X, pady=(0, 4))

        lbl_exp = ttk.Label(linha_opcoes, text="Exportação:", font=("Segoe UI", 9, "bold"), style="Card.TLabel")
        lbl_exp.pack(side=tk.LEFT, padx=(0, 8))

        rb_excel = ttk.Radiobutton(linha_opcoes, text="Excel (.xlsx)", value="excel", variable=self.var_formato_exportacao, style="Card.TRadiobutton")
        rb_excel.pack(side=tk.LEFT, padx=(0, 8))

        rb_pdf = ttk.Radiobutton(
            linha_opcoes,
            text="PDF Paisagem (.pdf)",
            value="pdf",
            variable=self.var_formato_exportacao,
            style="Card.TRadiobutton",
            state="normal" if REPORTLAB_DISPONIVEL else "disabled"
        )
        rb_pdf.pack(side=tk.LEFT, padx=(0, 8))

        rb_ambos = ttk.Radiobutton(
            linha_opcoes,
            text="Ambos (Excel + PDF)",
            value="ambos",
            variable=self.var_formato_exportacao,
            style="Card.TRadiobutton",
            state="normal" if REPORTLAB_DISPONIVEL else "disabled"
        )
        rb_ambos.pack(side=tk.LEFT, padx=(0, 15))

        self.chk_headless = ttk.Checkbutton(
            linha_opcoes,
            text="Navegador invisível em segundo plano",
            variable=self.var_headless,
            style="Card.TCheckbutton"
        )
        self.chk_headless.pack(side=tk.LEFT)

        # Botões de Ação (Iniciar / Cancelar) posicionados à direita
        self.btn_cancelar = ttk.Button(
            linha_opcoes,
            text="⏹️ Cancelar",
            style="Danger.TButton",
            state="disabled",
            command=self._cancelar_processamento
        )
        self.btn_cancelar.pack(side=tk.RIGHT, padx=(6, 0))

        self.btn_iniciar = ttk.Button(
            linha_opcoes,
            text="▶️ Iniciar Consulta",
            style="Action.TButton",
            state="disabled",
            command=self._iniciar_processamento
        )
        self.btn_iniciar.pack(side=tk.RIGHT)

        # 3. CARD: TABELA INTERATIVA EM TEMPO REAL E LOGS
        frame_centro = ttk.LabelFrame(container, text=" Acompanhamento em Tempo Real ", style="Card.TLabelframe", padding="8 8 8 8")
        frame_centro.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.notebook = ttk.Notebook(frame_centro)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Aba 1: Tabela Visual
        aba_tabela = ttk.Frame(self.notebook, padding="4 4 4 4")
        self.notebook.add(aba_tabela, text=" 📊 Tabela de Veículos ")

        colunas = ("num", "placa", "renavam", "licenciamento", "ipva", "multas", "situacao")
        self.tree = ttk.Treeview(aba_tabela, columns=colunas, show="headings", selectmode="extended")

        self.tree.heading("num", text="Item")
        self.tree.heading("placa", text="Placa")
        self.tree.heading("renavam", text="RENAVAM")
        self.tree.heading("licenciamento", text="Licenciamento (Detran)")
        self.tree.heading("ipva", text="IPVA (Sefaz)")
        self.tree.heading("multas", text="Multas (Detran)")
        self.tree.heading("situacao", text="Situação")

        self.tree.column("num", width=40, anchor="center")
        self.tree.column("placa", width=85, anchor="center")
        self.tree.column("renavam", width=115, anchor="center")
        self.tree.column("licenciamento", width=130, anchor="e")
        self.tree.column("ipva", width=130, anchor="e")
        self.tree.column("multas", width=130, anchor="e")
        self.tree.column("situacao", width=145, anchor="center")

        scroll_y = ttk.Scrollbar(aba_tabela, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # Menu de contexto no botão direito da tabela (Melhoria 8)
        self.menu_contexto = tk.Menu(self.tree, tearoff=0)
        self.menu_contexto.add_command(label="🔄 Reconsultar Veículo(s) Selecionado(s)", command=self._reconsultar_selecionado)
        self.menu_contexto.add_command(label="⚠️ Reconsultar Apenas Falhas", command=self._reconsultar_apenas_falhas)
        self.tree.bind("<Button-3>", self._exibir_menu_contexto)

        # Aba 2: Log Textual
        aba_log = ttk.Frame(self.notebook, padding="4 4 4 4")
        self.notebook.add(aba_log, text=" 📝 Log de Execução ")

        self.txt_log = tk.Text(
            aba_log,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#1A202C",
            fg="#E2E8F0",
            insertbackground="#FFFFFF",
            relief=tk.FLAT
        )
        scroll_log = ttk.Scrollbar(aba_log, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scroll_log.set)
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_log.pack(side=tk.RIGHT, fill=tk.Y)

        self._adicionar_log("Sistema iniciado (v2.1). Selecione uma planilha de frota para começar.")

        # 4. DASHBOARD DE MÉTRICAS (5 MINI-CARDS)
        frame_metricas = ttk.Frame(container)
        frame_metricas.pack(fill=tk.X, pady=(0, 8))

        def criar_card(pai, titulo, var, cor_l, cor_d):
            card = tk.Frame(pai, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1, padx=10, pady=5)
            card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=3)
            lbl_t = tk.Label(card, text=titulo, font=("Segoe UI", 8, "bold"), fg="#64748B", bg="#FFFFFF")
            lbl_t.pack(anchor="w")
            lbl_v = tk.Label(card, textvariable=var, font=("Segoe UI", 11, "bold"), fg=cor_l, bg="#FFFFFF")
            lbl_v.pack(anchor="w")
            self.cards_metricas.append((card, lbl_t, lbl_v, cor_l, cor_d))
            return card

        criar_card(frame_metricas, "VEÍCULOS", self.var_metricas_veiculos, "#1E3A8A", "#9BB3D6")
        criar_card(frame_metricas, "LICENCIAMENTO", self.var_metricas_lic, "#B45309", "#D8B983")
        criar_card(frame_metricas, "IPVA EM ABERTO", self.var_metricas_ipva, "#B91C1C", "#DDA0A8")
        criar_card(frame_metricas, "MULTAS DETRAN", self.var_metricas_multas, "#7C2D12", "#DBA392")
        criar_card(frame_metricas, "TOTAL A PAGAR", self.var_metricas_total, "#15803D", "#90BE9F")

        # 5. BARRA DE PROGRESSO E AÇÕES PÓS-PROCESSAMENTO
        frame_rodape = ttk.Frame(container)
        frame_rodape.pack(fill=tk.X)

        self.lbl_status = ttk.Label(
            frame_rodape,
            textvariable=self.var_status,
            font=("Segoe UI", 9, "bold"),
            foreground="#1A365D"
        )
        self.lbl_status.pack(anchor="w", pady=(0, 3))

        self.barra_progresso = ttk.Progressbar(frame_rodape, mode="determinate")
        self.barra_progresso.pack(fill=tk.X, pady=(0, 6))

        frame_botoes_finais = ttk.Frame(frame_rodape)
        frame_botoes_finais.pack(fill=tk.X)

        self.btn_abrir_excel = ttk.Button(
            frame_botoes_finais,
            text="📊 Abrir Planilha (Excel)",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_excel_gerado
        )
        self.btn_abrir_excel.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_abrir_pdf = ttk.Button(
            frame_botoes_finais,
            text="📄 Abrir Relatório (PDF)",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_pdf_gerado
        )
        self.btn_abrir_pdf.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_abrir_pasta = ttk.Button(
            frame_botoes_finais,
            text="📂 Abrir Pasta de Destino",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_pasta_destino
        )
        self.btn_abrir_pasta.pack(side=tk.LEFT, padx=(0, 6))

        # Botão de Reconsulta de Falhas (Melhoria 8)
        self.btn_reconsultar_falhas = ttk.Button(
            frame_botoes_finais,
            text="⚠️ Reconsultar Falhas",
            style="Secondary.TButton",
            state="disabled",
            command=self._reconsultar_apenas_falhas
        )
        self.btn_reconsultar_falhas.pack(side=tk.LEFT)

    # ==================================================================================
    # INTERAÇÕES E LÓGICA DE CONTROLE
    # ==================================================================================

    def _adicionar_log(self, msg: str):
        """Insere linha formatada no log visual."""
        hora = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert(tk.END, f"[{hora}] {msg}\n")
        self.txt_log.see(tk.END)

    def _exibir_menu_contexto(self, event):
        """Abre menu de contexto com clique com botão direito na tabela."""
        item = self.tree.identify_row(event.y)
        if item:
            if item not in self.tree.selection():
                self.tree.selection_set(item)
            self.menu_contexto.post(event.x_root, event.y_root)

    def _alternar_todos_filtros(self):
        """Marca ou desmarca todos os filtros de débito simultaneamente."""
        estado_atual = self.var_ipva.get() and self.var_lic.get() and self.var_multas.get()
        novo_estado = not estado_atual
        self.var_ipva.set(novo_estado)
        self.var_lic.set(novo_estado)
        self.var_multas.set(novo_estado)

    def _selecionar_pasta_destino(self):
        """Permite ao usuário definir diretório de destino customizado (Melhoria 7)."""
        pasta = filedialog.askdirectory(title="Selecione a Pasta de Destino dos Relatórios")
        if pasta:
            self.pasta_destino_var.set(pasta)
            self.lbl_pasta_destino.config(
                text=f"Destino: {pasta}",
                foreground="#9BB3D6" if self.modo_noturno else "#1D4ED8"
            )
            self._adicionar_log(f"Pasta de destino personalizada definida: {pasta}")

    def _selecionar_arquivo(self):
        """Lê o arquivo Excel com resiliência, validação e popula a grade inicial."""
        caminho = filedialog.askopenfilename(
            title="Selecione a Planilha de Veículos",
            filetypes=[("Planilhas Excel (*.xlsx)", "*.xlsx")]
        )
        if not caminho:
            return

        try:
            df, col_p, col_r = FrotaDataProcessor.carregar_planilha(caminho)
            self.df_carregado = df
            self.col_placa = col_p
            self.col_renavam = col_r
            self.caminho_arquivo.set(caminho)

            # Se não houver pasta customizada, usa a pasta da planilha como padrão
            if not self.pasta_destino_var.get():
                pasta_padrao = os.path.dirname(os.path.abspath(caminho))
                self.pasta_destino_var.set(pasta_padrao)
                self.lbl_pasta_destino.config(
                    text=f"Destino: {pasta_padrao}",
                    foreground="#90BE9F" if self.modo_noturno else "#15803D"
                )

            nome = os.path.basename(caminho)
            self.lbl_caminho.config(
                text=f"Carregado: {nome} ({len(df)} veículos) | Colunas: '{col_p}', '{col_r}'",
                foreground="#90BE9F" if self.modo_noturno else "#15803D",
                font=("Segoe UI", 9, "bold")
            )

            self._popular_tabela(df, col_p, col_r)

            # Reseta métricas
            self.var_metricas_veiculos.set(f"0 / {len(df)}")
            self.var_metricas_lic.set("R$ 0,00")
            self.var_metricas_ipva.set("R$ 0,00")
            self.var_metricas_multas.set("R$ 0,00")
            self.var_metricas_total.set("R$ 0,00")

            self.btn_iniciar.config(state="normal")
            self.btn_reconsultar_falhas.config(state="disabled")
            self.var_status.set(f"Planilha pronta com {len(df)} veículos higienizados. Clique em 'Iniciar Consulta'.")
            self._adicionar_log(f"Planilha '{nome}' carregada e validada com sucesso ({len(df)} linhas).")

        except Exception as e:
            messagebox.showerror("Erro na Leitura da Planilha", str(e))

    def _popular_tabela(self, df: pd.DataFrame, col_p: str, col_r: str):
        """Preenche a Treeview com os veículos da planilha e valida o formato das placas (Melhoria 4)."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.resultados_memoria = {}
        self.indices_em_falha = []

        for idx, row in df.iterrows():
            placa = str(row[col_p]).strip().upper()
            renavam = str(row[col_r]).strip()

            placa_valida = bool(PADRAO_PLACA.match(placa))
            if not placa_valida:
                situacao_inicial = "Placa Inválida"
                tag_inicial = "falha"
                self.indices_em_falha.append(idx)
            else:
                situacao_inicial = "Aguardando..."
                tag_inicial = "pendente"

            self.resultados_memoria[idx] = {
                "lic": "-",
                "ipva": "-",
                "multas": "-",
                "situacao": situacao_inicial
            }

            self.tree.insert(
                "",
                tk.END,
                iid=str(idx),
                values=(idx + 1, placa, renavam, "-", "-", "-", situacao_inicial),
                tags=(tag_inicial,)
            )

        if self.indices_em_falha:
            self._adicionar_log(f"Aviso: {len(self.indices_em_falha)} placas com formato divergente do padrão nacional.")

    def _iniciar_processamento(self):
        """Inicia a consulta completa para todos os veículos da planilha."""
        self._disparar_raspagem(indices_alvo=None)

    def _reconsultar_selecionado(self):
        """Reconsulta apenas o(s) veículo(s) selecionado(s) na tabela (Melhoria 8)."""
        selecionados = self.tree.selection()
        if not selecionados:
            messagebox.showwarning("Seleção Necessária", "Selecione ao menos um veículo na tabela para reconsultar.")
            return
        indices = [int(i) for i in selecionados]
        self._disparar_raspagem(indices_alvo=indices)

    def _reconsultar_apenas_falhas(self):
        """Reconsulta apenas as linhas com falha ou erro identificadas (Melhoria 8)."""
        indices_falhas = []
        for item_id in self.tree.get_children():
            tags = self.tree.item(item_id, "tags")
            if "falha" in tags:
                indices_falhas.append(int(item_id))

        if not indices_falhas:
            messagebox.showinfo("Nenhuma Falha", "Não há veículos marcados com falha na tabela atual.")
            return

        self._disparar_raspagem(indices_alvo=indices_falhas)

    def _disparar_raspagem(self, indices_alvo: list[int] | None = None):
        """Valida pré-requisitos e lança a thread de raspagem para o lote especificado."""
        if self.is_executando:
            messagebox.showwarning("Aviso", "Já existe um processamento em andamento.")
            return

        caminho = self.caminho_arquivo.get()
        if not caminho or not os.path.exists(caminho):
            messagebox.showerror("Aviso", "Selecione uma planilha válida antes de iniciar.")
            return

        c_ipva = self.var_ipva.get()
        c_lic = self.var_lic.get()
        c_multas = self.var_multas.get()

        if not (c_ipva or c_lic or c_multas):
            messagebox.showwarning("Filtro Necessário", "Selecione ao menos um débito para consulta (IPVA, Licenciamento ou Multas).")
            return

        self.is_executando = True
        self.cancelar_solicitado = False

        self.btn_iniciar.config(state="disabled")
        self.btn_selecionar.config(state="disabled")
        self.btn_pasta_destino.config(state="disabled")
        self.btn_reconsultar_falhas.config(state="disabled")
        self.btn_cancelar.config(state="normal")
        self.btn_abrir_excel.config(state="disabled")
        self.btn_abrir_pdf.config(state="disabled")
        self.btn_abrir_pasta.config(state="disabled")
        self.barra_progresso["value"] = 0

        self.notebook.select(0)

        thread = threading.Thread(
            target=self._executar_raspagem_thread,
            args=(caminho, self.var_headless.get(), c_ipva, c_lic, c_multas, self.var_formato_exportacao.get(), indices_alvo),
            daemon=True
        )
        thread.start()

    def _cancelar_processamento(self):
        """Solicita interrupção amigável da rotina de consulta."""
        if messagebox.askyesno("Confirmar Cancelamento", "Deseja realmente interromper as consultas ao concluir o veículo atual?"):
            self.cancelar_solicitado = True
            self.var_status.set("Cancelamento solicitado! Finalizando o veículo atual e salvando parciais...")
            self.btn_cancelar.config(state="disabled")
            self._adicionar_log("Solicitação de cancelamento registrada.")

    def _salvar_checkpoint(self, pasta_destino: str, df: pd.DataFrame):
        """Salva arquivo de progresso incremental / checkpoint (Melhoria 6)."""
        try:
            qtd_total = len(df)
            lics = [self.resultados_memoria.get(i, {}).get("lic", "-") for i in range(qtd_total)]
            ipvas = [self.resultados_memoria.get(i, {}).get("ipva", "-") for i in range(qtd_total)]
            multas = [self.resultados_memoria.get(i, {}).get("multas", "-") for i in range(qtd_total)]

            df_chk = FrotaDataProcessor.consolidar_resultados(df, lics, ipvas, multas, qtd_total)
            chk_path = os.path.join(pasta_destino, "~checkpoint_frota_ce.xlsx")
            df_chk.to_excel(chk_path, index=False)
        except Exception:
            pass

    def _executar_raspagem_thread(
        self,
        caminho_arquivo: str,
        modo_headless: bool,
        consultar_ipva: bool,
        consultar_lic: bool,
        consultar_multas: bool,
        formato_export: str,
        indices_alvo: list[int] | None = None
    ):
        """Rotina assíncrona de consulta, consolidação, salvamento incremental e exportação."""
        pasta_destino = self.pasta_destino_var.get().strip() or os.path.dirname(os.path.abspath(caminho_arquivo))
        os.makedirs(pasta_destino, exist_ok=True)

        arquivo_excel = os.path.join(pasta_destino, "resultados_frota_ce.xlsx")
        arquivo_pdf = os.path.join(pasta_destino, "resultados_frota_ce.pdf")

        scraper = None
        try:
            df = self.df_carregado.copy()
            total_geral = len(df)

            # Define lote de trabalho
            lista_indices = indices_alvo if indices_alvo is not None else list(df.index)
            total_lote = len(lista_indices)

            self.root.after(0, lambda: self.var_status.set("Inicializando motor Selenium / Chrome..."))
            filtros_ativos = []
            if consultar_ipva: filtros_ativos.append("IPVA (Sefaz)")
            if consultar_lic: filtros_ativos.append("Licenciamento (Detran)")
            if consultar_multas: filtros_ativos.append("Multas (Detran)")

            self._adicionar_log(f"Iniciando lote de {total_lote} veículos. Filtros: {', '.join(filtros_ativos)}.")

            scraper = FrotaScraperEngine(headless=modo_headless)
            self.scraper_engine = scraper
            scraper.iniciar()

            for pos_i, index in enumerate(lista_indices):
                if self.cancelar_solicitado:
                    self._adicionar_log("Processamento interrompido pelo usuário. Consolidando veículos já consultados...")
                    break

                row = df.loc[index]
                placa = str(row[self.col_placa]).strip().upper()
                renavam = str(row[self.col_renavam]).strip()
                pos = pos_i + 1
                pct = int((pos_i / total_lote) * 100)

                self.root.after(0, self._marcar_linha_processando, index, placa, renavam)
                self.root.after(0, lambda p=pos, t=total_lote, plc=placa, pc=pct: self._atualizar_progresso_ui(p, t, plc, pc))

                val_detran_lic = "-"
                val_detran_multas = "-"
                val_sefaz_ipva = "-"

                # 1. Consulta Detran-CE (Licenciamento e/ou Multas em sessão única)
                if consultar_lic or consultar_multas:
                    self._adicionar_log(f"[{pos}/{total_lote}] Consultando {placa} no Detran-CE...")
                    val_detran_lic, val_detran_multas = scraper.consultar_detran(
                        placa, renavam, consultar_lic, consultar_multas
                    )
                    time.sleep(0.5)

                # 2. Consulta Sefaz-CE (IPVA)
                if consultar_ipva:
                    self._adicionar_log(f"[{pos}/{total_lote}] Consultando {placa} na Sefaz-CE...")
                    val_sefaz_ipva = scraper.consultar_sefaz_ipva(placa, renavam)
                    time.sleep(0.5)

                # Salva na memória global de resultados
                self.resultados_memoria[index] = {
                    "lic": val_detran_lic,
                    "ipva": val_sefaz_ipva,
                    "multas": val_detran_multas
                }

                # Recalcula somas globais
                soma_lic = sum(formatar_moeda_float(d.get("lic", "0")) for d in self.resultados_memoria.values())
                soma_ipva = sum(formatar_moeda_float(d.get("ipva", "0")) for d in self.resultados_memoria.values())
                soma_multas = sum(formatar_moeda_float(d.get("multas", "0")) for d in self.resultados_memoria.values())
                total_acumulado = soma_lic + soma_ipva + soma_multas
                qtd_concluidos = sum(1 for d in self.resultados_memoria.values() if d.get("lic") != "-" or d.get("ipva") != "-")

                self.root.after(
                    0,
                    self._atualizar_linha_tabela,
                    index,
                    val_detran_lic,
                    val_sefaz_ipva,
                    val_detran_multas,
                    qtd_concluidos,
                    total_geral,
                    soma_lic,
                    soma_ipva,
                    soma_multas,
                    total_acumulado
                )
                self._adicionar_log(f"  -> {placa} | Lic: {val_detran_lic} | IPVA: {val_sefaz_ipva} | Multas: {val_detran_multas}")

                # Checkpoint a cada 5 veículos (Melhoria 6)
                if (pos_i + 1) % 5 == 0:
                    self._salvar_checkpoint(pasta_destino, df)

            # CONSOLIDAÇÃO E EXPORTAÇÃO
            self.root.after(0, lambda: self.var_status.set("Consolidando dados e gerando relatórios..."))

            lics_final = [self.resultados_memoria.get(i, {}).get("lic", "-") for i in range(total_geral)]
            ipvas_final = [self.resultados_memoria.get(i, {}).get("ipva", "-") for i in range(total_geral)]
            multas_final = [self.resultados_memoria.get(i, {}).get("multas", "-") for i in range(total_geral)]

            df_final = FrotaDataProcessor.consolidar_resultados(
                df, lics_final, ipvas_final, multas_final, total_geral
            )

            totais_dict = {
                "lic": formatar_float_moeda(sum(formatar_moeda_float(v) for v in lics_final)),
                "ipva": formatar_float_moeda(sum(formatar_moeda_float(v) for v in ipvas_final)),
                "multas": formatar_float_moeda(sum(formatar_moeda_float(v) for v in multas_final)),
                "total": formatar_float_moeda(
                    sum(formatar_moeda_float(v) for v in lics_final) +
                    sum(formatar_moeda_float(v) for v in ipvas_final) +
                    sum(formatar_moeda_float(v) for v in multas_final)
                )
            }

            # Exporta Excel (com proteção PermissionError)
            if formato_export in ["excel", "ambos"]:
                caminho_excel_salvo = FrotaRelatorioExporter.exportar_excel(df_final, arquivo_excel)
                self.arquivo_excel_gerado = caminho_excel_salvo
                self._adicionar_log(f"Planilha Excel gravada: {caminho_excel_salvo}")

            # Exporta PDF Paisagem (com proteção PermissionError)
            if formato_export in ["pdf", "ambos"]:
                caminho_pdf_salvo = FrotaRelatorioExporter.exportar_pdf(df_final, arquivo_pdf, totais_dict, filtros_ativos)
                self.arquivo_pdf_gerado = caminho_pdf_salvo
                self._adicionar_log(f"Relatório PDF gravado: {caminho_pdf_salvo}")

            # Remove checkpoint temporário após conclusão bem-sucedida
            chk_tmp = os.path.join(pasta_destino, "~checkpoint_frota_ce.xlsx")
            if os.path.exists(chk_tmp):
                try:
                    os.remove(chk_tmp)
                except Exception:
                    pass

            qtd_proc = sum(1 for d in self.resultados_memoria.values() if d.get("lic") != "-" or d.get("ipva") != "-")
            self.root.after(0, self._concluir_sucesso, qtd_proc, total_geral)

        except Exception as e:
            self._adicionar_log(f"ERRO CRÍTICO: {str(e)}")
            # Tenta salvar salvamento de emergência caso haja dados parciais
            try:
                self._salvar_checkpoint(pasta_destino, self.df_carregado)
                self._adicionar_log("Dados parciais gravados no checkpoint com segurança.")
            except Exception:
                pass
            self.root.after(0, self._concluir_erro, str(e))

        finally:
            self.is_executando = False
            if scraper:
                scraper.fechar()
            self.scraper_engine = None
            self.root.after(0, self._restaurar_controles)

    # ==================================================================================
    # MÉTODOS DE ATUALIZAÇÃO THREAD-SAFE (root.after)
    # ==================================================================================

    def _marcar_linha_processando(self, idx: int, placa: str, renavam: str):
        """Destaca o item atual na tabela visual."""
        item_id = str(idx)
        if self.tree.exists(item_id):
            self.tree.item(
                item_id,
                values=(idx + 1, placa, renavam, "...", "...", "...", "⏳ Processando"),
                tags=("processando",)
            )
            self.tree.see(item_id)

    def _atualizar_linha_tabela(
        self,
        idx: int,
        lic: str,
        ipva: str,
        multas: str,
        pos: int,
        total: int,
        s_lic: float,
        s_ipva: float,
        s_multas: float,
        total_g: float
    ):
        """Atualiza a linha do veículo diferenciando Débitos, Regular e Falhas (Melhoria 1)."""
        item_id = str(idx)
        if self.tree.exists(item_id):
            tem_falha = any(
                termo in str(v).lower()
                for v in [lic, ipva, multas]
                for termo in ["falha", "erro", "timeout", "inválid", "não localizado"]
            )
            tem_debito = (
                formatar_moeda_float(lic) > 0 or
                formatar_moeda_float(ipva) > 0 or
                formatar_moeda_float(multas) > 0
            )

            if tem_falha:
                tag = "falha"
                situacao = "⚠️ Falha na Consulta"
            elif tem_debito:
                tag = "com_debito"
                situacao = "⚠️ Com Débitos"
            else:
                tag = "sem_debito"
                situacao = "✔️ Sem Débitos"

            valores_antigos = self.tree.item(item_id, "values")
            self.tree.item(
                item_id,
                values=(valores_antigos[0], valores_antigos[1], valores_antigos[2], lic, ipva, multas, situacao),
                tags=(tag,)
            )

        # Atualiza Mini-Cards
        self.var_metricas_veiculos.set(f"{pos} / {total}")
        self.var_metricas_lic.set(formatar_float_moeda(s_lic))
        self.var_metricas_ipva.set(formatar_float_moeda(s_ipva))
        self.var_metricas_multas.set(formatar_float_moeda(s_multas))
        self.var_metricas_total.set(formatar_float_moeda(total_g))

    def _atualizar_progresso_ui(self, pos: int, total: int, placa: str, pct: int):
        """Atualiza a barra e a descrição de progresso."""
        self.var_status.set(f"Consultando [{pos}/{total}]: Veículo {placa}...")
        self.barra_progresso.configure(value=pct)

    def _restaurar_controles(self):
        """Reabilita os botões da interface e verifica itens em falha (Melhoria 8)."""
        self.btn_selecionar.config(state="normal")
        self.btn_pasta_destino.config(state="normal")
        self.btn_iniciar.config(state="normal")
        self.btn_cancelar.config(state="disabled")

        if self.arquivo_excel_gerado and os.path.exists(self.arquivo_excel_gerado):
            self.btn_abrir_excel.config(state="normal")
        if self.arquivo_pdf_gerado and os.path.exists(self.arquivo_pdf_gerado):
            self.btn_abrir_pdf.config(state="normal")
        if self.arquivo_excel_gerado or self.arquivo_pdf_gerado:
            self.btn_abrir_pasta.config(state="normal")

        # Verifica se há falhas na tabela para habilitar reconsulta rápida
        tem_falhas = any("falha" in self.tree.item(item_id, "tags") for item_id in self.tree.get_children())
        self.btn_reconsultar_falhas.config(state="normal" if tem_falhas else "disabled")

    def _concluir_sucesso(self, concluidos: int, total: int):
        """Feedback de sucesso."""
        self.barra_progresso.configure(value=100)
        tem_falhas = any("falha" in self.tree.item(item_id, "tags") for item_id in self.tree.get_children())

        if tem_falhas:
            msg = f"Processamento concluído com alertas ({concluidos}/{total} processados)."
        else:
            msg = "Processamento concluído com 100% de sucesso!"

        self.var_status.set(msg)

        msg_corpo = f"{msg}\n\nArquivos gerados com sucesso:\n"
        if self.arquivo_excel_gerado:
            msg_corpo += f"- Planilha: {os.path.basename(self.arquivo_excel_gerado)}\n"
        if self.arquivo_pdf_gerado:
            msg_corpo += f"- Relatório PDF: {os.path.basename(self.arquivo_pdf_gerado)}\n"

        if tem_falhas:
            msg_corpo += "\n💡 Dica: Você pode clicar em '⚠️ Reconsultar Falhas' para tentar os veículos não concluídos."

        messagebox.showinfo("Consulta Finalizada", msg_corpo)

    def _concluir_erro(self, erro: str):
        """Feedback de falha."""
        self.var_status.set("Erro durante o processamento.")
        messagebox.showerror(
            "Erro de Execução",
            f"Ocorreu uma falha durante o processo:\n\n{erro}\n\n"
            f"Os dados consultados até o momento foram preservados no checkpoint."
        )

    def _abrir_excel_gerado(self):
        """Abre a planilha Excel gerada."""
        if self.arquivo_excel_gerado and os.path.exists(self.arquivo_excel_gerado):
            try:
                os.startfile(self.arquivo_excel_gerado)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o arquivo Excel:\n{e}")
        else:
            messagebox.showwarning("Aviso", "O arquivo Excel ainda não foi gerado.")

    def _abrir_pdf_gerado(self):
        """Abre o relatório PDF gerado."""
        if self.arquivo_pdf_gerado and os.path.exists(self.arquivo_pdf_gerado):
            try:
                os.startfile(self.arquivo_pdf_gerado)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o arquivo PDF:\n{e}")
        else:
            messagebox.showwarning("Aviso", "O relatório PDF ainda não foi gerado.")

    def _abrir_pasta_destino(self):
        """Abre a pasta contendo os arquivos no Windows Explorer."""
        caminho_base = self.arquivo_excel_gerado or self.arquivo_pdf_gerado or self.pasta_destino_var.get()
        if caminho_base and os.path.exists(caminho_base):
            try:
                if os.path.isfile(caminho_base):
                    subprocess.Popen(f'explorer /select,"{os.path.normpath(caminho_base)}"')
                else:
                    os.startfile(caminho_base)
            except Exception:
                pasta = os.path.dirname(caminho_base) if os.path.isfile(caminho_base) else caminho_base
                os.startfile(pasta)
        else:
            messagebox.showwarning("Aviso", "A pasta de destino ainda não possui arquivos gerados.")

    def _ao_fechar_janela(self):
        """Encerramento seguro e sem processos órfãos."""
        if self.is_executando:
            if not messagebox.askyesno(
                "Consulta em Andamento",
                "O processamento ainda está em execução.\nDeseja realmente abortar e fechar a aplicação?"
            ):
                return
            self.cancelar_solicitado = True
            if self.scraper_engine:
                self.scraper_engine.fechar()

        self.root.destroy()
        sys.exit(0)


# ======================================================================================
# PONTO DE ENTRADA PRINCIPAL
# ======================================================================================

def main():
    ocultar_console()
    try:
        root = tk.Tk()
        app = AppConsultaFrotaGUI(root)
        root.mainloop()
    except Exception as e:
        import traceback
        err_msg = f"Falha ao iniciar o aplicativo:\n\n{e}\n\nDetalhes:\n{traceback.format_exc()}"
        try:
            from tkinter import messagebox
            root_err = tk.Tk()
            root_err.withdraw()
            messagebox.showerror("Erro de Inicialização", err_msg)
            root_err.destroy()
        except Exception:
            print(err_msg, file=sys.stderr)


if __name__ == "__main__":
    main()
