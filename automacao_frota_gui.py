"""
========================================================================================
Aplicativo GUI: Consulta de Frota - Detran/Sefaz CE
========================================================================================
Interface Gráfica (Tkinter) para automação de consultas veiculares em lote no Ceará.
Executa o Web Scraping (Selenium + Pandas) em thread dedicada para manter a interface responsiva.

Stack:
- Python 3
- Tkinter (Interface Gráfica Nativa)
- Selenium WebDriver + webdriver-manager (Automação Web)
- Pandas + Openpyxl (Processamento e Exportação de Planilhas)
- Threading (Execução Assíncrona e Segura)
========================================================================================
"""

import os
import sys
import time
import threading
from datetime import datetime
import pandas as pd

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager


# ======================================================================================
# FUNÇÕES DO MOTOR DE SCRAPING (SELENIUM / DETRAN / SEFAZ)
# ======================================================================================

def configurar_driver(headless: bool = True) -> webdriver.Chrome:
    """Configura e inicializa o Chrome WebDriver com parâmetros corporativos."""
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless=new")

    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    )

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


def clicar_elemento_seguro(driver: webdriver.Chrome, elemento) -> None:
    """Executa o clique com scrollIntoView e fallback JavaScript para evitar interceptações de UI."""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
        elemento.click()
    except Exception:
        driver.execute_script("arguments[0].click();", elemento)


def consultar_detran(driver: webdriver.Chrome, wait: WebDriverWait, placa: str, renavam: str) -> str:
    """Consulta o valor do Licenciamento em aberto no portal Detran-CE."""
    url_detran = "https://sistemas.detran.ce.gov.br/central/#veiculo"
    try:
        driver.get(url_detran)

        # 1. Clica no menu lateral 'Licenciamento'
        menu_lic = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Licenciamento')]")))
        clicar_elemento_seguro(driver, menu_lic)

        # 2. Localiza inputs visíveis (0: Placa, 1: Renavam)
        wait.until(lambda d: len([i for i in d.find_elements(By.TAG_NAME, "input") if i.is_displayed()]) >= 2)
        inputs_visiveis = [i for i in driver.find_elements(By.TAG_NAME, "input") if i.is_displayed()]

        inputs_visiveis[0].clear()
        inputs_visiveis[0].send_keys(str(placa).strip().upper())
        inputs_visiveis[1].clear()
        inputs_visiveis[1].send_keys(str(renavam).strip())

        # 3. Clica no botão CONFIRMAR
        btn_confirmar = wait.until(
            EC.presence_of_element_located((
                By.XPATH, "//button[contains(text(), 'CONFIRMAR') or contains(., 'CONFIRMAR') or @id='submit-veiculo']"
            ))
        )
        clicar_elemento_seguro(driver, btn_confirmar)

        # 4. Clica no link de emissão de licenciamento
        link_emissao = wait.until(
            EC.presence_of_element_located((
                By.XPATH, "//*[contains(text(), 'Clique AQUI se deseja imprimir seu licenciamento.')]"
            ))
        )
        clicar_elemento_seguro(driver, link_emissao)

        # 5. Extrai totalização no modal
        modal_total = wait.until(
            EC.visibility_of_element_located((By.XPATH, "//tr[contains(., 'TOTAL')]/td[last()]"))
        )
        valor = modal_total.text.strip()
        return valor if valor else "R$ 0,00"

    except TimeoutException:
        return "R$ 0,00"
    except Exception:
        return "R$ 0,00"


def consultar_sefaz(driver: webdriver.Chrome, wait: WebDriverWait, placa: str, renavam: str) -> str:
    """Consulta o saldo de IPVA em aberto no portal Sefaz-CE."""
    url_sefaz = "https://ipva.sefaz.ce.gov.br/#/impostos/emitir-dae"
    try:
        driver.get(url_sefaz)

        # 1. Preenche Placa
        xpath_placa = "//input[contains(@ng-model, 'placa') or @name='placa' or @formcontrolname='placa' or @id='mat-input-1']"
        placa_inp = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_placa)))
        placa_inp.clear()
        placa_inp.send_keys(str(placa).strip().upper())

        # 2. Preenche Renavam
        xpath_renavam = "//input[contains(@ng-model, 'renavam') or @name='renavam' or @formcontrolname='renavam' or @id='mat-input-2']"
        renavam_inp = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_renavam)))
        renavam_inp.clear()
        renavam_inp.send_keys(str(renavam).strip())

        # 3. Clica em Pesquisar
        xpath_btn = "//button[contains(translate(., 'pesquisar', 'PESQUISAR'), 'PESQUISAR') or contains(@class, 'search-button')]"
        btn_pesq = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn)))
        clicar_elemento_seguro(driver, btn_pesq)

        # 4. Aguarda tabela de débitos não quitados
        xpath_tabela = "//*[contains(text(), 'Débitos de IPVA não quitados')]"
        wait.until(EC.visibility_of_element_located((By.XPATH, xpath_tabela)))

        # 5. Extrai valor em aberto
        xpath_valor = "//td[contains(., 'Em aberto')]/preceding-sibling::td[contains(., 'R$')][1]"
        saldo_aberto = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_valor)))
        valor = saldo_aberto.text.strip()
        return valor if valor else "R$ 0,00"

    except TimeoutException:
        return "R$ 0,00"
    except Exception:
        return "R$ 0,00"


def formatar_planilha_saida(caminho_arquivo: str) -> None:
    """Aplica formatação executiva (cabeçalhos, bordas e alinhamentos) na planilha Excel de saída."""
    from openpyxl import load_workbook
    wb = load_workbook(caminho_arquivo)
    ws = wb.active

    cor_cabecalho = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fonte_linhas = Font(name="Calibri", size=10)
    alinhamento_centro = Alignment(horizontal="center", vertical="center")
    alinhamento_direita = Alignment(horizontal="right", vertical="center")
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
        for idx, cell in enumerate(row):
            cell.font = fonte_linhas
            cell.border = borda_fina
            if idx in [0, 1]:
                cell.alignment = alinhamento_centro
            else:
                cell.alignment = alinhamento_direita

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    wb.save(caminho_arquivo)


# ======================================================================================
# CLASSE PRINCIPAL DA INTERFACE GRÁFICA (TKINTER)
# ======================================================================================

class AppConsultaFrota:
    """Interface Gráfica para Automação de Frota (Detran-CE e Sefaz-CE)."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Consulta de Frota - Detran/Sefaz CE")
        self.root.geometry("640x580")
        self.root.minsize(580, 520)

        # Variáveis de Estado
        self.caminho_arquivo = tk.StringVar(value="")
        self.var_headless = tk.BooleanVar(value=True)
        self.var_status = tk.StringVar(value="Status: Aguardando seleção de planilha...")
        self.is_executando = False
        self.driver_ativo = None

        # Estilo ttk
        self.style = ttk.Style()
        self._configurar_estilos()

        # Construção dos Componentes Visuais
        self._criar_layout()

        # Evento de fechamento de janela
        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar_janela)

    def _configurar_estilos(self):
        """Define o esquema visual moderno da interface."""
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.root.configure(bg="#F4F6F9")
        self.style.configure(".", background="#F4F6F9", font=("Segoe UI", 10))
        self.style.configure("TLabel", background="#F4F6F9", foreground="#333333")
        self.style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), foreground="#1F497D")
        self.style.configure("SubHeader.TLabel", font=("Segoe UI", 9), foreground="#666666")
        self.style.configure("TLabelframe", background="#F4F6F9")
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#1F497D", background="#F4F6F9")
        
        # Botões estilizados
        self.style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            background="#1F497D",
            foreground="#FFFFFF",
            padding=8
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", "#16365C"), ("disabled", "#CCCCCC")],
            foreground=[("disabled", "#888888")]
        )

        self.style.configure(
            "Action.TButton",
            font=("Segoe UI", 11, "bold"),
            background="#2E7D32",
            foreground="#FFFFFF",
            padding=10
        )
        self.style.map(
            "Action.TButton",
            background=[("active", "#1B5E20"), ("disabled", "#CCCCCC")],
            foreground=[("disabled", "#888888")]
        )

    def _criar_layout(self):
        """Monta o grid de widgets com espaçamentos adequados."""
        container = ttk.Frame(self.root, padding="20 20 20 20")
        container.pack(fill=tk.BOTH, expand=True)

        # 1. Cabeçalho
        header_frame = ttk.Frame(container)
        header_frame.pack(fill=tk.X, pady=(0, 15))

        lbl_titulo = ttk.Label(header_frame, text="🚗 Consulta de Frota - Detran/Sefaz CE", style="Header.TLabel")
        lbl_titulo.pack(anchor="w")

        lbl_subtitulo = ttk.Label(
            header_frame,
            text="Automação de extração de débitos em aberto de Licenciamento e IPVA.",
            style="SubHeader.TLabel"
        )
        lbl_subtitulo.pack(anchor="w", pady=(2, 0))

        # 2. Seção de Arquivo de Entrada
        frame_arquivo = ttk.LabelFrame(container, text="1. Planilha de Veículos (Origem)", padding="12 12 12 12")
        frame_arquivo.pack(fill=tk.X, pady=(0, 12))

        btn_box = ttk.Frame(frame_arquivo)
        btn_box.pack(fill=tk.X)

        self.btn_selecionar = ttk.Button(
            btn_box,
            text="📁 Selecionar Planilha de Frota (.xlsx)",
            style="Primary.TButton",
            command=self._selecionar_arquivo
        )
        self.btn_selecionar.pack(side=tk.LEFT)

        self.lbl_caminho = ttk.Label(
            frame_arquivo,
            text="Nenhum arquivo selecionado",
            font=("Segoe UI", 9, "italic"),
            foreground="#777777",
            wraplength=560
        )
        self.lbl_caminho.pack(anchor="w", pady=(8, 0))

        # 3. Seção de Opções de Execução
        frame_opcoes = ttk.LabelFrame(container, text="2. Configurações de Navegação", padding="12 10 12 10")
        frame_opcoes.pack(fill=tk.X, pady=(0, 12))

        chk_headless = ttk.Checkbutton(
            frame_opcoes,
            text="Executar navegador em segundo plano (invisível)?",
            variable=self.var_headless
        )
        chk_headless.pack(anchor="w")

        # 4. Botão de Ação Principal
        frame_acao = ttk.Frame(container)
        frame_acao.pack(fill=tk.X, pady=(0, 15))

        self.btn_iniciar = ttk.Button(
            frame_acao,
            text="🚀 Iniciar Consulta de Débitos",
            style="Action.TButton",
            state="disabled",
            command=self._iniciar_processamento
        )
        self.btn_iniciar.pack(fill=tk.X)

        # 5. Seção de Acompanhamento / Status e Logs
        frame_status = ttk.LabelFrame(container, text="3. Progresso e Monitoramento", padding="12 10 12 10")
        frame_status.pack(fill=tk.BOTH, expand=True)

        self.lbl_status = ttk.Label(
            frame_status,
            textvariable=self.var_status,
            font=("Segoe UI", 10, "bold"),
            foreground="#1F497D"
        )
        self.lbl_status.pack(anchor="w", pady=(0, 6))

        # Barra de Progresso
        self.barra_progresso = ttk.Progressbar(frame_status, mode="determinate")
        self.barra_progresso.pack(fill=tk.X, pady=(0, 8))

        # Caixa de Log
        self.txt_log = tk.Text(
            frame_status,
            height=6,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#FFFFFF",
            fg="#222222",
            relief=tk.SOLID,
            borderwidth=1
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True)
        self._adicionar_log("Pronto. Selecione a planilha para habilitar a consulta.")

    # ==================================================================================
    # EVENTOS E CONTROLE DA GUI
    # ==================================================================================

    def _adicionar_log(self, mensagem: str):
        """Insere uma linha no log textual da interface com rolagem automática."""
        hora = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert(tk.END, f"[{hora}] {mensagem}\n")
        self.txt_log.see(tk.END)

    def _selecionar_arquivo(self):
        """Abre o diálogo de seleção de arquivo filtrando apenas por arquivos Excel."""
        caminho = filedialog.askopenfilename(
            title="Selecione a Planilha de Frota",
            filetypes=[("Arquivos Excel (*.xlsx)", "*.xlsx")]
        )
        if caminho:
            self.caminho_arquivo.set(caminho)
            self.lbl_caminho.config(
                text=f"Arquivo selecionado: {caminho}",
                foreground="#006622",
                font=("Segoe UI", 9, "bold")
            )
            self.btn_iniciar.config(state="normal")
            self.var_status.set("Status: Planilha carregada. Clique em 'Iniciar Consulta'.")
            self._adicionar_log(f"Planilha carregada: {os.path.basename(caminho)}")
        else:
            if not self.caminho_arquivo.get():
                self.btn_iniciar.config(state="disabled")

    def _iniciar_processamento(self):
        """Prepara o estado da interface e dispara a thread do Selenium."""
        caminho = self.caminho_arquivo.get()
        if not caminho or not os.path.exists(caminho):
            messagebox.showerror("Erro", "O arquivo selecionado não foi encontrado.")
            return

        # Bloqueia cliques para evitar duplicidade
        self.is_executando = True
        self.btn_iniciar.config(state="disabled")
        self.btn_selecionar.config(state="disabled")
        self.barra_progresso["value"] = 0

        # Dispara thread de processamento
        thread = threading.Thread(
            target=self._executar_raspagem_em_background,
            args=(caminho, self.var_headless.get()),
            daemon=True
        )
        thread.start()

    def _executar_raspagem_em_background(self, caminho_arquivo: str, modo_headless: bool):
        """
        Rotina executada em thread secundária para não congelar o Tkinter.
        Realiza leitura, consultas web e salvamento dos resultados.
        """
        driver = None
        pasta_destino = os.path.dirname(os.path.abspath(caminho_arquivo))
        arquivo_saida = os.path.join(pasta_destino, "resultados_frota_ce.xlsx")

        try:
            self._atualizar_ui_status("Lendo dados da planilha com Pandas...", 5)
            # Lê o Excel garantindo que RENAVAM mantenha zeros à esquerda
            df_frota = pd.read_excel(caminho_arquivo, dtype={"Placa": str, "RENAVAM": str})

            if "Placa" not in df_frota.columns or "RENAVAM" not in df_frota.columns:
                raise ValueError("A planilha deve conter obrigatoriamente as colunas 'Placa' e 'RENAVAM'.")

            total = len(df_frota)
            if total == 0:
                raise ValueError("A planilha selecionada não possui registros de veículos.")

            self._adicionar_log(f"Iniciando consulta para {total} veículos (Headless={modo_headless})...")
            self._atualizar_ui_status("Inicializando WebDriver do Chrome...", 10)

            # Inicializa Selenium
            driver = configurar_driver(headless=modo_headless)
            self.driver_ativo = driver
            wait = WebDriverWait(driver, 15)

            licenciamentos = []
            ipvas = []

            for index, row in df_frota.iterrows():
                if not self.is_executando:
                    break

                placa = str(row["Placa"]).strip().upper()
                renavam = str(row["RENAVAM"]).strip()
                pos = index + 1
                progresso_pct = 10 + int((pos / total) * 80)

                # Atualiza Label e Log
                msg_status = f"Status: [{pos}/{total}] Consultando placa {placa}..."
                self._atualizar_ui_status(msg_status, progresso_pct)
                self._adicionar_log(f"Consultando veículo {placa} (RENAVAM: {renavam})...")

                # Consulta Detran-CE
                val_detran = consultar_detran(driver, wait, placa, renavam)
                time.sleep(1.0)  # Rate limit

                # Consulta Sefaz-CE
                val_sefaz = consultar_sefaz(driver, wait, placa, renavam)
                time.sleep(1.0)  # Rate limit

                licenciamentos.append(val_detran)
                ipvas.append(val_sefaz)

                self._adicionar_log(f"  -> {placa} | Licenciamento: {val_detran} | IPVA: {val_sefaz}")

            if not self.is_executando:
                return

            # Finaliza gravação da planilha
            self._atualizar_ui_status("Gravando e formatando planilha de saída...", 95)
            df_frota["Licenciamento_CE"] = licenciamentos
            df_frota["IPVA_CE"] = ipvas
            df_frota["Data_Consulta"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            df_frota.to_excel(arquivo_saida, index=False, engine="openpyxl")
            formatar_planilha_saida(arquivo_saida)

            # Sucesso
            self._atualizar_ui_status("Concluído com sucesso!", 100)
            self._adicionar_log(f"Resultados gravados com sucesso em: {arquivo_saida}")
            self.root.after(0, self._concluir_com_sucesso, arquivo_saida)

        except Exception as e:
            self._adicionar_log(f"ERRO: {str(e)}")
            self.root.after(0, self._concluir_com_erro, str(e))

        finally:
            self.is_executando = False
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
            self.driver_ativo = None
            self.root.after(0, self._reabilitar_botoes)

    # ==================================================================================
    # COMUNICAÇÃO SEGURA COM A THREAD PRINCIPAL (root.after)
    # ==================================================================================

    def _atualizar_ui_status(self, texto_status: str, valor_progresso: int):
        """Atualiza a barra de status e o indicador de progresso na thread principal."""
        self.root.after(0, lambda: self.var_status.set(texto_status))
        self.root.after(0, lambda: self.barra_progresso.configure(value=valor_progresso))

    def _reabilitar_botoes(self):
        """Restaura o estado dos botões após término ou falha."""
        self.btn_selecionar.config(state="normal")
        if self.caminho_arquivo.get():
            self.btn_iniciar.config(state="normal")

    def _concluir_com_sucesso(self, arquivo_saida: str):
        """Exibe o popup de sucesso estipulado nos requisitos."""
        messagebox.showinfo(
            "Processo Concluído",
            "Processo concluído com sucesso!\nO arquivo de resultados foi salvo na mesma pasta da planilha original."
        )

    def _concluir_com_erro(self, mensagem_erro: str):
        """Exibe alerta com os detalhes do erro."""
        messagebox.showerror(
            "Falha no Processamento",
            f"Ocorreu um erro durante a consulta:\n\n{mensagem_erro}"
        )
        self.var_status.set("Status: Processamento interrompido devido a erro.")

    def _ao_fechar_janela(self):
        """Garante encerramento seguro do driver ao fechar a janela."""
        if self.is_executando:
            if not messagebox.askyesno(
                "Confirmação",
                "Uma consulta está em andamento. Deseja realmente interromper e fechar o aplicativo?"
            ):
                return
            self.is_executando = False
            if self.driver_ativo:
                try:
                    self.driver_ativo.quit()
                except Exception:
                    pass
        self.root.destroy()


# ======================================================================================
# PONTO DE ENTRADA DA APLICAÇÃO
# ======================================================================================

def main():
    root = tk.Tk()
    app = AppConsultaFrota(root)
    root.mainloop()


if __name__ == "__main__":
    main()
