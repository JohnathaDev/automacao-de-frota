"""
========================================================================================
Aplicativo GUI Profissional: Consulta de Frota - Detran/Sefaz CE (v2.0)
========================================================================================
Interface Gráfica com foco em Experiência do Usuário (UI/UX) para terceiros:
- Pré-visualização da planilha carregada em grade interativa (Treeview).
- Atualização em tempo real das consultas com cores por situação de débito.
- Métricas e Dashboard de totais financeiros (Licenciamento, IPVA e Total Geral).
- Botões pós-processamento: "Abrir Planilha no Excel" e "Abrir Pasta de Destino".
- Otimização de velocidade: detecção precoce de veículos sem débitos (sem espera inútil).
- Controle de cancelamento limpo ("Parar Consulta") e proteção contra cliques duplos.
- Threading dedicada para interface 100% fluida e responsiva.
========================================================================================
"""

import os
import sys
import time
import subprocess
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
# MOTOR DE SCRAPING OTIMIZADO (SELENIUM / DETRAN-CE / SEFAZ-CE)
# ======================================================================================

def configurar_driver(headless: bool = True) -> webdriver.Chrome:
    """Inicializa o Chrome WebDriver com parâmetros otimizados e silenciosos."""
    chrome_options = Options()
    if headless:
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
    """Consulta o valor do Licenciamento no Detran-CE com fallback gracioso."""
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
    """
    Consulta o saldo de IPVA na Sefaz-CE de forma acelerada.
    Se a tabela de 'Quitados' aparecer antes de 'Não quitados', encerra imediatamente sem timeout.
    """
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

        # 4. Otimização UI/UX: Aguarda qualquer uma das respostas da tela
        # (débitos pendentes OU débitos quitados/sem débitos)
        def resposta_carregada(d):
            tem_debitos = d.find_elements(By.XPATH, "//*[contains(text(), 'Débitos de IPVA não quitados')]")
            quitado = d.find_elements(By.XPATH, "//*[contains(text(), 'Débitos de IPVA Regularizados/Quitados') or contains(text(), 'Nenhum débito')]")
            return tem_debitos or quitado

        wait.until(resposta_carregada)

        # Verifica se há tabela de débitos em aberto
        debitos_elementos = driver.find_elements(By.XPATH, "//*[contains(text(), 'Débitos de IPVA não quitados')]")
        if not debitos_elementos:
            # Veículo quitado / sem débitos: retorno imediato!
            return "R$ 0,00"

        # 5. Extrai valor em aberto
        xpath_valor = "//td[contains(., 'Em aberto')]/preceding-sibling::td[contains(., 'R$')][1]"
        saldo_aberto = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_valor)))
        valor = saldo_aberto.text.strip()
        return valor if valor else "R$ 0,00"

    except TimeoutException:
        return "R$ 0,00"
    except Exception:
        return "R$ 0,00"


def formatar_moeda_float(valor_str: str) -> float:
    """Converte strings financeiras como 'R$ 188,96' em float 188.96."""
    try:
        limpo = valor_str.replace("R$", "").replace(".", "").replace(",", ".").strip()
        return float(limpo)
    except Exception:
        return 0.0


def formatar_planilha_saida(caminho_arquivo: str) -> None:
    """Aplica formatação executiva na planilha Excel de saída."""
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
# INTERFACE GRÁFICA PROFISSIONAL (TKINTER + TTK)
# ======================================================================================

class AppConsultaFrota:
    """Interface Gráfica otimizada para terceiros e usuários corporativos."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Consulta de Frota - Detran/Sefaz CE")
        self.root.geometry("880x690")
        self.root.minsize(780, 600)

        # Variáveis de Estado
        self.caminho_arquivo = tk.StringVar(value="")
        self.arquivo_saida_gerado = ""
        self.var_headless = tk.BooleanVar(value=True)
        self.var_status = tk.StringVar(value="Pronto para iniciar. Selecione a planilha de veículos.")
        self.var_metricas_veiculos = tk.StringVar(value="0 / 0")
        self.var_metricas_lic = tk.StringVar(value="R$ 0,00")
        self.var_metricas_ipva = tk.StringVar(value="R$ 0,00")
        self.var_metricas_total = tk.StringVar(value="R$ 0,00")

        self.is_executando = False
        self.cancelar_solicitado = False
        self.driver_ativo = None
        self.df_carregado = None

        # Estilos Visuais Modernos
        self.style = ttk.Style()
        self._configurar_estilos()

        # Construção da Interface
        self._criar_layout()

        # Protocolo de Encerramento Seguro
        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar_janela)

    def _configurar_estilos(self):
        """Define cores corporativas e fontes modernas."""
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.root.configure(bg="#F0F4F8")
        self.style.configure(".", background="#F0F4F8", font=("Segoe UI", 10))
        self.style.configure("TLabel", background="#F0F4F8", foreground="#2D3748")
        
        # Cabeçalhos
        self.style.configure("Header.TLabel", font=("Segoe UI", 15, "bold"), foreground="#1A365D")
        self.style.configure("SubHeader.TLabel", font=("Segoe UI", 9), foreground="#4A5568")
        
        # Cards / LabelFrames
        self.style.configure("Card.TLabelframe", background="#FFFFFF", relief=tk.SOLID, borderwidth=1)
        self.style.configure("Card.TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#1A365D", background="#F0F4F8")

        # Botão Primário (Selecionar)
        self.style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            background="#2B6CB0",
            foreground="#FFFFFF",
            padding=(12, 7)
        )
        self.style.map("Primary.TButton", background=[("active", "#2C5282"), ("disabled", "#CBD5E0")])

        # Botão de Execução (Verde)
        self.style.configure(
            "Action.TButton",
            font=("Segoe UI", 10, "bold"),
            background="#2E7D32",
            foreground="#FFFFFF",
            padding=(15, 8)
        )
        self.style.map("Action.TButton", background=[("active", "#1B5E20"), ("disabled", "#CBD5E0")])

        # Botão de Cancelamento (Vermelho)
        self.style.configure(
            "Danger.TButton",
            font=("Segoe UI", 10, "bold"),
            background="#C53030",
            foreground="#FFFFFF",
            padding=(12, 8)
        )
        self.style.map("Danger.TButton", background=[("active", "#9B2C2C"), ("disabled", "#CBD5E0")])

        # Botões de Ação Final (Cinza Corporativo)
        self.style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9, "bold"),
            background="#4A5568",
            foreground="#FFFFFF",
            padding=(10, 6)
        )
        self.style.map("Secondary.TButton", background=[("active", "#2D3748"), ("disabled", "#E2E8F0")])

        # Estilo da Tabela Treeview
        self.style.configure(
            "Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground="#2D3748",
            rowheight=26,
            font=("Segoe UI", 9)
        )
        self.style.configure(
            "Treeview.Heading",
            background="#E2E8F0",
            foreground="#1A365D",
            font=("Segoe UI", 9, "bold")
        )
        self.style.map("Treeview", background=[("selected", "#BEE3F8")], foreground=[("selected", "#1A365D")])

    def _criar_layout(self):
        """Monta o grid estruturado com hierarquia visual clara."""
        container = ttk.Frame(self.root, padding="15 15 15 15")
        container.pack(fill=tk.BOTH, expand=True)

        # -------------------------------------------------------------------------
        # 1. CABEÇALHO COM TÍTULO E VERSÃO
        # -------------------------------------------------------------------------
        header = ttk.Frame(container)
        header.pack(fill=tk.X, pady=(0, 10))

        lbl_tit = ttk.Label(header, text="🚗 Consulta de Frota - Detran/Sefaz CE", style="Header.TLabel")
        lbl_tit.pack(side=tk.LEFT)

        lbl_sub = ttk.Label(
            header,
            text="Automação Inteligente de Débitos de Licenciamento & IPVA",
            style="SubHeader.TLabel"
        )
        lbl_sub.pack(side=tk.LEFT, padx=(15, 0), pady=(4, 0))

        # -------------------------------------------------------------------------
        # 2. CARD: SELEÇÃO DE PLANILHA E CONTROLES
        # -------------------------------------------------------------------------
        frame_topo = ttk.LabelFrame(container, text=" Configurações e Seleção ", style="Card.TLabelframe", padding="12 10 12 10")
        frame_topo.pack(fill=tk.X, pady=(0, 10))

        linha1 = ttk.Frame(frame_topo)
        linha1.pack(fill=tk.X, pady=(0, 8))

        self.btn_selecionar = ttk.Button(
            linha1,
            text="📁 Selecionar Planilha (.xlsx)",
            style="Primary.TButton",
            command=self._selecionar_arquivo
        )
        self.btn_selecionar.pack(side=tk.LEFT, padx=(0, 10))

        self.lbl_caminho = ttk.Label(
            linha1,
            text="Nenhum arquivo selecionado",
            font=("Segoe UI", 9, "italic"),
            foreground="#718096"
        )
        self.lbl_caminho.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Linha de Ações e Opções
        linha2 = ttk.Frame(frame_topo)
        linha2.pack(fill=tk.X)

        self.chk_headless = ttk.Checkbutton(
            linha2,
            text="Navegador invisível em segundo plano (Recomendado)",
            variable=self.var_headless
        )
        self.chk_headless.pack(side=tk.LEFT, pady=2)

        # Botões de Ação (Iniciar / Cancelar)
        self.btn_cancelar = ttk.Button(
            linha2,
            text="⏹️ Cancelar Consulta",
            style="Danger.TButton",
            state="disabled",
            command=self._cancelar_processamento
        )
        self.btn_cancelar.pack(side=tk.RIGHT, padx=(6, 0))

        self.btn_iniciar = ttk.Button(
            linha2,
            text="▶️ Iniciar Consulta",
            style="Action.TButton",
            state="disabled",
            command=self._iniciar_processamento
        )
        self.btn_iniciar.pack(side=tk.RIGHT)

        # -------------------------------------------------------------------------
        # 3. CARD: TABELA INTERATIVA EM TEMPO REAL E LOGS
        # -------------------------------------------------------------------------
        frame_centro = ttk.LabelFrame(container, text=" Acompanhamento em Tempo Real ", style="Card.TLabelframe", padding="10 10 10 10")
        frame_centro.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Caderno de Abas (Notebook)
        self.notebook = ttk.Notebook(frame_centro)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Aba 1: Tabela Visual
        aba_tabela = ttk.Frame(self.notebook, padding="6 6 6 6")
        self.notebook.add(aba_tabela, text=" 📊 Tabela de Veículos ")

        colunas = ("num", "placa", "renavam", "licenciamento", "ipva", "situacao")
        self.tree = ttk.Treeview(aba_tabela, columns=colunas, show="headings", selectmode="browse")

        self.tree.heading("num", text="Item")
        self.tree.heading("placa", text="Placa")
        self.tree.heading("renavam", text="RENAVAM")
        self.tree.heading("licenciamento", text="Licenciamento (Detran)")
        self.tree.heading("ipva", text="IPVA em Aberto (Sefaz)")
        self.tree.heading("situacao", text="Situação")

        self.tree.column("num", width=45, anchor="center")
        self.tree.column("placa", width=95, anchor="center")
        self.tree.column("renavam", width=125, anchor="center")
        self.tree.column("licenciamento", width=150, anchor="e")
        self.tree.column("ipva", width=150, anchor="e")
        self.tree.column("situacao", width=150, anchor="center")

        # Scrollbars para a Tabela
        scroll_y = ttk.Scrollbar(aba_tabela, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # Tags de Cor para Linhas
        self.tree.tag_configure("pendente", foreground="#718096")
        self.tree.tag_configure("processando", background="#EBF8FF", foreground="#2B6CB0", font=("Segoe UI", 9, "bold"))
        self.tree.tag_configure("com_debito", background="#FFF5F5", foreground="#C53030", font=("Segoe UI", 9, "bold"))
        self.tree.tag_configure("sem_debito", background="#F0FFF4", foreground="#22543D")

        # Aba 2: Log Textual
        aba_log = ttk.Frame(self.notebook, padding="6 6 6 6")
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

        self._adicionar_log("Sistema iniciado. Aguardando carregamento da planilha.")

        # -------------------------------------------------------------------------
        # 4. DASHBOARD DE MÉTRICAS E TOTAIS FINANCEIROS
        # -------------------------------------------------------------------------
        frame_metricas = ttk.Frame(container)
        frame_metricas.pack(fill=tk.X, pady=(0, 10))

        # 4 Mini-Cards de Estatísticas
        def criar_card_metrica(pai, titulo, variavel, cor_destaque):
            card = tk.Frame(pai, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=12, pady=6)
            card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
            lbl_t = tk.Label(card, text=titulo, font=("Segoe UI", 8, "bold"), fg="#718096", bg="#FFFFFF")
            lbl_t.pack(anchor="w")
            lbl_v = tk.Label(card, textvariable=variavel, font=("Segoe UI", 12, "bold"), fg=cor_destaque, bg="#FFFFFF")
            lbl_v.pack(anchor="w")
            return card

        criar_card_metrica(frame_metricas, "VEÍCULOS CONSULTADOS", self.var_metricas_veiculos, "#2B6CB0")
        criar_card_metrica(frame_metricas, "TOTAL LICENCIAMENTO", self.var_metricas_lic, "#D69E2E")
        criar_card_metrica(frame_metricas, "TOTAL IPVA EM ABERTO", self.var_metricas_ipva, "#E53E3E")
        criar_card_metrica(frame_metricas, "TOTAL GERAL A PAGAR", self.var_metricas_total, "#2E7D32")

        # -------------------------------------------------------------------------
        # 5. BARRA DE PROGRESSO E AÇÕES PÓS-PROCESSAMENTO
        # -------------------------------------------------------------------------
        frame_rodape = ttk.Frame(container)
        frame_rodape.pack(fill=tk.X)

        self.lbl_status = ttk.Label(
            frame_rodape,
            textvariable=self.var_status,
            font=("Segoe UI", 9, "bold"),
            foreground="#1A365D"
        )
        self.lbl_status.pack(anchor="w", pady=(0, 4))

        self.barra_progresso = ttk.Progressbar(frame_rodape, mode="determinate")
        self.barra_progresso.pack(fill=tk.X, pady=(0, 8))

        # Botões de Acesso Rápido aos Resultados
        frame_botoes_finais = ttk.Frame(frame_rodape)
        frame_botoes_finais.pack(fill=tk.X)

        self.btn_abrir_excel = ttk.Button(
            frame_botoes_finais,
            text="📊 Abrir Planilha Gerada no Excel",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_planilha_gerada
        )
        self.btn_abrir_excel.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_abrir_pasta = ttk.Button(
            frame_botoes_finais,
            text="📂 Abrir Pasta de Destino",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_pasta_destino
        )
        self.btn_abrir_pasta.pack(side=tk.LEFT)

    # ==================================================================================
    # MÉTODOS DE CONTROLE E INTERAÇÃO
    # ==================================================================================

    def _adicionar_log(self, mensagem: str):
        """Insere uma linha com carimbo de hora no terminal de log."""
        hora = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert(tk.END, f"[{hora}] {mensagem}\n")
        self.txt_log.see(tk.END)

    def _selecionar_arquivo(self):
        """Abre o explorador para escolher a planilha e a pré-carrega na tabela visual."""
        caminho = filedialog.askopenfilename(
            title="Selecione a Planilha de Frota",
            filetypes=[("Arquivos Excel (*.xlsx)", "*.xlsx")]
        )
        if not caminho:
            return

        try:
            # Validação e Leitura prévia
            df = pd.read_excel(caminho, dtype={"Placa": str, "RENAVAM": str})
            if "Placa" not in df.columns or "RENAVAM" not in df.columns:
                messagebox.showerror(
                    "Formato Inválido",
                    "A planilha selecionada deve conter obrigatoriamente as colunas:\n- Placa\n- RENAVAM"
                )
                return

            if len(df) == 0:
                messagebox.showwarning("Aviso", "A planilha selecionada não possui registros.")
                return

            self.df_carregado = df
            self.caminho_arquivo.set(caminho)
            nome_arquivo = os.path.basename(caminho)
            self.lbl_caminho.config(
                text=f"Carregado: {nome_arquivo} ({len(df)} veículos)",
                foreground="#22543D",
                font=("Segoe UI", 9, "bold")
            )

            # Preenche a Tabela Treeview com os veículos pendentes
            self._popular_tabela(df)

            # Atualiza métricas iniciais
            self.var_metricas_veiculos.set(f"0 / {len(df)}")
            self.var_metricas_lic.set("R$ 0,00")
            self.var_metricas_ipva.set("R$ 0,00")
            self.var_metricas_total.set("R$ 0,00")

            self.btn_iniciar.config(state="normal")
            self.var_status.set(f"Planilha pronta com {len(df)} veículos. Clique em 'Iniciar Consulta'.")
            self._adicionar_log(f"Planilha '{nome_arquivo}' carregada com {len(df)} veículos.")

        except Exception as e:
            messagebox.showerror("Erro de Leitura", f"Não foi possível ler o arquivo Excel:\n\n{str(e)}")

    def _popular_tabela(self, df: pd.DataFrame):
        """Limpa e insere os veículos carregados na Treeview com estado Pendente."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, row in df.iterrows():
            placa = str(row["Placa"]).strip().upper()
            renavam = str(row["RENAVAM"]).strip()
            self.tree.insert(
                "",
                tk.END,
                iid=str(idx),
                values=(idx + 1, placa, renavam, "-", "-", "Aguardando..."),
                tags=("pendente",)
            )

    def _iniciar_processamento(self):
        """Configura os botões e dispara a thread com o Selenium."""
        caminho = self.caminho_arquivo.get()
        if not caminho or not os.path.exists(caminho):
            messagebox.showerror("Erro", "O arquivo selecionado não foi encontrado.")
            return

        self.is_executando = True
        self.cancelar_solicitado = False

        # Ajuste dos botões de controle
        self.btn_iniciar.config(state="disabled")
        self.btn_selecionar.config(state="disabled")
        self.btn_cancelar.config(state="normal")
        self.btn_abrir_excel.config(state="disabled")
        self.btn_abrir_pasta.config(state="disabled")
        self.barra_progresso["value"] = 0

        # Altera para a aba da tabela automaticamente
        self.notebook.select(0)

        # Inicia Threading
        thread = threading.Thread(
            target=self._executar_raspagem_em_background,
            args=(caminho, self.var_headless.get()),
            daemon=True
        )
        thread.start()

    def _cancelar_processamento(self):
        """Solicita interrupção amigável do processo."""
        if messagebox.askyesno("Confirmar Cancelamento", "Deseja realmente interromper as consultas ao finalizar o veículo atual?"):
            self.cancelar_solicitado = True
            self.var_status.set("Interrupção solicitada! Concluindo consulta atual...")
            self.btn_cancelar.config(state="disabled")
            self._adicionar_log("Solicitação de cancelamento enviada pelo usuário.")

    def _executar_raspagem_em_background(self, caminho_arquivo: str, modo_headless: bool):
        """Rotina assíncrona do Web Scraping."""
        driver = None
        pasta_destino = os.path.dirname(os.path.abspath(caminho_arquivo))
        arquivo_saida = os.path.join(pasta_destino, "resultados_frota_ce.xlsx")
        self.arquivo_saida_gerado = arquivo_saida

        try:
            df = self.df_carregado.copy()
            total = len(df)

            self.root.after(0, lambda: self.var_status.set("Inicializando navegador Chrome..."))
            self._adicionar_log(f"Iniciando consultas para {total} veículos (Modo invisível={modo_headless})...")

            driver = configurar_driver(headless=modo_headless)
            self.driver_ativo = driver
            wait = WebDriverWait(driver, 15)

            licenciamentos = []
            ipvas = []
            soma_lic = 0.0
            soma_ipva = 0.0

            for index, row in df.iterrows():
                if self.cancelar_solicitado:
                    self._adicionar_log("Processamento interrompido pelo usuário.")
                    break

                placa = str(row["Placa"]).strip().upper()
                renavam = str(row["RENAVAM"]).strip()
                pos = index + 1
                progresso_pct = int(((pos - 1) / total) * 100)

                # Marca linha atual como "Processando"
                self.root.after(0, self._marcar_linha_processando, index, placa, renavam)
                self.root.after(0, lambda p=pos, t=total, plc=placa, pct=progresso_pct: self._atualizar_progresso_ui(p, t, plc, pct))

                # 1. Consulta Detran
                self._adicionar_log(f"[{pos}/{total}] Consultando {placa} no Detran-CE...")
                val_detran = consultar_detran(driver, wait, placa, renavam)
                time.sleep(1.0)

                # 2. Consulta Sefaz
                self._adicionar_log(f"[{pos}/{total}] Consultando {placa} na Sefaz-CE...")
                val_sefaz = consultar_sefaz(driver, wait, placa, renavam)
                time.sleep(1.0)

                licenciamentos.append(val_detran)
                ipvas.append(val_sefaz)

                # Atualiza somas
                soma_lic += formatar_moeda_float(val_detran)
                soma_ipva += formatar_moeda_float(val_sefaz)
                total_geral = soma_lic + soma_ipva

                # Atualiza linha na tabela Treeview e Dashboard
                self.root.after(
                    0,
                    self._atualizar_linha_tabela,
                    index,
                    val_detran,
                    val_sefaz,
                    pos,
                    total,
                    soma_lic,
                    soma_ipva,
                    total_geral
                )
                self._adicionar_log(f"  -> Concluído {placa} | Licenciamento: {val_detran} | IPVA: {val_sefaz}")

            # Salva resultados parciais ou totais
            qtd_concluidos = len(licenciamentos)
            if qtd_concluidos > 0:
                self.root.after(0, lambda: self.var_status.set("Gravando e formatando planilha final..."))
                df_salvar = df.iloc[:qtd_concluidos].copy()
                df_salvar["Licenciamento_CE"] = licenciamentos
                df_salvar["IPVA_CE"] = ipvas
                df_salvar["Data_Consulta"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

                df_salvar.to_excel(arquivo_saida, index=False, engine="openpyxl")
                formatar_planilha_saida(arquivo_saida)

                self._adicionar_log(f"Planilha gravada com sucesso em: {arquivo_saida}")
                self.root.after(0, self._concluir_sucesso, qtd_concluidos, total)
            else:
                self.root.after(0, lambda: self.var_status.set("Consulta cancelada antes do primeiro veículo."))

        except Exception as e:
            self._adicionar_log(f"ERRO CRÍTICO: {str(e)}")
            self.root.after(0, self._concluir_erro, str(e))

        finally:
            self.is_executando = False
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
            self.driver_ativo = None
            self.root.after(0, self._restaurar_controles)

    # ==================================================================================
    # ATUALIZAÇÕES THREAD-SAFE DA INTERFACE (root.after)
    # ==================================================================================

    def _marcar_linha_processando(self, idx: int, placa: str, renavam: str):
        """Destaca visualmente na Treeview o veículo atualmente sob consulta."""
        item_id = str(idx)
        if self.tree.exists(item_id):
            self.tree.item(item_id, values=(idx + 1, placa, renavam, "Consultando...", "Consultando...", "⏳ Processando"), tags=("processando",))
            self.tree.see(item_id)

    def _atualizar_linha_tabela(self, idx: int, lic: str, ipva: str, pos: int, total: int, soma_l: float, soma_i: float, total_g: float):
        """Atualiza a linha do veículo com os valores finais e define cores (verde/vermelho)."""
        item_id = str(idx)
        if self.tree.exists(item_id):
            tem_debito = (lic != "R$ 0,00" and lic != "-") or (ipva != "R$ 0,00" and ipva != "-")
            tag = "com_debito" if tem_debito else "sem_debito"
            situacao = "⚠️ Com Débitos" if tem_debito else "✔️ Sem Débitos"

            valores_anteriores = self.tree.item(item_id, "values")
            self.tree.item(
                item_id,
                values=(valores_anteriores[0], valores_anteriores[1], valores_anteriores[2], lic, ipva, situacao),
                tags=(tag,)
            )

        # Atualiza Mini-Cards de Estatísticas
        self.var_metricas_veiculos.set(f"{pos} / {total}")
        self.var_metricas_lic.set(f"R$ {soma_l:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        self.var_metricas_ipva.set(f"R$ {soma_i:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        self.var_metricas_total.set(f"R$ {total_g:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    def _atualizar_progresso_ui(self, pos: int, total: int, placa: str, pct: int):
        """Atualiza rótulo de status e a barra de progresso."""
        self.var_status.set(f"Consultando [{pos}/{total}]: Placa {placa}...")
        self.barra_progresso.configure(value=pct)

    def _restaurar_controles(self):
        """Restaura botões ao término ou interrupção."""
        self.btn_selecionar.config(state="normal")
        self.btn_iniciar.config(state="normal")
        self.btn_cancelar.config(state="disabled")

        if os.path.exists(self.arquivo_saida_gerado):
            self.btn_abrir_excel.config(state="normal")
            self.btn_abrir_pasta.config(state="normal")

    def _concluir_sucesso(self, concluidos: int, total: int):
        """Exibe feedback final e atualiza a barra."""
        self.barra_progresso.configure(value=100)
        msg_status = "Processamento concluído com sucesso!" if concluidos == total else f"Concluído parcialmente ({concluidos}/{total} veículos)."
        self.var_status.set(msg_status)

        messagebox.showinfo(
            "Consulta Concluída",
            f"{msg_status}\n\nO arquivo 'resultados_frota_ce.xlsx' está pronto para visualização!"
        )

    def _concluir_erro(self, erro: str):
        """Exibe alerta de erro."""
        self.var_status.set("Erro durante o processamento.")
        messagebox.showerror("Erro", f"Ocorreu uma falha durante as consultas:\n\n{erro}")

    def _abrir_planilha_gerada(self):
        """Abre o arquivo Excel gerado diretamente no programa padrão do sistema."""
        if os.path.exists(self.arquivo_saida_gerado):
            try:
                os.startfile(self.arquivo_saida_gerado)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{e}")
        else:
            messagebox.showwarning("Aviso", "O arquivo de resultados ainda não foi gerado.")

    def _abrir_pasta_destino(self):
        """Abre a pasta de destino no Windows Explorer com o arquivo selecionado."""
        if os.path.exists(self.arquivo_saida_gerado):
            try:
                subprocess.Popen(f'explorer /select,"{os.path.normpath(self.arquivo_saida_gerado)}"')
            except Exception:
                pasta = os.path.dirname(self.arquivo_saida_gerado)
                os.startfile(pasta)
        else:
            messagebox.showwarning("Aviso", "A pasta de destino ainda não possui resultados gerados.")

    def _ao_fechar_janela(self):
        """Garante encerramento seguro ao fechar a janela."""
        if self.is_executando:
            if not messagebox.askyesno(
                "Consulta em Andamento",
                "O processamento ainda está em execução.\nDeseja realmente abortar e fechar a aplicação?"
            ):
                return
            self.cancelar_solicitado = True
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
