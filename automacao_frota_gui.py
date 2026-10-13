"""
========================================================================================
Aplicativo GUI Profissional: Consulta de Frota - Detran/Sefaz CE (Versão 3.0)
========================================================================================
Sistema corporativo de automação e web scraping com Interface Gráfica nativa (Tkinter),
resiliência dinâmica de dados (Pandas), extração de Licenciamento, IPVA e Multas,
geração de relatórios executivos em Excel e PDF Paisagem, persistência de preferências,
filtro/ordenação em tempo real, suporte a workers paralelos e Central de Download de
Boletos (Licenciamento e IPVA).
========================================================================================
"""

import os
import sys
import time
import threading
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Importações do Pacote Core Modularizado
from core.config_manager import ConfigManager
from core.logger_config import obter_logger
from core.data_processor import (
    FrotaDataProcessor,
    formatar_moeda_float,
    formatar_float_moeda,
    obter_estilo_financeiro,
    obter_ultimo_dia_util_mes,
    obter_proximo_dia_util,
    PADRAO_PLACA
)
from core.scraper_engine import FrotaScraperEngine
from core.boleto_downloader import FrotaBoletoDownloader
from core.relatorio_exporter import FrotaRelatorioExporter, REPORTLAB_DISPONIVEL

logger = obter_logger("GUI")


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
# JANELA MODAL: CENTRAL DE DOWNLOAD DE BOLETOS (LICENCIAMENTO E IPVA)
# ======================================================================================

class JanelaDownloadBoletos(tk.Toplevel):
    """
    Janela dedicada para seleção e download em lote dos boletos de Licenciamento e IPVA.
    Permite marcar individualmente ou em massa via 'Selecionar Todos'.
    """

    def __init__(
        self,
        parent: tk.Tk,
        df_frota: pd.DataFrame,
        col_placa: str,
        col_renavam: str,
        resultados_memoria: dict[int, dict],
        pasta_destino_base: str,
        modo_noturno: bool,
        modo_headless: bool,
        on_concluido_callback=None
    ):
        super().__init__(parent)
        self.title("Central de Download de Boletos - Licenciamento e IPVA")
        self.geometry("820x620")
        self.minsize(750, 520)
        self.transient(parent)
        self.grab_set()

        self.df_frota = df_frota
        self.col_placa = col_placa
        self.col_renavam = col_renavam
        self.resultados_memoria = resultados_memoria
        self.pasta_destino_base = pasta_destino_base
        self.modo_noturno = modo_noturno
        self.modo_headless = modo_headless
        self.on_concluido_callback = on_concluido_callback

        self.downloader: FrotaBoletoDownloader | None = None
        self.is_baixando = False
        self.cancelar_solicitado = False
        self.itens_boletos: list[dict] = []

        self._aplicar_estilo()
        self._identificar_debitos()
        self._criar_layout()

        self.protocol("WM_DELETE_WINDOW", self._ao_fechar)

    def _aplicar_estilo(self):
        self.bg_root = "#191B22" if self.modo_noturno else "#F0F4F8"
        self.bg_card = "#21242D" if self.modo_noturno else "#FFFFFF"
        self.fg_text = "#CBD5E1" if self.modo_noturno else "#2D3748"
        self.fg_title = "#E2E8F0" if self.modo_noturno else "#1A365D"
        self.fg_muted = "#8E99A8" if self.modo_noturno else "#64748B"
        self.configure(bg=self.bg_root)

    def _identificar_debitos(self):
        """Varre a memória de resultados para encontrar os veículos elegíveis a emissão de boleto."""
        self.itens_boletos.clear()
        total_veiculos = len(self.df_frota)

        for i in range(total_veiculos):
            dados = self.resultados_memoria.get(i, {})
            lic_str = dados.get("lic", "-")
            ipva_str = dados.get("ipva", "-")

            val_lic = formatar_moeda_float(lic_str)
            val_ipva = formatar_moeda_float(ipva_str)

            # Apenas se houver débito real maior que zero
            tem_lic = val_lic > 0 and "falha" not in str(lic_str).lower()
            tem_ipva = val_ipva > 0 and "falha" not in str(ipva_str).lower()

            if tem_lic or tem_ipva:
                row = self.df_frota.loc[i]
                placa = str(row[self.col_placa]).strip().upper()
                renavam = str(row[self.col_renavam]).strip()

                self.itens_boletos.append({
                    "index": i,
                    "placa": placa,
                    "renavam": renavam,
                    "val_lic_str": lic_str if tem_lic else "-",
                    "val_ipva_str": ipva_str if tem_ipva else "-",
                    "tem_lic": tem_lic,
                    "tem_ipva": tem_ipva,
                    "var_lic": tk.BooleanVar(value=tem_lic),
                    "var_ipva": tk.BooleanVar(value=tem_ipva),
                    "lbl_status": None
                })

    def _criar_layout(self):
        container = ttk.Frame(self, padding="14 14 14 14")
        container.pack(fill=tk.BOTH, expand=True)

        # 1. Cabeçalho
        frame_header = ttk.Frame(container)
        frame_header.pack(fill=tk.X, pady=(0, 10))

        lbl_tit = ttk.Label(
            frame_header,
            text="📥 Emissão e Download de Boletos em Lote",
            font=("Segoe UI", 13, "bold"),
            foreground=self.fg_title
        )
        lbl_tit.pack(anchor="w")

        data_hoje = datetime.now().strftime("%d-%m-%Y")
        pasta_prevista = os.path.join(self.pasta_destino_base, f"Licenciamento-IPVA ({data_hoje})")
        vencimento_ipva = obter_ultimo_dia_util_mes()

        lbl_sub = ttk.Label(
            frame_header,
            text=(
                f"Destino: {pasta_prevista}\n"
                f"Padrão dos arquivos: [PLACA] LICENCIAMENTO {data_hoje}.pdf  |  [PLACA] IPVA {data_hoje}.pdf\n"
                f"📅 Vencimento do IPVA: {vencimento_ipva} (Último dia útil do mês)"
            ),
            font=("Segoe UI", 8),
            foreground=self.fg_muted
        )
        lbl_sub.pack(anchor="w", pady=(2, 0))

        # 2. Barra de Seleção Rápida
        frame_controles = ttk.LabelFrame(container, text=" Seleção Rápida ", style="Card.TLabelframe", padding="8 6 8 6")
        frame_controles.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(
            frame_controles,
            text="☑ Selecionar Todos Licenciamento",
            style="Muted.TButton",
            command=self._marcar_todos_licenciamento
        ).pack(side=tk.LEFT, padx=(0, 6))

        ttk.Button(
            frame_controles,
            text="☑ Selecionar Todos IPVA",
            style="Muted.TButton",
            command=self._marcar_todos_ipva
        ).pack(side=tk.LEFT, padx=(0, 6))

        ttk.Button(
            frame_controles,
            text="☑ Marcar Todos",
            style="Muted.TButton",
            command=self._marcar_tudo
        ).pack(side=tk.LEFT, padx=(0, 6))

        ttk.Button(
            frame_controles,
            text="☐ Desmarcar Todos",
            style="Muted.TButton",
            command=self._desmarcar_tudo
        ).pack(side=tk.LEFT)

        lbl_info = ttk.Label(
            frame_controles,
            text=f"Veículos com débitos: {len(self.itens_boletos)}",
            font=("Segoe UI", 9, "bold"),
            foreground=self.fg_title
        )
        lbl_info.pack(side=tk.RIGHT)

        # 3. Lista Rolável de Veículos
        frame_lista = ttk.LabelFrame(container, text=" Veículos com Débitos Pendentes ", style="Card.TLabelframe", padding="6 6 6 6")
        frame_lista.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        canvas = tk.Canvas(frame_lista, bg=self.bg_card, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame_lista, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style="Card.TFrame")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Cabeçalho da tabela interna
        f_hd = ttk.Frame(scrollable_frame, style="Card.TFrame")
        f_hd.pack(fill=tk.X, pady=(0, 4), padx=4)
        ttk.Label(f_hd, text="Placa / RENAVAM", font=("Segoe UI", 9, "bold"), width=22, style="Card.TLabel").pack(side=tk.LEFT)
        ttk.Label(f_hd, text="Boleto de Licenciamento", font=("Segoe UI", 9, "bold"), width=30, style="Card.TLabel").pack(side=tk.LEFT)
        ttk.Label(f_hd, text="DAE de IPVA", font=("Segoe UI", 9, "bold"), width=30, style="Card.TLabel").pack(side=tk.LEFT)
        ttk.Label(f_hd, text="Status", font=("Segoe UI", 9, "bold"), width=16, style="Card.TLabel").pack(side=tk.LEFT)

        if not self.itens_boletos:
            ttk.Label(
                scrollable_frame,
                text="Nenhum veículo com débitos pendentes de Licenciamento ou IPVA.",
                font=("Segoe UI", 9, "italic"),
                foreground=self.fg_muted,
                style="Card.TLabel"
            ).pack(pady=20)
        else:
            for item in self.itens_boletos:
                f_row = ttk.Frame(scrollable_frame, style="Card.TFrame")
                f_row.pack(fill=tk.X, pady=2, padx=4)

                # Coluna 1: Placa e RENAVAM
                texto_id = f"🚘 {item['placa']} ({item['renavam']})"
                ttk.Label(f_row, text=texto_id, width=22, font=("Segoe UI", 9, "bold"), style="Card.TLabel").pack(side=tk.LEFT)

                # Coluna 2: Licenciamento Checkbox
                if item["tem_lic"]:
                    txt_chk_lic = f"Licenciamento: {item['val_lic_str']}"
                    chk_lic = ttk.Checkbutton(
                        f_row,
                        text=txt_chk_lic,
                        variable=item["var_lic"],
                        style="Card.TCheckbutton",
                        width=28
                    )
                    chk_lic.pack(side=tk.LEFT)
                else:
                    ttk.Label(f_row, text="Regular (R$ 0,00)", foreground=self.fg_muted, width=28, style="Card.TLabel").pack(side=tk.LEFT)

                # Coluna 3: IPVA Checkbox
                if item["tem_ipva"]:
                    txt_chk_ipva = f"IPVA: {item['val_ipva_str']}"
                    chk_ipva = ttk.Checkbutton(
                        f_row,
                        text=txt_chk_ipva,
                        variable=item["var_ipva"],
                        style="Card.TCheckbutton",
                        width=28
                    )
                    chk_ipva.pack(side=tk.LEFT)
                else:
                    ttk.Label(f_row, text="Regular (R$ 0,00)", foreground=self.fg_muted, width=28, style="Card.TLabel").pack(side=tk.LEFT)

                # Coluna 4: Status do Download
                lbl_st = ttk.Label(f_row, text="Pendente", font=("Segoe UI", 8), foreground=self.fg_muted, width=16, style="Card.TLabel")
                lbl_st.pack(side=tk.LEFT)
                item["lbl_status"] = lbl_st

        # 4. Progresso e Ações Finais
        frame_rodape = ttk.Frame(container)
        frame_rodape.pack(fill=tk.X)

        self.var_progresso_txt = tk.StringVar(value="Selecione os boletos desejados e clique em Iniciar Download.")
        self.lbl_progresso = ttk.Label(
            frame_rodape,
            textvariable=self.var_progresso_txt,
            font=("Segoe UI", 9, "bold"),
            foreground=self.fg_title
        )
        self.lbl_progresso.pack(anchor="w", pady=(0, 4))

        self.barra_progresso = ttk.Progressbar(frame_rodape, mode="determinate")
        self.barra_progresso.pack(fill=tk.X, pady=(0, 8))

        frame_botoes = ttk.Frame(frame_rodape)
        frame_botoes.pack(fill=tk.X)

        self.btn_iniciar_dl = ttk.Button(
            frame_botoes,
            text="🚀 Iniciar Download dos Boletos",
            style="Action.TButton",
            state="normal" if self.itens_boletos else "disabled",
            command=self._iniciar_downloads
        )
        self.btn_iniciar_dl.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_cancelar_dl = ttk.Button(
            frame_botoes,
            text="⏹️ Cancelar",
            style="Danger.TButton",
            state="disabled",
            command=self._cancelar_downloads
        )
        self.btn_cancelar_dl.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_abrir_pasta_boletos = ttk.Button(
            frame_botoes,
            text="📂 Abrir Pasta dos Boletos",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_pasta_boletos
        )
        self.btn_abrir_pasta_boletos.pack(side=tk.LEFT)

        ttk.Button(
            frame_botoes,
            text="Fechar",
            style="Secondary.TButton",
            command=self._ao_fechar
        ).pack(side=tk.RIGHT)

    def _marcar_todos_licenciamento(self):
        for item in self.itens_boletos:
            if item["tem_lic"]:
                item["var_lic"].set(True)

    def _marcar_todos_ipva(self):
        for item in self.itens_boletos:
            if item["tem_ipva"]:
                item["var_ipva"].set(True)

    def _marcar_tudo(self):
        self._marcar_todos_licenciamento()
        self._marcar_todos_ipva()

    def _desmarcar_tudo(self):
        for item in self.itens_boletos:
            item["var_lic"].set(False)
            item["var_ipva"].set(False)

    def _iniciar_downloads(self):
        """Inicia a rotina em thread separada para download sequencial seguro com Chrome."""
        tarefas = []
        for item in self.itens_boletos:
            if item["tem_lic"] and item["var_lic"].get():
                tarefas.append((item, "lic"))
            if item["tem_ipva"] and item["var_ipva"].get():
                tarefas.append((item, "ipva"))

        if not tarefas:
            messagebox.showwarning("Nenhum Boleto Selecionado", "Marque ao menos um boleto (Licenciamento ou IPVA) para baixar.")
            return

        self.is_baixando = True
        self.cancelar_solicitado = False
        self.btn_iniciar_dl.config(state="disabled")
        self.btn_cancelar_dl.config(state="normal")
        self.btn_abrir_pasta_boletos.config(state="disabled")
        self.barra_progresso["value"] = 0

        thread = threading.Thread(target=self._executar_downloads_thread, args=(tarefas,), daemon=True)
        thread.start()

    def _executar_downloads_thread(self, tarefas: list[tuple[dict, str]]):
        total = len(tarefas)
        self.downloader = FrotaBoletoDownloader(self.pasta_destino_base, headless=self.modo_headless)

        try:
            self.downloader.iniciar()
            sucessos = 0

            for idx, (item, tipo) in enumerate(tarefas):
                if self.cancelar_solicitado:
                    break

                placa = item["placa"]
                renavam = item["renavam"]
                lbl_tipo = "Licenciamento" if tipo == "lic" else "IPVA"
                pos = idx + 1
                pct = int((idx / total) * 100)

                self.after(0, lambda p=pos, t=total, plc=placa, tp=lbl_tipo, pc=pct: self._atualizar_status_ui(p, t, plc, tp, pc))

                # Atualiza label na linha
                if item["lbl_status"]:
                    self.after(0, lambda lbl=item["lbl_status"], tp=lbl_tipo: lbl.config(
                        text=f"Baixando {tp}...", foreground="#2563EB"
                    ))

                if tipo == "lic":
                    sucesso, msg = self.downloader.baixar_boleto_licenciamento(placa, renavam)
                else:
                    sucesso, msg = self.downloader.baixar_dae_ipva(placa, renavam)

                if sucesso:
                    sucessos += 1
                    cor = "#15803D"
                    txt = f"{lbl_tipo}: Concluído"
                else:
                    cor = "#DC2626"
                    txt = f"{lbl_tipo}: Erro"

                if item["lbl_status"]:
                    self.after(0, lambda lbl=item["lbl_status"], t=txt, c=cor: lbl.config(text=t, foreground=c))

            # Finalização
            self.after(0, lambda s=sucessos, t=total: self._finalizar_downloads_ui(s, t))

        except Exception as e:
            logger.error(f"Erro fatal no lote de downloads: {e}")
            self.after(0, lambda: self.var_progresso_txt.set("Falha no processo de download."))
        finally:
            if self.downloader:
                self.downloader.fechar()
            self.is_baixando = False

    def _atualizar_status_ui(self, pos: int, total: int, placa: str, tipo: str, pct: int):
        self.var_progresso_txt.set(f"Baixando boleto ({pos}/{total}): {placa} - {tipo}...")
        self.barra_progresso["value"] = pct

    def _finalizar_downloads_ui(self, sucessos: int, total: int):
        self.barra_progresso["value"] = 100
        self.btn_iniciar_dl.config(state="normal")
        self.btn_cancelar_dl.config(state="disabled")
        self.btn_abrir_pasta_boletos.config(state="normal")

        if self.cancelar_solicitado:
            self.var_progresso_txt.set(f"Downloads interrompidos! {sucessos} de {total} boletos baixados.")
            messagebox.showinfo("Cancelado", f"Processamento interrompido. Foram salvos {sucessos} boletos.")
        else:
            self.var_progresso_txt.set(f"Download concluído com sucesso! {sucessos} de {total} boletos salvos.")
            if sys.platform == "win32":
                try:
                    import winsound
                    winsound.MessageBeep(winsound.MB_ICONASTERISK)
                except Exception:
                    pass
            messagebox.showinfo(
                "Boletos Baixados",
                f"Foram baixados {sucessos} de {total} boletos selecionados com sucesso!\n\n"
                f"Pasta de destino:\n{self.downloader.pasta_destino_final if self.downloader else self.pasta_destino_base}"
            )

    def _cancelar_downloads(self):
        if messagebox.askyesno("Cancelar", "Deseja interromper o download dos boletos restantes?"):
            self.cancelar_solicitado = True
            self.btn_cancelar_dl.config(state="disabled")

    def _abrir_pasta_boletos(self):
        caminho = self.downloader.pasta_destino_final if self.downloader else self.pasta_destino_base
        if os.path.exists(caminho):
            os.startfile(caminho)
        else:
            os.startfile(self.pasta_destino_base)

    def _ao_fechar(self):
        if self.is_baixando:
            if not messagebox.askyesno("Fechar", "O download de boletos está em andamento. Deseja cancelar e fechar?"):
                return
            self.cancelar_solicitado = True
        self.destroy()


# ======================================================================================
# INTERFACE PRINCIPAL PROFISSIONAL
# ======================================================================================

class AppConsultaFrotaGUI:
    """Interface Gráfica Corporativa para Automação de Frotas CE."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Consulta de Frota - Detran/Sefaz CE (v3.0)")
        self.root.geometry("980x800")
        self.root.minsize(900, 720)

        # Carrega preferências persistentes
        self.cfg = ConfigManager.carregar()

        # Variáveis de Estado
        self.caminho_arquivo = tk.StringVar(value=self.cfg.get("ultimo_arquivo", ""))
        self.pasta_destino_var = tk.StringVar(value=self.cfg.get("pasta_destino", ""))
        self.arquivo_excel_gerado = ""
        self.arquivo_pdf_gerado = ""

        # Filtros de Débito
        self.var_ipva = tk.BooleanVar(value=self.cfg.get("consultar_ipva", True))
        self.var_lic = tk.BooleanVar(value=self.cfg.get("consultar_lic", True))
        self.var_multas = tk.BooleanVar(value=self.cfg.get("consultar_multas", True))

        # Formato de Exportação e Opções Gerais
        padrao_exp = self.cfg.get("formato_exportacao", "ambos")
        if not REPORTLAB_DISPONIVEL and padrao_exp in ["pdf", "ambos"]:
            padrao_exp = "excel"
        self.var_formato_exportacao = tk.StringVar(value=padrao_exp)
        self.var_headless = tk.BooleanVar(value=self.cfg.get("modo_headless", True))
        self.var_workers = tk.IntVar(value=self.cfg.get("qtd_workers", 1))

        self.var_status = tk.StringVar(value="Pronto para iniciar. Selecione a planilha de frota.")
        self.var_busca = tk.StringVar(value="")

        # Métricas do Dashboard
        self.var_metricas_veiculos = tk.StringVar(value="0 / 0")
        self.var_metricas_lic = tk.StringVar(value="R$ 0,00")
        self.var_metricas_ipva = tk.StringVar(value="R$ 0,00")
        self.var_metricas_multas = tk.StringVar(value="R$ 0,00")
        self.var_metricas_total = tk.StringVar(value="R$ 0,00")

        # Controle de Execução e Memória
        self.is_executando = False
        self.is_pausado = False
        self.cancelar_solicitado = False
        self.df_carregado: pd.DataFrame | None = None
        self.col_placa = "Placa"
        self.col_renavam = "RENAVAM"
        self.resultados_memoria: dict[int, dict] = {}
        self.indices_em_falha: list[int] = []

        # Ordenação da Tabela
        self.coluna_ordenacao = ""
        self.ordem_reversa = False

        # Modo Noturno
        self.modo_noturno = self.cfg.get("modo_noturno", False)
        self.cards_metricas = []

        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self._criar_layout()
        self._aplicar_tema()

        # Se houver último arquivo válido, tenta carregar automaticamente
        if self.caminho_arquivo.get() and os.path.exists(self.caminho_arquivo.get()):
            self.root.after(200, lambda: self._carregar_planilha_automatica(self.caminho_arquivo.get()))

        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar_janela)

    def _salvar_preferencias(self):
        """Grava o estado atual no config.json."""
        self.cfg["modo_noturno"] = self.modo_noturno
        self.cfg["pasta_destino"] = self.pasta_destino_var.get()
        self.cfg["consultar_ipva"] = self.var_ipva.get()
        self.cfg["consultar_lic"] = self.var_lic.get()
        self.cfg["consultar_multas"] = self.var_multas.get()
        self.cfg["modo_headless"] = self.var_headless.get()
        self.cfg["formato_exportacao"] = self.var_formato_exportacao.get()
        self.cfg["qtd_workers"] = self.var_workers.get()
        ConfigManager.salvar(self.cfg)

    def _aplicar_tema(self):
        """Aplica tokens visuais confortáveis para Modo Claro e Modo Noturno."""
        if self.modo_noturno:
            bg_root = "#191B22"
            bg_card = "#21242D"
            border_card = "#2F333F"
            fg_title = "#E2E8F0"
            fg_sub = "#8E99A8"
            fg_text = "#CBD5E1"
            bg_tree = "#21242D"
            fg_tree = "#CBD5E1"
            bg_tree_hd = "#2A2E3B"
            fg_tree_hd = "#E2E8F0"
            bg_sel = "#353D52"
            fg_sel = "#F8FAFC"
            bg_tab = "#1B1D24"
            fg_tab = "#8E99A8"
            bg_tab_sel = "#21242D"
            fg_tab_sel = "#9BB3D6"
            bg_log = "#181A20"
            fg_log = "#A5B4C7"
            status_fg = "#9BB3D6"
            bg_btn_sec = "#2D323E"
            fg_btn_sec = "#CBD5E1"
            bg_btn_sec_act = "#393F4E"
            bg_btn_muted = "#2A2E3A"
            fg_btn_muted = "#CBD5E1"
            bg_btn_theme = "#2A2E3A"
            fg_btn_theme = "#CBD5E1"
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
        self.style.configure(".", background=bg_root, foreground=fg_text, font=("Segoe UI", 9))
        self.style.configure("TFrame", background=bg_root)
        self.style.configure("Card.TFrame", background=bg_card)
        self.style.configure("TLabel", background=bg_root, foreground=fg_text)
        self.style.configure("Card.TLabel", background=bg_card, foreground=fg_text)
        self.style.configure("Header.TLabel", background=bg_root, font=("Segoe UI", 14, "bold"), foreground=fg_title)
        self.style.configure("SubHeader.TLabel", background=bg_root, font=("Segoe UI", 9), foreground=fg_sub)

        self.style.configure("Card.TLabelframe", background=bg_card, relief=tk.SOLID, borderwidth=1)
        self.style.configure("Card.TLabelframe.Label", font=("Segoe UI", 9, "bold"), foreground=fg_title, background=bg_root)
        self.style.configure("Card.TCheckbutton", background=bg_card, foreground=fg_text)
        self.style.map("Card.TCheckbutton", background=[("active", bg_card)], foreground=[("active", fg_title)])
        self.style.configure("Card.TRadiobutton", background=bg_card, foreground=fg_text)
        self.style.map("Card.TRadiobutton", background=[("active", bg_card)], foreground=[("active", fg_title)])

        self.style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"), background=bg_btn_pri, foreground="#FFFFFF", padding=(10, 6))
        self.style.map("Primary.TButton", background=[("active", bg_btn_pri_act), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        self.style.configure("Action.TButton", font=("Segoe UI", 9, "bold"), background=bg_btn_act, foreground="#FFFFFF", padding=(12, 6))
        self.style.map("Action.TButton", background=[("active", bg_btn_act_act), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        self.style.configure("Danger.TButton", font=("Segoe UI", 9, "bold"), background=bg_btn_dng, foreground="#FFFFFF", padding=(9, 6))
        self.style.map("Danger.TButton", background=[("active", bg_btn_dng_act), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        self.style.configure("Secondary.TButton", font=("Segoe UI", 9, "bold"), background=bg_btn_sec, foreground=fg_btn_sec, padding=(9, 5))
        self.style.map("Secondary.TButton", background=[("active", bg_btn_sec_act), ("disabled", "#4B5563" if self.modo_noturno else "#64748B")])

        self.style.configure("Muted.TButton", font=("Segoe UI", 8), background=bg_btn_muted, foreground=fg_btn_muted, padding=(6, 3))
        self.style.map("Muted.TButton", background=[("active", bg_btn_sec_act if self.modo_noturno else "#CBD5E1")])

        self.style.configure("Theme.TButton", font=("Segoe UI", 9, "bold"), background=bg_btn_theme, foreground=fg_btn_theme, padding=(8, 4))
        self.style.map("Theme.TButton", background=[("active", bg_btn_sec_act if self.modo_noturno else "#CBD5E1")])

        self.style.configure("TNotebook", background=bg_card, borderwidth=0)
        self.style.configure("TNotebook.Tab", background=bg_tab, foreground=fg_tab, padding=(12, 6), font=("Segoe UI", 9, "bold"))
        self.style.map("TNotebook.Tab", background=[("selected", bg_tab_sel)], foreground=[("selected", fg_tab_sel)])

        self.style.configure("Treeview", background=bg_tree, foreground=fg_tree, fieldbackground=bg_tree, rowheight=24, font=("Segoe UI", 9))
        self.style.configure("Treeview.Heading", background=bg_tree_hd, foreground=fg_tree_hd, font=("Segoe UI", 9, "bold"), padding=(4, 5))
        self.style.map("Treeview", background=[("selected", bg_sel)], foreground=[("selected", fg_sel)])

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
                self.tree.tag_configure("sem_debito", background="#F0FFF4", foreground="#276749")
                self.tree.tag_configure("falha", background="#FEF3C7", foreground="#B45309", font=("Segoe UI", 9, "bold"))

        if hasattr(self, "txt_logs"):
            self.txt_logs.configure(bg=bg_log, fg=fg_log, insertbackground=fg_log)

        if hasattr(self, "lbl_status"):
            self.lbl_status.configure(foreground=status_fg)

        for card_frame, lbl_titulo, lbl_valor, cor_l, cor_d in self.cards_metricas:
            card_frame.configure(bg=bg_card, highlightbackground=border_card)
            lbl_titulo.configure(bg=bg_card, fg=fg_sub)
            lbl_valor.configure(bg=bg_card, fg=cor_d if self.modo_noturno else cor_l)

        if hasattr(self, "btn_tema"):
            self.btn_tema.configure(text="☀️ Modo Claro" if self.modo_noturno else "🌙 Modo Noturno")

    def _alternar_tema(self):
        self.modo_noturno = not self.modo_noturno
        self._salvar_preferencias()
        self._aplicar_tema()

    def _criar_layout(self):
        container = ttk.Frame(self.root, padding="14 14 14 14")
        container.pack(fill=tk.BOTH, expand=True)

        # 1. BARRA SUPERIOR
        frame_topo = ttk.Frame(container)
        frame_topo.pack(fill=tk.X, pady=(0, 8))

        frame_titulos = ttk.Frame(frame_topo)
        frame_titulos.pack(side=tk.LEFT, fill=tk.X, expand=True)

        lbl_titulo = ttk.Label(
            frame_titulos,
            text="🚗 Consulta de Débitos de Frota - Detran/Sefaz CE",
            style="Header.TLabel"
        )
        lbl_titulo.pack(anchor="w")

        lbl_subtitulo = ttk.Label(
            frame_titulos,
            text="Versão 3.0 Corporativa | Automação com auto-heal, download de boletos e múltiplos workers",
            style="SubHeader.TLabel"
        )
        lbl_subtitulo.pack(anchor="w", pady=(1, 0))

        self.btn_tema = ttk.Button(
            frame_topo,
            text="🌙 Modo Noturno",
            style="Theme.TButton",
            command=self._alternar_tema
        )
        self.btn_tema.pack(side=tk.RIGHT, padx=(8, 0))

        # 2. CARD: ENTRADA E CONFIGURAÇÕES
        frame_configs = ttk.LabelFrame(container, text=" Configurações e Arquivo ", style="Card.TLabelframe", padding="8 8 8 8")
        frame_configs.pack(fill=tk.X, pady=(0, 8))

        # Linha 1: Seleção de Planilha e Destino
        linha_arquivo = ttk.Frame(frame_configs, style="Card.TFrame")
        linha_arquivo.pack(fill=tk.X, pady=(0, 6))

        self.btn_selecionar = ttk.Button(
            linha_arquivo,
            text="📁 Selecionar Planilha",
            style="Primary.TButton",
            command=self._selecionar_arquivo
        )
        self.btn_selecionar.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_pasta_destino = ttk.Button(
            linha_arquivo,
            text="📂 Pasta de Destino",
            style="Secondary.TButton",
            command=self._selecionar_pasta_destino
        )
        self.btn_pasta_destino.pack(side=tk.LEFT, padx=(0, 10))

        self.lbl_caminho = ttk.Label(
            linha_arquivo,
            text="Nenhuma planilha selecionada",
            font=("Segoe UI", 9, "italic"),
            foreground="#64748B",
            style="Card.TLabel"
        )
        self.lbl_caminho.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Linha 2: Filtros de Débito e Marcar Todos
        linha_filtros = ttk.Frame(frame_configs, style="Card.TFrame")
        linha_filtros.pack(fill=tk.X, pady=(0, 6))

        lbl_filtro_txt = ttk.Label(linha_filtros, text="Consultar:", font=("Segoe UI", 9, "bold"), style="Card.TLabel")
        lbl_filtro_txt.pack(side=tk.LEFT, padx=(0, 8))

        self.chk_ipva = ttk.Checkbutton(linha_filtros, text="IPVA (Sefaz-CE)", variable=self.var_ipva, style="Card.TCheckbutton")
        self.chk_ipva.pack(side=tk.LEFT, padx=(0, 10))

        self.chk_lic = ttk.Checkbutton(linha_lic := linha_filtros, text="Licenciamento (Detran-CE)", variable=self.var_lic, style="Card.TCheckbutton")
        self.chk_lic.pack(side=tk.LEFT, padx=(0, 10))

        self.chk_multas = ttk.Checkbutton(linha_filtros, text="Multas (Detran-CE)", variable=self.var_multas, style="Card.TCheckbutton")
        self.chk_multas.pack(side=tk.LEFT, padx=(0, 12))

        self.btn_toggle_filtros = ttk.Button(
            linha_filtros,
            text="☑️ Marcar/Desmarcar",
            style="Muted.TButton",
            command=self._alternar_todos_filtros
        )
        self.btn_toggle_filtros.pack(side=tk.LEFT, padx=(0, 14))

        # Seletor de Concorrência (Workers)
        lbl_workers = ttk.Label(linha_filtros, text="Velocidade:", font=("Segoe UI", 9, "bold"), style="Card.TLabel")
        lbl_workers.pack(side=tk.LEFT, padx=(0, 4))

        self.combo_workers = ttk.Combobox(
            linha_filtros,
            values=["1x (Sequencial)", "2x Paralelo", "3x Paralelo", "4x Máximo"],
            state="readonly",
            width=14
        )
        idx_w = max(0, min(self.var_workers.get() - 1, 3))
        self.combo_workers.current(idx_w)
        self.combo_workers.bind("<<ComboboxSelected>>", self._ao_mudar_workers)
        self.combo_workers.pack(side=tk.LEFT)

        # Linha 3: Opções de Exportação, Headless e Botões Principais
        linha_opcoes = ttk.Frame(frame_configs, style="Card.TFrame")
        linha_opcoes.pack(fill=tk.X, pady=(0, 2))

        lbl_exp = ttk.Label(linha_opcoes, text="Exportação:", font=("Segoe UI", 9, "bold"), style="Card.TLabel")
        lbl_exp.pack(side=tk.LEFT, padx=(0, 8))

        rb_excel = ttk.Radiobutton(linha_opcoes, text="Excel (.xlsx)", value="excel", variable=self.var_formato_exportacao, style="Card.TRadiobutton")
        rb_excel.pack(side=tk.LEFT, padx=(0, 8))

        rb_pdf = ttk.Radiobutton(
            linha_opcoes,
            text="PDF (.pdf)",
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
        rb_ambos.pack(side=tk.LEFT, padx=(0, 12))

        self.chk_headless = ttk.Checkbutton(
            linha_opcoes,
            text="Chrome Oculto",
            variable=self.var_headless,
            style="Card.TCheckbutton"
        )
        self.chk_headless.pack(side=tk.LEFT)

        # Botões Iniciar, Pausar e Cancelar à direita
        self.btn_cancelar = ttk.Button(
            linha_opcoes,
            text="⏹️ Cancelar",
            style="Danger.TButton",
            state="disabled",
            command=self._cancelar_processamento
        )
        self.btn_cancelar.pack(side=tk.RIGHT, padx=(6, 0))

        self.btn_pausar = ttk.Button(
            linha_opcoes,
            text="⏸️ Pausar",
            style="Secondary.TButton",
            state="disabled",
            command=self._alternar_pausa
        )
        self.btn_pausar.pack(side=tk.RIGHT, padx=(6, 0))

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

        # Aba 1: Tabela Visual com Busca e Ordenação
        aba_tabela = ttk.Frame(self.notebook, padding="4 4 4 4")
        self.notebook.add(aba_tabela, text=" 📊 Tabela de Veículos ")

        # Barra de Busca na Tabela
        frame_busca = ttk.Frame(aba_tabela)
        frame_busca.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frame_busca, text="🔍 Buscar na Tabela:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 6))
        self.entry_busca = ttk.Entry(frame_busca, textvariable=self.var_busca, width=28)
        self.entry_busca.pack(side=tk.LEFT, padx=(0, 6))
        self.var_busca.trace_add("write", lambda *_: self._filtrar_tabela())

        ttk.Button(frame_busca, text="✖ Limpar Busca", style="Muted.TButton", command=lambda: self.var_busca.set("")).pack(side=tk.LEFT)

        colunas = ("num", "placa", "renavam", "licenciamento", "ipva", "multas", "situacao")
        self.tree = ttk.Treeview(aba_tabela, columns=colunas, show="headings", selectmode="extended")

        cabecalhos_nomes = [
            ("num", "Item", 45, "center"),
            ("placa", "Placa", 85, "center"),
            ("renavam", "RENAVAM", 115, "center"),
            ("licenciamento", "Licenciamento", 130, "e"),
            ("ipva", "IPVA", 130, "e"),
            ("multas", "Multas", 130, "e"),
            ("situacao", "Situação", 145, "center"),
        ]

        for col_id, titulo, larg, alinhamento in cabecalhos_nomes:
            self.tree.heading(col_id, text=f"{titulo} ↕", command=lambda c=col_id: self._ordenar_coluna(c))
            self.tree.column(col_id, width=larg, anchor=alinhamento)

        scroll_y = ttk.Scrollbar(aba_tabela, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(aba_tabela, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Menu de Contexto
        self.menu_contexto = tk.Menu(self.tree, tearoff=0)
        self.menu_contexto.add_command(label="🔄 Reconsultar Veículo Selecionado", command=self._reconsultar_selecionado)
        self.tree.bind("<Button-3>", self._exibir_menu_contexto)

        # Aba 2: Console de Logs
        aba_logs = ttk.Frame(self.notebook, padding="4 4 4 4")
        self.notebook.add(aba_logs, text=" 📜 Logs do Sistema ")

        self.txt_logs = tk.Text(aba_logs, wrap=tk.WORD, font=("Consolas", 8), relief=tk.FLAT)
        scroll_logs = ttk.Scrollbar(aba_logs, orient="vertical", command=self.txt_logs.yview)
        self.txt_logs.configure(yscrollcommand=scroll_logs.set)
        self.txt_logs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_logs.pack(side=tk.RIGHT, fill=tk.Y)

        # 4. CARD: DASHBOARD DE MÉTRICAS CONSOLIDADAS
        frame_metricas = ttk.Frame(container)
        frame_metricas.pack(fill=tk.X, pady=(0, 8))

        def criar_card(pai, titulo, var_val, cor_l, cor_d):
            card = tk.Frame(pai, bg="#FFFFFF", relief=tk.SOLID, bd=1, highlightthickness=1)
            card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=3)
            lbl_t = tk.Label(card, text=titulo, font=("Segoe UI", 8, "bold"), fg="#64748B", bg="#FFFFFF")
            lbl_t.pack(pady=(4, 1))
            lbl_v = tk.Label(card, textvariable=var_val, font=("Segoe UI", 11, "bold"), fg=cor_l, bg="#FFFFFF")
            lbl_v.pack(pady=(0, 4))
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
            text="📊 Abrir Excel",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_excel_gerado
        )
        self.btn_abrir_excel.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_abrir_pdf = ttk.Button(
            frame_botoes_finais,
            text="📄 Abrir PDF",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_pdf_gerado
        )
        self.btn_abrir_pdf.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_abrir_pasta = ttk.Button(
            frame_botoes_finais,
            text="📂 Abrir Pasta",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_pasta_destino
        )
        self.btn_abrir_pasta.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_reconsultar_falhas = ttk.Button(
            frame_botoes_finais,
            text="⚠️ Reconsultar Falhas",
            style="Secondary.TButton",
            state="disabled",
            command=self._reconsultar_apenas_falhas
        )
        self.btn_reconsultar_falhas.pack(side=tk.LEFT, padx=(0, 10))

        # BOTÃO EXCLUSIVO DE DOWNLOAD DE BOLETOS (Habilitado apenas pós-consulta com débitos)
        self.btn_baixar_boletos = ttk.Button(
            frame_botoes_finais,
            text="📥 Baixar Boletos (Licenciamento e IPVA)",
            style="Action.TButton",
            state="disabled",
            command=self._abrir_central_download_boletos
        )
        self.btn_baixar_boletos.pack(side=tk.RIGHT)

    def _ao_mudar_workers(self, _=None):
        val_str = self.combo_workers.get()
        if "1x" in val_str: self.var_workers.set(1)
        elif "2x" in val_str: self.var_workers.set(2)
        elif "3x" in val_str: self.var_workers.set(3)
        elif "4x" in val_str: self.var_workers.set(4)
        self._salvar_preferencias()

    def _alternar_pausa(self):
        if not self.is_executando:
            return
        self.is_pausado = not self.is_pausado
        if self.is_pausado:
            self.btn_pausar.config(text="▶️ Continuar")
            self.var_status.set("Consulta pausada temporariamente pelo usuário.")
            self._adicionar_log("Processamento pausado.")
        else:
            self.btn_pausar.config(text="⏸️ Pausar")
            self.var_status.set("Retomando consultas de onde parou...")
            self._adicionar_log("Processamento retomado.")

    def _adicionar_log(self, mensagem: str):
        logger.info(mensagem)
        hora = datetime.now().strftime("%H:%M:%S")
        linha = f"[{hora}] {mensagem}\n"
        self.txt_logs.insert(tk.END, linha)
        self.txt_logs.see(tk.END)

    def _exibir_menu_contexto(self, event):
        item_selecionado = self.tree.identify_row(event.y)
        if item_selecionado:
            if item_selecionado not in self.tree.selection():
                self.tree.selection_set(item_selecionado)
            self.menu_contexto.post(event.x_root, event.y_root)

    def _alternar_todos_filtros(self):
        todos_ativos = self.var_ipva.get() and self.var_lic.get() and self.var_multas.get()
        novo_estado = not todos_ativos
        self.var_ipva.set(novo_estado)
        self.var_lic.set(novo_estado)
        self.var_multas.set(novo_estado)

    def _selecionar_pasta_destino(self):
        pasta = filedialog.askdirectory(title="Selecione a Pasta de Destino dos Relatórios e Boletos")
        if pasta:
            self.pasta_destino_var.set(pasta)
            self._salvar_preferencias()
            self._adicionar_log(f"Pasta de destino definida: {pasta}")

    def _selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a Planilha de Frota",
            filetypes=[("Planilhas Excel", "*.xlsx *.xls"), ("Todos os Arquivos", "*.*")]
        )
        if caminho:
            self._carregar_planilha_automatica(caminho)

    def _carregar_planilha_automatica(self, caminho: str):
        try:
            self.caminho_arquivo.set(caminho)
            df, col_p, col_r = FrotaDataProcessor.carregar_planilha(caminho)
            self.df_carregado = df
            self.col_placa = col_p
            self.col_renavam = col_r

            ConfigManager.adicionar_arquivo_recente(caminho)

            if not self.pasta_destino_var.get():
                pasta_padrao = os.path.dirname(os.path.abspath(caminho))
                self.pasta_destino_var.set(pasta_padrao)

            nome = os.path.basename(caminho)
            self.lbl_caminho.config(
                text=f"Carregado: {nome} ({len(df)} veículos) | Colunas: '{col_p}', '{col_r}'",
                foreground="#90BE9F" if self.modo_noturno else "#15803D",
                font=("Segoe UI", 9, "bold")
            )

            self.btn_iniciar.config(state="normal")
            self.var_status.set(f"Arquivo carregado com sucesso: {len(df)} veículos válidos prontos.")
            self._preencher_tabela_inicial()
            self._adicionar_log(f"Planilha '{nome}' carregada com {len(df)} registros válidos.")

        except Exception as e:
            messagebox.showerror("Erro ao Carregar", f"Não foi possível processar a planilha:\n\n{e}")
            self.lbl_caminho.config(text="Erro ao carregar o arquivo", foreground="#DC2626")
            self.btn_iniciar.config(state="disabled")

    def _preencher_tabela_inicial(self):
        self.tree.delete(*self.tree.get_children())
        self.resultados_memoria.clear()
        self.indices_em_falha.clear()

        df = self.df_carregado
        total = len(df)
        self.var_metricas_veiculos.set(f"0 / {total}")
        self.var_metricas_lic.set("R$ 0,00")
        self.var_metricas_ipva.set("R$ 0,00")
        self.var_metricas_multas.set("R$ 0,00")
        self.var_metricas_total.set("R$ 0,00")

        for idx, row in df.iterrows():
            placa = str(row[self.col_placa]).strip().upper()
            renavam = str(row[self.col_renavam]).strip()

            tag = "pendente"
            if not PADRAO_PLACA.match(placa):
                tag = "falha"
                self.indices_em_falha.append(idx)

            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(idx + 1, placa, renavam, "-", "-", "-", "Aguardando"),
                tags=(tag,)
            )

    def _filtrar_tabela(self):
        """Filtra as linhas visíveis na Treeview em tempo real conforme a busca."""
        termo = self.var_busca.get().strip().upper()
        if self.df_carregado is None:
            return

        for idx in range(len(self.df_carregado)):
            row = self.df_carregado.loc[idx]
            placa = str(row[self.col_placa]).strip().upper()
            renavam = str(row[self.col_renavam]).strip()
            dados_mem = self.resultados_memoria.get(idx, {})

            situacao = "Aguardando"
            if idx in self.resultados_memoria:
                lic = dados_mem.get("lic", "-")
                ipva = dados_mem.get("ipva", "-")
                multas = dados_mem.get("multas", "-")
                soma = formatar_moeda_float(lic) + formatar_moeda_float(ipva) + formatar_moeda_float(multas)
                if soma > 0: situacao = "Com Débitos"
                else: situacao = "Sem Débitos"

            corresponde = (not termo) or (termo in placa) or (termo in renavam) or (termo in situacao.upper())

            if corresponde:
                if not self.tree.exists(str(idx)):
                    # Reinsere o item
                    lic_v = dados_mem.get("lic", "-")
                    ipva_v = dados_mem.get("ipva", "-")
                    multas_v = dados_mem.get("multas", "-")
                    self.tree.insert(
                        "",
                        "end",
                        iid=str(idx),
                        values=(idx + 1, placa, renavam, lic_v, ipva_v, multas_v, situacao),
                        tags=("sem_debito" if situacao == "Sem Débitos" else "com_debito" if situacao == "Com Débitos" else "pendente",)
                    )
            else:
                if self.tree.exists(str(idx)):
                    self.tree.delete(str(idx))

    def _ordenar_coluna(self, col: str):
        """Ordena a Treeview dinamicamente pelo cabeçalho clicado (crescente/decrescente)."""
        itens = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]

        # Detecta se é coluna de moeda ou numérica
        if col in ["licenciamento", "ipva", "multas", "num"]:
            itens.sort(key=lambda t: formatar_moeda_float(t[0]), reverse=self.ordem_reversa)
        else:
            itens.sort(key=lambda t: t[0].upper(), reverse=self.ordem_reversa)

        for index, (_, k) in enumerate(itens):
            self.tree.move(k, "", index)

        self.ordem_reversa = not self.ordem_reversa
        self.coluna_ordenacao = col

    def _iniciar_processamento(self):
        self._disparar_raspagem(indices_alvo=None)

    def _reconsultar_selecionado(self):
        selecionados = self.tree.selection()
        if not selecionados:
            messagebox.showwarning("Seleção Necessária", "Selecione ao menos um veículo na tabela para reconsultar.")
            return
        indices = [int(i) for i in selecionados]
        self._disparar_raspagem(indices_alvo=indices)

    def _reconsultar_apenas_falhas(self):
        indices_falhas = []
        for item_id in self.tree.get_children():
            tags = self.tree.item(item_id, "tags")
            if "falha" in tags:
                indices_falhas.append(int(item_id))

        if not indices_falhas:
            messagebox.showinfo("Nenhuma Falha", "Não há veículos com status de falha na tabela.")
            return

        self._disparar_raspagem(indices_alvo=indices_falhas)

    def _disparar_raspagem(self, indices_alvo: list[int] | None = None):
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
            messagebox.showwarning("Filtro Necessário", "Selecione ao menos um débito para consulta.")
            return

        self.is_executando = True
        self.is_pausado = False
        self.cancelar_solicitado = False

        self.btn_iniciar.config(state="disabled")
        self.btn_selecionar.config(state="disabled")
        self.btn_pasta_destino.config(state="disabled")
        self.btn_reconsultar_falhas.config(state="disabled")
        self.btn_baixar_boletos.config(state="disabled")
        self.btn_pausar.config(state="normal", text="⏸️ Pausar")
        self.btn_cancelar.config(state="normal")
        self.btn_abrir_excel.config(state="disabled")
        self.btn_abrir_pdf.config(state="disabled")
        self.btn_abrir_pasta.config(state="disabled")
        self.barra_progresso["value"] = 0

        self.notebook.select(0)

        thread = threading.Thread(
            target=self._executar_raspagem_thread,
            args=(caminho, self.var_headless.get(), c_ipva, c_lic, c_multas, self.var_formato_exportacao.get(), indices_alvo, self.var_workers.get()),
            daemon=True
        )
        thread.start()

    def _cancelar_processamento(self):
        if messagebox.askyesno("Confirmar Cancelamento", "Deseja realmente interromper as consultas ao concluir o veículo atual?"):
            self.cancelar_solicitado = True
            self.var_status.set("Cancelamento solicitado! Finalizando o veículo atual e salvando parciais...")
            self.btn_cancelar.config(state="disabled")
            self.btn_pausar.config(state="disabled")
            self._adicionar_log("Solicitação de cancelamento registrada.")

    def _salvar_checkpoint(self, pasta_destino: str, df: pd.DataFrame):
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
        indices_alvo: list[int] | None = None,
        num_workers: int = 1
    ):
        pasta_destino = self.pasta_destino_var.get().strip() or os.path.dirname(os.path.abspath(caminho_arquivo))
        os.makedirs(pasta_destino, exist_ok=True)

        arquivo_excel = os.path.join(pasta_destino, "resultados_frota_ce.xlsx")
        arquivo_pdf = os.path.join(pasta_destino, "resultados_frota_ce.pdf")

        scrapers_ativos = []
        try:
            df = self.df_carregado.copy()
            total_geral = len(df)

            lista_indices = indices_alvo if indices_alvo is not None else list(df.index)
            total_lote = len(lista_indices)

            self.root.after(0, lambda: self.var_status.set("Inicializando navegadores Chrome..."))
            self._adicionar_log(f"Iniciando lote de {total_lote} veículos ({num_workers} worker(s)).")

            # Fila ou partição simples por worker
            lock_memoria = threading.Lock()
            concluidos_contador = [0]

            def worker_task(worker_id: int, sub_indices: list[int]):
                scraper = FrotaScraperEngine(headless=modo_headless)
                scrapers_ativos.append(scraper)
                try:
                    scraper.iniciar()
                    for idx in sub_indices:
                        if self.cancelar_solicitado:
                            break

                        # Pausa amigável
                        while self.is_pausado and not self.cancelar_solicitado:
                            time.sleep(0.5)

                        row = df.loc[idx]
                        placa = str(row[self.col_placa]).strip().upper()
                        renavam = str(row[self.col_renavam]).strip()

                        self.root.after(0, self._marcar_linha_processando, idx, placa, renavam)

                        val_detran_lic = "-"
                        val_detran_multas = "-"
                        val_sefaz_ipva = "-"

                        if consultar_lic or consultar_multas:
                            val_detran_lic, val_detran_multas = scraper.consultar_detran(
                                placa, renavam, consultar_lic, consultar_multas
                            )
                            time.sleep(0.3)

                        if consultar_ipva:
                            val_sefaz_ipva = scraper.consultar_sefaz_ipva(placa, renavam)
                            time.sleep(0.3)

                        with lock_memoria:
                            self.resultados_memoria[idx] = {
                                "lic": val_detran_lic,
                                "ipva": val_sefaz_ipva,
                                "multas": val_detran_multas
                            }
                            concluidos_contador[0] += 1
                            c_atual = concluidos_contador[0]

                            soma_lic = sum(formatar_moeda_float(d.get("lic", "0")) for d in self.resultados_memoria.values())
                            soma_ipva = sum(formatar_moeda_float(d.get("ipva", "0")) for d in self.resultados_memoria.values())
                            soma_multas = sum(formatar_moeda_float(d.get("multas", "0")) for d in self.resultados_memoria.values())
                            total_acumulado = soma_lic + soma_ipva + soma_multas
                            pct = int((c_atual / total_lote) * 100)

                        self.root.after(0, lambda p=c_atual, t=total_lote, plc=placa, pc=pct: self._atualizar_progresso_ui(p, t, plc, pc))
                        self.root.after(
                            0,
                            self._atualizar_linha_tabela,
                            idx,
                            val_detran_lic,
                            val_sefaz_ipva,
                            val_detran_multas,
                            c_atual,
                            total_geral,
                            soma_lic,
                            soma_ipva,
                            soma_multas,
                            total_acumulado
                        )
                        self._adicionar_log(f"[{c_atual}/{total_lote}] {placa} | Lic: {val_detran_lic} | IPVA: {val_sefaz_ipva} | Multas: {val_detran_multas}")

                        if c_atual % 5 == 0:
                            self._salvar_checkpoint(pasta_destino, df)

                finally:
                    scraper.fechar()

            # Dispara workers em paralelo
            num_workers = max(1, min(num_workers, len(lista_indices)))
            sub_listas = [lista_indices[i::num_workers] for i in range(num_workers)]
            threads_workers = []

            for w_id, sub_l in enumerate(sub_listas):
                if sub_l:
                    t = threading.Thread(target=worker_task, args=(w_id, sub_l), daemon=True)
                    threads_workers.append(t)
                    t.start()

            for t in threads_workers:
                t.join()

            # Consolidação dos Dados
            self.root.after(0, lambda: self.var_status.set("Consolidando dados e gerando relatórios..."))

            lics_final = [self.resultados_memoria.get(i, {}).get("lic", "-") for i in range(total_geral)]
            ipvas_final = [self.resultados_memoria.get(i, {}).get("ipva", "-") for i in range(total_geral)]
            multas_final = [self.resultados_memoria.get(i, {}).get("multas", "-") for i in range(total_geral)]

            df_final = FrotaDataProcessor.consolidar_resultados(
                df, lics_final, ipvas_final, multas_final, total_geral
            )

            soma_lic_f = sum(formatar_moeda_float(v) for v in lics_final)
            soma_ipva_f = sum(formatar_moeda_float(v) for v in ipvas_final)
            soma_multas_f = sum(formatar_moeda_float(v) for v in multas_final)
            totais_dict = {
                "lic": formatar_float_moeda(soma_lic_f),
                "ipva": formatar_float_moeda(soma_ipva_f),
                "multas": formatar_float_moeda(soma_multas_f),
                "total": formatar_float_moeda(soma_lic_f + soma_ipva_f + soma_multas_f)
            }

            caminho_excel_gerado = ""
            caminho_pdf_gerado = ""

            filtros_ativos = []
            if consultar_ipva: filtros_ativos.append("IPVA")
            if consultar_lic: filtros_ativos.append("Licenciamento")
            if consultar_multas: filtros_ativos.append("Multas")

            if formato_export in ["excel", "ambos"]:
                caminho_excel_gerado = FrotaRelatorioExporter.exportar_excel(df_final, arquivo_excel)
                self.arquivo_excel_gerado = caminho_excel_gerado

            if formato_export in ["pdf", "ambos"] and REPORTLAB_DISPONIVEL:
                caminho_pdf_gerado = FrotaRelatorioExporter.exportar_pdf(
                    df_final, arquivo_pdf, totais_dict, filtros_ativos
                )
                self.arquivo_pdf_gerado = caminho_pdf_gerado

            # Remove checkpoint temporário se tudo correu normalmente
            chk_path = os.path.join(pasta_destino, "~checkpoint_frota_ce.xlsx")
            if os.path.exists(chk_path):
                try: os.remove(chk_path)
                except Exception: pass

            self.root.after(
                0,
                self._finalizar_processamento_ui,
                caminho_excel_gerado,
                caminho_pdf_gerado,
                self.cancelar_solicitado
            )

        except Exception as e:
            logger.error(f"Erro durante o processamento da frota: {e}")
            self.root.after(0, lambda: messagebox.showerror("Erro de Processamento", f"Ocorreu uma falha no processamento:\n\n{e}"))
            self.root.after(0, lambda: self.var_status.set("Erro durante o processamento. Verifique a aba de logs."))
        finally:
            for sc in scrapers_ativos:
                sc.fechar()
            self.is_executando = False
            self.is_pausado = False

    def _marcar_linha_processando(self, index: int, placa: str, renavam: str):
        if self.tree.exists(str(index)):
            self.tree.item(str(index), values=(index + 1, placa, renavam, "...", "...", "...", "Processando..."), tags=("processando",))
            self.tree.see(str(index))

    def _atualizar_progresso_ui(self, atual: int, total: int, placa: str, pct: int):
        self.var_status.set(f"Consultando ({atual}/{total}): {placa}...")
        self.barra_progresso["value"] = pct

    def _atualizar_linha_tabela(
        self,
        index: int,
        lic: str,
        ipva: str,
        multas: str,
        qtd_concluidos: int,
        total_geral: int,
        soma_lic: float,
        soma_ipva: float,
        soma_multas: float,
        total_acumulado: float
    ):
        soma_veiculo = formatar_moeda_float(lic) + formatar_moeda_float(ipva) + formatar_moeda_float(multas)
        tem_falha = any("falha" in str(v).lower() or "erro" in str(v).lower() for v in [lic, ipva, multas])

        if tem_falha:
            situacao = "Falha na Consulta"
            tag = "falha"
            if index not in self.indices_em_falha:
                self.indices_em_falha.append(index)
        elif soma_veiculo > 0:
            situacao = "Com Débitos"
            tag = "com_debito"
            if index in self.indices_em_falha:
                self.indices_em_falha.remove(index)
        else:
            situacao = "Sem Débitos"
            tag = "sem_debito"
            if index in self.indices_em_falha:
                self.indices_em_falha.remove(index)

        row = self.df_carregado.loc[index]
        placa = str(row[self.col_placa]).strip().upper()
        renavam = str(row[self.col_renavam]).strip()

        if self.tree.exists(str(index)):
            self.tree.item(
                str(index),
                values=(index + 1, placa, renavam, lic, ipva, multas, situacao),
                tags=(tag,)
            )

        self.var_metricas_veiculos.set(f"{qtd_concluidos} / {total_geral}")
        self.var_metricas_lic.set(formatar_float_moeda(soma_lic))
        self.var_metricas_ipva.set(formatar_float_moeda(soma_ipva))
        self.var_metricas_multas.set(formatar_float_moeda(soma_multas))
        self.var_metricas_total.set(formatar_float_moeda(total_acumulado))

    def _finalizar_processamento_ui(self, arq_excel: str, arq_pdf: str, foi_cancelado: bool):
        self.barra_progresso["value"] = 100
        self.btn_iniciar.config(state="normal")
        self.btn_selecionar.config(state="normal")
        self.btn_pasta_destino.config(state="normal")
        self.btn_cancelar.config(state="disabled")
        self.btn_pausar.config(state="disabled", text="⏸️ Pausar")
        self.btn_abrir_pasta.config(state="normal")

        if arq_excel and os.path.exists(arq_excel):
            self.btn_abrir_excel.config(state="normal")
        if arq_pdf and os.path.exists(arq_pdf):
            self.btn_abrir_pdf.config(state="normal")

        if self.indices_em_falha:
            self.btn_reconsultar_falhas.config(state="normal")

        # Verifica se há débitos para habilitar o botão de Download de Boletos
        total_com_debitos = sum(
            1 for d in self.resultados_memoria.values()
            if formatar_moeda_float(d.get("lic", "0")) > 0 or formatar_moeda_float(d.get("ipva", "0")) > 0
        )
        if total_com_debitos > 0:
            self.btn_baixar_boletos.config(state="normal")

        # Notificação Sonora no Windows
        if sys.platform == "win32":
            try:
                import winsound
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
            except Exception:
                pass

        if foi_cancelado:
            self.var_status.set("Processamento interrompido. Resultados parciais salvos.")
            messagebox.showinfo("Interrompido", "O processamento foi interrompido e os veículos concluídos foram gravados com sucesso.")
        else:
            self.var_status.set("Processamento concluído com sucesso!")
            msg = f"Consultas concluídas com sucesso!\n\nTotal a Pagar na Frota: {self.var_metricas_total.get()}"
            if total_com_debitos > 0:
                msg += f"\n\n💡 {total_com_debitos} veículo(s) possuem débitos de Licenciamento ou IPVA.\nVocê pode usar o botão 'Baixar Boletos' abaixo para emitir os PDFs."
            messagebox.showinfo("Sucesso", msg)

    def _abrir_central_download_boletos(self):
        """Abre a Janela Modal de Download de Boletos para os veículos com débitos de Licenciamento e IPVA."""
        if not self.df_carregado is not None or not self.resultados_memoria:
            messagebox.showwarning("Aviso", "Realize uma consulta de frota antes de baixar boletos.")
            return

        pasta_base = self.pasta_destino_var.get().strip() or os.path.dirname(os.path.abspath(self.caminho_arquivo.get()))

        JanelaDownloadBoletos(
            parent=self.root,
            df_frota=self.df_carregado,
            col_placa=self.col_placa,
            col_renavam=self.col_renavam,
            resultados_memoria=self.resultados_memoria,
            pasta_destino_base=pasta_base,
            modo_noturno=self.modo_noturno,
            modo_headless=self.var_headless.get()
        )

    def _abrir_excel_gerado(self):
        if self.arquivo_excel_gerado and os.path.exists(self.arquivo_excel_gerado):
            os.startfile(self.arquivo_excel_gerado)

    def _abrir_pdf_gerado(self):
        if self.arquivo_pdf_gerado and os.path.exists(self.arquivo_pdf_gerado):
            os.startfile(self.arquivo_pdf_gerado)

    def _abrir_pasta_destino(self):
        pasta = self.pasta_destino_var.get().strip()
        if not pasta and self.caminho_arquivo.get():
            pasta = os.path.dirname(os.path.abspath(self.caminho_arquivo.get()))
        if pasta and os.path.exists(pasta):
            os.startfile(pasta)

    def _ao_fechar_janela(self):
        if self.is_executando:
            if not messagebox.askyesno("Encerrar", "Existe uma consulta em andamento. Deseja realmente interromper e fechar?"):
                return
            self.cancelar_solicitado = True
        self._salvar_preferencias()
        self.root.destroy()


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
            root_err = tk.Tk()
            root_err.withdraw()
            messagebox.showerror("Erro de Inicialização", err_msg)
            root_err.destroy()
        except Exception:
            print(err_msg, file=sys.stderr)


if __name__ == "__main__":
    main()
