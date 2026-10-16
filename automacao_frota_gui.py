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
import re
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
    calcular_total_a_pagar_veiculo,
    obter_estilo_financeiro,
    obter_ultimo_dia_util_mes,
    obter_proximo_dia_util,
    PADRAO_PLACA,
    gerar_nome_base_resultado,
    obter_caminhos_arquivos_saida
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


def configurar_dpi():
    """
    Habilita suporte a High-DPI no Windows para evitar distorções,
    desfoque e incompatibilidades de resolução entre diferentes monitores.
    """
    if sys.platform == "win32":
        try:
            import ctypes
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass



# ======================================================================================
# UTILITÁRIOS GRÁFICOS: DESENHO DE GEOMETRIAS ARREDONDADAS SUAVES (BEZIER)
# ======================================================================================

def desenhar_retangulo_arredondado(canvas: tk.Canvas, x1: float, y1: float, x2: float, y2: float, radius: int = 10, **kwargs):
    """
    Gera um polígono com bordas curvas suaves via algoritmo Bezier nativo do Tkinter (smooth=True).
    Proporciona o aspecto autêntico de cartões modernos de alta fidelidade (SaaS / Fintech).
    """
    r = radius
    pontos = [
        x1 + r, y1,
        x1 + r, y1,
        x2 - r, y1,
        x2 - r, y1,
        x2, y1,
        x2, y1 + r,
        x2, y1 + r,
        x2, y2 - r,
        x2, y2 - r,
        x2, y2,
        x2 - r, y2,
        x2 - r, y2,
        x1 + r, y2,
        x1 + r, y2,
        x1, y2,
        x1, y2 - r,
        x1, y2 - r,
        x1, y1 + r,
        x1, y1 + r,
        x1, y1
    ]
    return canvas.create_polygon(pontos, **kwargs, smooth=True)


# ======================================================================================
# COMPONENTE VISUAL: LUXURY METRIC CARD (DASHBOARD FINTECH)
# ======================================================================================

class LuxuryMetricCard(tk.Canvas):
    """
    Card de Métrica Executiva renderizado em Canvas de alta precisão.
    Conta com cantos arredondados, barra de acento funcional, tipografia hierárquica,
    ícone de nicho, legenda descritiva e micro-interação de brilho (Hover Glow) ao passar o cursor.
    """

    def __init__(
        self,
        parent: tk.Widget,
        titulo: str,
        var_valor: tk.StringVar,
        icone: str = "📊",
        cor_acento: str = "#3B82F6",
        subtexto: str = "✦ Tempo real",
        modo_noturno: bool = False
    ):
        super().__init__(parent, highlightthickness=0, height=76, width=100)
        self.titulo = titulo
        self.var_valor = var_valor
        self.icone = icone
        self.cor_acento = cor_acento
        self.subtexto = subtexto
        self.modo_noturno = modo_noturno

        self.is_hover = False
        self._definir_paleta()

        self.bind("<Configure>", self._ao_redimensionar)
        self.bind("<Enter>", self._ao_entrar)
        self.bind("<Leave>", self._ao_sair)

        # Atualização automática em tempo real sempre que a variável mudar
        self.var_valor.trace_add("write", lambda *_: self.redesenhar())

    def _definir_paleta(self):
        if self.modo_noturno:
            self.bg_parent = "#080C14"
            self.bg_card = "#0F172A"
            self.border_normal = "#1E293B"
            self.border_glow = self.cor_acento
            self.fg_title = "#94A3B8"
            self.fg_sub = "#64748B"
            self.fg_val_default = "#F8FAFC"
        else:
            self.bg_parent = "#F1F5F9"
            self.bg_card = "#FFFFFF"
            self.border_normal = "#E2E8F0"
            self.border_glow = self.cor_acento
            self.fg_title = "#475569"
            self.fg_sub = "#94A3B8"
            self.fg_val_default = "#0F172A"

        self.configure(bg=self.bg_parent)

    def atualizar_tema(self, modo_noturno: bool):
        self.modo_noturno = modo_noturno
        self._definir_paleta()
        self.redesenhar()

    def _ao_entrar(self, _):
        self.is_hover = True
        self.redesenhar()

    def _ao_sair(self, _):
        self.is_hover = False
        self.redesenhar()

    def _ao_redimensionar(self, event):
        self.redesenhar(event.width, event.height)

    def redesenhar(self, w: int = None, h: int = None):
        w = w or self.winfo_width()
        h = h or self.winfo_height()
        if w < 30 or h < 30:
            return

        self.delete("all")
        x1, y1, x2, y2 = 4, 3, w - 5, h - 4
        cor_borda = self.border_glow if self.is_hover else self.border_normal
        largura_borda = 2 if self.is_hover else 1

        # 1. Fundo com cantos arredondados
        desenhar_retangulo_arredondado(
            self, x1, y1, x2, y2, radius=10,
            fill=self.bg_card, outline=cor_borda, width=largura_borda
        )

        # 2. Barra de Acento Superior Refinada
        self.create_rectangle(
            x1 + 16, y1 + 1, x2 - 16, y1 + 3,
            fill=self.cor_acento, outline=""
        )

        # 3. Ícone e Título Superior
        self.create_text(
            x1 + 12, y1 + 16,
            text=f"{self.icone}  {self.titulo}",
            anchor="w",
            fill=self.fg_title,
            font=("Segoe UI", 8, "bold")
        )

        # 4. Valor Formatado (tamanho adaptativo)
        val = self.var_valor.get() or "R$ 0,00"
        cor_val = self.cor_acento if ("TOTAL" in self.titulo or "R$" in val) else self.fg_val_default
        tam_fonte = 12 if w >= 155 else 10
        self.create_text(
            x1 + 12, y1 + 40,
            text=val,
            anchor="w",
            fill=cor_val,
            font=("Segoe UI", tam_fonte, "bold")
        )

        # 5. Micro-Legenda Inferior
        self.create_text(
            x1 + 12, y1 + 60,
            text=self.subtexto,
            anchor="w",
            fill=self.fg_sub,
            font=("Segoe UI", 7)
        )


# ======================================================================================
# COMPONENTE VISUAL: LIVE STATUS ORB (INDICADOR DE ESTADO PULSANTE)
# ======================================================================================

class LiveStatusOrb(tk.Canvas):
    """
    Badge indicador de conectividade e status com visual de software de missão crítica.
    Exibe um orbe colorido luminoso acompanhado de texto descritivo de telemetria.
    """

    def __init__(self, parent: tk.Widget, modo_noturno: bool = False):
        super().__init__(parent, width=220, height=32, highlightthickness=0)
        self.modo_noturno = modo_noturno
        self.status_tipo = "pronto"
        self.texto_status = "SISTEMA ONLINE & PRONTO"
        self._atualizar_cores()

    def _atualizar_cores(self):
        self.bg_canvas = "#080C14" if self.modo_noturno else "#F1F5F9"
        self.bg_pill = "#111827" if self.modo_noturno else "#FFFFFF"
        self.border_pill = "#1F2937" if self.modo_noturno else "#E2E8F0"
        self.fg_text = "#F8FAFC" if self.modo_noturno else "#0F172A"
        self.configure(bg=self.bg_canvas)
        self.redesenhar()

    def atualizar_tema(self, modo_noturno: bool):
        self.modo_noturno = modo_noturno
        self._atualizar_cores()

    def definir_status(self, tipo: str, texto: str):
        self.status_tipo = tipo
        self.texto_status = texto
        self.redesenhar()

    def redesenhar(self):
        self.delete("all")
        w, h = 216, 28
        desenhar_retangulo_arredondado(
            self, 2, 2, w, h, radius=8,
            fill=self.bg_pill, outline=self.border_pill, width=1
        )

        cores_orbe = {
            "pronto": ("#10B981", "#34D399"),      # Esmeralda
            "processando": ("#3B82F6", "#60A5FA"), # Azul Elétrico
            "pausado": ("#F59E0B", "#FCD34D"),     # Âmbar
            "erro": ("#EF4444", "#FCA5A5")         # Carmim
        }
        cor_base, cor_brilho = cores_orbe.get(self.status_tipo, ("#10B981", "#34D399"))

        # Efeito de Halo/Glow no Orbe
        self.create_oval(10, 8, 22, 20, fill=cor_brilho, outline="")
        self.create_oval(12, 10, 20, 18, fill=cor_base, outline="")

        # Texto do Status
        self.create_text(
            28, 14,
            text=self.texto_status,
            anchor="w",
            fill=self.fg_text,
            font=("Segoe UI", 8, "bold")
        )


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
        on_concluido_callback=None,
        caminho_arquivo: str = ""
    ):
        super().__init__(parent)
        self.title("Central de Download de Boletos - Licenciamento e IPVA")
        self.transient(parent)
        self.grab_set()

        self._ajustar_geometria_responsiva()

        self.df_frota = df_frota
        self.col_placa = col_placa
        self.col_renavam = col_renavam
        self.resultados_memoria = resultados_memoria
        self.pasta_destino_base = pasta_destino_base
        self.modo_noturno = modo_noturno
        self.modo_headless = modo_headless
        self.on_concluido_callback = on_concluido_callback
        self.caminho_arquivo = caminho_arquivo

        self.nome_frota = ""
        if caminho_arquivo:
            base = FrotaDataProcessor.gerar_nome_base_resultado(caminho_arquivo)
            self.nome_frota = re.sub(r"^resultados[_\-\s]*", "", base, flags=re.IGNORECASE).strip() or base

        self.downloader: FrotaBoletoDownloader | None = None
        self.is_baixando = False
        self.is_pausado = False
        self.cancelar_solicitado = False
        self.itens_boletos: list[dict] = []
        self.tarefas_executadas_sucesso: set[tuple[int, str]] = set()

        self._aplicar_estilo()
        self._identificar_debitos()
        self._criar_layout()

        self.protocol("WM_DELETE_WINDOW", self._ao_fechar)

    def _ajustar_geometria_responsiva(self):
        """
        Ajusta e centraliza a Janela Modal de Download de Boletos de acordo com a
        resolução do monitor e a janela principal, garantindo que botões e lista fiquem visíveis.
        """
        self.update_idletasks()
        tela_w = self.winfo_screenwidth()
        tela_h = self.winfo_screenheight()

        margem_h = 40
        margem_v = 70

        min_w = 680
        min_h = 420

        largura = min(840, max(min_w, tela_w - margem_h))
        altura = min(560, max(min_h, tela_h - margem_v))

        if tela_w <= 1366:
            largura = min(largura, max(min_w, int(tela_w * 0.90)))
        if tela_h <= 800:
            altura = min(altura, max(min_h, int((tela_h - margem_v) * 0.90)))

        try:
            p = self.master
            p.update_idletasks()
            pos_x = p.winfo_rootx() + (p.winfo_width() - largura) // 2
            pos_y = p.winfo_rooty() + (p.winfo_height() - altura) // 2
        except Exception:
            pos_x = (tela_w - largura) // 2
            pos_y = (tela_h - margem_v - altura) // 2

        pos_x = max(10, min(pos_x, max(10, tela_w - largura - 10)))
        pos_y = max(10, min(pos_y, max(10, tela_h - margem_v - altura)))

        self.minsize(min_w, min_h)
        self.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    def _aplicar_estilo(self):
        if self.modo_noturno:
            self.bg_root = "#080C14"
            self.bg_card = "#0F172A"
            self.fg_text = "#E2E8F0"
            self.fg_title = "#F8FAFC"
            self.fg_muted = "#94A3B8"
            self.border_card = "#1F2937"
            self.bg_item = "#151F32"
        else:
            self.bg_root = "#F1F5F9"
            self.bg_card = "#FFFFFF"
            self.fg_text = "#1E293B"
            self.fg_title = "#0F172A"
            self.fg_muted = "#64748B"
            self.border_card = "#CBD5E1"
            self.bg_item = "#F8FAFC"

        self.configure(bg=self.bg_root)

        style = ttk.Style(self)
        style.configure("PremiumBoleto.TFrame", background=self.bg_root)
        style.configure("PremiumBoletoCard.TFrame", background=self.bg_card)
        style.configure("PremiumItem.TFrame", background=self.bg_item)
        style.configure("PrimaryDownload.TButton", font=("Segoe UI", 9, "bold"), background="#059669", foreground="#FFFFFF", padding=(14, 7))
        style.map("PrimaryDownload.TButton", background=[("active", "#047857"), ("disabled", "#4B5563")])

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
        container = ttk.Frame(self, style="PremiumBoleto.TFrame", padding="14 12 14 12")
        container.pack(fill=tk.BOTH, expand=True)

        # 1. Header Executivo do Modal (Fixado no topo)
        frame_header = tk.Frame(container, bg=self.bg_card, relief=tk.FLAT, bd=0, highlightthickness=1, highlightbackground=self.border_card)
        frame_header.pack(side=tk.TOP, fill=tk.X, pady=(0, 6), ipady=5, ipadx=10)

        lbl_tit = tk.Label(
            frame_header,
            text="📥 Central Executiva de Emissão de Boletos (PDF)",
            font=("Segoe UI", 11, "bold"),
            bg=self.bg_card,
            fg=self.fg_title
        )
        lbl_tit.pack(anchor="w")

        data_hoje = datetime.now().strftime("%d-%m-%Y")
        nome_pasta_prevista = (
            f"Licenciamento-IPVA ({self.nome_frota} - {data_hoje})"
            if self.nome_frota
            else f"Licenciamento-IPVA ({data_hoje})"
        )
        pasta_prevista = os.path.join(self.pasta_destino_base, nome_pasta_prevista)
        vencimento_ipva = obter_ultimo_dia_util_mes()

        sub_msg = (
            f"Diretório Destino: {pasta_prevista}\n"
            f"📅 Vencimento Legal do IPVA: {vencimento_ipva} (Último dia útil do mês)"
        )
        lbl_sub = tk.Label(frame_header, text=sub_msg, font=("Segoe UI", 8), bg=self.bg_card, fg=self.fg_muted, justify=tk.LEFT)
        lbl_sub.pack(anchor="w", pady=(2, 0))

        # 2. Barra de Seleção Rápida em Pílulas (Fixada no topo)
        frame_acoes = tk.Frame(container, bg=self.bg_card, relief=tk.FLAT, bd=0, highlightthickness=1, highlightbackground=self.border_card)
        frame_acoes.pack(side=tk.TOP, fill=tk.X, pady=(0, 6), ipady=4, ipadx=8)

        ttk.Button(frame_acoes, text="☑️ Marcar Licenciamento", style="Muted.TButton", command=self._marcar_todos_licenciamento).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(frame_acoes, text="☑️ Marcar IPVA", style="Muted.TButton", command=self._marcar_todos_ipva).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(frame_acoes, text="☑️ Marcar Tudo", style="Muted.TButton", command=self._marcar_tudo).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(frame_acoes, text="☐ Desmarcar", style="Muted.TButton", command=self._desmarcar_tudo).pack(side=tk.LEFT, padx=(0, 10))

        self.lbl_contador_modal = tk.Label(
            frame_acoes,
            text=f"Veículos com débitos: {len(self.itens_boletos)}",
            font=("Segoe UI", 9, "bold"),
            bg=self.bg_card,
            fg=self.fg_title
        )
        self.lbl_contador_modal.pack(side=tk.RIGHT)

        # 4. Rodapé e Botões de Comando (FIXADO NO FUNDO - NUNCA ESCONDIDO!)
        frame_rodape = ttk.Frame(container, style="PremiumBoleto.TFrame")
        frame_rodape.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))

        self.var_progresso_txt = tk.StringVar(value="Selecione os boletos desejados e clique em Iniciar Download.")
        self.lbl_progresso = ttk.Label(frame_rodape, textvariable=self.var_progresso_txt, font=("Segoe UI", 8, "bold"))
        self.lbl_progresso.pack(anchor="w", pady=(0, 3))

        self.barra_progresso = ttk.Progressbar(frame_rodape, mode="determinate")
        self.barra_progresso.pack(fill=tk.X, pady=(0, 6))

        frame_botoes = ttk.Frame(frame_rodape, style="PremiumBoleto.TFrame")
        frame_botoes.pack(fill=tk.X)

        self.btn_iniciar_dl = ttk.Button(
            frame_botoes,
            text="⚡ Iniciar Download dos Boletos",
            style="PrimaryDownload.TButton",
            state="normal" if self.itens_boletos else "disabled",
            command=self._iniciar_downloads
        )
        self.btn_iniciar_dl.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_pausar_dl = ttk.Button(frame_botoes, text="⏸️ Pausar", style="Secondary.TButton", state="disabled", command=self._alternar_pausa_downloads)
        self.btn_pausar_dl.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_cancelar_dl = ttk.Button(frame_botoes, text="⏹️ Cancelar", style="Danger.TButton", state="disabled", command=self._cancelar_downloads)
        self.btn_cancelar_dl.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_abrir_pasta_boletos = ttk.Button(frame_botoes, text="📂 Abrir Pasta", style="Secondary.TButton", state="disabled", command=self._abrir_pasta_boletos)
        self.btn_abrir_pasta_boletos.pack(side=tk.LEFT)

        ttk.Button(frame_botoes, text="Fechar", style="Muted.TButton", command=self._ao_fechar).pack(side=tk.RIGHT)

        # 3. Lista de Veículos com Contêiner Rolável (OCUPA O ESPAÇO RESTANTE NO CENTRO)
        frame_scroll = tk.Frame(container, bg=self.bg_card, relief=tk.FLAT, bd=0, highlightthickness=1, highlightbackground=self.border_card)
        frame_scroll.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 6))

        canvas = tk.Canvas(frame_scroll, bg=self.bg_card, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame_scroll, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style="PremiumBoletoCard.TFrame")

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Responsividade do frame interno no Canvas
        def _ao_redimensionar_canvas(event):
            canvas.itemconfig(canvas_window_id, width=event.width)
        canvas.bind("<Configure>", _ao_redimensionar_canvas)

        # Rolagem suave pelo mousewheel do Windows
        def _rolar_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self.bind_all("<MouseWheel>", _rolar_mousewheel)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=4)

        if not self.itens_boletos:
            lbl_vazio = tk.Label(
                scrollable_frame,
                text="🎉 Nenhum veículo com débitos pendentes elegíveis para emissão de boletos.",
                font=("Segoe UI", 9, "italic"),
                bg=self.bg_card,
                fg=self.fg_muted
            )
            lbl_vazio.pack(padx=20, pady=40)
        else:
            for item in self.itens_boletos:
                card_item = tk.Frame(scrollable_frame, bg=self.bg_item, relief=tk.FLAT, bd=0, highlightthickness=1, highlightbackground=self.border_card)
                card_item.pack(fill=tk.X, expand=True, padx=4, pady=3, ipady=4, ipadx=6)

                # Badge Monospace para Placa
                lbl_placa = tk.Label(card_item, text=f"🚗 {item['placa']}", font=("Consolas", 10, "bold"), bg=self.bg_item, fg=self.fg_title)
                lbl_placa.pack(side=tk.LEFT, padx=(6, 12))

                lbl_ren = tk.Label(card_item, text=f"RENAVAM: {item['renavam']}", font=("Segoe UI", 8), bg=self.bg_item, fg=self.fg_muted)
                lbl_ren.pack(side=tk.LEFT, padx=(0, 16))

                # Opção de Licenciamento
                if item["tem_lic"]:
                    chk_lic = ttk.Checkbutton(card_item, text=f"Licenciamento ({item['val_lic_str']})", variable=item["var_lic"], style="Card.TCheckbutton")
                    chk_lic.pack(side=tk.LEFT, padx=(0, 14))
                else:
                    lbl_sem_lic = tk.Label(card_item, text="Licenciamento: Isento/Pago", font=("Segoe UI", 8), bg=self.bg_item, fg="#10B981")
                    lbl_sem_lic.pack(side=tk.LEFT, padx=(0, 14))

                # Opção de IPVA
                if item["tem_ipva"]:
                    chk_ipva = ttk.Checkbutton(card_item, text=f"IPVA ({item['val_ipva_str']})", variable=item["var_ipva"], style="Card.TCheckbutton")
                    chk_ipva.pack(side=tk.LEFT, padx=(0, 14))
                else:
                    lbl_sem_ipva = tk.Label(card_item, text="IPVA: Isento/Pago", font=("Segoe UI", 8), bg=self.bg_item, fg="#10B981")
                    lbl_sem_ipva.pack(side=tk.LEFT, padx=(0, 14))

                # Status Individual
                lbl_st = tk.Label(card_item, text="Pendente", font=("Segoe UI", 8, "bold"), bg=self.bg_item, fg=self.fg_muted)
                lbl_st.pack(side=tk.RIGHT, padx=6)
                item["lbl_status"] = lbl_st

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

    def _alternar_pausa_downloads(self):
        """Pausa ou retoma a execução dos downloads."""
        if not self.is_baixando:
            return
        if not self.is_pausado:
            self.is_pausado = True
            self.btn_pausar_dl.config(text="▶️ Retomar")
            self.var_progresso_txt.set("⏸️ Downloads pausados. Clique em Retomar para prosseguir...")
            logger.info("Downloads pausados pelo usuário.")
        else:
            self.is_pausado = False
            self.btn_pausar_dl.config(text="⏸️ Pausar")
            self.var_progresso_txt.set("▶️ Retomando downloads...")
            logger.info("Downloads retomados pelo usuário.")

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
        self.is_pausado = False
        self.cancelar_solicitado = False
        self.tarefas_executadas_sucesso.clear()

        self.btn_iniciar_dl.config(state="disabled")
        self.btn_pausar_dl.config(state="normal", text="⏸️ Pausar")
        self.btn_cancelar_dl.config(state="normal")
        self.btn_abrir_pasta_boletos.config(state="disabled")
        self.barra_progresso["value"] = 0

        thread = threading.Thread(target=self._executar_downloads_thread, args=(tarefas,), daemon=True)
        thread.start()

    def _executar_downloads_thread(self, tarefas: list[tuple[dict, str]]):
        total_geral = len(tarefas)
        self.downloader = FrotaBoletoDownloader(
            self.pasta_destino_base,
            headless=self.modo_headless,
            nome_frota=self.nome_frota
        )

        try:
            self.downloader.iniciar()
            fila_pendentes = list(tarefas)
            rodada = 1

            while fila_pendentes and not self.cancelar_solicitado:
                if rodada > 1:
                    msg_rodada = f"Reiniciando download para {len(fila_pendentes)} boletos pendentes (Tentativa {rodada})..."
                    self.after(0, lambda m=msg_rodada: self.var_progresso_txt.set(m))
                    logger.info(msg_rodada)
                    # Reinicia a sessão do navegador para limpar abas residuais ou falhas de rede
                    try:
                        self.downloader.reiniciar_sessao()
                    except Exception as err:
                        logger.warning(f"Aviso ao reiniciar sessão do downloader: {err}")
                    time.sleep(2.0)

                falhas_desta_rodada = []

                for item, tipo in fila_pendentes:
                    if self.cancelar_solicitado:
                        break

                    # Suporte à pausa
                    while self.is_pausado and not self.cancelar_solicitado:
                        time.sleep(0.3)
                    if self.cancelar_solicitado:
                        break

                    placa = item["placa"]
                    renavam = item["renavam"]
                    lbl_tipo = "Licenciamento" if tipo == "lic" else "IPVA"

                    concluidos = len(self.tarefas_executadas_sucesso)
                    pct = int((concluidos / total_geral) * 100)
                    self.after(0, lambda c=concluidos, t=total_geral, plc=placa, tp=lbl_tipo, pc=pct, r=rodada:
                        self._atualizar_status_ui(c, t, plc, tp, pc, r)
                    )

                    # Atualiza status na linha da tabela
                    if item["lbl_status"]:
                        txt_tentativa = f"Baixando {lbl_tipo}..." if rodada == 1 else f"Baixando {lbl_tipo} (T{rodada})..."
                        self.after(0, lambda lbl=item["lbl_status"], t=txt_tentativa: lbl.config(
                            text=t, foreground="#2563EB"
                        ))

                    # Executa a geração e salvamento do boleto
                    if tipo == "lic":
                        sucesso, msg = self.downloader.baixar_boleto_licenciamento(placa, renavam)
                    else:
                        sucesso, msg = self.downloader.baixar_dae_ipva(placa, renavam)

                    if sucesso:
                        self.tarefas_executadas_sucesso.add((item["index"], tipo))
                        cor = "#15803D"
                        txt = f"{lbl_tipo}: Concluído"
                    else:
                        falhas_desta_rodada.append((item, tipo))
                        if not self.cancelar_solicitado:
                            cor = "#D97706"
                            txt = f"{lbl_tipo}: Erro (Repetirá)"
                        else:
                            cor = "#DC2626"
                            txt = f"{lbl_tipo}: Erro"

                    if item["lbl_status"]:
                        self.after(0, lambda lbl=item["lbl_status"], t=txt, c=cor: lbl.config(text=t, foreground=c))

                    time.sleep(0.5)

                fila_pendentes = falhas_desta_rodada
                rodada += 1

                if fila_pendentes and not self.cancelar_solicitado:
                    time.sleep(1.5)

            # Finalização (após concluir 100% ou se tiver sido cancelado)
            sucessos = len(self.tarefas_executadas_sucesso)
            self.after(0, lambda s=sucessos, t=total_geral: self._finalizar_downloads_ui(s, t))

        except Exception as e:
            logger.error(f"Erro fatal no lote de downloads: {e}")
            self.after(0, lambda: self.var_progresso_txt.set("Falha no processo de download."))
        finally:
            if self.downloader:
                self.downloader.fechar()
            self.is_baixando = False
            self.is_pausado = False

    def _atualizar_status_ui(self, concluidos: int, total: int, placa: str, tipo: str, pct: int, rodada: int = 1):
        info_rodada = f" [Tentativa {rodada}]" if rodada > 1 else ""
        self.var_progresso_txt.set(f"Baixando ({concluidos}/{total}){info_rodada}: {placa} - {tipo}...")
        self.barra_progresso["value"] = pct

    def _finalizar_downloads_ui(self, sucessos: int, total: int):
        self.barra_progresso["value"] = 100
        self.btn_iniciar_dl.config(state="normal")
        self.btn_pausar_dl.config(state="disabled", text="⏸️ Pausar")
        self.btn_cancelar_dl.config(state="disabled")
        self.btn_abrir_pasta_boletos.config(state="normal")

        if self.cancelar_solicitado:
            self.var_progresso_txt.set(f"Downloads interrompidos! {sucessos} de {total} boletos baixados.")
            messagebox.showinfo("Cancelado", f"Processamento interrompido.\n\nForam salvos {sucessos} de {total} boletos selecionados.")
        else:
            self.var_progresso_txt.set(f"🎉 Todos os downloads foram concluídos com sucesso! ({sucessos}/{total})")
            if sys.platform == "win32" and ConfigManager.carregar().get("som_notificacao", True):
                try:
                    import winsound
                    winsound.MessageBeep(winsound.MB_ICONASTERISK)
                except Exception:
                    pass
            messagebox.showinfo(
                "Downloads Concluídos",
                f"Todos os {total} boletos da lista foram baixados e salvos com sucesso na pasta de destino!\n\n"
                f"Pasta de destino:\n{self.downloader.pasta_destino_final if self.downloader else self.pasta_destino_base}"
            )

    def _cancelar_downloads(self):
        if messagebox.askyesno("Cancelar", "Deseja interromper o download dos boletos restantes?"):
            self.cancelar_solicitado = True
            self.is_pausado = False
            self.btn_cancelar_dl.config(state="disabled")
            self.btn_pausar_dl.config(state="disabled")
            self.var_progresso_txt.set("Interrompendo downloads... Aguarde finalizar tarefa em andamento.")

    def _abrir_pasta_boletos(self):
        caminho = self.downloader.pasta_destino_final if self.downloader else self.pasta_destino_base
        if os.path.exists(caminho):
            os.startfile(caminho)
        else:
            os.startfile(self.pasta_destino_base)

    def _ao_fechar(self):
        try:
            self.unbind_all("<MouseWheel>")
        except Exception:
            pass
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
        self._ajustar_geometria_responsiva()

        # Carrega preferências persistentes
        self.cfg = ConfigManager.carregar()

        # Variáveis de Estado
        self.caminho_arquivo = tk.StringVar(value=self.cfg.get("ultimo_arquivo", ""))
        self.pasta_destino_var = tk.StringVar(value=self.cfg.get("pasta_destino", ""))
        self.arquivo_excel_gerado = ""
        self.arquivo_pdf_gerado = ""
        self.arquivo_pdf_bi_gerado = ""

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
        self.luxury_cards: list[LuxuryMetricCard] = []
        self.status_orb: LiveStatusOrb | None = None
        self.botoes_filtro: dict[str, ttk.Button] = {}
        self.filtro_ativo = "TODOS"

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
        """Aplica tokens visuais de alta classe para Modo Claro e Modo Noturno."""
        if self.modo_noturno:
            bg_root = "#080C14"        # Fundo Noturno Ultra Deep
            bg_card = "#0F172A"        # Superfície de Cartão Noturno
            border_card = "#1F2937"    # Contorno Discreto 1px
            fg_title = "#F8FAFC"
            fg_sub = "#94A3B8"
            fg_text = "#E2E8F0"
            bg_tree = "#0B0F19"
            fg_tree = "#E2E8F0"
            bg_tree_hd = "#141D2E"
            fg_tree_hd = "#F8FAFC"
            bg_sel = "#1E293B"
            fg_sel = "#38BDF8"
            bg_tab = "#080C14"
            fg_tab = "#94A3B8"
            bg_tab_sel = "#0F172A"
            fg_tab_sel = "#38BDF8"
            bg_log = "#05070B"
            fg_log = "#E2E8F0"
            status_fg = "#38BDF8"
            bg_btn_sec = "#1E293B"
            fg_btn_sec = "#E2E8F0"
            bg_btn_sec_act = "#334155"
            bg_btn_muted = "#141C2B"
            fg_btn_muted = "#94A3B8"
            bg_btn_muted_act = "#1E293B"
            bg_btn_theme = "#141C2B"
            fg_btn_theme = "#38BDF8"
            bg_btn_pri = "#2563EB"
            bg_btn_pri_act = "#1D4ED8"
            bg_btn_act = "#059669"
            bg_btn_act_act = "#047857"
            bg_btn_dng = "#9F1239"
            bg_btn_dng_act = "#881337"
            prog_bg = "#38BDF8"
            prog_trough = "#1E293B"
            self.border_card_atual = border_card
            self.border_card_hover_atual = "#38BDF8"
        else:
            bg_root = "#F1F5F9"        # Fundo Claro Porcelana
            bg_card = "#FFFFFF"        # Superfície Branco Puro
            border_card = "#CBD5E1"    # Contorno Suave 1px
            fg_title = "#0F172A"       # Título Titânio Profundo
            fg_sub = "#475569"
            fg_text = "#1E293B"
            bg_tree = "#FFFFFF"
            fg_tree = "#1E293B"
            bg_tree_hd = "#E2E8F0"
            fg_tree_hd = "#0F172A"
            bg_sel = "#E0E7FF"
            fg_sel = "#1E3A8A"
            bg_tab = "#F1F5F9"
            fg_tab = "#64748B"
            bg_tab_sel = "#FFFFFF"
            fg_tab_sel = "#1E40AF"
            bg_log = "#0F172A"
            fg_log = "#F8FAFC"
            status_fg = "#1E40AF"
            bg_btn_sec = "#334155"
            fg_btn_sec = "#FFFFFF"
            bg_btn_sec_act = "#1E293B"
            bg_btn_muted = "#E2E8F0"
            fg_btn_muted = "#334155"
            bg_btn_muted_act = "#CBD5E1"
            bg_btn_theme = "#E2E8F0"
            fg_btn_theme = "#0F172A"
            bg_btn_pri = "#1E40AF"
            bg_btn_pri_act = "#1D4ED8"
            bg_btn_act = "#0D9488"
            bg_btn_act_act = "#0F766E"
            bg_btn_dng = "#E11D48"
            bg_btn_dng_act = "#BE123C"
            prog_bg = "#2563EB"
            prog_trough = "#E2E8F0"
            self.border_card_atual = border_card
            self.border_card_hover_atual = "#2563EB"

        self.root.configure(bg=bg_root)
        self.style.configure(".", background=bg_root, foreground=fg_text, font=("Segoe UI", 9))
        self.style.configure("TFrame", background=bg_root)
        self.style.configure("Card.TFrame", background=bg_card)
        self.style.configure("TLabel", background=bg_root, foreground=fg_text)
        self.style.configure("Card.TLabel", background=bg_card, foreground=fg_text)
        self.style.configure("Header.TLabel", background=bg_root, font=("Segoe UI", 13, "bold"), foreground=fg_title)
        self.style.configure("SubHeader.TLabel", background=bg_root, font=("Segoe UI", 8), foreground=fg_sub)

        self.style.configure("Card.TLabelframe", background=bg_card, relief=tk.SOLID, borderwidth=1)
        self.style.configure("Card.TLabelframe.Label", font=("Segoe UI", 9, "bold"), foreground=fg_title, background=bg_root)
        self.style.configure("Card.TCheckbutton", background=bg_card, foreground=fg_text, font=("Segoe UI", 9))
        self.style.map("Card.TCheckbutton", background=[("active", bg_card)], foreground=[("active", fg_title)])
        self.style.configure("Card.TRadiobutton", background=bg_card, foreground=fg_text, font=("Segoe UI", 9))
        self.style.map("Card.TRadiobutton", background=[("active", bg_card)], foreground=[("active", fg_title)])

        self.style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"), background=bg_btn_pri, foreground="#FFFFFF", padding=(14, 7))
        self.style.map("Primary.TButton", background=[("active", bg_btn_pri_act), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        self.style.configure("PrimaryDownload.TButton", font=("Segoe UI", 9, "bold"), background="#059669", foreground="#FFFFFF", padding=(12, 6))
        self.style.map("PrimaryDownload.TButton", background=[("active", "#047857"), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        self.style.configure("Action.TButton", font=("Segoe UI", 9, "bold"), background=bg_btn_act, foreground="#FFFFFF", padding=(14, 7))
        self.style.map("Action.TButton", background=[("active", bg_btn_act_act), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        self.style.configure("Danger.TButton", font=("Segoe UI", 9, "bold"), background=bg_btn_dng, foreground="#FFFFFF", padding=(10, 6))
        self.style.map("Danger.TButton", background=[("active", bg_btn_dng_act), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        self.style.configure("Secondary.TButton", font=("Segoe UI", 9, "bold"), background=bg_btn_sec, foreground=fg_btn_sec, padding=(10, 5))
        self.style.map("Secondary.TButton", background=[("active", bg_btn_sec_act), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        self.style.configure("Muted.TButton", font=("Segoe UI", 8), background=bg_btn_muted, foreground=fg_btn_muted, padding=(8, 4))
        self.style.map("Muted.TButton", background=[("active", bg_btn_muted_act)])

        self.style.configure("ActiveFilter.TButton", font=("Segoe UI", 8, "bold"), background=bg_btn_pri, foreground="#FFFFFF", padding=(9, 4))
        self.style.map("ActiveFilter.TButton", background=[("active", bg_btn_pri_act)])

        self.style.configure("Theme.TButton", font=("Segoe UI", 9, "bold"), background=bg_btn_theme, foreground=fg_btn_theme, padding=(10, 4))
        self.style.map("Theme.TButton", background=[("active", bg_btn_sec_act if self.modo_noturno else "#CBD5E1")])

        self.style.configure("TNotebook", background=bg_card, borderwidth=0)
        self.style.configure("TNotebook.Tab", background=bg_tab, foreground=fg_tab, padding=(18, 7), font=("Segoe UI", 9, "bold"))
        self.style.map("TNotebook.Tab", background=[("selected", bg_tab_sel)], foreground=[("selected", fg_tab_sel)])

        self.style.configure("Treeview", background=bg_tree, foreground=fg_tree, fieldbackground=bg_tree, rowheight=28, font=("Segoe UI", 9))
        self.style.configure("Treeview.Heading", background=bg_tree_hd, foreground=fg_tree_hd, font=("Segoe UI", 9, "bold"), padding=(6, 7))
        self.style.map("Treeview", background=[("selected", bg_sel)], foreground=[("selected", fg_sel)])

        self.style.configure("Horizontal.TProgressbar", background=prog_bg, troughcolor=prog_trough, bordercolor=prog_trough, thickness=8)
        self.style.configure("TEntry", fieldbackground=bg_card, foreground=fg_text, padding=(6, 4))
        self.style.configure("TCombobox", fieldbackground=bg_card, background=bg_btn_muted, foreground=fg_text, padding=(5, 4))

        if hasattr(self, "tree"):
            if self.modo_noturno:
                self.tree.tag_configure("row_even", background="#0B0F19")
                self.tree.tag_configure("row_odd", background="#0E1422")
                self.tree.tag_configure("pendente", foreground="#64748B")
                self.tree.tag_configure("processando", background="#162742", foreground="#93C5FD", font=("Segoe UI", 9, "bold"))
                self.tree.tag_configure("com_debito", background="#38181C", foreground="#FCA5A5", font=("Segoe UI", 9, "bold"))
                self.tree.tag_configure("sem_debito", background="#0F172A", foreground="#94A3B8")
                self.tree.tag_configure("falha", background="#382914", foreground="#FCD34D", font=("Segoe UI", 9, "bold"))
                self.tree.tag_configure("restricao", background="#382115", foreground="#FDBA74", font=("Segoe UI", 9, "bold"))
                self.tree.tag_configure("pago", background="#132E22", foreground="#86EFAC")
            else:
                self.tree.tag_configure("row_even", background="#FFFFFF")
                self.tree.tag_configure("row_odd", background="#F8FAFC")
                self.tree.tag_configure("pendente", foreground="#64748B")
                self.tree.tag_configure("processando", background="#EFF6FF", foreground="#1D4ED8", font=("Segoe UI", 9, "bold"))
                self.tree.tag_configure("com_debito", background="#FEE2E2", foreground="#991B1B", font=("Segoe UI", 9, "bold"))
                self.tree.tag_configure("sem_debito", background="#F8FAFC", foreground="#475569")
                self.tree.tag_configure("falha", background="#FEF3C7", foreground="#92400E", font=("Segoe UI", 9, "bold"))
                self.tree.tag_configure("restricao", background="#FFEDD5", foreground="#9A3412", font=("Segoe UI", 9, "bold"))
                self.tree.tag_configure("pago", background="#DCFCE7", foreground="#166534")

        if hasattr(self, "txt_logs"):
            self.txt_logs.configure(bg=bg_log, fg=fg_log, insertbackground=fg_text)
            cor_time = "#64748B"
            cor_info = "#38BDF8" if self.modo_noturno else "#0284C7"
            cor_suc = "#34D399" if self.modo_noturno else "#059669"
            cor_warn = "#FBBF24" if self.modo_noturno else "#D97706"
            cor_err = "#F87171" if self.modo_noturno else "#DC2626"
            self.txt_logs.tag_configure("tag_time", foreground=cor_time, font=("Consolas", 8))
            self.txt_logs.tag_configure("tag_info", foreground=cor_info, font=("Consolas", 9, "bold"))
            self.txt_logs.tag_configure("tag_sucesso", foreground=cor_suc, font=("Consolas", 9, "bold"))
            self.txt_logs.tag_configure("tag_alerta", foreground=cor_warn, font=("Consolas", 9, "bold"))
            self.txt_logs.tag_configure("tag_erro", foreground=cor_err, font=("Consolas", 9, "bold"))
            self.txt_logs.tag_configure("tag_corpo", foreground=fg_log, font=("Consolas", 9))

        if hasattr(self, "lbl_status"):
            self.lbl_status.configure(foreground=status_fg)

        if hasattr(self, "btn_tema"):
            self.btn_tema.configure(text="☀️ Modo Claro" if self.modo_noturno else "🌙 Modo Noturno")

        for card in self.luxury_cards:
            card.atualizar_tema(self.modo_noturno)

        if self.status_orb:
            self.status_orb.atualizar_tema(self.modo_noturno)

    def _alternar_tema(self):
        self.modo_noturno = not self.modo_noturno
        self._salvar_preferencias()
        self._aplicar_tema()

    def _ajustar_geometria_responsiva(self):
        """
        Calcula e ajusta as dimensões da janela de forma adaptativa conforme a
        resolução da tela e proporção do monitor do usuário, garantindo que
        o rodapé e os botões de ação nunca fiquem escondidos.
        """
        self.root.update_idletasks()
        tela_w = self.root.winfo_screenwidth()
        tela_h = self.root.winfo_screenheight()

        margem_h = 40
        margem_v = 70

        min_w = 820
        min_h = 520

        if tela_w >= 1920:
            largura = 1140
        elif tela_w >= 1600:
            largura = 1060
        elif tela_w >= 1366:
            largura = 980
        elif tela_w >= 1200:
            largura = 920
        else:
            largura = max(min_w, tela_w - margem_h)

        if tela_h >= 1080:
            altura = 760
        elif tela_h >= 900:
            altura = 710
        elif tela_h >= 768:
            altura = 640
        elif tela_h >= 720:
            altura = 600
        else:
            altura = max(min_h, tela_h - margem_v)

        largura = min(largura, max(min_w, tela_w - margem_h))
        altura = min(altura, max(min_h, tela_h - margem_v))

        pos_x = max(0, (tela_w - largura) // 2)
        pos_y = max(0, (tela_h - margem_v - altura) // 2)

        self.root.minsize(min_w, min_h)
        self.root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    def _criar_layout(self):
        container = ttk.Frame(self.root, padding="12 8 12 8")
        container.pack(fill=tk.BOTH, expand=True)

        # 1. BARRA SUPERIOR ELEGANTE (BRANDING & TELEMETRIA)
        frame_topo = ttk.Frame(container)
        frame_topo.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        frame_titulos = ttk.Frame(frame_topo)
        frame_titulos.pack(side=tk.LEFT, fill=tk.X, expand=True)

        frame_marca = ttk.Frame(frame_titulos)
        frame_marca.pack(anchor="w")

        lbl_badge = tk.Label(
            frame_marca,
            text="⚡ FLEET COMMAND",
            font=("Segoe UI", 7, "bold"),
            bg="#2563EB",
            fg="#FFFFFF",
            padx=6,
            pady=1
        )
        lbl_badge.pack(side=tk.LEFT, padx=(0, 8))

        lbl_titulo = ttk.Label(
            frame_marca,
            text="Automação & Inteligência de Frota",
            style="Header.TLabel"
        )
        lbl_titulo.pack(side=tk.LEFT)

        lbl_subtitulo = ttk.Label(
            frame_titulos,
            text="Detran-CE & Sefaz-CE  •  Auditoria Financeira em Lote, Download de Boletos & Relatórios Executivos",
            style="SubHeader.TLabel"
        )
        lbl_subtitulo.pack(anchor="w", pady=(2, 0))

        # Orbe de Status
        self.status_orb = LiveStatusOrb(frame_topo, modo_noturno=self.modo_noturno)
        self.status_orb.pack(side=tk.RIGHT, padx=(10, 0))

        # Botão de Tema
        self.btn_tema = ttk.Button(
            frame_topo,
            text="🌙 Modo Noturno",
            style="Theme.TButton",
            command=self._alternar_tema
        )
        self.btn_tema.pack(side=tk.RIGHT, padx=(8, 0), pady=2)

        # 2. PAINEL DE CONTROLE EXECUTIVO (CONFIGURAÇÕES E ARQUIVO)
        frame_configs = ttk.LabelFrame(container, text=" Painel de Controle e Parâmetros da Frota ", style="Card.TLabelframe", padding="10 6 10 6")
        frame_configs.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        # Linha 1: Seleção de Planilha e Pasta Destino
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
            text="Nenhuma planilha carregada",
            font=("Segoe UI", 9, "italic"),
            foreground="#64748B",
            style="Card.TLabel"
        )
        self.lbl_caminho.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Linha 2: Motores de Consulta e Velocidade
        linha_filtros = ttk.Frame(frame_configs, style="Card.TFrame")
        linha_filtros.pack(fill=tk.X, pady=(0, 6))

        lbl_filtro_txt = ttk.Label(linha_filtros, text="Motores Ativos:", font=("Segoe UI", 9, "bold"), style="Card.TLabel")
        lbl_filtro_txt.pack(side=tk.LEFT, padx=(0, 8))

        self.chk_ipva = ttk.Checkbutton(linha_filtros, text="🏛️ IPVA (Sefaz)", variable=self.var_ipva, style="Card.TCheckbutton")
        self.chk_ipva.pack(side=tk.LEFT, padx=(0, 8))

        self.chk_lic = ttk.Checkbutton(linha_filtros, text="📋 Licenciamento (Detran)", variable=self.var_lic, style="Card.TCheckbutton")
        self.chk_lic.pack(side=tk.LEFT, padx=(0, 8))

        self.chk_multas = ttk.Checkbutton(linha_filtros, text="🚨 Multas (Detran)", variable=self.var_multas, style="Card.TCheckbutton")
        self.chk_multas.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_toggle_filtros = ttk.Button(
            linha_filtros,
            text="☑️ Todos",
            style="Muted.TButton",
            command=self._alternar_todos_filtros
        )
        self.btn_toggle_filtros.pack(side=tk.LEFT, padx=(0, 12))

        self.combo_workers = ttk.Combobox(
            linha_filtros,
            values=["1x (Sequencial)", "2x Paralelo", "3x Paralelo", "4x Máximo"],
            state="readonly",
            width=14
        )
        idx_w = max(0, min(self.var_workers.get() - 1, 3))
        self.combo_workers.current(idx_w)
        self.combo_workers.bind("<<ComboboxSelected>>", self._ao_mudar_workers)
        self.combo_workers.pack(side=tk.RIGHT)

        lbl_workers = ttk.Label(linha_filtros, text="Velocidade:", font=("Segoe UI", 9, "bold"), style="Card.TLabel")
        lbl_workers.pack(side=tk.RIGHT, padx=(0, 6))

        # Linha 3: Exportação, Headless e Comandos Principais
        linha_opcoes = ttk.Frame(frame_configs, style="Card.TFrame")
        linha_opcoes.pack(fill=tk.X, pady=(0, 2))

        lbl_exp = ttk.Label(linha_opcoes, text="Exportação:", font=("Segoe UI", 9, "bold"), style="Card.TLabel")
        lbl_exp.pack(side=tk.LEFT, padx=(0, 6))

        rb_excel = ttk.Radiobutton(linha_opcoes, text="Excel", value="excel", variable=self.var_formato_exportacao, style="Card.TRadiobutton")
        rb_excel.pack(side=tk.LEFT, padx=(0, 6))

        rb_pdf = ttk.Radiobutton(
            linha_opcoes,
            text="PDF",
            value="pdf",
            variable=self.var_formato_exportacao,
            style="Card.TRadiobutton",
            state="normal" if REPORTLAB_DISPONIVEL else "disabled"
        )
        rb_pdf.pack(side=tk.LEFT, padx=(0, 6))

        rb_ambos = ttk.Radiobutton(
            linha_opcoes,
            text="Ambos",
            value="ambos",
            variable=self.var_formato_exportacao,
            style="Card.TRadiobutton",
            state="normal" if REPORTLAB_DISPONIVEL else "disabled"
        )
        rb_ambos.pack(side=tk.LEFT, padx=(0, 10))

        self.chk_headless = ttk.Checkbutton(
            linha_opcoes,
            text="Chrome Oculto (Headless)",
            variable=self.var_headless,
            style="Card.TCheckbutton"
        )
        self.chk_headless.pack(side=tk.LEFT)

        # Botões de Execução (fixados à direita)
        self.btn_cancelar = ttk.Button(
            linha_opcoes,
            text="⏹️ Cancelar",
            style="Danger.TButton",
            state="disabled",
            command=self._cancelar_processamento
        )
        self.btn_cancelar.pack(side=tk.RIGHT, padx=(4, 0))

        self.btn_pausar = ttk.Button(
            linha_opcoes,
            text="⏸️ Pausar",
            style="Secondary.TButton",
            state="disabled",
            command=self._alternar_pausa
        )
        self.btn_pausar.pack(side=tk.RIGHT, padx=(4, 0))

        self.btn_iniciar = ttk.Button(
            linha_opcoes,
            text="⚡ INICIAR CONSULTA",
            style="Action.TButton",
            state="disabled",
            command=self._iniciar_processamento
        )
        self.btn_iniciar.pack(side=tk.RIGHT)

        # 5. BARRA DE PROGRESSO & PAINEL DE AÇÕES EXECUTIVAS (FIXADO NO FUNDO - NUNCA ESCONDIDO!)
        frame_rodape = ttk.Frame(container)
        frame_rodape.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))

        self.lbl_status = ttk.Label(
            frame_rodape,
            textvariable=self.var_status,
            font=("Segoe UI", 9, "bold")
        )
        self.lbl_status.pack(anchor="w", pady=(0, 3))

        self.barra_progresso = ttk.Progressbar(frame_rodape, mode="determinate")
        self.barra_progresso.pack(fill=tk.X, pady=(0, 6))

        frame_botoes_finais = ttk.Frame(frame_rodape)
        frame_botoes_finais.pack(fill=tk.X)

        # BOTÃO GLAMOUROSO DE DOWNLOAD DE BOLETOS (Fixado com prioridade máxima à direita)
        self.btn_baixar_boletos = ttk.Button(
            frame_botoes_finais,
            text="📥 Central de Download de Boletos (PDF)",
            style="PrimaryDownload.TButton",
            state="disabled",
            command=self._abrir_central_download_boletos
        )
        self.btn_baixar_boletos.pack(side=tk.RIGHT)

        self.btn_abrir_excel = ttk.Button(
            frame_botoes_finais,
            text="📊 Relatório Excel",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_excel_gerado
        )
        self.btn_abrir_excel.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_abrir_pdf = ttk.Button(
            frame_botoes_finais,
            text="📄 Relatório PDF",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_pdf_gerado
        )
        self.btn_abrir_pdf.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_abrir_pdf_bi = ttk.Button(
            frame_botoes_finais,
            text="📈 Relatório BI (PDF)",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_pdf_bi_gerado
        )
        self.btn_abrir_pdf_bi.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_abrir_pasta = ttk.Button(
            frame_botoes_finais,
            text="📂 Pasta Destino",
            style="Secondary.TButton",
            state="disabled",
            command=self._abrir_pasta_destino
        )
        self.btn_abrir_pasta.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_reconsultar_falhas = ttk.Button(
            frame_botoes_finais,
            text="⚠️ Reconsultar Falhas",
            style="Secondary.TButton",
            state="disabled",
            command=self._reconsultar_apenas_falhas
        )
        self.btn_reconsultar_falhas.pack(side=tk.LEFT, padx=(0, 8))

        # 4. DASHBOARD DE MÉTRICAS ORNAMENTADO (FIXADO LOGO ACIMA DO RODAPÉ)
        frame_metricas = ttk.Frame(container)
        frame_metricas.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))

        self.luxury_cards.clear()
        card_configs = [
            ("FROTA TOTAL", self.var_metricas_veiculos, "🚘", "#3B82F6", "✦ Registros Válidos"),
            ("LICENCIAMENTO", self.var_metricas_lic, "📋", "#F59E0B", "✦ Detran-CE"),
            ("IPVA EM ABERTO", self.var_metricas_ipva, "🏛️", "#F43F5E", "✦ Sefaz-CE"),
            ("MULTAS DETRAN", self.var_metricas_multas, "🚨", "#EA580C", "✦ Infrações Ativas"),
            ("TOTAL CONSOLIDADO", self.var_metricas_total, "💎", "#10B981", "✦ Saldo Total a Liquidar"),
        ]

        for titulo, var_val, icone, cor_acento, sub in card_configs:
            c = LuxuryMetricCard(
                frame_metricas,
                titulo=titulo,
                var_valor=var_val,
                icone=icone,
                cor_acento=cor_acento,
                subtexto=sub,
                modo_noturno=self.modo_noturno
            )
            c.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=3)
            self.luxury_cards.append(c)

        # 3. ACOMPANHAMENTO EM TEMPO REAL: TABELA E CONSOLE SINTÁTICO (EXPANDE NO CENTRO)
        frame_centro = ttk.LabelFrame(container, text=" Central de Monitoramento e Auditoria ", style="Card.TLabelframe", padding="8 6 8 6")
        frame_centro.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 4))

        self.notebook = ttk.Notebook(frame_centro)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Aba 1: Tabela com Barra de Filtros em Pílulas
        aba_tabela = ttk.Frame(self.notebook, padding="4 4 4 4")
        self.notebook.add(aba_tabela, text=" 📊 Tabela Dinâmica da Frota ")

        # Barra de Busca & Filtros por Pílulas
        frame_filtros_bar = ttk.Frame(aba_tabela)
        frame_filtros_bar.pack(fill=tk.X, pady=(0, 6))

        # Botão redundante de atalho direto para Boletos na barra superior da tabela
        self.btn_boletos_tabela = ttk.Button(
            frame_filtros_bar,
            text="📥 Central de Boletos (PDF)",
            style="PrimaryDownload.TButton",
            state="disabled",
            command=self._abrir_central_download_boletos
        )
        self.btn_boletos_tabela.pack(side=tk.RIGHT)

        ttk.Label(frame_filtros_bar, text="🔍 Busca:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.entry_busca = ttk.Entry(frame_filtros_bar, textvariable=self.var_busca, width=20)
        self.entry_busca.pack(side=tk.LEFT, padx=(0, 6))
        self.var_busca.trace_add("write", lambda *_: self._filtrar_tabela())

        ttk.Button(frame_filtros_bar, text="✖ Limpar", style="Muted.TButton", command=lambda: self.var_busca.set("")).pack(side=tk.LEFT, padx=(0, 10))

        # Pílulas Rápidas de Filtro
        self.botoes_filtro["TODOS"] = ttk.Button(
            frame_filtros_bar, text="✦ Todos (0)", style="ActiveFilter.TButton",
            command=lambda: self._definir_filtro_pílula("TODOS")
        )
        self.botoes_filtro["TODOS"].pack(side=tk.LEFT, padx=(0, 3))

        self.botoes_filtro["DEBITOS"] = ttk.Button(
            frame_filtros_bar, text="⚠️ Débitos (0)", style="Muted.TButton",
            command=lambda: self._definir_filtro_pílula("DEBITOS")
        )
        self.botoes_filtro["DEBITOS"].pack(side=tk.LEFT, padx=(0, 3))

        self.botoes_filtro["PAGOS"] = ttk.Button(
            frame_filtros_bar, text="✅ Pagos (0)", style="Muted.TButton",
            command=lambda: self._definir_filtro_pílula("PAGOS")
        )
        self.botoes_filtro["PAGOS"].pack(side=tk.LEFT, padx=(0, 3))

        self.botoes_filtro["RESTRICAO"] = ttk.Button(
            frame_filtros_bar, text="🚨 Restrições (0)", style="Muted.TButton",
            command=lambda: self._definir_filtro_pílula("RESTRICAO")
        )
        self.botoes_filtro["RESTRICAO"].pack(side=tk.LEFT, padx=(0, 3))

        self.botoes_filtro["FALHAS"] = ttk.Button(
            frame_filtros_bar, text="❌ Falhas (0)", style="Muted.TButton",
            command=lambda: self._definir_filtro_pílula("FALHAS")
        )
        self.botoes_filtro["FALHAS"].pack(side=tk.LEFT)

        # Treeview com Colunas Alinhadas
        colunas = ("num", "placa", "renavam", "licenciamento", "ipva", "multas", "total_pagar", "situacao")
        self.tree = ttk.Treeview(aba_tabela, columns=colunas, show="headings", selectmode="extended", height=6)

        cabecalhos_nomes = [
            ("num", "#", 45, "center"),
            ("placa", "Placa", 95, "center"),
            ("renavam", "RENAVAM", 125, "center"),
            ("licenciamento", "Licenciamento", 130, "e"),
            ("ipva", "IPVA", 130, "e"),
            ("multas", "Multas", 130, "e"),
            ("total_pagar", "Total a Pagar", 135, "e"),
            ("situacao", "Situação Auditada", 150, "center"),
        ]

        for col_id, titulo, larg, alinhamento in cabecalhos_nomes:
            self.tree.heading(col_id, text=f"{titulo} ↕", command=lambda c=col_id: self._ordenar_coluna(c))
            self.tree.column(col_id, width=larg, anchor=alinhamento)

        scroll_y = ttk.Scrollbar(aba_tabela, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(aba_tabela, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Menu de Contexto
        self.menu_contexto = tk.Menu(self.tree, tearoff=0)
        self.menu_contexto.add_command(label="🔄 Reconsultar Veículo Selecionado", command=self._reconsultar_selecionado)
        self.tree.bind("<Button-3>", self._exibir_menu_contexto)

        # Aba 2: Console de Logs com Destaque Sintático
        aba_logs = ttk.Frame(self.notebook, padding="4 4 4 4")
        self.notebook.add(aba_logs, text=" 📜 Terminal Executivo de Logs ")

        self.txt_logs = tk.Text(aba_logs, height=6, wrap=tk.WORD, font=("Consolas", 9), relief=tk.FLAT, padx=10, pady=10)
        scroll_logs = ttk.Scrollbar(aba_logs, orient="vertical", command=self.txt_logs.yview)
        self.txt_logs.configure(yscrollcommand=scroll_logs.set)
        self.txt_logs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_logs.pack(side=tk.RIGHT, fill=tk.Y)

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
            self._adicionar_log("Processamento pausado.", nivel="ALERTA")
            if self.status_orb:
                self.status_orb.definir_status("pausado", "PROCESSAMENTO PAUSADO")
        else:
            self.btn_pausar.config(text="⏸️ Pausar")
            self.var_status.set("Retomando consultas de onde parou...")
            self._adicionar_log("Processamento retomado.", nivel="INFO")
            if self.status_orb:
                self.status_orb.definir_status("processando", "RETOMANDO CONSULTAS...")

    @staticmethod
    def _possui_indicador_falha(*valores: str) -> bool:
        """Verifica de forma unificada com o Core se algum valor indica falha, timeout ou veículo não localizado."""
        termos = ("falha", "erro", "timeout", "inválid", "invalid", "não localizado", "nao localizado")
        return any(
            termo in str(v).lower()
            for v in valores
            for termo in termos
        )

    def _adicionar_log(self, mensagem: str, nivel: str = "INFO"):
        """Exibe logs com realce sintático de cores e marcas de tempo de forma thread-safe."""
        logger.info(f"[{nivel}] {mensagem}")
        hora = datetime.now().strftime("%H:%M:%S")

        tag_nivel = "tag_info"
        if "SUCESSO" in nivel.upper() or "PAGO" in mensagem.upper():
            tag_nivel = "tag_sucesso"
        elif "ALERTA" in nivel.upper() or "AVISO" in nivel.upper() or "FALHA" in mensagem.upper():
            tag_nivel = "tag_alerta"
        elif "ERRO" in nivel.upper():
            tag_nivel = "tag_erro"

        def _inserir_no_widget():
            try:
                if hasattr(self, "txt_logs") and self.txt_logs.winfo_exists():
                    self.txt_logs.insert(tk.END, f"[{hora}] ", ("tag_time",))
                    self.txt_logs.insert(tk.END, f"[{nivel}] ", (tag_nivel,))
                    self.txt_logs.insert(tk.END, f"{mensagem}\n", ("tag_corpo",))
                    self.txt_logs.see(tk.END)
            except Exception:
                pass

        if threading.current_thread() is threading.main_thread():
            _inserir_no_widget()
        else:
            try:
                self.root.after(0, _inserir_no_widget)
            except Exception:
                pass

    def _definir_filtro_pílula(self, tipo: str):
        self.filtro_ativo = tipo
        for k, btn in self.botoes_filtro.items():
            btn.configure(style="ActiveFilter.TButton" if k == tipo else "Muted.TButton")
        self._filtrar_tabela()

    def _atualizar_contadores_filtro(self):
        if self.df_carregado is None:
            return

        total = len(self.df_carregado)
        debitos = 0
        pagos = 0
        restricoes = 0
        falhas = len(self.indices_em_falha)

        for idx, dados in self.resultados_memoria.items():
            lic = dados.get("lic", "-")
            ipva = dados.get("ipva", "-")
            multas = dados.get("multas", "-")
            m_av = dados.get("multas_a_vencer")
            if all(str(v).strip() in ["-", "N/A", "Não Consultado", ""] for v in [lic, ipva, multas]):
                continue
            tem_falha = self._possui_indicador_falha(lic, ipva, multas)
            if tem_falha:
                continue

            tot = calcular_total_a_pagar_veiculo(lic, ipva, multas, m_av)
            is_restr = "restriç" in str(lic).lower() or "restric" in str(lic).lower()
            is_pago = (str(lic).strip().upper() == "PAGO" or str(ipva).strip().upper() == "PAGO")

            if is_restr:
                restricoes += 1
            if tot > 0:
                debitos += 1
            elif is_pago or tot == 0:
                pagos += 1

        self.botoes_filtro["TODOS"].configure(text=f"✦ Todos ({total})")
        self.botoes_filtro["DEBITOS"].configure(text=f"⚠️ Com Débitos ({debitos})")
        self.botoes_filtro["PAGOS"].configure(text=f"✅ Pagos ({pagos})")
        self.botoes_filtro["RESTRICAO"].configure(text=f"🚨 Restrições ({restricoes})")
        self.botoes_filtro["FALHAS"].configure(text=f"❌ Falhas ({falhas})")

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
            nome_base_saida = FrotaDataProcessor.gerar_nome_base_resultado(caminho)
            self.lbl_caminho.config(
                text=f"Carregado: {nome} ({len(df)} veículos) ➔ Saída: '{nome_base_saida}.xlsx/.pdf'",
                foreground="#90BE9F" if self.modo_noturno else "#15803D",
                font=("Segoe UI", 9, "bold")
            )

            # Reseta caminhos de relatórios anteriores para evitar conflitos entre frotas
            self.arquivo_excel_gerado = ""
            self.arquivo_pdf_gerado = ""
            self.btn_abrir_excel.config(state="disabled")
            self.btn_abrir_pdf.config(state="disabled")

            self.btn_iniciar.config(state="normal")
            self.var_status.set(f"Arquivo carregado com sucesso: {len(df)} veículos prontos. Saída: {nome_base_saida}")
            self._preencher_tabela_inicial()
            self._adicionar_log(f"Planilha '{nome}' carregada ({len(df)} registros). Arquivos de saída configurados: '{nome_base_saida}.xlsx' e '{nome_base_saida}.pdf'")

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

        self.btn_baixar_boletos.config(state="disabled")
        if hasattr(self, "btn_boletos_tabela"):
            self.btn_boletos_tabela.config(state="disabled")

        for idx, row in df.iterrows():
            placa = str(row[self.col_placa]).strip().upper()
            renavam = str(row[self.col_renavam]).strip()

            tag = "pendente"
            if not PADRAO_PLACA.match(placa):
                tag = "falha"
                self.indices_em_falha.append(idx)

            zebra = "row_even" if idx % 2 == 0 else "row_odd"
            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(idx + 1, placa, renavam, "-", "-", "-", "-", "Aguardando"),
                tags=(zebra, tag)
            )

        self._atualizar_contadores_filtro()
        if self.status_orb:
            self.status_orb.definir_status("pronto", f"{len(self.df_carregado)} VEÍCULOS PRONTOS")

    def _filtrar_tabela(self):
        """Filtra as linhas visíveis na Treeview combinando busca textual e pílula ativa."""
        termo = self.var_busca.get().strip().upper()
        if self.df_carregado is None:
            return

        for idx in range(len(self.df_carregado)):
            row = self.df_carregado.loc[idx]
            placa = str(row[self.col_placa]).strip().upper()
            renavam = str(row[self.col_renavam]).strip()
            dados_mem = self.resultados_memoria.get(idx, {})

            situacao = "Aguardando"
            total_real = 0.0
            tag_linha = "pendente"
            tem_falha = idx in self.indices_em_falha

            if idx in self.resultados_memoria:
                lic = dados_mem.get("lic", "-")
                ipva = dados_mem.get("ipva", "-")
                multas = dados_mem.get("multas", "-")
                m_av = dados_mem.get("multas_a_vencer")
                tem_falha = self._possui_indicador_falha(lic, ipva, multas)
                total_real = calcular_total_a_pagar_veiculo(lic, ipva, multas, m_av)
                is_restr = "restriç" in str(lic).lower() or "restric" in str(lic).lower()
                is_pago = (str(lic).strip().upper() == "PAGO" or str(ipva).strip().upper() == "PAGO")

                if tem_falha:
                    situacao = "Falha na Consulta"
                    tag_linha = "falha"
                elif is_restr:
                    situacao = "Com Débitos / Restrição" if total_real > 0 else "Restrição"
                    tag_linha = "restricao"
                elif total_real > 0:
                    situacao = "Com Débitos"
                    tag_linha = "com_debito"
                else:
                    situacao = "Sem Débitos"
                    tag_linha = "pago" if is_pago else "sem_debito"

            # 1. Filtro de Busca Textual
            passa_termo = (not termo) or (termo in placa) or (termo in renavam) or (termo in situacao.upper())

            # 2. Filtro de Pílula
            passa_pilula = True
            if self.filtro_ativo == "DEBITOS":
                passa_pilula = (total_real > 0 and not tem_falha)
            elif self.filtro_ativo == "PAGOS":
                passa_pilula = (tag_linha in ["pago", "sem_debito"] and not tem_falha)
            elif self.filtro_ativo == "RESTRICAO":
                passa_pilula = (tag_linha == "restricao")
            elif self.filtro_ativo == "FALHAS":
                passa_pilula = tem_falha

            corresponde = passa_termo and passa_pilula

            if corresponde:
                if not self.tree.exists(str(idx)):
                    total_str = formatar_float_moeda(total_real) if idx in self.resultados_memoria else "-"
                    lic_v = dados_mem.get("lic", "-")
                    ipva_v = dados_mem.get("ipva", "-")
                    multas_v = dados_mem.get("multas", "-")
                    zebra = "row_even" if idx % 2 == 0 else "row_odd"
                    self.tree.insert(
                        "",
                        "end",
                        iid=str(idx),
                        values=(idx + 1, placa, renavam, lic_v, ipva_v, multas_v, total_str, situacao),
                        tags=(zebra, tag_linha)
                    )
            else:
                if self.tree.exists(str(idx)):
                    self.tree.delete(str(idx))

    def _ordenar_coluna(self, col: str):
        """Ordena a Treeview dinamicamente pelo cabeçalho clicado (crescente/decrescente)."""
        itens = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]

        # Detecta se é coluna numérica de índice ou de moeda
        if col == "num":
            itens.sort(key=lambda t: int(t[0]) if str(t[0]).isdigit() else 0, reverse=self.ordem_reversa)
        elif col in ["licenciamento", "ipva", "multas", "total_pagar"]:
            itens.sort(key=lambda t: (formatar_moeda_float(t[0]), str(t[0]).upper()), reverse=self.ordem_reversa)
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
        indices_falhas = list(self.indices_em_falha)
        for item_id in self.tree.get_children():
            tags = self.tree.item(item_id, "tags")
            if "falha" in tags and int(item_id) not in indices_falhas:
                indices_falhas.append(int(item_id))

        if not indices_falhas:
            messagebox.showinfo("Nenhuma Falha", "Não há veículos com status de falha na tabela.")
            return

        indices_falhas.sort()
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

        # Ao iniciar uma nova rodada completa, limpa a tabela e o cache de memória
        # para garantir que filtros modificados (ex: desmarcar multas) não mantenham resíduos
        if indices_alvo is None:
            self._preencher_tabela_inicial()

        self.is_executando = True
        self.is_pausado = False
        self.cancelar_solicitado = False

        self.btn_iniciar.config(state="disabled")
        self.btn_selecionar.config(state="disabled")
        self.btn_pasta_destino.config(state="disabled")
        self.btn_reconsultar_falhas.config(state="disabled")
        self.btn_baixar_boletos.config(state="disabled")
        if hasattr(self, "btn_boletos_tabela"):
            self.btn_boletos_tabela.config(state="disabled")
        self.btn_pausar.config(state="normal", text="⏸️ Pausar")
        self.btn_cancelar.config(state="normal")
        self.btn_abrir_excel.config(state="disabled")
        self.btn_abrir_pdf.config(state="disabled")
        if hasattr(self, "btn_abrir_pdf_bi"):
            self.btn_abrir_pdf_bi.config(state="disabled")
        self.btn_abrir_pasta.config(state="disabled")
        self.barra_progresso["value"] = 0

        if self.status_orb:
            self.status_orb.definir_status("processando", "INICIANDO CONSULTA...")

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
            if self.status_orb:
                self.status_orb.definir_status("pausado", "CANCELAMENTO SOLICITADO")

    def _salvar_checkpoint(self, pasta_destino: str, df: pd.DataFrame, caminho_arquivo: str = ""):
        try:
            qtd_total = len(df)
            lics = [self.resultados_memoria.get(i, {}).get("lic", "-") for i in range(qtd_total)]
            ipvas = [self.resultados_memoria.get(i, {}).get("ipva", "-") for i in range(qtd_total)]
            multas = [self.resultados_memoria.get(i, {}).get("multas", "-") for i in range(qtd_total)]
            multas_av = [self.resultados_memoria.get(i, {}).get("multas_a_vencer") for i in range(qtd_total)]

            df_chk = FrotaDataProcessor.consolidar_resultados(
                df, lics, ipvas, multas, qtd_total, resultados_multas_a_vencer=multas_av
            )
            _, _, chk_path = FrotaDataProcessor.obter_caminhos_arquivos_saida(
                caminho_arquivo or self.caminho_arquivo.get(), pasta_destino
            )
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

        arquivo_excel, arquivo_pdf, chk_path = FrotaDataProcessor.obter_caminhos_arquivos_saida(
            caminho_arquivo, pasta_destino
        )
        arquivo_pdf_bi = FrotaDataProcessor.obter_caminho_pdf_bi(caminho_arquivo, pasta_destino)
        nome_frota_log = FrotaDataProcessor.gerar_nome_base_resultado(caminho_arquivo)

        scrapers_ativos = []
        try:
            df = self.df_carregado.copy()
            total_geral = len(df)

            lista_indices = indices_alvo if indices_alvo is not None else list(df.index)
            total_lote = len(lista_indices)

            self.root.after(0, lambda: self.var_status.set("Inicializando navegadores Chrome..."))
            self._adicionar_log(f"Iniciando lote de {total_lote} veículos ({num_workers} worker(s)). Destino: '{nome_frota_log}.xlsx' e '{nome_frota_log}.pdf'")

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
                        multas_a_vencer = "R$ 0,00"
                        multas_vencidas = "R$ 0,00"
                        val_sefaz_ipva = "-"

                        if not PADRAO_PLACA.match(placa):
                            if consultar_lic:
                                val_detran_lic = "Placa Inválida"
                            if consultar_multas:
                                val_detran_multas = "Placa Inválida"
                            if consultar_ipva:
                                val_sefaz_ipva = "Placa Inválida"
                        else:
                            if consultar_lic or consultar_multas:
                                res_detran = scraper.consultar_detran(
                                    placa, renavam, consultar_lic, consultar_multas
                                )
                                val_detran_lic = res_detran[0]
                                if str(val_detran_lic).strip() in ["R$ 0,00", "0,00"]:
                                    val_detran_lic = "PAGO"
                                val_detran_multas = res_detran[1]
                                multas_a_vencer = getattr(res_detran, "multas_a_vencer", "R$ 0,00")
                                multas_vencidas = getattr(res_detran, "multas_vencidas", "R$ 0,00")
                                time.sleep(0.3)

                            if consultar_ipva:
                                val_sefaz_ipva = scraper.consultar_sefaz_ipva(placa, renavam)
                                if str(val_sefaz_ipva).strip() in ["R$ 0,00", "0,00"]:
                                    val_sefaz_ipva = "PAGO"
                                time.sleep(0.3)

                        with lock_memoria:
                            self.resultados_memoria[idx] = {
                                "lic": val_detran_lic,
                                "ipva": val_sefaz_ipva,
                                "multas": val_detran_multas,
                                "multas_a_vencer": multas_a_vencer,
                                "multas_vencidas": multas_vencidas
                            }
                            concluidos_contador[0] += 1
                            c_atual = concluidos_contador[0]

                            soma_lic = sum(formatar_moeda_float(d.get("lic", "0")) for d in self.resultados_memoria.values())
                            soma_ipva = sum(formatar_moeda_float(d.get("ipva", "0")) for d in self.resultados_memoria.values())
                            soma_multas = sum(formatar_moeda_float(d.get("multas", "0")) for d in self.resultados_memoria.values())

                            # Total a pagar real: Licenciamento (já contabiliza multas vencidas) + IPVA + multas a vencer
                            total_acumulado = sum(
                                calcular_total_a_pagar_veiculo(
                                    d.get("lic", "0"),
                                    d.get("ipva", "0"),
                                    d.get("multas", "0"),
                                    d.get("multas_a_vencer")
                                )
                                for d in self.resultados_memoria.values()
                            )
                            pct = int((c_atual / total_lote) * 100)

                            if c_atual % 5 == 0:
                                self._salvar_checkpoint(pasta_destino, df, caminho_arquivo)

                        self.root.after(0, lambda p=c_atual, t=total_lote, plc=placa, pc=pct: self._atualizar_progresso_ui(p, t, plc, pc))
                        self.root.after(
                            0,
                            self._atualizar_linha_tabela,
                            idx,
                            val_detran_lic,
                            val_sefaz_ipva,
                            val_detran_multas,
                            multas_a_vencer,
                            c_atual,
                            total_geral,
                            soma_lic,
                            soma_ipva,
                            soma_multas,
                            total_acumulado
                        )
                        self._adicionar_log(f"[{c_atual}/{total_lote}] {placa} | Lic: {val_detran_lic} | IPVA: {val_sefaz_ipva} | Multas: {val_detran_multas}")

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
            multas_a_vencer_final = [self.resultados_memoria.get(i, {}).get("multas_a_vencer") for i in range(total_geral)]

            df_final = FrotaDataProcessor.consolidar_resultados(
                df, lics_final, ipvas_final, multas_final, total_geral,
                resultados_multas_a_vencer=multas_a_vencer_final
            )

            soma_lic_f = sum(formatar_moeda_float(v) for v in lics_final)
            soma_ipva_f = sum(formatar_moeda_float(v) for v in ipvas_final)
            soma_multas_f = sum(formatar_moeda_float(v) for v in multas_final)
            soma_total_real_f = sum(
                calcular_total_a_pagar_veiculo(
                    self.resultados_memoria.get(i, {}).get("lic", "0"),
                    self.resultados_memoria.get(i, {}).get("ipva", "0"),
                    self.resultados_memoria.get(i, {}).get("multas", "0"),
                    self.resultados_memoria.get(i, {}).get("multas_a_vencer")
                )
                for i in range(total_geral)
            )
            totais_dict = {
                "lic": formatar_float_moeda(soma_lic_f),
                "ipva": formatar_float_moeda(soma_ipva_f),
                "multas": formatar_float_moeda(soma_multas_f),
                "total": formatar_float_moeda(soma_total_real_f)
            }

            caminho_excel_gerado = ""
            caminho_pdf_gerado = ""
            caminho_pdf_bi_gerado = ""

            filtros_ativos = []
            if consultar_ipva: filtros_ativos.append("IPVA")
            if consultar_lic: filtros_ativos.append("Licenciamento")
            if consultar_multas: filtros_ativos.append("Multas")

            identificador_frota = re.sub(r"^resultados[_\-\s]*", "", nome_frota_log, flags=re.IGNORECASE).strip() or nome_frota_log

            if formato_export in ["excel", "ambos"]:
                caminho_excel_gerado = FrotaRelatorioExporter.exportar_excel(df_final, arquivo_excel)
                self.arquivo_excel_gerado = caminho_excel_gerado
                self._adicionar_log(f"Planilha Excel gravada com sucesso: {os.path.basename(caminho_excel_gerado)}")

            if formato_export in ["pdf", "ambos"] and REPORTLAB_DISPONIVEL:
                caminho_pdf_gerado = FrotaRelatorioExporter.exportar_pdf(
                    df_final, arquivo_pdf, totais_dict, filtros_ativos, identificador_frota=identificador_frota
                )
                self.arquivo_pdf_gerado = caminho_pdf_gerado
                self._adicionar_log(f"Relatório PDF gerado com sucesso: {os.path.basename(caminho_pdf_gerado)}")

                caminho_pdf_bi_gerado = FrotaRelatorioExporter.exportar_pdf_bi(
                    df_final, arquivo_pdf_bi, totais_dict, filtros_ativos, identificador_frota=identificador_frota
                )
                self.arquivo_pdf_bi_gerado = caminho_pdf_bi_gerado
                self._adicionar_log(f"Relatório PDF BI gerado com sucesso: {os.path.basename(caminho_pdf_bi_gerado)}")

            # Remove checkpoint temporário se tudo correu normalmente
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
            self.tree.item(str(index), values=(index + 1, placa, renavam, "...", "...", "...", "...", "Processando..."), tags=("processando",))
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
        multas_a_vencer: str,
        qtd_concluidos: int,
        total_geral: int,
        soma_lic: float,
        soma_ipva: float,
        soma_multas: float,
        total_acumulado: float
    ):
        if str(lic).strip() in ["R$ 0,00", "0,00"]:
            lic = "PAGO"
        if str(ipva).strip() in ["R$ 0,00", "0,00"]:
            ipva = "PAGO"

        total_veiculo = calcular_total_a_pagar_veiculo(lic, ipva, multas, multas_a_vencer)
        total_veiculo_str = formatar_float_moeda(total_veiculo)
        tem_falha = self._possui_indicador_falha(lic, ipva, multas)
        is_restricao = "restriç" in str(lic).lower() or "restric" in str(lic).lower()
        is_pago = (str(lic).strip().upper() == "PAGO" or str(ipva).strip().upper() == "PAGO")

        if tem_falha:
            situacao = "Falha na Consulta"
            tag = "falha"
            if index not in self.indices_em_falha:
                self.indices_em_falha.append(index)
        elif is_restricao:
            tag = "restricao"
            if total_veiculo > 0:
                situacao = "Com Débitos / Restrição"
            else:
                situacao = "Restrição"
            if index in self.indices_em_falha:
                self.indices_em_falha.remove(index)
        elif total_veiculo > 0:
            situacao = "Com Débitos"
            tag = "com_debito"
            if index in self.indices_em_falha:
                self.indices_em_falha.remove(index)
        else:
            situacao = "Sem Débitos"
            tag = "pago" if is_pago else "sem_debito"
            if index in self.indices_em_falha:
                self.indices_em_falha.remove(index)

        row = self.df_carregado.loc[index]
        placa = str(row[self.col_placa]).strip().upper()
        renavam = str(row[self.col_renavam]).strip()

        zebra = "row_even" if index % 2 == 0 else "row_odd"
        if self.tree.exists(str(index)):
            self.tree.item(
                str(index),
                values=(index + 1, placa, renavam, lic, ipva, multas, total_veiculo_str, situacao),
                tags=(zebra, tag)
            )
        if self.filtro_ativo != "TODOS" or self.var_busca.get().strip():
            self._filtrar_tabela()

        self.var_metricas_veiculos.set(f"{qtd_concluidos} / {total_geral}")
        self.var_metricas_lic.set(formatar_float_moeda(soma_lic))
        self.var_metricas_ipva.set(formatar_float_moeda(soma_ipva))
        self.var_metricas_multas.set(formatar_float_moeda(soma_multas))
        self.var_metricas_total.set(formatar_float_moeda(total_acumulado))

        self._atualizar_contadores_filtro()
        if self.status_orb:
            self.status_orb.definir_status("processando", f"AUDITANDO: #{index + 1} / {total_geral}")

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
        if self.arquivo_pdf_bi_gerado and os.path.exists(self.arquivo_pdf_bi_gerado) and hasattr(self, "btn_abrir_pdf_bi"):
            self.btn_abrir_pdf_bi.config(state="normal")

        if self.indices_em_falha:
            self.btn_reconsultar_falhas.config(state="normal")

        # Verifica se há débitos para habilitar o botão de Download de Boletos
        total_com_debitos = sum(
            1 for d in self.resultados_memoria.values()
            if formatar_moeda_float(d.get("lic", "0")) > 0 or formatar_moeda_float(d.get("ipva", "0")) > 0
        )
        if total_com_debitos > 0:
            self.btn_baixar_boletos.config(state="normal")
            if hasattr(self, "btn_boletos_tabela"):
                self.btn_boletos_tabela.config(state="normal")

        # Notificação Sonora no Windows
        if sys.platform == "win32" and self.cfg.get("som_notificacao", True):
            try:
                import winsound
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
            except Exception:
                pass

        if self.status_orb:
            if foi_cancelado:
                self.status_orb.definir_status("pausado", "INTERROMPIDO")
            else:
                self.status_orb.definir_status("pronto", "AUDITORIA CONCLUÍDA")

        if foi_cancelado:
            self.var_status.set("Processamento interrompido. Resultados parciais salvos.")
            messagebox.showinfo("Interrompido", "O processamento foi interrompido e os veículos concluídos foram gravados com sucesso.")
        else:
            self.var_status.set("Processamento concluído com sucesso!")
            msg = f"Consultas concluídas com sucesso!\n\nTotal a Pagar na Frota: {self.var_metricas_total.get()}"
            if arq_excel and os.path.exists(arq_excel):
                msg += f"\n\n📊 Planilha Excel: {os.path.basename(arq_excel)}"
            if arq_pdf and os.path.exists(arq_pdf):
                msg += f"\n📑 Relatório PDF: {os.path.basename(arq_pdf)}"
            if self.arquivo_pdf_bi_gerado and os.path.exists(self.arquivo_pdf_bi_gerado):
                msg += f"\n📈 Relatório BI PDF: {os.path.basename(self.arquivo_pdf_bi_gerado)}"
            if total_com_debitos > 0:
                msg += f"\n\n💡 {total_com_debitos} veículo(s) possuem débitos de Licenciamento ou IPVA.\nVocê pode usar o botão 'Baixar Boletos' abaixo para emitir os PDFs."
            messagebox.showinfo("Sucesso", msg)

    def _abrir_central_download_boletos(self):
        """Abre a Janela Modal de Download de Boletos para os veículos com débitos de Licenciamento e IPVA."""
        if self.df_carregado is None or not self.resultados_memoria:
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
            modo_headless=self.var_headless.get(),
            caminho_arquivo=self.caminho_arquivo.get()
        )

    def _abrir_excel_gerado(self):
        if self.arquivo_excel_gerado and os.path.exists(self.arquivo_excel_gerado):
            os.startfile(self.arquivo_excel_gerado)

    def _abrir_pdf_gerado(self):
        if self.arquivo_pdf_gerado and os.path.exists(self.arquivo_pdf_gerado):
            os.startfile(self.arquivo_pdf_gerado)

    def _abrir_pdf_bi_gerado(self):
        if self.arquivo_pdf_bi_gerado and os.path.exists(self.arquivo_pdf_bi_gerado):
            os.startfile(self.arquivo_pdf_bi_gerado)

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
    configurar_dpi()
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
