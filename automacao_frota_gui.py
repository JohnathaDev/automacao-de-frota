"""
========================================================================================
Aplicativo GUI Principal (v4.1): Consulta de Frota - Detran/Sefaz CE
========================================================================================
Interface Gráfica Corporativa & Comercial (SaaS / Fintech B2B Edition):
- Barra de título nativa do Windows sincronizada com Modo Claro / Noturno (DWM API).
- KPI Cards no topo com barra de proporção financeira em tempo real e badges circulares.
- Toggle Switches modernos em formato de pílula (substituindo checkboxes clássicos).
- Orbe de telemetria com animação pulsante suave durante o processamento.
- Barra de progresso customizada em Canvas com cálculo de velocidade (s/veículo) e ETA.
- Aba "📈 Painel Analítico (BI da Frota)" com Gráfico de Rosca (Donut), Saúde da
  Frota e Ranking Top 10 Veículos/Proprietários de maior impacto financeiro.
- Geração automática do 2º PDF Executivo com Relatório de BI (*_BI.pdf).
- Menu rápido "🕘 Recentes" integrado ao histórico e arquivos da pasta.
- Modal "🔍 Raio-X Executivo do Veículo" (duplo clique na linha) com botão de cópia
  formatada para WhatsApp/E-mail.
- Indicadores dinâmicos de ordenação (▲ / ▼ / ↕) nas colunas da tabela.
- Carregamento inteligente de planilhas já auditadas (ex: resultados_frota_ce.xlsx)
  para visualização imediata do BI e métricas.
========================================================================================
"""

import os
import re
import sys
import math
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
    obter_ultimo_dia_util_mes,
    PADRAO_PLACA,
    gerar_nome_base_resultado,
    obter_caminhos_arquivos_saida
)
from core.scraper_engine import FrotaScraperEngine
from core.boleto_downloader import FrotaBoletoDownloader
from core.relatorio_exporter import FrotaRelatorioExporter, REPORTLAB_DISPONIVEL

logger = obter_logger("GUI")


# ======================================================================================
# SUPRESSÃO DE CONSOLE, HIGH-DPI E BARRA DE TÍTULO ESCURA NO WINDOWS (DWM)
# ======================================================================================

def ocultar_console():
    """Oculta a janela preta de console do Python no Windows."""
    if sys.platform == "win32" and "--console" not in sys.argv:
        try:
            import ctypes
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 0)
        except Exception:
            pass


def configurar_dpi():
    """Habilita suporte a High-DPI no Windows para nitidez máxima."""
    if sys.platform == "win32":
        try:
            import ctypes
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def aplicar_barra_titulo_windows(janela: tk.Misc, modo_noturno: bool):
    """Sincroniza a cor da barra de título nativa do Windows 10/11 com o tema da aplicação."""
    if sys.platform == "win32":
        try:
            import ctypes
            janela.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(janela.winfo_id())
            if not hwnd:
                hwnd = janela.winfo_id()
            valor = ctypes.c_int(1 if modo_noturno else 0)
            # DWMWA_USE_IMMERSIVE_DARK_MODE = 20 (Windows 11 / Win10 20H1+) ou 19
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 20, ctypes.byref(valor), ctypes.sizeof(valor)
            )
        except Exception:
            pass


# ======================================================================================
# UTILITÁRIOS GRÁFICOS: DESENHO DE GEOMETRIAS ARREDONDADAS SUAVES (BEZIER)
# ======================================================================================

def desenhar_retangulo_arredondado(
    canvas: tk.Canvas, x1: float, y1: float, x2: float, y2: float, radius: int = 10, **kwargs
):
    """Gera um polígono com bordas curvas suaves via algoritmo Bezier nativo do Tkinter."""
    r = min(radius, max(2, int((x2 - x1) / 2)), max(2, int((y2 - y1) / 2)))
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
# COMPONENTE VISUAL: TOGGLE SWITCH EM PÍLULA (ESTILO iOS / SAAS)
# ======================================================================================

class ModernToggleSwitch(tk.Canvas):
    """
    Interruptor deslizante em formato de pílula vinculado a um tk.BooleanVar.
    Substitui os checkboxes clássicos por uma estética comercial moderna.
    """

    def __init__(
        self,
        parent: tk.Widget,
        texto: str,
        variable: tk.BooleanVar,
        cor_ativo: str = "#2563EB",
        modo_noturno: bool = False,
        command=None
    ):
        super().__init__(parent, highlightthickness=0, cursor="hand2", height=26, width=165)
        self.texto = texto
        self.variable = variable
        self.cor_ativo = cor_ativo
        self.modo_noturno = modo_noturno
        self.command = command
        self.is_hover = False

        self._definir_cores()
        self.bind("<Button-1>", self._alternar)
        self.bind("<Enter>", self._ao_entrar)
        self.bind("<Leave>", self._ao_sair)
        self.bind("<Configure>", lambda e: self.redesenhar())
        self.variable.trace_add("write", lambda *_: self.redesenhar())

    def _definir_cores(self):
        if self.modo_noturno:
            self.bg_parent = "#0F172A"
            self.bg_off = "#334155"
            self.fg_text = "#E2E8F0"
            self.knob_color = "#F8FAFC"
        else:
            self.bg_parent = "#FFFFFF"
            self.bg_off = "#CBD5E1"
            self.fg_text = "#1E293B"
            self.knob_color = "#FFFFFF"
        self.configure(bg=self.bg_parent)

    def atualizar_tema(self, modo_noturno: bool, bg_override: str | None = None):
        self.modo_noturno = modo_noturno
        self._definir_cores()
        if bg_override:
            self.bg_parent = bg_override
            self.configure(bg=bg_override)
        self.redesenhar()

    def _ao_entrar(self, _):
        self.is_hover = True
        self.redesenhar()

    def _ao_sair(self, _):
        self.is_hover = False
        self.redesenhar()

    def _alternar(self, _=None):
        self.variable.set(not self.variable.get())
        if self.command:
            self.command()

    def redesenhar(self):
        self.delete("all")
        ativo = bool(self.variable.get())
        sw_w, sw_h = 34, 18
        x1, y1 = 4, 4
        x2, y2 = x1 + sw_w, y1 + sw_h

        cor_fundo = self.cor_ativo if ativo else self.bg_off
        cor_borda = self.cor_ativo if (ativo or self.is_hover) else self.bg_off

        desenhar_retangulo_arredondado(
            self, x1, y1, x2, y2, radius=9,
            fill=cor_fundo, outline=cor_borda, width=1
        )

        # Círculo interno deslizante (Knob)
        r_knob = 6
        cx = (x2 - 9) if ativo else (x1 + 9)
        cy = (y1 + y2) / 2
        self.create_oval(
            cx - r_knob, cy - r_knob, cx + r_knob, cy + r_knob,
            fill=self.knob_color, outline=""
        )

        # Rótulo de texto com largura auto-ajustável
        id_txt = self.create_text(
            x2 + 8, cy,
            text=self.texto,
            anchor="w",
            fill=self.fg_text,
            font=("Segoe UI", 9, "bold" if ativo else "normal")
        )
        bbox = self.bbox(id_txt)
        if bbox:
            largura_necessaria = bbox[2] + 8
            if abs(self.winfo_reqwidth() - largura_necessaria) > 4:
                self.configure(width=largura_necessaria)


# ======================================================================================
# COMPONENTE VISUAL: BARRA DE PROGRESSO COMERCIAL COM TELEMETRIA E ETA
# ======================================================================================

class ModernProgressBar(tk.Canvas):
    """
    Barra de progresso estilizada em Canvas com cantos arredondados,
    badge percentual integrado e exibição de telemetria (velocidade + ETA).
    """

    def __init__(self, parent: tk.Widget, modo_noturno: bool = False):
        super().__init__(parent, height=18, highlightthickness=0)
        self.modo_noturno = modo_noturno
        self._valor = 0.0
        self._texto_eta = ""
        self._definir_cores()
        self.bind("<Configure>", lambda e: self.redesenhar())

    def _definir_cores(self):
        if self.modo_noturno:
            self.bg_root = "#080C14"
            self.bg_trough = "#1E293B"
            self.border_trough = "#334155"
            self.fill_bar = "#2563EB"
            self.fill_highlight = "#38BDF8"
            self.fg_pct = "#F8FAFC"
        else:
            self.bg_root = "#F1F5F9"
            self.bg_trough = "#E2E8F0"
            self.border_trough = "#CBD5E1"
            self.fill_bar = "#2563EB"
            self.fill_highlight = "#3B82F6"
            self.fg_pct = "#0F172A"
        self.configure(bg=self.bg_root)

    def atualizar_tema(self, modo_noturno: bool):
        self.modo_noturno = modo_noturno
        self._definir_cores()
        self.redesenhar()

    def definir_telemetria(self, texto_eta: str):
        self._texto_eta = texto_eta
        self.redesenhar()

    def __setitem__(self, key: str, value):
        if key == "value":
            self._valor = max(0.0, min(100.0, float(value)))
            if self._valor == 0 or self._valor >= 100:
                self._texto_eta = ""
            self.redesenhar()
        else:
            super().__setitem__(key, value)

    def __getitem__(self, key: str):
        if key == "value":
            return self._valor
        return super().__getitem__(key)

    def redesenhar(self):
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 40 or h < 10:
            return

        self.delete("all")
        x1, y1, x2, y2 = 2, 2, w - 2, h - 2

        # Calha de fundo
        desenhar_retangulo_arredondado(
            self, x1, y1, x2, y2, radius=7,
            fill=self.bg_trough, outline=self.border_trough, width=1
        )

        # Preenchimento proporcional
        if self._valor > 0:
            larg_util = (x2 - x1) * (self._valor / 100.0)
            x_fim = max(x1 + 12, x1 + larg_util)
            cor_barra = "#10B981" if self._valor >= 100 else self.fill_bar
            desenhar_retangulo_arredondado(
                self, x1, y1, min(x2, x_fim), y2, radius=7,
                fill=cor_barra, outline="", width=0
            )
            # Faixa superior de brilho sutil
            if x_fim - x1 > 20:
                self.create_line(
                    x1 + 8, y1 + 3, min(x2 - 8, x_fim - 8), y1 + 3,
                    fill="#60A5FA" if self._valor < 100 else "#34D399", width=1
                )

        # Texto central de porcentagem + telemetria
        pct_str = f"{int(self._valor)}%"
        if self._texto_eta:
            pct_str += f"  •  {self._texto_eta}"
        cor_txt = "#FFFFFF" if self._valor > 48 else self.fg_pct
        self.create_text(
            w / 2, h / 2,
            text=pct_str,
            fill=cor_txt,
            font=("Segoe UI", 8, "bold")
        )


# ======================================================================================
# COMPONENTE VISUAL: LUXURY METRIC CARD (DASHBOARD FINTECH COM BARRA DE PROPORÇÃO)
# ======================================================================================

class LuxuryMetricCard(tk.Canvas):
    """
    Card de Métrica Executiva renderizado em Canvas de alta precisão.
    Inclui badge circular de ícone, barra de proporção financeira na base e Hover Glow.
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
        super().__init__(parent, highlightthickness=0, height=84, width=110)
        self.titulo = titulo
        self.var_valor = var_valor
        self.icone = icone
        self.cor_acento = cor_acento
        self.subtexto = subtexto
        self.modo_noturno = modo_noturno
        self.proporcao = 0.0  # 0.0 a 1.0 para mini-barra de participação

        self.is_hover = False
        self._definir_paleta()

        self.bind("<Configure>", self._ao_redimensionar)
        self.bind("<Enter>", self._ao_entrar)
        self.bind("<Leave>", self._ao_sair)

        self.var_valor.trace_add("write", lambda *_: self.redesenhar())

    def _definir_paleta(self):
        if self.modo_noturno:
            self.bg_parent = "#080C14"
            self.bg_card = "#0F172A"
            self.bg_badge = "#1E293B"
            self.bg_bar_track = "#1E293B"
            self.border_normal = "#1E293B"
            self.border_glow = self.cor_acento
            self.fg_title = "#94A3B8"
            self.fg_sub = "#64748B"
            self.fg_val_default = "#F8FAFC"
        else:
            self.bg_parent = "#F1F5F9"
            self.bg_card = "#FFFFFF"
            self.bg_badge = "#F1F5F9"
            self.bg_bar_track = "#E2E8F0"
            self.border_normal = "#CBD5E1"
            self.border_glow = self.cor_acento
            self.fg_title = "#475569"
            self.fg_sub = "#64748B"
            self.fg_val_default = "#0F172A"

        self.configure(bg=self.bg_parent)

    def definir_proporcao(self, frac: float, subtexto_novo: str | None = None):
        self.proporcao = max(0.0, min(1.0, float(frac)))
        if subtexto_novo is not None:
            self.subtexto = subtexto_novo
        self.redesenhar()

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
        x1, y1, x2, y2 = 3, 2, w - 4, h - 3
        cor_borda = self.border_glow if self.is_hover else self.border_normal
        largura_borda = 2 if self.is_hover else 1

        # 1. Fundo arredondado
        desenhar_retangulo_arredondado(
            self, x1, y1, x2, y2, radius=11,
            fill=self.bg_card, outline=cor_borda, width=largura_borda
        )

        # 2. Barra de acento superior
        self.create_rectangle(
            x1 + 14, y1 + 1, x2 - 14, y1 + 3,
            fill=self.cor_acento, outline=""
        )

        # 3. Badge circular do ícone à direita
        tem_badge = w >= 145
        limite_dir_topo = (x2 - 38) if tem_badge else (x2 - 10)
        if tem_badge:
            bx = x2 - 22
            by = y1 + 20
            self.create_oval(bx - 11, by - 11, bx + 11, by + 11, fill=self.bg_badge, outline="", tags=("badge_bg",))
            self.create_text(bx, by, text=self.icone, font=("Segoe UI Emoji", 8), tags=("badge_icon",))

        # 4. Título superior (com ajuste automático para nunca tocar o badge)
        id_tit = self.create_text(
            x1 + 12, y1 + 15,
            text=self.titulo,
            anchor="w",
            fill=self.fg_title,
            font=("Segoe UI", 8, "bold")
        )
        b_tit = self.bbox(id_tit)
        if b_tit and b_tit[2] > limite_dir_topo:
            txt_t = self.titulo
            while len(txt_t) > 4 and self.bbox(id_tit)[2] > limite_dir_topo:
                txt_t = txt_t[:-1].rstrip()
                self.itemconfigure(id_tit, text=txt_t + "…")

        # 5. Valor principal (com ajuste dinâmico de fonte se necessário)
        val = self.var_valor.get() or "R$ 0,00"
        cor_val = self.cor_acento if ("TOTAL" in self.titulo or "R$" in val) else self.fg_val_default
        tam_fonte = 13 if w >= 175 else (11 if w >= 140 else 10)
        id_val = self.create_text(
            x1 + 12, y1 + 39,
            text=val,
            anchor="w",
            fill=cor_val,
            font=("Segoe UI", tam_fonte, "bold")
        )
        b_val = self.bbox(id_val)
        while b_val and b_val[2] > (x2 - 10) and tam_fonte > 8:
            tam_fonte -= 1
            self.itemconfigure(id_val, font=("Segoe UI", tam_fonte, "bold"))
            b_val = self.bbox(id_val)

        # 6. Subtexto + Mini-barra de proporção
        id_sub = self.create_text(
            x1 + 12, y1 + 59,
            text=self.subtexto,
            anchor="w",
            fill=self.fg_sub,
            font=("Segoe UI", 7)
        )
        b_sub = self.bbox(id_sub)
        if b_sub and b_sub[2] > (x2 - 10):
            txt_s = self.subtexto
            while len(txt_s) > 4 and self.bbox(id_sub)[2] > (x2 - 10):
                txt_s = txt_s[:-1].rstrip()
                self.itemconfigure(id_sub, text=txt_s + "…")

        bar_y = y2 - 7
        bar_x1 = x1 + 12
        bar_x2 = x2 - 12
        if bar_x2 > bar_x1 + 10:
            self.create_line(bar_x1, bar_y, bar_x2, bar_y, fill=self.bg_bar_track, width=3)
            if self.proporcao > 0:
                fill_x2 = bar_x1 + max(4, (bar_x2 - bar_x1) * self.proporcao)
                self.create_line(bar_x1, bar_y, fill_x2, bar_y, fill=self.cor_acento, width=3)


# ======================================================================================
# COMPONENTE VISUAL: LIVE STATUS ORB COM ANIMAÇÃO PULSANTE
# ======================================================================================

class LiveStatusOrb(tk.Canvas):
    """
    Badge indicador de telemetria com halo luminoso pulsante em tempo real.
    """

    def __init__(self, parent: tk.Widget, modo_noturno: bool = False):
        super().__init__(parent, width=230, height=32, highlightthickness=0)
        self.modo_noturno = modo_noturno
        self.status_tipo = "pronto"
        self.texto_status = "SISTEMA ONLINE & PRONTO"
        self._pulso_fase = 0
        self._after_id = None
        self._atualizar_cores()

    def _atualizar_cores(self):
        self.bg_canvas = "#080C14" if self.modo_noturno else "#F1F5F9"
        self.bg_pill = "#0F172A" if self.modo_noturno else "#FFFFFF"
        self.border_pill = "#1E293B" if self.modo_noturno else "#CBD5E1"
        self.fg_text = "#F8FAFC" if self.modo_noturno else "#0F172A"
        self.configure(bg=self.bg_canvas)
        self.redesenhar()

    def atualizar_tema(self, modo_noturno: bool):
        self.modo_noturno = modo_noturno
        self._atualizar_cores()

    def definir_status(self, tipo: str, texto: str):
        self.status_tipo = tipo
        self.texto_status = texto
        if tipo == "processando":
            self._iniciar_pulso()
        else:
            self._parar_pulso()
        self.redesenhar()

    def _iniciar_pulso(self):
        if self._after_id is None:
            self._animar_pulso()

    def _parar_pulso(self):
        if self._after_id is not None:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        self._pulso_fase = 0

    def _animar_pulso(self):
        if not self.winfo_exists() or self.status_tipo != "processando":
            self._after_id = None
            return
        self._pulso_fase = (self._pulso_fase + 1) % 8
        self.redesenhar()
        self._after_id = self.after(180, self._animar_pulso)

    def redesenhar(self):
        self.delete("all")
        w, h = 226, 28
        desenhar_retangulo_arredondado(
            self, 2, 2, w, h, radius=9,
            fill=self.bg_pill, outline=self.border_pill, width=1
        )

        cores_orbe = {
            "pronto": ("#10B981", "#34D399"),
            "processando": ("#2563EB", "#93C5FD"),
            "pausado": ("#F59E0B", "#FCD34D"),
            "erro": ("#EF4444", "#FCA5A5")
        }
        cor_base, cor_brilho = cores_orbe.get(self.status_tipo, ("#10B981", "#34D399"))

        # Halo pulsante dinâmico
        expansao = (1.5 * math.sin(self._pulso_fase * math.pi / 4)) if self.status_tipo == "processando" else 0
        r_halo = 6 + expansao
        cx, cy = 16, 15
        self.create_oval(cx - r_halo, cy - r_halo, cx + r_halo, cy + r_halo, fill=cor_brilho, outline="")
        self.create_oval(cx - 4, cy - 4, cx + 4, cy + 4, fill=cor_base, outline="")

        self.create_text(
            29, 15,
            text=self.texto_status,
            anchor="w",
            fill=self.fg_text,
            font=("Segoe UI", 8, "bold")
        )


# ======================================================================================
# COMPONENTE VISUAL: PAINEL ANALÍTICO DE BI (GRÁFICO DONUT, SAÚDE DA FROTA & TOP 10)
# ======================================================================================

class PainelAnaliticoBI(tk.Canvas):
    """
    Dashboard Analítico nativo em Canvas para visualização executiva:
    - Coluna 1: Gráfico de Rosca (Donut Chart) de Composição Financeira (Licenciamento, IPVA, Multas).
    - Coluna 2: Medidor de Saúde da Frota (Regulares, Com Débitos, Restrições, Falhas).
    - Coluna 3: Ranking Top 10 Veículos / Proprietários de Maior Impacto Financeiro.
    """

    def __init__(self, parent: tk.Widget, modo_noturno: bool = False):
        super().__init__(parent, highlightthickness=0)
        self.modo_noturno = modo_noturno
        self.dados_bi = {
            "total_veiculos": 0,
            "soma_lic": 0.0,
            "soma_ipva": 0.0,
            "soma_multas": 0.0,
            "soma_total": 0.0,
            "qtd_pagos": 0,
            "qtd_debitos": 0,
            "qtd_restricao": 0,
            "qtd_falhas": 0,
            "qtd_boletos": 0,
            "top_veiculos": []  # list of (placa, proprietario, valor_float, status)
        }
        self._definir_cores()
        self.bind("<Configure>", lambda e: self.redesenhar())

    def _definir_cores(self):
        if self.modo_noturno:
            self.bg_canvas = "#080C14"
            self.bg_card = "#0F172A"
            self.border_card = "#1E293B"
            self.fg_title = "#F8FAFC"
            self.fg_sub = "#94A3B8"
            self.fg_muted = "#64748B"
            self.bg_bar = "#1E293B"
        else:
            self.bg_canvas = "#F8FAFC"
            self.bg_card = "#FFFFFF"
            self.border_card = "#E2E8F0"
            self.fg_title = "#0F172A"
            self.fg_sub = "#475569"
            self.fg_muted = "#64748B"
            self.bg_bar = "#F1F5F9"
        self.configure(bg=self.bg_canvas)

    def atualizar_tema(self, modo_noturno: bool):
        self.modo_noturno = modo_noturno
        self._definir_cores()
        self.redesenhar()

    def atualizar_dados(self, dados_bi: dict):
        self.dados_bi.update(dados_bi)
        self.redesenhar()

    def _ajustar_texto_largura(self, item_id: int, texto_completo: str, largura_max: float) -> str:
        """Trunca dinamicamente um texto no Canvas caso seu bbox ultrapasse largura_max."""
        self.itemconfigure(item_id, text=texto_completo)
        bbox = self.bbox(item_id)
        if not bbox or (bbox[2] - bbox[0]) <= largura_max:
            return texto_completo
        txt = texto_completo
        while len(txt) > 3:
            txt = txt[:-1].rstrip()
            tentativa = txt + "…"
            self.itemconfigure(item_id, text=tentativa)
            bbox = self.bbox(item_id)
            if bbox and (bbox[2] - bbox[0]) <= largura_max:
                return tentativa
        return txt

    def redesenhar(self):
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 100 or h < 80:
            return

        self.delete("all")
        pad = 8
        col_w = (w - pad * 4) / 3.0
        y1, y2 = pad, h - pad

        # Card 1: Composição dos Débitos (Donut Chart)
        x1_c1, x2_c1 = pad, pad + col_w
        desenhar_retangulo_arredondado(
            self, x1_c1, y1, x2_c1, y2, radius=10,
            fill=self.bg_card, outline=self.border_card, width=1
        )
        self._desenhar_donut_card(x1_c1, y1, x2_c1, y2)

        # Card 2: Índice de Regularidade & Saúde da Frota
        x1_c2, x2_c2 = x2_c1 + pad, x2_c1 + pad + col_w
        desenhar_retangulo_arredondado(
            self, x1_c2, y1, x2_c2, y2, radius=10,
            fill=self.bg_card, outline=self.border_card, width=1
        )
        self._desenhar_saude_card(x1_c2, y1, x2_c2, y2)

        # Card 3: Top 10 Veículos de Maior Impacto
        x1_c3, x2_c3 = x2_c2 + pad, w - pad
        desenhar_retangulo_arredondado(
            self, x1_c3, y1, x2_c3, y2, radius=10,
            fill=self.bg_card, outline=self.border_card, width=1
        )
        self._desenhar_top10_card(x1_c3, y1, x2_c3, y2)

    def _desenhar_donut_card(self, x1: float, y1: float, x2: float, y2: float):
        id_tit = self.create_text(
            x1 + 14, y1 + 18,
            text="🍩 Composição Financeira por Tributo",
            anchor="w", fill=self.fg_title, font=("Segoe UI", 9, "bold")
        )
        self._ajustar_texto_largura(id_tit, "🍩 Composição Financeira por Tributo", (x2 - x1) - 28)

        lic = self.dados_bi["soma_lic"]
        ipva = self.dados_bi["soma_ipva"]
        multas = self.dados_bi["soma_multas"]
        soma_total = self.dados_bi["soma_total"]
        soma_bruta = lic + ipva + multas

        card_w = x2 - x1
        h_disp = y2 - y1 - 38
        tem_espaco_resumo = h_disp >= 280
        zona_donut_h = int(h_disp * 0.56) if tem_espaco_resumo else h_disp

        # Garante que o raio + legenda lateral caibam folgados dentro de card_w
        larg_min_legenda = min(145, max(105, int(card_w * 0.44)))
        max_raio_w = max(32, int((card_w - 32 - larg_min_legenda - 14) / 2))
        max_raio_h = max(32, int(zona_donut_h * 0.42))
        raio = min(70, max_raio_w, max_raio_h)

        cx = x1 + 18 + raio
        cy = y1 + 36 + zona_donut_h / 2

        fatias = [
            ("Licenciamento", lic, "#F59E0B"),
            ("IPVA (Sefaz)", ipva, "#F43F5E"),
            ("Multas Detran", multas, "#EA580C")
        ]

        if soma_bruta <= 0:
            self.create_oval(
                cx - raio, cy - raio, cx + raio, cy + raio,
                fill=self.bg_bar, outline=self.border_card, width=2
            )
        else:
            ang_inicio = 90.0
            for _, val, cor in fatias:
                if val <= 0:
                    continue
                ext = -359.9 if val >= soma_bruta else -(val / soma_bruta) * 360.0
                self.create_arc(
                    cx - raio, cy - raio, cx + raio, cy + raio,
                    start=ang_inicio, extent=ext, fill=cor, outline=self.bg_card, width=2
                )
                ang_inicio += ext

        # Furo central do Donut com proteção geométrica contra toque no anel interno
        r_interno = max(24, int(raio * 0.68))
        self.create_oval(
            cx - r_interno, cy - r_interno, cx + r_interno, cy + r_interno,
            fill=self.bg_card, outline="", tags=("donut_hole",)
        )
        self.create_text(
            cx, cy - 8,
            text="A Liquidar",
            fill=self.fg_muted, font=("Segoe UI", 7),
            tags=("donut_center_lbl",)
        )
        val_centro_str = formatar_float_moeda(soma_total).replace("R$ ", "")
        id_val_c = self.create_text(
            cx, cy + 8,
            text=val_centro_str,
            fill="#10B981", font=("Segoe UI", 8, "bold"),
            tags=("donut_center_val",)
        )
        larg_corda_interna = max(28.0, 2.0 * math.sqrt(max(1.0, (r_interno - 3) ** 2 - 12 ** 2)))
        bbox_c = self.bbox(id_val_c)
        if bbox_c and (bbox_c[2] - bbox_c[0]) > larg_corda_interna:
            self.itemconfigure(id_val_c, font=("Segoe UI", 7, "bold"))
            bbox_c = self.bbox(id_val_c)
            if bbox_c and (bbox_c[2] - bbox_c[0]) > larg_corda_interna:
                self._ajustar_texto_largura(id_val_c, val_centro_str, larg_corda_interna)

        # Legenda lateral com espaçamento vertical seguro (16px entre título e valor)
        lx = cx + raio + 16
        larg_max_leg_txt = max(50, (x2 - 12) - (lx + 14))
        passo_leg = max(34, min(42, int(zona_donut_h / 3.2)))
        ly = cy - passo_leg

        for nome, val, cor in fatias:
            pct = (val / soma_bruta * 100.0) if soma_bruta > 0 else 0.0
            self.create_oval(lx, ly - 4, lx + 8, ly + 4, fill=cor, outline="")
            txt_leg = f"{nome} ({pct:.0f}%)"
            id_lt = self.create_text(
                lx + 14, ly - 7,
                text=txt_leg,
                anchor="w", fill=self.fg_title, font=("Segoe UI", 8, "bold")
            )
            self._ajustar_texto_largura(id_lt, txt_leg, larg_max_leg_txt)

            txt_lv = formatar_float_moeda(val)
            id_lv = self.create_text(
                lx + 14, ly + 9,
                text=txt_lv,
                anchor="w", fill=self.fg_sub, font=("Segoe UI", 8)
            )
            self._ajustar_texto_largura(id_lv, txt_lv, larg_max_leg_txt)
            ly += passo_leg

        # Se houver espaço vertical, exibe o Resumo Executivo de Auditoria abaixo do Donut
        y_base_donut = max(cy + raio + 16, (ly - passo_leg) + 24)
        if tem_espaco_resumo and (y2 - y_base_donut) >= 65:
            self.create_line(x1 + 14, y_base_donut, x2 - 14, y_base_donut, fill=self.border_card, width=1)
            ry = y_base_donut + 16
            abatimento = max(0.0, soma_bruta - soma_total)
            qtd_dev = max(1, self.dados_bi.get("qtd_debitos", 0))
            ticket_dev = (soma_total / qtd_dev) if self.dados_bi.get("qtd_debitos", 0) > 0 else 0.0

            linhas_resumo = [
                ("Soma Bruta (Órgãos):", formatar_float_moeda(soma_bruta), self.fg_sub),
                ("Multas Inclusas no Lic. (Abatidas):", f"- {formatar_float_moeda(abatimento)}", "#10B981" if abatimento > 0 else self.fg_muted),
                ("Total Líquido (Sem Duplicidade):", formatar_float_moeda(soma_total), "#EF4444" if soma_total > 0 else "#10B981"),
                ("Ticket Médio por Devedor:", formatar_float_moeda(ticket_dev), self.fg_title),
            ]
            passo_r = max(20, min(26, int((y2 - ry - 8) / max(1, len(linhas_resumo)))))
            for rotulo, val_s, cor_v in linhas_resumo:
                if ry + 8 > y2 - 8:
                    break
                id_rv = self.create_text(
                    x2 - 14, ry,
                    text=val_s, anchor="e", fill=cor_v, font=("Segoe UI", 8, "bold")
                )
                bbox_rv = self.bbox(id_rv)
                lim_rot = (bbox_rv[0] - (x1 + 14) - 10) if bbox_rv else ((x2 - x1) * 0.55)
                id_rr = self.create_text(
                    x1 + 14, ry,
                    text=rotulo, anchor="w", fill=self.fg_sub, font=("Segoe UI", 8)
                )
                self._ajustar_texto_largura(id_rr, rotulo, max(40, lim_rot))
                ry += passo_r

    def _desenhar_saude_card(self, x1: float, y1: float, x2: float, y2: float):
        id_tit = self.create_text(
            x1 + 14, y1 + 18,
            text="🛡️ Índice de Regularidade da Frota",
            anchor="w", fill=self.fg_title, font=("Segoe UI", 9, "bold")
        )
        self._ajustar_texto_largura(id_tit, "🛡️ Índice de Regularidade da Frota", (x2 - x1) - 28)

        tot = max(1, self.dados_bi["total_veiculos"])
        pagos = self.dados_bi["qtd_pagos"]
        debitos = self.dados_bi["qtd_debitos"]
        restr = self.dados_bi["qtd_restricao"]
        falhas = self.dados_bi["qtd_falhas"]
        auditados = pagos + debitos + falhas

        pct_reg = (pagos / auditados * 100.0) if auditados > 0 else 0.0
        pct_deb = (debitos / auditados * 100.0) if auditados > 0 else 0.0

        # Barra horizontal empilhada
        bx1, bx2 = x1 + 14, x2 - 14
        by1, by2 = y1 + 40, y1 + 56
        larg_total = max(10, bx2 - bx1)

        self.create_rectangle(bx1, by1, bx2, by2, fill=self.bg_bar, outline="", tags=("bar_track",))

        if auditados > 0:
            denom = max(tot, auditados, 1)
            x_cur = bx1
            for qtd, cor in [(pagos, "#10B981"), (debitos, "#EF4444"), (falhas, "#F59E0B")]:
                if qtd > 0 and x_cur < bx2:
                    w_seg = (qtd / denom) * larg_total
                    x_fim_seg = min(bx2, x_cur + w_seg)
                    self.create_rectangle(x_cur, by1, x_fim_seg, by2, fill=cor, outline="", tags=("bar_seg",))
                    x_cur = x_fim_seg

        # Indicadores em grade 2x2 com espaçamento vertical de 18px entre rótulo e valor
        metricas = [
            ("✅ Regulares / Quitados", f"{pagos} veíc. ({pct_reg:.0f}%)", "#10B981"),
            ("⚠️ Com Débitos Ativos", f"{debitos} veíc. ({pct_deb:.0f}%)", "#EF4444"),
            ("🚨 Com Restrição", f"{restr} veíc.", "#F97316"),
            ("❌ Falhas / Pendentes", f"{falhas} veíc.", "#F59E0B")
        ]

        gy = y1 + 76
        h_restante = y2 - gy - 10
        passo_row = max(42, min(54, int(h_restante * 0.28)))
        meio_x = (x1 + x2) / 2
        larg_col = max(50, (meio_x - x1) - 22)

        for idx, (rotulo, valor_str, cor) in enumerate(metricas):
            col = idx % 2
            row = idx // 2
            mx = (x1 + 14) if col == 0 else (meio_x + 6)
            my = gy + row * passo_row
            if my + 24 > y2 - 6:
                break
            id_r = self.create_text(mx, my, text=rotulo, anchor="w", fill=self.fg_sub, font=("Segoe UI", 8))
            self._ajustar_texto_largura(id_r, rotulo, larg_col)

            id_v = self.create_text(mx, my + 18, text=valor_str, anchor="w", fill=cor, font=("Segoe UI", 10, "bold"))
            self._ajustar_texto_largura(id_v, valor_str, larg_col)

        # Se houver espaço inferior, exibe Indicadores Executivos Complementares da Frota
        y_pos_grid = gy + 2 * passo_row + 4
        if (y2 - y_pos_grid) >= 65:
            self.create_line(x1 + 14, y_pos_grid, x2 - 14, y_pos_grid, fill=self.border_card, width=1)
            ry = y_pos_grid + 16
            soma_total = self.dados_bi.get("soma_total", 0.0)
            ticket_frota = (soma_total / max(1, auditados)) if auditados > 0 else 0.0
            top_v = self.dados_bi.get("top_veiculos", [])
            maior_txt = f"{top_v[0][0]} ({formatar_float_moeda(top_v[0][2])})" if top_v else "Nenhum"
            qtd_bol = self.dados_bi.get("qtd_boletos", debitos)

            extras = [
                ("🚗 Veículos Auditados na Sessão:", f"{auditados} de {self.dados_bi['total_veiculos']}", self.fg_title),
                ("📥 Elegíveis p/ Emissão de Boleto:", f"{qtd_bol} veíc.", "#3B82F6" if qtd_bol > 0 else self.fg_muted),
                ("💳 Custo Médio por Veículo da Frota:", formatar_float_moeda(ticket_frota), self.fg_title),
                ("🔥 Maior Passivo Individual:", maior_txt, "#EF4444" if top_v else self.fg_muted),
            ]
            passo_e = max(20, min(26, int((y2 - ry - 8) / max(1, len(extras)))))
            for rotulo, val_s, cor_v in extras:
                if ry + 8 > y2 - 8:
                    break
                id_ev = self.create_text(
                    x2 - 14, ry,
                    text=val_s, anchor="e", fill=cor_v, font=("Segoe UI", 8, "bold")
                )
                bbox_ev = self.bbox(id_ev)
                lim_rot = (bbox_ev[0] - (x1 + 14) - 10) if bbox_ev else ((x2 - x1) * 0.55)
                id_er = self.create_text(
                    x1 + 14, ry,
                    text=rotulo, anchor="w", fill=self.fg_sub, font=("Segoe UI", 8)
                )
                self._ajustar_texto_largura(id_er, rotulo, max(40, lim_rot))
                ry += passo_e

    def _desenhar_top10_card(self, x1: float, y1: float, x2: float, y2: float):
        id_tit = self.create_text(
            x1 + 14, y1 + 18,
            text="🏆 Top 10 Veículos de Maior Impacto",
            anchor="w", fill=self.fg_title, font=("Segoe UI", 9, "bold")
        )
        self._ajustar_texto_largura(id_tit, "🏆 Top 10 Veículos de Maior Impacto", (x2 - x1) - 28)

        top_lista = self.dados_bi.get("top_veiculos", [])[:10]
        if not top_lista:
            self.create_text(
                (x1 + x2) / 2, (y1 + y2) / 2,
                text="Nenhum débito pendente registrado",
                fill=self.fg_muted, font=("Segoe UI", 8, "italic")
            )
            return

        maior_val = max((item[2] for item in top_lista), default=1.0) or 1.0
        ry_inicio = y1 + 36
        h_disp = max(60.0, float((y2 - 8) - ry_inicio))

        # Garante que os 10 itens caibam mesmo em resoluções compactas (ex: 1024x600)
        max_linhas_visiveis = max(1, min(len(top_lista), int(h_disp // 14)))
        itens_visiveis = top_lista[:max_linhas_visiveis]
        n_slots = max(len(itens_visiveis), min(10, int(h_disp // 22)))
        passo_f = min(38.0, h_disp / max(1, n_slots))
        tam_fonte = 7 if passo_f < 20.0 else 8
        meia_barra = 3 if passo_f < 20.0 else 4

        card_w = x2 - x1
        linhas_renderizadas = []

        # Passo 1: Criar os textos de valor à direita (Coluna 3) e medir seu limite esquerdo exato
        for idx, (placa, prop, val_f, _) in enumerate(itens_visiveis):
            rank = idx + 1
            ry = int(round(ry_inicio + (idx + 0.5) * passo_f))
            val_str = formatar_float_moeda(val_f)
            id_val = self.create_text(
                x2 - 14, ry,
                text=val_str,
                anchor="e", fill=self.fg_sub, font=("Segoe UI", tam_fonte, "bold"),
                tags=("top10_val",)
            )
            linhas_renderizadas.append({
                "rank": rank,
                "placa": placa,
                "prop": prop,
                "val_f": val_f,
                "ry": ry,
                "id_val": id_val,
                "id_lbl": None
            })

        bboxes_val = [self.bbox(item["id_val"]) for item in linhas_renderizadas]
        min_val_x1 = min((b[0] for b in bboxes_val if b), default=x2 - 85)

        # Passo 2: Reservar largura mínima garantida para a barra proporcional (Coluna 2)
        larg_min_barra = max(55, int(card_w * 0.24))
        margem_seguranca = 12
        bx_max = min_val_x1 - margem_seguranca
        limite_max_lbl_x2 = min(
            x1 + 14 + int(card_w * 0.48),
            bx_max - larg_min_barra - margem_seguranca
        )
        larg_max_lbl = max(65, limite_max_lbl_x2 - (x1 + 14))

        # Passo 3: Criar os rótulos da esquerda (Coluna 1) truncando apenas se excederem limite_max_lbl_x2
        for item in linhas_renderizadas:
            rank = item["rank"]
            placa = item["placa"]
            prop = str(item["prop"] or "").strip()
            ry = item["ry"]

            prefixo = f"{rank:>2}. {placa}"
            texto_ideal = f"{prefixo} ({prop})" if prop else prefixo

            id_lbl = self.create_text(
                x1 + 14, ry,
                text=texto_ideal,
                anchor="w", fill=self.fg_title, font=("Consolas", tam_fonte, "bold"),
                tags=("top10_lbl",)
            )
            item["id_lbl"] = id_lbl

            bbox_l = self.bbox(id_lbl)
            if bbox_l and (bbox_l[2] - bbox_l[0]) > larg_max_lbl:
                if prop:
                    prop_limpo = prop
                    ajustou = False
                    while len(prop_limpo) > 1:
                        prop_limpo = prop_limpo[:-1].rstrip()
                        tentativa = f"{prefixo} ({prop_limpo}…)"
                        self.itemconfigure(id_lbl, text=tentativa)
                        b_try = self.bbox(id_lbl)
                        if b_try and (b_try[2] - b_try[0]) <= larg_max_lbl:
                            ajustou = True
                            break
                    if not ajustou:
                        self._ajustar_texto_largura(id_lbl, prefixo, larg_max_lbl)
                else:
                    self._ajustar_texto_largura(id_lbl, prefixo, larg_max_lbl)

        # Passo 4: Medir o limite direito real dos rótulos da Coluna 1 e posicionar as barras sem sobreposição
        bboxes_lbl = [self.bbox(item["id_lbl"]) for item in linhas_renderizadas]
        max_lbl_x2 = max((b[2] for b in bboxes_lbl if b), default=x1 + 110)

        bx_ini = max_lbl_x2 + margem_seguranca
        larg_barra_disp = max(16, bx_max - bx_ini)

        for item in linhas_renderizadas:
            rank = item["rank"]
            val_f = item["val_f"]
            ry = item["ry"]
            frac = min(1.0, max(0.0, val_f / maior_val))
            bx_fim = bx_ini + max(4, larg_barra_disp * frac)

            if rank == 1:
                cor_barra = "#EF4444"
            elif rank in (2, 3):
                cor_barra = "#F97316"
            else:
                cor_barra = "#3B82F6"

            # Calha de fundo sutil (0% a 100%)
            self.create_rectangle(
                bx_ini, ry - meia_barra, bx_max, ry + meia_barra,
                fill=self.bg_bar, outline="", tags=("bar_track",)
            )
            # Barra proporcional de valor
            self.create_rectangle(
                bx_ini, ry - meia_barra, min(bx_max, bx_fim), ry + meia_barra,
                fill=cor_barra, outline="", tags=("bar_fill",)
            )


# ======================================================================================
# JANELA MODAL: RAIO-X EXECUTIVO DO VEÍCULO (DUPLO CLIQUE NA TABELA)
# ======================================================================================

class JanelaRaioXVeiculo(tk.Toplevel):
    """
    Ficha executiva detalhada de um veículo individual com decomposição financeira
    e botão de cópia rápida formatada para WhatsApp / E-mail.
    """

    def __init__(
        self,
        parent: tk.Tk,
        placa: str,
        renavam: str,
        dados_extras: dict,
        dados_financeiros: dict,
        modo_noturno: bool = False,
        on_baixar_boletos=None
    ):
        super().__init__(parent)
        self.title(f"Raio-X Executivo do Veículo — {placa}")
        self.transient(parent)
        self.grab_set()
        self.modo_noturno = modo_noturno
        self.placa = placa
        self.renavam = renavam
        self.dados_extras = dados_extras
        self.dados_financeiros = dados_financeiros
        self.on_baixar_boletos = on_baixar_boletos

        aplicar_barra_titulo_windows(self, self.modo_noturno)

        larg, alt = 520, 410
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width() - larg) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - alt) // 2
        self.geometry(f"{larg}x{alt}+{max(20, px)}+{max(20, py)}")
        self.resizable(False, False)

        bg_root = "#080C14" if modo_noturno else "#F1F5F9"
        bg_card = "#0F172A" if modo_noturno else "#FFFFFF"
        fg_title = "#F8FAFC" if modo_noturno else "#0F172A"
        fg_sub = "#94A3B8" if modo_noturno else "#475569"
        border = "#1E293B" if modo_noturno else "#CBD5E1"
        self.configure(bg=bg_root)

        lic = dados_financeiros.get("lic", "-")
        ipva = dados_financeiros.get("ipva", "-")
        multas = dados_financeiros.get("multas", "-")
        m_av = dados_financeiros.get("multas_a_vencer", "R$ 0,00")
        m_venc = dados_financeiros.get("multas_vencidas", "R$ 0,00")
        total_f = calcular_total_a_pagar_veiculo(lic, ipva, multas, m_av)
        self.total_str = formatar_float_moeda(total_f)

        container = tk.Frame(self, bg=bg_root, padx=16, pady=14)
        container.pack(fill=tk.BOTH, expand=True)

        # Header Card
        header = tk.Frame(container, bg=bg_card, highlightthickness=1, highlightbackground=border, padx=14, pady=10)
        header.pack(fill=tk.X, pady=(0, 10))

        tk.Label(
            header, text=f"🚘  PLACA: {placa}",
            font=("Consolas", 14, "bold"), bg=bg_card, fg=fg_title
        ).pack(anchor="w")

        sub_info = f"RENAVAM: {renavam}"
        for k, v in dados_extras.items():
            if v and str(v).strip():
                sub_info += f"   •   {k}: {v}"
        tk.Label(header, text=sub_info, font=("Segoe UI", 9), bg=bg_card, fg=fg_sub).pack(anchor="w", pady=(2, 0))

        # Breakdown Card
        corpo = tk.Frame(container, bg=bg_card, highlightthickness=1, highlightbackground=border, padx=14, pady=12)
        corpo.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        tk.Label(
            corpo, text="Decomposição Auditada de Débitos",
            font=("Segoe UI", 9, "bold"), bg=bg_card, fg=fg_title
        ).pack(anchor="w", pady=(0, 8))

        linhas_detalhe = [
            ("📋 Licenciamento (Detran-CE):", lic),
            ("🏛️ IPVA em Aberto (Sefaz-CE):", ipva),
            ("🚨 Multas Totais (Detran-CE):", multas),
            ("   ↳ Multas no Prazo (A Vencer):", m_av),
            ("   ↳ Multas Vencidas (Inclusas no Lic.):", m_venc),
        ]

        for rotulo, valor in linhas_detalhe:
            f_row = tk.Frame(corpo, bg=bg_card)
            f_row.pack(fill=tk.X, pady=2)
            tk.Label(f_row, text=rotulo, font=("Segoe UI", 9), bg=bg_card, fg=fg_sub).pack(side=tk.LEFT)
            cor_v = "#10B981" if str(valor).upper() == "PAGO" else fg_title
            tk.Label(f_row, text=str(valor), font=("Segoe UI", 9, "bold"), bg=bg_card, fg=cor_v).pack(side=tk.RIGHT)

        tk.Frame(corpo, bg=border, height=1).pack(fill=tk.X, pady=8)

        f_tot = tk.Frame(corpo, bg=bg_card)
        f_tot.pack(fill=tk.X)
        tk.Label(
            f_tot, text="💎 TOTAL LÍQUIDO A PAGAR (Sem Duplicidade):",
            font=("Segoe UI", 9, "bold"), bg=bg_card, fg=fg_title
        ).pack(side=tk.LEFT)
        tk.Label(
            f_tot, text=self.total_str,
            font=("Segoe UI", 12, "bold"), bg=bg_card, fg="#10B981" if total_f == 0 else "#EF4444"
        ).pack(side=tk.RIGHT)

        # Botões de Ação
        rodape = tk.Frame(container, bg=bg_root)
        rodape.pack(fill=tk.X)

        self.btn_copiar = ttk.Button(
            rodape, text="📋 Copiar Resumo (WhatsApp / E-mail)",
            style="Primary.TButton", command=self._copiar_resumo
        )
        self.btn_copiar.pack(side=tk.LEFT, padx=(0, 6))

        if total_f > 0 and on_baixar_boletos:
            ttk.Button(
                rodape, text="📥 Central de Boletos",
                style="PrimaryDownload.TButton",
                command=lambda: [self.destroy(), on_baixar_boletos()]
            ).pack(side=tk.LEFT)

        ttk.Button(rodape, text="Fechar", style="Muted.TButton", command=self.destroy).pack(side=tk.RIGHT)

    def _copiar_resumo(self):
        prop_txt = ""
        for k, v in self.dados_extras.items():
            if v and str(v).strip():
                prop_txt += f"\n👤 *{k}:* {v}"
        resumo = (
            f"🚗 *AUDITORIA VEICULAR - FROTA CE*\n"
            f"• *Placa:* `{self.placa}`\n"
            f"• *RENAVAM:* `{self.renavam}`{prop_txt}\n"
            f"──────────────────────\n"
            f"📋 *Licenciamento:* {self.dados_financeiros.get('lic', '-')}\n"
            f"🏛️ *IPVA (Sefaz):* {self.dados_financeiros.get('ipva', '-')}\n"
            f"🚨 *Multas (Detran):* {self.dados_financeiros.get('multas', '-')}\n"
            f"💎 *Total Líquido a Pagar:* *{self.total_str}*"
        )
        self.clipboard_clear()
        self.clipboard_append(resumo)
        self.btn_copiar.configure(text="✅ Copiado para a Área de Transferência!")


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

        aplicar_barra_titulo_windows(self, self.modo_noturno)

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
        self.itens_boletos.clear()
        total_veiculos = len(self.df_frota)
        col_chassi = FrotaDataProcessor.localizar_coluna_chassi(self.df_frota)

        for i in range(total_veiculos):
            dados = self.resultados_memoria.get(i, {})
            lic_str = dados.get("lic", "-")
            ipva_str = dados.get("ipva", "-")

            val_lic = formatar_moeda_float(lic_str)
            val_ipva = formatar_moeda_float(ipva_str)

            tem_lic = val_lic > 0 and "falha" not in str(lic_str).lower()
            tem_ipva = val_ipva > 0 and "falha" not in str(ipva_str).lower()

            if tem_lic or tem_ipva:
                row = self.df_frota.loc[i]
                placa = str(row[self.col_placa]).strip().upper()
                renavam = str(row[self.col_renavam]).strip()
                chassi = (
                    str(row[col_chassi]).strip()
                    if (col_chassi and col_chassi in self.df_frota.columns and pd.notna(row[col_chassi]))
                    else ""
                )

                self.itens_boletos.append({
                    "index": i,
                    "placa": placa,
                    "renavam": renavam,
                    "chassi": chassi,
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

        frame_rodape = ttk.Frame(container, style="PremiumBoleto.TFrame")
        frame_rodape.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))

        self.var_progresso_txt = tk.StringVar(value="Selecione os boletos desejados e clique em Iniciar Download.")
        self.lbl_progresso = ttk.Label(frame_rodape, textvariable=self.var_progresso_txt, font=("Segoe UI", 8, "bold"))
        self.lbl_progresso.pack(anchor="w", pady=(0, 3))

        self.barra_progresso = ModernProgressBar(frame_rodape, modo_noturno=self.modo_noturno)
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

        frame_scroll = tk.Frame(container, bg=self.bg_card, relief=tk.FLAT, bd=0, highlightthickness=1, highlightbackground=self.border_card)
        frame_scroll.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 6))

        canvas = tk.Canvas(frame_scroll, bg=self.bg_card, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame_scroll, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style="PremiumBoletoCard.TFrame")

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def _ao_redimensionar_canvas(event):
            canvas.itemconfig(canvas_window_id, width=event.width)
        canvas.bind("<Configure>", _ao_redimensionar_canvas)

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

                lbl_placa = tk.Label(card_item, text=f"🚗 {item['placa']}", font=("Consolas", 10, "bold"), bg=self.bg_item, fg=self.fg_title)
                lbl_placa.pack(side=tk.LEFT, padx=(6, 12))

                lbl_ren = tk.Label(card_item, text=f"RENAVAM: {item['renavam']}", font=("Segoe UI", 8), bg=self.bg_item, fg=self.fg_muted)
                lbl_ren.pack(side=tk.LEFT, padx=(0, 16))

                if item["tem_lic"]:
                    chk_lic = ttk.Checkbutton(card_item, text=f"Licenciamento ({item['val_lic_str']})", variable=item["var_lic"], style="Card.TCheckbutton")
                    chk_lic.pack(side=tk.LEFT, padx=(0, 14))
                else:
                    lbl_sem_lic = tk.Label(card_item, text="Licenciamento: Isento/Pago", font=("Segoe UI", 8), bg=self.bg_item, fg="#10B981")
                    lbl_sem_lic.pack(side=tk.LEFT, padx=(0, 14))

                if item["tem_ipva"]:
                    chk_ipva = ttk.Checkbutton(card_item, text=f"IPVA ({item['val_ipva_str']})", variable=item["var_ipva"], style="Card.TCheckbutton")
                    chk_ipva.pack(side=tk.LEFT, padx=(0, 14))
                else:
                    lbl_sem_ipva = tk.Label(card_item, text="IPVA: Isento/Pago", font=("Segoe UI", 8), bg=self.bg_item, fg="#10B981")
                    lbl_sem_ipva.pack(side=tk.LEFT, padx=(0, 14))

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
                    try:
                        self.downloader.reiniciar_sessao()
                    except Exception as err:
                        logger.warning(f"Aviso ao reiniciar sessão do downloader: {err}")
                    time.sleep(2.0)

                falhas_desta_rodada = []

                for item, tipo in fila_pendentes:
                    if self.cancelar_solicitado:
                        break

                    while self.is_pausado and not self.cancelar_solicitado:
                        time.sleep(0.3)
                    if self.cancelar_solicitado:
                        break

                    placa = item["placa"]
                    renavam = item["renavam"]
                    chassi = item.get("chassi", "")
                    lbl_tipo = "Licenciamento" if tipo == "lic" else "IPVA"

                    concluidos = len(self.tarefas_executadas_sucesso)
                    pct = int((concluidos / total_geral) * 100)
                    self.after(0, lambda c=concluidos, t=total_geral, plc=placa, tp=lbl_tipo, pc=pct, r=rodada:
                        self._atualizar_status_ui(c, t, plc, tp, pc, r)
                    )

                    if item["lbl_status"]:
                        txt_tentativa = f"Baixando {lbl_tipo}..." if rodada == 1 else f"Baixando {lbl_tipo} (T{rodada})..."
                        self.after(0, lambda lbl=item["lbl_status"], t=txt_tentativa: lbl.config(
                            text=t, foreground="#2563EB"
                        ))

                    if tipo == "lic":
                        sucesso, msg = self.downloader.baixar_boleto_licenciamento(placa, renavam, chassi=chassi)
                    else:
                        sucesso, msg = self.downloader.baixar_dae_ipva(placa, renavam, chassi=chassi)

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
# INTERFACE PRINCIPAL COMERCIAL (V4.1)
# ======================================================================================

class AppConsultaFrotaGUI:
    """Interface Gráfica Comercial (SaaS / Fintech Edition) para Automação de Frotas CE."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Fleet Command Pro • Auditoria & Inteligência de Frota CE (v4.1)")
        self._ajustar_geometria_responsiva()

        self.cfg = ConfigManager.carregar()

        # Variáveis de Estado
        self.caminho_arquivo = tk.StringVar(value=self.cfg.get("ultimo_arquivo", ""))
        self.pasta_destino_var = tk.StringVar(value=self.cfg.get("pasta_destino", ""))
        self.arquivo_excel_gerado = ""
        self.arquivo_pdf_gerado = ""
        self.arquivo_pdf_bi_gerado = ""
        self.ultimo_comparativo_bi: dict | None = None

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

        self.var_status = tk.StringVar(value="Pronto para iniciar. Selecione a planilha de frota ou use o menu Recentes.")
        self.var_busca = tk.StringVar(value="")

        # Métricas do Dashboard
        self.var_metricas_veiculos = tk.StringVar(value="0 / 0")
        self.var_metricas_lic = tk.StringVar(value="R$ 0,00")
        self.var_metricas_ipva = tk.StringVar(value="R$ 0,00")
        self.var_metricas_multas = tk.StringVar(value="R$ 0,00")
        self.var_metricas_total = tk.StringVar(value="R$ 0,00")

        # Controle de Execução, Telemetria e Memória
        self.is_executando = False
        self.is_pausado = False
        self.cancelar_solicitado = False
        self._tempo_inicio_lote = 0.0
        self.df_carregado: pd.DataFrame | None = None
        self.col_placa = "Placa"
        self.col_renavam = "RENAVAM"
        self.col_chassi: str | None = None
        self.resultados_memoria: dict[int, dict] = {}
        self.indices_em_falha: list[int] = []

        # Ordenação da Tabela
        self.coluna_ordenacao = ""
        self.ordem_reversa = False
        self._titulos_colunas = {
            "num": "#",
            "placa": "Placa",
            "renavam": "RENAVAM",
            "licenciamento": "Licenciamento",
            "ipva": "IPVA",
            "multas": "Multas",
            "total_pagar": "Total a Pagar",
            "situacao": "Situação Auditada"
        }

        # Modo Noturno e Componentes Visuais
        self.modo_noturno = self.cfg.get("modo_noturno", False)
        self.cards_metricas = []
        self.luxury_cards: list[LuxuryMetricCard] = []
        self.toggle_switches: list[ModernToggleSwitch] = []
        self.paineis_borda: list[tk.Frame] = []
        self.status_orb: LiveStatusOrb | None = None
        self.painel_bi: PainelAnaliticoBI | None = None
        self.botoes_filtro: dict[str, ttk.Button] = {}
        self.filtro_ativo = "TODOS"

        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self._criar_layout()
        self._aplicar_tema()

        if self.caminho_arquivo.get() and os.path.exists(self.caminho_arquivo.get()):
            self.root.after(200, lambda: self._carregar_planilha_automatica(self.caminho_arquivo.get()))

        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar_janela)

    def _salvar_preferencias(self):
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
        """Aplica tokens visuais comerciais para Modo Claro e Modo Noturno + DWM Titlebar."""
        aplicar_barra_titulo_windows(self.root, self.modo_noturno)

        if self.modo_noturno:
            bg_root = "#080C14"
            bg_card = "#0F172A"
            border_card = "#1E293B"
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
            self.border_card_atual = border_card
        else:
            bg_root = "#F1F5F9"
            bg_card = "#FFFFFF"
            border_card = "#CBD5E1"
            fg_title = "#0F172A"
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
            self.border_card_atual = border_card

        self.root.configure(bg=bg_root)
        self.style.configure(".", background=bg_root, foreground=fg_text, font=("Segoe UI", 9))
        self.style.configure("TFrame", background=bg_root)
        self.style.configure("Card.TFrame", background=bg_card)
        self.style.configure("TLabel", background=bg_root, foreground=fg_text)
        self.style.configure("Card.TLabel", background=bg_card, foreground=fg_text)
        self.style.configure("Header.TLabel", background=bg_root, font=("Segoe UI", 13, "bold"), foreground=fg_title)
        self.style.configure("SubHeader.TLabel", background=bg_root, font=("Segoe UI", 8), foreground=fg_sub)

        self.style.configure("Card.TCheckbutton", background=bg_card, foreground=fg_text, font=("Segoe UI", 9))
        self.style.map("Card.TCheckbutton", background=[("active", bg_card)], foreground=[("active", fg_title)])
        self.style.configure("Card.TRadiobutton", background=bg_card, foreground=fg_text, font=("Segoe UI", 9))
        self.style.map("Card.TRadiobutton", background=[("active", bg_card)], foreground=[("active", fg_title)])

        self.style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"), background=bg_btn_pri, foreground="#FFFFFF", padding=(13, 6))
        self.style.map("Primary.TButton", background=[("active", bg_btn_pri_act), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        self.style.configure("PrimaryDownload.TButton", font=("Segoe UI", 9, "bold"), background="#059669", foreground="#FFFFFF", padding=(12, 6))
        self.style.map("PrimaryDownload.TButton", background=[("active", "#047857"), ("disabled", "#4B5563" if self.modo_noturno else "#94A3B8")])

        self.style.configure("Action.TButton", font=("Segoe UI", 9, "bold"), background=bg_btn_act, foreground="#FFFFFF", padding=(14, 6))
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
        self.style.configure("TNotebook.Tab", background=bg_tab, foreground=fg_tab, padding=(16, 7), font=("Segoe UI", 9, "bold"))
        self.style.map("TNotebook.Tab", background=[("selected", bg_tab_sel)], foreground=[("selected", fg_tab_sel)])

        self.style.configure("Treeview", background=bg_tree, foreground=fg_tree, fieldbackground=bg_tree, rowheight=29, font=("Segoe UI", 9))
        self.style.configure("Treeview.Heading", background=bg_tree_hd, foreground=fg_tree_hd, font=("Segoe UI", 9, "bold"), padding=(6, 7))
        self.style.map("Treeview", background=[("selected", bg_sel)], foreground=[("selected", fg_sel)])

        self.style.configure("TEntry", fieldbackground=bg_card, foreground=fg_text, padding=(6, 4))
        self.style.configure("TCombobox", fieldbackground=bg_card, background=bg_btn_muted, foreground=fg_text, padding=(5, 4))

        for pnl in self.paineis_borda:
            pnl.configure(bg=bg_card, highlightbackground=border_card)

        for sw in self.toggle_switches:
            sw.atualizar_tema(self.modo_noturno, bg_override=bg_card)

        if hasattr(self, "barra_progresso") and isinstance(self.barra_progresso, ModernProgressBar):
            self.barra_progresso.atualizar_tema(self.modo_noturno)

        if self.painel_bi:
            self.painel_bi.atualizar_tema(self.modo_noturno)

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
        self.root.update_idletasks()
        tela_w = self.root.winfo_screenwidth()
        tela_h = self.root.winfo_screenheight()

        margem_h = 40
        margem_v = 70

        min_w = 860
        min_h = 560

        if tela_w >= 1920:
            largura = 1180
        elif tela_w >= 1600:
            largura = 1100
        elif tela_w >= 1366:
            largura = 1000
        elif tela_w >= 1200:
            largura = 940
        else:
            largura = max(min_w, tela_w - margem_h)

        if tela_h >= 1080:
            altura = 780
        elif tela_h >= 900:
            altura = 730
        elif tela_h >= 768:
            altura = 650
        elif tela_h >= 720:
            altura = 610
        else:
            altura = max(min_h, tela_h - margem_v)

        largura = min(largura, max(min_w, tela_w - margem_h))
        altura = min(altura, max(min_h, tela_h - margem_v))

        pos_x = max(0, (tela_w - largura) // 2)
        pos_y = max(0, (tela_h - margem_v - altura) // 2)

        self.root.minsize(min_w, min_h)
        self.root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    def _criar_layout(self):
        container = ttk.Frame(self.root, padding="14 10 14 10")
        container.pack(fill=tk.BOTH, expand=True)

        # ==============================================================================
        # 1. HEADER COMERCIAL (BRANDING, TELEMETRIA PULSANTE E TEMA)
        # ==============================================================================
        frame_topo = ttk.Frame(container)
        frame_topo.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))

        frame_titulos = ttk.Frame(frame_topo)
        frame_titulos.pack(side=tk.LEFT, fill=tk.X, expand=True)

        frame_marca = ttk.Frame(frame_titulos)
        frame_marca.pack(anchor="w")

        lbl_badge = tk.Label(
            frame_marca,
            text="⚡ FLEET COMMAND PRO",
            font=("Segoe UI", 7, "bold"),
            bg="#2563EB",
            fg="#FFFFFF",
            padx=7,
            pady=2
        )
        lbl_badge.pack(side=tk.LEFT, padx=(0, 8))

        lbl_badge_v4 = tk.Label(
            frame_marca,
            text="B2B EDITION v4.1",
            font=("Segoe UI", 7, "bold"),
            bg="#059669",
            fg="#FFFFFF",
            padx=6,
            pady=2
        )
        lbl_badge_v4.pack(side=tk.LEFT, padx=(0, 10))

        lbl_titulo = ttk.Label(
            frame_marca,
            text="Plataforma Executiva de Auditoria & Gestão de Frota",
            style="Header.TLabel"
        )
        lbl_titulo.pack(side=tk.LEFT)

        lbl_subtitulo = ttk.Label(
            frame_titulos,
            text="Detran-CE & Sefaz-CE  •  Auditoria em Lote, BI Financeiro em Tempo Real & Emissão Automatizada de Boletos",
            style="SubHeader.TLabel"
        )
        lbl_subtitulo.pack(anchor="w", pady=(2, 0))

        self.status_orb = LiveStatusOrb(frame_topo, modo_noturno=self.modo_noturno)
        self.status_orb.pack(side=tk.RIGHT, padx=(10, 0))

        self.btn_tema = ttk.Button(
            frame_topo,
            text="🌙 Modo Noturno",
            style="Theme.TButton",
            command=self._alternar_tema
        )
        self.btn_tema.pack(side=tk.RIGHT, padx=(8, 0), pady=2)

        # ==============================================================================
        # 2. DASHBOARD DE KPI CARDS NO TOPO (HIERARQUIA EXECUTIVA)
        # ==============================================================================
        frame_metricas = ttk.Frame(container)
        frame_metricas.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))

        self.luxury_cards.clear()
        card_configs = [
            ("FROTA AUDITADA", self.var_metricas_veiculos, "🚘", "#3B82F6", "✦ Progresso do Lote"),
            ("LICENCIAMENTO", self.var_metricas_lic, "📋", "#F59E0B", "✦ Detran-CE (0%)"),
            ("IPVA EM ABERTO", self.var_metricas_ipva, "🏛️", "#F43F5E", "✦ Sefaz-CE (0%)"),
            ("MULTAS DETRAN", self.var_metricas_multas, "🚨", "#EA580C", "✦ Infrações (0%)"),
            ("TOTAL A LIQUIDAR", self.var_metricas_total, "💎", "#10B981", "✦ Líquido sem Duplicidade"),
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

        # ==============================================================================
        # 3. PAINEL DE CONTROLE ESTILIZADO COM TOGGLE SWITCHES & MENU RECENTES
        # ==============================================================================
        frame_configs = tk.Frame(container, highlightthickness=1, padx=12, pady=8)
        frame_configs.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))
        self.paineis_borda.append(frame_configs)

        # Linha 1: Seleção de Planilha, Menu Recentes e Pasta Destino
        linha_arquivo = ttk.Frame(frame_configs, style="Card.TFrame")
        linha_arquivo.pack(fill=tk.X, pady=(0, 6))

        self.btn_selecionar = ttk.Button(
            linha_arquivo,
            text="📁 Selecionar Planilha",
            style="Primary.TButton",
            command=self._selecionar_arquivo
        )
        self.btn_selecionar.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_recentes = ttk.Button(
            linha_arquivo,
            text="🕘 Recentes ▾",
            style="Secondary.TButton",
            command=self._abrir_menu_recentes
        )
        self.btn_recentes.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_pasta_destino = ttk.Button(
            linha_arquivo,
            text="📂 Pasta Destino",
            style="Secondary.TButton",
            command=self._selecionar_pasta_destino
        )
        self.btn_pasta_destino.pack(side=tk.LEFT, padx=(0, 10))

        self.lbl_caminho = ttk.Label(
            linha_arquivo,
            text="Nenhuma planilha carregada • Selecione um arquivo .xlsx ou use 'Recentes'",
            font=("Segoe UI", 9, "italic"),
            foreground="#64748B",
            style="Card.TLabel"
        )
        self.lbl_caminho.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Linha 2: Toggle Switches em Pílula, Exportação e Botões de Execução
        linha_controles = ttk.Frame(frame_configs, style="Card.TFrame")
        linha_controles.pack(fill=tk.X)

        lbl_filtro_txt = ttk.Label(linha_controles, text="Motores:", font=("Segoe UI", 9, "bold"), style="Card.TLabel")
        lbl_filtro_txt.pack(side=tk.LEFT, padx=(0, 6))

        sw_ipva = ModernToggleSwitch(
            linha_controles, "🏛️ IPVA", self.var_ipva,
            cor_ativo="#E11D48", modo_noturno=self.modo_noturno, command=self._salvar_preferencias
        )
        sw_ipva.pack(side=tk.LEFT, padx=(0, 4))

        sw_lic = ModernToggleSwitch(
            linha_controles, "📋 Licenciamento", self.var_lic,
            cor_ativo="#D97706", modo_noturno=self.modo_noturno, command=self._salvar_preferencias
        )
        sw_lic.pack(side=tk.LEFT, padx=(0, 4))

        sw_multas = ModernToggleSwitch(
            linha_controles, "🚨 Multas", self.var_multas,
            cor_ativo="#EA580C", modo_noturno=self.modo_noturno, command=self._salvar_preferencias
        )
        sw_multas.pack(side=tk.LEFT, padx=(0, 4))

        sw_headless = ModernToggleSwitch(
            linha_controles, "🕶️ Invisível", self.var_headless,
            cor_ativo="#2563EB", modo_noturno=self.modo_noturno, command=self._salvar_preferencias
        )
        sw_headless.pack(side=tk.LEFT, padx=(0, 8))

        self.toggle_switches.extend([sw_ipva, sw_lic, sw_multas, sw_headless])

        # Checkbuttons ocultos de compatibilidade para testes automatizados
        self.chk_ipva = ttk.Checkbutton(linha_controles, variable=self.var_ipva)
        self.chk_lic = ttk.Checkbutton(linha_controles, variable=self.var_lic)
        self.chk_multas = ttk.Checkbutton(linha_controles, variable=self.var_multas)
        self.chk_headless = ttk.Checkbutton(linha_controles, variable=self.var_headless)

        # Seletores de Exportação e Workers
        rb_ambos = ttk.Radiobutton(
            linha_controles, text="Excel+PDF", value="ambos",
            variable=self.var_formato_exportacao, style="Card.TRadiobutton",
            state="normal" if REPORTLAB_DISPONIVEL else "disabled"
        )
        rb_ambos.pack(side=tk.LEFT, padx=(4, 4))

        rb_excel = ttk.Radiobutton(
            linha_controles, text="Só Excel", value="excel",
            variable=self.var_formato_exportacao, style="Card.TRadiobutton"
        )
        rb_excel.pack(side=tk.LEFT, padx=(0, 8))

        self.combo_workers = ttk.Combobox(
            linha_controles,
            values=["1x (Sequencial)", "2x Paralelo", "3x Paralelo", "4x Máximo"],
            state="readonly",
            width=13
        )
        idx_w = max(0, min(self.var_workers.get() - 1, 3))
        self.combo_workers.current(idx_w)
        self.combo_workers.bind("<<ComboboxSelected>>", self._ao_mudar_workers)
        self.combo_workers.pack(side=tk.LEFT, padx=(0, 8))

        # Comandos Principais à Direita
        self.btn_cancelar = ttk.Button(
            linha_controles, text="⏹️ Cancelar", style="Danger.TButton",
            state="disabled", command=self._cancelar_processamento
        )
        self.btn_cancelar.pack(side=tk.RIGHT, padx=(4, 0))

        self.btn_pausar = ttk.Button(
            linha_controles, text="⏸️ Pausar", style="Secondary.TButton",
            state="disabled", command=self._alternar_pausa
        )
        self.btn_pausar.pack(side=tk.RIGHT, padx=(4, 0))

        self.btn_iniciar = ttk.Button(
            linha_controles, text="⚡ INICIAR AUDITORIA", style="Action.TButton",
            state="disabled", command=self._iniciar_processamento
        )
        self.btn_iniciar.pack(side=tk.RIGHT)

        # ==============================================================================
        # 5. RODAPÉ EXECUTIVO COM BARRA DE PROGRESSO & TELEMETRIA DE ETA (FIXADO NO FUNDO)
        # ==============================================================================
        frame_rodape = ttk.Frame(container)
        frame_rodape.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))

        self.lbl_status = ttk.Label(
            frame_rodape,
            textvariable=self.var_status,
            font=("Segoe UI", 9, "bold")
        )
        self.lbl_status.pack(anchor="w", pady=(0, 3))

        self.barra_progresso = ModernProgressBar(frame_rodape, modo_noturno=self.modo_noturno)
        self.barra_progresso.pack(fill=tk.X, pady=(0, 6))

        frame_botoes_finais = ttk.Frame(frame_rodape)
        frame_botoes_finais.pack(fill=tk.X)

        self.btn_baixar_boletos = ttk.Button(
            frame_botoes_finais,
            text="📥 Central de Download de Boletos (PDF)",
            style="PrimaryDownload.TButton",
            state="disabled",
            command=self._abrir_central_download_boletos
        )
        self.btn_baixar_boletos.pack(side=tk.RIGHT)

        self.btn_abrir_excel = ttk.Button(
            frame_botoes_finais, text="📊 Relatório Excel", style="Secondary.TButton",
            state="disabled", command=self._abrir_excel_gerado
        )
        self.btn_abrir_excel.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_abrir_pdf = ttk.Button(
            frame_botoes_finais, text="📄 Relatório PDF", style="Secondary.TButton",
            state="disabled", command=self._abrir_pdf_gerado
        )
        self.btn_abrir_pdf.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_abrir_pdf_bi = ttk.Button(
            frame_botoes_finais, text="📈 Relatório BI (PDF)", style="Secondary.TButton",
            state="disabled", command=self._abrir_pdf_bi_gerado
        )
        self.btn_abrir_pdf_bi.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_abrir_pasta = ttk.Button(
            frame_botoes_finais, text="📂 Pasta Destino", style="Secondary.TButton",
            state="disabled", command=self._abrir_pasta_destino
        )
        self.btn_abrir_pasta.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_reconsultar_falhas = ttk.Button(
            frame_botoes_finais, text="⚠️ Reconsultar Falhas", style="Secondary.TButton",
            state="disabled", command=self._reconsultar_apenas_falhas
        )
        self.btn_reconsultar_falhas.pack(side=tk.LEFT, padx=(0, 8))

        # ==============================================================================
        # 4. CENTRAL DE MONITORAMENTO E BI (TABELA DINÂMICA + PAINEL ANALÍTICO + LOGS)
        # ==============================================================================
        frame_centro = tk.Frame(container, highlightthickness=1, padx=8, pady=6)
        frame_centro.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.paineis_borda.append(frame_centro)

        self.notebook = ttk.Notebook(frame_centro)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # ABA 1: Tabela Dinâmica da Frota
        aba_tabela = ttk.Frame(self.notebook, padding="4 6 4 4")
        self.notebook.add(aba_tabela, text=" 📊 Tabela Dinâmica da Frota ")

        frame_filtros_bar = ttk.Frame(aba_tabela)
        frame_filtros_bar.pack(fill=tk.X, pady=(0, 6))

        self.btn_boletos_tabela = ttk.Button(
            frame_filtros_bar,
            text="📥 Central de Boletos (PDF)",
            style="PrimaryDownload.TButton",
            state="disabled",
            command=self._abrir_central_download_boletos
        )
        self.btn_boletos_tabela.pack(side=tk.RIGHT)

        ttk.Label(frame_filtros_bar, text="🔍 Busca:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.entry_busca = ttk.Entry(frame_filtros_bar, textvariable=self.var_busca, width=18)
        self.entry_busca.pack(side=tk.LEFT, padx=(0, 4))
        self.var_busca.trace_add("write", lambda *_: self._filtrar_tabela())

        ttk.Button(frame_filtros_bar, text="✖", style="Muted.TButton", command=lambda: self.var_busca.set("")).pack(side=tk.LEFT, padx=(0, 10))

        self.botoes_filtro["TODOS"] = ttk.Button(
            frame_filtros_bar, text="✦ Todos (0)", style="ActiveFilter.TButton",
            command=lambda: self._definir_filtro_pílula("TODOS")
        )
        self.botoes_filtro["TODOS"].pack(side=tk.LEFT, padx=(0, 3))

        self.botoes_filtro["DEBITOS"] = ttk.Button(
            frame_filtros_bar, text="⚠️ Com Débitos (0)", style="Muted.TButton",
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
        self.botoes_filtro["FALHAS"].pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(
            frame_filtros_bar,
            text="💡 Duplo clique no veículo abre o Raio-X",
            font=("Segoe UI", 8, "italic"),
            foreground="#64748B"
        ).pack(side=tk.LEFT)

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

        # Eventos de Duplo Clique (Raio-X) e Menu de Contexto
        self.tree.bind("<Double-1>", self._abrir_raio_x_evento)
        self.menu_contexto = tk.Menu(self.tree, tearoff=0)
        self.menu_contexto.add_command(label="🔍 Abrir Raio-X Executivo do Veículo", command=self._abrir_raio_x_selecionado)
        self.menu_contexto.add_separator()
        self.menu_contexto.add_command(label="🔄 Reconsultar Veículo Selecionado", command=self._reconsultar_selecionado)
        self.tree.bind("<Button-3>", self._exibir_menu_contexto)

        # ABA 2: Painel Analítico (BI da Frota)
        aba_bi = ttk.Frame(self.notebook, padding="4 6 4 4")
        self.notebook.add(aba_bi, text=" 📈 Painel Analítico (BI da Frota) ")
        self.painel_bi = PainelAnaliticoBI(aba_bi, modo_noturno=self.modo_noturno)
        self.painel_bi.pack(fill=tk.BOTH, expand=True)

        # ABA 3: Console de Logs com Destaque Sintático
        aba_logs = ttk.Frame(self.notebook, padding="4 6 4 4")
        self.notebook.add(aba_logs, text=" 📜 Terminal Executivo de Logs ")

        self.txt_logs = tk.Text(aba_logs, height=6, wrap=tk.WORD, font=("Consolas", 9), relief=tk.FLAT, padx=10, pady=10)
        scroll_logs = ttk.Scrollbar(aba_logs, orient="vertical", command=self.txt_logs.yview)
        self.txt_logs.configure(yscrollcommand=scroll_logs.set)
        self.txt_logs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_logs.pack(side=tk.RIGHT, fill=tk.Y)

    # ==================================================================================
    # MENU DE ARQUIVOS RECENTES & RAIO-X DO VEÍCULO
    # ==================================================================================

    def _abrir_menu_recentes(self):
        menu = tk.Menu(
            self.root, tearoff=0,
            bg="#0F172A" if self.modo_noturno else "#FFFFFF",
            fg="#F8FAFC" if self.modo_noturno else "#0F172A",
            font=("Segoe UI", 9)
        )
        cfg_atual = ConfigManager.carregar()
        recentes = [p for p in cfg_atual.get("arquivos_recentes", []) if os.path.exists(p)]

        # Adiciona também planilhas .xlsx presentes no diretório atual
        try:
            for f in sorted(os.listdir(".")):
                if f.lower().endswith(".xlsx") and not f.startswith("~"):
                    caminho_abs = os.path.abspath(f)
                    if caminho_abs not in recentes:
                        recentes.append(caminho_abs)
        except Exception:
            pass

        if not recentes:
            menu.add_command(label="(Nenhuma planilha recente encontrada)", state="disabled")
        else:
            for caminho in recentes[:8]:
                nome = os.path.basename(caminho)
                icone = "📊" if "resultado" in nome.lower() else "📁"
                menu.add_command(
                    label=f"{icone}  {nome}",
                    command=lambda c=caminho: self._carregar_planilha_automatica(c)
                )

        x = self.btn_recentes.winfo_rootx()
        y = self.btn_recentes.winfo_rooty() + self.btn_recentes.winfo_height()
        menu.tk_popup(x, y)

    def _abrir_raio_x_evento(self, event):
        item_id = self.tree.identify_row(event.y)
        if item_id:
            self._abrir_modal_raio_x(int(item_id))

    def _abrir_raio_x_selecionado(self):
        sel = self.tree.selection()
        if sel:
            self._abrir_modal_raio_x(int(sel[0]))

    def _abrir_modal_raio_x(self, idx: int):
        if self.df_carregado is None or idx not in self.df_carregado.index:
            return
        row = self.df_carregado.loc[idx]
        placa = str(row[self.col_placa]).strip().upper()
        renavam = str(row[self.col_renavam]).strip()

        colunas_ignorar = {
            self.col_placa, self.col_renavam, self.col_chassi,
            "Valor_Licenciamento", "Valor_IPVA", "Valor_Multas",
            "Total_a_Pagar", "Status", "Data_Consulta"
        }
        extras = {
            str(col): str(row[col])
            for col in self.df_carregado.columns
            if col not in colunas_ignorar
            and not any(k in str(col).strip().upper() for k in ["CHASSI", "CHASSIS", "UNNAMED"])
            and str(col).strip().upper() != "VIN"
            and pd.notna(row[col])
        }
        fin = self.resultados_memoria.get(idx, {"lic": "-", "ipva": "-", "multas": "-"})
        JanelaRaioXVeiculo(
            parent=self.root,
            placa=placa,
            renavam=renavam,
            dados_extras=extras,
            dados_financeiros=fin,
            modo_noturno=self.modo_noturno,
            on_baixar_boletos=self._abrir_central_download_boletos
        )

    # ==================================================================================
    # LÓGICA DE CONTROLE, FILTROS, BI E PROCESSAMENTO
    # ==================================================================================

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
        termos = ("falha", "erro", "timeout", "inválid", "invalid", "não localizado", "nao localizado")
        return any(
            termo in str(v).lower()
            for v in valores
            for termo in termos
        )

    def _adicionar_log(self, mensagem: str, nivel: str = "INFO"):
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
        qtd_boletos = 0
        falhas = len(self.indices_em_falha)

        soma_lic = 0.0
        soma_ipva = 0.0
        soma_multas = 0.0
        soma_total = 0.0
        ranking_debitos = []

        col_prop = next(
            (c for c in self.df_carregado.columns if "PROPRIET" in str(c).upper() or "CLIENTE" in str(c).upper()),
            None
        )

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

            val_lic_f = formatar_moeda_float(lic)
            val_ipva_f = formatar_moeda_float(ipva)
            val_multas_f = formatar_moeda_float(multas)
            tot = calcular_total_a_pagar_veiculo(lic, ipva, multas, m_av)

            soma_lic += val_lic_f
            soma_ipva += val_ipva_f
            soma_multas += val_multas_f
            soma_total += tot

            if val_lic_f > 0 or val_ipva_f > 0:
                qtd_boletos += 1

            is_restr = "restriç" in str(lic).lower() or "restric" in str(lic).lower()
            is_pago = (str(lic).strip().upper() == "PAGO" or str(ipva).strip().upper() == "PAGO")

            if is_restr:
                restricoes += 1
            if tot > 0:
                debitos += 1
                if idx in self.df_carregado.index:
                    row = self.df_carregado.loc[idx]
                    plc = str(row[self.col_placa]).strip().upper()
                    prop = str(row[col_prop]).strip() if col_prop and pd.notna(row[col_prop]) else ""
                    ranking_debitos.append((plc, prop, tot, "Com Débitos"))
            elif is_pago or tot == 0:
                pagos += 1

        self.botoes_filtro["TODOS"].configure(text=f"✦ Todos ({total})")
        self.botoes_filtro["DEBITOS"].configure(text=f"⚠️ Com Débitos ({debitos})")
        self.botoes_filtro["PAGOS"].configure(text=f"✅ Pagos ({pagos})")
        self.botoes_filtro["RESTRICAO"].configure(text=f"🚨 Restrições ({restricoes})")
        self.botoes_filtro["FALHAS"].configure(text=f"❌ Falhas ({falhas})")

        # Atualiza proporções visuais nos LuxuryMetricCards
        if len(self.luxury_cards) == 5:
            concluidos = len(self.resultados_memoria)
            self.luxury_cards[0].definir_proporcao(concluidos / max(1, total), f"✦ {concluidos} de {total} auditados")
            soma_bruta = max(1.0, soma_lic + soma_ipva + soma_multas)
            p_lic = (soma_lic / soma_bruta) if (soma_lic + soma_ipva + soma_multas) > 0 else 0.0
            p_ipva = (soma_ipva / soma_bruta) if (soma_lic + soma_ipva + soma_multas) > 0 else 0.0
            p_mul = (soma_multas / soma_bruta) if (soma_lic + soma_ipva + soma_multas) > 0 else 0.0
            self.luxury_cards[1].definir_proporcao(p_lic, f"✦ Detran-CE ({p_lic * 100:.0f}%)")
            self.luxury_cards[2].definir_proporcao(p_ipva, f"✦ Sefaz-CE ({p_ipva * 100:.0f}%)")
            self.luxury_cards[3].definir_proporcao(p_mul, f"✦ Infrações ({p_mul * 100:.0f}%)")
            self.luxury_cards[4].definir_proporcao(1.0 if soma_total > 0 else 0.0, f"✦ {debitos} veíc. com pendências")

        # Atualiza o Painel Analítico de BI
        if self.painel_bi:
            ranking_debitos.sort(key=lambda x: x[2], reverse=True)
            self.painel_bi.atualizar_dados({
                "total_veiculos": total,
                "soma_lic": soma_lic,
                "soma_ipva": soma_ipva,
                "soma_multas": soma_multas,
                "soma_total": soma_total,
                "qtd_pagos": pagos,
                "qtd_debitos": debitos,
                "qtd_restricao": restricoes,
                "qtd_falhas": falhas,
                "qtd_boletos": qtd_boletos,
                "top_veiculos": ranking_debitos[:10]
            })

    def _exibir_menu_contexto(self, event):
        item_selecionado = self.tree.identify_row(event.y)
        if item_selecionado:
            if item_selecionado not in self.tree.selection():
                self.tree.selection_set(item_selecionado)
            self.menu_contexto.post(event.x_root, event.y_root)

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
            self.col_chassi = FrotaDataProcessor.localizar_coluna_chassi(df)

            ConfigManager.adicionar_arquivo_recente(caminho)

            if not self.pasta_destino_var.get():
                pasta_padrao = os.path.dirname(os.path.abspath(caminho))
                self.pasta_destino_var.set(pasta_padrao)

            nome = os.path.basename(caminho)
            nome_base_saida = FrotaDataProcessor.gerar_nome_base_resultado(caminho)
            self.lbl_caminho.config(
                text=f"Carregado: {nome} ({len(df)} veículos) ➔ Saída: '{nome_base_saida}.xlsx/.pdf/{nome_base_saida}_BI.pdf'",
                foreground="#34D399" if self.modo_noturno else "#15803D",
                font=("Segoe UI", 9, "bold")
            )

            self.arquivo_excel_gerado = ""
            self.arquivo_pdf_gerado = ""
            self.arquivo_pdf_bi_gerado = ""
            self.btn_abrir_excel.config(state="disabled")
            self.btn_abrir_pdf.config(state="disabled")
            if hasattr(self, "btn_abrir_pdf_bi"):
                self.btn_abrir_pdf_bi.config(state="disabled")

            self.btn_iniciar.config(state="normal")
            self.var_status.set(f"Arquivo carregado com sucesso: {len(df)} veículos prontos. Saída: {nome_base_saida}")
            self._preencher_tabela_inicial()

            # Se a planilha carregada já possuir colunas auditadas (ex: resultados_frota_ce.xlsx),
            # popula automaticamente os dados para inspeção imediata do BI e da tabela!
            cols_existentes = set(df.columns)
            if {"Valor_Licenciamento", "Valor_IPVA", "Valor_Multas"}.issubset(cols_existentes):
                self._hidratar_resultados_existentes(df)
                self._adicionar_log(
                    f"Planilha auditada '{nome}' carregada com {len(df)} resultados pré-existentes (BI e métricas ativados).",
                    nivel="SUCESSO"
                )
            else:
                self._adicionar_log(
                    f"Planilha '{nome}' carregada ({len(df)} registros). Arquivos de saída configurados: '{nome_base_saida}.xlsx', '{nome_base_saida}.pdf' e '{nome_base_saida}_BI.pdf'"
                )

        except Exception as e:
            messagebox.showerror("Erro ao Carregar", f"Não foi possível processar a planilha:\n\n{e}")
            self.lbl_caminho.config(text="Erro ao carregar o arquivo", foreground="#DC2626")
            self.btn_iniciar.config(state="disabled")

    def _hidratar_resultados_existentes(self, df: pd.DataFrame):
        """Popula a tabela, os cards e o BI quando o usuário abre uma planilha já auditada."""
        soma_lic = 0.0
        soma_ipva = 0.0
        soma_multas = 0.0
        soma_tot = 0.0
        total = len(df)

        for idx, row in df.iterrows():
            lic = str(row.get("Valor_Licenciamento", "-")).strip()
            ipva = str(row.get("Valor_IPVA", "-")).strip()
            multas = str(row.get("Valor_Multas", "-")).strip()
            tot_salvo = str(row.get("Total_a_Pagar", "")).strip()

            lic_f = formatar_moeda_float(lic)
            ipva_f = formatar_moeda_float(ipva)
            multas_f = formatar_moeda_float(multas)
            if tot_salvo and "R$" in tot_salvo:
                tot_f = formatar_moeda_float(tot_salvo)
                m_av = formatar_float_moeda(max(0.0, round(tot_f - lic_f - ipva_f, 2))) if lic_f > 0 else multas
            else:
                m_av = None
                tot_f = calcular_total_a_pagar_veiculo(lic, ipva, multas)

            soma_lic += lic_f
            soma_ipva += ipva_f
            soma_multas += multas_f
            soma_tot += tot_f

            self.resultados_memoria[idx] = {
                "lic": lic,
                "ipva": ipva,
                "multas": multas,
                "multas_a_vencer": m_av or "R$ 0,00"
            }
            self._atualizar_linha_tabela(
                idx, lic, ipva, multas, m_av, idx + 1, total,
                soma_lic, soma_ipva, soma_multas, soma_tot
            )

        total_com_debitos = sum(
            1 for d in self.resultados_memoria.values()
            if formatar_moeda_float(d.get("lic", "0")) > 0 or formatar_moeda_float(d.get("ipva", "0")) > 0
        )
        if total_com_debitos > 0:
            self.btn_baixar_boletos.config(state="normal")
            self.btn_boletos_tabela.config(state="normal")
        if hasattr(self, "btn_abrir_pdf_bi") and REPORTLAB_DISPONIVEL:
            self.btn_abrir_pdf_bi.config(state="normal")
        if self.status_orb:
            self.status_orb.definir_status("pronto", "DADOS AUDITADOS PRONTOS")

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

            passa_termo = (not termo) or (termo in placa) or (termo in renavam) or (termo in situacao.upper())

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
        itens = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]

        if col == "num":
            itens.sort(key=lambda t: int(t[0]) if str(t[0]).isdigit() else 0, reverse=self.ordem_reversa)
        elif col in ["licenciamento", "ipva", "multas", "total_pagar"]:
            itens.sort(key=lambda t: (formatar_moeda_float(t[0]), str(t[0]).upper()), reverse=self.ordem_reversa)
        else:
            itens.sort(key=lambda t: t[0].upper(), reverse=self.ordem_reversa)

        for index, (_, k) in enumerate(itens):
            self.tree.move(k, "", index)

        # Atualiza indicadores visuais de ordenação nos cabeçalhos (▲ / ▼ / ↕)
        seta = "▼" if self.ordem_reversa else "▲"
        for cid, titulo_base in self._titulos_colunas.items():
            marcador = seta if cid == col else "↕"
            self.tree.heading(cid, text=f"{titulo_base} {marcador}")

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

        if indices_alvo is None:
            self._preencher_tabela_inicial()

        self.is_executando = True
        self.is_pausado = False
        self.cancelar_solicitado = False
        self._tempo_inicio_lote = time.time()

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
            self.status_orb.definir_status("processando", "INICIANDO AUDITORIA...")

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

        # Captura a última planilha de resultados existente no disco ANTES de ser sobrescrita
        df_anterior_disco, data_anterior_disco = FrotaDataProcessor.carregar_resultado_anterior(arquivo_excel)
        if df_anterior_disco is not None:
            self._adicionar_log(
                f"Planilha anterior localizada ('{os.path.basename(arquivo_excel)}' de {data_anterior_disco}). "
                f"As alterações e pagamentos serão comparados no Relatório de BI."
            )

        scrapers_ativos = []
        try:
            df = self.df_carregado.copy()
            total_geral = len(df)

            lista_indices = indices_alvo if indices_alvo is not None else list(df.index)
            total_lote = len(lista_indices)

            self.root.after(0, lambda: self.var_status.set("Inicializando navegadores Chrome..."))
            self._adicionar_log(f"Iniciando lote de {total_lote} veículos ({num_workers} worker(s)). Destino: '{nome_frota_log}.xlsx' e '{nome_frota_log}.pdf'")

            lock_memoria = threading.Lock()
            concluidos_contador = [0]
            erros_workers = []

            def worker_task(worker_id: int, sub_indices: list[int]):
                scraper = None
                try:
                    scraper = FrotaScraperEngine(headless=modo_headless)
                    scrapers_ativos.append(scraper)
                    scraper.iniciar()
                    for idx in sub_indices:
                        if self.cancelar_solicitado:
                            break

                        while self.is_pausado and not self.cancelar_solicitado:
                            time.sleep(0.5)

                        row = df.loc[idx]
                        placa = str(row[self.col_placa]).strip().upper()
                        renavam = str(row[self.col_renavam]).strip()
                        chassi = (
                            str(row[self.col_chassi]).strip()
                            if (self.col_chassi and self.col_chassi in df.columns and pd.notna(row[self.col_chassi]))
                            else ""
                        )

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
                                    placa, renavam, consultar_lic, consultar_multas, chassi=chassi
                                )
                                val_detran_lic = res_detran[0]
                                if str(val_detran_lic).strip() in ["R$ 0,00", "0,00"]:
                                    val_detran_lic = "PAGO"
                                val_detran_multas = res_detran[1]
                                multas_a_vencer = getattr(res_detran, "multas_a_vencer", "R$ 0,00")
                                multas_vencidas = getattr(res_detran, "multas_vencidas", "R$ 0,00")
                                time.sleep(0.3)

                            if consultar_ipva:
                                val_sefaz_ipva = scraper.consultar_sefaz_ipva(placa, renavam, chassi=chassi)
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

                except Exception as exc:
                    logger.error(f"Erro no Worker {worker_id + 1}: {exc}", exc_info=True)
                    with lock_memoria:
                        erros_workers.append(f"Worker {worker_id + 1}: {exc}")
                    self._adicionar_log(f"Erro no Worker {worker_id + 1}: {exc}", nivel="ERRO")
                finally:
                    if scraper:
                        try:
                            scraper.fechar()
                        except Exception:
                            pass

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

            if concluidos_contador[0] == 0:
                if self.cancelar_solicitado:
                    self.root.after(0, lambda: self.var_status.set("Auditoria cancelada pelo operador."))
                    self.root.after(0, self._restaurar_botoes_apos_conclusao_ou_erro)
                    return
                detalhe = "\n".join(erros_workers) if erros_workers else "Nenhum veículo pôde ser consultado (verifique os logs)."
                self._adicionar_log(f"Falha na auditoria: nenhum veículo consultado. Detalhes:\n{detalhe}", nivel="ERRO")
                self.root.after(0, lambda: self.var_status.set("Falha na auditoria. Verifique a aba de logs."))
                self.root.after(0, lambda: messagebox.showerror(
                    "Falha na Auditoria",
                    f"Nenhum veículo pôde ser auditado!\n\nMotivo:\n{detalhe}\n\nConsulte a aba 'Logs de Operação' para mais detalhes."
                ))
                self.root.after(0, self._restaurar_botoes_apos_conclusao_ou_erro)
                return

            self.root.after(0, lambda: self.var_status.set("Consolidando dados e gerando relatórios..."))

            lics_final = [self.resultados_memoria.get(i, {}).get("lic", "-") for i in range(total_geral)]
            ipvas_final = [self.resultados_memoria.get(i, {}).get("ipva", "-") for i in range(total_geral)]
            multas_final = [self.resultados_memoria.get(i, {}).get("multas", "-") for i in range(total_geral)]
            multas_a_vencer_final = [self.resultados_memoria.get(i, {}).get("multas_a_vencer") for i in range(total_geral)]

            df_final = FrotaDataProcessor.consolidar_resultados(
                df, lics_final, ipvas_final, multas_final, total_geral,
                resultados_multas_a_vencer=multas_a_vencer_final
            )

            # Compara com a última planilha feita antes de sobrescrevê-la
            comparativo_anterior = None
            if df_anterior_disco is not None:
                comparativo_anterior = FrotaDataProcessor.comparar_auditorias(
                    df_anterior_disco,
                    df_final,
                    data_anterior=data_anterior_disco
                )
                self.ultimo_comparativo_bi = comparativo_anterior
                if comparativo_anterior:
                    nv = "SUCESSO" if comparativo_anterior.get("teve_alteracao") else "INFO"
                    self._adicionar_log(
                        f"Comparativo de evolução ({comparativo_anterior['data_anterior']} -> {comparativo_anterior['data_atual']}):",
                        nivel=nv
                    )
                    for frase in comparativo_anterior.get("frases_resumo", []):
                        self._adicionar_log(f"  ↳ {frase}", nivel=nv)
            else:
                self.ultimo_comparativo_bi = None

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
                    df_final,
                    arquivo_pdf_bi,
                    totais_dict,
                    filtros_ativos,
                    identificador_frota=identificador_frota,
                    comparativo_anterior=comparativo_anterior
                )
                self.arquivo_pdf_bi_gerado = caminho_pdf_bi_gerado
                self._adicionar_log(f"Relatório PDF de BI gerado com sucesso: {os.path.basename(caminho_pdf_bi_gerado)}")

            if os.path.exists(chk_path):
                try: os.remove(chk_path)
                except Exception: pass

            self.root.after(
                0,
                self._finalizar_processamento_ui,
                caminho_excel_gerado,
                caminho_pdf_gerado,
                self.cancelar_solicitado,
                concluidos_contador[0],
                total_lote
            )

        except Exception as e:
            logger.error(f"Erro durante o processamento da frota: {e}", exc_info=True)
            self.root.after(0, lambda: messagebox.showerror("Erro de Processamento", f"Ocorreu uma falha no processamento:\n\n{e}"))
            self.root.after(0, lambda: self.var_status.set("Erro durante o processamento. Verifique a aba de logs."))
            self.root.after(0, self._restaurar_botoes_apos_conclusao_ou_erro)
        finally:
            for sc in scrapers_ativos:
                try:
                    sc.fechar()
                except Exception:
                    pass
            self.is_executando = False
            self.is_pausado = False

    def _marcar_linha_processando(self, index: int, placa: str, renavam: str):
        if self.tree.exists(str(index)):
            self.tree.item(str(index), values=(index + 1, placa, renavam, "...", "...", "...", "...", "Processando..."), tags=("processando",))
            self.tree.see(str(index))

    def _atualizar_progresso_ui(self, atual: int, total: int, placa: str, pct: int):
        decorrido = max(0.1, time.time() - self._tempo_inicio_lote) if self._tempo_inicio_lote else 0.0
        eta_str = ""
        if atual > 0 and decorrido > 0.5:
            vel = decorrido / atual
            restante = max(0.0, (total - atual) * vel)
            mins, segs = divmod(int(restante), 60)
            eta_str = f"{vel:.1f}s/veíc  •  ETA: {mins:02d}m {segs:02d}s"
            self.barra_progresso.definir_telemetria(eta_str)

        self.var_status.set(f"Consultando ({atual}/{total}): {placa}..." + (f"  [{eta_str}]" if eta_str else ""))
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
        if self.status_orb and self.is_executando:
            self.status_orb.definir_status("processando", f"AUDITANDO: #{index + 1} / {total_geral}")

    def _restaurar_botoes_apos_conclusao_ou_erro(self):
        """Restaura o estado dos botões da interface após cancelamento ou falha."""
        self.btn_iniciar.config(state="normal")
        self.btn_selecionar.config(state="normal")
        self.btn_pasta_destino.config(state="normal")
        self.btn_cancelar.config(state="disabled")
        self.btn_pausar.config(state="disabled", text="⏸️ Pausar")
        self.btn_abrir_pasta.config(state="normal")
        if self.indices_em_falha:
            self.btn_reconsultar_falhas.config(state="normal")
        if self.status_orb:
            self.status_orb.definir_status("pronto", "PRONTO")

    def _finalizar_processamento_ui(self, arq_excel: str, arq_pdf: str, foi_cancelado: bool, concluidos: int = 0, total_lote: int = 0):
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
        if hasattr(self, "btn_abrir_pdf_bi") and self.arquivo_pdf_bi_gerado and os.path.exists(self.arquivo_pdf_bi_gerado):
            self.btn_abrir_pdf_bi.config(state="normal")

        if self.indices_em_falha:
            self.btn_reconsultar_falhas.config(state="normal")

        total_com_debitos = sum(
            1 for d in self.resultados_memoria.values()
            if formatar_moeda_float(d.get("lic", "0")) > 0 or formatar_moeda_float(d.get("ipva", "0")) > 0
        )
        if total_com_debitos > 0:
            self.btn_baixar_boletos.config(state="normal")
            if hasattr(self, "btn_boletos_tabela"):
                self.btn_boletos_tabela.config(state="normal")

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
            messagebox.showinfo("Interrompido", f"O processamento foi interrompido e os {concluidos} veículos concluídos foram gravados com sucesso.")
        else:
            if total_lote > 0 and concluidos < total_lote:
                self.var_status.set(f"Processamento parcial: {concluidos}/{total_lote} veículos concluídos.")
                msg = f"Consultas parcialmente concluídas ({concluidos}/{total_lote} veículos)!\n\nTotal a Pagar na Frota: {self.var_metricas_total.get()}"
            else:
                self.var_status.set("Processamento concluído com sucesso!")
                msg = f"Consultas concluídas com sucesso!\n\nTotal a Pagar na Frota: {self.var_metricas_total.get()}"
            if arq_excel and os.path.exists(arq_excel):
                msg += f"\n\n📊 Planilha Excel: {os.path.basename(arq_excel)}"
            if arq_pdf and os.path.exists(arq_pdf):
                msg += f"\n📑 Relatório PDF: {os.path.basename(arq_pdf)}"
            if self.arquivo_pdf_bi_gerado and os.path.exists(self.arquivo_pdf_bi_gerado):
                msg += f"\n📈 Relatório BI (PDF): {os.path.basename(self.arquivo_pdf_bi_gerado)}"
            if self.ultimo_comparativo_bi and self.ultimo_comparativo_bi.get("teve_alteracao"):
                resumos = self.ultimo_comparativo_bi.get("frases_resumo", [])[:4]
                msg += f"\n\n🔄 Alterações desde a última planilha ({self.ultimo_comparativo_bi.get('data_anterior')}):\n"
                msg += "\n".join(f"• {r}" for r in resumos)
            if total_com_debitos > 0:
                msg += f"\n\n💡 {total_com_debitos} veículo(s) possuem débitos de Licenciamento ou IPVA.\nVocê pode usar o botão 'Baixar Boletos' abaixo para emitir os PDFs."
            messagebox.showinfo("Sucesso", msg)

    def _abrir_central_download_boletos(self):
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
        if not self.arquivo_pdf_bi_gerado or not os.path.exists(self.arquivo_pdf_bi_gerado):
            if self.df_carregado is not None and self.resultados_memoria and REPORTLAB_DISPONIVEL:
                try:
                    caminho = self.caminho_arquivo.get()
                    pasta_destino = self.pasta_destino_var.get().strip() or os.path.dirname(os.path.abspath(caminho))
                    total_geral = len(self.df_carregado)
                    lics = [self.resultados_memoria.get(i, {}).get("lic", "-") for i in range(total_geral)]
                    ipvas = [self.resultados_memoria.get(i, {}).get("ipva", "-") for i in range(total_geral)]
                    multas = [self.resultados_memoria.get(i, {}).get("multas", "-") for i in range(total_geral)]
                    m_av = [self.resultados_memoria.get(i, {}).get("multas_a_vencer") for i in range(total_geral)]
                    df_final = FrotaDataProcessor.consolidar_resultados(
                        self.df_carregado, lics, ipvas, multas, total_geral, resultados_multas_a_vencer=m_av
                    )
                    arq_bi = FrotaDataProcessor.obter_caminho_pdf_bi(caminho, pasta_destino)
                    nome_frota_log = FrotaDataProcessor.gerar_nome_base_resultado(caminho)
                    ident = re.sub(r"^resultados[_\-\s]*", "", nome_frota_log, flags=re.IGNORECASE).strip() or nome_frota_log
                    filtros = []
                    if self.var_ipva.get(): filtros.append("IPVA")
                    if self.var_lic.get(): filtros.append("Licenciamento")
                    if self.var_multas.get(): filtros.append("Multas")
                    self.arquivo_pdf_bi_gerado = FrotaRelatorioExporter.exportar_pdf_bi(
                        df_final,
                        arq_bi,
                        filtros_ativos=filtros or ["IPVA", "Licenciamento", "Multas"],
                        identificador_frota=ident,
                        comparativo_anterior=self.ultimo_comparativo_bi
                    )
                    self._adicionar_log(f"Relatório PDF de BI gerado sob demanda: {os.path.basename(self.arquivo_pdf_bi_gerado)}", nivel="SUCESSO")
                except Exception as e:
                    messagebox.showerror("Erro ao Gerar BI PDF", f"Não foi possível gerar o Relatório PDF de BI:\n\n{e}")
                    return
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
        if self.status_orb:
            self.status_orb._parar_pulso()
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
