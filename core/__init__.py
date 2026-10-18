"""
Pacote Central (Core) do Sistema de Automação de Frotas CE.
"""
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
    extrair_placa_limpa,
    sanitizar_renavam,
    sanitizar_chassi,
    gerar_nome_base_resultado,
    obter_caminhos_arquivos_saida,
    obter_caminho_pdf_bi,
    carregar_resultado_anterior,
    comparar_auditorias
)
from core.scraper_engine import FrotaScraperEngine
from core.boleto_downloader import FrotaBoletoDownloader
from core.relatorio_exporter import FrotaRelatorioExporter

__all__ = [
    "ConfigManager",
    "obter_logger",
    "FrotaDataProcessor",
    "formatar_moeda_float",
    "formatar_float_moeda",
    "calcular_total_a_pagar_veiculo",
    "obter_estilo_financeiro",
    "obter_ultimo_dia_util_mes",
    "obter_proximo_dia_util",
    "PADRAO_PLACA",
    "extrair_placa_limpa",
    "sanitizar_renavam",
    "sanitizar_chassi",
    "gerar_nome_base_resultado",
    "obter_caminhos_arquivos_saida",
    "obter_caminho_pdf_bi",
    "carregar_resultado_anterior",
    "comparar_auditorias",
    "FrotaScraperEngine",
    "FrotaBoletoDownloader",
    "FrotaRelatorioExporter"
]
