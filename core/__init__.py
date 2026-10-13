"""
Pacote Central (Core) do Sistema de Automação de Frotas CE.
"""
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
from core.relatorio_exporter import FrotaRelatorioExporter

__all__ = [
    "ConfigManager",
    "obter_logger",
    "FrotaDataProcessor",
    "formatar_moeda_float",
    "formatar_float_moeda",
    "obter_estilo_financeiro",
    "obter_ultimo_dia_util_mes",
    "obter_proximo_dia_util",
    "PADRAO_PLACA",
    "FrotaScraperEngine",
    "FrotaBoletoDownloader",
    "FrotaRelatorioExporter"
]
