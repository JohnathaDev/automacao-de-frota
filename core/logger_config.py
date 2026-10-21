"""
Configuração centralizada de logs com rotação em disco (logs/automacao_frota.log).
"""
import os
import logging
from logging.handlers import RotatingFileHandler

from core.config_manager import obter_diretorio_base

PASTA_LOGS = os.path.join(obter_diretorio_base(), "logs")
ARQUIVO_LOG = os.path.join(PASTA_LOGS, "automacao_frota.log")


def obter_logger(nome: str = "AutomacaoFrota") -> logging.Logger:
    """Configura e retorna uma instância de Logger com rotação de arquivo (máx 5MB, 3 backups)."""
    os.makedirs(PASTA_LOGS, exist_ok=True)
    logger = logging.getLogger(nome)
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        # Formato padronizado corporativo
        formato = logging.Formatter(
            "[%(asctime)s] [%(levelname)s] [%(name)s]: %(message)s",
            datefmt="%d/%m/%Y %H:%M:%S"
        )

        # Handler de arquivo rotativo (5 MB, até 3 arquivos antigos)
        file_handler = RotatingFileHandler(
            ARQUIVO_LOG, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
        )
        file_handler.setFormatter(formato)
        file_handler.setLevel(logging.INFO)
        logger.addHandler(file_handler)

        # Handler de console
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formato)
        console_handler.setLevel(logging.INFO)
        logger.addHandler(console_handler)

    return logger
