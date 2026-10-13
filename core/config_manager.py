"""
Módulo de Configuração e Persistência de Preferências do Usuário (config.json).
"""
import os
import json
from typing import Any, Dict

ARQUIVO_CONFIG = "config.json"

CONFIG_PADRAO: Dict[str, Any] = {
    "modo_noturno": False,
    "pasta_destino": "",
    "ultimo_arquivo": "",
    "consultar_ipva": True,
    "consultar_lic": True,
    "consultar_multas": True,
    "modo_headless": True,
    "formato_exportacao": "ambos",
    "qtd_workers": 1,
    "som_notificacao": True,
    "arquivos_recentes": []
}


class ConfigManager:
    """Gerencia a leitura e persistência das configurações do usuário."""

    @classmethod
    def carregar(cls) -> Dict[str, Any]:
        """Carrega o config.json se existir, mesclando com os padrões."""
        if not os.path.exists(ARQUIVO_CONFIG):
            return dict(CONFIG_PADRAO)
        try:
            with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as f:
                dados = json.load(f)
            config = dict(CONFIG_PADRAO)
            config.update(dados)
            return config
        except Exception:
            return dict(CONFIG_PADRAO)

    @classmethod
    def salvar(cls, config: Dict[str, Any]) -> None:
        """Salva as configurações atualizadas no config.json."""
        try:
            config_atual = cls.carregar()
            config_atual.update(config)
            with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as f:
                json.dump(config_atual, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

    @classmethod
    def adicionar_arquivo_recente(cls, caminho_arquivo: str) -> None:
        """Registra um arquivo no histórico de recentes (máximo 5 itens)."""
        if not caminho_arquivo or not os.path.exists(caminho_arquivo):
            return
        cfg = cls.carregar()
        recentes = cfg.get("arquivos_recentes", [])
        caminho_abs = os.path.abspath(caminho_arquivo)
        if caminho_abs in recentes:
            recentes.remove(caminho_abs)
        recentes.insert(0, caminho_abs)
        cfg["arquivos_recentes"] = recentes[:5]
        cfg["ultimo_arquivo"] = caminho_abs
        cls.salvar(cfg)
