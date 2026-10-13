"""
Testes Unitários Automatizados para o Pacote Core da Automação de Frotas.
"""
import os
import shutil
import unittest
import pandas as pd
from datetime import datetime

from core.config_manager import ConfigManager, ARQUIVO_CONFIG
from core.data_processor import (
    FrotaDataProcessor,
    formatar_moeda_float,
    formatar_float_moeda,
    obter_estilo_financeiro,
    obter_ultimo_dia_util_mes,
    obter_proximo_dia_util,
    PADRAO_PLACA
)
from core.relatorio_exporter import FrotaRelatorioExporter
from core.boleto_downloader import FrotaBoletoDownloader


class TestCore(unittest.TestCase):

    def setUp(self):
        self.temp_dir = os.path.abspath("temp_test_dir")
        os.makedirs(self.temp_dir, exist_ok=True)

    def tearDown(self):
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_formatacao_moeda(self):
        self.assertEqual(formatar_moeda_float("R$ 188,96"), 188.96)
        self.assertEqual(formatar_moeda_float("1.234,56"), 1234.56)
        self.assertEqual(formatar_moeda_float("R$ 0,00"), 0.0)
        self.assertEqual(formatar_moeda_float("-"), 0.0)
        self.assertEqual(formatar_moeda_float("Falha Detran"), 0.0)

        self.assertEqual(formatar_float_moeda(188.96), "R$ 188,96")
        self.assertEqual(formatar_float_moeda(1234.56), "R$ 1.234,56")

    def test_regras_estilo_financeiro(self):
        # Valor <= 1000: sem bold, sem underline
        bold, under, cor = obter_estilo_financeiro(500.0)
        self.assertFalse(bold)
        self.assertFalse(under)
        self.assertIsNone(cor)

        # Valor > 1000 e <= 1500: bold = True
        bold, under, cor = obter_estilo_financeiro(1200.0)
        self.assertTrue(bold)
        self.assertFalse(under)
        self.assertIsNone(cor)

        # Valor > 1500: bold = True e under = True
        bold, under, cor = obter_estilo_financeiro(1600.0)
        self.assertTrue(bold)
        self.assertTrue(under)
        self.assertIsNone(cor)

        # Licenciamento > 188.96: cor vermelha #DC2626
        bold, under, cor = obter_estilo_financeiro(189.0, is_licenciamento=True)
        self.assertEqual(cor, "#DC2626")

        # Licenciamento > 1500: bold, underline E vermelho
        bold, under, cor = obter_estilo_financeiro(2000.0, is_licenciamento=True)
        self.assertTrue(bold)
        self.assertTrue(under)
        self.assertEqual(cor, "#DC2626")

    def test_validacao_placa_e_dia_util(self):
        self.assertTrue(bool(PADRAO_PLACA.match("THS2C78")))
        self.assertTrue(bool(PADRAO_PLACA.match("ABC1234")))
        self.assertFalse(bool(PADRAO_PLACA.match("INVALIDA")))

        # Teste do último dia útil do mês (outubro de 2026 termina no sábado 31, então último dia útil é sexta 30)
        dt_outubro = datetime(2026, 10, 12)
        ultimo_dia = obter_ultimo_dia_util_mes(dt_outubro)
        self.assertEqual(ultimo_dia, "30/10/2026")

        # Teste para dezembro de 2026 (31/12 é quinta-feira)
        dt_dezembro = datetime(2026, 12, 1)
        self.assertEqual(obter_ultimo_dia_util_mes(dt_dezembro), "31/12/2026")

    def test_config_manager(self):
        cfg = ConfigManager.carregar()
        self.assertIn("modo_noturno", cfg)
        self.assertIn("formato_exportacao", cfg)

        ConfigManager.salvar({"modo_noturno": True, "qtd_workers": 2})
        cfg_atual = ConfigManager.carregar()
        self.assertTrue(cfg_atual["modo_noturno"])
        self.assertEqual(cfg_atual["qtd_workers"], 2)

        # Limpeza
        ConfigManager.salvar({"modo_noturno": False, "qtd_workers": 1})

    def test_exportacao_excel_e_pdf(self):
        df = pd.DataFrame([
            {"Placa": "THS2C78", "RENAVAM": "01415443316", "Valor_Licenciamento": "R$ 1.267,93", "Valor_IPVA": "R$ 0,00", "Valor_Multas": "R$ 3.529,89", "Status": "Com Débitos Pendentes"},
            {"Placa": "SBO1F58", "RENAVAM": "01353629012", "Valor_Licenciamento": "R$ 0,00", "Valor_IPVA": "R$ 1.943,68", "Valor_Multas": "R$ 0,00", "Status": "Com Débitos Pendentes"},
            {"Placa": "TII5F89", "RENAVAM": "01458899249", "Valor_Licenciamento": "R$ 0,00", "Valor_IPVA": "R$ 0,00", "Valor_Multas": "R$ 0,00", "Status": "Sem Débitos"}
        ])

        excel_path = os.path.join(self.temp_dir, "teste_relatorio.xlsx")
        caminho_gerado = FrotaRelatorioExporter.exportar_excel(df, excel_path)
        self.assertTrue(os.path.exists(caminho_gerado))

        pdf_path = os.path.join(self.temp_dir, "teste_relatorio.pdf")
        totais = {"lic": "R$ 1.267,93", "ipva": "R$ 1.943,68", "multas": "R$ 3.529,89", "total": "R$ 6.741,50"}
        caminho_pdf = FrotaRelatorioExporter.exportar_pdf(df, pdf_path, totais, ["Licenciamento", "IPVA", "Multas"])
        self.assertTrue(os.path.exists(caminho_pdf))

    def test_boleto_downloader_naming(self):
        downloader = FrotaBoletoDownloader(self.temp_dir, headless=True)
        self.assertIn("Licenciamento-IPVA", downloader.nome_pasta_boletos)
        self.assertTrue(os.path.isabs(downloader.pasta_destino_final))


if __name__ == "__main__":
    unittest.main()
