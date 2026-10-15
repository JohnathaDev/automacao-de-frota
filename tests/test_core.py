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
    calcular_total_a_pagar_veiculo,
    obter_estilo_financeiro,
    obter_ultimo_dia_util_mes,
    obter_proximo_dia_util,
    PADRAO_PLACA,
    gerar_nome_base_resultado,
    obter_caminhos_arquivos_saida
)
from core.scraper_engine import DetranResultado
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

    def test_calculo_total_a_pagar_sem_duplicidade(self):
        # Cenário 1: Exemplo fornecido pelo usuário:
        # Moto (taxa R$ 188,96) com 2 multas: R$ 293,47 (já vencida) e R$ 234,78 (a vencer com 20% de desconto).
        # Licenciamento = 188,96 + 293,47 = R$ 482,43
        # Multas total = 293,47 + 234,78 = R$ 528,25
        # Total real esperado = Licenciamento (482,43) + Multas a vencer (234,78) = R$ 717,21
        total_com_a_vencer = calcular_total_a_pagar_veiculo(
            lic="R$ 482,43",
            ipva="R$ 0,00",
            multas="R$ 528,25",
            multas_a_vencer="R$ 234,78"
        )
        self.assertEqual(total_com_a_vencer, 717.21)

        # Sem passar multas_a_vencer explicitamente (usando fallback analítico inteligente):
        total_fallback = calcular_total_a_pagar_veiculo(
            lic="R$ 482,43",
            ipva="R$ 0,00",
            multas="R$ 528,25"
        )
        self.assertEqual(total_fallback, 717.21)

        # Cenário 2: Veículo com apenas multas a vencer (Licenciamento na taxa base de R$ 188,96):
        # Licenciamento = R$ 188,96, Multas a vencer = R$ 260,31
        # Total esperado = 188,96 + 260,31 = R$ 449,27
        total_a_vencer_puro = calcular_total_a_pagar_veiculo(
            lic="R$ 188,96",
            ipva="R$ 0,00",
            multas="R$ 260,31",
            multas_a_vencer="R$ 260,31"
        )
        self.assertEqual(total_a_vencer_puro, 449.27)

        # Cenário 3: Veículo com multas apenas vencidas (Licenciamento = R$ 495,31 = 188,96 + 306,35):
        # Multas = R$ 306,35 (vencida), multas_a_vencer = 0
        # Total esperado = R$ 495,31 (não pode somar 306,35 duas vezes!)
        total_so_vencidas = calcular_total_a_pagar_veiculo(
            lic="R$ 495,31",
            ipva="R$ 0,00",
            multas="R$ 306,35",
            multas_a_vencer="R$ 0,00"
        )
        self.assertEqual(total_so_vencidas, 495.31)

        # Cenário 4: Carro (taxa base R$ 220,45) com multa vencida de R$ 200,00 e IPVA de R$ 500,00:
        # Licenciamento = R$ 420,45 (220,45 + 200,00)
        # Total esperado = 420,45 + 500,00 = R$ 920,45
        total_carro = calcular_total_a_pagar_veiculo(
            lic="R$ 420,45",
            ipva="R$ 500,00",
            multas="R$ 200,00",
            multas_a_vencer="R$ 0,00"
        )
        self.assertEqual(total_carro, 920.45)

        # Cenário 5: Veículo com Licenciamento = R$ 0,00, IPVA = R$ 1.943,68, Multas = R$ 0,00:
        total_ipva = calcular_total_a_pagar_veiculo(
            lic="R$ 0,00",
            ipva="R$ 1.943,68",
            multas="R$ 0,00"
        )
        self.assertEqual(total_ipva, 1943.68)

    def test_consolidar_resultados_com_total_a_pagar(self):
        df_original = pd.DataFrame([
            {"Placa": "THS2C78", "RENAVAM": "01415443316"},
            {"Placa": "SBO1F58", "RENAVAM": "01353629012"}
        ])
        lics = ["R$ 482,43", "R$ 0,00"]
        ipvas = ["R$ 0,00", "R$ 1.943,68"]
        multas = ["R$ 528,25", "R$ 0,00"]
        multas_av = ["R$ 234,78", "R$ 0,00"]

        df_consolidado = FrotaDataProcessor.consolidar_resultados(
            df_original, lics, ipvas, multas, 2, resultados_multas_a_vencer=multas_av
        )

        self.assertIn("Total_a_Pagar", df_consolidado.columns)
        self.assertEqual(df_consolidado.loc[0, "Total_a_Pagar"], "R$ 717,21")
        self.assertEqual(df_consolidado.loc[1, "Total_a_Pagar"], "R$ 1.943,68")
        self.assertEqual(df_consolidado.loc[0, "Status"], "Com Débitos Pendentes")

    def test_detran_resultado_compatibilidade(self):
        res = DetranResultado("R$ 482,43", "R$ 528,25", "R$ 234,78", "R$ 293,47")
        # Deve desempacotar como tupla tradicional de 2 valores:
        val_lic, val_multas = res
        self.assertEqual(val_lic, "R$ 482,43")
        self.assertEqual(val_multas, "R$ 528,25")
    def test_boleto_downloader_features(self):
        downloader = FrotaBoletoDownloader(self.temp_dir, headless=True)
        self.assertIn("Licenciamento-IPVA", downloader.nome_pasta_boletos)
        self.assertTrue(hasattr(downloader, "_salvar_pagina_como_pdf"))
        self.assertTrue(hasattr(downloader, "reiniciar_sessao"))
        self.assertTrue(hasattr(downloader, "_capturar_boleto_ou_arquivo"))

    def test_downloader_retry_queue_logic(self):
        # Simula o algoritmo de repetição usado pela GUI
        tarefas = [
            ({"placa": "AAA1111", "renavam": "123", "index": 0}, "lic"),
            ({"placa": "BBB2222", "renavam": "456", "index": 1}, "ipva"),
            ({"placa": "CCC3333", "renavam": "789", "index": 2}, "lic")
        ]
        
        # Simula que BBB2222 falha na primeira tentativa e sucede na segunda
        tentativa = 1
        concluidos = set()
        fila = list(tarefas)
        
        while fila:
            falhas = []
            for item, tipo in fila:
                if item["placa"] == "BBB2222" and tentativa == 1:
                    falhas.append((item, tipo))
                else:
                    concluidos.add((item["index"], tipo))
            fila = falhas
            tentativa += 1
            if tentativa > 5:
                break
                
    def test_pago_e_restricao_impedimento(self):
        # 1. formatar_moeda_float deve converter PAGO e Restrição em 0.0
        self.assertEqual(formatar_moeda_float("PAGO"), 0.0)
        self.assertEqual(formatar_moeda_float("Restrição"), 0.0)
        self.assertEqual(formatar_moeda_float("Restricao"), 0.0)

        # 2. Total a pagar com PAGO e com Restrição
        tot_pago = calcular_total_a_pagar_veiculo(lic="PAGO", ipva="R$ 594,06", multas="R$ 0,00")
        self.assertEqual(tot_pago, 594.06)

        tot_restr = calcular_total_a_pagar_veiculo(lic="Restrição", ipva="R$ 0,00", multas="R$ 1.019,44")
        self.assertEqual(tot_restr, 1019.44)

        # 3. Consolidar resultados com status de Restrição e PAGO
        df_in = pd.DataFrame([
            {"Placa": "POT5I85", "RENAVAM": "01247217938"},
            {"Placa": "THY7G95", "RENAVAM": "01450027048"},
            {"Placa": "ABC1234", "RENAVAM": "01234567890"}
        ])
        lics = ["Restrição", "PAGO", "Restrição"]
        ipvas = ["R$ 0,00", "R$ 594,06", "R$ 0,00"]
        multas = ["R$ 1.019,44", "R$ 0,00", "R$ 0,00"]

        df_out = FrotaDataProcessor.consolidar_resultados(df_in, lics, ipvas, multas, 3)
        self.assertEqual(df_out.loc[0, "Status"], "Com Débitos / Restrição")
        self.assertEqual(df_out.loc[1, "Status"], "Com Débitos Pendentes")
        self.assertEqual(df_out.loc[2, "Status"], "Restrição de Impedimento")
        # IPVA com valor R$ 0,00 registrado como 'PAGO'
        self.assertEqual(df_out.loc[0, "Valor_IPVA"], "PAGO")
        self.assertEqual(df_out.loc[1, "Valor_IPVA"], "R$ 594,06")
        self.assertEqual(df_out.loc[2, "Valor_IPVA"], "PAGO")

        # 4. Estilos no Excel para PAGO (verde) e Restrição (laranja)
        excel_path = os.path.join(self.temp_dir, "teste_estilos_especiais.xlsx")
        FrotaRelatorioExporter.exportar_excel(df_out, excel_path)
        
        import openpyxl
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        # Linha 2 é POT5I85 (Restrição) - coluna C (Valor_Licenciamento)
        c_restr = ws["C2"]
        self.assertEqual(c_restr.value, "Restrição")
        self.assertEqual(c_restr.fill.start_color.rgb, "00FFEDD5")
        self.assertEqual(c_restr.font.color.rgb, "009A3412")
        self.assertTrue(c_restr.font.bold)

        # Linha 2 é POT5I85 - coluna D (Valor_IPVA) deve ser PAGO (verde)
        c_ipva_pago = ws["D2"]
        self.assertEqual(c_ipva_pago.value, "PAGO")
        self.assertEqual(c_ipva_pago.fill.start_color.rgb, "00DCFCE7")
        self.assertEqual(c_ipva_pago.font.color.rgb, "00166534")
        self.assertTrue(c_ipva_pago.font.bold)

        # Linha 3 é THY7G95 (PAGO) - coluna C (Valor_Licenciamento)
        c_pago = ws["C3"]
        self.assertEqual(c_pago.value, "PAGO")
        self.assertEqual(c_pago.fill.start_color.rgb, "00DCFCE7")
        self.assertEqual(c_pago.font.color.rgb, "00166534")
        self.assertTrue(c_pago.font.bold)

    def test_nomenclatura_dinamica_arquivos(self):
        # 1. Casos solicitados pelo usuário
        self.assertEqual(gerar_nome_base_resultado("frota_final_8.xlsx"), "resultados_final_8")
        self.assertEqual(gerar_nome_base_resultado("frota_final_8"), "resultados_final_8")
        self.assertEqual(gerar_nome_base_resultado("frota_ferreira_gomes.xlsx"), "resultados_ferreira_gomes")
        self.assertEqual(gerar_nome_base_resultado("frota_ferreira_gomes"), "resultados_ferreira_gomes")

        # 2. Variações com hífen e espaço
        self.assertEqual(gerar_nome_base_resultado("frota-final-8.xlsx"), "resultados-final-8")
        self.assertEqual(gerar_nome_base_resultado("frota final 8.xlsx"), "resultados final 8")
        self.assertEqual(gerar_nome_base_resultado("frota_ce.xlsx"), "resultados_ce")

        # 3. Nomes sem prefixo frota ou somente frota
        self.assertEqual(gerar_nome_base_resultado("frota.xlsx"), "resultados_frota")
        self.assertEqual(gerar_nome_base_resultado("ferreira_gomes.xlsx"), "resultados_ferreira_gomes")
        self.assertEqual(gerar_nome_base_resultado("final_8.xlsx"), "resultados_final_8")
        self.assertEqual(gerar_nome_base_resultado("resultados_final_8.xlsx"), "resultados_final_8")
        self.assertEqual(gerar_nome_base_resultado(""), "resultados_frota")

        # 4. Geração de caminhos completos (Excel, PDF, Checkpoint)
        excel, pdf, chk = obter_caminhos_arquivos_saida(
            "C:/frotas/frota_final_8.xlsx",
            pasta_destino=self.temp_dir
        )
        self.assertEqual(os.path.basename(excel), "resultados_final_8.xlsx")
        self.assertEqual(os.path.basename(pdf), "resultados_final_8.pdf")
        self.assertEqual(os.path.basename(chk), "~checkpoint_resultados_final_8.xlsx")
        self.assertEqual(os.path.dirname(excel), self.temp_dir)

        # 5. Exportação real de relatório com nome dinâmico
        df = pd.DataFrame([{"Placa": "AAA1111", "RENAVAM": "01234567890", "Total_a_Pagar": "R$ 100,00"}])
        c_excel = FrotaRelatorioExporter.exportar_excel(df, excel)
        self.assertTrue(os.path.exists(c_excel))
        self.assertEqual(c_excel, excel)

        totais = {"lic": "R$ 0,00", "ipva": "R$ 0,00", "multas": "R$ 0,00", "total": "R$ 100,00"}
        c_pdf = FrotaRelatorioExporter.exportar_pdf(df, pdf, totais, ["Licenciamento"], identificador_frota="final_8")
        self.assertTrue(os.path.exists(c_pdf))
        self.assertEqual(c_pdf, pdf)

        # 6. Downloader de Boletos com separação por frota
        dl = FrotaBoletoDownloader(self.temp_dir, headless=True, nome_frota="final_8")
        self.assertIn("final_8", dl.nome_pasta_boletos)
        self.assertIn("Licenciamento-IPVA", dl.nome_pasta_boletos)


if __name__ == "__main__":
    unittest.main()

