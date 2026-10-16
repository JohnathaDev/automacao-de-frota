"""
Testes de Inicialização, Widgets e Design System da Interface Gráfica Principal (Tkinter).
"""
import unittest
import tkinter as tk
import pandas as pd
from automacao_frota_gui import AppConsultaFrotaGUI, JanelaDownloadBoletos


class TestGUI(unittest.TestCase):

    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()  # Não exibe a janela durante os testes

    def tearDown(self):
        try:
            self.root.destroy()
        except Exception:
            pass

    def test_gui_init(self):
        app = AppConsultaFrotaGUI(self.root)
        self.assertIsNotNone(app.tree)
        self.assertIsNotNone(app.btn_iniciar)
        self.assertIsNotNone(app.btn_baixar_boletos)
        self.assertIsNotNone(app.combo_workers)
        self.assertEqual(len(app.luxury_cards), 5)
        self.assertIsNotNone(app.status_orb)

        # Testa alternância de temas (Claro / Noturno)
        estado_inicial = app.modo_noturno
        app._alternar_tema()
        self.assertNotEqual(app.modo_noturno, estado_inicial)
        if app.modo_noturno:
            self.assertIn("Modo Claro", app.btn_tema.cget("text"))
        else:
            self.assertIn("Modo Noturno", app.btn_tema.cget("text"))

        app._alternar_tema()
        self.assertEqual(app.modo_noturno, estado_inicial)

    def test_gui_filtros_pilula(self):
        app = AppConsultaFrotaGUI(self.root)
        app.df_carregado = pd.DataFrame([
            {"Placa": "BRA1A11", "RENAVAM": "01111111111"},
            {"Placa": "BRA2B22", "RENAVAM": "02222222222"}
        ])
        app.col_placa = "Placa"
        app.col_renavam = "RENAVAM"
        app._preencher_tabela_inicial()

        app.resultados_memoria[0] = {"lic": "R$ 1.250,00", "ipva": "PAGO", "multas": "R$ 0,00"}
        app._atualizar_linha_tabela(0, "R$ 1.250,00", "PAGO", "R$ 0,00", "R$ 0,00", 1, 2, 1250.0, 0.0, 0.0, 1250.0)

        app.resultados_memoria[1] = {"lic": "PAGO", "ipva": "PAGO", "multas": "R$ 0,00"}
        app._atualizar_linha_taca = 1
        app._atualizar_linha_tabela(1, "PAGO", "PAGO", "R$ 0,00", "R$ 0,00", 2, 2, 1250.0, 0.0, 0.0, 1250.0)

        app._definir_filtro_pílula("DEBITOS")
        self.assertEqual(len(app.tree.get_children()), 1)

        app._definir_filtro_pílula("PAGOS")
        self.assertEqual(len(app.tree.get_children()), 1)

        app._definir_filtro_pílula("TODOS")
        self.assertEqual(len(app.tree.get_children()), 2)

    def test_janela_download_boletos_init(self):
        df_frota = pd.DataFrame([
            {"Placa": "BRA1A11", "RENAVAM": "01111111111"},
            {"Placa": "BRA2B22", "RENAVAM": "02222222222"}
        ])
        resultados_memoria = {
            0: {"lic": "R$ 1.267,93", "ipva": "R$ 0,00", "multas": "R$ 0,00"},
            1: {"lic": "R$ 0,00", "ipva": "R$ 1.943,68", "multas": "R$ 0,00"}
        }

        janela = JanelaDownloadBoletos(
            parent=self.root,
            df_frota=df_frota,
            col_placa="Placa",
            col_renavam="RENAVAM",
            resultados_memoria=resultados_memoria,
            pasta_destino_base="temp_test_dir",
            modo_noturno=False,
            modo_headless=True
        )

        self.assertEqual(len(janela.itens_boletos), 2)
        self.assertTrue(janela.itens_boletos[0]["tem_lic"])
        self.assertTrue(janela.itens_boletos[1]["tem_ipva"])

        # Testa seleção rápida
        janela._desmarcar_tudo()
        self.assertFalse(janela.itens_boletos[0]["var_lic"].get())
        self.assertFalse(janela.itens_boletos[1]["var_ipva"].get())

        janela._marcar_tudo()
        self.assertTrue(janela.itens_boletos[0]["var_lic"].get())
        self.assertTrue(janela.itens_boletos[1]["var_ipva"].get())

        janela.destroy()

    def test_gui_nomenclatura_dinamica_ao_carregar(self):
        df_frota = pd.DataFrame([{"Placa": "BRA1A11", "RENAVAM": "01111111111"}])
        resultados = {0: {"lic": "R$ 100,00", "ipva": "0"}}
        janela = JanelaDownloadBoletos(
            parent=self.root,
            df_frota=df_frota,
            col_placa="Placa",
            col_renavam="RENAVAM",
            resultados_memoria=resultados,
            pasta_destino_base="temp_test_dir",
            modo_noturno=False,
            modo_headless=True,
            caminho_arquivo="C:/frotas/frota_ferreira_gomes.xlsx"
        )
        self.assertEqual(janela.nome_frota, "ferreira_gomes")
        janela.destroy()


if __name__ == "__main__":
    unittest.main()
