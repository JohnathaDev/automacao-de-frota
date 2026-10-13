"""
Testes de Inicialização e Widgets da Interface Gráfica (Tkinter).
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

    def test_janela_download_boletos_init(self):
        df_frota = pd.DataFrame([
            {"Placa": "THS2C78", "RENAVAM": "01415443316"},
            {"Placa": "SBO1F58", "RENAVAM": "01353629012"}
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


if __name__ == "__main__":
    unittest.main()
