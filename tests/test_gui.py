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
            self.root.update_idletasks()
            self.root.update()
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
            {"Placa": "BRA2B22", "RENAVAM": "02222222222"},
            {"Placa": "BRA3C33", "RENAVAM": "03333333333"}
        ])
        app.col_placa = "Placa"
        app.col_renavam = "RENAVAM"
        app._preencher_tabela_inicial()

        app.resultados_memoria[0] = {"lic": "R$ 1.250,00", "ipva": "PAGO", "multas": "R$ 0,00"}
        app._atualizar_linha_tabela(0, "R$ 1.250,00", "PAGO", "R$ 0,00", "R$ 0,00", 1, 3, 1250.0, 0.0, 0.0, 1250.0)

        app.resultados_memoria[1] = {"lic": "PAGO", "ipva": "PAGO", "multas": "R$ 0,00"}
        app._atualizar_linha_tabela(1, "PAGO", "PAGO", "R$ 0,00", "R$ 0,00", 2, 3, 1250.0, 0.0, 0.0, 1250.0)

        app.resultados_memoria[2] = {"lic": "Não Localizado", "ipva": "Não Localizado", "multas": "Não Localizado"}
        app._atualizar_linha_tabela(2, "Não Localizado", "Não Localizado", "Não Localizado", "R$ 0,00", 3, 3, 1250.0, 0.0, 0.0, 1250.0)

        self.assertIn(2, app.indices_em_falha)

        app._definir_filtro_pílula("DEBITOS")
        self.assertEqual(len(app.tree.get_children()), 1)

        app._definir_filtro_pílula("PAGOS")
        self.assertEqual(len(app.tree.get_children()), 1)

        app._definir_filtro_pílula("FALHAS")
        self.assertEqual(len(app.tree.get_children()), 1)

        app._definir_filtro_pílula("TODOS")
        self.assertEqual(len(app.tree.get_children()), 3)

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

    def test_gui_comercial_bi_top10_sem_sobreposicao(self):
        from unittest.mock import patch
        import automacao_frota_gui as gui_com

        def rects_intersect(b1, b2, tol=0):
            return not (
                b1[2] - tol <= b2[0]
                or b2[2] - tol <= b1[0]
                or b1[3] - tol <= b2[1]
                or b2[3] - tol <= b1[1]
            )

        with patch.object(gui_com.AppConsultaFrotaGUI, "_carregar_planilha_automatica"), \
             patch("tkinter.messagebox.askyesno", return_value=False):
            app = gui_com.AppConsultaFrotaGUI(self.root)
            app.df_carregado = pd.DataFrame([
                {
                    "Placa": f"THN{i}B{i:02d}",
                    "RENAVAM": f"{10000000000 + i}",
                    "PROPRIETÁRIO": "FERREIRA GOMES TRANSPORTES E LOGISTICA"
                }
                for i in range(15)
            ])
            app.col_placa = "Placa"
            app.col_renavam = "RENAVAM"
            app._preencher_tabela_inicial()

            for i in range(15):
                val = 4500.0 - (i * 210.5)
                val_str = f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                app.resultados_memoria[i] = {
                    "lic": "R$ 188,36",
                    "ipva": "R$ 540,10",
                    "multas": val_str,
                    "multas_a_vencer": val_str
                }
            app._atualizar_contadores_filtro()

            bi = app.painel_bi
            self.assertEqual(len(bi.dados_bi["top_veiculos"]), 10)

            for bi_w, bi_h in [(980, 230), (1320, 390), (1880, 680)]:
                bi.winfo_width = lambda bw=bi_w: bw
                bi.winfo_height = lambda bh=bi_h: bh
                bi.redesenhar()

                all_items = bi.find_all()
                text_items = [i for i in all_items if bi.type(i) == "text"]
                bar_items = (
                    list(bi.find_withtag("bar_track"))
                    + list(bi.find_withtag("bar_fill"))
                    + list(bi.find_withtag("bar_seg"))
                )

                # Verifica Top 10 renderizado
                self.assertEqual(len(bi.find_withtag("top10_lbl")), 10)
                self.assertEqual(len(bi.find_withtag("top10_val")), 10)

                # Verifica ausência de colisão Texto x Texto
                for idx_a in range(len(text_items)):
                    for idx_b in range(idx_a + 1, len(text_items)):
                        ba = bi.bbox(text_items[idx_a])
                        bb = bi.bbox(text_items[idx_b])
                        self.assertFalse(
                            rects_intersect(ba, bb, tol=1),
                            f"Colisão texto x texto em {bi_w}x{bi_h}: {ba} vs {bb}"
                        )

                # Verifica ausência de colisão Texto x Barra
                for it in text_items:
                    bt = bi.bbox(it)
                    for ibar in bar_items:
                        bbar = bi.bbox(ibar)
                        self.assertFalse(
                            rects_intersect(bt, bbar, tol=0),
                            f"Colisão texto x barra em {bi_w}x{bi_h}: {bt} vs {bbar}"
                        )

    def test_gui_raio_x_omite_chassi_e_repassa_boletos(self):
        from unittest.mock import patch
        import automacao_frota_gui as gui_com

        df_frota = pd.DataFrame([
            {
                "PROPRIETÁRIO": "ALEXCO",
                "PLACA": "PMY0362",
                "RENAVAM": "01077542633",
                "CHASSI": "9BD27805MGB109244"
            }
        ])
        resultados = {0: {"lic": "R$ 199,66", "ipva": "R$ 850,00", "multas": "R$ 0,00"}}

        janela_dl = JanelaDownloadBoletos(
            parent=self.root,
            df_frota=df_frota,
            col_placa="PLACA",
            col_renavam="RENAVAM",
            resultados_memoria=resultados,
            pasta_destino_base="temp_test_dir",
            modo_noturno=False,
            modo_headless=True
        )
        self.assertEqual(len(janela_dl.itens_boletos), 1)
        self.assertEqual(janela_dl.itens_boletos[0]["chassi"], "9BD27805MGB109244")
        janela_dl.destroy()

        with patch.object(gui_com.AppConsultaFrotaGUI, "_carregar_planilha_automatica"), \
             patch.object(gui_com, "JanelaRaioXVeiculo") as mock_raio_x:
            app = gui_com.AppConsultaFrotaGUI(self.root)
            app.df_carregado = df_frota
            app.col_placa = "PLACA"
            app.col_renavam = "RENAVAM"
            app.col_chassi = "CHASSI"
            app.resultados_memoria = resultados

            app._abrir_modal_raio_x(0)
            self.assertTrue(mock_raio_x.called)
            kwargs = mock_raio_x.call_args.kwargs
            self.assertIn("PROPRIETÁRIO", kwargs["dados_extras"])
            self.assertNotIn("CHASSI", kwargs["dados_extras"])


if __name__ == "__main__":
    unittest.main()

