"""
Motor de Web Scraping para Detran-CE e Sefaz-CE com Selenium.
Inclui resiliência avançada, auto-heal (reinicialização automática do navegador em falha)
e sessão unificada de consulta.
"""
import os
import sys
import time
import subprocess
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

from core.logger_config import obter_logger

logger = obter_logger("ScraperEngine")


class FrotaScraperEngine:
    """
    Executa a automação de consultas no Detran-CE e Sefaz-CE com alta resiliência,
    sessão única unificada para Licenciamento e Multas e tratamento explícito de falhas.
    Suporta auto-heal (recriação automática do WebDriver em caso de crash).
    """

    def __init__(self, headless: bool = True):
        self.headless = headless
        self.driver: webdriver.Chrome | None = None
        self.wait: WebDriverWait | None = None

    def iniciar(self):
        """Inicializa o WebDriver com configurações silenciosas e robustas."""
        chrome_options = Options()
        if self.headless:
            chrome_options.add_argument("--headless=new")

        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        )

        os.environ["WDM_LOG"] = "0"
        os.environ["WDM_LOG_LEVEL"] = "0"

        service = Service(ChromeDriverManager().install())
        if sys.platform == "win32":
            service.creation_flags = subprocess.CREATE_NO_WINDOW

        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.wait = WebDriverWait(self.driver, 15)
        logger.info(f"Instância do ChromeDriver iniciada com sucesso (Headless={self.headless}).")

    def fechar(self):
        """Encerra a instância do navegador de forma segura."""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None
            self.wait = None
            logger.info("Instância do ChromeDriver encerrada.")

    def reiniciar(self):
        """Auto-heal: reinicia a instância do navegador após crash ou travamento."""
        logger.warning("Auto-heal ativado: reiniciando instância do ChromeDriver...")
        self.fechar()
        time.sleep(1)
        self.iniciar()

    def _clicar_seguro(self, elemento) -> None:
        """Executa o clique com scrollIntoView e fallback JavaScript."""
        try:
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
            elemento.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", elemento)

    def consultar_detran(
        self,
        placa: str,
        renavam: str,
        consultar_lic: bool,
        consultar_multas: bool,
        tentativa_atual: int = 1
    ) -> tuple[str, str]:
        """
        Consulta Detran-CE na mesma sessão:
        Extrai Licenciamento e/ou Multas com navegação unificada para velocidade máxima.
        Diferencia veículos regulares (R$ 0,00) de falhas/erros de conexão.
        Possui auto-heal se o navegador travar.
        """
        val_lic = "-" if not consultar_lic else "Falha Detran"
        val_multas = "-" if not consultar_multas else "Falha Detran"

        if not (consultar_lic or consultar_multas):
            return "-", "-"

        url_detran = "https://sistemas.detran.ce.gov.br/central/#veiculo"
        try:
            if not self.driver:
                self.iniciar()

            self.driver.get(url_detran)

            # 1. Clica no menu lateral de Licenciamento (ou Taxas/Multas)
            menu_btn = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Licenciamento') or contains(text(), 'Taxas / Multas')]"))
            )
            self._clicar_seguro(menu_btn)

            # 2. Localiza inputs visíveis (Placa e RENAVAM)
            self.wait.until(lambda d: len([i for i in d.find_elements(By.TAG_NAME, "input") if i.is_displayed()]) >= 2)
            inputs = [i for i in self.driver.find_elements(By.TAG_NAME, "input") if i.is_displayed()]

            inputs[0].clear()
            inputs[0].send_keys(str(placa).strip().upper())
            inputs[1].clear()
            inputs[1].send_keys(str(renavam).strip())

            # 3. Clica em CONFIRMAR
            btn_confirmar = self.wait.until(
                EC.presence_of_element_located((
                    By.XPATH, "//button[contains(text(), 'CONFIRMAR') or @id='submit-veiculo']"
                ))
            )
            self._clicar_seguro(btn_confirmar)
            time.sleep(1.2)

            # Checa se houve alerta de veículo inexistente/inválido
            alertas = self.driver.find_elements(
                By.XPATH, "//*[contains(text(), 'não encontrado') or contains(text(), 'Inexistente') or contains(text(), 'Dados incorretos')]"
            )
            if alertas:
                return ("Não Localizado" if consultar_lic else "-"), ("Não Localizado" if consultar_multas else "-")

            # Helper interno para fechar o modal compartilhado do Detran-CE
            def fechar_modal_se_aberto():
                try:
                    fechar_btns = self.driver.find_elements(
                        By.XPATH, "//*[@id='outros-servicos-modal']//button[contains(text(), 'Fechar') or contains(@class, 'close') or @data-dismiss='modal']"
                    )
                    for b in fechar_btns:
                        if b.is_displayed():
                            self._clicar_seguro(b)
                            time.sleep(0.5)
                            break
                except Exception:
                    pass

            # 4. Extração de LICENCIAMENTO (Executado PRIMEIRO em modal limpo)
            if consultar_lic:
                try:
                    link_lic_els = self.driver.find_elements(
                        By.XPATH, "//*[contains(text(), 'Clique AQUI se deseja imprimir seu licenciamento.')]"
                    )
                    if not link_lic_els:
                        val_lic = "R$ 0,00"
                    else:
                        self._clicar_seguro(link_lic_els[0])
                        wait_modal = WebDriverWait(self.driver, 10)

                        # Aguarda o modal carregar especificamente o conteúdo do Licenciamento
                        wait_modal.until(EC.visibility_of_element_located((
                            By.XPATH, "//*[@id='outros-servicos-modal']//*[contains(text(), 'licenciamento') or contains(text(), 'Licenciamento')]"
                        )))
                        time.sleep(0.8)

                        el_total_lic = wait_modal.until(EC.visibility_of_element_located((
                            By.XPATH, "//*[@id='outros-servicos-modal']//tr[contains(., 'TOTAL')]/td[last()]"
                        )))
                        texto_lic = el_total_lic.text.strip()
                        val_lic = texto_lic if texto_lic else "R$ 0,00"
                except TimeoutException:
                    val_lic = "R$ 0,00"
                except Exception:
                    val_lic = "Falha Licenciamento"
                finally:
                    fechar_modal_se_aberto()

            # 5. Extração de MULTAS (Executado DEPOIS com modal anterior fechado)
            if consultar_multas:
                try:
                    sem_multas = self.driver.find_elements(
                        By.XPATH, "//*[contains(text(), 'não possui multas') or contains(text(), 'Nenhuma multa') or contains(text(), 'Não constam multas')]"
                    )
                    if sem_multas:
                        val_multas = "R$ 0,00"
                    else:
                        links_multa = self.driver.find_elements(
                            By.XPATH, "//*[contains(text(), 'multas, CLIQUE AQUI') or contains(text(), 'multa, CLIQUE AQUI')]"
                        )
                        if links_multa:
                            self._clicar_seguro(links_multa[0])
                            wait_modal = WebDriverWait(self.driver, 10)

                            wait_modal.until(EC.visibility_of_element_located((
                                By.XPATH, "//*[@id='outros-servicos-modal']//table//tr"
                            )))
                            time.sleep(0.8)

                            linhas_multa = self.driver.find_elements(By.XPATH, "//*[@id='outros-servicos-modal']//table//tr")
                            soma_multas = 0.0
                            encontrou_multas = False

                            from core.data_processor import formatar_moeda_float, formatar_float_moeda

                            for linha in linhas_multa:
                                tds = linha.find_elements(By.TAG_NAME, "td")
                                if len(tds) >= 6:
                                    txt_valor = tds[-1].text.strip()
                                    if "R$" in txt_valor:
                                        soma_multas += formatar_moeda_float(txt_valor)
                                        encontrou_multas = True

                            if encontrou_multas:
                                val_multas = formatar_float_moeda(soma_multas)
                            else:
                                el_total = self.driver.find_elements(
                                    By.XPATH, "//*[@id='outros-servicos-modal']//tr[contains(., 'TOTAL')]/td[last()]"
                                )
                                if el_total and "R$" in el_total[0].text:
                                    val_multas = el_total[0].text.strip()
                                else:
                                    val_multas = "R$ 0,00"
                        else:
                            val_multas = "R$ 0,00"
                except TimeoutException:
                    val_multas = "R$ 0,00"
                except Exception:
                    val_multas = "Falha Multas"
                finally:
                    fechar_modal_se_aberto()

            return val_lic, val_multas

        except (WebDriverException, Exception) as e:
            logger.warning(f"Erro ao consultar Detran para {placa} (Tentativa {tentativa_atual}): {e}")
            if tentativa_atual == 1:
                # Auto-heal: reinicia o navegador e tenta mais uma vez
                self.reiniciar()
                return self.consultar_detran(placa, renavam, consultar_lic, consultar_multas, tentativa_atual=2)
            return ("Falha Detran" if consultar_lic else "-"), ("Falha Detran" if consultar_multas else "-")

    def consultar_sefaz_ipva(self, placa: str, renavam: str, tentativa_atual: int = 1) -> str:
        """
        Consulta o saldo de IPVA em aberto na Sefaz-CE com detecção precoce de veículos quitados.
        Diferencia veículos regulares (R$ 0,00) de falhas/erros de conexão.
        Possui auto-heal se o navegador travar.
        """
        url_sefaz = "https://ipva.sefaz.ce.gov.br/#/impostos/emitir-dae"
        try:
            if not self.driver:
                self.iniciar()

            self.driver.get(url_sefaz)

            # 1. Preenche Placa
            xpath_placa = "//input[contains(@ng-model, 'placa') or @name='placa' or @formcontrolname='placa' or @id='mat-input-1']"
            inp_placa = self.wait.until(EC.visibility_of_element_located((By.XPATH, xpath_placa)))
            inp_placa.clear()
            inp_placa.send_keys(str(placa).strip().upper())

            # 2. Preenche RENAVAM
            xpath_renavam = "//input[contains(@ng-model, 'renavam') or @name='renavam' or @formcontrolname='renavam' or @id='mat-input-2']"
            inp_renavam = self.wait.until(EC.visibility_of_element_located((By.XPATH, xpath_renavam)))
            inp_renavam.clear()
            inp_renavam.send_keys(str(renavam).strip())

            # 3. Clica em Pesquisar
            xpath_btn = "//button[contains(translate(., 'pesquisar', 'PESQUISAR'), 'PESQUISAR') or contains(@class, 'search-button')]"
            btn_pesquisar = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn)))
            self._clicar_seguro(btn_pesquisar)

            # 4. Aguarda carregar resposta (débitos em aberto OU débitos quitados OU aviso de erro)
            def resposta_pronta(d):
                tem_aberto = d.find_elements(By.XPATH, "//*[contains(text(), 'Débitos de IPVA não quitados')]")
                quitado = d.find_elements(
                    By.XPATH, "//*[contains(text(), 'Débitos de IPVA Regularizados/Quitados') or contains(text(), 'Nenhum débito')]"
                )
                erro = d.find_elements(
                    By.XPATH, "//*[contains(text(), 'não encontrado') or contains(text(), 'inválido') or contains(text(), 'Não há registro')]"
                )
                return tem_aberto or quitado or erro

            self.wait.until(resposta_pronta)

            # Se acusar não encontrado
            if self.driver.find_elements(By.XPATH, "//*[contains(text(), 'não encontrado') or contains(text(), 'Não há registro')]"):
                return "Não Localizado"

            # Se não há tabela de não quitados, encerra como R$ 0,00 imediatamente
            if not self.driver.find_elements(By.XPATH, "//*[contains(text(), 'Débitos de IPVA não quitados')]"):
                return "R$ 0,00"

            # 5. Extrai saldo em aberto
            xpath_valor = "//td[contains(., 'Em aberto')]/preceding-sibling::td[contains(., 'R$')][1]"
            el_valor = self.wait.until(EC.visibility_of_element_located((By.XPATH, xpath_valor)))
            txt_ipva = el_valor.text.strip()
            return txt_ipva if txt_ipva else "R$ 0,00"

        except (WebDriverException, Exception) as e:
            logger.warning(f"Erro ao consultar Sefaz para {placa} (Tentativa {tentativa_atual}): {e}")
            if tentativa_atual == 1:
                # Auto-heal: reinicia o navegador e tenta mais uma vez
                self.reiniciar()
                return self.consultar_sefaz_ipva(placa, renavam, tentativa_atual=2)
            return "Falha Sefaz"
