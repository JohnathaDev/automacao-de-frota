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
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

from core.logger_config import obter_logger

logger = obter_logger("ScraperEngine")


class DetranResultado(tuple):
    """
    Subclasse de tuple (lic, multas) para retrocompatibilidade perfeita com desempacotamento de 2 valores:
    val_lic, val_multas = scraper.consultar_detran(...)
    Expõe também multas_a_vencer e multas_vencidas para cálculo exato do Total a Pagar.
    """
    def __new__(cls, val_lic: str, val_multas: str, multas_a_vencer: str = "R$ 0,00", multas_vencidas: str = "R$ 0,00"):
        return super().__new__(cls, (val_lic, val_multas))

    def __init__(self, val_lic: str, val_multas: str, multas_a_vencer: str = "R$ 0,00", multas_vencidas: str = "R$ 0,00"):
        self.lic = val_lic
        self.multas = val_multas
        self.multas_a_vencer = multas_a_vencer
        self.multas_vencidas = multas_vencidas


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
    ) -> DetranResultado:
        """
        Consulta Detran-CE na mesma sessão:
        Extrai Licenciamento e/ou Multas com navegação unificada para velocidade máxima.
        Diferencia multas a vencer (no prazo) de multas já vencidas (embutidas no licenciamento).
        Diferencia veículos regulares (R$ 0,00) de falhas/erros de conexão.
        Possui auto-heal se o navegador travar.
        """
        val_lic = "-" if not consultar_lic else "Falha Detran"
        val_multas = "-" if not consultar_multas else "Falha Detran"
        val_multas_a_vencer = "-" if not consultar_multas else "R$ 0,00"
        val_multas_vencidas = "-" if not consultar_multas else "R$ 0,00"

        if not (consultar_lic or consultar_multas):
            return DetranResultado("-", "-", "-", "-")

        try:
            if not self.driver:
                self.iniciar()

            # Força navegação limpa para resetar o DOM e modais entre veículos
            self.driver.get("https://sistemas.detran.ce.gov.br/central/")
            time.sleep(0.2)
            self.driver.get("https://sistemas.detran.ce.gov.br/central/#veiculo")

            # 1. Clica no menu lateral de Licenciamento / Taxas / Multas
            menu_btn = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Licenciamento') or contains(text(), 'Taxas / Multas')]"))
            )
            self.driver.execute_script("arguments[0].click();", menu_btn)

            # 2. Localiza inputs visíveis (Placa e RENAVAM) e preenche de forma limpa
            self.wait.until(lambda d: len([i for i in d.find_elements(By.TAG_NAME, "input") if i.is_displayed()]) >= 2)
            inputs = [i for i in self.driver.find_elements(By.TAG_NAME, "input") if i.is_displayed()]

            placa_limpa = str(placa).strip().upper()
            renavam_limpo = str(renavam).strip()

            inputs[0].clear()
            inputs[0].send_keys(placa_limpa)
            inputs[1].clear()
            inputs[1].send_keys(renavam_limpo)

            # 3. Clica em CONFIRMAR
            btn_confirmar = self.wait.until(
                EC.element_to_be_clickable((
                    By.XPATH, "//button[contains(text(), 'CONFIRMAR') or @id='submit-veiculo']"
                ))
            )
            self.driver.execute_script("arguments[0].click();", btn_confirmar)
            time.sleep(1.0)

            # 4. Aguarda a estabilização da resposta do Detran pós-confirmação
            def resposta_detran_pronta(d):
                tem_alerta = len(d.find_elements(By.XPATH, "//*[contains(text(), 'não encontrado') or contains(text(), 'Inexistente') or contains(text(), 'Dados incorretos')]")) > 0
                tem_link_lic = len(d.find_elements(By.XPATH, "//*[contains(text(), 'licenciamento') and contains(text(), 'AQUI')] | //*[contains(text(), 'imprimir seu licenciamento')]")) > 0
                tem_link_multa = len(d.find_elements(By.XPATH, "//*[contains(text(), 'CLIQUE AQUI')]")) > 0
                tem_sem_multa = len(d.find_elements(By.XPATH, "//*[contains(text(), 'não possui multas') or contains(text(), 'Nenhuma multa')]")) > 0
                tem_dados = len(d.find_elements(By.XPATH, f"//*[contains(text(), '{placa_limpa}') and not(self::input)]")) > 0
                return tem_alerta or tem_link_lic or tem_link_multa or tem_sem_multa or tem_dados

            try:
                WebDriverWait(self.driver, 10).until(resposta_detran_pronta)
            except Exception:
                pass
            time.sleep(1.0)

            # Checa se houve alerta de veículo inexistente/inválido VISÍVEL
            alertas = [
                a for a in self.driver.find_elements(
                    By.XPATH, "//*[contains(text(), 'não encontrado') or contains(text(), 'Inexistente') or contains(text(), 'Dados incorretos')]"
                ) if a.is_displayed()
            ]
            if alertas:
                return DetranResultado(
                    "Não Localizado" if consultar_lic else "-",
                    "Não Localizado" if consultar_multas else "-",
                    "Não Localizado" if consultar_multas else "-",
                    "Não Localizado" if consultar_multas else "-"
                )

            # Helper interno seguro para fechar o modal compartilhado do Detran-CE
            def fechar_modal_se_aberto():
                try:
                    self.driver.execute_script("""
                        var m = document.getElementById('outros-servicos-modal');
                        if (m && typeof $ !== 'undefined') { $(m).modal('hide'); }
                        var bds = document.querySelectorAll('.modal-backdrop');
                        bds.forEach(function(b){ b.remove(); });
                        document.body.classList.remove('modal-open');
                    """)
                    time.sleep(0.4)
                except Exception:
                    pass

            aits_no_licenciamento = set()

            # 5. Extração de LICENCIAMENTO
            if consultar_lic:
                try:
                    link_lic_els = [
                        l for l in self.driver.find_elements(
                            By.XPATH, "//*[contains(text(), 'licenciamento') and contains(text(), 'AQUI')] | //*[contains(text(), 'imprimir seu licenciamento')]"
                        ) if l.is_displayed()
                    ]
                    if not link_lic_els:
                        val_lic = "PAGO"
                    else:
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'}); arguments[0].click();", link_lic_els[0])
                        wait_modal = WebDriverWait(self.driver, 14)

                        # Aguarda o modal abrir
                        wait_modal.until(lambda d: d.find_element(By.ID, "outros-servicos-modal").is_displayed())
                        time.sleep(1.0)

                        modal_el = self.driver.find_element(By.ID, "outros-servicos-modal")
                        modal_txt = modal_el.text.lower()

                        if any(t in modal_txt for t in ["restricao de impedimento", "restrição de impedimento", "com restricao", "com restrição"]):
                            val_lic = "Restrição"
                        elif any(t in modal_txt for t in ["já licenciado", "ja licenciado", "sem débitos", "sem debitos"]):
                            val_lic = "PAGO"
                        else:
                            el_total_lic = wait_modal.until(EC.visibility_of_element_located((
                                By.XPATH, "//*[@id='outros-servicos-modal']//tr[contains(., 'TOTAL')]/td[last()]"
                            )))
                            texto_lic = el_total_lic.text.strip()
                            from core.data_processor import formatar_moeda_float
                            if not texto_lic or formatar_moeda_float(texto_lic) == 0.0 or texto_lic == "R$ 0,00":
                                val_lic = "PAGO"
                            else:
                                val_lic = texto_lic

                        # Captura possíveis AITs das multas vencidas embutidas no licenciamento
                        import re
                        linhas_lic = self.driver.find_elements(By.XPATH, "//*[@id='outros-servicos-modal']//table//tr")
                        for tr in linhas_lic:
                            texto_tr = tr.text.strip().upper()
                            possiveis_aits = re.findall(r"\b[A-Z0-9]{8,12}\b", texto_tr)
                            for a in possiveis_aits:
                                aits_no_licenciamento.add(a)

                except Exception as e_lic:
                    fechar_modal_se_aberto()
                    if tentativa_atual == 1:
                        raise e_lic
                    logger.debug(f"Aviso na extração de licenciamento ({placa}): {e_lic}")
                    val_lic = "Falha Licenciamento"
                finally:
                    fechar_modal_se_aberto()

            # 6. Extração de MULTAS
            if consultar_multas:
                try:
                    sem_multas = [
                        s for s in self.driver.find_elements(
                            By.XPATH, "//*[contains(text(), 'não possui multas') or contains(text(), 'Nenhuma multa') or contains(text(), 'Não constam multas')]"
                        ) if s.is_displayed()
                    ]
                    if sem_multas:
                        val_multas = "R$ 0,00"
                        val_multas_a_vencer = "R$ 0,00"
                        val_multas_vencidas = "R$ 0,00"
                    else:
                        links_multa = [
                            l for l in self.driver.find_elements(
                                By.XPATH, "//*[contains(text(), 'multas, CLIQUE AQUI') or contains(text(), 'multa, CLIQUE AQUI')]"
                            ) if l.is_displayed()
                        ]
                        if links_multa:
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'}); arguments[0].click();", links_multa[0])
                            wait_modal = WebDriverWait(self.driver, 14)

                            wait_modal.until(EC.visibility_of_element_located((
                                By.XPATH, "//*[@id='outros-servicos-modal']//table//tr"
                            )))
                            time.sleep(0.8)

                            linhas_multa = self.driver.find_elements(By.XPATH, "//*[@id='outros-servicos-modal']//table//tr")
                            soma_multas = 0.0
                            soma_multas_a_vencer = 0.0
                            soma_multas_vencidas = 0.0
                            encontrou_multas = False

                            from core.data_processor import formatar_moeda_float, formatar_float_moeda
                            from datetime import datetime
                            hoje = datetime.now().date()

                            for linha in linhas_multa:
                                tds = linha.find_elements(By.TAG_NAME, "td")
                                if len(tds) >= 6:
                                    txt_valor = tds[-1].text.strip()
                                    if "R$" in txt_valor:
                                        valor_multa = formatar_moeda_float(txt_valor)
                                        soma_multas += valor_multa
                                        encontrou_multas = True

                                        ait_multa = ""
                                        data_venc_str = ""
                                        if len(tds) >= 8:
                                            ait_multa = tds[1].text.strip().upper()
                                            data_venc_str = tds[5].text.strip()
                                        elif len(tds) >= 7:
                                            data_venc_str = tds[4].text.strip()

                                        is_vencida = False
                                        if data_venc_str:
                                            try:
                                                dt_venc = datetime.strptime(data_venc_str, "%d/%m/%Y").date()
                                                if dt_venc < hoje:
                                                    is_vencida = True
                                            except Exception:
                                                pass

                                        if ait_multa and ait_multa in aits_no_licenciamento:
                                            is_vencida = True

                                        if is_vencida:
                                            soma_multas_vencidas += valor_multa
                                        else:
                                            soma_multas_a_vencer += valor_multa

                            if encontrou_multas:
                                val_multas = formatar_float_moeda(soma_multas)
                                val_multas_a_vencer = formatar_float_moeda(soma_multas_a_vencer)
                                val_multas_vencidas = formatar_float_moeda(soma_multas_vencidas)
                            else:
                                el_total = self.driver.find_elements(
                                    By.XPATH, "//*[@id='outros-servicos-modal']//tr[contains(., 'TOTAL')]/td[last()]"
                                )
                                if el_total and "R$" in el_total[0].text:
                                    val_multas = el_total[0].text.strip()
                                    val_multas_a_vencer = val_multas
                                    val_multas_vencidas = "R$ 0,00"
                                else:
                                    val_multas = "R$ 0,00"
                                    val_multas_a_vencer = "R$ 0,00"
                                    val_multas_vencidas = "R$ 0,00"
                        else:
                            val_multas = "R$ 0,00"
                            val_multas_a_vencer = "R$ 0,00"
                            val_multas_vencidas = "R$ 0,00"
                except Exception as e_mult:
                    fechar_modal_se_aberto()
                    if tentativa_atual == 1:
                        raise e_mult
                    logger.debug(f"Aviso na extração de multas ({placa}): {e_mult}")
                    val_multas = "Falha Multas"
                    val_multas_a_vencer = "Falha Multas"
                    val_multas_vencidas = "Falha Multas"
                finally:
                    fechar_modal_se_aberto()

            return DetranResultado(val_lic, val_multas, val_multas_a_vencer, val_multas_vencidas)

        except (WebDriverException, Exception) as e:
            logger.warning(f"Erro ao consultar Detran para {placa} (Tentativa {tentativa_atual}): {e}")
            if tentativa_atual == 1:
                # Auto-heal: reinicia o navegador e tenta mais uma vez
                self.reiniciar()
                return self.consultar_detran(placa, renavam, consultar_lic, consultar_multas, tentativa_atual=2)
            return DetranResultado(
                "Falha Detran" if consultar_lic else "-",
                "Falha Detran" if consultar_multas else "-",
                "Falha Detran" if consultar_multas else "-",
                "Falha Detran" if consultar_multas else "-"
            )

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

            # Se não há tabela de não quitados, encerra como PAGO imediatamente
            if not self.driver.find_elements(By.XPATH, "//*[contains(text(), 'Débitos de IPVA não quitados')]"):
                return "PAGO"

            # 5. Extrai saldo em aberto
            xpath_valor = "//td[contains(., 'Em aberto')]/preceding-sibling::td[contains(., 'R$')][1]"
            el_valor = self.wait.until(EC.visibility_of_element_located((By.XPATH, xpath_valor)))
            txt_ipva = el_valor.text.strip()
            from core.data_processor import formatar_moeda_float
            if not txt_ipva or formatar_moeda_float(txt_ipva) == 0.0 or txt_ipva == "R$ 0,00":
                return "PAGO"
            return txt_ipva

        except (WebDriverException, Exception) as e:
            logger.warning(f"Erro ao consultar Sefaz para {placa} (Tentativa {tentativa_atual}): {e}")
            if tentativa_atual == 1:
                # Auto-heal: reinicia o navegador e tenta mais uma vez
                self.reiniciar()
                return self.consultar_sefaz_ipva(placa, renavam, tentativa_atual=2)
            return "Falha Sefaz"
