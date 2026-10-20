"""
Motor de Web Scraping para Detran-CE e Sefaz-CE com Selenium.
Inclui resiliência avançada, auto-heal (reinicialização automática do navegador em falha)
e sessão unificada de consulta.
"""
import os
import re
import sys
import time
import shutil
import tempfile
import threading
import subprocess
from datetime import datetime
from selenium.webdriver.chrome.webdriver import WebDriver as ChromeDriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

from core.data_processor import formatar_moeda_float, formatar_float_moeda
from core.logger_config import obter_logger

logger = obter_logger("ScraperEngine")

_CHROMEDRIVER_LOCK = threading.Lock()
_CHROMEDRIVER_PATH: str | None = None


def obter_caminho_chromedriver() -> str | None:
    """Resolve e faz cache do caminho do ChromeDriver de forma thread-safe com múltiplos fallbacks."""
    global _CHROMEDRIVER_PATH
    with _CHROMEDRIVER_LOCK:
        if _CHROMEDRIVER_PATH and os.path.exists(_CHROMEDRIVER_PATH):
            return _CHROMEDRIVER_PATH

        user_home = os.path.expanduser("~")
        caminhos_conhecidos = [
            os.path.join(user_home, r".cache\selenium\chromedriver\win64\155.0.8059.84\chromedriver.exe"),
            os.path.join(user_home, r".wdm\drivers\chromedriver\win64\155.0.8059.84\chromedriver-win64\chromedriver.exe"),
            os.path.join(os.path.dirname(sys.executable), "chromedriver.exe"),
            os.path.join(os.getcwd(), "chromedriver.exe")
        ]
        for c in caminhos_conhecidos:
            if os.path.exists(c):
                _CHROMEDRIVER_PATH = os.path.normpath(c)
                logger.info(f"ChromeDriver localizado no cache do sistema: {_CHROMEDRIVER_PATH}")
                return _CHROMEDRIVER_PATH

        # 2. Tenta download/resolução via ChromeDriverManager
        try:
            p = ChromeDriverManager().install()
            if p and os.path.exists(p):
                if not p.endswith(".exe"):
                    p_dir = os.path.dirname(p)
                    exe_cand = os.path.join(p_dir, "chromedriver.exe")
                    if os.path.exists(exe_cand):
                        p = exe_cand
                _CHROMEDRIVER_PATH = os.path.normpath(p)
                return _CHROMEDRIVER_PATH
        except Exception as e_wdm:
            logger.warning(f"ChromeDriverManager falhou ({e_wdm}). Delegando para o Selenium Manager nativo...")

        # 3. Retorna None para deixar o Selenium 4 usar seu Selenium Manager nativo integrado
        return None


class DetranResultado(tuple):
    """
    Subclasse de tuple (lic, multas) para retrocompatibilidade perfeita com desempacotamento de 2 valores:
    val_lic, val_multas = scraper.consultar_detran(...)
    Expõe também multas_a_vencer, multas_vencidas e usou_chassi para rastreabilidade.
    """
    def __new__(
        cls,
        val_lic: str,
        val_multas: str,
        multas_a_vencer: str = "R$ 0,00",
        multas_vencidas: str = "R$ 0,00",
        usou_chassi: bool = False
    ):
        return super().__new__(cls, (val_lic, val_multas))

    def __init__(
        self,
        val_lic: str,
        val_multas: str,
        multas_a_vencer: str = "R$ 0,00",
        multas_vencidas: str = "R$ 0,00",
        usou_chassi: bool = False
    ):
        self.lic = val_lic
        self.multas = val_multas
        self.multas_a_vencer = multas_a_vencer
        self.multas_vencidas = multas_vencidas
        self.usou_chassi = usou_chassi


class FrotaScraperEngine:
    """
    Executa a automação de consultas no Detran-CE e Sefaz-CE com alta resiliência,
    sessão única unificada para Licenciamento e Multas e tratamento explícito de falhas.
    Suporta auto-heal (recriação automática do WebDriver em caso de crash) e fallback
    automático por Chassi quando o RENAVAM está incorreto ou ausente.
    """

    def __init__(self, headless: bool = True):
        self.headless = headless
        self.driver: ChromeDriver | None = None
        self.wait: WebDriverWait | None = None
        self.temp_user_data_dir: str | None = None

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

        # Perfil temporário isolado por worker para evitar conflitos de sessão/concorrência
        try:
            self.temp_user_data_dir = tempfile.mkdtemp(prefix="chrome_worker_")
            chrome_options.add_argument(f"--user-data-dir={self.temp_user_data_dir}")
        except Exception:
            self.temp_user_data_dir = None

        os.environ["WDM_LOG"] = "0"
        os.environ["WDM_LOG_LEVEL"] = "0"

        caminho_driver = obter_caminho_chromedriver()
        service = Service(executable_path=caminho_driver) if (caminho_driver and os.path.exists(caminho_driver)) else Service()
        if sys.platform == "win32":
            service.creation_flags = subprocess.CREATE_NO_WINDOW

        self.driver = ChromeDriver(service=service, options=chrome_options)
        self.wait = WebDriverWait(self.driver, 15)
        logger.info(f"Instância do ChromeDriver iniciada com sucesso (Headless={self.headless}).")

    def fechar(self):
        """Encerra a instância do navegador e limpa perfil temporário de forma segura."""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None
            self.wait = None
            logger.info("Instância do ChromeDriver encerrada.")

        if self.temp_user_data_dir and os.path.exists(self.temp_user_data_dir):
            try:
                shutil.rmtree(self.temp_user_data_dir, ignore_errors=True)
            except Exception:
                pass
            self.temp_user_data_dir = None

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
        tentativa_atual: int = 1,
        chassi: str = "",
        usar_chassi: bool = False
    ) -> DetranResultado:
        """
        Consulta Detran-CE na mesma sessão:
        Extrai Licenciamento e/ou Multas com navegação unificada para velocidade máxima.
        Se o RENAVAM estiver ausente, inválido ou retornar 'Veículo não encontrado',
        realiza fallback automático pesquisando por Placa + Chassi (caso o chassi esteja disponível).
        """
        val_lic = "-" if not consultar_lic else "Falha Detran"
        val_multas = "-" if not consultar_multas else "Falha Detran"
        val_multas_a_vencer = "-" if not consultar_multas else "R$ 0,00"
        val_multas_vencidas = "-" if not consultar_multas else "R$ 0,00"

        if not (consultar_lic or consultar_multas):
            return DetranResultado("-", "-", "-", "-")

        from core.data_processor import sanitizar_chassi
        placa_limpa = str(placa).strip().upper()
        renavam_limpo = str(renavam).strip()
        chassi_limpo = sanitizar_chassi(chassi)

        renavam_valido = bool(renavam_limpo) and renavam_limpo.isdigit() and int(renavam_limpo) > 0
        if not usar_chassi and not renavam_valido and chassi_limpo:
            usar_chassi = True

        credencial_busca = chassi_limpo if (usar_chassi and chassi_limpo) else renavam_limpo

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

            # 2. Localiza inputs visíveis (Placa e RENAVAM/Chassi) e preenche de forma limpa
            self.wait.until(lambda d: len([i for i in d.find_elements(By.TAG_NAME, "input") if i.is_displayed()]) >= 2)
            inputs = [i for i in self.driver.find_elements(By.TAG_NAME, "input") if i.is_displayed()]

            inputs[0].clear()
            inputs[0].send_keys(placa_limpa)
            inputs[1].clear()
            inputs[1].send_keys(credencial_busca)

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
                tem_alerta = len([
                    a for a in d.find_elements(
                        By.XPATH,
                        "//*[contains(@class, 'msg-validacao') or contains(text(), 'não encontrado') "
                        "or contains(text(), 'Inexistente') or contains(text(), 'Dados incorretos') "
                        "or contains(text(), 'inválido')]"
                    ) if a.is_displayed() and a.text.strip()
                ]) > 0
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

            # Checa se houve alerta de veículo inexistente/inválido VISÍVEL ou se a seção de dados não carregou
            alertas = [
                a for a in self.driver.find_elements(
                    By.XPATH,
                    "//*[contains(@class, 'msg-validacao') or contains(text(), 'não encontrado') "
                    "or contains(text(), 'Inexistente') or contains(text(), 'Dados incorretos') "
                    "or contains(text(), 'inválido')]"
                ) if a.is_displayed() and a.text.strip()
            ]
            tem_secao_veiculo = len(self.driver.find_elements(
                By.XPATH,
                f"//*[contains(text(), 'licenciamento') and contains(text(), 'AQUI')] | "
                f"//*[contains(text(), 'imprimir seu licenciamento')] | "
                f"//*[contains(text(), 'CLIQUE AQUI')] | "
                f"//*[contains(text(), 'não possui multas') or contains(text(), 'Nenhuma multa')] | "
                f"//*[contains(text(), '{placa_limpa}') and not(self::input)]"
            )) > 0

            if alertas or not tem_secao_veiculo:
                if not usar_chassi and chassi_limpo:
                    logger.info(
                        f"Detran-CE: Placa {placa_limpa} não localizada via RENAVAM ({renavam_limpo}). "
                        f"Acionando fallback automático com Placa + Chassi ({chassi_limpo})..."
                    )
                    return self.consultar_detran(
                        placa_limpa,
                        renavam_limpo,
                        consultar_lic,
                        consultar_multas,
                        tentativa_atual=tentativa_atual,
                        chassi=chassi_limpo,
                        usar_chassi=True
                    )
                return DetranResultado(
                    "Não Localizado" if consultar_lic else "-",
                    "Não Localizado" if consultar_multas else "-",
                    "Não Localizado" if consultar_multas else "-",
                    "Não Localizado" if consultar_multas else "-",
                    usou_chassi=usar_chassi
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
                            if not texto_lic or formatar_moeda_float(texto_lic) == 0.0 or texto_lic == "R$ 0,00":
                                val_lic = "PAGO"
                            else:
                                val_lic = texto_lic

                        # Captura possíveis AITs das multas vencidas embutidas no licenciamento
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

            return DetranResultado(val_lic, val_multas, val_multas_a_vencer, val_multas_vencidas, usou_chassi=usar_chassi)

        except (WebDriverException, Exception) as e:
            logger.warning(f"Erro ao consultar Detran para {placa} (Tentativa {tentativa_atual}): {e}")
            if tentativa_atual == 1:
                # Auto-heal: reinicia o navegador e tenta mais uma vez
                self.reiniciar()
                return self.consultar_detran(
                    placa, renavam, consultar_lic, consultar_multas,
                    tentativa_atual=2, chassi=chassi, usar_chassi=usar_chassi
                )
            if not usar_chassi and chassi_limpo:
                logger.info(f"Detran-CE: Tentando fallback com Placa + Chassi após erro para {placa_limpa}...")
                return self.consultar_detran(
                    placa, renavam, consultar_lic, consultar_multas,
                    tentativa_atual=2, chassi=chassi_limpo, usar_chassi=True
                )
            return DetranResultado(
                "Falha Detran" if consultar_lic else "-",
                "Falha Detran" if consultar_multas else "-",
                "Falha Detran" if consultar_multas else "-",
                "Falha Detran" if consultar_multas else "-",
                usou_chassi=usar_chassi
            )

    def consultar_sefaz_ipva(
        self,
        placa: str,
        renavam: str,
        tentativa_atual: int = 1,
        chassi: str = "",
        usar_apenas_chassi: bool = False
    ) -> str:
        """
        Consulta o saldo de IPVA em aberto na Sefaz-CE com detecção precoce de veículos quitados.
        Se o RENAVAM estiver incorreto, ausente ou retornar 'Veículo não encontrado', realiza
        fallback pesquisando apenas pelo Chassi no portal da Sefaz-CE.
        """
        url_sefaz = "https://ipva.sefaz.ce.gov.br/#/impostos/emitir-dae"
        from core.data_processor import sanitizar_chassi
        placa_limpa = str(placa).strip().upper()
        renavam_limpo = str(renavam).strip()
        chassi_limpo = sanitizar_chassi(chassi)

        renavam_valido = bool(renavam_limpo) and renavam_limpo.isdigit() and int(renavam_limpo) > 0
        if not usar_apenas_chassi and not renavam_valido and chassi_limpo:
            usar_apenas_chassi = True

        self.ultimo_ipva_usou_chassi = usar_apenas_chassi

        try:
            if not self.driver:
                self.iniciar()

            if usar_apenas_chassi:
                # Limpa o estado Angular anterior (ex: snackbar 'Veículo não encontrado')
                self.driver.get("about:blank")
                time.sleep(0.15)

            self.driver.get(url_sefaz)

            if usar_apenas_chassi and chassi_limpo:
                # Pesquisa usando APENAS o Chassi no site da Sefaz-CE
                xpath_chassi = "//input[contains(@ng-model, 'chassi') or @name='chassi' or @formcontrolname='chassi' or @id='mat-input-0']"
                inp_chassi = self.wait.until(EC.visibility_of_element_located((By.XPATH, xpath_chassi)))
                inp_chassi.clear()
                inp_chassi.send_keys(chassi_limpo)
            else:
                # 1. Preenche Placa
                xpath_placa = "//input[contains(@ng-model, 'placa') or @name='placa' or @formcontrolname='placa' or @id='mat-input-1']"
                inp_placa = self.wait.until(EC.visibility_of_element_located((By.XPATH, xpath_placa)))
                inp_placa.clear()
                inp_placa.send_keys(placa_limpa)

                # 2. Preenche RENAVAM
                xpath_renavam = "//input[contains(@ng-model, 'renavam') or @name='renavam' or @formcontrolname='renavam' or @id='mat-input-2']"
                inp_renavam = self.wait.until(EC.visibility_of_element_located((By.XPATH, xpath_renavam)))
                inp_renavam.clear()
                inp_renavam.send_keys(renavam_limpo)

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

            # Se acusar não encontrado ou inválido
            if self.driver.find_elements(By.XPATH, "//*[contains(text(), 'não encontrado') or contains(text(), 'inválido') or contains(text(), 'Não há registro')]"):
                if not usar_apenas_chassi and chassi_limpo:
                    logger.info(
                        f"Sefaz-CE: Placa {placa_limpa} / RENAVAM {renavam_limpo} não localizado. "
                        f"Acionando fallback automático apenas pelo Chassi ({chassi_limpo})..."
                    )
                    return self.consultar_sefaz_ipva(
                        placa_limpa,
                        renavam_limpo,
                        tentativa_atual=tentativa_atual,
                        chassi=chassi_limpo,
                        usar_apenas_chassi=True
                    )
                return "Não Localizado"

            # Se não há tabela de não quitados, encerra como PAGO imediatamente
            if not self.driver.find_elements(By.XPATH, "//*[contains(text(), 'Débitos de IPVA não quitados')]"):
                return "PAGO"

            # 5. Extrai saldo em aberto
            xpath_valor = "//td[contains(., 'Em aberto')]/preceding-sibling::td[contains(., 'R$')][1]"
            el_valor = self.wait.until(EC.visibility_of_element_located((By.XPATH, xpath_valor)))
            txt_ipva = el_valor.text.strip()
            if not txt_ipva or formatar_moeda_float(txt_ipva) == 0.0 or txt_ipva == "R$ 0,00":
                return "PAGO"
            return txt_ipva

        except (WebDriverException, Exception) as e:
            logger.warning(f"Erro ao consultar Sefaz para {placa} (Tentativa {tentativa_atual}): {e}")
            if tentativa_atual == 1:
                # Auto-heal: reinicia o navegador e tenta mais uma vez
                self.reiniciar()
                return self.consultar_sefaz_ipva(
                    placa, renavam, tentativa_atual=2, chassi=chassi, usar_apenas_chassi=usar_apenas_chassi
                )
            if not usar_apenas_chassi and chassi_limpo:
                logger.info(f"Sefaz-CE: Tentando fallback pelo Chassi após erro para {placa_limpa}...")
                return self.consultar_sefaz_ipva(
                    placa, renavam, tentativa_atual=2, chassi=chassi_limpo, usar_apenas_chassi=True
                )
            return "Falha Sefaz"
