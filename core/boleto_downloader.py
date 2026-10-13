"""
Módulo de Download Automatizado de Boletos (Licenciamento Detran-CE e IPVA Sefaz-CE).
Salva e renomeia os boletos no padrão corporativo:
{PLACA} LICENCIAMENTO {DD-MM-YYYY}.pdf
{PLACA} IPVA {DD-MM-YYYY}.pdf
"""
import os
import sys
import time
import shutil
import subprocess
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

from core.data_processor import obter_ultimo_dia_util_mes, obter_proximo_dia_util
from core.logger_config import obter_logger

logger = obter_logger("BoletoDownloader")


class FrotaBoletoDownloader:
    """Realiza o download em lote e renomeação de boletos de Licenciamento e IPVA."""

    def __init__(self, pasta_base_destino: str, headless: bool = True):
        self.pasta_base_destino = pasta_base_destino
        self.headless = headless
        self.data_hoje_str = datetime.now().strftime("%d-%m-%Y")
        
        # Nome da pasta: "Licenciamento-IPVA DD-MM-AAAA" (sem barras para compatibilidade total com o Windows)
        self.nome_pasta_boletos = f"Licenciamento-IPVA ({self.data_hoje_str})"
        self.pasta_destino_final = os.path.join(self.pasta_base_destino, self.nome_pasta_boletos)
        
        self.pasta_temp_downloads = os.path.join(self.pasta_destino_final, ".temp_downloads")
        self.driver: webdriver.Chrome | None = None
        self.wait: WebDriverWait | None = None
        self.cancelar_solicitado = False

    def iniciar(self):
        """Inicializa o navegador com configurações dedicadas para download silencioso de PDFs."""
        os.makedirs(self.pasta_destino_final, exist_ok=True)
        os.makedirs(self.pasta_temp_downloads, exist_ok=True)

        chrome_options = Options()
        if self.headless:
            chrome_options.add_argument("--headless=new")

        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-popup-blocking")
        chrome_options.add_argument("--log-level=3")

        # Configurações do Chrome para salvar PDFs diretamente sem abrir o visualizador interno
        prefs = {
            "download.default_directory": os.path.abspath(self.pasta_temp_downloads),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)

        service = Service(ChromeDriverManager().install())
        if sys.platform == "win32":
            service.creation_flags = subprocess.CREATE_NO_WINDOW

        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.wait = WebDriverWait(self.driver, 15)
        logger.info(f"Downloader de Boletos inicializado. Pasta de destino: {self.pasta_destino_final}")

    def fechar(self):
        """Fecha o navegador e remove a pasta temporária."""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None
            self.wait = None
        
        # Limpa pasta temporária
        if os.path.exists(self.pasta_temp_downloads):
            try:
                shutil.rmtree(self.pasta_temp_downloads, ignore_errors=True)
            except Exception:
                pass

    def _clicar_seguro(self, elemento) -> None:
        """Executa clique seguro com rolagem."""
        try:
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
            elemento.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", elemento)

    def _limpar_pasta_temp(self):
        """Remove arquivos antigos da pasta temporária."""
        if os.path.exists(self.pasta_temp_downloads):
            for f in os.listdir(self.pasta_temp_downloads):
                p = os.path.join(self.pasta_temp_downloads, f)
                try:
                    if os.path.isfile(p):
                        os.remove(p)
                except Exception:
                    pass

    def _aguardar_arquivo_baixado(self, timeout_segundos: int = 15) -> str | None:
        """Monitora a pasta temporária até que um arquivo PDF termine o download (sem .crdownload)."""
        inicio = time.time()
        while time.time() - inicio < timeout_segundos:
            if not os.path.exists(self.pasta_temp_downloads):
                time.sleep(0.5)
                continue
            arquivos = os.listdir(self.pasta_temp_downloads)
            # Verifica se há arquivos que não sejam temporários do Chrome
            prontos = [f for f in arquivos if not f.endswith(".crdownload") and not f.endswith(".tmp")]
            if prontos:
                caminho_pronto = os.path.join(self.pasta_temp_downloads, prontos[0])
                # Confirma que o tamanho do arquivo não está mais mudando
                tam1 = os.path.getsize(caminho_pronto)
                time.sleep(0.5)
                tam2 = os.path.getsize(caminho_pronto)
                if tam1 == tam2 and tam1 > 0:
                    return caminho_pronto
            time.sleep(0.5)
        return None

    def baixar_boleto_licenciamento(self, placa: str, renavam: str) -> tuple[bool, str]:
        """
        Gera e faz o download do boleto de Licenciamento no Detran-CE.
        Retorna (sucesso: bool, mensagem_ou_caminho: str).
        """
        self._limpar_pasta_temp()
        nome_arquivo_final = f"{placa} LICENCIAMENTO {self.data_hoje_str}.pdf"
        destino_final = os.path.join(self.pasta_destino_final, nome_arquivo_final)

        url_detran = "https://sistemas.detran.ce.gov.br/central/#veiculo"
        try:
            self.driver.get(url_detran)

            menu_btn = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Licenciamento') or contains(text(), 'Taxas / Multas')]"))
            )
            self._clicar_seguro(menu_btn)

            self.wait.until(lambda d: len([i for i in d.find_elements(By.TAG_NAME, "input") if i.is_displayed()]) >= 2)
            inputs = [i for i in self.driver.find_elements(By.TAG_NAME, "input") if i.is_displayed()]
            inputs[0].clear()
            inputs[0].send_keys(placa)
            inputs[1].clear()
            inputs[1].send_keys(renavam)

            btn_confirmar = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'CONFIRMAR') or @id='submit-veiculo']"))
            )
            self._clicar_seguro(btn_confirmar)
            time.sleep(1.5)

            # Verifica link de licenciamento
            link_lic = self.driver.find_elements(
                By.XPATH, "//*[contains(text(), 'Clique AQUI se deseja imprimir seu licenciamento.')]"
            )
            if not link_lic:
                return False, "Nenhum boleto de licenciamento em aberto disponível."

            self._clicar_seguro(link_lic[0])
            btn_baixar = self.wait.until(EC.element_to_be_clickable((By.ID, "btn-emitir-licenciamento")))
            time.sleep(0.5)
            self._clicar_seguro(btn_baixar)

            arquivo_baixado = self._aguardar_arquivo_baixado(timeout_segundos=15)
            if arquivo_baixado and os.path.exists(arquivo_baixado):
                shutil.move(arquivo_baixado, destino_final)
                logger.info(f"Boleto de Licenciamento salvo: {destino_final}")
                return True, destino_final
            else:
                return False, "Tempo esgotado aguardando download do Detran."

        except Exception as e:
            logger.error(f"Erro ao baixar licenciamento {placa}: {e}")
            return False, f"Falha: {e}"

    def baixar_dae_ipva(self, placa: str, renavam: str) -> tuple[bool, str]:
        """
        Gera e faz o download do DAE de IPVA na Sefaz-CE.
        Retorna (sucesso: bool, mensagem_ou_caminho: str).
        """
        self._limpar_pasta_temp()
        nome_arquivo_final = f"{placa} IPVA {self.data_hoje_str}.pdf"
        destino_final = os.path.join(self.pasta_destino_final, nome_arquivo_final)

        url_sefaz = "https://ipva.sefaz.ce.gov.br/#/impostos/emitir-dae"
        try:
            self.driver.get(url_sefaz)

            inp_placa = self.wait.until(EC.visibility_of_element_located((
                By.XPATH, "//input[contains(@ng-model, 'placa') or @name='placa' or @formcontrolname='placa' or @id='mat-input-1']"
            )))
            inp_placa.clear()
            inp_placa.send_keys(placa)

            inp_renavam = self.wait.until(EC.visibility_of_element_located((
                By.XPATH, "//input[contains(@ng-model, 'renavam') or @name='renavam' or @formcontrolname='renavam' or @id='mat-input-2']"
            )))
            inp_renavam.clear()
            inp_renavam.send_keys(renavam)

            btn_pesquisar = self.wait.until(EC.element_to_be_clickable((
                By.XPATH, "//button[contains(translate(., 'pesquisar', 'PESQUISAR'), 'PESQUISAR') or contains(@class, 'search-button')]"
            )))
            self._clicar_seguro(btn_pesquisar)
            time.sleep(2.5)

            # Verifica se há débitos não quitados
            if not self.driver.find_elements(By.XPATH, "//*[contains(text(), 'Débitos de IPVA não quitados')]"):
                return False, "Nenhum débito de IPVA em aberto para este veículo."

            # Ajusta campo "Pagar em" para o último dia útil do mês
            try:
                campo_pagar = self.driver.find_elements(By.XPATH, "//mat-form-field[contains(., 'Pagar em')]//input")
                if campo_pagar:
                    data_vencimento = obter_ultimo_dia_util_mes()
                    cp = campo_pagar[0]
                    cp.click()
                    cp.send_keys(Keys.CONTROL + "a")
                    cp.send_keys(Keys.BACKSPACE)
                    cp.send_keys(data_vencimento)
                    cp.send_keys(Keys.TAB)
                    time.sleep(0.5)
                    logger.info(f"Vencimento do IPVA ({placa}) definido para o último dia útil do mês: {data_vencimento}")
            except Exception as e:
                logger.warning(f"Não foi possível ajustar a data de vencimento do IPVA: {e}")

            # Seleciona o checkbox da parcela em aberto
            mat_chk = self.driver.find_elements(By.CSS_SELECTOR, "mat-checkbox")
            if not mat_chk:
                return False, "Não foi possível localizar o seletor da parcela de IPVA."

            ActionChains(self.driver).move_to_element(mat_chk[0]).click().perform()
            time.sleep(1)

            btn_gerar = self.driver.find_element(By.ID, "dae-veiculo-novo")
            if "disabled" in btn_gerar.get_attribute("class") or btn_gerar.get_attribute("disabled"):
                # Tenta segundo clique se necessário
                ActionChains(self.driver).move_to_element(mat_chk[0]).click().perform()
                time.sleep(0.5)

            self._clicar_seguro(btn_gerar)
            time.sleep(2.5)

            # Localiza botão "Imprimir DAE"
            btn_imprimir = self.wait.until(EC.element_to_be_clickable((
                By.XPATH, "//button[contains(., 'Imprimir DAE')]"
            )))
            self._clicar_seguro(btn_imprimir)

            arquivo_baixado = self._aguardar_arquivo_baixado(timeout_segundos=15)
            if arquivo_baixado and os.path.exists(arquivo_baixado):
                shutil.move(arquivo_baixado, destino_final)
                logger.info(f"DAE de IPVA salvo: {destino_final}")
                return True, destino_final
            else:
                return False, "Tempo esgotado aguardando download do DAE."

        except Exception as e:
            logger.error(f"Erro ao baixar IPVA {placa}: {e}")
            return False, f"Falha: {e}"
