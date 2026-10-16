"""
Módulo de Download Automatizado de Boletos (Licenciamento Detran-CE e IPVA Sefaz-CE).
Salva e renomeia os boletos no padrão corporativo:
{PLACA} LICENCIAMENTO {DD-MM-YYYY}.pdf
{PLACA} IPVA {DD-MM-YYYY}.pdf
Suporta captura por download direto e renderização via 'Salvar como PDF' (Chrome CDP).
"""
import os
import re
import sys
import time
import base64
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

    def __init__(self, pasta_base_destino: str, headless: bool = True, nome_frota: str = ""):
        self.pasta_base_destino = pasta_base_destino
        self.headless = headless
        self.nome_frota = nome_frota.strip()
        self.data_hoje_str = datetime.now().strftime("%d-%m-%Y")
        
        # Nome da pasta: inclui identificação da frota se fornecida para evitar colisão entre frotas
        if self.nome_frota:
            identificador_limpo = re.sub(r'[\\/*?:"<>|]', "", self.nome_frota).strip()
            self.nome_pasta_boletos = f"Licenciamento-IPVA ({identificador_limpo} - {self.data_hoje_str})"
        else:
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

    def reiniciar_sessao(self):
        """Reinicia a sessão do navegador para garantir um estado limpo em caso de falha de conexão."""
        logger.info("Reiniciando sessão do navegador para reprocessamento limpo...")
        self.fechar()
        self.iniciar()

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
                time.sleep(0.4)
                continue
            arquivos = os.listdir(self.pasta_temp_downloads)
            prontos = [f for f in arquivos if not f.endswith(".crdownload") and not f.endswith(".tmp")]
            if prontos:
                caminho_pronto = os.path.join(self.pasta_temp_downloads, prontos[0])
                try:
                    tam1 = os.path.getsize(caminho_pronto)
                    time.sleep(0.4)
                    tam2 = os.path.getsize(caminho_pronto)
                    if tam1 == tam2 and tam1 > 0:
                        return caminho_pronto
                except Exception:
                    pass
            time.sleep(0.4)
        return None

    def _salvar_pagina_como_pdf(self, destino_final: str) -> bool:
        """
        Salva a página atualmente em foco no navegador como PDF via Chrome DevTools Protocol (CDP).
        Equivale exatamente a selecionar 'Salvar como PDF' na caixa de diálogo de impressão do navegador.
        """
        try:
            pdf_res = self.driver.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True,
                "preferCSSPageSize": True
            })
            if pdf_res and "data" in pdf_res:
                pdf_bytes = base64.b64decode(pdf_res["data"])
                if len(pdf_bytes) > 1000:
                    with open(destino_final, "wb") as f:
                        f.write(pdf_bytes)
                    logger.info(f"Página salva como PDF via 'Salvar como PDF' ({len(pdf_bytes)} bytes): {destino_final}")
                    return True
        except Exception as e:
            logger.warning(f"Erro ao salvar página como PDF via CDP: {e}")
        return False

    def _fechar_janelas_extras(self, janela_principal: str):
        """Fecha quaisquer abas/janelas extras abertas pelo Detran/Sefaz e restaura o foco na principal."""
        try:
            todas = self.driver.window_handles
            for handle in todas:
                if handle != janela_principal:
                    try:
                        self.driver.switch_to.window(handle)
                        self.driver.close()
                    except Exception:
                        pass
            self.driver.switch_to.window(janela_principal)
        except Exception as e:
            logger.debug(f"Aviso ao restaurar janela principal: {e}")
            try:
                if self.driver.window_handles:
                    self.driver.switch_to.window(self.driver.window_handles[0])
            except Exception:
                pass

    def _capturar_boleto_ou_arquivo(self, destino_final: str, janela_principal: str, timeout_espera: int = 15) -> bool:
        """
        Captura o documento gerado:
        1. Tenta identificar se o arquivo foi baixado fisicamente na pasta de downloads.
        2. Caso tenha sido aberta uma página de impressão/extrato (ex: /gerar_boleto ou window.print()),
           salva o conteúdo exatamente como 'Salvar como PDF' via Chrome CDP.
        3. Restaura o foco na janela principal original.
        """
        inicio = time.time()

        # 1. Aguarda brevemente abertura de nova janela (caso ocorra)
        while time.time() - inicio < 4:
            if len(self.driver.window_handles) > 1:
                break
            time.sleep(0.3)

        todas_janelas = self.driver.window_handles
        janela_boleto = todas_janelas[-1] if len(todas_janelas) > 1 else janela_principal
        if janela_boleto != self.driver.current_window_handle:
            self.driver.switch_to.window(janela_boleto)

        # 2. Verifica se o Chrome já baixou o arquivo físico diretamente
        arquivo_baixado = self._aguardar_arquivo_baixado(timeout_segundos=2)
        if arquivo_baixado and os.path.exists(arquivo_baixado):
            shutil.move(arquivo_baixado, destino_final)
            self._fechar_janelas_extras(janela_principal)
            return True

        # 3. Monitora se a nova janela carregou o boleto/extrato para 'Salvar como PDF'
        while time.time() - inicio < timeout_espera:
            arquivo_baixado = self._aguardar_arquivo_baixado(timeout_segundos=1)
            if arquivo_baixado and os.path.exists(arquivo_baixado):
                shutil.move(arquivo_baixado, destino_final)
                self._fechar_janelas_extras(janela_principal)
                return True

            try:
                url_atual = self.driver.current_url.lower()
                corpo = self.driver.page_source.lower()

                # Verifica se a página contém extrato de taxas/multas ou boleto
                if (
                    "gerar_boleto" in url_atual
                    or "extrato" in corpo
                    or "arrecadação" in corpo
                    or "arrecadacao" in corpo
                    or "dae sefaz" in corpo
                    or "taxas" in corpo
                ):
                    time.sleep(1.0)  # Aguarda estabilização dos estilos e código de barras
                    if self._salvar_pagina_como_pdf(destino_final):
                        self._fechar_janelas_extras(janela_principal)
                        return True
            except Exception:
                pass

            time.sleep(0.5)

        # 4. Fallback final via CDP caso haja qualquer conteúdo renderizado
        if self._salvar_pagina_como_pdf(destino_final):
            self._fechar_janelas_extras(janela_principal)
            return True

        self._fechar_janelas_extras(janela_principal)
        return False

    def baixar_boleto_licenciamento(self, placa: str, renavam: str) -> tuple[bool, str]:
        """
        Gera e faz o download do boleto de Licenciamento no Detran-CE.
        Trata casos padrão e casos com débitos de IPVA em aberto (que abrem extrato para impressão),
        salvando como PDF fiel através de renderização nativa de impressão ('Salvar como PDF').
        Retorna (sucesso: bool, mensagem_ou_caminho: str).
        """
        self._limpar_pasta_temp()
        nome_arquivo_final = f"{placa} LICENCIAMENTO {self.data_hoje_str}.pdf"
        destino_final = os.path.join(self.pasta_destino_final, nome_arquivo_final)

        url_detran = "https://sistemas.detran.ce.gov.br/central/#veiculo"
        try:
            self.driver.get(url_detran)
            janela_principal = self.driver.current_window_handle

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

            # Verifica se há link para emissão de licenciamento
            link_lic = self.driver.find_elements(
                By.XPATH, "//*[contains(text(), 'Clique AQUI se deseja imprimir seu licenciamento.')]"
            )
            if not link_lic:
                return False, "Nenhum boleto de licenciamento em aberto disponível."

            self._clicar_seguro(link_lic[0])
            time.sleep(1.5)

            # Caso haja débito de IPVA em aberto, o modal exibe o campo data_vencimento
            try:
                campos_data = self.driver.find_elements(By.XPATH, "//input[@id='data_vencimento' or @name='data_vencimento']")
                for cd in campos_data:
                    if cd.is_displayed():
                        data_limite = obter_ultimo_dia_util_mes()
                        cd.clear()
                        cd.send_keys(data_limite)
                        time.sleep(0.3)
                        logger.info(f"Vencimento do Licenciamento ({placa}) ajustado: {data_limite}")
                        break
            except Exception as e:
                logger.debug(f"Aviso ao verificar campo data_vencimento: {e}")

            # Localiza o botão de emissão
            botoes_emitir = self.driver.find_elements(
                By.XPATH,
                "//input[@value='Emitir Taxa(s)'] | //button[@id='btn-emitir-licenciamento'] | "
                "//input[@id='btn-emitir-licenciamento'] | //input[@type='submit' and contains(@value, 'Emitir')] | "
                "//button[contains(., 'Emitir')]"
            )
            if not botoes_emitir:
                return False, "Botão de emissão de licenciamento não encontrado."

            btn_emitir = next((b for b in botoes_emitir if b.is_displayed()), botoes_emitir[0])
            self.driver.execute_script("arguments[0].click();", btn_emitir)

            # Captura via download direto ou 'Salvar como PDF' (CDP)
            sucesso = self._capturar_boleto_ou_arquivo(destino_final, janela_principal, timeout_espera=14)
            if sucesso and os.path.exists(destino_final) and os.path.getsize(destino_final) > 1000:
                logger.info(f"Boleto de Licenciamento salvo com sucesso: {destino_final}")
                return True, destino_final
            else:
                return False, "Não foi possível gerar ou salvar o boleto do Detran."

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
            janela_principal = self.driver.current_window_handle

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

            # Ajusta campo 'Pagar em' para o último dia útil do mês
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
                ActionChains(self.driver).move_to_element(mat_chk[0]).click().perform()
                time.sleep(0.5)

            self._clicar_seguro(btn_gerar)
            time.sleep(2.5)

            # Localiza botão 'Imprimir DAE'
            btn_imprimir = self.wait.until(EC.element_to_be_clickable((
                By.XPATH, "//button[contains(., 'Imprimir DAE')]"
            )))
            self._clicar_seguro(btn_imprimir)

            # Captura via download direto ou 'Salvar como PDF' (CDP)
            sucesso = self._capturar_boleto_ou_arquivo(destino_final, janela_principal, timeout_espera=14)
            if sucesso and os.path.exists(destino_final) and os.path.getsize(destino_final) > 1000:
                logger.info(f"DAE de IPVA salvo com sucesso: {destino_final}")
                return True, destino_final
            else:
                return False, "Tempo esgotado aguardando download ou geração do DAE."

        except Exception as e:
            logger.error(f"Erro ao baixar IPVA {placa}: {e}")
            return False, f"Falha: {e}"
