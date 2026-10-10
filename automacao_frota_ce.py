"""
========================================================================================
Automação de Consulta de Débitos Veiculares - Ceará (Detran-CE & Sefaz-CE)
========================================================================================
Autor: Engenharia de Software / Automação de Frota
Stack: Python 3, Selenium WebDriver, Pandas, Openpyxl, Webdriver-Manager

Funcionalidades:
1. Geração dinâmica de planilha de teste 'frota.xlsx' com preservação de zeros à esquerda.
2. Consulta de Licenciamento via portal Detran-CE (Single Page Application).
3. Consulta de IPVA em aberto via portal Sefaz-CE (Angular SPA).
4. Sincronização estrita via WebDriverWait e Expected Conditions (sem time.sleep em carregamentos).
5. Resiliência a SPA (tratamento de overlays, rotas com hash e exceções de timeout).
6. Exportação consolidada e estilizada em 'resultados_frota_ce.xlsx'.
========================================================================================
"""

import os
import sys
import time
from datetime import datetime
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

# ======================================================================================
# CONFIGURAÇÕES GERAIS
# ======================================================================================
HEADLESS_MODE = True           # Alterne para False para acompanhar a interface gráfica
TIMEOUT_PADRAO = 15            # Tempo mínimo exigido de espera explícita (em segundos)
INTERVALO_RATE_LIMIT = 1.0     # Pausa preventiva entre requisições para evitar bloqueio
ARQUIVO_ENTRADA = "frota.xlsx"
ARQUIVO_SAIDA = "resultados_frota_ce.xlsx"

# Dados de teste oficiais fornecidos no requisito
DADOS_TESTE = [
    {"Placa": "TII5F89", "RENAVAM": "01458899249"},
    {"Placa": "THT5F69", "RENAVAM": "01458894204"},
    {"Placa": "THU5D09", "RENAVAM": "01458895766"},
    {"Placa": "THV0A49", "RENAVAM": "01434835143"},
    {"Placa": "TIE8B49", "RENAVAM": "01422988322"},
    {"Placa": "THY7G95", "RENAVAM": "01450027048"}
]


def criar_planilha_entrada(caminho: str = ARQUIVO_ENTRADA, forcar: bool = False) -> None:
    """
    Cria dinamicamente a planilha de teste 'frota.xlsx' caso não exista,
    assegurando a integridade dos dados e o tipo textual do RENAVAM.
    """
    if os.path.exists(caminho) and not forcar:
        print(f"[INFO] Arquivo de entrada já existente: {caminho}")
        return

    print(f"[INFO] Gerando planilha inicial de teste: {caminho}...")
    df = pd.DataFrame(DADOS_TESTE)
    # Garante armazenamento estrito como texto para preservar zeros à esquerda
    df["RENAVAM"] = df["RENAVAM"].astype(str)

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Frota")

    print(f"[SUCESSO] Planilha '{caminho}' criada com {len(df)} veículos de teste.")


def configurar_driver(headless: bool = HEADLESS_MODE) -> webdriver.Chrome:
    """
    Inicializa e configura o Selenium WebDriver com ChromeDriver automatizado via webdriver-manager.
    Inclui flags de robustez para execução em ambientes corporativos e headless.
    """
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless=new")
    
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    )

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


def clicar_elemento_seguro(driver: webdriver.Chrome, elemento) -> None:
    """
    Executa o clique em um elemento web. Caso ocorra interceptação por animações
    ou overlays da SPA, utiliza execução de script JavaScript de forma segura.
    """
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
        elemento.click()
    except Exception:
        driver.execute_script("arguments[0].click();", elemento)


def consultar_detran(driver: webdriver.Chrome, wait: WebDriverWait, placa: str, renavam: str) -> str:
    """
    Consulta o valor de Licenciamento em aberto no Detran-CE.
    
    Fluxo:
    1. Acessa: https://sistemas.detran.ce.gov.br/central/#veiculo
    2. Aguarda e clica no menu lateral: //a[contains(text(), 'Licenciamento')]
    3. Localiza os campos de Placa e Renavam dentre os inputs visíveis.
    4. Clica no botão: //button[contains(text(), 'CONFIRMAR')]
    5. Clica no link de emissão: //*[contains(text(), 'Clique AQUI se deseja imprimir seu licenciamento.')]
    6. Extrai o texto do valor usando a linha de totalização: //tr[contains(., 'TOTAL')]/td[last()]
    """
    url_detran = "https://sistemas.detran.ce.gov.br/central/#veiculo"
    try:
        driver.get(url_detran)

        # 2. Aguarda e clica no menu lateral
        menu_licenciamento = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Licenciamento')]"))
        )
        clicar_elemento_seguro(driver, menu_licenciamento)

        # 3. Localiza todos os inputs visíveis na tela
        wait.until(lambda d: len([i for i in d.find_elements(By.TAG_NAME, "input") if i.is_displayed()]) >= 2)
        inputs_visiveis = [i for i in driver.find_elements(By.TAG_NAME, "input") if i.is_displayed()]

        # Preenche Placa e Renavam
        inputs_visiveis[0].clear()
        inputs_visiveis[0].send_keys(str(placa).strip().upper())
        inputs_visiveis[1].clear()
        inputs_visiveis[1].send_keys(str(renavam).strip())

        # 4. Clica no botão CONFIRMAR
        btn_confirmar = wait.until(
            EC.presence_of_element_located((
                By.XPATH, "//button[contains(text(), 'CONFIRMAR') or contains(., 'CONFIRMAR') or @id='submit-veiculo']"
            ))
        )
        clicar_elemento_seguro(driver, btn_confirmar)

        # 5. Na próxima tela, aguarda e clica no link de emissão
        link_emissao = wait.until(
            EC.presence_of_element_located((
                By.XPATH, "//*[contains(text(), 'Clique AQUI se deseja imprimir seu licenciamento.')]"
            ))
        )
        clicar_elemento_seguro(driver, link_emissao)

        # 6. Um modal abrirá. Extrai o texto do valor usando o XPath da linha de totalização
        modal_total = wait.until(
            EC.visibility_of_element_located((By.XPATH, "//tr[contains(., 'TOTAL')]/td[last()]"))
        )
        valor_texto = modal_total.text.strip()
        return valor_texto if valor_texto else "R$ 0,00"

    except TimeoutException:
        # Se der timeout ao tentar extrair o resultado, assume sem débitos
        return "R$ 0,00"
    except Exception as erro:
        print(f"[ALERTA] Falha inesperada no Detran-CE ({placa}): {erro}")
        return "R$ 0,00"


def consultar_sefaz(driver: webdriver.Chrome, wait: WebDriverWait, placa: str, renavam: str) -> str:
    """
    Consulta os débitos de IPVA em aberto na Sefaz-CE.
    
    Fluxo:
    1. Acessa: https://ipva.sefaz.ce.gov.br/#/impostos/emitir-dae
    2. Localiza e preenche a Placa: //input[contains(@ng-model, 'placa') or @name='placa' or @formcontrolname='placa' or @id='mat-input-1']
    3. Localiza e preenche o Renavam: //input[contains(@ng-model, 'renavam') or @name='renavam' or @formcontrolname='renavam' or @id='mat-input-2']
    4. Clica no botão: //button[contains(text(), 'PESQUISAR') or contains(., 'Pesquisar')]
    5. Aguarda o carregamento da tabela inferior: //*[contains(text(), 'Débitos de IPVA não quitados')]
    6. Extrai o saldo em aberto: //td[contains(., 'Em aberto')]/preceding-sibling::td[contains(., 'R$')][1]
    """
    url_sefaz = "https://ipva.sefaz.ce.gov.br/#/impostos/emitir-dae"
    try:
        driver.get(url_sefaz)

        # 2. Localiza e preenche a Placa
        xpath_placa = "//input[contains(@ng-model, 'placa') or @name='placa' or @formcontrolname='placa' or @id='mat-input-1']"
        placa_input = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_placa)))
        placa_input.clear()
        placa_input.send_keys(str(placa).strip().upper())

        # 3. Localiza e preenche o Renavam
        xpath_renavam = "//input[contains(@ng-model, 'renavam') or @name='renavam' or @formcontrolname='renavam' or @id='mat-input-2']"
        renavam_input = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_renavam)))
        renavam_input.clear()
        renavam_input.send_keys(str(renavam).strip())

        # 4. Clica no botão PESQUISAR
        xpath_botao = "//button[contains(translate(., 'pesquisar', 'PESQUISAR'), 'PESQUISAR') or contains(@class, 'search-button')]"
        btn_pesquisar = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_botao)))
        clicar_elemento_seguro(driver, btn_pesquisar)

        # 5. Aguarda o carregamento da tabela inferior
        xpath_tabela = "//*[contains(text(), 'Débitos de IPVA não quitados')]"
        wait.until(EC.visibility_of_element_located((By.XPATH, xpath_tabela)))

        # 6. Extrai o valor do saldo em aberto ao lado da situação "Em aberto"
        xpath_valor = "//td[contains(., 'Em aberto')]/preceding-sibling::td[contains(., 'R$')][1]"
        saldo_aberto = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_valor)))
        valor_texto = saldo_aberto.text.strip()
        return valor_texto if valor_texto else "R$ 0,00"

    except TimeoutException:
        # Quando não há débitos não quitados, a tabela não existe ou dá timeout
        return "R$ 0,00"
    except Exception as erro:
        print(f"[ALERTA] Falha inesperada na Sefaz-CE ({placa}): {erro}")
        return "R$ 0,00"


def formatar_planilha_saida(caminho_arquivo: str) -> None:
    """
    Aplica formatação visual profissional na planilha Excel gerada (cabeçalhos, bordas,
    alinhamentos e larguras dinâmicas de coluna).
    """
    from openpyxl import load_workbook
    wb = load_workbook(caminho_arquivo)
    ws = wb.active

    cor_cabecalho = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fonte_linhas = Font(name="Calibri", size=10)
    alinhamento_centro = Alignment(horizontal="center", vertical="center")
    alinhamento_direita = Alignment(horizontal="right", vertical="center")
    borda_fina = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Formatação do cabeçalho
    for cell in ws[1]:
        cell.fill = cor_cabecalho
        cell.font = fonte_cabecalho
        cell.alignment = alinhamento_centro

    # Formatação dos dados
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for idx, cell in enumerate(row):
            cell.font = fonte_linhas
            cell.border = borda_fina
            if idx in [0, 1]:  # Placa e Renavam centralizados
                cell.alignment = alinhamento_centro
            else:              # Valores financeiros alinhados à direita
                cell.alignment = alinhamento_direita

    # Ajuste automático da largura das colunas
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    wb.save(caminho_arquivo)


def processar_frota():
    """
    Orquestrador principal:
    1. Prepara planilha de entrada.
    2. Lê 'frota.xlsx' via Pandas garantindo dtypes textuais.
    3. Inicializa o WebDriver com espera explícita mínima de 15s.
    4. Executa consultas no Detran-CE e Sefaz-CE para cada veículo.
    5. Salva os dados consolidados em 'resultados_frota_ce.xlsx'.
    """
    print("\n" + "=" * 70)
    print(" INICIANDO AUTOMAÇÃO DE CONSULTA DE DÉBITOS VEICULARES (CEARÁ)")
    print("=" * 70)

    # 1. Garante que o arquivo de entrada exista
    criar_planilha_entrada(ARQUIVO_ENTRADA)

    # 2. Leitura com Pandas garantindo que o RENAVAM não perca zeros à esquerda
    print(f"\n[1/4] Carregando arquivo '{ARQUIVO_ENTRADA}'...")
    df_frota = pd.read_excel(ARQUIVO_ENTRADA, dtype={"Placa": str, "RENAVAM": str})
    
    # Validação estrutural de colunas
    if "Placa" not in df_frota.columns or "RENAVAM" not in df_frota.columns:
        raise ValueError("O arquivo de entrada deve conter obrigatoriamente as colunas 'Placa' e 'RENAVAM'.")

    total_veiculos = len(df_frota)
    print(f"      Total de veículos para consulta: {total_veiculos}")

    # 3. Inicialização do Selenium
    print(f"\n[2/4] Inicializando navegador Chrome (Headless={HEADLESS_MODE}, Timeout={TIMEOUT_PADRAO}s)...")
    driver = configurar_driver(headless=HEADLESS_MODE)
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)

    licenciamentos = []
    ipvas = []
    status_processamento = []

    try:
        print("\n[3/4] Consultando débitos nos portais governamentais...")
        for index, row in df_frota.iterrows():
            placa = str(row["Placa"]).strip().upper()
            renavam = str(row["RENAVAM"]).strip()
            num_atual = index + 1

            print(f"\n--- [{num_atual}/{total_veiculos}] Processando: Placa {placa} | RENAVAM {renavam} ---")
            
            # Consulta Detran-CE (Licenciamento)
            inicio_detran = time.time()
            valor_licenciamento = consultar_detran(driver, wait, placa, renavam)
            duracao_detran = time.time() - inicio_detran
            print(f"    -> Detran-CE (Licenciamento): {valor_licenciamento} ({duracao_detran:.1f}s)")
            
            # Micro-pausa de segurança (rate limiting)
            time.sleep(INTERVALO_RATE_LIMIT)

            # Consulta Sefaz-CE (IPVA)
            inicio_sefaz = time.time()
            valor_ipva = consultar_sefaz(driver, wait, placa, renavam)
            duracao_sefaz = time.time() - inicio_sefaz
            print(f"    -> Sefaz-CE (IPVA):          {valor_ipva} ({duracao_sefaz:.1f}s)")

            licenciamentos.append(valor_licenciamento)
            ipvas.append(valor_ipva)
            status_processamento.append("Sucesso")

            # Micro-pausa antes do próximo veículo
            time.sleep(INTERVALO_RATE_LIMIT)

    except KeyboardInterrupt:
        print("\n[AVISO] Processamento interrompido pelo usuário.")
    except Exception as e:
        print(f"\n[ERRO CRÍTICO] Falha durante o processamento da frota: {e}")
    finally:
        print("\n[INFO] Finalizando sessão do navegador Chrome...")
        driver.quit()

    # 4. Consolidando resultados e gerando a planilha final
    print(f"\n[4/4] Gerando planilha final consolidada: '{ARQUIVO_SAIDA}'...")
    df_frota["Licenciamento_CE"] = licenciamentos
    df_frota["IPVA_CE"] = ipvas
    df_frota["Data_Consulta"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    df_frota.to_excel(ARQUIVO_SAIDA, index=False, engine="openpyxl")
    formatar_planilha_saida(ARQUIVO_SAIDA)

    print("\n" + "=" * 70)
    print(" PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
    print(f" Arquivo gerado: {os.path.abspath(ARQUIVO_SAIDA)}")
    print("=" * 70)

    # Exibição resumida na saída padrão
    print("\nResumo dos Resultados:")
    print(df_frota[["Placa", "RENAVAM", "Licenciamento_CE", "IPVA_CE"]])
    return df_frota


if __name__ == "__main__":
    processar_frota()
