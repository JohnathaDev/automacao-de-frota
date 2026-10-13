"""
========================================================================================
Automação de Consulta de Débitos Veiculares - Ceará (Detran-CE & Sefaz-CE) [CLI v3.0]
========================================================================================
Módulo de linha de comando que consome a arquitetura unificada de raspagem (v3.0):
- Leitura resiliente com Pandas, zfill(11) no RENAVAM e validação de placas.
- Consultas a Licenciamento, IPVA e Multas sem falsos negativos.
- Salvamento parcial/checkpoint em caso de interrupção.
- Exportação executiva para Excel e PDF Paisagem com proteção contra PermissionError.
========================================================================================
"""

import os
import sys
import time
from datetime import datetime
import pandas as pd

from core import (
    FrotaDataProcessor,
    FrotaScraperEngine,
    FrotaRelatorioExporter,
    formatar_moeda_float,
    formatar_float_moeda,
    calcular_total_a_pagar_veiculo
)

HEADLESS_MODE = True
ARQUIVO_ENTRADA = "frota.xlsx"
ARQUIVO_SAIDA_EXCEL = "resultados_frota_ce.xlsx"
ARQUIVO_SAIDA_PDF = "resultados_frota_ce.pdf"

DADOS_TESTE = [
    {"Placa": "TII5F89", "RENAVAM": "01458899249"},
    {"Placa": "THT5F69", "RENAVAM": "01458894204"},
    {"Placa": "THU5D09", "RENAVAM": "01458895766"},
    {"Placa": "THV0A49", "RENAVAM": "01434835143"},
    {"Placa": "TIE8B49", "RENAVAM": "01422988322"},
    {"Placa": "THY7G95", "RENAVAM": "01450027048"}
]


def criar_planilha_entrada(caminho: str = ARQUIVO_ENTRADA, forcar: bool = False) -> None:
    """Cria a planilha de teste 'frota.xlsx' caso não exista, garantindo RENAVAM como texto."""
    if os.path.exists(caminho) and not forcar:
        print(f"[INFO] Arquivo de teste já existente: {caminho}")
        return

    print(f"[INFO] Gerando planilha inicial de teste: {caminho}...")
    df = pd.DataFrame(DADOS_TESTE)
    df["RENAVAM"] = df["RENAVAM"].astype(str).str.zfill(11)

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Frota")

    print(f"[SUCESSO] Planilha '{caminho}' criada com {len(df)} veículos de teste.")


def processar_frota(
    caminho_entrada: str = ARQUIVO_ENTRADA,
    consultar_ipva: bool = True,
    consultar_lic: bool = True,
    consultar_multas: bool = True
) -> pd.DataFrame | None:
    """Executa a rotina de automação via linha de comando."""
    print("=" * 70)
    print(" AUTOMAÇÃO DE FROTA CE (v3.0) - DETRAN & SEFAZ")
    print("=" * 70)

    criar_planilha_entrada(caminho_entrada)

    # 1. Carregamento resiliente com higienização e zfill(11)
    print(f"\n[1/4] Carregando e validando arquivo: '{caminho_entrada}'...")
    try:
        df_frota, col_placa, col_renavam = FrotaDataProcessor.carregar_planilha(caminho_entrada)
        print(f"      Colunas detectadas: Placa='{col_placa}' | RENAVAM='{col_renavam}'")
        print(f"      Total de veículos válidos: {len(df_frota)}")
    except Exception as e:
        print(f"[ERRO] Falha ao carregar a planilha: {e}")
        return None

    total = len(df_frota)
    scraper = FrotaScraperEngine(headless=HEADLESS_MODE)

    licenciamentos = []
    ipvas = []
    multas = []
    multas_a_vencer_lista = []

    soma_lic = 0.0
    soma_ipva = 0.0
    soma_multas = 0.0

    try:
        print(f"\n[2/4] Inicializando Selenium WebDriver (Headless={HEADLESS_MODE})...")
        scraper.iniciar()

        print("\n[3/4] Consultando débitos nos portais...")
        for index, row in df_frota.iterrows():
            placa = str(row[col_placa]).strip().upper()
            renavam = str(row[col_renavam]).strip()
            pos = index + 1

            print(f"\n[{pos}/{total}] Processando veículo: Placa {placa} | RENAVAM {renavam}")

            val_lic = "-"
            val_multas = "-"
            m_a_vencer = "R$ 0,00"
            val_ipva = "-"

            if consultar_lic or consultar_multas:
                res_detran = scraper.consultar_detran(placa, renavam, consultar_lic, consultar_multas)
                val_lic = res_detran[0]
                val_multas = res_detran[1]
                m_a_vencer = getattr(res_detran, "multas_a_vencer", "R$ 0,00")
                print(f"    -> Detran-CE: Licenciamento={val_lic} | Multas={val_multas} (A vencer: {m_a_vencer})")
                time.sleep(0.4)

            if consultar_ipva:
                val_ipva = scraper.consultar_sefaz_ipva(placa, renavam)
                print(f"    -> Sefaz-CE:  IPVA em aberto={val_ipva}")
                time.sleep(0.4)

            licenciamentos.append(val_lic)
            ipvas.append(val_ipva)
            multas.append(val_multas)
            multas_a_vencer_lista.append(m_a_vencer)

            soma_lic += formatar_moeda_float(val_lic)
            soma_ipva += formatar_moeda_float(val_ipva)
            soma_multas += formatar_moeda_float(val_multas)

    except KeyboardInterrupt:
        print("\n[AVISO] Processamento interrompido pelo usuário. Consolidando resultados parciais...")
    except Exception as e:
        print(f"\n[ERRO CRÍTICO] Falha durante o processamento: {e}")
    finally:
        print("\n[INFO] Finalizando sessão do navegador...")
        scraper.fechar()

    # 4. Consolidação e Exportação (Salva progresso parcial se interrompido)
    qtd = len(licenciamentos)
    if qtd == 0:
        print("[AVISO] Nenhum registro processado.")
        return None

    print(f"\n[4/4] Consolidando resultados e gerando relatórios ({qtd}/{total} veículos)...")
    df_final = FrotaDataProcessor.consolidar_resultados(
        df_frota, licenciamentos, ipvas, multas, qtd,
        resultados_multas_a_vencer=multas_a_vencer_lista
    )

    soma_total_real = sum(
        calcular_total_a_pagar_veiculo(
            licenciamentos[i],
            ipvas[i],
            multas[i],
            multas_a_vencer_lista[i]
        )
        for i in range(qtd)
    )

    totais_dict = {
        "lic": formatar_float_moeda(soma_lic),
        "ipva": formatar_float_moeda(soma_ipva),
        "multas": formatar_float_moeda(soma_multas),
        "total": formatar_float_moeda(soma_total_real)
    }

    # Gera Excel com proteção contra PermissionError
    caminho_excel = FrotaRelatorioExporter.exportar_excel(df_final, ARQUIVO_SAIDA_EXCEL)
    print(f"  -> Planilha Excel gerada: {os.path.abspath(caminho_excel)}")

    # Gera PDF Paisagem com proteção contra PermissionError
    caminho_pdf = FrotaRelatorioExporter.exportar_pdf(
        df_final,
        ARQUIVO_SAIDA_PDF,
        totais_dict,
        ["IPVA (Sefaz)", "Licenciamento (Detran)", "Multas (Detran)"]
    )
    print(f"  -> Relatório PDF gerado:  {os.path.abspath(caminho_pdf)}")

    print("\n" + "=" * 70)
    print(" PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
    print(f" Totais: Lic={totais_dict['lic']} | IPVA={totais_dict['ipva']} | Multas={totais_dict['multas']} | Total={totais_dict['total']}")
    print("=" * 70)

    return df_final


if __name__ == "__main__":
    processar_frota()
