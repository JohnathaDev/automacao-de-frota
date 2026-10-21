"""
========================================================================================
Automação de Consulta de Débitos Veiculares - Ceará (Detran-CE & Sefaz-CE) [CLI v4.2]
========================================================================================
Módulo de linha de comando que consome a arquitetura unificada de raspagem (v4.2):
- Leitura resiliente com Pandas, zfill(11) no RENAVAM e validação de placas.
- Consultas a Licenciamento, IPVA e Multas sem falsos negativos.
- Salvamento parcial/checkpoint em caso de interrupção.
- Exportação executiva para Excel e PDF Paisagem com proteção contra PermissionError.
========================================================================================
"""

import os
import re
import sys
import time
import pandas as pd

from core import (
    FrotaDataProcessor,
    FrotaScraperEngine,
    FrotaRelatorioExporter,
    PADRAO_PLACA,
    formatar_moeda_float,
    formatar_float_moeda,
    calcular_total_a_pagar_veiculo
)

HEADLESS_MODE = True
ARQUIVO_ENTRADA = "frota.xlsx"

DADOS_TESTE = [
    {"Placa": "BRA1A11", "RENAVAM": "01234567891"},
    {"Placa": "BRA2B22", "RENAVAM": "01234567892"},
    {"Placa": "BRA3C33", "RENAVAM": "01234567893"},
    {"Placa": "BRA4D44", "RENAVAM": "01234567894"},
    {"Placa": "BRA5E55", "RENAVAM": "01234567895"},
    {"Placa": "BRA6F66", "RENAVAM": "01234567896"}
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
    consultar_multas: bool = True,
    pasta_destino: str | None = None
) -> pd.DataFrame | None:
    """Executa a rotina de automação via linha de comando."""
    print("=" * 70)
    print(" AUTOMAÇÃO DE FROTA CE (v4.2) - DETRAN & SEFAZ")
    print("=" * 70)

    if not os.path.exists(caminho_entrada):
        if os.path.abspath(caminho_entrada) != os.path.abspath(ARQUIVO_ENTRADA):
            # Nunca gera dados fictícios no lugar de uma planilha informada pelo usuário
            print(f"[ERRO] Planilha não encontrada: '{caminho_entrada}'")
            return None
        criar_planilha_entrada(caminho_entrada)

    # 1. Carregamento resiliente com higienização e zfill(11)
    print(f"\n[1/4] Carregando e validando arquivo: '{caminho_entrada}'...")
    try:
        df_frota, col_placa, col_renavam = FrotaDataProcessor.carregar_planilha(caminho_entrada)
        col_chassi = FrotaDataProcessor.localizar_coluna_chassi(df_frota)
        info_chassi = f" | Chassi='{col_chassi}' (Fallback ativo)" if col_chassi else ""
        print(f"      Colunas detectadas: Placa='{col_placa}' | RENAVAM='{col_renavam}'{info_chassi}")
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
            chassi = (
                str(row[col_chassi]).strip()
                if (col_chassi and col_chassi in df_frota.columns and pd.notna(row[col_chassi]))
                else ""
            )
            pos = index + 1

            print(f"\n[{pos}/{total}] Processando veículo: Placa {placa} | RENAVAM {renavam}")

            val_lic = "-"
            val_multas = "-"
            m_a_vencer = "R$ 0,00"
            val_ipva = "-"

            if not PADRAO_PLACA.match(placa):
                if consultar_lic:
                    val_lic = "Placa Inválida"
                if consultar_multas:
                    val_multas = "Placa Inválida"
                if consultar_ipva:
                    val_ipva = "Placa Inválida"
                print(f"    -> [AVISO] Placa fora do padrão Mercosul/Tradicional: {placa}")
            else:
                try:
                    if consultar_lic or consultar_multas:
                        res_detran = scraper.consultar_detran(placa, renavam, consultar_lic, consultar_multas, chassi=chassi)
                        val_lic = res_detran[0]
                        if str(val_lic).strip() in ["R$ 0,00", "0,00"]:
                            val_lic = "PAGO"
                        val_multas = res_detran[1]
                        m_a_vencer = getattr(res_detran, "multas_a_vencer", "R$ 0,00")
                        print(f"    -> Detran-CE: Licenciamento={val_lic} | Multas={val_multas} (A vencer: {m_a_vencer})")
                        time.sleep(0.4)

                    if consultar_ipva:
                        val_ipva = scraper.consultar_sefaz_ipva(placa, renavam, chassi=chassi)
                        if str(val_ipva).strip() in ["R$ 0,00", "0,00"]:
                            val_ipva = "PAGO"
                        print(f"    -> Sefaz-CE:  IPVA em aberto={val_ipva}")
                        time.sleep(0.4)
                except Exception as e_veiculo:
                    # Erro inesperado em um veículo não interrompe o restante do lote
                    print(f"    -> [ERRO] Falha inesperada ao consultar {placa}: {e_veiculo}")
                    if consultar_lic and val_lic == "-":
                        val_lic = "Falha Detran"
                    if consultar_multas and val_multas == "-":
                        val_multas = "Falha Detran"
                        m_a_vencer = "Falha Detran"
                    if consultar_ipva and val_ipva == "-":
                        val_ipva = "Falha Sefaz"

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

    if not pasta_destino:
        pasta_destino = os.path.dirname(os.path.abspath(caminho_entrada)) if os.path.dirname(caminho_entrada) else "."

    arquivo_excel, arquivo_pdf, chk_path = FrotaDataProcessor.obter_caminhos_arquivos_saida(
        caminho_entrada, pasta_destino
    )
    arquivo_pdf_bi = FrotaDataProcessor.obter_caminho_pdf_bi(caminho_entrada, pasta_destino)
    nome_base = FrotaDataProcessor.gerar_nome_base_resultado(caminho_entrada)
    identificador_frota = re.sub(r"^resultados[_\-\s]*", "", nome_base, flags=re.IGNORECASE).strip() or nome_base

    # Lê a última planilha existente antes de sobrescrevê-la para identificar baixas/pagamentos
    df_anterior, data_anterior = FrotaDataProcessor.carregar_resultado_anterior(arquivo_excel)
    comparativo_anterior = None
    if df_anterior is not None:
        comparativo_anterior = FrotaDataProcessor.comparar_auditorias(
            df_anterior,
            df_final,
            data_anterior=data_anterior
        )
        if comparativo_anterior:
            print(f"\n  [Comparativo com planilha anterior de {comparativo_anterior['data_anterior']}]")
            for frase in comparativo_anterior.get("frases_resumo", []):
                print(f"    • {frase}")

    # Gera Excel com proteção contra PermissionError
    caminho_excel = FrotaRelatorioExporter.exportar_excel(df_final, arquivo_excel)
    print(f"  -> Planilha Excel gerada:    {os.path.abspath(caminho_excel)}")

    filtros_ativos = []
    if consultar_ipva:
        filtros_ativos.append("IPVA (Sefaz)")
    if consultar_lic:
        filtros_ativos.append("Licenciamento (Detran)")
    if consultar_multas:
        filtros_ativos.append("Multas (Detran)")

    # Gera PDF Paisagem com proteção contra PermissionError
    caminho_pdf = FrotaRelatorioExporter.exportar_pdf(
        df_final,
        arquivo_pdf,
        totais_dict,
        filtros_ativos,
        identificador_frota=identificador_frota
    )
    print(f"  -> Relatório PDF gerado:     {os.path.abspath(caminho_pdf)}")

    # Gera 2º PDF com o Relatório Analítico de BI + Comparativo de Evolução
    caminho_pdf_bi = FrotaRelatorioExporter.exportar_pdf_bi(
        df_final,
        arquivo_pdf_bi,
        totais_dict,
        filtros_ativos,
        identificador_frota=identificador_frota,
        comparativo_anterior=comparativo_anterior
    )
    print(f"  -> Relatório BI PDF gerado:  {os.path.abspath(caminho_pdf_bi)}")

    print("\n" + "=" * 70)
    print(" PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
    print(f" Totais: Lic={totais_dict['lic']} | IPVA={totais_dict['ipva']} | Multas={totais_dict['multas']} | Total={totais_dict['total']}")
    print("=" * 70)

    return df_final


if __name__ == "__main__":
    arq = sys.argv[1] if len(sys.argv) > 1 else ARQUIVO_ENTRADA
    processar_frota(caminho_entrada=arq)
