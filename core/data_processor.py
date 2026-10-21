"""
Processamento, higienização e validação de dados de frotas com Pandas.
Regras financeiras de formatação para relatórios e exibição.
"""
import os
import re
import calendar
from datetime import datetime, timedelta
import pandas as pd

from core.logger_config import obter_logger

logger = obter_logger("DataProcessor")

# Padrão oficial de placas veiculares brasileiras (Mercosul e Tradicional)
PADRAO_PLACA = re.compile(r"^[A-Z]{3}[0-9][A-Z0-9][0-9]{2}$|^[A-Z]{3}[0-9]{4}$")


def formatar_moeda_float(valor_str: str) -> float:
    """Converte valores monetários textuais (ex: 'R$ 188,96' ou '188.96') em float 188.96."""
    try:
        if not valor_str:
            return 0.0
        val_lower = str(valor_str).lower().strip()
        termos_ignorar = [
            "-", "n/a", "não consultado", "nao consultado", "falha", "erro", "aguardando",
            "inválid", "invalid", "não localizado", "nao localizado", "timeout", "...", "pago", "restrição", "restricao"
        ]
        if any(termo in val_lower for termo in termos_ignorar):
            return 0.0
        limpo = val_lower.replace("r$", "").strip()
        if "," in limpo:
            limpo = limpo.replace(".", "").replace(",", ".").strip()
        elif limpo.count(".") > 1:
            limpo = limpo.replace(".", "").strip()
        return float(limpo)
    except Exception:
        return 0.0


def formatar_float_moeda(valor: float) -> str:
    """Converte um float monetário em formato brasileiro: R$ 1.234,56."""
    try:
        formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return formatado
    except Exception:
        return "R$ 0,00"


def calcular_total_a_pagar_veiculo(
    lic: str | float,
    ipva: str | float,
    multas: str | float,
    multas_a_vencer: str | float | None = None
) -> float:
    """
    Calcula o Total a Pagar real para um veículo, eliminando duplicações:
    - O valor do Licenciamento no Detran-CE (R$ 188,96 para motos e R$ 220,45 para carros de taxa base)
      já inclui as multas vencidas em seu extrato oficial.
    - Portanto, quando há Licenciamento em aberto (lic > 0), apenas as multas A VENCER
      (ainda dentro do prazo de pagamento) devem ser somadas ao Licenciamento + IPVA.
    - As multas já vencidas não devem ser somadas novamente, pois já foram contabilizadas no Licenciamento.
    - Caso não haja débito de Licenciamento (lic == 0), ou o Licenciamento não tenha sido consultado,
      as multas em aberto são somadas normalmente.
    """
    val_lic_f = formatar_moeda_float(lic) if isinstance(lic, str) else float(lic or 0.0)
    val_ipva_f = formatar_moeda_float(ipva) if isinstance(ipva, str) else float(ipva or 0.0)
    val_multas_f = formatar_moeda_float(multas) if isinstance(multas, str) else float(multas or 0.0)

    if multas_a_vencer is not None:
        val_a_vencer_f = (
            formatar_moeda_float(multas_a_vencer)
            if isinstance(multas_a_vencer, str)
            else float(multas_a_vencer or 0.0)
        )
    else:
        # Fallback analítico baseado nas taxas oficiais informadas:
        # Moto = R$ 188,96, Carro = R$ 220,45
        if val_lic_f > 0 and val_multas_f > 0:
            if abs(val_lic_f - 188.96) < 0.5 or abs(val_lic_f - 220.45) < 0.5:
                # Nenhuma multa vencida no licenciamento; todo o saldo de multas está a vencer
                val_a_vencer_f = val_multas_f
            elif val_lic_f > 220.45:
                taxa_base = 188.96 if (val_lic_f - 188.96) <= val_multas_f + 1.0 else 220.45
                multas_vencidas_est = max(0.0, val_lic_f - taxa_base)
                val_a_vencer_f = max(0.0, val_multas_f - multas_vencidas_est)
            elif val_lic_f > 188.96:
                multas_vencidas_est = max(0.0, val_lic_f - 188.96)
                val_a_vencer_f = max(0.0, val_multas_f - multas_vencidas_est)
            else:
                val_a_vencer_f = val_multas_f
        else:
            val_a_vencer_f = val_multas_f

    if val_lic_f > 0:
        return round(val_lic_f + val_ipva_f + val_a_vencer_f, 2)
    else:
        return round(val_ipva_f + val_multas_f, 2)


def obter_estilo_financeiro(valor_float: float, is_licenciamento: bool = False) -> tuple[bool, bool, str | None]:
    """
    Retorna regras de estilo financeiro cumulativas:
    - Valor > R$ 1.000,00: negrito = True
    - Valor > R$ 1.500,00: negrito = True, sublinhado = True
    - Licenciamento > R$ 188,96 (multas/taxas atreladas): cor_fonte = '#DC2626' (vermelho corporativo)
    Retorna: (bold: bool, underline: bool, color_hex: str | None)
    """
    bold = False
    underline = False
    color_hex = None

    if valor_float > 1500.00:
        bold = True
        underline = True
    elif valor_float > 1000.00:
        bold = True

    if is_licenciamento and valor_float > 188.96:
        color_hex = "#DC2626"

    return bold, underline, color_hex


def obter_ultimo_dia_util_mes(data_base: datetime | None = None) -> str:
    """
    Retorna o último dia útil do mês atual no formato DD/MM/AAAA para vencimento do DAE de IPVA,
    retrocedendo automaticamente caso o último dia caia em fim de semana ou feriado nacional.
    """
    if data_base is None:
        data_base = datetime.now()

    ano = data_base.year
    mes = data_base.month
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    d = datetime(ano, mes, ultimo_dia)

    # Feriados nacionais fixos no Brasil (dia, mês)
    feriados = [
        (1, 1), (21, 4), (1, 5), (7, 9), (12, 10), (2, 11), (15, 11), (20, 11), (25, 12)
    ]

    while d.weekday() >= 5 or (d.day, d.month) in feriados:
        d -= timedelta(days=1)

    # Se o último dia útil já passou (ex: consulta feita no sábado/domingo de fim de mês),
    # usa a própria data base para que o vencimento nunca seja retroativo
    if d.date() < data_base.date():
        d = data_base

    return d.strftime("%d/%m/%Y")


def extrair_placa_limpa(val: str) -> str:
    """
    Extrai apenas a Placa válida (Mercosul ou Tradicional) mesmo quando a célula
    contém textos extras, símbolos, hífens, pontos ou múltiplos caracteres.
    """
    if val is None or pd.isna(val):
        return ""
    val_str = str(val).strip().upper()
    if val_str in ["NAN", "NONE", "", "N/A", "-", "NULL"]:
        return ""

    # 1. Busca padrão de placa isolado ou cercado por delimitadores/textos
    m = re.search(r"(?<![A-Z0-9])([A-Z]{3})[\s\-\._/]*([0-9][A-Z0-9][0-9]{2})(?![A-Z0-9])", val_str)
    if m:
        return f"{m.group(1)}{m.group(2)}"

    # 2. Fallback: remove caracteres especiais e verifica padrão
    limpo = re.sub(r"[^A-Z0-9]", "", val_str)
    m2 = re.search(r"([A-Z]{3}[0-9][A-Z0-9][0-9]{2}|[A-Z]{3}[0-9]{4})", limpo)
    if m2 and len(limpo) == 7:
        return m2.group(1)
    return limpo


def sanitizar_renavam(val: str) -> str:
    """
    Extrai e higieniza o código RENAVAM (11 dígitos com zeros à esquerda),
    ignorando textos, rótulos, pontuações ou caracteres extras na célula.
    """
    if val is None or pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.upper() in ["NAN", "NONE", "", "N/A", "-", "NULL"]:
        return ""
    if val_str.endswith(".0"):
        val_str = val_str[:-2]

    # Se houver pontuação de milhar/dígito (ex: 01.486.619.662-0), unifica dígitos próximos
    sem_pontuacao_num = re.sub(r"(?<=\d)[\.\-\s/](?=\d)", "", val_str)

    # Prioriza sequências numéricas de 9 a 11 dígitos caso haja outros números no texto da célula
    candidatos = re.findall(r"(?<!\d)\d{9,11}(?!\d)", sem_pontuacao_num)
    if candidatos:
        return candidatos[0].zfill(11)

    digitos = "".join(c for c in sem_pontuacao_num if c.isdigit())
    if digitos and len(digitos) <= 11:
        return digitos.zfill(11)
    return digitos


def sanitizar_chassi(val: str) -> str:
    """
    Extrai e higieniza o código do Chassi (VIN) do veículo (tipicamente 17 caracteres alfanuméricos),
    removendo rótulos, espaços ou símbolos extras.
    """
    if val is None or pd.isna(val):
        return ""
    val_str = str(val).strip().upper()
    if val_str in ["NAN", "NONE", "", "N/A", "-", "NULL"]:
        return ""

    # Procura token alfanumérico de 17 caracteres (padrão internacional VIN/Chassi)
    candidatos_17 = re.findall(r"(?<![A-Z0-9])[A-Z0-9]{17}(?![A-Z0-9])", val_str)
    if candidatos_17:
        return candidatos_17[0]

    limpo = re.sub(r"[^A-Z0-9]", "", val_str)
    if len(limpo) >= 7:
        return limpo[:25]
    return ""


class FrotaDataProcessor:
    """Higienização de dados, validação de regras de frota e consolidação."""

    @staticmethod
    def localizar_coluna_chassi(df: pd.DataFrame) -> str | None:
        """Detecta automaticamente se a planilha possui coluna de Chassi do veículo."""
        if df is None or df.empty:
            return None

        for col in df.columns:
            nome_norm = re.sub(r"[^A-Z0-9]", "", str(col).strip().upper())
            if any(termo in nome_norm for termo in ["CHASSI", "CHASSIS", "VIN"]):
                return col

        # Fallback por inspeção de conteúdo: coluna com códigos VIN de 17 caracteres alfanuméricos (com letras e números)
        for col in df.columns:
            nome_norm = re.sub(r"[^A-Z0-9]", "", str(col).strip().upper())
            if any(t in nome_norm for t in ["PLACA", "RENAVAM", "RENAVAN", "RNV", "PROPRIET", "CLIENTE", "STATUS", "VALOR", "TOTAL"]):
                continue
            amostra = [sanitizar_chassi(v) for v in df[col].dropna().head(10)]
            validos_17 = [
                v for v in amostra
                if len(v) == 17 and any(c.isalpha() for c in v) and any(c.isdigit() for c in v)
            ]
            if validos_17 and len(validos_17) >= max(1, len(amostra) // 2):
                return col

        return None

    @staticmethod
    def localizar_colunas_chave(df: pd.DataFrame) -> tuple[str, str]:
        """Detecta automaticamente os nomes das colunas de Placa e RENAVAM na planilha."""
        col_placa = None
        col_renavam = None

        colunas_upper = {str(col).strip().upper(): col for col in df.columns}

        for nome_chave, col_original in colunas_upper.items():
            if "PLACA" in nome_chave:
                col_placa = col_original
                break

        for nome_chave, col_original in colunas_upper.items():
            if any(termo in nome_chave for termo in ["RENAVAM", "RENAVAN", "RNV"]):
                col_renavam = col_original
                break

        # Fallback heurístico caso a planilha tenha cabeçalhos atípicos ou vários caracteres
        if not col_placa:
            for col in df.columns:
                if col == col_renavam:
                    continue
                amostra = [extrair_placa_limpa(v) for v in df[col].dropna().head(15)]
                placas_validas = [p for p in amostra if PADRAO_PLACA.match(p)]
                if placas_validas and len(placas_validas) >= max(1, len(amostra) // 2):
                    col_placa = col
                    break

        if not col_renavam:
            for col in df.columns:
                if col == col_placa:
                    continue
                nome_u = str(col).strip().upper()
                if any(t in nome_u for t in ["CHASSI", "CHASSIS", "VIN", "PROPRIET", "CLIENTE"]):
                    continue
                amostra = [sanitizar_renavam(v) for v in df[col].dropna().head(15)]
                rens_validos = [r for r in amostra if len(r) == 11 and r.isdigit()]
                if rens_validos and len(rens_validos) >= max(1, len(amostra) // 2):
                    col_renavam = col
                    break

        if not col_placa or not col_renavam:
            faltantes = []
            if not col_placa:
                faltantes.append("'Placa'")
            if not col_renavam:
                faltantes.append("'RENAVAM'")
            raise ValueError(
                f"A planilha selecionada não possui as seguintes colunas obrigatórias: {', '.join(faltantes)}.\n"
                f"Colunas encontradas: {list(df.columns)}"
            )

        return col_placa, col_renavam

    @classmethod
    def carregar_planilha(cls, caminho_arquivo: str) -> tuple[pd.DataFrame, str, str]:
        """
        Carrega a planilha Excel assegurando que:
        - Planilhas com títulos no topo ou várias colunas/caracteres extras sejam suportadas
        - Linhas vazias sejam removidas
        - Valores de RENAVAM não percam zeros à esquerda (zfill 11)
        - Placas sejam higienizadas (extraindo apenas a placa limpa)
        - Caso exista coluna de Chassi, ela seja higienizada e preservada em memória para fallback,
          mas colunas irrelevantes de ruído sejam filtradas.
        """
        df = pd.read_excel(caminho_arquivo, dtype=str)
        if len(df) == 0:
            raise ValueError("O arquivo Excel selecionado não contém nenhum registro.")

        # Se o cabeçalho real estiver abaixo da primeira linha (ex: planilhas com título no topo)
        try:
            col_placa, col_renavam = cls.localizar_colunas_chave(df)
        except ValueError:
            df_raw = pd.read_excel(caminho_arquivo, dtype=str, header=None)
            linha_cabecalho = None
            for idx_row in range(min(10, len(df_raw))):
                valores_linha = [str(v).strip().upper() for v in df_raw.iloc[idx_row].values if pd.notna(v)]
                texto_linha = " ".join(valores_linha)
                if "PLACA" in texto_linha and any(t in texto_linha for t in ["RENAVAM", "RENAVAN", "RNV", "CHASSI"]):
                    linha_cabecalho = idx_row
                    break
            if linha_cabecalho is not None:
                df = pd.read_excel(caminho_arquivo, dtype=str, header=linha_cabecalho)
            col_placa, col_renavam = cls.localizar_colunas_chave(df)

        col_chassi = cls.localizar_coluna_chassi(df)

        # 1. Remove linhas onde placa é nula
        df = df.dropna(subset=[col_placa]).copy()

        # 2. Higienização de Placa (extrai apenas a placa limpa mesmo com textos/caracteres ao redor)
        df[col_placa] = df[col_placa].apply(extrair_placa_limpa)

        # 3. Higienização de RENAVAM (preserva 11 dígitos com zeros à esquerda)
        df[col_renavam] = df[col_renavam].apply(sanitizar_renavam)

        # 4. Higienização de Chassi (se presente na planilha)
        if col_chassi and col_chassi in df.columns:
            df[col_chassi] = df[col_chassi].apply(sanitizar_chassi)

        # 5. Remove linhas sem Placa válida ou sem ambos (RENAVAM e Chassi)
        placas_invalidas = ["NAN", "NONE", "", "N/A", "-", "NULL"]
        mask_placa_ok = ~df[col_placa].isin(placas_invalidas)
        mask_ren_ok = ~df[col_renavam].isin(placas_invalidas)
        if col_chassi and col_chassi in df.columns:
            mask_chassi_ok = ~df[col_chassi].isin(placas_invalidas)
            df = df[mask_placa_ok & (mask_ren_ok | mask_chassi_ok)].copy()
            # Onde o RENAVAM estiver vazio mas houver Chassi, preenche com '-' para exibição
            df[col_renavam] = df[col_renavam].apply(lambda r: r if r and r not in placas_invalidas else "-")
        else:
            df = df[mask_placa_ok & mask_ren_ok].copy()

        if len(df) == 0:
            raise ValueError("Nenhum registro com Placa e RENAVAM/Chassi válidos foi encontrado na planilha.")

        # 6. Filtra colunas de ruído caso a planilha possua dezenas de colunas extras com caracteres diversos,
        # mantendo apenas Proprietário/Cliente (se existir), Placa, RENAVAM, Chassi (para consulta) e colunas de auditoria
        colunas_manter = []
        col_prop_encontrada = False
        colunas_auditoria = {"VALOR_LICENCIAMENTO", "VALOR_IPVA", "VALOR_MULTAS", "TOTAL_A_PAGAR", "STATUS", "DATA_CONSULTA"}
        for c in df.columns:
            cu = str(c).strip().upper()
            if c in (col_placa, col_renavam, col_chassi):
                colunas_manter.append(c)
            elif cu in colunas_auditoria:
                colunas_manter.append(c)
            elif not col_prop_encontrada and any(t in cu for t in ["PROPRIET", "CLIENTE", "EMPRESA", "LOCAT"]):
                colunas_manter.append(c)
                col_prop_encontrada = True

        if colunas_manter:
            df = df[colunas_manter].copy()

        df = df.reset_index(drop=True)
        return df, col_placa, col_renavam

    @staticmethod
    def consolidar_resultados(
        df_original: pd.DataFrame,
        resultados_lic: list[str],
        resultados_ipva: list[str],
        resultados_multas: list[str],
        qtd_processada: int,
        resultados_multas_a_vencer: list[str] | None = None
    ) -> pd.DataFrame:
        """
        Constrói o DataFrame final preservando as colunas de identificação (sem expor o Chassi nas descrições/saídas)
        e adicionando as colunas oficiais de débitos, total real a pagar e situação.
        Evita duplicação de multas vencidas já inclusas no Licenciamento.
        """
        df_final = df_original.iloc[:qtd_processada].copy()

        # Remove a coluna de Chassi (e eventuais colunas Unnamed) na transferência para Excel/PDF
        # para manter os relatórios exatamente no padrão atual
        cols_remover = [
            c for c in df_final.columns
            if any(t in re.sub(r"[^A-Z0-9]", "", str(c).strip().upper()) for t in ["CHASSI", "CHASSIS", "VIN"])
            or str(c).strip().upper().startswith("UNNAMED")
        ]
        if cols_remover:
            df_final = df_final.drop(columns=cols_remover)


        def normalizar_valor_pago(v: str) -> str:
            v_str = str(v).strip()
            v_lower = v_str.lower()
            if v_str in ["R$ 0,00", "0,00", "0.00", "0"]:
                return "PAGO"
            if formatar_moeda_float(v_str) == 0.0 and v_str not in ["-", "N/A", "Não Consultado", ""] and not any(
                f in v_lower for f in ["falha", "erro", "inválid", "invalid", "timeout", "não localizado", "nao localizado", "restriç", "restric"]
            ):
                return "PAGO"
            return v_str

        resultados_lic_proc = [normalizar_valor_pago(v) for v in resultados_lic[:qtd_processada]]
        resultados_ipva_proc = [normalizar_valor_pago(v) for v in resultados_ipva[:qtd_processada]]

        df_final["Valor_Licenciamento"] = resultados_lic_proc
        df_final["Valor_IPVA"] = resultados_ipva_proc
        df_final["Valor_Multas"] = resultados_multas[:qtd_processada]

        # Determinação do Total a Pagar e Status consolidado
        totais_a_pagar_lista = []
        status_lista = []

        for i in range(qtd_processada):
            lic = str(resultados_lic_proc[i])
            ipva = str(resultados_ipva_proc[i])
            multas = str(resultados_multas[i])
            m_a_vencer = (
                str(resultados_multas_a_vencer[i])
                if resultados_multas_a_vencer and i < len(resultados_multas_a_vencer)
                else None
            )

            nao_consultado = all(
                v.strip() in ["-", "N/A", "Não Consultado", ""]
                for v in [lic, ipva, multas]
            )

            tem_falha = any(
                termo in v.lower()
                for v in [lic, ipva, multas]
                for termo in ["falha", "erro", "timeout", "inválid", "invalid", "não localizado", "nao localizado"]
            )

            total_real = calcular_total_a_pagar_veiculo(lic, ipva, multas, m_a_vencer)
            totais_a_pagar_lista.append("-" if nao_consultado else formatar_float_moeda(total_real))

            if nao_consultado:
                status_lista.append("Não Consultado")
            elif tem_falha:
                status_lista.append("Falha na Consulta")
            elif "restriç" in lic.lower() or "restric" in lic.lower():
                if total_real > 0:
                    status_lista.append("Com Débitos / Restrição")
                else:
                    status_lista.append("Restrição de Impedimento")
            elif total_real > 0:
                status_lista.append("Com Débitos Pendentes")
            else:
                status_lista.append("Sem Débitos")

        df_final["Total_a_Pagar"] = totais_a_pagar_lista
        df_final["Status"] = status_lista
        df_final["Data_Consulta"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        return df_final

    @staticmethod
    def gerar_nome_base_resultado(caminho_ou_nome_arquivo: str) -> str:
        """Delega para a função utilitária gerar_nome_base_resultado."""
        return gerar_nome_base_resultado(caminho_ou_nome_arquivo)

    @staticmethod
    def obter_caminhos_arquivos_saida(
        caminho_entrada: str,
        pasta_destino: str | None = None
    ) -> tuple[str, str, str]:
        """Delega para a função utilitária obter_caminhos_arquivos_saida."""
        return obter_caminhos_arquivos_saida(caminho_entrada, pasta_destino)

    @staticmethod
    def obter_caminho_pdf_bi(
        caminho_entrada: str,
        pasta_destino: str | None = None
    ) -> str:
        """Delega para a função utilitária obter_caminho_pdf_bi."""
        return obter_caminho_pdf_bi(caminho_entrada, pasta_destino)

    @staticmethod
    def carregar_resultado_anterior(caminho_excel_resultado: str) -> tuple[pd.DataFrame | None, str]:
        """Delega para a função utilitária carregar_resultado_anterior."""
        return carregar_resultado_anterior(caminho_excel_resultado)

    @staticmethod
    def comparar_auditorias(
        df_anterior: pd.DataFrame | None,
        df_atual: pd.DataFrame | None,
        data_anterior: str = "",
        data_atual: str = ""
    ) -> dict | None:
        """Delega para a função utilitária comparar_auditorias."""
        return comparar_auditorias(df_anterior, df_atual, data_anterior, data_atual)


def gerar_nome_base_resultado(caminho_ou_nome_arquivo: str) -> str:
    """
    Gera o nome base para os arquivos de resultados (PDF e Excel)
    a partir do nome do arquivo consultado.

    Regras e Exemplos:
      - 'frota_final_8.xlsx' -> 'resultados_final_8'
      - 'frota_ferreira_gomes.xlsx' -> 'resultados_ferreira_gomes'
      - 'frota-final-8.xlsx' -> 'resultados-final-8'
      - 'frota final 8.xlsx' -> 'resultados final 8'
      - 'frota_ce.xlsx' -> 'resultados_ce'
      - 'frota.xlsx' -> 'resultados_frota'
      - 'ferreira_gomes.xlsx' -> 'resultados_ferreira_gomes'
      - 'final_8.xlsx' -> 'resultados_final_8'
      - 'resultados_final_8.xlsx' -> 'resultados_final_8'
    """
    if not caminho_ou_nome_arquivo:
        return "resultados_frota"

    nome_arquivo = os.path.basename(str(caminho_ou_nome_arquivo).strip())
    nome_base, _ = os.path.splitext(nome_arquivo)
    nome_base = nome_base.strip()

    if not nome_base:
        return "resultados_frota"

    # Se já começa com 'resultados_' ou 'resultados-' ou 'resultados '
    if re.match(r"^resultados[_\-\s]", nome_base, re.IGNORECASE):
        return nome_base

    # Se começa com 'frota' seguido de separador (_ - ou espaço)
    padrao_frota = re.match(r"^frota([_\-\s]+)(.+)$", nome_base, re.IGNORECASE)
    if padrao_frota:
        sep = padrao_frota.group(1)
        resto = padrao_frota.group(2).strip()
        if resto:
            return f"resultados{sep}{resto}"

    # Se for exatamente 'frota'
    if nome_base.lower() == "frota":
        return "resultados_frota"

    # Caso geral: adiciona prefixo resultados_
    return f"resultados_{nome_base}"


def obter_caminhos_arquivos_saida(
    caminho_entrada: str,
    pasta_destino: str | None = None
) -> tuple[str, str, str]:
    """
    Gera as rotas completas para os arquivos de saída:
    (caminho_excel, caminho_pdf, caminho_checkpoint)
    baseando-se no nome da planilha que está sendo consultada.
    """
    nome_base = gerar_nome_base_resultado(caminho_entrada)
    if not pasta_destino:
        pasta_destino = os.path.dirname(os.path.abspath(caminho_entrada)) if caminho_entrada else os.getcwd()

    arquivo_excel = os.path.join(pasta_destino, f"{nome_base}.xlsx")
    arquivo_pdf = os.path.join(pasta_destino, f"{nome_base}.pdf")
    arquivo_checkpoint = os.path.join(pasta_destino, f"~checkpoint_{nome_base}.xlsx")
    return arquivo_excel, arquivo_pdf, arquivo_checkpoint


def obter_caminho_pdf_bi(
    caminho_entrada: str,
    pasta_destino: str | None = None
) -> str:
    """
    Gera o caminho completo para o 2º relatório em PDF contendo o Painel Executivo de BI:
    Exemplo: 'frota_ferreira_gomes.xlsx' -> 'resultados_ferreira_gomes_BI.pdf'
    """
    nome_base = gerar_nome_base_resultado(caminho_entrada)
    if not pasta_destino:
        pasta_destino = os.path.dirname(os.path.abspath(caminho_entrada)) if caminho_entrada else os.getcwd()
    return os.path.join(pasta_destino, f"{nome_base}_BI.pdf")


def _formatar_lista_placas(placas: list[str], limite: int = 15) -> str:
    """Formata uma lista de placas em português natural (ex: 'X, Y e Z')."""
    if not placas:
        return ""
    exibidas = placas[:limite]
    restante = len(placas) - len(exibidas)
    if len(exibidas) == 1:
        base = exibidas[0]
    elif len(exibidas) == 2:
        base = f"{exibidas[0]} e {exibidas[1]}"
    else:
        base = ", ".join(exibidas[:-1]) + f" e {exibidas[-1]}"
    if restante > 0:
        base += f" (+{restante} veículo(s))"
    return base


_VALORES_NAO_CONSULTADOS = {"", "-", "N/A", "NÃO CONSULTADO", "NAO CONSULTADO", "NAN", "NONE"}


def _valor_nao_consultado(valor: str) -> bool:
    """Indica se o campo de débito não foi consultado na auditoria (ex: '-')."""
    return str(valor).strip().upper() in _VALORES_NAO_CONSULTADOS


def _equalizar_nao_comparavel(
    anterior: str,
    atual: str,
    restricao_atual_sem_valor: bool = False
) -> tuple[str, str, bool]:
    """
    Quando apenas um dos lados possui valor efetivamente consultado, replica esse valor no outro lado
    para que a comparação não interprete a ausência de consulta como quitação (ou como novo débito).
    Retorna (anterior_efetivo, atual_efetivo, houve_substituicao).
    """
    ant_nc = _valor_nao_consultado(anterior)
    atu_nc = _valor_nao_consultado(atual)
    if restricao_atual_sem_valor and not atu_nc:
        atual_lower = str(atual).lower()
        if ("restriç" in atual_lower or "restric" in atual_lower) and formatar_moeda_float(anterior) > 0:
            return anterior, anterior, True
    if atu_nc and not ant_nc:
        return anterior, anterior, True
    if ant_nc and not atu_nc:
        return atual, atual, True
    return anterior, atual, False


def _total_comparavel(
    row: pd.Series,
    mapa: dict,
    valores_originais: tuple[str, str, str],
    valores_efetivos: tuple[str, str, str]
) -> float:
    """
    Total a Pagar de uma auditoria para fins de comparação.
    Usa o Total_a_Pagar gravado (quando > 0) e, se algum tributo foi equalizado por não ter sido
    consultado, ajusta esse total apenas pela diferença que o tributo equalizado representa.
    """
    total_salvo = formatar_moeda_float(row.get(mapa["total"], "0")) if mapa["total"] else 0.0
    if total_salvo <= 0:
        return calcular_total_a_pagar_veiculo(*valores_efetivos)
    if tuple(valores_originais) == tuple(valores_efetivos):
        return total_salvo
    ajuste = calcular_total_a_pagar_veiculo(*valores_efetivos) - calcular_total_a_pagar_veiculo(*valores_originais)
    return round(total_salvo + ajuste, 2)


def carregar_resultado_anterior(caminho_excel_resultado: str) -> tuple[pd.DataFrame | None, str]:
    """
    Carrega a última planilha de resultados existente antes de ser sobrescrita,
    retornando (df_anterior, data_hora_anterior_str).
    Se o arquivo não existir ou não contiver colunas de auditoria, retorna (None, "").
    """
    if not caminho_excel_resultado or not os.path.exists(caminho_excel_resultado):
        return None, ""
    try:
        df_ant = pd.read_excel(caminho_excel_resultado, dtype=str).fillna("")
        cols_upper = {str(c).strip().upper(): c for c in df_ant.columns}
        tem_col_auditoria = any(
            k in cols_upper
            for k in ["VALOR_LICENCIAMENTO", "VALOR_IPVA", "VALOR_MULTAS", "TOTAL_A_PAGAR"]
        )
        if not tem_col_auditoria or df_ant.empty:
            return None, ""

        data_ant = ""
        if "DATA_CONSULTA" in cols_upper:
            col_dt = cols_upper["DATA_CONSULTA"]
            vals_dt = [str(v).strip() for v in df_ant[col_dt] if str(v).strip() and str(v).strip() != "-"]
            if vals_dt:
                data_ant = vals_dt[-1]
        if not data_ant:
            mtime = os.path.getmtime(caminho_excel_resultado)
            data_ant = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M:%S")
        return df_ant, data_ant
    except Exception as e:
        logger.warning(f"Não foi possível ler planilha anterior para comparativo ({caminho_excel_resultado}): {e}")
        return None, ""


def comparar_auditorias(
    df_anterior: pd.DataFrame | None,
    df_atual: pd.DataFrame | None,
    data_anterior: str = "",
    data_atual: str = ""
) -> dict | None:
    """
    Compara a auditoria anterior (ex: 09h) com a auditoria atual (ex: 18h) antes da sobrescrita
    e identifica detalhadamente todos os pagamentos/baixas de IPVA, Licenciamento e Multas,
    além de eventuais novos débitos ou baixas de restrição.
    """
    if df_anterior is None or df_atual is None or df_anterior.empty or df_atual.empty:
        return None

    def _mapear_colunas(df: pd.DataFrame) -> dict[str, str | None]:
        mapa = {"placa": None, "prop": None, "lic": None, "ipva": None, "multas": None, "total": None, "status": None, "data": None}
        for c in df.columns:
            cu = str(c).strip().upper()
            if not mapa["placa"] and "PLACA" in cu:
                mapa["placa"] = c
            elif not mapa["prop"] and any(t in cu for t in ["PROPRIET", "CLIENTE", "EMPRESA", "LOCAT"]):
                mapa["prop"] = c
            elif not mapa["lic"] and "LICENCIAMENTO" in cu:
                mapa["lic"] = c
            elif not mapa["ipva"] and "IPVA" in cu:
                mapa["ipva"] = c
            elif not mapa["multas"] and "MULTA" in cu:
                mapa["multas"] = c
            elif not mapa["total"] and "TOTAL" in cu:
                mapa["total"] = c
            elif not mapa["status"] and ("STATUS" in cu or "SITUA" in cu):
                mapa["status"] = c
            elif not mapa["data"] and "DATA_CONSULTA" in cu:
                mapa["data"] = c
        return mapa

    m_ant = _mapear_colunas(df_anterior)
    m_atu = _mapear_colunas(df_atual)
    if not m_ant["placa"] or not m_atu["placa"]:
        return None

    if not data_anterior and m_ant["data"]:
        vals = [str(v).strip() for v in df_anterior[m_ant["data"]] if str(v).strip() and str(v).strip() != "-"]
        if vals:
            data_anterior = vals[-1]
    if not data_atual and m_atu["data"]:
        vals = [str(v).strip() for v in df_atual[m_atu["data"]] if str(v).strip() and str(v).strip() != "-"]
        if vals:
            data_atual = vals[-1]
    if not data_atual:
        data_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    def _linha_valida(row: pd.Series, mapa: dict) -> bool:
        st = str(row.get(mapa["status"], "") if mapa["status"] else "").lower()
        if any(t in st for t in ["não consultado", "nao consultado", "falha"]):
            return False
        vals = [
            str(row.get(mapa[k], "") if mapa[k] else "").strip()
            for k in ["lic", "ipva", "multas"]
        ]
        if all(v in ["", "-", "N/A", "Não Consultado"] for v in vals):
            return False
        if any(t in " ".join(vals).lower() for t in ["falha", "erro", "timeout", "não localizado", "nao localizado"]):
            return False
        return True

    # Indexa auditoria anterior por placa higienizada
    idx_ant = {}
    for _, row in df_anterior.iterrows():
        plc = re.sub(r"[^A-Za-z0-9]", "", str(row.get(m_ant["placa"], "")).strip().upper())
        if plc and _linha_valida(row, m_ant):
            idx_ant[plc] = row

    if not idx_ant:
        return None

    ipvas_quitados = []
    ipvas_parciais = []
    lics_quitados = []
    lics_parciais = []
    multas_pagas = []
    restricoes_baixadas = []
    acrescimos = []
    movimentacoes_detalhadas = []

    total_anterior_comparado = 0.0
    total_atual_comparado = 0.0
    veiculos_comparados = 0

    for _, row_atu in df_atual.iterrows():
        plc = re.sub(r"[^A-Za-z0-9]", "", str(row_atu.get(m_atu["placa"], "")).strip().upper())
        if not plc or plc not in idx_ant or not _linha_valida(row_atu, m_atu):
            continue

        row_ant = idx_ant[plc]
        veiculos_comparados += 1

        prop = ""
        if m_atu["prop"]:
            prop = str(row_atu.get(m_atu["prop"], "")).strip()
        elif m_ant["prop"]:
            prop = str(row_ant.get(m_ant["prop"], "")).strip()

        lic_ant_s = str(row_ant.get(m_ant["lic"], "0") if m_ant["lic"] else "0")
        ipva_ant_s = str(row_ant.get(m_ant["ipva"], "0") if m_ant["ipva"] else "0")
        mul_ant_s = str(row_ant.get(m_ant["multas"], "0") if m_ant["multas"] else "0")

        lic_atu_s = str(row_atu.get(m_atu["lic"], "0") if m_atu["lic"] else "0")
        ipva_atu_s = str(row_atu.get(m_atu["ipva"], "0") if m_atu["ipva"] else "0")
        mul_atu_s = str(row_atu.get(m_atu["multas"], "0") if m_atu["multas"] else "0")

        # Tributo não consultado em uma das auditorias (ou Licenciamento atual com Restrição, sem valor
        # informado pelo portal) não é comparável: assume-se o valor da outra auditoria (sem alteração),
        # evitando registrar pagamentos ou acréscimos que não ocorreram.
        lic_ant_e, lic_atu_e, _ = _equalizar_nao_comparavel(lic_ant_s, lic_atu_s, restricao_atual_sem_valor=True)
        ipva_ant_e, ipva_atu_e, _ = _equalizar_nao_comparavel(ipva_ant_s, ipva_atu_s)
        mul_ant_e, mul_atu_e, _ = _equalizar_nao_comparavel(mul_ant_s, mul_atu_s)

        v_lic_ant = formatar_moeda_float(lic_ant_e)
        v_ipva_ant = formatar_moeda_float(ipva_ant_e)
        v_mul_ant = formatar_moeda_float(mul_ant_e)
        v_tot_ant = _total_comparavel(
            row_ant, m_ant, (lic_ant_s, ipva_ant_s, mul_ant_s), (lic_ant_e, ipva_ant_e, mul_ant_e)
        )

        v_lic_atu = formatar_moeda_float(lic_atu_e)
        v_ipva_atu = formatar_moeda_float(ipva_atu_e)
        v_mul_atu = formatar_moeda_float(mul_atu_e)
        v_tot_atu = _total_comparavel(
            row_atu, m_atu, (lic_atu_s, ipva_atu_s, mul_atu_s), (lic_atu_e, ipva_atu_e, mul_atu_e)
        )

        total_anterior_comparado += v_tot_ant
        total_atual_comparado += v_tot_atu

        # 1. Comparação de IPVA
        diff_ipva = round(v_ipva_ant - v_ipva_atu, 2)
        if diff_ipva >= 0.01:
            info = {"placa": plc, "prop": prop, "anterior": v_ipva_ant, "atual": v_ipva_atu, "pago": diff_ipva}
            if v_ipva_atu == 0.0:
                ipvas_quitados.append(info)
                movimentacoes_detalhadas.append({
                    "placa": plc, "prop": prop, "tributo": "IPVA (Quitado)",
                    "anterior": formatar_float_moeda(v_ipva_ant), "atual": "PAGO",
                    "variacao": f"- {formatar_float_moeda(diff_ipva)}", "tipo": "BAIXA"
                })
            else:
                ipvas_parciais.append(info)
                movimentacoes_detalhadas.append({
                    "placa": plc, "prop": prop, "tributo": "IPVA (Baixa Parcial)",
                    "anterior": formatar_float_moeda(v_ipva_ant), "atual": formatar_float_moeda(v_ipva_atu),
                    "variacao": f"- {formatar_float_moeda(diff_ipva)}", "tipo": "BAIXA"
                })
        elif diff_ipva <= -0.01:
            acr = abs(diff_ipva)
            acrescimos.append({"placa": plc, "prop": prop, "tributo": "IPVA", "anterior": v_ipva_ant, "atual": v_ipva_atu, "acrescimo": acr})
            movimentacoes_detalhadas.append({
                "placa": plc, "prop": prop, "tributo": "IPVA (Acréscimo)",
                "anterior": formatar_float_moeda(v_ipva_ant), "atual": formatar_float_moeda(v_ipva_atu),
                "variacao": f"+ {formatar_float_moeda(acr)}", "tipo": "ACRESCIMO"
            })

        # 2. Comparação de Multas
        diff_mul = round(v_mul_ant - v_mul_atu, 2)
        if diff_mul >= 0.01:
            info_m = {
                "placa": plc, "prop": prop, "anterior": v_mul_ant,
                "atual": v_mul_atu, "pago": diff_mul, "quitado": (v_mul_atu == 0.0)
            }
            multas_pagas.append(info_m)
            movimentacoes_detalhadas.append({
                "placa": plc, "prop": prop,
                "tributo": "Multas (Quitadas)" if v_mul_atu == 0.0 else "Multas (Baixa Parcial)",
                "anterior": formatar_float_moeda(v_mul_ant),
                "atual": "R$ 0,00" if v_mul_atu == 0.0 else formatar_float_moeda(v_mul_atu),
                "variacao": f"- {formatar_float_moeda(diff_mul)}", "tipo": "BAIXA"
            })
        elif diff_mul <= -0.01:
            acr = abs(diff_mul)
            acrescimos.append({"placa": plc, "prop": prop, "tributo": "Multas", "anterior": v_mul_ant, "atual": v_mul_atu, "acrescimo": acr})
            movimentacoes_detalhadas.append({
                "placa": plc, "prop": prop, "tributo": "Multas (Nova/Acréscimo)",
                "anterior": formatar_float_moeda(v_mul_ant), "atual": formatar_float_moeda(v_mul_atu),
                "variacao": f"+ {formatar_float_moeda(acr)}", "tipo": "ACRESCIMO"
            })

        # 3. Comparação de Licenciamento
        diff_lic = round(v_lic_ant - v_lic_atu, 2)
        if diff_lic >= 0.01:
            if v_lic_atu == 0.0:
                # Se parte do licenciamento anterior era multa vencida que também foi baixada,
                # destaca o valor líquido das taxas de licenciamento quitadas
                pago_lic_liq = max(0.0, round(diff_lic - max(0.0, diff_mul), 2)) if diff_mul > 0 else diff_lic
                if pago_lic_liq == 0.0:
                    pago_lic_liq = diff_lic
                lics_quitados.append({
                    "placa": plc, "prop": prop, "anterior": v_lic_ant,
                    "atual": 0.0, "pago": pago_lic_liq
                })
                movimentacoes_detalhadas.append({
                    "placa": plc, "prop": prop, "tributo": "Licenciamento (Quitado)",
                    "anterior": formatar_float_moeda(v_lic_ant), "atual": "PAGO",
                    "variacao": f"- {formatar_float_moeda(pago_lic_liq)}", "tipo": "BAIXA"
                })
            else:
                # Redução parcial no licenciamento além de eventual baixa de multa vencida
                reducao_liquida_lic = round(diff_lic - max(0.0, diff_mul), 2)
                if reducao_liquida_lic >= 0.05:
                    lics_parciais.append({
                        "placa": plc, "prop": prop, "anterior": v_lic_ant,
                        "atual": v_lic_atu, "pago": reducao_liquida_lic
                    })
                    movimentacoes_detalhadas.append({
                        "placa": plc, "prop": prop, "tributo": "Licenciamento (Baixa Parcial)",
                        "anterior": formatar_float_moeda(v_lic_ant), "atual": formatar_float_moeda(v_lic_atu),
                        "variacao": f"- {formatar_float_moeda(reducao_liquida_lic)}", "tipo": "BAIXA"
                    })
        elif diff_lic <= -0.01:
            acr_lic = round(abs(diff_lic) - max(0.0, -diff_mul), 2)
            if acr_lic >= 0.05:
                acrescimos.append({"placa": plc, "prop": prop, "tributo": "Licenciamento", "anterior": v_lic_ant, "atual": v_lic_atu, "acrescimo": acr_lic})
                movimentacoes_detalhadas.append({
                    "placa": plc, "prop": prop, "tributo": "Licenciamento (Acréscimo)",
                    "anterior": formatar_float_moeda(v_lic_ant), "atual": formatar_float_moeda(v_lic_atu),
                    "variacao": f"+ {formatar_float_moeda(acr_lic)}", "tipo": "ACRESCIMO"
                })

        # 4. Baixa de Restrição de Impedimento
        tinha_restr = "restriç" in lic_ant_s.lower() or "restric" in lic_ant_s.lower()
        tem_restr = "restriç" in lic_atu_s.lower() or "restric" in lic_atu_s.lower()
        if tinha_restr and not tem_restr:
            restricoes_baixadas.append({"placa": plc, "prop": prop})
            movimentacoes_detalhadas.append({
                "placa": plc, "prop": prop, "tributo": "Restrição Detran",
                "anterior": "Restrição", "atual": "Regularizado" if v_lic_atu == 0 else formatar_float_moeda(v_lic_atu),
                "variacao": "Baixada", "tipo": "BAIXA"
            })

    if veiculos_comparados == 0:
        return None

    # Construção das frases executivas em linguagem natural
    frases_resumo = []

    if ipvas_quitados:
        placas_ipva = [x["placa"] for x in ipvas_quitados]
        soma_ipva_q = sum(x["pago"] for x in ipvas_quitados)
        str_placas = _formatar_lista_placas(placas_ipva)
        if len(placas_ipva) == 1:
            frases_resumo.append(
                f"Foi pago o IPVA da placa {str_placas} (valor abatido: {formatar_float_moeda(soma_ipva_q)})."
            )
        else:
            frases_resumo.append(
                f"Foram pagos os IPVAs das placas {str_placas} (total abatido: {formatar_float_moeda(soma_ipva_q)})."
            )

    for item in ipvas_parciais:
        frases_resumo.append(
            f"Foram pagos {formatar_float_moeda(item['pago'])} de IPVA da placa {item['placa']} "
            f"(saldo reduzido de {formatar_float_moeda(item['anterior'])} para {formatar_float_moeda(item['atual'])})."
        )

    if lics_quitados:
        placas_lic = [x["placa"] for x in lics_quitados]
        soma_lic_q = sum(x["pago"] for x in lics_quitados)
        str_placas = _formatar_lista_placas(placas_lic)
        if len(placas_lic) == 1:
            frases_resumo.append(
                f"Foi pago o Licenciamento da placa {str_placas} (valor abatido: {formatar_float_moeda(soma_lic_q)})."
            )
        else:
            frases_resumo.append(
                f"Foram pagos os Licenciamentos das placas {str_placas} (total abatido: {formatar_float_moeda(soma_lic_q)})."
            )

    for item in lics_parciais:
        frases_resumo.append(
            f"Foram pagos {formatar_float_moeda(item['pago'])} em taxas de Licenciamento da placa {item['placa']} "
            f"(reduzido de {formatar_float_moeda(item['anterior'])} para {formatar_float_moeda(item['atual'])})."
        )

    for item in multas_pagas:
        if item["quitado"]:
            frases_resumo.append(
                f"Foram pagos {formatar_float_moeda(item['pago'])} em multas da placa {item['placa']} (quitando 100% das multas)."
            )
        else:
            frases_resumo.append(
                f"Foram pagos {formatar_float_moeda(item['pago'])} em multas da placa {item['placa']} "
                f"(saldo de multas reduzido de {formatar_float_moeda(item['anterior'])} para {formatar_float_moeda(item['atual'])})."
            )

    if restricoes_baixadas:
        placas_r = _formatar_lista_placas([x["placa"] for x in restricoes_baixadas])
        frases_resumo.append(f"Restrição de impedimento baixada/regularizada para a(s) placa(s) {placas_r}.")

    if acrescimos:
        placas_a = _formatar_lista_placas(list(dict.fromkeys(x["placa"] for x in acrescimos)))
        soma_acr = sum(x["acrescimo"] for x in acrescimos)
        frases_resumo.append(
            f"Atenção: identificado acréscimo de {formatar_float_moeda(soma_acr)} em novos débitos/atualizações na(s) placa(s) {placas_a}."
        )

    total_pago_ipva = round(sum(x["pago"] for x in ipvas_quitados) + sum(x["pago"] for x in ipvas_parciais), 2)
    total_pago_lic = round(sum(x["pago"] for x in lics_quitados) + sum(x["pago"] for x in lics_parciais), 2)
    total_pago_multas = round(sum(x["pago"] for x in multas_pagas), 2)
    saldo_abatido_liquido = round(max(0.0, total_anterior_comparado - total_atual_comparado), 2)

    teve_alteracao = bool(movimentacoes_detalhadas)
    if not teve_alteracao:
        frases_resumo.append(
            f"Nenhuma alteração financeira (pagamentos ou novos débitos) foi detectada entre a auditoria de {data_anterior or 'anterior'} e a atual."
        )

    return {
        "data_anterior": data_anterior or "Auditoria Anterior",
        "data_atual": data_atual,
        "veiculos_comparados": veiculos_comparados,
        "teve_alteracao": teve_alteracao,
        "total_anterior": round(total_anterior_comparado, 2),
        "total_atual": round(total_atual_comparado, 2),
        "saldo_abatido_liquido": saldo_abatido_liquido,
        "total_pago_ipva": total_pago_ipva,
        "total_pago_lic": total_pago_lic,
        "total_pago_multas": total_pago_multas,
        "ipvas_quitados": ipvas_quitados,
        "ipvas_parciais": ipvas_parciais,
        "lics_quitados": lics_quitados,
        "lics_parciais": lics_parciais,
        "multas_pagas": multas_pagas,
        "restricoes_baixadas": restricoes_baixadas,
        "acrescimos": acrescimos,
        "frases_resumo": frases_resumo,
        "movimentacoes": movimentacoes_detalhadas,
    }


