"""
Processamento, higienização e validação de dados de frotas com Pandas.
Regras financeiras de formatação para relatórios e exibição.
"""
import os
import re
from datetime import datetime, timedelta
import pandas as pd

# Padrão oficial de placas veiculares brasileiras (Mercosul e Tradicional)
PADRAO_PLACA = re.compile(r"^[A-Z]{3}[0-9][A-Z0-9][0-9]{2}$|^[A-Z]{3}[0-9]{4}$")


def formatar_moeda_float(valor_str: str) -> float:
    """Converte valores monetários textuais (ex: 'R$ 188,96' ou '188.96') em float 188.96."""
    try:
        if not valor_str:
            return 0.0
        val_lower = str(valor_str).lower().strip()
        termos_ignorar = [
            "-", "n/a", "não consultado", "falha", "erro", "aguardando",
            "inválid", "não localizado", "timeout", "...", "pago", "restrição", "restricao"
        ]
        if any(termo in val_lower for termo in termos_ignorar):
            return 0.0
        limpo = val_lower.replace("r$", "").replace(".", "").replace(",", ".").strip()
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


import calendar


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

    # Feriados nacionais comuns no Brasil (dia, mês)
    feriados = [
        (1, 1), (21, 4), (1, 5), (7, 9), (12, 10), (2, 11), (15, 11), (25, 12)
    ]

    while d.weekday() >= 5 or (d.day, d.month) in feriados:
        d -= timedelta(days=1)

    return d.strftime("%d/%m/%Y")


def obter_proximo_dia_util(data_base: datetime | None = None) -> str:
    """Mantido para compatibilidade, retornando o último dia útil do mês."""
    return obter_ultimo_dia_util_mes(data_base)


class FrotaDataProcessor:
    """Higienização de dados, validação de regras de frota e consolidação."""

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
        - Linhas vazias sejam removidas
        - Valores de RENAVAM não percam zeros à esquerda (zfill 11)
        - Placas sejam higienizadas
        """
        df = pd.read_excel(caminho_arquivo, dtype=str)
        if len(df) == 0:
            raise ValueError("O arquivo Excel selecionado não contém nenhum registro.")

        col_placa, col_renavam = cls.localizar_colunas_chave(df)

        # 1. Remove linhas onde placa ou renavam são nulos
        df = df.dropna(subset=[col_placa, col_renavam]).copy()

        # 2. Higienização de Placa (remove hífens, espaços e transforma em maiúsculo)
        df[col_placa] = (
            df[col_placa]
            .astype(str)
            .str.strip()
            .str.upper()
            .str.replace("-", "", regex=False)
            .str.replace(" ", "", regex=False)
        )

        # 3. Higienização de RENAVAM (preserva 11 dígitos com zeros à esquerda)
        def sanitizar_renavam(val):
            val_str = str(val).strip()
            if val_str.endswith(".0"):
                val_str = val_str[:-2]
            digitos = "".join(c for c in val_str if c.isdigit())
            if digitos and len(digitos) <= 11:
                return digitos.zfill(11)
            return digitos or val_str

        df[col_renavam] = df[col_renavam].apply(sanitizar_renavam)

        # 4. Remove linhas com campos 'NAN', 'NONE', vazios ou hífens
        df = df[
            ~df[col_placa].isin(["NAN", "NONE", "", "N/A", "-", "NULL"]) &
            ~df[col_renavam].isin(["NAN", "NONE", "", "N/A", "-", "NULL"])
        ].copy()

        if len(df) == 0:
            raise ValueError("Nenhum registro com Placa e RENAVAM válidos foi encontrado na planilha.")

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
        Constrói o DataFrame final preservando todas as colunas originais do usuário
        e adicionando as colunas oficiais de débitos, total real a pagar e situação.
        Evita duplicação de multas vencidas já inclusas no Licenciamento.
        """
        df_final = df_original.iloc[:qtd_processada].copy()

        def normalizar_valor_pago(v: str) -> str:
            v_str = str(v).strip()
            v_lower = v_str.lower()
            if v_str in ["R$ 0,00", "0,00", "0.00", "0"]:
                return "PAGO"
            if formatar_moeda_float(v_str) == 0.0 and v_str not in ["-", "N/A", "Não Consultado"] and not any(
                f in v_lower for f in ["falha", "erro", "inválid", "timeout", "não localizado", "restriç", "restric"]
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

            tem_falha = any(
                termo in v.lower()
                for v in [lic, ipva, multas]
                for termo in ["falha", "erro", "timeout", "inválid", "não localizado"]
            )

            total_real = calcular_total_a_pagar_veiculo(lic, ipva, multas, m_a_vencer)
            totais_a_pagar_lista.append(formatar_float_moeda(total_real))

            if tem_falha:
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


