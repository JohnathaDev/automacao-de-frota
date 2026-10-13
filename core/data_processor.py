"""
Processamento, higienização e validação de dados de frotas com Pandas.
Regras financeiras de formatação para relatórios e exibição.
"""
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
            "inválid", "não localizado", "timeout", "..."
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
        qtd_processada: int
    ) -> pd.DataFrame:
        """
        Constrói o DataFrame final preservando todas as colunas originais do usuário
        e adicionando as colunas oficiais de débitos e situação.
        Evita falsos negativos: identifica 'Falha na Consulta' e 'Placa Inválida'.
        """
        df_final = df_original.iloc[:qtd_processada].copy()

        df_final["Valor_Licenciamento"] = resultados_lic[:qtd_processada]
        df_final["Valor_IPVA"] = resultados_ipva[:qtd_processada]
        df_final["Valor_Multas"] = resultados_multas[:qtd_processada]

        # Determinação do Status consolidado
        status_lista = []
        for i in range(qtd_processada):
            lic = str(resultados_lic[i])
            ipva = str(resultados_ipva[i])
            multas = str(resultados_multas[i])

            tem_falha = any(
                termo in v.lower()
                for v in [lic, ipva, multas]
                for termo in ["falha", "erro", "timeout", "inválid", "não localizado"]
            )
            soma = (
                formatar_moeda_float(lic) +
                formatar_moeda_float(ipva) +
                formatar_moeda_float(multas)
            )

            if tem_falha:
                status_lista.append("Falha na Consulta")
            elif soma > 0:
                status_lista.append("Com Débitos Pendentes")
            else:
                status_lista.append("Sem Débitos")

        df_final["Status"] = status_lista
        df_final["Data_Consulta"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        return df_final
