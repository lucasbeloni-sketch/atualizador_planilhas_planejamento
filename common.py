"""
Funções e configurações compartilhadas entre os blocos 1, 2 e 3.

Centraliza autenticação, leitura/escrita no Google Sheets, retry e os
helpers de manipulação de matrizes. Corrigir um bug aqui corrige nos três
blocos de uma vez.
"""
import base64
import json
import os
import random
import time
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
import requests
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, WorksheetNotFound
from gspread.utils import rowcol_to_a1, a1_to_rowcol


# =========================================================
# CONFIGURAÇÕES COMUNS
# =========================================================
LISTA_PLANILHAS_SPREADSHEET_ID = os.getenv(
    "LISTA_PLANILHAS_SPREADSHEET_ID",
    "1kMJedysNlxxPU2PtCwICHlBbZVL4YpvyHSsR7Xl71Ig",
)

ABA_LISTA_PLANILHAS = os.getenv(
    "ABA_LISTA_PLANILHAS",
    "BD_Planilhas",
)

TIMEZONE = ZoneInfo("America/Sao_Paulo")

CHUNK_SIZE = int(os.getenv("CHUNK_SIZE", "5000"))

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# =========================================================
# AUTENTICAÇÃO
# =========================================================
def get_gspread_client() -> gspread.Client:
    """
    Aceita credencial de duas formas:
    1) Secret GOOGLE_CREDENTIALS_B64 no GitHub Actions
    2) Arquivo local service_account.json, para teste local
    """
    credentials_b64 = os.getenv("GOOGLE_CREDENTIALS_B64", "").strip()

    if credentials_b64:
        service_account_info = json.loads(base64.b64decode(credentials_b64).decode("utf-8"))
        credentials = Credentials.from_service_account_info(
            service_account_info,
            scopes=SCOPES,
        )
    else:
        credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
        credentials = Credentials.from_service_account_file(
            credentials_path,
            scopes=SCOPES,
        )

    return gspread.authorize(credentials)


# =========================================================
# RETRY
# =========================================================
def _retry_after_segundos(erro) -> float | None:
    """Lê o header Retry-After da resposta, se houver."""
    resposta = getattr(erro, "response", None)

    if resposta is None:
        return None

    retry_after = resposta.headers.get("Retry-After") if resposta.headers else None

    if not retry_after:
        return None

    try:
        return float(retry_after)
    except (TypeError, ValueError):
        return None


def executar_com_retry(
    func,
    tentativas: int = 5,
    espera_inicial: float = 2.0,
    espera_maxima: float = 60.0,
):
    """
    Executa func com retry em erros de API e de rede.

    Backoff exponencial com jitter, respeitando o header Retry-After
    quando o Google o envia (rate limit 429 / indisponibilidade 503).
    """
    ultimo_erro = None

    for tentativa in range(1, tentativas + 1):
        try:
            return func()
        except (APIError, requests.exceptions.RequestException) as erro:
            ultimo_erro = erro

            if tentativa == tentativas:
                raise

            espera = min(espera_inicial * (2 ** (tentativa - 1)), espera_maxima)
            espera += random.uniform(0, espera_inicial)

            retry_after = _retry_after_segundos(erro)
            if retry_after is not None:
                espera = max(espera, retry_after)

            print(
                f"[AVISO] Erro Google API/rede ({type(erro).__name__}). "
                f"Tentativa {tentativa}/{tentativas}. Nova tentativa em {espera:.0f}s."
            )
            time.sleep(espera)

    raise ultimo_erro


# =========================================================
# HELPERS DE VALOR
# =========================================================
def is_blank(valor) -> bool:
    return valor is None or valor == ""


def as_text(valor) -> str:
    if valor is None:
        return ""

    return str(valor)


def pad_row(row: list, n_cols: int) -> list:
    row = list(row or [])

    if len(row) < n_cols:
        row += [""] * (n_cols - len(row))

    return row[:n_cols]


def pad_matrix(values: list[list], n_rows: int, n_cols: int) -> list[list]:
    saida = []

    for i in range(n_rows):
        row = values[i] if i < len(values) else []
        saida.append(pad_row(row, n_cols))

    return saida


def dimensoes_range(range_a1: str) -> tuple[int, int, int, int]:
    """
    Retorna:
    linha_inicial, coluna_inicial, qtd_linhas, qtd_colunas
    """
    if ":" not in range_a1:
        row, col = a1_to_rowcol(range_a1)
        return row, col, 1, 1

    inicio, fim = range_a1.split(":")
    row_ini, col_ini = a1_to_rowcol(inicio)
    row_fim, col_fim = a1_to_rowcol(fim)

    qtd_linhas = row_fim - row_ini + 1
    qtd_colunas = col_fim - col_ini + 1

    return row_ini, col_ini, qtd_linhas, qtd_colunas


# =========================================================
# ABAS
# =========================================================
def abrir_aba(spreadsheet: gspread.Spreadsheet, nome_aba: str) -> gspread.Worksheet:
    try:
        return spreadsheet.worksheet(nome_aba)
    except WorksheetNotFound as erro:
        raise RuntimeError(
            f"A aba '{nome_aba}' não foi encontrada na planilha '{spreadsheet.title}'."
        ) from erro


# =========================================================
# LEITURA
# =========================================================
def ler_range(
    worksheet: gspread.Worksheet,
    range_a1: str,
    n_rows: int | None = None,
    n_cols: int | None = None,
    date_time_render_option: str = "FORMATTED_STRING",
) -> list[list]:
    valores = executar_com_retry(
        lambda: worksheet.get(
            range_a1,
            value_render_option="UNFORMATTED_VALUE",
            date_time_render_option=date_time_render_option,
        )
    )

    if n_rows is None and n_cols is None:
        return valores

    if n_rows is None:
        n_rows = len(valores)

    if n_cols is None:
        n_cols = max((len(row) for row in valores), default=0)

    return pad_matrix(valores, n_rows, n_cols)


def ler_coluna(worksheet: gspread.Worksheet, col: int) -> list:
    valores = executar_com_retry(
        lambda: worksheet.col_values(
            col,
            value_render_option="UNFORMATTED_VALUE",
        )
    )
    return valores


def ultima_linha_preenchida_por_coluna(worksheet: gspread.Worksheet, col: int) -> int:
    valores = ler_coluna(worksheet, col)

    for idx in range(len(valores) - 1, -1, -1):
        if not is_blank(valores[idx]):
            return idx + 1

    return 0


# =========================================================
# ESCRITA / LIMPEZA
# =========================================================
def limpar_intervalos(worksheet: gspread.Worksheet, ranges: list[str]) -> None:
    if not ranges:
        return

    executar_com_retry(lambda: worksheet.batch_clear(ranges))


def escrever_celula(worksheet: gspread.Worksheet, a1: str, valor, raw: bool = True) -> None:
    value_input_option = "RAW" if raw else "USER_ENTERED"

    executar_com_retry(
        lambda: worksheet.update(
            range_name=a1,
            values=[[valor]],
            value_input_option=value_input_option,
        )
    )


def escrever_matriz(
    worksheet: gspread.Worksheet,
    start_row: int,
    start_col: int,
    values: list[list],
    raw: bool = True,
    chunk_size: int = CHUNK_SIZE,
) -> None:
    if not values:
        return

    value_input_option = "RAW" if raw else "USER_ENTERED"
    total_cols = max(len(row) for row in values)

    for offset in range(0, len(values), chunk_size):
        bloco = values[offset : offset + chunk_size]

        row_ini = start_row + offset
        row_fim = row_ini + len(bloco) - 1
        col_fim = start_col + total_cols - 1

        range_a1 = f"{rowcol_to_a1(row_ini, start_col)}:{rowcol_to_a1(row_fim, col_fim)}"
        bloco_padronizado = [pad_row(row, total_cols) for row in bloco]

        executar_com_retry(
            lambda range_a1=range_a1, bloco_padronizado=bloco_padronizado: worksheet.update(
                range_name=range_a1,
                values=bloco_padronizado,
                value_input_option=value_input_option,
            )
        )


def escrever_formulas_matriz(
    worksheet: gspread.Worksheet,
    start_row: int,
    start_col: int,
    formulas: list[list[str]],
    chunk_size: int = CHUNK_SIZE,
) -> None:
    if not formulas:
        return

    total_cols = max(len(row) for row in formulas)

    for offset in range(0, len(formulas), chunk_size):
        bloco = formulas[offset : offset + chunk_size]

        row_ini = start_row + offset
        row_fim = row_ini + len(bloco) - 1
        col_fim = start_col + total_cols - 1

        range_a1 = f"{rowcol_to_a1(row_ini, start_col)}:{rowcol_to_a1(row_fim, col_fim)}"
        bloco_padronizado = [pad_row(row, total_cols) for row in bloco]

        executar_com_retry(
            lambda range_a1=range_a1, bloco_padronizado=bloco_padronizado: worksheet.update(
                range_name=range_a1,
                values=bloco_padronizado,
                value_input_option="USER_ENTERED",
            )
        )


def congelar_intervalo(
    worksheet: gspread.Worksheet,
    range_a1: str,
    date_time_render_option: str = "FORMATTED_STRING",
) -> None:
    """
    Lê os resultados calculados das fórmulas e cola como valores.
    """
    row_ini, col_ini, qtd_linhas, qtd_colunas = dimensoes_range(range_a1)

    valores = ler_range(
        worksheet=worksheet,
        range_a1=range_a1,
        n_rows=qtd_linhas,
        n_cols=qtd_colunas,
        date_time_render_option=date_time_render_option,
    )

    escrever_matriz(
        worksheet=worksheet,
        start_row=row_ini,
        start_col=col_ini,
        values=valores,
        raw=True,
    )


def formatar_data_hora(worksheet: gspread.Worksheet, range_a1: str) -> None:
    try:
        executar_com_retry(
            lambda: worksheet.format(
                range_a1,
                {
                    "numberFormat": {
                        "type": "DATE_TIME",
                        "pattern": "dd/MM/yyyy HH:mm:ss",
                    }
                },
            )
        )
    except Exception as erro:
        print(f"[AVISO] Não foi possível aplicar formato de data/hora em {range_a1}: {erro}")


# =========================================================
# LISTA DE PLANILHAS (aba BD_Planilhas)
# =========================================================
def buscar_planilhas(ss_lista: gspread.Spreadsheet) -> list[dict[str, str]]:
    """
    Busca os dados das planilhas na aba BD_Planilhas, a partir da linha 3.
    Nome: coluna B. ID: coluna C. Valor BE: coluna D.
    Remove IDs vazios e duplicados.

    Sempre retorna a chave "valor_be" (usada apenas pelo Bloco 3); os demais
    blocos a ignoram.
    """
    aba_lista = abrir_aba(ss_lista, ABA_LISTA_PLANILHAS)

    ultima_linha = ultima_linha_preenchida_por_coluna(aba_lista, 3)

    if ultima_linha < 3:
        return []

    valores = ler_range(
        aba_lista,
        f"B3:D{ultima_linha}",
        ultima_linha - 2,
        3,
    )

    planilhas = []
    ids_vistos = set()

    for row in valores:
        nome_planilha = as_text(row[0]).strip() if len(row) > 0 else ""
        id_planilha = as_text(row[1]).strip() if len(row) > 1 else ""
        valor_be = as_text(row[2]).strip() if len(row) > 2 else ""

        if not id_planilha:
            continue

        if id_planilha in ids_vistos:
            continue

        if not nome_planilha:
            nome_planilha = "Sem nome informado"

        planilhas.append(
            {
                "nome": nome_planilha,
                "id": id_planilha,
                "valor_be": valor_be,
            }
        )

        ids_vistos.add(id_planilha)

    return planilhas


def agora_formatado() -> str:
    """Data/hora atual no fuso configurado, formato dd/MM/yyyy HH:mm:ss."""
    return datetime.now(TIMEZONE).strftime("%d/%m/%Y %H:%M:%S")
