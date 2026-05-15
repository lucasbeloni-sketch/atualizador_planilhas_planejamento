import base64
import json
import os
import time
import traceback
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, WorksheetNotFound
from gspread.utils import rowcol_to_a1, a1_to_rowcol


# =========================================================
# CONFIGURAÇÕES
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

# Tempo de espera para o Google Sheets calcular as fórmulas antes de congelar.
# Se alguma fórmula for mais pesada, pode aumentar no workflow via env.
CALC_WAIT_SECONDS = int(os.getenv("CALC_WAIT_SECONDS", "15"))

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
# HELPERS
# =========================================================
def executar_com_retry(func, tentativas: int = 5, espera_inicial: float = 2.0):
    ultimo_erro = None

    for tentativa in range(1, tentativas + 1):
        try:
            return func()
        except APIError as erro:
            ultimo_erro = erro

            if tentativa == tentativas:
                raise

            espera = espera_inicial * tentativa
            print(
                f"[AVISO] Erro Google API. "
                f"Tentativa {tentativa}/{tentativas}. Nova tentativa em {espera:.0f}s."
            )
            time.sleep(espera)

    raise ultimo_erro


def abrir_aba(spreadsheet: gspread.Spreadsheet, nome_aba: str) -> gspread.Worksheet:
    try:
        return spreadsheet.worksheet(nome_aba)
    except WorksheetNotFound as erro:
        raise RuntimeError(
            f"A aba '{nome_aba}' não foi encontrada na planilha '{spreadsheet.title}'."
        ) from erro


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


def ler_range(
    worksheet: gspread.Worksheet,
    range_a1: str,
    n_rows: int | None = None,
    n_cols: int | None = None,
) -> list[list]:
    valores = executar_com_retry(
        lambda: worksheet.get(
            range_a1,
            value_render_option="UNFORMATTED_VALUE",
            date_time_render_option="FORMATTED_STRING",
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


def copiar_formula_com_referencia_ajustada(
    spreadsheet: gspread.Spreadsheet,
    worksheet: gspread.Worksheet,
    source_start_col: int,
    source_end_col: int,
    dest_start_row: int,
    dest_end_row: int,
) -> None:
    """
    Replica a lógica do Apps Script:
    worksheet.getRange('B1:G1').copyTo(B6:GlastRow, PASTE_FORMULA)

    Usa CopyPasteRequest da API do Google Sheets, que ajusta referências relativas.
    """
    if dest_end_row < dest_start_row:
        return

    sheet_id = worksheet.id

    request = {
        "requests": [
            {
                "copyPaste": {
                    "source": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": source_start_col - 1,
                        "endColumnIndex": source_end_col,
                    },
                    "destination": {
                        "sheetId": sheet_id,
                        "startRowIndex": dest_start_row - 1,
                        "endRowIndex": dest_end_row,
                        "startColumnIndex": source_start_col - 1,
                        "endColumnIndex": source_end_col,
                    },
                    "pasteType": "PASTE_FORMULA",
                    "pasteOrientation": "NORMAL",
                }
            }
        ]
    }

    executar_com_retry(lambda: spreadsheet.batch_update(request))


def congelar_intervalo(worksheet: gspread.Worksheet, range_a1: str) -> None:
    """
    Lê os resultados calculados das fórmulas e cola como valores.
    """
    row_ini, col_ini, qtd_linhas, qtd_colunas = dimensoes_range(range_a1)

    valores = ler_range(
        worksheet=worksheet,
        range_a1=range_a1,
        n_rows=qtd_linhas,
        n_cols=qtd_colunas,
    )

    escrever_matriz(
        worksheet=worksheet,
        start_row=row_ini,
        start_col=col_ini,
        values=valores,
        raw=True,
    )


def copiar_calcular_e_congelar(
    spreadsheet: gspread.Spreadsheet,
    worksheet: gspread.Worksheet,
    source_start_col: int,
    source_end_col: int,
    dest_start_row: int,
    dest_end_row: int,
    nome_bloco: str,
) -> None:
    if dest_end_row < dest_start_row:
        return

    range_destino = (
        f"{rowcol_to_a1(dest_start_row, source_start_col)}:"
        f"{rowcol_to_a1(dest_end_row, source_end_col)}"
    )

    print(f"Aplicando fórmulas em {nome_bloco}: {range_destino}")

    copiar_formula_com_referencia_ajustada(
        spreadsheet=spreadsheet,
        worksheet=worksheet,
        source_start_col=source_start_col,
        source_end_col=source_end_col,
        dest_start_row=dest_start_row,
        dest_end_row=dest_end_row,
    )

    print(f"Aguardando cálculo das fórmulas em {nome_bloco} por {CALC_WAIT_SECONDS}s...")
    time.sleep(CALC_WAIT_SECONDS)

    print(f"Congelando valores em {nome_bloco}: {range_destino}")
    congelar_intervalo(worksheet, range_destino)


def buscar_planilhas(ss_lista: gspread.Spreadsheet) -> list[dict[str, str]]:
    """
    Busca os nomes e IDs das planilhas na aba BD_Planilhas.
    Nome: coluna B, a partir da linha 3.
    ID: coluna C, a partir da linha 3.
    Remove vazios e IDs duplicados.
    """
    aba_lista = abrir_aba(ss_lista, ABA_LISTA_PLANILHAS)

    ultima_linha = ultima_linha_preenchida_por_coluna(aba_lista, 3)

    if ultima_linha < 3:
        return []

    valores = ler_range(
        aba_lista,
        f"B3:C{ultima_linha}",
        ultima_linha - 2,
        2,
    )

    planilhas = []
    ids_vistos = set()

    for row in valores:
        nome_planilha = as_text(row[0]).strip() if len(row) > 0 else ""
        id_planilha = as_text(row[1]).strip() if len(row) > 1 else ""

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
            }
        )

        ids_vistos.add(id_planilha)

    return planilhas


# =========================================================
# BLOCO 2 - Carteira_Planejador
# =========================================================
def executar_bloco2_carteira_planejador(ss_dest: gspread.Spreadsheet) -> None:
    worksheet = abrir_aba(ss_dest, "Carteira_Planejador")

    print("Atualizando aba Carteira_Planejador...")

    escrever_celula(worksheet, "F3", "Em Atualização")

    last_row = ultima_linha_preenchida_por_coluna(worksheet, 13)

    print(f"Última linha preenchida pela coluna M: {last_row}")

    if last_row < 6:
        finalizar_execucao(worksheet)
        print("Nenhuma linha para atualizar a partir da linha 6.")
        return

    # =====================================================
    # Limpezas equivalentes ao Apps Script
    # =====================================================
    limpar_intervalos(
        worksheet,
        [
            f"B6:G{last_row}",
            f"P6:P{last_row}",
            f"S6:AF{last_row}",
            f"AI6:AJ{last_row}",
            f"AM6:AV{last_row}",
            f"AX6:BD{last_row}",
            f"BF6:BF{last_row}",
        ],
    )

    # =====================================================
    # Parte 1
    # B1:G1 -> B6:GlastRow
    # P1:P1 -> P6:PlastRow
    # =====================================================
    print("Executando lógica Parte 1...")

    copiar_calcular_e_congelar(
        spreadsheet=ss_dest,
        worksheet=worksheet,
        source_start_col=2,
        source_end_col=7,
        dest_start_row=6,
        dest_end_row=last_row,
        nome_bloco="B:G",
    )

    copiar_calcular_e_congelar(
        spreadsheet=ss_dest,
        worksheet=worksheet,
        source_start_col=16,
        source_end_col=16,
        dest_start_row=6,
        dest_end_row=last_row,
        nome_bloco="P:P",
    )

    # =====================================================
    # Parte 2
    # BF1:BF1 -> BF6:BFlastRow
    # AI1:AJ1 -> AI6:AJlastRow
    # =====================================================
    print("Executando lógica Parte 2...")

    copiar_calcular_e_congelar(
        spreadsheet=ss_dest,
        worksheet=worksheet,
        source_start_col=58,
        source_end_col=58,
        dest_start_row=6,
        dest_end_row=last_row,
        nome_bloco="BF:BF",
    )

    copiar_calcular_e_congelar(
        spreadsheet=ss_dest,
        worksheet=worksheet,
        source_start_col=35,
        source_end_col=36,
        dest_start_row=6,
        dest_end_row=last_row,
        nome_bloco="AI:AJ",
    )

    # =====================================================
    # Parte 3
    # S1:AF1 -> S6:AFlastRow
    # AM1:AV1 -> AM6:AVlastRow
    # AX1:BD1 -> AX6:BDlastRow
    # =====================================================
    print("Executando lógica Parte 3...")

    copiar_calcular_e_congelar(
        spreadsheet=ss_dest,
        worksheet=worksheet,
        source_start_col=19,
        source_end_col=32,
        dest_start_row=6,
        dest_end_row=last_row,
        nome_bloco="S:AF",
    )

    copiar_calcular_e_congelar(
        spreadsheet=ss_dest,
        worksheet=worksheet,
        source_start_col=39,
        source_end_col=48,
        dest_start_row=6,
        dest_end_row=last_row,
        nome_bloco="AM:AV",
    )

    copiar_calcular_e_congelar(
        spreadsheet=ss_dest,
        worksheet=worksheet,
        source_start_col=50,
        source_end_col=56,
        dest_start_row=6,
        dest_end_row=last_row,
        nome_bloco="AX:BD",
    )

    finalizar_execucao(worksheet)

    print("Carteira_Planejador atualizada com sucesso.")


def finalizar_execucao(worksheet: gspread.Worksheet) -> None:
    data_hora = datetime.now(TIMEZONE).strftime("%d/%m/%Y %H:%M:%S")
    escrever_celula(worksheet, "F3", data_hora, raw=False)
    formatar_data_hora(worksheet, "F3")


# =========================================================
# EXECUÇÃO DE UMA PLANILHA
# =========================================================
def executar_bloco2_para_planilha(
    client: gspread.Client,
    dest_spreadsheet_id: str,
    nome_planilha: str,
    indice: int,
    total: int,
) -> None:
    print("")
    print("=" * 80)
    print(f"Executando planilha {indice}/{total}")
    print(f"Nome destino: {nome_planilha}")
    print(f"ID destino: {dest_spreadsheet_id}")
    print("=" * 80)

    ss_dest = executar_com_retry(lambda: client.open_by_key(dest_spreadsheet_id))

    executar_bloco2_carteira_planejador(ss_dest)

    print(f"[OK] Bloco 2 concluído: {nome_planilha} | {dest_spreadsheet_id}")


# =========================================================
# MAIN
# =========================================================
def main() -> None:
    inicio = datetime.now(TIMEZONE)

    print(f"Início geral do Bloco 2 - Carteira_Planejador: {inicio.strftime('%d/%m/%Y %H:%M:%S')}")

    client = get_gspread_client()

    ss_lista = executar_com_retry(lambda: client.open_by_key(LISTA_PLANILHAS_SPREADSHEET_ID))

    planilhas = buscar_planilhas(ss_lista)

    if not planilhas:
        raise RuntimeError(
            f"Nenhum ID encontrado em {ABA_LISTA_PLANILHAS}!C3:C "
            f"na planilha {LISTA_PLANILHAS_SPREADSHEET_ID}."
        )

    print(f"Total de planilhas encontradas: {len(planilhas)}")

    sucessos = []
    erros = []

    for indice, item in enumerate(planilhas, start=1):
        nome_planilha = item["nome"]
        dest_spreadsheet_id = item["id"]

        try:
            executar_bloco2_para_planilha(
                client=client,
                dest_spreadsheet_id=dest_spreadsheet_id,
                nome_planilha=nome_planilha,
                indice=indice,
                total=len(planilhas),
            )

            sucessos.append(
                {
                    "nome": nome_planilha,
                    "id": dest_spreadsheet_id,
                }
            )

        except Exception as erro:
            print("")
            print("[ERRO] Falha ao processar uma planilha no Bloco 2.")
            print(f"Nome: {nome_planilha}")
            print(f"ID: {dest_spreadsheet_id}")
            print(f"Erro: {erro}")
            print(traceback.format_exc())

            erros.append(
                {
                    "nome": nome_planilha,
                    "id": dest_spreadsheet_id,
                    "erro": str(erro),
                }
            )

            continue

    fim = datetime.now(TIMEZONE)
    duracao = (fim - inicio).total_seconds()

    print("")
    print("=" * 80)
    print("RESUMO FINAL - BLOCO 2")
    print("=" * 80)
    print(f"Início: {inicio.strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"Fim: {fim.strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"Duração total: {duracao:.1f}s")
    print(f"Planilhas com sucesso: {len(sucessos)}")
    print(f"Planilhas com erro: {len(erros)}")

    if sucessos:
        print("")
        print("Planilhas concluídas com sucesso:")
        for item in sucessos:
            print(f"- {item['nome']} | {item['id']}")

    if erros:
        print("")
        print("Planilhas com erro:")
        for item in erros:
            print(f"- {item['nome']} | {item['id']} | Erro: {item['erro']}")

        raise RuntimeError(
            f"Bloco 2 finalizado com erro em {len(erros)} planilha(s). "
            f"Verifique o log acima."
        )

    print("")
    print("Todas as planilhas foram processadas com sucesso no Bloco 2.")


if __name__ == "__main__":
    main()
