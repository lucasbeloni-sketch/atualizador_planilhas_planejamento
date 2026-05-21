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

# Tempo para aguardar o Google Sheets calcular as fórmulas antes de congelar.
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
            date_time_render_option="SERIAL_NUMBER",
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


def ultima_linha_preenchida_da_planilha(
    worksheet: gspread.Worksheet,
    range_a1: str = "A:BT",
) -> int:
    """
    Equivalente aproximado ao getLastRow() do Apps Script,
    considerando o intervalo principal usado no Plan_Principal.
    """
    valores = executar_com_retry(
        lambda: worksheet.get(
            range_a1,
            value_render_option="UNFORMATTED_VALUE",
            date_time_render_option="SERIAL_NUMBER",
        )
    )

    for idx in range(len(valores) - 1, -1, -1):
        row = valores[idx]
        if any(not is_blank(cell) for cell in row):
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


def aplicar_formatacoes_fixas_plan_principal(worksheet: gspread.Worksheet) -> None:
    """
    Reaplica formatações fixas da aba Plan_Principal da linha 6 até o fim da aba,
    mesmo em linhas vazias.

    Moeda:
      AL, AM, AO, AQ, BQ

    Porcentagem:
      AN, AP, AR

    Duração:
      BL, BM, BN, BO, BP
    """
    ultima_linha = max(worksheet.row_count, 1000)

    formatos = [
        {
            "ranges": [
                f"AL6:AL{ultima_linha}",
                f"AM6:AM{ultima_linha}",
                f"AO6:AO{ultima_linha}",
                f"AQ6:AQ{ultima_linha}",
                f"BQ6:BQ{ultima_linha}",
            ],
            "format": {
                "numberFormat": {
                    "type": "CURRENCY",
                    "pattern": 'R$ #,##0.00',
                }
            },
        },
        {
            "ranges": [
                f"AN6:AN{ultima_linha}",
                f"AP6:AP{ultima_linha}",
                f"AR6:AR{ultima_linha}",
            ],
            "format": {
                "numberFormat": {
                    "type": "PERCENT",
                    "pattern": "0%",
                }
            },
        },
        {
            "ranges": [
                f"BL6:BL{ultima_linha}",
                f"BM6:BM{ultima_linha}",
                f"BN6:BN{ultima_linha}",
                f"BO6:BO{ultima_linha}",
                f"BP6:BP{ultima_linha}",
            ],
            "format": {
                "numberFormat": {
                    "type": "TIME",
                    "pattern": "[h]:mm:ss",
                }
            },
        },
    ]

    for item in formatos:
        for range_a1 in item["ranges"]:
            try:
                executar_com_retry(
                    lambda range_a1=range_a1, fmt=item["format"]: worksheet.format(
                        range_a1,
                        fmt,
                    )
                )
            except Exception as erro:
                print(f"[AVISO] Não foi possível aplicar formatação em {range_a1}: {erro}")

    print("Formatações fixas reaplicadas em Plan_Principal.")


def remover_filtro_basico(spreadsheet: gspread.Spreadsheet, worksheet: gspread.Worksheet) -> None:
    """
    Remove filtro ativo da aba, se houver.
    Caso não exista filtro, apenas segue.
    """
    try:
        request = {
            "requests": [
                {
                    "clearBasicFilter": {
                        "sheetId": worksheet.id,
                    }
                }
            ]
        }

        executar_com_retry(lambda: spreadsheet.batch_update(request))
        print("Filtro ativo removido, se existia.")
    except Exception as erro:
        print(f"[AVISO] Não foi possível remover filtro ativo ou não havia filtro: {erro}")


def escape_formula_text(texto: str) -> str:
    """
    Escapa aspas duplas para usar texto dentro de fórmula do Google Sheets.
    """
    return as_text(texto).replace('"', '""')


def buscar_planilhas(ss_lista: gspread.Spreadsheet) -> list[dict[str, str]]:
    """
    Busca os dados das planilhas na aba BD_Planilhas.
    Nome: coluna B, a partir da linha 3.
    ID: coluna C, a partir da linha 3.
    Valor BE: coluna D, a partir da linha 3.
    Remove IDs vazios e duplicados.
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


# =========================================================
# FÓRMULAS - PLAN_PRINCIPAL
# =========================================================
def formulas_j_l(row: int) -> list[str]:
    return [
        f'=XLOOKUP(H{row};Carteira!$C:$C;Carteira!$S:$S;"")',
        f'=XLOOKUP(H{row};Carteira!$C:$C;Carteira!$Q:$Q;"")',
        f'=XLOOKUP(H{row};Carteira!$C:$C;Carteira!$R:$R;"")',
    ]


def formula_al(row: int) -> str:
    colunas = [
        "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
        "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ",
    ]

    partes = []

    for idx, coluna in enumerate(colunas, start=5):
        partes.append(f"IFERROR({coluna}{row}*BD_Config!$W${idx};0)")

    soma = "+".join(partes)

    return (
        f'=IF(B{row}="";"";'
        f'IF(BQ{row}<>"";BQ{row};({soma})))'
    )


def formula_am(row: int) -> str:
    row_anterior = row - 1

    return (
        f'=IF(B{row}="";"";'
        f'IF(COUNTIFS($B$1:B{row_anterior};B{row};$G$1:G{row_anterior};G{row})=0;'
        f'XLOOKUP(G{row}&B{row};BD_Metas!$A:$A;BD_Metas!$D:$D;0);0))'
    )


def formula_ao(row: int) -> str:
    return (
        f'=IF(B{row}="";"";'
        f'IF(AL{row}=0;0;'
        f'SUMIFS(BD_Serv_GPM!$G:$G;BD_Serv_GPM!$H:$H;H{row};BD_Serv_GPM!$D:$D;G{row};BD_Serv_GPM!$F:$F;B{row})'
        f'+SUMIFS(BD_Serv_GPM!$G:$G;BD_Serv_GPM!$J:$J;H{row};BD_Serv_GPM!$D:$D;G{row};BD_Serv_GPM!$F:$F;B{row})))'
    )


def formula_aq(row: int) -> str:
    return (
        f'=IF(B{row}="";"";'
        f'SUMIFS(BD_Serv_GPM!$G:$G;BD_Serv_GPM!$D:$D;G{row};BD_Serv_GPM!$F:$F;B{row}))'
    )


def formula_as(row: int) -> str:
    return (
        f'=IF(B{row}="";"";'
        f'IF(XLOOKUP(H{row}&B{row}*1&G{row};BD_Serv_GPM!$I:$I;BD_Serv_GPM!$E:$E;"-")="-";'
        f'XLOOKUP(H{row}&B{row}*1&G{row};BD_Serv_GPM!$K:$K;BD_Serv_GPM!$E:$E;"-");'
        f'XLOOKUP(H{row}&B{row}*1&G{row};BD_Serv_GPM!$I:$I;BD_Serv_GPM!$E:$E;"-")))'
    )


def formula_be(row: int, valor_be: str) -> str:
    texto = escape_formula_text(valor_be)

    return f'=IF(B{row}<>"";"{texto}";"")'


def formulas_derivadas_an_ap_ar(row: int) -> list[str]:
    return [
        f'=IF(B{row}="";"";IFERROR(AL{row}/AM{row};0))',
        f'=IF(B{row}="";"";IFERROR(AO{row}/AL{row};0))',
        f'=IF(B{row}="";"";IFERROR(AQ{row}/AM{row};0))',
    ]


def formulas_br_bt(row: int) -> list[str]:
    return [
        f'=XLOOKUP($H{row};Carteira!$C:$C;Carteira!$AU:$AU;"")',
        f'=XLOOKUP($H{row};Carteira!$C:$C;Carteira!$AV:$AV;"")',
        f'=XLOOKUP($H{row};Carteira!$C:$C;Carteira!$AA:$AA;"")',
    ]


def formula_ak(row: int) -> str:
    return (
        f'=IFERROR('
        f'IF(AND(AQ{row}<>"";AL{row}<>"");'
        f'IF(AND(NOT(OR(N(AL{row})=0;N(AQ{row})/N(AL{row})<1,1));'
        f'SUMIF(BD_Precificacao!$A:$A;H{row};BD_Precificacao!$F:$F)>0);'
        f'"PROD/ORC ACIMA";'
        f'IF(NOT(OR(N(AL{row})=0;N(AQ{row})/N(AL{row})<1,1));'
        f'"PROD. ACIMA";'
        f'IF(SUMIF(BD_Precificacao!$A:$A;H{row};BD_Precificacao!$F:$F)>0;'
        f'"ORC. ACIMA";"OPCIONAL")));'
        f'"NÃO");"NÃO")'
    )


# =========================================================
# BLOCO 3 - PLAN_PRINCIPAL
# =========================================================
def executar_bloco3_plan_principal(
    ss_dest: gspread.Spreadsheet,
    valor_be: str,
) -> None:
    aba = abrir_aba(ss_dest, "Plan_Principal")

    print("Atualizando aba Plan_Principal...")

    remover_filtro_basico(ss_dest, aba)

    escrever_celula(aba, "F3", "Em Atualização")

    last_row_sheet = ultima_linha_preenchida_da_planilha(aba, "A:BT")

    print(f"Última linha preenchida da Plan_Principal: {last_row_sheet}")

    if last_row_sheet < 6:
        aplicar_formatacoes_fixas_plan_principal(aba)
        finalizar_execucao(aba)
        print("Nenhuma linha para atualizar a partir da linha 6.")
        return

    # =====================================================
    # Limpezas iniciais
    # =====================================================
    limpar_intervalos(
        aba,
        [
            "J6:L",
            "AL6:AS",
            "AK6:AK",
            "BE6:BE",
            "BR6:BT",
        ],
    )

    linhas = list(range(6, last_row_sheet + 1))

    # =====================================================
    # Fórmulas iniciais
    # J:L
    # AL, AM, AO, AQ, AS
    # BE
    # =====================================================
    print("Aplicando fórmulas iniciais em J:L...")
    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=10,
        formulas=[formulas_j_l(row) for row in linhas],
    )

    print("Aplicando fórmulas em AL...")
    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=38,
        formulas=[[formula_al(row)] for row in linhas],
    )

    print("Aplicando fórmulas em AM...")
    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=39,
        formulas=[[formula_am(row)] for row in linhas],
    )

    print("Aplicando fórmulas em AO...")
    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=41,
        formulas=[[formula_ao(row)] for row in linhas],
    )

    print("Aplicando fórmulas em AQ...")
    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=43,
        formulas=[[formula_aq(row)] for row in linhas],
    )

    print("Aplicando fórmulas em AS...")
    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=45,
        formulas=[[formula_as(row)] for row in linhas],
    )

    print(f'Aplicando fórmulas em BE com valor da BD_Planilhas coluna D: "{valor_be}"')
    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=57,
        formulas=[[formula_be(row, valor_be)] for row in linhas],
    )

    print(f"Aguardando cálculo inicial por {CALC_WAIT_SECONDS}s...")
    time.sleep(CALC_WAIT_SECONDS)

    # =====================================================
    # Fórmulas derivadas
    # AN, AP, AR
    # BR:BT
    # =====================================================
    print("Aplicando fórmulas derivadas em AN, AP e AR...")

    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=40,
        formulas=[[formulas_derivadas_an_ap_ar(row)[0]] for row in linhas],
    )

    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=42,
        formulas=[[formulas_derivadas_an_ap_ar(row)[1]] for row in linhas],
    )

    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=44,
        formulas=[[formulas_derivadas_an_ap_ar(row)[2]] for row in linhas],
    )

    print("Aplicando fórmulas em BR:BT...")
    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=70,
        formulas=[formulas_br_bt(row) for row in linhas],
    )

    print(f"Aguardando cálculo das derivadas por {CALC_WAIT_SECONDS}s...")
    time.sleep(CALC_WAIT_SECONDS)

    # =====================================================
    # Congelar valores
    # AL:AS, J:L, BR:BT
    # BE fica como fórmula, igual ao Apps Script original.
    # =====================================================
    print("Congelando valores em AL:AS...")
    congelar_intervalo(aba, f"AL6:AS{last_row_sheet}")

    print("Congelando valores em J:L...")
    congelar_intervalo(aba, f"J6:L{last_row_sheet}")

    print("Congelando valores em BR:BT...")
    congelar_intervalo(aba, f"BR6:BT{last_row_sheet}")

    # =====================================================
    # AK baseado na última linha preenchida em H
    # =====================================================
    aplicar_ak_por_coluna_h(aba, last_row_sheet)

    # Reaplica as formatações fixas mesmo nas linhas vazias/não atualizadas
    aplicar_formatacoes_fixas_plan_principal(aba)

    finalizar_execucao(aba)

    print("Plan_Principal atualizada com sucesso.")


def aplicar_ak_por_coluna_h(
    aba: gspread.Worksheet,
    last_row_sheet: int,
) -> None:
    num_rows = max(last_row_sheet - 5, 0)

    if num_rows <= 0:
        limpar_intervalos(aba, ["AK6:AK"])
        return

    valores_h = ler_range(
        worksheet=aba,
        range_a1=f"H6:H{last_row_sheet}",
        n_rows=num_rows,
        n_cols=1,
    )

    last_row_h = 5

    for idx in range(len(valores_h) - 1, -1, -1):
        valor = valores_h[idx][0] if valores_h[idx] else ""

        if not is_blank(valor):
            last_row_h = 6 + idx
            break

    if last_row_h < 6:
        limpar_intervalos(aba, ["AK6:AK"])
        print("Coluna H sem dados a partir da linha 6. AK limpa.")
        return

    print(f"Última linha preenchida pela coluna H: {last_row_h}")

    limpar_intervalos(aba, [f"AK6:AK{last_row_h}"])

    linhas_ak = list(range(6, last_row_h + 1))

    print("Aplicando nova lógica em AK...")
    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=37,
        formulas=[[formula_ak(row)] for row in linhas_ak],
    )

    print(f"Aguardando cálculo de AK por {CALC_WAIT_SECONDS}s...")
    time.sleep(CALC_WAIT_SECONDS)

    print("Congelando valores em AK...")
    congelar_intervalo(aba, f"AK6:AK{last_row_h}")

    if last_row_h < last_row_sheet:
        limpar_intervalos(aba, [f"AK{last_row_h + 1}:AK"])


def finalizar_execucao(worksheet: gspread.Worksheet) -> None:
    data_hora = datetime.now(TIMEZONE).strftime("%d/%m/%Y %H:%M:%S")
    escrever_celula(worksheet, "F3", data_hora, raw=False)
    formatar_data_hora(worksheet, "F3")


# =========================================================
# EXECUÇÃO DE UMA PLANILHA
# =========================================================
def executar_bloco3_para_planilha(
    client: gspread.Client,
    dest_spreadsheet_id: str,
    nome_planilha: str,
    valor_be: str,
    indice: int,
    total: int,
) -> None:
    print("")
    print("=" * 80)
    print(f"Executando planilha {indice}/{total}")
    print(f"Nome destino: {nome_planilha}")
    print(f"ID destino: {dest_spreadsheet_id}")
    print(f"Valor para Plan_Principal!BE: {valor_be}")
    print("=" * 80)

    ss_dest = executar_com_retry(lambda: client.open_by_key(dest_spreadsheet_id))

    executar_bloco3_plan_principal(
        ss_dest=ss_dest,
        valor_be=valor_be,
    )

    print(f"[OK] Bloco 3 concluído: {nome_planilha} | {dest_spreadsheet_id}")


# =========================================================
# MAIN
# =========================================================
def main() -> None:
    inicio = datetime.now(TIMEZONE)

    print(f"Início geral do Bloco 3 - Plan_Principal: {inicio.strftime('%d/%m/%Y %H:%M:%S')}")

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
        valor_be = item["valor_be"]

        try:
            executar_bloco3_para_planilha(
                client=client,
                dest_spreadsheet_id=dest_spreadsheet_id,
                nome_planilha=nome_planilha,
                valor_be=valor_be,
                indice=indice,
                total=len(planilhas),
            )

            sucessos.append(
                {
                    "nome": nome_planilha,
                    "id": dest_spreadsheet_id,
                    "valor_be": valor_be,
                }
            )

        except Exception as erro:
            print("")
            print("[ERRO] Falha ao processar uma planilha no Bloco 3.")
            print(f"Nome: {nome_planilha}")
            print(f"ID: {dest_spreadsheet_id}")
            print(f"Valor BE: {valor_be}")
            print(f"Erro: {erro}")
            print(traceback.format_exc())

            erros.append(
                {
                    "nome": nome_planilha,
                    "id": dest_spreadsheet_id,
                    "valor_be": valor_be,
                    "erro": str(erro),
                }
            )

            continue

    fim = datetime.now(TIMEZONE)
    duracao = (fim - inicio).total_seconds()

    print("")
    print("=" * 80)
    print("RESUMO FINAL - BLOCO 3")
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
            print(f"- {item['nome']} | {item['id']} | BE: {item['valor_be']}")

    if erros:
        print("")
        print("Planilhas com erro:")
        for item in erros:
            print(
                f"- {item['nome']} | {item['id']} | "
                f"BE: {item['valor_be']} | Erro: {item['erro']}"
            )

        raise RuntimeError(
            f"Bloco 3 finalizado com erro em {len(erros)} planilha(s). "
            f"Verifique o log acima."
        )

    print("")
    print("Todas as planilhas foram processadas com sucesso no Bloco 3.")


if __name__ == "__main__":
    main()
