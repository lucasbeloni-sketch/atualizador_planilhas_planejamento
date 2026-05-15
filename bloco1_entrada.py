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
from gspread.utils import rowcol_to_a1


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

ORIGEM_SPREADSHEET_ID = os.getenv(
    "ORIGEM_SPREADSHEET_ID",
    "1lUNIeWCddfmvJEjWJpQMtuR4oRuMsI3VImDY0xBp3Bs",
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


def limpar_intervalos(worksheet: gspread.Worksheet, ranges: list[str]) -> None:
    if not ranges:
        return

    executar_com_retry(lambda: worksheet.batch_clear(ranges))


def limpar_coluna_preservando_validacao(
    worksheet: gspread.Worksheet,
    col: int,
    start_row: int = 6,
) -> None:
    """
    Limpa somente os valores da coluna, sem remover lista suspensa,
    validação de dados ou formatação.
    """
    last_row = ultima_linha_preenchida_por_coluna(worksheet, col)

    if last_row < start_row:
        return

    qtd_linhas = last_row - start_row + 1
    valores_vazios = [[""] for _ in range(qtd_linhas)]

    escrever_matriz(
        worksheet=worksheet,
        start_row=start_row,
        start_col=col,
        values=valores_vazios,
        raw=True,
    )


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


def is_blank(valor) -> bool:
    return valor is None or valor == ""


def as_text(valor) -> str:
    if valor is None:
        return ""

    return str(valor)


def is_zero(valor) -> bool:
    if valor == 0:
        return True

    if isinstance(valor, str):
        texto = valor.strip().replace(",", ".")
        return texto in {"0", "0.0"}

    return False


def is_date_like(valor) -> bool:
    """
    Aproxima o comportamento do ISDATE do Google Sheets para as colunas usadas no Bloco 1.
    Considera datas vindas como texto formatado ou serial numérico plausível.
    """
    if valor is None or valor == "":
        return False

    if isinstance(valor, datetime):
        return True

    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return 20000 <= float(valor) <= 90000

    if isinstance(valor, str):
        texto = valor.strip()

        if not texto:
            return False

        formatos = [
            "%d/%m/%Y",
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%Y %H:%M",
            "%Y-%m-%d",
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%dT%H:%M:%S",
        ]

        for formato in formatos:
            try:
                datetime.strptime(texto[:19], formato)
                return True
            except ValueError:
                pass

    return False


def montar_mapa_primeira_ocorrencia(
    rows: list[list],
    key_col_idx: int,
    value_col_idx: int,
) -> dict[str, object]:
    mapa = {}

    for row in rows:
        chave = row[key_col_idx] if key_col_idx < len(row) else ""

        if is_blank(chave):
            continue

        chave_txt = as_text(chave).strip()

        if chave_txt not in mapa:
            mapa[chave_txt] = row[value_col_idx] if value_col_idx < len(row) else ""

    return mapa


def valor_coluna(row: list, col_planilha: int, col_inicial_range: int = 2):
    """
    Para ranges lidos a partir da coluna B.
    Exemplo: se range é B:AW, B=índice 0, C=índice 1, etc.
    """
    idx = col_planilha - col_inicial_range

    if idx < 0 or idx >= len(row):
        return ""

    return row[idx]


def buscar_ids_planilhas(ss_lista: gspread.Spreadsheet) -> list[str]:
    """
    Busca os IDs das planilhas na aba BD_Planilhas, coluna C, a partir da linha 3.
    Remove vazios e IDs duplicados.
    """
    aba_lista = abrir_aba(ss_lista, ABA_LISTA_PLANILHAS)
    valores_coluna_c = ler_coluna(aba_lista, 3)

    ids = []
    ids_vistos = set()

    for linha, valor in enumerate(valores_coluna_c, start=1):
        if linha < 3:
            continue

        id_planilha = as_text(valor).strip()

        if not id_planilha:
            continue

        if id_planilha in ids_vistos:
            continue

        ids.append(id_planilha)
        ids_vistos.add(id_planilha)

    return ids


# =========================================================
# BLOCO 1.1 - atualizarCart_Validador
# =========================================================
def atualizar_cart_validador(
    ss_dest: gspread.Spreadsheet,
    ss_orig: gspread.Spreadsheet,
) -> None:
    print("[1/3] Atualizando Cart_Validador...")

    aba_entrada = abrir_aba(ss_dest, "Entrada")
    aba_cart_validador_dest = abrir_aba(ss_dest, "Cart_Validador")
    aba_validacoes_orig = abrir_aba(ss_orig, "Carteira_Validações")

    escrever_celula(aba_entrada, "F2", "Etapa 1 de 3")

    limpar_intervalos(aba_cart_validador_dest, ["A1:B"])

    ultima_linha_as = ultima_linha_preenchida_por_coluna(aba_validacoes_orig, 45)

    if ultima_linha_as == 0:
        print("[1/3] Carteira_Validações sem dados na coluna AS. Cart_Validador ficou vazio.")
        return

    col_g = ler_range(aba_validacoes_orig, f"G1:G{ultima_linha_as}", ultima_linha_as, 1)
    col_as = ler_range(aba_validacoes_orig, f"AS1:AS{ultima_linha_as}", ultima_linha_as, 1)

    saida = []

    for i in range(ultima_linha_as):
        saida.append(
            [
                col_g[i][0] if i < len(col_g) else "",
                col_as[i][0] if i < len(col_as) else "",
            ]
        )

    escrever_matriz(aba_cart_validador_dest, 1, 1, saida)

    print(f"[1/3] Cart_Validador atualizado com {len(saida)} linhas.")


# =========================================================
# BLOCO 1.2 - atualizarCarteira
# =========================================================
def atualizar_carteira(
    ss_dest: gspread.Spreadsheet,
    ss_orig: gspread.Spreadsheet,
) -> None:
    print("[2/3] Atualizando Carteira...")

    aba_entrada = abrir_aba(ss_dest, "Entrada")
    aba_carteira_dest = abrir_aba(ss_dest, "Carteira")
    aba_carteira_orig = abrir_aba(ss_orig, "Carteira")

    escrever_celula(aba_entrada, "F2", "Etapa 2 de 3")

    start_row = 5
    last_src = ultima_linha_preenchida_por_coluna(aba_carteira_orig, 1)
    num_rows = max(0, last_src - (start_row - 1))

    limpar_intervalos(aba_carteira_dest, ["C1:AW", "B2:B"])

    if num_rows == 0:
        print("[2/3] Nenhum dado encontrado na origem a partir da linha 5.")
        return

    blk_a_ad = ler_range(aba_carteira_orig, f"A{start_row}:AD{last_src}", num_rows, 30)
    blk_aj_ak = ler_range(aba_carteira_orig, f"AJ{start_row}:AK{last_src}", num_rows, 2)
    blk_aw_ax = ler_range(aba_carteira_orig, f"AW{start_row}:AX{last_src}", num_rows, 2)
    blk_ca_cf = ler_range(aba_carteira_orig, f"CA{start_row}:CF{last_src}", num_rows, 6)
    blk_bq_cy = ler_range(aba_carteira_orig, f"BQ{start_row}:CY{last_src}", num_rows, 35)
    blk_cg_ch = ler_range(aba_carteira_orig, f"CG{start_row}:CH{last_src}", num_rows, 2)

    saida_c_aw = []

    for i in range(num_rows):
        row = [""] * 47

        r_a = blk_a_ad[i]
        r_aj = blk_aj_ak[i]
        r_aw = blk_aw_ax[i]
        r_ca = blk_ca_cf[i]
        r_bq = blk_bq_cy[i]
        r_cg = blk_cg_ch[i]

        # C = A
        row[0] = r_a[0]

        # D = U
        row[1] = r_a[20]

        # E:G = C:E
        row[2] = r_a[2]
        row[3] = r_a[3]
        row[4] = r_a[4]

        # H:N = N:T
        for k in range(7):
            row[5 + k] = r_a[13 + k]

        # O:W = V:AD
        for k in range(9):
            row[12 + k] = r_a[21 + k]

        # X:Y = AJ:AK
        row[21] = r_aj[0]
        row[22] = r_aj[1]

        # Z = CA, AB = CC, AC = CF, AF = CE
        row[23] = r_ca[0]
        row[25] = r_ca[2]
        row[26] = r_ca[5]
        row[29] = r_ca[4]

        # AA = BQ
        row[24] = r_bq[0]

        # AD:AE = AW:AX
        row[27] = r_aw[0]
        row[28] = r_aw[1]

        # AG:AM = CQ:CW
        for k in range(7):
            row[30 + k] = r_bq[26 + k]

        # AN:AR = BR:BV
        row[37] = r_bq[1]
        row[38] = r_bq[2]
        row[39] = r_bq[3]
        row[40] = r_bq[4]
        row[41] = r_bq[5]

        # AS:AT = CX:CY
        row[42] = r_bq[33]
        row[43] = r_bq[34]

        # AU:AV = CG:CH
        row[44] = r_cg[0]
        row[45] = r_cg[1]

        saida_c_aw.append(row)

    escrever_matriz(aba_carteira_dest, 1, 3, saida_c_aw)

    # Calcula coluna B via script
    aba_config = abrir_aba(ss_dest, "BD_Config")
    cfg_vals = ler_range(aba_config, "B4:B9", 6, 1)
    cfg_set = {as_text(row[0]).strip() for row in cfg_vals if not is_blank(row[0])}

    aba_plan = abrir_aba(ss_dest, "Carteira_Planejador")
    last_plan = ultima_linha_preenchida_por_coluna(aba_plan, 13)

    q_vals = (
        ler_range(aba_plan, f"M1:M{last_plan}", last_plan, 1)
        if last_plan > 0
        else []
    )

    freq = {}

    for row in q_vals:
        valor = row[0] if row else ""

        if is_blank(valor):
            continue

        chave = as_text(valor).strip()
        freq[chave] = freq.get(chave, 0) + 1

    if num_rows >= 2:
        saida_b = []

        for i in range(1, num_rows):
            col_c = saida_c_aw[i][0]
            col_d = saida_c_aw[i][1]
            col_o = saida_c_aw[i][12]

            if is_blank(col_c):
                valor_b = ""
            elif as_text(col_d).strip() != "OBRA RETIRADA" and as_text(col_o).strip() in cfg_set:
                valor_b = freq.get(as_text(col_c).strip(), 0)
            else:
                valor_b = "-"

            saida_b.append([valor_b])

        escrever_matriz(aba_carteira_dest, 2, 2, saida_b)

    print(f"[2/3] Carteira atualizada com {num_rows} linhas em C:AW.")


# =========================================================
# BLOCO 1.3 - atualizarEntradaNova
# =========================================================
def atualizar_entrada_nova(ss_dest: gspread.Spreadsheet) -> None:
    print("[3/3] Atualizando Entrada...")

    aba_entrada = abrir_aba(ss_dest, "Entrada")
    aba_carteira = abrir_aba(ss_dest, "Carteira")
    aba_cart_validador = abrir_aba(ss_dest, "Cart_Validador")

    escrever_celula(aba_entrada, "F2", "Etapa 3 de 3")

    # Limpa a coluna B somente nos valores, preservando lista suspensa/validação
    limpar_coluna_preservando_validacao(aba_entrada, col=2, start_row=6)

    # Limpa normalmente as demais colunas
    limpar_intervalos(aba_entrada, ["C6:AD"])

    last_carteira = ultima_linha_preenchida_por_coluna(aba_carteira, 3)

    if last_carteira == 0:
        finalizar_execucao(aba_entrada)
        print("[3/3] Carteira sem dados. Entrada ficou vazia.")
        return

    # Lê B:AW porque o filtro usa B e as demais colunas estão até AW.
    carteira_rows = ler_range(aba_carteira, f"B1:AW{last_carteira}", last_carteira, 48)

    last_validador = ultima_linha_preenchida_por_coluna(aba_cart_validador, 1)

    cart_validador_rows = (
        ler_range(aba_cart_validador, f"A1:B{last_validador}", last_validador, 2)
        if last_validador > 0
        else []
    )

    mapa_observacao = montar_mapa_primeira_ocorrencia(cart_validador_rows, 0, 1)

    # Mapas para simular XLOOKUP em Carteira!C:C
    mapa_c_para_h = montar_mapa_primeira_ocorrencia(carteira_rows, 1, 6)    # C -> H
    mapa_c_para_aa = montar_mapa_primeira_ocorrencia(carteira_rows, 1, 25)  # C -> AA
    mapa_c_para_af = montar_mapa_primeira_ocorrencia(carteira_rows, 1, 30)  # C -> AF

    saida_c_ad = []

    for row in carteira_rows:
        valor_b = valor_coluna(row, 2)

        if is_blank(valor_b) or not is_zero(valor_b):
            continue

        projeto = valor_coluna(row, 3)
        projeto_key = as_text(projeto).strip()

        flag_c = 2 if as_text(mapa_c_para_h.get(projeto_key, "")).strip() == "APTA" else 0
        flag_d = 2 if is_date_like(mapa_c_para_aa.get(projeto_key, "")) else 0
        flag_e = 2 if is_date_like(mapa_c_para_af.get(projeto_key, "")) else 0

        # Saída C:AD = 28 colunas
        out = [""] * 28

        # C:E
        out[0] = flag_c
        out[1] = flag_d
        out[2] = flag_e

        # F = Carteira!G
        out[3] = valor_coluna(row, 7)

        # G = Carteira!E
        out[4] = valor_coluna(row, 5)

        # H = Carteira!D
        out[5] = valor_coluna(row, 4)

        # I:K ficam vazias

        # L = Carteira!C
        out[9] = projeto

        # M = Carteira!S
        out[10] = valor_coluna(row, 19)

        # N:O = Carteira!Q:R
        out[11] = valor_coluna(row, 17)
        out[12] = valor_coluna(row, 18)

        # P fica vazia

        # Q = Carteira!U
        out[14] = valor_coluna(row, 21)

        # R:S = Carteira!V:W
        out[15] = valor_coluna(row, 22)
        out[16] = valor_coluna(row, 23)

        # T:U = Carteira!X:Y
        out[17] = valor_coluna(row, 24)
        out[18] = valor_coluna(row, 25)

        # V:W ficam vazias

        # X = Carteira!T
        out[21] = valor_coluna(row, 20)

        # Y = Observações Carteira via Cart_Validador
        out[22] = mapa_observacao.get(projeto_key, "")

        # Z:AA = Carteira!Z:AA
        out[23] = valor_coluna(row, 26)
        out[24] = valor_coluna(row, 27)

        # AB:AC = Carteira!AD:AE
        out[25] = valor_coluna(row, 30)
        out[26] = valor_coluna(row, 31)

        # AD = Carteira!O
        out[27] = valor_coluna(row, 15)

        saida_c_ad.append(out)

    if saida_c_ad:
        escrever_matriz(aba_entrada, 6, 3, saida_c_ad)

    finalizar_execucao(aba_entrada)

    print(f"[3/3] Entrada atualizada com {len(saida_c_ad)} linhas.")


def finalizar_execucao(aba_entrada: gspread.Worksheet) -> None:
    data_hora = datetime.now(TIMEZONE).strftime("%d/%m/%Y %H:%M:%S")
    escrever_celula(aba_entrada, "F2", data_hora)


# =========================================================
# EXECUÇÃO DE UMA PLANILHA
# =========================================================
def executar_bloco1_para_planilha(
    client: gspread.Client,
    ss_orig: gspread.Spreadsheet,
    dest_spreadsheet_id: str,
    indice: int,
    total: int,
) -> None:
    print("")
    print("=" * 80)
    print(f"Executando planilha {indice}/{total}")
    print(f"ID destino: {dest_spreadsheet_id}")
    print("=" * 80)

    ss_dest = executar_com_retry(lambda: client.open_by_key(dest_spreadsheet_id))

    atualizar_cart_validador(ss_dest, ss_orig)
    atualizar_carteira(ss_dest, ss_orig)
    atualizar_entrada_nova(ss_dest)

    print(f"[OK] Planilha concluída: {dest_spreadsheet_id}")


# =========================================================
# MAIN
# =========================================================
def main() -> None:
    inicio = datetime.now(TIMEZONE)

    print(f"Início geral do Bloco 1 - Entrada: {inicio.strftime('%d/%m/%Y %H:%M:%S')}")

    client = get_gspread_client()

    ss_lista = executar_com_retry(lambda: client.open_by_key(LISTA_PLANILHAS_SPREADSHEET_ID))
    ss_orig = executar_com_retry(lambda: client.open_by_key(ORIGEM_SPREADSHEET_ID))

    ids_planilhas = buscar_ids_planilhas(ss_lista)

    if not ids_planilhas:
        raise RuntimeError(
            f"Nenhum ID encontrado em {ABA_LISTA_PLANILHAS}!C3:C "
            f"na planilha {LISTA_PLANILHAS_SPREADSHEET_ID}."
        )

    print(f"Total de planilhas encontradas: {len(ids_planilhas)}")

    sucessos = []
    erros = []

    for indice, dest_spreadsheet_id in enumerate(ids_planilhas, start=1):
        try:
            executar_bloco1_para_planilha(
                client=client,
                ss_orig=ss_orig,
                dest_spreadsheet_id=dest_spreadsheet_id,
                indice=indice,
                total=len(ids_planilhas),
            )
            sucessos.append(dest_spreadsheet_id)

        except Exception as erro:
            print("")
            print("[ERRO] Falha ao processar uma planilha.")
            print(f"ID: {dest_spreadsheet_id}")
            print(f"Erro: {erro}")
            print(traceback.format_exc())

            erros.append(
                {
                    "id": dest_spreadsheet_id,
                    "erro": str(erro),
                }
            )

            continue

    fim = datetime.now(TIMEZONE)
    duracao = (fim - inicio).total_seconds()

    print("")
    print("=" * 80)
    print("RESUMO FINAL")
    print("=" * 80)
    print(f"Início: {inicio.strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"Fim: {fim.strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"Duração total: {duracao:.1f}s")
    print(f"Planilhas com sucesso: {len(sucessos)}")
    print(f"Planilhas com erro: {len(erros)}")

    if sucessos:
        print("")
        print("IDs concluídos com sucesso:")
        for id_ok in sucessos:
            print(f"- {id_ok}")

    if erros:
        print("")
        print("IDs com erro:")
        for item in erros:
            print(f"- {item['id']} | Erro: {item['erro']}")

        raise RuntimeError(
            f"Processamento finalizado com erro em {len(erros)} planilha(s). "
            f"Verifique o log acima."
        )

    print("")
    print("Todas as planilhas foram processadas com sucesso.")


if __name__ == "__main__":
    main()
