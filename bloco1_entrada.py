import base64
import json
import os
import re
import time
from datetime import datetime
from typing import Any, Dict, List, Optional
from zoneinfo import ZoneInfo

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# =========================
# CONFIGURAÇÕES PRINCIPAIS
# =========================

PLANILHA_DESTINO_ID = os.getenv(
    "PLANILHA_DESTINO_ID",
    "1QUolOl1Sk1ZdLMLjG9RSNxUsSB8ucydvqXRifxFUOH1jR5Mlc7c33Yu0",
)

PLANILHA_ORIGEM_ID = os.getenv(
    "PLANILHA_ORIGEM_ID",
    "1lUNIeWCddfmvJEjWJpQMtuR4oRuMsI3VImDY0xBp3Bs",
)

TIMEZONE = os.getenv("TIMEZONE", "America/Sao_Paulo")
CHUNK_SIZE = int(os.getenv("CHUNK_SIZE", "5000"))

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


# =========================
# FUNÇÕES BASE
# =========================

def obter_credenciais() -> Credentials:
    """
    Prioridade:
    1) GOOGLE_CREDENTIALS_B64: secret em Base64 usado no GitHub Actions.
    2) GOOGLE_APPLICATION_CREDENTIALS: caminho local do JSON.
    3) service_account.json no diretório do projeto.
    """
    credentials_b64 = os.getenv("GOOGLE_CREDENTIALS_B64", "").strip()

    if credentials_b64:
        info = json.loads(base64.b64decode(credentials_b64).decode("utf-8"))
        return Credentials.from_service_account_info(info, scopes=SCOPES)

    credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
    if os.path.exists(credentials_path):
        return Credentials.from_service_account_file(credentials_path, scopes=SCOPES)

    raise RuntimeError(
        "Credenciais não encontradas. Configure GOOGLE_CREDENTIALS_B64 no GitHub "
        "ou use GOOGLE_APPLICATION_CREDENTIALS/local service_account.json."
    )


def col_para_num(coluna: str) -> int:
    total = 0
    for char in coluna.upper():
        if not ("A" <= char <= "Z"):
            continue
        total = total * 26 + (ord(char) - ord("A") + 1)
    return total


def num_para_col(numero: int) -> str:
    letras = ""
    while numero:
        numero, resto = divmod(numero - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def aba(nome: str) -> str:
    return "'" + nome.replace("'", "''") + "'"


def montar_range(nome_aba: str, linha: int, coluna: int, qtd_linhas: int, qtd_colunas: int) -> str:
    col_ini = num_para_col(coluna)
    col_fim = num_para_col(coluna + qtd_colunas - 1)
    lin_fim = linha + qtd_linhas - 1
    return f"{aba(nome_aba)}!{col_ini}{linha}:{col_fim}{lin_fim}"


def matriz_padrao(matriz: List[List[Any]], linhas: int, colunas: int) -> List[List[Any]]:
    saida = []

    for i in range(linhas):
        linha = matriz[i] if i < len(matriz) else []
        linha = list(linha[:colunas])

        if len(linha) < colunas:
            linha.extend([""] * (colunas - len(linha)))

        saida.append(linha)

    return saida


def valor_linha(linha: List[Any], coluna_inicial: str, coluna_desejada: str) -> Any:
    idx = col_para_num(coluna_desejada) - col_para_num(coluna_inicial)

    if 0 <= idx < len(linha):
        return linha[idx]

    return ""


def definir_valor_entrada(linha: List[Any], coluna_destino: str, valor: Any) -> None:
    idx = col_para_num(coluna_destino) - col_para_num("C")

    if 0 <= idx < len(linha):
        linha[idx] = valor


def vazio(valor: Any) -> bool:
    return valor is None or str(valor).strip() == ""


def chave(valor: Any) -> str:
    if valor is None:
        return ""

    if isinstance(valor, (int, float)):
        if float(valor).is_integer():
            return str(int(valor))
        return str(valor).strip()

    return str(valor).strip()


def eh_zero(valor: Any) -> bool:
    if vazio(valor):
        return False

    if isinstance(valor, (int, float)):
        return float(valor) == 0

    texto = str(valor).strip().replace(".", "").replace(",", ".")

    try:
        return float(texto) == 0
    except ValueError:
        return False


def eh_data(valor: Any) -> bool:
    if vazio(valor):
        return False

    if isinstance(valor, (int, float)):
        return float(valor) > 0

    texto = str(valor).strip()

    if texto in {"-", "0", "0,0", "0.0"}:
        return False

    padroes = [
        r"^\d{1,2}/\d{1,2}/\d{2,4}$",
        r"^\d{1,2}/\d{1,2}/\d{2,4}\s+\d{1,2}:\d{2}(:\d{2})?$",
        r"^\d{4}-\d{2}-\d{2}$",
        r"^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(:\d{2})?",
    ]

    return any(re.match(padrao, texto) for padrao in padroes)


def serial_google_sheets(dt: datetime) -> float:
    epoch = datetime(1899, 12, 30, tzinfo=dt.tzinfo)
    return (dt - epoch).total_seconds() / 86400


class SheetsClient:
    def __init__(self) -> None:
        self.service = build(
            "sheets",
            "v4",
            credentials=obter_credenciais(),
            cache_discovery=False,
        )

    def executar(self, fabrica_requisicao, descricao: str = "") -> Dict[str, Any]:
        tentativas = 6

        for tentativa in range(tentativas):
            try:
                return fabrica_requisicao().execute()

            except HttpError as erro:
                status = getattr(erro.resp, "status", None)

                if status not in {429, 500, 502, 503, 504} or tentativa == tentativas - 1:
                    raise

                espera = min(60, (2 ** tentativa) + 1)
                print(
                    f"Aviso: falha temporária em {descricao or 'requisição'} "
                    f"({status}). Nova tentativa em {espera}s."
                )
                time.sleep(espera)

        return {}

    def get_values(
        self,
        spreadsheet_id: str,
        range_a1: str,
        value_render_option: str = "UNFORMATTED_VALUE",
    ) -> List[List[Any]]:
        resposta = self.executar(
            lambda: self.service.spreadsheets()
            .values()
            .get(
                spreadsheetId=spreadsheet_id,
                range=range_a1,
                valueRenderOption=value_render_option,
                dateTimeRenderOption="SERIAL_NUMBER",
            ),
            f"leitura {range_a1}",
        )

        return resposta.get("values", [])

    def batch_get(
        self,
        spreadsheet_id: str,
        ranges: List[str],
        value_render_option: str = "UNFORMATTED_VALUE",
    ) -> List[List[List[Any]]]:
        resposta = self.executar(
            lambda: self.service.spreadsheets()
            .values()
            .batchGet(
                spreadsheetId=spreadsheet_id,
                ranges=ranges,
                valueRenderOption=value_render_option,
                dateTimeRenderOption="SERIAL_NUMBER",
            ),
            "batchGet",
        )

        value_ranges = resposta.get("valueRanges", [])
        return [vr.get("values", []) for vr in value_ranges]

    def update_values(
        self,
        spreadsheet_id: str,
        range_a1: str,
        values: List[List[Any]],
        value_input_option: str = "RAW",
    ) -> None:
        self.executar(
            lambda: self.service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_a1,
                valueInputOption=value_input_option,
                body={"values": values},
            ),
            f"escrita {range_a1}",
        )

    def clear_values(self, spreadsheet_id: str, ranges: List[str]) -> None:
        self.executar(
            lambda: self.service.spreadsheets()
            .values()
            .batchClear(
                spreadsheetId=spreadsheet_id,
                body={"ranges": ranges},
            ),
            "limpeza de ranges",
        )

    def batch_update(self, spreadsheet_id: str, requests: List[Dict[str, Any]]) -> None:
        if not requests:
            return

        self.executar(
            lambda: self.service.spreadsheets()
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests},
            ),
            "batchUpdate",
        )

    def get_sheet_id(self, spreadsheet_id: str, sheet_name: str) -> Optional[int]:
        resposta = self.executar(
            lambda: self.service.spreadsheets()
            .get(
                spreadsheetId=spreadsheet_id,
                fields="sheets(properties(sheetId,title))",
            ),
            "metadata da planilha",
        )

        for sheet in resposta.get("sheets", []):
            props = sheet.get("properties", {})

            if props.get("title") == sheet_name:
                return props.get("sheetId")

        return None


def atualizar_status(client: SheetsClient, texto: str) -> None:
    client.update_values(
        PLANILHA_DESTINO_ID,
        f"{aba('Entrada')}!F2",
        [[texto]],
        value_input_option="USER_ENTERED",
    )


def finalizar_status_com_data(client: SheetsClient) -> None:
    dt = datetime.now(ZoneInfo(TIMEZONE))
    serial = serial_google_sheets(dt)

    client.update_values(
        PLANILHA_DESTINO_ID,
        f"{aba('Entrada')}!F2",
        [[serial]],
        value_input_option="USER_ENTERED",
    )

    sheet_id = client.get_sheet_id(PLANILHA_DESTINO_ID, "Entrada")

    if sheet_id is None:
        return

    client.batch_update(
        PLANILHA_DESTINO_ID,
        [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": 2,
                        "startColumnIndex": 5,
                        "endColumnIndex": 6,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "DATE_TIME",
                                "pattern": "dd/MM/yyyy HH:mm:ss",
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat",
                }
            }
        ],
    )


def update_chunked(
    client: SheetsClient,
    spreadsheet_id: str,
    sheet_name: str,
    start_row: int,
    start_col: int,
    values: List[List[Any]],
    value_input_option: str = "RAW",
) -> None:
    if not values:
        return

    qtd_colunas = max(len(row) for row in values)
    valores_normalizados = []

    for row in values:
        nova = list(row[:qtd_colunas])

        if len(nova) < qtd_colunas:
            nova.extend([""] * (qtd_colunas - len(nova)))

        valores_normalizados.append(nova)

    for inicio in range(0, len(valores_normalizados), CHUNK_SIZE):
        bloco = valores_normalizados[inicio: inicio + CHUNK_SIZE]

        range_a1 = montar_range(
            sheet_name,
            start_row + inicio,
            start_col,
            len(bloco),
            qtd_colunas,
        )

        client.update_values(
            spreadsheet_id,
            range_a1,
            bloco,
            value_input_option=value_input_option,
        )


# =========================
# BLOCO 1.1 - Cart_Validador
# =========================

def atualizar_cart_validador(client: SheetsClient) -> None:
    print("Iniciando etapa 1/3: Cart_Validador")
    atualizar_status(client, "Etapa 1 de 3")

    client.clear_values(
        PLANILHA_DESTINO_ID,
        [f"{aba('Cart_Validador')}!A1:B"],
    )

    col_g, col_as = client.batch_get(
        PLANILHA_ORIGEM_ID,
        [
            f"{aba('Carteira_Validações')}!G1:G",
            f"{aba('Carteira_Validações')}!AS1:AS",
        ],
    )

    last_row = len(col_as)

    if last_row == 0:
        print("Cart_Validador sem dados para importar.")
        return

    saida = []

    for i in range(last_row):
        valor_g = col_g[i][0] if i < len(col_g) and col_g[i] else ""
        valor_as = col_as[i][0] if i < len(col_as) and col_as[i] else ""

        saida.append([valor_g, valor_as])

    update_chunked(
        client,
        PLANILHA_DESTINO_ID,
        "Cart_Validador",
        1,
        1,
        saida,
    )

    print(f"Etapa 1/3 concluída. Linhas importadas: {len(saida)}")


# =========================
# BLOCO 1.2 - Carteira
# =========================

def atualizar_carteira(client: SheetsClient) -> None:
    print("Iniciando etapa 2/3: Carteira")
    atualizar_status(client, "Etapa 2 de 3")

    client.clear_values(
        PLANILHA_DESTINO_ID,
        [
            f"{aba('Carteira')}!C1:AW",
            f"{aba('Carteira')}!B2:B",
        ],
    )

    start_row = 5

    col_a = client.get_values(
        PLANILHA_ORIGEM_ID,
        f"{aba('Carteira')}!A:A",
    )

    last_src = len(col_a)

    if last_src < start_row:
        print("Carteira origem sem dados a partir da linha 5.")
        return

    num_rows = last_src - (start_row - 1)

    blocos = client.batch_get(
        PLANILHA_ORIGEM_ID,
        [
            f"{aba('Carteira')}!A{start_row}:AD{last_src}",
            f"{aba('Carteira')}!AJ{start_row}:AK{last_src}",
            f"{aba('Carteira')}!AW{start_row}:AX{last_src}",
            f"{aba('Carteira')}!CA{start_row}:CF{last_src}",
            f"{aba('Carteira')}!BQ{start_row}:CY{last_src}",
            f"{aba('Carteira')}!CG{start_row}:CH{last_src}",
        ],
    )

    blk_a_ad = matriz_padrao(blocos[0], num_rows, 30)
    blk_aj_ak = matriz_padrao(blocos[1], num_rows, 2)
    blk_aw_ax = matriz_padrao(blocos[2], num_rows, 2)
    blk_ca_cf = matriz_padrao(blocos[3], num_rows, 6)
    blk_bq_cy = matriz_padrao(blocos[4], num_rows, 35)
    blk_cg_ch = matriz_padrao(blocos[5], num_rows, 2)

    saida = []

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

        saida.append(row)

    update_chunked(
        client,
        PLANILHA_DESTINO_ID,
        "Carteira",
        1,
        3,
        saida,
    )

    # Cálculo da coluna B, equivalente ao bloco do Apps Script.
    cfg_vals = client.get_values(
        PLANILHA_DESTINO_ID,
        f"{aba('BD_Config')}!B4:B9",
    )

    cfg_set = {
        chave(row[0])
        for row in cfg_vals
        if row and not vazio(row[0])
    }

    plan_vals = client.get_values(
        PLANILHA_DESTINO_ID,
        f"{aba('Carteira_Planejador')}!M:M",
    )

    freq: Dict[str, int] = {}

    for row in plan_vals:
        if not row or vazio(row[0]):
            continue

        key = chave(row[0])
        freq[key] = freq.get(key, 0) + 1

    if num_rows >= 2:
        b_out = []

        for i in range(1, num_rows):
            col_c = saida[i][0]
            col_d = saida[i][1]
            col_o = saida[i][12]

            if vazio(col_c):
                valor = ""

            elif chave(col_d) != "OBRA RETIRADA" and chave(col_o) in cfg_set:
                valor = freq.get(chave(col_c), 0)

            else:
                valor = "-"

            b_out.append([valor])

        update_chunked(
            client,
            PLANILHA_DESTINO_ID,
            "Carteira",
            2,
            2,
            b_out,
        )

    print(f"Etapa 2/3 concluída. Linhas processadas: {num_rows}")


# =========================
# BLOCO 1.3 - Entrada
# =========================

def atualizar_entrada_nova(client: SheetsClient) -> None:
    print("Iniciando etapa 3/3: Entrada")
    atualizar_status(client, "Etapa 3 de 3")

    client.clear_values(
        PLANILHA_DESTINO_ID,
        [f"{aba('Entrada')}!B6:AD"],
    )

    validador = client.get_values(
        PLANILHA_DESTINO_ID,
        f"{aba('Cart_Validador')}!A:B",
    )

    obs_lookup: Dict[str, Any] = {}

    for row in validador:
        if not row or vazio(row[0]):
            continue

        key = chave(row[0])

        if key not in obs_lookup:
            obs_lookup[key] = row[1] if len(row) > 1 else ""

    carteira_raw = client.get_values(
        PLANILHA_DESTINO_ID,
        f"{aba('Carteira')}!B1:AW",
    )

    carteira = matriz_padrao(carteira_raw, len(carteira_raw), 48)

    primeira_ocorrencia_por_projeto: Dict[str, Dict[str, Any]] = {}

    for row in carteira:
        projeto = valor_linha(row, "B", "C")
        key = chave(projeto)

        if vazio(key) or key in primeira_ocorrencia_por_projeto:
            continue

        primeira_ocorrencia_por_projeto[key] = {
            "H": valor_linha(row, "B", "H"),
            "AA": valor_linha(row, "B", "AA"),
            "AF": valor_linha(row, "B", "AF"),
        }

    saida = []

    for row in carteira:
        status_b = valor_linha(row, "B", "B")

        if not eh_zero(status_b):
            continue

        projeto = valor_linha(row, "B", "C")
        projeto_key = chave(projeto)

        lookup = primeira_ocorrencia_por_projeto.get(projeto_key, {})

        nova = [""] * 28  # Entrada C:AD

        # C, D, E
        definir_valor_entrada(
            nova,
            "C",
            2 if chave(lookup.get("H", "")) == "APTA" else 0,
        )

        definir_valor_entrada(
            nova,
            "D",
            2 if eh_data(lookup.get("AA", "")) else 0,
        )

        definir_valor_entrada(
            nova,
            "E",
            2 if eh_data(lookup.get("AF", "")) else 0,
        )

        # Mesma distribuição das fórmulas FILTER do Apps Script.
        definir_valor_entrada(nova, "F", valor_linha(row, "B", "G"))
        definir_valor_entrada(nova, "G", valor_linha(row, "B", "E"))
        definir_valor_entrada(nova, "H", valor_linha(row, "B", "D"))

        definir_valor_entrada(nova, "L", valor_linha(row, "B", "C"))
        definir_valor_entrada(nova, "M", valor_linha(row, "B", "S"))

        definir_valor_entrada(nova, "N", valor_linha(row, "B", "Q"))
        definir_valor_entrada(nova, "O", valor_linha(row, "B", "R"))

        definir_valor_entrada(nova, "Q", valor_linha(row, "B", "U"))

        definir_valor_entrada(nova, "R", valor_linha(row, "B", "V"))
        definir_valor_entrada(nova, "S", valor_linha(row, "B", "W"))

        definir_valor_entrada(nova, "T", valor_linha(row, "B", "X"))
        definir_valor_entrada(nova, "U", valor_linha(row, "B", "Y"))

        definir_valor_entrada(nova, "X", valor_linha(row, "B", "T"))
        definir_valor_entrada(nova, "Y", obs_lookup.get(projeto_key, ""))

        definir_valor_entrada(nova, "Z", valor_linha(row, "B", "Z"))
        definir_valor_entrada(nova, "AA", valor_linha(row, "B", "AA"))

        definir_valor_entrada(nova, "AB", valor_linha(row, "B", "AD"))
        definir_valor_entrada(nova, "AC", valor_linha(row, "B", "AE"))

        definir_valor_entrada(nova, "AD", valor_linha(row, "B", "O"))

        saida.append(nova)

    if saida:
        update_chunked(
            client,
            PLANILHA_DESTINO_ID,
            "Entrada",
            6,
            3,
            saida,
        )

    finalizar_status_com_data(client)

    print(f"Etapa 3/3 concluída. Linhas enviadas para Entrada: {len(saida)}")


# =========================
# EXECUÇÃO PRINCIPAL
# =========================

def main() -> None:
    inicio = time.time()

    print("Robô Bloco 1 - Entrada iniciado.")
    print(f"Planilha destino: {PLANILHA_DESTINO_ID}")
    print(f"Planilha origem: {PLANILHA_ORIGEM_ID}")

    client = SheetsClient()

    atualizar_cart_validador(client)
    atualizar_carteira(client)
    atualizar_entrada_nova(client)

    duracao = round(time.time() - inicio, 2)

    print(f"Robô Bloco 1 finalizado com sucesso em {duracao}s.")


if __name__ == "__main__":
    main()
