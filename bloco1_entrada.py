import os
import traceback
from datetime import datetime

import gspread

from common import (
    LISTA_PLANILHAS_SPREADSHEET_ID,
    TIMEZONE,
    abrir_aba,
    agora_formatado,
    as_text,
    buscar_planilhas,
    escrever_celula,
    escrever_matriz,
    executar_com_retry,
    formatar_data,
    get_gspread_client,
    is_blank,
    ler_range,
    limpar_intervalos,
    ultima_linha_preenchida_por_coluna,
)


# =========================================================
# CONFIGURAÇÕES ESPECÍFICAS DO BLOCO 1
# =========================================================
ORIGEM_SPREADSHEET_ID = os.getenv(
    "ORIGEM_SPREADSHEET_ID",
    "1lUNIeWCddfmvJEjWJpQMtuR4oRuMsI3VImDY0xBp3Bs",
)


# =========================================================
# HELPERS ESPECÍFICOS DO BLOCO 1
# =========================================================
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

    Faixa de serial 20000..90000 ~ 1954..2146 (datas plausíveis de obra).
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


# =========================================================
# CACHE DA ORIGEM (lida 1x, reusada em todas as planilhas)
# =========================================================
def preparar_dados_origem(ss_orig: gspread.Spreadsheet) -> dict:
    """
    Lê a planilha de origem UMA vez e devolve as matrizes já transformadas.

    A origem é idêntica para todos os destinos, então tanto a saída do
    Cart_Validador (G+AS) quanto a matriz C:AW da Carteira são iguais em todas
    as planilhas. Calcular aqui uma única vez evita reler a origem 10x e foi o
    que estourava a quota "Read requests per minute per user" do Sheets (429).
    """
    print("Preparando dados da origem (leitura única)...")

    # --- Cart_Validador: G + AS de Carteira_Validações ---
    aba_validacoes_orig = abrir_aba(ss_orig, "Carteira_Validações")
    ultima_linha_as = ultima_linha_preenchida_por_coluna(aba_validacoes_orig, 45)

    saida_validador = []

    if ultima_linha_as == 0:
        print("[origem] Carteira_Validações sem dados na coluna AS.")
    else:
        col_g = ler_range(aba_validacoes_orig, f"G1:G{ultima_linha_as}", ultima_linha_as, 1)
        col_as = ler_range(aba_validacoes_orig, f"AS1:AS{ultima_linha_as}", ultima_linha_as, 1)

        for i in range(ultima_linha_as):
            saida_validador.append(
                [
                    col_g[i][0] if i < len(col_g) else "",
                    col_as[i][0] if i < len(col_as) else "",
                ]
            )

    # --- Carteira: matriz C:AW a partir dos blocos da origem ---
    aba_carteira_orig = abrir_aba(ss_orig, "Carteira")
    start_row = 5
    last_src = ultima_linha_preenchida_por_coluna(aba_carteira_orig, 1)
    num_rows = max(0, last_src - (start_row - 1))

    saida_c_aw = []

    if num_rows == 0:
        print("[origem] Nenhum dado encontrado na Carteira a partir da linha 5.")
    else:
        blk_a_ad = ler_range(aba_carteira_orig, f"A{start_row}:AD{last_src}", num_rows, 30)
        blk_aj_ak = ler_range(aba_carteira_orig, f"AJ{start_row}:AK{last_src}", num_rows, 2)
        blk_aw_ax = ler_range(aba_carteira_orig, f"AW{start_row}:AX{last_src}", num_rows, 2)
        blk_ca_cf = ler_range(aba_carteira_orig, f"CA{start_row}:CF{last_src}", num_rows, 6)
        blk_bq_cy = ler_range(aba_carteira_orig, f"BQ{start_row}:CY{last_src}", num_rows, 35)
        blk_cg_ch = ler_range(aba_carteira_orig, f"CG{start_row}:CH{last_src}", num_rows, 2)

        # Colunas de data da origem, lidas com SERIAL_NUMBER.
        #
        # O FORMATTED_STRING padrão devolve a data já renderizada pelo formato da
        # célula (ex.: "dez./24") e a escrita RAW gravava isso como TEXTO no destino.
        # Com serial, data real vira número (valor de data de verdade, renderizado
        # pelo formato que o destino já tem) e o que é texto na origem
        # (ex.: "jan./25") continua vindo como texto.
        #
        # Destino -> origem: D=U, AA=BQ, AF=CE, AG=CQ, AU=CG.
        # BQ, CE, CQ e CG caem dentro de BQ:CY, então uma única leitura serial
        # desse range cobre as quatro; só o U precisa de leitura própria.
        blk_u_serial = ler_range(
            aba_carteira_orig,
            f"U{start_row}:U{last_src}",
            num_rows,
            1,
            date_time_render_option="SERIAL_NUMBER",
        )
        blk_bq_cy_serial = ler_range(
            aba_carteira_orig,
            f"BQ{start_row}:CY{last_src}",
            num_rows,
            35,
            date_time_render_option="SERIAL_NUMBER",
        )

        # Índices dentro de blk_bq_cy_serial (BQ = 0)
        idx_bq = 0
        idx_ce = 14
        idx_cq = 26
        idx_cg = 16

        for i in range(num_rows):
            row = [""] * 47

            r_a = blk_a_ad[i]
            r_aj = blk_aj_ak[i]
            r_aw = blk_aw_ax[i]
            r_ca = blk_ca_cf[i]
            r_bq = blk_bq_cy[i]
            r_cg = blk_cg_ch[i]
            r_bq_ser = blk_bq_cy_serial[i]

            # C = A
            row[0] = r_a[0]

            # D = U (serial, para não colar data como texto)
            row[1] = blk_u_serial[i][0]

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

            # Z = CA, AB = CC, AC = CF, AF = CE (serial: data, não texto)
            row[23] = r_ca[0]
            row[25] = r_ca[2]
            row[26] = r_ca[5]
            row[29] = r_bq_ser[idx_ce]

            # AA = BQ (serial: data, não texto)
            row[24] = r_bq_ser[idx_bq]

            # AD:AE = AW:AX
            row[27] = r_aw[0]
            row[28] = r_aw[1]

            # AG:AM = CQ:CW
            for k in range(7):
                row[30 + k] = r_bq[26 + k]

            # AG = CQ (serial: data, não texto)
            row[30] = r_bq_ser[idx_cq]

            # AN:AR = BR:BV
            row[37] = r_bq[1]
            row[38] = r_bq[2]
            row[39] = r_bq[3]
            row[40] = r_bq[4]
            row[41] = r_bq[5]

            # AS:AT = CX:CY
            row[42] = r_bq[33]
            row[43] = r_bq[34]

            # AU:AV = CG:CH (AU serial: data, não texto)
            row[44] = r_bq_ser[idx_cg]
            row[45] = r_cg[1]

            saida_c_aw.append(row)

    return {
        "saida_validador": saida_validador,
        "saida_c_aw": saida_c_aw,
        "num_rows": num_rows,
    }


# =========================================================
# BLOCO 1.1 - atualizarCart_Validador
# =========================================================
def atualizar_cart_validador(
    ss_dest: gspread.Spreadsheet,
    dados_origem: dict,
) -> None:
    print("[1/3] Atualizando Cart_Validador...")

    aba_entrada = abrir_aba(ss_dest, "Entrada")
    aba_cart_validador_dest = abrir_aba(ss_dest, "Cart_Validador")

    escrever_celula(aba_entrada, "F2", "Etapa 1 de 3")

    limpar_intervalos(aba_cart_validador_dest, ["A1:B"])

    saida = dados_origem["saida_validador"]

    if not saida:
        print("[1/3] Origem sem dados. Cart_Validador ficou vazio.")
        return

    escrever_matriz(aba_cart_validador_dest, 1, 1, saida)

    print(f"[1/3] Cart_Validador atualizado com {len(saida)} linhas.")


# =========================================================
# BLOCO 1.2 - atualizarCarteira
# =========================================================
def atualizar_carteira(
    ss_dest: gspread.Spreadsheet,
    dados_origem: dict,
) -> None:
    print("[2/3] Atualizando Carteira...")

    aba_entrada = abrir_aba(ss_dest, "Entrada")
    aba_carteira_dest = abrir_aba(ss_dest, "Carteira")

    escrever_celula(aba_entrada, "F2", "Etapa 2 de 3")

    num_rows = dados_origem["num_rows"]
    saida_c_aw = dados_origem["saida_c_aw"]

    limpar_intervalos(aba_carteira_dest, ["C1:AW", "B2:B"])

    if num_rows == 0:
        print("[2/3] Nenhum dado encontrado na origem a partir da linha 5.")
        return

    escrever_matriz(aba_carteira_dest, 1, 3, saida_c_aw)

    # Garante formato de data nas colunas gravadas como serial. Sem isso, numa
    # célula em "Automático" o serial aparece como 45627 em vez de 01/12/2024
    # (metade das planilhas de destino estava assim). Da linha 2 pra baixo, para
    # não tocar o cabeçalho da linha 1. Um único request por planilha.
    ultima_linha_carteira = num_rows

    if ultima_linha_carteira >= 2:
        formatar_data(
            aba_carteira_dest,
            [
                f"D2:D{ultima_linha_carteira}",
                f"AA2:AA{ultima_linha_carteira}",
                f"AF2:AF{ultima_linha_carteira}",
                f"AG2:AG{ultima_linha_carteira}",
                f"AU2:AU{ultima_linha_carteira}",
            ],
        )

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

    # Segunda leitura das colunas de data da Carteira, agora como serial.
    #
    # carteira_rows vem com FORMATTED_STRING, então uma data em D ou AA volta
    # renderizada pelo formato da célula (ex.: "dez./24") e a escrita RAW gravava
    # isso como TEXTO na Entrada, desfazendo na Entrada o que o Bloco 1.2 acertou
    # na Carteira. C:AF num único range cobre a chave (C) e as três colunas
    # necessárias, custando 1 request em vez de 3.
    #
    # Também alimenta os flags das colunas D e E da Entrada: is_date_like devolve
    # False para "dez./24" (nenhum formato de data casa com esse texto) e True
    # para o serial, que é o comportamento do ISDATE original.
    carteira_datas = ler_range(
        aba_carteira,
        f"C1:AF{last_carteira}",
        last_carteira,
        30,
        date_time_render_option="SERIAL_NUMBER",
    )

    # Índices dentro de carteira_datas (C = 0)
    idx_dt_c = 0
    idx_dt_d = 1
    idx_dt_aa = 24
    idx_dt_af = 29

    last_validador = ultima_linha_preenchida_por_coluna(aba_cart_validador, 1)

    cart_validador_rows = (
        ler_range(aba_cart_validador, f"A1:B{last_validador}", last_validador, 2)
        if last_validador > 0
        else []
    )

    mapa_observacao = montar_mapa_primeira_ocorrencia(cart_validador_rows, 0, 1)

    # Mapas para simular XLOOKUP em Carteira!C:C
    mapa_c_para_h = montar_mapa_primeira_ocorrencia(carteira_rows, 1, 6)    # C -> H

    # AA e AF saem da leitura serial: os flags precisam ver data, não texto.
    mapa_c_para_aa = montar_mapa_primeira_ocorrencia(carteira_datas, idx_dt_c, idx_dt_aa)
    mapa_c_para_af = montar_mapa_primeira_ocorrencia(carteira_datas, idx_dt_c, idx_dt_af)

    saida_c_ad = []

    for i, row in enumerate(carteira_rows):
        row_dt = carteira_datas[i]

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

        # H = Carteira!D (serial: data, não texto)
        out[5] = row_dt[idx_dt_d]

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

        # Z:AA = Carteira!Z:AA (AA serial: data, não texto)
        out[23] = valor_coluna(row, 26)
        out[24] = row_dt[idx_dt_aa]

        # AB:AC = Carteira!AD:AE
        out[25] = valor_coluna(row, 30)
        out[26] = valor_coluna(row, 31)

        # AD = Carteira!O
        out[27] = valor_coluna(row, 15)

        saida_c_ad.append(out)

    if saida_c_ad:
        escrever_matriz(aba_entrada, 6, 3, saida_c_ad)

        # H recebe serial (Carteira!D): garante formato de data. AA também vem de
        # coluna de data, mas na amostra só apareceu texto ("-"); formatada junto
        # porque quando vier data precisa exibir data.
        ultima_linha_saida = 6 + len(saida_c_ad) - 1

        formatar_data(
            aba_entrada,
            [
                f"H6:H{ultima_linha_saida}",
                f"AA6:AA{ultima_linha_saida}",
            ],
        )

    finalizar_execucao(aba_entrada)

    print(f"[3/3] Entrada atualizada com {len(saida_c_ad)} linhas.")


def finalizar_execucao(aba_entrada: gspread.Worksheet) -> None:
    escrever_celula(aba_entrada, "F2", agora_formatado())


# =========================================================
# EXECUÇÃO DE UMA PLANILHA
# =========================================================
def executar_bloco1_para_planilha(
    client: gspread.Client,
    dados_origem: dict,
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

    atualizar_cart_validador(ss_dest, dados_origem)
    atualizar_carteira(ss_dest, dados_origem)
    atualizar_entrada_nova(ss_dest)

    print(f"[OK] Planilha concluída: {nome_planilha} | {dest_spreadsheet_id}")


# =========================================================
# MAIN
# =========================================================
def main() -> None:
    inicio = datetime.now(TIMEZONE)

    print(f"Início geral do Bloco 1 - Entrada: {inicio.strftime('%d/%m/%Y %H:%M:%S')}")

    client = get_gspread_client()

    ss_lista = executar_com_retry(lambda: client.open_by_key(LISTA_PLANILHAS_SPREADSHEET_ID))
    ss_orig = executar_com_retry(lambda: client.open_by_key(ORIGEM_SPREADSHEET_ID))

    planilhas = buscar_planilhas(ss_lista)

    if not planilhas:
        raise RuntimeError(
            f"Nenhum ID encontrado na aba BD_Planilhas!C3:C "
            f"na planilha {LISTA_PLANILHAS_SPREADSHEET_ID}."
        )

    print(f"Total de planilhas encontradas: {len(planilhas)}")

    # Lê a origem uma única vez; reusada em todas as planilhas (evita 429).
    dados_origem = preparar_dados_origem(ss_orig)

    sucessos = []
    erros = []

    for indice, item in enumerate(planilhas, start=1):
        nome_planilha = item["nome"]
        dest_spreadsheet_id = item["id"]

        try:
            executar_bloco1_para_planilha(
                client=client,
                dados_origem=dados_origem,
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
            print("[ERRO] Falha ao processar uma planilha.")
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
    print("RESUMO FINAL")
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
            f"Processamento finalizado com erro em {len(erros)} planilha(s). "
            f"Verifique o log acima."
        )

    print("")
    print("Todas as planilhas foram processadas com sucesso.")


if __name__ == "__main__":
    main()
