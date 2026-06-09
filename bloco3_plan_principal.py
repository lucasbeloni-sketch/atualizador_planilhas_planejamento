import traceback
from datetime import datetime

import gspread

from common import (
    LISTA_PLANILHAS_SPREADSHEET_ID,
    TIMEZONE,
    abrir_aba,
    agora_formatado,
    aguardar_estabilizacao,
    as_text,
    buscar_planilhas,
    congelar_intervalo,
    escrever_celula,
    escrever_formulas_matriz,
    executar_com_retry,
    formatar_data_hora,
    get_gspread_client,
    is_blank,
    ler_coluna,
    ler_range,
    limpar_intervalos,
)


# =========================================================
# CONFIGURAÇÕES ESPECÍFICAS DO BLOCO 3
# =========================================================
# O Plan_Principal lê datas como serial numérico para congelar e reaplicar formato.
RENDER_SERIAL = "SERIAL_NUMBER"


# =========================================================
# HELPERS ESPECÍFICOS DO BLOCO 3
# =========================================================
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
            date_time_render_option=RENDER_SERIAL,
        )
    )

    for idx in range(len(valores) - 1, -1, -1):
        row = valores[idx]
        if any(not is_blank(cell) for cell in row):
            return idx + 1

    return 0


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


def _norm_chave(valor):
    """
    Normaliza um valor para casar com a semântica do COUNTIFS do Sheets:
    números comparados por valor (5 == 5.0) e texto case-insensitive.
    """
    if isinstance(valor, bool):
        return valor

    if isinstance(valor, (int, float)):
        return float(valor)

    return as_text(valor).strip().lower()


def construir_formulas_am(
    aba: gspread.Worksheet,
    linhas: list[int],
    last_row_sheet: int,
) -> list[list]:
    """
    AM = meta (BD_Metas!D via XLOOKUP por G&B) apenas na PRIMEIRA ocorrência de
    cada par (B, G) lendo de cima para baixo; nas repetições, 0; com B vazio, "".

    Substitui a fórmula original que usava COUNTIFS sobre intervalo crescente
    ($B$1:B{row-1}), de custo O(n^2) e principal gargalo do cálculo de AL:AS,
    por detecção da primeira ocorrência em Python (O(n)). O XLOOKUP em BD_Metas
    continua sendo feito pelo Sheets, preservando a concatenação G&B original.
    """
    col_b = ler_coluna(aba, 2)
    col_g = ler_coluna(aba, 7)

    def valor(col: list, idx: int):
        return col[idx] if idx < len(col) else ""

    vistos = set()
    am_por_linha = {}

    # Percorre desde a linha 1 (o COUNTIFS original considerava $B$1 em diante),
    # mas só emite fórmula para as linhas de dados (>= 6).
    for idx in range(last_row_sheet):
        row = idx + 1
        b = valor(col_b, idx)
        g = valor(col_g, idx)
        chave = (_norm_chave(b), _norm_chave(g))

        if row >= 6:
            if is_blank(b):
                am_por_linha[row] = ""
            elif chave in vistos:
                am_por_linha[row] = 0
            else:
                am_por_linha[row] = (
                    f'=XLOOKUP(G{row}&B{row};BD_Metas!$A:$A;BD_Metas!$D:$D;0)'
                )

        vistos.add(chave)

    return [[am_por_linha[row]] for row in linhas]


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

    print("Aplicando AM (primeira ocorrência de (B,G) calculada em Python)...")
    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=39,
        formulas=construir_formulas_am(aba, linhas, last_row_sheet),
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

    # As derivadas (AN/AP/AR) dividem por AL/AM/AO/AQ; aguarda o bloco inicial
    # estabilizar antes de aplicá-las.
    print("Aguardando cálculo inicial em AL:AS...")
    aguardar_estabilizacao(
        aba,
        f"AL6:AS{last_row_sheet}",
        date_time_render_option=RENDER_SERIAL,
        descricao="cálculo inicial AL:AS",
    )

    # =====================================================
    # Fórmulas derivadas
    # AN, AP, AR
    # BR:BT
    # =====================================================
    print("Aplicando fórmulas derivadas em AN, AP e AR...")

    derivadas = [formulas_derivadas_an_ap_ar(row) for row in linhas]

    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=40,
        formulas=[[d[0]] for d in derivadas],
    )

    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=42,
        formulas=[[d[1]] for d in derivadas],
    )

    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=44,
        formulas=[[d[2]] for d in derivadas],
    )

    print("Aplicando fórmulas em BR:BT...")
    escrever_formulas_matriz(
        worksheet=aba,
        start_row=6,
        start_col=70,
        formulas=[formulas_br_bt(row) for row in linhas],
    )

    # =====================================================
    # Congelar valores
    # AL:AS, J:L, BR:BT
    # BE fica como fórmula, igual ao Apps Script original.
    # Cada congelar_intervalo aguarda o cálculo estabilizar antes de ler.
    # =====================================================
    print("Congelando valores em AL:AS...")
    congelar_intervalo(aba, f"AL6:AS{last_row_sheet}", date_time_render_option=RENDER_SERIAL)

    print("Congelando valores em J:L...")
    congelar_intervalo(aba, f"J6:L{last_row_sheet}", date_time_render_option=RENDER_SERIAL)

    print("Congelando valores em BR:BT...")
    congelar_intervalo(aba, f"BR6:BT{last_row_sheet}", date_time_render_option=RENDER_SERIAL)

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
        date_time_render_option=RENDER_SERIAL,
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

    print("Aguardando cálculo e congelando valores em AK...")
    congelar_intervalo(aba, f"AK6:AK{last_row_h}", date_time_render_option=RENDER_SERIAL)

    if last_row_h < last_row_sheet:
        limpar_intervalos(aba, [f"AK{last_row_h + 1}:AK"])


def finalizar_execucao(worksheet: gspread.Worksheet) -> None:
    escrever_celula(worksheet, "F3", agora_formatado(), raw=False)
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
            f"Nenhum ID encontrado na aba BD_Planilhas!C3:C "
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
