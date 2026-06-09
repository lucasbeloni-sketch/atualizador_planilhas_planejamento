import traceback
from datetime import datetime

import gspread

from common import (
    LISTA_PLANILHAS_SPREADSHEET_ID,
    TIMEZONE,
    abrir_aba,
    agora_formatado,
    buscar_planilhas,
    congelar_intervalo,
    escrever_celula,
    executar_com_retry,
    formatar_data_hora,
    get_gspread_client,
    limpar_intervalos,
    ultima_linha_preenchida_por_coluna,
)
from gspread.utils import rowcol_to_a1


# =========================================================
# HELPERS ESPECÍFICOS DO BLOCO 2
# =========================================================
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

    # congelar_intervalo aguarda o cálculo estabilizar antes de ler/congelar.
    print(f"Aguardando cálculo e congelando {nome_bloco}: {range_destino}")
    congelar_intervalo(worksheet, range_destino)


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
    escrever_celula(worksheet, "F3", agora_formatado(), raw=False)
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
            f"Nenhum ID encontrado na aba BD_Planilhas!C3:C "
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
