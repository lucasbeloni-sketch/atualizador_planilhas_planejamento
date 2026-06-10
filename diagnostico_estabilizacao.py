"""
Diagnóstico SOMENTE-LEITURA da causa do timeout de AL:AS no Bloco 3.

Hipótese: _contem_carregando faz match de SUBSTRING ("loading"/"carregando")
em qualquer célula do range. A coluna AS traz texto de serviço (BD_Serv_GPM!E
via XLOOKUP). Se esse texto contiver as substrings, aguardar_estabilizacao
nunca declara "estável" e estoura o timeout — independente do nº de linhas.

Este script NÃO escreve em nenhuma planilha. Apenas lê AL6:AS e procura os
marcadores, replicando exatamente a leitura de aguardar_estabilizacao.
"""
from gspread.utils import rowcol_to_a1

from common import (
    _MARCADORES_CARREGANDO,
    abrir_aba,
    executar_com_retry,
    get_gspread_client,
)

# (nome, id, resultado_observado_no_run_cancelado)
PLANILHAS = [
    ("BARREIRAS",            "1OTHF2ytEOjGgfE49paARXkz9GjaklOQC_UhiXwUjC2E", "timeout"),
    ("BOM JESUS DA LAPA",    "1rj2V7CxbZwkan63eCeLkH9G00Gi041IZNC6vwEgq6yI", "timeout"),
    ("BRUMADO",              "1oS619l3x_D1mXkvDpw8vs91G6ipZmsK83JqEIwPj7Uk", "timeout"),
    ("GUANAMBI",             "1FO5tyhXygbbzSmmTGdnm45j4DD_rRFQgEheN8T8Wy70", "timeout"),
    ("IBOTIRAMA",            "1dNwj8qWTl1k92PxI9iXwaNZYITnxuKP-kOF1QnZK3Iw", "ok"),
    ("IRECE",                "1NV0oObhLHAqnSpJKmeBBHQQxcxwlRh14TKQwO561GEw", "ok"),
    ("ITAPETINGA",           "1rzT8o6XZi4v8j7CYLky3BD3sT5IPjv1PRb45ipBfbw4", "timeout"),
    ("JEQUIE",               "1sGHf-zWXoxjnO20QBw2KWX39BSCzT8rzHdEz1hL7jyU", "ok"),
    ("LIVRAMENTO",           "1gN2tR_LCuRnVCQ9tm2UURnVuMlJPVNEjvmo02TwFQCI", "ok"),
    ("VITORIA DA CONQUISTA", "1XmpY8mqkRou-CRY68j1ljHH8W8zcROy7wnwMMSfbV7o", "cancelada"),
]

# AL..AS -> colunas 38..45 (AL=38). Mesmo range esperado em aguardar_estabilizacao.
COL_INICIAL = 38
RANGE_AL_AS = "AL6:AS"


def marcadores_na_celula(cell) -> list[str]:
    if not isinstance(cell, str):
        return []
    texto = cell.strip().lower()
    return [m for m in _MARCADORES_CARREGANDO if m in texto]


def diagnosticar(client, nome: str, sid: str, esperado: str) -> dict:
    print("")
    print("=" * 80)
    print(f"Planilha: {nome} | esperado no run: {esperado}")
    print(f"ID: {sid}")
    print("=" * 80)

    ss = executar_com_retry(lambda: client.open_by_key(sid))
    aba = abrir_aba(ss, "Plan_Principal")

    # Lê exatamente como aguardar_estabilizacao: UNFORMATTED_VALUE + SERIAL.
    valores = executar_com_retry(
        lambda: aba.get(
            RANGE_AL_AS,
            value_render_option="UNFORMATTED_VALUE",
            date_time_render_option="SERIAL_NUMBER",
        )
    )

    total_celulas = sum(len(row) for row in valores)
    matches = []  # (endereco, marcador, valor)

    for i, row in enumerate(valores):
        for j, cell in enumerate(row):
            achados = marcadores_na_celula(cell)
            for m in achados:
                endereco = rowcol_to_a1(6 + i, COL_INICIAL + j)
                matches.append((endereco, m, cell))

    print(f"Linhas lidas em AL:AS: {len(valores)} | células não vazias: {total_celulas}")
    print(f"Células que casam com {_MARCADORES_CARREGANDO}: {len(matches)}")

    if matches:
        print(">>> CONFIRMA a hipótese: aguardar_estabilizacao nunca estabilizaria.")
        for endereco, marcador, valor in matches[:10]:
            print(f'    {endereco}  [{marcador}]  ->  "{valor}"')
        if len(matches) > 10:
            print(f"    ... (+{len(matches) - 10} outras células)")
    else:
        print(">>> Sem marcadores. Estabilização não seria bloqueada por _contem_carregando.")

    return {
        "nome": nome,
        "esperado": esperado,
        "linhas": len(valores),
        "matches": len(matches),
        "exemplos": matches[:5],
    }


def main() -> None:
    client = get_gspread_client()

    resumo = []
    for nome, sid, esperado in PLANILHAS:
        try:
            resumo.append(diagnosticar(client, nome, sid, esperado))
        except Exception as erro:
            print(f"[ERRO] Falha ao ler {nome}: {erro}")
            resumo.append(
                {"nome": nome, "esperado": esperado, "linhas": -1, "matches": -1, "exemplos": []}
            )

    print("")
    print("=" * 80)
    print("RESUMO — marcadores 'loading'/'carregando' em AL6:AS por planilha")
    print("=" * 80)
    print(f"{'Planilha':<24}{'Run':<12}{'Linhas':>8}{'Matches':>10}")
    for r in resumo:
        print(f"{r['nome']:<24}{r['esperado']:<12}{r['linhas']:>8}{r['matches']:>10}")

    # Correlação esperada: matches > 0  <=>  resultado 'timeout'/'cancelada'.
    print("")
    print("Correlação (matches>0 deve casar com timeout/cancelada):")
    for r in resumo:
        tem = r["matches"] > 0
        problema = r["esperado"] in ("timeout", "cancelada")
        marca = "OK casa" if tem == problema else "NÃO casa"
        print(f"  {r['nome']:<24} matches={r['matches']:>4}  esperado={r['esperado']:<10} -> {marca}")


if __name__ == "__main__":
    main()
