# Atualizador de Planilhas de Planejamento

Automatiza a atualização de planilhas Google Sheets de planejamento, replicando
em Python a lógica que antes rodava em Apps Script. Roda via GitHub Actions.

## Blocos

| Script | Aba alvo | O que faz |
|--------|----------|-----------|
| `bloco1_entrada.py` | `Cart_Validador`, `Carteira`, `Entrada` | Copia dados da planilha de origem, monta a Carteira e gera a aba Entrada. |
| `bloco2_carteira_planejador.py` | `Carteira_Planejador` | Aplica fórmulas por blocos de colunas, aguarda o cálculo e congela como valores. |
| `bloco3_plan_principal.py` | `Plan_Principal` | Aplica fórmulas (XLOOKUP/SUMIFS/etc), congela valores e reaplica formatações fixas. |

`common.py` reúne tudo que é compartilhado: autenticação, retry, leitura/escrita
no Sheets e helpers de matriz. **Corrigir um bug nele corrige nos três blocos.**

A lista de planilhas-alvo vem da aba `BD_Planilhas` (coluna B = nome, C = ID,
D = valor usado em `Plan_Principal!BE` no Bloco 3), a partir da linha 3.

## Configuração

Credencial de service account, por uma de duas vias:

- **GitHub Actions:** secret `GOOGLE_CREDENTIALS_B64` (JSON da service account em base64).
- **Local:** arquivo `service_account.json` na raiz (ignorado pelo git), ou variável
  `GOOGLE_APPLICATION_CREDENTIALS` apontando para ele.

Variáveis de ambiente opcionais:

| Variável | Default | Uso |
|----------|---------|-----|
| `LISTA_PLANILHAS_SPREADSHEET_ID` | (id fixo no código) | Planilha com a aba `BD_Planilhas`. |
| `ORIGEM_SPREADSHEET_ID` | (id fixo no código) | Origem dos dados (só Bloco 1). |
| `CHUNK_SIZE` | `5000` | Linhas por requisição de escrita. |
| `CALC_WAIT_SECONDS` | `15` | Espera para o Sheets calcular fórmulas (Blocos 2 e 3). |

## Rodar local

```bash
pip install -r requirements.txt
python bloco1_entrada.py
python bloco2_carteira_planejador.py
python bloco3_plan_principal.py
```

## Workflows (GitHub Actions)

Todos os workflows compartilham o mesmo grupo de concorrência
(`planejamento-sheets-write`) para **nunca rodarem em paralelo** sobre as mesmas
planilhas. O pipeline B1+B2 roda agendado (12h e 20h de Brasília); os demais são
disparados manualmente (`workflow_dispatch`).
