# Dados de Preços — Formato e Instruções

O arquivo `precos_b3.csv` **não está incluído neste repositório** por poder ser originário de fonte proprietária. Você precisa fornecer o seu próprio arquivo.

## Formato esperado

```
data,PETR4,VALE3,ITUB4,...,IBOV,BOVA11,DIVO11
2025-06-03,38.42,61.10,33.20,...,127834,127.45,143.20
2025-06-04,38.91,60.87,33.55,...,128102,128.01,143.85
...
```

### Requisitos

| Campo | Requisito |
|---|---|
| Primeira coluna | `data` no formato `YYYY-MM-DD` |
| Demais colunas | Tickers B3 em maiúsculas (ex: `PETR4`, `VALE3`) |
| Coluna `IBOV` | **Obrigatória** — usada para Beta e Alpha de Jensen |
| Coluna `DIVO11` | Opcional — usada para o Information Ratio vs. IDIV |
| Tipo de preço | Fechamento **ajustado** por proventos e eventos corporativos |
| Período | Mínimo recomendado: 252 pregões (~1 ano) |

## Como obter os dados

### Opção 1: yfinance (gratuito)

```python
import yfinance as yf
import pandas as pd

tickers_b3 = [
    "PETR4.SA", "VALE3.SA", "ITUB4.SA", "BBDC4.SA", "ABEV3.SA",
    "TAEE11.SA", "CPLE3.SA", "VIVT3.SA", "ALOS3.SA", "BBAS3.SA",
    "BOVA11.SA", "DIVO11.SA"
]
ibov_ticker = "^BVSP"

all_tickers = tickers_b3 + [ibov_ticker]
df = yf.download(all_tickers, start="2025-06-01", end="2026-06-01", auto_adjust=True)["Close"]

# Renomear colunas
rename_map = {t: t.replace(".SA", "") for t in tickers_b3}
rename_map["^BVSP"] = "IBOV"
df = df.rename(columns=rename_map)

df.index.name = "data"
df.to_csv("precos_b3.csv")
print(f"Arquivo gerado com {len(df)} pregões e {len(df.columns)} ativos.")
```

### Opção 2: Economática / Bloomberg / Reuters

Exporte o histórico de preços ajustados em CSV, certifique-se de que o formato seja compatível (data na primeira coluna, tickers nas demais).

### Opção 3: B3 / Bora Investir

A B3 disponibiliza séries históricas de preços em [bvmf.bmfbovespa.com.br](https://www.b3.com.br/pt_br/market-data-e-indices/servicos-de-dados/market-data/historico/mercado-a-vista/cotacoes-historicas/).

## Tickers das carteiras analisadas

Todos precisam estar presentes no seu CSV para reproduzir a análise das 5 carteiras:

```
ABEV3, ALOS3, AXIA3, BBAS3, BBDC4, BPAC11, BRAP4, CMIN3,
CPLE3, CXSE3, CURY3, EGIE3, ENGI11, ITUB4, PETR4, PRIO3,
SAPR11, SLCE3, VALE3, VIVT3
```

Mais IBOV e DIVO11 como benchmarks.

## Problemas comuns

**Ticker não encontrado:** Verifique se o código B3 está correto. Alguns ETFs têm sufixo diferente (ex: BOVA11, DIVO11). Use a substituição no script para casos especiais (ex: AXIA6 → AXIA3).

**Valores faltantes (NaN):** Os scripts fazem `ffill(limit=5)` — preenchimento forward com limite de 5 dias. Para gaps maiores, o ativo pode ser excluído automaticamente.

**Índices vs. ETFs:** O IBOV é o índice (ponto), não o ETF BOVA11. O yfinance retorna `^BVSP` para o índice. O DIVO11 é o ETF que replica o IDIV (com diferença de come-cotas e taxa de administração).
