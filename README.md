# Otimização de Markowitz — Carteiras de Dividendos B3

> Apliquei otimização de Markowitz (Max Sharpe) às carteiras de dividendos publicadas pelas principais casas de análise brasileiras em junho de 2026. O modelo perdeu. O motivo importa mais que o resultado.

**Artigo completo:** [LinkedIn — *Apliquei otimização de Markowitz às carteiras de dividendos de junho*](#) *(link será atualizado após publicação)*

---

## O que este repositório contém

Três scripts Python independentes que reproduzem integralmente a análise publicada no artigo:

| Script | O que faz | Output |
|---|---|---|
| `analisar_carteiras.py` | Calcula métricas de risco-retorno para as 5 carteiras e aplica o otimizador Max Sharpe | `comparacao_carteiras.xlsx` |
| `backtest_benchmark.py` | Backtest out-of-sample (split 50/50) e comparação com DIVO11 e IBOV | `backtest_benchmark.xlsx` |
| `monte_carlo.py` | Bootstrap com 500 splits aleatórios + simulação Monte Carlo forward (5.000 trajetórias) | `monte_carlo.xlsx` |
| `analisar_minha_carteira.py` | **Para o leitor:** analise sua própria carteira preenchendo `minha_carteira.xlsx` | `resultado_carteira.xlsx` |

---

## Para analisar sua própria carteira

1. Preencha `minha_carteira.xlsx` com seus tickers e pesos (aba "Minha Carteira")
2. Forneça seu CSV de preços (veja formato abaixo)
3. Execute:

```bash
python analisar_minha_carteira.py
```

4. Abra `resultado_carteira.xlsx` para ver os indicadores

---

## Dados de Preços

Os scripts precisam de um arquivo `precos_b3.csv` com preços de fechamento ajustados. **O arquivo não está incluído neste repositório** — você deve fornecer o seu próprio.

### Formato esperado

```
data,PETR4,VALE3,ITUB4,...,IBOV,BOVA11,DIVO11
2025-06-03,38.42,61.10,33.20,...,127834,127.45,143.20
2025-06-04,38.91,60.87,33.55,...,128102,128.01,143.85
...
```

- **Primeira coluna:** `data` no formato `YYYY-MM-DD`
- **Demais colunas:** tickers B3 em maiúsculas (ex: `PETR4`, `VALE3`)
- **Obrigatório:** coluna `IBOV` para o cálculo de Beta e Alpha de Jensen
- **Opcional:** coluna `DIVO11` para o Information Ratio vs. benchmark de dividendos
- **Preços:** fechamento ajustado por proventos e eventos corporativos

### Onde obter dados

- [Economática](https://economatica.com) — base profissional, exportação direta em CSV
- [Yahoo Finance](https://finance.yahoo.com) via `yfinance` (adicione `.SA` ao ticker: `PETR4.SA`)
- [Fundamentus](https://fundamentus.com.br) — dados históricos de fechamento
- [B3](https://b3.com.br) — histórico via Bora Investir

**Exemplo com yfinance:**
```python
import yfinance as yf
import pandas as pd

tickers = ["PETR4.SA","VALE3.SA","ITUB4.SA","^BVSP"]
df = yf.download(tickers, start="2025-06-01", end="2026-06-01")["Close"]
df.columns = [c.replace(".SA","").replace("^BVSP","IBOV") for c in df.columns]
df.to_csv("precos_b3.csv", index_label="data")
```

---

## Carteiras Analisadas

Carteiras de dividendos publicadas pelas 5 casas para junho de 2026. Fonte: Money Times (01–04/06/2026).

Detalhes completos em `Carteiras/carteiras_jun2026.json`.

| Casa | Nº Ativos | Destaques |
|---|---|---|
| XP Dividendos | 11 | CPLE3 15%, ITUB4 15%, VALE3 12,5% |
| BTG Dividendos | 12 | Pesos entre 5% e 10% |
| Itaú BBA | 5 | Pesos iguais 20%: AXIA3, ALOS3, BBDC4, VALE3, PETR4 |
| Santander | 10 | Pesos iguais 10% |
| BB Investimentos | 10 | Pesos iguais 10% |

**Nota:** AXIA6 (PN) substituído por AXIA3 (ON) — mesma empresa, série com histórico disponível.

---

## Modelo de Otimização

```
Maximizar  Sharpe(w) = [R_p(w) − RF] / σ_p(w)

Sujeito a  Σ wᵢ = 1
           α · wᵢ_rec ≤ wᵢ ≤ β · wᵢ_rec   ∀ i

Parâmetros: α = 0,5   β = 2,0   RF = 14,40% a.a.
```

- **Solver:** `scipy.optimize.minimize` com método SLSQP
- **Retornos:** log-retornos diários anualizados por 252 pregões
- **Covariância:** estimador amostral (sem shrinkage — ver limitações)

---

## Resultados Principais

### Full-Sample (pesos recomendados, jun/2025 – mai/2026)

| Carteira | Retorno Anual | Sharpe | IR vs DIVO11 | Alpha Jensen |
|---|---|---|---|---|
| **DIVO11** (proxy IDIV) | 21,3% | 0,44 | — | — |
| **IBOV** | 23,9% | 0,57 | — | — |
| XP Dividendos | 31,6% | 0,98 | 1,52 | +7,6% |
| BTG Dividendos | 28,7% | 0,77 | 0,98 | +4,4% |
| **Itaú BBA Div.** | **35,2%** | **1,16** | **1,63** | **+11,4%** |
| Santander Div. | 33,9% | 1,03 | 1,59 | +9,3% |
| BB Dividendos | 26,6% | 0,74 | 0,92 | +3,4% |

### Bootstrap Backtest (500 splits aleatórios)

| Carteira | Sharpe OOS Rec. | Sharpe IS Opt. | Sharpe OOS Opt. | Agrega? |
|---|---|---|---|---|
| XP Dividendos | 0,911 | 2,153 | 0,591 | ❌ Não |
| BTG Dividendos | 0,250 | 2,537 | 0,250 | ⚠ Inconclusivo |
| Itaú BBA | 0,470 | 3,007 | 0,273 | ❌ Não |
| Santander | 0,496 | 2,572 | 0,377 | ❌ Não |
| BB Investimentos | 0,369 | 1,909 | 0,164 | ❌ Não |

**Achado central:** Sharpe IS médio ~2,5 → Sharpe OOS mediano ~0,3. Degradação de ~88%. O otimizador clássico com janela curta não agrega valor sistematicamente.

---

## Instalação

```bash
git clone https://github.com/alissondpoliveira/Markowitz-carteiras-dividendos.git
cd markowitz-carteiras-dividendos
pip install -r requirements.txt
```

### Pré-requisitos

- Python 3.9+
- Bibliotecas: `pandas`, `numpy`, `scipy`, `openpyxl`

---

## Como Reproduzir

```bash
# 1. Coloque precos_b3.csv na pasta raiz do projeto

# 2. Análise completa das 5 carteiras
python analisar_carteiras.py       # → comparacao_carteiras.xlsx

# 3. Backtest OOS + comparação com benchmarks
python backtest_benchmark.py       # → backtest_benchmark.xlsx

# 4. Bootstrap 500 splits + Monte Carlo 5000 trajetórias
python monte_carlo.py              # → monte_carlo.xlsx

# 5. Análise da sua carteira personalizada
# (preencha minha_carteira.xlsx primeiro)
python analisar_minha_carteira.py  # → resultado_carteira.xlsx
```

---

## Limitações do Modelo

**Estimation error maximization:** O problema clássico do Markowitz com janelas curtas. Com ~100 pregões de dados, o erro padrão do retorno esperado anualizado é da ordem de ±3 p.p. por ativo. Com 10 ativos, o otimizador maximiza, junto com o Sharpe verdadeiro, o ruído de estimação. O resultado é Sharpe inflado in-sample e degradado out-of-sample.

Alternativas mais robustas (não implementadas neste repositório):
- **Ledoit-Wolf shrinkage** na matriz de covariância
- **Black-Litterman** combinando prior de mercado com views do analista
- **Janela mais longa:** mínimo recomendado de 3–5 anos para estimação estável

---

## Estrutura de Arquivos

```
markowitz-carteiras-dividendos/
├── README.md
├── requirements.txt
├── .gitignore
├── analisar_carteiras.py         # Análise principal
├── backtest_benchmark.py         # Backtest OOS + benchmark
├── monte_carlo.py                # Bootstrap + Monte Carlo
├── analisar_minha_carteira.py    # Script para carteira do leitor
├── minha_carteira.xlsx           # Template de entrada para o leitor
├── Carteiras/
│   └── carteiras_jun2026.json    # Metadados das 5 carteiras
├── dados/
│   └── README_dados.md           # Instruções sobre formato de dados
└── outputs/
    └── .gitkeep
```

---

## Referências

- Markowitz, H. (1952). Portfolio Selection. *Journal of Finance*, 7(1), 77–91.
- Ledoit, O. & Wolf, M. (2004). A well-conditioned estimator for large-dimensional covariance matrices. *Journal of Multivariate Analysis*, 88(2), 365–411.
- Black, F. & Litterman, R. (1992). Global Portfolio Optimization. *Financial Analysts Journal*, 48(5), 28–43.
- Michaud, R. (1989). The Markowitz Optimization Enigma: Is "Optimized" Optimal? *Financial Analysts Journal*, 45(1), 31–42.

---

## Disclaimer

> **Este repositório é exclusivamente para fins educacionais e de pesquisa.** O autor não é analista de valores mobiliários credenciado pela CVM. Nada neste repositório constitui recomendação de investimento, análise de valores mobiliários ou consultoria financeira. Rentabilidade passada não é garantia de rentabilidade futura. Antes de tomar qualquer decisão de investimento, consulte um profissional habilitado.
>
> A análise foi desenvolvida com suporte de Claude (Anthropic) para implementação do código e estruturação analítica. As decisões metodológicas, a interpretação dos resultados e a curadoria dos dados foram do autor.

---

*Alisson D. P. Oliveira · Engenharia de Produção · CFA Candidate · Junho 2026*
