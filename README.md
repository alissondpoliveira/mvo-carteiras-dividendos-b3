# Otimização de Markowitz — Carteiras de Dividendos B3

> Apliquei otimização de Markowitz (Max Sharpe) às carteiras de dividendos publicadas pelas principais casas de análise brasileiras em junho de 2026. O modelo perdeu. O motivo importa mais que o resultado.

**Artigo completo:** [LinkedIn — *ETFs no Brasil em 2025: crescimento real, causas ainda indetermináveis*](#) *(link será atualizado após publicação)*

---

## Estrutura do Repositório

```
markowitz-carteiras-dividendos/
├── scripts/
│   ├── analisar_carteiras.py        # Análise das 5 carteiras de mercado + otimizador
│   ├── backtest_benchmark.py        # Backtest OOS (split 50/50) + DIVO11 e IBOV
│   ├── monte_carlo.py               # Bootstrap 500 splits + Monte Carlo 5.000 trajetórias
│   └── analisar_minha_carteira.py   # Para o leitor: analise sua própria carteira
├── planilhas/
│   ├── comparacao_carteiras.xlsx    # Output: comparação das 5 carteiras
│   ├── backtest_benchmark.xlsx      # Output: backtest OOS + benchmarks
│   ├── monte_carlo.xlsx             # Output: bootstrap + Monte Carlo
│   └── minha_carteira.xlsx          # Template de entrada: ticker + quantidade
├── dados/
│   ├── precos_b3.csv                # Série histórica de preços (não inclusa — veja abaixo)
│   ├── README_dados.md              # Instruções sobre formato do CSV
│   └── Carteiras/
│       └── carteiras_jun2026.json   # Metadados das 5 carteiras analisadas
├── README.md
├── requirements.txt
└── .gitignore
```

---

## Para Analisar Sua Própria Carteira

1. Preencha `planilhas/minha_carteira.xlsx` com seus tickers e **quantidade de papéis**
2. Forneça o CSV de preços em `dados/precos_b3.csv` (formato abaixo)
3. Execute:

```bash
python scripts/analisar_minha_carteira.py
```

4. Abra `resultado_carteira.xlsx` — dashboard com composição, métricas de risco e sugestão de pesos otimizados

O script calcula automaticamente os pesos pela última cotação disponível no CSV. Não é necessário informar percentuais.

---

## Dados de Preços

Os scripts precisam de `dados/precos_b3.csv` com preços de fechamento ajustados. **O arquivo não está incluído neste repositório.**

### Formato esperado

```
data,PETR4,VALE3,ITUB4,...,IBOV,BOVA11,DIVO11
2025-06-03,38.42,61.10,33.20,...,127834,127.45,143.20
2025-06-04,38.91,60.87,33.55,...,128102,128.01,143.85
```

- Primeira coluna: `data` em `YYYY-MM-DD`
- Colunas: tickers B3 em maiúsculas
- Coluna `IBOV` obrigatória para Beta e Alpha de Jensen
- Coluna `DIVO11` opcional para Information Ratio vs. benchmark de dividendos
- Preços: fechamento ajustado por proventos e eventos corporativos

### Como obter dados com yfinance

```python
import yfinance as yf, pandas as pd

tickers = ["PETR4.SA","VALE3.SA","ITUB4.SA","^BVSP"]
df = yf.download(tickers, start="2025-06-01", end="2026-06-01")["Close"]
df.columns = [c.replace(".SA","").replace("^BVSP","IBOV") for c in df.columns]
df.to_csv("dados/precos_b3.csv", index_label="data")
```

---

## Carteiras Analisadas

Carteiras de dividendos publicadas em junho de 2026. Fonte: Money Times (01–04/06/2026) e Santander PDF.

| Casa | Ativos | Destaques |
|---|---|---|
| Casa A | 11 | CPLE3 15%, ITUB4 15%, VALE3 12,5% |
| Casa B | 12 | Pesos entre 5% e 10% |
| Itaú BBA | 5 | Pesos iguais 20%: AXIA3, ALOS3, BBDC4, VALE3, PETR4 |
| Santander | 10 | Pesos iguais 10% |
| BB Investimentos | 10 | Pesos iguais 10% |

Detalhes completos em `dados/Carteiras/carteiras_jun2026.json`.

---

## Modelo Matemático

### Definições

Dado um universo de $n$ ativos com log-retornos diários $r_{i,t}$, define-se:

$$\mu_i = \bar{r}_i \times 252 \qquad \text{(retorno esperado anualizado)}$$

$$\Sigma = \hat{S} \times 252 \qquad \text{(matriz de covariância anualizada)}$$

onde $\hat{S}$ é o estimador amostral com correção de Bessel ($ddof = 1$).

O retorno e a volatilidade de uma carteira com vetor de pesos $w \in \mathbb{R}^n$ são:

$$R_p(w) = w^\top \mu$$

$$\sigma_p(w) = \sqrt{w^\top \Sigma\, w}$$

### Problema de Otimização

$$\max_{w \in \mathbb{R}^n} \; S(w) = \frac{R_p(w) - R_f}{\sigma_p(w)}$$

sujeito a:

$$\sum_{i=1}^{n} w_i = 1$$

$$\alpha \cdot w_i^{\text{rec}} \leq w_i \leq \beta \cdot w_i^{\text{rec}}, \quad \forall\, i = 1, \ldots, n$$

**Parâmetros:** $\alpha = 0{,}5$ (mínimo 50% do peso recomendado) $\quad \beta = 2{,}0$ (máximo 200%) $\quad R_f = 14{,}40\%$ a.a.

**Solver:** `scipy.optimize.minimize` com método SLSQP (Sequential Least Squares Programming).

### Métricas Calculadas

**Índice de Sharpe**

$$S = \frac{R_p - R_f}{\sigma_p}$$

**Índice de Sortino**

$$So = \frac{R_p - R_f}{\sigma_d}, \qquad \sigma_d = \sqrt{\frac{252}{T} \sum_{t:\, r_t < R_f/252} \!\!\left(r_t - \frac{R_f}{252}\right)^{\!2}}$$

**Beta e Alpha de Jensen** (benchmark IBOV)

$$\beta_p = \frac{\text{Cov}(r_p,\, r_m)}{\text{Var}(r_m)}, \qquad \alpha_J = R_p - \bigl[R_f + \beta_p\,(R_m - R_f)\bigr]$$

**Índice de Treynor**

$$T = \frac{R_p - R_f}{\beta_p}$$

**Value at Risk e CVaR (histórico)**

$$\text{VaR}_{95\%} = -\text{Percentil}_{5\%}(r_p)$$

$$\text{CVaR}_{95\%} = -\,\mathbb{E}\!\left[r_p \mid r_p \leq -\text{VaR}_{95\%}\right]$$

**Maximum Drawdown**

$$\text{MDD} = \min_{t} \frac{V_t - \max_{s \leq t} V_s}{\max_{s \leq t} V_s}, \qquad V_t = \exp\!\left(\sum_{\tau=1}^{t} r_{p,\tau}\right)$$

**Information Ratio** (benchmark DIVO11)

$$\text{IR} = \frac{\overline{r_p - r_b}}{\text{TE}}, \qquad \text{TE} = \sigma(r_p - r_b) \times \sqrt{252}$$

---

## Resultados Principais

### Full-Sample — Pesos Recomendados (jun/2025 – mai/2026)

| Carteira | Retorno Anual | Sharpe | IR vs DIVO11 | $\alpha_J$ |
|---|---|---|---|---|
| **DIVO11** (proxy IDIV) | 21,3% | 0,44 | — | — |
| **IBOV** | 23,9% | 0,57 | — | — |
| Casa A | 31,6% | 0,98 | 1,52 | +7,6% |
| Casa B | 28,7% | 0,77 | 0,98 | +4,4% |
| **Itaú BBA Div.** | **35,2%** | **1,16** | **1,63** | **+11,4%** |
| Santander Div. | 33,9% | 1,03 | 1,59 | +9,3% |
| BB Dividendos | 26,6% | 0,74 | 0,92 | +3,4% |

### Bootstrap Backtest — 500 Splits Aleatórios (Sharpe OOS mediano)

| Carteira | Rec OOS | Opt IS | Opt OOS | $\Delta$ mediana | Agrega? |
|---|---|---|---|---|---|
| Casa A | 0,911 | 2,153 | 0,591 | −0,284 | ❌ Não |
| Casa B | 0,250 | 2,537 | 0,250 | +0,043 | ⚠ Inconclusivo |
| Itaú BBA | 0,470 | 3,007 | 0,273 | −0,287 | ❌ Não |
| Santander | 0,496 | 2,572 | 0,377 | −0,163 | ❌ Não |
| BB Investimentos | 0,369 | 1,909 | 0,164 | −0,298 | ❌ Não |

**Achado central:** Sharpe IS médio $\approx 2{,}5$ → Sharpe OOS mediano $\approx 0{,}3$. Degradação de ~88%. O otimizador clássico com janela curta não agrega valor sistematicamente — maximiza o Sharpe verdadeiro junto com o erro de estimação dos parâmetros.

---

## Limitações do Modelo

**Estimation error maximization.** Com ~100 pregões de dados e 10 ativos, o erro padrão do retorno esperado anualizado é da ordem de $\pm 3$ p.p. por ativo. O otimizador maximiza, junto com o Sharpe verdadeiro, o ruído de estimação. O resultado é Sharpe inflado in-sample e degradado out-of-sample.

Alternativas mais robustas (não implementadas neste repositório):

- **Ledoit-Wolf shrinkage** — regularização da matriz $\Sigma$ para reduzir o impacto de correlações estimadas com ruído
- **Black-Litterman** — combina prior de mercado (CAPM) com views do analista via atualização bayesiana
- **Janela mais longa** — mínimo recomendado de 3–5 anos para estimação estável de $\mu$ e $\Sigma$

---

## Como Reproduzir

```bash
git clone https://github.com/alissondpoliveira/Markowitz-carteiras-dividendos.git
cd markowitz-carteiras-dividendos
pip install -r requirements.txt

# Coloque precos_b3.csv em dados/

python scripts/analisar_carteiras.py        # → planilhas/comparacao_carteiras.xlsx
python scripts/backtest_benchmark.py        # → planilhas/backtest_benchmark.xlsx
python scripts/monte_carlo.py               # → planilhas/monte_carlo.xlsx
python scripts/analisar_minha_carteira.py   # → resultado_carteira.xlsx
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
