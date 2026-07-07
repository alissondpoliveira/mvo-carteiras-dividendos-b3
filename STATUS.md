# Otimização de Portfolio — Status do Projeto

> **Última atualização:** 07/06/2026  
> **Modelo:** Claude Sonnet 4.6 · Cowork Mode  
> **Para continuar:** abra esta pasta no Cowork e mencione este arquivo. Todas as análises são reproduzíveis.

---

## Objetivo

Aplicar otimização de máximo Sharpe às carteiras de **dividendos** recomendadas por casas de análise brasileiras em junho de 2022, comparar via métricas quantitativas e avaliar se o otimizador gera valor real via bootstrap e Monte Carlo.

---

## Dados

- **Arquivo:** `precos_b3.csv`
- **Período:** 03/06/2025 → 29/05/2026 — **248 pregões**
- **Colunas:** 228 (ações B3 + índices)
- **Benchmarks disponíveis:** `IBOV`, `BOVA11`, `DIVO11`
- **IDIV:** não está no CSV. Usar **DIVO11** como proxy (ETF que replica o IDIV)
- **Taxa livre de risco:** RF_AA = 14,40% a.a. → RF_DIÁRIO = ln(1,144)/252

---

## Carteiras (Dividendos — Junho 2026)

Fonte: Money Times 01–04/06/2026 + Santander PDF.  
Detalhes completos em `Carteiras/carteiras_jun2026.json`.

| Casa | Nº Ativos | Principais Posições |
|---|---|---|
| **Casa A** | 11 | CPLE3 15%, ITUB4 15%, VALE3 12.5%, AXIA3 10% |
| **Casa B** | 12 | PETR4 10%, ITUB4 10%, BBDC4 10%, CURY3 10% |
| **Casa C** | 5 | Pesos iguais 20% |
| **Casa D** | 10 | Pesos iguais 10% |
| **Casa E** | 10 | Pesos iguais 10% |

**Substituição:** AXIA6 → AXIA3 (mesma empresa Axia Energia, PN→ON).

---

## Modelo de Otimização

```
Max  Sharpe(w) = [R_p(w) - RF] / σ_p(w)

s.a. Σ wᵢ = 1
     α · wᵢ_rec ≤ wᵢ ≤ β · wᵢ_rec   ∀ i
     (α = 0.5,  β = 2.0)

R_p  = w' μ         (μ anualizado: E[log-ret] × 252)
σ_p  = √(w' Σ w)    (Σ anualizada: Cov × 252)
```

Solver: `scipy.optimize.minimize` com método SLSQP.

---

## Métricas Calculadas

| Grupo | Indicadores |
|---|---|
| Retorno/Risco | R_anual, Vol_anual, Sharpe, Sortino, Treynor |
| Risco de Cauda | VaR 95/99% (histórico + paramétrico), CVaR 95/99%, MDD |
| Benchmark | Beta vs IBOV, Alpha de Jensen, IR vs DIVO11 |
| Otimização | Pesos recomendados vs. otimizados, Δ Sharpe |

---

## Arquivos do Projeto

| Arquivo | Descrição | Status |
|---|---|---|
| `precos_b3.csv` | Dados de preços — somente leitura | ✅ |
| `Carteiras/carteiras_jun2026.json` | Metadados das 5 carteiras | ✅ |
| `analisar_carteiras.py` | Análise principal + otimização + Excel | ✅ |
| `comparacao_carteiras.xlsx` | Excel com Comparação, Pesos, Disclaimers | ✅ |
| `backtest_benchmark.py` | Backtest OOS (split 50/50) + DIVO11 benchmark | ✅ |
| `backtest_benchmark.xlsx` | Excel com Backtest OOS + Benchmark DIVO11 | ✅ |
| `monte_carlo.py` | Bootstrap (500 splits) + Monte Carlo forward (5000 trajetórias) | ✅ |
| `monte_carlo.xlsx` | Excel com Bootstrap Backtest + Monte Carlo + Distribuição Sharpe | ✅ |
| `analisar_minha_carteira.py` | Script para o leitor analisar sua própria carteira | ✅ |
| `minha_carteira.xlsx` | Template Excel de entrada para o leitor (tickers + pesos) | ✅ |
| `README.md` | README do repositório GitHub | ✅ |
| `requirements.txt` | Dependências Python para o GitHub | ✅ |
| `.gitignore` | Ignora dados proprietários e outputs gerados | ✅ |
| `dados/README_dados.md` | Instruções sobre formato do CSV de preços | ✅ |
| `artigo-markowitz-rascunho-v1.md` | Rascunho do artigo LinkedIn (v1) | ✅ |
| `STATUS.md` | Este arquivo — resumo do projeto para continuidade | ✅ |

---

## Resultados-Chave

### Análise Full-Sample (período completo, pesos recomendados)

| Carteira | Retorno Anual | Sharpe | IR vs DIVO11 | Alpha Jensen |
|---|---|---|---|---|
| **DIVO11** (benchmark) | 21.3% | 0.44 | — | — |
| **IBOV** | 23.9% | 0.57 | — | — |
| Casa A | 31.6% | 0.98 | 1.52 | +7.6% |
| Casa B | 28.7% | 0.77 | 0.98 | +4.4% |
| **Casa C** | **35.2%** | **1.16** | **1.63** | **+11.4%** |
| Casa D | 33.9% | 1.03 | 1.59 | +9.3% |
| Casa E | 26.6% | 0.74 | 0.92 | +3.4% |

→ Todas as 5 carteiras superaram DIVO11 e IBOV no período amostral.

---

### Bootstrap Backtest (500 splits aleatórios — Sharpe OOS mediano)

| Carteira | Rec OOS | Opt IS | Opt OOS | Δ mediana | Opt>Rec | Veredicto |
|---|---|---|---|---|---|---|
| Casa A | 0.911 | 2.153 | 0.591 | −0.284 | 5.6% | ❌ NÃO AGREGA |
| Casa B | 0.250 | 2.537 | 0.250 | +0.043 | 57.6% | ⚠️ INCONCLUSIVO |
| Casa C | 0.470 | 3.007 | 0.273 | −0.287 | 8.8% | ❌ NÃO AGREGA |
| Casa D | 0.496 | 2.572 | 0.377 | −0.163 | 13.6% | ❌ NÃO AGREGA |
| Casa E | 0.369 | 1.909 | 0.164 | −0.298 | 5.2% | ❌ NÃO AGREGA |

**Achado central:** O otimizador com janela curta (~6 meses de treino) não agrega valor sistematicamente. O Sharpe IS inflado (~2-3×) não se transfere para OOS. Os pesos "ingênuos" das casas resistem melhor — evidência de que a seleção qualitativa dos analistas funciona como regularização implícita. Apenas BTG é inconclusivo.

---

### Monte Carlo Forward (5000 trajetórias × 252 pregões)

| Carteira | P5 | Mediana | P95 | P(>RF=14.4%) | P(>DIVO11) | P(MDD>20%) |
|---|---|---|---|---|---|---|
| Casa A | 4.1% | 36.9% | 83.3% | 86% | 72% | 3% |
| Casa B | −1.7% | 32.8% | 80.0% | 79% | 65% | 8% |
| **Casa C** | **6.0%** | **42.0%** | **92.0%** | **88%** | **77%** | **3%** |
| Casa D | 3.1% | 40.2% | 90.6% | 86% | 74% | 5% |
| Casa E | −0.4% | 30.3% | 70.6% | 79% | 62% | 3% |

Parâmetros μ e Σ estimados no histórico completo. Assume normalidade (fat tails reais podem piorar P5).

---

## Fenômeno Observado: Estimation Error Maximization

O otimizador clássico com janela curta maximiza junto com o Sharpe o **erro de estimação** dos parâmetros μ e Σ. Com apenas ~80–170 pregões de treino:

- Sharpe IS médio: **~2.5** (inflado pelo overfitting)
- Sharpe OOS mediano: **~0.3** (degradação de ~88%)
- Pesos recomendados OOS mediano: **~0.5** (mais estável)

Possíveis mitigações (não implementadas ainda):
- **Shrinkage Ledoit-Wolf** na matriz de covariância
- **Black-Litterman** para combinar prior de mercado com views
- **Regularização L2** (adição de termo de penalidade nos pesos)
- **Janela mais longa** (mínimo recomendado: 3–5 anos)

---

## Próximos Passos — GitHub

**Para publicar no GitHub:**
1. Criar repositório público: `markowitz-carteiras-dividendos`
2. Copiar os arquivos: `analisar_carteiras.py`, `backtest_benchmark.py`, `monte_carlo.py`, `analisar_minha_carteira.py`, `minha_carteira.xlsx`, `README.md`, `requirements.txt`, `.gitignore`, `dados/README_dados.md`, `Carteiras/carteiras_jun2026.json`
3. **Não subir:** `precos_b3.csv`, arquivos `.xlsx` de output (já estão no `.gitignore`)
4. Atualizar `alissondpoliveira` nas referências do artigo e dos scripts com o handle do GitHub

**Pendências do artigo:**
- [ ] Revisar e aprovar `artigo-markowitz-rascunho-v1.md`
- [ ] Atualizar `alissondpoliveira` → handle real do GitHub após criar o repo
- [ ] Gerar `artigo-markowitz-rascunho-v1.docx` atualizado (após aprovação do .md)
- [ ] Verificação de compliance antes de publicar

## Melhorias Futuras Possíveis

- [ ] Aplicar Ledoit-Wolf shrinkage na matriz Σ e comparar degradação OOS
- [ ] Black-Litterman com views sobre dividend yield esperado
- [ ] Incluir dividendos recebidos no retorno (cotação ≠ total return)
- [ ] Exportar gráficos das trajetórias Monte Carlo (matplotlib → PNG → embed no Excel)
- [ ] Análise de correlação entre ativos por carteira (heatmap)
- [ ] Comparar com carteiras de outras casas

---

## Como Reproduzir

```bash
# Na pasta do projeto:
python analisar_carteiras.py    # → comparacao_carteiras.xlsx
python backtest_benchmark.py    # → backtest_benchmark.xlsx
python monte_carlo.py           # → monte_carlo.xlsx
```

Todos os scripts são independentes e usam `precos_b3.csv` como única fonte de dados.

---

## Disclaimers

> ⚠️ **Análise retroativa e hipotética. NÃO é recomendação de investimento.**  
> Desempenho passado não garante resultados futuros. Os modelos assumem premissas simplificadoras (normalidade dos retornos, liquidez, sem custos de transação). Utilize estas análises apenas como exercício acadêmico de pesquisa operacional e finanças quantitativas.
