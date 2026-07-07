"""
backtest_estatico.py — Etapa 2 de 3
=====================================
Experimento estático sem look-ahead bias.

Metodologia:
  - Calibração IS : 2019-06-01 → 2022-05-31  (~756 pregões)
    * Otimizador de máximo Sharpe calibrado APENAS com dados anteriores a junho/2022
    * CDI: média geométrica anualizada do período IS (BCB SGS 11)
    * Bounds: α=0.5 × w_analista ≤ w_i ≤ β=2.0 × w_analista

  - Teste OOS     : 2022-06-01 → 2026-06-28  (~1.016 pregões)
    * Buy-and-hold estático — pesos NÃO são rebalanceados
    * Compara pesos do analista (junho/2022) vs pesos otimizados (calibrados no IS)
    * CDI: média geométrica anualizada do período OOS (BCB SGS 11)

  - Sharpe (CORRIGIDO):
    * R_anual = (1 + R_acum)^(252 / N_dias) - 1   ← composto, não linear
    * vol = std(log_ret_diários) × √252
    * Sharpe = (R_anual - CDI_anual) / vol

  - Pesos das carteiras verificados nos PDFs originais (junho/2022):
    * Casa A: BBAS3 20%, CPLE6 20%, EGIE3 20%, GRND3 10%, PETR4 20%, TIMS3 10%
    * Casa B: todos com 10% — ALUP11, BBAS3, BBDC4, CYRE3, ENGI11,
                              ITUB4, SBSP3, TRPL4, VALE3, VIVT3

Saída:
  planilhas/backtest_2022.xlsx
    Aba "Comparacao"   — métricas IS e OOS para analista e otimizado
    Aba "Pesos"        — comparação de pesos analista vs otimizado
    Aba "Benchmarks"   — BOVA11, DIVO11 e CDI no período OOS
    Aba "ValorDiario"  — série diária de valor normalizado (base 1,0)

Execute:
  python backtest_estatico.py
"""

import os
import warnings
import numpy as np
import pandas as pd
from scipy.optimize import minimize

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
# Carteiras verificadas — PDFs originais — junho de 2022
# ---------------------------------------------------------------------------
CARTEIRAS = {
    "Casa A — Dividendos": {
        "tickers": ["BBAS3", "CPLE6", "EGIE3", "GRND3", "PETR4", "TIMS3"],
        "pesos":   [0.20,    0.20,    0.20,    0.10,    0.20,    0.10],
        # Fonte: Carteira Top Dividendos — 1º de junho de 2022 (PDF verificado)
        # Nota: VIVT3 substituída por TIMS3 nesta edição.
    },
    "Casa B — Dividendos": {
        "tickers": ["ALUP11", "BBAS3", "BBDC4", "CYRE3", "ENGI11",
                    "ITUB4",  "SBSP3", "TRPL4", "VALE3", "VIVT3"],
        "pesos":   [0.10,     0.10,    0.10,    0.10,    0.10,
                    0.10,     0.10,    0.10,    0.10,    0.10],
        # Fonte: PDF verificado
        # Mudanças em junho/2022: saíram AGRO3 e EGIE3; entraram CYRE3 e SBSP3.
    },
}

BENCHMARKS = ["BOVA11", "DIVO11"]

# Parâmetros
DATA_IS_INI  = "2019-06-01"
DATA_IS_FIM  = "2022-05-31"
DATA_OOS_INI = "2022-06-01"
DATA_OOS_FIM = "2026-06-28"

ALPHA = 0.50   # limite inferior: 50% do peso do analista
BETA  = 2.00   # limite superior: 200% do peso do analista

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DADOS_DIR  = os.path.join(BASE_DIR, "dados")
OUTPUT_DIR = os.path.join(BASE_DIR, "planilhas")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# Funções auxiliares
# ---------------------------------------------------------------------------

def cdi_anualizado(cdi_diario_decimal: pd.Series) -> tuple[float, float]:
    """
    Recebe série de CDI diário em decimal (ex: 0.000499).
    Retorna (cdi_acumulado, cdi_anual_geometrico).
    """
    acum = float((1 + cdi_diario_decimal).prod() - 1)
    n_anos = len(cdi_diario_decimal) / 252
    anual  = float((1 + acum) ** (1 / n_anos) - 1)
    return acum, anual


def valor_portfolio(precos_periodo: pd.DataFrame, pesos: dict) -> pd.Series:
    """
    Calcula a série de valor de R$ 1 investido em buy-and-hold estático.
    Os pesos não são rebalanceados — cada ativo cresce livremente.
    pesos: {ticker: peso_inicial, ...}
    """
    tickers = list(pesos.keys())
    w = np.array([pesos[t] for t in tickers])
    p = precos_periodo[tickers].copy()
    # Preço relativo ao primeiro pregão disponível
    p_rel = p.div(p.iloc[0])
    # Valor do portfólio = soma ponderada dos valores relativos
    v = (p_rel * w).sum(axis=1)
    v.name = "valor"
    return v


def metricas(valor_serie: pd.Series, cdi_anual: float) -> dict:
    """
    Calcula métricas de performance a partir de uma série de valor (base 1,0).
    Sharpe calculado com retorno anualizado composto — não linear.
    """
    n_dias = len(valor_serie)

    # Retorno acumulado
    r_acum = float(valor_serie.iloc[-1] / valor_serie.iloc[0] - 1)

    # Retorno anualizado COMPOSTO (corrigido)
    r_anual = float((1 + r_acum) ** (252 / n_dias) - 1)

    # Volatilidade anualizada via log-retornos
    log_ret = np.log(valor_serie / valor_serie.shift(1)).dropna()
    vol_anual = float(log_ret.std() * np.sqrt(252))

    # Sharpe (usa retorno composto anualizado)
    sharpe = (r_anual - cdi_anual) / vol_anual if vol_anual > 1e-8 else np.nan

    # Maximum Drawdown
    rolling_max = valor_serie.cummax()
    drawdown    = (valor_serie / rolling_max - 1)
    mdd         = float(drawdown.min())

    return {
        "N pregões":          n_dias,
        "Retorno Acumulado":  r_acum,
        "Retorno Anualizado": r_anual,
        "Volatilidade Anual": vol_anual,
        "Sharpe Ratio":       sharpe,
        "Max Drawdown":       mdd,
    }


def otimizar_max_sharpe(
    log_ret_is: pd.DataFrame,
    pesos_rec: np.ndarray,
    cdi_is_anual: float,
    alpha: float = 0.5,
    beta:  float = 2.0,
) -> np.ndarray:
    """
    Otimizador de máximo Sharpe calibrado no IS.
    Bounds: [alpha × w_rec, beta × w_rec]
    Restrição: sum(w) = 1
    Solver: SLSQP com múltiplos pontos de partida.
    """
    mu    = log_ret_is.mean().values * 252
    sigma = log_ret_is.cov().values  * 252
    n     = len(mu)

    def neg_sharpe(w: np.ndarray) -> float:
        ret = float(w @ mu)
        vol = float(np.sqrt(w @ sigma @ w))
        if vol < 1e-10:
            return np.inf
        return -(ret - cdi_is_anual) / vol

    bounds = []
    for i in range(n):
        lb = max(0.0,  alpha * pesos_rec[i])
        ub = min(1.0,  beta  * pesos_rec[i])
        bounds.append((lb, ub))

    constraints = [{"type": "eq", "fun": lambda w: np.sum(w) - 1.0}]

    # Pontos de partida: pesos do analista + centro dos bounds
    centros = np.array([(b[0] + b[1]) / 2 for b in bounds])
    centros /= centros.sum()

    pontos_partida = [pesos_rec.copy(), centros]

    best_result = None
    best_val    = np.inf

    for w0 in pontos_partida:
        # Clipar e normalizar para garantir viabilidade inicial
        w0c = np.clip(w0, [b[0] for b in bounds], [b[1] for b in bounds])
        w0c = w0c / w0c.sum()

        res = minimize(
            neg_sharpe, w0c,
            method="SLSQP",
            bounds=bounds,
            constraints=constraints,
            options={"ftol": 1e-12, "maxiter": 2000},
        )
        if res.fun < best_val:
            best_val    = res.fun
            best_result = res

    if not best_result.success:
        print(f"    Aviso: SLSQP não convergiu — {best_result.message}")

    return best_result.x


# ---------------------------------------------------------------------------
# Carregar dados
# ---------------------------------------------------------------------------
print("=" * 65)
print("backtest_estatico.py — Experimento sem look-ahead bias")
print("=" * 65)

print("\nCarregando dados...")
precos = pd.read_parquet(os.path.join(DADOS_DIR, "precos.parquet"))
cdi_df = pd.read_parquet(os.path.join(DADOS_DIR, "cdi.parquet"))

# CDI em decimal (série BCB SGS 11 está em % ao dia)
cdi_decimal = cdi_df["CDI"] / 100.0

# CDI para cada período
cdi_is_acum,  cdi_is_anual  = cdi_anualizado(cdi_decimal.loc[DATA_IS_INI:DATA_IS_FIM])
cdi_oos_acum, cdi_oos_anual = cdi_anualizado(cdi_decimal.loc[DATA_OOS_INI:DATA_OOS_FIM])

print(f"  CDI IS  ({DATA_IS_INI}→{DATA_IS_FIM}): {cdi_is_acum:.2%} acum | {cdi_is_anual:.2%} a.a.")
print(f"  CDI OOS ({DATA_OOS_INI}→{DATA_OOS_FIM}): {cdi_oos_acum:.2%} acum | {cdi_oos_anual:.2%} a.a.")

precos_is  = precos.loc[DATA_IS_INI:DATA_IS_FIM]
precos_oos = precos.loc[DATA_OOS_INI:DATA_OOS_FIM]

# ---------------------------------------------------------------------------
# Processamento por carteira
# ---------------------------------------------------------------------------
lista_resultados  = []
lista_pesos       = []
series_valor_oos  = {}   # para aba "ValorDiario"

for nome, cfg in CARTEIRAS.items():
    print(f"\n{'─'*65}")
    print(f"Carteira: {nome}")

    tickers = list(cfg["tickers"])
    w_rec   = np.array(cfg["pesos"], dtype=float)

    # ---- Remover tickers sem dados e redistribuir pesos ----
    tickers_sem_dados = [
        t for t in tickers
        if t not in precos.columns or precos[t].dropna().empty
    ]
    if tickers_sem_dados:
        print(f"  Aviso: {tickers_sem_dados} sem dados — removidos e pesos redistribuídos.")
        idx_ok  = [i for i, t in enumerate(tickers) if t not in tickers_sem_dados]
        tickers = [tickers[i] for i in idx_ok]
        w_rec   = w_rec[idx_ok]
        w_rec   = w_rec / w_rec.sum()   # normalizar para soma = 1
        print(f"  Carteira ajustada: {dict(zip(tickers, [f'{w:.1%}' for w in w_rec]))}")

    # ---- IS: verificar dados e calcular log-retornos ----
    p_is = precos_is[tickers].dropna(how="any")
    if p_is.empty or len(p_is) < 60:
        print(f"  Erro: dados IS insuficientes ({len(p_is)} linhas). Pulando.")
        continue

    log_ret_is = np.log(p_is / p_is.shift(1)).dropna()
    print(f"  IS: {p_is.index[0].date()} → {p_is.index[-1].date()} | {len(log_ret_is)} pregões")

    # ---- Otimização ----
    print("  Calibrando Max Sharpe (IS)...")
    w_opt = otimizar_max_sharpe(log_ret_is, w_rec, cdi_is_anual, ALPHA, BETA)

    print(f"\n  {'Ticker':<8}  {'Analista':>10}  {'Otimizado':>10}  {'Delta':>10}")
    print(f"  {'─'*44}")
    for t, wr, wo in zip(tickers, w_rec, w_opt):
        print(f"  {t:<8}  {wr:>10.1%}  {wo:>10.1%}  {wo-wr:>+10.1%}")
    print(f"  {'TOTAL':<8}  {sum(w_rec):>10.1%}  {sum(w_opt):>10.1%}")

    # ---- Métricas IS (hipotéticas — só para comparação de overfitting) ----
    v_rec_is = valor_portfolio(precos_is[tickers], dict(zip(tickers, w_rec)))
    v_opt_is = valor_portfolio(precos_is[tickers], dict(zip(tickers, w_opt)))
    m_rec_is = metricas(v_rec_is, cdi_is_anual)
    m_opt_is = metricas(v_opt_is, cdi_is_anual)

    # ---- OOS: tratar tickers com histórico truncado ----
    p_oos = precos_oos[tickers].copy()

    # Informar gaps de dados (ex: CPLE6 após privatização da Copel em 2023)
    for t in tickers:
        ultimo = p_oos[t].last_valid_index()
        if ultimo is not None and ultimo < p_oos.index[-1]:
            dias_faltando = (p_oos.index[-1] - ultimo).days
            print(f"\n  Aviso: {t} sem dados após {ultimo.date()} "
                  f"({dias_faltando} dias até o fim do período).")
            print(f"         Análise OOS do portfólio {nome} encerrada em {ultimo.date()}.")

    # Usar apenas pregões com todos os tickers disponíveis (buy-and-hold puro)
    p_oos_valido = p_oos.dropna(how="any")

    if p_oos_valido.empty or len(p_oos_valido) < 20:
        print(f"  Erro: dados OOS insuficientes. Pulando.")
        continue

    print(f"\n  OOS: {p_oos_valido.index[0].date()} → {p_oos_valido.index[-1].date()} | {len(p_oos_valido)} pregões")

    # ---- Métricas OOS ----
    v_rec_oos = valor_portfolio(p_oos_valido, dict(zip(tickers, w_rec)))
    v_opt_oos = valor_portfolio(p_oos_valido, dict(zip(tickers, w_opt)))
    m_rec_oos = metricas(v_rec_oos, cdi_oos_anual)
    m_opt_oos = metricas(v_opt_oos, cdi_oos_anual)

    # Impressão comparativa
    print(f"\n  {'Métrica':<25}  {'IS Analista':>11}  {'IS Otim.':>10}  {'OOS Analista':>13}  {'OOS Otim.':>10}")
    print(f"  {'─'*75}")
    for k in ["Retorno Acumulado", "Retorno Anualizado", "Volatilidade Anual", "Sharpe Ratio", "Max Drawdown"]:
        ri = m_rec_is.get(k, np.nan)
        oi = m_opt_is.get(k, np.nan)
        ro = m_rec_oos.get(k, np.nan)
        oo = m_opt_oos.get(k, np.nan)
        if k == "Sharpe Ratio":
            print(f"  {k:<25}  {ri:>11.2f}  {oi:>10.2f}  {ro:>13.2f}  {oo:>10.2f}")
        else:
            print(f"  {k:<25}  {ri:>11.2%}  {oi:>10.2%}  {ro:>13.2%}  {oo:>10.2%}")

    # ---- Acumular resultados ----
    for periodo, tipo, m in [
        ("IS (calibração)", "Analista (hipotético)", m_rec_is),
        ("IS (calibração)", "Otimizado",   m_opt_is),
        ("OOS (teste)",     "Analista",               m_rec_oos),
        ("OOS (teste)",     "Otimizado",    m_opt_oos),
    ]:
        lista_resultados.append({
            "Carteira": nome,
            "Período":  periodo,
            "Tipo":     tipo,
            **m,
        })

    for t, wr, wo in zip(tickers, w_rec, w_opt):
        lista_pesos.append({
            "Carteira":       nome,
            "Ticker":         t,
            "Peso Analista":  wr,
            "Peso Otimizado": wo,
            "Diferença":      wo - wr,
        })

    # Salvar séries diárias OOS (normalizadas em 1,0)
    series_valor_oos[f"{nome} — Analista"]   = v_rec_oos
    series_valor_oos[f"{nome} — Otimizado"]  = v_opt_oos

# ---------------------------------------------------------------------------
# Benchmarks OOS
# ---------------------------------------------------------------------------
print(f"\n{'─'*65}")
print("Benchmarks OOS:")

lista_bench = []

for bench in BENCHMARKS:
    if bench not in precos_oos.columns:
        print(f"  {bench}: não disponível")
        continue
    p_b = precos_oos[[bench]].dropna()
    if p_b.empty:
        continue
    v_b = valor_portfolio(p_b, {bench: 1.0})
    m_b = metricas(v_b, cdi_oos_anual)
    print(f"  {bench:<8}: Ret={m_b['Retorno Acumulado']:>8.2%}  "
          f"Anual={m_b['Retorno Anualizado']:>8.2%}  "
          f"Sharpe={m_b['Sharpe Ratio']:>5.2f}  "
          f"MDD={m_b['Max Drawdown']:>8.2%}")
    lista_bench.append({"Benchmark": bench, **m_b})
    series_valor_oos[bench] = v_b

# CDI como benchmark (sem MDD nem vol — referência de custo de oportunidade)
print(f"  {'CDI':<8}: Ret={cdi_oos_acum:>8.2%}  Anual={cdi_oos_anual:>8.2%}")
lista_bench.append({
    "Benchmark":          "CDI (SGS 11)",
    "N pregões":          len(cdi_decimal.loc[DATA_OOS_INI:DATA_OOS_FIM]),
    "Retorno Acumulado":  cdi_oos_acum,
    "Retorno Anualizado": cdi_oos_anual,
    "Volatilidade Anual": 0.0,
    "Sharpe Ratio":       0.0,
    "Max Drawdown":       0.0,
})

# Série CDI diária (valor acumulado)
cdi_oos_serie = cdi_decimal.loc[DATA_OOS_INI:DATA_OOS_FIM]
cdi_valor     = (1 + cdi_oos_serie).cumprod()
cdi_valor     = cdi_valor / cdi_valor.iloc[0]
cdi_valor.name = "CDI"
series_valor_oos["CDI"] = cdi_valor

# ---------------------------------------------------------------------------
# Salvar Excel
# ---------------------------------------------------------------------------
print(f"\n{'─'*65}")
print("Salvando Excel...")

df_resultados = pd.DataFrame(lista_resultados)
df_pesos      = pd.DataFrame(lista_pesos)
df_bench      = pd.DataFrame(lista_bench)
df_valor      = pd.DataFrame(series_valor_oos)

# Formatar colunas de porcentagem para legibilidade
pct_cols = ["Retorno Acumulado", "Retorno Anualizado", "Volatilidade Anual", "Max Drawdown"]

path_excel = os.path.join(OUTPUT_DIR, "backtest_2022.xlsx")

with pd.ExcelWriter(path_excel, engine="openpyxl") as writer:
    df_resultados.to_excel(writer, sheet_name="Comparacao", index=False)
    df_pesos.to_excel(writer,      sheet_name="Pesos",      index=False)
    df_bench.to_excel(writer,      sheet_name="Benchmarks", index=False)
    df_valor.to_excel(writer,      sheet_name="ValorDiario")

    # Formatação automática de colunas
    for sheet_name, df in [
        ("Comparacao", df_resultados),
        ("Benchmarks", df_bench),
    ]:
        ws = writer.sheets[sheet_name]
        for col in ws.iter_cols(1, ws.max_column, 1, 1):
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = 22

print(f"Excel salvo: {path_excel}")
print("\nEtapa 2 concluída. Execute: python monte_carlo.py")
print("=" * 65)

# Exportar variáveis para uso pelo monte_carlo.py
__all__ = [
    "CARTEIRAS", "DATA_IS_INI", "DATA_IS_FIM", "DATA_OOS_INI", "DATA_OOS_FIM",
    "ALPHA", "BETA", "DADOS_DIR", "OUTPUT_DIR",
    "cdi_is_anual", "cdi_oos_anual",
    "valor_portfolio", "metricas", "otimizar_max_sharpe",
]
