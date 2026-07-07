"""
monte_carlo.py — Etapa 3 de 3
================================
Bootstrap histórico sobre retornos do período OOS.

Método:
  - Reamostrar com reposição os retornos diários reais do OOS (iid bootstrap)
  - 5.000 trajetórias de 252 pregões (horizonte de 1 ano prospectivo)
  - Sem contaminação de dados IS — reamostragem restrita ao OOS

Para cada portfólio (analista e otimizado, Casa A e Casa B):
  - Distribuição de: Retorno Acumulado, Volatilidade Anual, Sharpe, Max Drawdown
  - Estatísticas: p5, p25, p50, p75, p95, média

Nota: o bootstrap histórico assume que os retornos OOS observados constituem
      a distribuição preditiva para períodos futuros similares. Não assume
      normalidade, mas assume independência temporal — condição aproximada para
      retornos diários de renda variável.

Saída:
  planilhas/monte_carlo_2022.xlsx
    Aba "Percentis"   — percentis p5/p25/p50/p75/p95 por métrica
    Aba "Simulacoes"  — matriz completa de simulações (5000 × 4 métricas)

Execute:
  python monte_carlo.py
"""

import os
import sys
import warnings
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
# Configurações (espelhadas de backtest_estatico.py — sem import circular)
# ---------------------------------------------------------------------------
CARTEIRAS = {
    "Casa A — Dividendos": {
        "tickers": ["BBAS3", "CPLE6", "EGIE3", "GRND3", "PETR4", "TIMS3"],
        "pesos":   [0.20,    0.20,    0.20,    0.10,    0.20,    0.10],
    },
    "Casa B — Dividendos": {
        "tickers": ["ALUP11", "BBAS3", "BBDC4", "CYRE3", "ENGI11",
                    "ITUB4",  "SBSP3", "TRPL4", "VALE3", "VIVT3"],
        "pesos":   [0.10,     0.10,    0.10,    0.10,    0.10,
                    0.10,     0.10,    0.10,    0.10,    0.10],
    },
}

DATA_IS_INI  = "2019-06-01"
DATA_IS_FIM  = "2022-05-31"
DATA_OOS_INI = "2022-06-01"
DATA_OOS_FIM = "2026-06-28"

ALPHA     = 0.50
BETA      = 2.00
N_SIMS    = 5_000
HORIZONTE = 252      # pregões (≈ 1 ano)
SEED      = 42

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DADOS_DIR  = os.path.join(BASE_DIR, "dados")
OUTPUT_DIR = os.path.join(BASE_DIR, "planilhas")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# Funções auxiliares (copiadas de backtest_estatico para evitar dependência)
# ---------------------------------------------------------------------------

def cdi_anualizado(cdi_diario_decimal: pd.Series) -> tuple[float, float]:
    acum  = float((1 + cdi_diario_decimal).prod() - 1)
    anual = float((1 + acum) ** (252 / len(cdi_diario_decimal)) - 1)
    return acum, anual


def valor_portfolio(precos_periodo: pd.DataFrame, pesos: dict) -> pd.Series:
    tickers = list(pesos.keys())
    w = np.array([pesos[t] for t in tickers])
    p = precos_periodo[tickers].copy()
    p_rel = p.div(p.iloc[0])
    return (p_rel * w).sum(axis=1)


def otimizar_max_sharpe(
    log_ret_is: pd.DataFrame,
    pesos_rec: np.ndarray,
    cdi_is_anual: float,
    alpha: float = 0.5,
    beta:  float = 2.0,
) -> np.ndarray:
    from scipy.optimize import minimize

    mu    = log_ret_is.mean().values * 252
    sigma = log_ret_is.cov().values  * 252
    n     = len(mu)

    def neg_sharpe(w: np.ndarray) -> float:
        ret = float(w @ mu)
        vol = float(np.sqrt(w @ sigma @ w))
        return -(ret - cdi_is_anual) / vol if vol > 1e-10 else np.inf

    bounds = [
        (max(0.0, alpha * pesos_rec[i]), min(1.0, beta * pesos_rec[i]))
        for i in range(n)
    ]
    constraints = [{"type": "eq", "fun": lambda w: np.sum(w) - 1.0}]

    centros = np.array([(b[0]+b[1])/2 for b in bounds])
    centros /= centros.sum()

    best, best_val = None, np.inf
    for w0 in [pesos_rec.copy(), centros]:
        w0c = np.clip(w0, [b[0] for b in bounds], [b[1] for b in bounds])
        w0c = w0c / w0c.sum()
        res = minimize(neg_sharpe, w0c, method="SLSQP",
                       bounds=bounds, constraints=constraints,
                       options={"ftol": 1e-12, "maxiter": 2000})
        if res.fun < best_val:
            best_val = res.fun
            best = res
    if best is None:
        print("    Aviso: otimizador nao encontrou solucao — retornando pesos do analista.")
        return pesos_rec.copy()
    return best.x


def simular_bootstrap(ret_diarios: np.ndarray, cdi_oos_anual: float,
                      n_sims: int = 5_000, horizonte: int = 252,
                      seed: int = 42) -> dict:
    """
    Bootstrap histórico iid.
    ret_diarios: array 1D de retornos diários do portfólio no OOS.
    Retorna: dict de arrays com n_sims resultados por métrica.
    """
    rng = np.random.default_rng(seed)

    ret_acum_arr  = np.empty(n_sims)
    ret_anual_arr = np.empty(n_sims)
    vol_arr       = np.empty(n_sims)
    sharpe_arr    = np.empty(n_sims)
    mdd_arr       = np.empty(n_sims)

    for i in range(n_sims):
        # Reamostrar com reposição
        sample = rng.choice(ret_diarios, size=horizonte, replace=True)

        # Trajetória de valor
        v = np.cumprod(1.0 + sample)

        # Retorno acumulado e anualizado composto
        r_acum  = float(v[-1] - 1)
        r_anual = float((1 + r_acum) ** (252 / horizonte) - 1)

        # Volatilidade anualizada
        log_s = np.log(1.0 + sample)
        vol   = float(log_s.std(ddof=1) * np.sqrt(252))

        # Sharpe (retorno composto anualizado)
        sharpe = (r_anual - cdi_oos_anual) / vol if vol > 1e-8 else np.nan

        # Max Drawdown
        rolling_max = np.maximum.accumulate(v)
        dd          = v / rolling_max - 1.0
        mdd         = float(dd.min())

        ret_acum_arr[i]  = r_acum
        ret_anual_arr[i] = r_anual
        vol_arr[i]       = vol
        sharpe_arr[i]    = sharpe
        mdd_arr[i]       = mdd

    return {
        "Retorno Acumulado (252d)":  ret_acum_arr,
        "Retorno Anualizado":        ret_anual_arr,
        "Volatilidade Anual":        vol_arr,
        "Sharpe Ratio":              sharpe_arr,
        "Max Drawdown":              mdd_arr,
    }


def percentis(arr: np.ndarray) -> dict:
    arr_clean = arr[~np.isnan(arr)]
    return {
        "p5":    float(np.percentile(arr_clean,  5)),
        "p25":   float(np.percentile(arr_clean, 25)),
        "p50":   float(np.percentile(arr_clean, 50)),
        "p75":   float(np.percentile(arr_clean, 75)),
        "p95":   float(np.percentile(arr_clean, 95)),
        "média": float(arr_clean.mean()),
        "desvio": float(arr_clean.std()),
    }


# ---------------------------------------------------------------------------
# Carregar dados
# ---------------------------------------------------------------------------
print("=" * 65)
print("monte_carlo.py — Bootstrap histórico sobre período OOS")
print(f"  Simulações : {N_SIMS:,}")
print(f"  Horizonte  : {HORIZONTE} pregões (≈ 1 ano)")
print(f"  OOS        : {DATA_OOS_INI} → {DATA_OOS_FIM}")
print("=" * 65)

print("\nCarregando dados...")
precos = pd.read_parquet(os.path.join(DADOS_DIR, "precos.parquet"))
cdi_df = pd.read_parquet(os.path.join(DADOS_DIR, "cdi.parquet"))

cdi_decimal = cdi_df["CDI"] / 100.0
_, cdi_is_anual  = cdi_anualizado(cdi_decimal.loc[DATA_IS_INI:DATA_IS_FIM])
_, cdi_oos_anual = cdi_anualizado(cdi_decimal.loc[DATA_OOS_INI:DATA_OOS_FIM])

precos_is  = precos.loc[DATA_IS_INI:DATA_IS_FIM]
precos_oos = precos.loc[DATA_OOS_INI:DATA_OOS_FIM]

# ---------------------------------------------------------------------------
# Loop por carteira
# ---------------------------------------------------------------------------
lista_percentis  = []
lista_simulacoes = []   # todas as simulações para conferência

for nome, cfg in CARTEIRAS.items():
    print(f"\n{'─'*65}")
    print(f"Carteira: {nome}")

    tickers = list(cfg["tickers"])
    w_rec   = np.array(cfg["pesos"], dtype=float)

    # Remover tickers sem dados e redistribuir pesos (mesma lógica do backtest)
    tickers_sem_dados = [
        t for t in tickers
        if t not in precos.columns or precos[t].dropna().empty
    ]
    if tickers_sem_dados:
        print(f"  Aviso: {tickers_sem_dados} sem dados — removidos e pesos redistribuídos.")
        idx_ok  = [i for i, t in enumerate(tickers) if t not in tickers_sem_dados]
        tickers = [tickers[i] for i in idx_ok]
        w_rec   = w_rec[idx_ok] / w_rec[idx_ok].sum()

    # Recalcular w_opt (mesma lógica de backtest_estatico.py)
    p_is = precos_is[tickers].dropna(how="any")
    log_ret_is = np.log(p_is / p_is.shift(1)).dropna()
    w_opt = otimizar_max_sharpe(log_ret_is, w_rec, cdi_is_anual, ALPHA, BETA)

    # Série de valor OOS (buy-and-hold estático)
    p_oos = precos_oos[tickers].dropna(how="any")
    if p_oos.empty:
        print(f"  Erro: sem dados OOS. Pulando.")
        continue

    print(f"  OOS: {p_oos.index[0].date()} → {p_oos.index[-1].date()} | {len(p_oos)} pregões")

    for label, w in [("Analista", w_rec), ("Otimizado", w_opt)]:
        print(f"\n  Simulando: {label} ...", end=" ", flush=True)

        v_port     = valor_portfolio(p_oos, dict(zip(tickers, w)))
        ret_diarios = v_port.pct_change().dropna().values

        resultados = simular_bootstrap(
            ret_diarios, cdi_oos_anual,
            n_sims=N_SIMS, horizonte=HORIZONTE, seed=SEED
        )

        print(f"concluído ({N_SIMS:,} trajetórias)")

        # Percentis
        for metrica, arr in resultados.items():
            p = percentis(arr)
            row = {
                "Carteira": nome,
                "Tipo":     label,
                "Métrica":  metrica,
                **p,
            }
            lista_percentis.append(row)

            # Resumo no terminal
            if metrica in ("Retorno Acumulado (252d)", "Sharpe Ratio"):
                fmt = ".2%" if "Retorno" in metrica else ".2f"
                print(f"    {metrica:<32}  "
                      f"p5={p['p5']:{fmt}}  p50={p['p50']:{fmt}}  p95={p['p95']:{fmt}}")

        # Salvar todas as simulações para auditoria
        df_sim = pd.DataFrame(resultados)
        df_sim.insert(0, "Tipo",     label)
        df_sim.insert(0, "Carteira", nome)
        lista_simulacoes.append(df_sim)

# ---------------------------------------------------------------------------
# Salvar Excel
# ---------------------------------------------------------------------------
print(f"\n{'─'*65}")
print("Salvando Excel...")

if not lista_percentis:
    print("\nNenhuma simulacao concluida (verifique os dados de entrada).")
    sys.exit(1)

df_percentis  = pd.DataFrame(lista_percentis)
df_simulacoes = pd.concat(lista_simulacoes, ignore_index=True) if lista_simulacoes else pd.DataFrame()

path_mc = os.path.join(OUTPUT_DIR, "monte_carlo_2022.xlsx")

with pd.ExcelWriter(path_mc, engine="openpyxl") as writer:
    df_percentis.to_excel(writer,  sheet_name="Percentis",  index=False)
    df_simulacoes.to_excel(writer, sheet_name="Simulacoes", index=False)

    for sheet_name in ["Percentis", "Simulacoes"]:
        ws = writer.sheets[sheet_name]
        for col in ws.iter_cols(1, ws.max_column, 1, 1):
            ws.column_dimensions[col[0].column_letter].width = 20

print(f"Monte Carlo salvo: {path_mc}")
print("\nEtapa 3 concluída.")
print("=" * 65)
print("Resultados em: planilhas/")
print("  backtest_2022.xlsx")
print("  monte_carlo_2022.xlsx")
print("=" * 65)
