# -*- coding: utf-8 -*-
"""
coletar_dados_b3_4a.py
═══════════════════════════════════════════════════════════════════════════
Coleta 4 anos de preços via Yahoo Finance para as 5 carteiras de dividendos
+ benchmarks DIVO11 e IBOV.

Tratamento de tickers com histórico incompleto:
  - ALOS3: criado em dez/2022 (fusão Aliansce Sonae + BR Malls).
           Predecessor: ALSO3. Retornos ALSO3 (jun/2022–dez/2022) são
           emendados com retornos ALOS3 (jan/2023–hoje) ao nível de
           log-retorno, preservando a distribuição empírica.
  - AXIA3: listada em jun/2022 (privatização Eletrobras). Predecessor:
           ELET6. Emenda igual — cobre os ~10 pregões iniciais ausentes.
  - VBBR3: renomeada de BRDT3 em 2021. Yahoo Finance já ajusta
           automaticamente; verificação explícita de cobertura garante
           que o histórico chegue a jun/2022.

Saída: analise_4anos/dados/precos_b3.csv
═══════════════════════════════════════════════════════════════════════════
"""

import os, sys, time, datetime
import requests
import numpy as np
import pandas as pd

# ── Caminhos ───────────────────────────────────────────────────────────────────
PASTA   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_CSV = os.path.join(PASTA, "dados", "precos_b3.csv")
os.makedirs(os.path.join(PASTA, "dados"), exist_ok=True)

# ── Parâmetros ─────────────────────────────────────────────────────────────────
START_DATE   = "2022-06-01"      # início da janela de 4 anos
MAX_FFILL    = 5                 # pregões consecutivos sem negociação: ffill ok
MIN_PREGOES  = 700               # mínimo de dias originais para aprovação

PERIOD1 = int(datetime.datetime(2022, 6, 1).timestamp())
PERIOD2 = int(datetime.datetime.today().replace(hour=23, minute=59).timestamp())

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
}

# ── Tickers alvo e predecessores ───────────────────────────────────────────────
# Apenas os tickers efetivamente usados nas 5 carteiras + benchmarks.
# Predecessores são baixados separadamente e usados apenas para emendar os gaps.

TICKERS_ALVO = [
    # Benchmarks
    "DIVO11", "BOVA11",
    # Carteira Casa A
    "PETR4", "PRIO3", "VALE3", "CPLE3", "ENGI11", "AXIA3", "VIVT3",
    "ALOS3", "B3SA3", "CXSE3", "ITUB4",
    # Carteira Casa B (incrementais)
    "BBDC4", "EQTL3", "MOTV3", "CSMG3", "CURY3",
    # Carteira Casa D (incrementais)
    "BPAC11", "VBBR3", "PETR3",
    # Carteira BB (incrementais)
    "ABEV3", "BRAP4", "DIRR3", "ITSA4", "TAEE11", "TIMS3",
]

# Predecessor → (ticker_predecessor, descrição)
# Usados para emendar histórico quando o ticker alvo tem < MIN_PREGOES pregões.
PREDECESSORES = {
    "ALOS3": ("ALSO3", "Aliansche Sonae (pre-fusão ALOS3, dez/2022)"),
    "AXIA3": ("ELET6", "Eletrobras PN (pre-privatização AXIA3, jun/2022)"),
}

# ── Funções de download ────────────────────────────────────────────────────────
def fetch_yahoo(ticker: str, symbol_override: str = None) -> pd.Series | None:
    """Retorna série de fechamentos ajustados (DatetimeIndex)."""
    symbol = symbol_override if symbol_override else ticker + ".SA"
    url = (
        f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"
        f"?interval=1d&period1={PERIOD1}&period2={PERIOD2}&events=adjclose"
    )
    for tentativa in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=15)
            if r.status_code == 429:
                time.sleep(5 * (tentativa + 1))
                continue
            if r.status_code != 200:
                return None
            j = r.json()
            result = j.get("chart", {}).get("result")
            if not result:
                return None
            result = result[0]
            timestamps = result.get("timestamp", [])
            indicators = result.get("indicators", {})
            adj = indicators.get("adjclose", [{}])
            prices_raw = (
                adj[0].get("adjclose", [])
                if adj and adj[0].get("adjclose")
                else indicators.get("quote", [{}])[0].get("close", [])
            )
            if not timestamps or not prices_raw:
                return None
            idx, vals = [], []
            for ts, p in zip(timestamps, prices_raw):
                if p is not None and p > 0:
                    idx.append(datetime.date.fromtimestamp(ts))
                    vals.append(float(p))
            if not idx:
                return None
            s = pd.Series(vals, index=pd.to_datetime(idx)).sort_index()
            return s[~s.index.duplicated(keep="last")]
        except Exception:
            time.sleep(2)
    return None


def reconstruct_prices_from_returns(log_rets: np.ndarray,
                                    dates: pd.DatetimeIndex,
                                    base: float = 100.0) -> pd.Series:
    """
    Reconstrói série de preços a partir de log-retornos.
    price_0 = base; price_t = base * exp(cumsum(log_rets[:t]))
    A série terá len(dates) pontos, com price[0] = base.
    """
    cum = np.concatenate([[0.0], np.cumsum(log_rets)])
    prices = base * np.exp(cum)
    return pd.Series(prices, index=dates)


def emenda_retornos(serie_pred: pd.Series,
                    serie_alvo: pd.Series,
                    ticker: str,
                    pred_ticker: str) -> pd.Series:
    """
    Emenda série predecessor + série alvo ao nível de log-retorno.

    Procedimento:
      1. Garante que predecessor cobre o período anterior ao alvo.
      2. Calcula log-retornos de cada série individualmente.
      3. Concatena os log-retornos (predecessor até handoff, alvo a partir do handoff).
      4. Reconstrói série de preços a partir dos retornos concatenados.
      5. Retorna série sintética com DatetimeIndex consistente.
    """
    # Alinhar índices aos dias úteis (sem hora)
    s_pred = serie_pred.copy(); s_pred.index = pd.to_datetime(s_pred.index.date)
    s_alvo = serie_alvo.copy(); s_alvo.index = pd.to_datetime(s_alvo.index.date)

    # Datas antes do início do alvo usam predecessor
    handoff = s_alvo.index[0]
    pred_window = s_pred[s_pred.index < handoff].copy()

    if len(pred_window) < 5:
        print(f"    AVISO [{ticker}]: predecessor {pred_ticker} tem apenas "
              f"{len(pred_window)} dias antes do alvo — emenda ignorada.")
        return s_alvo

    print(f"    Emenda {pred_ticker}→{ticker}: "
          f"{pred_window.index[0].date()}→{pred_window.index[-1].date()} "
          f"({len(pred_window)} dias pred) + "
          f"{s_alvo.index[0].date()}→{s_alvo.index[-1].date()} "
          f"({len(s_alvo)} dias alvo)")

    # Log-retornos do predecessor no período da lacuna
    pred_vals = pred_window.ffill(limit=MAX_FFILL).values.astype(float)
    lr_pred = np.where(
        np.isnan(pred_vals[1:]) | np.isnan(pred_vals[:-1]), 0.0,
        np.log(np.clip(pred_vals[1:] / pred_vals[:-1], 1e-8, None))
    )

    # Log-retornos do alvo
    alvo_vals = s_alvo.ffill(limit=MAX_FFILL).values.astype(float)
    lr_alvo = np.where(
        np.isnan(alvo_vals[1:]) | np.isnan(alvo_vals[:-1]), 0.0,
        np.log(np.clip(alvo_vals[1:] / alvo_vals[:-1], 1e-8, None))
    )

    # Datas combinadas: predecessor (sem último ponto) + alvo
    datas_pred_retornos = pred_window.index[1:]          # N-1 datas de retornos pred
    datas_alvo_retornos = s_alvo.index[1:]               # N-1 datas de retornos alvo
    datas_todas_retornos = datas_pred_retornos.append(datas_alvo_retornos)

    # Verificar sobreposição de datas
    overlapping = datas_pred_retornos[datas_pred_retornos.isin(datas_alvo_retornos)]
    if len(overlapping) > 0:
        print(f"    AVISO [{ticker}]: {len(overlapping)} datas sobrepostas entre "
              f"predecessor e alvo — mantendo apenas alvo.")
        mask_pred = ~datas_pred_retornos.isin(datas_alvo_retornos)
        lr_pred = lr_pred[mask_pred.values]
        datas_pred_retornos = datas_pred_retornos[mask_pred]
        datas_todas_retornos = datas_pred_retornos.append(datas_alvo_retornos)

    lr_concat = np.concatenate([lr_pred, lr_alvo])

    # Datas do preço: data_inicial_pred + todas datas de retornos
    datas_precos = pred_window.index[:1].append(datas_todas_retornos)
    serie_sintetica = reconstruct_prices_from_returns(lr_concat, datas_precos)

    return serie_sintetica


# ══════════════════════════════════════════════════════════════════════════════
# DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════════
print("═" * 70)
print(f"Coleta de dados: {START_DATE} → hoje  |  Janela: 4 anos (~1.008 pregões)")
print("═" * 70)

# Benchmark IBOV
print("\n[BM] Baixando IBOV (^BVSP)...")
s_ibov = fetch_yahoo("IBOV", symbol_override="^BVSP")
if s_ibov is None or len(s_ibov) < 200:
    print("ERRO: IBOV não disponível. Verifique conexão.")
    sys.exit(1)
print(f"[BM] OK  IBOV  ({len(s_ibov)} dias)\n")

series = {"IBOV": s_ibov}
predecessores_cache = {}
falhas = []

# Baixar predecessores antecipadamente
print("Baixando predecessores para emenda de histórico...")
for ticker_alvo, (pred_tk, desc) in PREDECESSORES.items():
    print(f"  [{pred_tk}] {desc}")
    s = fetch_yahoo(pred_tk)
    if s is not None and len(s) >= 10:
        predecessores_cache[pred_tk] = s
        print(f"  OK {pred_tk}: {len(s)} dias ({s.index[0].date()} → {s.index[-1].date()})")
    else:
        print(f"  -- {pred_tk}: sem dados — emenda de {ticker_alvo} será ignorada")
    time.sleep(0.8)

# Baixar tickers alvo
print(f"\nBaixando {len(TICKERS_ALVO)} tickers alvo...")
for i, tk in enumerate(TICKERS_ALVO):
    s = fetch_yahoo(tk)
    n = len(s) if s is not None else 0
    if s is not None and n >= 50:
        # Verificar se precisa emenda
        if tk in PREDECESSORES and n < MIN_PREGOES:
            pred_tk, _ = PREDECESSORES[tk]
            if pred_tk in predecessores_cache:
                print(f"  [{i+1:2d}] {tk:8s}: {n:4d} dias → emendando com {pred_tk}...", end="")
                s_stitched = emenda_retornos(predecessores_cache[pred_tk], s, tk, pred_tk)
                series[tk] = s_stitched
                print(f" → {len(s_stitched)} dias após emenda")
            else:
                series[tk] = s
                print(f"  [{i+1:2d}] OK  {tk:8s}: {n:4d} dias (sem predecessor disponível)")
        elif tk in PREDECESSORES:
            # Tem cobertura suficiente, mas avisa caso precise emenda de borda inicial
            series[tk] = s
            print(f"  [{i+1:2d}] OK  {tk:8s}: {n:4d} dias (cobertura OK, sem emenda necessária)")
        else:
            series[tk] = s
            print(f"  [{i+1:2d}] OK  {tk:8s}: {n:4d} dias")
    else:
        falhas.append(tk)
        print(f"  [{i+1:2d}] --  {tk:8s}: sem dados suficientes ({n} dias)")
    time.sleep(0.5)

print(f"\nBaixados: {len(series)-1} tickers alvo + IBOV | Falhas: {len(falhas)}")
if falhas:
    print(f"Falhas: {falhas}")

# ══════════════════════════════════════════════════════════════════════════════
# MONTAR DataFrame
# ══════════════════════════════════════════════════════════════════════════════
df_raw = pd.DataFrame(series).sort_index()
df_raw = df_raw[df_raw.index.dayofweek < 5].dropna(how="all")

# Reindexar ao calendário de pregões do IBOV
cal_ibov = df_raw.index[df_raw["IBOV"].notna()]
df_cal   = df_raw.reindex(cal_ibov)

# Forward-fill com limite
df_filled = df_cal.ffill(limit=MAX_FFILL)

# Filtro de qualidade: cobertura mínima (dias com preço ORIGINAL)
counts_orig = df_cal.notna().sum()
ok = counts_orig[counts_orig >= MIN_PREGOES].index.tolist()
descartados = [t for t in df_filled.columns if t not in ok]
if descartados:
    print(f"\nFiltro cobertura >= {MIN_PREGOES} dias: descartados {descartados}")

df_q = df_filled[ok].dropna(how="all")

# IBOV como primeira coluna
if "IBOV" in df_q.columns:
    df_q = df_q[["IBOV"] + [c for c in df_q.columns if c != "IBOV"]]

# ══════════════════════════════════════════════════════════════════════════════
# RELATÓRIO
# ══════════════════════════════════════════════════════════════════════════════
print(f"\n{'═'*60}")
print(f"  TICKERS APROVADOS : {len(df_q.columns)}")
print(f"  PREGÕES NO CSV    : {len(df_q)}")
print(f"  PERÍODO           : {df_q.index[0].date()} → {df_q.index[-1].date()}")
print(f"\n  Cobertura por ticker:")
for tk in df_q.columns:
    dias = int(counts_orig.get(tk, 0))
    sufixo = "  ← emenda de predecessor" if tk in PREDECESSORES else ""
    print(f"    {tk:10s}: {dias:4d} dias originais{sufixo}")
print(f"{'═'*60}")

# ══════════════════════════════════════════════════════════════════════════════
# SALVAR
# ══════════════════════════════════════════════════════════════════════════════
df_q.to_csv(OUT_CSV, date_format="%Y-%m-%d")
print(f"\nSalvo: {OUT_CSV}")
print(f"({len(df_q.columns)} colunas × {len(df_q)} linhas)")
print("\nPróximo passo: python analise_4anos/scripts/analisar_carteiras.py")
