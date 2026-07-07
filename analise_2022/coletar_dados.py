"""
coletar_dados.py — Etapa 1 de 3
================================
Baixa preços ajustados e CDI para o experimento IS/OOS 2022.

Tickers (fonte: PDFs originais das casas — junho de 2022):
  Casa A — Dividendos:  BBAS3, CPLE6, EGIE3, GRND3, PETR4, TIMS3
  Casa B — Dividendos:     ALUP11, BBAS3, BBDC4, CYRE3, ENGI11, ITUB4, SBSP3, TRPL4, VALE3, VIVT3
  Benchmarks:         BOVA11, DIVO11

Período:
  Calibração IS : 2019-06-01 → 2022-05-31
  Teste OOS     : 2022-06-01 → 2026-06-28

Saída:
  dados/precos.parquet
  dados/cdi.parquet
"""

import os
import sys
import time

# ---------------------------------------------------------------------------
# Python 3.14 compatibility — protobuf C extension patch
# ---------------------------------------------------------------------------
# yfinance >= 0.2.55 depende de protobuf cujo C extension (google._upb._message)
# quebra no Python 3.14 (tp_new metaclass issue).
# Solução: mockar os módulos antes que yfinance os importe.
# Funcional: não usamos WebSocket/live data — só histórico.
if sys.version_info >= (3, 14):
    from unittest.mock import MagicMock as _Mock
    for _m in ['yfinance.pricing_pb2', 'yfinance.live']:
        sys.modules[_m] = _Mock()

import pandas as pd
import yfinance as yf

# ---------------------------------------------------------------------------
# Configurações
# ---------------------------------------------------------------------------
DATA_INICIO = "2019-06-01"
DATA_FIM    = "2026-06-28"

TICKERS_A    = ["BBAS3", "CPLE6", "EGIE3", "GRND3", "PETR4", "TIMS3"]
TICKERS_B   = ["ALUP11", "BBAS3", "BBDC4", "CYRE3", "ENGI11",
                  "ITUB4", "SBSP3", "TRPL4", "VALE3", "VIVT3"]
TICKERS_BENCH = ["BOVA11", "DIVO11"]

TODOS_TICKERS = sorted(set(TICKERS_A + TICKERS_B + TICKERS_BENCH))
TICKERS_YF    = [t + ".SA" for t in TODOS_TICKERS]

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
DADOS_DIR = os.path.join(BASE_DIR, "dados")
os.makedirs(DADOS_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# Teste de conectividade
# ---------------------------------------------------------------------------
print("Testando conectividade com Yahoo Finance...")
try:
    import urllib.request
    req = urllib.request.Request(
        "https://query1.finance.yahoo.com/v8/finance/chart/BBAS3.SA?interval=1d&range=5d",
        headers={"User-Agent": "Mozilla/5.0"}
    )
    with urllib.request.urlopen(req, timeout=10) as resp:
        body = resp.read(200)
        if body:
            print("  Yahoo Finance: OK")
        else:
            print("  Yahoo Finance: resposta vazia — pode haver bloqueio de IP ou VPN.")
except Exception as e:
    print(f"  Yahoo Finance: ERRO — {e}")
    print("  Verifique conexão, VPN ativa, ou firewall corporativo.")
    print("  O script continua — yfinance fará novas tentativas.")
print()

# ---------------------------------------------------------------------------
# Funções de download
# ---------------------------------------------------------------------------

def download_batch(tickers_yf, start, end):
    """Download em batch via yf.download(). Rápido mas pode falhar por rate-limit."""
    print("  Tentando download em batch...")
    raw = yf.download(tickers_yf, start=start, end=end, auto_adjust=True,
                      progress=True, threads=False)
    if raw.empty:
        return None

    precos = raw["Close"].copy() if "Close" in raw.columns else raw.xs("Close", axis=1, level=0)
    precos.columns = [str(c).replace(".SA", "") for c in precos.columns]
    precos.index = pd.to_datetime(precos.index)

    n_ok = int((~precos.isna().all()).sum())
    print(f"  Batch: {n_ok}/{len(precos.columns)} tickers com dados.")
    return precos if n_ok > 0 else None


def download_individual(tickers_yf, start, end, delay=0.8):
    """Fallback: download um ticker por vez com delay. Mais robusto contra rate-limit."""
    print(f"  Tentando download individual (delay={delay}s entre tickers)...")
    dfs = {}
    for t in tickers_yf:
        ticker_clean = t.replace(".SA", "")
        try:
            time.sleep(delay)
            obj = yf.Ticker(t)
            hist = obj.history(start=start, end=end, auto_adjust=True)
            if not hist.empty and "Close" in hist.columns:
                dfs[ticker_clean] = hist["Close"]
                print(f"    {ticker_clean:8s}: {len(hist)} obs ({hist.index[0].date()} → {hist.index[-1].date()})")
            else:
                print(f"    {ticker_clean:8s}: sem dados")
        except Exception as e:
            print(f"    {ticker_clean:8s}: erro — {e}")

    if not dfs:
        return None

    df = pd.DataFrame(dfs)
    df.index = pd.to_datetime(df.index)
    df.index = df.index.tz_localize(None)  # remover timezone para compatibilidade
    return df


# ---------------------------------------------------------------------------
# 1. Preços — Yahoo Finance
# ---------------------------------------------------------------------------
print("=" * 60)
print("Etapa 1/2: Preços via Yahoo Finance")
print(f"  Tickers : {', '.join(TODOS_TICKERS)}")
print(f"  Período : {DATA_INICIO} → {DATA_FIM}")
print("=" * 60)

precos = download_batch(TICKERS_YF, DATA_INICIO, DATA_FIM)

if precos is None:
    print("\n  Batch falhou. Tentando download individual...")
    precos = download_individual(TICKERS_YF, DATA_INICIO, DATA_FIM)

if precos is None:
    print("\nERRO CRÍTICO: nenhum dado foi obtido do Yahoo Finance.")
    print("Possíveis causas:")
    print("  1. Yahoo Finance bloqueando IPs brasileiros")
    print("  2. VPN ou proxy interferindo")
    print("  3. Rate limiting (espere alguns minutos e tente novamente)")
    sys.exit(1)

precos.index.name = "data"

# ---------------------------------------------------------------------------
# Substituições para tickers descontinuados
# ---------------------------------------------------------------------------
# CPLE6 (Copel PN): privatizada em ago/2023, ações PN convertidas para ON (CPLE3)
#   ratio 1:1 — o retorno de CPLE3 representa fielmente o retorno do investidor.
# TRPL4 (ISA CTEEP PN): ticker descontinuado — TRPL3 (ON) é o mesmo emissor
#   e tem dinâmica de retorno praticamente idêntica.
SUBSTITUICOES = {
    "CPLE6": "CPLE3",
    "TRPL4": "TRPL3",
}

for original, substituto in SUBSTITUICOES.items():
    if original not in precos.columns or precos[original].isna().all():
        print(f"\n  {original}: sem dados — baixando substituto {substituto}...")
        sub_yf = substituto + ".SA"
        try:
            time.sleep(1.0)
            sub_hist = yf.Ticker(sub_yf).history(
                start=DATA_INICIO, end=DATA_FIM, auto_adjust=True
            )
            if not sub_hist.empty and "Close" in sub_hist.columns:
                sub_series = sub_hist["Close"].copy()
                sub_series.index = pd.to_datetime(sub_series.index)
                if hasattr(sub_series.index, "tz") and sub_series.index.tz is not None:
                    sub_series.index = sub_series.index.tz_localize(None)
                sub_series.name = original   # manter nome original para compatibilidade
                # Alinhar ao índice do DataFrame principal
                precos[original] = sub_series.reindex(precos.index)
                n_ok = int(precos[original].notna().sum())
                print(f"  OK: {original} → {substituto} | {n_ok} obs disponíveis")
                print(f"  NOTA metodológica: {original} substituído por {substituto}.")
                print(f"        Mesmo emissor; retornos usados para calibração de covariâncias.")
            else:
                print(f"  {substituto}: também sem dados. {original} permanece vazio.")
        except Exception as e:
            print(f"  Erro ao baixar {substituto}: {e}")

# Relatório de cobertura
print("\nCobertura de dados por ticker:")
print(f"  {'Ticker':<8}  {'Início':>12}  {'Fim':>12}  {'Obs':>6}  {'NaN':>5}")
print("  " + "-" * 50)
for t in sorted(precos.columns):
    serie   = precos[t]
    validos = serie.dropna()
    inicio  = str(validos.index[0].date())  if not validos.empty else "N/D"
    fim     = str(validos.index[-1].date()) if not validos.empty else "N/D"
    n_obs   = len(validos)
    n_nan   = serie.isna().sum()
    print(f"  {t:<8}  {inicio:>12}  {fim:>12}  {n_obs:>6}  {n_nan:>5}")

path_precos = os.path.join(DADOS_DIR, "precos.parquet")
precos.to_parquet(path_precos)
print(f"\nPreços salvos: {path_precos}  {precos.shape}")

# ---------------------------------------------------------------------------
# 2. CDI diário — BCB SGS série 11
# ---------------------------------------------------------------------------
print("\n" + "=" * 60)
print("Etapa 2/2: CDI via BCB SGS (série 11)")
print("=" * 60)

try:
    from bcb import sgs
except ImportError:
    print("\nErro: python-bcb não instalado. Execute: pip install python-bcb")
    sys.exit(1)

cdi_raw = sgs.get({"CDI": 11}, start=DATA_INICIO, end=DATA_FIM)
cdi_raw.index = pd.to_datetime(cdi_raw.index)
cdi_raw.index.name = "data"

print(f"  Observações: {len(cdi_raw)}")
print(f"  Período    : {cdi_raw.index[0].date()} → {cdi_raw.index[-1].date()}")
print(f"  CDI inicial: {cdi_raw['CDI'].iloc[0]:.6f}% ao dia")
print(f"  CDI final  : {cdi_raw['CDI'].iloc[-1]:.6f}% ao dia")

path_cdi = os.path.join(DADOS_DIR, "cdi.parquet")
cdi_raw.to_parquet(path_cdi)
print(f"\nCDI salvo: {path_cdi}")

print("\n" + "=" * 60)
print("Etapa 1 concluída. Execute: python backtest_estatico.py")
print("=" * 60)
