@echo off
cd /d "D:\OneDrive\06. Projetos e Escrita\Otimização de Portfolio"

echo ============================================================
echo  ANALISE 4 ANOS - Carteiras de Dividendos
echo ============================================================
echo.

echo [1/4] Coletando dados (4 anos via Yahoo Finance)...
python analise_4anos\scripts\coletar_dados_b3_4a.py
if errorlevel 1 (
    echo ERRO na coleta de dados. Abortando.
    pause
    exit /b 1
)

echo.
echo [2/4] Analisando carteiras (full-sample)...
python analise_4anos\scripts\analisar_carteiras.py
if errorlevel 1 (
    echo ERRO em analisar_carteiras. Abortando.
    pause
    exit /b 1
)

echo.
echo [3/4] Backtest benchmark (split 75/25)...
python analise_4anos\scripts\backtest_benchmark.py
if errorlevel 1 (
    echo ERRO em backtest_benchmark. Abortando.
    pause
    exit /b 1
)

echo.
echo [4/4] Monte Carlo / Bootstrap historico...
python analise_4anos\scripts\monte_carlo.py
if errorlevel 1 (
    echo ERRO em monte_carlo. Abortando.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  CONCLUIDO. Planilhas em: analise_4anos\planilhas\
echo ============================================================
pause
