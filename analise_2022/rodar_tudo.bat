@echo off
setlocal

echo.
echo ================================================================
echo  Backtest IS/OOS 2022 - Casa A e Casa B
echo ================================================================
echo.

cd /d "%~dp0"

REM ---- Verificar Python ----
where python >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=python
    goto :python_ok
)
where py >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=py
    goto :python_ok
)
echo ERRO: Python nao encontrado no PATH.
echo Instale Python 3.10+ em https://www.python.org/downloads/
echo Marque "Add Python to PATH" durante a instalacao.
pause
exit /b 1

:python_ok
%PYTHON% --version
echo.

REM ---- yfinance recente usa protobuf com extensao C incompativel com Python 3.14 ----
REM ---- Fixar versao 0.2.50 (sem WebSocket/live, sem protobuf) ----
REM ---- e forcar implementacao pura-Python do protobuf caso instalado ----
set PROTOCOL_BUFFERS_PYTHON_IMPLEMENTATION=python

REM ---- Instalar dependencias ----
echo [1/4] Instalando dependencias...
%PYTHON% -m pip install --quiet --upgrade yfinance python-bcb pandas numpy scipy openpyxl pyarrow
if %errorlevel% neq 0 (
    echo ERRO: falha na instalacao de dependencias.
    pause
    exit /b 1
)
echo OK
echo.

REM ---- Coletar dados ----
echo [2/4] Coletando dados (Yahoo Finance + BCB)...
%PYTHON% coletar_dados.py
if %errorlevel% neq 0 (
    echo ERRO: falha na coleta de dados.
    echo Verifique conexao com a internet.
    pause
    exit /b 1
)
echo.

REM ---- Backtest estatico ----
echo [3/4] Backtest estatico (IS + OOS)...
%PYTHON% backtest_estatico.py
if %errorlevel% neq 0 (
    echo ERRO: falha no backtest.
    pause
    exit /b 1
)
echo.

REM ---- Monte Carlo ----
echo [4/4] Monte Carlo bootstrap (5.000 simulacoes)...
%PYTHON% monte_carlo.py
if %errorlevel% neq 0 (
    echo ERRO: falha no Monte Carlo.
    pause
    exit /b 1
)

echo.
echo ================================================================
echo  Concluido!
echo  Resultados em: planilhas\
echo    backtest_2022.xlsx
echo    monte_carlo_2022.xlsx
echo ================================================================
echo.
pause
endlocal
