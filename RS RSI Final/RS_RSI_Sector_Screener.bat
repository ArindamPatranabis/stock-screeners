@echo off
title RS Cross + Sector RS + RSI Stock Screener
cd /d "%~dp0"
set PY=
py --version >nul 2>&1 && set PY=py
if not defined PY python --version >nul 2>&1 && set PY=python
if not defined PY if exist "%USERPROFILE%\python.exe" set PY="%USERPROFILE%\python.exe"
if not defined PY (
    echo Python was not found.
    pause
    exit /b
)
echo Checking required packages...
%PY% -m pip install --quiet --upgrade yfinance pandas openpyxl
echo.
%PY% rs_rsi_sector_screener.py %1
echo.
pause
