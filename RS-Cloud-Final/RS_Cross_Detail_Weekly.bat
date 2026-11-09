@echo off
title RS Cross Screener (Detail) - Weekly
cd /d "%~dp0"

rem ---- find Python ----
set PY=
py --version >nul 2>&1 && set PY=py
if not defined PY python --version >nul 2>&1 && set PY=python
if not defined PY if exist "%USERPROFILE%\python.exe" set PY="%USERPROFILE%\python.exe"
if not defined PY (
    echo Python was not found. Install it from python.org and tick "Add Python to PATH".
    pause
    exit /b
)

echo Checking required packages...
%PY% -m pip install --quiet --upgrade yfinance pandas openpyxl
echo.
%PY% rs_cross_detail_weekly.py %1
echo.
pause
