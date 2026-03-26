@echo off
title Tri-Tech TIA - HTML to PPTX
echo ========================================
echo   Tri-Tech TIA - HTML to PPTX Converter
echo ========================================
echo.

cd /d "%~dp0"

set PYTHON=
where python >nul 2>nul && set PYTHON=python
if not defined PYTHON where py >nul 2>nul && set PYTHON=py -3
if not defined PYTHON if exist "C:\Program Files\Python313\python.exe" set PYTHON="C:\Program Files\Python313\python.exe"
if not defined PYTHON (
    echo [ERROR] Python not found. Install Python 3.13+ from https://python.org
    pause
    exit /b 1
)

echo Using: %PYTHON%
echo.

%PYTHON% html_to_pptx.py presentazione_html
if %ERRORLEVEL% NEQ 0 goto :error

if not exist "Slides1.pptx" (
    echo [WARNING] Slides1.pptx was not created.
    pause
    exit /b 0
)

echo.
echo Done. Opening Slides1.pptx...
start "" "Slides1.pptx"
pause
exit /b 0

:error
echo.
echo [ERROR] Conversion failed.
pause
exit /b 1
