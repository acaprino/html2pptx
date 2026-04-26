@echo off
setlocal
title html2pptx - HTML to PPTX
echo ========================================
echo   html2pptx - HTML to PPTX Converter
echo ========================================
echo.

cd /d "%~dp0"

set PYTHON=
where python >nul 2>nul && set PYTHON=python
if not defined PYTHON where py >nul 2>nul && set PYTHON=py -3
if not defined PYTHON (
    echo [ERROR] Python not found in PATH. Install Python 3.10+ from https://python.org
    pause
    exit /b 1
)

echo Using: %PYTHON%
echo.

%PYTHON% html_to_pptx.py -i presentazione_html -s 3
set RC=%ERRORLEVEL%

REM Distinguish exit codes: 0=success, 2=partial save, other=hard failure.
if %RC% EQU 2 (
    echo.
    echo [WARNING] Partial save - check .partial.pptx alongside the requested output.
    pause
    exit /b 2
)
if %RC% NEQ 0 goto :error

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
echo [ERROR] Conversion failed (exit code %RC%).
pause
exit /b %RC%
