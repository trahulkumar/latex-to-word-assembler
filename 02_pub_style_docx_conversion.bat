@echo off
echo ===========================================================
echo   Latex to Publisher-Style DOCX Conversion Utility
echo   -----------------------------------------------------------
echo   This script automates the conversion of LaTeX chapters 
echo   into professional DOCX files with specific publisher 
echo   styles (Lora font, centered images, red figure details).
echo ===========================================================
echo.

echo Enter Chapter Number (e.g. 1) [10 second timeout]:
set "id="
for /f "delims=" %%I in ('powershell -command "$t = [Console]::In.ReadLineAsync(); if($t.Wait(10000)){Write-Output $t.Result}"') do set "id=%%I"

if "%id%"=="" (
    echo.
    echo No input received or timeout occurred. Exiting.
    echo Exiting in 10 seconds...
    timeout /t 10 >nul
    exit /b 1
)

echo.
echo Running Publisher Style Conversion for Chapter %id%...
echo.

uv run src/convert_to_pub_docx.py --chapter %id%

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ===========================================================
    echo   ERROR: Conversion failed for Chapter %id%.
    echo ===========================================================
    echo.
    echo Script will exit in 30 seconds...
    timeout /t 30
    exit /b %ERRORLEVEL%
)

echo.
echo ===========================================================
echo   SUCCESS: Conversion complete for Chapter %id%.
echo ===========================================================
echo.
echo Output window will close in 30 seconds...
timeout /t 30
