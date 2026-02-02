@echo off
echo Enter Chapter Number (e.g. 1) [10 second timeout]:
set "id="
for /f "delims=" %%I in ('powershell -command "$t = [Console]::In.ReadLineAsync(); if($t.Wait(10000)){Write-Output $t.Result}"') do set "id=%%I"

if "%id%"=="" (
    echo.
    echo No input received or timeout occurred. Exiting.
    exit /b 1
)
echo Running Publisher Style Conversion for Chapter %id%...
uv run src/convert_to_pub_docx.py --chapter %id%
if %ERRORLEVEL% NEQ 0 (
    echo Conversion failed.
    exit /b %ERRORLEVEL%
)
echo Done.
