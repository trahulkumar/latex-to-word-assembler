@echo off
set /p "id=Enter Chapter Number (e.g. 1): "
echo Running Conversion for Chapter %id%...
uv run src/convert_to_docx.py --chapter %id%
if %ERRORLEVEL% NEQ 0 (
    echo Conversion failed.
    exit /b %ERRORLEVEL%
)
echo Done.

