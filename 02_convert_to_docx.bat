@echo off
echo Running Conversion...
uv run src/convert_to_docx.py
if %ERRORLEVEL% NEQ 0 (
    echo Conversion failed.
    exit /b %ERRORLEVEL%
)
echo Done.

