@echo off
echo Running Metadata Generation...
uv run src/generate_metadata.py
if %ERRORLEVEL% NEQ 0 (
    echo Metadata generation failed.
    exit /b %ERRORLEVEL%
)
echo Done.

