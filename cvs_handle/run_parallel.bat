@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "..\\.venv\\Scripts\\activate.bat" (
    call "..\\.venv\\Scripts\\activate.bat"
)

python -u fix_xhs_note_urls_parallel.py --workers 10 %*

pause
