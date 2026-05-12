@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

:: 示例用法:
:: run_parallel.bat --excel input.csv --output result.csv --workers 3 --sessions browser_session_1,browser_session_2 --speed-mode turbo

if exist "..\\.venv\\Scripts\\activate.bat" (
    call "..\\.venv\\Scripts\\activate.bat"
)

python -u fix_xhs_note_urls_parallel.py %*

pause
