@echo off
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "fix_xhs_note_urls_gui.py"
    exit /b 0
)

if exist ".venv\Scripts\python.exe" (
    start "" ".venv\Scripts\python.exe" "fix_xhs_note_urls_gui.py"
    exit /b 0
)

where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "fix_xhs_note_urls_gui.py"
    exit /b 0
)

python "fix_xhs_note_urls_gui.py"
