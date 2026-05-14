@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "GUI_SCRIPT=%~dp0split_csv_upload_gui.py"

if exist "..\.venv\Scripts\pythonw.exe" (
    start "" "..\.venv\Scripts\pythonw.exe" "%GUI_SCRIPT%"
    exit /b 0
)

if exist "..\.venv\Scripts\python.exe" (
    start "" "..\.venv\Scripts\python.exe" "%GUI_SCRIPT%"
    exit /b 0
)

where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "%GUI_SCRIPT%"
    exit /b 0
)

python "%GUI_SCRIPT%"
