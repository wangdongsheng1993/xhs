@echo off
setlocal
cd /d "%~dp0"
set "XHS_FEISHU_SOURCE=https://mv21kbvltn.feishu.cn/wiki/PzBaw9C66iZtRFkPvfbcVCDqnBb"
set "XHS_FEISHU_UPLOAD_SOURCE=https://mv21kbvltn.feishu.cn/wiki/PzBaw9C66iZtRFkPvfbcVCDqnBb"

if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "xhs_gui_launcher.py"
    exit /b 0
)

if exist ".venv\Scripts\python.exe" (
    start "" ".venv\Scripts\python.exe" "xhs_gui_launcher.py"
    exit /b 0
)

where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "xhs_gui_launcher.py"
    exit /b 0
)

python "xhs_gui_launcher.py"
