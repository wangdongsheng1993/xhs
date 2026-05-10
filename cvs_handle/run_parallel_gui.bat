@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "..\\.venv\\Scripts\\pythonw.exe" (
    start "" "..\\.venv\\Scripts\\pythonw.exe" fix_xhs_note_urls_parallel_gui.py
) else if exist "..\\.venv\\Scripts\\python.exe" (
    start "" "..\\.venv\\Scripts\\python.exe" fix_xhs_note_urls_parallel_gui.py
) else (
    start "" pythonw fix_xhs_note_urls_parallel_gui.py
)
