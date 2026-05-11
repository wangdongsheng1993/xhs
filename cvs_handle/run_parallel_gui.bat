@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "GUI_SCRIPT=%~dp0fix_xhs_note_urls_parallel_gui.py"
set "LOG_DIR=%~dp0logs"
set "PYTHON_CMD="

if exist "..\\.venv\\Scripts\\pythonw.exe" (
    set "PYTHON_CMD=%~dp0..\.venv\Scripts\pythonw.exe"
) else if exist "..\\.venv\\Scripts\\python.exe" (
    set "PYTHON_CMD=%~dp0..\.venv\Scripts\python.exe"
) else (
    set "PYTHON_CMD=pythonw"
)

powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command ^
  "$guiScript = '%GUI_SCRIPT%';" ^
  "$pythonCmd = '%PYTHON_CMD%';" ^
  "$logDir = '%LOG_DIR%';" ^
  "New-Item -ItemType Directory -Force -Path $logDir | Out-Null;" ^
  "$logPath = Join-Path $logDir 'parallel_gui_runtime.log';" ^
  "$start = Get-Date;" ^
  "$process = Start-Process -FilePath $pythonCmd -ArgumentList @($guiScript) -PassThru;" ^
  "Wait-Process -Id $process.Id;" ^
  "$end = Get-Date;" ^
  "$elapsed = $end - $start;" ^
  "$line = '{0} | start={1:yyyy-MM-dd HH:mm:ss} | end={2:yyyy-MM-dd HH:mm:ss} | elapsed={3:hh\:mm\:ss}' -f $process.Id, $start, $end, $elapsed;" ^
  "Add-Content -Path $logPath -Value $line -Encoding UTF8"
