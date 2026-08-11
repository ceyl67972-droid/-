@echo off
setlocal
cd /d "%~dp0"

set "CODEX_RUNTIME=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python"
if exist "%CODEX_RUNTIME%\python.exe" (
  set "PYTHON=%CODEX_RUNTIME%\python.exe"
  set "PYTHONW=%CODEX_RUNTIME%\pythonw.exe"
) else (
  where py >nul 2>nul
  if errorlevel 1 (
    echo Python 3.10 or newer is required.
    pause
    exit /b 1
  )
  set "PYTHON=py -3"
  set "PYTHONW=pyw -3"
)

%PYTHON% -c "import flask, openpyxl, pdfplumber, pypdfium2, waitress, webview" >nul 2>nul
if errorlevel 1 %PYTHON% -m pip install -r requirements.txt

start "" %PYTHONW% "%~dp0desktop_app.py"
