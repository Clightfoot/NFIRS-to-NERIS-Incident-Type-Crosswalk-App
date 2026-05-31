@echo off
setlocal
cd /d "%~dp0"

set "BUNDLED_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "APP_URL=http://127.0.0.1:8765"

start "" "%APP_URL%"

if exist "%BUNDLED_PYTHON%" (
  "%BUNDLED_PYTHON%" app.py
) else (
  python app.py
)
