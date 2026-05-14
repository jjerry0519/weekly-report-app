@echo off
setlocal
cd /d "%~dp0"

set "BUNDLED_PY=C:\Users\user\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" server.py
) else (
  py -3 server.py
  if errorlevel 1 (
    python server.py
  )
)
pause
