@echo off
setlocal
cd /d "%~dp0"

rem Use a fresh port so older sandbox-started services do not interfere.
set PORT=8796

set "BUNDLED_PY=C:\Users\user\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

echo.
echo 同業送件明細網頁正式版啟動中...
echo 網址：http://localhost:%PORT%
echo.
echo 請保持這個視窗開著。要停止服務時，直接關閉此視窗即可。
echo.

start "" "http://localhost:%PORT%"

if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" server.py
) else (
  py -3 server.py
  if errorlevel 1 (
    python server.py
  )
)

pause
