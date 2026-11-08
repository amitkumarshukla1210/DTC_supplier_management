@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
  echo Python was not found. Install Python 3.10 or later and enable "Add Python to PATH".
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating the app's Python environment...
  python -m venv .venv
  if errorlevel 1 (
    echo Could not create the Python environment.
    pause
    exit /b 1
  )
)

echo Checking and installing required packages...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo Package installation failed. Check your internet connection and try again.
  pause
  exit /b 1
)

set "APP_PORT=8000"
powershell -NoProfile -Command "if (Get-NetTCPConnection -LocalPort 8000 -State Listen -ErrorAction SilentlyContinue) { exit 1 } else { exit 0 }"
if errorlevel 1 (
  set "APP_PORT=8001"
  powershell -NoProfile -Command "if (Get-NetTCPConnection -LocalPort 8001 -State Listen -ErrorAction SilentlyContinue) { exit 1 } else { exit 0 }"
  if errorlevel 1 (
    echo Ports 8000 and 8001 are already in use. Close the other app server, then run this launcher again.
    pause
    exit /b 1
  )
  echo Port 8000 is already in use. Starting this project separately on port 8001.
)

echo Starting Supply Chain Management at http://127.0.0.1:%APP_PORT%
echo Keep the server window open while using the app. Close it to stop the app.
start "Supply Chain Management server" "%~dp0.venv\Scripts\python.exe" -m uvicorn app:app --host 127.0.0.1 --port %APP_PORT%
timeout /t 3 /nobreak >nul
start "" "http://127.0.0.1:%APP_PORT%"
