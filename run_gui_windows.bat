@echo off
setlocal ENABLEDELAYEDEXPANSION

REM NoticeForge Modern GUI v3.1 - Windows runner

if not exist ".venv\Scripts\python.exe" (
  echo [INFO] Creating venv...
  python -m venv .venv
  if errorlevel 1 (
    echo [ERROR] Failed to create venv. Ensure Python is installed and 'python' works.
    pause
    exit /b 1
  )
)

call .venv\Scripts\activate.bat
if errorlevel 1 (
  echo [ERROR] Failed to activate venv.
  pause
  exit /b 1
)

echo [INFO] Upgrading pip...
.venv\Scripts\python.exe -m pip install --upgrade pip

echo [INFO] Installing requirements...
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
  echo [ERROR] pip install failed. Check network/proxy and retry.
  pause
  exit /b 1
)

echo [INFO] Launching NoticeForge Modern GUI v3.1...
.venv\Scripts\python.exe noticeforge_gui.py

pause
