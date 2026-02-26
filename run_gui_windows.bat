@echo off
setlocal ENABLEDELAYEDEXPANSION

REM ============================================================
REM  NoticeForge GUI - Windows Launcher
REM ============================================================

if not exist ".venv\Scripts\python.exe" (
  echo [INFO] Creating virtual environment...
  python -m venv .venv
  if errorlevel 1 (
    echo [ERROR] Failed to create venv. Please ensure Python is installed.
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

echo [INFO] Checking packages...
.venv\Scripts\python.exe -m pip install --upgrade pip -q
.venv\Scripts\python.exe -m pip install -r requirements.txt -q
if errorlevel 1 (
  echo [ERROR] Package installation failed. Check network and retry.
  pause
  exit /b 1
)

echo [INFO] Launching NoticeForge...

REM --- Launch without console window ---
REM Method 1: pythonw.exe (no console, available in most Python installs)
if exist ".venv\Scripts\pythonw.exe" (
  start "" .venv\Scripts\pythonw.exe noticeforge_gui.py
  exit /b 0
)

REM Method 2: VBScript launcher (hides console window for python.exe)
set VBS=%TEMP%\nf_launch_%RANDOM%.vbs
echo Set sh = CreateObject("WScript.Shell") > "%VBS%"
echo sh.Run """" ^& sh.CurrentDirectory ^& "\.venv\Scripts\python.exe" ^& """ noticeforge_gui.py", 0, False >> "%VBS%"
wscript "%VBS%"
del "%VBS%" 2>nul
