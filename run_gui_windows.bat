@echo off
setlocal ENABLEDELAYEDEXPANSION

REM NoticeForge GUI - Windows 起動スクリプト

if not exist ".venv\Scripts\python.exe" (
  echo [INFO] 初回セットアップ中（少し時間がかかります）...
  python -m venv .venv
  if errorlevel 1 (
    echo [ERROR] Pythonの仮想環境の作成に失敗しました。
    echo         Python がインストールされているか確認してください。
    pause
    exit /b 1
  )
)

call .venv\Scripts\activate.bat
if errorlevel 1 (
  echo [ERROR] 仮想環境の起動に失敗しました。
  pause
  exit /b 1
)

echo [INFO] パッケージを確認しています...
.venv\Scripts\python.exe -m pip install --upgrade pip -q
.venv\Scripts\python.exe -m pip install -r requirements.txt -q
if errorlevel 1 (
  echo [ERROR] 必要なパッケージのインストールに失敗しました。
  echo         ネットワーク接続を確認して再試行してください。
  pause
  exit /b 1
)

REM GUIをコンソールウィンドウなしで起動 (pythonw.exe = ウィンドウなしPython)
start "" .venv\Scripts\pythonw.exe noticeforge_gui.py
