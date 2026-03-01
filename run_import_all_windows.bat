@echo off
setlocal
chcp 65001 >nul

set /p IMPORT_DIR=取り込みたい資料フォルダのパスを貼り付けて Enterしてください: 

if "%IMPORT_DIR%"=="" (
  echo フォルダが未入力です。終了します。
  pause
  exit /b 1
)

python egov_downloader.py --import-all-dir "%IMPORT_DIR%"
if errorlevel 1 (
  echo.
  echo 取り込みでエラーが発生しました。
  pause
  exit /b 1
)

echo.
echo 完了しました。次は run_gui_windows.bat を起動して処理開始してください。
pause
endlocal
