@echo off
cd /d "%~dp0"
echo 必要な部品をインストールします...
py -m pip install --upgrade pandas openpyxl
if errorlevel 1 (
  echo.
  echo 部品のインストールに失敗しました。この画面の内容を教えてください。
  pause
  exit /b 1
)
echo.
echo ヤフーメールへの接続を確認します...
py airregi_monthly.py --check
pause
