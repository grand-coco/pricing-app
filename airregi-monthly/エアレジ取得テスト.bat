@echo off
cd /d "%~dp0"
rem ブラウザの動きを表示しながら、先月分を「取得テスト」フォルダにダウンロードします
py airregi_download.py --test
start "" explorer "%~dp0取得テスト"
if exist "%~dp0取得ログ" start "" explorer "%~dp0取得ログ"
pause
