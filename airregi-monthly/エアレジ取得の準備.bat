@echo off
cd /d "%~dp0"
echo エアレジ自動ダウンロード用の部品をインストールします...
py -m pip install --upgrade playwright
echo.
echo 終わりました。この画面は閉じてかまいません。
pause
