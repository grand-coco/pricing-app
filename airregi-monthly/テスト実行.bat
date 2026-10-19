@echo off
cd /d "%~dp0"
rem メールは送らずに、zipとメール下書きだけを作ります（PayPayメールの取得も試します）
py airregi_monthly.py --test --mail
rem 結果をメモ帳でも開く（黒い画面を閉じても残ります）
start "" notepad "%~dp0実行結果.txt"
pause
