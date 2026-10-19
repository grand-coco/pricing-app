@echo off
cd /d "%~dp0"
rem メールは送らずに、zipとメール下書きだけを作ります（PayPayメールの取得も試します）
py airregi_monthly.py --test --mail
pause
