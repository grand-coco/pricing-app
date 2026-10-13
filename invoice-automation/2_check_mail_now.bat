@echo off
chcp 65001 > nul
py "%~dp0invoice_tool.py" --collect
pause
