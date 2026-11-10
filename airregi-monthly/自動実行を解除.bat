@echo off
schtasks /delete /tn "AirregiMonthly" /f
schtasks /delete /tn "エアレジ月次送付" /f >nul 2>&1
pause
