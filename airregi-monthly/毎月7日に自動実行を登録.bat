@echo off
cd /d "%~dp0"
schtasks /create /tn "エアレジ月次送付" /tr "\"%~dp0今すぐ送信.bat\"" /sc monthly /d 7 /st 09:00 /f
rem 7日の9時にパソコンが切れていた場合、次に起動したときに実行する
powershell -NoProfile -Command "$t = Get-ScheduledTask -TaskName 'エアレジ月次送付'; $t.Settings.StartWhenAvailable = $true; Set-ScheduledTask -InputObject $t | Out-Null"
echo.
echo 毎月7日 9:00 に実行するよう登録しました。
pause
