@echo off
cd /d "%~dp0"
rem 前の「毎月7日／10日」の登録が残っていれば消す
schtasks /delete /tn "エアレジ月次送付" /f >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0register_task.ps1"
if %errorlevel%==0 (
  echo 登録しました。
  echo 毎日9時と、Windowsへのサインイン・ロック解除の5分後に確認し、
  echo 10日以降で先月分をまだ送っていなければ、確認画面を出します。
) else if %errorlevel%==2 (
  echo 毎日9時の確認だけ登録しました（サインイン時・ロック解除時の登録は、このパソコンでは許可されませんでした）。
  echo 9時にパソコンが使えなかった日は、次にパソコンが使えるようになってから数分後に確認画面が出ます。
) else (
  echo 登録できませんでした。この画面の内容を教えてください。
)
pause
