@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0register_task.ps1"
if %errorlevel%==0 (
  echo 登録しました。
  echo 毎日10時と、Windowsへのサインイン・ロック解除の5分後に、請求書メールを確認します。
  echo 20日以降でまだ送っていなければ、送信の確認メッセージを出します。
) else if %errorlevel%==2 (
  echo 毎日10時の確認だけ登録しました（サインイン時・ロック解除時の登録は、このパソコンでは許可されませんでした）。
  echo 10時にパソコンが使えなかった日は、次にパソコンが使えるようになったときに確認します。
) else (
  echo 登録できませんでした。Python（pyw.exe）が入っているか確認して、この画面の内容を教えてください。
)
pause
