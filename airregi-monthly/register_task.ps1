# Registers the scheduled task that checks every day at 9:00, and 5 minutes after sign-in or unlock
# (the delay gives OneDrive time to sync the invoices),
# whether last month's data still has to be sent (airregi_monthly.py --auto decides; it shows nothing otherwise).
# Exit code: 0 = all three triggers, 2 = daily 9:00 only, 1 = error.
$ErrorActionPreference = "Stop"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pyw = Get-Command pyw.exe -ErrorAction SilentlyContinue
if (-not $pyw) { exit 1 }
$user = "$env:USERDOMAIN\$env:USERNAME"
$name = "AirregiMonthly"
$action = New-ScheduledTaskAction -Execute $pyw.Source -Argument "`"$dir\airregi_monthly.py`" --auto" -WorkingDirectory $dir
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive
$daily = New-ScheduledTaskTrigger -Daily -At "09:00"
try {
    $logon = New-ScheduledTaskTrigger -AtLogOn -User $user
    $logon.Delay = "PT5M"
    $cls = Get-CimClass -Namespace ROOT\Microsoft\Windows\TaskScheduler -ClassName MSFT_TaskSessionStateChangeTrigger
    $unlock = New-CimInstance -CimClass $cls -ClientOnly
    $unlock.StateChange = 8  # unlock
    $unlock.UserId = $user
    $unlock.Enabled = $true
    $unlock.Delay = "PT5M"
    Register-ScheduledTask -TaskName $name -Action $action -Trigger @($daily, $logon, $unlock) -Settings $settings -Principal $principal -Force | Out-Null
    exit 0
} catch {
    try {
        Register-ScheduledTask -TaskName $name -Action $action -Trigger $daily -Settings $settings -Principal $principal -Force | Out-Null
        exit 2
    } catch { exit 1 }
}
