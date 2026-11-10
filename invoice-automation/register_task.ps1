# Registers the scheduled task "InvoiceTool": every day at 10:00, and 5 minutes after sign-in or unlock
# (the delay gives OneDrive and the network time to be ready).
# It runs invoice_tool.py with pyw.exe, so no window is shown unless there is something to confirm.
# Exit code: 0 = all three triggers, 2 = daily 10:00 only, 1 = error.
$ErrorActionPreference = "Stop"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pyw = Get-Command pyw.exe -ErrorAction SilentlyContinue
if (-not $pyw) { exit 1 }
$user = "$env:USERDOMAIN\$env:USERNAME"
$name = "InvoiceTool"
$action = New-ScheduledTaskAction -Execute $pyw.Source -Argument "`"$dir\invoice_tool.py`"" -WorkingDirectory $dir
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive
$daily = New-ScheduledTaskTrigger -Daily -At "10:00"
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
