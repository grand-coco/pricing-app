# Registers a daily Windows scheduled task that runs invoice_tool.py.
$ErrorActionPreference = "Stop"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pyw = Get-Command pyw.exe -ErrorAction SilentlyContinue
if (-not $pyw) {
    Write-Host "ERROR: pyw.exe was not found. Please install Python from python.org first." -ForegroundColor Red
    exit 1
}
$action   = New-ScheduledTaskAction -Execute $pyw.Source -Argument "`"$dir\invoice_tool.py`"" -WorkingDirectory $dir
$trigger  = New-ScheduledTaskTrigger -Daily -At "10:00"
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName "InvoiceTool" -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
Write-Host ""
Write-Host "OK: The task 'InvoiceTool' was registered. It runs every day at 10:00." -ForegroundColor Green
