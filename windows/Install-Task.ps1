# Registers the daily task. No administrator rights needed - it runs as you,
# only while you are logged on, because it has to open a Chrome window.

$TaskName = 'ConZoL Daily Download'
$RunAt    = '08:00'

$ps1 = Join-Path $PSScriptRoot 'ConZoL-Daily.ps1'
if (-not (Test-Path $ps1)) {
    Write-Host 'ConZoL-Daily.ps1 is not next to this script - keep the two files together.'
    exit 1
}

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument ('-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $ps1 + '"')

$trigger = New-ScheduledTaskTrigger -Daily -At ([datetime]$RunAt)

# StartWhenAvailable catches up if the machine was off at the scheduled time.
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Force `
    -Description 'Opens ConZoL in Chrome so the ConZoL Auto Download userscript runs its daily check.' | Out-Null

Write-Host ''
Write-Host ('Task "' + $TaskName + '" created - runs every day at ' + $RunAt + '.')
Write-Host 'Open Task Scheduler to change the time, or run Uninstall-Task.bat to remove it.'
Write-Host ''
