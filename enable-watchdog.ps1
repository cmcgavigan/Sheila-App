# enable-watchdog.ps1 - schedule watchdog.py every 4 hours via Task Scheduler,
# so it keeps the receipt servers alive and alerts you (ntfy) if one stays down
# for over a day of actual uptime.
#
# Run once, from this folder:
#     powershell -ExecutionPolicy Bypass -File .\enable-watchdog.ps1
#
# To DISABLE later:
#     Unregister-ScheduledTask -TaskName "Sheila Watchdog" -Confirm:$false
#
# Set NTFY_TOPIC in .env and subscribe to it in the ntfy phone app, or it can
# monitor + restart but not alert you.

param([string]$ProjectDir = $PSScriptRoot)

try {
    $script = Join-Path $ProjectDir 'watchdog.py'
    if (-not (Test-Path $script)) {
        Write-Host "[X] watchdog.py not found at: $script" -ForegroundColor Red
        exit 1
    }

    $py = (Get-Command python -ErrorAction SilentlyContinue).Source
    if (-not $py) { $py = (Get-Command py -ErrorAction SilentlyContinue).Source }
    if (-not $py) {
        Write-Host "[X] Python not found on PATH. Install Python 3 first." -ForegroundColor Red
        exit 1
    }

    $taskName = "Sheila Watchdog"
    $action = New-ScheduledTaskAction -Execute $py -Argument "`"$script`"" -WorkingDirectory $ProjectDir

    $trigDaily = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5) -RepetitionInterval (New-TimeSpan -Hours 4)
    $trigLogon = New-ScheduledTaskTrigger -AtLogOn

    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    }
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($trigDaily, $trigLogon) -Settings $settings -Description "Health-check + auto-restart the receipt servers; ntfy alert if down 24h." | Out-Null

    Write-Host ""
    Write-Host "[OK] Watchdog scheduled - runs every 4 hours + at logon." -ForegroundColor Green
    Write-Host "     Logs: $ProjectDir\watchdog.log" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Test it now with:  python `"$script`" --dry-run" -ForegroundColor DarkGray
    exit 0
}
catch {
    Write-Host "[X] Failed to schedule watchdog: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
