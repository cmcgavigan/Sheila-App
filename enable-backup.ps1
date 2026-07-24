# enable-backup.ps1 - schedule backup.py to run once a day via Task Scheduler.
# It zips the data and copies it into your Google Drive folder (set
# GDRIVE_BACKUP_DIR in .env), keeping 30 days of backups.
#
# Run once:  powershell -ExecutionPolicy Bypass -File .\enable-backup.ps1
# Disable:   Unregister-ScheduledTask -TaskName "Sheila Backup" -Confirm:$false

param([string]$ProjectDir = $PSScriptRoot)

try {
    $script = Join-Path $ProjectDir 'backup.py'
    if (-not (Test-Path $script)) {
        Write-Host "[X] backup.py not found at: $script" -ForegroundColor Red
        exit 1
    }

    $py = (Get-Command python -ErrorAction SilentlyContinue).Source
    if (-not $py) { $py = (Get-Command py -ErrorAction SilentlyContinue).Source }
    if (-not $py) {
        Write-Host "[X] Python not found on PATH. Install Python 3 first." -ForegroundColor Red
        exit 1
    }

    $taskName = "Sheila Backup"
    $action = New-ScheduledTaskAction -Execute $py -Argument "`"$script`"" -WorkingDirectory $ProjectDir

    # Daily at 02:30 (quiet hour). StartWhenAvailable catches up if the laptop was
    # off at that time - it runs at the next opportunity.
    $trigger = New-ScheduledTaskTrigger -Daily -At 2:30AM
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 15)

    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    }
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description "Daily zip of Sheila's data to Google Drive, 30-day retention." | Out-Null

    Write-Host ""
    Write-Host "[OK] Backup scheduled - runs daily at 02:30 (catches up if the laptop was off)." -ForegroundColor Green
    Write-Host ""
    Write-Host "Make sure GDRIVE_BACKUP_DIR in .env points at a real folder in your Drive." -ForegroundColor Yellow
    Write-Host "Test it now with:  python `"$script`"" -ForegroundColor DarkGray
    exit 0
}
catch {
    Write-Host "[X] Failed to schedule backup: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
