# enable-autostart.ps1 — make BOTH receipt servers start automatically when you
# log in to Windows: your Receipts server (green tray) AND Sheila's server (gold
# tray). It installs ONE Startup shortcut that launches both via
# start-both-servers.vbs, and removes older single-app Startup shortcuts so
# nothing double-launches.
#
# Run once, from this folder, in PowerShell:
#     powershell -ExecutionPolicy Bypass -File .\enable-autostart.ps1
#
# To DISABLE later: delete "Receipt Servers.lnk" from your Startup folder
#     (Win+R -> shell:startup -> delete the shortcut)

param([string]$ProjectDir = $PSScriptRoot)

try {
    $ProjectDir = $ProjectDir -replace '[\\"]+$', ''
    $target = Join-Path $ProjectDir 'start-both-servers.vbs'
    if (-not (Test-Path $target)) {
        Write-Host "[X] start-both-servers.vbs not found at: $target" -ForegroundColor Red
        exit 1
    }

    $w       = New-Object -ComObject WScript.Shell
    $startup = [Environment]::GetFolderPath('Startup')

    # Remove older single-app Startup shortcuts so we don't launch twice.
    foreach ($old in @('Sheila Server.lnk', 'Receipts Server.lnk')) {
        $p = Join-Path $startup $old
        if (Test-Path $p) { Remove-Item $p -Force; Write-Host "  Removed old shortcut: $old" -ForegroundColor DarkGray }
    }

    $linkPath = Join-Path $startup 'Receipt Servers.lnk'
    $sc = $w.CreateShortcut($linkPath)
    $sc.TargetPath       = 'wscript.exe'
    $sc.Arguments        = '"' + $target + '"'
    $sc.WorkingDirectory = $ProjectDir
    $sc.Description      = 'Auto-start both receipt servers (Receipts + Sheila) at login'
    $sc.WindowStyle      = 7
    $sc.Save()

    Write-Host ""
    Write-Host "[OK] Both servers will now start at login." -ForegroundColor Green
    Write-Host "     Shortcut: $linkPath"
    Write-Host "     Tray: GREEN = Receipts, GOLD = Sheila." -ForegroundColor Gray
    Write-Host ""
    Write-Host "TIP: stop the laptop sleeping, or the servers are unreachable while asleep:" -ForegroundColor Yellow
    Write-Host "     Settings > System > Power > 'When plugged in, put my device to sleep' = Never" -ForegroundColor DarkGray
    exit 0
}
catch {
    Write-Host "[X] Failed to set auto-start: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
