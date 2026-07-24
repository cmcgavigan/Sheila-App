# Sheila App v2 - installer (run elevated; INSTALL.cmd does that for you).
# Idempotent: safe to run again after an update or to repair the install.
#
# Optional: -V1Source "C:\path\to\old Sheilas app" runs a READ-ONLY migration
# dry run against that explicit folder. The installer never guesses where v1
# lives and never migrates automatically.
param(
    [string]$V1Source = ''
)
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot   # project folder (parent of install\)
$ServiceName = 'SheilaApp'
Set-Location $Root

function Say($msg) { Write-Host "  $msg" -ForegroundColor Cyan }
function Ok($msg)  { Write-Host "  [OK] $msg" -ForegroundColor Green }
function Fail($msg){ Write-Host "  [!!] $msg" -ForegroundColor Red; Read-Host 'Press Enter to close'; exit 1 }

Write-Host ''
Write-Host '  Sheila App v2 - installer' -ForegroundColor Magenta
Write-Host '  =========================' -ForegroundColor Magenta

# --- 1. Python ---------------------------------------------------------------
Say 'Checking Python...'
$py = $null
foreach ($cand in @('py -3', 'python', 'python3')) {
    try {
        $v = Invoke-Expression "$cand --version" 2>$null
        if ($v -match 'Python 3\.(\d+)') { if ([int]$Matches[1] -ge 9) { $py = $cand; break } }
    } catch {}
}
if (-not $py) { Fail 'Python 3.9+ not found. Install it from python.org (tick "Add to PATH"), then run INSTALL.cmd again.' }
Ok "Python found ($py)"

# --- 2. venv + dependencies --------------------------------------------------
Say 'Creating virtual environment + installing dependencies...'
if (-not (Test-Path "$Root\.venv")) { Invoke-Expression "$py -m venv `"$Root\.venv`"" }
$venvPy = "$Root\.venv\Scripts\python.exe"
& $venvPy -m pip install --quiet --upgrade pip
& $venvPy -m pip install --quiet -r "$Root\requirements.txt"
if ($LASTEXITCODE -ne 0) { Fail 'pip install failed - check your internet connection and rerun.' }
Ok 'Dependencies installed'

# --- 3. .env -----------------------------------------------------------------
if (-not (Test-Path "$Root\.env")) {
    $parentEnv = Join-Path (Split-Path -Parent $Root) '.env'
    if (Test-Path $parentEnv) {
        Say 'No .env yet - copying keys from the v1 app and adjusting ports...'
        $content = Get-Content $parentEnv -Raw
        $content = $content -replace '(?m)^PORT=.*$', 'PORT=3002'
        $content = $content -replace '(?m)^TS_HTTPS_PORT=.*$', 'TS_HTTPS_PORT=8444'
        Set-Content "$Root\.env" $content -Encoding UTF8
        Ok '.env created from v1 (PORT=3002, TS_HTTPS_PORT=8444 so both apps can run)'
    } else {
        Copy-Item "$Root\.env.example" "$Root\.env"
        Ok '.env created from template - EDIT IT to add your GROQ_API_KEY before use'
    }
}

# --- 3a. Authentication secrets -----------------------------------------------
# Fresh installs must not inherit the historical default PIN or start without a
# usable login secret. Values are written only to the local .env file.
$envText = Get-Content "$Root\.env" -Raw
if ($envText -notmatch '(?m)^AUTH_PASSWORD=') {
    $authSecret = (& $venvPy -c "import secrets; print(secrets.token_urlsafe(24))").Trim()
    Add-Content "$Root\.env" "`nAUTH_PASSWORD=$authSecret"
    Ok 'Generated a strong AUTH_PASSWORD in .env (use it to sign in)'
}
if ($envText -notmatch '(?m)^TREATMENTS_PIN=') {
    $treatPin = (& $venvPy -c "import secrets; print(secrets.randbelow(90000000)+10000000)").Trim()
    Add-Content "$Root\.env" "TREATMENTS_PIN=$treatPin"
    Ok 'Generated a unique treatment-editor PIN in .env'
}

# --- 3b. v1 data migration check (READ-ONLY dry run, explicit source only) -----
# Never migrates automatically and never guesses where the v1 app lives: the
# dry run only happens when the caller explicitly passed -V1Source. Applying
# is always a deliberate, separate command.
if ($V1Source) {
    if (-not (Test-Path (Join-Path $V1Source 'her-expenses.xlsx'))) {
        Say "-V1Source '$V1Source' has no her-expenses.xlsx - not a v1 app folder, skipping the migration dry run."
    } else {
        Say "Running migration DRY RUN against '$V1Source' (reads only, changes nothing)..."
        try {
            & $venvPy -m app.migrate --source $V1Source
            if ($LASTEXITCODE -ne 0) { Say '(dry run reported a problem - see above; the install continues)' }
            Write-Host ''
            Write-Host '  To import that v1 data into v2, run this yourself afterwards:' -ForegroundColor Yellow
            Write-Host "    & `"$venvPy`" -m app.migrate --source `"$V1Source`" --apply" -ForegroundColor Yellow
            Write-Host '  (dry-run first is the default; --apply is required to write anything)' -ForegroundColor Yellow
            Write-Host ''
        } catch { Say "(migration dry run failed: $_ - the install continues)" }
    }
} elseif (-not (Test-Path "$Root\data\sheila.db")) {
    Say 'No -V1Source given, so no migration preview was run. To preview importing v1 data:'
    Say '  powershell -ExecutionPolicy Bypass -File install\install.ps1 -V1Source "C:\path\to\old Sheilas app"'
}

# --- 4. NSSM (service manager) -------------------------------------------------
$nssm = "$Root\install\nssm.exe"
if (-not (Test-Path $nssm)) {
    Say 'Fetching NSSM (service manager)...'
    $nssmUrl = 'https://nssm.cc/release/nssm-2.24.zip'
    $nssmSha256 = '727D1E42275C605E0F04ABA98095C38A8E1E46DEF453CDFFCE42869428AA6743'
    $zip = Join-Path $env:TEMP 'nssm-2.24.zip'
    $nssmTemp = Join-Path $env:TEMP ("nssm-" + [guid]::NewGuid().ToString('N'))
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest $nssmUrl -OutFile $zip -UseBasicParsing
        $actual = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
        if ($actual -ne $nssmSha256) { throw "NSSM SHA-256 mismatch: got $actual" }
        Expand-Archive $zip $nssmTemp -Force
        Copy-Item "$nssmTemp\nssm-2.24\win64\nssm.exe" $nssm
    } catch { Fail "Could not download or verify NSSM ($_). Download the pinned artifact yourself and put nssm.exe into the install\ folder, then rerun." }
    finally { Remove-Item $zip -Force -ErrorAction SilentlyContinue; Remove-Item $nssmTemp -Recurse -Force -ErrorAction SilentlyContinue }
}
Ok 'NSSM ready'

# --- 5. Install / update the service ------------------------------------------
Say "Installing Windows service '$ServiceName'..."
& $nssm stop $ServiceName 2>$null | Out-Null
& $nssm remove $ServiceName confirm 2>$null | Out-Null
& $nssm install $ServiceName $venvPy "`"$Root\run.py`"" | Out-Null
& $nssm set $ServiceName AppDirectory $Root | Out-Null
& $nssm set $ServiceName ObjectName 'LocalService' | Out-Null
& $nssm set $ServiceName DisplayName 'Sheila PMU Receipts (v2)' | Out-Null
& $nssm set $ServiceName Description 'Receipts/expenses/trips PWA server for Sheila - starts at boot, restarts on crash.' | Out-Null
& $nssm set $ServiceName Start SERVICE_AUTO_START | Out-Null
& $nssm set $ServiceName AppStdout "$Root\data\logs\service.log" | Out-Null
& $nssm set $ServiceName AppStderr "$Root\data\logs\service.log" | Out-Null
& $nssm set $ServiceName AppRotateFiles 1 | Out-Null
& $nssm set $ServiceName AppRotateBytes 5242880 | Out-Null
& $nssm set $ServiceName AppExit Default Restart | Out-Null
& $nssm set $ServiceName AppRestartDelay 5000 | Out-Null
New-Item -ItemType Directory -Force -Path "$Root\data\logs" | Out-Null
icacls "$Root\data" /grant 'NT AUTHORITY\LOCAL SERVICE:(OI)(CI)M' /T /C | Out-Null
Ok 'Service installed'

# --- 6. Firewall rule ----------------------------------------------------------
$port = 3002
if (Test-Path "$Root\.env") {
    $m = Select-String -Path "$Root\.env" -Pattern '^PORT=(\d+)' | Select-Object -First 1
    if ($m) { $port = [int]$m.Matches[0].Groups[1].Value }
}
Say "Opening firewall port $port..."
netsh advfirewall firewall delete rule name="Sheila App v2" | Out-Null
netsh advfirewall firewall add rule name="Sheila App v2" dir=in action=allow protocol=TCP localport=$port profile=private,domain | Out-Null
Ok 'Firewall rule set'

# --- 7. Start ------------------------------------------------------------------
Say 'Starting the service...'
& $nssm start $ServiceName | Out-Null
Start-Sleep -Seconds 4
$svc = Get-Service $ServiceName -ErrorAction SilentlyContinue
if ($svc -and $svc.Status -eq 'Running') {
    Ok 'Service is RUNNING'
    Write-Host ''
    Write-Host "  Open on this laptop:  https://localhost:$port" -ForegroundColor Yellow
    Write-Host "  Phone setup QR wall:  https://localhost:$port/setup" -ForegroundColor Yellow
    Write-Host "  Logs:                 $Root\data\logs\service.log" -ForegroundColor Yellow
} else {
    Fail "Service did not start - check $Root\data\logs\service.log"
}
Write-Host ''
Read-Host 'Done. Press Enter to close'
