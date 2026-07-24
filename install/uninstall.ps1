# Removes the SheilaApp service and firewall rule. Data (data\) is NOT touched.
$ErrorActionPreference = 'SilentlyContinue'
$Root = Split-Path -Parent $PSScriptRoot
$nssm = "$Root\install\nssm.exe"
if (Test-Path $nssm) {
    & $nssm stop SheilaApp | Out-Null
    & $nssm remove SheilaApp confirm | Out-Null
}
netsh advfirewall firewall delete rule name="Sheila App v2" | Out-Null
Write-Host 'Service and firewall rule removed. Data folder left untouched.'
Read-Host 'Press Enter to close'
