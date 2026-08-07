# install_service.ps1 — install the Lorry Assignment web app as a 24/7 Windows
# service using NSSM. Run in an ELEVATED PowerShell (Administrator), from the
# project folder:  .\install_service.ps1
#
# Re-running is safe: it removes any existing service first, then reinstalls.

$ErrorActionPreference = "Stop"
$Name    = "LorryAssignment"
$Port    = 8000
$AppDir  = $PSScriptRoot
$Script  = Join-Path $AppDir "web_app.py"
$Nssm    = Join-Path $AppDir "nssm.exe"

Write-Host "== Lorry Assignment service installer ==" -ForegroundColor Cyan

# 1. Locate NSSM
if (-not (Test-Path $Nssm)) {
    throw "nssm.exe not found in $AppDir. Download it from https://nssm.cc/download and copy the win64\nssm.exe here."
}

# 2. Locate the REAL python.exe — NOT the WindowsApps Store alias stub, which
#    cannot run as a service.
$python = $null
$candidates = @()
Get-Command python.exe -All -ErrorAction SilentlyContinue | ForEach-Object { $candidates += $_.Source }
$candidates += (Get-ChildItem "$env:LOCALAPPDATA\Python\*\python.exe" -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
$candidates += (Get-ChildItem "$env:LOCALAPPDATA\Programs\Python\*\python.exe" -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
foreach ($c in $candidates) {
    if ($c -and (Test-Path $c) -and ($c -notmatch "WindowsApps")) { $python = $c; break }
}
if (-not $python) { throw "Could not find a real python.exe (only the WindowsApps stub). Install Python from python.org and re-run." }
Write-Host "Using Python: $python" -ForegroundColor Green
Write-Host "App script:   $Script"

# 3. Verify the web app file exists
if (-not (Test-Path $Script)) { throw "web_app.py not found in $AppDir. Are you in the project folder?" }

# 4. Remove any existing service, then (re)install
& $Nssm stop   $Name 2>$null | Out-Null
& $Nssm remove $Name confirm 2>$null | Out-Null

& $Nssm install $Name "$python" "$Script"
& $Nssm set $Name AppDirectory "$AppDir"
& $Nssm set $Name Start SERVICE_AUTO_START
$logDir = Join-Path $AppDir "logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
& $Nssm set $Name AppStdout (Join-Path $logDir "web.log")
& $Nssm set $Name AppStderr (Join-Path $logDir "web.log")

# 5. Firewall rule (idempotent)
if (-not (Get-NetFirewallRule -DisplayName "Lorry Assignment $Port" -ErrorAction SilentlyContinue)) {
    New-NetFirewallRule -DisplayName "Lorry Assignment $Port" -Direction Inbound -Protocol TCP -LocalPort $Port -Action Allow | Out-Null
    Write-Host "Firewall opened for port $Port." -ForegroundColor Green
}

# 6. Start it
& $Nssm start $Name
Start-Sleep -Seconds 2
$status = & $Nssm status $Name
Write-Host "Service status: $status" -ForegroundColor Yellow

$ip = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
       Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.*" } |
       Select-Object -First 1).IPAddress
Write-Host ""
Write-Host "Done. Open:  http://$ip`:$Port/login   (or http://127.0.0.1:$Port/login on this PC)" -ForegroundColor Cyan
Write-Host "Manage:  .\nssm.exe restart $Name  |  .\nssm.exe stop $Name  |  .\nssm.exe status $Name"
