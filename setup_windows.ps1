# DO Bot — Windows Production Setup
# Run this script ONCE in PowerShell as Administrator
# Usage: Right-click PowerShell → "Run as Administrator", then:
#   cd "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot"
#   .\setup_windows.ps1

$ErrorActionPreference = "Stop"
$APP_DIR = "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot"
$NSSM    = "$APP_DIR\nssm.exe"

# Use sys.executable to get the real python.exe, not the Windows Store alias
# (the Store alias at WindowsApps\python.exe is inaccessible to the SYSTEM account)
$PYTHON = python -c "import sys; print(sys.executable)"
if (-not $PYTHON -or -not (Test-Path $PYTHON)) {
    Write-Host "ERROR: Could not locate real python.exe. Got: $PYTHON" -ForegroundColor Red
    exit 1
}
Write-Host "Using Python: $PYTHON" -ForegroundColor Gray

Set-Location $APP_DIR

# ── Create logs directory FIRST (NSSM needs it before starting the service) ───
New-Item -ItemType Directory -Force -Path "$APP_DIR\logs" | Out-Null
Write-Host "   Logs directory: $APP_DIR\logs" -ForegroundColor Gray

# ── Step 1: Install Python dependencies ──────────────────────────────────────
Write-Host "`n=== [1/4] Installing Python dependencies ===" -ForegroundColor Cyan
python -m pip install -r requirements.txt

# Resolve the Python site-packages path so NSSM services can import them
$SITE_PACKAGES = python -c "import site; print(site.getsitepackages()[0])" 2>$null
if (-not $SITE_PACKAGES) {
    $SITE_PACKAGES = python -c "import site; print(site.getusersitepackages())" 2>$null
}

# ── Step 2: Download NSSM (service manager) ───────────────────────────────────
Write-Host "`n=== [2/4] Downloading NSSM ===" -ForegroundColor Cyan
if (-not (Test-Path $NSSM)) {
    $nssmZip = "$env:TEMP\nssm.zip"
    Invoke-WebRequest -Uri "https://nssm.cc/release/nssm-2.24.zip" -OutFile $nssmZip
    Expand-Archive -Path $nssmZip -DestinationPath "$env:TEMP\nssm_extract" -Force
    Copy-Item "$env:TEMP\nssm_extract\nssm-2.24\win64\nssm.exe" $NSSM
    Write-Host "   nssm.exe saved to $NSSM"
} else {
    Write-Host "   nssm.exe already present, skipping."
}

# ── Step 3: Install DO Bot as Windows Service ─────────────────────────────────
Write-Host "`n=== [3/4] Installing DO Bot service ===" -ForegroundColor Cyan
try { & $NSSM stop   do_bot         2>$null } catch {}; $LASTEXITCODE = 0
try { & $NSSM remove do_bot confirm 2>$null } catch {}; $LASTEXITCODE = 0

& $NSSM install do_bot $PYTHON
& $NSSM set     do_bot AppParameters    "$APP_DIR\app.py"
& $NSSM set     do_bot AppDirectory     $APP_DIR
& $NSSM set     do_bot DisplayName      "DO Bot - WhatsApp Lorry Assignment"
& $NSSM set     do_bot Description      "WhatsApp AI Logistics bot (Flask + waitress)"
& $NSSM set     do_bot Start            SERVICE_AUTO_START
& $NSSM set     do_bot AppStdout        "$APP_DIR\logs\do_bot.log"
& $NSSM set     do_bot AppStderr        "$APP_DIR\logs\do_bot_error.log"
& $NSSM set     do_bot AppRotateFiles   1
& $NSSM set     do_bot AppRotateBytes   5242880
& $NSSM set     do_bot AppNoConsole     1

# Run as the current user so it can access user-scoped Python and packages.
# Services default to SYSTEM which cannot reach AppData\Local installs.
$svcUser = ".\$env:USERNAME"
Write-Host "   Service account: $svcUser" -ForegroundColor Gray
Write-Host "   Enter your Windows login password when prompted:" -ForegroundColor Yellow
$svcCred = Get-Credential -UserName $svcUser -Message "Password for DO Bot service account ($svcUser)"
& $NSSM set do_bot ObjectName $svcCred.UserName $svcCred.GetNetworkCredential().Password

# Pass Python's site-packages path as an extra safety net
if ($SITE_PACKAGES) {
    & $NSSM set do_bot AppEnvironmentExtra "PYTHONPATH=$SITE_PACKAGES"
    Write-Host "   PYTHONPATH set to: $SITE_PACKAGES" -ForegroundColor Gray
}

# ── Step 4: Install ngrok as Windows Service ──────────────────────────────────
Write-Host "`n=== [4/4] Installing ngrok service ===" -ForegroundColor Cyan

$ngrokCmd  = Get-Command ngrok -ErrorAction SilentlyContinue
$ngrokPath = if ($ngrokCmd) { $ngrokCmd.Source } else { $null }
if (-not $ngrokPath) {
    Write-Host "   ngrok not found in PATH." -ForegroundColor Yellow
    Write-Host "   Download from https://ngrok.com/download, extract ngrok.exe to $APP_DIR" -ForegroundColor Yellow
    $ngrokPath = "$APP_DIR\ngrok.exe"
}

try { & $NSSM stop   ngrok_do_bot         2>$null } catch {}; $LASTEXITCODE = 0
try { & $NSSM remove ngrok_do_bot confirm 2>$null } catch {}; $LASTEXITCODE = 0

& $NSSM install ngrok_do_bot $ngrokPath
& $NSSM set     ngrok_do_bot AppParameters   "start do_bot --config `"$APP_DIR\ngrok_do_bot.yml`""
& $NSSM set     ngrok_do_bot AppDirectory    $APP_DIR
& $NSSM set     ngrok_do_bot DisplayName     "DO Bot - ngrok Tunnel"
& $NSSM set     ngrok_do_bot Start           SERVICE_AUTO_START
& $NSSM set     ngrok_do_bot AppStdout       "$APP_DIR\logs\ngrok.log"
& $NSSM set     ngrok_do_bot AppStderr       "$APP_DIR\logs\ngrok_error.log"
& $NSSM set     ngrok_do_bot AppNoConsole    1
& $NSSM set     ngrok_do_bot DependOnService do_bot
& $NSSM set     ngrok_do_bot ObjectName      $svcCred.UserName $svcCred.GetNetworkCredential().Password

Write-Host "`n=== Setup complete! ===" -ForegroundColor Green
Write-Host ""
Write-Host "Before starting, make sure:" -ForegroundColor Yellow
Write-Host "  1. config.txt has your real META_ACCESS_TOKEN and META_PHONE_NUMBER_ID"
Write-Host "  2. ngrok_do_bot.yml has your real ngrok authtoken"
Write-Host "     (copy from ngrok_do_bot.yml.example and fill in the token)"
Write-Host "     Get your token at: https://dashboard.ngrok.com/get-started/your-authtoken"
Write-Host ""
Write-Host "Then start both services:" -ForegroundColor Cyan
Write-Host "  Start-Service do_bot"
Write-Host "  Start-Service ngrok_do_bot"
Write-Host ""
Write-Host "Check logs in: $APP_DIR\logs\"
Write-Host ""
Write-Host "If services fail to start, run: .\diagnose_windows.ps1" -ForegroundColor Yellow
